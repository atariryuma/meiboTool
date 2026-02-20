"""レイアウトエディター メインウィンドウ

LayFile を視覚的に編集する CTkToplevel ウィンドウ。
Canvas 上でオブジェクトを選択・移動・リサイズし、
プロパティパネルで属性を編集する。
"""

from __future__ import annotations

import dataclasses
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from typing import TYPE_CHECKING

import customtkinter as ctk

from core.lay_parser import (
    LayFile,
    LayoutObject,
    ObjectType,
    new_field,
    new_label,
    new_line,
)
from core.lay_serializer import load_layout, save_layout
from gui.editor.layout_canvas import LayoutCanvas
from gui.editor.object_list import ObjectListPanel
from gui.editor.properties_panel import PropertiesPanel
from gui.editor.toolbar import EditorToolbar

if TYPE_CHECKING:
    pass

# ── 定数 ─────────────────────────────────────────────────────────────────────

_FILE_TYPES = [
    ('レイアウトファイル', '*.json *.lay'),
    ('JSON レイアウト', '*.json'),
    ('スズキ校務 .lay', '*.lay'),
    ('すべて', '*.*'),
]

_SAVE_FILE_TYPES = [
    ('JSON レイアウト', '*.json'),
]

# Undo スナップショット
_UndoEntry = tuple[int, LayoutObject, LayoutObject]  # (index, before, after)


class EditorWindow(ctk.CTkToplevel):
    """レイアウトエディター メインウィンドウ。"""

    def __init__(
        self, master: ctk.CTk,
        lay: LayFile | None = None,
    ) -> None:
        super().__init__(master)
        self.title('レイアウトエディター')
        self.geometry('1200x800')
        self.transient(master)

        self._lay = lay or LayFile(
            title='新規レイアウト',
            page_width=840,
            page_height=1188,
        )
        self._current_file: str | None = None
        self._dirty = False
        self._undo_stack: list[_UndoEntry] = []
        self._redo_stack: list[_UndoEntry] = []
        self._add_mode: str | None = None  # 'label', 'field', 'line' or None

        self._build_ui()
        self._canvas_panel.set_layout(self._lay)
        self._object_list.set_layout(self._lay)
        self._update_status()

    # ── UI 構築 ───────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── ツールバー ──
        callbacks = {
            'new': self._on_new,
            'open': self._on_open,
            'save': self._on_save,
            'save_as': self._on_save_as,
            'undo': self._on_undo,
            'redo': self._on_redo,
            'add_label': self._on_add_label,
            'add_field': self._on_add_field,
            'add_line': self._on_add_line,
            'delete': self._on_delete,
            'data_fill': self._on_data_fill,
            'zoom_in': self._on_zoom_in,
            'zoom_out': self._on_zoom_out,
        }
        self._toolbar = EditorToolbar(self, callbacks)
        self._toolbar.grid(row=0, column=0, sticky='ew')

        # ── メインエリア ──
        main = ctk.CTkFrame(self)
        main.grid(row=1, column=0, sticky='nsew')
        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=1)

        # Object List (左サイド)
        self._object_list = ObjectListPanel(
            main, on_select=self._on_list_select,
        )
        self._object_list.grid(row=0, column=0, sticky='ns', padx=(2, 0), pady=2)

        # Canvas (中央)
        self._canvas_panel = LayoutCanvas(
            main,
            on_select=self._on_object_selected,
            on_modify=self._on_object_modified,
            on_cursor=self._on_cursor_move,
        )
        self._canvas_panel.grid(row=0, column=1, sticky='nsew')

        # Canvas クリック（追加モード用）
        self._canvas_panel.canvas.bind('<Button-1>', self._on_canvas_click, add=True)

        # Properties Panel (右サイド)
        self._props = PropertiesPanel(main, on_change=self._on_prop_change)
        self._props.grid(row=0, column=2, sticky='ns', padx=(0, 2), pady=2)

        # ── ステータスバー ──
        self._status = ctk.CTkLabel(
            self, text='', height=24, anchor='w',
            font=ctk.CTkFont(size=11),
        )
        self._status.grid(row=2, column=0, sticky='ew', padx=8)

        # ── キーバインド ──
        self.bind('<Delete>', lambda _e: self._on_delete())
        self.bind('<Control-z>', lambda _e: self._on_undo())
        self.bind('<Control-y>', lambda _e: self._on_redo())
        self.bind('<Control-s>', lambda _e: self._on_save())
        self.bind('<Control-o>', lambda _e: self._on_open())
        self.bind('<Control-n>', lambda _e: self._on_new())
        self.bind('<Escape>', lambda _e: self._cancel_add_mode())

    # ── コールバック ─────────────────────────────────────────────────────

    def _on_object_selected(self, index: int) -> None:
        """Canvas でオブジェクト選択時。"""
        if 0 <= index < len(self._lay.objects):
            self._props.set_object(self._lay.objects[index], index)
            self._object_list.select(index)
        else:
            self._props.set_object(None)
        self._update_status()

    def _on_list_select(self, index: int) -> None:
        """オブジェクト一覧で選択時。"""
        if 0 <= index < len(self._lay.objects):
            self._canvas_panel.select(index)
            self._props.set_object(self._lay.objects[index], index)
        self._update_status()

    def _on_object_modified(self) -> None:
        """ドラッグ移動/リサイズ完了時。"""
        self._dirty = True
        self._object_list.refresh()
        self._update_status()

    def _on_prop_change(self, obj: LayoutObject) -> None:
        """プロパティパネルで値が変更された時。"""
        self._dirty = True
        self._canvas_panel.refresh()
        self._object_list.refresh()
        self._update_status()

    def _on_cursor_move(self, mx: int, my: int) -> None:
        """カーソル移動時（モデル座標）。"""
        mm_x = mx * 0.25
        mm_y = my * 0.25
        self._update_status(cursor=(mx, my, mm_x, mm_y))

    # ── ファイル操作 ─────────────────────────────────────────────────────

    def _on_new(self) -> None:
        if self._dirty and not self._confirm_discard():
            return
        self._lay = LayFile(
            title='新規レイアウト',
            page_width=840,
            page_height=1188,
        )
        self._current_file = None
        self._dirty = False
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._canvas_panel.set_layout(self._lay)
        self._object_list.set_layout(self._lay)
        self._props.set_object(None)
        self._update_status()

    def _on_open(self) -> None:
        if self._dirty and not self._confirm_discard():
            return

        path = fd.askopenfilename(
            title='レイアウトを開く',
            filetypes=_FILE_TYPES,
        )
        if not path:
            return

        try:
            if path.lower().endswith('.lay'):
                from core.lay_parser import parse_lay
                self._lay = parse_lay(path)
            else:
                self._lay = load_layout(path)
            self._current_file = path if path.endswith('.json') else None
            self._dirty = False
            self._undo_stack.clear()
            self._redo_stack.clear()
            self._canvas_panel.set_layout(self._lay)
            self._object_list.set_layout(self._lay)
            self._props.set_object(None)
            self._update_status()
        except Exception as e:
            mb.showerror('読み込みエラー', str(e), parent=self)

    def _on_save(self) -> None:
        if self._current_file:
            self._save_to(self._current_file)
        else:
            self._on_save_as()

    def _on_save_as(self) -> None:
        path = fd.asksaveasfilename(
            title='名前を付けて保存',
            defaultextension='.json',
            filetypes=_SAVE_FILE_TYPES,
        )
        if path:
            self._save_to(path)

    def _save_to(self, path: str) -> None:
        try:
            save_layout(self._lay, path)
            self._current_file = path
            self._dirty = False
            self._update_status()
        except Exception as e:
            mb.showerror('保存エラー', str(e), parent=self)

    def _confirm_discard(self) -> bool:
        return mb.askyesno(
            '未保存の変更',
            '変更が保存されていません。破棄しますか？',
            parent=self,
        )

    # ── 編集操作 ─────────────────────────────────────────────────────────

    def _on_undo(self) -> None:
        if not self._undo_stack:
            return
        idx, before, after = self._undo_stack.pop()
        self._redo_stack.append((idx, before, after))
        if idx < len(self._lay.objects):
            self._lay.objects[idx] = before
        self._dirty = True
        self._canvas_panel.refresh()
        self._props.set_object(None)
        self._update_status()

    def _on_redo(self) -> None:
        if not self._redo_stack:
            return
        idx, before, after = self._redo_stack.pop()
        self._undo_stack.append((idx, before, after))
        if idx < len(self._lay.objects):
            self._lay.objects[idx] = after
        self._dirty = True
        self._canvas_panel.refresh()
        self._props.set_object(None)
        self._update_status()

    def _push_undo(self, index: int, before: LayoutObject) -> None:
        """Undo スタックに変更前のオブジェクトを保存する。"""
        after = dataclasses.replace(self._lay.objects[index])
        self._undo_stack.append((index, before, after))
        self._redo_stack.clear()

    # ── オブジェクト追加 ─────────────────────────────────────────────────

    def _on_add_label(self) -> None:
        self._add_mode = 'label'
        self._update_status()

    def _on_add_field(self) -> None:
        self._add_mode = 'field'
        self._update_status()

    def _on_add_line(self) -> None:
        self._add_mode = 'line'
        self._update_status()

    def _cancel_add_mode(self) -> None:
        self._add_mode = None
        self._update_status()

    def _on_canvas_click(self, event: tk.Event) -> None:
        """Canvas クリック時（追加モード処理）。"""
        if self._add_mode is None:
            return

        canvas = self._canvas_panel.canvas
        cx = canvas.canvasx(event.x)
        cy = canvas.canvasy(event.y)
        scale = self._canvas_panel.get_zoom()
        from core.lay_renderer import canvas_to_model
        mx, my = canvas_to_model(cx, cy, scale, 30, 30)

        if self._add_mode == 'label':
            obj = new_label(mx, my, mx + 100, my + 30, text='新しいラベル')
            self._lay.objects.append(obj)
        elif self._add_mode == 'field':
            obj = new_field(mx, my, mx + 100, my + 30, field_id=108)
            self._lay.objects.append(obj)
        elif self._add_mode == 'line':
            obj = new_line(mx, my, mx + 100, my)
            self._lay.objects.append(obj)

        self._add_mode = None
        self._dirty = True
        self._canvas_panel.set_layout(self._lay)
        self._object_list.set_layout(self._lay)
        new_idx = len(self._lay.objects) - 1
        self._canvas_panel.select(new_idx)
        self._on_object_selected(new_idx)
        self._update_status()

    # ── オブジェクト削除 ─────────────────────────────────────────────────

    def _on_delete(self) -> None:
        idx = self._canvas_panel.get_selected_index()
        if idx < 0 or idx >= len(self._lay.objects):
            return
        self._lay.objects.pop(idx)
        self._dirty = True
        self._canvas_panel.set_layout(self._lay)
        self._object_list.set_layout(self._lay)
        self._props.set_object(None)
        self._update_status()

    # ── データ差込 ─────────────────────────────────────────────────────

    def _on_data_fill(self) -> None:
        """データ差込ダイアログを開く。"""
        from gui.editor.data_fill_dialog import DataFillDialog
        DataFillDialog(
            self, self._lay,
            on_preview=self._on_fill_preview,
            on_print=self._on_fill_print,
        )

    def _on_fill_preview(self, filled_lay: LayFile) -> None:
        """差込プレビューを Canvas に表示する。"""
        self._canvas_panel.set_layout(filled_lay)

    def _on_fill_print(self, filled_layouts: list[LayFile]) -> None:
        """差込済みレイアウトを印刷する。"""
        from gui.editor.print_dialog import PrintDialog
        PrintDialog(self, filled_layouts)

    # ── ズーム ───────────────────────────────────────────────────────────

    def _on_zoom_in(self) -> None:
        self._canvas_panel.zoom_in()
        pct = int(self._canvas_panel.get_zoom() / 0.5 * 100)
        self._toolbar.set_zoom_label(pct)

    def _on_zoom_out(self) -> None:
        self._canvas_panel.zoom_out()
        pct = int(self._canvas_panel.get_zoom() / 0.5 * 100)
        self._toolbar.set_zoom_label(pct)

    # ── ステータス ───────────────────────────────────────────────────────

    def _update_status(
        self, cursor: tuple[int, int, float, float] | None = None,
    ) -> None:
        parts = []

        # ファイル名
        if self._current_file:
            import os
            parts.append(os.path.basename(self._current_file))
        else:
            parts.append('新規')
        if self._dirty:
            parts.append('(変更あり)')

        # オブジェクト数
        parts.append(f'オブジェクト: {len(self._lay.objects)}')

        # 選択中
        idx = self._canvas_panel.get_selected_index()
        if 0 <= idx < len(self._lay.objects):
            obj = self._lay.objects[idx]
            type_name = {
                ObjectType.LABEL: 'ラベル',
                ObjectType.FIELD: 'フィールド',
                ObjectType.LINE: '罫線',
            }.get(obj.obj_type, '?')
            parts.append(f'選択: #{idx} {type_name}')

        # カーソル位置
        if cursor:
            mx, my, mm_x, mm_y = cursor
            parts.append(f'({mx}, {my}) = ({mm_x:.1f}mm, {mm_y:.1f}mm)')

        # 追加モード
        if self._add_mode:
            mode_name = {'label': 'ラベル', 'field': 'フィールド', 'line': '罫線'}
            parts.append(f'[{mode_name.get(self._add_mode, "?")}追加モード - クリックで配置]')

        # ズーム
        pct = int(self._canvas_panel.get_zoom() / 0.5 * 100)
        parts.append(f'ズーム: {pct}%')

        self._status.configure(text='  |  '.join(parts))
