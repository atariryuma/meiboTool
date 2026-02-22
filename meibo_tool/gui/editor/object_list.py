"""オブジェクト一覧パネル

LayFile 内のオブジェクトを Treeview で一覧表示し、
クリックで Canvas の選択と連動する。
"""

from __future__ import annotations

import tkinter as tk
import tkinter.ttk as ttk
from collections.abc import Callable

import customtkinter as ctk

from core.lay_parser import (
    LayFile,
    LayoutObject,
    ObjectType,
    resolve_field_display,
)

# ── 定数 ─────────────────────────────────────────────────────────────────────

_TYPE_NAMES = {
    ObjectType.LABEL: 'ラベル',
    ObjectType.FIELD: 'フィールド',
    ObjectType.LINE: '罫線',
    ObjectType.GROUP: 'グループ',
}


class ObjectListPanel(ctk.CTkFrame):
    """左サイドのオブジェクト一覧パネル。"""

    def __init__(
        self, master: ctk.CTkBaseClass,
        on_select: Callable[[int], None] | None = None,
    ) -> None:
        super().__init__(master, width=220)
        self._on_select = on_select
        self._lay: LayFile | None = None
        self._updating = False
        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── タイトル ──
        ctk.CTkLabel(
            self, text='オブジェクト一覧',
            font=ctk.CTkFont(size=14, weight='bold'),
        ).grid(row=0, column=0, padx=8, pady=(8, 4), sticky='w')

        # ── Treeview ──
        tree_frame = ctk.CTkFrame(self, fg_color='transparent')
        tree_frame.grid(row=1, column=0, sticky='nsew', padx=4, pady=4)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = ('type', 'text', 'pos')
        self._tree = ttk.Treeview(
            tree_frame, columns=columns, show='headings',
            selectmode='browse', height=20,
        )

        self._tree.heading('type', text='種類')
        self._tree.heading('text', text='内容')
        self._tree.heading('pos', text='位置')

        self._tree.column('type', width=60, minwidth=50)
        self._tree.column('text', width=90, minwidth=60)
        self._tree.column('pos', width=60, minwidth=50)

        self._tree.grid(row=0, column=0, sticky='nsew')

        # スクロールバー
        scrollbar = ttk.Scrollbar(
            tree_frame, orient='vertical', command=self._tree.yview,
        )
        scrollbar.grid(row=0, column=1, sticky='ns')
        self._tree.configure(yscrollcommand=scrollbar.set)

        # 選択イベント
        self._tree.bind('<<TreeviewSelect>>', self._on_tree_select)

        # ── オブジェクト数 ──
        self._count_label = ctk.CTkLabel(self, text='0 件')
        self._count_label.grid(row=2, column=0, padx=8, pady=(0, 8), sticky='w')

    # ── 公開メソッド ─────────────────────────────────────────────────────

    def set_layout(self, lay: LayFile) -> None:
        """レイアウトを設定して一覧を更新する。"""
        self._lay = lay
        self._refresh_list()

    def refresh(self) -> None:
        """一覧を再描画する。"""
        self._refresh_list()

    def select(self, index: int) -> None:
        """指定インデックスのオブジェクトを選択する。"""
        children = self._tree.get_children()
        if 0 <= index < len(children):
            self._updating = True
            self._tree.selection_set(children[index])
            self._tree.see(children[index])
            self._updating = False

    # ── 内部メソッド ─────────────────────────────────────────────────────

    def _refresh_list(self) -> None:
        """Treeview を再構築する。"""
        self._updating = True

        # 既存行を削除
        for item in self._tree.get_children():
            self._tree.delete(item)

        if self._lay is None:
            self._count_label.configure(text='0 件')
            self._updating = False
            return

        for i, obj in enumerate(self._lay.objects):
            type_name = _TYPE_NAMES.get(obj.obj_type, '?')
            text = self._get_display_text(obj)
            pos = self._get_position_text(obj)
            self._tree.insert('', 'end', iid=str(i), values=(type_name, text, pos))

        self._count_label.configure(text=f'{len(self._lay.objects)} 件')
        self._updating = False

    @staticmethod
    def _get_display_text(obj: LayoutObject) -> str:
        """オブジェクトの表示用テキストを取得する。"""
        if obj.obj_type == ObjectType.LABEL:
            text = obj.text
            return text[:12] + '…' if len(text) > 12 else text
        if obj.obj_type == ObjectType.FIELD:
            return resolve_field_display(obj.field_id)
        if obj.obj_type == ObjectType.LINE:
            return '—'
        return ''

    @staticmethod
    def _get_position_text(obj: LayoutObject) -> str:
        """位置を簡潔に表示する。"""
        if obj.rect:
            return f'{obj.rect.left},{obj.rect.top}'
        if obj.line_start:
            return f'{obj.line_start.x},{obj.line_start.y}'
        return ''

    def _on_tree_select(self, _event: tk.Event) -> None:
        """Treeview 選択時のコールバック。"""
        if self._updating:
            return
        selection = self._tree.selection()
        if not selection:
            return
        try:
            index = int(selection[0])
        except ValueError:
            return
        if self._on_select:
            self._on_select(index)
