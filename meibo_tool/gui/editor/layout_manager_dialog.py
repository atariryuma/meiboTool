"""レイアウトライブラリ管理ダイアログ

保存済みレイアウトの一覧表示・インポート・削除・リネーム・エディター起動。
"""

from __future__ import annotations

import contextlib
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from collections.abc import Callable
from typing import Any

import customtkinter as ctk

from core.config import get_layout_dir
from core.layout_registry import (
    delete_layout,
    import_json_file,
    import_lay_file,
    import_lay_file_multi,
    rename_layout,
    scan_layout_dir,
)


class LayoutManagerDialog(ctk.CTkToplevel):
    """レイアウトライブラリ管理ダイアログ。"""

    def __init__(
        self, master: ctk.CTkBaseClass,
        config: dict[str, Any],
        on_open: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.title('レイアウトライブラリ')
        self.geometry('700x500')
        self.transient(master)

        self._config = config
        self._on_open = on_open
        self._layout_dir = get_layout_dir(config)
        self._layouts: list[dict[str, Any]] = []

        self._build_ui()
        self._refresh_list()

    # ── UI 構築 ───────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # タイトル
        ctk.CTkLabel(
            self, text='レイアウトライブラリ',
            font=ctk.CTkFont(size=16, weight='bold'),
        ).grid(row=0, column=0, padx=15, pady=(15, 5), sticky='w')

        # Treeview
        tree_frame = ctk.CTkFrame(self)
        tree_frame.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = ('title', 'page_size', 'objects', 'fields')
        self._tree = ttk.Treeview(
            tree_frame, columns=columns, show='headings',
            selectmode='browse',
        )
        self._tree.heading('title', text='タイトル')
        self._tree.heading('page_size', text='用紙サイズ')
        self._tree.heading('objects', text='オブジェクト数')
        self._tree.heading('fields', text='フィールド数')
        self._tree.column('title', width=250)
        self._tree.column('page_size', width=120)
        self._tree.column('objects', width=100)
        self._tree.column('fields', width=100)
        self._tree.grid(row=0, column=0, sticky='nsew')
        self._tree.bind('<Double-1>', lambda _e: self._on_open_click())

        scrollbar = ttk.Scrollbar(
            tree_frame, orient='vertical', command=self._tree.yview,
        )
        scrollbar.grid(row=0, column=1, sticky='ns')
        self._tree.configure(yscrollcommand=scrollbar.set)

        # ボタン
        btn_frame = ctk.CTkFrame(self, fg_color='transparent')
        btn_frame.grid(row=2, column=0, padx=10, pady=(5, 10), sticky='ew')

        for col, (label, cmd, width) in enumerate([
            ('インポート', self._on_import, 100),
            ('エディターで開く', self._on_open_click, 130),
            ('リネーム', self._on_rename, 80),
            ('削除', self._on_delete, 80),
            ('新規作成', self._on_new, 80),
            ('閉じる', self.destroy, 80),
        ]):
            ctk.CTkButton(
                btn_frame, text=label, width=width, command=cmd,
            ).grid(row=0, column=col, padx=5, pady=5)

    # ── リスト管理 ─────────────────────────────────────────────────────

    def _refresh_list(self) -> None:
        """レイアウト一覧を再読み込みする。"""
        for item in self._tree.get_children():
            self._tree.delete(item)

        self._layouts = scan_layout_dir(self._layout_dir)
        for i, meta in enumerate(self._layouts):
            self._tree.insert(
                '', 'end', iid=str(i),
                values=(
                    meta.get('title') or meta.get('name', ''),
                    meta.get('page_size_mm', ''),
                    meta.get('object_count', 0),
                    meta.get('field_count', 0),
                ),
            )

    def _get_selected_index(self) -> int | None:
        """選択中のレイアウトのインデックスを返す。"""
        sel = self._tree.selection()
        if not sel:
            return None
        return int(sel[0])

    # ── ボタン操作 ─────────────────────────────────────────────────────

    def _on_import(self) -> None:
        """レイアウトファイルをインポートする。

        .lay ファイルの場合、マルチレイアウトを検出して一括インポートを提案する。
        """
        path = fd.askopenfilename(
            parent=self,
            title='レイアウトファイルを選択',
            filetypes=[
                ('レイアウトファイル', '*.lay *.json'),
                ('.lay ファイル', '*.lay'),
                ('.json ファイル', '*.json'),
            ],
        )
        if not path:
            return

        try:
            if path.lower().endswith('.lay'):
                self._import_lay(path)
            else:
                import_json_file(path, self._layout_dir)
                self._refresh_list()
                mb.showinfo('完了', 'レイアウトをインポートしました。', parent=self)
        except Exception as e:
            mb.showerror('インポートエラー', str(e), parent=self)

    def _import_lay(self, path: str) -> None:
        """マルチレイアウト対応の .lay インポート。"""
        from core.lay_parser import parse_lay_multi

        layouts = parse_lay_multi(path)
        count = len(layouts)

        if count <= 1:
            # 単一レイアウト: 従来のインポート
            import_lay_file(path, self._layout_dir)
            self._refresh_list()
            mb.showinfo('完了', 'レイアウトをインポートしました。', parent=self)
        else:
            # マルチレイアウト: 確認ダイアログ
            if mb.askyesno(
                '一括インポート',
                f'このファイルには {count} 件のレイアウトが含まれています。\n'
                f'すべてインポートしますか？',
                parent=self,
            ):
                results = import_lay_file_multi(path, self._layout_dir)
                self._refresh_list()
                mb.showinfo(
                    '完了',
                    f'{len(results)} 件のレイアウトをインポートしました。',
                    parent=self,
                )

    def _on_open_click(self) -> None:
        """選択中のレイアウトをエディターで開く。"""
        idx = self._get_selected_index()
        if idx is None or idx >= len(self._layouts):
            mb.showinfo('選択なし', 'レイアウトを選択してください。', parent=self)
            return

        path = self._layouts[idx]['path']
        if self._on_open:
            self._on_open(path)
        self.destroy()

    def _on_rename(self) -> None:
        """選択中のレイアウトをリネームする。"""
        idx = self._get_selected_index()
        if idx is None or idx >= len(self._layouts):
            mb.showinfo('選択なし', 'レイアウトを選択してください。', parent=self)
            return

        meta = self._layouts[idx]
        dialog = ctk.CTkInputDialog(
            text='新しい名前を入力してください:',
            title='リネーム',
        )
        new_name = dialog.get_input()
        if not new_name:
            return

        try:
            rename_layout(meta['path'], new_name)
            self._refresh_list()
        except FileExistsError as e:
            mb.showerror('リネームエラー', str(e), parent=self)

    def _on_delete(self) -> None:
        """選択中のレイアウトを削除する。"""
        idx = self._get_selected_index()
        if idx is None or idx >= len(self._layouts):
            mb.showinfo('選択なし', 'レイアウトを選択してください。', parent=self)
            return

        meta = self._layouts[idx]
        name = meta.get('title') or meta.get('name', '')
        if not mb.askyesno(
            '削除確認', f'「{name}」を削除しますか？', parent=self,
        ):
            return

        delete_layout(meta['path'])
        self._refresh_list()

    def _on_new(self) -> None:
        """新規レイアウトでエディターを起動する。"""
        if self._on_open:
            self._on_open('')  # 空パス = 新規
        self.destroy()

    def focus_set(self) -> None:
        """CTkToplevel の after(150, focus_set) による TclError を防止する。

        tkinter.Misc で ``focus = focus_set`` というエイリアスが定義されているため、
        focus_set だけオーバーライドしても CTkToplevel が ``.focus`` 経由で呼ぶと
        元の Misc.focus_set が使われてしまう。focus も同時にオーバーライドする。
        """
        if self.winfo_exists():
            super().focus_set()

    focus = focus_set

    def destroy(self) -> None:
        """Treeview のイベントバインドを解除してから破棄する。"""
        with contextlib.suppress(tk.TclError, AttributeError):
            self._tree.unbind('<Double-1>')
        # 破棄後に focus が呼ばれても TclError にならないようにする
        with contextlib.suppress(AttributeError):
            self._tree.focus = lambda *_a, **_kw: ''
        super().destroy()
