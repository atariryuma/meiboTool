"""名簿ファイル読込パネル（セクション①）

SPEC.md §3.2 参照。
"""

from __future__ import annotations

import os
import tkinter.filedialog as fd
from collections.abc import Callable

import customtkinter as ctk

from core.importer import import_c4th_excel


class ImportFrame(ctk.CTkFrame):
    """名簿ファイル選択・読み込みセクション。"""

    def __init__(self, master, on_load: Callable) -> None:
        """
        on_load(df_mapped, unmapped, source_path) を読込完了時に呼ぶ。
        """
        super().__init__(master, corner_radius=6)
        self.on_load = on_load
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self,
            text='① 名簿ファイル読込',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=0, column=0, sticky='w', padx=10, pady=(8, 4))

        self._btn = ctk.CTkButton(
            self, text='📂 別のファイルを選択…', command=self._pick_file
        )
        self._btn.grid(row=1, column=0, padx=10, pady=4, sticky='ew')

        self._path_label = ctk.CTkLabel(
            self, text='（未選択）', text_color='gray',
            wraplength=270, anchor='w',
        )
        self._path_label.grid(row=2, column=0, padx=10, pady=2, sticky='w')

        self._count_label = ctk.CTkLabel(self, text='')
        self._count_label.grid(row=3, column=0, padx=10, pady=(2, 4), sticky='w')

        self._sync_label = ctk.CTkLabel(
            self, text='', font=ctk.CTkFont(size=11),
        )
        self._sync_label.grid(row=4, column=0, padx=10, pady=(0, 8), sticky='w')

    # ── 外部 API ─────────────────────────────────────────────────────────

    def load_from_path(self, path: str) -> None:
        """指定パスから直接ファイルを読み込む（起動時の自動再読込用）。"""
        self._load(path)

    def show_sync_status(self, message: str, warning: bool = False) -> None:
        """同期ステータスメッセージを表示する。"""
        color = '#B45309' if warning else '#059669'
        self._sync_label.configure(text=message, text_color=color)

    # ── 内部 ─────────────────────────────────────────────────────────────

    def _pick_file(self) -> None:
        path = fd.askopenfilename(
            title='名簿 Excel を選択',
            filetypes=[('Excel ファイル', '*.xlsx *.xls'), ('すべて', '*.*')],
        )
        if path:
            self._load(path)

    def _load(self, path: str) -> None:
        try:
            df_mapped, unmapped = import_c4th_excel(path)
            n = len(df_mapped)
            self._path_label.configure(text=os.path.basename(path), text_color='black')
            self._count_label.configure(
                text=f'✅ {n} 名 読み込み完了', text_color='green'
            )
            self.on_load(df_mapped, unmapped, path)
        except Exception as e:
            msg = str(e)
            self._path_label.configure(text=msg[:80], text_color='red')
            self._count_label.configure(text='❌ 読込エラー', text_color='red')
            if len(msg) > 80:
                import tkinter.messagebox as _mb
                _mb.showerror('読込エラー', msg)
