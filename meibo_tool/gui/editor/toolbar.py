"""エディターツールバー

ファイル操作・編集ツール・ズームコントロールを提供する。
"""

from __future__ import annotations

from collections.abc import Callable

import customtkinter as ctk


class EditorToolbar(ctk.CTkFrame):
    """エディター上部のツールバー。"""

    def __init__(
        self, master: ctk.CTkBaseClass,
        callbacks: dict[str, Callable],
    ) -> None:
        super().__init__(master, height=40, corner_radius=0)
        self._callbacks = callbacks
        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(20, weight=1)

        col = 0

        # ── ファイル操作 ──
        for label, key in [
            ('新規', 'new'), ('開く', 'open'),
            ('保存', 'save'), ('名前を付けて保存', 'save_as'),
        ]:
            btn = ctk.CTkButton(
                self, text=label, width=80 if len(label) <= 2 else 120,
                height=28, command=self._callbacks.get(key, lambda: None),
            )
            btn.grid(row=0, column=col, padx=2, pady=4)
            col += 1

        # ── セパレーター ──
        sep = ctk.CTkFrame(self, width=2, height=24, fg_color='gray60')
        sep.grid(row=0, column=col, padx=6, pady=6)
        col += 1

        # ── 編集ツール ──
        for label, key in [
            ('元に戻す', 'undo'), ('やり直し', 'redo'),
        ]:
            btn = ctk.CTkButton(
                self, text=label, width=80, height=28,
                command=self._callbacks.get(key, lambda: None),
            )
            btn.grid(row=0, column=col, padx=2, pady=4)
            col += 1

        # ── セパレーター ──
        sep2 = ctk.CTkFrame(self, width=2, height=24, fg_color='gray60')
        sep2.grid(row=0, column=col, padx=6, pady=6)
        col += 1

        # ── オブジェクト追加 ──
        for label, key in [
            ('+ ラベル', 'add_label'), ('+ フィールド', 'add_field'),
            ('+ 罫線', 'add_line'), ('削除', 'delete'),
        ]:
            btn = ctk.CTkButton(
                self, text=label, width=90, height=28,
                command=self._callbacks.get(key, lambda: None),
            )
            btn.grid(row=0, column=col, padx=2, pady=4)
            col += 1

        # ── セパレーター ──
        sep3 = ctk.CTkFrame(self, width=2, height=24, fg_color='gray60')
        sep3.grid(row=0, column=col, padx=6, pady=6)
        col += 1

        # ── スペーサー ──
        col = 20

        # ── ズーム ──
        self._zoom_label = ctk.CTkLabel(self, text='100%', width=50)
        self._zoom_label.grid(row=0, column=col + 1, padx=2, pady=4)

        ctk.CTkButton(
            self, text='-', width=30, height=28,
            command=self._callbacks.get('zoom_out', lambda: None),
        ).grid(row=0, column=col, padx=2, pady=4)

        ctk.CTkButton(
            self, text='+', width=30, height=28,
            command=self._callbacks.get('zoom_in', lambda: None),
        ).grid(row=0, column=col + 2, padx=2, pady=4)

    def set_zoom_label(self, pct: int) -> None:
        """ズーム率表示を更新する。"""
        self._zoom_label.configure(text=f'{pct}%')
