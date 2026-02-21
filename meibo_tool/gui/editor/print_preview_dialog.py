"""印刷プレビューダイアログ

差込済みレイアウトのページ送り表示 + 印刷ボタン。
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

import customtkinter as ctk
from PIL import Image as PILImage

from core.lay_renderer import render_layout_to_image

if TYPE_CHECKING:
    from core.lay_parser import LayFile


class PrintPreviewDialog(ctk.CTkToplevel):
    """印刷プレビュー: ページ送り + 印刷ボタン。"""

    _MAX_DISPLAY_WIDTH = 600

    def __init__(
        self, master: ctk.CTkBaseClass,
        layouts: list[LayFile],
        on_print: Callable[[list[LayFile]], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.title('印刷プレビュー')
        self.geometry('680x920')
        self.transient(master)

        self._layouts = layouts
        self._on_print = on_print
        self._current_page = 0
        self._preview_images: list[PILImage.Image] = []
        self._tk_image: ctk.CTkImage | None = None

        self._build_ui()
        self._render_all_pages()
        if self._preview_images:
            self._show_page(0)

    # ── UI 構築 ───────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ナビゲーションバー
        nav = ctk.CTkFrame(self, fg_color='transparent')
        nav.grid(row=0, column=0, padx=10, pady=(10, 5), sticky='ew')
        nav.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            nav, text='<', width=40, command=self._on_prev,
        ).grid(row=0, column=0, padx=2)

        self._page_label = ctk.CTkLabel(
            nav, text='0 / 0', font=ctk.CTkFont(size=14),
        )
        self._page_label.grid(row=0, column=1, padx=5)

        ctk.CTkButton(
            nav, text='>', width=40, command=self._on_next,
        ).grid(row=0, column=2, padx=2)

        # プレビュー画像
        self._img_label = ctk.CTkLabel(self, text='レンダリング中…')
        self._img_label.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')

        # ボタン
        btn_frame = ctk.CTkFrame(self, fg_color='transparent')
        btn_frame.grid(row=2, column=0, padx=10, pady=(5, 10), sticky='e')

        ctk.CTkButton(
            btn_frame, text='印刷', width=100, command=self._on_print_click,
        ).grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkButton(
            btn_frame, text='閉じる', width=80, command=self.destroy,
        ).grid(row=0, column=1, padx=5, pady=5)

    # ── ページレンダリング ─────────────────────────────────────────────

    def _render_all_pages(self) -> None:
        """全ページを PIL 画像にレンダリングする。"""
        for lay in self._layouts:
            img = render_layout_to_image(lay, dpi=150)
            self._preview_images.append(img)

    def _show_page(self, idx: int) -> None:
        """指定ページを表示する。"""
        total = len(self._preview_images)
        if idx < 0 or idx >= total:
            return
        self._current_page = idx
        self._page_label.configure(text=f'{idx + 1} / {total}')

        img = self._preview_images[idx]
        # ダイアログ幅に合わせて縮小
        ratio = self._MAX_DISPLAY_WIDTH / img.width
        display_w = int(img.width * ratio)
        display_h = int(img.height * ratio)
        display_img = img.resize((display_w, display_h), PILImage.LANCZOS)

        self._tk_image = ctk.CTkImage(
            light_image=display_img, size=(display_w, display_h),
        )
        self._img_label.configure(image=self._tk_image, text='')

    # ── ナビゲーション ─────────────────────────────────────────────────

    def _on_prev(self) -> None:
        if self._current_page > 0:
            self._show_page(self._current_page - 1)

    def _on_next(self) -> None:
        if self._current_page < len(self._preview_images) - 1:
            self._show_page(self._current_page + 1)

    def _on_print_click(self) -> None:
        if self._on_print:
            self._on_print(self._layouts)
            self.destroy()

    def destroy(self) -> None:
        """ダイアログ破棄時に PIL 画像を解放する。"""
        for img in self._preview_images:
            img.close()
        self._preview_images.clear()
        self._tk_image = None
        super().destroy()
