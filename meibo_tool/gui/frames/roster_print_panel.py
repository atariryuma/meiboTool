"""名簿印刷パネル

レイアウト一覧から選択し、名簿データを差し込んで印刷する。
"""

from __future__ import annotations

import contextlib
import logging
import os
import threading
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from typing import Any

import customtkinter as ctk
import pandas as pd
from PIL import Image as PILImage

from core.config import get_layout_dir
from core.importer import import_c4th_excel
from core.lay_parser import LayFile
from core.lay_renderer import fill_layout, render_layout_to_image
from core.lay_serializer import load_layout
from core.layout_registry import scan_layout_dir
from core.special_needs import (
    detect_regular_students,
    detect_special_needs_students,
    merge_special_needs_students,
)

logger = logging.getLogger(__name__)


class RosterPrintPanel(ctk.CTkFrame):
    """名簿印刷タブのメインパネル。"""

    def __init__(
        self, master: ctk.CTkBaseClass,
        config: dict[str, Any],
    ) -> None:
        super().__init__(master, corner_radius=0, fg_color='transparent')
        self._config = config
        self._layout_dir = get_layout_dir(config)
        self._layouts: list[dict[str, Any]] = []
        self._selected_lay: LayFile | None = None
        self._selected_path: str | None = None
        self._df: pd.DataFrame | None = None
        self._render_generation = 0

        self._build_ui()
        self._refresh_layouts()

    # ── UI 構築 ───────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=0, minsize=320)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── 左パネル ──────────────────────────────────────────────────
        left = ctk.CTkFrame(self, width=310, corner_radius=6)
        left.grid(row=0, column=0, sticky='nsew', padx=(5, 2), pady=5)
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)  # Treeview が伸びる

        # レイアウト選択ヘッダー
        ctk.CTkLabel(
            left, text='レイアウト選択',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=0, column=0, sticky='w', padx=8, pady=(8, 4))

        # Treeview
        tree_frame = ctk.CTkFrame(left, fg_color='transparent')
        tree_frame.grid(row=1, column=0, sticky='nsew', padx=5, pady=(0, 5))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = ('title', 'page_size', 'fields')
        self._tree = ttk.Treeview(
            tree_frame, columns=columns, show='headings',
            selectmode='browse', height=6,
        )
        self._tree.heading('title', text='タイトル')
        self._tree.heading('page_size', text='用紙')
        self._tree.heading('fields', text='フィールド')
        self._tree.column('title', width=160)
        self._tree.column('page_size', width=60)
        self._tree.column('fields', width=60)
        self._tree.grid(row=0, column=0, sticky='nsew')
        self._tree.bind('<<TreeviewSelect>>', self._on_layout_select)

        scrollbar = ttk.Scrollbar(
            tree_frame, orient='vertical', command=self._tree.yview,
        )
        scrollbar.grid(row=0, column=1, sticky='ns')
        self._tree.configure(yscrollcommand=scrollbar.set)

        # 名簿データセクション
        data_frame = ctk.CTkFrame(left, fg_color='transparent')
        data_frame.grid(row=2, column=0, sticky='ew', padx=5, pady=5)
        data_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            data_frame, text='名簿データ',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=0, column=0, columnspan=2, sticky='w', padx=3, pady=(0, 4))

        ctk.CTkButton(
            data_frame, text='Excel 読込', width=100,
            command=self._on_load_excel,
        ).grid(row=1, column=0, padx=3, pady=2)

        self._file_label = ctk.CTkLabel(
            data_frame, text='ファイル未選択', text_color='gray',
        )
        self._file_label.grid(row=1, column=1, padx=5, sticky='w')

        self._count_label = ctk.CTkLabel(data_frame, text='')
        self._count_label.grid(
            row=2, column=0, columnspan=2, padx=3, pady=(2, 0), sticky='w',
        )

        # オプションセクション
        opt_frame = ctk.CTkFrame(left, fg_color='transparent')
        opt_frame.grid(row=3, column=0, sticky='ew', padx=5, pady=5)
        opt_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            opt_frame, text='オプション',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=0, column=0, columnspan=2, sticky='w', padx=3, pady=(0, 4))

        self._school_var = ctk.StringVar(
            value=self._config.get('school_name', ''),
        )
        self._year_var = ctk.StringVar(
            value=str(self._config.get('fiscal_year', 2025)),
        )
        self._teacher_var = ctk.StringVar()

        for r, (label, var) in enumerate([
            ('学校名', self._school_var),
            ('年度', self._year_var),
            ('担任名', self._teacher_var),
        ], start=1):
            ctk.CTkLabel(opt_frame, text=label).grid(
                row=r, column=0, padx=3, pady=2, sticky='w',
            )
            ctk.CTkEntry(opt_frame, textvariable=var, width=180).grid(
                row=r, column=1, padx=3, pady=2, sticky='ew',
            )

        # 特別支援学級配置セクション
        placement_frame = ctk.CTkFrame(left, fg_color='transparent')
        placement_frame.grid(row=4, column=0, sticky='ew', padx=5, pady=5)

        ctk.CTkLabel(
            placement_frame, text='特別支援学級',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=0, column=0, columnspan=2, sticky='w', padx=3, pady=(0, 4))

        self._placement_var = ctk.StringVar(
            value=self._config.get('special_needs_placement', 'appended'),
        )
        ctk.CTkRadioButton(
            placement_frame, text='末尾に追加',
            variable=self._placement_var, value='appended',
        ).grid(row=1, column=0, padx=(10, 5), pady=2, sticky='w')
        ctk.CTkRadioButton(
            placement_frame, text='出席番号順に統合',
            variable=self._placement_var, value='integrated',
        ).grid(row=2, column=0, padx=(10, 5), pady=2, sticky='w')

        # ボタンセクション
        btn_frame = ctk.CTkFrame(left, fg_color='transparent')
        btn_frame.grid(row=5, column=0, sticky='ew', padx=5, pady=(5, 8))

        ctk.CTkButton(
            btn_frame, text='プレビュー', width=120,
            command=self._on_preview,
        ).grid(row=0, column=0, padx=5, pady=5)

        ctk.CTkButton(
            btn_frame, text='印刷', width=120,
            command=self._on_print,
        ).grid(row=0, column=1, padx=5, pady=5)

        # ── 右パネル（プレビュー） ────────────────────────────────────────
        right = ctk.CTkFrame(self, corner_radius=6)
        right.grid(row=0, column=1, sticky='nsew', padx=(2, 5), pady=5)
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        self._preview_scroll = ctk.CTkScrollableFrame(
            right, fg_color='transparent',
        )
        self._preview_scroll.grid(row=0, column=0, sticky='nsew')
        self._preview_scroll.grid_columnconfigure(0, weight=1)

        self._preview_label = ctk.CTkLabel(
            self._preview_scroll,
            text='レイアウトを選択すると\nプレビューが表示されます',
            text_color='gray',
        )
        self._preview_label.grid(row=0, column=0, padx=20, pady=40)

        # CTkImage の GC 防止用リファレンス
        self._preview_image_ref: ctk.CTkImage | None = None

    # ── レイアウト一覧 ─────────────────────────────────────────────────

    def refresh_layouts(self) -> None:
        """レイアウト一覧を再読み込みする（外部から呼び出し可能）。"""
        self._refresh_layouts()

    def _refresh_layouts(self) -> None:
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
                    meta.get('field_count', 0),
                ),
            )

    def _on_layout_select(self, _event: tk.Event) -> None:
        """Treeview でレイアウト選択時。"""
        sel = self._tree.selection()
        if not sel:
            return

        idx = int(sel[0])
        if idx >= len(self._layouts):
            return

        meta = self._layouts[idx]
        path = meta['path']

        try:
            if path.lower().endswith('.lay'):
                from core.lay_parser import parse_lay
                self._selected_lay = parse_lay(path)
            else:
                self._selected_lay = load_layout(path)
            self._selected_path = path
            self._render_preview()
        except Exception as e:
            mb.showerror('読み込みエラー', str(e))
            self._selected_lay = None

    # ── プレビュー ─────────────────────────────────────────────────────

    def _render_preview(self) -> None:
        """選択中のレイアウトをプレビュー表示する。"""
        if self._selected_lay is None:
            return

        lay = self._selected_lay

        # データがある場合は1人目を差し込んでプレビュー
        if self._df is not None and not self._df.empty:
            try:
                row = self._df.iloc[0].to_dict()
                lay = fill_layout(self._selected_lay, row, self._get_options())
            except Exception:
                pass  # 差込失敗時はテンプレートのまま表示

        # バックグラウンドでレンダリング（世代カウンターでキャンセル管理）
        self._render_generation += 1
        gen = self._render_generation
        threading.Thread(
            target=self._render_worker,
            args=(lay, gen),
            daemon=True,
        ).start()

    def _render_worker(self, lay: LayFile, generation: int) -> None:
        """バックグラウンドでレイアウトをレンダリングする。"""
        try:
            img = render_layout_to_image(lay, dpi=150)
            if generation == self._render_generation and self.winfo_exists():
                self.after(0, lambda: self._show_preview_image(img))
        except Exception:
            logger.exception('レイアウトプレビュー生成エラー')
            if generation == self._render_generation and self.winfo_exists():
                self.after(0, lambda: self._show_preview_error(
                    'プレビューの生成中にエラーが発生しました',
                ))

    def _show_preview_image(self, pil_image: PILImage.Image) -> None:
        """PIL Image をプレビューに表示する。"""
        try:
            w, h = pil_image.size
            max_w = 500
            if w > max_w:
                ratio = max_w / w
                w = max_w
                h = int(h * ratio)

            ctk_img = ctk.CTkImage(
                light_image=pil_image,
                size=(w, h),
            )
            self._preview_image_ref = ctk_img
            self._preview_label.configure(text='', image=ctk_img)
            self._preview_label.grid(row=0, column=0, padx=4, pady=4)
        except tk.TclError:
            pass

    def _show_preview_error(self, msg: str) -> None:
        """プレビューエラーメッセージを表示する。"""
        self._preview_image_ref = None
        with contextlib.suppress(tk.TclError):
            self._preview_label.configure(
                text=msg, text_color='#CC4444', image=None,
            )

    # ── データ読込 ─────────────────────────────────────────────────────

    def _on_load_excel(self) -> None:
        """C4th Excel ファイルを読み込む。"""
        path = fd.askopenfilename(
            title='C4th Excel を選択',
            filetypes=[('Excel ファイル', '*.xlsx *.xls'), ('すべて', '*.*')],
        )
        if not path:
            return

        try:
            df, _unmapped = import_c4th_excel(path)
            self._df = df

            self._file_label.configure(
                text=os.path.basename(path),
                text_color=ctk.ThemeManager.theme['CTkLabel']['text_color'],
            )

            # 特支検出して件数表示
            n_regular = len(detect_regular_students(df))
            n_special = len(detect_special_needs_students(df))
            total = len(df)
            if n_special > 0:
                self._count_label.configure(
                    text=f'印刷対象: {total} 名'
                    f'（通常 {n_regular} + 特支 {n_special}）',
                )
            else:
                self._count_label.configure(text=f'印刷対象: {total} 名')

            # プレビュー更新（データ差込あり）
            self._render_preview()

        except Exception as e:
            mb.showerror('読み込みエラー', str(e))

    # ── オプション ─────────────────────────────────────────────────────

    def _get_options(self) -> dict[str, Any]:
        """差込オプションを返す。"""
        opts: dict[str, Any] = {
            'school_name': self._school_var.get(),
            'teacher_name': self._teacher_var.get(),
            'name_display': 'furigana',
        }
        with contextlib.suppress(ValueError):
            opts['fiscal_year'] = int(self._year_var.get())
        return opts

    # ── 差込・印刷 ─────────────────────────────────────────────────────

    def _build_filled_layouts(self) -> list[LayFile]:
        """全生徒のデータを差し込んだレイアウト一覧を生成する。"""
        if self._selected_lay is None or self._df is None:
            return []

        regular = detect_regular_students(self._df)
        special = detect_special_needs_students(self._df)

        placement = self._placement_var.get()
        merged = merge_special_needs_students(regular, special, placement)

        opts = self._get_options()
        filled: list[LayFile] = []
        errors: list[str] = []

        for i, (_idx, row) in enumerate(merged.iterrows()):
            try:
                filled.append(
                    fill_layout(self._selected_lay, row.to_dict(), opts),
                )
            except Exception as e:
                errors.append(f'行 {i + 1}: {e}')

        if errors:
            msg = '\n'.join(errors[:10])
            if len(errors) > 10:
                msg += f'\n... 他 {len(errors) - 10} 件'
            mb.showwarning('差込エラー', f'一部の行でエラー:\n{msg}')

        return filled

    def _on_preview(self) -> None:
        """プレビューボタン押下時。"""
        if self._selected_lay is None:
            mb.showinfo('レイアウト未選択', 'レイアウトを選択してください。')
            return
        if self._df is None:
            mb.showinfo('データ未読込', 'Excel ファイルを先に読み込んでください。')
            return

        filled = self._build_filled_layouts()
        if not filled:
            mb.showerror('エラー', '差込可能なデータがありません。')
            return

        from gui.editor.print_preview_dialog import PrintPreviewDialog
        PrintPreviewDialog(
            self.winfo_toplevel(), filled,
            on_print=self._do_print,
        )

    def _on_print(self) -> None:
        """印刷ボタン押下時。"""
        if self._selected_lay is None:
            mb.showinfo('レイアウト未選択', 'レイアウトを選択してください。')
            return
        if self._df is None:
            mb.showinfo('データ未読込', 'Excel ファイルを先に読み込んでください。')
            return

        filled = self._build_filled_layouts()
        if not filled:
            mb.showerror('エラー', '差込可能なデータがありません。')
            return

        from gui.editor.print_dialog import PrintDialog
        PrintDialog(self.winfo_toplevel(), filled)

    def _do_print(self, layouts: list[LayFile]) -> None:
        """PrintPreviewDialog からの印刷コールバック。"""
        from gui.editor.print_dialog import PrintDialog
        PrintDialog(self.winfo_toplevel(), layouts)
