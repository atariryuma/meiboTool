"""データ差込ダイアログ

C4th Excel ファイルを読み込み、学生データをレイアウトに差し込んで
プレビュー表示する。印刷にも使用する。
"""

from __future__ import annotations

import contextlib
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from collections.abc import Callable

import customtkinter as ctk
import pandas as pd

from core.importer import import_c4th_excel
from core.lay_parser import LayFile
from core.lay_renderer import fill_layout


class DataFillDialog(ctk.CTkToplevel):
    """C4th データ差込ダイアログ。"""

    def __init__(
        self, master: ctk.CTkBaseClass,
        lay: LayFile,
        on_preview: Callable[[LayFile], None] | None = None,
        on_print: Callable[[list[LayFile]], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.title('データ差込')
        self.geometry('500x500')
        self.transient(master)

        self._lay = lay
        self._on_preview = on_preview
        self._on_print = on_print
        self._df: pd.DataFrame | None = None
        self._options: dict = {
            'fiscal_year': 2025,
            'school_name': '',
            'teacher_name': '',
            'name_display': 'furigana',
        }
        self._current_row_idx = 0

        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        row = 0

        # ── ファイル選択 ──
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=row, column=0, padx=10, pady=(10, 5), sticky='ew')
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            file_frame, text='Excel 読込', width=100,
            command=self._on_load,
        ).grid(row=0, column=0, padx=5, pady=5)

        self._file_label = ctk.CTkLabel(file_frame, text='ファイル未選択')
        self._file_label.grid(row=0, column=1, padx=5, sticky='w')
        row += 1

        # ── オプション ──
        opt_frame = ctk.CTkFrame(self)
        opt_frame.grid(row=row, column=0, padx=10, pady=5, sticky='ew')
        opt_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(opt_frame, text='学校名').grid(
            row=0, column=0, padx=5, pady=2, sticky='w',
        )
        self._school_var = ctk.StringVar()
        ctk.CTkEntry(opt_frame, textvariable=self._school_var).grid(
            row=0, column=1, padx=5, pady=2, sticky='ew',
        )

        ctk.CTkLabel(opt_frame, text='担任名').grid(
            row=1, column=0, padx=5, pady=2, sticky='w',
        )
        self._teacher_var = ctk.StringVar()
        ctk.CTkEntry(opt_frame, textvariable=self._teacher_var).grid(
            row=1, column=1, padx=5, pady=2, sticky='ew',
        )

        ctk.CTkLabel(opt_frame, text='年度').grid(
            row=2, column=0, padx=5, pady=2, sticky='w',
        )
        self._year_var = ctk.StringVar(value='2025')
        ctk.CTkEntry(opt_frame, textvariable=self._year_var, width=80).grid(
            row=2, column=1, padx=5, pady=2, sticky='w',
        )
        row += 1

        # ── 学生選択 ──
        nav_frame = ctk.CTkFrame(self)
        nav_frame.grid(row=row, column=0, padx=10, pady=5, sticky='ew')
        nav_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            nav_frame, text='<', width=40, command=self._on_prev,
        ).grid(row=0, column=0, padx=2, pady=5)

        self._student_label = ctk.CTkLabel(nav_frame, text='— / —')
        self._student_label.grid(row=0, column=1, padx=5)

        ctk.CTkButton(
            nav_frame, text='>', width=40, command=self._on_next,
        ).grid(row=0, column=2, padx=2, pady=5)

        self._student_name_label = ctk.CTkLabel(nav_frame, text='')
        self._student_name_label.grid(row=0, column=3, padx=10, sticky='w')
        row += 1

        # ── 学生一覧（Listbox） ──
        self._listbox = ctk.CTkTextbox(self, height=200)
        self._listbox.grid(row=row, column=0, padx=10, pady=5, sticky='nsew')
        row += 1

        # ── アクションボタン ──
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=row, column=0, padx=10, pady=(5, 10), sticky='ew')

        ctk.CTkButton(
            btn_frame, text='プレビュー', width=120,
            command=self._on_preview_click,
        ).grid(row=0, column=0, padx=5, pady=5)

        ctk.CTkButton(
            btn_frame, text='全員印刷', width=120,
            command=self._on_print_all,
        ).grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkButton(
            btn_frame, text='閉じる', width=80,
            command=self.destroy,
        ).grid(row=0, column=2, padx=5, pady=5)

    # ── データ読込 ───────────────────────────────────────────────────────

    def _on_load(self) -> None:
        path = fd.askopenfilename(
            title='C4th Excel を選択',
            filetypes=[('Excel ファイル', '*.xlsx *.xls'), ('すべて', '*.*')],
        )
        if not path:
            return

        try:
            df, unmapped = import_c4th_excel(path)
            self._df = df
            self._current_row_idx = 0

            import os
            self._file_label.configure(text=os.path.basename(path))

            # 一覧表示
            self._listbox.delete('1.0', 'end')
            for i, row in df.iterrows():
                name = row.get('氏名', row.get('正式氏名', f'No.{i}'))
                klass = row.get('組', '?')
                num = row.get('出席番号', '?')
                self._listbox.insert('end', f'{klass}組 {num}番 {name}\n')

            self._update_nav()

        except Exception as e:
            mb.showerror('読み込みエラー', str(e), parent=self)

    # ── ナビゲーション ───────────────────────────────────────────────────

    def _on_prev(self) -> None:
        if self._df is not None and self._current_row_idx > 0:
            self._current_row_idx -= 1
            self._update_nav()
            self._auto_preview()

    def _on_next(self) -> None:
        if self._df is not None and self._current_row_idx < len(self._df) - 1:
            self._current_row_idx += 1
            self._update_nav()
            self._auto_preview()

    def _update_nav(self) -> None:
        if self._df is None or len(self._df) == 0:
            self._student_label.configure(text='— / —')
            self._student_name_label.configure(text='')
            return

        total = len(self._df)
        idx = self._current_row_idx
        self._student_label.configure(text=f'{idx + 1} / {total}')

        row = self._df.iloc[idx]
        name = row.get('氏名', row.get('正式氏名', ''))
        self._student_name_label.configure(text=str(name))

    # ── プレビュー・印刷 ─────────────────────────────────────────────────

    def _get_options(self) -> dict:
        opts = dict(self._options)
        opts['school_name'] = self._school_var.get()
        opts['teacher_name'] = self._teacher_var.get()
        with contextlib.suppress(ValueError):
            opts['fiscal_year'] = int(self._year_var.get())
        return opts

    def _get_filled_lay(self, row_idx: int) -> LayFile | None:
        if self._df is None:
            return None
        row = self._df.iloc[row_idx].to_dict()
        return fill_layout(self._lay, row, self._get_options())

    def _auto_preview(self) -> None:
        if self._on_preview:
            filled = self._get_filled_lay(self._current_row_idx)
            if filled:
                self._on_preview(filled)

    def _on_preview_click(self) -> None:
        if self._df is None:
            mb.showwarning('データ未読込', 'Excel ファイルを先に読み込んでください。', parent=self)
            return
        self._auto_preview()

    def _on_print_all(self) -> None:
        if self._df is None:
            mb.showwarning('データ未読込', 'Excel ファイルを先に読み込んでください。', parent=self)
            return
        if self._on_print is None:
            mb.showinfo('印刷', '印刷機能は Phase 3 で実装予定です。', parent=self)
            return

        filled_layouts = []
        opts = self._get_options()
        for i in range(len(self._df)):
            row = self._df.iloc[i].to_dict()
            filled = fill_layout(self._lay, row, opts)
            filled_layouts.append(filled)

        self._on_print(filled_layouts)
