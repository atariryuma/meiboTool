"""印刷ダイアログ

プリンター選択 → 印刷実行を行うダイアログ。
データ差込済みの LayFile リストを受け取り、全ページを印刷する。
"""

from __future__ import annotations

import threading
import tkinter.messagebox as mb
from typing import TYPE_CHECKING

import customtkinter as ctk

from core.win_printer import HAS_WIN32, PrintJob, enumerate_printers, get_default_printer

if TYPE_CHECKING:
    from core.lay_parser import LayFile


class PrintDialog(ctk.CTkToplevel):
    """プリンター選択ダイアログ。"""

    def __init__(
        self, master: ctk.CTkBaseClass,
        layouts: list[LayFile],
    ) -> None:
        super().__init__(master)
        self.title('印刷')
        self.geometry('400x250')
        self.transient(master)
        self.grab_set()

        self._layouts = layouts
        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        row = 0

        # ── タイトル ──
        ctk.CTkLabel(
            self, text='印刷設定',
            font=ctk.CTkFont(size=16, weight='bold'),
        ).grid(row=row, column=0, padx=15, pady=(15, 5), sticky='w')
        row += 1

        if not HAS_WIN32:
            ctk.CTkLabel(
                self,
                text='pywin32 がインストールされていません。\n'
                     'pip install pywin32 を実行してください。',
                text_color='red',
            ).grid(row=row, column=0, padx=15, pady=10)
            return

        # ── プリンター選択 ──
        printers = enumerate_printers()
        printer_names = [p['name'] for p in printers]
        default = get_default_printer() or (printer_names[0] if printer_names else '')

        ctk.CTkLabel(self, text='プリンター:').grid(
            row=row, column=0, padx=15, pady=(10, 2), sticky='w',
        )
        row += 1

        self._printer_var = ctk.StringVar(value=default)
        if printer_names:
            self._printer_combo = ctk.CTkComboBox(
                self, values=printer_names, variable=self._printer_var,
                width=350,
            )
            self._printer_combo.grid(row=row, column=0, padx=15, pady=2, sticky='ew')
        else:
            ctk.CTkLabel(self, text='プリンターが見つかりません', text_color='red').grid(
                row=row, column=0, padx=15, pady=2,
            )
        row += 1

        # ── ページ数 ──
        ctk.CTkLabel(
            self, text=f'印刷ページ数: {len(self._layouts)} ページ',
        ).grid(row=row, column=0, padx=15, pady=(10, 2), sticky='w')
        row += 1

        # ── 進捗バー ──
        self._progress = ctk.CTkProgressBar(self, width=350)
        self._progress.grid(row=row, column=0, padx=15, pady=5, sticky='ew')
        self._progress.set(0)
        row += 1

        self._status_label = ctk.CTkLabel(self, text='')
        self._status_label.grid(row=row, column=0, padx=15, pady=2, sticky='w')
        row += 1

        # ── ボタン ──
        btn_frame = ctk.CTkFrame(self, fg_color='transparent')
        btn_frame.grid(row=row, column=0, padx=15, pady=(10, 15), sticky='e')

        self._print_btn = ctk.CTkButton(
            btn_frame, text='印刷', width=100,
            command=self._on_print,
        )
        self._print_btn.grid(row=0, column=0, padx=5)

        ctk.CTkButton(
            btn_frame, text='キャンセル', width=100,
            command=self.destroy,
        ).grid(row=0, column=1, padx=5)

    def _on_print(self) -> None:
        """印刷を実行する（バックグラウンドスレッド）。"""
        printer_name = self._printer_var.get()
        if not printer_name:
            mb.showwarning('プリンター未選択', 'プリンターを選択してください。', parent=self)
            return

        self._print_btn.configure(state='disabled')
        self._status_label.configure(text='印刷準備中...')

        thread = threading.Thread(
            target=self._print_worker,
            args=(printer_name,),
            daemon=True,
        )
        thread.start()

    def _print_worker(self, printer_name: str) -> None:
        """バックグラウンドで印刷を実行する。"""
        total = len(self._layouts)
        try:
            with PrintJob(printer_name) as job:
                job.start(f'名簿印刷 ({total}ページ)')
                for i, lay in enumerate(self._layouts):
                    job.print_page(lay)
                    progress = (i + 1) / total
                    self.after(0, self._update_progress, progress, i + 1, total)

            self.after(0, self._on_print_done, None)

        except Exception as e:
            self.after(0, self._on_print_done, str(e))

    def _update_progress(self, progress: float, current: int, total: int) -> None:
        """進捗を更新する（メインスレッドで実行）。"""
        if not self.winfo_exists():
            return
        self._progress.set(progress)
        self._status_label.configure(text=f'{current} / {total} ページ印刷中...')

    def _on_print_done(self, error: str | None) -> None:
        """印刷完了（メインスレッドで実行）。"""
        if not self.winfo_exists():
            return
        self._print_btn.configure(state='normal')
        if error:
            self._status_label.configure(text='印刷エラー')
            mb.showerror('印刷エラー', error, parent=self)
        else:
            self._progress.set(1.0)
            self._status_label.configure(text='印刷完了')
            mb.showinfo('完了', '印刷が完了しました。', parent=self)
