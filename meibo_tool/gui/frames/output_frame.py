"""生成ボタン + 進捗バー + システム情報（セクション⑤⑥）

SPEC.md §3.2 参照。
"""

from __future__ import annotations

import os
import threading
from collections.abc import Callable

import customtkinter as ctk


class OutputFrame(ctk.CTkFrame):
    """Excel 生成ボタン・進捗バー・バージョン表示。"""

    def __init__(self, master, on_generate: Callable, config: dict) -> None:
        """
        on_generate() は BaseGenerator インスタンス（or None）を返す Callable。
        スレッドセーフである必要はない（スレッド内から呼ばれる）。
        """
        super().__init__(master, corner_radius=6)
        self._on_generate = on_generate
        self.grid_columnconfigure(0, weight=1)

        row = 0

        # ── ⑤ 生成 ───────────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text='⑤ 生成',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=row, column=0, sticky='w', padx=10, pady=(8, 4))
        row += 1

        self._gen_btn = ctk.CTkButton(
            self,
            text='📄 Excel を生成',
            font=ctk.CTkFont(size=14, weight='bold'),
            height=40,
            command=self._on_click,
        )
        self._gen_btn.grid(row=row, column=0, padx=10, pady=4, sticky='ew')
        row += 1

        self._progress = ctk.CTkProgressBar(self)
        self._progress.grid(row=row, column=0, padx=10, pady=4, sticky='ew')
        self._progress.set(0)
        row += 1

        self._status = ctk.CTkLabel(
            self, text='', text_color='gray', wraplength=270, anchor='w'
        )
        self._status.grid(row=row, column=0, padx=10, pady=(2, 4), sticky='w')
        row += 1

        # ── ⑥ システム ───────────────────────────────────────────────────
        ctk.CTkFrame(self, height=1, fg_color='gray80').grid(
            row=row, column=0, sticky='ew', padx=10, pady=6
        )
        row += 1

        self._ver_label = ctk.CTkLabel(
            self,
            text=f'v{config.get("app_version", "1.0.0")}',
            text_color='gray',
            font=ctk.CTkFont(size=10),
        )
        self._ver_label.grid(row=row, column=0, padx=10, pady=(0, 8), sticky='w')

    # ── 外部 API ─────────────────────────────────────────────────────────

    def set_enabled(self, enabled: bool) -> None:
        self._gen_btn.configure(state='normal' if enabled else 'disabled')

    def show_update_notice(self, info: dict) -> None:
        ver = info.get('version', '?')
        self._status.configure(text=f'🔔 更新あり: v{ver}', text_color='orange')

    # ── 内部 ─────────────────────────────────────────────────────────────

    def _on_click(self) -> None:
        # メインスレッドでジェネレーターを作成（Tkinter ウィジェットへのアクセスはここで完結）
        try:
            gen = self._on_generate()
        except Exception as e:
            self._done(None, f'エラー: {e}')
            return
        if gen is None:
            self._done(None, 'テンプレートを選択してください')
            return
        self._gen_btn.configure(state='disabled')
        self._progress.set(0.1)
        self._status.configure(text='生成中…', text_color='gray')
        threading.Thread(target=self._run, args=(gen,), daemon=True).start()

    def _run(self, gen) -> None:
        try:
            out = gen.generate()
            self.after(0, lambda: self._done(out, None))
        except Exception as e:
            err = str(e)
            self.after(0, lambda: self._done(None, f'エラー: {err}'))

    def _done(self, output_path: str | None, error: str | None) -> None:
        self._progress.set(1.0 if output_path else 0)
        self._gen_btn.configure(state='normal')
        if output_path:
            fname = os.path.basename(output_path)
            self._status.configure(text=f'✅ {fname}', text_color='green')
            import contextlib
            with contextlib.suppress(Exception):
                os.startfile(os.path.dirname(output_path))
        else:
            self._status.configure(
                text=error or 'エラーが発生しました', text_color='red'
            )
