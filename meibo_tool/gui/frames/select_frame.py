"""テンプレート選択 + 出力オプションパネル（セクション③④）

SPEC.md §3.2 参照。
"""

from __future__ import annotations

from collections.abc import Callable

import customtkinter as ctk

from templates.template_registry import get_display_groups


class SelectFrame(ctk.CTkFrame):
    """テンプレート選択ラジオボタン + 名前表示モード + 出力オプション。"""

    def __init__(
        self,
        master,
        config: dict,
        on_teacher_save: Callable | None = None,
        on_school_name_save: Callable | None = None,
        on_template_change: Callable | None = None,
    ) -> None:
        super().__init__(master, corner_radius=6)
        self._config = config
        self._on_teacher_save = on_teacher_save
        self._on_school_name_save = on_school_name_save
        self._on_template_change = on_template_change
        self.grid_columnconfigure(0, weight=1)

        row = 0

        # ── ③ テンプレート選択 ────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text='③ テンプレート選択',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=row, column=0, sticky='w', padx=10, pady=(8, 4))
        row += 1

        self._tmpl_var = ctk.StringVar(value='')
        self._radio_widgets: list[ctk.CTkRadioButton] = []

        for display_name, key, icon, _desc in get_display_groups():
            rb = ctk.CTkRadioButton(
                self,
                text=f'{icon} {display_name}',
                value=key,
                variable=self._tmpl_var,
            )
            rb.grid(row=row, column=0, padx=(20, 10), pady=2, sticky='w')
            self._radio_widgets.append(rb)
            row += 1

        # ── 名前の表示方法 ────────────────────────────────────────────────
        ctk.CTkFrame(self, height=1, fg_color='gray80').grid(
            row=row, column=0, sticky='ew', padx=10, pady=6
        )
        row += 1

        ctk.CTkLabel(
            self,
            text='名前の表示方法:',
            font=ctk.CTkFont(size=11),
        ).grid(row=row, column=0, sticky='w', padx=10, pady=(4, 2))
        row += 1

        self._mode_var = ctk.StringVar(value='furigana')
        self._mode_widgets: list[ctk.CTkRadioButton] = []
        for val, label in (
            ('furigana', 'ふりがなつき（上: かな　下: 漢字）'),
            ('kanji', '漢字のみ'),
            ('kana', 'ひらがなのみ'),
        ):
            rb = ctk.CTkRadioButton(
                self, text=label, value=val, variable=self._mode_var
            )
            rb.grid(row=row, column=0, padx=(30, 10), pady=1, sticky='w')
            self._mode_widgets.append(rb)
            row += 1

        # テンプレート / 名前表示変更コールバック
        self._tmpl_var.trace_add('write', self._fire_template_change)
        self._mode_var.trace_add('write', self._fire_template_change)

        # ── ④ 出力オプション ──────────────────────────────────────────────
        ctk.CTkFrame(self, height=1, fg_color='gray80').grid(
            row=row, column=0, sticky='ew', padx=10, pady=6
        )
        row += 1

        ctk.CTkLabel(
            self,
            text='④ 出力オプション',
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=row, column=0, sticky='w', padx=10, pady=(4, 4))
        row += 1

        # 学校名（入力フィールド + 保存ボタン）
        ctk.CTkLabel(self, text='学校名:').grid(row=row, column=0, padx=10, sticky='w')
        row += 1

        school_row_frame = ctk.CTkFrame(self, fg_color='transparent')
        school_row_frame.grid(row=row, column=0, padx=10, pady=(2, 4), sticky='ew')
        school_row_frame.grid_columnconfigure(0, weight=1)
        row += 1

        self._school_var = ctk.StringVar(value=config.get('school_name', ''))
        self._school_entry = ctk.CTkEntry(
            school_row_frame,
            textvariable=self._school_var,
            placeholder_text='例：那覇市立天久小学校',
        )
        self._school_entry.grid(row=0, column=0, sticky='ew')

        self._school_save_btn = ctk.CTkButton(
            school_row_frame,
            text='保存',
            width=48,
            command=self._on_school_save_click,
        )
        self._school_save_btn.grid(row=0, column=1, padx=(4, 0))

        # 年度
        ctk.CTkLabel(self, text='年度（西暦）:').grid(
            row=row, column=0, padx=10, sticky='w'
        )
        row += 1
        self._fy_var = ctk.StringVar(value=str(config.get('fiscal_year', '')))
        ctk.CTkEntry(self, textvariable=self._fy_var, width=90).grid(
            row=row, column=0, padx=(20, 10), pady=2, sticky='w'
        )
        row += 1

        # 担任名（入力フィールド + 保存ボタン）
        ctk.CTkLabel(self, text='担任名:').grid(row=row, column=0, padx=10, sticky='w')
        row += 1

        teacher_row_frame = ctk.CTkFrame(self, fg_color='transparent')
        teacher_row_frame.grid(row=row, column=0, padx=10, pady=(2, 8), sticky='ew')
        teacher_row_frame.grid_columnconfigure(0, weight=1)
        row += 1

        self._teacher_var = ctk.StringVar(value='')
        self._teacher_entry = ctk.CTkEntry(
            teacher_row_frame,
            textvariable=self._teacher_var,
            placeholder_text='例：山田先生',
        )
        self._teacher_entry.grid(row=0, column=0, sticky='ew')

        self._save_btn = ctk.CTkButton(
            teacher_row_frame,
            text='保存',
            width=48,
            command=self._on_save_click,
        )
        self._save_btn.grid(row=0, column=1, padx=(4, 0))

        self._all_widgets = self._radio_widgets + self._mode_widgets + [
            self._school_entry,
            self._school_save_btn,
            self._teacher_entry,
            self._save_btn,
        ]

    # ── 外部 API ─────────────────────────────────────────────────────────

    def set_enabled(self, enabled: bool) -> None:
        state = 'normal' if enabled else 'disabled'
        for w in self._all_widgets:
            w.configure(state=state)

    def set_teacher(self, name: str) -> None:
        """担任名フィールドを設定する（クラス選択時に App から呼ぶ）。"""
        self._teacher_var.set(name)

    def set_school_name(self, name: str) -> None:
        """学校名フィールドを設定する（App から呼ぶ）。"""
        self._school_var.set(name)

    def get_selected_template(self) -> str:
        return self._tmpl_var.get()

    def get_options(self) -> dict:
        default_fy = self._config.get('fiscal_year', 2025)
        try:
            fy = int(self._fy_var.get())
            if not (2000 <= fy <= 2100):
                fy = default_fy
        except ValueError:
            fy = default_fy
        return {
            'fiscal_year': fy,
            'teacher_name': self._teacher_var.get(),
            'school_name': self._school_var.get(),
            'name_display': self._mode_var.get(),
        }

    # ── 内部 ─────────────────────────────────────────────────────────────

    def _fire_template_change(self, *_args) -> None:
        """テンプレートまたは名前表示モード変更時にコールバックを発火する。"""
        if self._on_template_change:
            self._on_template_change()

    def _on_school_save_click(self) -> None:
        if self._on_school_name_save:
            self._on_school_name_save(self._school_var.get())

    def _on_save_click(self) -> None:
        if self._on_teacher_save:
            self._on_teacher_save(self._teacher_var.get())
