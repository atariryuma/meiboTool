"""交流学級割り当てダイアログ

特別支援学級の児童を通常学級の名簿に含めるために、
各児童の交流学級（通常学級）を割り当てる。
ファイル読込時に未割り当ての特支児童がいる場合に自動表示される。
"""

from __future__ import annotations

from collections.abc import Callable

import customtkinter as ctk
import pandas as pd

from core.special_needs import make_student_key


class ExchangeClassDialog(ctk.CTkToplevel):
    """特支児童の交流学級割り当てダイアログ。

    Parameters
    ----------
    master : CTk ウィジェット
    special_students_df : 特別支援学級在籍の児童 DataFrame
    regular_classes : 通常クラスの ``(学年, 組)`` リスト
    existing_assignments : 既存の割り当て辞書 ``{student_key: "学年-組"}``
    on_confirm : 確定コールバック。``{student_key: "学年-組"}`` を引数に呼ばれる。
    """

    def __init__(
        self,
        master: ctk.CTk,
        special_students_df: pd.DataFrame,
        regular_classes: list[tuple[str, str]],
        existing_assignments: dict[str, str],
        on_confirm: Callable[[dict[str, str]], None],
    ) -> None:
        super().__init__(master)
        self.title('交流学級の割り当て')
        self.geometry('520x480')
        self.resizable(False, True)
        self.transient(master)
        self.grab_set()

        self._special_df = special_students_df
        self._regular_classes = regular_classes
        self._existing = existing_assignments
        self._on_confirm = on_confirm
        self._combo_vars: list[tuple[str, ctk.StringVar]] = []

        self._build_ui()

    # ────────────────────────────────────────────────────────────────────────
    # UI 構築
    # ────────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── 説明ヘッダー ─────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color='transparent')
        header.grid(row=0, column=0, sticky='ew', padx=12, pady=(12, 4))

        ctk.CTkLabel(
            header,
            text='交流学級の割り当て',
            font=ctk.CTkFont(size=14, weight='bold'),
        ).pack(anchor='w')

        ctk.CTkLabel(
            header,
            text=(
                '特別支援学級の児童が所属する通常学級（交流学級）を\n'
                '選択してください。名簿に自動で含まれるようになります。'
            ),
            font=ctk.CTkFont(size=11),
            text_color='gray30',
            wraplength=480,
            justify='left',
        ).pack(anchor='w', pady=(2, 0))

        # ── 児童リスト（スクロール） ──────────────────────────────────────
        scroll = ctk.CTkScrollableFrame(self, corner_radius=6)
        scroll.grid(row=1, column=0, sticky='nsew', padx=12, pady=6)
        scroll.grid_columnconfigure(0, weight=1)
        scroll.grid_columnconfigure(1, weight=0)
        scroll.grid_columnconfigure(2, weight=1)

        # ヘッダー行
        ctk.CTkLabel(
            scroll, text='児童名（在籍学級）',
            font=ctk.CTkFont(size=11, weight='bold'),
        ).grid(row=0, column=0, sticky='w', padx=(4, 8), pady=(0, 4))
        ctk.CTkLabel(
            scroll, text='→', font=ctk.CTkFont(size=11),
        ).grid(row=0, column=1, padx=4, pady=(0, 4))
        ctk.CTkLabel(
            scroll, text='交流学級',
            font=ctk.CTkFont(size=11, weight='bold'),
        ).grid(row=0, column=2, sticky='w', padx=(8, 4), pady=(0, 4))

        # 学年ごとにグループ化した選択肢を構築
        class_by_grade: dict[str, list[str]] = {}
        for g, k in sorted(self._regular_classes):
            class_by_grade.setdefault(g, []).append(f'{g}年{k}組')

        # 全クラスのフラットリスト（コンボボックス用）
        all_options = ['（なし）']
        for g in sorted(class_by_grade.keys()):
            all_options.extend(class_by_grade[g])

        # 児童行を生成
        for i, (_, row) in enumerate(self._special_df.iterrows(), start=1):
            student_key = make_student_key(row)
            name = row.get('氏名', '?')
            grade = row['学年']
            klass = row['組']

            ctk.CTkLabel(
                scroll,
                text=f'{name}（{grade}年{klass}）',
                font=ctk.CTkFont(size=11),
                wraplength=200,
                anchor='w',
            ).grid(row=i, column=0, sticky='w', padx=(4, 8), pady=3)

            ctk.CTkLabel(scroll, text='→').grid(row=i, column=1, padx=4, pady=3)

            # 既存割り当てがあればプリセット
            existing_val = self._existing.get(student_key, '')
            default = '（なし）'
            if existing_val:
                # "1-1" → "1年1組" に変換
                parts = existing_val.split('-', 1)
                if len(parts) == 2:
                    display = f'{parts[0]}年{parts[1]}組'
                    if display in all_options:
                        default = display

            # 同学年のクラスを先頭に配置した選択肢
            grade_options = ['（なし）']
            if grade in class_by_grade:
                grade_options.extend(class_by_grade[grade])
            # 他学年も追加（区切り付き）
            for g in sorted(class_by_grade.keys()):
                if g != grade:
                    grade_options.extend(class_by_grade[g])

            var = ctk.StringVar(value=default)
            combo = ctk.CTkComboBox(
                scroll, variable=var, values=grade_options,
                width=160, state='readonly',
            )
            combo.grid(row=i, column=2, sticky='ew', padx=(8, 4), pady=3)
            self._combo_vars.append((student_key, var))

        # ── ボタン ────────────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(self, fg_color='transparent')
        btn_frame.grid(row=2, column=0, sticky='ew', padx=12, pady=(4, 12))
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btn_frame, text='保存', command=self._on_save, width=120,
        ).grid(row=0, column=0, padx=5, pady=5, sticky='e')

        ctk.CTkButton(
            btn_frame, text='閉じる', command=self.destroy,
            width=120, fg_color='gray',
        ).grid(row=0, column=1, padx=5, pady=5, sticky='w')

    # ────────────────────────────────────────────────────────────────────────
    # イベントハンドラ
    # ────────────────────────────────────────────────────────────────────────

    def _on_save(self) -> None:
        """割り当てを収集してコールバックを呼ぶ。"""
        assignments: dict[str, str] = {}
        for student_key, var in self._combo_vars:
            val = var.get()
            if val and val != '（なし）':
                # "1年1組" → "1-1" に変換
                assignments[student_key] = _display_to_key(val)
            else:
                assignments[student_key] = ''
        self._on_confirm(assignments)
        self.destroy()


def _display_to_key(display: str) -> str:
    """``"1年2組"`` → ``"1-2"`` に変換する。"""
    # "1年2組" パターン
    display = display.strip()
    if '年' in display and display.endswith('組'):
        parts = display.rstrip('組').split('年', 1)
        if len(parts) == 2:
            return f'{parts[0]}-{parts[1]}'
    return display
