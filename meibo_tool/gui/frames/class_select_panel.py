"""クラス・学年選択パネル（セクション②）

ファイル読込後に有効化される。
  set_data(df)         : df から学年/組の組み合わせを抽出してラジオボタンを生成
  get_filter() -> dict : 選択された {'学年': '3', '組': '2'} 等を返す
                         学年全体の場合 {'学年': '3', '組': None}
                         学校全体の場合 {'学年': None, '組': None}
  get_available_classes(学年) : バッチ生成用 (学年, 組) タプルのリストを返す
  has_special_needs()  : 特別支援学級が存在するかどうかを返す
"""

from __future__ import annotations

from collections.abc import Callable

import customtkinter as ctk
import pandas as pd

from core.special_needs import is_special_needs_class


class ClassSelectPanel(ctk.CTkFrame):
    """学年・クラス選択パネル。"""

    def __init__(
        self, master, on_select: Callable,
        title: str = '② クラス選択',
    ) -> None:
        """
        on_select(filter_dict) をラジオボタン変更時に呼ぶ。
        filter_dict = {'学年': str | None, '組': str | None}
        """
        super().__init__(master, corner_radius=6)
        self._on_select = on_select
        self._var = ctk.StringVar(value='')
        self._classes: dict[str, tuple[str | None, str | None, int]] = {}
        self._radio_frame: ctk.CTkFrame | None = None
        self._has_special_needs = False

        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(size=13, weight='bold'),
        ).grid(row=0, column=0, sticky='w', padx=10, pady=(8, 4))

        self._placeholder = ctk.CTkLabel(
            self,
            text='（名簿ファイルを読み込むと\nクラス一覧が表示されます）',
            text_color='gray',
            wraplength=260,
            justify='left',
        )
        self._placeholder.grid(row=1, column=0, padx=10, pady=(4, 8), sticky='w')

    # ── 外部 API ──────────────────────────────────────────────────────────

    def set_data(self, df: pd.DataFrame) -> None:
        """データをセットしてラジオボタンを再生成する。"""
        if self._radio_frame is not None:
            self._radio_frame.destroy()
            self._radio_frame = None
        self._classes.clear()
        self._var.set('')
        self._has_special_needs = False

        if '学年' not in df.columns or '組' not in df.columns:
            self._placeholder.grid()
            return

        self._placeholder.grid_remove()

        self._radio_frame = ctk.CTkFrame(self, fg_color='transparent')
        self._radio_frame.grid(row=1, column=0, sticky='ew', padx=5, pady=(0, 8))
        self._radio_frame.grid_columnconfigure(0, weight=1)

        groups = (
            df.groupby(['学年', '組'], sort=True)
            .size()
            .reset_index(name='count')
        )

        # 通常学級と特別支援学級を分離
        regular_dict: dict[str, list[tuple[str, int]]] = {}
        special_dict: dict[str, list[tuple[str, int]]] = {}
        for _, row in groups.iterrows():
            g = str(row['学年'])
            k = str(row['組'])
            c = int(row['count'])
            if is_special_needs_class(k):
                special_dict.setdefault(g, []).append((k, c))
            else:
                regular_dict.setdefault(g, []).append((k, c))

        self._has_special_needs = bool(special_dict)

        row_num = 0
        total = len(df)
        first_class_val: str | None = None

        # 全校
        ctk.CTkRadioButton(
            self._radio_frame,
            text=f'全校（{total}名）',
            value='-',
            variable=self._var,
            command=self._on_change,
        ).grid(row=row_num, column=0, padx=10, pady=(4, 2), sticky='w')
        self._classes['-'] = (None, None, total)
        row_num += 1

        def _int_key(s: str) -> int:
            try:
                return int(s)
            except ValueError:
                return 0

        # 通常学級を表示
        for grade in sorted(regular_dict.keys(), key=_int_key):
            class_list = regular_dict[grade]
            grade_total = sum(c for _, c in class_list)
            grade_key = f'{grade}-'

            # 学年全体
            ctk.CTkRadioButton(
                self._radio_frame,
                text=f'{grade}年 全体（{grade_total}名）',
                value=grade_key,
                variable=self._var,
                command=self._on_change,
            ).grid(row=row_num, column=0, padx=10, pady=(6, 1), sticky='w')
            self._classes[grade_key] = (grade, None, grade_total)
            row_num += 1

            for (kumi, count) in sorted(class_list, key=lambda x: _int_key(x[0])):
                val = f'{grade}-{kumi}'
                ctk.CTkRadioButton(
                    self._radio_frame,
                    text=f'　{grade}年{kumi}組（{count}名）',
                    value=val,
                    variable=self._var,
                    command=self._on_change,
                ).grid(row=row_num, column=0, padx=(20, 10), pady=1, sticky='w')
                self._classes[val] = (grade, kumi, count)
                row_num += 1

                if first_class_val is None:
                    first_class_val = val

        # 特別支援学級がある場合はラベルで表示（情報のみ）
        if special_dict:
            ctk.CTkLabel(
                self._radio_frame,
                text='── 特別支援学級 ──',
                font=ctk.CTkFont(size=11, weight='bold'),
                text_color='gray40',
            ).grid(row=row_num, column=0, padx=10, pady=(8, 2), sticky='w')
            row_num += 1

            for grade in sorted(special_dict.keys(), key=_int_key):
                for (kumi, count) in sorted(special_dict[grade]):
                    ctk.CTkLabel(
                        self._radio_frame,
                        text=f'　{grade}年 {kumi}（{count}名）',
                        text_color='gray50',
                        font=ctk.CTkFont(size=11),
                    ).grid(row=row_num, column=0, padx=(20, 10), pady=1, sticky='w')
                    row_num += 1

        # 最初の通常クラスを自動選択
        if first_class_val:
            self._var.set(first_class_val)
            self._on_change()

    def get_filter(self) -> dict[str, str | None]:
        """現在の選択フィルターを返す。"""
        val = self._var.get()
        if not val or val == '-':
            return {'学年': None, '組': None}
        parts = val.split('-', 1)
        return {
            '学年': parts[0] or None,
            '組': parts[1] or None,
        }

    def get_available_classes(
        self, 学年: str | None = None
    ) -> list[tuple[str, str]]:
        """
        バッチ生成用: フィルタに該当する全通常クラスの (学年, 組) リストを返す。
        学年=None の場合は全通常クラスを返す。特別支援学級は含まない。
        """
        result = []
        for _val, (g, k, _) in self._classes.items():
            if g is None or k is None:
                continue
            if is_special_needs_class(k):
                continue
            if 学年 is None or g == 学年:
                result.append((g, k))
        return sorted(result)

    def has_special_needs(self) -> bool:
        """データ内に特別支援学級が存在するかどうかを返す。"""
        return self._has_special_needs

    # ── 内部 ─────────────────────────────────────────────────────────────

    def _on_change(self) -> None:
        self._on_select(self.get_filter())
