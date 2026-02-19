"""カラムマッピング手動調整ダイアログ

自動マッピングで対応できなかった列を、ユーザーが手動で内部論理名に紐づける。
import 直後に unmapped が存在する場合に表示される。
"""

from __future__ import annotations

from collections.abc import Callable

import customtkinter as ctk

from core.mapper import COLUMN_ALIASES, EXACT_MAP

# 選択肢として表示する内部論理名一覧（重複排除・ソート）
_ALL_LOGICAL_NAMES: list[str] = sorted(
    {*EXACT_MAP.values(), *COLUMN_ALIASES.keys()}
)


class MappingDialog(ctk.CTkToplevel):
    """未マップ列の手動マッピングダイアログ。

    Parameters
    ----------
    master : CTk ウィジェット
    unmapped : 自動マッピングできなかった元カラム名のリスト
    existing_mapped : 既にマッピング済みの論理名セット（重複防止用）
    on_confirm : 確定コールバック。``{元カラム名: 論理名}`` を引数に呼ばれる。
                 スキップされた列は含まない。
    """

    def __init__(
        self,
        master: ctk.CTk,
        unmapped: list[str],
        existing_mapped: set[str],
        on_confirm: Callable[[dict[str, str]], None],
    ) -> None:
        super().__init__(master)
        self.title('カラムマッピング')
        self.geometry('560x440')
        self.resizable(False, True)
        self.transient(master)
        self.grab_set()

        self._unmapped = unmapped
        self._existing_mapped = existing_mapped
        self._on_confirm = on_confirm
        self._combo_vars: list[ctk.StringVar] = []

        self._build_ui()

    # ────────────────────────────────────────────────────────────────────────
    # UI 構築
    # ────────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── 説明ヘッダー ───────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color='transparent')
        header.grid(row=0, column=0, sticky='ew', padx=12, pady=(12, 4))

        ctk.CTkLabel(
            header,
            text='カラムマッピング',
            font=ctk.CTkFont(size=14, weight='bold'),
        ).pack(anchor='w')

        ctk.CTkLabel(
            header,
            text=(
                f'自動マッピングできなかった {len(self._unmapped)} 列があります。\n'
                '必要な列を選択するか、「（スキップ）」のまま無視してください。'
            ),
            font=ctk.CTkFont(size=11),
            text_color='gray30',
            wraplength=520,
            justify='left',
        ).pack(anchor='w', pady=(2, 0))

        # ── マッピング行リスト（スクロール） ──────────────────────────────
        scroll = ctk.CTkScrollableFrame(self, corner_radius=6)
        scroll.grid(row=1, column=0, sticky='nsew', padx=12, pady=6)
        scroll.grid_columnconfigure(0, weight=1)
        scroll.grid_columnconfigure(1, weight=0)
        scroll.grid_columnconfigure(2, weight=1)

        # ヘッダー行
        ctk.CTkLabel(
            scroll, text='元のカラム名', font=ctk.CTkFont(size=11, weight='bold'),
        ).grid(row=0, column=0, sticky='w', padx=(4, 8), pady=(0, 4))
        ctk.CTkLabel(
            scroll, text='→', font=ctk.CTkFont(size=11),
        ).grid(row=0, column=1, padx=4, pady=(0, 4))
        ctk.CTkLabel(
            scroll, text='マッピング先', font=ctk.CTkFont(size=11, weight='bold'),
        ).grid(row=0, column=2, sticky='w', padx=(8, 4), pady=(0, 4))

        # ドロップダウンの選択肢: 既にマップ済みの論理名は除外
        available = [n for n in _ALL_LOGICAL_NAMES if n not in self._existing_mapped]
        options = ['（スキップ）', *available]

        for i, col_name in enumerate(self._unmapped, start=1):
            ctk.CTkLabel(
                scroll, text=col_name, font=ctk.CTkFont(size=11),
                wraplength=200, anchor='w',
            ).grid(row=i, column=0, sticky='w', padx=(4, 8), pady=3)

            ctk.CTkLabel(scroll, text='→').grid(row=i, column=1, padx=4, pady=3)

            var = ctk.StringVar(value='（スキップ）')
            combo = ctk.CTkComboBox(
                scroll, variable=var, values=options,
                width=200, state='readonly',
            )
            combo.grid(row=i, column=2, sticky='ew', padx=(8, 4), pady=3)
            self._combo_vars.append(var)

        # ── ボタン ─────────────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(self, fg_color='transparent')
        btn_frame.grid(row=2, column=0, sticky='ew', padx=12, pady=(4, 12))
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btn_frame, text='確定', command=self._on_ok, width=120,
        ).grid(row=0, column=0, padx=5, pady=5, sticky='e')

        ctk.CTkButton(
            btn_frame, text='すべてスキップ', command=self._on_skip_all,
            width=120, fg_color='gray',
        ).grid(row=0, column=1, padx=5, pady=5, sticky='w')

    # ────────────────────────────────────────────────────────────────────────
    # イベントハンドラ
    # ────────────────────────────────────────────────────────────────────────

    def _collect_mapping(self) -> dict[str, str]:
        """ユーザー選択からマッピング辞書を構築する。スキップ分は含まない。"""
        mapping: dict[str, str] = {}
        for col_name, var in zip(self._unmapped, self._combo_vars, strict=True):
            val = var.get()
            if val and val != '（スキップ）':
                mapping[col_name] = val
        return mapping

    def _on_ok(self) -> None:
        mapping = self._collect_mapping()
        self._on_confirm(mapping)
        self.destroy()

    def _on_skip_all(self) -> None:
        self._on_confirm({})
        self.destroy()
