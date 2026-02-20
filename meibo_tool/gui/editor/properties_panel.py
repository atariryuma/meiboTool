"""オブジェクトプロパティ編集パネル

選択されたオブジェクトの属性（位置、サイズ、テキスト、フォント等）を
フォーム UI で表示・編集する。
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable

import customtkinter as ctk

from core.lay_parser import (
    FIELD_ID_MAP,
    FontInfo,
    LayoutObject,
    ObjectType,
    Point,
    Rect,
)

# ── 定数 ─────────────────────────────────────────────────────────────────────

_FIELD_CHOICES = [''] + [
    f'{fid}: {name}' for fid, name in sorted(FIELD_ID_MAP.items())
]

_H_ALIGN_CHOICES = ['左揃え', '中央', '右揃え']
_V_ALIGN_CHOICES = ['上揃え', '中央', '下揃え']

_FONT_CHOICES = [
    'IPAmj明朝', 'ＭＳ 明朝', 'ＭＳ ゴシック', 'メイリオ', '游ゴシック',
]


class PropertiesPanel(ctk.CTkFrame):
    """右サイドのプロパティ編集パネル。"""

    def __init__(
        self, master: ctk.CTkBaseClass,
        on_change: Callable[[LayoutObject], None] | None = None,
    ) -> None:
        super().__init__(master, width=220)
        self._on_change = on_change
        self._obj: LayoutObject | None = None
        self._obj_index: int = -1
        self._updating = False  # 変更通知の再帰防止
        self._build_ui()
        self._clear()

    # ── UI 構築 ───────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)

        row = 0
        self._title_label = ctk.CTkLabel(
            self, text='プロパティ', font=ctk.CTkFont(size=14, weight='bold'),
        )
        self._title_label.grid(
            row=row, column=0, columnspan=2, padx=8, pady=(8, 4), sticky='w',
        )
        row += 1

        # タイプ表示
        row = self._add_label_row(row, 'タイプ')
        self._type_label = ctk.CTkLabel(self, text='—')
        self._type_label.grid(row=row - 1, column=1, padx=4, sticky='w')

        # ── 位置・サイズ ──
        self._pos_section = ctk.CTkLabel(
            self, text='位置・サイズ', font=ctk.CTkFont(weight='bold'),
        )
        self._pos_section.grid(
            row=row, column=0, columnspan=2, padx=8, pady=(8, 2), sticky='w',
        )
        row += 1

        self._left_var = ctk.StringVar()
        self._top_var = ctk.StringVar()
        self._width_var = ctk.StringVar()
        self._height_var = ctk.StringVar()

        for label, var in [
            ('X', self._left_var), ('Y', self._top_var),
            ('幅', self._width_var), ('高さ', self._height_var),
        ]:
            row = self._add_entry_row(row, label, var)

        # ── テキスト ──
        self._text_section = ctk.CTkLabel(
            self, text='テキスト', font=ctk.CTkFont(weight='bold'),
        )
        self._text_section.grid(
            row=row, column=0, columnspan=2, padx=8, pady=(8, 2), sticky='w',
        )
        row += 1

        self._text_var = ctk.StringVar()
        row = self._add_entry_row(row, '内容', self._text_var)

        self._prefix_var = ctk.StringVar()
        self._suffix_var = ctk.StringVar()
        row = self._add_entry_row(row, '前置', self._prefix_var)
        row = self._add_entry_row(row, '後置', self._suffix_var)

        # ── フィールド ──
        self._field_section = ctk.CTkLabel(
            self, text='フィールド', font=ctk.CTkFont(weight='bold'),
        )
        self._field_section.grid(
            row=row, column=0, columnspan=2, padx=8, pady=(8, 2), sticky='w',
        )
        row += 1

        self._field_var = ctk.StringVar()
        self._field_combo = ctk.CTkComboBox(
            self, values=_FIELD_CHOICES, variable=self._field_var,
            width=180, command=self._on_field_change,
        )
        ctk.CTkLabel(self, text='ID').grid(row=row, column=0, padx=(8, 2), sticky='w')
        self._field_combo.grid(row=row, column=1, padx=4, pady=2, sticky='ew')
        row += 1

        # ── フォント ──
        self._font_section = ctk.CTkLabel(
            self, text='フォント', font=ctk.CTkFont(weight='bold'),
        )
        self._font_section.grid(
            row=row, column=0, columnspan=2, padx=8, pady=(8, 2), sticky='w',
        )
        row += 1

        self._font_name_var = ctk.StringVar()
        self._font_combo = ctk.CTkComboBox(
            self, values=_FONT_CHOICES, variable=self._font_name_var,
            width=180, command=self._on_property_edited,
        )
        ctk.CTkLabel(self, text='名前').grid(row=row, column=0, padx=(8, 2), sticky='w')
        self._font_combo.grid(row=row, column=1, padx=4, pady=2, sticky='ew')
        row += 1

        self._font_size_var = ctk.StringVar()
        row = self._add_entry_row(row, 'サイズ', self._font_size_var)

        # ── 揃え ──
        self._align_section = ctk.CTkLabel(
            self, text='揃え', font=ctk.CTkFont(weight='bold'),
        )
        self._align_section.grid(
            row=row, column=0, columnspan=2, padx=8, pady=(8, 2), sticky='w',
        )
        row += 1

        self._h_align_var = ctk.StringVar()
        self._h_align_combo = ctk.CTkComboBox(
            self, values=_H_ALIGN_CHOICES, variable=self._h_align_var,
            width=180, command=self._on_property_edited,
        )
        ctk.CTkLabel(self, text='水平').grid(row=row, column=0, padx=(8, 2), sticky='w')
        self._h_align_combo.grid(row=row, column=1, padx=4, pady=2, sticky='ew')
        row += 1

        self._v_align_var = ctk.StringVar()
        self._v_align_combo = ctk.CTkComboBox(
            self, values=_V_ALIGN_CHOICES, variable=self._v_align_var,
            width=180, command=self._on_property_edited,
        )
        ctk.CTkLabel(self, text='垂直').grid(row=row, column=0, padx=(8, 2), sticky='w')
        self._v_align_combo.grid(row=row, column=1, padx=4, pady=2, sticky='ew')
        row += 1

        # ── LINE 用座標 ──
        self._line_section = ctk.CTkLabel(
            self, text='線の座標', font=ctk.CTkFont(weight='bold'),
        )
        self._line_section.grid(
            row=row, column=0, columnspan=2, padx=8, pady=(8, 2), sticky='w',
        )
        row += 1

        self._lx1_var = ctk.StringVar()
        self._ly1_var = ctk.StringVar()
        self._lx2_var = ctk.StringVar()
        self._ly2_var = ctk.StringVar()
        row = self._add_entry_row(row, '始X', self._lx1_var)
        row = self._add_entry_row(row, '始Y', self._ly1_var)
        row = self._add_entry_row(row, '終X', self._lx2_var)
        row = self._add_entry_row(row, '終Y', self._ly2_var)

        # 全 StringVar に変更トレース登録
        for var in [
            self._left_var, self._top_var, self._width_var, self._height_var,
            self._text_var, self._prefix_var, self._suffix_var,
            self._font_size_var,
            self._lx1_var, self._ly1_var, self._lx2_var, self._ly2_var,
        ]:
            var.trace_add('write', lambda *_a: self._on_property_edited())

    def _add_label_row(self, row: int, text: str) -> int:
        ctk.CTkLabel(self, text=text).grid(
            row=row, column=0, padx=(8, 2), pady=2, sticky='w',
        )
        return row + 1

    def _add_entry_row(
        self, row: int, label: str, var: ctk.StringVar,
    ) -> int:
        ctk.CTkLabel(self, text=label).grid(
            row=row, column=0, padx=(8, 2), pady=2, sticky='w',
        )
        entry = ctk.CTkEntry(self, textvariable=var, width=120)
        entry.grid(row=row, column=1, padx=4, pady=2, sticky='ew')
        return row + 1

    # ── 公開メソッド ─────────────────────────────────────────────────────

    def set_object(self, obj: LayoutObject | None, index: int = -1) -> None:
        """選択オブジェクトを表示する。"""
        self._obj = obj
        self._obj_index = index

        if obj is None:
            self._clear()
            return

        self._updating = True
        try:
            self._populate(obj)
        finally:
            self._updating = False

    def _clear(self) -> None:
        """全フィールドをクリアする。"""
        self._updating = True
        self._type_label.configure(text='—')
        for var in [
            self._left_var, self._top_var, self._width_var, self._height_var,
            self._text_var, self._prefix_var, self._suffix_var,
            self._font_size_var, self._font_name_var,
            self._lx1_var, self._ly1_var, self._lx2_var, self._ly2_var,
        ]:
            var.set('')
        self._field_var.set('')
        self._h_align_var.set('')
        self._v_align_var.set('')
        self._updating = False

    def _populate(self, obj: LayoutObject) -> None:
        """オブジェクトのプロパティを UI に反映する。"""
        type_names = {
            ObjectType.LABEL: 'ラベル',
            ObjectType.FIELD: 'フィールド',
            ObjectType.LINE: '罫線',
            ObjectType.GROUP: 'グループ',
        }
        self._type_label.configure(
            text=type_names.get(obj.obj_type, '不明'),
        )

        # 位置・サイズ
        if obj.rect:
            self._left_var.set(str(obj.rect.left))
            self._top_var.set(str(obj.rect.top))
            self._width_var.set(str(obj.rect.width))
            self._height_var.set(str(obj.rect.height))
        else:
            for var in [self._left_var, self._top_var,
                        self._width_var, self._height_var]:
                var.set('')

        # テキスト
        self._text_var.set(obj.text)
        self._prefix_var.set(obj.prefix)
        self._suffix_var.set(obj.suffix)

        # フィールド
        if obj.obj_type == ObjectType.FIELD and obj.field_id:
            name = FIELD_ID_MAP.get(obj.field_id, '')
            self._field_var.set(f'{obj.field_id}: {name}' if name else str(obj.field_id))
        else:
            self._field_var.set('')

        # フォント
        self._font_name_var.set(obj.font.name or '')
        self._font_size_var.set(str(obj.font.size_pt))

        # 揃え
        self._h_align_var.set(
            _H_ALIGN_CHOICES[obj.h_align] if 0 <= obj.h_align <= 2 else '',
        )
        self._v_align_var.set(
            _V_ALIGN_CHOICES[obj.v_align] if 0 <= obj.v_align <= 2 else '',
        )

        # LINE 座標
        if obj.line_start:
            self._lx1_var.set(str(obj.line_start.x))
            self._ly1_var.set(str(obj.line_start.y))
        if obj.line_end:
            self._lx2_var.set(str(obj.line_end.x))
            self._ly2_var.set(str(obj.line_end.y))

    # ── 変更通知 ─────────────────────────────────────────────────────────

    def _on_field_change(self, _value: str = '') -> None:
        self._on_property_edited()

    def _on_property_edited(self, *_args: object) -> None:
        """プロパティが変更された時にモデルを更新する。"""
        if self._updating or self._obj is None:
            return

        obj = self._obj

        # 位置・サイズ → rect
        if obj.rect is not None:
            left = self._safe_int(self._left_var.get(), obj.rect.left)
            top = self._safe_int(self._top_var.get(), obj.rect.top)
            w = self._safe_int(self._width_var.get(), obj.rect.width)
            h = self._safe_int(self._height_var.get(), obj.rect.height)
            obj.rect = Rect(left, top, left + w, top + h)

        # テキスト
        obj.text = self._text_var.get()
        obj.prefix = self._prefix_var.get()
        obj.suffix = self._suffix_var.get()

        # フィールド ID
        field_str = self._field_var.get()
        if ':' in field_str:
            with contextlib.suppress(ValueError):
                obj.field_id = int(field_str.split(':')[0].strip())

        # フォント
        font_name = self._font_name_var.get()
        font_size = self._safe_float(
            self._font_size_var.get(), obj.font.size_pt,
        )
        obj.font = FontInfo(font_name, font_size)

        # 揃え
        h_idx = _H_ALIGN_CHOICES.index(self._h_align_var.get()) \
            if self._h_align_var.get() in _H_ALIGN_CHOICES else obj.h_align
        v_idx = _V_ALIGN_CHOICES.index(self._v_align_var.get()) \
            if self._v_align_var.get() in _V_ALIGN_CHOICES else obj.v_align
        obj.h_align = h_idx
        obj.v_align = v_idx

        # LINE 座標
        if obj.line_start is not None:
            obj.line_start = Point(
                self._safe_int(self._lx1_var.get(), obj.line_start.x),
                self._safe_int(self._ly1_var.get(), obj.line_start.y),
            )
        if obj.line_end is not None:
            obj.line_end = Point(
                self._safe_int(self._lx2_var.get(), obj.line_end.x),
                self._safe_int(self._ly2_var.get(), obj.line_end.y),
            )

        if self._on_change:
            self._on_change(obj)

    @staticmethod
    def _safe_int(s: str, default: int) -> int:
        try:
            return int(s)
        except (ValueError, TypeError):
            return default

    @staticmethod
    def _safe_float(s: str, default: float) -> float:
        try:
            return float(s)
        except (ValueError, TypeError):
            return default
