"""LayFile マルチバックエンド描画エンジン

LayFile オブジェクトを異なる描画先（Canvas / PIL / GDI）に統一的に描画する。
座標変換（モデル → 描画先）もここで一元管理する。

使用方法:
    from core.lay_renderer import CanvasBackend, LayRenderer
    backend = CanvasBackend(canvas, scale=0.5, offset_x=20, offset_y=20)
    renderer = LayRenderer(lay, backend)
    renderer.render_all()
"""

from __future__ import annotations

import tkinter as tk
from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

from core.lay_parser import (
    LayFile,
    LayoutObject,
    ObjectType,
    resolve_field_name,
)

if TYPE_CHECKING:
    pass

# ── 定数 ─────────────────────────────────────────────────────────────────────

# ページ背景・グリッド
_PAGE_BG = '#FFFFFF'
_PAGE_BORDER = '#999999'

# オブジェクト色
_LABEL_BG = ''            # 透明（印刷イメージと一致）
_FIELD_BG = '#F0F6FF'     # ごく薄い水色（編集時のみ見える目印）
_LABEL_OUTLINE = ''       # 透明（枠線なし）
_FIELD_OUTLINE = '#B0C4DE' # 薄い点線枠（編集時の目印）
_LINE_COLOR = '#000000'
_TEXT_COLOR = '#000000'
_FIELD_TEXT_COLOR = '#4A7AB5'

# 選択表示
_SELECT_OUTLINE = '#1A73E8'
_HANDLE_COLOR = '#1A73E8'
_HANDLE_SIZE = 5

# Windows フォントパス候補
_FONT_PATHS = [
    'C:/Windows/Fonts/ipamjm.ttf',
    'C:/Windows/Fonts/meiryo.ttc',
    'C:/Windows/Fonts/YuGothR.ttc',
    'C:/Windows/Fonts/msgothic.ttc',
]


# ── 座標変換 ─────────────────────────────────────────────────────────────────


def model_to_canvas(
    x: int, y: int, scale: float,
    offset_x: float = 0, offset_y: float = 0,
) -> tuple[float, float]:
    """モデル座標 (0.25mm単位) → Canvas ピクセル。"""
    return x * scale + offset_x, y * scale + offset_y


def canvas_to_model(
    cx: float, cy: float, scale: float,
    offset_x: float = 0, offset_y: float = 0,
) -> tuple[int, int]:
    """Canvas ピクセル → モデル座標 (0.25mm単位)。"""
    return round((cx - offset_x) / scale), round((cy - offset_y) / scale)


def model_to_printer(x: int, dpi: int) -> int:
    """モデル座標 → プリンタドット。"""
    return round(x * 0.25 * dpi / 25.4)


# ── 抽象バックエンド ─────────────────────────────────────────────────────────


class RenderBackend(ABC):
    """描画先の抽象インターフェース。"""

    @abstractmethod
    def draw_rect(
        self, left: float, top: float, right: float, bottom: float,
        fill: str = '', outline: str = '#000000', width: int = 1,
    ) -> None:
        """矩形を描画する。"""

    @abstractmethod
    def draw_line(
        self, x1: float, y1: float, x2: float, y2: float,
        color: str = '#000000', width: int = 1,
    ) -> None:
        """直線を描画する。"""

    @abstractmethod
    def draw_text(
        self, x: float, y: float, w: float, h: float,
        text: str, font_name: str, font_size: float,
        h_align: int = 0, v_align: int = 0,
        color: str = '#000000',
    ) -> None:
        """テキストを描画する。"""


# ── Canvas バックエンド ──────────────────────────────────────────────────────


class CanvasBackend(RenderBackend):
    """tkinter Canvas への描画バックエンド。"""

    def __init__(
        self, canvas: tk.Canvas, scale: float,
        offset_x: float = 0, offset_y: float = 0,
    ) -> None:
        self._canvas = canvas
        self._scale = scale
        self._ox = offset_x
        self._oy = offset_y

    @property
    def scale(self) -> float:
        return self._scale

    @property
    def offset_x(self) -> float:
        return self._ox

    @property
    def offset_y(self) -> float:
        return self._oy

    def _to_px(self, x: int, y: int) -> tuple[float, float]:
        """モデル座標をキャンバスピクセルに変換する。"""
        return model_to_canvas(x, y, self._scale, self._ox, self._oy)

    def draw_rect(
        self, left: float, top: float, right: float, bottom: float,
        fill: str = '', outline: str = '#000000', width: int = 1,
        tags: tuple[str, ...] = (),
    ) -> int:
        return self._canvas.create_rectangle(
            left, top, right, bottom,
            fill=fill, outline=outline, width=width, tags=tags,
        )

    def draw_line(
        self, x1: float, y1: float, x2: float, y2: float,
        color: str = '#000000', width: int = 1,
        tags: tuple[str, ...] = (),
    ) -> int:
        return self._canvas.create_line(
            x1, y1, x2, y2, fill=color, width=width, tags=tags,
        )

    def draw_text(
        self, x: float, y: float, w: float, h: float,
        text: str, font_name: str, font_size: float,
        h_align: int = 0, v_align: int = 0,
        color: str = '#000000', tags: tuple[str, ...] = (),
        bold: bool = False, italic: bool = False,
    ) -> int | None:
        import tkinter.font as tkfont

        if not text:
            return None

        # \r\n → \n に正規化
        text = text.replace('\r\n', '\n').replace('\r', '\n')

        # Canvas フォントサイズ（ポイント → Canvas 用スケール）
        fname = font_name or 'IPAmj明朝'
        scaled_size = max(int(font_size * self._scale / 0.5), 6)

        # フォントスタイル
        weight = 'bold' if bold else 'normal'
        slant = 'italic' if italic else 'roman'

        # ── 縦書き判定 ──
        # 幅が狭くて高さがある矩形 + 改行なしテキスト → 縦書き
        is_vertical = (
            '\n' not in text
            and len(text) > 1
            and w < h * 0.5
            and w < 120 * self._scale
        )
        has_newline = '\n' in text

        if is_vertical:
            return self._draw_vertical_text(
                x, y, w, h, text, fname, scaled_size,
                weight, slant, h_align, v_align, color, tags,
            )

        # ── 複数行テキスト ──
        if has_newline:
            return self._draw_multiline_text(
                x, y, w, h, text, fname, scaled_size,
                weight, slant, h_align, v_align, color, tags,
            )

        # ── 通常の単一行テキスト ──
        # 自動文字サイズ調整（ボックス幅に収まるよう縮小）
        box_w = max(w - 4, 1)
        if box_w > 0:
            try:
                font_obj = tkfont.Font(
                    family=fname, size=scaled_size,
                    weight=weight, slant=slant,
                )
                text_w = font_obj.measure(text)
                while text_w > box_w and scaled_size > 5:
                    scaled_size -= 1
                    font_obj.configure(size=scaled_size)
                    text_w = font_obj.measure(text)
            except Exception:
                pass

        font_spec = (fname, scaled_size, weight, slant)
        tx, ty, anchor = self._calc_text_pos(
            x, y, w, h, h_align, v_align,
        )

        return self._canvas.create_text(
            tx, ty, text=text, font=font_spec, fill=color,
            anchor=anchor, width=0, tags=tags,
        )

    def _calc_text_pos(
        self, x: float, y: float, w: float, h: float,
        h_align: int, v_align: int,
    ) -> tuple[float, float, str]:
        """テキスト描画位置と anchor を計算する。"""
        if h_align == 0:
            tx = x + 2
        elif h_align == 2:
            tx = x + w - 2
        else:
            tx = x + w / 2

        if v_align == 0:
            ty = y + 2
        elif v_align == 2:
            ty = y + h - 2
        else:
            ty = y + h / 2

        anchor_h = {0: 'w', 1: 'center', 2: 'e'}.get(h_align, 'center')
        anchor_v = {0: 'n', 1: 'center', 2: 's'}.get(v_align, 'center')
        anchor_map = {
            ('n', 'w'): 'nw', ('n', 'center'): 'n', ('n', 'e'): 'ne',
            ('center', 'w'): 'w', ('center', 'center'): 'center',
            ('center', 'e'): 'e',
            ('s', 'w'): 'sw', ('s', 'center'): 's', ('s', 'e'): 'se',
        }
        anchor = anchor_map.get((anchor_v, anchor_h), 'center')
        return tx, ty, anchor

    def _draw_vertical_text(
        self, x: float, y: float, w: float, h: float,
        text: str, fname: str, scaled_size: int,
        weight: str, slant: str,
        h_align: int, v_align: int,
        color: str, tags: tuple[str, ...],
    ) -> int | None:
        """縦書きテキスト: 各文字を縦に1文字ずつ描画する。"""
        import tkinter.font as tkfont

        chars = list(text)
        n = len(chars)
        if n == 0:
            return None

        # 各文字の高さを計測してサイズ調整
        try:
            font_obj = tkfont.Font(
                family=fname, size=scaled_size,
                weight=weight, slant=slant,
            )
            char_h = font_obj.metrics('linespace')
        except Exception:
            char_h = scaled_size

        # 全文字の高さがボックスに収まるまで縮小
        total_h = char_h * n
        box_h = max(h - 4, 1)
        while total_h > box_h and scaled_size > 5:
            scaled_size -= 1
            try:
                font_obj = tkfont.Font(
                    family=fname, size=scaled_size,
                    weight=weight, slant=slant,
                )
                char_h = font_obj.metrics('linespace')
            except Exception:
                char_h = scaled_size
            total_h = char_h * n

        font_spec = (fname, scaled_size, weight, slant)

        # 水平位置: ボックス中央
        cx = x + w / 2

        # 垂直位置: 揃え方に応じて計算
        if v_align == 0:  # top
            start_y = y + 2
        elif v_align == 2:  # bottom
            start_y = y + h - total_h - 2
        else:  # center
            start_y = y + (h - total_h) / 2

        last_id = None
        for i, ch in enumerate(chars):
            cy = start_y + char_h * i + char_h / 2
            last_id = self._canvas.create_text(
                cx, cy, text=ch, font=font_spec, fill=color,
                anchor='center', width=0, tags=tags,
            )
        return last_id

    def _draw_multiline_text(
        self, x: float, y: float, w: float, h: float,
        text: str, fname: str, scaled_size: int,
        weight: str, slant: str,
        h_align: int, v_align: int,
        color: str, tags: tuple[str, ...],
    ) -> int:
        """複数行テキスト: \\n を尊重して描画する。"""
        font_spec = (fname, scaled_size, weight, slant)
        tx, ty, anchor = self._calc_text_pos(
            x, y, w, h, h_align, v_align,
        )

        # tkinter create_text は \n を尊重する（width=0 でも改行される）
        return self._canvas.create_text(
            tx, ty, text=text, font=font_spec, fill=color,
            anchor=anchor, width=0, tags=tags,
        )


# ── レンダラー ───────────────────────────────────────────────────────────────


class LayRenderer:
    """LayFile を指定バックエンドに描画するレンダラー。"""

    def __init__(self, lay: LayFile, backend: CanvasBackend) -> None:
        self._lay = lay
        self._b = backend

    def render_all(self) -> None:
        """ページ外枠 + 全オブジェクトを描画する。

        罫線（LINE）は最前面に描画する。
        """
        self._render_page_outline()
        # LABEL / FIELD を先に描画
        for i, obj in enumerate(self._lay.objects):
            if obj.obj_type != ObjectType.LINE:
                self.render_object(obj, index=i)
        # LINE を最前面に描画
        for i, obj in enumerate(self._lay.objects):
            if obj.obj_type == ObjectType.LINE:
                self.render_object(obj, index=i)

    def _render_page_outline(self) -> None:
        """ページ背景と外枠を描画する。"""
        px1, py1 = self._b._to_px(0, 0)
        px2, py2 = self._b._to_px(
            self._lay.page_width, self._lay.page_height,
        )
        self._b.draw_rect(
            px1, py1, px2, py2,
            fill=_PAGE_BG, outline=_PAGE_BORDER, width=1,
            tags=('page',),
        )

    def render_object(
        self, obj: LayoutObject, index: int = 0,
    ) -> None:
        """1つのオブジェクトを描画する。"""
        tag = f'obj_{index}'

        if obj.obj_type == ObjectType.LABEL:
            self._render_label(obj, tag)
        elif obj.obj_type == ObjectType.FIELD:
            self._render_field(obj, tag)
        elif obj.obj_type == ObjectType.LINE:
            self._render_line(obj, tag)

    def _render_label(self, obj: LayoutObject, tag: str) -> None:
        """LABEL オブジェクトを描画する（透明背景）。"""
        if obj.rect is None:
            return
        r = obj.rect
        px1, py1 = self._b._to_px(r.left, r.top)
        px2, py2 = self._b._to_px(r.right, r.bottom)

        # 透明なヒットエリア（選択用の不可視矩形）
        self._b.draw_rect(
            px1, py1, px2, py2,
            fill='', outline='', width=0,
            tags=(tag, 'label'),
        )

        text = obj.prefix + obj.text + obj.suffix
        if text:
            self._b.draw_text(
                px1, py1, px2 - px1, py2 - py1,
                text, obj.font.name, obj.font.size_pt,
                obj.h_align, obj.v_align, _TEXT_COLOR,
                tags=(tag, 'label_text'),
                bold=obj.font.bold, italic=obj.font.italic,
            )

    def _render_field(self, obj: LayoutObject, tag: str) -> None:
        """FIELD オブジェクトを描画する（薄い点線枠 + プレースホルダー）。"""
        if obj.rect is None:
            return
        r = obj.rect
        px1, py1 = self._b._to_px(r.left, r.top)
        px2, py2 = self._b._to_px(r.right, r.bottom)

        # 薄い背景 + 点線枠（編集時の目印）
        self._b.draw_rect(
            px1, py1, px2, py2,
            fill=_FIELD_BG, outline=_FIELD_OUTLINE, width=1,
            tags=(tag, 'field'),
        )

        name = resolve_field_name(obj.field_id)
        display = f'{obj.prefix}{{{{{name}}}}}{obj.suffix}'
        self._b.draw_text(
            px1, py1, px2 - px1, py2 - py1,
            display, obj.font.name, obj.font.size_pt,
            obj.h_align, obj.v_align, _FIELD_TEXT_COLOR,
            tags=(tag, 'field_text'),
            bold=obj.font.bold, italic=obj.font.italic,
        )

    def _render_line(self, obj: LayoutObject, tag: str) -> None:
        """LINE オブジェクトを描画する。"""
        if obj.line_start is None or obj.line_end is None:
            return
        px1, py1 = self._b._to_px(obj.line_start.x, obj.line_start.y)
        px2, py2 = self._b._to_px(obj.line_end.x, obj.line_end.y)

        self._b.draw_line(
            px1, py1, px2, py2,
            color=_LINE_COLOR, width=1,
            tags=(tag, 'line'),
        )


# ── 選択ハンドル描画 ─────────────────────────────────────────────────────────


def draw_selection_handles(
    canvas: tk.Canvas, obj: LayoutObject, backend: CanvasBackend,
) -> None:
    """選択オブジェクトの周囲にリサイズハンドルを描画する。"""
    canvas.delete('handles')

    if obj.rect is not None:
        r = obj.rect
        px1, py1 = backend._to_px(r.left, r.top)
        px2, py2 = backend._to_px(r.right, r.bottom)
    elif obj.line_start is not None and obj.line_end is not None:
        px1, py1 = backend._to_px(obj.line_start.x, obj.line_start.y)
        px2, py2 = backend._to_px(obj.line_end.x, obj.line_end.y)
    else:
        return

    # 選択枠
    canvas.create_rectangle(
        px1 - 1, py1 - 1, px2 + 1, py2 + 1,
        outline=_SELECT_OUTLINE, width=2, dash=(4, 4),
        tags=('handles',),
    )

    # 8 ハンドル (矩形) / LINE は 2 ハンドル
    hs = _HANDLE_SIZE
    if obj.rect is not None:
        cx, cy = (px1 + px2) / 2, (py1 + py2) / 2
        handle_positions = {
            'nw': (px1, py1), 'n': (cx, py1), 'ne': (px2, py1),
            'w': (px1, cy), 'e': (px2, cy),
            'sw': (px1, py2), 's': (cx, py2), 'se': (px2, py2),
        }
    else:
        handle_positions = {
            'start': (px1, py1),
            'end': (px2, py2),
        }

    for name, (hx, hy) in handle_positions.items():
        canvas.create_rectangle(
            hx - hs, hy - hs, hx + hs, hy + hs,
            fill=_HANDLE_COLOR, outline='white', width=1,
            tags=('handles', f'handle_{name}'),
        )


def clear_selection_handles(canvas: tk.Canvas) -> None:
    """選択ハンドルを削除する。"""
    canvas.delete('handles')


# ── データ差込 ───────────────────────────────────────────────────────────────


def fill_layout(
    lay: LayFile, data_row: dict, options: dict | None = None,
) -> LayFile:
    """FIELD オブジェクトを実データで置換した LayFile のコピーを返す。

    元の LayFile は変更しない。FIELD は対応するデータ値を持つ LABEL に変換される。

    Args:
        lay: 元のレイアウト
        data_row: 1 名分のデータ dict（内部論理名がキー）
        options: {'fiscal_year', 'school_name', 'teacher_name', 'name_display'} etc.

    Returns:
        差込済み LayFile（コピー）
    """
    from utils.address import build_address
    from utils.date_fmt import DATE_KEYS, format_date
    from utils.wareki import to_wareki

    options = options or {}
    mode = options.get('name_display', 'furigana')

    # name_display モード用
    _NAME_KANA_KEYS = frozenset({'氏名かな', '正式氏名かな'})
    _NAME_KANJI_KEYS = frozenset({'氏名', '正式氏名'})
    _KANJI_TO_KANA = {'氏名': '氏名かな', '正式氏名': '正式氏名かな'}

    def _resolve(key: str) -> str:
        """フィールド論理名からデータ値を解決する。"""
        original_key = key

        # 特殊キー
        if key == '年度':
            return str(options.get('fiscal_year', ''))
        if key == '年度和暦':
            fy = options.get('fiscal_year', 2025)
            return to_wareki(fy, 4, 1).replace('年', '年度')
        if key == '学校名':
            return options.get('school_name', '')
        if key == '担任名':
            return options.get('teacher_name', '')
        if key == '住所':
            return build_address(data_row)

        # name_display モード
        if mode == 'kanji' and key in _NAME_KANA_KEYS:
            return ''
        if mode == 'kana':
            if key in _NAME_KANA_KEYS:
                return ''
            if key in _NAME_KANJI_KEYS:
                key = _KANJI_TO_KANA[key]

        v = data_row.get(key, '')
        if v is None:
            return ''
        s = str(v).strip()
        if s.lower() == 'nan':
            return ''
        if original_key in DATE_KEYS:
            return format_date(s)
        return s

    filled_objects: list[LayoutObject] = []
    for obj in lay.objects:
        if obj.obj_type == ObjectType.FIELD:
            logical_name = resolve_field_name(obj.field_id)
            value = _resolve(logical_name)
            text = obj.prefix + value + obj.suffix
            filled_obj = LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=obj.rect,
                text=text,
                font=obj.font,
                h_align=obj.h_align,
                v_align=obj.v_align,
            )
            filled_objects.append(filled_obj)
        else:
            filled_objects.append(obj)

    return LayFile(
        title=lay.title,
        version=lay.version,
        page_width=lay.page_width,
        page_height=lay.page_height,
        objects=filled_objects,
    )
