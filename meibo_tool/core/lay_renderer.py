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

import os
import tkinter as tk
from abc import ABC, abstractmethod

from core.lay_parser import (
    FontInfo,
    LayFile,
    LayoutObject,
    ObjectType,
    Point,
    Rect,
    resolve_field_name,
)

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

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

# Windows フォント名 → パスマッピング（.lay で使われるフォント名に対応）
_FONT_NAME_MAP: dict[str, str] = {
    'ＭＳ 明朝': 'C:/Windows/Fonts/msmincho.ttc',
    'ＭＳ ゴシック': 'C:/Windows/Fonts/msgothic.ttc',
    'ＭＳ Ｐ明朝': 'C:/Windows/Fonts/msmincho.ttc',
    'ＭＳ Ｐゴシック': 'C:/Windows/Fonts/msgothic.ttc',
    'IPAmj明朝': 'C:/Windows/Fonts/ipamjm.ttf',
    'メイリオ': 'C:/Windows/Fonts/meiryo.ttc',
    'Yu Gothic': 'C:/Windows/Fonts/YuGothR.ttc',
    '游ゴシック': 'C:/Windows/Fonts/YuGothR.ttc',
}

# フォールバック候補（指定フォントが見つからない場合）
_FALLBACK_FONT_PATHS = [
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
        bold: bool = False, italic: bool = False,
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


# ── PIL バックエンド ──────────────────────────────────────────────────────────


class PILBackend(RenderBackend):
    """PIL (Pillow) への描画バックエンド。印刷プレビュー用。"""

    def __init__(self, image: Image.Image, dpi: int = 150) -> None:
        self._img = image
        self._draw = ImageDraw.Draw(image)
        self._dpi = dpi
        self._scale = 0.25 * dpi / 25.4  # model unit → pixels

    def _to_px(self, x: int, y: int) -> tuple[float, float]:
        """モデル座標を画像ピクセルに変換する。"""
        return x * self._scale, y * self._scale

    def draw_rect(
        self, left: float, top: float, right: float, bottom: float,
        fill: str = '', outline: str = '#000000', width: int = 1,
        **kwargs: object,
    ) -> None:
        fill_c = fill if fill else None
        outline_c = outline if outline else None
        if fill_c is None and outline_c is None:
            return
        self._draw.rectangle(
            [left, top, right, bottom],
            fill=fill_c, outline=outline_c, width=max(1, width),
        )

    def draw_line(
        self, x1: float, y1: float, x2: float, y2: float,
        color: str = '#000000', width: int = 1,
        **kwargs: object,
    ) -> None:
        self._draw.line(
            [(x1, y1), (x2, y2)], fill=color, width=max(1, width),
        )

    def draw_text(
        self, x: float, y: float, w: float, h: float,
        text: str, font_name: str, font_size: float,
        h_align: int = 0, v_align: int = 0,
        color: str = '#000000',
        bold: bool = False, italic: bool = False,
        **kwargs: object,
    ) -> None:
        if not text:
            return

        text = text.replace('\r\n', '\n').replace('\r', '\n')

        # 縦書き判定（CanvasBackend と同じロジック）
        is_vertical = (
            '\n' not in text
            and len(text) > 1
            and w < h * 0.5
            and w < 120 * self._scale
        )
        has_newline = '\n' in text

        if is_vertical:
            self._draw_vertical(x, y, w, h, text, font_size, h_align, v_align, color, font_name)
        elif has_newline:
            self._draw_multiline(x, y, w, h, text, font_size, h_align, v_align, color, font_name)
        else:
            self._draw_single_line(x, y, w, h, text, font_size, h_align, v_align, color, font_name)

    def _load_font(
        self, size_pt: float, font_name: str = '',
    ) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
        """PIL フォントをロードする。

        font_name が指定されていれば、対応するフォントファイルを優先的に使用する。
        見つからない場合はフォールバックリストから順に試す。
        """
        size_px = max(8, int(size_pt * self._dpi / 72))

        # 指定フォント名でマッチするパスを優先
        if font_name:
            path = _FONT_NAME_MAP.get(font_name)
            if path and os.path.exists(path):
                try:
                    return ImageFont.truetype(path, size=size_px)
                except (OSError, IndexError):
                    pass

        # フォールバック
        for path in _FALLBACK_FONT_PATHS:
            if os.path.exists(path):
                try:
                    return ImageFont.truetype(path, size=size_px)
                except (OSError, IndexError):
                    continue
        return ImageFont.load_default()

    def _draw_single_line(
        self, x: float, y: float, w: float, h: float,
        text: str, font_size: float,
        h_align: int, v_align: int, color: str,
        font_name: str = '',
    ) -> None:
        font = self._load_font(font_size, font_name)

        # 自動フォント縮小
        bbox = self._draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        current_size = font_size
        while tw > w * 1.05 and current_size > 4:
            current_size *= 0.9
            font = self._load_font(current_size, font_name)
            bbox = self._draw.textbbox((0, 0), text, font=font)
            tw = bbox[2] - bbox[0]

        bbox = self._draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]

        # 水平アライメント
        if h_align == 1:     # 中央
            tx = x + (w - tw) / 2
        elif h_align == 2:   # 右揃え
            tx = x + w - tw
        else:                # 左揃え
            tx = x

        # 垂直アライメント
        if v_align == 1:     # 中央
            ty = y + (h - th) / 2
        elif v_align == 2:   # 下揃え
            ty = y + h - th
        else:                # 上揃え
            ty = y

        self._draw.text((tx, ty), text, fill=color, font=font)

    def _draw_vertical(
        self, x: float, y: float, w: float, h: float,
        text: str, font_size: float,
        h_align: int, v_align: int, color: str,
        font_name: str = '',
    ) -> None:
        n = len(text)
        char_h = h / n if n > 0 else h
        char_size = min(font_size, char_h * 72 / self._dpi * 0.9)
        font = self._load_font(char_size, font_name)

        for i, ch in enumerate(text):
            bbox = self._draw.textbbox((0, 0), ch, font=font)
            cw = bbox[2] - bbox[0]
            ch_h = bbox[3] - bbox[1]
            cx = x + (w - cw) / 2
            cy = y + i * char_h + (char_h - ch_h) / 2
            self._draw.text((cx, cy), ch, fill=color, font=font)

    def _draw_multiline(
        self, x: float, y: float, w: float, h: float,
        text: str, font_size: float,
        h_align: int, v_align: int, color: str,
        font_name: str = '',
    ) -> None:
        lines = text.split('\n')
        font = self._load_font(font_size, font_name)

        # 行高さ計算
        line_h = h / len(lines) if lines else h

        for i, line in enumerate(lines):
            if not line:
                continue
            bbox = self._draw.textbbox((0, 0), line, font=font)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]

            # 水平アライメント
            if h_align == 1:
                tx = x + (w - tw) / 2
            elif h_align == 2:
                tx = x + w - tw
            else:
                tx = x

            ty = y + i * line_h + (line_h - th) / 2
            self._draw.text((tx, ty), line, fill=color, font=font)


# ── レンダラー ───────────────────────────────────────────────────────────────


class LayRenderer:
    """LayFile を指定バックエンドに描画するレンダラー。"""

    def __init__(self, lay: LayFile, backend: RenderBackend) -> None:
        self._lay = lay
        self._b = backend

    def render_all(self, *, skip_page_outline: bool = False) -> None:
        """ページ外枠 + 全オブジェクトを描画する。

        罫線（LINE）は最前面に描画する。

        Args:
            skip_page_outline: True の場合、ページ背景・外枠を描画しない（印刷用）。
        """
        if not skip_page_outline:
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


# ── PIL レンダリング ──────────────────────────────────────────────────────────


def render_layout_to_image(
    lay: LayFile, dpi: int = 150, *, for_print: bool = False,
) -> Image.Image:
    """LayFile を PIL 画像にレンダリングする。

    Args:
        lay: レンダリング対象のレイアウト
        dpi: 画像解像度 (default 150)
        for_print: True の場合、ページ外枠を描画しない（印刷用）。

    Returns:
        PIL.Image.Image (RGB)
    """
    scale = 0.25 * max(1, dpi) / 25.4
    w = max(1, int(lay.page_width * scale))
    h = max(1, int(lay.page_height * scale))
    img = Image.new('RGB', (w, h), (255, 255, 255))
    backend = PILBackend(img, dpi)
    renderer = LayRenderer(lay, backend)
    renderer.render_all(skip_page_outline=for_print)
    return img


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


# ── タイル配置 ───────────────────────────────────────────────────────────────

# A4 用紙サイズ（0.25mm 単位）
A4_WIDTH = 840     # 210mm
A4_HEIGHT = 1188   # 297mm


def calculate_page_arrangement(
    lay: LayFile,
    paper_width: int = A4_WIDTH,
    paper_height: int = A4_HEIGHT,
) -> tuple[int, int, int, float]:
    """レイアウトサイズから1ページあたりの配置数を計算する。

    レイアウトが用紙より大きい場合は、縮小して複数枚配置を試みる。

    Args:
        lay: レイアウト
        paper_width: 用紙幅（0.25mm 単位）
        paper_height: 用紙高さ（0.25mm 単位）

    Returns:
        (cols, rows, per_page, scale) タプル。
        scale は 1.0（等倍）または縮小率。
    """
    card_w = max(1, lay.page_width)
    card_h = max(1, lay.page_height)

    # 等倍で複数枚入るか試す
    cols = max(1, paper_width // card_w)
    rows = max(1, paper_height // card_h)
    if cols * rows > 1:
        return cols, rows, cols * rows, 1.0

    # カードが用紙に収まるなら等倍 1 枚
    if card_w <= paper_width and card_h <= paper_height:
        return 1, 1, 1, 1.0

    # カードが用紙より大きい — 縮小して複数枚配置を試みる
    # 候補配置を試し、縮小率 >= 25% の最良を採用
    _MIN_SCALE = 0.25
    candidates = [(2, 1), (1, 2), (2, 2), (3, 1), (1, 3), (3, 2), (2, 3)]
    for c, r in candidates:
        s = min(paper_width / (c * card_w), paper_height / (r * card_h))
        if s >= _MIN_SCALE:
            return c, r, c * r, s

    # どの配置も縮小率が足りない — 用紙に合わせて 1 枚
    scale = min(paper_width / card_w, paper_height / card_h)
    return 1, 1, 1, scale


def tile_layouts(
    layouts: list[LayFile],
    cols: int,
    rows: int,
    paper_width: int = A4_WIDTH,
    paper_height: int = A4_HEIGHT,
    scale: float = 1.0,
) -> list[LayFile]:
    """複数のレイアウトを1ページにタイル配置した LayFile のリストを返す。

    各レイアウトを用紙上のグリッドに配置し、中央揃えする。
    既に per_page == 1 かつ scale == 1.0 の場合は元のリストをそのまま返す。

    Args:
        layouts: 差込済みの個別レイアウト
        cols: 列数
        rows: 行数
        paper_width: 用紙幅（0.25mm 単位）
        paper_height: 用紙高さ（0.25mm 単位）
        scale: 縮小率（1.0 = 等倍）

    Returns:
        タイル配置された page LayFile のリスト
    """
    per_page = cols * rows
    if (per_page <= 1 and scale >= 1.0) or not layouts:
        return layouts

    cell_w = int(layouts[0].page_width * scale)
    cell_h = int(layouts[0].page_height * scale)

    # カード間ガター (0.25mm 単位, 上限 20 = 5mm)
    _MAX_GUTTER = 20
    spare_x = paper_width - cols * cell_w
    spare_y = paper_height - rows * cell_h
    gutter_x = min(_MAX_GUTTER, spare_x // cols) if cols > 1 and spare_x > 0 else 0
    gutter_y = min(_MAX_GUTTER, spare_y // rows) if rows > 1 and spare_y > 0 else 0

    # 中央揃えマージン（ガター分を含む）
    total_w = cols * cell_w + (cols - 1) * gutter_x
    total_h = rows * cell_h + (rows - 1) * gutter_y
    margin_x = (paper_width - total_w) // 2
    margin_y = (paper_height - total_h) // 2

    pages: list[LayFile] = []
    for page_start in range(0, len(layouts), per_page):
        page_items = layouts[page_start:page_start + per_page]

        page = LayFile(
            title='印刷ページ',
            page_width=paper_width,
            page_height=paper_height,
            objects=[],
        )

        for i, lay in enumerate(page_items):
            col = i % cols
            row = i // cols
            ox = margin_x + col * (cell_w + gutter_x)
            oy = margin_y + row * (cell_h + gutter_y)

            for obj in lay.objects:
                new_obj = _scale_and_offset_object(obj, scale, ox, oy)
                page.objects.append(new_obj)

        pages.append(page)

    return pages


def _scale_and_offset_object(
    obj: LayoutObject, scale: float, dx: int, dy: int,
) -> LayoutObject:
    """オブジェクトを縮小してからオフセットしたコピーを返す。"""
    if scale >= 1.0:
        return _offset_object(obj, dx, dy)

    s = scale
    new_rect = None
    if obj.rect is not None:
        new_rect = Rect(
            int(obj.rect.left * s) + dx,
            int(obj.rect.top * s) + dy,
            int(obj.rect.right * s) + dx,
            int(obj.rect.bottom * s) + dy,
        )

    new_start = None
    if obj.line_start is not None:
        new_start = Point(int(obj.line_start.x * s) + dx,
                          int(obj.line_start.y * s) + dy)

    new_end = None
    if obj.line_end is not None:
        new_end = Point(int(obj.line_end.x * s) + dx,
                        int(obj.line_end.y * s) + dy)

    # フォントサイズも縮小
    new_font = FontInfo(
        name=obj.font.name,
        size_pt=obj.font.size_pt * s,
        bold=obj.font.bold,
        italic=obj.font.italic,
    )

    return LayoutObject(
        obj_type=obj.obj_type,
        rect=new_rect,
        line_start=new_start,
        line_end=new_end,
        text=obj.text,
        field_id=obj.field_id,
        font=new_font,
        h_align=obj.h_align,
        v_align=obj.v_align,
        prefix=obj.prefix,
        suffix=obj.suffix,
    )


def _offset_object(obj: LayoutObject, dx: int, dy: int) -> LayoutObject:
    """オブジェクトの座標をオフセットしたコピーを返す。"""
    new_rect = None
    if obj.rect is not None:
        new_rect = Rect(
            obj.rect.left + dx,
            obj.rect.top + dy,
            obj.rect.right + dx,
            obj.rect.bottom + dy,
        )

    new_start = None
    if obj.line_start is not None:
        new_start = Point(obj.line_start.x + dx, obj.line_start.y + dy)

    new_end = None
    if obj.line_end is not None:
        new_end = Point(obj.line_end.x + dx, obj.line_end.y + dy)

    return LayoutObject(
        obj_type=obj.obj_type,
        rect=new_rect,
        line_start=new_start,
        line_end=new_end,
        text=obj.text,
        field_id=obj.field_id,
        font=obj.font,
        h_align=obj.h_align,
        v_align=obj.v_align,
        prefix=obj.prefix,
        suffix=obj.suffix,
    )


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
            # データフィールドは IPAmj明朝 に上書き（IVS 異体字対応）
            # 静的テキスト（LABEL）は .lay のフォントをそのまま使う
            filled_font = FontInfo(
                name='IPAmj明朝',
                size_pt=obj.font.size_pt,
                bold=obj.font.bold,
                italic=obj.font.italic,
            )
            filled_obj = LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=obj.rect,
                text=text,
                font=filled_font,
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
