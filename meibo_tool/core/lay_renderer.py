"""LayFile 描画エンジン

LayFile オブジェクトを PIL (Pillow) で統一的に描画する。
座標変換（モデル → 描画先）もここで一元管理する。

使用方法:
    from core.lay_renderer import PILBackend, LayRenderer, render_layout_to_image
    img = render_layout_to_image(lay, dpi=150)
"""

from __future__ import annotations

import os
from dataclasses import replace
from io import BytesIO

from core.lay_parser import (
    EmbeddedImage,
    FontInfo,
    LayFile,
    LayoutObject,
    ObjectType,
    PaperLayout,
    Point,
    Rect,
    TableColumn,
    resolve_field_display,
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

# テーブル色
_TABLE_HEADER_BG = '#F0F0F0'
_TABLE_DATA_BG = '#FFE0E0'
_TABLE_BORDER = '#000000'
_TABLE_HEADER_TEXT = '#000000'
_TABLE_DATA_TEXT = '#4A7AB5'

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
    """モデル座標 → Canvas ピクセル。"""
    return x * scale + offset_x, y * scale + offset_y


def canvas_to_model(
    cx: float, cy: float, scale: float,
    offset_x: float = 0, offset_y: float = 0,
) -> tuple[int, int]:
    """Canvas ピクセル → モデル座標。"""
    return round((cx - offset_x) / scale), round((cy - offset_y) / scale)


def model_to_printer(x: int, dpi: int, unit_mm: float = 0.25) -> int:
    """モデル座標 → プリンタドット。"""
    return round(x * unit_mm * dpi / 25.4)


def _normalize_text(text: str) -> str:
    """描画前の改行コードを統一する。"""
    return text.replace('\r\n', '\n').replace('\r', '\n')


def _should_render_vertical_text(
    text: str,
    w: float,
    h: float,
    scale: float,
    *,
    vertical: bool = False,
) -> bool:
    """縦書き描画にするかを判定する。"""
    if vertical:
        return True
    return (
        '\n' not in text
        and len(text) > 1
        and w < h * 0.5
        and w < 120 * scale
    )


# ── PIL バックエンド ──────────────────────────────────────────────────────────


class PILBackend:
    """PIL (Pillow) 描画バックエンド。エディター・プレビュー・印刷で共通使用。"""

    def __init__(
        self, image: Image.Image, dpi: int = 150,
        unit_mm: float = 0.25,
    ) -> None:
        self._img = image
        self._draw = ImageDraw.Draw(image)
        self._dpi = dpi
        self._unit_mm = unit_mm
        self._scale = unit_mm * dpi / 25.4  # model unit → pixels

    def _to_px(self, x: int, y: int) -> tuple[float, float]:
        """モデル座標を画像ピクセルに変換する。"""
        return x * self._scale, y * self._scale

    def draw_rect(
        self, left: float, top: float, right: float, bottom: float,
        fill: str = '', outline: str = '#000000', width: int = 1,
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
        vertical: bool = False,
        auto_wrap: bool = False,
        underline: bool = False,
        strikethrough: bool = False,
    ) -> None:
        if not text:
            return

        text = _normalize_text(text)

        # 縦書き判定: font.vertical フラグ優先、フォールバックはヒューリスティック
        is_vertical = _should_render_vertical_text(
            text, w, h, self._scale, vertical=vertical,
        )
        has_newline = '\n' in text

        if is_vertical:
            self._draw_vertical(x, y, w, h, text, font_size, h_align, v_align, color, font_name)
        elif has_newline:
            self._draw_multiline(x, y, w, h, text, font_size, h_align, v_align, color, font_name)
        else:
            self._draw_single_line(
                x, y, w, h, text, font_size, h_align, v_align, color, font_name,
                auto_wrap=auto_wrap,
                underline=underline,
                strikethrough=strikethrough,
            )

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

    @staticmethod
    def _align_x(
        x: float, w: float, content_w: float, h_align: int, pad: float = 2.0,
    ) -> float:
        """横方向の揃え位置を返す。"""
        if h_align == 1:
            return x + (w - content_w) / 2
        if h_align == 2:
            return x + w - content_w - pad
        return x + pad

    @staticmethod
    def _align_y(
        y: float, h: float, content_h: float, v_align: int, pad: float = 2.0,
    ) -> float:
        """縦方向の揃え位置を返す。"""
        if v_align == 1:
            return y + (h - content_h) / 2
        if v_align == 2:
            return y + h - content_h - pad
        return y + pad

    def _draw_single_line(
        self, x: float, y: float, w: float, h: float,
        text: str, font_size: float,
        h_align: int, v_align: int, color: str,
        font_name: str = '',
        auto_wrap: bool = False,
        underline: bool = False,
        strikethrough: bool = False,
    ) -> None:
        font = self._load_font(font_size, font_name)
        avail_w = max(1.0, w - 4.0)

        # auto_wrap: テキストが幅を超え、高さに余裕があれば複数行に折り返す
        bbox = self._draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        if auto_wrap and tw > avail_w * 1.05 and h >= (bbox[3] - bbox[1]) * 1.8:
            self._draw_multiline(
                x, y, w, h, text, font_size, h_align, v_align, color, font_name,
            )
            return

        bbox = self._draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]

        tx = self._align_x(x, w, tw, h_align)
        ty = self._align_y(y, h, th, v_align)

        self._draw.text((tx, ty), text, fill=color, font=font)

        # 下線
        if underline:
            uy = ty + th + 1
            self._draw.line([(tx, uy), (tx + tw, uy)], fill=color, width=1)
        # 取り消し線
        if strikethrough:
            sy = ty + th // 2
            self._draw.line([(tx, sy), (tx + tw, sy)], fill=color, width=1)

    def _draw_vertical(
        self, x: float, y: float, w: float, h: float,
        text: str, font_size: float,
        h_align: int, v_align: int, color: str,
        font_name: str = '',
    ) -> None:
        chars = list(text)
        n = len(chars)
        if n == 0:
            return

        box_h = max(h - 4.0, 1.0)
        char_h = box_h / n
        char_size = min(font_size, char_h * 72 / self._dpi * 0.9)
        font = self._load_font(char_size, font_name)
        total_h = char_h * n
        start_y = self._align_y(y, h, total_h, v_align)

        for i, ch in enumerate(chars):
            bbox = self._draw.textbbox((0, 0), ch, font=font)
            cw = bbox[2] - bbox[0]
            ch_h = bbox[3] - bbox[1]
            tx = self._align_x(x, w, cw, h_align)
            ty = start_y + i * char_h + (char_h - ch_h) / 2
            self._draw.text((tx, ty), ch, fill=color, font=font)

    def _draw_multiline(
        self, x: float, y: float, w: float, h: float,
        text: str, font_size: float,
        h_align: int, v_align: int, color: str,
        font_name: str = '',
    ) -> None:
        lines = text.split('\n')
        font = self._load_font(font_size, font_name)
        max_w = max(1.0, w - 4.0)

        wrapped_lines: list[str] = []
        for line in lines:
            wrapped_lines.extend(self._wrap_line(line, font, max_w))
        lines = wrapped_lines or ['']

        if not lines:
            return
        line_bbox = self._draw.textbbox((0, 0), 'Ag', font=font)
        natural_line_h = max(1.0, float(line_bbox[3] - line_bbox[1]))
        line_h = natural_line_h
        box_h = max(h - 4.0, 1.0)
        if line_h * len(lines) > box_h:
            line_h = box_h / len(lines)
        total_h = line_h * len(lines)
        start_y = self._align_y(y, h, total_h, v_align)

        for i, line in enumerate(lines):
            if not line:
                continue
            bbox = self._draw.textbbox((0, 0), line, font=font)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]

            tx = self._align_x(x, w, tw, h_align)
            ty = start_y + i * line_h + (line_h - th) / 2
            self._draw.text((tx, ty), line, fill=color, font=font)

    def _wrap_line(
        self,
        line: str,
        font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
        max_w: float,
    ) -> list[str]:
        """1行テキストを描画幅で折り返す（空白の有無に依存しない）。"""
        if not line:
            return ['']

        wrapped: list[str] = []
        current = ''
        for ch in line:
            candidate = current + ch
            bbox = self._draw.textbbox((0, 0), candidate, font=font)
            cand_w = bbox[2] - bbox[0]
            if current and cand_w > max_w:
                wrapped.append(current)
                current = ch
            else:
                current = candidate
        if current:
            wrapped.append(current)
        return wrapped or ['']

    def draw_image(
        self, left: float, top: float, right: float, bottom: float,
        image_data: bytes,
    ) -> None:
        """埋め込み画像を PIL 画像に描画する。"""
        try:
            img = Image.open(BytesIO(image_data))
            w = max(1, int(right - left))
            h = max(1, int(bottom - top))
            # 透過画像（パレット/LA 等も含む）を RGBA に統一して合成
            img = img.convert('RGBA')
            img = img.resize((w, h), Image.Resampling.LANCZOS)
            self._img.paste(img, (int(left), int(top)), img)
        except Exception:
            pass


# ── レンダラー ───────────────────────────────────────────────────────────────


class LayRenderer:
    """LayFile を指定バックエンドに描画するレンダラー。"""

    def __init__(
        self, lay: LayFile, backend: PILBackend,
        layout_registry: dict[str, LayFile] | None = None,
        editor_mode: bool = True,
    ) -> None:
        self._lay = lay
        self._b = backend
        self._registry = layout_registry or {}
        self._editor_mode = editor_mode

    def render_all(self, *, skip_page_outline: bool = False) -> None:
        """ページ外枠 + 全オブジェクトを描画する。

        描画順: TABLE → LABEL/FIELD → LINE（最前面）

        Args:
            skip_page_outline: True の場合、ページ背景・外枠を描画しない（印刷用）。
        """
        if not skip_page_outline:
            self._render_page_outline()
        # TABLE を最背面に描画
        for i, obj in enumerate(self._lay.objects):
            if obj.obj_type == ObjectType.TABLE:
                self.render_object(obj, index=i)
        # MEIBO を描画（参照先レイアウトを展開）
        for i, obj in enumerate(self._lay.objects):
            if obj.obj_type == ObjectType.MEIBO:
                self.render_object(obj, index=i)
        # LABEL / FIELD / IMAGE を描画
        for i, obj in enumerate(self._lay.objects):
            if obj.obj_type not in (
                ObjectType.LINE, ObjectType.TABLE, ObjectType.MEIBO,
            ):
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
        )

    def render_object(
        self, obj: LayoutObject, index: int = 0,
    ) -> None:
        """1つのオブジェクトを描画する。"""
        if obj.obj_type == ObjectType.LABEL:
            self._render_label(obj)
        elif obj.obj_type == ObjectType.FIELD:
            self._render_field(obj)
        elif obj.obj_type == ObjectType.GROUP:
            self._render_group(obj)
        elif obj.obj_type == ObjectType.LINE:
            self._render_line(obj)
        elif obj.obj_type == ObjectType.TABLE:
            self._render_table(obj)
        elif obj.obj_type == ObjectType.MEIBO:
            self._render_meibo(obj)
        elif obj.obj_type == ObjectType.IMAGE:
            self._render_image(obj)

    def _render_group(self, obj: LayoutObject) -> None:
        """GROUP オブジェクトを外枠として描画する。"""
        if obj.rect is None:
            return
        r = obj.rect
        px1, py1 = self._b._to_px(r.left, r.top)
        px2, py2 = self._b._to_px(r.right, r.bottom)

        # style_1002: 8 → 太枠(2px)、2 → 標準枠(1px)、その他 → 1px
        group_w = 2 if obj.style_1002 == 8 else 1

        self._b.draw_rect(
            px1, py1, px2, py2,
            fill='', outline=_LINE_COLOR, width=group_w,
        )

    def _render_label(self, obj: LayoutObject) -> None:
        """LABEL オブジェクトを描画する（透明背景）。"""
        if obj.rect is None:
            return
        r = obj.rect
        px1, py1 = self._b._to_px(r.left, r.top)
        px2, py2 = self._b._to_px(r.right, r.bottom)

        # 一部レイアウトでは LABEL の style_1002=10 が枠付きテキストボックスを表す。
        label_outline = _LINE_COLOR if obj.style_1002 == 10 else ''
        label_width = 1 if obj.style_1002 == 10 else 0

        # 枠線描画（style_1002=10 の場合のみ）
        self._b.draw_rect(
            px1, py1, px2, py2,
            fill='', outline=label_outline, width=label_width,
        )

        text = obj.prefix + obj.text + obj.suffix
        if text:
            self._b.draw_text(
                px1, py1, px2 - px1, py2 - py1,
                text, obj.font.name, obj.font.size_pt,
                obj.h_align, obj.v_align, _TEXT_COLOR,
                bold=obj.font.bold, italic=obj.font.italic,
                vertical=obj.font.vertical,
                auto_wrap=True,
                underline=obj.font.underline,
                strikethrough=obj.font.strikethrough,
            )

    def _render_field(self, obj: LayoutObject) -> None:
        """FIELD オブジェクトを描画する。

        editor_mode=True: 薄い背景 + 点線枠 + フィールド名表示
        editor_mode=False: 背景なし + ○ 表示（プレビュー/印刷用）
        """
        if obj.rect is None:
            return
        r = obj.rect
        px1, py1 = self._b._to_px(r.left, r.top)
        px2, py2 = self._b._to_px(r.right, r.bottom)

        if self._editor_mode:
            # エディター: 薄い背景 + 点線枠 + フィールド名
            self._b.draw_rect(
                px1, py1, px2, py2,
                fill=_FIELD_BG, outline=_FIELD_OUTLINE, width=1,
            )
            name = resolve_field_display(obj.field_id)
            display = f'{obj.prefix}{name}{obj.suffix}'
            color = _FIELD_TEXT_COLOR
        else:
            # プレビュー: 背景なし + ○○○ 表示
            display = f'{obj.prefix}○○○{obj.suffix}'
            color = _TEXT_COLOR

        self._b.draw_text(
            px1, py1, px2 - px1, py2 - py1,
            display, obj.font.name, obj.font.size_pt,
            obj.h_align, obj.v_align, color,
            bold=obj.font.bold, italic=obj.font.italic,
            vertical=obj.font.vertical,
            underline=obj.font.underline,
            strikethrough=obj.font.strikethrough,
        )

    def _render_line(self, obj: LayoutObject) -> None:
        """LINE オブジェクトを描画する。"""
        if obj.line_start is None or obj.line_end is None:
            return
        px1, py1 = self._b._to_px(obj.line_start.x, obj.line_start.y)
        px2, py2 = self._b._to_px(obj.line_end.x, obj.line_end.y)

        # style_1001: 正の値 → その値を線幅として使用、-1/0/None → デフォルト(1)
        line_w = 1
        if obj.style_1001 is not None and obj.style_1001 > 0:
            line_w = obj.style_1001

        self._b.draw_line(
            px1, py1, px2, py2,
            color=_LINE_COLOR, width=line_w,
        )

    def _render_table(self, obj: LayoutObject) -> None:
        """TABLE オブジェクトを描画する。

        テーブル rect 内をカラム幅の比率で分割し、
        ヘッダー行（灰色背景）+ データ行（ピンク背景）を描画する。
        """
        if obj.rect is None or not obj.table_columns:
            return

        r = obj.rect
        px1, py1 = self._b._to_px(r.left, r.top)
        px2, py2 = self._b._to_px(r.right, r.bottom)
        table_w = px2 - px1
        table_h = py2 - py1

        cols = obj.table_columns
        total_width = sum(c.width for c in cols) or 1

        # フォントサイズ（パーサーで解析済み）
        font_size = obj.font.size_pt if obj.font.size_pt > 0 else 9.0

        # 行数: パーサーで取得した明示的行数を使用
        # 0x0BBB タグに格納された行数（ヘッダー行を含む場合と含まない場合がある）
        if obj.table_row_count > 0:
            n_data_rows = obj.table_row_count
            data_row_h = table_h / (n_data_rows + 1)
            header_h = data_row_h
        else:
            # フォールバック: フォントサイズベースのヒューリスティック
            header_h = min(table_h * 0.08, font_size * 2.5)
            if header_h < 10:
                header_h = min(table_h * 0.15, 20)
            data_row_h = header_h
            n_data_rows = max(1, int((table_h - header_h) / data_row_h))

        # 全データ行に薄いピンク背景を描画（エディターモードのみ）
        if self._editor_mode:
            for row_i in range(n_data_rows):
                data_y = py1 + header_h + row_i * data_row_h
                if data_y + data_row_h > py2:
                    break
                self._b.draw_rect(
                    px1, data_y, px2,
                    min(data_y + data_row_h, py2),
                    fill=_TABLE_DATA_BG, outline='', width=0,
                )

        # 外枠
        self._b.draw_rect(
            px1, py1, px2, py2,
            fill='', outline=_TABLE_BORDER, width=1,
        )

        # ヘッダー行背景
        self._b.draw_rect(
            px1, py1, px2, py1 + header_h,
            fill=_TABLE_HEADER_BG, outline=_TABLE_BORDER, width=1,
        )

        # カラムの描画
        col_x = px1
        for col in cols:
            col_w = table_w * col.width / total_width

            # ヘッダーテキスト
            if col.header:
                self._b.draw_text(
                    col_x, py1, col_w, header_h,
                    col.header, '', font_size,
                    h_align=1, v_align=1,
                    color=_TABLE_HEADER_TEXT,
                )

            # 縦罫線
            if col_x > px1:
                self._b.draw_line(
                    col_x, py1, col_x, py2,
                    color=_TABLE_BORDER, width=1,
                )

            # 全データ行にプレースホルダーを表示
            if self._editor_mode:
                display_name = resolve_field_display(col.field_id)
                data_color = _TABLE_DATA_TEXT
            else:
                display_name = '○○○'
                data_color = _TEXT_COLOR
            for row_i in range(n_data_rows):
                data_y = py1 + header_h + row_i * data_row_h
                if data_y + data_row_h > py2:
                    break
                self._b.draw_text(
                    col_x, data_y, col_w, data_row_h,
                    display_name,
                    '', font_size * 0.9,
                    h_align=col.h_align, v_align=1,
                    color=data_color,
                )

            col_x += col_w

        # データ行の横罫線
        for i in range(n_data_rows + 1):
            line_y = py1 + header_h + i * data_row_h
            if line_y <= py2:
                self._b.draw_line(
                    px1, line_y, px2, line_y,
                    color=_TABLE_BORDER, width=1,
                )

    def _render_meibo(self, obj: LayoutObject) -> None:
        """MEIBO オブジェクトを描画する（参照先レイアウトを繰り返し配置）。"""
        if obj.meibo is None or not self._registry:
            return
        meibo = obj.meibo
        ref_lay = self._registry.get(meibo.ref_name)
        if ref_lay is None:
            return
        for i in range(meibo.row_count):
            if meibo.direction == 0:  # 縦並び (vertical)
                dx = meibo.origin_x
                dy = meibo.origin_y + i * meibo.cell_height
            else:  # 横並び (horizontal)
                dx = meibo.origin_x + i * meibo.cell_width
                dy = meibo.origin_y
            for ref_obj in ref_lay.objects:
                offset_obj = _offset_object(ref_obj, dx, dy)
                if offset_obj.obj_type == ObjectType.LABEL:
                    self._render_label(offset_obj)
                elif offset_obj.obj_type == ObjectType.FIELD:
                    self._render_field(offset_obj)
                elif offset_obj.obj_type == ObjectType.GROUP:
                    self._render_group(offset_obj)
                elif offset_obj.obj_type == ObjectType.LINE:
                    self._render_line(offset_obj)
                elif offset_obj.obj_type == ObjectType.TABLE:
                    self._render_table(offset_obj)
                elif offset_obj.obj_type == ObjectType.IMAGE:
                    self._render_image(offset_obj)

    def _render_image(self, obj: LayoutObject) -> None:
        """IMAGE オブジェクトを描画する（埋め込み PNG 画像）。"""
        if obj.image is None or not obj.image.image_data:
            return
        if obj.rect is None:
            return
        r = obj.rect
        px1, py1 = self._b._to_px(r.left, r.top)
        px2, py2 = self._b._to_px(r.right, r.bottom)
        self._b.draw_image(
            px1, py1, px2, py2,
            obj.image.image_data,
        )


# ── PIL レンダリング ──────────────────────────────────────────────────────────


def render_layout_to_image(
    lay: LayFile, dpi: int = 150, *, for_print: bool = False,
    layout_registry: dict[str, LayFile] | None = None,
    editor_mode: bool = False,
) -> Image.Image:
    """LayFile を PIL 画像にレンダリングする。

    PaperLayout が設定されていれば unit_mm を使い、
    なければ旧形式の 0.25mm/unit を使用する。

    Args:
        lay: レンダリング対象のレイアウト
        dpi: 画像解像度 (default 150)
        for_print: True の場合、ページ外枠を描画しない（印刷用）。
        layout_registry: MEIBO 参照解決用の {名前: LayFile} dict。
        editor_mode: True の場合、フィールドに背景色 + 名前表示（エディター用）。
            False の場合、背景なし + ○ 表示（プレビュー/印刷用）。

    Returns:
        PIL.Image.Image (RGB)
    """
    unit_mm = lay.paper.unit_mm if lay.paper else 0.25
    scale = unit_mm * max(1, dpi) / 25.4
    w = max(1, int(lay.page_width * scale))
    h = max(1, int(lay.page_height * scale))
    img = Image.new('RGB', (w, h), (255, 255, 255))
    backend = PILBackend(img, dpi, unit_mm=unit_mm)
    renderer = LayRenderer(
        lay, backend, layout_registry=layout_registry,
        editor_mode=editor_mode,
    )
    renderer.render_all(skip_page_outline=for_print)
    return img


# ── タイル配置 ───────────────────────────────────────────────────────────────

# A4 用紙サイズ（0.25mm/unit の旧形式でのデフォルト値）
A4_WIDTH = 840     # 210mm / 0.25
A4_HEIGHT = 1188   # 297mm / 0.25


def get_page_arrangement(
    lay: LayFile,
) -> tuple[int, int, int, float]:
    """LayFile の PaperLayout から配置情報を返す。

    PaperLayout があればそのまま cols/rows を使用する。
    なければ calculate_page_arrangement() にフォールバック。

    Returns:
        (cols, rows, per_page, scale) タプル。
    """
    p = lay.paper
    if p is None:
        return calculate_page_arrangement(lay)

    if p.mode == 0:
        # 全面モード: 1 アイテム/ページ
        return 1, 1, 1, 1.0
    else:
        # ラベルモード: PaperLayout の cols×rows をそのまま使う
        cols = max(1, p.cols)
        rows = max(1, p.rows)
        return cols, rows, cols * rows, 1.0


def calculate_page_arrangement(
    lay: LayFile,
    paper_width: int = A4_WIDTH,
    paper_height: int = A4_HEIGHT,
) -> tuple[int, int, int, float]:
    """レイアウトサイズから1ページあたりの配置数を計算する。

    レイアウトが用紙より大きい場合は、縮小して複数枚配置を試みる。

    Args:
        lay: レイアウト
        paper_width: 用紙幅（モデル座標単位）
        paper_height: 用紙高さ（モデル座標単位）

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


# 用紙サイズ (mm) — PaperLayout の paper_size から用紙寸法を得るためのテーブル
_TILE_PAPER_MM: dict[str, tuple[int, int]] = {
    'A3': (297, 420),
    'A4': (210, 297),
    'A5': (148, 210),
    'B4': (257, 364),
    'B5': (182, 257),
    'はがき': (100, 148),
}


def tile_layouts(
    layouts: list[LayFile],
    cols: int,
    rows: int,
    paper_width: int = A4_WIDTH,
    paper_height: int = A4_HEIGHT,
    scale: float = 1.0,
    paper: PaperLayout | None = None,
) -> list[LayFile]:
    """複数のレイアウトを1ページにタイル配置した LayFile のリストを返す。

    各レイアウトを用紙上のグリッドに配置し、中央揃えする。
    既に per_page == 1 かつ scale == 1.0 の場合は元のリストをそのまま返す。

    PaperLayout が指定された場合、用紙サイズ・余白・間隔を PaperLayout から取得する。

    Args:
        layouts: 差込済みの個別レイアウト
        cols: 列数
        rows: 行数
        paper_width: 用紙幅（モデル座標単位、paper 未指定時のフォールバック）
        paper_height: 用紙高さ（モデル座標単位、paper 未指定時のフォールバック）
        scale: 縮小率（1.0 = 等倍）
        paper: PaperLayout（ラベルモードの場合、余白・間隔を使用）

    Returns:
        タイル配置された page LayFile のリスト
    """
    per_page = cols * rows
    if (per_page <= 1 and scale >= 1.0) or not layouts:
        return layouts

    cell_w = int(layouts[0].page_width * scale)
    cell_h = int(layouts[0].page_height * scale)

    # PaperLayout からラベル配置情報を取得
    if paper is not None and paper.mode == 1:
        unit_mm = paper.unit_mm or 0.1
        # 用紙サイズをレイアウト座標単位に変換
        short_mm, long_mm = _TILE_PAPER_MM.get(
            paper.paper_size, (210, 297),
        )
        if paper.orientation == 'landscape':
            pw_mm, ph_mm = long_mm, short_mm
        else:
            pw_mm, ph_mm = short_mm, long_mm
        paper_width = int(pw_mm / unit_mm)
        paper_height = int(ph_mm / unit_mm)
        margin_x = int(paper.margin_left_mm / unit_mm)
        margin_y = int(paper.margin_top_mm / unit_mm)
        gutter_x = int(paper.spacing_h_mm / unit_mm)
        gutter_y = int(paper.spacing_v_mm / unit_mm)
    else:
        # フォールバック: 中央揃え
        _MAX_GUTTER = 20
        spare_x = paper_width - cols * cell_w
        spare_y = paper_height - rows * cell_h
        gutter_x = min(_MAX_GUTTER, spare_x // cols) if cols > 1 and spare_x > 0 else 0
        gutter_y = min(_MAX_GUTTER, spare_y // rows) if rows > 1 and spare_y > 0 else 0
        total_w = cols * cell_w + (cols - 1) * gutter_x
        total_h = rows * cell_h + (rows - 1) * gutter_y
        margin_x = (paper_width - total_w) // 2
        margin_y = (paper_height - total_h) // 2

    # ページ用の PaperLayout（unit_mm を保持するために mode=0 で作成）
    page_paper = None
    if paper is not None:
        page_paper = PaperLayout(
            mode=0,
            unit_mm=paper.unit_mm,
            paper_size=paper.paper_size,
            orientation=paper.orientation,
        )

    pages: list[LayFile] = []
    for page_start in range(0, len(layouts), per_page):
        page_items = layouts[page_start:page_start + per_page]

        page = LayFile(
            title='印刷ページ',
            page_width=paper_width,
            page_height=paper_height,
            objects=[],
            paper=page_paper,
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


def _clone_layout_object(obj: LayoutObject, **changes) -> LayoutObject:
    """LayoutObject のコピーを作り、必要な項目だけ差し替える。"""
    base = replace(
        obj,
        table_columns=list(obj.table_columns),
        raw_tags=list(obj.raw_tags),
    )
    for key, val in changes.items():
        setattr(base, key, val)
    return base


def _clone_layfile(lay: LayFile, **changes) -> LayFile:
    """LayFile のコピーを作り、必要な項目だけ差し替える。"""
    base = replace(
        lay,
        objects=list(lay.objects),
        raw_tags=list(lay.raw_tags),
    )
    for key, val in changes.items():
        setattr(base, key, val)
    return base


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
        vertical=obj.font.vertical,
    )

    return _clone_layout_object(
        obj,
        rect=new_rect,
        line_start=new_start,
        line_end=new_end,
        font=new_font,
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

    return _clone_layout_object(
        obj,
        rect=new_rect,
        line_start=new_start,
        line_end=new_end,
    )


# ── 外字検出 ─────────────────────────────────────────────────────────────────


def _contains_gaiji(text: str) -> bool:
    """テキストに外字（IVS 異体字・CJK拡張漢字）が含まれるか判定する。

    IVS (Ideographic Variation Sequence) の異体字セレクタ (U+E0100〜U+E01EF) や
    CJK統合漢字拡張B以降 (U+20000〜) の文字は IPAmj明朝 でないと正しく表示できない。
    """
    for ch in text:
        cp = ord(ch)
        # IVS 異体字セレクタ (Variation Selectors Supplement)
        if 0xE0100 <= cp <= 0xE01EF:
            return True
        # CJK統合漢字拡張B〜I (U+20000〜U+3134F)
        if cp >= 0x20000:
            return True
    return False


# ── データ差込 ───────────────────────────────────────────────────────────────


def _resolve_field_value(
    key: str, data_row: dict, options: dict,
) -> str:
    """フィールド論理名からデータ値を解決する。

    fill_layout / fill_meibo_layout 共通のデータ解決ロジック。
    """
    from utils.address import build_address, build_guardian_address
    from utils.date_fmt import DATE_KEYS, format_date
    from utils.wareki import to_wareki

    original_key = key
    mode = options.get('name_display', 'furigana')

    _NAME_KANA_KEYS = frozenset({'氏名かな', '正式氏名かな'})
    _NAME_KANJI_KEYS = frozenset({'氏名', '正式氏名'})
    _KANJI_TO_KANA = {'氏名': '氏名かな', '正式氏名': '正式氏名かな'}

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
    if key == '保護者住所':
        return build_guardian_address(data_row)
    if key == 'ページ番号':
        return str(options.get('page_number', ''))
    if key == '人数合計':
        return str(options.get('total_count', ''))

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


_PHOTO_FIELD_ID = 400


def _fill_photo_field(
    obj: LayoutObject, data_row: dict, options: dict,
) -> LayoutObject:
    """写真フィールドを IMAGE オブジェクトに変換する。"""
    empty_label = _clone_layout_object(
        obj,
        obj_type=ObjectType.LABEL,
        text='',
        field_id=0,
        image=None,
    )
    photo_map = options.get('_photo_map')
    if not photo_map:
        return empty_label

    from core.photo_manager import load_photo_bytes, match_photo_to_student

    photo_path = match_photo_to_student(data_row, photo_map)
    if photo_path is None:
        return empty_label

    # フィールドの rect からアスペクト比を計算
    target_rect = None
    if obj.rect is not None:
        w = abs(obj.rect.right - obj.rect.left)
        h = abs(obj.rect.bottom - obj.rect.top)
        if w > 0 and h > 0:
            target_rect = (w, h)

    image_data = load_photo_bytes(photo_path, target_rect=target_rect)
    if image_data is None:
        return empty_label

    return _clone_layout_object(
        obj,
        obj_type=ObjectType.IMAGE,
        field_id=0,
        image=EmbeddedImage(
            rect=(
                obj.rect.left, obj.rect.top, obj.rect.right, obj.rect.bottom,
            ) if obj.rect else (0, 0, 0, 0),
            image_data=image_data,
            original_path=photo_path,
        ),
    )


def _fill_field_object(
    obj: LayoutObject, data_row: dict, options: dict,
) -> LayoutObject:
    """FIELD オブジェクトをデータで埋めた LABEL に変換する。
    写真フィールド (field_id=400) は IMAGE に変換する。
    """
    if obj.field_id == _PHOTO_FIELD_ID:
        return _fill_photo_field(obj, data_row, options)
    logical_name = resolve_field_name(obj.field_id)
    value = _resolve_field_value(logical_name, data_row, options)
    text = obj.prefix + value + obj.suffix
    if _contains_gaiji(text):
        filled_font = FontInfo(
            name='IPAmj明朝',
            size_pt=obj.font.size_pt,
            bold=obj.font.bold,
            italic=obj.font.italic,
            vertical=obj.font.vertical,
        )
    else:
        filled_font = obj.font
    return _clone_layout_object(
        obj,
        obj_type=ObjectType.LABEL,
        text=text,
        font=filled_font,
        field_id=0,
        image=None,
    )


def _fill_table_object(
    obj: LayoutObject, data_row: dict, options: dict,
) -> LayoutObject:
    """TABLE オブジェクトの各カラムをデータで埋める。"""
    filled_cols = []
    for col in obj.table_columns:
        logical_name = resolve_field_name(col.field_id)
        value = _resolve_field_value(logical_name, data_row, options)
        filled_cols.append(TableColumn(
            field_id=col.field_id,
            width=col.width,
            h_align=col.h_align,
            header=value if value else col.header,
        ))
    return _clone_layout_object(obj, table_columns=filled_cols)


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
    opts = options or {}
    filled_objects: list[LayoutObject] = []
    for obj in lay.objects:
        if obj.obj_type == ObjectType.FIELD:
            filled_objects.append(_fill_field_object(obj, data_row, opts))
        elif obj.obj_type == ObjectType.TABLE:
            filled_objects.append(_fill_table_object(obj, data_row, opts))
        else:
            filled_objects.append(obj)

    return _clone_layfile(lay, objects=filled_objects)


def _meibo_page_capacity(lay: LayFile) -> int:
    """MEIBO レイアウトの 1 ページあたりの生徒収容数を返す。

    MEIBO がなければ 0 を返す。
    """
    cap = 0
    for obj in lay.objects:
        if obj.obj_type == ObjectType.MEIBO and obj.meibo:
            cap = max(cap, obj.meibo.data_start_index + obj.meibo.row_count)
    return cap


def has_meibo(lay: LayFile) -> bool:
    """レイアウトに MEIBO オブジェクトが含まれるか判定する。"""
    return any(o.obj_type == ObjectType.MEIBO and o.meibo for o in lay.objects)


def fill_meibo_layout(
    lay: LayFile,
    data_rows: list[dict],
    options: dict | None = None,
    layout_registry: dict[str, LayFile] | None = None,
) -> list[LayFile]:
    """MEIBO 付きレイアウトに複数名分のデータを差し込む。

    MEIBO オブジェクトを展開し、各セルに生徒データを差込んだ
    LABEL/LINE オブジェクトに変換する。ページ容量を超える生徒は
    次ページに配置される。

    Args:
        lay: MEIBO オブジェクトを含むレイアウト
        data_rows: 全生徒データ dict のリスト
        options: fill_layout と同じオプション
        layout_registry: MEIBO ref_name 解決用レジストリ

    Returns:
        ページごとの差込済み LayFile リスト
    """
    import logging

    logger = logging.getLogger(__name__)
    opts = options or {}
    registry = layout_registry or {}

    capacity = _meibo_page_capacity(lay)
    if capacity <= 0:
        # MEIBO なし — 通常の fill_layout にフォールバック
        if data_rows:
            return [fill_layout(lay, data_rows[0], opts)]
        return []

    pages: list[LayFile] = []
    total = len(data_rows)

    for page_start in range(0, max(total, 1), capacity):
        page_data = data_rows[page_start:page_start + capacity]
        # 共通データ（年度、学校名等）は先頭行から取得
        common_row = page_data[0] if page_data else {}
        page_opts = {**opts, 'page_number': (page_start // capacity) + 1}

        filled_objects: list[LayoutObject] = []

        for obj in lay.objects:
            if obj.obj_type == ObjectType.FIELD:
                filled_objects.append(
                    _fill_field_object(obj, common_row, page_opts),
                )
            elif obj.obj_type == ObjectType.TABLE:
                filled_objects.append(
                    _fill_table_object(obj, common_row, page_opts),
                )
            elif obj.obj_type == ObjectType.MEIBO and obj.meibo:
                meibo = obj.meibo
                ref_lay = registry.get(meibo.ref_name)
                if ref_lay is None:
                    logger.warning(
                        'MEIBO ref_name 未解決: %s', meibo.ref_name,
                    )
                    continue

                # MEIBO セルを展開
                for i in range(meibo.row_count):
                    student_idx = meibo.data_start_index + i
                    if meibo.direction == 0:  # 縦並び (vertical)
                        dx = meibo.origin_x
                        dy = meibo.origin_y + i * meibo.cell_height
                    else:  # 横並び (horizontal)
                        dx = meibo.origin_x + i * meibo.cell_width
                        dy = meibo.origin_y

                    if student_idx < len(page_data):
                        student_row = page_data[student_idx]
                        # パーツレイアウトの各オブジェクトをデータ差込
                        for ref_obj in ref_lay.objects:
                            offset_obj = _offset_object(ref_obj, dx, dy)
                            if offset_obj.obj_type == ObjectType.FIELD:
                                filled_objects.append(
                                    _fill_field_object(
                                        offset_obj, student_row, page_opts,
                                    ),
                                )
                            elif offset_obj.obj_type == ObjectType.TABLE:
                                filled_objects.append(
                                    _fill_table_object(
                                        offset_obj, student_row, page_opts,
                                    ),
                                )
                            else:
                                filled_objects.append(offset_obj)
                    else:
                        # 生徒データなし — 罫線やラベル（枠線）のみ配置
                        for ref_obj in ref_lay.objects:
                            if ref_obj.obj_type in (
                                ObjectType.LINE, ObjectType.LABEL, ObjectType.GROUP,
                            ):
                                filled_objects.append(
                                    _offset_object(ref_obj, dx, dy),
                                )
            else:
                filled_objects.append(obj)

        pages.append(_clone_layfile(lay, objects=filled_objects))

    return pages
