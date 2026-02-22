"""PILBackend + render_layout_to_image のユニットテスト

テスト対象:
  - PILBackend の描画プリミティブ（矩形・線・テキスト）
  - render_layout_to_image のサイズ・基本動作
  - 縦書き・複数行テキスト
"""

from __future__ import annotations

import numpy as np
from PIL import Image

from core.lay_parser import (
    FontInfo,
    LayFile,
    LayoutObject,
    ObjectType,
    Rect,
    new_field,
    new_label,
    new_line,
)
from core.lay_renderer import PILBackend, render_layout_to_image

# ── PILBackend プリミティブテスト ─────────────────────────────────────────────


class TestPILBackendPrimitives:
    """PILBackend の基本描画テスト。"""

    def test_draw_rect(self):
        img = Image.new('RGB', (200, 200), (255, 255, 255))
        backend = PILBackend(img, dpi=150)
        backend.draw_rect(10, 10, 100, 100, fill='#000000', outline='#000000')

        arr = np.array(img)
        # 黒ピクセルが存在する
        assert (arr[10:100, 10:100] == 0).any()

    def test_draw_rect_skip_empty(self):
        """fill と outline が両方空のとき何も描画しない。"""
        img = Image.new('RGB', (200, 200), (255, 255, 255))
        backend = PILBackend(img, dpi=150)
        backend.draw_rect(10, 10, 100, 100, fill='', outline='')

        arr = np.array(img)
        assert (arr == 255).all()

    def test_draw_line(self):
        img = Image.new('RGB', (200, 200), (255, 255, 255))
        backend = PILBackend(img, dpi=150)
        backend.draw_line(0, 100, 200, 100, color='#000000', width=2)

        arr = np.array(img)
        # 線の位置に黒ピクセル
        assert (arr[99:102, :, :] == 0).any()

    def test_draw_text(self):
        img = Image.new('RGB', (300, 100), (255, 255, 255))
        backend = PILBackend(img, dpi=150)
        backend.draw_text(
            10, 10, 280, 80, 'テスト',
            font_name='', font_size=14.0,
        )

        arr = np.array(img)
        # テキスト描画により白でないピクセルが存在する
        assert (arr != 255).any()

    def test_draw_text_empty_string(self):
        """空文字列は何も描画しない。"""
        img = Image.new('RGB', (200, 100), (255, 255, 255))
        backend = PILBackend(img, dpi=150)
        backend.draw_text(10, 10, 180, 80, '', font_name='', font_size=14.0)

        arr = np.array(img)
        assert (arr == 255).all()


# ── render_layout_to_image テスト ─────────────────────────────────────────────


class TestRenderLayoutToImage:
    """render_layout_to_image のテスト。"""

    def test_returns_pil_image(self):
        lay = LayFile(objects=[new_label(10, 10, 200, 50, text='テスト')])
        img = render_layout_to_image(lay, dpi=150)
        assert isinstance(img, Image.Image)
        assert img.mode == 'RGB'

    def test_a4_dimensions(self):
        """A4 (840x1188) を 150DPI でレンダリングすると正しいサイズになる。"""
        lay = LayFile(page_width=840, page_height=1188)
        img = render_layout_to_image(lay, dpi=150)

        # 840 * 0.25 * 150 / 25.4 ≈ 1240
        # 1188 * 0.25 * 150 / 25.4 ≈ 1754
        assert abs(img.width - 1240) <= 2
        assert abs(img.height - 1754) <= 2

    def test_mixed_objects(self):
        """LABEL + FIELD + LINE 混在レイアウトがエラーなくレンダリングされる。"""
        lay = LayFile(objects=[
            new_label(10, 10, 200, 50, text='ラベル', font_size=12.0),
            new_field(200, 10, 400, 50, field_id=108, font_size=11.0),
            new_line(10, 60, 400, 60),
        ])
        img = render_layout_to_image(lay, dpi=100)
        assert isinstance(img, Image.Image)
        # 白一色ではない（何か描画されている）
        arr = np.array(img)
        assert (arr != 255).any()

    def test_different_dpi(self):
        """DPI を変えるとサイズが変わる。"""
        lay = LayFile(page_width=840, page_height=1188)
        img_low = render_layout_to_image(lay, dpi=72)
        img_high = render_layout_to_image(lay, dpi=300)
        assert img_high.width > img_low.width
        assert img_high.height > img_low.height


# ── テキストモードテスト ──────────────────────────────────────────────────────


class TestPILTextModes:
    """縦書き・複数行テキストの PILBackend テスト。"""

    def test_vertical_text(self):
        """狭い矩形 + 複数文字 → 縦書きが描画される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 10, 25, 200),  # w=15, h=190 → 狭い
                text='特記事項',
                font=FontInfo('', 10.0),
            ),
        ])
        img = render_layout_to_image(lay, dpi=150)
        arr = np.array(img)
        assert (arr != 255).any()

    def test_multiline_text(self):
        """改行テキストが描画される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 10, 100, 80),
                text='氏\r\n名',
                font=FontInfo('', 10.0),
            ),
        ])
        img = render_layout_to_image(lay, dpi=150)
        arr = np.array(img)
        assert (arr != 255).any()

    def test_alignment_center(self):
        """中央揃えテキストがエラーなく描画される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 10, 300, 60),
                text='中央揃え',
                font=FontInfo('', 12.0),
                h_align=1,
                v_align=1,
            ),
        ])
        img = render_layout_to_image(lay, dpi=150)
        assert isinstance(img, Image.Image)

    def test_multiline_long_line_wraps_inside_box(self):
        """複数行テキストの長い行がボックス外にはみ出さない。"""
        long_line = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' * 20
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 10, 130, 140),  # 幅を狭くして折り返し必須にする
                text=f'{long_line}\nEND',
                font=FontInfo('', 10.0),
                h_align=0,
                v_align=0,
            ),
        ])
        img = render_layout_to_image(lay, dpi=150)
        arr = np.array(img)
        # ボックスの十分右側にテキストの描画が漏れていないこと
        spill = (arr[20:120, 450:650] != 255).any()
        assert not spill

    def test_label_style_1002_10_draws_border(self):
        """style_1002=10 の LABEL は枠線を描画する。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(100, 100, 300, 200),
                text='',
                style_1002=10,
            ),
        ])
        img = render_layout_to_image(lay, dpi=150)
        arr = np.array(img)
        scale = 0.25 * 150 / 25.4

        edge_x = int(100 * scale)
        edge_y = int(150 * scale)
        inside_x = int(150 * scale)
        inside_y = int(150 * scale)

        # 枠線上は白以外（黒線）が描画される
        assert (arr[edge_y, edge_x] != 255).any()
        # 内部は塗りつぶされない
        assert (arr[inside_y, inside_x] == 255).all()

    def test_single_line_long_label_wraps_instead_of_truncating(self):
        """長い1行 LABEL は折り返して複数行で描画される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(20, 20, 220, 150),  # 高さを確保
                text='※上位のレベルに同意いただいた場合、その下位レベルにも同意したものと見なします。',
                font=FontInfo('ＭＳ 明朝', 8.0),
                h_align=0,
                v_align=0,
            ),
        ])
        img = render_layout_to_image(lay, dpi=150)
        arr = np.array(img)

        scale = 0.25 * 150 / 25.4
        x1 = int(20 * scale)
        x2 = int(220 * scale)
        y1 = int(20 * scale)
        y2 = int(150 * scale)
        crop = arr[y1:y2, x1:x2]

        gray = 0.299 * crop[:, :, 0] + 0.587 * crop[:, :, 1] + 0.114 * crop[:, :, 2]
        mask = gray < 180
        row_counts = mask.sum(axis=1)
        th = max(3, int(mask.shape[1] * 0.01))

        seg_count = 0
        in_seg = False
        for c in row_counts:
            if c >= th and not in_seg:
                in_seg = True
                seg_count += 1
            elif c < th and in_seg:
                in_seg = False

        # 1行縮小ではなく、複数行として描画されていること
        assert seg_count >= 2

    def test_multiline_v_align_changes_vertical_position(self):
        """複数行テキストで v_align により描画開始位置が変わる。"""
        rect = Rect(20, 20, 220, 260)
        top_lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=rect,
                text='LINE1\nLINE2',
                font=FontInfo('', 12.0),
                h_align=0,
                v_align=0,
            ),
        ])
        bottom_lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=rect,
                text='LINE1\nLINE2',
                font=FontInfo('', 12.0),
                h_align=0,
                v_align=2,
            ),
        ])

        img_top = render_layout_to_image(top_lay, dpi=150, for_print=True)
        img_bottom = render_layout_to_image(bottom_lay, dpi=150, for_print=True)

        scale = 0.25 * 150 / 25.4
        x1 = int(rect.left * scale)
        x2 = int(rect.right * scale)
        y1 = int(rect.top * scale)
        y2 = int(rect.bottom * scale)
        crop_top = np.array(img_top)[y1:y2, x1:x2]
        crop_bottom = np.array(img_bottom)[y1:y2, x1:x2]

        def _first_ink_row(crop: np.ndarray) -> int:
            mask = (crop != 255).any(axis=2)
            rows = np.where(mask.any(axis=1))[0]
            return int(rows[0]) if rows.size > 0 else -1

        top_row = _first_ink_row(crop_top)
        bottom_row = _first_ink_row(crop_bottom)
        assert top_row >= 0 and bottom_row >= 0
        assert bottom_row > top_row

    def test_vertical_h_align_changes_horizontal_position(self):
        """縦書きテキストで h_align により描画位置が左右に変わる。"""
        rect = Rect(20, 20, 140, 220)
        left_lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=rect,
                text='ABCD',
                font=FontInfo('', 12.0, vertical=True),
                h_align=0,
                v_align=1,
            ),
        ])
        right_lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=rect,
                text='ABCD',
                font=FontInfo('', 12.0, vertical=True),
                h_align=2,
                v_align=1,
            ),
        ])

        img_left = render_layout_to_image(left_lay, dpi=150, for_print=True)
        img_right = render_layout_to_image(right_lay, dpi=150, for_print=True)

        scale = 0.25 * 150 / 25.4
        x1 = int(rect.left * scale)
        x2 = int(rect.right * scale)
        y1 = int(rect.top * scale)
        y2 = int(rect.bottom * scale)
        crop_left = np.array(img_left)[y1:y2, x1:x2]
        crop_right = np.array(img_right)[y1:y2, x1:x2]

        def _first_ink_col(crop: np.ndarray) -> int:
            mask = (crop != 255).any(axis=2)
            cols = np.where(mask.any(axis=0))[0]
            return int(cols[0]) if cols.size > 0 else -1

        left_col = _first_ink_col(crop_left)
        right_col = _first_ink_col(crop_right)
        assert left_col >= 0 and right_col >= 0
        assert right_col > left_col


# ── エッジケーステスト ────────────────────────────────────────────────────────


class TestEdgeCases:
    """PILBackend と render_layout_to_image のエッジケーステスト。"""

    def test_zero_page_dimensions(self):
        """ページサイズ 0 でもクラッシュしない。"""
        lay = LayFile(page_width=0, page_height=0)
        img = render_layout_to_image(lay, dpi=150)
        assert img.width >= 1
        assert img.height >= 1

    def test_zero_dpi_clamped(self):
        """DPI=0 でもクラッシュしない。"""
        lay = LayFile(page_width=840, page_height=1188)
        img = render_layout_to_image(lay, dpi=0)
        assert img.width >= 1
        assert img.height >= 1

    def test_empty_layout(self):
        """オブジェクトなしでも画像が返る（ページ枠描画あり）。"""
        lay = LayFile(objects=[])
        img = render_layout_to_image(lay, dpi=72)
        assert isinstance(img, Image.Image)
        assert img.width > 0
        assert img.height > 0

    def test_bold_italic_text(self):
        """bold/italic テキストがエラーなく描画される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 10, 200, 50),
                text='太字斜体',
                font=FontInfo('', 12.0, bold=True, italic=True),
            ),
        ])
        img = render_layout_to_image(lay, dpi=100)
        arr = np.array(img)
        assert (arr != 255).any()

    def test_very_long_text(self):
        """非常に長いテキストでクラッシュしない。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 10, 100, 30),
                text='あ' * 200,
                font=FontInfo('', 10.0),
            ),
        ])
        img = render_layout_to_image(lay, dpi=72)
        assert isinstance(img, Image.Image)

    def test_field_with_prefix_suffix(self):
        """prefix/suffix 付きフィールドが描画される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.FIELD,
                rect=Rect(10, 10, 200, 50),
                field_id=108,
                font=FontInfo('', 10.0),
                prefix='〒',
                suffix='様',
            ),
        ])
        img = render_layout_to_image(lay, dpi=100)
        assert isinstance(img, Image.Image)
