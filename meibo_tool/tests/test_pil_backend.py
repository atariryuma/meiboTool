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

    def test_draw_text_kwargs_absorb_tags(self):
        """tags パラメータがエラーにならないことを確認。"""
        img = Image.new('RGB', (200, 100), (255, 255, 255))
        backend = PILBackend(img, dpi=150)
        # tags= は Canvas 専用だが、PILBackend は **kwargs で吸収する
        backend.draw_text(
            10, 10, 180, 80, 'ABC',
            font_name='', font_size=14.0,
            tags=('obj_0', 'label_text'),
        )
        # エラーが出なければ OK


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
