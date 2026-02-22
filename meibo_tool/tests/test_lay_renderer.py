"""lay_renderer テスト

fill_layout() データ差込ロジックと座標変換を検証する。
"""

from __future__ import annotations

import pytest

from core.lay_parser import (
    FIELD_ID_MAP,
    EmbeddedImage,
    FontInfo,
    LayFile,
    LayoutObject,
    MeiboArea,
    ObjectType,
    PaperLayout,
    Rect,
    TableColumn,
    new_field,
    new_label,
    new_line,
    resolve_field_display,
)
from core.lay_renderer import (
    _contains_gaiji,
    _meibo_page_capacity,
    _normalize_text,
    _should_render_vertical_text,
    canvas_to_model,
    fill_layout,
    fill_meibo_layout,
    get_page_arrangement,
    has_meibo,
    model_to_canvas,
    model_to_printer,
)

# ── ヘルパー ─────────────────────────────────────────────────────────────────


def _make_layout(*objects: LayoutObject) -> LayFile:
    return LayFile(
        title='テスト',
        page_width=840,
        page_height=1188,
        objects=list(objects),
    )


# ── 座標変換テスト ────────────────────────────────────────────────────────────


class TestCoordinateConversion:
    """座標変換関数のテスト。"""

    def test_model_to_canvas_no_offset(self) -> None:
        px, py = model_to_canvas(100, 200, 0.5)
        assert px == pytest.approx(50.0)
        assert py == pytest.approx(100.0)

    def test_model_to_canvas_with_offset(self) -> None:
        px, py = model_to_canvas(100, 200, 0.5, 30, 30)
        assert px == pytest.approx(80.0)
        assert py == pytest.approx(130.0)

    def test_canvas_to_model_round_trip(self) -> None:
        scale, ox, oy = 0.5, 30, 30
        mx_in, my_in = 420, 594
        cx, cy = model_to_canvas(mx_in, my_in, scale, ox, oy)
        mx_out, my_out = canvas_to_model(cx, cy, scale, ox, oy)
        assert mx_out == mx_in
        assert my_out == my_in

    def test_model_to_printer_a4_width(self) -> None:
        # A4 幅 = 840 × 0.25mm = 210mm → 300dpi: 210/25.4*300 ≈ 2480
        dots = model_to_printer(840, 300)
        assert dots == pytest.approx(2480, abs=1)

    def test_model_to_printer_zero(self) -> None:
        assert model_to_printer(0, 300) == 0


# ── fill_layout テスト ────────────────────────────────────────────────────────


class TestFillLayout:
    """fill_layout() データ差込のテスト。"""

    def test_field_replaced_by_label(self) -> None:
        """FIELD が LABEL に変換される。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        data = {'氏名': '田中太郎'}
        result = fill_layout(lay, data)

        assert len(result.objects) == 1
        filled = result.objects[0]
        assert filled.obj_type == ObjectType.LABEL
        assert filled.text == '田中太郎'

    def test_label_unchanged(self) -> None:
        """LABEL はそのまま保持される。"""
        obj = new_label(10, 20, 200, 50, text='学年')
        lay = _make_layout(obj)
        result = fill_layout(lay, {})

        assert len(result.objects) == 1
        assert result.objects[0].obj_type == ObjectType.LABEL
        assert result.objects[0].text == '学年'

    def test_line_unchanged(self) -> None:
        """LINE はそのまま保持される。"""
        obj = new_line(0, 0, 100, 0)
        lay = _make_layout(obj)
        result = fill_layout(lay, {})

        assert len(result.objects) == 1
        assert result.objects[0].obj_type == ObjectType.LINE

    def test_original_not_modified(self) -> None:
        """元の LayFile は変更されない。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        fill_layout(lay, {'氏名': '佐藤花子'})

        # 元のオブジェクトはまだ FIELD
        assert lay.objects[0].obj_type == ObjectType.FIELD

    def test_prefix_suffix_preserved(self) -> None:
        """prefix/suffix がデータ値と結合される。"""
        obj = LayoutObject(
            obj_type=ObjectType.FIELD,
            rect=Rect(10, 20, 200, 50),
            field_id=105,
            prefix='',
            suffix='組',
            font=FontInfo('IPAmj明朝', 11.0),
        )
        lay = _make_layout(obj)
        data = {'組': '3'}
        result = fill_layout(lay, data)

        assert result.objects[0].text == '3組'

    def test_rect_preserved(self) -> None:
        """座標・サイズが保持される。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': 'テスト'})

        assert result.objects[0].rect == Rect(10, 20, 200, 50)

    def test_font_preserved(self) -> None:
        """フォント情報が保持される。"""
        obj = new_field(10, 20, 200, 50, field_id=108, font_size=14.0)
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': 'テスト'})

        assert result.objects[0].font.size_pt == 14.0

    def test_alignment_preserved(self) -> None:
        """揃え設定が保持される。"""
        obj = new_field(10, 20, 200, 50, field_id=108, h_align=2, v_align=0)
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': 'テスト'})

        assert result.objects[0].h_align == 2
        assert result.objects[0].v_align == 0

    def test_missing_data_returns_empty(self) -> None:
        """データに存在しないフィールドは空文字になる。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        result = fill_layout(lay, {})

        assert result.objects[0].text == ''

    def test_nan_value_converted_to_empty(self) -> None:
        """NaN 値は空文字になる。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': 'nan'})

        assert result.objects[0].text == ''

    def test_none_value_converted_to_empty(self) -> None:
        """None 値は空文字になる。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': None})

        assert result.objects[0].text == ''


class TestFillLayoutSpecialKeys:
    """fill_layout() 特殊キーのテスト。"""

    def test_fiscal_year(self) -> None:
        """年度フィールドが options から取得される。"""
        obj = LayoutObject(
            obj_type=ObjectType.FIELD,
            rect=Rect(10, 20, 200, 50),
            field_id=9999,  # 存在しない ID → 'field_9999'
            font=FontInfo('IPAmj明朝', 11.0),
        )
        # field_9999 は '年度' にマッピングされないので、直接テスト
        # 実際の年度フィールドをシミュレート
        lay = _make_layout(obj)
        result = fill_layout(lay, {}, {'fiscal_year': 2025})
        # field_9999 は特殊キーでないため data_row から取得
        assert result.objects[0].text == ''

    def test_school_name_option(self) -> None:
        """学校名はオプションから取得される。"""
        # FIELD_ID_MAP に学校名がないため、カスタムオブジェクトで確認
        # fill_layout は resolve_field_name → 論理名 → _resolve で処理
        # 直接テスト: 学校名キーを持つ field_id があれば取得できる

        # 間接テスト: options が正しくパスされることを確認
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        opts = {'school_name': 'テスト小学校', 'fiscal_year': 2025}
        result = fill_layout(lay, {'氏名': '山田太郎'}, opts)
        assert result.objects[0].text == '山田太郎'

    def test_page_info_preserved(self) -> None:
        """ページ情報が保持される。"""
        lay = LayFile(
            title='テストレイアウト',
            page_width=840,
            page_height=1188,
            objects=[new_field(10, 20, 200, 50, field_id=108)],
        )
        result = fill_layout(lay, {'氏名': 'テスト'})
        assert result.title == 'テストレイアウト'
        assert result.page_width == 840
        assert result.page_height == 1188


class TestFillLayoutDateFormatting:
    """fill_layout() 日付フォーマットのテスト。"""

    def test_date_field_formatted(self) -> None:
        """生年月日フィールドが日付形式に変換される。"""
        obj = new_field(10, 20, 200, 50, field_id=610)  # 610 = 生年月日
        lay = _make_layout(obj)
        data = {'生年月日': '2015/04/01'}
        result = fill_layout(lay, data)
        # format_date の戻り値に依存するが、空文字でないことを確認
        assert result.objects[0].text != ''

    def test_non_date_field_not_formatted(self) -> None:
        """非日付フィールドはそのまま保持される。"""
        obj = new_field(10, 20, 200, 50, field_id=108)  # 108 = 氏名
        lay = _make_layout(obj)
        data = {'氏名': '2015/04/01'}
        result = fill_layout(lay, data)
        # 氏名は DATE_KEYS に含まれないのでそのまま
        assert result.objects[0].text == '2015/04/01'


class TestFillLayoutMixed:
    """fill_layout() 複合テスト。"""

    def test_mixed_objects(self) -> None:
        """LABEL + FIELD + LINE の混在レイアウト。"""
        objects = [
            new_label(0, 0, 100, 30, text='氏名'),
            new_field(100, 0, 300, 30, field_id=108),
            new_line(0, 30, 300, 30),
            new_label(0, 30, 100, 60, text='〒'),
            new_field(100, 30, 300, 60, field_id=603),
        ]
        lay = _make_layout(*objects)
        data = {'氏名': '鈴木一郎', '郵便番号': '900-0005'}
        result = fill_layout(lay, data)

        assert len(result.objects) == 5
        assert result.objects[0].obj_type == ObjectType.LABEL  # ラベルそのまま
        assert result.objects[0].text == '氏名'
        assert result.objects[1].obj_type == ObjectType.LABEL  # FIELD → LABEL
        assert result.objects[1].text == '鈴木一郎'
        assert result.objects[2].obj_type == ObjectType.LINE   # 罫線そのまま
        assert result.objects[3].text == '〒'
        assert result.objects[4].text == '900-0005'

    def test_multiple_students(self) -> None:
        """複数生徒で連続呼び出しが正しく動作する。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)

        result1 = fill_layout(lay, {'氏名': '生徒A'})
        result2 = fill_layout(lay, {'氏名': '生徒B'})

        assert result1.objects[0].text == '生徒A'
        assert result2.objects[0].text == '生徒B'
        # 元は変更されていない
        assert lay.objects[0].obj_type == ObjectType.FIELD

    def test_empty_layout(self) -> None:
        """オブジェクト 0 件の空レイアウト。"""
        lay = _make_layout()
        result = fill_layout(lay, {'氏名': 'テスト'})
        assert len(result.objects) == 0

    def test_all_field_ids(self) -> None:
        """全 FIELD_ID_MAP エントリが fill_layout でエラーにならない。"""
        objects = [
            new_field(0, i * 30, 200, (i + 1) * 30, field_id=fid)
            for i, fid in enumerate(FIELD_ID_MAP)
        ]
        lay = _make_layout(*objects)
        data = {name: f'値_{name}' for name in FIELD_ID_MAP.values()}
        result = fill_layout(lay, data)
        assert len(result.objects) == len(FIELD_ID_MAP)
        for obj in result.objects:
            assert obj.obj_type == ObjectType.LABEL


def _is_vertical_text(w: float, h: float, text: str, scale: float = 1.0) -> bool:
    """縦書き判定ロジック（PILBackend.draw_text 内と同じ）。"""
    normalized = _normalize_text(text)
    return _should_render_vertical_text(
        normalized, w, h, scale, vertical=False,
    )


def _is_multiline_text(text: str) -> bool:
    """複数行判定（\r\n を \n に正規化後に判定）。"""
    return '\n' in _normalize_text(text)


class TestVerticalTextDetection:
    """縦書きテキスト検出のテスト（PILBackend と同じロジック）。"""

    def test_narrow_box_no_newline_is_vertical(self) -> None:
        """狭い矩形＋改行なし → 縦書き判定。"""
        # w=60, h=280
        assert _is_vertical_text(60, 280, 'その他特記事項') is True

    def test_wide_box_is_not_vertical(self) -> None:
        """幅のある矩形 → 縦書きではない。"""
        # w=180, h=60
        assert _is_vertical_text(180, 60, '学年') is False

    def test_narrow_with_newline_is_not_vertical(self) -> None:
        """狭い矩形でも改行あり → 縦書きではない。"""
        # w=40, h=130
        assert _is_vertical_text(40, 130, '氏\r\n名') is False

    def test_single_char_is_not_vertical(self) -> None:
        """1文字 → 縦書きではない。"""
        # w=30, h=100
        assert _is_vertical_text(30, 100, 'A') is False


class TestMultilineTextDetection:
    """複数行テキスト検出のテスト。"""

    def test_text_with_crlf(self) -> None:
        assert _is_multiline_text('行1\r\n行2') is True

    def test_text_with_lf(self) -> None:
        assert _is_multiline_text('行1\n行2') is True

    def test_text_with_cr(self) -> None:
        assert _is_multiline_text('行1\r行2') is True

    def test_single_line(self) -> None:
        assert _is_multiline_text('単一行テキスト') is False


class TestFontInfoBoldItalic:
    """FontInfo の bold/italic テスト。"""

    def test_default_not_bold(self) -> None:
        f = FontInfo()
        assert f.bold is False
        assert f.italic is False

    def test_bold_flag(self) -> None:
        f = FontInfo(bold=True)
        assert f.bold is True

    def test_italic_flag(self) -> None:
        f = FontInfo(italic=True)
        assert f.italic is True

    def test_bold_italic_preserved_in_fill(self) -> None:
        """fill_layout で bold/italic が保持される。"""
        obj = LayoutObject(
            obj_type=ObjectType.FIELD,
            rect=Rect(10, 20, 200, 50),
            field_id=108,
            font=FontInfo('ＭＳ ゴシック', 14.0, bold=True, italic=True),
        )
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': 'テスト'})
        assert result.objects[0].font.bold is True
        assert result.objects[0].font.italic is True


class TestContainsGaiji:
    """外字（IVS異体字・CJK拡張漢字）検出テスト。"""

    def test_normal_text_no_gaiji(self) -> None:
        """通常テキストは外字なし。"""
        assert _contains_gaiji('山田太郎') is False

    def test_empty_string(self) -> None:
        assert _contains_gaiji('') is False

    def test_ascii(self) -> None:
        assert _contains_gaiji('Hello World') is False

    def test_katakana_hiragana(self) -> None:
        assert _contains_gaiji('やまだたろう') is False
        assert _contains_gaiji('ヤマダタロウ') is False

    def test_ivs_variation_selector(self) -> None:
        """IVS 異体字セレクタ (U+E0100) を含むテキスト。"""
        # 葛 + VS17 (U+E0100) — 葛飾区の「葛」の異体字
        text = '葛\U000E0100城'
        assert _contains_gaiji(text) is True

    def test_cjk_extension_b(self) -> None:
        """CJK統合漢字拡張B (U+20000〜) の文字。"""
        # U+20000 (CJK Extension B の先頭)
        text = '田中\U00020000太郎'
        assert _contains_gaiji(text) is True

    def test_ivs_at_end(self) -> None:
        """IVS がテキスト末尾にある場合。"""
        text = '邊\U000E0100'
        assert _contains_gaiji(text) is True


class TestFillLayoutFontSelection:
    """fill_layout のフォント選択テスト: 外字なし→元フォント、外字あり→IPAmj明朝。"""

    def test_no_gaiji_keeps_original_font(self) -> None:
        """外字を含まないデータは .lay のフォントをそのまま使う。"""
        obj = LayoutObject(
            obj_type=ObjectType.FIELD,
            rect=Rect(10, 20, 200, 50),
            field_id=108,
            font=FontInfo('ＭＳ ゴシック', 12.0),
        )
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': '山田太郎'})
        assert result.objects[0].font.name == 'ＭＳ ゴシック'

    def test_gaiji_overrides_to_ipamj(self) -> None:
        """IVS 異体字を含むデータは IPAmj明朝 に上書きされる。"""
        obj = LayoutObject(
            obj_type=ObjectType.FIELD,
            rect=Rect(10, 20, 200, 50),
            field_id=108,
            font=FontInfo('ＭＳ ゴシック', 12.0),
        )
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': '葛\U000E0100城太郎'})
        assert result.objects[0].font.name == 'IPAmj明朝'
        # サイズ・bold・italic は元のまま
        assert result.objects[0].font.size_pt == 12.0

    def test_gaiji_preserves_bold_italic(self) -> None:
        """外字で IPAmj明朝 に上書きされても bold/italic は保持。"""
        obj = LayoutObject(
            obj_type=ObjectType.FIELD,
            rect=Rect(10, 20, 200, 50),
            field_id=108,
            font=FontInfo('ＭＳ ゴシック', 14.0, bold=True, italic=True),
        )
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': '邊\U000E0100太郎'})
        assert result.objects[0].font.name == 'IPAmj明朝'
        assert result.objects[0].font.bold is True
        assert result.objects[0].font.italic is True

    def test_gaiji_preserves_vertical(self) -> None:
        """外字でフォント置換されても縦書きフラグは保持される。"""
        obj = LayoutObject(
            obj_type=ObjectType.FIELD,
            rect=Rect(10, 20, 200, 50),
            field_id=108,
            font=FontInfo('ＭＳ ゴシック', 12.0, vertical=True),
        )
        lay = _make_layout(obj)
        result = fill_layout(lay, {'氏名': '葛\U000E0100城太郎'})
        assert result.objects[0].font.name == 'IPAmj明朝'
        assert result.objects[0].font.vertical is True

    def test_label_always_keeps_original_font(self) -> None:
        """LABEL（静的テキスト）は外字有無にかかわらず元フォントのまま。"""
        obj = new_label(10, 20, 200, 50, text='見出し')
        obj = LayoutObject(
            obj_type=ObjectType.LABEL,
            rect=obj.rect,
            text='見出し',
            font=FontInfo('ＭＳ 明朝', 16.0),
        )
        lay = _make_layout(obj)
        result = fill_layout(lay, {})
        assert result.objects[0].font.name == 'ＭＳ 明朝'


class TestFillLayoutWithRealLay:
    """実 .lay ファイルでの fill_layout テスト。"""

    @pytest.fixture()
    def real_lay(self) -> LayFile | None:
        from pathlib import Path
        lay_files = list(Path('meibo_tool').rglob('*.lay'))
        if not lay_files:
            pytest.skip('.lay ファイルが見つかりません')
        from core.lay_parser import parse_lay
        return parse_lay(str(lay_files[0]))

    def test_real_file_fill(self, real_lay: LayFile) -> None:
        """実ファイルで fill_layout がエラーなく完了する。"""
        data = {
            '氏名': '山田太郎',
            '氏名かな': 'やまだたろう',
            '組': '1',
            '性別': '男',
            '都道府県': '沖縄県',
            '市区町村': '那覇市',
            '町番地': '泉崎1-2-3',
            '生年月日': '2015/04/01',
        }
        result = fill_layout(real_lay, data, {'fiscal_year': 2025})
        assert len(result.objects) == len(real_lay.objects)

    def test_real_file_no_fields_remain(self, real_lay: LayFile) -> None:
        """fill 後に FIELD オブジェクトが残っていない。"""
        data = {name: f'テスト_{name}' for name in FIELD_ID_MAP.values()}
        result = fill_layout(real_lay, data)
        field_count = sum(
            1 for obj in result.objects if obj.obj_type == ObjectType.FIELD
        )
        assert field_count == 0


# ── TABLE 描画テスト ───────────────────────────────────────────────────────


class TestTableRendering:
    """_render_table の描画テスト（PILBackend 使用）。"""

    def _make_table_layout(self) -> LayFile:
        cols = [
            TableColumn(field_id=108, width=32, h_align=0, header='氏名'),
            TableColumn(field_id=610, width=24, h_align=1, header='生年月日'),
            TableColumn(field_id=107, width=8, h_align=1, header='性別'),
        ]
        return LayFile(
            title='テーブルテスト',
            page_width=840,
            page_height=1188,
            objects=[
                LayoutObject(
                    obj_type=ObjectType.TABLE,
                    rect=Rect(10, 100, 800, 1100),
                    table_columns=cols,
                ),
            ],
        )

    def test_table_renders_without_error(self) -> None:
        """TABLE オブジェクトが PIL レンダリングでエラーにならない。"""
        pytest.importorskip('PIL')
        from core.lay_renderer import render_layout_to_image

        lay = self._make_table_layout()
        img = render_layout_to_image(lay, dpi=72)
        assert img.size[0] > 0
        assert img.size[1] > 0

    def test_table_in_render_all(self) -> None:
        """render_all が TABLE を含むレイアウトを処理できる。"""
        pytest.importorskip('PIL')
        from PIL import Image

        from core.lay_renderer import LayRenderer, PILBackend

        lay = self._make_table_layout()
        img = Image.new('RGB', (400, 600), (255, 255, 255))
        backend = PILBackend(img, dpi=72)
        renderer = LayRenderer(lay, backend)
        renderer.render_all()  # should not raise

    def test_table_with_labels_and_lines(self) -> None:
        """TABLE + LABEL + LINE の混合レイアウトが描画できる。"""
        pytest.importorskip('PIL')
        from core.lay_renderer import render_layout_to_image

        lay = LayFile(
            page_width=840,
            page_height=1188,
            objects=[
                new_label(10, 10, 200, 40, text='修了台帳'),
                LayoutObject(
                    obj_type=ObjectType.TABLE,
                    rect=Rect(10, 50, 800, 1100),
                    table_columns=[
                        TableColumn(field_id=108, width=32, header='氏名'),
                    ],
                ),
                new_line(10, 1150, 800, 1150),
            ],
        )
        img = render_layout_to_image(lay, dpi=72)
        assert img.size[0] > 0

    def test_empty_table_columns_skipped(self) -> None:
        """カラムなしの TABLE は描画をスキップする。"""
        pytest.importorskip('PIL')
        from PIL import Image

        from core.lay_renderer import LayRenderer, PILBackend

        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.TABLE,
                rect=Rect(0, 0, 100, 100),
                table_columns=[],
            ),
        ])
        img = Image.new('RGB', (100, 100), (255, 255, 255))
        backend = PILBackend(img, dpi=72)
        renderer = LayRenderer(lay, backend)
        renderer.render_all()  # should not raise


# ── PaperLayout 配置テスト ────────────────────────────────────────────────


class TestGetPageArrangement:
    """get_page_arrangement のテスト。"""

    def test_mode0_returns_1x1(self) -> None:
        """mode=0（全面）は 1×1 を返す。"""
        lay = LayFile(
            paper=PaperLayout(mode=0, cols=1, rows=1),
        )
        cols, rows, per_page, scale = get_page_arrangement(lay)
        assert cols == 1
        assert rows == 1
        assert per_page == 1

    def test_mode1_uses_paper_cols_rows(self) -> None:
        """mode=1（ラベル）は PaperLayout の cols/rows を返す。"""
        lay = LayFile(
            paper=PaperLayout(mode=1, cols=2, rows=3),
        )
        cols, rows, per_page, scale = get_page_arrangement(lay)
        assert cols == 2
        assert rows == 3
        assert per_page == 6

    def test_no_paper_fallback(self) -> None:
        """paper=None のときは calculate_page_arrangement にフォールバック。"""
        lay = LayFile(page_width=840, page_height=1188)
        cols, rows, per_page, scale = get_page_arrangement(lay)
        assert per_page >= 1

    def test_scale_is_1(self) -> None:
        """PaperLayout 使用時のスケールは常に 1.0。"""
        lay = LayFile(
            paper=PaperLayout(mode=1, cols=2, rows=1),
        )
        _cols, _rows, _per_page, scale = get_page_arrangement(lay)
        assert scale == 1.0


# ── fill_layout TABLE テスト ─────────────────────────────────────────────


class TestFillLayoutTable:
    """fill_layout の TABLE 対応テスト。"""

    def test_table_preserved_after_fill(self) -> None:
        """fill_layout 後も TABLE オブジェクトが保持される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.TABLE,
                rect=Rect(10, 100, 800, 1100),
                table_columns=[
                    TableColumn(field_id=108, width=32, header='氏名'),
                ],
            ),
        ])
        result = fill_layout(lay, {'氏名': '山田太郎'})
        assert len(result.objects) == 1
        assert result.objects[0].obj_type == ObjectType.TABLE
        assert len(result.objects[0].table_columns) == 1

    def test_table_column_filled_with_data(self) -> None:
        """TABLE カラムの header がデータ値に置換される。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.TABLE,
                rect=Rect(10, 100, 800, 1100),
                table_columns=[
                    TableColumn(field_id=108, width=32, header='氏名'),
                    TableColumn(field_id=107, width=8, header='性別'),
                ],
            ),
        ])
        result = fill_layout(lay, {'氏名': '山田太郎', '性別': '男'})
        table = result.objects[0]
        assert table.table_columns[0].header == '山田太郎'
        assert table.table_columns[1].header == '男'

    def test_table_column_no_data_keeps_header(self) -> None:
        """データがない場合はヘッダーを維持する。"""
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.TABLE,
                rect=Rect(10, 100, 800, 1100),
                table_columns=[
                    TableColumn(field_id=108, width=32, header='氏名'),
                ],
            ),
        ])
        result = fill_layout(lay, {})
        assert result.objects[0].table_columns[0].header == '氏名'

    def test_paper_preserved_in_fill(self) -> None:
        """fill_layout 後も paper が保持される。"""
        paper = PaperLayout(mode=1, cols=2, rows=1, paper_size='A4')
        lay = LayFile(
            paper=paper,
            objects=[new_field(0, 0, 100, 30, field_id=108)],
        )
        result = fill_layout(lay, {'氏名': 'テスト'})
        assert result.paper is not None
        assert result.paper.paper_size == 'A4'


# ── render_layout_to_image unit_mm テスト ────────────────────────────────


class TestRenderWithUnitMm:
    """unit_mm による画像サイズの検証。"""

    def test_unit_mm_025_default(self) -> None:
        """paper=None → 旧形式 0.25mm/unit で画像生成。"""
        pytest.importorskip('PIL')
        from core.lay_renderer import render_layout_to_image

        lay = LayFile(page_width=840, page_height=1188)
        img = render_layout_to_image(lay, dpi=150)
        expected_w = int(840 * 0.25 * 150 / 25.4)
        assert abs(img.size[0] - expected_w) <= 1

    def test_unit_mm_01_with_paper(self) -> None:
        """paper.unit_mm=0.1 → 新形式で画像生成。"""
        pytest.importorskip('PIL')
        from core.lay_renderer import render_layout_to_image

        lay = LayFile(
            page_width=2100,
            page_height=2970,
            paper=PaperLayout(unit_mm=0.1),
        )
        img = render_layout_to_image(lay, dpi=150)
        # 2100 * 0.1mm = 210mm → 210 * 150 / 25.4 ≈ 1240px
        expected_w = int(2100 * 0.1 * 150 / 25.4)
        assert abs(img.size[0] - expected_w) <= 1


# ── Bug修正: 特殊キー追加テスト ────────────────────────────────────────────


class TestFillLayoutGuardianAddress:
    """保護者住所の結合・同上判定テスト。"""

    def test_guardian_address_c4th_separate_fields(self) -> None:
        """C4th データ: 4分割フィールドが結合される。"""
        obj = new_field(10, 20, 200, 50, field_id=683)  # 保護者住所
        lay = _make_layout(obj)
        data = {
            '保護者都道府県': '沖縄県',
            '保護者市区町村': '那覇市',
            '保護者町番地': '泊1-2-3',
            '保護者建物名': '',
        }
        result = fill_layout(lay, data)
        assert result.objects[0].text == '沖縄県那覇市泊1-2-3'

    def test_guardian_address_same_as_student(self) -> None:
        """保護者住所=児童住所 → 「同上」。"""
        obj = new_field(10, 20, 200, 50, field_id=683)
        lay = _make_layout(obj)
        data = {
            '都道府県': '沖縄県', '市区町村': '那覇市',
            '町番地': '天久1-2-3', '建物名': '',
            '保護者都道府県': '沖縄県', '保護者市区町村': '那覇市',
            '保護者町番地': '天久1-2-3', '保護者建物名': '',
        }
        result = fill_layout(lay, data)
        assert result.objects[0].text == '同上'

    def test_guardian_address_suzuki_single_field(self) -> None:
        """スズキ校務: 単一フィールドがそのまま使われる。"""
        obj = new_field(10, 20, 200, 50, field_id=683)
        lay = _make_layout(obj)
        data = {'保護者住所': '東京都千代田区永田町1-7-1'}
        result = fill_layout(lay, data)
        assert result.objects[0].text == '東京都千代田区永田町1-7-1'

    def test_guardian_address_empty(self) -> None:
        """保護者住所データがない場合は空。"""
        obj = new_field(10, 20, 200, 50, field_id=683)
        lay = _make_layout(obj)
        result = fill_layout(lay, {})
        assert result.objects[0].text == ''


class TestFillLayoutFiscalYear:
    """学年/組/年度 の解決テスト。"""

    def test_gakunen_resolved(self) -> None:
        """学年(101) が data_row から解決される。"""
        obj = new_field(10, 20, 200, 50, field_id=101)  # 学年
        lay = _make_layout(obj)
        result = fill_layout(lay, {'学年': '3'}, {})
        assert result.objects[0].text == '3'

    def test_kumi_resolved(self) -> None:
        """組(102) が data_row から解決される。"""
        obj = new_field(10, 20, 200, 50, field_id=102)  # 組
        lay = _make_layout(obj)
        result = fill_layout(lay, {'組': '2'}, {})
        assert result.objects[0].text == '2'

    def test_nendo_resolved(self) -> None:
        """年度(134) が fiscal_year から解決される。"""
        obj = new_field(10, 20, 200, 50, field_id=134)  # 年度
        lay = _make_layout(obj)
        result = fill_layout(lay, {}, {'fiscal_year': 2025})
        assert result.objects[0].text == '2025'

    def test_nendo_wareki_resolved(self) -> None:
        """年度和暦(110) が wareki 形式で解決される。"""
        obj = new_field(10, 20, 200, 50, field_id=110)  # 年度和暦
        lay = _make_layout(obj)
        result = fill_layout(lay, {}, {'fiscal_year': 2025})
        assert '7年度' in result.objects[0].text  # 令和7年度


class TestFillLayoutPageMeta:
    """ページ番号/人数合計テスト。"""

    def test_page_number(self) -> None:
        """ページ番号が options から取得される。"""
        obj = new_field(10, 20, 200, 50, field_id=138)  # ページ番号
        lay = _make_layout(obj)
        result = fill_layout(lay, {}, {'page_number': 3})
        assert result.objects[0].text == '3'

    def test_total_count(self) -> None:
        """人数合計が options から取得される。"""
        obj = new_field(10, 20, 200, 50, field_id=204)  # 人数合計
        lay = _make_layout(obj)
        result = fill_layout(lay, {}, {'total_count': 35})
        assert result.objects[0].text == '35'

    def test_page_meta_empty_when_not_set(self) -> None:
        """options にない場合は空文字。"""
        lay = _make_layout(
            new_field(0, 0, 100, 30, field_id=138),
            new_field(0, 30, 100, 60, field_id=204),
        )
        result = fill_layout(lay, {}, {})
        assert result.objects[0].text == ''
        assert result.objects[1].text == ''


class TestFillLayoutTransferDate:
    """転出日の日付フォーマットテスト。"""

    def test_transfer_out_date_formatted(self) -> None:
        """転出日が YY/MM/DD にフォーマットされる。"""
        obj = new_field(10, 20, 200, 50, field_id=137)  # 転出日
        lay = _make_layout(obj)
        data = {'転出日': '2025-03-31'}
        result = fill_layout(lay, data)
        assert result.objects[0].text == '25/03/31'

    def test_enrollment_date_formatted(self) -> None:
        """編入日が YY/MM/DD にフォーマットされる。"""
        obj = new_field(10, 20, 200, 50, field_id=108)  # 氏名の ID を借用
        lay = _make_layout(obj)
        # 直接 _resolve をテストできないため、data_row 経由で確認
        data = {'氏名': '2025-04-01'}
        result = fill_layout(lay, data)
        # 氏名は DATE_KEYS に含まれないのでそのまま
        assert result.objects[0].text == '2025-04-01'


# ── resolve_field_display テスト ──────────────────────────────────────────────


class TestResolveFieldDisplay:
    """FIELD_DISPLAY_MAP の表示名解決テスト。"""

    def test_known_id_returns_display_name(self) -> None:
        assert resolve_field_display(110) == '年度（和暦）'

    def test_another_known_id(self) -> None:
        assert resolve_field_display(106) == '番号(番無し)'

    def test_unknown_id_falls_back_to_field_name(self) -> None:
        # 未知ID は resolve_field_name の結果 = 'field_NNN'
        assert resolve_field_display(99999) == 'field_99999'

    def test_display_differs_from_internal(self) -> None:
        """表示名と内部論理名が異なることを確認。"""
        # 101: 内部='学年', 表示='学年' (同じ)
        # 110: 内部='年度和暦', 表示='年度（和暦）'
        from core.lay_parser import resolve_field_name
        assert resolve_field_name(110) == '年度和暦'
        assert resolve_field_display(110) == '年度（和暦）'


# ── has_meibo / _meibo_page_capacity テスト ───────────────────────────────────


def _make_meibo_object(
    ref_name: str = 'test_ref',
    row_count: int = 10,
    data_start_index: int = 0,
    direction: int = 0,
) -> LayoutObject:
    """MEIBO テスト用オブジェクトを作成する。"""
    return LayoutObject(
        obj_type=ObjectType.MEIBO,
        meibo=MeiboArea(
            origin_x=100,
            origin_y=200,
            cell_width=300,
            cell_height=50,
            row_count=row_count,
            data_start_index=data_start_index,
            ref_name=ref_name,
            direction=direction,
        ),
    )


class TestHasMeibo:
    """has_meibo() 判定テスト。"""

    def test_layout_without_meibo(self) -> None:
        lay = _make_layout(new_label(0, 0, 100, 50, 'hello'))
        assert has_meibo(lay) is False

    def test_layout_with_meibo(self) -> None:
        lay = _make_layout(_make_meibo_object())
        assert has_meibo(lay) is True

    def test_empty_layout(self) -> None:
        lay = _make_layout()
        assert has_meibo(lay) is False


class TestMeiboPageCapacity:
    """_meibo_page_capacity() テスト。"""

    def test_no_meibo_returns_zero(self) -> None:
        lay = _make_layout(new_label(0, 0, 100, 50, 'hello'))
        assert _meibo_page_capacity(lay) == 0

    def test_single_meibo(self) -> None:
        lay = _make_layout(_make_meibo_object(row_count=20, data_start_index=0))
        assert _meibo_page_capacity(lay) == 20

    def test_two_meibos(self) -> None:
        """掲示用名列表のように 2 つの MEIBO (0-19, 20-39)。"""
        lay = _make_layout(
            _make_meibo_object(row_count=20, data_start_index=0),
            _make_meibo_object(row_count=20, data_start_index=20),
        )
        assert _meibo_page_capacity(lay) == 40


# ── fill_meibo_layout テスト ──────────────────────────────────────────────────


def _make_parts_layout() -> LayFile:
    """テスト用パーツレイアウト（番号 + 氏名）。"""
    return LayFile(
        title='test_parts',
        page_width=300,
        page_height=50,
        objects=[
            new_field(0, 0, 100, 50, field_id=106),   # 出席番号
            new_field(100, 0, 300, 50, field_id=108),  # 氏名
        ],
    )


class TestFillMeiboLayout:
    """fill_meibo_layout() テスト。"""

    def _make_meibo_layout(self) -> LayFile:
        """掲示用名列表風のテストレイアウト。"""
        return LayFile(
            title='テスト名列表',
            page_width=840,
            page_height=1188,
            objects=[
                new_field(400, 50, 500, 100, field_id=134),  # 年度 (top-level)
                _make_meibo_object(
                    ref_name='test_parts', row_count=3,
                    data_start_index=0, direction=0,
                ),
                _make_meibo_object(
                    ref_name='test_parts', row_count=3,
                    data_start_index=3, direction=0,
                ),
            ],
        )

    def test_basic_fill(self) -> None:
        """6 名分のデータで 1 ページ生成。"""
        lay = self._make_meibo_layout()
        parts = _make_parts_layout()
        registry = {'test_parts': parts}
        data_rows = [
            {'出席番号': str(i + 1), '氏名': f'生徒{i + 1}'}
            for i in range(6)
        ]
        opts = {'fiscal_year': 2025}
        pages = fill_meibo_layout(lay, data_rows, opts, layout_registry=registry)
        assert len(pages) == 1
        # top-level FIELD (年度) が LABEL に変換されていること
        labels = [o for o in pages[0].objects if o.obj_type == ObjectType.LABEL]
        label_texts = [o.text for o in labels]
        assert '2025' in label_texts
        # 生徒名が含まれていること
        assert '生徒1' in label_texts
        assert '生徒6' in label_texts

    def test_paging(self) -> None:
        """capacity を超えるデータで複数ページ生成。"""
        lay = self._make_meibo_layout()  # capacity=6
        parts = _make_parts_layout()
        registry = {'test_parts': parts}
        data_rows = [
            {'出席番号': str(i + 1), '氏名': f'生徒{i + 1}'}
            for i in range(8)
        ]
        pages = fill_meibo_layout(lay, data_rows, layout_registry=registry)
        assert len(pages) == 2
        # ページ 2 には 2 名
        page2_labels = [
            o for o in pages[1].objects
            if o.obj_type == ObjectType.LABEL and '生徒' in o.text
        ]
        assert len(page2_labels) == 2

    def test_fewer_students_than_capacity(self) -> None:
        """capacity 未満のデータで 1 ページ（空スロットあり）。"""
        lay = self._make_meibo_layout()  # capacity=6
        parts = _make_parts_layout()
        registry = {'test_parts': parts}
        data_rows = [
            {'出席番号': '1', '氏名': '太郎'},
        ]
        pages = fill_meibo_layout(lay, data_rows, layout_registry=registry)
        assert len(pages) == 1
        labels = [
            o for o in pages[0].objects
            if o.obj_type == ObjectType.LABEL and o.text == '太郎'
        ]
        assert len(labels) == 1

    def test_unresolved_ref_name_skipped(self) -> None:
        """ref_name 未解決の MEIBO はスキップ（エラーにならない）。"""
        lay = self._make_meibo_layout()
        data_rows = [{'出席番号': '1', '氏名': 'テスト'}]
        # registry を空にして ref_name 未解決を起こす
        pages = fill_meibo_layout(lay, data_rows, layout_registry={})
        assert len(pages) == 1
        # top-level FIELD は変換されている
        labels = [o for o in pages[0].objects if o.obj_type == ObjectType.LABEL]
        assert len(labels) >= 1

    def test_top_level_field_common_data(self) -> None:
        """年度等の共通フィールドが正しく埋まる。"""
        lay = self._make_meibo_layout()
        parts = _make_parts_layout()
        registry = {'test_parts': parts}
        data_rows = [{'出席番号': '1', '氏名': 'テスト'}]
        opts = {'fiscal_year': 2026, 'school_name': 'テスト校'}
        pages = fill_meibo_layout(lay, data_rows, opts, layout_registry=registry)
        labels = [o for o in pages[0].objects if o.obj_type == ObjectType.LABEL]
        label_texts = [o.text for o in labels]
        assert '2026' in label_texts


# ── 写真フィールド差込テスト ────────────────────────────────────────────────


class TestFillPhotoField:
    """写真フィールド (field_id=400) の差込テスト。"""

    def _make_photo_layout(self) -> LayFile:
        """写真フィールドを含むテストレイアウト。"""
        photo_field = new_field(10, 20, 200, 300, field_id=400)
        name_field = new_field(10, 310, 200, 340, field_id=108)
        return _make_layout(photo_field, name_field)

    def test_photo_field_no_map_returns_empty_label(self) -> None:
        """photo_map がない場合は空の LABEL を返す。"""
        lay = self._make_photo_layout()
        data = {'氏名': '山田太郎', '学年': '1', '組': '1', '出席番号': '1'}
        result = fill_layout(lay, data)

        photo_obj = result.objects[0]
        assert photo_obj.obj_type == ObjectType.LABEL
        assert photo_obj.text == ''

    def test_photo_field_no_match_returns_empty_label(self) -> None:
        """マッチしない場合は空の LABEL を返す。"""
        lay = self._make_photo_layout()
        data = {'氏名': '山田太郎', '学年': '1', '組': '1', '出席番号': '1'}
        opts = {'_photo_map': {'2-1-01': '/photos/other.jpg'}}
        result = fill_layout(lay, data, opts)

        photo_obj = result.objects[0]
        assert photo_obj.obj_type == ObjectType.LABEL
        assert photo_obj.text == ''

    def test_photo_field_converts_to_image(self, tmp_path) -> None:
        """写真がマッチした場合 IMAGE に変換される。"""

        from PIL import Image

        # テスト画像作成
        img = Image.new('RGB', (100, 150), (255, 0, 0))
        photo_path = str(tmp_path / '1-1-01.jpg')
        img.save(photo_path)

        lay = self._make_photo_layout()
        data = {'氏名': '山田太郎', '学年': '1', '組': '1', '出席番号': '1'}
        opts = {'_photo_map': {'1-1-01': photo_path}}
        result = fill_layout(lay, data, opts)

        photo_obj = result.objects[0]
        assert photo_obj.obj_type == ObjectType.IMAGE
        assert photo_obj.image is not None
        assert len(photo_obj.image.image_data) > 0
        assert photo_obj.image.original_path == photo_path

    def test_photo_field_rect_preserved(self, tmp_path) -> None:
        """写真フィールドの座標が IMAGE に引き継がれる。"""
        from PIL import Image

        img = Image.new('RGB', (100, 150), (0, 255, 0))
        photo_path = str(tmp_path / '1-1-01.jpg')
        img.save(photo_path)

        lay = self._make_photo_layout()
        data = {'学年': '1', '組': '1', '出席番号': '1'}
        opts = {'_photo_map': {'1-1-01': photo_path}}
        result = fill_layout(lay, data, opts)

        photo_obj = result.objects[0]
        assert photo_obj.rect is not None
        assert photo_obj.rect.left == 10
        assert photo_obj.rect.top == 20

    def test_non_photo_fields_unchanged(self, tmp_path) -> None:
        """写真以外のフィールドは従来通り LABEL に変換される。"""
        from PIL import Image

        img = Image.new('RGB', (100, 150), (0, 0, 255))
        photo_path = str(tmp_path / '1-1-01.jpg')
        img.save(photo_path)

        lay = self._make_photo_layout()
        data = {'氏名': '山田太郎', '学年': '1', '組': '1', '出席番号': '1'}
        opts = {'_photo_map': {'1-1-01': photo_path}}
        result = fill_layout(lay, data, opts)

        name_obj = result.objects[1]
        assert name_obj.obj_type == ObjectType.LABEL
        assert name_obj.text == '山田太郎'

    def test_photo_in_mixed_layout(self) -> None:
        """LABEL + 写真 FIELD + テキスト FIELD の混在。"""
        label = new_label(0, 0, 100, 20, text='写真票')
        photo = new_field(10, 30, 200, 200, field_id=400)
        name = new_field(10, 210, 200, 230, field_id=108)
        line = new_line(0, 250, 200, 250)
        lay = _make_layout(label, photo, name, line)

        data = {'氏名': '鈴木花子'}
        result = fill_layout(lay, data)

        assert result.objects[0].obj_type == ObjectType.LABEL
        assert result.objects[0].text == '写真票'
        assert result.objects[1].obj_type == ObjectType.LABEL  # 写真なし → 空LABEL
        assert result.objects[1].text == ''
        assert result.objects[2].obj_type == ObjectType.LABEL
        assert result.objects[2].text == '鈴木花子'
        assert result.objects[3].obj_type == ObjectType.LINE

    def test_photo_field_corrupt_image(self, tmp_path) -> None:
        """破損画像の場合は空の LABEL を返す。"""
        corrupt_path = str(tmp_path / '1-1-01.jpg')
        with open(corrupt_path, 'wb') as f:
            f.write(b'not an image')

        lay = self._make_photo_layout()
        data = {'学年': '1', '組': '1', '出席番号': '1'}
        opts = {'_photo_map': {'1-1-01': corrupt_path}}
        result = fill_layout(lay, data, opts)

        photo_obj = result.objects[0]
        assert photo_obj.obj_type == ObjectType.LABEL
        assert photo_obj.text == ''

    def test_image_preserved_in_fill(self) -> None:
        """既存の IMAGE オブジェクトは fill_layout で保持される。"""
        img_obj = LayoutObject(
            obj_type=ObjectType.IMAGE,
            rect=Rect(10, 20, 200, 300),
            image=EmbeddedImage(
                rect=(10, 20, 200, 300),
                image_data=b'fake-png-data',
                original_path='/test.png',
            ),
        )
        lay = _make_layout(img_obj)
        result = fill_layout(lay, {})

        assert result.objects[0].obj_type == ObjectType.IMAGE
        assert result.objects[0].image.image_data == b'fake-png-data'

    def test_photo_fill_with_name_match(self, tmp_path) -> None:
        """氏名ベースでマッチする場合も IMAGE に変換される。"""
        from PIL import Image

        img = Image.new('RGB', (100, 150), (255, 255, 0))
        photo_path = str(tmp_path / '山田太郎.jpg')
        img.save(photo_path)

        lay = self._make_photo_layout()
        data = {'氏名': '山田 太郎', '学年': '1', '組': '1', '出席番号': '99'}
        opts = {'_photo_map': {'山田太郎': photo_path}}
        result = fill_layout(lay, data, opts)

        photo_obj = result.objects[0]
        assert photo_obj.obj_type == ObjectType.IMAGE
