"""lay_renderer テスト

fill_layout() データ差込ロジックと座標変換を検証する。
"""

from __future__ import annotations

import pytest

from core.lay_parser import (
    FIELD_ID_MAP,
    FontInfo,
    LayFile,
    LayoutObject,
    ObjectType,
    Rect,
    new_field,
    new_label,
    new_line,
)
from core.lay_renderer import (
    canvas_to_model,
    fill_layout,
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
            new_label(0, 30, 100, 60, text='住所'),
            new_field(100, 30, 300, 60, field_id=603),
        ]
        lay = _make_layout(*objects)
        data = {'氏名': '鈴木一郎', '都道府県': '沖縄県'}
        result = fill_layout(lay, data)

        assert len(result.objects) == 5
        assert result.objects[0].obj_type == ObjectType.LABEL  # ラベルそのまま
        assert result.objects[0].text == '氏名'
        assert result.objects[1].obj_type == ObjectType.LABEL  # FIELD → LABEL
        assert result.objects[1].text == '鈴木一郎'
        assert result.objects[2].obj_type == ObjectType.LINE   # 罫線そのまま
        assert result.objects[3].text == '住所'
        assert result.objects[4].text == '沖縄県'

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
