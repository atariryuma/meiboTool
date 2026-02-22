"""lay_serializer.py のユニットテスト

テスト対象:
  - LayFile → dict → LayFile ラウンドトリップ
  - JSON ファイル保存・読込
  - エラーハンドリング
"""

from __future__ import annotations

import json
import os

import pytest

from core.lay_parser import (
    FontInfo,
    LayFile,
    LayoutObject,
    ObjectType,
    PaperLayout,
    Point,
    RawTag,
    Rect,
    TableColumn,
    new_field,
    new_label,
    new_line,
)
from core.lay_serializer import (
    dict_to_layfile,
    layfile_to_dict,
    load_layout,
    save_layout,
)

# ── ヘルパー ─────────────────────────────────────────────────────────────────

def _make_full_layout() -> LayFile:
    """テスト用のフル構成 LayFile を構築する。"""
    return LayFile(
        title='テスト帳票',
        version=1600,
        page_width=840,
        page_height=1188,
        objects=[
            new_label(10, 20, 200, 50, text='氏名', font_size=12.0),
            new_field(200, 20, 500, 50, field_id=108, font_size=11.0),
            new_line(10, 60, 500, 60),
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 70, 200, 100),
                text='住所',
                font=FontInfo('ＭＳ 明朝', 10.0),
                h_align=0,
                v_align=2,
                prefix='〒',
                suffix='様',
            ),
        ],
    )


# ── テストクラス ─────────────────────────────────────────────────────────────


class TestRoundTrip:
    """dict ラウンドトリップのテスト。"""

    def test_empty_layout(self):
        lay = LayFile()
        d = layfile_to_dict(lay)
        restored = dict_to_layfile(d)
        assert restored.title == ''
        assert restored.page_width == 840
        assert restored.page_height == 1188
        assert restored.objects == []

    def test_full_layout(self):
        lay = _make_full_layout()
        d = layfile_to_dict(lay)
        restored = dict_to_layfile(d)

        assert restored.title == 'テスト帳票'
        assert restored.version == 1600
        assert restored.page_width == 840
        assert len(restored.objects) == 4

    def test_label_preserved(self):
        lay = _make_full_layout()
        restored = dict_to_layfile(layfile_to_dict(lay))
        label = restored.labels[0]

        assert label.obj_type == ObjectType.LABEL
        assert label.text == '氏名'
        assert label.rect == Rect(10, 20, 200, 50)
        assert label.font.size_pt == 12.0

    def test_field_preserved(self):
        lay = _make_full_layout()
        restored = dict_to_layfile(layfile_to_dict(lay))
        fld = restored.fields[0]

        assert fld.obj_type == ObjectType.FIELD
        assert fld.field_id == 108
        assert fld.rect == Rect(200, 20, 500, 50)

    def test_line_preserved(self):
        lay = _make_full_layout()
        restored = dict_to_layfile(layfile_to_dict(lay))
        line = restored.lines[0]

        assert line.obj_type == ObjectType.LINE
        assert line.line_start == Point(10, 60)
        assert line.line_end == Point(500, 60)

    def test_font_info_preserved(self):
        lay = _make_full_layout()
        restored = dict_to_layfile(layfile_to_dict(lay))
        label_with_font = restored.objects[3]

        assert label_with_font.font.name == 'ＭＳ 明朝'
        assert label_with_font.font.size_pt == 10.0

    def test_alignment_preserved(self):
        lay = _make_full_layout()
        restored = dict_to_layfile(layfile_to_dict(lay))
        obj = restored.objects[3]

        assert obj.h_align == 0
        assert obj.v_align == 2

    def test_prefix_suffix_preserved(self):
        lay = _make_full_layout()
        restored = dict_to_layfile(layfile_to_dict(lay))
        obj = restored.objects[3]

        assert obj.prefix == '〒'
        assert obj.suffix == '様'

    def test_style_flags_preserved(self):
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(10, 20, 200, 100),
                text='枠付きラベル',
                style_1001=0,
                style_1002=10,
                style_1003=-1,
                raw_tags=[
                    RawTag(path=[0x03EB, 0x03E9], payload=b'\x00\x00\x00\x00', payload_len=4),
                    RawTag(path=[0x03EB, 0x0BBC], payload=b'', payload_len=14),
                    RawTag(path=[0x03EB, 0x0BBC, 0x03EA], payload=b'\x78\x00\x00\x00', payload_len=4),
                ],
            ),
        ], raw_tags=[
            RawTag(path=[0x05E0], payload=b'\x02\x00\x00\x00', payload_len=4),
        ])
        restored = dict_to_layfile(layfile_to_dict(lay))
        obj = restored.objects[0]
        assert obj.style_1001 == 0
        assert obj.style_1002 == 10
        assert obj.style_1003 == -1
        assert obj.raw_tags == [
            RawTag(path=[0x03EB, 0x03E9], payload=b'\x00\x00\x00\x00', payload_len=4),
            RawTag(path=[0x03EB, 0x0BBC], payload=b'', payload_len=14),
            RawTag(path=[0x03EB, 0x0BBC, 0x03EA], payload=b'\x78\x00\x00\x00', payload_len=4),
        ]
        assert restored.raw_tags == [
            RawTag(path=[0x05E0], payload=b'\x02\x00\x00\x00', payload_len=4),
        ]

    def test_page_dimensions_preserved(self):
        lay = LayFile(page_width=1188, page_height=840)
        restored = dict_to_layfile(layfile_to_dict(lay))

        assert restored.page_width == 1188
        assert restored.page_height == 840


class TestFileSaveLoad:
    """ファイル保存・読込のテスト。"""

    def test_save_and_load(self, tmp_path):
        lay = _make_full_layout()
        path = str(tmp_path / 'test.json')
        save_layout(lay, path)

        assert os.path.isfile(path)

        restored = load_layout(path)
        assert restored.title == 'テスト帳票'
        assert len(restored.objects) == 4

    def test_json_is_valid(self, tmp_path):
        lay = _make_full_layout()
        path = str(tmp_path / 'test.json')
        save_layout(lay, path)

        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        assert data['format'] == 'meibo_layout_v1'
        assert isinstance(data['objects'], list)

    def test_creates_parent_dirs(self, tmp_path):
        path = str(tmp_path / 'sub' / 'dir' / 'test.json')
        save_layout(LayFile(), path)
        assert os.path.isfile(path)


class TestErrorHandling:
    """エラーハンドリングのテスト。"""

    def test_invalid_format_raises(self):
        with pytest.raises(ValueError, match='Unknown format'):
            dict_to_layfile({'format': 'wrong_format'})

    def test_missing_format_raises(self):
        with pytest.raises(ValueError, match='Unknown format'):
            dict_to_layfile({'title': 'test'})

    def test_unknown_object_type_defaults_to_label(self):
        data = {
            'format': 'meibo_layout_v1',
            'objects': [{'type': 'UNKNOWN_TYPE'}],
        }
        lay = dict_to_layfile(data)
        assert lay.objects[0].obj_type == ObjectType.LABEL


class TestBoldItalicRoundTrip:
    """bold/italic の JSON ラウンドトリップテスト。"""

    @pytest.mark.parametrize(
        ('bold', 'italic'),
        [(True, False), (False, True), (True, True)],
    )
    def test_font_style_preserved(self, bold: bool, italic: bool):
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(0, 0, 100, 30),
                text='スタイル',
                font=FontInfo('ＭＳ ゴシック', 12.0, bold=bold, italic=italic),
            ),
        ])
        restored = dict_to_layfile(layfile_to_dict(lay))
        assert restored.objects[0].font.bold is bold
        assert restored.objects[0].font.italic is italic

    def test_normal_no_bold_italic_in_json(self):
        """通常フォントでは bold/italic キーが JSON に含まれない。"""
        lay = LayFile(objects=[
            new_label(0, 0, 100, 30, text='通常'),
        ])
        d = layfile_to_dict(lay)
        font_dict = d['objects'][0].get('font', {})
        assert 'bold' not in font_dict
        assert 'italic' not in font_dict
        assert 'underline' not in font_dict
        assert 'strikethrough' not in font_dict

    @pytest.mark.parametrize(
        ('underline', 'strikethrough'),
        [(True, False), (False, True), (True, True)],
    )
    def test_underline_strikethrough_preserved(self, underline: bool, strikethrough: bool):
        lay = LayFile(objects=[
            LayoutObject(
                obj_type=ObjectType.LABEL,
                rect=Rect(0, 0, 100, 30),
                text='装飾',
                font=FontInfo(
                    'ＭＳ ゴシック', 12.0,
                    underline=underline, strikethrough=strikethrough,
                ),
            ),
        ])
        restored = dict_to_layfile(layfile_to_dict(lay))
        assert restored.objects[0].font.underline is underline
        assert restored.objects[0].font.strikethrough is strikethrough


# ── アトミック書き込みテスト ──────────────────────────────────────────────────


class TestAtomicSave:
    """save_layout のアトミック書き込みテスト。"""

    def test_overwrite_existing_file(self, tmp_path):
        """既存ファイルの上書きが成功すること。"""
        path = str(tmp_path / 'test.json')
        save_layout(LayFile(title='v1'), path)
        save_layout(LayFile(title='v2'), path)

        restored = load_layout(path)
        assert restored.title == 'v2'

    def test_original_preserved_on_error(self, tmp_path):
        """シリアライズエラー時に元ファイルが残ること。"""
        path = str(tmp_path / 'test.json')
        save_layout(LayFile(title='元データ'), path)

        # json.dump がエラーを起こすオブジェクトを仕込む
        from unittest.mock import patch
        with patch('core.lay_serializer.json.dump', side_effect=TypeError('bad')), \
             pytest.raises(TypeError):
            save_layout(LayFile(title='壊れる'), path)

        # 元ファイルが無傷であること
        restored = load_layout(path)
        assert restored.title == '元データ'

    def test_no_temp_file_left_on_error(self, tmp_path):
        """エラー時に一時ファイルが残らないこと。"""
        path = str(tmp_path / 'test.json')

        from unittest.mock import patch
        with patch('core.lay_serializer.json.dump', side_effect=OSError('disk full')), \
             pytest.raises(OSError):
            save_layout(LayFile(), path)

        # 一時ファイルが残っていないこと
        files = list(tmp_path.iterdir())
        assert all(not f.name.startswith('.layout_') for f in files)


# ── デシリアライズ エッジケーステスト ─────────────────────────────────────────


class TestDeserializationEdgeCases:
    """dict_to_layfile / _dict_to_object のエッジケーステスト。"""

    def test_object_with_empty_text(self):
        """text='' のオブジェクトが正しく復元される。"""
        data = {
            'format': 'meibo_layout_v1',
            'objects': [{'type': 'LABEL', 'rect': [0, 0, 100, 30]}],
        }
        lay = dict_to_layfile(data)
        assert lay.objects[0].text == ''

    def test_object_with_no_rect(self):
        """rect なしのオブジェクトが正しく復元される。"""
        data = {
            'format': 'meibo_layout_v1',
            'objects': [{'type': 'LABEL', 'text': 'no rect'}],
        }
        lay = dict_to_layfile(data)
        assert lay.objects[0].rect is None

    def test_font_defaults(self):
        """font キーなしで FontInfo のデフォルトが使われる。"""
        data = {
            'format': 'meibo_layout_v1',
            'objects': [{'type': 'LABEL'}],
        }
        lay = dict_to_layfile(data)
        assert lay.objects[0].font.name == ''
        assert lay.objects[0].font.size_pt == 10.0
        assert lay.objects[0].font.bold is False

    def test_load_nonexistent_file_raises(self, tmp_path):
        """存在しないファイルを読むと FileNotFoundError。"""
        with pytest.raises(FileNotFoundError):
            load_layout(str(tmp_path / 'no_such.json'))

    def test_load_invalid_json_raises(self, tmp_path):
        """不正な JSON ファイルを読むと json.JSONDecodeError。"""
        path = str(tmp_path / 'bad.json')
        with open(path, 'w') as f:
            f.write('{broken json}')
        with pytest.raises(json.JSONDecodeError):
            load_layout(path)

# ── V2 フォーマット テスト ───────────────────────────────────────────────────


class TestV2FormatDetection:
    """v1/v2 フォーマット自動判定のテスト。"""

    def test_no_paper_no_table_emits_v1(self):
        lay = LayFile(objects=[new_label(0, 0, 100, 30, text='test')])
        d = layfile_to_dict(lay)
        assert d['format'] == 'meibo_layout_v1'

    def test_paper_emits_v2(self):
        lay = LayFile(paper=PaperLayout(mode=0))
        d = layfile_to_dict(lay)
        assert d['format'] == 'meibo_layout_v2'

    def test_table_columns_emits_v2(self):
        obj = LayoutObject(
            obj_type=ObjectType.TABLE,
            rect=Rect(0, 0, 800, 600),
            table_columns=[TableColumn(field_id=108, header='氏名')],
        )
        lay = LayFile(objects=[obj])
        d = layfile_to_dict(lay)
        assert d['format'] == 'meibo_layout_v2'

    def test_v2_readable_by_dict_to_layfile(self):
        data = {
            'format': 'meibo_layout_v2',
            'title': 'v2テスト',
            'objects': [],
        }
        lay = dict_to_layfile(data)
        assert lay.title == 'v2テスト'


class TestPaperLayoutRoundTrip:
    """PaperLayout の JSON ラウンドトリップテスト。"""

    def test_paper_preserved(self):
        paper = PaperLayout(
            mode=1,
            unit_mm=0.1,
            item_width_mm=84.0,
            item_height_mm=272.0,
            cols=2,
            rows=1,
            margin_left_mm=14.0,
            margin_top_mm=11.0,
            spacing_h_mm=10.0,
            spacing_v_mm=0.0,
            paper_size='A4',
            orientation='portrait',
        )
        lay = LayFile(paper=paper)
        d = layfile_to_dict(lay)
        restored = dict_to_layfile(d)

        assert restored.paper is not None
        p = restored.paper
        assert p.mode == 1
        assert p.unit_mm == 0.1
        assert p.item_width_mm == 84.0
        assert p.item_height_mm == 272.0
        assert p.cols == 2
        assert p.rows == 1
        assert p.margin_left_mm == 14.0
        assert p.margin_top_mm == 11.0
        assert p.spacing_h_mm == 10.0
        assert p.spacing_v_mm == 0.0
        assert p.paper_size == 'A4'
        assert p.orientation == 'portrait'

    def test_no_paper_returns_none(self):
        lay = LayFile()
        d = layfile_to_dict(lay)
        restored = dict_to_layfile(d)
        assert restored.paper is None

    def test_paper_json_keys(self):
        lay = LayFile(paper=PaperLayout(mode=0, paper_size='A3'))
        d = layfile_to_dict(lay)
        assert 'paper' in d
        assert d['paper']['paper_size'] == 'A3'

    def test_paper_file_roundtrip(self, tmp_path):
        paper = PaperLayout(
            mode=0,
            unit_mm=0.1,
            item_width_mm=210.0,
            item_height_mm=297.0,
            paper_size='A4',
            orientation='portrait',
        )
        lay = LayFile(title='用紙テスト', paper=paper)
        path = str(tmp_path / 'paper.json')
        save_layout(lay, path)
        restored = load_layout(path)
        assert restored.paper.paper_size == 'A4'
        assert restored.paper.unit_mm == 0.1


class TestTableColumnRoundTrip:
    """TABLE オブジェクト + カラムの JSON ラウンドトリップテスト。"""

    def test_table_columns_preserved(self):
        cols = [
            TableColumn(field_id=108, width=32, h_align=0, header='氏名'),
            TableColumn(field_id=610, width=24, h_align=1, header='生年月日'),
            TableColumn(field_id=603, width=16, h_align=2, header='都道府県'),
        ]
        obj = LayoutObject(
            obj_type=ObjectType.TABLE,
            rect=Rect(10, 100, 800, 1100),
            table_columns=cols,
        )
        lay = LayFile(objects=[obj])
        d = layfile_to_dict(lay)
        restored = dict_to_layfile(d)

        table = restored.tables[0]
        assert len(table.table_columns) == 3
        assert table.table_columns[0].field_id == 108
        assert table.table_columns[0].header == '氏名'
        assert table.table_columns[1].h_align == 1
        assert table.table_columns[2].width == 16

    def test_table_type_preserved(self):
        obj = LayoutObject(
            obj_type=ObjectType.TABLE,
            rect=Rect(0, 0, 100, 100),
            table_columns=[TableColumn(field_id=108)],
        )
        lay = LayFile(objects=[obj])
        restored = dict_to_layfile(layfile_to_dict(lay))
        assert restored.objects[0].obj_type == ObjectType.TABLE

    def test_empty_columns_not_in_json(self):
        obj = LayoutObject(
            obj_type=ObjectType.LABEL,
            rect=Rect(0, 0, 100, 30),
            text='no cols',
        )
        lay = LayFile(objects=[obj])
        d = layfile_to_dict(lay)
        assert 'table_columns' not in d['objects'][0]

    def test_table_json_structure(self):
        obj = LayoutObject(
            obj_type=ObjectType.TABLE,
            rect=Rect(10, 20, 800, 600),
            table_columns=[
                TableColumn(field_id=108, width=32, h_align=0, header='氏名'),
            ],
        )
        lay = LayFile(objects=[obj])
        d = layfile_to_dict(lay)
        obj_d = d['objects'][0]
        assert obj_d['type'] == 'TABLE'
        assert len(obj_d['table_columns']) == 1
        col_d = obj_d['table_columns'][0]
        assert col_d['field_id'] == 108
        assert col_d['width'] == 32
        assert col_d['header'] == '氏名'

    def test_table_column_h_align_zero_omitted(self):
        """h_align=0 は JSON に含まれない（デフォルト値）。"""
        obj = LayoutObject(
            obj_type=ObjectType.TABLE,
            rect=Rect(0, 0, 100, 100),
            table_columns=[TableColumn(field_id=108, h_align=0, header='氏名')],
        )
        d = layfile_to_dict(LayFile(objects=[obj]))
        assert 'h_align' not in d['objects'][0]['table_columns'][0]

    def test_full_v2_file_roundtrip(self, tmp_path):
        """v2: paper + TABLE 含む完全なファイルラウンドトリップ。"""
        paper = PaperLayout(
            mode=0, unit_mm=0.1,
            item_width_mm=297.0, item_height_mm=210.0,
            paper_size='A4', orientation='landscape',
        )
        cols = [
            TableColumn(field_id=108, width=32, header='氏名'),
            TableColumn(field_id=610, width=24, h_align=1, header='生年月日'),
        ]
        lay = LayFile(
            title='修了台帳',
            version=1600,
            page_width=2970,
            page_height=2100,
            objects=[
                new_label(10, 20, 200, 50, text='修了台帳'),
                LayoutObject(
                    obj_type=ObjectType.TABLE,
                    rect=Rect(10, 100, 2900, 2000),
                    table_columns=cols,
                ),
                new_line(10, 2050, 2900, 2050),
            ],
            paper=paper,
        )

        path = str(tmp_path / 'v2_full.json')
        save_layout(lay, path)

        # JSON が v2 フォーマットであること
        with open(path, encoding='utf-8') as f:
            raw = json.load(f)
        assert raw['format'] == 'meibo_layout_v2'
        assert 'paper' in raw

        # ラウンドトリップ
        restored = load_layout(path)
        assert restored.title == '修了台帳'
        assert restored.paper.paper_size == 'A4'
        assert restored.paper.orientation == 'landscape'
        assert len(restored.objects) == 3
        assert restored.tables[0].table_columns[0].header == '氏名'
        assert restored.tables[0].table_columns[1].h_align == 1


