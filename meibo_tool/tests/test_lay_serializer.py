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
    Point,
    Rect,
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

_SAMPLE_LAY = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
    'R8年度小学校個票20260130.lay',
)


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


# ── 実ファイルラウンドトリップ ─────────────────────────────────────────────────


@pytest.mark.skipif(
    not os.path.isfile(_SAMPLE_LAY),
    reason='サンプル .lay ファイルが見つかりません',
)
class TestRealFileRoundTrip:
    """実 .lay → JSON → LayFile ラウンドトリップ。"""

    def test_roundtrip_preserves_object_count(self, tmp_path):
        from core.lay_parser import parse_lay

        lay = parse_lay(_SAMPLE_LAY)
        path = str(tmp_path / 'roundtrip.json')
        save_layout(lay, path)
        restored = load_layout(path)

        assert len(restored.objects) == len(lay.objects)

    def test_roundtrip_preserves_title(self, tmp_path):
        from core.lay_parser import parse_lay

        lay = parse_lay(_SAMPLE_LAY)
        path = str(tmp_path / 'roundtrip.json')
        save_layout(lay, path)
        restored = load_layout(path)

        assert restored.title == lay.title

    def test_roundtrip_preserves_field_ids(self, tmp_path):
        from core.lay_parser import parse_lay

        lay = parse_lay(_SAMPLE_LAY)
        path = str(tmp_path / 'roundtrip.json')
        save_layout(lay, path)
        restored = load_layout(path)

        orig_ids = sorted(f.field_id for f in lay.fields)
        rest_ids = sorted(f.field_id for f in restored.fields)
        assert orig_ids == rest_ids
