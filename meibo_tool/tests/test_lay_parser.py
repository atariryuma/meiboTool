"""lay_parser.py のユニットテスト

テスト対象:
  - ヘッダー検証
  - zlib 展開
  - TLV パース
  - オブジェクト分類（LABEL / FIELD / LINE / GROUP）
  - フィールド ID マッピング
  - 座標・フォントパース
  - 実ファイルパース
"""

from __future__ import annotations

import os
import struct
import zlib

import pytest

from core.lay_parser import (
    FIELD_ID_MAP,
    FontInfo,
    LayFile,
    ObjectType,
    Point,
    Rect,
    parse_lay,
    parse_lay_bytes,
    resolve_field_name,
)

# ── ヘルパー ─────────────────────────────────────────────────────────────────

_MAGIC = 'EXCMIDataContainer01'.encode('utf-16-le')


def _make_tlv(tag: int, payload: bytes) -> bytes:
    """TLV エントリをバイト列で構築する。"""
    return struct.pack('<HI', tag, len(payload)) + payload


def _make_lay_bytes(decompressed: bytes) -> bytes:
    """最小限の .lay バイト列を構築する（テスト用）。"""
    compressed = zlib.compress(decompressed)
    return _MAGIC + struct.pack('<I', len(decompressed)) + compressed


def _make_minimal_decompressed(
    title: str = 'テスト',
    objects_payload: bytes = b'',
) -> bytes:
    """展開後データの最小構造を構築する。"""
    # タイトル TLV (実ファイルのタイトルには先頭に2バイトのプレフィックスがある)
    title_bytes = title.encode('utf-16-le')
    title_tlv = _make_tlv(0x05DC, title_bytes)

    # コンテンツブロック: ルートプロパティ + オブジェクトリスト + ページジオメトリ
    obj_list_tlv = _make_tlv(0x03EA, objects_payload)
    page_rect = struct.pack('<II', 840, 1188)
    page_tlv = _make_tlv(0x07D1, page_rect)
    content_inner = (
        _make_tlv(0x03E8, b'\x02')
        + obj_list_tlv
        + page_tlv
    )
    content_tlv = _make_tlv(0x05E1, content_inner)

    # ヘッダー: version(uint16) + content_size(uint32) + TLV列
    body = title_tlv + content_tlv
    header = struct.pack('<HI', 1600, len(body) + 6)
    return header + body


# ── サンプルファイルパス ─────────────────────────────────────────────────────

_SAMPLE_LAY = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
    'R8年度小学校個票20260130.lay',
)


# ── テストクラス ─────────────────────────────────────────────────────────────


class TestHeader:
    """ヘッダー検証のテスト。"""

    def test_valid_header_accepted(self):
        data = _make_lay_bytes(_make_minimal_decompressed())
        lay = parse_lay_bytes(data)
        assert isinstance(lay, LayFile)

    def test_invalid_header_raises(self):
        bad = b'X' * 40 + struct.pack('<I', 10) + zlib.compress(b'\x00' * 10)
        with pytest.raises(ValueError, match='Invalid .lay header'):
            parse_lay_bytes(bad)

    def test_truncated_file_raises(self):
        with pytest.raises(ValueError, match='too short'):
            parse_lay_bytes(b'\x00' * 10)


class TestDecompression:
    """zlib 展開のテスト。"""

    def test_decompressed_size_matches(self):
        raw = _make_minimal_decompressed()
        data = _make_lay_bytes(raw)
        lay = parse_lay_bytes(data)
        assert lay.version == 1600

    def test_corrupt_zlib_raises(self):
        bad = _MAGIC + struct.pack('<I', 100) + b'\x00' * 50
        with pytest.raises(ValueError, match='zlib'):
            parse_lay_bytes(bad)

    def test_size_mismatch_raises(self):
        raw = _make_minimal_decompressed()
        compressed = zlib.compress(raw)
        # 偽のサイズを設定
        bad = _MAGIC + struct.pack('<I', len(raw) + 999) + compressed
        with pytest.raises(ValueError, match='mismatch'):
            parse_lay_bytes(bad)


class TestFieldIdMap:
    """フィールド ID マッピングのテスト。"""

    def test_known_id_mapped(self):
        assert resolve_field_name(108) == '氏名'
        assert resolve_field_name(610) == '生年月日'
        assert resolve_field_name(105) == '組'

    def test_unknown_id_fallback(self):
        assert resolve_field_name(99999) == 'field_99999'

    def test_all_known_ids_have_names(self):
        for _fid, name in FIELD_ID_MAP.items():
            assert isinstance(name, str) and len(name) > 0


class TestObjectTypes:
    """オブジェクト分類のテスト。"""

    def test_label_object(self):
        text = 'テスト'.encode('utf-16-le')
        font_name = 'ＭＳ 明朝'.encode('utf-16-le')
        font_block = (
            _make_tlv(0x03E8, font_name)
            + _make_tlv(0x03EA, struct.pack('<I', 120))
        )
        inner = (
            _make_tlv(0x07D1, struct.pack('<IIII', 10, 20, 100, 50))
            + _make_tlv(0x0BB9, text)
            + _make_tlv(0x0BBC, font_block)
        )
        label_tlv = _make_tlv(0x03EB, inner)

        raw = _make_minimal_decompressed(objects_payload=label_tlv)
        lay = parse_lay_bytes(_make_lay_bytes(raw))

        assert len(lay.labels) == 1
        obj = lay.labels[0]
        assert obj.obj_type == ObjectType.LABEL
        assert obj.text == 'テスト'
        assert obj.rect is not None
        assert obj.rect.left == 10
        assert obj.rect.right == 100
        assert obj.font.name == 'ＭＳ 明朝'
        assert obj.font.size_pt == 12.0

    def test_field_object(self):
        inner = (
            _make_tlv(0x07D1, struct.pack('<IIII', 50, 60, 200, 100))
            + _make_tlv(0x0BB9, struct.pack('<I', 108))
        )
        field_tlv = _make_tlv(0x03EC, inner)

        raw = _make_minimal_decompressed(objects_payload=field_tlv)
        lay = parse_lay_bytes(_make_lay_bytes(raw))

        assert len(lay.fields) == 1
        obj = lay.fields[0]
        assert obj.obj_type == ObjectType.FIELD
        assert obj.field_id == 108

    def test_line_object(self):
        inner = (
            _make_tlv(0x07D1, struct.pack('<II', 10, 100))
            + _make_tlv(0x07D2, struct.pack('<II', 500, 100))
        )
        line_tlv = _make_tlv(0x03E9, inner)

        raw = _make_minimal_decompressed(objects_payload=line_tlv)
        lay = parse_lay_bytes(_make_lay_bytes(raw))

        assert len(lay.lines) == 1
        obj = lay.lines[0]
        assert obj.obj_type == ObjectType.LINE
        assert obj.line_start == Point(10, 100)
        assert obj.line_end == Point(500, 100)


class TestRect:
    """座標データクラスのテスト。"""

    def test_rect_properties(self):
        r = Rect(10, 20, 100, 80)
        assert r.width == 90
        assert r.height == 60

    def test_point(self):
        p = Point(42, 99)
        assert p.x == 42
        assert p.y == 99


class TestFontInfo:
    """フォント情報のテスト。"""

    def test_default_values(self):
        f = FontInfo()
        assert f.name == ''
        assert f.size_pt == 10.0
        assert f.bold is False
        assert f.italic is False

    def test_bold_italic(self):
        f = FontInfo(bold=True, italic=True)
        assert f.bold is True
        assert f.italic is True


# ── 実ファイルテスト ─────────────────────────────────────────────────────────


@pytest.mark.skipif(
    not os.path.isfile(_SAMPLE_LAY),
    reason='サンプル .lay ファイルが見つかりません',
)
class TestRealFile:
    """実際の .lay ファイルをパースするテスト。"""

    @pytest.fixture
    def lay(self) -> LayFile:
        return parse_lay(_SAMPLE_LAY)

    def test_parse_succeeds(self, lay: LayFile):
        assert isinstance(lay, LayFile)

    def test_title_extracted(self, lay: LayFile):
        # タイトルには 'R' prefix byte がある可能性
        assert '小学校個票' in lay.title

    def test_version(self, lay: LayFile):
        assert lay.version == 1600

    def test_page_width(self, lay: LayFile):
        assert lay.page_width == 840  # A4 幅 = 210mm

    def test_object_count(self, lay: LayFile):
        assert len(lay.objects) >= 60  # 71 オブジェクト期待

    def test_has_fields(self, lay: LayFile):
        assert len(lay.fields) >= 14

    def test_has_labels(self, lay: LayFile):
        assert len(lay.labels) >= 18

    def test_has_lines(self, lay: LayFile):
        assert len(lay.lines) >= 20

    def test_field_108_is_name(self, lay: LayFile):
        name_fields = [f for f in lay.fields if f.field_id == 108]
        assert len(name_fields) == 1
        assert name_fields[0].rect is not None

    def test_label_has_shimei(self, lay: LayFile):
        shimei = [lb for lb in lay.labels if '氏' in lb.text]
        assert len(shimei) >= 1

    def test_label_font(self, lay: LayFile):
        for lb in lay.labels:
            if lb.font.name:
                assert 'Ｓ' in lb.font.name or '明朝' in lb.font.name
                break
        else:
            pytest.fail('フォント名を持つラベルが見つかりません')

    def test_lines_have_endpoints(self, lay: LayFile):
        for line in lay.lines:
            assert line.line_start is not None
            assert line.line_end is not None
