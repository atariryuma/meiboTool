"""lay_parser.py のユニットテスト

テスト対象:
  - ヘッダー検証
  - zlib 展開
  - TLV パース
  - オブジェクト分類（LABEL / FIELD / LINE / GROUP / TABLE）
  - フィールド ID マッピング
  - 座標・フォントパース
  - PaperLayout 自動検出
  - parse_lay_multi マルチレイアウト
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
    LayoutObject,
    ObjectType,
    PaperLayout,
    Point,
    Rect,
    TableColumn,
    _detect_paper,
    new_image,
    parse_lay,
    parse_lay_bytes,
    parse_lay_multi,
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


# ── PaperLayout テスト ──────────────────────────────────────────────────────


class TestPaperLayout:
    """PaperLayout の自動検出テスト。"""

    def test_a4_portrait_detection(self):
        p = PaperLayout(
            item_width_mm=91.0, item_height_mm=55.0,
            cols=2, rows=5,
            margin_left_mm=14.0, margin_top_mm=11.0,
        )
        _detect_paper(p)
        assert p.paper_size == 'A4'
        assert p.orientation == 'portrait'

    def test_a4_portrait_label_mode(self):
        """学級編成用個票: A4 縦, 2×1"""
        p = PaperLayout(
            item_width_mm=84.0, item_height_mm=272.0,
            cols=2, rows=1,
            margin_left_mm=14.0, margin_top_mm=11.0,
            spacing_h_mm=10.0,
        )
        _detect_paper(p)
        assert p.paper_size == 'A4'
        assert p.orientation == 'portrait'

    def test_a3_landscape_detection(self):
        """机用氏名ラベル: A3 横, 2×5"""
        p = PaperLayout(
            item_width_mm=175.0, item_height_mm=55.0,
            cols=2, rows=5,
            margin_left_mm=35.0, margin_top_mm=11.0,
        )
        _detect_paper(p)
        # 紙幅 = 35*2 + 2*175 = 420mm, 高さ = 11*2 + 5*55 = 297mm → A3 横
        assert p.paper_size == 'A3'
        assert p.orientation == 'landscape'

    def test_paper_width_mm(self):
        p = PaperLayout(
            item_width_mm=91.0, item_height_mm=55.0,
            cols=2, rows=5,
            margin_left_mm=14.0, margin_top_mm=11.0,
            spacing_h_mm=0.0,
        )
        assert p.paper_width_mm == pytest.approx(210.0)
        assert p.paper_height_mm == pytest.approx(297.0)

    def test_paper_size_with_spacing(self):
        p = PaperLayout(
            item_width_mm=84.0, item_height_mm=272.0,
            cols=2, rows=1,
            margin_left_mm=14.0, margin_top_mm=11.0,
            spacing_h_mm=10.0, spacing_v_mm=0.0,
        )
        # Width: 14 + 2*84 + 10 + 14 = 206mm
        assert p.paper_width_mm == pytest.approx(206.0)
        # Height: 11 + 272 + 11 = 294mm
        assert p.paper_height_mm == pytest.approx(294.0)

    def test_no_match_too_large(self):
        """標準用紙に収まらない巨大サイズは検出しない。"""
        p = PaperLayout(
            item_width_mm=500.0, item_height_mm=500.0,
            cols=1, rows=1,
            margin_left_mm=5.0, margin_top_mm=5.0,
        )
        _detect_paper(p)
        assert p.paper_size == ''
        assert p.orientation == ''

    def test_unit_mm_default(self):
        p = PaperLayout()
        assert p.unit_mm == 0.25


# ── TableColumn テスト ──────────────────────────────────────────────────────


class TestTableColumn:
    """TableColumn データクラスのテスト。"""

    def test_default_values(self):
        col = TableColumn()
        assert col.field_id == 0
        assert col.width == 1
        assert col.h_align == 0
        assert col.header == ''

    def test_custom_values(self):
        col = TableColumn(field_id=108, width=32, h_align=1, header='氏名')
        assert col.field_id == 108
        assert col.header == '氏名'


# ── TABLE オブジェクトパーステスト ──────────────────────────────────────────


class TestTableObjectParse:
    """TABLE オブジェクト (0x03EF) のパーステスト。"""

    def _make_table_tlv(self, columns: list[dict]) -> bytes:
        """TABLE オブジェクトの TLV バイト列を構築する。"""
        # TABLE rect
        rect_data = struct.pack('<IIII', 100, 200, 800, 600)
        inner = _make_tlv(0x07D1, rect_data)

        # 各カラムを 0x0BBC タグ内に構築
        for col in columns:
            col_inner = b''
            col_inner += _make_tlv(0x03E8, struct.pack('<I', col['field_id']))
            col_inner += _make_tlv(0x03E9, struct.pack('<I', col.get('width', 10)))
            col_inner += _make_tlv(0x03EA, struct.pack('<I', col.get('h_align', 0)))
            if 'header' in col:
                col_inner += _make_tlv(0x03EB, col['header'].encode('utf-16-le'))
            inner += _make_tlv(0x0BBC, col_inner)

        return _make_tlv(0x03EF, inner)

    def test_table_with_columns(self):
        columns = [
            {'field_id': 108, 'width': 32, 'h_align': 0, 'header': '氏名'},
            {'field_id': 107, 'width': 8, 'h_align': 1, 'header': '性別'},
        ]
        table_tlv = self._make_table_tlv(columns)

        # オブジェクトリストに TABLE を入れたレイアウト
        raw = _make_minimal_decompressed(objects_payload=table_tlv)
        lay = parse_lay_bytes(_make_lay_bytes(raw))

        assert len(lay.tables) == 1
        tbl = lay.tables[0]
        assert tbl.obj_type == ObjectType.TABLE
        assert tbl.rect == Rect(100, 200, 800, 600)
        assert len(tbl.table_columns) == 2

        col0 = tbl.table_columns[0]
        assert col0.field_id == 108
        assert col0.width == 32
        assert col0.h_align == 0
        assert col0.header == '氏名'

        col1 = tbl.table_columns[1]
        assert col1.field_id == 107
        assert col1.header == '性別'

    def test_table_empty_columns(self):
        table_tlv = _make_tlv(0x03EF, _make_tlv(0x07D1, struct.pack('<IIII', 0, 0, 100, 50)))
        raw = _make_minimal_decompressed(objects_payload=table_tlv)
        lay = parse_lay_bytes(_make_lay_bytes(raw))

        assert len(lay.tables) == 1
        assert lay.tables[0].table_columns == []

    def test_tables_property(self):
        """LayFile.tables プロパティのテスト。"""
        lay = LayFile(objects=[
            LayoutObject(obj_type=ObjectType.LABEL, text='x'),
            LayoutObject(obj_type=ObjectType.TABLE, table_columns=[
                TableColumn(field_id=108, header='氏名'),
            ]),
        ])
        assert len(lay.tables) == 1
        assert lay.tables[0].table_columns[0].header == '氏名'


# ── FIELD_ID_MAP 拡充テスト ──────────────────────────────────────────────────


class TestFieldIdMapExpanded:
    """拡充された FIELD_ID_MAP のテスト。"""

    def test_map_has_at_least_25_entries(self):
        assert len(FIELD_ID_MAP) >= 25

    def test_new_entries_mapped(self):
        assert resolve_field_name(104) == '学年'
        assert resolve_field_name(106) == '出席番号'
        assert resolve_field_name(133) == '行番号'
        assert resolve_field_name(137) == '転出日'
        assert resolve_field_name(601) == '在校兄弟'
        assert resolve_field_name(602) == '保護者名'
        assert resolve_field_name(608) == '緊急連絡先'
        assert resolve_field_name(685) == '公簿名'
        assert resolve_field_name(686) == '公簿名かな'
        assert resolve_field_name(1515) == '新学級1'

    def test_existing_entries_unchanged(self):
        assert resolve_field_name(105) == '組'
        assert resolve_field_name(108) == '氏名'
        assert resolve_field_name(610) == '生年月日'

    def test_corrected_field_ids(self):
        """修正後の field_id マッピングを検証。"""
        assert resolve_field_name(101) == '学年'
        assert resolve_field_name(102) == '組'
        assert resolve_field_name(110) == '年度和暦'
        assert resolve_field_name(134) == '年度'


class TestFieldDisplayMapCorrected:
    """修正後の FIELD_DISPLAY_MAP テスト。"""

    def test_display_101(self):
        from core.lay_parser import resolve_field_display
        assert resolve_field_display(101) == '学年'

    def test_display_102(self):
        from core.lay_parser import resolve_field_display
        assert resolve_field_display(102) == '学級'

    def test_display_110(self):
        from core.lay_parser import resolve_field_display
        assert resolve_field_display(110) == '年度（和暦）'

    def test_display_134(self):
        from core.lay_parser import resolve_field_display
        assert resolve_field_display(134) == '年度'


# ── マルチレイアウト .lay 実ファイルテスト ────────────────────────────────────

_MULTI_LAY = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
    '名簿レイアウト_20260221.lay',
)


@pytest.mark.skipif(
    not os.path.isfile(_MULTI_LAY),
    reason='マルチレイアウト .lay ファイルが見つかりません',
)
class TestMultiLayoutRealFile:
    """マルチレイアウト .lay ファイルの実ファイルテスト。"""

    @pytest.fixture
    def layouts(self) -> list[LayFile]:
        return parse_lay_multi(_MULTI_LAY)

    def test_parse_multi_count(self, layouts: list[LayFile]):
        assert len(layouts) == 27

    def test_main_layout_title(self, layouts: list[LayFile]):
        assert '家庭調査票' in layouts[0].title

    def test_all_layouts_have_paper(self, layouts: list[LayFile]):
        for lay in layouts:
            assert lay.paper is not None

    def test_all_layouts_have_objects(self, layouts: list[LayFile]):
        # 1年生用机ラベル (index 10) は 0 オブジェクトだが他は > 0
        has_objects = sum(1 for lay in layouts if len(lay.objects) > 0)
        assert has_objects >= 25

    def test_mode0_paper_size_expanded(self, layouts: list[LayFile]):
        """mode=0 レイアウトの page_width/page_height が用紙サイズに拡張される。"""
        main = layouts[0]
        assert main.paper is not None
        assert main.paper.mode == 0
        # A4 portrait in 0.1mm units: 2100×2970
        assert main.page_width == 2100
        assert main.page_height == 2970

    def test_mode1_item_size(self, layouts: list[LayFile]):
        """mode=1 レイアウトの page_width/page_height がアイテムサイズ。"""
        kojihyo = layouts[24]  # 学級編成用個票
        assert kojihyo.paper is not None
        assert kojihyo.paper.mode == 1
        assert kojihyo.page_width == 840
        assert kojihyo.page_height == 2720

    def test_a4_portrait_detected(self, layouts: list[LayFile]):
        main = layouts[0]
        assert main.paper.paper_size == 'A4'
        assert main.paper.orientation == 'portrait'

    def test_a3_landscape_detected(self, layouts: list[LayFile]):
        desk_label = layouts[7]  # 机用氏名ラベル
        assert desk_label.paper.paper_size == 'A3'
        assert desk_label.paper.orientation == 'landscape'

    def test_doc_flag2_landscape_shuryou(self, layouts: list[LayFile]):
        """修了台帳: mode=0 + DOC_FLAG2=2 → A4 横 (2970×2100)。"""
        daicho = layouts[18]
        assert daicho.paper is not None
        assert daicho.paper.mode == 0
        assert daicho.paper.orientation == 'landscape'
        assert daicho.page_width == 2970
        assert daicho.page_height == 2100

    def test_doc_flag2_landscape_sotsugyou(self, layouts: list[LayFile]):
        """卒業台帳: mode=0 + DOC_FLAG2=2 → A4 横 (2970×2100)。"""
        sotsugyou = layouts[21]
        assert sotsugyou.paper is not None
        assert sotsugyou.paper.mode == 0
        assert sotsugyou.paper.orientation == 'landscape'
        assert sotsugyou.page_width == 2970
        assert sotsugyou.page_height == 2100

    def test_doc_flag2_portrait_unchanged(self, layouts: list[LayFile]):
        """家庭調査票: mode=0 + DOC_FLAG2=1 → A4 縦 (2100×2970) のまま。"""
        main = layouts[0]
        assert main.paper.mode == 0
        assert main.paper.orientation == 'portrait'
        assert main.page_width == 2100
        assert main.page_height == 2970

    def test_paper_size_flag_a3(self, layouts: list[LayFile]):
        """1年生用机ラベル: mode=0 + paper_size_flag=0 → A3 縦 (2970×4200)。"""
        label_1nen = layouts[10]
        assert '1年生' in label_1nen.title
        assert label_1nen.paper is not None
        assert label_1nen.paper.mode == 0
        assert label_1nen.paper.paper_size == 'A3'
        assert label_1nen.paper.orientation == 'portrait'
        assert label_1nen.page_width == 2970
        assert label_1nen.page_height == 4200

    def test_paper_size_flag_b4(self, layouts: list[LayFile]):
        """複数学級名列表: mode=0 + paper_size_flag=1 → B4 横 (3640×2570)。"""
        fukusu = layouts[17]
        assert '複数学級' in fukusu.title
        assert fukusu.paper is not None
        assert fukusu.paper.mode == 0
        assert fukusu.paper.paper_size == 'B4'
        assert fukusu.paper.orientation == 'landscape'
        assert fukusu.page_width == 3640
        assert fukusu.page_height == 2570

    def test_table_found_in_shuryou_daicho(self, layouts: list[LayFile]):
        """修了台帳に TABLE オブジェクト (10カラム) が含まれる。"""
        daicho = layouts[18]  # 修了台帳
        assert len(daicho.tables) == 1
        tbl = daicho.tables[0]
        assert len(tbl.table_columns) == 10

    def test_table_column_headers(self, layouts: list[LayFile]):
        """修了台帳のカラムヘッダーが正しくパースされる。"""
        daicho = layouts[18]
        tbl = daicho.tables[0]
        headers = [c.header for c in tbl.table_columns]
        assert '氏名' in headers
        assert 'ふりがな' in headers
        assert '性別' in headers
        assert '生年月日' in headers

    def test_table_column_field_ids(self, layouts: list[LayFile]):
        """修了台帳のカラム field_id が正しい。"""
        daicho = layouts[18]
        tbl = daicho.tables[0]
        fids = [c.field_id for c in tbl.table_columns]
        assert 108 in fids  # 氏名
        assert 107 in fids  # 性別

    def test_table_has_rect(self, layouts: list[LayFile]):
        daicho = layouts[18]
        tbl = daicho.tables[0]
        assert tbl.rect is not None
        assert tbl.rect.width > 0
        assert tbl.rect.height > 0

    def test_unit_mm_is_01(self, layouts: list[LayFile]):
        """マルチレイアウトの座標単位は 0.1mm。"""
        for lay in layouts:
            assert lay.paper.unit_mm == pytest.approx(0.1)

    def test_parse_lay_single_returns_first(self):
        """parse_lay() は最初のレイアウトだけを返す。"""
        lay = parse_lay(_MULTI_LAY)
        assert '家庭調査票' in lay.title


# ── LayoutObject 新フィールドテスト ──────────────────────────────────────────


class TestLayoutObjectTableColumns:
    """LayoutObject.table_columns のテスト。"""

    def test_default_empty_list(self):
        obj = LayoutObject(obj_type=ObjectType.LABEL)
        assert obj.table_columns == []

    def test_table_with_columns(self):
        cols = [TableColumn(field_id=108, header='氏名')]
        obj = LayoutObject(
            obj_type=ObjectType.TABLE,
            rect=Rect(0, 0, 100, 50),
            table_columns=cols,
        )
        assert len(obj.table_columns) == 1
        assert obj.table_columns[0].header == '氏名'


# ── new_image ────────────────────────────────────────────────────────────────


class TestNewImage:
    """new_image() ヘルパーのテスト。"""

    def test_creates_image_object(self) -> None:
        obj = new_image(10, 20, 200, 300)
        assert obj.obj_type == ObjectType.IMAGE
        assert obj.image is not None

    def test_rect_set_correctly(self) -> None:
        obj = new_image(10, 20, 200, 300)
        assert obj.rect == Rect(10, 20, 200, 300)
        assert obj.image.rect == (10, 20, 200, 300)

    def test_embedded_image_data(self) -> None:
        data = b'\x89PNG\r\n\x1a\n...'
        obj = new_image(0, 0, 100, 100, image_data=data, original_path='/test.png')
        assert obj.image.image_data == data
        assert obj.image.original_path == '/test.png'
