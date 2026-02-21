"""スズキ校務 .lay バイナリレイアウトファイルのパーサー

EXCMIDataContainer01 形式を解析し、レイアウトオブジェクトを抽出する。

フォーマット概要:
  - ヘッダー: "EXCMIDataContainer01" (UTF-16LE, 40 bytes)
  - uint32: 展開後サイズ
  - zlib 圧縮データ → TLV 再帰構造

座標単位:
  - 旧形式（単一レイアウト）: 0.25mm/unit (A4 幅 = 840 単位 = 210mm)
  - 新形式（マルチレイアウト）: 0.1mm/unit (A4 幅 = 2100 単位 = 210mm)
  PaperLayout.unit_mm で座標単位を示す。
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field
from enum import IntEnum

# ── 定数 ─────────────────────────────────────────────────────────────────────

_MAGIC = 'EXCMIDataContainer01'
_MAGIC_BYTES = _MAGIC.encode('utf-16-le')  # 40 bytes
_HEADER_LEN = len(_MAGIC_BYTES)  # 40

# TLV タグ定数 ── ドキュメントレベル
_TAG_DOC_TITLE = 0x05DC   # 1500
_TAG_DOC_FLAG1 = 0x05DF   # 1503
_TAG_DOC_FLAG2 = 0x05E0   # 1504
_TAG_DOC_CONTENT = 0x05E1  # 1505
_TAG_DOC_LCID = 0x05E2    # 1506

# TLV タグ定数 ── オブジェクトレベル
_TAG_OBJ_STYLE = 0x03E8     # 1000
_TAG_OBJ_LINE = 0x03E9      # 1001 (LINE object / style flag)
_TAG_OBJ_CONTAINER = 0x03EA  # 1002 (GROUP/container / pen type)
_TAG_OBJ_LABEL = 0x03EB     # 1003 (LABEL object / line pattern)
_TAG_OBJ_FIELD = 0x03EC     # 1004 (FIELD object)
_TAG_OBJ_PROP5 = 0x03ED     # 1005
_TAG_OBJ_PROP6 = 0x03EE     # 1006
_TAG_OBJ_TABLE = 0x03EF     # 1007 (TABLE / 名簿オブジェクト)

# TLV タグ定数 ── マルチレイアウトコンテナ
_TAG_LAYOUT_ENTRY = 0x0640  # 1600: 追加レイアウトブロック

# TLV タグ定数 ── ジオメトリ
_TAG_GEO_FLAG = 0x07D0   # 2000
_TAG_GEO_RECT = 0x07D1   # 2001
_TAG_GEO_POINT2 = 0x07D2  # 2002
_TAG_GEO_PROP3 = 0x07D3   # 2003
_TAG_GEO_PROP4 = 0x07D4   # 2004

# TLV タグ定数 ── コンテンツ
_TAG_TEXT = 0x0BB9    # 3001 (テキスト / フィールドID)
_TAG_HALIGN = 0x0BBA  # 3002
_TAG_VALIGN = 0x0BBB  # 3003
_TAG_FONT = 0x0BBC    # 3004 (フォントブロック)
_TAG_PREFIX = 0x0BBD  # 3005
_TAG_SUFFIX = 0x0BBE  # 3006

# フォントブロック内タグ
_TAG_FONT_NAME = 0x03E8   # 1000
_TAG_FONT_STYLE = 0x03E9  # 1001
_TAG_FONT_SIZE = 0x03EA   # 1002

# ── ジオメトリ単位 ─────────────────────────────────────────────────────────────

_GEO_UNIT_MM = 0.1  # ジオメトリタグの座標単位 (0.1mm = 10 units/mm)

# ── フィールド ID マッピング ──────────────────────────────────────────────────

FIELD_ID_MAP: dict[int, str] = {
    # 基本情報
    104: '学年',
    105: '組',
    106: '出席番号',
    107: '性別',
    108: '氏名',
    109: '氏名かな',
    133: '行番号',
    137: '転出日',
    400: '写真',
    # 連絡先・住所
    601: '在校兄弟',
    602: '保護者名',
    603: '都道府県',
    604: '市区町村',
    607: '町番地',
    608: '緊急連絡先',
    610: '生年月日',
    # 公簿名
    685: '公簿名',
    686: '公簿名かな',
    # 学級編成
    1500: '評定1',
    1501: '家庭環境',
    1502: '評定2',
    1504: '学級配慮',
    1505: '不適児童',
    1506: '欠席',
    1515: '新学級1',
}


def resolve_field_name(field_id: int) -> str:
    """フィールドIDを論理名に変換する。未知IDは 'field_NNN' を返す。"""
    return FIELD_ID_MAP.get(field_id, f'field_{field_id}')


# ── データクラス ─────────────────────────────────────────────────────────────


class ObjectType(IntEnum):
    """レイアウトオブジェクトの種類。"""
    LINE = 1
    GROUP = 2
    LABEL = 3
    FIELD = 4
    TABLE = 5


@dataclass
class Point:
    """座標点 (PaperLayout.unit_mm 単位)。"""
    x: int
    y: int


@dataclass
class Rect:
    """矩形 (PaperLayout.unit_mm 単位)。"""
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top


@dataclass
class FontInfo:
    """フォント情報。"""
    name: str = ''
    size_pt: float = 10.0
    bold: bool = False
    italic: bool = False


@dataclass
class TableColumn:
    """テーブルオブジェクトのカラム定義。"""
    field_id: int = 0
    width: int = 1       # 相対幅
    h_align: int = 0     # 0=左, 1=中央, 2=右
    header: str = ''     # カラムヘッダー文字列


# ── 用紙・配置サイズ ────────────────────────────────────────────────────────

# 標準用紙サイズ (mm)
_PAPER_SIZES: dict[str, tuple[int, int]] = {
    'A3': (297, 420),
    'A4': (210, 297),
    'A5': (148, 210),
    'B4': (257, 364),
    'B5': (182, 257),
    'はがき': (100, 148),
}


@dataclass
class PaperLayout:
    """用紙配置情報（.lay のジオメトリタグから自動計算）。

    Attributes:
        mode: 0=全面（オブジェクトが用紙全体にわたる）、1=ラベル（1アイテム内に収まる）
        unit_mm: 座標1単位あたりの mm 数 (旧形式=0.25, 新形式=0.1)
        item_width_mm: アイテム幅 (mm)
        item_height_mm: アイテム高さ (mm)
        cols: 列数
        rows: 行数
        margin_left_mm: 左余白 (mm)
        margin_top_mm: 上余白 (mm)
        spacing_h_mm: 水平間隔 (mm)
        spacing_v_mm: 垂直間隔 (mm)
        paper_size: 自動検出された用紙サイズ ('A4' 等)
        orientation: 'portrait' or 'landscape'
    """
    mode: int = 0
    unit_mm: float = 0.25
    item_width_mm: float = 0.0
    item_height_mm: float = 0.0
    cols: int = 1
    rows: int = 1
    margin_left_mm: float = 0.0
    margin_top_mm: float = 0.0
    spacing_h_mm: float = 0.0
    spacing_v_mm: float = 0.0
    paper_size: str = ''
    orientation: str = ''

    @property
    def paper_width_mm(self) -> float:
        """用紙幅を mm で返す。"""
        return (self.margin_left_mm * 2
                + self.cols * self.item_width_mm
                + max(0, self.cols - 1) * self.spacing_h_mm)

    @property
    def paper_height_mm(self) -> float:
        """用紙高さを mm で返す。"""
        return (self.margin_top_mm * 2
                + self.rows * self.item_height_mm
                + max(0, self.rows - 1) * self.spacing_v_mm)


def _detect_paper(p: PaperLayout) -> None:
    """ジオメトリから用紙サイズと向きを自動判定する。"""
    total_w = p.paper_width_mm
    total_h = p.paper_height_mm

    best_name = ''
    best_orient = ''
    best_diff = float('inf')

    for name, (short, long) in _PAPER_SIZES.items():
        for orient, pw, ph in [
            ('portrait', short, long),
            ('landscape', long, short),
        ]:
            diff = abs(total_w - pw) + abs(total_h - ph)
            if diff < best_diff:
                best_diff = diff
                best_name = name
                best_orient = orient

    if best_diff < 20:  # 20mm 以内の誤差で一致とみなす
        p.paper_size = best_name
        p.orientation = best_orient


@dataclass
class LayoutObject:
    """1つのレイアウト要素。"""
    obj_type: ObjectType
    rect: Rect | None = None
    line_start: Point | None = None
    line_end: Point | None = None
    text: str = ''
    field_id: int = 0
    font: FontInfo = field(default_factory=FontInfo)
    h_align: int = 0
    v_align: int = 0
    prefix: str = ''
    suffix: str = ''
    table_columns: list[TableColumn] = field(default_factory=list)


@dataclass
class LayFile:
    """パース済み .lay ファイル。"""
    title: str = ''
    version: int = 0
    page_width: int = 840
    page_height: int = 1188
    objects: list[LayoutObject] = field(default_factory=list)
    paper: PaperLayout | None = None

    @property
    def labels(self) -> list[LayoutObject]:
        return [o for o in self.objects if o.obj_type == ObjectType.LABEL]

    @property
    def fields(self) -> list[LayoutObject]:
        return [o for o in self.objects if o.obj_type == ObjectType.FIELD]

    @property
    def lines(self) -> list[LayoutObject]:
        return [o for o in self.objects if o.obj_type == ObjectType.LINE]

    @property
    def tables(self) -> list[LayoutObject]:
        return [o for o in self.objects if o.obj_type == ObjectType.TABLE]


# ── オブジェクト生成ヘルパー ──────────────────────────────────────────────────

_DEFAULT_FONT_NAME = 'IPAmj明朝'


def new_label(
    left: int, top: int, right: int, bottom: int,
    text: str = '', font_name: str = _DEFAULT_FONT_NAME,
    font_size: float = 11.0, h_align: int = 1, v_align: int = 1,
) -> LayoutObject:
    """新しい LABEL オブジェクトを生成する。"""
    return LayoutObject(
        obj_type=ObjectType.LABEL,
        rect=Rect(left, top, right, bottom),
        text=text,
        font=FontInfo(font_name, font_size),
        h_align=h_align, v_align=v_align,
    )


def new_field(
    left: int, top: int, right: int, bottom: int,
    field_id: int, font_name: str = _DEFAULT_FONT_NAME,
    font_size: float = 11.0, h_align: int = 1, v_align: int = 1,
) -> LayoutObject:
    """新しい FIELD オブジェクトを生成する。"""
    return LayoutObject(
        obj_type=ObjectType.FIELD,
        rect=Rect(left, top, right, bottom),
        field_id=field_id,
        font=FontInfo(font_name, font_size),
        h_align=h_align, v_align=v_align,
    )


def new_line(x1: int, y1: int, x2: int, y2: int) -> LayoutObject:
    """新しい LINE オブジェクトを生成する。"""
    return LayoutObject(
        obj_type=ObjectType.LINE,
        line_start=Point(x1, y1),
        line_end=Point(x2, y2),
    )


# ── TLV パース ───────────────────────────────────────────────────────────────


def _read_tlv(data: bytes, offset: int) -> tuple[int, int, bytes, int]:
    """1つの TLV エントリを読み取る。

    Returns:
        (tag, length, payload, next_offset)
    """
    if offset + 6 > len(data):
        raise ValueError(f'TLV read beyond data at offset {offset}')
    tag = struct.unpack_from('<H', data, offset)[0]
    length = struct.unpack_from('<I', data, offset + 2)[0]
    end = offset + 6 + length
    if end > len(data):
        raise ValueError(
            f'TLV payload exceeds data: tag=0x{tag:04X}, '
            f'offset={offset}, length={length}, data_size={len(data)}'
        )
    payload = data[offset + 6:end]
    return tag, length, payload, end


def _iter_tlv(data: bytes, start: int = 0,
              end: int | None = None) -> list[tuple[int, bytes]]:
    """指定範囲内の TLV エントリをすべて読み取る。"""
    if end is None:
        end = len(data)
    entries: list[tuple[int, bytes]] = []
    pos = start
    while pos + 6 <= end:
        tag, length, payload, next_pos = _read_tlv(data, pos)
        if next_pos > end:
            break
        entries.append((tag, payload))
        pos = next_pos
    return entries


# ── フォントパース ───────────────────────────────────────────────────────────


def _parse_font(payload: bytes) -> FontInfo:
    """タグ 0x0BBC (3004) のフォントブロックをパースする。"""
    info = FontInfo()
    for tag, data in _iter_tlv(payload):
        if tag == _TAG_FONT_NAME and len(data) >= 2:
            info.name = data.decode('utf-16-le', errors='replace')
        elif tag == _TAG_FONT_STYLE and len(data) == 4:
            style = struct.unpack_from('<I', data)[0]
            # bit 0 = bold, bit 1 = italic (推定)
            info.bold = bool(style & 0x01)
            info.italic = bool(style & 0x02)
        elif tag == _TAG_FONT_SIZE and len(data) == 4:
            tenths = struct.unpack_from('<I', data)[0]
            info.size_pt = tenths / 10.0
    return info


# ── オブジェクトパース ───────────────────────────────────────────────────────


def _parse_object_block(outer_tag: int, payload: bytes) -> LayoutObject:
    """1つのオブジェクトブロックをパースする。

    outer_tag がオブジェクト種別を決定:
      0x03E9 (1001) = LINE
      0x03EA (1002) = GROUP
      0x03EB (1003) = LABEL
      0x03EC (1004) = FIELD
    """
    type_map = {
        _TAG_OBJ_LINE: ObjectType.LINE,
        _TAG_OBJ_CONTAINER: ObjectType.GROUP,
        _TAG_OBJ_LABEL: ObjectType.LABEL,
        _TAG_OBJ_FIELD: ObjectType.FIELD,
    }
    obj = LayoutObject(obj_type=type_map.get(outer_tag, ObjectType.LABEL))

    for tag, data in _iter_tlv(payload):
        if tag == _TAG_GEO_RECT:
            if len(data) == 16:
                left, top, right, bottom = struct.unpack_from('<IIII', data)
                obj.rect = Rect(left, top, right, bottom)
            elif len(data) == 8:
                x, y = struct.unpack_from('<II', data)
                obj.line_start = Point(x, y)
        elif tag == _TAG_GEO_POINT2 and len(data) == 8:
            x, y = struct.unpack_from('<II', data)
            obj.line_end = Point(x, y)
        elif tag == _TAG_TEXT:
            if obj.obj_type == ObjectType.LABEL and len(data) >= 2:
                obj.text = data.decode('utf-16-le', errors='replace')
            elif obj.obj_type == ObjectType.FIELD and len(data) >= 4:
                obj.field_id = struct.unpack_from('<I', data)[0]
        elif tag == _TAG_HALIGN and len(data) == 4:
            obj.h_align = struct.unpack_from('<i', data)[0]
        elif tag == _TAG_VALIGN and len(data) == 4:
            obj.v_align = struct.unpack_from('<i', data)[0]
        elif tag == _TAG_PREFIX and len(data) >= 2:
            obj.prefix = data.decode('utf-16-le', errors='replace')
        elif tag == _TAG_SUFFIX and len(data) >= 2:
            obj.suffix = data.decode('utf-16-le', errors='replace')
        elif tag == _TAG_FONT:
            obj.font = _parse_font(data)

    return obj


def _parse_table_columns(payload: bytes) -> list[TableColumn]:
    """TABLE オブジェクト内のカラム定義をパースする。

    TABLE (0x03EF) 内の 0x0BBC タグがカラム定義を含む。
    カラム内の TLV:
      0x03E8 = field_id (uint32)
      0x03E9 = 幅 (uint32)
      0x03EA = 配置 (uint32)
      0x03EB = ヘッダー文字列 (UTF-16LE)
    """
    columns: list[TableColumn] = []
    for tag, data in _iter_tlv(payload):
        if tag == _TAG_FONT:  # 0x0BBC = カラム定義（TABLE コンテキスト）
            col = TableColumn()
            for ct, cp in _iter_tlv(data):
                if ct == _TAG_FONT_NAME and len(cp) >= 4:
                    # TABLE 内では 0x03E8 = field_id
                    col.field_id = struct.unpack_from('<I', cp)[0]
                elif ct == _TAG_FONT_STYLE and len(cp) >= 4:
                    # TABLE 内では 0x03E9 = 幅
                    col.width = struct.unpack_from('<I', cp)[0]
                elif ct == _TAG_FONT_SIZE and len(cp) >= 4:
                    # TABLE 内では 0x03EA = 配置
                    col.h_align = struct.unpack_from('<I', cp)[0]
                elif ct == _TAG_OBJ_LABEL and len(cp) >= 2:
                    # 0x03EB = ヘッダー文字列
                    col.header = cp.decode('utf-16-le', errors='replace')
            columns.append(col)
    return columns


def _parse_table_object(payload: bytes) -> LayoutObject:
    """TABLE オブジェクト (0x03EF) をパースする。"""
    obj = LayoutObject(obj_type=ObjectType.TABLE)

    for tag, data in _iter_tlv(payload):
        if tag == _TAG_GEO_RECT:
            if len(data) == 16:
                left, top, right, bottom = struct.unpack_from('<IIII', data)
                obj.rect = Rect(left, top, right, bottom)
        elif tag == _TAG_FONT:
            # TABLE 内の 0x0BBC はカラム定義として扱う
            pass  # _parse_table_columns で一括処理
        elif tag == _TAG_TEXT and len(data) >= 2:
            obj.text = data.decode('utf-16-le', errors='replace')

    obj.table_columns = _parse_table_columns(payload)

    # テーブルのフォント情報（全体フォント）を取得
    for tag, data in _iter_tlv(payload):
        if tag == _TAG_GEO_PROP3 and len(data) == 4:
            # テーブルフォントサイズ候補
            pass

    return obj


def _parse_object_list(payload: bytes) -> list[LayoutObject]:
    """タグ 1002 のオブジェクトリストをパースする。

    GROUP (CONTAINER) オブジェクトは再帰的に子要素を展開し、
    GROUP 自身の rect は外枠として 4 本の LINE に変換する。
    TABLE オブジェクトはカラム定義を含むテーブルとしてパースする。
    """
    objects: list[LayoutObject] = []
    for tag, data in _iter_tlv(payload):
        if tag == _TAG_OBJ_CONTAINER:
            # GROUP: 子オブジェクトを再帰的に展開
            children = _parse_object_list(data)
            objects.extend(children)
            # GROUP 自身の rect を外枠 LINE に変換
            obj = _parse_object_block(tag, data)
            if obj.rect is not None:
                r = obj.rect
                for x1, y1, x2, y2 in [
                    (r.left, r.top, r.right, r.top),       # 上辺
                    (r.left, r.bottom, r.right, r.bottom),  # 下辺
                    (r.left, r.top, r.left, r.bottom),      # 左辺
                    (r.right, r.top, r.right, r.bottom),    # 右辺
                ]:
                    objects.append(LayoutObject(
                        obj_type=ObjectType.LINE,
                        line_start=Point(x1, y1),
                        line_end=Point(x2, y2),
                    ))
        elif tag == _TAG_OBJ_TABLE:
            obj = _parse_table_object(data)
            objects.append(obj)
        elif tag in (_TAG_OBJ_LINE, _TAG_OBJ_LABEL, _TAG_OBJ_FIELD):
            obj = _parse_object_block(tag, data)
            objects.append(obj)
    return objects


# ── コンテンツブロックパース ─────────────────────────────────────────────────


def _parse_content_block(
    payload: bytes,
) -> tuple[list[LayoutObject], int, int, PaperLayout | None]:
    """タグ 1505 のメインコンテンツブロックをパースする。

    Returns:
        (objects, page_width, page_height, paper_layout)
    """
    objects: list[LayoutObject] = []
    page_w, page_h = 840, 1188
    found_object_list = False

    # ジオメトリタグ値 (生値)
    geo_mode: int | None = None
    geo_item: tuple[int, int] | None = None  # (w, h)
    geo_count: tuple[int, int] | None = None  # (cols, rows)
    geo_spacing: tuple[int, int] | None = None  # (h, v)
    geo_margin: tuple[int, int] | None = None  # (left, top)

    for tag, data in _iter_tlv(payload):
        if tag == _TAG_OBJ_CONTAINER and not found_object_list:
            # 最初の tag-1002 ブロック = オブジェクトリスト
            objects = _parse_object_list(data)
            found_object_list = True
        elif tag == _TAG_GEO_FLAG:
            # 0x07D0: モード (1バイト値)
            if len(data) >= 1:
                geo_mode = data[0]
        elif tag == _TAG_GEO_RECT and len(data) == 8:
            # 0x07D1: アイテムサイズ (w, h)
            w, h = struct.unpack_from('<II', data)
            if w > 0 and h > 0:
                geo_item = (w, h)
                page_w, page_h = w, h
        elif tag == _TAG_GEO_POINT2 and len(data) == 8:
            # 0x07D2: 個数 (cols, rows)
            c, r = struct.unpack_from('<II', data)
            geo_count = (c, r)
        elif tag == _TAG_GEO_PROP3 and len(data) == 8:
            # 0x07D3: 間隔 (horizontal, vertical)
            geo_spacing = struct.unpack_from('<II', data)
        elif tag == _TAG_GEO_PROP4 and len(data) == 8:
            # 0x07D4: 左上余白 (left, top)
            geo_margin = struct.unpack_from('<II', data)

    # PaperLayout を構築
    paper: PaperLayout | None = None
    if geo_item is not None and geo_count is not None:
        cols, rows = geo_count
        iw, ih = geo_item
        sp_h, sp_v = geo_spacing or (0, 0)
        ml, mt = geo_margin or (0, 0)

        paper = PaperLayout(
            mode=geo_mode if geo_mode is not None else 0,
            unit_mm=_GEO_UNIT_MM,
            item_width_mm=iw * _GEO_UNIT_MM,
            item_height_mm=ih * _GEO_UNIT_MM,
            cols=cols,
            rows=rows,
            margin_left_mm=ml * _GEO_UNIT_MM,
            margin_top_mm=mt * _GEO_UNIT_MM,
            spacing_h_mm=sp_h * _GEO_UNIT_MM,
            spacing_v_mm=sp_v * _GEO_UNIT_MM,
        )
        _detect_paper(paper)

        # mode=0 (全面): オブジェクトは用紙全体に配置されるため
        # page_width/page_height を用紙サイズに拡張する
        if paper.mode == 0:
            paper_w_units = int(paper.paper_width_mm / _GEO_UNIT_MM)
            paper_h_units = int(paper.paper_height_mm / _GEO_UNIT_MM)
            if paper_w_units > page_w:
                page_w = paper_w_units
            if paper_h_units > page_h:
                page_h = paper_h_units

    return objects, page_w, page_h, paper


# ── トップレベルパース ───────────────────────────────────────────────────────


def _parse_layout_from_tlv(
    entries: list[tuple[int, bytes]],
) -> LayFile:
    """TLV エントリ列から1つの LayFile をパースする。

    entries は (tag, payload) のリスト。
    タイトル (0x05DC) とコンテンツ (0x05E1) を含む。
    """
    lay = LayFile()
    for tag, payload in entries:
        if tag == _TAG_DOC_TITLE and len(payload) >= 2:
            lay.title = payload.decode('utf-16-le', errors='replace')
        elif tag == _TAG_DOC_CONTENT:
            objs, pw, ph, paper = _parse_content_block(payload)
            lay.objects = objs
            lay.page_width = pw
            lay.page_height = ph
            lay.paper = paper
    return lay


def _parse_decompressed(data: bytes) -> LayFile:
    """展開済みデータをパースして LayFile を返す（メインレイアウトのみ）。"""
    if len(data) < 6:
        raise ValueError('Decompressed data too short')

    lay = LayFile()
    lay.version = struct.unpack_from('<H', data, 0)[0]

    # メインレイアウト用の TLV エントリを収集
    main_entries: list[tuple[int, bytes]] = []
    for tag, payload in _iter_tlv(data, start=6):
        if tag in (_TAG_DOC_TITLE, _TAG_DOC_CONTENT,
                   _TAG_DOC_FLAG1, _TAG_DOC_FLAG2, _TAG_DOC_LCID):
            main_entries.append((tag, payload))

    lay = _parse_layout_from_tlv(main_entries)
    lay.version = struct.unpack_from('<H', data, 0)[0]
    return lay


def _parse_decompressed_multi(data: bytes) -> list[LayFile]:
    """展開済みデータから全レイアウトをパースする。"""
    if len(data) < 6:
        raise ValueError('Decompressed data too short')

    version = struct.unpack_from('<H', data, 0)[0]
    layouts: list[LayFile] = []

    # メインレイアウト用 TLV + 追加レイアウト (0x0640)
    main_entries: list[tuple[int, bytes]] = []
    for tag, payload in _iter_tlv(data, start=6):
        if tag == _TAG_LAYOUT_ENTRY:
            # 追加レイアウトブロック: 内部に 0x05DC + 0x05E1 を含む
            sub_entries = _iter_tlv(payload)
            sub_lay = _parse_layout_from_tlv(sub_entries)
            sub_lay.version = version
            layouts.append(sub_lay)
        elif tag in (_TAG_DOC_TITLE, _TAG_DOC_CONTENT,
                     _TAG_DOC_FLAG1, _TAG_DOC_FLAG2, _TAG_DOC_LCID):
            main_entries.append((tag, payload))

    # メインレイアウトを先頭に挿入
    main_lay = _parse_layout_from_tlv(main_entries)
    main_lay.version = version
    layouts.insert(0, main_lay)

    return layouts


# ── 公開 API ─────────────────────────────────────────────────────────────────


def _decompress_lay(data: bytes) -> bytes:
    """生のバイト列からヘッダー検証 + zlib 展開を行う。"""
    if len(data) < _HEADER_LEN + 4:
        raise ValueError(
            f'File too short: {len(data)} bytes '
            f'(minimum {_HEADER_LEN + 4})'
        )
    header = data[:_HEADER_LEN]
    if header != _MAGIC_BYTES:
        try:
            header_text = header.decode('utf-16-le', errors='replace')
        except Exception:
            header_text = header.hex()
        raise ValueError(f'Invalid .lay header: {header_text!r}')

    expected_size = struct.unpack_from('<I', data, _HEADER_LEN)[0]
    compressed = data[_HEADER_LEN + 4:]
    try:
        decompressed = zlib.decompress(compressed)
    except zlib.error as e:
        raise ValueError(f'zlib decompression failed: {e}') from e

    if len(decompressed) != expected_size:
        raise ValueError(
            f'Decompressed size mismatch: '
            f'expected {expected_size}, got {len(decompressed)}'
        )
    return decompressed


def parse_lay(path: str) -> LayFile:
    """`.lay` ファイルをパースして LayFile を返す（メインレイアウトのみ）。

    Args:
        path: .lay ファイルのパス

    Returns:
        パース済み LayFile オブジェクト

    Raises:
        ValueError: ヘッダーが不正、またはデータが壊れている場合
    """
    with open(path, 'rb') as f:
        data = f.read()
    return _parse_decompressed(_decompress_lay(data))


def parse_lay_multi(path: str) -> list[LayFile]:
    """`.lay` ファイルから全レイアウトをパースする。

    マルチレイアウト対応: メインレイアウト + 追加レイアウト (0x0640 ブロック)
    をすべて返す。単一レイアウトのファイルでも 1 要素のリストを返す。

    Args:
        path: .lay ファイルのパス

    Returns:
        パース済み LayFile のリスト（メインが先頭）
    """
    with open(path, 'rb') as f:
        data = f.read()
    return _parse_decompressed_multi(_decompress_lay(data))


def parse_lay_bytes(data: bytes) -> LayFile:
    """バイト列から .lay ファイルをパースする（テスト用）。"""
    return _parse_decompressed(_decompress_lay(data))
