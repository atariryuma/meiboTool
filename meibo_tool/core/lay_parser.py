"""スズキ校務 .lay バイナリレイアウトファイルのパーサー

EXCMIDataContainer01 形式を解析し、レイアウトオブジェクトを抽出する。

フォーマット概要:
  - ヘッダー: "EXCMIDataContainer01" (UTF-16LE, 40 bytes)
  - uint32: 展開後サイズ
  - zlib 圧縮データ → TLV 再帰構造

座標単位: 0.25mm (A4 幅 = 840 単位 = 210mm)
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

# ── フィールド ID マッピング ──────────────────────────────────────────────────

FIELD_ID_MAP: dict[int, str] = {
    105: '組',
    107: '性別',
    108: '氏名',
    109: '氏名かな',
    400: '写真',
    603: '都道府県',
    604: '市区町村',
    607: '町番地',
    610: '生年月日',
    1500: '評定1',
    1501: '家庭環境',
    1502: '評定2',
    1504: '学級配慮',
    1505: '不適児童',
    1506: '欠席',
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


@dataclass
class Point:
    """座標点 (0.25mm 単位)。"""
    x: int
    y: int


@dataclass
class Rect:
    """矩形 (0.25mm 単位)。"""
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


@dataclass
class LayFile:
    """パース済み .lay ファイル。"""
    title: str = ''
    version: int = 0
    page_width: int = 840
    page_height: int = 1188
    objects: list[LayoutObject] = field(default_factory=list)

    @property
    def labels(self) -> list[LayoutObject]:
        return [o for o in self.objects if o.obj_type == ObjectType.LABEL]

    @property
    def fields(self) -> list[LayoutObject]:
        return [o for o in self.objects if o.obj_type == ObjectType.FIELD]

    @property
    def lines(self) -> list[LayoutObject]:
        return [o for o in self.objects if o.obj_type == ObjectType.LINE]


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


def _parse_object_list(payload: bytes) -> list[LayoutObject]:
    """タグ 1002 のオブジェクトリストをパースする。"""
    objects: list[LayoutObject] = []
    for tag, data in _iter_tlv(payload):
        if tag in (_TAG_OBJ_LINE, _TAG_OBJ_CONTAINER, _TAG_OBJ_LABEL,
                   _TAG_OBJ_FIELD):
            obj = _parse_object_block(tag, data)
            objects.append(obj)
    return objects


# ── コンテンツブロックパース ─────────────────────────────────────────────────


def _parse_content_block(payload: bytes) -> tuple[list[LayoutObject], int, int]:
    """タグ 1505 のメインコンテンツブロックをパースする。

    Returns:
        (objects, page_width, page_height)
    """
    objects: list[LayoutObject] = []
    page_w, page_h = 840, 1188
    found_object_list = False

    for tag, data in _iter_tlv(payload):
        if tag == _TAG_OBJ_CONTAINER and not found_object_list:
            # 最初の tag-1002 ブロック = オブジェクトリスト
            objects = _parse_object_list(data)
            found_object_list = True
        elif tag == _TAG_GEO_RECT and len(data) == 8:
            # ドキュメント末尾のページサイズ
            w, h = struct.unpack_from('<II', data)
            if w > 0 and h > 0:
                page_w, page_h = w, h

    return objects, page_w, page_h


# ── トップレベルパース ───────────────────────────────────────────────────────


def _parse_decompressed(data: bytes) -> LayFile:
    """展開済みデータをパースして LayFile を返す。"""
    lay = LayFile()

    # 先頭 6 バイト: version(uint16) + content_size(uint32)
    if len(data) < 6:
        raise ValueError('Decompressed data too short')
    lay.version = struct.unpack_from('<H', data, 0)[0]
    # content_size = struct.unpack_from('<I', data, 2)[0]  # 検証用

    # オフセット 6 から TLV 列
    for tag, payload in _iter_tlv(data, start=6):
        if tag == _TAG_DOC_TITLE and len(payload) >= 2:
            lay.title = payload.decode('utf-16-le', errors='replace')
        elif tag == _TAG_DOC_CONTENT:
            objs, pw, ph = _parse_content_block(payload)
            lay.objects = objs
            lay.page_width = pw
            lay.page_height = ph

    return lay


# ── 公開 API ─────────────────────────────────────────────────────────────────


def parse_lay(path: str) -> LayFile:
    """`.lay` ファイルをパースして LayFile を返す。

    Args:
        path: .lay ファイルのパス

    Returns:
        パース済み LayFile オブジェクト

    Raises:
        ValueError: ヘッダーが不正、またはデータが壊れている場合
    """
    with open(path, 'rb') as f:
        data = f.read()

    # ヘッダー検証
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

    # 展開後サイズ
    expected_size = struct.unpack_from('<I', data, _HEADER_LEN)[0]

    # zlib 展開
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

    return _parse_decompressed(decompressed)


def parse_lay_bytes(data: bytes) -> LayFile:
    """バイト列から .lay ファイルをパースする（テスト用）。"""
    if len(data) < _HEADER_LEN + 4:
        raise ValueError(f'Data too short: {len(data)} bytes')
    header = data[:_HEADER_LEN]
    if header != _MAGIC_BYTES:
        raise ValueError('Invalid .lay header')
    expected_size = struct.unpack_from('<I', data, _HEADER_LEN)[0]
    compressed = data[_HEADER_LEN + 4:]
    try:
        decompressed = zlib.decompress(compressed)
    except zlib.error as e:
        raise ValueError(f'zlib decompression failed: {e}') from e
    if len(decompressed) != expected_size:
        raise ValueError('Decompressed size mismatch')
    return _parse_decompressed(decompressed)
