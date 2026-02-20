"""LayFile ↔ JSON シリアライザー

LayFile オブジェクトを人間可読な JSON 形式で保存・読込する。
.lay バイナリの代替として、エディターのネイティブ保存形式に使用する。

使用方法:
    from core.lay_serializer import save_layout, load_layout
    save_layout(lay, 'output.json')
    lay = load_layout('output.json')
"""

from __future__ import annotations

import json
from pathlib import Path

from core.lay_parser import (
    FontInfo,
    LayFile,
    LayoutObject,
    ObjectType,
    Point,
    Rect,
)

# ── 定数 ─────────────────────────────────────────────────────────────────────

_FORMAT_KEY = 'meibo_layout_v1'


# ── シリアライズ ─────────────────────────────────────────────────────────────


def _rect_to_list(r: Rect) -> list[int]:
    return [r.left, r.top, r.right, r.bottom]


def _point_to_list(p: Point) -> list[int]:
    return [p.x, p.y]


def _font_to_dict(f: FontInfo) -> dict:
    d: dict = {'name': f.name, 'size_pt': f.size_pt}
    if f.bold:
        d['bold'] = True
    if f.italic:
        d['italic'] = True
    return d


def _object_to_dict(obj: LayoutObject) -> dict:
    """LayoutObject を JSON 互換の dict に変換する。"""
    d: dict = {'type': obj.obj_type.name}

    if obj.rect is not None:
        d['rect'] = _rect_to_list(obj.rect)
    if obj.line_start is not None:
        d['line_start'] = _point_to_list(obj.line_start)
    if obj.line_end is not None:
        d['line_end'] = _point_to_list(obj.line_end)
    if obj.text:
        d['text'] = obj.text
    if obj.field_id:
        d['field_id'] = obj.field_id
    if obj.font.name or obj.font.size_pt != 10.0:
        d['font'] = _font_to_dict(obj.font)
    if obj.h_align:
        d['h_align'] = obj.h_align
    if obj.v_align:
        d['v_align'] = obj.v_align
    if obj.prefix:
        d['prefix'] = obj.prefix
    if obj.suffix:
        d['suffix'] = obj.suffix

    return d


def layfile_to_dict(lay: LayFile) -> dict:
    """LayFile を JSON 互換の dict に変換する。"""
    return {
        'format': _FORMAT_KEY,
        'title': lay.title,
        'version': lay.version,
        'page_width': lay.page_width,
        'page_height': lay.page_height,
        'objects': [_object_to_dict(obj) for obj in lay.objects],
    }


# ── デシリアライズ ───────────────────────────────────────────────────────────


def _list_to_rect(data: list[int]) -> Rect:
    return Rect(data[0], data[1], data[2], data[3])


def _list_to_point(data: list[int]) -> Point:
    return Point(data[0], data[1])


def _dict_to_font(data: dict) -> FontInfo:
    return FontInfo(
        name=data.get('name', ''),
        size_pt=data.get('size_pt', 10.0),
        bold=data.get('bold', False),
        italic=data.get('italic', False),
    )


_TYPE_MAP = {t.name: t for t in ObjectType}


def _dict_to_object(d: dict) -> LayoutObject:
    """dict から LayoutObject を復元する。"""
    type_name = d.get('type', 'LABEL')
    obj_type = _TYPE_MAP.get(type_name, ObjectType.LABEL)

    obj = LayoutObject(obj_type=obj_type)

    if 'rect' in d:
        obj.rect = _list_to_rect(d['rect'])
    if 'line_start' in d:
        obj.line_start = _list_to_point(d['line_start'])
    if 'line_end' in d:
        obj.line_end = _list_to_point(d['line_end'])
    if 'text' in d:
        obj.text = d['text']
    if 'field_id' in d:
        obj.field_id = d['field_id']
    if 'font' in d:
        obj.font = _dict_to_font(d['font'])
    if 'h_align' in d:
        obj.h_align = d['h_align']
    if 'v_align' in d:
        obj.v_align = d['v_align']
    if 'prefix' in d:
        obj.prefix = d['prefix']
    if 'suffix' in d:
        obj.suffix = d['suffix']

    return obj


def dict_to_layfile(data: dict) -> LayFile:
    """dict から LayFile を復元する。"""
    fmt = data.get('format', '')
    if fmt != _FORMAT_KEY:
        raise ValueError(
            f"Unknown format: {fmt!r} (expected '{_FORMAT_KEY}')"
        )

    return LayFile(
        title=data.get('title', ''),
        version=data.get('version', 0),
        page_width=data.get('page_width', 840),
        page_height=data.get('page_height', 1188),
        objects=[_dict_to_object(o) for o in data.get('objects', [])],
    )


# ── 公開 API ─────────────────────────────────────────────────────────────────


def save_layout(lay: LayFile, path: str) -> None:
    """LayFile を JSON ファイルに保存する。"""
    out_dir = Path(path).parent
    if out_dir != Path('.'):
        out_dir.mkdir(parents=True, exist_ok=True)
    data = layfile_to_dict(lay)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_layout(path: str) -> LayFile:
    """JSON ファイルから LayFile を読み込む。"""
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    return dict_to_layfile(data)
