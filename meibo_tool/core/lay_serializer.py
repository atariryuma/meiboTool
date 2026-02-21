"""LayFile ↔ JSON シリアライザー

LayFile オブジェクトを人間可読な JSON 形式で保存・読込する。
.lay バイナリの代替として、エディターのネイティブ保存形式に使用する。

使用方法:
    from core.lay_serializer import save_layout, load_layout
    save_layout(lay, 'output.json')
    lay = load_layout('output.json')
"""

from __future__ import annotations

import base64
import json
import tempfile
from pathlib import Path

from core.lay_parser import (
    EmbeddedImage,
    FontInfo,
    LayFile,
    LayoutObject,
    MeiboArea,
    ObjectType,
    PaperLayout,
    Point,
    Rect,
    TableColumn,
)

# ── 定数 ─────────────────────────────────────────────────────────────────────

_FORMAT_V1 = 'meibo_layout_v1'
_FORMAT_V2 = 'meibo_layout_v2'
_SUPPORTED_FORMATS = {_FORMAT_V1, _FORMAT_V2}


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


def _table_column_to_dict(col: TableColumn) -> dict:
    d: dict = {'field_id': col.field_id, 'width': col.width}
    if col.h_align:
        d['h_align'] = col.h_align
    if col.header:
        d['header'] = col.header
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
    if obj.table_columns:
        d['table_columns'] = [_table_column_to_dict(c) for c in obj.table_columns]
    if obj.meibo is not None:
        m = obj.meibo
        d['meibo'] = {
            'origin_x': m.origin_x,
            'origin_y': m.origin_y,
            'cell_width': m.cell_width,
            'cell_height': m.cell_height,
            'row_count': m.row_count,
            'data_start_index': m.data_start_index,
            'ref_name': m.ref_name,
            'direction': m.direction,
        }
    if obj.image is not None:
        img = obj.image
        d['image'] = {
            'rect': list(img.rect),
            'image_data': base64.b64encode(img.image_data).decode('ascii'),
            'original_path': img.original_path,
        }

    return d


def _paper_to_dict(p: PaperLayout) -> dict:
    """PaperLayout を JSON 互換の dict に変換する。"""
    return {
        'mode': p.mode,
        'unit_mm': p.unit_mm,
        'item_width_mm': p.item_width_mm,
        'item_height_mm': p.item_height_mm,
        'cols': p.cols,
        'rows': p.rows,
        'margin_left_mm': p.margin_left_mm,
        'margin_top_mm': p.margin_top_mm,
        'spacing_h_mm': p.spacing_h_mm,
        'spacing_v_mm': p.spacing_v_mm,
        'paper_size': p.paper_size,
        'orientation': p.orientation,
    }


def layfile_to_dict(lay: LayFile) -> dict:
    """LayFile を JSON 互換の dict に変換する。

    paper または table_columns があれば v2 フォーマットで出力する。
    """
    has_v2_data = lay.paper is not None or any(
        obj.table_columns or obj.meibo is not None or obj.image is not None
        for obj in lay.objects
    )
    fmt = _FORMAT_V2 if has_v2_data else _FORMAT_V1

    d: dict = {
        'format': fmt,
        'title': lay.title,
        'version': lay.version,
        'page_width': lay.page_width,
        'page_height': lay.page_height,
        'objects': [_object_to_dict(obj) for obj in lay.objects],
    }

    if lay.paper is not None:
        d['paper'] = _paper_to_dict(lay.paper)

    return d


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


def _dict_to_table_column(d: dict) -> TableColumn:
    return TableColumn(
        field_id=d.get('field_id', 0),
        width=d.get('width', 1),
        h_align=d.get('h_align', 0),
        header=d.get('header', ''),
    )


def _dict_to_paper(d: dict) -> PaperLayout:
    return PaperLayout(
        mode=d.get('mode', 0),
        unit_mm=d.get('unit_mm', 0.25),
        item_width_mm=d.get('item_width_mm', 0.0),
        item_height_mm=d.get('item_height_mm', 0.0),
        cols=d.get('cols', 1),
        rows=d.get('rows', 1),
        margin_left_mm=d.get('margin_left_mm', 0.0),
        margin_top_mm=d.get('margin_top_mm', 0.0),
        spacing_h_mm=d.get('spacing_h_mm', 0.0),
        spacing_v_mm=d.get('spacing_v_mm', 0.0),
        paper_size=d.get('paper_size', ''),
        orientation=d.get('orientation', ''),
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
    if 'table_columns' in d:
        obj.table_columns = [
            _dict_to_table_column(c) for c in d['table_columns']
        ]
    if 'meibo' in d:
        m = d['meibo']
        obj.meibo = MeiboArea(
            origin_x=m.get('origin_x', 0),
            origin_y=m.get('origin_y', 0),
            cell_width=m.get('cell_width', 0),
            cell_height=m.get('cell_height', 0),
            row_count=m.get('row_count', 0),
            data_start_index=m.get('data_start_index', 0),
            ref_name=m.get('ref_name', ''),
            direction=m.get('direction', 0),
        )
    if 'image' in d:
        img = d['image']
        obj.image = EmbeddedImage(
            rect=tuple(img.get('rect', [0, 0, 0, 0])),
            image_data=base64.b64decode(img.get('image_data', '')),
            original_path=img.get('original_path', ''),
        )

    return obj


def dict_to_layfile(data: dict) -> LayFile:
    """dict から LayFile を復元する。v1 / v2 両対応。"""
    fmt = data.get('format', '')
    if fmt not in _SUPPORTED_FORMATS:
        raise ValueError(
            f"Unknown format: {fmt!r} (expected one of {_SUPPORTED_FORMATS})"
        )

    paper = None
    if 'paper' in data:
        paper = _dict_to_paper(data['paper'])

    return LayFile(
        title=data.get('title', ''),
        version=data.get('version', 0),
        page_width=data.get('page_width', 840),
        page_height=data.get('page_height', 1188),
        objects=[_dict_to_object(o) for o in data.get('objects', [])],
        paper=paper,
    )


# ── 公開 API ─────────────────────────────────────────────────────────────────


def save_layout(lay: LayFile, path: str) -> None:
    """LayFile を JSON ファイルに保存する。

    アトミック書き込み: 一時ファイルに書いてからリネームする。
    書き込み中にエラーが起きても元ファイルは壊れない。
    """
    out_dir = Path(path).parent
    if out_dir != Path('.'):
        out_dir.mkdir(parents=True, exist_ok=True)
    data = layfile_to_dict(lay)

    # 同じディレクトリに一時ファイルを作成（rename がアトミックになるよう同一FS）
    fd, tmp_path = tempfile.mkstemp(
        dir=str(out_dir), suffix='.tmp', prefix='.layout_',
    )
    try:
        with open(fd, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        # Windows では既存ファイルへの rename が失敗するため replace を使う
        Path(tmp_path).replace(path)
    except BaseException:
        Path(tmp_path).unlink(missing_ok=True)
        raise


def load_layout(path: str) -> LayFile:
    """JSON ファイルから LayFile を読み込む。"""
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    return dict_to_layfile(data)
