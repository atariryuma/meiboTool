"""レイアウトライブラリ管理

内部保存された .json レイアウトファイルをスキャンし、
メタデータを一覧で返す。.lay ファイルのインポートにも対応。
"""

from __future__ import annotations

import json
import os
import shutil
from pathlib import Path
from typing import Any

from core.lay_parser import parse_lay
from core.lay_serializer import load_layout, save_layout


def scan_layout_dir(layout_dir: str) -> list[dict[str, Any]]:
    """レイアウトフォルダ内の .json ファイルをスキャンしてメタデータ一覧を返す。

    Returns:
        [{name, file, path, title, page_width, page_height,
          page_size_mm, object_count, field_count, label_count, line_count}, ...]
    """
    results: list[dict[str, Any]] = []
    if not os.path.isdir(layout_dir):
        return results

    for fname in sorted(os.listdir(layout_dir)):
        if not fname.lower().endswith('.json'):
            continue
        path = os.path.join(layout_dir, fname)
        meta = _read_layout_meta(path)
        if meta is not None:
            meta['file'] = fname
            meta['path'] = path
            results.append(meta)

    return results


def _read_layout_meta(path: str) -> dict[str, Any] | None:
    """JSON レイアウトファイルからメタデータを読み取る。"""
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return None

    if data.get('format') != 'meibo_layout_v1':
        return None

    objects = data.get('objects', [])
    field_count = sum(1 for o in objects if o.get('type') == 'FIELD')
    label_count = sum(1 for o in objects if o.get('type') == 'LABEL')
    line_count = sum(1 for o in objects if o.get('type') == 'LINE')

    pw = data.get('page_width', 840)
    ph = data.get('page_height', 1188)
    pw_mm = pw * 0.25
    ph_mm = ph * 0.25

    return {
        'name': Path(path).stem,
        'title': data.get('title', ''),
        'page_width': pw,
        'page_height': ph,
        'page_size_mm': f'{pw_mm:.0f}x{ph_mm:.0f}mm',
        'object_count': len(objects),
        'field_count': field_count,
        'label_count': label_count,
        'line_count': line_count,
    }


def import_lay_file(src_path: str, layout_dir: str) -> str:
    """.lay ファイルを解析し、JSON として layout_dir に保存する。

    Returns:
        保存先の JSON ファイルパス。
    """
    lay = parse_lay(src_path)
    stem = Path(src_path).stem
    dest_path = _unique_path(layout_dir, stem)
    save_layout(lay, dest_path)
    return dest_path


def import_json_file(src_path: str, layout_dir: str) -> str:
    """既存の .json レイアウトファイルをライブラリにコピーする。

    Returns:
        コピー先のファイルパス。

    Raises:
        ValueError: meibo_layout_v1 形式でない場合。
    """
    load_layout(src_path)  # バリデーション（ValueError on invalid format）
    stem = Path(src_path).stem
    dest_path = _unique_path(layout_dir, stem)
    shutil.copy2(src_path, dest_path)
    return dest_path


def delete_layout(path: str) -> None:
    """レイアウトファイルを削除する。"""
    if os.path.exists(path):
        os.remove(path)


def rename_layout(path: str, new_name: str) -> str:
    """レイアウトファイルをリネームする。JSON 内の title も更新する。

    Returns:
        新しいファイルパス。

    Raises:
        FileExistsError: 同名ファイルが既に存在する場合。
    """
    new_path = os.path.join(os.path.dirname(path), f'{new_name}.json')
    if os.path.exists(new_path) and os.path.abspath(new_path) != os.path.abspath(path):
        raise FileExistsError(f'ファイルが既に存在します: {new_name}.json')

    lay = load_layout(path)
    from core.lay_parser import LayFile
    updated = LayFile(
        title=new_name,
        version=lay.version,
        page_width=lay.page_width,
        page_height=lay.page_height,
        objects=lay.objects,
    )
    save_layout(updated, new_path)
    if os.path.abspath(new_path) != os.path.abspath(path):
        os.remove(path)
    return new_path


def _unique_path(layout_dir: str, stem: str) -> str:
    """名前衝突を回避したファイルパスを返す。"""
    dest = os.path.join(layout_dir, f'{stem}.json')
    counter = 1
    while os.path.exists(dest):
        dest = os.path.join(layout_dir, f'{stem}_{counter}.json')
        counter += 1
    return dest
