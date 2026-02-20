"""layout_registry.py のユニットテスト

テスト対象:
  - レイアウトフォルダのスキャン
  - .lay / .json インポート
  - 削除・リネーム
"""

from __future__ import annotations

import json
import os

import pytest

from core.lay_parser import LayFile, new_field, new_label, new_line
from core.lay_serializer import save_layout
from core.layout_registry import (
    delete_layout,
    import_json_file,
    import_lay_file,
    rename_layout,
    scan_layout_dir,
)

# ── ヘルパー ─────────────────────────────────────────────────────────────────

_SAMPLE_LAY = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
    'R8年度小学校個票20260130.lay',
)


def _make_layout_json(path: str, title: str = 'テスト', n_fields: int = 2) -> None:
    """テスト用の JSON レイアウトファイルを作成する。"""
    lay = LayFile(
        title=title,
        objects=[
            new_label(0, 0, 100, 30, text='見出し'),
            *[new_field(0, 30 * (i + 1), 100, 30 * (i + 2), field_id=100 + i)
              for i in range(n_fields)],
            new_line(0, 0, 100, 0),
        ],
    )
    save_layout(lay, path)


# ── scan_layout_dir テスト ────────────────────────────────────────────────────


class TestScanLayoutDir:
    """レイアウトフォルダスキャンのテスト。"""

    def test_empty_dir(self, tmp_path):
        results = scan_layout_dir(str(tmp_path))
        assert results == []

    def test_nonexistent_dir(self, tmp_path):
        results = scan_layout_dir(str(tmp_path / 'no_such_dir'))
        assert results == []

    def test_scan_with_layouts(self, tmp_path):
        _make_layout_json(str(tmp_path / 'layout1.json'), title='帳票A')
        _make_layout_json(str(tmp_path / 'layout2.json'), title='帳票B', n_fields=5)

        results = scan_layout_dir(str(tmp_path))
        assert len(results) == 2
        names = [r['name'] for r in results]
        assert 'layout1' in names
        assert 'layout2' in names

    def test_scan_meta_fields(self, tmp_path):
        _make_layout_json(str(tmp_path / 'test.json'), title='テスト帳票', n_fields=3)

        results = scan_layout_dir(str(tmp_path))
        assert len(results) == 1
        meta = results[0]
        assert meta['title'] == 'テスト帳票'
        assert meta['field_count'] == 3
        assert meta['label_count'] == 1
        assert meta['line_count'] == 1
        assert meta['object_count'] == 5  # 1 label + 3 fields + 1 line
        assert meta['page_size_mm'] == '210x297mm'

    def test_scan_ignores_non_json(self, tmp_path):
        _make_layout_json(str(tmp_path / 'layout.json'))
        (tmp_path / 'readme.txt').write_text('not a layout')
        (tmp_path / 'data.xlsx').write_bytes(b'\x00')

        results = scan_layout_dir(str(tmp_path))
        assert len(results) == 1

    def test_scan_ignores_invalid_json(self, tmp_path):
        _make_layout_json(str(tmp_path / 'valid.json'))
        # 不正な JSON
        (tmp_path / 'broken.json').write_text('{invalid json}')
        # 正しい JSON だが format が違う
        (tmp_path / 'other.json').write_text(
            json.dumps({'format': 'other_format'}),
        )

        results = scan_layout_dir(str(tmp_path))
        assert len(results) == 1
        assert results[0]['name'] == 'valid'


# ── import テスト ─────────────────────────────────────────────────────────────


class TestImport:
    """インポートのテスト。"""

    @pytest.mark.skipif(
        not os.path.isfile(_SAMPLE_LAY),
        reason='サンプル .lay ファイルが見つかりません',
    )
    def test_import_lay_file(self, tmp_path):
        dest = import_lay_file(_SAMPLE_LAY, str(tmp_path))
        assert os.path.isfile(dest)
        assert dest.endswith('.json')

        results = scan_layout_dir(str(tmp_path))
        assert len(results) == 1
        assert results[0]['object_count'] > 0

    def test_import_json_file(self, tmp_path):
        src = str(tmp_path / 'source' / 'layout.json')
        os.makedirs(os.path.dirname(src), exist_ok=True)
        _make_layout_json(src, title='インポート用')

        lib_dir = str(tmp_path / 'library')
        os.makedirs(lib_dir)
        dest = import_json_file(src, lib_dir)

        assert os.path.isfile(dest)
        results = scan_layout_dir(lib_dir)
        assert len(results) == 1
        assert results[0]['title'] == 'インポート用'

    def test_import_name_collision(self, tmp_path):
        src = str(tmp_path / 'source' / 'layout.json')
        os.makedirs(os.path.dirname(src), exist_ok=True)
        _make_layout_json(src)

        lib_dir = str(tmp_path / 'library')
        os.makedirs(lib_dir)
        # 同名ファイルを先に作成
        _make_layout_json(os.path.join(lib_dir, 'layout.json'))

        dest = import_json_file(src, lib_dir)
        assert 'layout_1.json' in dest

    def test_import_json_validates_format(self, tmp_path):
        src = str(tmp_path / 'bad.json')
        with open(src, 'w') as f:
            json.dump({'format': 'wrong'}, f)

        with pytest.raises(ValueError, match='Unknown format'):
            import_json_file(src, str(tmp_path / 'lib'))


# ── delete / rename テスト ────────────────────────────────────────────────────


class TestDeleteRename:
    """削除・リネームのテスト。"""

    def test_delete_layout(self, tmp_path):
        path = str(tmp_path / 'to_delete.json')
        _make_layout_json(path)
        assert os.path.isfile(path)

        delete_layout(path)
        assert not os.path.isfile(path)

    def test_delete_nonexistent(self, tmp_path):
        # エラーにならないことを確認
        delete_layout(str(tmp_path / 'no_such_file.json'))

    def test_rename_layout(self, tmp_path):
        path = str(tmp_path / 'old_name.json')
        _make_layout_json(path, title='旧名称')

        new_path = rename_layout(path, '新名称')
        assert os.path.isfile(new_path)
        assert not os.path.isfile(path)
        assert os.path.basename(new_path) == '新名称.json'

    def test_rename_updates_title(self, tmp_path):
        path = str(tmp_path / 'original.json')
        _make_layout_json(path, title='元の名前')

        rename_layout(path, '変更後')
        results = scan_layout_dir(str(tmp_path))
        assert len(results) == 1
        assert results[0]['title'] == '変更後'

    def test_rename_conflict_raises(self, tmp_path):
        path1 = str(tmp_path / 'a.json')
        path2 = str(tmp_path / 'b.json')
        _make_layout_json(path1)
        _make_layout_json(path2)

        with pytest.raises(FileExistsError):
            rename_layout(path1, 'b')
