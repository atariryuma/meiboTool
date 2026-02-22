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

from core.lay_parser import (
    LayFile,
    LayoutObject,
    MeiboArea,
    ObjectType,
    PaperLayout,
    RawTag,
    Rect,
    TableColumn,
    new_field,
    new_label,
    new_line,
)
from core.lay_serializer import save_layout
from core.layout_registry import (
    build_layout_registry,
    collect_part_layout_keys,
    delete_layout,
    import_json_file,
    rename_layout,
    scan_layout_dir,
)

# ── ヘルパー ─────────────────────────────────────────────────────────────────

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


def _make_meibo_layout_json(path: str, title: str, ref_name: str) -> None:
    """MEIBO 参照を持つテスト用レイアウト JSON を作成する。"""
    lay = LayFile(
        title=title,
        objects=[
            LayoutObject(
                obj_type=ObjectType.MEIBO,
                rect=Rect(0, 0, 100, 100),
                meibo=MeiboArea(ref_name=ref_name, row_count=1),
            ),
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

    def test_rename_same_name(self, tmp_path):
        """同じ名前にリネームしてもエラーにならない。"""
        path = str(tmp_path / 'same.json')
        _make_layout_json(path, title='同名テスト')

        new_path = rename_layout(path, 'same')
        assert os.path.isfile(new_path)


# ── unique_path テスト ───────────────────────────────────────────────────────


class TestUniquePath:
    """unique_path のテスト。"""

    def test_no_collision(self, tmp_path):
        from core.layout_registry import unique_path
        result = unique_path(str(tmp_path), 'new_layout')
        assert result.endswith('new_layout.json')

    def test_collision_increments(self, tmp_path):
        from core.layout_registry import unique_path
        # 衝突ファイルを作成
        (tmp_path / 'test.json').write_text('{}')
        (tmp_path / 'test_1.json').write_text('{}')

        result = unique_path(str(tmp_path), 'test')
        assert result.endswith('test_2.json')

    def test_import_json_nonexistent_raises(self, tmp_path):
        """存在しない JSON ファイルのインポートはエラー。"""
        with pytest.raises(FileNotFoundError):
            import_json_file(
                str(tmp_path / 'no_such.json'),
                str(tmp_path / 'lib'),
            )


# ── v2 メタデータ読み取りテスト ──────────────────────────────────────────


class TestV2MetaData:
    """v2 フォーマットのスキャン・メタデータ読み取りテスト。"""

    def test_v2_layout_scanned(self, tmp_path):
        """v2 JSON がスキャンで検出される。"""
        paper = PaperLayout(mode=1, unit_mm=0.1, paper_size='A4')
        lay = LayFile(
            title='v2テスト',
            page_width=2100,
            page_height=2970,
            objects=[new_label(0, 0, 100, 30, text='test')],
            paper=paper,
        )
        save_layout(lay, str(tmp_path / 'v2.json'))
        results = scan_layout_dir(str(tmp_path))
        assert len(results) == 1
        assert results[0]['title'] == 'v2テスト'

    def test_v2_page_size_mm_uses_unit_mm(self, tmp_path):
        """v2 の page_size_mm が unit_mm=0.1 を使って計算される。"""
        paper = PaperLayout(unit_mm=0.1)
        lay = LayFile(
            page_width=2100,
            page_height=2970,
            paper=paper,
        )
        save_layout(lay, str(tmp_path / 'v2.json'))
        results = scan_layout_dir(str(tmp_path))
        # 2100 * 0.1 = 210mm, 2970 * 0.1 = 297mm
        assert results[0]['page_size_mm'] == '210x297mm'

    def test_rename_preserves_paper(self, tmp_path):
        """リネーム後も paper が保持される。"""
        paper = PaperLayout(mode=1, paper_size='A4', orientation='portrait')
        lay = LayFile(
            title='元の名前',
            objects=[new_label(0, 0, 100, 30, text='test')],
            paper=paper,
            raw_tags=[RawTag(path=[0x05E0], payload=b'\x02\x00\x00\x00', payload_len=4)],
        )
        path = str(tmp_path / 'original.json')
        save_layout(lay, path)
        new_path = rename_layout(path, '新しい名前')

        from core.lay_serializer import load_layout
        restored = load_layout(new_path)
        assert restored.paper is not None
        assert restored.paper.paper_size == 'A4'
        assert restored.raw_tags == [
            RawTag(path=[0x05E0], payload=b'\x02\x00\x00\x00', payload_len=4),
        ]

    def test_table_in_meta_count(self, tmp_path):
        """TABLE オブジェクトが object_count に含まれる。"""
        lay = LayFile(
            objects=[
                new_label(0, 0, 100, 30, text='test'),
                LayoutObject(
                    obj_type=ObjectType.TABLE,
                    rect=Rect(0, 30, 800, 600),
                    table_columns=[
                        TableColumn(field_id=108, header='氏名'),
                    ],
                ),
            ],
            paper=PaperLayout(unit_mm=0.1),
        )
        save_layout(lay, str(tmp_path / 'table.json'))
        results = scan_layout_dir(str(tmp_path))
        assert results[0]['object_count'] == 2


# ── build_layout_registry テスト ──────────────────────────────────────────


class TestBuildLayoutRegistry:
    """build_layout_registry() のテスト。"""

    def test_empty_dir(self, tmp_path):
        registry = build_layout_registry(str(tmp_path))
        assert registry == {}

    def test_nonexistent_dir(self, tmp_path):
        registry = build_layout_registry(str(tmp_path / 'no_such'))
        assert registry == {}

    def test_registers_by_title(self, tmp_path):
        _make_layout_json(str(tmp_path / 'file1.json'), title='帳票タイトルA')
        registry = build_layout_registry(str(tmp_path))
        assert '帳票タイトルA' in registry

    def test_registers_by_stem(self, tmp_path):
        _make_layout_json(str(tmp_path / 'my_layout.json'), title='異なるタイトル')
        registry = build_layout_registry(str(tmp_path))
        assert 'my_layout' in registry
        assert '異なるタイトル' in registry

    def test_stem_equals_title_no_duplicate(self, tmp_path):
        """stem == title のとき重複エントリにならない。"""
        _make_layout_json(str(tmp_path / 'same.json'), title='same')
        registry = build_layout_registry(str(tmp_path))
        assert 'same' in registry
        # 値は同一オブジェクト
        assert registry['same'].title == 'same'

    def test_alias_gakkyu(self, tmp_path):
        """'gakkyu' エイリアスが 'takara_simei' に解決される。"""
        _make_layout_json(str(tmp_path / 'takara_simei.json'), title='takara_simei')
        registry = build_layout_registry(str(tmp_path))
        assert 'gakkyu' in registry
        assert registry['gakkyu'] is registry['takara_simei']

    def test_alias_sirabe(self, tmp_path):
        """'sirabe' エイリアスが 'takara_sirabe' に解決される。"""
        _make_layout_json(str(tmp_path / 'takara_sirabe.json'), title='takara_sirabe')
        registry = build_layout_registry(str(tmp_path))
        assert 'sirabe' in registry

    def test_alias_photo_copy(self, tmp_path):
        """写真名簿用 ref_name エイリアスが写真ラベルに解決される。"""
        _make_layout_json(
            str(tmp_path / 'photo_label.json'),
            title='【天久小】写真ラベル（名前なし）',
        )
        registry = build_layout_registry(str(tmp_path))
        assert 'コピー：写真＋番号＋ふりがな' in registry
        assert (
            registry['コピー：写真＋番号＋ふりがな']
            is registry['【天久小】写真ラベル（名前なし）']
        )

    def test_alias_unresolved_without_target(self, tmp_path):
        """ターゲットが無いエイリアスは登録されない。"""
        _make_layout_json(str(tmp_path / 'unrelated.json'), title='無関係')
        registry = build_layout_registry(str(tmp_path))
        assert 'gakkyu' not in registry

    def test_multiple_layouts(self, tmp_path):
        _make_layout_json(str(tmp_path / 'a.json'), title='レイアウトA')
        _make_layout_json(str(tmp_path / 'b.json'), title='レイアウトB')
        _make_layout_json(str(tmp_path / 'c.json'), title='レイアウトC')
        registry = build_layout_registry(str(tmp_path))
        assert 'レイアウトA' in registry
        assert 'レイアウトB' in registry
        assert 'レイアウトC' in registry
        assert 'a' in registry
        assert 'b' in registry
        assert 'c' in registry


class TestCollectPartLayoutKeys:
    """collect_part_layout_keys() のテスト。"""

    def test_collects_referenced_layout_title(self, tmp_path):
        _make_layout_json(str(tmp_path / 'part.json'), title='パーツ')
        _make_meibo_layout_json(str(tmp_path / 'parent.json'), title='親', ref_name='パーツ')

        keys = collect_part_layout_keys(str(tmp_path))
        assert 'パーツ' in keys

    def test_collects_all_keys_of_referenced_layout(self, tmp_path):
        _make_layout_json(str(tmp_path / 'part_layout.json'), title='takara_simei')
        _make_meibo_layout_json(str(tmp_path / 'parent.json'), title='親', ref_name='gakkyu')

        keys = collect_part_layout_keys(str(tmp_path))
        assert 'takara_simei' in keys
        assert 'part_layout' in keys

    def test_ignores_self_reference(self, tmp_path):
        _make_meibo_layout_json(str(tmp_path / 'self.json'), title='self', ref_name='self')

        keys = collect_part_layout_keys(str(tmp_path))
        assert keys == set()

    def test_ignores_unresolved_reference(self, tmp_path):
        _make_meibo_layout_json(str(tmp_path / 'parent.json'), title='親', ref_name='missing')

        keys = collect_part_layout_keys(str(tmp_path))
        assert keys == set()
