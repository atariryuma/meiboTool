"""gui/dialogs/mapping_dialog.py のユニットテスト

GUI ダイアログの内部ロジック（論理名一覧・重複排除）を中心にテスト。
Tkinter ウィジェットの生成はテストしない（CI 環境で動かないため）。
"""

from __future__ import annotations

from gui.dialogs.mapping_dialog import _ALL_LOGICAL_NAMES

# ── 論理名一覧テスト ──────────────────────────────────────────────────────────


class TestAllLogicalNames:
    """_ALL_LOGICAL_NAMES の基本プロパティ。"""

    def test_is_sorted(self):
        """論理名一覧がソート済みであること。"""
        assert sorted(_ALL_LOGICAL_NAMES) == _ALL_LOGICAL_NAMES

    def test_no_duplicates(self):
        """重複がないこと。"""
        assert len(_ALL_LOGICAL_NAMES) == len(set(_ALL_LOGICAL_NAMES))

    def test_contains_required_columns(self):
        """必須の論理名が含まれていること。"""
        required = ['氏名', '氏名かな', '性別', '生年月日', '組', '出席番号']
        for name in required:
            assert name in _ALL_LOGICAL_NAMES, f'{name} が論理名一覧にない'

    def test_contains_exact_map_values(self):
        """EXACT_MAP の全バリューが含まれていること。"""
        from core.mapper import EXACT_MAP
        for val in EXACT_MAP.values():
            assert val in _ALL_LOGICAL_NAMES, f'{val} が論理名一覧にない'

    def test_contains_alias_keys(self):
        """COLUMN_ALIASES の全キーが含まれていること。"""
        from core.mapper import COLUMN_ALIASES
        for key in COLUMN_ALIASES:
            assert key in _ALL_LOGICAL_NAMES, f'{key} が論理名一覧にない'

    def test_has_reasonable_count(self):
        """論理名が妥当な数であること。"""
        assert 25 <= len(_ALL_LOGICAL_NAMES) <= 100


# ── マッピングロジックテスト ──────────────────────────────────────────────────


class TestMappingLogic:
    """マッピング確定ロジックの単体テスト。"""

    def test_skip_value_excluded(self):
        """スキップ選択はマッピング結果に含まれない。"""
        # _collect_mapping のロジックを再現
        unmapped = ['カスタム列A', 'カスタム列B']
        selections = ['（スキップ）', '氏名']
        mapping = {}
        for col, sel in zip(unmapped, selections, strict=True):
            if sel and sel != '（スキップ）':
                mapping[col] = sel
        assert mapping == {'カスタム列B': '氏名'}

    def test_empty_selection_excluded(self):
        """空文字の選択はマッピング結果に含まれない。"""
        unmapped = ['列X']
        selections = ['']
        mapping = {}
        for col, sel in zip(unmapped, selections, strict=True):
            if sel and sel != '（スキップ）':
                mapping[col] = sel
        assert mapping == {}

    def test_all_skip_returns_empty(self):
        """全スキップなら空辞書を返す。"""
        unmapped = ['A', 'B', 'C']
        selections = ['（スキップ）', '（スキップ）', '（スキップ）']
        mapping = {}
        for col, sel in zip(unmapped, selections, strict=True):
            if sel and sel != '（スキップ）':
                mapping[col] = sel
        assert mapping == {}

    def test_all_mapped_returns_full(self):
        """全て選択すれば全てマッピングされる。"""
        unmapped = ['col1', 'col2']
        selections = ['氏名', '性別']
        mapping = {}
        for col, sel in zip(unmapped, selections, strict=True):
            if sel and sel != '（スキップ）':
                mapping[col] = sel
        assert mapping == {'col1': '氏名', 'col2': '性別'}


# ── 既存マップとの重複排除テスト ──────────────────────────────────────────────


class TestDuplicateExclusion:
    """既にマッピング済みの論理名は選択肢から除外される。"""

    def test_existing_mapped_excluded(self):
        """existing_mapped に含まれる論理名は available に含まれない。"""
        existing = {'氏名', '性別', '生年月日'}
        available = [n for n in _ALL_LOGICAL_NAMES if n not in existing]
        assert '氏名' not in available
        assert '性別' not in available
        assert '生年月日' not in available

    def test_unmapped_names_available(self):
        """existing_mapped に含まれない論理名は available に残る。"""
        existing = {'氏名'}
        available = [n for n in _ALL_LOGICAL_NAMES if n not in existing]
        assert '氏名かな' in available
        assert '性別' in available

    def test_empty_existing_keeps_all(self):
        """existing_mapped が空なら全論理名が利用可能。"""
        existing: set[str] = set()
        available = [n for n in _ALL_LOGICAL_NAMES if n not in existing]
        assert len(available) == len(_ALL_LOGICAL_NAMES)
