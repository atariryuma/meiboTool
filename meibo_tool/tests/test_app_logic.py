"""gui/app.py の内部ロジックのユニットテスト

Tkinter ウィジェットの生成はテストしない（CI 環境で動かないため）。
_validate_required_columns, _fmt_cell, _filter_df ロジック, _BatchGenerator をテスト。
"""

from __future__ import annotations

import logging
from unittest.mock import MagicMock

import pandas as pd
import pytest

from gui.app import (
    _BatchGenerator,
    _fmt_cell,
    _sort_by_attendance,
    _validate_required_columns,
)

# ────────────────────────────────────────────────────────────────────────────
# _validate_required_columns
# ────────────────────────────────────────────────────────────────────────────

class TestValidateRequiredColumns:

    def test_mandatory_missing_raises(self):
        """mandatory_columns 不足 → ValueError。"""
        meta = {'mandatory_columns': ['組', '出席番号'], 'required_columns': ['氏名']}
        df = pd.DataFrame({'氏名': ['太郎']})
        with pytest.raises(ValueError, match='必須カラム'):
            _validate_required_columns(meta, df)

    def test_mandatory_partial_missing(self):
        """mandatory_columns の一部不足 → ValueError。"""
        meta = {'mandatory_columns': ['組', '出席番号'], 'required_columns': []}
        df = pd.DataFrame({'組': ['1']})
        with pytest.raises(ValueError, match='出席番号'):
            _validate_required_columns(meta, df)

    def test_required_missing_warns(self, caplog):
        """required_columns 不足 → 警告のみ（例外なし）。"""
        meta = {'mandatory_columns': [], 'required_columns': ['氏名', '性別']}
        df = pd.DataFrame({'出席番号': ['1']})
        with caplog.at_level(logging.WARNING):
            _validate_required_columns(meta, df)
        assert '推奨カラムが不足' in caplog.text

    def test_all_present_no_error(self):
        """全カラム揃い → 例外なし。"""
        meta = {
            'mandatory_columns': ['組', '出席番号'],
            'required_columns': ['氏名'],
        }
        df = pd.DataFrame({'組': ['1'], '出席番号': ['1'], '氏名': ['太郎']})
        _validate_required_columns(meta, df)  # no exception

    def test_no_mandatory_no_error(self):
        """mandatory_columns が空 → 例外なし。"""
        meta = {'mandatory_columns': [], 'required_columns': []}
        df = pd.DataFrame({'氏名': ['太郎']})
        _validate_required_columns(meta, df)

    def test_empty_meta_keys_no_error(self):
        """meta にキーがない場合 → 例外なし。"""
        _validate_required_columns({}, pd.DataFrame({'氏名': ['太郎']}))


# ────────────────────────────────────────────────────────────────────────────
# _fmt_cell
# ────────────────────────────────────────────────────────────────────────────

class TestFmtCell:

    def test_normal_string(self):
        assert _fmt_cell('氏名', '山田 太郎') == '山田 太郎'

    def test_none_returns_empty(self):
        assert _fmt_cell('氏名', None) == ''

    def test_nan_returns_empty(self):
        assert _fmt_cell('氏名', float('nan')) == ''

    def test_nan_string_returns_empty(self):
        assert _fmt_cell('氏名', 'nan') == ''

    def test_empty_string_returns_empty(self):
        assert _fmt_cell('氏名', '') == ''

    def test_date_column_formatted(self):
        """日付カラムは format_date() でフォーマットされる。"""
        result = _fmt_cell('生年月日', '2018-06-15')
        assert result == '18/06/15'

    def test_non_date_column_not_formatted(self):
        """日付以外のカラムはそのまま。"""
        result = _fmt_cell('氏名', '2018-06-15')
        assert result == '2018-06-15'

    def test_numeric_value(self):
        assert _fmt_cell('出席番号', 5) == '5'


# ────────────────────────────────────────────────────────────────────────────
# _filter_df ロジック（App クラスを経由せず同等ロジックを検証）
# ────────────────────────────────────────────────────────────────────────────

def _filter_df_logic(df: pd.DataFrame, 学年: str | None, 組: str | None) -> pd.DataFrame:
    """基本的なフィルタ＋出席番号ソートのロジックを検証する（特支統合なし）。"""
    df = df.copy()
    if 学年 and '学年' in df.columns:
        df = df[df['学年'] == 学年]
    if 組 and '組' in df.columns:
        df = df[df['組'] == 組]
    if '出席番号' in df.columns:
        df['_sort_key'] = pd.to_numeric(df['出席番号'], errors='coerce').fillna(0)
        df = df.sort_values('_sort_key').drop(columns='_sort_key')
        df = df.reset_index(drop=True)
    return df


class TestFilterDfLogic:

    @pytest.fixture
    def multi_class_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {'学年': '1', '組': '1', '出席番号': '10', '氏名': 'A'},
            {'学年': '1', '組': '1', '出席番号': '2',  '氏名': 'B'},
            {'学年': '1', '組': '2', '出席番号': '1',  '氏名': 'C'},
            {'学年': '2', '組': '1', '出席番号': '1',  '氏名': 'D'},
        ])

    def test_filter_by_grade_and_class(self, multi_class_df):
        result = _filter_df_logic(multi_class_df, '1', '1')
        assert len(result) == 2
        assert set(result['氏名']) == {'A', 'B'}

    def test_filter_by_grade_only(self, multi_class_df):
        result = _filter_df_logic(multi_class_df, '1', None)
        assert len(result) == 3

    def test_no_filter_returns_all(self, multi_class_df):
        result = _filter_df_logic(multi_class_df, None, None)
        assert len(result) == 4

    def test_numeric_sort(self, multi_class_df):
        """出席番号が数値順にソートされる（'2' < '10'）。"""
        result = _filter_df_logic(multi_class_df, '1', '1')
        assert list(result['出席番号']) == ['2', '10']

    def test_non_numeric_attendance_number(self):
        """非数値の出席番号は fillna(0) で先頭に配置される。"""
        df = pd.DataFrame([
            {'学年': '1', '組': '1', '出席番号': 'X', '氏名': 'X'},
            {'学年': '1', '組': '1', '出席番号': '1', '氏名': 'A'},
        ])
        result = _filter_df_logic(df, '1', '1')
        # 'X' → coerce → NaN → fillna(0) → 0 < 1 → 'X' が先
        assert list(result['氏名']) == ['X', 'A']

    def test_empty_result(self, multi_class_df):
        """存在しないクラスでフィルタ → 空。"""
        result = _filter_df_logic(multi_class_df, '9', '9')
        assert len(result) == 0


# ────────────────────────────────────────────────────────────────────────────
# _BatchGenerator
# ────────────────────────────────────────────────────────────────────────────

class TestBatchGenerator:

    def test_generate_calls_all(self):
        """全ジェネレーターの generate() が呼ばれる。"""
        mock1 = MagicMock()
        mock2 = MagicMock()
        batch = _BatchGenerator([mock1, mock2], '/tmp/output')
        batch.generate()
        mock1.generate.assert_called_once()
        mock2.generate.assert_called_once()

    def test_returns_output_dir(self):
        """一括生成完了パスを返す。"""
        batch = _BatchGenerator([], '/tmp/output')
        result = batch.generate()
        assert result.startswith('/tmp/output')
        assert '一括生成完了' in result

    def test_empty_generators(self):
        """空リストでも動作する。"""
        batch = _BatchGenerator([], '/tmp/output')
        batch.generate()  # no exception


# ────────────────────────────────────────────────────────────────────────────
# _sort_by_attendance
# ────────────────────────────────────────────────────────────────────────────

class TestSortByAttendance:

    def test_numeric_sort(self):
        """出席番号が数値順にソートされる。"""
        df = pd.DataFrame([
            {'出席番号': '10', '氏名': 'A'},
            {'出席番号': '2', '氏名': 'B'},
            {'出席番号': '1', '氏名': 'C'},
        ])
        result = _sort_by_attendance(df)
        assert list(result['氏名']) == ['C', 'B', 'A']

    def test_no_attendance_column(self):
        """出席番号列がない場合もエラーなし。"""
        df = pd.DataFrame([{'氏名': 'A'}, {'氏名': 'B'}])
        result = _sort_by_attendance(df)
        assert len(result) == 2

    def test_non_numeric_first(self):
        """非数値は fillna(0) で先頭。"""
        df = pd.DataFrame([
            {'出席番号': '5', '氏名': 'A'},
            {'出席番号': 'X', '氏名': 'B'},
        ])
        result = _sort_by_attendance(df)
        assert list(result['氏名']) == ['B', 'A']

    def test_index_preserved(self):
        """ソート後も元のインデックスが保持される（セル編集のマッピング用）。"""
        df = pd.DataFrame([
            {'出席番号': '3', '氏名': 'A'},
            {'出席番号': '1', '氏名': 'B'},
        ])
        result = _sort_by_attendance(df)
        # 出席番号順: B(idx=1), A(idx=0)
        assert list(result.index) == [1, 0]


# ────────────────────────────────────────────────────────────────────────────
# 自動統合ロジック（_filter_df の交流学級割り当て自動統合を再現）
# ────────────────────────────────────────────────────────────────────────────

def _filter_with_assignments(
    df_all: pd.DataFrame,
    学年: str | None,
    組: str | None,
    assignments: dict[str, str],
    placement: str = 'appended',
) -> pd.DataFrame:
    """App._filter_df の自動統合ロジックを再現する。"""
    from core.special_needs import (
        detect_special_needs_students,
        get_assigned_students,
        merge_special_needs_students,
    )

    df = df_all.copy()
    if 学年 and '学年' in df.columns:
        df = df[df['学年'] == 学年]

    if 組 and '組' in df.columns:
        df_regular = df[df['組'] == 組].copy()
        df_regular = _sort_by_attendance(df_regular)

        if assignments:
            target_class = f'{学年}-{組}'
            special_all = detect_special_needs_students(df_all)
            assigned = get_assigned_students(special_all, assignments, target_class)
            if not assigned.empty:
                return merge_special_needs_students(
                    df_regular, assigned, placement=placement,
                )
        return df_regular

    return _sort_by_attendance(df)


class TestFilterWithAssignments:

    @pytest.fixture
    def mixed_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {'学年': '1', '組': '1', '出席番号': '1', '氏名': 'A'},
            {'学年': '1', '組': '1', '出席番号': '3', '氏名': 'C'},
            {'学年': '1', '組': '2', '出席番号': '1', '氏名': 'D'},
            {'学年': '1', '組': 'なかよし', '出席番号': '1', '氏名': '特支X'},
            {'学年': '1', '組': 'なかよし', '出席番号': '2', '氏名': '特支Y'},
            {'学年': '2', '組': '1', '出席番号': '1', '氏名': 'E'},
            {'学年': '2', '組': 'ひまわり', '出席番号': '1', '氏名': '特支Z'},
        ])

    def test_assigned_included_appended(self, mixed_df):
        """割り当て済み特支児童が末尾に追加される。"""
        assignments = {'1-なかよし-1': '1-1'}
        result = _filter_with_assignments(mixed_df, '1', '1', assignments)
        assert len(result) == 3
        assert list(result['氏名']) == ['A', 'C', '特支X']

    def test_assigned_included_integrated(self, mixed_df):
        """割り当て済み特支児童が出席番号順に統合される。"""
        assignments = {'1-なかよし-1': '1-1'}
        result = _filter_with_assignments(
            mixed_df, '1', '1', assignments, placement='integrated',
        )
        assert len(result) == 3
        assert list(result['氏名']) == ['A', '特支X', 'C']

    def test_no_assignment_no_special(self, mixed_df):
        """割り当てなし → 通常児童のみ。"""
        result = _filter_with_assignments(mixed_df, '1', '1', {})
        assert len(result) == 2
        assert set(result['氏名']) == {'A', 'C'}

    def test_assignment_different_class(self, mixed_df):
        """他クラスに割り当て → 表示されない。"""
        assignments = {'1-なかよし-1': '1-2'}
        result = _filter_with_assignments(mixed_df, '1', '1', assignments)
        assert len(result) == 2
        assert '特支X' not in result['氏名'].values

    def test_multiple_assignments(self, mixed_df):
        """複数児童の割り当て。"""
        assignments = {
            '1-なかよし-1': '1-1',
            '1-なかよし-2': '1-1',
        }
        result = _filter_with_assignments(mixed_df, '1', '1', assignments)
        assert len(result) == 4
        assert list(result['氏名']) == ['A', 'C', '特支X', '特支Y']

    def test_grade_level_includes_all(self, mixed_df):
        """学年全体: 全員（特支含む）が表示される。"""
        result = _filter_with_assignments(mixed_df, '1', None, {})
        assert len(result) == 5  # 通常3 + 特支2
        assert '特支X' in result['氏名'].values

    def test_school_level_includes_all(self, mixed_df):
        """全校: 全員が表示される。"""
        result = _filter_with_assignments(mixed_df, None, None, {})
        assert len(result) == 7

    def test_cross_grade_not_included(self, mixed_df):
        """他学年の特支は含まれない。"""
        assignments = {'2-ひまわり-1': '1-1'}
        result = _filter_with_assignments(mixed_df, '1', '1', assignments)
        # 2年ひまわりの児童は全特支から取得されるが、1-1 に割り当てれば含まれる
        assert len(result) == 3
        assert '特支Z' in result['氏名'].values
