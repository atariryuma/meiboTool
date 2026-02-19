"""gui/app.py の内部ロジックのユニットテスト

Tkinter ウィジェットの生成はテストしない（CI 環境で動かないため）。
_validate_required_columns, _fmt_cell, _filter_df ロジック, _BatchGenerator をテスト。
"""

from __future__ import annotations

import logging
from unittest.mock import MagicMock

import pandas as pd
import pytest

from gui.app import _BatchGenerator, _fmt_cell, _validate_required_columns

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
    """App._filter_df と同じロジックを再現。"""
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
