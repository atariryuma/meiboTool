"""core/special_needs.py のユニットテスト

特別支援学級判定・検出・統合ロジックをテストする。
"""

from __future__ import annotations

import pandas as pd

from core.special_needs import (
    detect_regular_students,
    detect_special_needs_students,
    get_assigned_students,
    get_special_needs_classes,
    get_unassigned_students,
    is_special_needs_class,
    make_student_key,
    merge_special_needs_students,
)

# ────────────────────────────────────────────────────────────────────────────
# is_special_needs_class
# ────────────────────────────────────────────────────────────────────────────

class TestIsSpecialNeedsClass:

    def test_numeric_halfwidth(self):
        """半角数字は通常学級。"""
        assert is_special_needs_class('1') is False
        assert is_special_needs_class('10') is False
        assert is_special_needs_class('3') is False

    def test_numeric_fullwidth(self):
        """全角数字は通常学級。"""
        assert is_special_needs_class('１') is False
        assert is_special_needs_class('１０') is False

    def test_hiragana_is_special(self):
        """ひらがな名は特別支援学級。"""
        assert is_special_needs_class('なかよし') is True
        assert is_special_needs_class('ひまわり') is True
        assert is_special_needs_class('たんぽぽ') is True

    def test_katakana_is_special(self):
        """カタカナ名も特別支援学級。"""
        assert is_special_needs_class('ヒマワリ') is True

    def test_kanji_is_special(self):
        """漢字名は特別支援学級。"""
        assert is_special_needs_class('若竹') is True

    def test_mixed_alphanumeric_is_special(self):
        """英字混在は特別支援学級。"""
        assert is_special_needs_class('A組') is True

    def test_empty_is_not_special(self):
        assert is_special_needs_class('') is False

    def test_none_is_not_special(self):
        assert is_special_needs_class(None) is False

    def test_whitespace_number(self):
        """前後空白付き数字は通常学級。"""
        assert is_special_needs_class(' 1 ') is False


# ────────────────────────────────────────────────────────────────────────────
# detect_special_needs_students / detect_regular_students
# ────────────────────────────────────────────────────────────────────────────

class TestDetectStudents:

    def _make_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {'学年': '1', '組': '1', '出席番号': '1', '氏名': '通常太郎'},
            {'学年': '1', '組': '1', '出席番号': '2', '氏名': '通常花子'},
            {'学年': '1', '組': 'なかよし', '出席番号': '1', '氏名': '特支一郎'},
            {'学年': '2', '組': 'ひまわり', '出席番号': '2', '氏名': '特支二郎'},
            {'学年': '2', '組': '1', '出席番号': '3', '氏名': '通常次郎'},
        ])

    def test_detect_special(self):
        df = self._make_df()
        result = detect_special_needs_students(df)
        assert len(result) == 2
        assert set(result['氏名']) == {'特支一郎', '特支二郎'}

    def test_detect_regular(self):
        df = self._make_df()
        result = detect_regular_students(df)
        assert len(result) == 3
        assert '特支一郎' not in result['氏名'].values

    def test_no_special_students(self):
        df = pd.DataFrame([
            {'学年': '1', '組': '1', '出席番号': '1', '氏名': 'A'},
        ])
        assert detect_special_needs_students(df).empty

    def test_all_special_students(self):
        df = pd.DataFrame([
            {'学年': '1', '組': 'なかよし', '出席番号': '1', '氏名': 'A'},
        ])
        assert len(detect_special_needs_students(df)) == 1
        assert detect_regular_students(df).empty

    def test_no_kumi_column(self):
        df = pd.DataFrame([{'学年': '1', '氏名': 'A'}])
        assert detect_special_needs_students(df).empty
        assert len(detect_regular_students(df)) == 1


# ────────────────────────────────────────────────────────────────────────────
# get_special_needs_classes
# ────────────────────────────────────────────────────────────────────────────

class TestGetSpecialNeedsClasses:

    def test_mixed(self):
        df = pd.DataFrame([
            {'組': '1'}, {'組': 'なかよし'}, {'組': '2'}, {'組': 'ひまわり'},
        ])
        result = get_special_needs_classes(df)
        assert result == ['なかよし', 'ひまわり']

    def test_no_special(self):
        df = pd.DataFrame([{'組': '1'}, {'組': '2'}])
        assert get_special_needs_classes(df) == []

    def test_no_kumi_column(self):
        df = pd.DataFrame([{'学年': '1'}])
        assert get_special_needs_classes(df) == []


# ────────────────────────────────────────────────────────────────────────────
# merge_special_needs_students
# ────────────────────────────────────────────────────────────────────────────

class TestMergeSpecialNeedsStudents:

    def _regular_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {'出席番号': '1', '氏名': 'A', '性別': '男'},
            {'出席番号': '3', '氏名': 'C', '性別': '女'},
            {'出席番号': '5', '氏名': 'E', '性別': '男'},
        ])

    def _special_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {'出席番号': '2', '氏名': 'B特', '性別': '男'},
            {'出席番号': '4', '氏名': 'D特', '性別': '女'},
        ])

    def test_appended(self):
        """末尾に追加: 通常学級の後に特支児童が並ぶ。"""
        result = merge_special_needs_students(
            self._regular_df(), self._special_df(), placement='appended',
        )
        assert len(result) == 5
        assert list(result['氏名']) == ['A', 'C', 'E', 'B特', 'D特']

    def test_integrated(self):
        """出席番号順に統合: 出席番号順でインターリーブ。"""
        result = merge_special_needs_students(
            self._regular_df(), self._special_df(), placement='integrated',
        )
        assert len(result) == 5
        assert list(result['氏名']) == ['A', 'B特', 'C', 'D特', 'E']

    def test_empty_special(self):
        """特支が空 → 通常のみ返す。"""
        result = merge_special_needs_students(
            self._regular_df(), pd.DataFrame(columns=['出席番号', '氏名', '性別']),
        )
        assert len(result) == 3

    def test_empty_regular(self):
        """通常が空 → 特支のみ返す。"""
        result = merge_special_needs_students(
            pd.DataFrame(columns=['出席番号', '氏名', '性別']),
            self._special_df(),
            placement='appended',
        )
        assert len(result) == 2

    def test_default_is_appended(self):
        """デフォルトは appended。"""
        result = merge_special_needs_students(
            self._regular_df(), self._special_df(),
        )
        assert list(result['氏名']) == ['A', 'C', 'E', 'B特', 'D特']


# ────────────────────────────────────────────────────────────────────────────
# make_student_key
# ────────────────────────────────────────────────────────────────────────────

class TestMakeStudentKey:

    def test_basic(self):
        row = pd.Series({'学年': '1', '組': 'なかよし', '出席番号': '1'})
        assert make_student_key(row) == '1-なかよし-1'

    def test_numeric_class(self):
        row = pd.Series({'学年': '2', '組': 'ひまわり', '出席番号': '3'})
        assert make_student_key(row) == '2-ひまわり-3'


# ────────────────────────────────────────────────────────────────────────────
# get_assigned_students
# ────────────────────────────────────────────────────────────────────────────

class TestGetAssignedStudents:

    def _special_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {'学年': '1', '組': 'なかよし', '出席番号': '1', '氏名': '特支A'},
            {'学年': '1', '組': 'なかよし', '出席番号': '2', '氏名': '特支B'},
            {'学年': '2', '組': 'ひまわり', '出席番号': '1', '氏名': '特支C'},
        ])

    def test_returns_assigned_only(self):
        """指定クラスに割り当て済みの児童のみ返す。"""
        assignments = {
            '1-なかよし-1': '1-1',
            '1-なかよし-2': '1-2',
            '2-ひまわり-1': '2-1',
        }
        result = get_assigned_students(self._special_df(), assignments, '1-1')
        assert len(result) == 1
        assert result.iloc[0]['氏名'] == '特支A'

    def test_multiple_assigned(self):
        """同じクラスに複数児童割り当て。"""
        assignments = {
            '1-なかよし-1': '1-1',
            '1-なかよし-2': '1-1',
            '2-ひまわり-1': '2-1',
        }
        result = get_assigned_students(self._special_df(), assignments, '1-1')
        assert len(result) == 2
        assert set(result['氏名']) == {'特支A', '特支B'}

    def test_no_assignments(self):
        """割り当てなし → 空。"""
        result = get_assigned_students(self._special_df(), {}, '1-1')
        assert result.empty

    def test_no_match(self):
        """対象クラスへの割り当てなし → 空。"""
        assignments = {'1-なかよし-1': '1-2'}
        result = get_assigned_students(self._special_df(), assignments, '1-1')
        assert result.empty

    def test_empty_df(self):
        """空 DataFrame → 空。"""
        empty = pd.DataFrame(columns=['学年', '組', '出席番号', '氏名'])
        result = get_assigned_students(empty, {'x': '1-1'}, '1-1')
        assert result.empty

    def test_missing_columns(self):
        """必須列がない DataFrame → 空。"""
        df = pd.DataFrame([{'氏名': 'A'}])
        result = get_assigned_students(df, {'x': '1-1'}, '1-1')
        assert result.empty


# ────────────────────────────────────────────────────────────────────────────
# get_unassigned_students
# ────────────────────────────────────────────────────────────────────────────

class TestGetUnassignedStudents:

    def _special_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {'学年': '1', '組': 'なかよし', '出席番号': '1', '氏名': '特支A'},
            {'学年': '1', '組': 'なかよし', '出席番号': '2', '氏名': '特支B'},
            {'学年': '2', '組': 'ひまわり', '出席番号': '1', '氏名': '特支C'},
        ])

    def test_all_unassigned(self):
        """全員未割り当て。"""
        result = get_unassigned_students(self._special_df(), {})
        assert len(result) == 3

    def test_partial_assigned(self):
        """一部割り当て済み → 残りが返る。"""
        assignments = {'1-なかよし-1': '1-1'}
        result = get_unassigned_students(self._special_df(), assignments)
        assert len(result) == 2
        assert set(result['氏名']) == {'特支B', '特支C'}

    def test_all_assigned(self):
        """全員割り当て済み → 空。"""
        assignments = {
            '1-なかよし-1': '1-1',
            '1-なかよし-2': '1-2',
            '2-ひまわり-1': '2-1',
        }
        result = get_unassigned_students(self._special_df(), assignments)
        assert result.empty

    def test_empty_string_counts_as_unassigned(self):
        """空文字の割り当ては未割り当てと見なす。"""
        assignments = {'1-なかよし-1': ''}
        result = get_unassigned_students(self._special_df(), assignments)
        assert len(result) == 3

    def test_empty_df(self):
        """空 DataFrame → 空。"""
        empty = pd.DataFrame(columns=['学年', '組', '出席番号', '氏名'])
        result = get_unassigned_students(empty, {})
        assert result.empty

    def test_missing_columns(self):
        """必須列がない → 全行返す（フィルタ不能）。"""
        df = pd.DataFrame([{'氏名': 'A'}])
        result = get_unassigned_students(df, {})
        assert len(result) == 1
