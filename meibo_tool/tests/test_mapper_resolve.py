"""core/mapper.py resolve_name_fields() のテスト"""

from __future__ import annotations

from core.mapper import resolve_name_fields


class TestResolveNameFields:
    def test_formal_mode_uses_formal_name(self):
        data = {
            '氏名': '山田 太郎',
            '氏名かな': 'やまだ たろう',
            '正式氏名': '山田 太郞',  # 異体字
            '正式氏名かな': 'やまだ たろう',
            '保護者名': '山田 幸子',
            '保護者正式名': '山田 幸子',
            '保護者名かな': 'やまだ さちこ',
            '保護者正式名かな': 'やまだ さちこ',
        }
        result = resolve_name_fields(data, use_formal=True)
        assert result['表示氏名'] == '山田 太郞'
        assert result['表示氏名かな'] == 'やまだ たろう'
        assert result['表示保護者名'] == '山田 幸子'

    def test_normal_mode_uses_regular_name(self):
        data = {
            '氏名': '山田 太郎',
            '氏名かな': 'やまだ たろう',
            '正式氏名': '山田 太郞',
            '正式氏名かな': 'やまだ たろう',
            '保護者名': '山田 幸子',
            '保護者名かな': 'やまだ さちこ',
        }
        result = resolve_name_fields(data, use_formal=False)
        assert result['表示氏名'] == '山田 太郎'
        assert result['表示氏名かな'] == 'やまだ たろう'

    def test_formal_fallback_to_regular(self):
        """正式氏名が空の場合、通常氏名にフォールバックする。"""
        data = {
            '氏名': '田中 花子',
            '氏名かな': 'たなか はなこ',
            '正式氏名': '',
            '正式氏名かな': '',
            '保護者名': '田中 美穂',
            '保護者正式名': '',
            '保護者名かな': 'たなか みほ',
            '保護者正式名かな': '',
        }
        result = resolve_name_fields(data, use_formal=True)
        assert result['表示氏名'] == '田中 花子'
        assert result['表示氏名かな'] == 'たなか はなこ'
        assert result['表示保護者名'] == '田中 美穂'

    def test_formal_fallback_nan_value(self):
        """正式氏名が 'nan' の場合もフォールバック。"""
        data = {
            '氏名': '鈴木 健太',
            '氏名かな': 'すずき けんた',
            '正式氏名': 'nan',
            '正式氏名かな': 'nan',
            '保護者名': '鈴木 典子',
            '保護者正式名': 'nan',
            '保護者名かな': 'すずき のりこ',
            '保護者正式名かな': 'nan',
        }
        result = resolve_name_fields(data, use_formal=True)
        assert result['表示氏名'] == '鈴木 健太'
        assert result['表示氏名かな'] == 'すずき けんた'

    def test_missing_keys_return_empty(self):
        """キーが存在しない場合は空文字列。"""
        data = {}
        result = resolve_name_fields(data, use_formal=False)
        assert result['表示氏名'] == ''
        assert result['表示氏名かな'] == ''
        assert result['表示保護者名'] == ''
        assert result['表示保護者名かな'] == ''

    def test_none_value_treated_as_empty(self):
        """None 値は空文字列として扱われる。"""
        data = {
            '氏名': None,
            '氏名かな': None,
            '保護者名': None,
            '保護者名かな': None,
        }
        result = resolve_name_fields(data, use_formal=False)
        assert result['表示氏名'] == ''
