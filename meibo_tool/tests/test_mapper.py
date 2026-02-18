"""core/mapper.py のユニットテスト"""

import pandas as pd

from core.mapper import EXACT_MAP, map_columns, normalize_header


class TestNormalizeHeader:
    def test_strip_spaces(self):
        assert normalize_header('  名前  ') == '名前'

    def test_zenkaku_space(self):
        # 全角スペースが U+3000 に統一されること
        result = normalize_header('保護者1\u3000続柄')
        assert result == '保護者1\u3000続柄'

    def test_hankaku_to_zenkaku(self):
        # 全角スペース「　」が U+3000 に変換されること
        result = normalize_header('保護者1　続柄')  # 全角スペース
        assert result == '保護者1\u3000続柄'


class TestMapColumns:
    def test_exact_match_all(self):
        """全 EXACT_MAP キーが正しくマッピングされること"""
        df = pd.DataFrame(columns=list(EXACT_MAP.keys()))
        df_mapped, unmapped = map_columns(df)
        assert len(unmapped) == 0
        for _original, logical in EXACT_MAP.items():
            assert logical in df_mapped.columns

    def test_guardian_zenkaku_space(self):
        """保護者カラムの全角スペースが正しく処理されること"""
        df = pd.DataFrame(columns=['保護者1\u3000続柄', '保護者1\u3000名前'])
        df_mapped, unmapped = map_columns(df)
        assert '保護者続柄' in df_mapped.columns
        assert '保護者名' in df_mapped.columns
        assert len(unmapped) == 0

    def test_unknown_columns(self):
        df = pd.DataFrame(columns=['未知のカラム', '学年'])
        df_mapped, unmapped = map_columns(df)
        assert '未知のカラム' in unmapped
        assert '学年' in df_mapped.columns
