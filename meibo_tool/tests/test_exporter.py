"""core/exporter.py のテスト"""

from __future__ import annotations

import pandas as pd

from core.exporter import export_csv, export_excel


def _sample_df() -> pd.DataFrame:
    return pd.DataFrame({
        '氏名': ['山田太郎', '田中花子'],
        '性別': ['男', '女'],
        '学年': ['1', '2'],
    })


class TestExportCSV:
    def test_basic_export(self, tmp_path):
        """CSV ファイルが正しく書き出される。"""
        path = tmp_path / 'out.csv'
        export_csv(_sample_df(), str(path))
        content = path.read_text(encoding='utf-8-sig')
        assert '山田太郎' in content
        assert '田中花子' in content

    def test_utf8_bom(self, tmp_path):
        """デフォルトで UTF-8 BOM 付きで書き出される。"""
        path = tmp_path / 'out.csv'
        export_csv(_sample_df(), str(path))
        raw = path.read_bytes()
        assert raw[:3] == b'\xef\xbb\xbf'  # UTF-8 BOM

    def test_shift_jis(self, tmp_path):
        """Shift_JIS エンコーディングで書き出せる。"""
        path = tmp_path / 'out.csv'
        export_csv(_sample_df(), str(path), encoding='shift_jis')
        content = path.read_text(encoding='shift_jis')
        assert '山田太郎' in content

    def test_roundtrip(self, tmp_path):
        """書き出した CSV を読み戻して同一内容を確認。"""
        path = tmp_path / 'out.csv'
        df = _sample_df()
        export_csv(df, str(path))
        df2 = pd.read_csv(str(path), dtype=str, encoding='utf-8-sig')
        assert list(df2['氏名']) == ['山田太郎', '田中花子']


class TestExportExcel:
    def test_basic_export(self, tmp_path):
        """Excel ファイルが正しく書き出される。"""
        path = tmp_path / 'out.xlsx'
        export_excel(_sample_df(), str(path))
        assert path.exists()

    def test_roundtrip(self, tmp_path):
        """書き出した Excel を読み戻して同一内容を確認。"""
        path = tmp_path / 'out.xlsx'
        df = _sample_df()
        export_excel(df, str(path))
        df2 = pd.read_excel(str(path), dtype=str)
        assert list(df2['氏名']) == ['山田太郎', '田中花子']
