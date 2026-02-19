"""core/importer.py のテスト

テスト対象:
  - detect_header_row: ヘッダー行自動検出
  - import_c4th_excel: C4th Excel の読み込み + カラムマッピング
"""

from __future__ import annotations

from openpyxl import Workbook

from core.importer import detect_header_row, import_c4th_excel


def _create_excel(path, rows: list[list], start_row: int = 1) -> str:
    """テスト用 Excel ファイルを作成する。"""
    wb = Workbook()
    ws = wb.active
    for i, row_data in enumerate(rows, start_row):
        for j, val in enumerate(row_data, 1):
            ws.cell(row=i, column=j, value=val)
    wb.save(str(path))
    return str(path)


# ── detect_header_row ─────────────────────────────────────────────────────────


class TestDetectHeaderRow:
    def test_header_on_first_row(self, tmp_path):
        """行1にヘッダーがある場合。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['名前', 'ふりがな', '性別', '生年月日', '学年', '電話番号1'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1', '098-123-4567'],
        ])
        assert detect_header_row(path) == 1

    def test_header_on_second_row(self, tmp_path):
        """行1にメタ情報、行2にヘッダーがある場合。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['C4th エクスポート', None, None, None, None],
            ['名前', 'ふりがな', '性別', '生年月日', '学年', '電話番号1'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1', '098-123-4567'],
        ])
        assert detect_header_row(path) == 2

    def test_header_on_third_row(self, tmp_path):
        """行1-2にメタ情報、行3にヘッダーがある場合。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['学校名: テスト小学校', None, None, None],
            ['出力日: 2025-01-01', None, None, None],
            ['名前', 'ふりがな', '性別', '生年月日', '学年'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1'],
        ])
        assert detect_header_row(path) == 3

    def test_no_header_defaults_to_1(self, tmp_path):
        """文字列セルが5つ未満の場合はデフォルト1を返す。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            [1, 2, 3],
            [4, 5, 6],
        ])
        assert detect_header_row(path) == 1

    def test_numeric_meta_rows_skipped(self, tmp_path):
        """数値のみの行はスキップされる。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            [2025, 1, 15, None, None],  # 数値メタ行
            ['名前', 'ふりがな', '性別', '生年月日', '学年'],
        ])
        assert detect_header_row(path) == 2


# ── import_c4th_excel ─────────────────────────────────────────────────────────


class TestImportC4thExcel:
    def test_basic_import(self, tmp_path):
        """基本的な読み込みとカラムマッピング。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['名前', 'ふりがな', '性別', '生年月日', '学年'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1'],
            ['田中花子', 'たなかはなこ', '女', '2018-04-03', '1'],
        ])
        df, unmapped = import_c4th_excel(path)

        assert len(df) == 2
        assert '氏名' in df.columns  # '名前' → '氏名' にマッピング
        assert '氏名かな' in df.columns
        assert df.iloc[0]['氏名'] == '山田太郎'

    def test_all_columns_are_string(self, tmp_path):
        """全カラムが文字列型で読み込まれる。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['名前', 'ふりがな', '性別', '生年月日', '学年'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1'],
        ])
        df, _ = import_c4th_excel(path)
        for col in df.columns:
            assert df[col].dtype == object  # pandas str dtype

    def test_unmapped_columns_reported(self, tmp_path):
        """マッピングできないカラムが unmapped に含まれる。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['名前', 'ふりがな', '性別', '生年月日', '学年', '独自カラム'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1', 'xxx'],
        ])
        _, unmapped = import_c4th_excel(path)
        assert '独自カラム' in unmapped

    def test_header_with_meta_rows(self, tmp_path):
        """メタ情報行ありのファイルを正しく読み込む。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['C4th データエクスポート', None, None, None, None],
            ['名前', 'ふりがな', '性別', '生年月日', '学年'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1'],
        ])
        df, _ = import_c4th_excel(path)
        assert len(df) == 1
        assert '氏名' in df.columns

    def test_empty_rows_dropped(self, tmp_path):
        """空白行が除去される。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['名前', 'ふりがな', '性別', '生年月日', '学年'],
            ['山田太郎', 'やまだたろう', '男', '2018-01-01', '1'],
            [None, None, None, None, None],
            ['田中花子', 'たなかはなこ', '女', '2018-04-03', '1'],
        ])
        df, _ = import_c4th_excel(path)
        assert len(df) == 2

    def test_fullwidth_space_header(self, tmp_path):
        """全角スペース付きヘッダー（C4th形式）が正しくマッピングされる。"""
        path = _create_excel(tmp_path / 'test.xlsx', [
            ['名前', 'ふりがな', '性別', '保護者1\u3000続柄', '保護者1\u3000名前'],
            ['山田太郎', 'やまだたろう', '男', '母', '山田幸子'],
        ])
        df, _ = import_c4th_excel(path)
        assert '保護者続柄' in df.columns
        assert '保護者名' in df.columns
