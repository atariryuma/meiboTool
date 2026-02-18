"""core/generator.py と templates/generators/gen_meireihyo.py のテスト

テスト対象:
  - gen_meireihyo.generate() が正しい構造の xlsx を生成すること
  - GridGenerator が番号付きプレースホルダーを実データに置換すること
  - GridGenerator がヘッダー行の非番号プレースホルダーも置換すること
  - fill_placeholders / setup_print の基本動作
"""

from __future__ import annotations

import os

import pandas as pd
import pytest
from openpyxl import load_workbook

from core.generator import (
    GridGenerator,
    fill_placeholders,
    setup_print,
)
from templates.generators import gen_meireihyo

# ────────────────────────────────────────────────────────────────────────────
# フィクスチャ
# ────────────────────────────────────────────────────────────────────────────

@pytest.fixture
def template_path(tmp_path) -> str:
    """テスト用に名列表テンプレートを一時ディレクトリに生成して返す。"""
    out = str(tmp_path / '掲示用名列表.xlsx')
    gen_meireihyo.generate(out)
    return out


@pytest.fixture
def large_df() -> pd.DataFrame:
    """GridGenerator 用の 25 名データ（2列目が 5 名）。"""
    rows = []
    for i in range(1, 26):
        rows.append({
            '生徒コード': f'S{i:06d}',
            '学年': '3',
            '組': '2',
            '出席番号': str(i),
            '氏名': f'テスト 太郎{i}',
            '氏名かな': f'てすと たろう{i}',
            '正式氏名': f'テスト 太郎{i}',
            '正式氏名かな': f'てすと たろう{i}',
            '性別': '男',
            '生年月日': '2015-04-01',
        })
    return pd.DataFrame(rows)


# ────────────────────────────────────────────────────────────────────────────
# gen_meireihyo: テンプレート生成の構造テスト
# ────────────────────────────────────────────────────────────────────────────

class TestGenMeireihyo:
    def test_generates_file(self, template_path):
        """generate() がファイルを生成すること。"""
        assert os.path.isfile(template_path)
        assert os.path.getsize(template_path) > 0

    def test_sheet_name(self, template_path):
        """シート名が '名列表' であること。"""
        wb = load_workbook(template_path)
        assert '名列表' in wb.sheetnames

    def test_row_count(self, template_path):
        """行数が 22（タイトル1 + ヘッダー1 + データ20）であること。"""
        wb = load_workbook(template_path)
        ws = wb.active
        assert ws.max_row == 22

    def test_column_count(self, template_path):
        """列数が 7（A〜G）であること。"""
        wb = load_workbook(template_path)
        ws = wb.active
        assert ws.max_column == 7

    def test_title_has_placeholders(self, template_path):
        """タイトルセルに {{学年}} {{組}} {{担任名}} が含まれること。"""
        wb = load_workbook(template_path)
        ws = wb.active
        title = ws['A1'].value or ''
        assert '{{学年}}' in title
        assert '{{組}}' in title
        assert '{{担任名}}' in title

    def test_first_data_row_placeholders(self, template_path):
        """Row 3（1番目のデータ行）が {{出席番号_1}} {{氏名_1}} {{氏名かな_1}} を持つこと。"""
        wb = load_workbook(template_path)
        ws = wb.active
        assert ws['A3'].value == '{{出席番号_1}}'
        assert ws['B3'].value == '{{氏名かな_1}}'
        assert ws['C3'].value == '{{氏名_1}}'

    def test_right_column_offset_placeholders(self, template_path):
        """Row 3 の右列が {{出席番号_21}} {{氏名_21}} {{氏名かな_21}} を持つこと。"""
        wb = load_workbook(template_path)
        ws = wb.active
        assert ws['E3'].value == '{{出席番号_21}}'
        assert ws['F3'].value == '{{氏名かな_21}}'
        assert ws['G3'].value == '{{氏名_21}}'

    def test_last_data_row_placeholders(self, template_path):
        """Row 22（最終データ行）が _20 と _40 のプレースホルダーを持つこと。"""
        wb = load_workbook(template_path)
        ws = wb.active
        assert ws['A22'].value == '{{出席番号_20}}'
        assert ws['C22'].value == '{{氏名_20}}'
        assert ws['E22'].value == '{{出席番号_40}}'
        assert ws['G22'].value == '{{氏名_40}}'

    def test_print_orientation(self, template_path):
        """印刷設定が縦（portrait）であること。"""
        wb = load_workbook(template_path)
        ws = wb.active
        assert ws.page_setup.orientation == 'portrait'

    def test_paper_size_a4(self, template_path):
        """用紙サイズが A4（9）であること。"""
        wb = load_workbook(template_path)
        ws = wb.active
        assert ws.page_setup.paperSize == 9


# ────────────────────────────────────────────────────────────────────────────
# GridGenerator: プレースホルダー置換テスト
# ────────────────────────────────────────────────────────────────────────────

class TestGridGeneratorMeireihyo:
    def _make_options(self, template_path: str) -> dict:
        return {
            'fiscal_year': 2025,
            'school_name': '那覇市立天久小学校',
            'teacher_name': '山田先生',
            'template_dir': os.path.dirname(template_path),
        }

    def test_output_file_created(self, template_path, dummy_df, tmp_path):
        """GridGenerator.generate() が出力ファイルを作成すること。"""
        out = str(tmp_path / 'output.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, dummy_df, options)
        result = gen.generate()
        assert result == out
        assert os.path.isfile(out)

    def test_student1_name_filled(self, template_path, dummy_df, tmp_path):
        """1番目の学生の氏名が {{氏名_1}} に正しく差し込まれること。"""
        out = str(tmp_path / 'output.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, dummy_df, options)
        gen.generate()

        wb = load_workbook(out)
        ws = wb.active
        assert ws['C3'].value == dummy_df.iloc[0]['氏名']

    def test_student1_shusseki_filled(self, template_path, dummy_df, tmp_path):
        """1番目の学生の出席番号が {{出席番号_1}} に正しく差し込まれること。"""
        out = str(tmp_path / 'output.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, dummy_df, options)
        gen.generate()

        wb = load_workbook(out)
        ws = wb.active
        assert ws['A3'].value == dummy_df.iloc[0]['出席番号']

    def test_student5_name_filled(self, template_path, dummy_df, tmp_path):
        """5番目の学生の氏名が {{氏名_5}} に差し込まれること。"""
        out = str(tmp_path / 'output.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, dummy_df, options)
        gen.generate()

        wb = load_workbook(out)
        ws = wb.active
        assert ws['C7'].value == dummy_df.iloc[4]['氏名']

    def test_header_placeholders_filled(self, template_path, dummy_df, tmp_path):
        """タイトル行の {{学年}} {{組}} {{担任名}} が置換されること。"""
        out = str(tmp_path / 'output.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, dummy_df, options)
        gen.generate()

        wb = load_workbook(out)
        ws = wb.active
        title = ws['A1'].value or ''
        assert '{{学年}}' not in title
        assert '{{組}}' not in title
        assert '{{担任名}}' not in title
        assert '山田先生' in title

    def test_unused_slots_cleared(self, template_path, dummy_df, tmp_path):
        """5名データでは No.6 以降のスロットが空（None）になること。
        fill_placeholders が未マッチの番号付きプレースホルダーを空文字に変換する。"""
        out = str(tmp_path / 'output.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, dummy_df, options)
        gen.generate()

        wb = load_workbook(out)
        ws = wb.active
        # Row 8 は 6 番目のスロット → データなし → 空セル（None）になる
        assert ws['C8'].value is None or ws['C8'].value == ''

    def test_right_column_cleared_for_small_data(
            self, template_path, dummy_df, tmp_path):
        """5名データでは右列（No.21〜）のスロットが空になること。"""
        out = str(tmp_path / 'output.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, dummy_df, options)
        gen.generate()

        wb = load_workbook(out)
        ws = wb.active
        # 右列の先頭 E3 は No.21 に対応するが 5 名しかいない → 空
        assert ws['E3'].value is None or ws['E3'].value == ''

    def test_25_students_right_column_partially_filled(
            self, template_path, large_df, tmp_path):
        """25名データでは右列の No.21〜25 が正しく差し込まれること。"""
        out = str(tmp_path / 'output_25.xlsx')
        options = self._make_options(template_path)
        gen = GridGenerator(template_path, out, large_df, options)
        gen.generate()

        wb = load_workbook(out)
        ws = wb.active
        # 右列 Row 3 (21番目) の氏名
        assert ws['G3'].value == large_df.iloc[20]['氏名']
        # 右列 Row 7 (25番目) の氏名
        assert ws['G7'].value == large_df.iloc[24]['氏名']
        # 右列 Row 8 (26番目) はデータなし → 空セル
        assert ws['G8'].value is None or ws['G8'].value == ''


# ────────────────────────────────────────────────────────────────────────────
# fill_placeholders: 特殊キーのテスト
# ────────────────────────────────────────────────────────────────────────────

class TestFillPlaceholders:
    def _make_ws(self):
        """テスト用の 1 セルだけのワークシートを返す。"""
        from openpyxl import Workbook
        wb = Workbook()
        return wb.active

    def test_fiscal_year(self):
        ws = self._make_ws()
        ws['A1'] = '{{年度}}'
        fill_placeholders(ws, {}, {'fiscal_year': 2025})
        assert ws['A1'].value == '2025'

    def test_school_name(self):
        ws = self._make_ws()
        ws['A1'] = '{{学校名}}'
        fill_placeholders(ws, {}, {'school_name': 'テスト小学校'})
        assert ws['A1'].value == 'テスト小学校'

    def test_teacher_name(self):
        ws = self._make_ws()
        ws['A1'] = '{{担任名}}'
        fill_placeholders(ws, {}, {'teacher_name': '鈴木先生'})
        assert ws['A1'].value == '鈴木先生'

    def test_wareki_nendo(self):
        ws = self._make_ws()
        ws['A1'] = '{{年度和暦}}'
        fill_placeholders(ws, {}, {'fiscal_year': 2025})
        assert ws['A1'].value == '令和7年度'

    def test_address_combined(self):
        ws = self._make_ws()
        ws['A1'] = '{{住所}}'
        data = {
            '都道府県': '沖縄県', '市区町村': '那覇市',
            '町番地': '天久1-2-3', '建物名': '',
        }
        fill_placeholders(ws, data, {})
        assert ws['A1'].value == '沖縄県那覇市天久1-2-3'

    def test_nan_value_becomes_empty(self):
        ws = self._make_ws()
        ws['A1'] = '{{氏名}}'
        fill_placeholders(ws, {'氏名': 'nan'}, {})
        assert ws['A1'].value == ''


# ────────────────────────────────────────────────────────────────────────────
# setup_print: 印刷設定テスト
# ────────────────────────────────────────────────────────────────────────────

class TestSetupPrint:
    def _make_ws(self):
        from openpyxl import Workbook
        return Workbook().active

    def test_portrait_orientation(self):
        ws = self._make_ws()
        setup_print(ws, orientation='portrait')
        assert ws.page_setup.orientation == 'portrait'

    def test_landscape_orientation(self):
        ws = self._make_ws()
        setup_print(ws, orientation='landscape')
        assert ws.page_setup.orientation == 'landscape'

    def test_paper_size_a4(self):
        ws = self._make_ws()
        setup_print(ws)
        assert ws.page_setup.paperSize == 9

    def test_fit_to_page(self):
        ws = self._make_ws()
        setup_print(ws)
        # fitToPage は sheet_properties.pageSetUpPr 経由で設定される
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True
        assert ws.page_setup.fitToWidth == 1
