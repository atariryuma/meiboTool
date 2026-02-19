"""core/generator.py と templates/generators/gen_meireihyo.py のテスト

テスト対象:
  - gen_meireihyo.generate() が正しい構造の xlsx を生成すること
  - GridGenerator が番号付きプレースホルダーを実データに置換すること
  - fill_placeholders / setup_print の基本動作
  - name_display モード（furigana/kanji/kana）の差込動作
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
def tmpl_meireihyo(tmp_path) -> str:
    """掲示用名列表テンプレート（単一・全モード共通）を一時ディレクトリに生成して返す。"""
    out = str(tmp_path / '掲示用名列表.xlsx')
    gen_meireihyo.generate(out)
    return out


@pytest.fixture
def large_df() -> pd.DataFrame:
    """GridGenerator 用の 25 名データ（右列に 5 名入る）。"""
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
# gen_meireihyo: テンプレート生成構造テスト
# ────────────────────────────────────────────────────────────────────────────

class TestGenMeireihyo:
    """単一テンプレートの共通構造テスト。"""

    def test_generates_file(self, tmpl_meireihyo):
        assert os.path.isfile(tmpl_meireihyo)
        assert os.path.getsize(tmpl_meireihyo) > 0

    def test_sheet_name(self, tmpl_meireihyo):
        wb = load_workbook(tmpl_meireihyo)
        assert '名列表' in wb.sheetnames

    def test_title_has_placeholders(self, tmpl_meireihyo):
        wb = load_workbook(tmpl_meireihyo)
        title = wb.active['A1'].value or ''
        assert '{{学年}}' in title
        assert '{{組}}' in title
        assert '{{担任名}}' in title

    def test_portrait_a4(self, tmpl_meireihyo):
        ws = load_workbook(tmpl_meireihyo).active
        assert ws.page_setup.orientation == 'portrait'
        assert ws.page_setup.paperSize == 9

    def test_row_count(self, tmpl_meireihyo):
        """Row 1: title, Row 2: header, Rows 3–42: 20 pairs × 2 = 40 data rows"""
        ws = load_workbook(tmpl_meireihyo).active
        assert ws.max_row == 42

    def test_kana_row_placeholder(self, tmpl_meireihyo):
        """かな行（Row 3）に {{氏名かな_1}} があること。"""
        ws = load_workbook(tmpl_meireihyo).active
        assert ws['B3'].value == '{{氏名かな_1}}'    # 左かな行
        assert ws['E3'].value == '{{氏名かな_21}}'   # 右かな行

    def test_name_row_placeholder(self, tmpl_meireihyo):
        """氏名行（Row 4）に {{氏名_1}} があること。"""
        ws = load_workbook(tmpl_meireihyo).active
        assert ws['B4'].value == '{{氏名_1}}'    # 左氏名行
        assert ws['E4'].value == '{{氏名_21}}'   # 右氏名行

    def test_no_col_placeholder(self, tmpl_meireihyo):
        """番号列 A3（2行マージ）に {{出席番号_1}} があること。"""
        ws = load_workbook(tmpl_meireihyo).active
        assert ws['A3'].value == '{{出席番号_1}}'
        assert ws['A4'].value is None  # マージ下側セルは None

    def test_last_student_placeholders(self, tmpl_meireihyo):
        """20 番目の学生（Rows 41–42）のプレースホルダーが正しいこと。"""
        ws = load_workbook(tmpl_meireihyo).active
        assert ws['B41'].value == '{{氏名かな_20}}'
        assert ws['B42'].value == '{{氏名_20}}'
        assert ws['E41'].value == '{{氏名かな_40}}'
        assert ws['E42'].value == '{{氏名_40}}'

    def test_kana_row_height_small(self, tmpl_meireihyo):
        """かな行の行高が 15pt 以下（9pt フォントが収まり空白時も目立たないサイズ）であること。"""
        ws = load_workbook(tmpl_meireihyo).active
        assert ws.row_dimensions[3].height <= 15  # KANA_HEIGHT = 11


# ────────────────────────────────────────────────────────────────────────────
# GridGenerator: name_display モード別の差込テスト
# ────────────────────────────────────────────────────────────────────────────

class TestGridGeneratorMeireihyo:
    """掲示用名列表テンプレートへの GridGenerator 差込テスト。

    行構成（2行/人）:
      Person n: かな行 = row 3 + (n-1)*2
                氏名行 = row 4 + (n-1)*2
    """

    def _options(self, tmpl_path: str, name_display: str = 'furigana') -> dict:
        return {
            'fiscal_year': 2025,
            'school_name': '那覇市立天久小学校',
            'teacher_name': '山田先生',
            'template_dir': os.path.dirname(tmpl_path),
            'name_display': name_display,
        }

    def test_output_file_created(self, tmpl_meireihyo, dummy_df, tmp_path):
        out = str(tmp_path / 'output.xlsx')
        gen = GridGenerator(tmpl_meireihyo, out, dummy_df,
                            self._options(tmpl_meireihyo))
        assert gen.generate() == out
        assert os.path.isfile(out)

    # ── furigana モード ──────────────────────────────────────────────────────

    def test_furigana_kana_row_filled(self, tmpl_meireihyo, dummy_df, tmp_path):
        """furigana: かな行 B3 にかな値が入ること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo, 'furigana')).generate()
        ws = load_workbook(out).active
        assert ws['B3'].value == dummy_df.iloc[0]['氏名かな']

    def test_furigana_name_row_filled(self, tmpl_meireihyo, dummy_df, tmp_path):
        """furigana: 氏名行 B4 に漢字値が入ること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo, 'furigana')).generate()
        ws = load_workbook(out).active
        assert ws['B4'].value == dummy_df.iloc[0]['氏名']

    # ── kanji モード ─────────────────────────────────────────────────────────

    def test_kanji_kana_row_blank(self, tmpl_meireihyo, dummy_df, tmp_path):
        """kanji: かな行 B3 が空白になること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo, 'kanji')).generate()
        ws = load_workbook(out).active
        assert ws['B3'].value in (None, '')

    def test_kanji_name_row_filled(self, tmpl_meireihyo, dummy_df, tmp_path):
        """kanji: 氏名行 B4 に漢字値が入ること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo, 'kanji')).generate()
        ws = load_workbook(out).active
        assert ws['B4'].value == dummy_df.iloc[0]['氏名']

    # ── kana モード ──────────────────────────────────────────────────────────

    def test_kana_kana_row_blank(self, tmpl_meireihyo, dummy_df, tmp_path):
        """kana: かな行 B3 が空白になること（かな値は下段に転写）。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo, 'kana')).generate()
        ws = load_workbook(out).active
        assert ws['B3'].value in (None, '')

    def test_kana_name_row_shows_kana(self, tmpl_meireihyo, dummy_df, tmp_path):
        """kana: 氏名行 B4 にかな値が転写されること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo, 'kana')).generate()
        ws = load_workbook(out).active
        assert ws['B4'].value == dummy_df.iloc[0]['氏名かな']

    # ── 共通：ヘッダー・未使用スロット・右列 ─────────────────────────────────

    def test_header_placeholders_filled(self, tmpl_meireihyo, dummy_df, tmp_path):
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo)).generate()
        title = load_workbook(out).active['A1'].value or ''
        assert '{{学年}}' not in title
        assert '{{組}}' not in title
        assert '山田先生' in title

    def test_no_col_filled(self, tmpl_meireihyo, dummy_df, tmp_path):
        """番号列 A3 に出席番号が入ること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo)).generate()
        ws = load_workbook(out).active
        assert ws['A3'].value == dummy_df.iloc[0]['出席番号']

    def test_unused_slots_blank(self, tmpl_meireihyo, dummy_df, tmp_path):
        """5名データでは No.6 のかな行 B13 が空になること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo)).generate()
        ws = load_workbook(out).active
        # Person5 → B11(kana), B12(name);  Person6 → B13(kana) = 空
        assert ws['B13'].value in (None, '')

    def test_right_column_blank_for_small_data(self, tmpl_meireihyo, dummy_df, tmp_path):
        """5名データでは右列 E3（No.21 のかな行）が空になること。"""
        out = str(tmp_path / 'output.xlsx')
        GridGenerator(tmpl_meireihyo, out, dummy_df,
                      self._options(tmpl_meireihyo)).generate()
        ws = load_workbook(out).active
        assert ws['E3'].value in (None, '')

    def test_25_students_right_column_filled(self, tmpl_meireihyo, large_df, tmp_path):
        """25名データでは右列 No.21–25 の氏名行が差し込まれること。"""
        out = str(tmp_path / 'output_25.xlsx')
        GridGenerator(tmpl_meireihyo, out, large_df,
                      self._options(tmpl_meireihyo, 'kanji')).generate()
        ws = load_workbook(out).active
        # kanji モード: かな行=空, 氏名行=漢字値
        # No.21 → E3(kana=blank), E4(name=iloc[20])
        # No.25 → E11(kana=blank), E12(name=iloc[24])
        assert ws['E3'].value in (None, '')           # No.21 かな行（kanji→空）
        assert ws['E4'].value == large_df.iloc[20]['氏名']   # No.21 氏名行
        assert ws['E12'].value == large_df.iloc[24]['氏名']  # No.25 氏名行
        assert ws['E13'].value in (None, '')          # No.26 かな行（空データ）


# ────────────────────────────────────────────────────────────────────────────
# fill_placeholders: 特殊キーのテスト
# ────────────────────────────────────────────────────────────────────────────

class TestFillPlaceholders:
    def _make_ws(self):
        from openpyxl import Workbook
        return Workbook().active

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

    # ── name_display モードテスト ──────────────────────────────────────────

    def test_furigana_mode_fills_both(self):
        """furigana モードでは氏名・氏名かなの両方が展開される。"""
        ws = self._make_ws()
        ws['A1'] = '{{氏名かな}}'
        ws['A2'] = '{{氏名}}'
        data = {'氏名': '山田 太郎', '氏名かな': 'やまだ たろう'}
        fill_placeholders(ws, data, {'name_display': 'furigana'})
        assert ws['A1'].value == 'やまだ たろう'
        assert ws['A2'].value == '山田 太郎'

    def test_kanji_mode_blanks_kana_field(self):
        """kanji モードでは {{氏名かな}} が空白になる。"""
        ws = self._make_ws()
        ws['A1'] = '{{氏名かな}}'
        ws['A2'] = '{{氏名}}'
        data = {'氏名': '山田 太郎', '氏名かな': 'やまだ たろう'}
        fill_placeholders(ws, data, {'name_display': 'kanji'})
        assert ws['A1'].value == ''
        assert ws['A2'].value == '山田 太郎'

    def test_kana_mode_kana_row_blank(self):
        """kana モードでは {{氏名かな}} が空白になる（かな行を細く保つ）。"""
        ws = self._make_ws()
        ws['A1'] = '{{氏名かな}}'
        ws['A2'] = '{{氏名}}'
        data = {'氏名': '山田 太郎', '氏名かな': 'やまだ たろう'}
        fill_placeholders(ws, data, {'name_display': 'kana'})
        assert ws['A1'].value == ''

    def test_kana_mode_puts_kana_in_name_field(self):
        """kana モードでは {{氏名}} にかな値が転写される（氏名行に大きく表示）。"""
        ws = self._make_ws()
        ws['A1'] = '{{氏名かな}}'
        ws['A2'] = '{{氏名}}'
        data = {'氏名': '山田 太郎', '氏名かな': 'やまだ たろう'}
        fill_placeholders(ws, data, {'name_display': 'kana'})
        assert ws['A2'].value == 'やまだ たろう'

    def test_default_mode_same_as_furigana(self):
        """name_display 未指定（デフォルト）は furigana と同じ動作。"""
        ws = self._make_ws()
        ws['A1'] = '{{氏名かな}}'
        ws['A2'] = '{{氏名}}'
        data = {'氏名': '田中 花子', '氏名かな': 'たなか はなこ'}
        fill_placeholders(ws, data, {})
        assert ws['A1'].value == 'たなか はなこ'
        assert ws['A2'].value == '田中 花子'

    def test_kanji_mode_blanks_formal_kana(self):
        """kanji モードでは {{正式氏名かな}} も空白になる。"""
        ws = self._make_ws()
        ws['A1'] = '{{正式氏名かな}}'
        ws['A2'] = '{{正式氏名}}'
        data = {'正式氏名': '山田 太郎', '正式氏名かな': 'やまだ たろう'}
        fill_placeholders(ws, data, {'name_display': 'kanji'})
        assert ws['A1'].value == ''
        assert ws['A2'].value == '山田 太郎'

    def test_kana_mode_puts_formal_kana_in_formal_name_field(self):
        """kana モードでは {{正式氏名}} に正式氏名かなの値が転写される。"""
        ws = self._make_ws()
        ws['A1'] = '{{正式氏名かな}}'
        ws['A2'] = '{{正式氏名}}'
        data = {'正式氏名': '山田 太郎', '正式氏名かな': 'やまだ たろう'}
        fill_placeholders(ws, data, {'name_display': 'kana'})
        assert ws['A1'].value == ''
        assert ws['A2'].value == 'やまだ たろう'


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
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True
        assert ws.page_setup.fitToWidth == 1
