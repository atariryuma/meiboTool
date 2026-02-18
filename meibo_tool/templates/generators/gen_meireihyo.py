"""掲示用名列表テンプレート生成スクリプト

SPEC.md §4.4 参照。
A4縦・2列（左:No.1〜20、右:No.21〜40）・タイトル・薄ピンク背景。

テンプレート内のプレースホルダー形式:
  ヘッダー: {{学年}} {{組}} {{担任名}}
  データ:   {{出席番号_1}} 〜 {{出席番号_40}}
            {{氏名かな_1}} 〜 {{氏名かな_40}}
            {{氏名_1}}     〜 {{氏名_40}}

このファイルが生成する xlsx を GridGenerator が読み込み、_fill_numbered() で
番号付きプレースホルダーを実データに置換する。
"""

from __future__ import annotations

import os
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins

# ────────────────────────────────────────────────────────────────────────────
# スタイル定数（SPEC §4.0.2 参照）
# ────────────────────────────────────────────────────────────────────────────

FONT_FAMILY = 'IPAmj明朝'

_PINK_BG     = 'FFE0E0'  # 薄ピンク（データセル）
_HEADER_BG   = 'FFCCCC'  # やや濃いピンク（ヘッダー行）
_WHITE_BG    = 'FFFFFF'

FILL_PINK   = PatternFill(fill_type='solid', fgColor=_PINK_BG)
FILL_HEADER = PatternFill(fill_type='solid', fgColor=_HEADER_BG)

_THIN   = Side(style='thin',   color='000000')
_MEDIUM = Side(style='medium', color='000000')

BORDER_DATA = Border(top=_THIN, bottom=_THIN, left=_THIN, right=_THIN)

ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_RIGHT  = Alignment(horizontal='right',  vertical='center', wrap_text=True)

FONT_TITLE  = Font(name=FONT_FAMILY, size=18, bold=True)
FONT_HEADER = Font(name=FONT_FAMILY, size=9,  bold=True)
FONT_NO     = Font(name=FONT_FAMILY, size=11, bold=False)
FONT_NAME   = Font(name=FONT_FAMILY, size=14, bold=True)
FONT_KANA   = Font(name=FONT_FAMILY, size=9,  bold=False)

MAX_PER_COL = 20   # 1列あたりの最大人数
TOTAL_SLOTS = MAX_PER_COL * 2  # 40名分


# ────────────────────────────────────────────────────────────────────────────
# ヘルパー
# ────────────────────────────────────────────────────────────────────────────

def _ph(base: str, n: int) -> str:
    """番号付きプレースホルダー文字列を返す。例: _ph('氏名', 1) → '{{氏名_1}}'"""
    return '{{' + base + '_' + str(n) + '}}'


def _cell(ws, row: int, col: int, *,
          value=None, font=None, fill=None, border=None, alignment=None):
    """セルに値とスタイルを一括設定する。"""
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if font is not None:
        c.font = font
    if fill is not None:
        c.fill = fill
    if border is not None:
        c.border = border
    if alignment is not None:
        c.alignment = alignment
    return c


# ────────────────────────────────────────────────────────────────────────────
# テンプレート生成
# ────────────────────────────────────────────────────────────────────────────

def generate(output_path: str) -> None:
    """掲示用名列表テンプレートを output_path に保存する。

    生成されるシート構成:
      Row 1      : タイトル ({{学年}}年{{組}}組 名列表 担任:{{担任名}})
      Row 2      : 列ヘッダー (番号 / ふりがな / 氏名 / ─ / 番号 / ふりがな / 氏名)
      Row 3〜22  : データ行 × 20（左列 No.1〜20、右列 No.21〜40）

    列レイアウト (A〜G):
      A: 番号（左）  B: ふりがな（左）  C: 氏名（左）
      D: 区切り
      E: 番号（右）  F: ふりがな（右）  G: 氏名（右）
    """
    wb = Workbook()
    ws = wb.active
    ws.title = '名列表'

    # ── 列幅 ──────────────────────────────────────────────────────────────
    col_widths = {
        'A': 5.0,   # 番号（左）
        'B': 11.0,  # ふりがな（左）
        'C': 16.0,  # 氏名（左）
        'D': 2.0,   # 区切り
        'E': 5.0,   # 番号（右）
        'F': 11.0,  # ふりがな（右）
        'G': 16.0,  # 氏名（右）
    }
    for letter, width in col_widths.items():
        ws.column_dimensions[letter].width = width

    # ── Row 1: タイトル ───────────────────────────────────────────────────
    ws.merge_cells('A1:G1')
    _cell(ws, 1, 1,
          value='{{学年}}年{{組}}組　名列表　　担任：{{担任名}}',
          font=FONT_TITLE,
          alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 42

    # ── Row 2: 列ヘッダー ─────────────────────────────────────────────────
    header_labels = [
        (1, '番号'), (2, 'ふりがな'), (3, '氏名'),   # 左列
        (4, ''),                                       # 区切り
        (5, '番号'), (6, 'ふりがな'), (7, '氏名'),   # 右列
    ]
    for col, label in header_labels:
        kwargs: dict = dict(value=label, font=FONT_HEADER, alignment=ALIGN_CENTER)
        if col != 4:
            kwargs['fill'] = FILL_HEADER
            kwargs['border'] = BORDER_DATA
        _cell(ws, 2, col, **kwargs)
    ws.row_dimensions[2].height = 18

    # ── Rows 3〜22: データ行 ──────────────────────────────────────────────
    for i in range(MAX_PER_COL):
        row  = 3 + i
        ln   = i + 1          # 左列インデックス (1〜20)
        rn   = i + 21         # 右列インデックス (21〜40)

        # 左列
        _cell(ws, row, 1,
              value=_ph('出席番号', ln), font=FONT_NO,
              fill=FILL_PINK, border=BORDER_DATA, alignment=ALIGN_CENTER)
        _cell(ws, row, 2,
              value=_ph('氏名かな', ln), font=FONT_KANA,
              fill=FILL_PINK, border=BORDER_DATA, alignment=ALIGN_RIGHT)
        _cell(ws, row, 3,
              value=_ph('氏名', ln), font=FONT_NAME,
              fill=FILL_PINK, border=BORDER_DATA, alignment=ALIGN_CENTER)

        # 区切り列（スタイルなし）
        _cell(ws, row, 4)

        # 右列
        _cell(ws, row, 5,
              value=_ph('出席番号', rn), font=FONT_NO,
              fill=FILL_PINK, border=BORDER_DATA, alignment=ALIGN_CENTER)
        _cell(ws, row, 6,
              value=_ph('氏名かな', rn), font=FONT_KANA,
              fill=FILL_PINK, border=BORDER_DATA, alignment=ALIGN_RIGHT)
        _cell(ws, row, 7,
              value=_ph('氏名', rn), font=FONT_NAME,
              fill=FILL_PINK, border=BORDER_DATA, alignment=ALIGN_CENTER)

        ws.row_dimensions[row].height = 30

    # ── 印刷設定（SPEC §5.4）─────────────────────────────────────────────
    from openpyxl.worksheet.properties import PageSetupProperties

    ws.page_setup.paperSize   = 9          # A4
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0          # 高さは自動

    # fitToPage は sheet_properties 経由で設定しないと AttributeError になる
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(
        left=0.39, right=0.39,
        top=0.39,  bottom=0.39,
        header=0.2, footer=0.2,
    )
    ws.print_options.horizontalCentered = True

    # ── 保存 ─────────────────────────────────────────────────────────────
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)
    print(f'Generated: {output_path}')


# ────────────────────────────────────────────────────────────────────────────
# スタンドアロン実行
# ────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # gen_meireihyo.py → generators/ → templates/ → meibo_tool/ → project_root
    _project_root = Path(__file__).resolve().parents[3]
    _out = _project_root / 'テンプレート' / '掲示用名列表.xlsx'
    generate(str(_out))
