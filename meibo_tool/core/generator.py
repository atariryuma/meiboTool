"""テンプレートへのデータ差込処理

SPEC.md 第5章 §5.1〜§5.9 参照。
3 種類のジェネレーター + 共通プレースホルダー置換。
"""

from __future__ import annotations

import os
import re
import shutil
from abc import ABC, abstractmethod
from copy import copy

import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

from templates.template_registry import TEMPLATES
from utils.address import build_address
from utils.font_helper import apply_font
from utils.wareki import to_wareki

PLACEHOLDER_RE = re.compile(r'\{\{(.+?)\}\}')


# ────────────────────────────────────────────────────────────────────────────
# プレースホルダー置換（共通）
# ────────────────────────────────────────────────────────────────────────────

def fill_placeholders(ws, data_row: dict | pd.Series, options: dict | None = None) -> None:
    """
    ワークシート内の全プレースホルダー {{...}} をデータで置換する。

    特殊キー:
        {{年度}}       → options['fiscal_year']
        {{年度和暦}}   → 和暦表記の年度
        {{学校名}}     → options['school_name']
        {{担任名}}     → options['teacher_name']
        {{住所}}       → 都道府県+市区町村+町番地+建物名 の結合
    """
    options = options or {}

    def _val(key: str) -> str:
        v = data_row.get(key, '')
        if v is None:
            return ''
        s = str(v).strip()
        return '' if s.lower() == 'nan' else s

    def replacer(match: re.Match) -> str:
        key = match.group(1)
        if key == '年度':
            return str(options.get('fiscal_year', ''))
        if key == '年度和暦':
            fy = options.get('fiscal_year', 2025)
            return to_wareki(fy, 4, 1).replace('年', '年度')
        if key == '学校名':
            return options.get('school_name', '')
        if key == '担任名':
            return options.get('teacher_name', '')
        if key == '住所':
            return build_address(data_row)
        return _val(key)

    from openpyxl.cell.cell import MergedCell
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value and isinstance(cell.value, str) and '{{' in cell.value:
                cell.value = PLACEHOLDER_RE.sub(replacer, cell.value)


# ────────────────────────────────────────────────────────────────────────────
# 画像付きシート複製
# ────────────────────────────────────────────────────────────────────────────

def copy_sheet_with_images(wb, source_ws, new_title: str):
    """
    画像も含めてワークシートを複製する。
    openpyxl の copy_worksheet は画像をコピーしないため手動で再挿入する。
    """
    target_ws = wb.copy_worksheet(source_ws)
    target_ws.title = new_title
    for img in source_ws._images:
        new_img = Image(img.ref)
        new_img.anchor = copy(img.anchor)
        new_img.width = img.width
        new_img.height = img.height
        target_ws.add_image(new_img)
    return target_ws


# ────────────────────────────────────────────────────────────────────────────
# 印刷設定
# ────────────────────────────────────────────────────────────────────────────

def setup_print(ws, orientation: str = 'portrait', paper_size: int = 9) -> None:
    """
    印刷設定を適用する。
    paper_size: 9=A4
    orientation: 'portrait' or 'landscape'
    """
    from openpyxl.worksheet.page import PageMargins

    ws.page_setup.paperSize = paper_size
    ws.page_setup.orientation = orientation
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.page_margins = PageMargins(
        left=0.39, right=0.39,
        top=0.39, bottom=0.39,
        header=0.2, footer=0.2,
    )
    ws.print_options.horizontalCentered = True


# ────────────────────────────────────────────────────────────────────────────
# ベースクラス
# ────────────────────────────────────────────────────────────────────────────

class BaseGenerator(ABC):
    def __init__(
        self,
        template_path: str,
        output_path: str,
        data: pd.DataFrame,
        options: dict,
    ) -> None:
        self.template_path = template_path
        self.output_path = output_path
        self.data = data
        self.options = options
        self.wb = None

    def generate(self) -> str:
        """テンプレートにデータを差込み、output_path に保存して返す。"""
        # テンプレートをコピーしてから開く（元ファイル保護＋画像消失防止）
        shutil.copy2(self.template_path, self.output_path)
        self.wb = load_workbook(self.output_path)
        self._populate()
        self._apply_font_all()
        self.wb.save(self.output_path)
        self.wb.close()
        return self.output_path

    def _apply_font_all(self) -> None:
        for ws in self.wb.worksheets:
            apply_font(ws)

    @abstractmethod
    def _populate(self) -> None:
        pass


# ────────────────────────────────────────────────────────────────────────────
# GridGenerator — 名札系（1 ページに N 名をグリッド配置）
# ────────────────────────────────────────────────────────────────────────────

class GridGenerator(BaseGenerator):
    """
    テンプレートの 1 枚目シートに「1 カード分」のレイアウトが定義されており、
    それを児童数分コピーして複数ページに配置する。
    実際の配置ロジックはテンプレート生成スクリプト側で行う方式のため、
    ここでは fill_placeholders のみ適用する簡易実装。
    """

    def _populate(self) -> None:
        ws = self.wb.active
        meta = self._get_template_meta()
        orientation = meta.get('orientation', 'portrait')
        setup_print(ws, orientation=orientation)

        # プレースホルダーを全児童分で順番に置換（テンプレート内に {{氏名_1}} 等の番号付き形式）
        for i, (_, row) in enumerate(self.data.iterrows(), 1):
            row_dict = row.to_dict()
            # 番号付きプレースホルダーに変換して置換
            _fill_numbered(ws, i, row_dict, self.options)

    def _get_template_meta(self) -> dict:
        for _name, meta in TEMPLATES.items():
            if meta['file'] == os.path.basename(self.template_path):
                return meta
        return {}


def _fill_numbered(ws, idx: int, data_row: dict, options: dict) -> None:
    """{{氏名_1}}, {{氏名_2}} ... 形式の番号付きプレースホルダーを置換する。"""
    from openpyxl.cell.cell import MergedCell

    def _val(key: str) -> str:
        v = data_row.get(key, '')
        if v is None:
            return ''
        s = str(v).strip()
        return '' if s.lower() == 'nan' else s

    suffix = f'_{idx}'
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value and isinstance(cell.value, str) and '{{' in cell.value:
                def replacer(match: re.Match) -> str:
                    key = match.group(1)
                    if key.endswith(suffix):
                        base_key = key[: -len(suffix)]
                        if base_key == '住所':
                            return build_address(data_row)
                        return _val(base_key)
                    return match.group(0)  # 他のインデックスは変更しない
                cell.value = PLACEHOLDER_RE.sub(replacer, cell.value)


# ────────────────────────────────────────────────────────────────────────────
# ListGenerator — 名列表・台帳系（データ行を児童数分コピー展開）
# ────────────────────────────────────────────────────────────────────────────

class ListGenerator(BaseGenerator):
    """
    テンプレート内の {{行開始}} 〜 {{行終了}} マーカーで囲まれた行を
    児童数分コピー展開する。
    マーカーがない場合は最初の {{氏名}} が含まれる行をテンプレート行として扱う。
    """

    def _populate(self) -> None:
        ws = self.wb.active
        meta = self._get_template_meta()
        orientation = meta.get('orientation', 'portrait')
        setup_print(ws, orientation=orientation)

        # ヘッダー行・フッター行を特定してデータ行を展開
        template_row = self._find_template_row(ws)
        if template_row is None:
            return

        self._expand_rows(ws, template_row)

        # ヘッダーのプレースホルダー（年度・学校名等）を置換
        first_row = self.data.iloc[0].to_dict() if len(self.data) > 0 else {}
        fill_placeholders(ws, first_row, self.options)

    def _get_template_meta(self) -> dict:
        for _name, meta in TEMPLATES.items():
            if meta['file'] == os.path.basename(self.template_path):
                return meta
        return {}

    def _find_template_row(self, ws) -> int | None:
        """{{氏名}} または {{表示氏名}} を含む最初の行番号を返す。"""
        for row in ws.iter_rows():
            for cell in row:
                if (cell.value and isinstance(cell.value, str)
                        and ('{{氏名}}' in cell.value or '{{表示氏名}}' in cell.value)):
                    return cell.row
        return None

    def _expand_rows(self, ws, template_row: int) -> None:
        """テンプレート行を児童数分コピーして展開する。"""
        from copy import copy as cp


        n = len(self.data)
        if n == 0:
            return

        # テンプレート行のセル情報を保存
        tmpl_cells = []
        for cell in ws[template_row]:
            tmpl_cells.append({
                'value': cell.value,
                'font': cp(cell.font),
                'alignment': cp(cell.alignment),
                'border': cp(cell.border),
                'fill': cp(cell.fill),
                'number_format': cell.number_format,
            })

        row_height = ws.row_dimensions[template_row].height

        # 既存行を 1 行だけ利用し、残りを挿入
        for i, (_, data_row) in enumerate(self.data.iterrows()):
            row_num = template_row + i
            if i > 0:
                ws.insert_rows(row_num)

            for col_idx, tmpl in enumerate(tmpl_cells, 1):
                cell = ws.cell(row=row_num, column=col_idx)
                cell.value = tmpl['value']
                cell.font = tmpl['font']
                cell.alignment = tmpl['alignment']
                cell.border = tmpl['border']
                cell.fill = tmpl['fill']
                cell.number_format = tmpl['number_format']

            ws.row_dimensions[row_num].height = row_height

            row_dict = data_row.to_dict()
            fill_placeholders(ws, row_dict, self.options)


# ────────────────────────────────────────────────────────────────────────────
# IndividualGenerator — 個票系（1 名/シートで複製）
# ────────────────────────────────────────────────────────────────────────────

class IndividualGenerator(BaseGenerator):
    """
    テンプレートの 1 枚目シートを児童数分 copy_worksheet で複製する。
    画像は copy_sheet_with_images で再挿入する。
    """

    def _populate(self) -> None:
        template_ws = self.wb.active
        meta = self._get_template_meta()
        orientation = meta.get('orientation', 'portrait')

        for i, (_, row) in enumerate(self.data.iterrows()):
            row_dict = row.to_dict()
            shusseki = row_dict.get('出席番号', str(i + 1))
            name = row_dict.get('氏名', row_dict.get('正式氏名', str(i + 1)))
            title = f'{str(shusseki).zfill(2)}_{name}'[:31]  # シート名 31 文字制限

            if i == 0:
                ws = template_ws
                ws.title = title
            else:
                ws = copy_sheet_with_images(self.wb, template_ws, title)

            fill_placeholders(ws, row_dict, self.options)
            setup_print(ws, orientation=orientation)

    def _get_template_meta(self) -> dict:
        for _name, meta in TEMPLATES.items():
            if meta['file'] == os.path.basename(self.template_path):
                return meta
        return {}


# ────────────────────────────────────────────────────────────────────────────
# ファクトリー関数
# ────────────────────────────────────────────────────────────────────────────

def create_generator(
    template_name: str,
    output_path: str,
    data: pd.DataFrame,
    options: dict,
) -> BaseGenerator:
    """
    テンプレート名から適切なジェネレーターを返すファクトリー関数。

    options に必要なキー:
        template_dir  : テンプレートフォルダの絶対パス
        fiscal_year   : 年度（整数）
        school_name   : 学校名
        teacher_name  : 担任名（省略可）
    """
    if template_name not in TEMPLATES:
        raise ValueError(f'未定義のテンプレート: {template_name}')

    meta = TEMPLATES[template_name]
    template_path = os.path.join(options['template_dir'], meta['file'])

    if not os.path.exists(template_path):
        raise FileNotFoundError(f'テンプレートが見つかりません: {template_path}')

    gen_type = meta['type']
    if gen_type == 'grid':
        return GridGenerator(template_path, output_path, data, options)
    if gen_type == 'list':
        return ListGenerator(template_path, output_path, data, options)
    if gen_type == 'individual':
        return IndividualGenerator(template_path, output_path, data, options)
    raise ValueError(f'不明なジェネレータータイプ: {gen_type}')
