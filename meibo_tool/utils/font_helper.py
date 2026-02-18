"""IPAmj明朝フォント設定ヘルパー

openpyxl のスタイルはイミュータブルなので、
必ず Font(...) で新インスタンスを作ること。
cell.font.bold = True は AttributeError になる。
"""

from openpyxl.styles import Font
from openpyxl.cell.cell import MergedCell


DEFAULT_FONT = 'IPAmj明朝'


def apply_font(ws, font_name: str = DEFAULT_FONT) -> None:
    """
    ワークシート内の全セルのフォント名を font_name に変更する。
    既存のサイズ・太字・斜体・色は保持する。
    MergedCell（結合セルの左上以外）はスキップする。
    """
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value is None:
                continue
            cur = cell.font or Font()
            cell.font = Font(
                name=font_name,
                size=cur.size,
                bold=cur.bold,
                italic=cur.italic,
                color=cur.color,
                underline=cur.underline,
                strike=cur.strike,
            )
