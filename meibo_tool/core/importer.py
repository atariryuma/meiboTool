"""C4th エクスポート Excel の読み込み

SPEC.md 第2章 §2.4 参照。
ヘッダー行自動検出 + 全列 dtype=str で読み込む。
"""

import pandas as pd
from openpyxl import load_workbook

from core.mapper import map_columns


def detect_header_row(filepath: str, max_scan: int = 10) -> int:
    """
    ヘッダー行を自動検出する。
    判定基準: 文字列セルが 5 つ以上連続する最初の行（1-indexed）。
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    result = 1  # フォールバック
    for row_idx, row in enumerate(ws.iter_rows(max_row=max_scan), 1):
        str_count = sum(
            1 for cell in row
            if cell.value is not None and isinstance(cell.value, str)
        )
        if str_count >= 5:
            result = row_idx
            break
    wb.close()
    return result


def import_c4th_excel(filepath: str) -> tuple[pd.DataFrame, list[str]]:
    """
    C4th エクスポート Excel を読み込み、マッピング済み DataFrame を返す。

    Returns:
        df_mapped: 内部論理名にリネーム済みの DataFrame
        unmapped:  マッピングできなかったカラム名リスト
    """
    header_row = detect_header_row(filepath)
    df = pd.read_excel(
        filepath,
        header=header_row - 1,  # 0-indexed
        dtype=str,               # 全列文字列（型変換は後工程）
    )
    # 空白カラム名・全空白行を除去
    df = df.loc[:, df.columns.notna()]
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)

    df_mapped, unmapped = map_columns(df)
    return df_mapped, unmapped
