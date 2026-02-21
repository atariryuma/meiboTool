"""名簿データの CSV/Excel エクスポート"""

from __future__ import annotations

import pandas as pd


def export_csv(
    df: pd.DataFrame, filepath: str, *, encoding: str = 'utf-8-sig',
) -> None:
    """DataFrame を CSV ファイルに書き出す。

    デフォルトは UTF-8 with BOM（Excel で開いた時に文字化けしない）。
    """
    df.to_csv(filepath, index=False, encoding=encoding)


def export_excel(df: pd.DataFrame, filepath: str) -> None:
    """DataFrame を Excel ファイルに書き出す。"""
    df.to_excel(filepath, index=False, engine='openpyxl')
