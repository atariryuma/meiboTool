"""住所フィールド結合ユーティリティ"""

import pandas as pd


def build_address(row: dict | pd.Series) -> str:
    """
    都道府県・市区町村・町番地・建物名を結合して住所文字列を返す。
    NaN・'nan'・空文字は除外する。

    Examples:
        >>> build_address({'都道府県': '沖縄県', '市区町村': '那覇市', '町番地': '天久1-2-3', '建物名': ''})
        '沖縄県那覇市天久1-2-3'
    """
    fields = ['都道府県', '市区町村', '町番地', '建物名']
    parts = []
    for f in fields:
        val = row.get(f, '')
        if val is None:
            continue
        s = str(val).strip()
        if s and s.lower() != 'nan':
            parts.append(s)
    return ''.join(parts)


def build_guardian_address(row: dict | pd.Series) -> str:
    """
    保護者の住所を結合して返す。
    全フィールドが空の場合は空文字を返す（同上と判断させるのは呼び出し元）。
    """
    fields = ['保護者都道府県', '保護者市区町村', '保護者町番地', '保護者建物名']
    parts = []
    for f in fields:
        val = row.get(f, '')
        if val is None:
            continue
        s = str(val).strip()
        if s and s.lower() != 'nan':
            parts.append(s)
    return ''.join(parts)
