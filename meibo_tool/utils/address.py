"""住所フィールド結合ユーティリティ"""

import pandas as pd


def _join_fields(row: dict | pd.Series, fields: list[str]) -> str:
    """指定フィールドを結合して住所文字列を返す。NaN・'nan'・空文字は除外。"""
    parts = []
    for f in fields:
        val = row.get(f, '')
        if val is None:
            continue
        s = str(val).strip()
        if s and s.lower() != 'nan':
            parts.append(s)
    return ''.join(parts)


def build_address(row: dict | pd.Series) -> str:
    """
    都道府県・市区町村・町番地・建物名を結合して住所文字列を返す。
    NaN・'nan'・空文字は除外する。

    Examples:
        >>> build_address({'都道府県': '沖縄県', '市区町村': '那覇市', '町番地': '天久1-2-3', '建物名': ''})
        '沖縄県那覇市天久1-2-3'
    """
    return _join_fields(row, ['都道府県', '市区町村', '町番地', '建物名'])


def build_guardian_address(row: dict | pd.Series) -> str:
    """保護者住所を結合して返す。児童住所と同一なら「同上」を返す。

    スズキ校務データ: '保護者住所' が単一フィールドで存在 → そのまま返す。
    C4th データ: 保護者都道府県/市区町村/町番地/建物名 を結合する。
    """
    # スズキ校務: 単一フィールドがあればそのまま使う
    single = row.get('保護者住所', '')
    if single is not None:
        s = str(single).strip()
        if s and s.lower() != 'nan':
            return s

    # C4th: 4 フィールドを結合
    guardian = _join_fields(
        row, ['保護者都道府県', '保護者市区町村', '保護者町番地', '保護者建物名'],
    )
    if not guardian:
        return ''

    # 児童住所と同一なら「同上」
    student = build_address(row)
    if guardian == student:
        return '同上'
    return guardian
