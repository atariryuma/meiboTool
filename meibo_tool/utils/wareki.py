"""西暦 → 和暦変換ユーティリティ"""

# 元号テーブル（開始日順、新しい順に並べる）
_GENGO = [
    ('令和', 2019, 5, 1),   # 2019-05-01 〜
    ('平成', 1989, 1, 8),   # 1989-01-08 〜 2019-04-30
    ('昭和', 1926, 12, 25), # 1926-12-25 〜 1989-01-07
    ('大正', 1912, 7, 30),  # 1912-07-30 〜 1926-12-24
    ('明治', 1868, 1, 25),  # 1868-01-25 〜 1912-07-29
]


def to_wareki(year: int, month: int = 1, day: int = 1) -> str:
    """
    西暦年（と任意の月日）を和暦文字列に変換する。

    Examples:
        >>> to_wareki(2025)
        '令和7年'
        >>> to_wareki(2019, 5, 1)
        '令和1年'
        >>> to_wareki(2019, 4, 30)
        '平成31年'
        >>> to_wareki(1989, 1, 8)
        '平成1年'
    """
    for gengo, g_year, g_month, g_day in _GENGO:
        start = (g_year, g_month, g_day)
        if (year, month, day) >= start:
            wareki_year = year - g_year + 1
            return f'{gengo}{wareki_year}年'
    return f'西暦{year}年'


def to_wareki_full(year: int, month: int = 1, day: int = 1) -> str:
    """
    和暦を「令和7年4月1日」形式で返す。

    Examples:
        >>> to_wareki_full(2025, 4, 1)
        '令和7年4月1日'
    """
    base = to_wareki(year, month, day)
    return f'{base}{month}月{day}日'


def fiscal_year_to_wareki(fiscal_year: int) -> str:
    """
    年度（4月1日起算）の和暦を返す。

    Examples:
        >>> fiscal_year_to_wareki(2025)
        '令和7年度'
    """
    label = to_wareki(fiscal_year, 4, 1)
    return label.replace('年', '年度')
