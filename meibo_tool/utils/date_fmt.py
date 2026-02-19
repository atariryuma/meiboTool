"""日付文字列のフォーマットユーティリティ"""

import re

# 日付型フィールド（YY/MM/DD にフォーマットする）
DATE_KEYS: frozenset[str] = frozenset({'生年月日', '入学日', '転入日'})


def format_date(s: str) -> str:
    """日付文字列を YY/MM/DD 形式に変換する。変換不能ならそのまま返す。

    対応形式:
        - "2018-06-15" / "2018/06/15" / "2018-06-15 00:00:00"
        - Excel シリアル値 ("43266.0")
    """
    m = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f'{y % 100:02d}/{mo:02d}/{d:02d}'
    # Excel serial number (float string like "43266.0")
    try:
        serial = float(s)
        if 1 < serial < 100000:
            from datetime import datetime, timedelta
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=int(serial))
            return dt.strftime('%y/%m/%d')
    except ValueError:
        pass
    return s
