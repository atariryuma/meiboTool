"""C4th カラム名 → 内部論理名マッピング

SPEC.md 第2章 §2.2〜§2.3 参照。
全角スペース（U+3000）は normalize_header() で統一する。
"""

# C4th 確定ヘッダー → 内部論理名（完全一致マップ）
EXACT_MAP: dict[str, str] = {
    '生徒コード': '生徒コード',
    '学年': '学年',
    '名前': '氏名',
    'ふりがな': '氏名かな',
    '正式名前': '正式氏名',
    '正式名前ふりがな': '正式氏名かな',
    '性別': '性別',
    '生年月日': '生年月日',
    '外国籍': '外国籍',
    '郵便番号': '郵便番号',
    '都道府県': '都道府県',
    '市区町村': '市区町村',
    '町番地': '町番地',
    'アパート/マンション名': '建物名',
    '電話番号1': '電話番号1',
    '電話番号2': '電話番号2',
    '電話番号3': '電話番号3',
    'FAX番号': 'FAX番号',
    '出身校': '出身校',
    '入学日': '入学日',
    '転入日': '転入日',
    '保護者1\u3000続柄': '保護者続柄',       # 全角スペース区切り
    '保護者1\u3000名前': '保護者名',
    '保護者1\u3000名前ふりがな': '保護者名かな',
    '保護者1\u3000正式名前': '保護者正式名',
    '保護者1\u3000正式名前ふりがな': '保護者正式名かな',
    '保護者1\u3000郵便番号': '保護者郵便番号',
    '保護者1\u3000都道府県': '保護者都道府県',
    '保護者1\u3000市区町村': '保護者市区町村',
    '保護者1\u3000町番地': '保護者町番地',
    '保護者1\u3000アパート/マンション名': '保護者建物名',
    '保護者1\u3000電話番号1': '保護者電話1',
    '保護者1\u3000電話番号2': '保護者電話2',
    '保護者1\u3000電話番号3': '保護者電話3',
    '保護者1\u3000FAX番号': '保護者FAX',
    '保護者1\u3000緊急連絡先': '緊急連絡先',
}

# 表記ゆれ対応エイリアス（他校・旧バージョン）
COLUMN_ALIASES: dict[str, list[str]] = {
    '氏名': ['名前', '児童氏名', '生徒氏名', '児童名'],
    '氏名かな': ['ふりがな', 'フリガナ', 'かな', 'カナ'],
    '正式氏名': ['正式名前', '正式氏名'],
    '出席番号': ['番号', '出席番号', 'No', 'NO', '席番'],
    '組': ['組', '学級', 'クラス'],
}


def normalize_header(s: str) -> str:
    """ヘッダー名を正規化する（前後空白除去・全角スペース統一）。"""
    if not isinstance(s, str):
        return str(s)
    s = s.strip()
    # 全角スペース（U+3000）で統一（念のため両種を変換）
    s = s.replace('　', '\u3000')
    return s


def map_columns(df):
    """
    DataFrame のカラム名を内部論理名にマッピングする。

    Returns:
        df_mapped: リネーム済み DataFrame
        unmapped:  マッピングできなかった元のカラム名リスト
    """

    renamed: dict[str, str] = {}
    unmapped: list[str] = []

    for col in df.columns:
        norm = normalize_header(col)
        if norm in EXACT_MAP:
            renamed[col] = EXACT_MAP[norm]
        else:
            # エイリアス検索
            found = False
            for logical, aliases in COLUMN_ALIASES.items():
                if norm in aliases or col in aliases:
                    renamed[col] = logical
                    found = True
                    break
            if not found:
                unmapped.append(col)

    df_mapped = df.rename(columns=renamed)
    return df_mapped, unmapped


# 正式名前系 → 通常名前系へのフォールバックマッピング
_FALLBACK_COLUMNS: list[tuple[str, str]] = [
    ('正式氏名', '氏名'),
    ('正式氏名かな', '氏名かな'),
    ('保護者正式名', '保護者名'),
    ('保護者正式名かな', '保護者名かな'),
]


def ensure_fallback_columns(df) -> None:
    """不足している論理列をフォールバック元からコピーして補完する（in-place）。

    例: '正式氏名' がなく '氏名' がある場合、'氏名' の値を '正式氏名' にコピー。
    逆方向（'氏名' がなく '正式氏名' がある場合）も補完する。
    """
    for formal, casual in _FALLBACK_COLUMNS:
        if formal not in df.columns and casual in df.columns:
            df[formal] = df[casual]
        elif casual not in df.columns and formal in df.columns:
            df[casual] = df[formal]


def resolve_name_fields(data_row: dict, use_formal: bool) -> dict[str, str]:
    """
    use_formal_name フラグに基づき表示用氏名フィールドを選択する。
    正式氏名が空の場合は通常氏名にフォールバックする。
    """
    def _val(key: str) -> str:
        v = data_row.get(key, '')
        return '' if (v is None or str(v).strip().lower() == 'nan') else str(v).strip()

    if use_formal:
        return {
            '表示氏名': _val('正式氏名') or _val('氏名'),
            '表示氏名かな': _val('正式氏名かな') or _val('氏名かな'),
            '表示保護者名': _val('保護者正式名') or _val('保護者名'),
            '表示保護者名かな': _val('保護者正式名かな') or _val('保護者名かな'),
        }
    return {
        '表示氏名': _val('氏名'),
        '表示氏名かな': _val('氏名かな'),
        '表示保護者名': _val('保護者名'),
        '表示保護者名かな': _val('保護者名かな'),
    }
