"""テンプレートメタデータレジストリ

SPEC.md 第4章 §4.0.1 参照。
各テンプレートの差込タイプ・ファイル名・必須カラム等を一元管理する。

TEMPLATES dict はハードコードされたオーバーライド定義。
テンプレートフォルダに .xlsx を置くだけで自動検出される仕組みも併用する。
`get_all_templates(template_dir)` がスキャン結果と TEMPLATES をマージして返す。

`enabled`: False のテンプレートは UI に表示しない（Phase 3 未実装分）。
"""

from __future__ import annotations

import os
from typing import Any

# カテゴリ表示順序
CATEGORY_ORDER: list[str] = [
    '名札・ラベル',
    '名簿・出欠表',
    'その他',
]

TEMPLATES: dict[str, dict] = {
    # ── Grid 型（名札・ラベル各種）────────────────────────────────────────
    'ラベル_色付き': {
        'file': 'ラベル_色付き.xlsx',
        'type': 'grid',
        'category': '名札・ラベル',
        'cards_per_page': 40,
        'orientation': 'portrait',
        'use_formal_name': False,
        'required_columns': ['氏名'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '🌸',
        'description': '装飾ラベル（カラー、40枚/ページ）',
    },
    'ラベル_大2': {
        'file': 'ラベル_大2.xlsx',
        'type': 'grid',
        'category': '名札・ラベル',
        'cards_per_page': 40,
        'orientation': 'portrait',
        'use_formal_name': False,
        'required_columns': ['氏名'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '🏷',
        'description': '大ラベル（黒・2列40枚/ページ）',
    },
    'ラベル_小': {
        'file': 'ラベル_小.xlsx',
        'type': 'grid',
        'category': '名札・ラベル',
        'cards_per_page': 40,
        'orientation': 'portrait',
        'use_formal_name': False,
        'required_columns': ['氏名'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '🏷',
        'description': '小ラベル（4列×20行、40枚/ページ×2部）',
    },
    'ラベル_特大': {
        'file': 'ラベル_特大.xlsx',
        'type': 'grid',
        'category': '名札・ラベル',
        'cards_per_page': 40,
        'orientation': 'portrait',
        'use_formal_name': False,
        'required_columns': ['氏名'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '🏷',
        'description': '特大ラベル（72pt、40枚/ページ）',
    },
    '名札_1年生用': {
        'file': '名札_1年生用.xlsx',
        'type': 'grid',
        'category': '名札・ラベル',
        'cards_per_page': 8,
        'grid_cols': 8,
        'grid_rows': 1,
        'use_formal_name': False,
        'required_columns': ['氏名かな'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '🎒',
        'description': '縦長短冊型（ふりがな縦書き）8枚/ページ',
    },
    'サンプル_名札': {
        'file': 'サンプル_名札.xlsx',
        'type': 'grid',
        'category': '名札・ラベル',
        'cards_per_page': 10,
        'orientation': 'landscape',
        'use_formal_name': False,
        'required_columns': ['氏名'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '🏷',
        'description': 'サンプル名札（A4横・10枚/ページ）',
    },

    # ── Grid 型（出欠帳票）───────────────────────────────────────────────
    '横名簿': {
        'file': '横名簿.xlsx',
        'type': 'grid',
        'category': '名簿・出欠表',
        'cards_per_page': 40,
        'orientation': 'landscape',
        'use_formal_name': False,
        'required_columns': ['氏名'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '📋',
        'description': '横型名簿（A4横・2列40名・出欠欄10日分）',
    },
    '縦一週間': {
        'file': '縦一週間.xlsx',
        'type': 'grid',
        'category': '名簿・出欠表',
        'orientation': 'portrait',
        'use_formal_name': False,
        'required_columns': ['氏名'],
        'mandatory_columns': ['組', '出席番号'],
        'icon': '📋',
        'description': '縦型週間出欠表（A4縦・40名固定）',
    },
    # ── Grid 型（性別一覧）─────────────────────────────────────────────────
    '男女一覧': {
        'file': '男女一覧.xlsx',
        'type': 'grid',
        'category': '名簿・出欠表',
        'orientation': 'landscape',
        'use_formal_name': False,
        'required_columns': ['氏名', '性別'],
        'mandatory_columns': ['組', '出席番号'],
        'sort_by': '性別',
        'sort_order': {'男': 0, '女': 1},
        'icon': '📋',
        'description': '男女別一覧表（A4横・性別ソート）',
    },

}

# ── TEMPLATES のファイル名 → キー名の逆引きインデックス ─────────────────────
_FILE_TO_KEY: dict[str, str] = {
    meta['file']: name for name, meta in TEMPLATES.items()
}


# ────────────────────────────────────────────────────────────────────────────
# テンプレート統合（スキャン + ハードコード）
# ────────────────────────────────────────────────────────────────────────────

def get_all_templates(template_dir: str) -> dict[str, dict[str, Any]]:
    """テンプレートフォルダのスキャン結果と TEMPLATES をマージして返す。

    マージルール:
      - TEMPLATES にあるエントリ: TEMPLATES の値を優先（アイコン・説明等を維持）
        ただしファイルがフォルダに存在しない場合は除外
      - フォルダにのみ存在する .xlsx: スキャン結果をそのまま使用
      - enabled=False のエントリは維持される（UIで非表示にするため）
    """
    from templates.template_scanner import scan_template_folder

    scanned = scan_template_folder(template_dir)
    result: dict[str, dict[str, Any]] = {}

    # 1. TEMPLATES のエントリを処理
    for name, meta in TEMPLATES.items():
        filename = meta['file']
        filepath = os.path.join(template_dir, filename) if template_dir else ''
        file_key = os.path.splitext(filename)[0]

        if meta.get('enabled') is False:
            # enabled=False は常に含める（UIで除外するため）
            result[name] = dict(meta)
        elif os.path.isfile(filepath):
            # ファイルが存在 → TEMPLATES の定義を使用
            result[name] = dict(meta)
            # スキャン結果から除外（重複防止）
            scanned.pop(file_key, None)
            # TEMPLATES のキー名とファイル名ベースのキーが異なる場合も除外
            scanned.pop(name, None)
        # ファイルが存在しない enabled テンプレートは除外

    # 2. スキャンのみで見つかったテンプレートを追加
    for key, meta in scanned.items():
        # TEMPLATES にファイル名が既に登録されている場合はスキップ
        if meta['file'] in _FILE_TO_KEY:
            continue
        result[key] = meta

    return result


# ────────────────────────────────────────────────────────────────────────────
# GUI 用ヘルパー関数
# ────────────────────────────────────────────────────────────────────────────

def get_display_groups(
    template_dir: str = '',
) -> list[tuple[str, list[tuple[str, str, str, str]]]]:
    """
    GUI テンプレート選択リスト用の **カテゴリ別** 一覧を返す。
    `enabled=False` のテンプレートは除外する。

    template_dir が指定された場合はフォルダスキャンも含めた全テンプレートを返す。
    空文字の場合は TEMPLATES のみを使用する（後方互換）。

    戻り値: [(category_name, [(display_name, key, icon, description), ...]), ...]
    カテゴリ順序は CATEGORY_ORDER に従う。未分類は「その他」に入る。
    """
    templates = get_all_templates(template_dir) if template_dir else TEMPLATES

    # カテゴリごとにグルーピング
    groups: dict[str, list[tuple[str, str, str, str]]] = {}
    for name, meta in templates.items():
        if not meta.get('enabled', True):
            continue
        cat = meta.get('category', 'その他')
        groups.setdefault(cat, []).append(
            (name, name, meta.get('icon', '📋'), meta.get('description', name))
        )

    # CATEGORY_ORDER 順でソート、未知のカテゴリは末尾
    order = {cat: i for i, cat in enumerate(CATEGORY_ORDER)}
    return sorted(groups.items(), key=lambda x: order.get(x[0], 999))
