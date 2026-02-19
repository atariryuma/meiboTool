"""特別支援学級のデータ判定・統合ロジック

C4th エクスポートでは「組」列に特別支援学級名（例: 「なかよし」「ひまわり」）が
入っている。数字の組名（半角/全角）は通常学級、それ以外を特別支援学級と判定する。

交流学級情報は C4th エクスポートに含まれないため、ユーザーが手動で割り当てる。
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd


def is_special_needs_class(class_name: str) -> bool:
    """「組」の値が特別支援学級かどうかを判定する。

    通常学級: 半角/全角の数字のみ（例: '1', '２', '10'）
    特別支援学級: それ以外（例: 'なかよし', 'ひまわり', 'A組'）
    """
    if not class_name or not isinstance(class_name, str):
        return False
    # 全角数字を半角に正規化
    normalized = unicodedata.normalize('NFKC', class_name.strip())
    return not bool(re.fullmatch(r'\d+', normalized))


def detect_special_needs_students(df: pd.DataFrame) -> pd.DataFrame:
    """DataFrame から特別支援学級在籍の児童を抽出する。"""
    if '組' not in df.columns:
        return pd.DataFrame(columns=df.columns)
    mask = df['組'].apply(is_special_needs_class)
    return df[mask].copy()


def detect_regular_students(df: pd.DataFrame) -> pd.DataFrame:
    """DataFrame から通常学級在籍の児童を抽出する。"""
    if '組' not in df.columns:
        return df.copy()
    mask = df['組'].apply(lambda v: not is_special_needs_class(v))
    return df[mask].copy()


def get_special_needs_classes(df: pd.DataFrame) -> list[str]:
    """DataFrame 内の特別支援学級名を返す（ソート済み）。"""
    if '組' not in df.columns:
        return []
    all_classes = df['組'].unique()
    return sorted([c for c in all_classes if is_special_needs_class(str(c))])


def merge_special_needs_students(
    regular_df: pd.DataFrame,
    special_df: pd.DataFrame,
    placement: str = 'appended',
) -> pd.DataFrame:
    """通常学級の児童に特別支援学級の児童を統合する。

    Args:
        regular_df: 通常学級の児童（出席番号ソート済み）
        special_df: 特別支援学級の児童（交流学級でフィルタ済み）
        placement: 'integrated' = 出席番号順に統合, 'appended' = 末尾に追加

    Returns:
        統合された DataFrame
    """
    if special_df.empty:
        return regular_df.copy()

    if placement == 'integrated':
        # 出席番号順に統合（特支児童の出席番号で挿入位置を決定）
        combined = pd.concat([regular_df, special_df], ignore_index=True)
        if '出席番号' in combined.columns:
            combined['_sort_key'] = pd.to_numeric(
                combined['出席番号'], errors='coerce',
            ).fillna(999)
            combined = combined.sort_values('_sort_key').drop(columns='_sort_key')
            combined = combined.reset_index(drop=True)
        return combined

    # 末尾に追加（デフォルト）
    return pd.concat([regular_df, special_df], ignore_index=True)
