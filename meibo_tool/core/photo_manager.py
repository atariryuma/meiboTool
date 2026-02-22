"""写真管理モジュール

生徒写真のスキャン・マッチング・読込を管理する。
写真は photo_dir フォルダに保存され、ファイル名規約で自動マッチされる。
"""

from __future__ import annotations

import contextlib
import io
import logging
import os
import re
import shutil

import pandas as pd
from PIL import Image, ImageOps

logger = logging.getLogger(__name__)

_SUPPORTED_EXTENSIONS = frozenset({
    '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp',
})

# ファイル名から番号キーを抽出する正規表現
# 例: '1-1-01' or '01' or '1'
_NUMBER_KEY_PATTERN = re.compile(r'^(\d+-\d+-\d+|\d+)$')


def scan_photos(photo_dir: str) -> dict[str, str]:
    """フォルダをスキャンして {マッチキー: 絶対パス} の辞書を返す。

    1ファイルにつき複数キーを登録する（番号系 + 名前系）。
    サブフォルダにも再帰的に対応する。

    Args:
        photo_dir: 写真フォルダのパス

    Returns:
        マッチキー → 絶対パス の辞書
    """
    result: dict[str, str] = {}
    if not os.path.isdir(photo_dir):
        return result

    for dirpath, _dirs, filenames in os.walk(photo_dir):
        for fname in filenames:
            ext = os.path.splitext(fname)[1].lower()
            if ext not in _SUPPORTED_EXTENSIONS:
                continue

            abs_path = os.path.join(dirpath, fname)
            stem = os.path.splitext(fname)[0]

            keys = _generate_match_keys(stem)
            for key in keys:
                if key in result:
                    logger.warning(
                        '写真キー重複: %s → 既存=%s, 新規=%s (スキップ)',
                        key, result[key], abs_path,
                    )
                    continue
                result[key] = abs_path

    return result


def _generate_match_keys(stem: str) -> list[str]:
    """ファイル名（拡張子なし）からマッチキーのリストを生成する。

    例:
        '1-1-01' → ['1-1-01', '1-1-1']
        '01' → ['01', '1']
        '山田太郎' → ['山田太郎']
    """
    keys: list[str] = []
    stem_stripped = stem.strip()
    if not stem_stripped:
        return keys

    # 常に元のファイル名をキーに追加
    keys.append(stem_stripped)

    # 番号キーの場合、ゼロ埋めなし版も追加
    m = re.match(r'^(\d+)-(\d+)-(\d+)$', stem_stripped)
    if m:
        grade, cls, num = m.groups()
        no_pad = f'{int(grade)}-{int(cls)}-{int(num)}'
        if no_pad != stem_stripped:
            keys.append(no_pad)
        # 出席番号のみのキーも追加（単学級対応）
        keys.append(num)
        num_no_pad = str(int(num))
        if num_no_pad != num:
            keys.append(num_no_pad)
        return keys

    # 純粋な数字の場合（出席番号のみ）
    if stem_stripped.isdigit():
        no_pad = str(int(stem_stripped))
        if no_pad != stem_stripped:
            keys.append(no_pad)
        return keys

    return keys


def match_photo_to_student(
    data_row: dict, photo_map: dict[str, str],
) -> str | None:
    """生徒データ行に対応する写真ファイルパスを返す。

    マッチング優先順位:
        1. {学年}-{組}-{出席番号} (ゼロ埋め2桁)
        2. {学年}-{組}-{出席番号} (ゼロ埋めなし)
        3. {出席番号} のみ (ゼロ埋め2桁)
        4. {出席番号} のみ (ゼロ埋めなし)
        5. {氏名} (空白除去)
        6. {正式氏名} (空白除去)
        7. None

    Args:
        data_row: 生徒データ dict（内部論理名がキー）
        photo_map: scan_photos() の返値

    Returns:
        写真ファイルの絶対パス、またはマッチしない場合 None
    """
    if not photo_map:
        return None

    grade = _normalize_field(data_row.get('学年', ''))
    cls = _normalize_field(data_row.get('組', ''))
    num = _normalize_field(data_row.get('出席番号', ''))

    # 番号ベースのマッチング
    if grade and cls and num:
        candidates = _build_number_candidates(grade, cls, num)
        for key in candidates:
            if key in photo_map:
                return photo_map[key]

    # 出席番号のみのマッチング（単学級向け）
    if num:
        for key in _build_num_only_candidates(num):
            if key in photo_map:
                return photo_map[key]

    # 氏名ベースのマッチング
    for name_field in ('氏名', '正式氏名'):
        name = data_row.get(name_field, '')
        if name:
            name_key = _strip_spaces(str(name))
            if name_key and name_key in photo_map:
                return photo_map[name_key]

    return None


def _normalize_field(value: str | None) -> str:
    """フィールド値を正規化する（NaN・空白除去）。"""
    if value is None:
        return ''
    s = str(value).strip()
    if s.lower() == 'nan':
        return ''
    # 全角数字→半角数字
    s = s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    return s


def _strip_spaces(s: str) -> str:
    """全角・半角スペースを除去する。"""
    return s.replace(' ', '').replace('\u3000', '')


def _build_number_candidates(
    grade: str, cls: str, num: str,
) -> list[str]:
    """番号ベースのマッチ候補キーを生成する。"""
    candidates: list[str] = []
    try:
        g = int(grade)
        c = int(cls)
        n = int(num)
    except ValueError:
        return candidates

    # ゼロ埋め2桁
    candidates.append(f'{g}-{c}-{n:02d}')
    # ゼロ埋めなし
    candidates.append(f'{g}-{c}-{n}')
    return candidates


def _build_num_only_candidates(num: str) -> list[str]:
    """出席番号のみのマッチ候補キーを生成する。"""
    candidates: list[str] = []
    try:
        n = int(num)
    except ValueError:
        return candidates
    candidates.append(f'{n:02d}')
    candidates.append(str(n))
    return candidates


def load_photo_bytes(
    photo_path: str,
    max_size: int = 600,
    target_rect: tuple[int, int] | None = None,
) -> bytes | None:
    """画像ファイルを読込、処理して PNG bytes を返す。

    処理内容:
        1. EXIF Orientation タグに基づく自動回転
        2. RGBA/P → RGB 変換（透過画像対応）
        3. target_rect 指定時: アスペクト比に合わせてセンタークロップ
        4. max_size px にリサイズ（長辺基準）
        5. PNG bytes に変換

    Args:
        photo_path: 画像ファイルパス
        max_size: リサイズの最大長辺 px
        target_rect: (width, height) アスペクト比ターゲット

    Returns:
        PNG bytes、エラー時は None
    """
    try:
        img = Image.open(photo_path)
    except Exception:
        logger.warning('写真読込エラー: %s', photo_path)
        return None

    # EXIF 回転補正（EXIF 情報がない or 壊れている場合は無視）
    with contextlib.suppress(Exception):
        img = ImageOps.exif_transpose(img)

    # RGBA/P → RGB 変換（白背景で合成）
    if img.mode in ('RGBA', 'LA', 'PA'):
        background = Image.new('RGB', img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        img = background
    elif img.mode == 'P' or img.mode != 'RGB':
        img = img.convert('RGB')

    # アスペクト比に合わせてセンタークロップ
    if target_rect is not None:
        tw, th = target_rect
        if tw > 0 and th > 0:
            img = _center_crop(img, tw, th)

    # リサイズ（長辺基準）
    w, h = img.size
    if max(w, h) > max_size:
        if w >= h:
            new_w = max_size
            new_h = max(1, int(h * max_size / w))
        else:
            new_h = max_size
            new_w = max(1, int(w * max_size / h))
        img = img.resize((new_w, new_h), Image.LANCZOS)

    # PNG bytes に変換
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()


def _center_crop(img: Image.Image, target_w: int, target_h: int) -> Image.Image:
    """アスペクト比に合わせてセンタークロップする。"""
    w, h = img.size
    target_ratio = target_w / target_h
    current_ratio = w / h

    if current_ratio > target_ratio:
        # 横長すぎ → 左右をクロップ
        new_w = int(h * target_ratio)
        left = (w - new_w) // 2
        img = img.crop((left, 0, left + new_w, h))
    elif current_ratio < target_ratio:
        # 縦長すぎ → 上下をクロップ
        new_h = int(w / target_ratio)
        top = (h - new_h) // 2
        img = img.crop((0, top, w, top + new_h))

    return img


def get_match_status(
    df: pd.DataFrame, photo_map: dict[str, str],
) -> tuple[int, int, list[str]]:
    """マッチ状況を返す。

    Args:
        df: 生徒データの DataFrame
        photo_map: scan_photos() の返値

    Returns:
        (matched_count, total_count, unmatched_names)
    """
    total = len(df)
    matched = 0
    unmatched_names: list[str] = []

    for _, row in df.iterrows():
        data_row = row.to_dict()
        if match_photo_to_student(data_row, photo_map) is not None:
            matched += 1
        else:
            name = str(data_row.get('氏名', '')).strip()
            if name and name.lower() != 'nan':
                unmatched_names.append(name)
            else:
                unmatched_names.append(f"(行 {_ + 1})")

    return matched, total, unmatched_names


def import_photos_from_folder(
    src_dir: str, photo_dir: str,
) -> tuple[int, list[str]]:
    """ソースフォルダから写真を写真フォルダにコピーする。

    既存ファイルと同名の場合はスキップ（上書きしない）。

    Args:
        src_dir: コピー元フォルダ
        photo_dir: コピー先フォルダ（写真フォルダ）

    Returns:
        (copied_count, skipped_filenames)
    """
    copied = 0
    skipped: list[str] = []

    if not os.path.isdir(src_dir):
        return copied, skipped

    os.makedirs(photo_dir, exist_ok=True)

    for fname in os.listdir(src_dir):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in _SUPPORTED_EXTENSIONS:
            continue

        src_path = os.path.join(src_dir, fname)
        if not os.path.isfile(src_path):
            continue

        dst_path = os.path.join(photo_dir, fname)
        if os.path.exists(dst_path):
            skipped.append(fname)
            continue

        shutil.copy2(src_path, dst_path)
        copied += 1

    return copied, skipped
