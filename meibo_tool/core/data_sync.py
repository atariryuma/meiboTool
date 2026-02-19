"""名簿データの自動同期

3 つのモード:
  - manual:  従来の手動ファイル選択（何もしない）
  - lan:     UNC パスからの自動読込
  - gdrive:  Google Drive からの暗号化ファイルダウンロード + 復号

いずれもバックグラウンドスレッドから呼び出し、
結果を SyncResult で返す。GUI はメインスレッドで処理する。
"""

from __future__ import annotations

import contextlib
import logging
import os
import shutil
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import requests

from core.config import get_cache_dir
from core.crypto import (
    DecryptionError,
    compute_file_hash,
    decrypt_file,
    unprotect_password,
)

logger = logging.getLogger(__name__)

# Google Drive 直接ダウンロード URL テンプレート
_GDRIVE_URL = 'https://drive.google.com/uc?export=download&id={file_id}'
_REQUEST_TIMEOUT = 15  # seconds


@dataclass
class SyncResult:
    """同期結果。"""
    status: str     # 'unchanged' | 'updated' | 'unavailable' | 'decrypt_error' | 'error'
    path: str = ''  # 同期されたファイルのパス（updated の場合）
    message: str = ''  # エラーメッセージ
    config_updates: dict | None = None  # メインスレッドで config に適用する変更


def sync(config: dict[str, Any]) -> SyncResult:
    """config['data_source']['mode'] に応じてデータ同期を実行する。"""
    ds = config.get('data_source', {})
    mode = ds.get('mode', 'manual')

    if mode == 'manual':
        return SyncResult(status='unchanged')
    if mode == 'lan':
        return _sync_lan(config)
    if mode == 'gdrive':
        return _sync_gdrive(config)
    return SyncResult(status='error', message=f'不明なモード: {mode}')


def _sync_lan(config: dict[str, Any]) -> SyncResult:
    """LAN モード: 共有フォルダからの同期。"""
    ds = config.get('data_source', {})
    lan_path = ds.get('lan_path', '')

    if not lan_path:
        return SyncResult(status='error', message='LAN パスが設定されていません')

    if not os.path.exists(lan_path):
        # パスが存在しない → キャッシュがあればそれを使う
        cache_file = ds.get('cache_file', '')
        if cache_file and os.path.exists(cache_file):
            return SyncResult(
                status='unavailable',
                path=cache_file,
                message='ネットワークフォルダに接続できません。前回のデータを使用します。',
            )
        return SyncResult(
            status='unavailable',
            message='ネットワークフォルダに接続できません。',
        )

    # ハッシュ比較
    try:
        current_hash = compute_file_hash(lan_path)
    except OSError as exc:
        return SyncResult(status='error', message=f'ファイル読込エラー: {exc}')

    last_hash = ds.get('last_sync_hash', '')
    if current_hash == last_hash:
        cache_file = ds.get('cache_file', '')
        if cache_file and os.path.exists(cache_file):
            return SyncResult(status='unchanged', path=cache_file)
        return SyncResult(status='unchanged')

    # 変更あり → キャッシュにコピー
    cache_dir = get_cache_dir()
    cache_path = os.path.join(cache_dir, 'roster_cache.xlsx')
    try:
        shutil.copy2(lan_path, cache_path)
    except OSError as exc:
        return SyncResult(status='error', message=f'キャッシュ保存エラー: {exc}')

    # config 更新はメインスレッドに委譲する
    updates = {
        'last_sync_hash': current_hash,
        'last_sync_time': datetime.now().isoformat(timespec='seconds'),
        'cache_file': cache_path,
    }

    return SyncResult(status='updated', path=cache_path, config_updates=updates)


def _sync_gdrive(config: dict[str, Any]) -> SyncResult:
    """Google Drive モード: 暗号化ファイルのダウンロード + 復号。"""
    ds = config.get('data_source', {})
    file_id = ds.get('gdrive_file_id', '')
    protected_pw = ds.get('encryption_password', '')

    if not file_id:
        return SyncResult(status='error', message='Google Drive ファイル ID が設定されていません')

    password = unprotect_password(protected_pw)
    if not password:
        return SyncResult(status='error', message='暗号化パスワードが設定されていません')

    cache_dir = get_cache_dir()
    encrypted_path = os.path.join(cache_dir, 'roster_encrypted.bin')

    # ダウンロード
    try:
        url = _GDRIVE_URL.format(file_id=file_id)
        resp = requests.get(url, timeout=_REQUEST_TIMEOUT, stream=True)
        resp.raise_for_status()

        with open(encrypted_path, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)

    except requests.RequestException as exc:
        logger.warning('Google Drive ダウンロード失敗: %s', exc)
        cache_file = ds.get('cache_file', '')
        if cache_file and os.path.exists(cache_file):
            return SyncResult(
                status='unavailable',
                path=cache_file,
                message='Google Drive に接続できません。前回のデータを使用します。',
            )
        return SyncResult(
            status='unavailable',
            message='Google Drive に接続できません。',
        )

    # ハッシュ比較
    try:
        current_hash = compute_file_hash(encrypted_path)
    except OSError as exc:
        return SyncResult(status='error', message=f'ハッシュ計算エラー: {exc}')

    last_hash = ds.get('last_sync_hash', '')
    if current_hash == last_hash:
        cache_file = ds.get('cache_file', '')
        if cache_file and os.path.exists(cache_file):
            return SyncResult(status='unchanged', path=cache_file)
        # ハッシュは同じだがキャッシュがない場合は復号する
        pass  # fall through to decryption

    # 復号
    cache_path = os.path.join(cache_dir, 'roster_cache.xlsx')
    try:
        decrypt_file(encrypted_path, cache_path, password)
    except DecryptionError as exc:
        return SyncResult(status='decrypt_error', message=str(exc))

    # 暗号化ファイルを削除（復号済みのため不要）
    with contextlib.suppress(OSError):
        os.remove(encrypted_path)

    # config 更新はメインスレッドに委譲する
    updates = {
        'last_sync_hash': current_hash,
        'last_sync_time': datetime.now().isoformat(timespec='seconds'),
        'cache_file': cache_path,
    }

    return SyncResult(status='updated', path=cache_path, config_updates=updates)
