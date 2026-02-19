"""GitHub Releases ベースのアプリ自動更新

起動時にバックグラウンドスレッドで check_for_update() を呼び出す。
新バージョンがある場合は UpdateInfo を返し、GUI がダイアログを表示する。
"""

from __future__ import annotations

import contextlib
import logging
import os
import subprocess
import sys
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import requests

from core.config import save_config

logger = logging.getLogger(__name__)

_API_BASE = 'https://api.github.com/repos'
_REQUEST_TIMEOUT = 5  # seconds


# ── データクラス ─────────────────────────────────────────────────────────────

@dataclass
class UpdateInfo:
    """更新情報。"""
    current_version: str
    new_version: str
    release_notes: str
    asset_url: str
    asset_name: str
    asset_size: int
    published_at: str


# ── バージョン比較 ───────────────────────────────────────────────────────────

def _parse_version(v: str) -> tuple[int, ...]:
    """'1.2.3' や 'v1.2.3' → (1, 2, 3)。"""
    return tuple(int(x) for x in v.lstrip('v').split('.'))


def is_newer(remote: str, local: str) -> bool:
    """remote が local より新しいバージョンか判定する。"""
    try:
        return _parse_version(remote) > _parse_version(local)
    except (ValueError, AttributeError):
        return False


# ── 更新チェック ─────────────────────────────────────────────────────────────

def check_for_update(config: dict[str, Any]) -> UpdateInfo | None:
    """GitHub Releases API で最新バージョンを確認する。

    Returns:
        UpdateInfo: 新バージョンがある場合
        None: 最新版 / ネットワークエラー / スキップ設定済み
    """
    update_cfg = config.get('update', {})
    repo = update_cfg.get('github_repo', '')
    if not repo:
        return None

    if not update_cfg.get('check_on_startup', True):
        return None

    current = update_cfg.get('current_app_version', '0.0.0')

    try:
        url = f'{_API_BASE}/{repo}/releases/latest'
        resp = requests.get(
            url,
            headers={'Accept': 'application/vnd.github.v3+json'},
            timeout=_REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
    except requests.RequestException as exc:
        logger.debug('GitHub API リクエスト失敗: %s', exc)
        return None

    data = resp.json()
    tag = data.get('tag_name', '')

    if not is_newer(tag, current):
        return None

    # スキップ設定
    skip = update_cfg.get('skip_version', '')
    if skip and tag == skip:
        return None

    # zip アセットを探す
    assets = data.get('assets', [])
    zip_asset = None
    for asset in assets:
        if asset.get('name', '').endswith('.zip'):
            zip_asset = asset
            break

    if not zip_asset:
        logger.warning('Release %s に zip アセットが見つかりません', tag)
        return None

    # last_check_time 更新
    update_cfg['last_check_time'] = datetime.now().isoformat(timespec='seconds')
    config['update'] = update_cfg
    with contextlib.suppress(OSError):
        save_config(config)

    return UpdateInfo(
        current_version=current,
        new_version=tag,
        release_notes=data.get('body', '') or '',
        asset_url=zip_asset['browser_download_url'],
        asset_name=zip_asset['name'],
        asset_size=zip_asset.get('size', 0),
        published_at=data.get('published_at', ''),
    )


# ── ダウンロード ─────────────────────────────────────────────────────────────

def download_release_asset(
    asset_url: str,
    dest_path: str,
    progress_cb: Callable[[float], None] | None = None,
) -> str:
    """Release asset をダウンロードする。

    Args:
        asset_url: GitHub Release asset の URL
        dest_path: 保存先パス
        progress_cb: 進捗コールバック (0.0〜1.0)

    Returns:
        保存先パス
    """
    resp = requests.get(asset_url, stream=True, timeout=30)
    resp.raise_for_status()

    total = int(resp.headers.get('content-length', 0))
    downloaded = 0

    with open(dest_path, 'wb') as f:
        for chunk in resp.iter_content(chunk_size=65536):
            f.write(chunk)
            downloaded += len(chunk)
            if progress_cb and total > 0:
                progress_cb(downloaded / total)

    return dest_path


# ── バッチファイル生成（--onedir 対応） ──────────────────────────────────────

def generate_update_batch(zip_path: str, current_dir: str) -> None:
    """自己更新バッチを生成して実行し、アプリを終了する。

    手順:
      1. zip を一時フォルダに展開
      2. バッチファイル生成（旧フォルダリネーム → 新フォルダリネーム → config 復元）
      3. バッチ実行
      4. sys.exit(0)

    Args:
        zip_path: ダウンロードした Release zip のパス
        current_dir: 現在のアプリフォルダ（例: dist/名簿帳票ツール）
    """
    parent = os.path.dirname(current_dir)
    folder_name = os.path.basename(current_dir)
    temp_dir = os.path.join(parent, '_update_temp')

    # zip 展開
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(temp_dir)

    # 展開されたフォルダ名を取得（zip 内のトップレベルフォルダ）
    extracted_items = os.listdir(temp_dir)
    if len(extracted_items) == 1 and os.path.isdir(os.path.join(temp_dir, extracted_items[0])):
        extracted_name = extracted_items[0]
    else:
        extracted_name = ''  # フォルダ構造がない場合は直接使う

    exe_name = os.path.basename(sys.executable) if getattr(sys, 'frozen', False) else ''

    config_src = os.path.join(current_dir, 'config.json')
    config_backup = os.path.join(parent, '_config_backup.json')

    new_dir = os.path.join(temp_dir, extracted_name) if extracted_name else temp_dir

    bat_content = f'''@echo off
chcp 65001 > nul
echo 名簿帳票ツールを更新しています...
timeout /t 3 /nobreak > nul
taskkill /IM "{exe_name}" /F > nul 2>&1
timeout /t 2 /nobreak > nul

rem config.json バックアップ
if exist "{config_src}" copy /Y "{config_src}" "{config_backup}" > nul

rem 旧フォルダをリネーム
move "{current_dir}" "{os.path.join(parent, '_old_' + folder_name)}"

rem 新フォルダをリネーム
move "{new_dir}" "{current_dir}"

rem config.json 復元
if exist "{config_backup}" (
    copy /Y "{config_backup}" "{os.path.join(current_dir, 'config.json')}" > nul
    del "{config_backup}" > nul
)

rem クリーンアップ
rmdir /s /q "{os.path.join(parent, '_old_' + folder_name)}" > nul 2>&1
rmdir /s /q "{temp_dir}" > nul 2>&1

rem 再起動
if "{exe_name}" NEQ "" (
    start "" "{os.path.join(current_dir, exe_name)}"
)

rem 自己削除
del "%~f0"
'''

    bat_path = os.path.join(parent, '_update.bat')
    with open(bat_path, 'w', encoding='utf-8') as f:
        f.write(bat_content)

    # バッチ実行
    subprocess.Popen(
        ['cmd', '/c', bat_path],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
    )

    sys.exit(0)
