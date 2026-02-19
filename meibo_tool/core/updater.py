"""GitHub Releases ベースのアプリ自動更新

起動時にバックグラウンドスレッドで check_for_update() を呼び出す。
新バージョンがある場合は UpdateInfo を返し、GUI がダイアログを表示する。
"""

from __future__ import annotations

import contextlib
import hashlib
import logging
import os
import subprocess
import sys
import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import requests

from core.config import save_config

logger = logging.getLogger(__name__)

_API_BASE = 'https://api.github.com/repos'
_REQUEST_TIMEOUT = 5  # seconds

# 差分アップデート時にスキップするファイル（ユーザー編集可能）
_SKIP_FILES: frozenset[str] = frozenset({'config.json'})


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
    manifest_url: str | None = field(default=None)


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

    # zip / manifest アセットを探す
    assets = data.get('assets', [])
    zip_asset = None
    manifest_asset = None
    for asset in assets:
        name = asset.get('name', '')
        if name.endswith('.zip'):
            zip_asset = asset
        elif name == 'manifest.json':
            manifest_asset = asset

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
        manifest_url=(
            manifest_asset['browser_download_url']
            if manifest_asset else None
        ),
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


# ── 差分アップデート ─────────────────────────────────────────────────────────

def download_manifest(url: str) -> dict[str, Any]:
    """manifest.json をダウンロードして dict で返す。"""
    resp = requests.get(url, timeout=_REQUEST_TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def compute_local_manifest(app_dir: str) -> dict[str, dict[str, Any]]:
    """ローカルファイルの SHA-256 ハッシュマニフェストを計算する。

    config.json 等のユーザー編集ファイルはスキップする。

    Returns:
        {"相対パス": {"sha256": "...", "size": N}, ...}
    """
    result: dict[str, dict[str, Any]] = {}
    for root, _dirs, files in os.walk(app_dir):
        for fname in files:
            full = os.path.join(root, fname)
            rel = os.path.relpath(full, app_dir).replace('\\', '/')
            if rel in _SKIP_FILES:
                continue
            sha = hashlib.sha256()
            with open(full, 'rb') as f:
                while True:
                    chunk = f.read(65536)
                    if not chunk:
                        break
                    sha.update(chunk)
            result[rel] = {
                'sha256': sha.hexdigest(),
                'size': os.path.getsize(full),
            }
    return result


def diff_manifests(
    remote: dict[str, dict[str, Any]],
    local: dict[str, dict[str, Any]],
) -> list[str]:
    """リモートとローカルのマニフェストを比較し、変更・追加ファイルのリストを返す。

    Returns:
        変更が必要なファイルの相対パスリスト
    """
    changed: list[str] = []
    for rel_path, remote_info in remote.items():
        if rel_path in _SKIP_FILES:
            continue
        local_info = local.get(rel_path)
        if local_info is None or local_info['sha256'] != remote_info['sha256']:
            changed.append(rel_path)
    return changed


def extract_changed_files(
    zip_path: str,
    staging_dir: str,
    changed_files: list[str],
) -> None:
    """zip から変更ファイルのみをステージングディレクトリに展開する。

    zip 内のトップレベルフォルダを除去してフラットに展開する。
    """
    with zipfile.ZipFile(zip_path, 'r') as zf:
        # トップレベルフォルダ名を検出
        names = zf.namelist()
        top_prefix = ''
        if names:
            first = names[0]
            if '/' in first:
                top_prefix = first.split('/')[0] + '/'

        changed_set = set(changed_files)
        for info in zf.infolist():
            if info.is_dir():
                continue
            # トップレベルフォルダを除去した相対パス
            inner_path = info.filename
            if top_prefix and inner_path.startswith(top_prefix):
                inner_path = inner_path[len(top_prefix):]
            if not inner_path:
                continue
            if inner_path in changed_set:
                dest = os.path.join(staging_dir, inner_path.replace('/', os.sep))
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with zf.open(info) as src, open(dest, 'wb') as dst:
                    while True:
                        chunk = src.read(65536)
                        if not chunk:
                            break
                        dst.write(chunk)


# ── バッチファイル生成（--onedir 対応） ──────────────────────────────────────

def generate_update_batch(
    zip_path: str,
    current_dir: str,
    changed_files: list[str] | None = None,
    staging_dir: str | None = None,
) -> None:
    """自己更新バッチを生成して実行し、アプリを終了する。

    差分モード (changed_files + staging_dir 指定時):
      ステージングから変更ファイルのみ上書きコピーする軽量バッチ。

    フル置換モード (changed_files=None):
      zip 全展開 → フォルダ全体を置換する既存方式。

    Args:
        zip_path: ダウンロードした Release zip のパス
        current_dir: 現在のアプリフォルダ
        changed_files: 差分更新対象の相対パスリスト
        staging_dir: 変更ファイルが展開されたステージングディレクトリ
    """
    parent = os.path.dirname(current_dir)
    exe_name = os.path.basename(sys.executable) if getattr(sys, 'frozen', False) else ''

    if changed_files is not None and staging_dir is not None:
        # ── 差分モード ──
        bat_content = _generate_diff_batch(
            current_dir, parent, staging_dir, changed_files, exe_name,
        )
    else:
        # ── フル置換モード（後方互換） ──
        bat_content = _generate_full_batch(
            zip_path, current_dir, parent, exe_name,
        )

    bat_path = os.path.join(parent, '_update.bat')
    with open(bat_path, 'w', encoding='utf-8') as f:
        f.write(bat_content)

    # バッチ実行
    subprocess.Popen(
        ['cmd', '/c', bat_path],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
    )

    sys.exit(0)


def _generate_diff_batch(
    current_dir: str,
    parent: str,
    staging_dir: str,
    changed_files: list[str],
    exe_name: str,
) -> str:
    """差分コピーバッチスクリプトを生成する。"""
    # xcopy コマンドのリストを生成
    copy_cmds: list[str] = []
    for rel in changed_files:
        src = os.path.join(staging_dir, rel.replace('/', os.sep))
        dst = os.path.join(current_dir, rel.replace('/', os.sep))
        dst_dir = os.path.dirname(dst)
        copy_cmds.append(f'if not exist "{dst_dir}" mkdir "{dst_dir}"')
        copy_cmds.append(f'copy /Y "{src}" "{dst}" > nul')

    copy_block = '\n'.join(copy_cmds)

    return f'''@echo off
chcp 65001 > nul
echo 名簿帳票ツールを更新しています（差分更新: {len(changed_files)} ファイル）...
timeout /t 3 /nobreak > nul
taskkill /IM "{exe_name}" /F > nul 2>&1
timeout /t 2 /nobreak > nul

rem 変更ファイルを上書きコピー
{copy_block}

rem ステージングをクリーンアップ
rmdir /s /q "{staging_dir}" > nul 2>&1

rem 再起動
if "{exe_name}" NEQ "" (
    start "" "{os.path.join(current_dir, exe_name)}"
)

rem 自己削除
del "%~f0"
'''


def _generate_full_batch(
    zip_path: str,
    current_dir: str,
    parent: str,
    exe_name: str,
) -> str:
    """フル置換バッチスクリプトを生成する。"""
    folder_name = os.path.basename(current_dir)
    temp_dir = os.path.join(parent, '_update_temp')

    # zip 展開
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(temp_dir)

    # 展開されたフォルダ名を取得
    extracted_items = os.listdir(temp_dir)
    if len(extracted_items) == 1 and os.path.isdir(os.path.join(temp_dir, extracted_items[0])):
        extracted_name = extracted_items[0]
    else:
        extracted_name = ''

    config_src = os.path.join(current_dir, 'config.json')
    config_backup = os.path.join(parent, '_config_backup.json')
    new_dir = os.path.join(temp_dir, extracted_name) if extracted_name else temp_dir

    return f'''@echo off
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
