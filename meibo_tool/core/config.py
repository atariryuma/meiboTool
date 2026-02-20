"""設定ファイル（config.json）管理"""

import json
import os
import sys
from datetime import date
from typing import Any


def _current_fiscal_year() -> int:
    """現在の年度（西暦）を返す。4月以降は当年、3月以前は前年。"""
    today = date.today()
    return today.year if today.month >= 4 else today.year - 1


def _get_app_dir() -> str:
    """アプリの実行ディレクトリを返す（ユーザー書き込み用）。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _get_bundle_dir() -> str:
    """PyInstaller バンドルデータのディレクトリを返す（読み取り専用リソース用）。

    frozen 時は sys._MEIPASS（_internal/）、開発時はプロジェクトルート。
    """
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _get_config_path() -> str:
    """config.json の絶対パスを返す。exe / 開発どちらでも動作する。

    ユーザーが編集した config.json は exe と同じディレクトリに保存される。
    """
    return os.path.join(_get_app_dir(), 'config.json')


def _default_config() -> dict[str, Any]:
    """デフォルト設定を返す。fiscal_year を呼び出し時点で計算する。"""
    return {
        'app_version': '1.0.0',
        'school_name': '',
        'school_type': 'elementary',
        'template_dir': './テンプレート',
        'output_dir': './出力',
        'layout_dir': './レイアウト',
        'default_font': 'IPAmj明朝',
        'fiscal_year': _current_fiscal_year(),
        'graduation_cert_start_number': 1,
        'homeroom_teachers': {},
        'last_loaded_file': '',
        'update': {
            'github_repo': '',
            'check_on_startup': True,
            'current_app_version': '1.0.0',
            'last_check_time': '',
            'skip_version': '',
        },
        'data_source': {
            'mode': 'manual',           # 'manual' | 'lan' | 'gdrive'
            'lan_path': '',
            'gdrive_file_id': '',
            'encryption_password': '',  # DPAPI で保護して保存
            'last_sync_hash': '',
            'last_sync_time': '',
            'cache_file': '',
        },
    }


def _deep_merge(base: dict, override: dict) -> dict:
    """ネストした辞書を再帰的にマージする。override が優先。"""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def load_config() -> dict[str, Any]:
    """config.json を読み込む。存在しない / 不正な場合はデフォルト値を返す。

    読み込み優先順位:
      1. exe ディレクトリの config.json（ユーザー編集版）
      2. バンドルディレクトリの config.json（初期同梱版）
      3. デフォルト値
    """
    defaults = _default_config()

    # 読み込み候補パス（ユーザー版 → バンドル版）
    candidates = [_get_config_path()]
    bundle_path = os.path.join(_get_bundle_dir(), 'config.json')
    if bundle_path != candidates[0]:
        candidates.append(bundle_path)

    for path in candidates:
        if not os.path.exists(path):
            continue
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            return _deep_merge(defaults, data)
        except (json.JSONDecodeError, OSError):
            continue

    return _deep_merge(defaults, {})


def save_config(config: dict[str, Any]) -> None:
    """config を config.json に保存する。"""
    path = _get_config_path()
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def get_template_dir(config: dict[str, Any]) -> str:
    """テンプレートフォルダの絶対パスを返す。

    解決順序:
      1. exe ディレクトリからの相対パス（ユーザーがテンプレートを上書きした場合）
      2. バンドルディレクトリ内のテンプレート（PyInstaller 同梱版）
    """
    raw = config.get('template_dir', './テンプレート')
    if os.path.isabs(raw):
        return raw
    # exe ディレクトリから解決
    app_path = os.path.normpath(os.path.join(_get_app_dir(), raw))
    if os.path.isdir(app_path):
        return app_path
    # バンドルディレクトリにフォールバック（PyInstaller frozen）
    bundle_path = os.path.normpath(os.path.join(_get_bundle_dir(), raw))
    if os.path.isdir(bundle_path):
        return bundle_path
    return app_path


def get_cache_dir() -> str:
    """キャッシュディレクトリの絶対パスを返す。存在しない場合は作成する。"""
    base = os.path.dirname(_get_config_path())
    path = os.path.join(base, 'cache')
    os.makedirs(path, exist_ok=True)
    return path


def get_output_dir(config: dict[str, Any]) -> str:
    """出力フォルダの絶対パスを返す。存在しない場合は作成する。"""
    raw = config.get('output_dir', './出力')
    if os.path.isabs(raw):
        path = raw
    else:
        base = os.path.dirname(_get_config_path())
        path = os.path.normpath(os.path.join(base, raw))
    os.makedirs(path, exist_ok=True)
    return path


def get_layout_dir(config: dict[str, Any]) -> str:
    """レイアウトライブラリフォルダの絶対パスを返す。存在しない場合は作成する。"""
    raw = config.get('layout_dir', './レイアウト')
    if os.path.isabs(raw):
        path = raw
    else:
        base = os.path.dirname(_get_config_path())
        path = os.path.normpath(os.path.join(base, raw))
    os.makedirs(path, exist_ok=True)
    return path
