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


def _get_config_path() -> str:
    """config.json の絶対パスを返す。exe / 開発どちらでも動作する。"""
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        # config.py は meibo_tool/core/ にある。
        # meibo_tool/core/config.py → core → meibo_tool → project_root
        base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, 'config.json')


def _default_config() -> dict[str, Any]:
    """デフォルト設定を返す。fiscal_year を呼び出し時点で計算する。"""
    return {
        'app_version': '1.0.0',
        'school_name': '',
        'school_type': 'elementary',
        'template_dir': './テンプレート',
        'output_dir': './出力',
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
    """config.json を読み込む。存在しない / 不正な場合はデフォルト値を返す。"""
    defaults = _default_config()
    path = _get_config_path()
    if not os.path.exists(path):
        return _deep_merge(defaults, {})
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return _deep_merge(defaults, {})
    # デフォルト値で補完（ネスト辞書も再帰マージで後方互換）
    return _deep_merge(defaults, data)


def save_config(config: dict[str, Any]) -> None:
    """config を config.json に保存する。"""
    path = _get_config_path()
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def get_template_dir(config: dict[str, Any]) -> str:
    """テンプレートフォルダの絶対パスを返す。"""
    raw = config.get('template_dir', './テンプレート')
    if os.path.isabs(raw):
        return raw
    base = os.path.dirname(_get_config_path())
    return os.path.normpath(os.path.join(base, raw))


def get_cache_dir(config: dict[str, Any] | None = None) -> str:
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
