"""設定ファイル（config.json）管理"""

import json
import os
import sys
from typing import Any


def _get_config_path() -> str:
    """config.json の絶対パスを返す。exe / 開発どちらでも動作する。"""
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        # config.py は meibo_tool/core/ にある。
        # meibo_tool/core/config.py → core → meibo_tool → project_root
        base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, 'config.json')


_DEFAULT_CONFIG: dict[str, Any] = {
    'app_version': '1.0.0',
    'school_name': '那覇市立天久小学校',
    'school_type': 'elementary',
    'template_dir': './テンプレート',
    'output_dir': './出力',
    'default_font': 'IPAmj明朝',
    'fiscal_year': 2025,
    'graduation_cert_start_number': 1,
    'homeroom_teachers': {},
    'update': {
        'version_file_id': '',
        'check_on_startup': True,
        'current_app_version': '1.0.0',
        'current_template_version': '1.0.0',
    },
    'column_mappings': {'last_used': {}},
    'manual_columns': {
        'last_file_hash': '',
        'mandatory': {},
        'optional': {},
    },
    'recent_files': [],
}


def load_config() -> dict[str, Any]:
    """config.json を読み込む。存在しない場合はデフォルト値を返す。"""
    path = _get_config_path()
    if not os.path.exists(path):
        return dict(_DEFAULT_CONFIG)
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    # デフォルト値で補完（キーが増えた場合の後方互換性）
    merged = dict(_DEFAULT_CONFIG)
    merged.update(data)
    return merged


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
