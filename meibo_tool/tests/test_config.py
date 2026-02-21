"""core/config.py のテスト

テスト対象:
  - _deep_merge: ネスト辞書のマージ
  - load_config / save_config: 読み書き
  - _current_fiscal_year: 年度計算
  - get_template_dir / get_output_dir / get_cache_dir: パス解決
"""

from __future__ import annotations

import json
import os
from unittest.mock import patch

from core.config import (
    _current_fiscal_year,
    _deep_merge,
    get_cache_dir,
    load_config,
    save_config,
)

# ── _deep_merge ───────────────────────────────────────────────────────────────


class TestDeepMerge:
    def test_flat_merge(self):
        base = {'a': 1, 'b': 2}
        override = {'b': 3, 'c': 4}
        result = _deep_merge(base, override)
        assert result == {'a': 1, 'b': 3, 'c': 4}

    def test_nested_merge(self):
        base = {'a': {'x': 1, 'y': 2}}
        override = {'a': {'y': 3, 'z': 4}}
        result = _deep_merge(base, override)
        assert result == {'a': {'x': 1, 'y': 3, 'z': 4}}

    def test_override_replaces_non_dict(self):
        base = {'a': 'string'}
        override = {'a': {'nested': True}}
        result = _deep_merge(base, override)
        assert result == {'a': {'nested': True}}

    def test_empty_override(self):
        base = {'a': 1, 'b': {'c': 2}}
        result = _deep_merge(base, {})
        assert result == base

    def test_empty_base(self):
        override = {'a': 1}
        result = _deep_merge({}, override)
        assert result == {'a': 1}

    def test_deeply_nested(self):
        base = {'a': {'b': {'c': 1, 'd': 2}}}
        override = {'a': {'b': {'d': 3}}}
        result = _deep_merge(base, override)
        assert result == {'a': {'b': {'c': 1, 'd': 3}}}


# ── _current_fiscal_year ──────────────────────────────────────────────────────


class TestCurrentFiscalYear:
    @patch('core.config.date')
    def test_april_returns_current_year(self, mock_date):
        from datetime import date
        mock_date.today.return_value = date(2025, 4, 1)
        mock_date.side_effect = lambda *args, **kwargs: date(*args, **kwargs)
        assert _current_fiscal_year() == 2025

    @patch('core.config.date')
    def test_march_returns_previous_year(self, mock_date):
        from datetime import date
        mock_date.today.return_value = date(2025, 3, 31)
        mock_date.side_effect = lambda *args, **kwargs: date(*args, **kwargs)
        assert _current_fiscal_year() == 2024


# ── load_config / save_config ─────────────────────────────────────────────────


class TestLoadSaveConfig:
    @patch('core.config._get_config_path')
    def test_load_nonexistent_returns_defaults(self, mock_path, tmp_path):
        mock_path.return_value = str(tmp_path / 'nonexistent.json')
        config = load_config()
        assert 'school_name' in config
        assert 'update' in config
        assert 'data_source' in config
        assert config['update']['check_on_startup'] is True

    @patch('core.config._get_config_path')
    def test_save_and_load_roundtrip(self, mock_path, tmp_path):
        config_path = str(tmp_path / 'config.json')
        mock_path.return_value = config_path

        config = load_config()
        config['school_name'] = 'テスト小学校'
        save_config(config)

        loaded = load_config()
        assert loaded['school_name'] == 'テスト小学校'

    @patch('core.config._get_bundle_dir')
    @patch('core.config._get_config_path')
    def test_load_malformed_json_returns_defaults(
        self, mock_path, mock_bundle, tmp_path,
    ):
        config_path = tmp_path / 'config.json'
        config_path.write_text('{invalid json!!!', encoding='utf-8')
        mock_path.return_value = str(config_path)
        # バンドルディレクトリも tmp_path にして実 config.json を読まないようにする
        mock_bundle.return_value = str(tmp_path / 'no_bundle')

        config = load_config()
        # エラーにならずデフォルト値が返される
        assert 'school_name' in config
        assert config['school_name'] == ''

    @patch('core.config._get_config_path')
    def test_backward_compat_old_config(self, mock_path, tmp_path):
        """旧構造の config.json を読んでも新キーが補完される。"""
        config_path = tmp_path / 'config.json'
        old_config = {
            'school_name': '旧小学校',
            'fiscal_year': 2024,
        }
        config_path.write_text(json.dumps(old_config), encoding='utf-8')
        mock_path.return_value = str(config_path)

        config = load_config()
        assert config['school_name'] == '旧小学校'
        assert config['fiscal_year'] == 2024
        # 新キーが補完されている
        assert 'update' in config
        assert 'data_source' in config
        assert config['data_source']['mode'] == 'manual'


# ── get_cache_dir ─────────────────────────────────────────────────────────────


class TestGetCacheDir:
    @patch('core.config._get_config_path')
    def test_creates_cache_dir(self, mock_path, tmp_path):
        mock_path.return_value = str(tmp_path / 'config.json')
        cache_dir = get_cache_dir()
        assert os.path.isdir(cache_dir)
        assert cache_dir.endswith('cache')
