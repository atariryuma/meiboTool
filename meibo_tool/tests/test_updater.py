"""core/updater.py のテスト"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from core.updater import _parse_version, check_for_update, is_newer

# ── バージョン比較テスト ─────────────────────────────────────────────────────


class TestParseVersion:
    def test_simple(self):
        assert _parse_version('1.2.3') == (1, 2, 3)

    def test_with_v_prefix(self):
        assert _parse_version('v1.2.3') == (1, 2, 3)

    def test_two_parts(self):
        assert _parse_version('1.0') == (1, 0)


class TestIsNewer:
    def test_newer_patch(self):
        assert is_newer('1.0.1', '1.0.0') is True

    def test_newer_minor(self):
        assert is_newer('1.1.0', '1.0.0') is True

    def test_newer_major(self):
        assert is_newer('2.0.0', '1.9.9') is True

    def test_same_version(self):
        assert is_newer('1.0.0', '1.0.0') is False

    def test_older_version(self):
        assert is_newer('1.0.0', '1.0.1') is False

    def test_with_v_prefix(self):
        assert is_newer('v1.1.0', '1.0.0') is True

    def test_invalid_remote(self):
        assert is_newer('invalid', '1.0.0') is False

    def test_none_remote(self):
        assert is_newer(None, '1.0.0') is False


# ── check_for_update テスト ──────────────────────────────────────────────────


def _make_config(**update_overrides) -> dict:
    config = {
        'update': {
            'github_repo': 'test-owner/test-repo',
            'check_on_startup': True,
            'current_app_version': '1.0.0',
            'last_check_time': '',
            'skip_version': '',
        },
    }
    config['update'].update(update_overrides)
    return config


def _mock_release_response(tag: str, body: str = 'release notes', zip_name: str = 'app.zip'):
    """GitHub Releases API のモックレスポンスを作成する。"""
    resp = MagicMock()
    resp.raise_for_status.return_value = None
    resp.json.return_value = {
        'tag_name': tag,
        'body': body,
        'published_at': '2026-02-19T00:00:00Z',
        'assets': [
            {
                'name': zip_name,
                'browser_download_url': f'https://github.com/releases/download/{tag}/{zip_name}',
                'size': 1024000,
            },
        ],
    }
    return resp


class TestCheckForUpdate:
    @patch('core.updater.save_config')
    @patch('core.updater.requests.get')
    def test_new_version_available(self, mock_get, mock_save):
        mock_get.return_value = _mock_release_response('v1.1.0')
        config = _make_config()

        result = check_for_update(config)

        assert result is not None
        assert result.new_version == 'v1.1.0'
        assert result.current_version == '1.0.0'
        assert result.release_notes == 'release notes'
        assert result.asset_url.endswith('app.zip')

    @patch('core.updater.requests.get')
    def test_same_version_returns_none(self, mock_get):
        mock_get.return_value = _mock_release_response('v1.0.0')
        config = _make_config()

        result = check_for_update(config)
        assert result is None

    @patch('core.updater.requests.get')
    def test_older_version_returns_none(self, mock_get):
        mock_get.return_value = _mock_release_response('v0.9.0')
        config = _make_config()

        result = check_for_update(config)
        assert result is None

    @patch('core.updater.requests.get')
    def test_skip_version(self, mock_get):
        mock_get.return_value = _mock_release_response('v1.1.0')
        config = _make_config(skip_version='v1.1.0')

        result = check_for_update(config)
        assert result is None

    def test_no_repo_configured(self):
        config = _make_config(github_repo='')
        result = check_for_update(config)
        assert result is None

    def test_check_disabled(self):
        config = _make_config(check_on_startup=False)
        result = check_for_update(config)
        assert result is None

    @patch('core.updater.requests.get')
    def test_network_error_returns_none(self, mock_get):
        import requests as req
        mock_get.side_effect = req.ConnectionError('offline')
        config = _make_config()

        result = check_for_update(config)
        assert result is None

    @patch('core.updater.requests.get')
    def test_no_zip_asset_returns_none(self, mock_get):
        resp = MagicMock()
        resp.raise_for_status.return_value = None
        resp.json.return_value = {
            'tag_name': 'v2.0.0',
            'body': '',
            'published_at': '',
            'assets': [
                {'name': 'readme.txt', 'browser_download_url': '', 'size': 100},
            ],
        }
        mock_get.return_value = resp
        config = _make_config()

        result = check_for_update(config)
        assert result is None

    @patch('core.updater.save_config')
    @patch('core.updater.requests.get')
    def test_updates_last_check_time(self, mock_get, mock_save):
        mock_get.return_value = _mock_release_response('v1.1.0')
        config = _make_config()

        check_for_update(config)

        assert config['update']['last_check_time'] != ''
