"""core/updater.py のテスト"""

from __future__ import annotations

import hashlib
import zipfile
from unittest.mock import MagicMock, patch

from core.updater import (
    _parse_version,
    check_for_update,
    compute_local_manifest,
    diff_manifests,
    extract_changed_files,
    is_newer,
)

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


def _mock_release_response(
    tag: str,
    body: str = 'release notes',
    zip_name: str = 'app.zip',
    include_manifest: bool = False,
):
    """GitHub Releases API のモックレスポンスを作成する。"""
    assets = [
        {
            'name': zip_name,
            'browser_download_url': f'https://github.com/releases/download/{tag}/{zip_name}',
            'size': 1024000,
        },
    ]
    if include_manifest:
        assets.append({
            'name': 'manifest.json',
            'browser_download_url': f'https://github.com/releases/download/{tag}/manifest.json',
            'size': 512,
        })
    resp = MagicMock()
    resp.raise_for_status.return_value = None
    resp.json.return_value = {
        'tag_name': tag,
        'body': body,
        'published_at': '2026-02-19T00:00:00Z',
        'assets': assets,
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

    @patch('core.updater.save_config')
    @patch('core.updater.requests.get')
    def test_manifest_url_included(self, mock_get, mock_save):
        mock_get.return_value = _mock_release_response('v1.1.0', include_manifest=True)
        config = _make_config()

        result = check_for_update(config)
        assert result is not None
        assert result.manifest_url is not None
        assert 'manifest.json' in result.manifest_url

    @patch('core.updater.save_config')
    @patch('core.updater.requests.get')
    def test_manifest_url_none_when_missing(self, mock_get, mock_save):
        mock_get.return_value = _mock_release_response('v1.1.0', include_manifest=False)
        config = _make_config()

        result = check_for_update(config)
        assert result is not None
        assert result.manifest_url is None


# ── diff_manifests テスト ────────────────────────────────────────────────────


class TestDiffManifests:
    def test_detects_changed_file(self):
        remote = {
            'app.exe': {'sha256': 'aaa', 'size': 100},
            'lib.dll': {'sha256': 'bbb', 'size': 200},
        }
        local = {
            'app.exe': {'sha256': 'aaa', 'size': 100},
            'lib.dll': {'sha256': 'old', 'size': 200},
        }
        result = diff_manifests(remote, local)
        assert result == ['lib.dll']

    def test_detects_new_file(self):
        remote = {
            'app.exe': {'sha256': 'aaa', 'size': 100},
            'new.dll': {'sha256': 'ccc', 'size': 300},
        }
        local = {
            'app.exe': {'sha256': 'aaa', 'size': 100},
        }
        result = diff_manifests(remote, local)
        assert result == ['new.dll']

    def test_no_changes(self):
        files = {
            'app.exe': {'sha256': 'aaa', 'size': 100},
            'lib.dll': {'sha256': 'bbb', 'size': 200},
        }
        result = diff_manifests(files, files)
        assert result == []

    def test_skips_config_json(self):
        remote = {
            'config.json': {'sha256': 'new_hash', 'size': 50},
            'app.exe': {'sha256': 'aaa', 'size': 100},
        }
        local = {
            'config.json': {'sha256': 'old_hash', 'size': 40},
            'app.exe': {'sha256': 'aaa', 'size': 100},
        }
        result = diff_manifests(remote, local)
        assert result == []

    def test_multiple_changes(self):
        remote = {
            'a.dll': {'sha256': 'new1', 'size': 100},
            'b.dll': {'sha256': 'new2', 'size': 200},
            'c.dll': {'sha256': 'same', 'size': 300},
        }
        local = {
            'a.dll': {'sha256': 'old1', 'size': 100},
            'b.dll': {'sha256': 'old2', 'size': 200},
            'c.dll': {'sha256': 'same', 'size': 300},
        }
        result = diff_manifests(remote, local)
        assert set(result) == {'a.dll', 'b.dll'}


# ── compute_local_manifest テスト ────────────────────────────────────────────


class TestComputeLocalManifest:
    def test_computes_hashes(self, tmp_path):
        # ファイル作成
        (tmp_path / 'app.exe').write_bytes(b'hello')
        sub = tmp_path / '_internal'
        sub.mkdir()
        (sub / 'lib.dll').write_bytes(b'world')

        result = compute_local_manifest(str(tmp_path))

        assert 'app.exe' in result
        assert '_internal/lib.dll' in result
        assert result['app.exe']['sha256'] == hashlib.sha256(b'hello').hexdigest()
        assert result['_internal/lib.dll']['sha256'] == hashlib.sha256(b'world').hexdigest()
        assert result['app.exe']['size'] == 5
        assert result['_internal/lib.dll']['size'] == 5

    def test_skips_config_json(self, tmp_path):
        (tmp_path / 'config.json').write_text('{}', encoding='utf-8')
        (tmp_path / 'app.exe').write_bytes(b'x')

        result = compute_local_manifest(str(tmp_path))

        assert 'config.json' not in result
        assert 'app.exe' in result

    def test_empty_directory(self, tmp_path):
        result = compute_local_manifest(str(tmp_path))
        assert result == {}


# ── extract_changed_files テスト ─────────────────────────────────────────────


class TestExtractChangedFiles:
    def test_extracts_only_changed(self, tmp_path):
        # zip を作成（トップレベルフォルダあり）
        zip_path = tmp_path / 'release.zip'
        with zipfile.ZipFile(zip_path, 'w') as zf:
            zf.writestr('app/app.exe', b'new_exe')
            zf.writestr('app/_internal/lib.dll', b'new_lib')
            zf.writestr('app/_internal/unchanged.dll', b'same')

        staging = tmp_path / 'staging'
        staging.mkdir()

        extract_changed_files(
            str(zip_path),
            str(staging),
            ['app.exe', '_internal/lib.dll'],
        )

        assert (staging / 'app.exe').read_bytes() == b'new_exe'
        assert (staging / '_internal' / 'lib.dll').read_bytes() == b'new_lib'
        assert not (staging / '_internal' / 'unchanged.dll').exists()

    def test_no_top_level_folder(self, tmp_path):
        zip_path = tmp_path / 'flat.zip'
        with zipfile.ZipFile(zip_path, 'w') as zf:
            zf.writestr('app.exe', b'data')
            zf.writestr('lib.dll', b'other')

        staging = tmp_path / 'staging'
        staging.mkdir()

        extract_changed_files(str(zip_path), str(staging), ['app.exe'])

        assert (staging / 'app.exe').read_bytes() == b'data'
        assert not (staging / 'lib.dll').exists()

    def test_nested_subdirectory(self, tmp_path):
        zip_path = tmp_path / 'release.zip'
        with zipfile.ZipFile(zip_path, 'w') as zf:
            zf.writestr('app/_internal/sub/deep.dll', b'deep_data')

        staging = tmp_path / 'staging'
        staging.mkdir()

        extract_changed_files(
            str(zip_path), str(staging), ['_internal/sub/deep.dll'],
        )

        assert (staging / '_internal' / 'sub' / 'deep.dll').read_bytes() == b'deep_data'
