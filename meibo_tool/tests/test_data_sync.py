"""core/data_sync.py のテスト"""

from __future__ import annotations

import os
from unittest.mock import MagicMock, patch

from core.crypto import compute_file_hash, encrypt_file, protect_password
from core.data_sync import sync


def _make_config(tmp_path, **ds_overrides) -> dict:
    """テスト用 config を生成する。cache_dir が tmp_path を使うようにする。"""
    config = {
        'data_source': {
            'mode': 'manual',
            'lan_path': '',
            'gdrive_file_id': '',
            'encryption_password': '',
            'last_sync_hash': '',
            'last_sync_time': '',
            'cache_file': '',
        },
    }
    config['data_source'].update(ds_overrides)
    return config


# ── Manual モード ────────────────────────────────────────────────────────────


class TestManualMode:
    def test_manual_returns_unchanged(self, tmp_path):
        config = _make_config(tmp_path, mode='manual')
        result = sync(config)
        assert result.status == 'unchanged'

    def test_unknown_mode_returns_error(self, tmp_path):
        config = _make_config(tmp_path, mode='unknown_mode')
        result = sync(config)
        assert result.status == 'error'
        assert '不明なモード' in result.message


# ── LAN モード ───────────────────────────────────────────────────────────────


class TestLanMode:
    def test_empty_path_returns_error(self, tmp_path):
        config = _make_config(tmp_path, mode='lan', lan_path='')
        result = sync(config)
        assert result.status == 'error'
        assert 'パス' in result.message

    def test_nonexistent_path_no_cache(self, tmp_path):
        config = _make_config(tmp_path, mode='lan', lan_path=str(tmp_path / 'no_such_file.xlsx'))
        result = sync(config)
        assert result.status == 'unavailable'
        assert '接続できません' in result.message

    def test_nonexistent_path_with_cache(self, tmp_path):
        cache_file = tmp_path / 'old_cache.xlsx'
        cache_file.write_bytes(b'old data')
        config = _make_config(
            tmp_path,
            mode='lan',
            lan_path=str(tmp_path / 'no_such_file.xlsx'),
            cache_file=str(cache_file),
        )
        result = sync(config)
        assert result.status == 'unavailable'
        assert result.path == str(cache_file)

    @patch('core.data_sync.get_cache_dir')
    def test_new_file_returns_updated(self, mock_cache_dir, tmp_path):
        mock_cache_dir.return_value = str(tmp_path / 'cache')
        os.makedirs(tmp_path / 'cache', exist_ok=True)

        source = tmp_path / 'roster.xlsx'
        source.write_bytes(b'student data here')

        config = _make_config(tmp_path, mode='lan', lan_path=str(source))
        result = sync(config)

        assert result.status == 'updated'
        assert os.path.exists(result.path)
        # キャッシュファイルの内容が一致
        with open(result.path, 'rb') as f:
            assert f.read() == b'student data here'
        # config_updates に更新情報が含まれている
        assert result.config_updates is not None
        assert result.config_updates['last_sync_hash'] == compute_file_hash(str(source))

    @patch('core.data_sync.get_cache_dir')
    def test_unchanged_file_returns_unchanged(self, mock_cache_dir, tmp_path):
        mock_cache_dir.return_value = str(tmp_path / 'cache')
        os.makedirs(tmp_path / 'cache', exist_ok=True)

        source = tmp_path / 'roster.xlsx'
        source.write_bytes(b'student data')
        file_hash = compute_file_hash(str(source))

        cache_file = tmp_path / 'cache' / 'roster_cache.xlsx'
        cache_file.write_bytes(b'student data')

        config = _make_config(
            tmp_path,
            mode='lan',
            lan_path=str(source),
            last_sync_hash=file_hash,
            cache_file=str(cache_file),
        )
        result = sync(config)

        assert result.status == 'unchanged'

    @patch('core.data_sync.get_cache_dir')
    def test_changed_file_returns_updated_with_new_hash(self, mock_cache_dir, tmp_path):
        mock_cache_dir.return_value = str(tmp_path / 'cache')
        os.makedirs(tmp_path / 'cache', exist_ok=True)

        source = tmp_path / 'roster.xlsx'
        source.write_bytes(b'old data')
        old_hash = compute_file_hash(str(source))

        # ファイルを更新
        source.write_bytes(b'new data')

        config = _make_config(
            tmp_path,
            mode='lan',
            lan_path=str(source),
            last_sync_hash=old_hash,
        )
        result = sync(config)

        assert result.status == 'updated'
        # config_updates のハッシュが更新されている
        assert result.config_updates is not None
        new_hash = result.config_updates['last_sync_hash']
        assert new_hash != old_hash
        assert new_hash == compute_file_hash(str(source))


# ── Google Drive モード ──────────────────────────────────────────────────────


class TestGdriveMode:
    def test_no_file_id_returns_error(self, tmp_path):
        config = _make_config(tmp_path, mode='gdrive', gdrive_file_id='')
        result = sync(config)
        assert result.status == 'error'
        assert 'ファイル ID' in result.message

    def test_no_password_returns_error(self, tmp_path):
        config = _make_config(
            tmp_path,
            mode='gdrive',
            gdrive_file_id='some_id',
            encryption_password='',
        )
        result = sync(config)
        assert result.status == 'error'
        assert 'パスワード' in result.message

    @patch('core.data_sync.get_cache_dir')
    @patch('core.data_sync.requests.get')
    def test_successful_download_and_decrypt(self, mock_get, mock_cache_dir, tmp_path):
        cache_dir = tmp_path / 'cache'
        os.makedirs(cache_dir, exist_ok=True)
        mock_cache_dir.return_value = str(cache_dir)

        # 暗号化ファイルを事前に作成
        source = tmp_path / 'roster.xlsx'
        source.write_bytes(b'student PII data')
        encrypted = tmp_path / 'roster.encrypted'
        encrypt_file(str(source), str(encrypted), 'school_password')
        encrypted_data = encrypted.read_bytes()

        # requests.get のモック
        mock_resp = MagicMock()
        mock_resp.raise_for_status.return_value = None
        mock_resp.iter_content.return_value = [encrypted_data]
        mock_get.return_value = mock_resp

        config = _make_config(
            tmp_path,
            mode='gdrive',
            gdrive_file_id='test_file_id',
            encryption_password=protect_password('school_password'),
        )
        result = sync(config)

        assert result.status == 'updated'
        assert os.path.exists(result.path)
        with open(result.path, 'rb') as f:
            assert f.read() == b'student PII data'
        # config_updates が返される
        assert result.config_updates is not None
        assert result.config_updates['last_sync_hash'] != ''

    @patch('core.data_sync.get_cache_dir')
    @patch('core.data_sync.requests.get')
    def test_wrong_password_returns_decrypt_error(self, mock_get, mock_cache_dir, tmp_path):
        cache_dir = tmp_path / 'cache'
        os.makedirs(cache_dir, exist_ok=True)
        mock_cache_dir.return_value = str(cache_dir)

        # 別のパスワードで暗号化
        source = tmp_path / 'roster.xlsx'
        source.write_bytes(b'data')
        encrypted = tmp_path / 'roster.encrypted'
        encrypt_file(str(source), str(encrypted), 'correct_password')

        mock_resp = MagicMock()
        mock_resp.raise_for_status.return_value = None
        mock_resp.iter_content.return_value = [encrypted.read_bytes()]
        mock_get.return_value = mock_resp

        config = _make_config(
            tmp_path,
            mode='gdrive',
            gdrive_file_id='test_id',
            encryption_password=protect_password('wrong_password'),
        )
        result = sync(config)

        assert result.status == 'decrypt_error'

    @patch('core.data_sync.requests.get')
    def test_network_failure_no_cache(self, mock_get, tmp_path):
        import requests as req
        mock_get.side_effect = req.ConnectionError('network down')

        config = _make_config(
            tmp_path,
            mode='gdrive',
            gdrive_file_id='test_id',
            encryption_password=protect_password('pw'),
        )
        result = sync(config)

        assert result.status == 'unavailable'
        assert '接続できません' in result.message

    @patch('core.data_sync.requests.get')
    def test_network_failure_with_cache(self, mock_get, tmp_path):
        import requests as req
        mock_get.side_effect = req.ConnectionError('network down')

        cache_file = tmp_path / 'old_cache.xlsx'
        cache_file.write_bytes(b'cached data')

        config = _make_config(
            tmp_path,
            mode='gdrive',
            gdrive_file_id='test_id',
            encryption_password=protect_password('pw'),
            cache_file=str(cache_file),
        )
        result = sync(config)

        assert result.status == 'unavailable'
        assert result.path == str(cache_file)

    @patch('core.data_sync.get_cache_dir')
    @patch('core.data_sync.requests.get')
    def test_unchanged_file_returns_unchanged(self, mock_get, mock_cache_dir, tmp_path):
        cache_dir = tmp_path / 'cache'
        os.makedirs(cache_dir, exist_ok=True)
        mock_cache_dir.return_value = str(cache_dir)

        source = tmp_path / 'roster.xlsx'
        source.write_bytes(b'data')
        encrypted = tmp_path / 'roster.encrypted'
        encrypt_file(str(source), str(encrypted), 'pw')
        encrypted_data = encrypted.read_bytes()
        enc_hash = compute_file_hash(str(encrypted))

        mock_resp = MagicMock()
        mock_resp.raise_for_status.return_value = None
        mock_resp.iter_content.return_value = [encrypted_data]
        mock_get.return_value = mock_resp

        cache_file = cache_dir / 'roster_cache.xlsx'
        cache_file.write_bytes(b'data')

        config = _make_config(
            tmp_path,
            mode='gdrive',
            gdrive_file_id='test_id',
            encryption_password=protect_password('pw'),
            last_sync_hash=enc_hash,
            cache_file=str(cache_file),
        )
        result = sync(config)

        assert result.status == 'unchanged'
