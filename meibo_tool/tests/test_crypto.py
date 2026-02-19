"""core/crypto.py のテスト"""

import sys

import pytest

from core.crypto import (
    DecryptionError,
    compute_file_hash,
    decrypt_file,
    encrypt_file,
    protect_password,
    unprotect_password,
)

# ── encrypt / decrypt 往復テスト ─────────────────────────────────────────────


class TestEncryptDecrypt:
    """ファイル暗号化 / 復号の往復テスト。"""

    def test_roundtrip_small_file(self, tmp_path):
        """小さいファイルの暗号化→復号が元に戻る。"""
        source = tmp_path / 'sample.txt'
        source.write_bytes(b'Hello, World!')
        encrypted = tmp_path / 'sample.encrypted'
        decrypted = tmp_path / 'sample_dec.txt'

        encrypt_file(str(source), str(encrypted), 'password123')
        decrypt_file(str(encrypted), str(decrypted), 'password123')

        assert decrypted.read_bytes() == b'Hello, World!'

    def test_roundtrip_binary_file(self, tmp_path):
        """バイナリデータ（Excel 等を想定）の往復。"""
        source = tmp_path / 'data.bin'
        content = bytes(range(256)) * 100  # 25.6 KB
        source.write_bytes(content)
        encrypted = tmp_path / 'data.encrypted'
        decrypted = tmp_path / 'data_dec.bin'

        encrypt_file(str(source), str(encrypted), 'テスト日本語パスワード')
        decrypt_file(str(encrypted), str(decrypted), 'テスト日本語パスワード')

        assert decrypted.read_bytes() == content

    def test_roundtrip_empty_file(self, tmp_path):
        """空ファイルの暗号化→復号。"""
        source = tmp_path / 'empty.txt'
        source.write_bytes(b'')
        encrypted = tmp_path / 'empty.encrypted'
        decrypted = tmp_path / 'empty_dec.txt'

        encrypt_file(str(source), str(encrypted), 'pw')
        decrypt_file(str(encrypted), str(decrypted), 'pw')

        assert decrypted.read_bytes() == b''

    def test_encrypted_file_has_magic_header(self, tmp_path):
        """暗号化ファイルの先頭 4 バイトが MBE1。"""
        source = tmp_path / 'test.txt'
        source.write_bytes(b'test data')
        encrypted = tmp_path / 'test.encrypted'

        encrypt_file(str(source), str(encrypted), 'pw')

        with open(encrypted, 'rb') as f:
            assert f.read(4) == b'MBE1'

    def test_different_salt_per_encryption(self, tmp_path):
        """同じファイル・同じパスワードでも暗号化結果が毎回異なる。"""
        source = tmp_path / 'test.txt'
        source.write_bytes(b'same content')
        enc1 = tmp_path / 'enc1.bin'
        enc2 = tmp_path / 'enc2.bin'

        encrypt_file(str(source), str(enc1), 'pw')
        encrypt_file(str(source), str(enc2), 'pw')

        assert enc1.read_bytes() != enc2.read_bytes()


class TestDecryptErrors:
    """復号エラーのテスト。"""

    def test_wrong_password(self, tmp_path):
        """パスワード不正で DecryptionError。"""
        source = tmp_path / 'test.txt'
        source.write_bytes(b'secret data')
        encrypted = tmp_path / 'test.encrypted'
        decrypted = tmp_path / 'test_dec.txt'

        encrypt_file(str(source), str(encrypted), 'correct_password')

        with pytest.raises(DecryptionError, match='パスワード'):
            decrypt_file(str(encrypted), str(decrypted), 'wrong_password')

    def test_invalid_magic(self, tmp_path):
        """マジックバイトが不正な場合。"""
        bad_file = tmp_path / 'bad.encrypted'
        bad_file.write_bytes(b'XXXX' + b'\x00' * 100)
        decrypted = tmp_path / 'dec.txt'

        with pytest.raises(DecryptionError, match='マジックバイト'):
            decrypt_file(str(bad_file), str(decrypted), 'pw')

    def test_truncated_file(self, tmp_path):
        """ヘッダーが途中で切れたファイル。"""
        truncated = tmp_path / 'truncated.encrypted'
        truncated.write_bytes(b'MBE1' + b'\x00' * 10)  # salt 不足
        decrypted = tmp_path / 'dec.txt'

        with pytest.raises(DecryptionError):
            decrypt_file(str(truncated), str(decrypted), 'pw')

    def test_corrupted_ciphertext(self, tmp_path):
        """暗号文の一部が破損した場合。"""
        source = tmp_path / 'test.txt'
        source.write_bytes(b'test data for corruption')
        encrypted = tmp_path / 'test.encrypted'
        decrypted = tmp_path / 'test_dec.txt'

        encrypt_file(str(source), str(encrypted), 'pw')

        # 末尾バイトを書き換え
        data = bytearray(encrypted.read_bytes())
        data[-1] ^= 0xFF
        encrypted.write_bytes(bytes(data))

        with pytest.raises(DecryptionError, match='パスワード'):
            decrypt_file(str(encrypted), str(decrypted), 'pw')


# ── compute_file_hash テスト ─────────────────────────────────────────────────


class TestFileHash:
    """ファイルハッシュのテスト。"""

    def test_same_content_same_hash(self, tmp_path):
        """同じ内容なら同じハッシュ。"""
        f1 = tmp_path / 'a.txt'
        f2 = tmp_path / 'b.txt'
        f1.write_bytes(b'hello')
        f2.write_bytes(b'hello')

        assert compute_file_hash(str(f1)) == compute_file_hash(str(f2))

    def test_different_content_different_hash(self, tmp_path):
        """異なる内容なら異なるハッシュ。"""
        f1 = tmp_path / 'a.txt'
        f2 = tmp_path / 'b.txt'
        f1.write_bytes(b'hello')
        f2.write_bytes(b'world')

        assert compute_file_hash(str(f1)) != compute_file_hash(str(f2))

    def test_hash_is_hex_sha256(self, tmp_path):
        """ハッシュ値が 64 文字の16進文字列 (SHA-256)。"""
        f = tmp_path / 'test.txt'
        f.write_bytes(b'data')

        h = compute_file_hash(str(f))
        assert len(h) == 64
        assert all(c in '0123456789abcdef' for c in h)


# ── protect / unprotect パスワード保護テスト ──────────────────────────────────


class TestPasswordProtection:
    """パスワード保護 / 復元のテスト。"""

    def test_empty_password(self):
        """空パスワードは空文字列を返す。"""
        assert protect_password('') == ''
        assert unprotect_password('') == ''

    @pytest.mark.skipif(sys.platform != 'win32', reason='DPAPI is Windows only')
    def test_dpapi_roundtrip(self):
        """DPAPI での保護→復元が往復する。"""
        original = 'テスト学校パスワード2025!'
        protected = protect_password(original)

        assert protected.startswith('dpapi:')
        assert protected != original
        assert unprotect_password(protected) == original

    @pytest.mark.skipif(sys.platform != 'win32', reason='DPAPI is Windows only')
    def test_dpapi_ascii_password(self):
        """ASCII パスワードの往復。"""
        original = 'simple_password_123'
        protected = protect_password(original)

        assert unprotect_password(protected) == original

    def test_unprotect_plain_fallback(self):
        """plain: プレフィックスのフォールバック復元。"""
        import base64
        encoded = 'plain:' + base64.b64encode('テスト'.encode()).decode('ascii')
        assert unprotect_password(encoded) == 'テスト'

    def test_unprotect_raw_string(self):
        """プレフィックスなし文字列はそのまま返す（旧 config 互換）。"""
        assert unprotect_password('raw_password') == 'raw_password'
