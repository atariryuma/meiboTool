"""AES-256-GCM ファイル暗号化 / 復号化

名簿データ（児童・生徒の PII 含む）を Google Drive 経由で共有する際の暗号化。
学校ごとの共有パスワードから鍵を導出し、ファイル全体を暗号化する。

暗号化ファイル形式 (Version 1):
    [4B magic: b'MBE1'] [16B salt] [12B nonce] [ciphertext + 16B GCM tag]
"""

from __future__ import annotations

import base64
import ctypes
import ctypes.wintypes
import hashlib
import secrets
import sys
from typing import Final

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

# ── 定数 ─────────────────────────────────────────────────────────────────────

MAGIC: Final[bytes] = b'MBE1'
SALT_SIZE: Final[int] = 16
NONCE_SIZE: Final[int] = 12
KEY_SIZE: Final[int] = 32  # 256 bits
PBKDF2_ITERATIONS: Final[int] = 600_000  # OWASP 2023 推奨値


class DecryptionError(Exception):
    """復号化失敗（パスワード不正 or ファイル破損）。"""


# ── 鍵導出 ───────────────────────────────────────────────────────────────────

def _derive_key(password: str, salt: bytes) -> bytes:
    """パスワード + salt から AES-256 鍵を導出する (PBKDF2-HMAC-SHA256)。"""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=KEY_SIZE,
        salt=salt,
        iterations=PBKDF2_ITERATIONS,
    )
    return kdf.derive(password.encode('utf-8'))


# ── ファイル暗号化 / 復号化 ──────────────────────────────────────────────────

def encrypt_file(source_path: str, dest_path: str, password: str) -> None:
    """ファイルを AES-256-GCM で暗号化して dest_path に書き出す。

    Args:
        source_path: 暗号化元ファイルパス
        dest_path: 暗号化済みファイルの出力先パス
        password: 暗号化パスワード（学校ごとの共有パスワード）
    """
    salt = secrets.token_bytes(SALT_SIZE)
    nonce = secrets.token_bytes(NONCE_SIZE)
    key = _derive_key(password, salt)

    with open(source_path, 'rb') as f:
        plaintext = f.read()

    aesgcm = AESGCM(key)
    ciphertext = aesgcm.encrypt(nonce, plaintext, None)  # AAD なし

    with open(dest_path, 'wb') as f:
        f.write(MAGIC)
        f.write(salt)
        f.write(nonce)
        f.write(ciphertext)


def decrypt_file(encrypted_path: str, dest_path: str, password: str) -> None:
    """暗号化ファイルを復号して dest_path に書き出す。

    Args:
        encrypted_path: 暗号化済みファイルパス
        dest_path: 復号結果の出力先パス
        password: 暗号化時と同じパスワード

    Raises:
        DecryptionError: マジックバイト不正、パスワード不正、ファイル破損
    """
    with open(encrypted_path, 'rb') as f:
        magic = f.read(len(MAGIC))
        if magic != MAGIC:
            raise DecryptionError(
                '暗号化ファイルの形式が正しくありません（マジックバイト不一致）'
            )
        salt = f.read(SALT_SIZE)
        nonce = f.read(NONCE_SIZE)
        ciphertext = f.read()

    if len(salt) < SALT_SIZE or len(nonce) < NONCE_SIZE or len(ciphertext) == 0:
        raise DecryptionError('暗号化ファイルが破損しています（データ不足）')

    key = _derive_key(password, salt)
    aesgcm = AESGCM(key)

    try:
        plaintext = aesgcm.decrypt(nonce, ciphertext, None)
    except Exception as exc:
        raise DecryptionError(
            'パスワードが正しくないか、ファイルが破損しています'
        ) from exc

    with open(dest_path, 'wb') as f:
        f.write(plaintext)


# ── ファイルハッシュ ─────────────────────────────────────────────────────────

def compute_file_hash(path: str) -> str:
    """ファイルの SHA-256 ハッシュを返す。"""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


# ── config.json 内パスワード保護（Windows DPAPI） ────────────────────────────

_DPAPI_PREFIX: Final[str] = 'dpapi:'
_PLAIN_PREFIX: Final[str] = 'plain:'


def _is_windows() -> bool:
    return sys.platform == 'win32'


def protect_password(password: str) -> str:
    """パスワードを保護して config.json に保存可能な文字列にする。

    Windows: DPAPI (CryptProtectData) で暗号化。
    非 Windows: base64 エンコード（フォールバック）。
    """
    if not password:
        return ''
    if _is_windows():
        return _dpapi_protect(password)
    return _PLAIN_PREFIX + base64.b64encode(password.encode('utf-8')).decode('ascii')


def unprotect_password(protected: str) -> str:
    """保護されたパスワード文字列を復元する。"""
    if not protected:
        return ''
    if protected.startswith(_DPAPI_PREFIX):
        return _dpapi_unprotect(protected)
    if protected.startswith(_PLAIN_PREFIX):
        return base64.b64decode(protected[len(_PLAIN_PREFIX):]).decode('utf-8')
    # 未保護（マイグレーション用: 旧 config の平文パスワード）
    return protected


# ── DPAPI 実装 ───────────────────────────────────────────────────────────────

class _DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ('cbData', ctypes.wintypes.DWORD),
        ('pbData', ctypes.POINTER(ctypes.c_char)),
    ]


def _dpapi_protect(password: str) -> str:
    """Windows DPAPI CryptProtectData でパスワードを保護する。"""
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32

    data = password.encode('utf-8')
    input_blob = _DATA_BLOB(
        len(data),
        ctypes.cast(ctypes.create_string_buffer(data, len(data)),
                     ctypes.POINTER(ctypes.c_char)),
    )
    output_blob = _DATA_BLOB()

    success = crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        None,   # description
        None,   # optional entropy
        None,   # reserved
        None,   # prompt struct
        0,      # flags
        ctypes.byref(output_blob),
    )
    if not success:
        raise OSError('CryptProtectData failed')

    protected_bytes = ctypes.string_at(output_blob.pbData, output_blob.cbData)
    kernel32.LocalFree(output_blob.pbData)
    return _DPAPI_PREFIX + base64.b64encode(protected_bytes).decode('ascii')


def _dpapi_unprotect(protected: str) -> str:
    """Windows DPAPI CryptUnprotectData でパスワードを復元する。"""
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32

    encrypted_bytes = base64.b64decode(protected[len(_DPAPI_PREFIX):])
    input_blob = _DATA_BLOB(
        len(encrypted_bytes),
        ctypes.cast(ctypes.create_string_buffer(encrypted_bytes, len(encrypted_bytes)),
                     ctypes.POINTER(ctypes.c_char)),
    )
    output_blob = _DATA_BLOB()

    success = crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,   # description out
        None,   # optional entropy
        None,   # reserved
        None,   # prompt struct
        0,      # flags
        ctypes.byref(output_blob),
    )
    if not success:
        raise OSError('CryptUnprotectData failed')

    data = ctypes.string_at(output_blob.pbData, output_blob.cbData)
    kernel32.LocalFree(output_blob.pbData)
    return data.decode('utf-8')
