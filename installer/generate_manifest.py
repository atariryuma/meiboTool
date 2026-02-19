"""ビルド成果物のファイルハッシュマニフェストを生成する。

差分アップデート用。CI で PyInstaller ビルド後に実行する。

Usage:
    python installer/generate_manifest.py dist/名簿帳票ツール v1.0.3
"""

import hashlib
import json
import os
import sys
from datetime import datetime, timezone


def generate_manifest(build_dir: str, version: str) -> dict:
    """build_dir 内の全ファイルの SHA-256 ハッシュを記録したマニフェストを返す。"""
    manifest: dict = {
        'version': version,
        'build_time': datetime.now(timezone.utc).isoformat(timespec='seconds'),
        'files': {},
    }
    for root, _dirs, files in os.walk(build_dir):
        for fname in files:
            full = os.path.join(root, fname)
            rel = os.path.relpath(full, build_dir).replace('\\', '/')
            with open(full, 'rb') as f:
                sha = hashlib.sha256(f.read()).hexdigest()
            manifest['files'][rel] = {
                'sha256': sha,
                'size': os.path.getsize(full),
            }
    return manifest


def main() -> None:
    if len(sys.argv) < 3:
        print(f'Usage: {sys.argv[0]} <build_dir> <version>')
        sys.exit(1)

    build_dir = sys.argv[1]
    version = sys.argv[2]

    if not os.path.isdir(build_dir):
        print(f'Error: {build_dir} is not a directory')
        sys.exit(1)

    m = generate_manifest(build_dir, version)
    out_path = os.path.join(os.path.dirname(build_dir), 'manifest.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(m, f, ensure_ascii=False, indent=2)

    print(f'Manifest written: {out_path} ({len(m["files"])} files)')


if __name__ == '__main__':
    main()
