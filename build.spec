# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build spec — 名簿帳票ツール

ビルドコマンド:
    pyinstaller build.spec

出力:
    dist/名簿帳票ツール/

注意:
    --onedir 必須。CustomTkinter の .json/.otf データファイルが
    --onefile では見つからずクラッシュする。
"""

from PyInstaller.utils.hooks import collect_data_files

# CustomTkinter のテーマ定義 (.json) + フォント (.otf) を収集
datas = collect_data_files('customtkinter')

# テンプレート Excel 群と config.json を同梱
datas += [('テンプレート', 'テンプレート')]
datas += [('config.json', '.')]

a = Analysis(
    ['meibo_tool/main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['openpyxl', 'customtkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='名簿帳票ツール',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='名簿帳票ツール',
)
