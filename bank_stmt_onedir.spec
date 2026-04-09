# -*- mode: python ; coding: utf-8 -*-

import sys, os

block_cipher = None

src_path = os.path.join(os.getcwd(), 'src')
spec_path = os.getcwd()

gui_assets = [
    (os.path.join(src_path, 'gui', 'assets'), os.path.join('gui', 'assets')),
]

a = Analysis(
    [os.path.join(src_path, 'main.py')],
    pathex=[spec_path],
    binaries=[],
    datas=gui_assets,
    hiddenimports=[
        'webview',
        'pymupdf',
        'rapidfuzz',
        'openpyxl',
        'tkinter',
        'xml.etree.ElementTree',
        'PIL',
        'rapidfuzz.distance',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='银行流水核对系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='银行流水核对系统',
)
