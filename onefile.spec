# -*- mode: python ; coding: utf-8 -*-
import os, sys
block_cipher = None

# Get repo root from environment variable set by workflow, fallback to spec dir
repo_root = os.environ.get('REPO_ROOT', os.path.dirname(os.path.abspath(__file__)))
src_dir = os.path.join(repo_root, 'src')

a = Analysis(
    [os.path.join(src_dir, 'main.py')],
    pathex=[repo_root],
    binaries=[],
    datas=[
        (os.path.join(src_dir, 'gui', 'assets'), 'gui', 'assets'),
        (src_dir, 'src'),
    ],
    hiddenimports=[
        'pywebview', 'pymupdf', 'rapidfuzz', 'openpyxl',
        'tkinter', 'xml.etree.ElementTree', 'PIL', 'rapidfuzz.distance',
    ],
    hookspath=[], runtime_hooks=[], excludes=[],
    win_no_prefer_redirects=False, win_private_assemblies=False,
    cipher=block_cipher, noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=False,
    name='银行流水核对系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
)
