# -*- mode: python ; coding: utf-8 -*-
import os, sys
block_cipher = None

spec_dir = os.path.dirname(os.path.abspath(SPEC))
src_dir = os.path.join(spec_dir, 'src')

gui_assets = [
    (os.path.join(src_dir, 'gui', 'assets'), 'gui', 'assets'),
]

a = Analysis(
    [os.path.join(src_dir, 'main.py')],
    pathex=[spec_dir],
    binaries=[],
    datas=gui_assets,
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
    exclude_binaries=True, name='银行流水核对系统',
    debug=False, bootloader_ignore_signals=False, strip=False, upx=True,
    console=False, disable_windowed_traceback=False, argv_emulation=False,
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, name='银行流水核对系统',
)
