# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

import sys
import os
from pathlib import Path

# Get tkinterdnd2 package location
tkdnd_path = None
for path in sys.path:
    potential_path = Path(path) / 'tkinterdnd2'
    if potential_path.exists():
        tkdnd_path = potential_path
        break

# Prepare datas for tkinterdnd2
datas = []
if tkdnd_path:
    # Include tkdnd directory with platform-specific libraries
    tkdnd_dir = tkdnd_path / 'tkdnd'
    if tkdnd_dir.exists():
        datas.append((str(tkdnd_dir), 'tkinterdnd2/tkdnd'))

a = Analysis(
    ['excel_cleaner.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['openpyxl', 'pandas', 'tkinter', 'tkinterdnd2'],
    hookspath=[],
    hooksconfig={},
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelCleaner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False to hide console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # You can add an .ico file path here if desired
)
