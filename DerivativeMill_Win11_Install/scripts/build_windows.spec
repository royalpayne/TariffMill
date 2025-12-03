# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for DerivativeMill Windows 11 installer
# This file configures how PyInstaller builds the executable

import sys
import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

a = Analysis(
    [os.path.join(project_root, 'DerivativeMill', 'derivativemill.py').replace('\\', '/')],
    pathex=[project_root],
    binaries=[],
    datas=[
        (os.path.join(project_root, 'DerivativeMill', 'Resources'), 'DerivativeMill/Resources'),
        (os.path.join(project_root, 'README.md'), '.'),
    ],
    hiddenimports=[
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'pandas',
        'openpyxl',
        'pdfplumber',
        'PIL',
    ] + collect_submodules('PyQt5'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[],
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
    name='DerivativeMill',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(project_root, 'DerivativeMill', 'Resources', 'derivativemill.ico').replace('\\', '/') if sys.platform == 'win32' else None,
)

# Uncomment for one-dir bundle (development)
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     strip=False,
#     upx=True,
#     upx_exclude=[],
#     name='DerivativeMill'
# )
