# -*- mode: python ; coding: utf-8 -*-

import sys
import os
import PySide6
from PyInstaller.utils.hooks import collect_submodules, collect_dynamic_libs

block_cipher = None

conda_plugins = os.path.join(sys.prefix, "Library", "plugins")

numpy_hiddenimports = collect_submodules('numpy')

numpy_binaries = collect_dynamic_libs('numpy')

a = Analysis(
    ['qt6_app.py'],
    pathex=[],
    binaries=[
        (os.path.join(sys.base_prefix, "python313.dll"), "."),
        (os.path.join(sys.base_prefix, "vcruntime140.dll"), "."),
        (os.path.join(sys.base_prefix, "vcruntime140_1.dll"), ".")
    ] + numpy_binaries,
    datas=[
    ],
    hiddenimports=[
        "qtpy",
        "numpy",
        "numpy.core._methods",
        "numpy.lib.format",
        "pandas",
        "openpyxl",
        "win32com",
        "win32com.client",
        "pythoncom",
        "pywintypes",
        "PyPDF2",
        "PySide6",
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
        "PySide6.QtNetwork",
        "PySide6.QtPrintSupport"
    ] + numpy_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='qt6_app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
