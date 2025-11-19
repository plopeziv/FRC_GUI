# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

block_cipher = None

base_dir = Path(SPECPATH).parent

streamlit_datas, streamlit_binaries, streamlit_hiddenimports = collect_all('streamlit')
blinker_datas, blinker_binaries, blinker_hiddenimports = collect_all('blinker')

a = Analysis(
    ['launcher.py'],
    pathex=[str(base_dir)],
    binaries=streamlit_binaries + blinker_binaries,
    datas=[
        ('app.py', '.'),
        ('data_manager', 'data_manager'),
        ('frontend_components', 'frontend_components'),
    ] + streamlit_datas + blinker_datas,
    hiddenimports=[
        'streamlit',
        'streamlit.web.cli',
        'streamlit.runtime',
        'streamlit.runtime.scriptrunner',
        'streamlit.runtime.state',
        'importlib_metadata',
        'pandas',
        'openpyxl',
        'xlrd',
        'altair',
        'pyarrow',
    ] + streamlit_hiddenimports + blinker_hiddenimports,
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
    name='FRC_Ticket_GUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Add this if you want a single executable file:
    # onefile=True,
)