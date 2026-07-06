# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

# openpyxl ships style/template data files that must be bundled
openpyxl_datas = collect_data_files('openpyxl')

a = Analysis(
    ['process_cbt.py'],
    pathex=[],
    binaries=[],
    datas=openpyxl_datas,
    hiddenimports=[
        'openpyxl.cell._writer',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.styles.alignment',
        'openpyxl.styles.borders',
        'cbt.normalize',
        'cbt.loader',
        'cbt.matcher',
        'cbt.reporter',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='cbt_report',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,          # keep console window — this is a CLI tool
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    icon=None,
)
