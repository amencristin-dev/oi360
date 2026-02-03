# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['oi_360_soa_reco_pyqt_final.py'],
    pathex=[],
    binaries=[],
    datas=[('Oi360 Logo_4.png', '.')],
    hiddenimports=['PyQt5', 'pandas', 'openpyxl', 'xlrd', 'xlsxwriter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Oi360_SOA_RECO',
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
)
