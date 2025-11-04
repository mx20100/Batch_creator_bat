# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['converter.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Max Moors\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\pandas', 'pandas'), ('C:\\Users\\Max Moors\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\numpy', 'numpy'), ('C:\\Users\\Max Moors\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\openpyxl', 'openpyxl')],
    hiddenimports=[],
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
    name='Converter',
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
)
