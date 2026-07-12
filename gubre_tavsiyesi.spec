# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gubre_tavsiye.py'],
    pathex=[],
    binaries=[],
    datas=[('app.ico', '.'), ('app.png', '.'), ('translations', 'translations'), ('C:\\Windows\\Fonts\\tahoma.ttf', '.'), ('C:\\Windows\\Fonts\\tahomabd.ttf', '.')],
    hiddenimports=['openpyxl', 'matplotlib', 'geopy'],
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
    name='gubre_tavsiyesi',
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
    icon=['app.ico'],
)
