# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['C:/Users/dacan/OneDrive/Desktop/Projects/TRAFFIC_REPORT'],
    binaries=[],
    datas=[("templates", "templates"), ("data", "data"), ('C:/Users/dacan/OneDrive/Desktop/Projects/TRAFFIC_REPORT/.venv/Lib/site-packages/docxcompose/templates', 'docxcompose/templates')],
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
    name='Traffic_Report v2.0.7',
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
    icon=['ui\\icono.ico'],
)
