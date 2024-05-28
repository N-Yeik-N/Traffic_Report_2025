Estoy utilizando el módulo de PDF2image
Requiere descargarlo y colocar la carpeta bin/ como Path del sistema.

Para compilar el programa deberás sí o sí utilizar pyinstaller --name 'Nombre' .main.py
Luego ver el archivo Nombre.spec
Hacer algunos cambios para que acabe así:

# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['C:/Users/dacan/OneDrive/Desktop/Projects/TRAFFIC_REPORT'],
    binaries=[],
    datas=[
        ('C:/Users/dacan/OneDrive/Desktop/Projects/TRAFFIC_REPORT/.venv/Lib/site-packages/docxcompose/templates', 'docxcompose/templates')
        ],
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
    [],
    exclude_binaries=True,
    name='REPORTER',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='C:/Users/dacan/OneDrive/Desktop/Projects/TRAFFIC_REPORT/ui/icono.ico'
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='REPORTER',
)

Con eso debería bastar para que pueda considerar los templates del docxcompose.