# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

pyarrow_datas, pyarrow_binaries, pyarrow_hidden = collect_all("pyarrow")
calamine_datas, calamine_binaries, calamine_hidden = collect_all("python_calamine")

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=pyarrow_binaries + calamine_binaries,
    datas=[
        ('conversores/config/configuracion.json', 'conversores/config'),
        ('assets/metlife_logo.jpg', 'assets'),
        ('icono.ico', '.'),
    ] + pyarrow_datas + calamine_datas,
    hiddenimports=[
        'python_calamine',
        'openpyxl',
        'xlsxwriter',
        'pyarrow',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
    ] + pyarrow_hidden + calamine_hidden,
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
    name='CommiFlow',
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
    icon=['icono.ico'],
)
