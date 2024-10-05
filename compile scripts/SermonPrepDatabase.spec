# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['..\\src\\sermon_prep_database.py'],
    pathex=[],
    binaries=[],
    datas=[('../gpl-3.0.rtf', './'), ('../src/resources', 'resources/'), ('../src', 'src'), ('../src/ghostscript', './ghostscript/'), ('../README.html', './'), ('../README.md', './')],
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
    name='SermonPrepDatabase',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['..\\src\\resources\\icons.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SermonPrepDatabase',
)
