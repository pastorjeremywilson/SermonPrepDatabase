# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src/SermonPrepDatabase.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/gpl-3.0.rtf', '.'), ('C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src/ghostscript', 'ghostscript/'), ('C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/resources', 'resources/'), ('C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src', 'src/')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

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
    icon=['C:\\Users\\pasto\\Nextcloud\\Documents\\Python Workspace\\Sermon Prep Database\\resources\\icons.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SermonPrepDatabase',
)
