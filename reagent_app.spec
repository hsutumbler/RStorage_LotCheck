# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['simple_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),  # 包含HTML模板
    ],
    hiddenimports=[
        'win32print',
        'win32api',
        'win32gui',
        'pywintypes',
        'engineio.async_drivers.threading',
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='試劑管理系統',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 顯示控制台視窗以便查看啟動訊息
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加圖標檔案路徑
)
