# -*- mode: python ; coding: utf-8 -*-
# md2word.spec — PyInstaller 打包配置（GUI 版）

a = Analysis(
    ['launcher_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app_gui.py', '.'),
        ('pandoc/pandoc.exe', 'pandoc'),
    ],
    hiddenimports=[
        'pypandoc',
        'docx',
        'docx.shared',
        'docx.oxml',
        'docx.oxml.ns',
        'customtkinter',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'streamlit',
        'altair',
        'pydeck',
        'blinker',
        'gitpython',
        'smmap',
        'gitdb',
        'watchdog',
    ],
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
    name='MD2Word_GUI',
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
    icon=None,
)
