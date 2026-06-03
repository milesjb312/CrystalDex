# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas,binaries,hiddenimports = collect_all('openpyxl')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas+[('README.md','.'),
    ('Resources/Crystal_Pictures','Crystal_Pictures'),
    ('Resources/Crystal_Screens.json','Resources'),
    ('Resources/crystaldex_icon.png','Resources'),
    ('Resources/crystaldex_icon.ico','Resources'),
    ('Resources/CrystalDex_Splash.png','Resources'),
    ('CS_FS.xlsx','.'),
    ('Index_FS_2021.xlsx','.'),
    ('PEGIon_FS.xlsx','.'),
    ('SaltRx_FS.xlsx','.'),
    ('Wizard_Screen_Formulation_&_Scoring.xlsx','.')],
        hiddenimports=hiddenimports+['pdfplumber',
    'pdfminer',
    'pdfminer.high_level',
    'pdfminer.layout',
    'pdfminer.converter',
    'pywinauto',
    'pyautogui',
    'pymsgbox',
    'pygetwindow',
    'pyrect',
    'pytweening',
    'mouseinfo',
    'pynput',
    'pynput.keyboard',
    'pynput.mouse',
    'openpyxl',
    'box_sdk_gen'],
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
    name='CrystalDex',
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
    icon='Resources/crystaldex_icon.ico'
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CrystalDex',
    icon='Resources/crystaldex_icon.ico'
)
