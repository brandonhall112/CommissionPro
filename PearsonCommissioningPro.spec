# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

block_cipher = None

# Avoid using __file__ in spec files (can be undefined in some runners)
PROJECT_DIR = os.path.abspath(os.getcwd())
ASSETS_DIR = os.path.join(PROJECT_DIR, "assets")

APP_NAME = "PearsonCommissioningPro"
ICON_PATH = os.path.join(ASSETS_DIR, "icon.ico")  # <-- your converted .ico goes here

# Bundle runtime files
datas = []

# Excel input file (the app can load it from the packaged folder)
excel_path = os.path.join(PROJECT_DIR, "Tech days and quote rates.xlsx")
if os.path.exists(excel_path):
    datas.append((excel_path, "."))

# Logo used in printable quote header (if your app references it)
logo_path = os.path.join(ASSETS_DIR, "Pearson Logo.png")
if os.path.exists(logo_path):
    datas.append((logo_path, "assets"))

a = Analysis(
    ["app.py"],
    pathex=[PROJECT_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "PySide6.QtSvg",
        "PySide6.QtXml",
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
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                   # GUI
    disable_windowed_traceback=False,
    icon=ICON_PATH,                  # <-- EXE icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=APP_NAME,
)
