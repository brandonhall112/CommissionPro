# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.building.build_main import Analysis, PYZ, EXE

block_cipher = None

PROJECT_DIR = os.path.abspath(os.getcwd())
ASSETS_DIR = os.path.join(PROJECT_DIR, "assets")

APP_NAME = "PearsonCommissioningPro"
ICON_PATH = os.path.join(ASSETS_DIR, "PearsonP.ico")

# Build datas as a list of (src, dest_folder) 2-tuples (what Analysis expects)
datas = []
if os.path.isdir(ASSETS_DIR):
    for root, dirs, files in os.walk(ASSETS_DIR):
        for fn in files:
            src = os.path.join(root, fn)
            rel = os.path.relpath(root, ASSETS_DIR)  # subfolder under assets
            dest = os.path.join("assets", rel) if rel != "." else "assets"
            datas.append((src, dest))

a = Analysis(
    ["app.py"],
    pathex=[PROJECT_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=["PySide6.QtSvg", "PySide6.QtXml"],
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

# Only apply icon if present (prevents CI failure)
icon_arg = ICON_PATH if os.path.exists(ICON_PATH) else None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    icon=icon_arg,
)
