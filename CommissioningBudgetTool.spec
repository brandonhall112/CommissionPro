# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
block_cipher = None
APP_DIR = Path(SPECPATH)

a = Analysis(
    ["app.py"],
    pathex=[str(APP_DIR)],
    binaries=[],
    datas=[(str(APP_DIR / "assets"), "assets")],
    hiddenimports=["PySide6.QtSvg", "PySide6.QtXml", "PySide6.QtPrintSupport"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name="CommissioningBudgetTool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
