# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

PROJECT_ROOT = Path(globals().get("SPECPATH", ".")).resolve()

hiddenimports = ["pythoncom", "pywintypes", "win32timezone"]
hiddenimports += collect_submodules("pptx")
hiddenimports += collect_submodules("PIL")
hiddenimports += collect_submodules("win32com")
hiddenimports += collect_submodules("tkinterdnd2")
datas = collect_data_files("tkinterdnd2")


a = Analysis(
    ["run_gui.py"],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="PPTTemplateReplacer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
