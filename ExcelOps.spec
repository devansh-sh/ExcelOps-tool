# -*- mode: python ; coding: utf-8 -*-
import os
import sys


project_root = os.path.abspath(SPECPATH)
datas = []
assets_dir = os.path.join(project_root, "assets")
if os.path.isdir(assets_dir):
    datas.append((assets_dir, "assets"))
presets_dir = os.path.join(project_root, "presets")
if os.path.isdir(presets_dir):
    datas.append((presets_dir, "presets"))

icon = None
if sys.platform == "darwin":
    candidate = os.path.join(project_root, "assets", "excelops.icns")
    icon = candidate if os.path.exists(candidate) else None
elif os.name == "nt":
    candidate = os.path.join(project_root, "assets", "excelops.ico")
    icon = candidate if os.path.exists(candidate) else None

a = Analysis(
    [os.path.join(project_root, "main.py")],
    pathex=[project_root],
    binaries=[],
    datas=datas,
    hiddenimports=["openpyxl", "xlrd"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ExcelOps",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=icon,
)

collection = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="ExcelOps",
)

if sys.platform == "darwin":
    app = BUNDLE(
        collection,
        name="ExcelOps.app",
        icon=icon,
        bundle_identifier="com.bds.excelops",
        info_plist={
            "CFBundleDisplayName": "ExcelOps",
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
