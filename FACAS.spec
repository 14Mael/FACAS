# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_all


# Playwright 的 Python 运行时和 openpyxl 的资源需要显式收集。
playwright_datas, playwright_binaries, playwright_hiddenimports = collect_all("playwright")
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all("openpyxl")


datas = [*playwright_datas, *openpyxl_datas]
binaries = [*playwright_binaries, *openpyxl_binaries]
hiddenimports = [*playwright_hiddenimports, *openpyxl_hiddenimports]


a = Analysis(
    ["app.py"],
    pathex=[str(Path(SPEC).parent)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["icuuc", "icudt78"],
    noarchive=False,
    optimize=0,
)

# Poppler/其他运行库可能把 ICU 78 放到打包根目录，文件名却是 Qt 使用的
# 通用名称。Qt6Core 加载到该 DLL 后会因导出符号不匹配而导入失败，必须
# 从最终二进制清单中再次过滤，避免自动分析重新带入。
_bad_icu_names = {"icuuc.dll", "icudt78.dll"}
a.binaries = [
    entry for entry in a.binaries
    if Path(entry[0]).name.lower() not in _bad_icu_names
]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="FACAS",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
