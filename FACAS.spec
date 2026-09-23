# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_all


# Playwright 的 Python 运行时和 openpyxl 的资源需要显式收集。
playwright_datas, playwright_binaries, playwright_hiddenimports = collect_all("playwright")
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all("openpyxl")
xlrd_datas, xlrd_binaries, xlrd_hiddenimports = collect_all("xlrd")


datas = [*playwright_datas, *openpyxl_datas, *xlrd_datas]
binaries = [*playwright_binaries, *openpyxl_binaries, *xlrd_binaries]
hiddenimports = [
    *playwright_hiddenimports,
    *openpyxl_hiddenimports,
    *xlrd_hiddenimports,
    "bank_entry",
]

# 程序只使用 QtCore、QtGui、QtWidgets；其余 Qt 模块由官方 hook 自动收集，
# 会显著增加体积，但不会参与本项目的界面、Excel 或 PDF 流程。
_unused_qt_modules = [
    "PySide6.QtNetwork",
    "PySide6.QtOpenGL",
    "PySide6.QtPdf",
    "PySide6.QtQml",
    "PySide6.QtQmlMeta",
    "PySide6.QtQmlModels",
    "PySide6.QtQmlWorkerScript",
    "PySide6.QtQuick",
    "PySide6.QtSvg",
    "PySide6.QtVirtualKeyboard",
]


a = Analysis(
    ["app.py"],
    pathex=[str(Path(SPEC).parent)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["icuuc", "icudt78", "PIL", *_unused_qt_modules],
    noarchive=False,
    optimize=0,
)

# Poppler/其他运行库可能把 ICU 78 放到打包根目录，文件名却是 Qt 使用的
# 通用名称。Qt6Core 加载到该 DLL 后会因导出符号不匹配而导入失败，必须
# 从最终二进制清单中再次过滤，避免自动分析重新带入。
_bad_icu_names = {"icuuc.dll", "icudt78.dll"}
_unused_qt_binaries = {
    "qt6network.dll",
    "qtnetwork.pyd",
    "qt6opengl.dll",
    "qtopengl.pyd",
    "qt6pdf.dll",
    "qtpdf.pyd",
    "qt6qml.dll",
    "qtqml.pyd",
    "qt6qmlmeta.dll",
    "qtqmlmeta.pyd",
    "qt6qmlmodels.dll",
    "qtqmlmodels.pyd",
    "qt6qmlworkerscript.dll",
    "qtqmlworkerscript.pyd",
    "qt6quick.dll",
    "qtquick.pyd",
    "qt6svg.dll",
    "qtsvg.pyd",
    "qt6virtualkeyboard.dll",
    "qtvirtualkeyboard.pyd",
}


def _keep_binary(entry):
    destination = entry[0].replace("/", "\\")
    normalized = destination.lower()
    name = Path(destination).name.lower()
    if name in _bad_icu_names or name in _unused_qt_binaries:
        return False
    if normalized.startswith("pil\\"):
        return False
    if normalized.startswith("pyside6\\translations\\"):
        return False
    if normalized.startswith("pyside6\\plugins\\"):
        # 桌面版只需要 Windows 平台插件；保留 offscreen 便于自动化测试。
        return normalized.endswith(("\\platforms\\qwindows.dll", "\\platforms\\qoffscreen.dll"))
    return True


a.binaries = [
    entry for entry in a.binaries
    if _keep_binary(entry)
]


def _keep_data(entry):
    destination = entry[0].replace("/", "\\").lower()
    return not (
        destination.startswith("pyside6\\translations\\")
        or destination.startswith("pyside6\\qml\\")
        or destination.startswith("pil\\")
    )


a.datas = [entry for entry in a.datas if _keep_data(entry)]

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
