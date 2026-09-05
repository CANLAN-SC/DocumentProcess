# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：生成不依赖 Python 的单文件 Windows 程序 PdfToWord.exe。"""

from PyInstaller.utils.hooks import collect_all

# 模板与静态资源一并打包，运行时从 sys._MEIPASS 解压目录加载。
datas = [
    ("web/templates", "web/templates"),
    ("web/static", "web/static"),
]
binaries = []
hiddenimports = []

# 收集 PyMuPDF / python-docx / docxcompose 等依赖的隐藏导入与数据文件。
for package in ("pymupdf", "docx", "docxcompose"):
    try:
        pkg_datas, pkg_binaries, pkg_hidden = collect_all(package)
        datas += pkg_datas
        binaries += pkg_binaries
        hiddenimports += pkg_hidden
    except Exception:
        # 某个包不存在时不影响整体打包。
        continue

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

# onefile 模式：EXE 直接包含 binaries 与 datas，不生成 COLLECT 目录。
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="PdfToWord",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
