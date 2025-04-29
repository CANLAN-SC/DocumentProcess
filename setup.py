from cx_Freeze import setup, Executable
import sys
import os

# 包含poppler文件夹
poppler_dir = "poppler"
include_files = [
    (os.path.join("poppler", "Library", "bin"), os.path.join("lib", "poppler", "Library", "bin"))
]
# 基础配置
build_options = {
    "packages": ["os", "sys", "tkinter", "PIL", "docx", "pdf2image", "docxcompose"],
    "excludes": ["tkinter.test"],
    "include_files": include_files,
    "optimize": 2
}

# 隐藏控制台窗口
base = "Win32GUI" if sys.platform == "win32" else None

executables = [
    Executable(
        "PDF_Converter_GUI.py",
        base=base,
        target_name="PDFConverter.exe",
    )
]

setup(
    name="PDF Converter",
    version="2.0",
    description="PDF/图片转Word工具",
    options={"build_exe": build_options},
    executables=executables
)