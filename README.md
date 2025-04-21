# 📄 PDF 批量转换为 Word（含图片）工具

这是一个 **自动将多个 PDF 或图片文件转换为 Word 文件（.docx）并合并输出** 的小工具。

📌 它能将 PDF 中的每一页变成高清图片插入 Word，同时也支持 PNG/JPG 图片直接插入 Word。所有生成的单个 Word 文件可自动合并成一个完整文档（保留图片）。

适合以下场景：

- 批量处理扫描件（PDF）或图片为 Word 文档格式
- 汇总图片、扫描图表等内容形成报告
- 不懂复杂排版，只需要“图文版” Word 汇总文件

------

## ✨ 功能特点

- 📂 自动处理指定文件夹下的所有 PDF / 图片（JPG、PNG）文件
- 🖼 PDF 每页转为高清图片插入 Word，图片直接插入 Word
- 📏 自动缩放图片，适配 A4 页面（不会超出页面）
- 📄 每个文件生成一个对应的 Word 文件
- 📚 可自动合并所有生成的 Word 文件，保留图片不丢失

------

## 📁 文件结构说明

```
项目目录/
┌── 2word.py                      # 转word脚本
├── mergeWord.py                  # 合并word脚本
└── 待处理文件夹/                  # 改成你自己的文件夹名字
    ├── *.pdf / *.jpg / *.png
    └── 待处理文件夹docx/          # 自动生成的 Word 文件存放目录
    |   └── *.docx                # 每个 PDF 对应一个 Word 文件
    └── 待处理文件夹_合并.docx     # 合并生成的总 Word 文档
```

------

## ✅ 环境要求
建议使用VSCode或者Pycharm运行，并安装Python环境。
确保终端运行安装以下 Python 库：

```bash
pip install pdf2image

pip install pymupdf python-docx pillow

pip install docxcompose
```

- Windows 用户需要解压缩`poppler.zip`压缩包
- Linux 用户：
终端运行
```bash
sudo apt-get install poppler-utils
```
并且注释`2word.py` 的 `poppler_path` 变量，以及将`pages = convert_from_path(file_path, dpi=200, poppler_path=poppler_path)`中的`, poppler_path=poppler_path`删除。

---

## 🛠 使用步骤（超简单！）

### 第一步：准备

1. 将所有 PDF 和图片文件放入一个文件夹，例如 `专利/`
2. 修改脚本中 `input_folder = '专利'`

### 第二步：转换

运行 `2word.py`，将 PDF / 图片 转换为多个 Word 文件：


生成的 Word 文件保存在 `专利/专利_docx/` 目录中。

### 第三步：合并 Word 文件

运行 `mergeWord.py`，合并所有 Word 文件为一个大文件：

最终合并结果为： `专利/专利_合并.docx`

------

## 📌 注意事项

- 本工具不会提取文字，只适用于图像内容转 Word
- 图片经过压缩与缩放，确保清晰且适配 A4 页面
- 若 PDF 中存在签章/注释报错，可通过 `pdf2image` 替代 `fitz`（已在脚本中处理）

------

## ❤️ 作者的话

这个工具是为了批量整理扫描资料和图片归档而写的，根据老婆大人的要求进行了多次迭代，并最简化安装步骤，希望能帮到同样有这种需求的小伙伴。如果你有更好的建议，欢迎提交 PR 或 issue！
