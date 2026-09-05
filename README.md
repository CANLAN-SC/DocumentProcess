# PDF / 图片批量转 Word

一个本地 Web 工具：把 PDF 的每一页渲染为高清图片并插入 Word，也可以把 JPG、JPEG、PNG 图片直接插入 Word。

新版已统一为浏览器操作界面，不再使用 Tkinter GUI；PDF 渲染改用 **PyMuPDF**，因此不再需要安装 Poppler，也不再依赖系统 `PATH`。

## 功能

- 批量上传 PDF、JPG、JPEG、PNG
- A4 Word 版式实时预览，可切换文件并查看 PDF 前 6 页
- 批量调整全部图片大小（页面可用区域的 10%～200%）
- 分别设置 Word 上、下、左、右页边距
- 高清和压缩两种输出质量
- 多文件分别导出为 ZIP，或合并为一个 Word 文档
- 所有文件只在运行该程序的本机处理，临时结果 24 小时后自动清理

## 安装

建议使用 Python 3.10 或更高版本：

```bash
python -m pip install -r requirements.txt
```

## 启动 Web 界面

```bash
python app.py
```

程序默认启动在：

```text
http://127.0.0.1:8000
```

启动后会自动打开浏览器。服务器环境或不希望自动打开浏览器时：

```bash
python app.py --no-browser
```

旧入口 `python PDF_Converter_GUI.py` 仍可使用，但现在同样启动 Web 界面，不会再创建桌面 GUI。

## 使用流程

1. 拖放或选择一个或多个 PDF/图片文件。
2. 在右侧查看 Word 页面预览。
3. 调整图片比例、页边距、输出质量和合并选项。
4. 点击“开始转换”，完成后直接下载 DOCX 或 ZIP。

## 项目结构

```text
app.py                         # 推荐启动入口
PDF_Converter_GUI.py           # 兼容旧入口，转发到 Web 应用
requirements.txt               # Python 依赖
web/
├── app.py                     # Flask 路由、上传和下载
├── core/converter.py          # PyMuPDF 渲染与 Word 生成
├── templates/index.html       # 统一 Web 页面
└── static/
    ├── css/app.css
    └── js/app.js
tests/test_pdf_converter.py    # 转换和 Web API 测试
```

## 打包为独立程序（Windows）

在已安装依赖和 PyInstaller 的 Python 环境中，使用项目根目录下的 `build.spec` 一键生成不依赖 Python 的单文件程序：

```bash
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --clean --noconfirm build.spec
```

生成的程序位于 `dist/PdfToWord.exe`。把 `PdfToWord.exe` 单独拷贝给未安装 Python 的同事即可使用：

- 双击运行，程序会在本机 `http://127.0.0.1:8000` 提供服务并自动打开浏览器；
- 控制台窗口会显示访问地址，直接关闭窗口或按 `Ctrl+C` 即可退出；
- 命令行参数与源码版一致，例如 `PdfToWord.exe --no-browser` 可关闭自动打开浏览器。

注意事项：

- 单文件程序首次启动时需要把运行文件解压到临时目录，启动会比源码版慢（通常数秒到数十秒），请耐心等待；
- 程序体积约 90～140MB，属于正常现象（内含 Python 解释器和 PyMuPDF 等运行库）；
- 个别杀毒软件可能对无签名的 PyInstaller 单文件程序误报，如遇拦截请加入信任列表或改用源码方式运行。

## 测试

```bash
python -m unittest discover -s tests -v
```

> 工具生成的是图像型 Word，不会对扫描内容执行 OCR 或文字提取。
