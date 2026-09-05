# 交付概览 · 纸页（PDF / 图片转 Word）三项功能完善

## 一句话总结
为本地 Web 工具「纸页」完成三项增强：文件列表字典序排序、图片预览放大上限放宽到 200%、打包为免 Python 的单文件 exe，均经独立 QA 验证通过。

## 交付状态
- 三项需求全部实现，单元测试 11/11 通过，冻结环境端到端实测通过。
- 已知问题：0（无源码 Bug、无测试 Bug）。

## 改动明细

### 需求一：文件列表字典序升序
- `web/static/js/app.js`：`addFiles()` 内按 `file.name` 排序（`localeCompare` + 自然数字排序 + 大小写不敏感），排序后按 `fileKey` 重定位选中项，避免追加文件后选中错位。

### 需求二：图片预览放大到 200%
- `web/templates/index.html`：range `max` 100 → 200，刻度改为「小 / 适应页面 / 放大」。
- `web/static/js/app.js`：`validateSettings()` 上限 100 → 200。
- `web/core/converter.py`：`LayoutConfig.validate()` 上限 100 → 200；`render_preview` 默认 dpi 110 → 150（200% 放大保持清晰）。
- `web/static/css/app.css`：图片用显式 `width/height: var(--image-scale)` + `object-fit: contain` + `flex-shrink: 0`；`.word-content` overflow 由 hidden → visible，>100% 时直观展示“超出页面”。

### 需求三：打包独立可执行程序
- 新增 `build.spec`（PyInstaller onefile，打包 `web/templates`、`web/static`）。
- `web/app.py`：新增 `_is_frozen()` / `_resource_root()`，冻结时从 `sys._MEIPASS` 加载模板静态、数据目录改到系统临时目录；非冻结行为完全不变。
- 产物：`dist/PdfToWord.exe`（约 52 MB）。

## 使用方式
- 源码运行：`python app.py`（需 Python 3.10+ 与 requirements.txt）。
- 免环境运行：双击 `dist/PdfToWord.exe`，自动打开浏览器访问 `http://127.0.0.1:8000`。

## 注意事项
- 单文件 exe 首次启动需解压到临时目录，启动稍慢；部分杀软可能误报，需加入信任。
- 重新打包：安装 PyInstaller 后执行 `pyinstaller --clean --noconfirm build.spec`。

## 后续建议
1. 若需 200% 以上（如 300%）或自由缩放，改动 `index.html`/`app.js`/`converter.py` 三处上限值即可。
2. 可考虑为 exe 增加托盘图标/自动关停，进一步优化免 Python 用户体验。
3. `build/` 为构建中间产物，可随时删除，不影响 exe。
