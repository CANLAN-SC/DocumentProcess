"""兼容旧的高清批处理入口；PDF 渲染已改用 PyMuPDF。"""

from pathlib import Path

from web.core.converter import SUPPORTED_EXTENSIONS, WordConverter


def run_batch(compress=False):
    input_folder = Path("待处理文件")
    if not input_folder.is_dir():
        raise SystemExit(f"输入文件夹不存在：{input_folder.resolve()}")
    output_folder = input_folder / f"{input_folder.name}_docx"
    output_folder.mkdir(parents=True, exist_ok=True)
    converter = WordConverter(compress=compress)
    files = sorted(
        (path for path in input_folder.iterdir() if path.suffix.lower() in SUPPORTED_EXTENSIONS),
        key=lambda path: path.name.lower(),
    )
    if not files:
        raise SystemExit("输入文件夹中没有可转换的 PDF 或图片")
    for path in files:
        output_path = output_folder / f"{path.stem}.docx"
        converter.convert(path, output_path)
        print(f"转换成功：{path.name}")
    print(f"全部完成，输出目录：{output_folder.resolve()}")


if __name__ == "__main__":
    run_batch()
