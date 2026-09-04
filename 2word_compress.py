"""兼容旧的压缩批处理入口；PDF 渲染已改用 PyMuPDF。"""

import runpy
from pathlib import Path


if __name__ == "__main__":
    namespace = runpy.run_path(str(Path(__file__).with_name("2word.py")))
    namespace["run_batch"](compress=True)
