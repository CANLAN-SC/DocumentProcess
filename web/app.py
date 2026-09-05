import argparse
import re
import shutil
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
import zipfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_from_directory, url_for

from web.core.converter import LayoutConfig, SUPPORTED_EXTENSIONS, WordConverter, render_preview


def _is_frozen():
    """判断是否运行在 PyInstaller 打包后的冻结环境中。"""
    return getattr(sys, "frozen", False)


def _resource_root():
    """返回模板与静态资源所在的根目录。

    冻结打包时资源被解压到临时目录 sys._MEIPASS，否则使用项目根目录。
    """
    if _is_frozen():
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


PROJECT_ROOT = _resource_root()
if _is_frozen():
    DEFAULT_DATA_DIR = Path(tempfile.gettempdir()) / "pdf2word" / "data"
else:
    DEFAULT_DATA_DIR = Path(__file__).resolve().parent / "data"
JOB_ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")
INVALID_FILENAME_CHARS = re.compile(r"[<>:\"/\\|?*\x00-\x1f]")
MAX_FILES = 50
JOB_LIFETIME_SECONDS = 24 * 60 * 60


def create_app(test_config=None):
    resource_root = _resource_root()
    app = Flask(
        __name__,
        template_folder=str(resource_root / "web" / "templates"),
        static_folder=str(resource_root / "web" / "static"),
    )
    app.config.update(
        MAX_CONTENT_LENGTH=500 * 1024 * 1024,
        DATA_DIR=str(DEFAULT_DATA_DIR),
        JSON_AS_ASCII=False,
    )
    if test_config:
        app.config.update(test_config)
    Path(app.config["DATA_DIR"]).mkdir(parents=True, exist_ok=True)

    @app.after_request
    def add_security_headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.errorhandler(413)
    def upload_too_large(_error):
        return jsonify(error="上传内容超过 500 MB 限制"), 413

    @app.get("/")
    def index():
        cleanup_expired_jobs(Path(app.config["DATA_DIR"]))
        return render_template("index.html")

    @app.post("/api/preview")
    def preview():
        uploaded = request.files.get("file")
        if uploaded is None or not uploaded.filename:
            return jsonify(error="请选择需要预览的文件"), 400
        try:
            filename = safe_upload_name(uploaded.filename)
            ensure_supported(filename)
            job_id, job_folder = create_job(Path(app.config["DATA_DIR"]))
            upload_folder = job_folder / "uploads"
            preview_folder = job_folder / "preview"
            input_path = upload_folder / filename
            uploaded.save(input_path)
            pages, total_pages = render_preview(input_path, preview_folder)
            for page in pages:
                page["url"] = url_for(
                    "preview_asset", job_id=job_id, filename=page["filename"]
                )
            return jsonify(
                job_id=job_id,
                filename=filename,
                title=Path(filename).stem,
                pages=pages,
                total_pages=total_pages,
                preview_limit=len(pages),
            )
        except (ValueError, RuntimeError, OSError) as exc:
            return jsonify(error=str(exc)), 400

    @app.get("/api/preview/<job_id>/<filename>")
    def preview_asset(job_id, filename):
        try:
            job_folder = get_job_folder(Path(app.config["DATA_DIR"]), job_id)
            return send_from_directory(job_folder / "preview", filename)
        except (ValueError, FileNotFoundError) as exc:
            return jsonify(error=str(exc)), 404

    @app.post("/api/convert")
    def convert():
        uploaded_files = [item for item in request.files.getlist("files") if item.filename]
        if not uploaded_files:
            return jsonify(error="请至少选择一个 PDF 或图片文件"), 400
        if len(uploaded_files) > MAX_FILES:
            return jsonify(error=f"一次最多处理 {MAX_FILES} 个文件"), 400

        try:
            layout = parse_layout(request.form)
            compress = request.form.get("quality", "normal") == "compress"
            merge = parse_boolean(request.form.get("merge"))
            job_id, job_folder = create_job(Path(app.config["DATA_DIR"]))
            upload_folder = job_folder / "uploads"
            output_folder = job_folder / "output"
            converter = WordConverter(layout, compress)
            word_files = []
            used_names = set()
            used_output_names = set()

            for uploaded in uploaded_files:
                filename = unique_filename(safe_upload_name(uploaded.filename), used_names)
                ensure_supported(filename)
                input_path = upload_folder / filename
                uploaded.save(input_path)
                output_name = unique_filename(
                    f"{Path(filename).stem}.docx", used_output_names
                )
                output_path = output_folder / output_name
                converter.convert(input_path, output_path)
                word_files.append(output_path)

            if merge and len(word_files) > 1:
                result_path = output_folder / "转换结果_合并.docx"
                converter.merge_documents(word_files, result_path)
                result_kind = "merged"
            elif len(word_files) == 1:
                result_path = word_files[0]
                result_kind = "single"
            else:
                result_path = output_folder / "Word转换结果.zip"
                with zipfile.ZipFile(result_path, "w", zipfile.ZIP_DEFLATED) as archive:
                    for word_file in word_files:
                        archive.write(word_file, arcname=word_file.name)
                result_kind = "archive"

            return jsonify(
                job_id=job_id,
                count=len(word_files),
                filename=result_path.name,
                kind=result_kind,
                download_url=url_for(
                    "download_result", job_id=job_id, filename=result_path.name
                ),
            )
        except (ValueError, RuntimeError, OSError) as exc:
            return jsonify(error=str(exc)), 400

    @app.get("/api/download/<job_id>/<filename>")
    def download_result(job_id, filename):
        try:
            job_folder = get_job_folder(Path(app.config["DATA_DIR"]), job_id)
            return send_from_directory(job_folder / "output", filename, as_attachment=True)
        except (ValueError, FileNotFoundError) as exc:
            return jsonify(error=str(exc)), 404

    return app


def safe_upload_name(original_name):
    name = Path(original_name).name.strip()
    name = INVALID_FILENAME_CHARS.sub("_", name).strip(". ")
    if not name:
        raise ValueError("文件名无效")
    path = Path(name)
    suffix = path.suffix.lower()
    stem = path.stem[:140].strip() or "未命名文件"
    return stem + suffix


def unique_filename(filename, used_names):
    path = Path(filename)
    candidate = filename
    counter = 2
    while candidate.casefold() in used_names:
        candidate = f"{path.stem}_{counter}{path.suffix}"
        counter += 1
    used_names.add(candidate.casefold())
    return candidate


def ensure_supported(filename):
    extension = Path(filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        supported = "、".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"不支持 {extension or '无扩展名'} 文件，仅支持：{supported}")


def parse_layout(form):
    try:
        layout = LayoutConfig(
            image_scale_percent=float(form.get("image_scale", 90)),
            margin_top_cm=float(form.get("margin_top", 2.54)),
            margin_bottom_cm=float(form.get("margin_bottom", 2.54)),
            margin_left_cm=float(form.get("margin_left", 2.54)),
            margin_right_cm=float(form.get("margin_right", 2.54)),
        )
    except (TypeError, ValueError) as exc:
        raise ValueError("图片大小或页边距格式无效") from exc
    layout.validate()
    return layout


def parse_boolean(value):
    return str(value).lower() in {"1", "true", "yes", "on"}


def create_job(data_dir):
    data_dir = data_dir.resolve()
    data_dir.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    job_folder = (data_dir / job_id).resolve()
    if job_folder.parent != data_dir:
        raise RuntimeError("无法创建安全的任务目录")
    for name in ("uploads", "preview", "output"):
        (job_folder / name).mkdir(parents=True, exist_ok=True)
    return job_id, job_folder


def get_job_folder(data_dir, job_id):
    if not JOB_ID_PATTERN.fullmatch(job_id):
        raise ValueError("任务编号无效")
    data_dir = data_dir.resolve()
    job_folder = (data_dir / job_id).resolve()
    if job_folder.parent != data_dir or not job_folder.is_dir():
        raise FileNotFoundError("任务不存在或已过期")
    return job_folder


def cleanup_expired_jobs(data_dir):
    if not data_dir.exists():
        return
    cutoff = time.time() - JOB_LIFETIME_SECONDS
    resolved_data_dir = data_dir.resolve()
    for item in data_dir.iterdir():
        if not item.is_dir() or not JOB_ID_PATTERN.fullmatch(item.name):
            continue
        resolved_item = item.resolve()
        if resolved_item.parent == resolved_data_dir and item.stat().st_mtime < cutoff:
            shutil.rmtree(resolved_item)


def main():
    parser = argparse.ArgumentParser(description="PDF/图片转 Word Web 工具")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=8000, type=int)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    local_url = f"http://{args.host}:{args.port}"
    if not args.no_browser and args.host in {"127.0.0.1", "localhost"}:
        threading.Timer(1.0, lambda: webbrowser.open(local_url)).start()
    print(f"Web 界面已启动：{local_url}")
    create_app().run(host=args.host, port=args.port, debug=False)


if __name__ == "__main__":
    main()
