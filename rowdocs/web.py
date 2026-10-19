"""Local web page (``rowdocs serve``). Binds to 127.0.0.1 only; nothing leaves the machine."""

from __future__ import annotations

import io
import shutil
import tempfile
import zipfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from flask import Flask, Response, jsonify, request, send_file
from werkzeug.datastructures import FileStorage
from werkzeug.exceptions import HTTPException

from .build import ValidationFailed, build
from .data import DATA_EXTENSIONS, DataError, Dataset, load_data, select_rows
from .naming import DEFAULT_PATTERN
from .pdf import INSTALL_HELP, find_soffice
from .template import TemplateFileError, WordTemplate, extract_text
from .validate import validate

HOST = "127.0.0.1"
DEFAULT_PORT = 8765
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
_ALLOWED_HOSTS = {"127.0.0.1", "localhost"}
_STATIC = Path(__file__).parent / "static"


class BadUpload(ValueError):
    pass


def _save_upload(field: str, allowed: tuple[str, ...], folder: Path) -> Path:
    upload: FileStorage | None = request.files.get(field)
    if upload is None or not upload.filename:
        raise BadUpload(f"missing file: {field}")
    suffix = Path(upload.filename).suffix.lower()
    if suffix not in allowed:
        raise BadUpload(f"{upload.filename!r} is not accepted for {field}; use {' or '.join(allowed)}")
    target = folder / f"{field}{suffix}"  # never trust the client's file name
    upload.save(target)
    if suffix in (".xlsx", ".docx") and not zipfile.is_zipfile(target):
        raise BadUpload(f"{upload.filename!r} is not a valid {suffix} file")
    return target


def _load_inputs(folder: Path) -> tuple[Dataset, WordTemplate, str]:
    data_path = _save_upload("data", DATA_EXTENSIONS, folder)
    template_path = _save_upload("template", (".docx",), folder)
    sheet = request.form.get("sheet") or None
    rows = request.form.get("rows") or None
    where = [w for w in request.form.getlist("where") if w.strip()]
    dataset = select_rows(load_data(data_path, sheet), rows, where)
    dataset.source = request.files["data"].filename or dataset.source
    template = WordTemplate(template_path, name=request.files["template"].filename)
    return dataset, template, (request.form.get("name") or DEFAULT_PATTERN)


def create_app(tmp_root: str | None = None) -> Flask:
    app = Flask(__name__, static_folder=None)
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES
    app.config["TMP_ROOT"] = tmp_root  # None = system temp folder

    @contextmanager
    def request_dir() -> Iterator[Path]:
        """A private temp folder for one request, always deleted afterwards."""
        folder = Path(tempfile.mkdtemp(prefix="rowdocs-req-", dir=app.config["TMP_ROOT"]))
        try:
            yield folder
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    @app.before_request
    def only_local_hosts() -> Response | None:
        # Blocks DNS-rebinding: a remote page cannot talk to us through a hostile host name.
        host = (request.host or "").rsplit(":", 1)[0].strip("[]")
        if host not in _ALLOWED_HOSTS:
            return Response("Forbidden host", status=403)
        return None

    @app.errorhandler(BadUpload)
    @app.errorhandler(DataError)
    @app.errorhandler(TemplateFileError)
    def bad_request(exc: Exception):
        return jsonify(error=str(exc)), 400

    @app.errorhandler(413)
    def too_large(_exc: HTTPException):
        return jsonify(error=f"upload too large (limit {MAX_UPLOAD_BYTES // (1024 * 1024)} MB)"), 413

    @app.get("/")
    def index():
        return send_file(_STATIC / "index.html", mimetype="text/html")

    @app.get("/api/info")
    def info():
        return jsonify(pdf_available=find_soffice() is not None, max_upload_mb=MAX_UPLOAD_BYTES // (1024 * 1024))

    @app.post("/api/check")
    def check():
        with request_dir() as folder:
            dataset, template, pattern = _load_inputs(folder)
            report = validate(dataset, template, pattern).to_dict()
            preview = {"row": None, "text": "", "error": ""}
            if dataset.selected:
                first = dataset.selected[0]
                preview["row"] = first
                try:
                    preview["text"] = extract_text(template.render(dataset.context(first)))
                except Exception as exc:  # show render problems in the page, not as a 500
                    preview["error"] = f"{exc.__class__.__name__}: {exc}"
            return jsonify(report=report, preview=preview)

    @app.post("/api/generate")
    def generate():
        want_pdf = request.form.get("pdf") in ("1", "true", "on")
        want_merge = request.form.get("merge_pdf") in ("1", "true", "on")
        with request_dir() as folder:
            dataset, template, pattern = _load_inputs(folder)
            if (want_pdf or want_merge) and not find_soffice():
                return jsonify(error=INSTALL_HELP), 400
            try:
                result = build(
                    dataset,
                    template,
                    folder / "rowdocs-output",
                    name_pattern=pattern,
                    pdf=want_pdf,
                    merge_pdf=want_merge,
                    make_zip=True,
                )
            except ValidationFailed as failed:
                return jsonify(error="validation failed; fix the errors first", report=failed.report.to_dict()), 422
            payload = io.BytesIO(result.zip_path.read_bytes())  # read before the folder is deleted
        return send_file(payload, mimetype="application/zip", as_attachment=True, download_name="rowdocs-output.zip")

    return app


def make_server(port: int = DEFAULT_PORT, app: Flask | None = None):
    """A WSGI server bound to 127.0.0.1 only (``port=0`` picks a free port)."""
    from werkzeug.serving import make_server as _make_server

    return _make_server(HOST, port, app or create_app(), threaded=True)


def serve(port: int = DEFAULT_PORT) -> None:  # pragma: no cover - interactive
    server = make_server(port)
    print(f"rowdocs is running at http://{HOST}:{server.server_port}/  (Ctrl+C to stop)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nstopped")
