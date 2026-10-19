"""Web UI tests (criterion 12) using Flask's test client; no network access."""

from __future__ import annotations

import io
import threading
import zipfile

import pytest

from rowdocs.web import HOST, MAX_UPLOAD_BYTES, create_app, make_server


@pytest.fixture
def tmp_root(tmp_path):
    root = tmp_path / "requests"
    root.mkdir()
    return root


@pytest.fixture
def client(tmp_root):
    app = create_app(tmp_root=str(tmp_root))
    app.config["TESTING"] = True
    return app.test_client()


def invoice_form(examples, **extra):
    form = {
        "data": (io.BytesIO((examples / "invoices.csv").read_bytes()), "invoices.csv"),
        "template": (io.BytesIO((examples / "invoice.docx").read_bytes()), "invoice.docx"),
    }
    form.update(extra)
    return form


def test_server_binds_to_localhost_only():
    server = make_server(port=0)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        assert server.server_address[0] == HOST == "127.0.0.1"
    finally:
        server.shutdown()
        thread.join(timeout=5)


def test_index_is_self_contained(client):
    response = client.get("/")
    html = response.get_data(as_text=True)
    assert response.status_code == 200
    assert "<script src=" not in html and "<link" not in html  # no CDN / external assets
    assert "/api/check" in html and "/api/generate" in html


def test_check_shows_warnings_and_preview(client, examples, tmp_root):
    response = client.post("/api/check", data=invoice_form(examples), content_type="multipart/form-data")
    assert response.status_code == 200
    body = response.get_json()
    assert body["report"]["ok"] is True
    assert any(i["level"] == "WARN" and i["row"] == 5 and i["variable"] == "amount" for i in body["report"]["issues"])
    assert body["preview"]["row"] == 1
    assert "Priya Nair" in body["preview"]["text"] and "₹18,500.00" in body["preview"]["text"]
    assert list(tmp_root.iterdir()) == []


def test_generate_downloads_valid_zip(client, examples, tmp_root):
    form = invoice_form(examples, name="{{ invoice_no }}_{{ full_name }}")
    response = client.post("/api/generate", data=form, content_type="multipart/form-data")
    assert response.status_code == 200
    assert response.mimetype == "application/zip"
    with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
        assert archive.testzip() is None
        names = archive.namelist()
    assert "manifest.csv" in names
    assert "INV-1001_Priya Nair.docx" in names and len([n for n in names if n.endswith(".docx")]) == 8
    assert list(tmp_root.iterdir()) == []  # the request's temp folder is gone


def test_exe_upload_rejected(client, examples, tmp_root):
    form = {
        "data": (io.BytesIO(b"MZ\x90\x00"), "setup.exe"),
        "template": (io.BytesIO((examples / "invoice.docx").read_bytes()), "invoice.docx"),
    }
    response = client.post("/api/check", data=form, content_type="multipart/form-data")
    assert response.status_code == 400
    assert "not accepted" in response.get_json()["error"]
    assert list(tmp_root.iterdir()) == []


def test_fake_docx_rejected(client, examples):
    form = invoice_form(examples)
    form["template"] = (io.BytesIO(b"not a zip"), "invoice.docx")
    assert client.post("/api/check", data=form, content_type="multipart/form-data").status_code == 400


def test_validation_errors_block_generate(client, examples, tmp_root):
    response = client.post(
        "/api/generate", data=invoice_form(examples, name="{{ nope }}"), content_type="multipart/form-data"
    )
    assert response.status_code == 422
    assert any("nope" in i["message"] for i in response.get_json()["report"]["issues"])
    assert list(tmp_root.iterdir()) == []


def test_upload_limit(client):
    big = io.BytesIO(b"a" * (MAX_UPLOAD_BYTES + 1))
    response = client.post("/api/check", data={"data": (big, "big.csv")}, content_type="multipart/form-data")
    assert response.status_code == 413


def test_foreign_host_header_rejected(client):
    assert client.get("/", headers={"Host": "evil.example:8765"}).status_code == 403
    assert client.get("/", headers={"Host": "127.0.0.1:8765"}).status_code == 200
