"""Repeated end-to-end conversions through the Flask app.

Runs the full upload -> convert -> download path several times with
different templates, page limits, and blank/filled name fields. The PDF
measurement is stubbed for speed; real LibreOffice measurement is covered
by the manual multi-run validation in the shell.
"""
import io
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app as app_module

REPORT = "/home/kotalakishanreddy/Downloads/report(1).docx"
TEMPLATE = "/home/kotalakishanreddy/Downloads/conference-template-a4.docx"
FILLED_TEMPLATE = "/home/kotalakishanreddy/Downloads/conference-template-a4-filled(1).docx"


@pytest.fixture()
def client(monkeypatch):
    monkeypatch.setattr(app_module, "pdf_page_count", lambda *a, **k: (3, None))
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as client:
        yield client


def _files(report_path=REPORT, template_path=TEMPLATE):
    with open(report_path, "rb") as f:
        rep = f.read()
    with open(template_path, "rb") as f:
        tpl = f.read()
    return rep, tpl


def _post(client, rep, tpl, title="", authors="", max_pages="4"):
    data = {
        "report": (io.BytesIO(rep), "report.docx"),
        "template": (io.BytesIO(tpl), "template.docx"),
        "title": title,
        "authors": authors,
        "keywords": "test",
        "max_pages": max_pages,
    }
    return client.post("/convert", data=data,
                       content_type="multipart/form-data")


@pytest.mark.parametrize("round_no", [1, 2, 3])
@pytest.mark.parametrize("max_pages", ["1", "4", "8"])
@pytest.mark.parametrize("template_path", [TEMPLATE, FILLED_TEMPLATE])
def test_convert_repeatedly(client, round_no, max_pages, template_path, tmp_path):
    rep, tpl = _files(template_path=template_path)
    resp = _post(client, rep, tpl, max_pages=max_pages)
    assert resp.status_code == 200, resp.data[:300]
    assert b"Conversion result" in resp.data
    # blank name fields must autofill, never 500
    assert b"Paper title:" in resp.data


@pytest.mark.parametrize("round_no", [1, 2])
def test_convert_with_typed_names(client, round_no):
    rep, tpl = _files()
    resp = _post(client, rep, tpl, title="My Paper", authors="A. One\nB. Two",
                 max_pages="6")
    assert resp.status_code == 200
    assert b"My Paper" in resp.data
    assert b"filled in automatically" not in resp.data


def test_blank_names_are_autofilled(client):
    rep, tpl = _files()
    resp = _post(client, rep, tpl, title="", authors="")
    assert resp.status_code == 200
    assert b"filled in automatically" in resp.data


def test_missing_template_is_400(client):
    rep, _ = _files()
    resp = client.post("/convert",
                       data={"report": (io.BytesIO(rep), "report.docx")},
                       content_type="multipart/form-data")
    assert resp.status_code == 400


def test_non_docx_is_400(client):
    resp = client.post(
        "/convert",
        data={"report": (io.BytesIO(b"nope"), "report.txt"),
              "template": (io.BytesIO(b"nope"), "template.txt")},
        content_type="multipart/form-data")
    assert resp.status_code == 400


def test_download_serves_paper(client):
    rep, tpl = _files()
    resp = _post(client, rep, tpl)
    assert resp.status_code == 200
    body = resp.data.decode()
    marker = '/download/'
    idx = body.index(marker)
    name = body[idx + len(marker):].split('"')[0]
    dl = client.get(f"/download/{name}")
    assert dl.status_code == 200
    assert len(dl.data) > 10000
    # output must be a readable .docx containing the typed title default
    out = os.path.join("/tmp", f"stress_{name}")
    with open(out, "wb") as f:
        f.write(dl.data)
    from docx import Document

    doc = Document(out)
    assert len(doc.paragraphs) > 20
