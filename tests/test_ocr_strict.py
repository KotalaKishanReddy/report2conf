import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.normalize import normalize_template, rebuild_strict
from core.ocr import text_to_report_dict, extract_any_report, unlimited_ocr_available
from core.build import build_paper


def test_rebuild_strict_contains_strict_namespace(tmp_path):
    from docx import Document

    src = tmp_path / "mini.docx"
    doc = Document()
    doc.add_paragraph("Hello strict world.")
    doc.save(src)
    strict = tmp_path / "mini_strict.docx"
    rebuild_strict(str(src), str(strict))
    import zipfile

    data = zipfile.ZipFile(strict).read("word/document.xml")
    assert b"purl.oclc.org/ooxml/wordprocessingml/main" in data
    assert b'w:conformance="strict"' in data
    # And it normalizes back to something python-docx can open.
    back = tmp_path / "back.docx"
    normalize_template(str(strict), str(back))
    doc2 = Document(str(back))
    assert any("Hello strict world." in p.text for p in doc2.paragraphs)


def test_text_to_report_dict_keeps_structure():
    text = "1 Introduction\n\nThis is the intro with number 17.0411 MW.\n\nConclusion\n\nFinal words here."
    d = text_to_report_dict(text, "sample.pdf")
    assert d["words"] > 5
    assert len(d["sections"]) >= 1
    joined = " ".join(p for s in d["sections"] for p in s["paras"])
    assert "17.0411" in joined


def test_extract_any_rejects_unknown_suffix(tmp_path):
    p = tmp_path / "file.txt"
    p.write_text("hello")
    try:
        extract_any_report(str(p))
    except ValueError as e:
        assert "Unsupported" in str(e)
    else:
        raise AssertionError("expected ValueError")


def test_unlimited_ocr_off_by_default():
    os.environ.pop("REPORT2CONF_OCR", None)
    assert unlimited_ocr_available() is False


def test_paper_css_uses_template_type(tmp_path):
    from docx import Document

    from core.typography import paper_css

    p = tmp_path / "t.docx"
    doc = Document()
    doc.add_paragraph("x", style="Title")
    doc.save(p)
    css = paper_css(str(p))
    assert ".doc-title" in css and ".doc-body" in css
    assert "Times New Roman" in css


def test_paper_css_unknown_alignment_falls_back(tmp_path):
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    from core.typography import paper_css

    p = tmp_path / "s.docx"
    doc = Document()
    doc.add_paragraph("x", style="Title")
    pPr = doc.styles["Title"]._element.get_or_add_pPr()
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "start")  # emitted by some generators; unmappable
    pPr.append(jc)
    doc.save(p)
    css = paper_css(str(p))
    assert ".doc-title" in css


def test_save_doc_json_flow():
    import json as _json

    import app as appmod
    from core.normalize import normalize_template

    job = "testdoc1"
    tpl = "/home/kotalakishanreddy/Downloads/conference-template-a4.docx"
    normalize_template(tpl, os.path.join(appmod.OUTDIR, f"{job}_normalized.docx"))
    appmod._save_job_state(
        job, "Old", ["A"], "Abs", "k",
        [{"heading": "H", "level": 1, "text": "old body",
          "orig_words": 2, "new_words": 2}],
        [], [], ["ref one"], 4)
    c = appmod.app.test_client()
    payload = {
        "title": "New", "authors": ["A", "B"], "abstract": "Abs2",
        "keywords": "k",
        "sections": [{"heading": "Intro", "level": 1, "paras": ["Hello 42."]}],
        "tables": [[["a", "b"], ["1", "2"]]],
        "references": ["ref two"],
    }
    r = c.post(f"/save/{job}", data={"doc_json": _json.dumps(payload)})
    assert r.status_code == 200
    html = r.data.decode()
    assert "Images included / total" in html


def test_build_then_strict_keeps_text(tmp_path):
    from core.normalize import normalize_template

    tpl = tmp_path / "tpl.docx"
    from docx import Document

    doc = Document()
    doc.add_paragraph("Paper Title", style="Title")
    doc.add_paragraph("Abstract placeholder")
    doc.add_paragraph("Introduction (heading 1)")
    doc.save(tpl)
    norm = tmp_path / "norm.docx"
    normalize_template(str(tpl), str(norm))
    sections = [{"heading": "Introduction", "level": 1,
                 "text": "Local OCR text with 42 results preserved."}]
    out = tmp_path / "paper.docx"
    build_paper(str(norm), "T", ["A"], "Abs.", "k",
                sections, [], [], [], str(out))
    strict = tmp_path / "paper_strict.docx"
    rebuild_strict(str(out), str(strict))
    assert strict.stat().st_size > 0
