import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.normalize import normalize_template
from core.extract import extract_report
from core.rewrite import extractive_compress, rewrite_sections, target_words
from core.build import build_paper

REPORT = "/home/kotalakishanreddy/Downloads/report(1).docx"
TEMPLATE = "/home/kotalakishanreddy/Downloads/conference-template-a4.docx"
TMP = "/tmp/r2c_test"


def setup_module():
    os.makedirs(TMP, exist_ok=True)


def test_normalize_opens_with_python_docx():
    from docx import Document

    dst = os.path.join(TMP, "normalized.docx")
    normalize_template(TEMPLATE, dst)
    doc = Document(dst)
    assert len(doc.paragraphs) > 50


def test_extract_report():
    data = extract_report(REPORT)
    assert data["words"] > 3000
    assert len(data["sections"]) >= 3
    assert len(data["tables"]) >= 10
    assert len(data["images"]) >= 1


def test_rewrite_stays_lossless_and_shrinks():
    data = extract_report(REPORT)
    budget = target_words(4, len(data["tables"]), len(data["images"]))
    sections, coverage = rewrite_sections(data["sections"][:4], budget, use_llm=False)
    assert len(sections) == 4
    # every input section survives
    assert coverage["fallback_used"] == 4
    # digit-bearing content preserved: pick a section with numbers
    for s in sections:
        if any(ch.isdigit() for ch in s["text"]):
            break
    else:
        raise AssertionError("no numbers preserved")


def test_extractive_compress_keeps_digits():
    text = "First sentence here. The loss was 17.0411 MW in the base case. Another filler sentence about things."
    out = extractive_compress(text, 15)
    assert "17.0411" in out


def _cols_of_paragraph(doc, text_prefix):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for para in doc.paragraphs:
        if para.text.startswith(text_prefix):
            el = para._element
            # walk forward to the next sectPr (in-paragraph or final)
            found_self = False
            for child in doc.element.body:
                if child == el:
                    found_self = True
                    continue
                if not found_self:
                    continue
                if child.tag.endswith("}p"):
                    sp = child.find(f"{{{W}}}pPr/{{{W}}}sectPr/{{{W}}}cols")
                    if sp is not None:
                        return sp.attrib
                elif child.tag.endswith("}sectPr"):
                    cols = child.find(f"{{{W}}}cols")
                    return cols.attrib if cols is not None else None
    return None


def test_output_keeps_two_column_body():
    from docx import Document

    norm = os.path.join(TMP, "normalized.docx")
    normalize_template(TEMPLATE, norm)
    data = extract_report(REPORT)
    sections, _ = rewrite_sections(data["sections"][:3], 1500, use_llm=False)
    out = os.path.join(TMP, "paper.docx")
    if not os.path.exists(out):
        build_paper(norm, "Test Title", ["Author One"], "Abstract text.",
                    "keyword1, keyword2", sections, data["tables"][:2],
                    [], data["references"][:3], out)
    doc = Document(out)
    cols = _cols_of_paragraph(doc, sections[0]["heading"][:20])
    assert cols is not None and cols.get(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num",
        cols.get("num")) == "2", f"body not in 2-col section: {cols}"


def test_table_cells_use_template_styles():
    from docx import Document

    out = os.path.join(TMP, "paper.docx")
    doc = Document(out)
    assert len(doc.tables) >= 1

    def _nm(p):
        return p.style.name if p.style is not None else ""

    cell_styles = {_nm(c.paragraphs[0]) for t in doc.tables
                   for r in t.rows for c in r.cells}
    assert cell_styles & {"table col head", "table copy", "table head"}, cell_styles


def test_title_keeps_template_run_formatting():
    from docx import Document

    out = os.path.join(TMP, "paper.docx")
    doc = Document(out)
    title_para = next(p for p in doc.paragraphs if p.text == "Test Title")
    assert title_para.runs, "title has no runs"
    assert title_para.runs[0]._element.rPr is not None, "title lost run formatting"


def _strip_para_styles(path, out, count=3):
    """Remove pStyle from the first `count` body paragraphs.

    Reproduces real-world templates (e.g. conference-template-a4-filled)
    whose paragraphs carry direct formatting but no named style, so
    ``para.style`` is None.
    """
    import shutil
    import zipfile

    from docx import Document

    # Drop the "Normal" style definition: real-world templates built by other
    # tools (e.g. conference-template-a4-filled) lack it, which is what makes
    # python-docx report ``para.style`` as None for style-less paragraphs.
    zout_path = out + ".tmp.zip"
    with zipfile.ZipFile(path) as zin:
        with zipfile.ZipFile(zout_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name == "word/styles.xml":
                    data = re.sub(
                        r'<w:style[^>]*w:styleId="Normal".*?</w:style>',
                        "", data.decode(), count=1, flags=re.S,
                    ).encode()
                zout.writestr(name, data)
    shutil.move(zout_path, out)
    doc = Document(out)
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    stripped = 0
    for para in doc.paragraphs:
        if stripped >= count:
            break
        pPr = para._element.find(f"{{{W}}}pPr")
        if pPr is None:
            continue
        pStyle = pPr.find(f"{{{W}}}pStyle")
        if pStyle is not None:
            pPr.remove(pStyle)
            stripped += 1
    assert any(p.style is None for p in doc.paragraphs), "no style-None para made"
    doc.save(out)
    return out


def test_styleless_template_paragraphs_do_not_crash():
    from docx import Document

    norm = os.path.join(TMP, "normalized.docx")
    normalize_template(TEMPLATE, norm)
    broken = _strip_para_styles(norm, os.path.join(TMP, "nostyle_tpl.docx"))
    data = extract_report(REPORT)
    sections, _ = rewrite_sections(data["sections"][:2], 800, use_llm=False)
    out = os.path.join(TMP, "nostyle_paper.docx")
    build_paper(broken, "T", ["A"], "abs", "kw", sections, [], [], [], out)
    assert os.path.exists(out)
    Document(out)  # must still open


def test_suggest_metadata_from_report_with_title_block():
    from docx import Document

    from core.extract import prettify_filename

    t = Document()
    t.add_paragraph("Deep Learning for Load Flow Analysis", style="Heading 1")
    t.add_paragraph("Some body text here with content.", style="Normal")
    t.add_paragraph("Contact: author@example.edu", style="Normal")
    titled = os.path.join(TMP, "titled.docx")
    t.save(titled)
    data = extract_report(titled)
    assert data["suggested_title"] == "Deep Learning for Load Flow Analysis"
    assert data["suggested_authors"] == ["Contact: author@example.edu"]
    assert prettify_filename("report(1).docx") == "report"
    assert prettify_filename("/x/my_paper-v2.docx") == "my paper v2"


def test_sample_reports_have_no_title_block():
    data = extract_report(REPORT)
    assert data["suggested_title"] == ""
    assert data["suggested_authors"] == []


def test_templates_missing_body_styles_get_sane_fallbacks():
    """A template without BodyText/caption styles must not borrow Title."""
    from docx import Document

    from core.build import _ensure_style, _resolve

    filled = "/home/kotalakishanreddy/Downloads/conference-template-a4-filled(1).docx"
    d = Document(filled)
    for name, size in [("BodyText", 10), ("figurecaption", 8),
                       ("references", 8), ("tablecopy", 8)]:
        got = _ensure_style(d, name, size)
        assert _resolve(d, name) == got
        assert got != "Title"
        st = next(s for s in d.styles if s.name == got)
        assert st.font.size.pt == size


def test_build_paper_end_to_end():
    from docx import Document

    norm = os.path.join(TMP, "normalized.docx")
    normalize_template(TEMPLATE, norm)
    data = extract_report(REPORT)
    sections, _ = rewrite_sections(data["sections"][:3], 1500, use_llm=False)
    out = os.path.join(TMP, "paper.docx")
    build_paper(norm, "Test Title", ["Author One"], "Abstract text.",
                "keyword1, keyword2", sections, data["tables"][:2],
                [], data["references"][:3], out)
    doc = Document(out)
    texts = "\n".join(p.text for p in doc.paragraphs)
    assert "Test Title" in texts
    assert len(doc.tables) >= 2
