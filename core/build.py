"""Build the filled conference paper from the normalized template.

Layout contract: the template's section breaks (single-column title block,
multi-column author block, two-column body) are structural — new content is
inserted INSIDE the template's two-column body section, never appended after
it. All formatting comes from the template's own styles.
"""
import copy
import os
import tempfile
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _style_exists(doc, name):
    return any(s.name == name or s.style_id == name for s in doc.styles)


def _resolve(doc, *names, default="Normal"):
    for n in names:
        if _style_exists(doc, n):
            return n
    if _style_exists(doc, default):
        return default
    # Templates built by other tools may lack every usual style; fall back
    # to the first paragraph style actually defined in the document.
    for s in doc.styles:
        if s.type == WD_STYLE_TYPE.PARAGRAPH:
            return s.name
    return default


def _ensure_style(doc, name, size_pt, bold=False, italic=False, centered=False):
    """Add a missing IEEE-like paragraph style; existing styles win."""
    if _style_exists(doc, name):
        return name
    try:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    except ValueError:
        return name
    st.font.name = "Times New Roman"
    st.font.size = Pt(size_pt)
    st.font.bold = bold
    st.font.italic = italic
    st.paragraph_format.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    return name


def _has_sectPr(para):
    return para._element.find(f"{{{W_NS}}}pPr/{{{W_NS}}}sectPr") is not None


def _sname(para):
    """Paragraph style name, robust to style-less paragraphs.

    Templates built by other tools may lack a Normal/default style, in which
    case python-docx reports ``para.style`` as None.
    """
    style = para.style
    return style.name if style is not None else ""


def _copy_first_run_format(para):
    """Snapshot the first run's rPr element so replacement text keeps its look."""
    for run in para.runs:
        if run._element.rPr is not None:
            return copy.deepcopy(run._element.rPr)
    return None


def _set_para_text(para, text):
    """Replace paragraph text, keeping the first run's character formatting."""
    rPr = _copy_first_run_format(para)
    p_el = para._element
    for child in list(p_el):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            p_el.remove(child)
    run = para.add_run(text)
    if rPr is not None:
        if run._element.rPr is not None:
            run._element.remove(run._element.rPr)
        run._element.insert(0, rPr)
    return run


def _clear_para_text(para):
    _set_para_text(para, "")


def build_paper(normalized_template: str, title: str, authors: list,
                abstract: str, keywords: str, sections: list,
                tables: list, images: list, references: list,
                out_path: str) -> str:
    doc = Document(normalized_template)

    # 1. Title / authors / abstract / keywords (header matter, keep in place).
    header_end = 0
    for idx, para in enumerate(doc.paragraphs):
        if _sname(para).lower().startswith("paper title"):
            _set_para_text(para, title)
            header_end = max(header_end, idx + 1)
            break
    author_paras = [(i, p) for i, p in enumerate(doc.paragraphs)
                    if _sname(p) == "Author"]
    for i, (idx, para) in enumerate(author_paras):
        _set_para_text(para, authors[i] if i < len(authors) else "")
        header_end = max(header_end, idx + 1)
    for idx, para in enumerate(doc.paragraphs):
        sname = _sname(para)
        if sname == "Abstract" and abstract:
            _set_para_text(para, "Abstract\u2014" + abstract)
            header_end = max(header_end, idx + 1)
        elif sname == "Keywords" and keywords:
            _set_para_text(para, "Keywords\u2014" + keywords)
            header_end = max(header_end, idx + 1)

    # Content styles: the template's own definitions win; missing ones are
    # added with IEEE-like formatting so output stays readable on any template.
    _ensure_style(doc, "BodyText", 10)
    _ensure_style(doc, "figurecaption", 8)
    _ensure_style(doc, "references", 8)
    _ensure_style(doc, "tablehead", 8, bold=True, centered=True)
    _ensure_style(doc, "tablecolhead", 8, bold=True, centered=True)
    _ensure_style(doc, "tablecopy", 8)
    body_style = _resolve(doc, "BodyText")
    h1 = _resolve(doc, "Heading1", "Heading 1")
    h2 = _resolve(doc, "Heading2", "Heading 2")
    cap_style = _resolve(doc, "figurecaption")
    ref_style = _resolve(doc, "references")
    ref_head = _resolve(doc, "Heading5", "Heading 5", "Heading1", "Heading 1")
    t_head = _resolve(doc, "tablecolhead", "tablehead", default=None)
    t_copy = _resolve(doc, "tablecopy", default=None)
    t_table = _resolve(doc, "Table Grid", "Normal Table", default=None)

    # 2. Locate the body start and the section-break anchors to preserve.
    paras = list(doc.paragraphs)
    cut_start = None
    for i, para in enumerate(paras):
        if i > 5 and (
            _sname(para) in ("Heading1", "Heading 1", "Heading 2")
            or para.text.strip().lower().startswith("introduction (heading 1)")
        ):
            cut_start = i
            break
    if cut_start is None:
        # Unknown template layout: keep all header matter, replace the rest.
        cut_start = min(header_end, len(paras))

    keepers = [p for p in paras[cut_start:] if _has_sectPr(p)]
    anchor = keepers[0] if keepers else None
    if anchor is not None:
        # The anchor carries the two-column body section. If it sits after the
        # template's References heading it is placeholder content: clear the
        # text but keep the paragraph (and its sectPr) as the section anchor.
        ref_seen = any(
            p.text.strip().lower() == "references"
            for p in paras[:paras.index(anchor)]
        )
        if ref_seen:
            _clear_para_text(anchor)

    # 3. Remove template instructional body (and its sample tables), keeping
    #    section-break anchors and everything above cut_start.
    body_el = doc.element.body
    for para in paras[cut_start:]:
        if para in keepers:
            continue
        body_el.remove(para._element)
    for table in list(doc.tables):
        body_el.remove(table._element)

    # 4. Insert new content directly before the two-column section anchor
    #    (python-docx appends at the end, which would land past the anchor
    #    in the final single-column section, so each block is moved up front).
    #    Templates without section breaks have no anchor: append at the end.
    anchor_el = anchor._element if anchor is not None else None

    def _emit(factory):
        obj = factory()
        if anchor_el is not None:
            anchor_el.addprevious(obj._element)
        return obj

    for sec in sections:
        _emit(lambda: doc.add_paragraph(sec["heading"],
                                        style=h1 if sec["level"] == 1 else h2))
        for chunk in sec["text"].split("\n"):
            chunk = chunk.strip()
            if chunk:
                _emit(lambda c=chunk: doc.add_paragraph(c, style=body_style))

    tmpimgs = []
    for idx, (blob, ctype) in enumerate(images[:8]):
        ext = ".png" if "png" in (ctype or "") else ".jpg"
        fd, ipath = tempfile.mkstemp(suffix=ext)
        with os.fdopen(fd, "wb") as f:
            f.write(blob)
        tmpimgs.append(ipath)
        _emit(lambda: doc.add_paragraph(f"Figure {idx + 1}. Report figure {idx + 1}.",
                                        style=cap_style))
        para = doc.add_paragraph(style=body_style)
        para.add_run().add_picture(ipath, width=Inches(2.8))
        if anchor_el is not None:
            anchor_el.addprevious(para._element)

    for ti, rows in enumerate(tables[:10]):
        if not rows:
            continue
        _emit(lambda: doc.add_paragraph(f"Table {ti + 1}. Report data.",
                                        style=cap_style))
        ncols = max(len(r) for r in rows)
        table = _emit(lambda: doc.add_table(rows=len(rows), cols=ncols))
        if t_table:
            try:
                table.style = t_table
            except (KeyError, ValueError):
                pass
        for ri, row in enumerate(rows):
            for ci in range(ncols):
                val = row[ci] if ci < len(row) else ""
                cell = table.cell(ri, ci)
                cell.text = ""
                cell_style = t_head if ri == 0 else t_copy
                p = cell.paragraphs[0]
                if cell_style:
                    try:
                        p.style = cell_style
                    except (KeyError, ValueError):
                        pass
                p.add_run(val[:300])

    if references:
        _emit(lambda: doc.add_paragraph("References", style=ref_head))
        for ref in references[:40]:
            _emit(lambda r=ref: doc.add_paragraph(r, style=ref_style))

    for p in tmpimgs:
        try:
            os.unlink(p)
        except OSError:
            pass
    doc.save(out_path)
    return out_path
