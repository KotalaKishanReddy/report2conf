"""Extract structured content from the uploaded report .docx."""
import os
import re
from docx import Document

GENERIC_HEADINGS = {
    "introduction", "abstract", "references", "conclusion", "conclusions",
    "acknowledgment", "acknowledgement", "keywords", "list of publications",
    "table of contents", "contents",
}

AUTHOR_HINTS = re.compile(
    r"@|university|institute|college|department|school of|faculty of|laboratory",
    re.IGNORECASE,
)

HEADING_RE = re.compile(r"^\d+(\.\d+)*\s+\S")


def _iter_images(doc):
    """Yield (blob, ext) for each embedded image in document order."""
    images = []
    seen = set()
    for para in doc.paragraphs:
        for run in para.runs:
            for el in run._element.iter():
                if el.tag.endswith("}blip"):
                    rid = el.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                    )
                    if rid and rid not in seen:
                        seen.add(rid)
                        try:
                            part = doc.part.related_parts[rid]
                            images.append((part.blob, part.content_type))
                        except KeyError:
                            pass
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        for el in run._element.iter():
                            if el.tag.endswith("}blip"):
                                rid = el.get(
                                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                                )
                                if rid and rid not in seen:
                                    seen.add(rid)
                                    try:
                                        part = doc.part.related_parts[rid]
                                        images.append((part.blob, part.content_type))
                                    except KeyError:
                                        pass
    return images


def extract_report(path: str) -> dict:
    doc = Document(path)
    sections = []
    current = {"heading": "Introduction", "level": 1, "paras": []}
    references = []
    in_refs = False
    tables_data = []
    for table in doc.tables:
        rows = []
        for row in table.rows:
            rows.append([c.text.strip() for c in row.cells])
        if any(any(c for c in r) for r in rows):
            tables_data.append(rows)

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name if para.style else "Normal"
        if text.upper().startswith("REFERENCES") or text.upper().startswith("LIST OF PUBLICATIONS"):
            in_refs = True
            continue
        if in_refs:
            references.append(text)
            continue
        is_heading = style.startswith("Heading") or (
            style in ("Normal",) and HEADING_RE.match(text) and len(text) < 120
        )
        if is_heading:
            if current["paras"]:
                sections.append(current)
            level = 1
            m = re.match(r"^(Heading)\s*(\d)", style)
            if m:
                level = int(m.group(2))
            current = {"heading": text, "level": min(level, 2), "paras": []}
        else:
            current["paras"].append(text)
    if current["paras"]:
        sections.append(current)

    opener = [p for s in sections[:2] for p in s["paras"][:12]]
    suggested_authors = []
    for para in opener:
        if AUTHOR_HINTS.search(para) and len(para) <= 160:
            cleaned = re.sub(r"\s+", " ", para).strip()
            if cleaned not in suggested_authors:
                suggested_authors.append(cleaned)
        if len(suggested_authors) >= 6:
            break

    images = _iter_images(doc)
    full_text = "\n".join(
        p for s in sections for p in s["paras"]
    )
    title_guess = ""
    for s in sections:
        h = s["heading"].strip()
        if (h and h.lower() not in GENERIC_HEADINGS
                and not re.match(r"^\d", h) and len(h) <= 160):
            title_guess = re.sub(r"\s+", " ", h)
            break

    return {
        "sections": sections,
        "tables": tables_data,
        "images": images,
        "references": references,
        "words": len(full_text.split()),
        "chars": len(full_text),
        "suggested_title": title_guess,
        "suggested_authors": suggested_authors,
    }


def prettify_filename(filename: str) -> str:
    """Turn an upload filename into a readable fallback title."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    stem = re.sub(r"\(\d+\)$", "", stem)  # report(1) -> report
    stem = re.sub(r"[_\-]+", " ", stem).strip()
    return stem[:120]
