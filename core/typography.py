"""Display typography read from the conference template.

The online editor must follow the paper's own font, size, and alignment
1-to-1, so this module reads the resolved styles out of the normalized
template (via the same resolution rules as ``core.build``) and renders
them as CSS. No hardcoded IEEE values: what the template defines wins,
sane defaults cover the rest.
"""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .build import _resolve

_ALIGN_CSS = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    WD_ALIGN_PARAGRAPH.DISTRIBUTE: "justify",
}

# role -> (candidate style names, fallback fmt)
ROLES = {
    "title": (["Paper Title"], {"size": 24.0, "align": "center"}),
    "author": (["Author"], {"size": 11.0, "align": "center"}),
    "abstract": (["Abstract"], {"size": 9.0, "align": "justify"}),
    "keywords": (["Keywords"], {"size": 9.0, "align": "justify"}),
    "h1": (["Heading1", "Heading 1"], {"size": 10.0, "align": "center", "bold": True}),
    "h2": (["Heading2", "Heading 2"], {"size": 10.0, "align": "left", "italic": True}),
    "body": (["BodyText"], {"size": 10.0, "align": "justify"}),
    "cap": (["figurecaption"], {"size": 8.0, "align": "center"}),
    "ref": (["references"], {"size": 8.0, "align": "justify"}),
    "thead": (["tablecolhead", "tablehead"], {"size": 8.0, "align": "center", "bold": True}),
    "tcopy": (["tablecopy"], {"size": 8.0, "align": "justify"}),
}


def _fmt(doc, names, fallback):
    name = _resolve(doc, *names, default=None)
    font_name = "Times New Roman"
    size = fallback.get("size", 10.0)
    align = fallback.get("align", "justify")
    bold = fallback.get("bold", False)
    italic = fallback.get("italic", False)
    small_caps = False
    if name is not None:
        try:
            style = doc.styles[name]
        except KeyError:
            style = None
        if style is not None:
            font = style.font
            if font.name:
                font_name = font.name
            if font.size is not None:
                size = font.size.pt
            if font.bold is not None:
                bold = font.bold
            if font.italic is not None:
                italic = font.italic
            if font.small_caps is not None:
                small_caps = font.small_caps
            pf = style.paragraph_format
            try:
                jc = pf.alignment
            except ValueError:
                # Newer OOXML justification values (e.g. w:val="start")
                # that python-docx cannot map: keep the role default.
                jc = None
            if jc is not None:
                align = _ALIGN_CSS.get(jc, align)
    return {
        "font": font_name, "size": size, "align": align,
        "bold": bold, "italic": italic, "small_caps": small_caps,
    }


def _rule(fmt, extra=""):
    fam = fmt["font"].replace('"', "")
    decls = [f'font-family: "{fam}", "Times New Roman", serif',
             f'font-size: {fmt["size"]}pt',
             f'text-align: {fmt["align"]}']
    if fmt["bold"]:
        decls.append("font-weight: bold")
    if fmt["italic"]:
        decls.append("font-style: italic")
    if fmt["small_caps"]:
        decls.append("font-variant: small-caps")
    if extra:
        decls.append(extra)
    return "; ".join(decls)


def paper_css(normalized_template: str) -> str:
    """CSS text mirroring the template's own title/body/table/reference type."""
    doc = Document(normalized_template)
    f = {role: _fmt(doc, names, fb) for role, (names, fb) in ROLES.items()}
    lines = [
        f'.doc-title {{{_rule(f["title"])}}}',
        f'.doc-author {{{_rule(f["author"])}}}',
        f'.doc-abstract {{{_rule(f["abstract"])}}}',
        f'.doc-keys {{{_rule(f["keywords"])}}}',
        f'.doc-h1 {{{_rule(f["h1"])}}}',
        f'.doc-h2 {{{_rule(f["h2"])}}}',
        f'.doc-body {{{_rule(f["body"], "line-height: 1.25; margin: 0 0 0.4em")}}}',
        f'.doc-cap {{{_rule(f["cap"], "margin: 0.8em 0 0.3em")}}}',
        f'.doc-ref {{{_rule(f["ref"], "margin: 0 0 0.3em")}}}',
        f'.doc-table table {{width: 100%; border-collapse: collapse}}',
        f'.doc-table th {{{_rule(f["thead"])}}}',
        f'.doc-table td {{{_rule(f["tcopy"])}}}',
        '.doc-table th, .doc-table td {border: 1px solid #999; padding: 3px 5px}',
    ]
    return "\n".join(lines)
