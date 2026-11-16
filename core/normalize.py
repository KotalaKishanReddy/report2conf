"""Normalize strict-OOXML conference templates to transitional OOXML.

The IEEE A4 template uses the strict namespace (purl.oclc.org) which
python-docx cannot open. We byte-replace the known strict namespaces
with their transitional equivalents inside a copied .docx zip.
"""
import zipfile

REPLACEMENTS = [
    (b"http://purl.oclc.org/ooxml/wordprocessingml/main",
     b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
    (b"http://purl.oclc.org/ooxml/officeDocument/math",
     b"http://schemas.openxmlformats.org/officeDocument/2006/math"),
    (b"http://purl.oclc.org/ooxml/officeDocument/relationships",
     b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
    (b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
     b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
    (b"http://purl.oclc.org/ooxml/drawingml/main",
     b"http://schemas.openxmlformats.org/drawingml/2006/main"),
    (b"http://purl.oclc.org/ooxml/drawingml/picture",
     b"http://schemas.openxmlformats.org/drawingml/2006/picture"),
    (b"http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
     b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
    (b"http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties",
     b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extendedProperties"),
    (b' w:conformance="strict"', b""),
]


def normalize_template(src_path: str, dst_path: str) -> str:
    """Copy src .docx to dst with transitional namespaces. Returns dst_path."""
    with zipfile.ZipFile(src_path) as zin:
        with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name.endswith(".xml") or name.endswith(".rels"):
                    for old, new in REPLACEMENTS:
                        data = data.replace(old, new)
                zout.writestr(name, data)
    return dst_path
