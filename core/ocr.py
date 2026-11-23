"""Local OCR ingest + optional baidu/Unlimited-OCR backend.

Default path is free + safe + offline:
  digital PDF -> text via pypdf (no OCR needed)
  scanned PDF / page images -> local Tesseract binary (already installed)
  .docx -> existing python-docx extractor (no OCR)

baidu/Unlimited-OCR is supported as an *optional* backend for one-shot
long-horizon parsing (multi-page images -> structured markdown). It is NOT
required and NOT installed by default: it needs an A800-class GPU
(weights: baidu/Unlimited-OCR-2) or a paid cloud endpoint. When enabled,
its output is mapped into the same ``extract_report`` dict shape so the
rest of the pipeline (rewrite -> build -> strict rebuild) is unchanged.

Enable it explicitly, e.g.:
  REPORT2CONF_OCR=unlimited python3 app.py
with ``transformers`` + ``torch`` + model weights available. Otherwise the
local path is used automatically.
"""

import io
import os
import re
import subprocess
import tempfile

TEXT_MIN_CHARS = 200


def _tesseract_available() -> bool:
    try:
        r = subprocess.run(
            ["tesseract", "--version"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        return r.returncode == 0
    except Exception:
        return False


def ocr_image_file(image_path: str) -> str:
    """OCR a single image file with the local Tesseract binary."""
    if not _tesseract_available():
        return ""
    try:
        r = subprocess.run(
            ["tesseract", image_path, "stdout", "-l", "eng"],
            capture_output=True,
            text=True,
            timeout=120,
        )
        return r.stdout.strip() if r.returncode == 0 else ""
    except Exception:
        return ""


def _pdf_text_pypdf(path: str) -> str:
    try:
        from pypdf import PdfReader

        reader = PdfReader(path)
        chunks = []
        for page in reader.pages:
            try:
                chunks.append(page.extract_text() or "")
            except Exception:
                chunks.append("")
        return "\n".join(chunks).strip()
    except Exception:
        return ""


def _pdf_page_images(path: str, dpi: int = 200):
    """Render PDF pages to temp PNGs. Returns (paths, tmpdir)."""
    try:
        import fitz  # PyMuPDF
    except Exception:
        return [], None
    tmpdir = tempfile.mkdtemp(prefix="r2c_ocr_")
    doc = fitz.open(path)
    out = []
    try:
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        for i, page in enumerate(doc):
            p = os.path.join(tmpdir, f"page_{i + 1:04d}.png")
            page.get_pixmap(matrix=mat).save(p)
            out.append(p)
    finally:
        doc.close()
    return out, tmpdir


def unlimited_ocr_available() -> bool:
    """True only when the Unlimited-OCR stack is installed + configured."""
    if os.environ.get("REPORT2CONF_OCR", "").lower() != "unlimited":
        return False
    try:
        import torch  # noqa: F401
        from transformers import AutoModel, AutoTokenizer  # noqa: F401
    except Exception:
        return False
    return bool(os.environ.get("UNLIMITED_OCR_MODEL", "baidu/Unlimited-OCR-2"))


def ocr_pages_with_unlimited_ocr(image_paths: list) -> str | None:
    """Parse page images with baidu/Unlimited-OCR. None when unavailable.

    Uses ``infer_multi`` so cross-page context is preserved (one <PAGE>
    block per input page, in order). Falls back to None on any error so
    callers can use the local Tesseract path instead.
    """
    if not unlimited_ocr_available() or not image_paths:
        return None
    try:
        import torch
        from transformers import AutoModel, AutoTokenizer

        model_name = os.environ.get("UNLIMITED_OCR_MODEL", "baidu/Unlimited-OCR-2")
        tokenizer = AutoTokenizer.from_pretrained(model_name, trust_remote_code=True)
        model = AutoModel.from_pretrained(
            model_name, trust_remote_code=True, use_safetensors=True,
            torch_dtype=torch.bfloat16,
        ).eval()
        if torch.cuda.is_available():
            model = model.cuda()
        outdir = tempfile.mkdtemp(prefix="r2c_uocr_")
        text, _ = model.infer_multi(
            tokenizer, image_files=image_paths,
            output_path=os.path.join(outdir, "multi"), save_results=False,
        )
        return text.strip() if text else None
    except Exception:
        return None


def text_to_report_dict(full_text: str, source_name: str = "") -> dict:
    """Map plain OCR/text output into the extract_report dict shape."""
    paras = [p.strip() for p in re.split(r"\n\s*\n", full_text) if p.strip()]
    if not paras:
        lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
        paras = lines
    sections = []
    current = {"heading": "Introduction", "level": 1, "paras": []}
    for para in paras:
        is_heading = (
            len(para) < 120
            and (re.match(r"^\d+(\.\d+)*\s+\S", para) or (
                len(para.split()) <= 8 and para[0].isupper()
                and not para.endswith(".")
            ))
        )
        if is_heading and current["paras"]:
            sections.append(current)
            current = {"heading": para, "level": 1, "paras": []}
        else:
            current["paras"].append(para)
    if current["paras"]:
        sections.append(current)
    if not sections:
        sections = [{"heading": "Introduction", "level": 1,
                     "paras": [source_name or "Empty document."]}]
    words = len(full_text.split())
    return {
        "sections": sections,
        "tables": [],
        "images": [],
        "references": [],
        "words": words,
        "chars": len(full_text),
        "suggested_title": "",
        "suggested_authors": [],
        "ocr_text": True,
    }


def extract_pdf_report(path: str) -> dict:
    """Extract a PDF report: digital text first, OCR only when needed."""
    text = _pdf_text_pypdf(path)
    if len(text) >= TEXT_MIN_CHARS:
        d = text_to_report_dict(text, os.path.basename(path))
        d["ocr_text"] = False
        return d
    # Scanned PDF: render pages, try Unlimited-OCR, else local Tesseract.
    page_paths, tmpdir = _pdf_page_images(path)
    try:
        if page_paths:
            uocr = ocr_pages_with_unlimited_ocr(page_paths)
            if uocr and len(uocr.strip()) >= TEXT_MIN_CHARS // 4:
                return text_to_report_dict(uocr, os.path.basename(path))
            chunks = [ocr_image_file(p) for p in page_paths]
            joined = "\n\n".join(c for c in chunks if c).strip()
            if joined:
                return text_to_report_dict(joined, os.path.basename(path))
    finally:
        if tmpdir:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
    if text:
        return text_to_report_dict(text, os.path.basename(path))
    raise ValueError("Could not extract text from PDF (scanned and OCR empty).")


def extract_image_report(path: str) -> dict:
    """Extract a single page-image report via Unlimited-OCR or Tesseract."""
    uocr = ocr_pages_with_unlimited_ocr([path])
    if uocr and uocr.strip():
        return text_to_report_dict(uocr, os.path.basename(path))
    text = ocr_image_file(path)
    if not text:
        # Last resort: describe the image so the pipeline still builds.
        text = f"Figure from {os.path.basename(path)}."
    d = text_to_report_dict(text, os.path.basename(path))
    try:
        with open(path, "rb") as f:
            blob = f.read()
        ext = os.path.splitext(path)[1].lower()
        ctype = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".tiff": "image/tiff",
            ".tif": "image/tiff",
            ".bmp": "image/bmp",
            ".webp": "image/webp",
        }.get(ext, "image/jpeg")
        d["images"] = [(blob, ctype)]
    except Exception:
        pass
    return d


def extract_any_report(path: str) -> dict:
    """Dispatch .docx/.pdf/.png/.jpg/.tiff to the right extractor."""
    lower = path.lower()
    if lower.endswith(".docx"):
        from .extract import extract_report

        d = extract_report(path)
        # Normalize shape with OCR-sourced dicts.
        d.setdefault("ocr_text", False)
        return d
    if lower.endswith(".pdf"):
        return extract_pdf_report(path)
    if lower.endswith((".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp", ".webp")):
        return extract_image_report(path)
    raise ValueError(f"Unsupported report type: {os.path.basename(path)}")


def png_bytes_to_text(blob: bytes) -> str:
    """OCR raw image bytes with local Tesseract (temp file, offline)."""
    fd, p = tempfile.mkstemp(suffix=".png")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(blob)
        return ocr_image_file(p)
    finally:
        try:
            os.unlink(p)
        except OSError:
            pass
