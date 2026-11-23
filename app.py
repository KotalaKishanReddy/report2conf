"""report2conf: upload a report + IEEE template, get a page-limited paper."""
import json
import os
import uuid
from flask import Flask, render_template, request, send_file, abort

from core.normalize import normalize_template, rebuild_strict
from core.extract import prettify_filename
from core.ocr import extract_any_report
from core.rewrite import rewrite_sections, target_words
from core.build import build_paper
from core.paginate import pdf_page_count

BASE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(BASE, "outputs")
os.makedirs(OUTDIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

REPORT_EXTS = (".docx", ".pdf", ".png", ".jpg", ".jpeg",
               ".tiff", ".tif", ".bmp", ".webp")


def _save_job_state(job, title, authors, abstract, keywords,
                    sections, tables, image_metas, references, max_pages,
                    ocr_used=False):
    state = {
        "title": title, "authors": authors, "abstract": abstract,
        "keywords": keywords, "sections": sections, "tables": tables,
        "image_metas": image_metas, "references": references,
        "max_pages": max_pages, "ocr_used": ocr_used,
    }
    with open(os.path.join(OUTDIR, f"{job}.json"), "w") as f:
        json.dump(state, f)


def _load_job_state(job):
    with open(os.path.join(OUTDIR, f"{job}.json")) as f:
        state = json.load(f)
    images = []
    for meta in state.get("image_metas", []):
        with open(os.path.join(OUTDIR, meta["file"]), "rb") as f:
            images.append((f.read(), meta["ctype"]))
    return state, images


def _build_dual(norm_path, title, authors, abstract, keywords,
                sections, tables, images, references, out_path):
    build_paper(norm_path, title, authors, abstract, keywords, sections,
                tables, images, references, out_path)
    strict_path = out_path.replace(".docx", "_strict.docx")
    rebuild_strict(out_path, strict_path)
    return out_path, strict_path


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/convert")
def convert():
    report = request.files.get("report")
    template = request.files.get("template")
    if not report or not template or not report.filename or not template.filename:
        abort(400, "Both report and template files are required.")
    rep_lower = report.filename.lower()
    if not rep_lower.endswith(REPORT_EXTS):
        abort(400, f"{report.filename} must be .docx, .pdf, or an image.")
    if not template.filename.lower().endswith(".docx"):
        abort(400, f"{template.filename} is not a .docx file.")
    try:
        max_pages = max(1, min(20, int(request.form.get("max_pages", "4"))))
    except ValueError:
        abort(400, "max_pages must be a number.")
    # Free + safe default: extractive fallback unless the user opts into API.
    use_llm = request.form.get("use_llm", "off") == "on"
    form_title = request.form.get("title", "").strip()
    authors_raw = request.form.get("authors", "").strip()

    job = uuid.uuid4().hex[:8]
    rep_ext = os.path.splitext(report.filename)[1].lower() or ".docx"
    rep_path = os.path.join(OUTDIR, f"{job}_report{rep_ext}")
    tpl_path = os.path.join(OUTDIR, f"{job}_template.docx")
    norm_path = os.path.join(OUTDIR, f"{job}_normalized.docx")
    out_path = os.path.join(OUTDIR, f"{job}_paper.docx")
    report.save(rep_path)
    template.save(tpl_path)

    try:
        normalize_template(tpl_path, norm_path)
        data = extract_any_report(rep_path)
    except Exception as e:  # noqa: BLE001
        app.logger.exception("input parsing failed")
        return render_template("error.html", message=f"Could not read: {e}"), 500

    # Autofill: blank fields fall back to report-derived suggestions.
    autofilled = []
    title = form_title
    if not title:
        title = (data["suggested_title"]
                 or prettify_filename(report.filename)
                 or "Converted Conference Paper")
        autofilled.append("title")
    authors = [a.strip() for a in authors_raw.splitlines() if a.strip()]
    if not authors:
        authors = data["suggested_authors"] or ["Author Name"]
        autofilled.append("authors")

    budget = target_words(max_pages, len(data["tables"]), len(data["images"]))
    try:
        sections, coverage = rewrite_sections(data["sections"], budget, use_llm=use_llm)

        abstract = ""
        if sections:
            abstract = sections[0]["text"][:1200]
        keywords = request.form.get("keywords", "conference paper, formatting, results").strip()
        out_path, strict_path = _build_dual(
            norm_path, title, authors, abstract, keywords, sections,
            data["tables"], data["images"], data["references"], out_path)

        pages, pdf_path = pdf_page_count(out_path)
        # One tightening pass if over budget and content can shrink.
        if pages and pages > max_pages:
            sections, coverage = rewrite_sections(
                data["sections"], int(budget * max_pages / pages * 0.9), use_llm=False)
            abstract = sections[0]["text"][:1200] if sections else ""
            out_path, strict_path = _build_dual(
                norm_path, title, authors, abstract, keywords, sections,
                data["tables"][:6], data["images"][:4],
                data["references"], out_path)
            pages2, pdf_path2 = pdf_page_count(out_path)
            if pages2:
                pages, pdf_path = pages2, pdf_path2

        tight = bool(pages and pages > max_pages)
        kept_tables = min(len(data["tables"]), 6 if tight else 10)
        kept_images = min(len(data["images"]), 4 if tight else 8)
        coverage_report = {
            "report_words": data["words"],
            "sections_kept": len(sections),
            "tables_total": len(data["tables"]),
            "tables_included": kept_tables,
            "tables_deferred": max(0, len(data["tables"]) - kept_tables),
            "images_total": len(data["images"]),
            "images_included": kept_images,
            "references_kept": min(len(data["references"]), 40),
            "references_total": len(data["references"]),
            "references_deferred": max(0, len(data["references"]) - 40),
            "pages": pages,
            "max_pages": max_pages,
            "within_budget": bool(pages and pages <= max_pages),
            "llm_used": coverage["llm_used"],
            "fallback_used": coverage["fallback_used"],
            "pdf_error": None if pages else pdf_path,
            "ocr_used": bool(data.get("ocr_text")),
        }
        # Persist images as files so the online editor can rebuild.
        image_metas = []
        for i, (blob, ctype) in enumerate(data["images"][:8]):
            iname = f"{job}_img{i}.bin"
            with open(os.path.join(OUTDIR, iname), "wb") as f:
                f.write(blob)
            image_metas.append({"file": iname, "ctype": ctype})
        _save_job_state(job, title, authors, abstract, keywords,
                        sections, data["tables"], image_metas,
                        data["references"][:40], max_pages,
                        ocr_used=bool(data.get("ocr_text")))
    except Exception as e:  # noqa: BLE001
        app.logger.exception("conversion failed")
        return render_template("error.html", message=str(e)[:500]), 500
    return render_template("result.html", job=job,
                           out_file=os.path.basename(out_path),
                           strict_file=os.path.basename(strict_path),
                           coverage=coverage_report,
                           sections=coverage["sections"],
                           title=title, autofilled=autofilled)


@app.get("/edit/<job>")
def edit(job):
    if ".." in job or "/" in job:
        abort(404)
    try:
        state, _images = _load_job_state(job)
    except FileNotFoundError:
        abort(404)
    return render_template("edit.html", job=job, state=state)


@app.post("/save/<job>")
def save(job):
    if ".." in job or "/" in job:
        abort(404)
    try:
        state, images = _load_job_state(job)
    except FileNotFoundError:
        abort(404)
    title = request.form.get("title", state["title"]).strip() or state["title"]
    authors = [a.strip() for a in request.form.get("authors", "\n".join(state["authors"])).splitlines() if a.strip()] or state["authors"]
    abstract = request.form.get("abstract", state["abstract"])
    keywords = request.form.get("keywords", state["keywords"])
    headings = request.form.getlist("heading")
    texts = request.form.getlist("text")
    sections = state["sections"]
    if headings and texts and len(headings) == len(texts):
        sections = [{"heading": h.strip() or f"Section {i + 1}", "level": 1, "text": t,
                     "orig_words": len(t.split()), "new_words": len(t.split())}
                    for i, (h, t) in enumerate(zip(headings, texts))]
    norm_path = os.path.join(OUTDIR, f"{job}_normalized.docx")
    out_path = os.path.join(OUTDIR, f"{job}_paper.docx")
    out_path, strict_path = _build_dual(
        norm_path, title, authors, abstract, keywords, sections,
        state["tables"], images, state["references"], out_path)
    pages, pdf_path = pdf_page_count(out_path)
    _save_job_state(job, title, authors, abstract, keywords,
                    sections, state["tables"], state["image_metas"],
                    state["references"], state["max_pages"],
                    ocr_used=state.get("ocr_used", False))
    coverage_report = {
        "report_words": state.get("sections") and sum(s.get("new_words", 0) for s in sections) or 0,
        "sections_kept": len(sections),
        "tables_total": len(state["tables"]),
        "tables_included": min(len(state["tables"]), 10),
        "tables_deferred": max(0, len(state["tables"]) - 10),
        "images_total": len(images),
        "images_included": len(images),
        "references_kept": min(len(state["references"]), 40),
        "references_total": len(state["references"]),
        "references_deferred": max(0, len(state["references"]) - 40),
        "pages": pages,
        "max_pages": state["max_pages"],
        "within_budget": bool(pages and pages <= state["max_pages"]),
        "llm_used": 0,
        "fallback_used": len(sections),
        "pdf_error": None if pages else pdf_path,
        "ocr_used": state.get("ocr_used", False),
        "edited": True,
    }
    coverage_sections = [{"heading": s["heading"],
                          "orig_words": s.get("orig_words", len(s["text"].split())),
                          "new_words": len(s["text"].split())} for s in sections]
    return render_template("result.html", job=job,
                           out_file=os.path.basename(out_path),
                           strict_file=os.path.basename(strict_path),
                           coverage=coverage_report,
                           sections=coverage_sections,
                           title=title, autofilled=[])


@app.get("/download/<name>")
def download(name):
    path = os.path.join(OUTDIR, name)
    if not os.path.isfile(path) or ".." in name:
        abort(404)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5057, debug=False)
