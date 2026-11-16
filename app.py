"""report2conf: upload a report + IEEE template, get a page-limited paper."""
import os
import uuid
from flask import Flask, render_template, request, send_file, abort

from core.normalize import normalize_template
from core.extract import extract_report, prettify_filename
from core.rewrite import rewrite_sections, target_words
from core.build import build_paper
from core.paginate import pdf_page_count

BASE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(BASE, "outputs")
os.makedirs(OUTDIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/convert")
def convert():
    report = request.files.get("report")
    template = request.files.get("template")
    if not report or not template or not report.filename or not template.filename:
        abort(400, "Both report and template files are required.")
    for f in (report, template):
        if not f.filename.lower().endswith(".docx"):
            abort(400, f"{f.filename} is not a .docx file.")
    try:
        max_pages = max(1, min(20, int(request.form.get("max_pages", "4"))))
    except ValueError:
        abort(400, "max_pages must be a number.")
    use_llm = request.form.get("use_llm", "on") == "on"
    form_title = request.form.get("title", "").strip()
    authors_raw = request.form.get("authors", "").strip()

    job = uuid.uuid4().hex[:8]
    rep_path = os.path.join(OUTDIR, f"{job}_report.docx")
    tpl_path = os.path.join(OUTDIR, f"{job}_template.docx")
    norm_path = os.path.join(OUTDIR, f"{job}_normalized.docx")
    out_path = os.path.join(OUTDIR, f"{job}_paper.docx")
    report.save(rep_path)
    template.save(tpl_path)

    try:
        normalize_template(tpl_path, norm_path)
        data = extract_report(rep_path)
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
        build_paper(norm_path, title, authors, abstract, keywords, sections,
                    data["tables"], data["images"], data["references"], out_path)

        pages, pdf_path = pdf_page_count(out_path)
        # One tightening pass if over budget and content can shrink.
        if pages and pages > max_pages:
            sections, coverage = rewrite_sections(
                data["sections"], int(budget * max_pages / pages * 0.9), use_llm=False)
            abstract = sections[0]["text"][:1200] if sections else ""
            build_paper(norm_path, title, authors, abstract, keywords, sections,
                        data["tables"][:6], data["images"][:4],
                        data["references"], out_path)
            pages2, pdf_path2 = pdf_page_count(out_path)
            if pages2:
                pages, pdf_path = pages2, pdf_path2

        kept_tables = min(len(data["tables"]), 10 if not (pages and pages > max_pages) else 6)
        coverage_report = {
            "report_words": data["words"],
            "sections_kept": len(sections),
            "tables_total": len(data["tables"]),
            "tables_included": kept_tables,
            "tables_deferred": max(0, len(data["tables"]) - kept_tables),
            "images_total": len(data["images"]),
            "references_kept": min(len(data["references"]), 40),
            "pages": pages,
            "max_pages": max_pages,
            "within_budget": bool(pages and pages <= max_pages),
            "llm_used": coverage["llm_used"],
            "fallback_used": coverage["fallback_used"],
            "pdf_error": None if pages else pdf_path,
        }
    except Exception as e:  # noqa: BLE001
        app.logger.exception("conversion failed")
        return render_template("error.html", message=str(e)[:500]), 500
    return render_template("result.html", job=job,
                           out_file=os.path.basename(out_path),
                           coverage=coverage_report,
                           sections=coverage["sections"],
                           title=title, autofilled=autofilled)


@app.get("/download/<name>")
def download(name):
    path = os.path.join(OUTDIR, name)
    if not os.path.isfile(path) or ".." in name:
        abort(404)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5057, debug=False)
