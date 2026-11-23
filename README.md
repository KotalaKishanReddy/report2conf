# report2conf

Upload a long report (`.docx`, `.pdf`, or image) plus an IEEE conference
template `.docx`, pick a maximum page count, and download a filled
conference paper (editable + strict rebuild). An online editor lets you
fix the text before the final download.

## Start

```bash
cd /home/kotalakishanreddy/muse/report2conf
python3 app.py
```

Open [http://127.0.0.1:5057](http://127.0.0.1:5057) in a browser.

## How it works

1. The template is normalized from strict OOXML to the transitional form
   `python-docx` can edit, preserving its styles.
2. The report is read by type: `.docx` via `python-docx`; digital `.pdf`
   via `pypdf`; scanned PDFs/images via local Tesseract OCR (offline,
   needs the `tesseract` binary — without it scanned PDFs error and images
   build with a placeholder). The OCR path recovers text only; tables and
   references from scanned PDFs/images are not carried over.
   Optional `baidu/Unlimited-OCR-2` backend, enabled explicitly with
   `REPORT2CONF_OCR=unlimited` and model weights installed (uses GPU when
   present, CPU otherwise), does one-shot multi-page parsing instead.
3. Each section is rewritten: extractive compressor by default (free + safe,
   keeps sentences with numbers). Tick "Use cloud AI rewrite" only when an
   OpenAI-compatible `OPENAI_API_KEY` is set (`REPORT2CONF_MODEL`, default
   `gpt-4o-mini`).
4. The paper is built with the template's styles, then shipped twice: the
   editable transitional `.docx` plus a `_strict.docx` strict-OOXML rebuild.
   Page count is measured with headless LibreOffice
   (`soffice --headless --convert-to pdf`); if it exceeds the page limit,
   one tightening pass runs automatically.
5. Anything that cannot fit is listed on the result page, never dropped
   silently. Use Edit online to fix title/authors/abstract/sections and
   rebuild (saving reuses the edited text as-is, with no second rewrite or
   tightening pass).

## Limits

- At most 10 tables / 8 images / 40 references (6 / 4 after the tightening
  pass); `max_pages` is clamped to 1–20, default 4.
- The tightening pass forces the extractive rewrite even if cloud rewrite
  was ticked.
- `tests/test_basic.py` and `tests/test_stress.py` need two sample files
  under `~/Downloads`; `tests/test_ocr_strict.py` runs anywhere.

## Test

```bash
cd /home/kotalakishanreddy/muse/report2conf
python3 -m pytest tests/ -q
```
