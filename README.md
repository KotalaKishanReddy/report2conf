# report2conf

Upload a long report `.docx` plus an IEEE conference template `.docx`, pick a
maximum page count, and download a filled conference paper.

## Start

```bash
cd /home/kotalakishanreddy/muse/report2conf
python3 app.py
```

Open [http://127.0.0.1:5057](http://127.0.0.1:5057) in a browser.

## How it works

1. The template is normalized from strict OOXML to the transitional form
   `python-docx` can edit, preserving its styles.
2. The report is split into sections, tables, images, and references.
3. Each section is rewritten: an OpenAI-compatible API when `OPENAI_API_KEY`
   is set (`REPORT2CONF_MODEL`, default `gpt-4o-mini`), otherwise a
   deterministic extractive compressor that keeps sentences with numbers.
4. The paper is built with the template's styles and measured with headless
   LibreOffice (`soffice --headless --convert-to pdf`); if it exceeds the
   page limit, one tightening pass runs automatically.
5. Anything that cannot fit is listed on the result page, never dropped
   silently.

## Test

```bash
cd /home/kotalakishanreddy/muse/report2conf
python3 -m pytest tests/ -q
```
