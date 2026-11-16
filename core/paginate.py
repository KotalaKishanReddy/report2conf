"""Page measurement via headless LibreOffice + pypdf."""
import os
import subprocess
import tempfile

from pypdf import PdfReader


def pdf_page_count(docx_path: str, timeout: int = 120) -> tuple[int | None, str | None]:
    """Convert docx to PDF headlessly; return (pages, pdf_path)."""
    outdir = tempfile.mkdtemp(prefix="r2c_pdf_")
    profile = tempfile.mkdtemp(prefix="r2c_profile_")
    try:
        proc = subprocess.run(
            ["soffice", "--headless",
             f"-env:UserInstallation=file://{profile}",
             "--convert-to", "pdf", "--outdir", outdir, docx_path],
            capture_output=True, text=True, timeout=timeout,
        )
        base = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
        pdf_path = os.path.join(outdir, base)
        if not os.path.exists(pdf_path):
            return None, proc.stderr[-2000:] or proc.stdout[-2000:]
        reader = PdfReader(pdf_path)
        return len(reader.pages), pdf_path
    except Exception as e:  # noqa: BLE001
        return None, str(e)
    finally:
        import shutil

        shutil.rmtree(profile, ignore_errors=True)
