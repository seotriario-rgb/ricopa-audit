#!/usr/bin/env python3
"""
FastAPI web app for RICOPA SEO Audit generation.
Users upload NDJSON zip → get Excel + PPT deliverables.
"""

import os, shutil, zipfile, uuid, tempfile
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from build_audit import build_audit

app = FastAPI(title="Auditoría SEO RICOPA")

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "_plantillas"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    return (BASE_DIR / "templates" / "upload.html").read_text(encoding="utf-8")


@app.post("/audit", response_class=HTMLResponse)
async def process_audit(
    request: Request,
    client: str = Form(...),
    month: str = Form("Julio"),
    year: str = Form("2026"),
    files: UploadFile = File(...),
):
    # Validate
    if not files.filename or not files.filename.endswith(".zip"):
        return HTMLResponse(content="<h3>Error: el archivo debe ser un .zip con los archivos NDJSON</h3>", status_code=400)

    job_id = uuid.uuid4().hex[:8]
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    errors = []
    try:
        # 1. Save and extract zip with NDJSONs
        zip_path = job_dir / "upload.zip"
        with open(zip_path, "wb") as f:
            shutil.copyfileobj(files.file, f)

        data_dir = job_dir / "data"
        data_dir.mkdir(exist_ok=True)
        with zipfile.ZipFile(zip_path) as zf:
            for member in zf.namelist():
                # Skip __MACOSX, .DS_Store, directories
                base = os.path.basename(member)
                if base.startswith(".") or base.startswith("__") or not base:
                    continue
                if member.endswith("/"):
                    continue
                zf.extract(member, data_dir)

        # Flatten: if files are inside a subfolder, move them to data_dir root
        for root, dirs, fnames in os.walk(data_dir):
            for fname in fnames:
                src = os.path.join(root, fname)
                dst = os.path.join(data_dir, fname)
                if src != dst:
                    shutil.move(src, dst)

        # 2. Copy templates
        xlsx_out = job_dir / f"{client} - Auditoria Seo - {month} {year}.xlsx"
        pptx_out = job_dir / f"{client} - Auditoria Seo - {month} {year}.pptx"
        shutil.copy2(TEMPLATES_DIR / "UNA SEO - AUDITORIA - plantilla.xlsx", xlsx_out)
        shutil.copy2(TEMPLATES_DIR / "UNA SEO - AUDITORIA - plantilla.pptx", pptx_out)

        # 3. Extract domain from client name or guess from NDJSON
        domain = ""
        try:
            ndjson_files = list(data_dir.glob("*.ndjson"))
            if ndjson_files:
                with open(ndjson_files[0], encoding="utf-8") as f:
                    for line in f:
                        if '"Dirección"' in line or '"Fuente"' in line:
                            import re, json
                            try:
                                d = json.loads(line.strip())
                                url = d.get("Dirección") or d.get("Fuente") or ""
                                from urllib.parse import urlparse
                                host = urlparse(url).hostname or ""
                                if host:
                                    parts = host.split(".")
                                    if len(parts) >= 2:
                                        domain = ".".join(parts[-3:]) if len(parts) >= 3 and parts[-3] != "www" else ".".join(parts[-2:])
                                    break
                            except Exception:
                                continue
        except Exception:
            pass
        if not domain:
            domain = client.lower().replace(" ", "") + ".com"

        # 4. Build audit
        counts = build_audit(
            data_dir=str(data_dir),
            xlsx_path=str(xlsx_out),
            pptx_path=str(pptx_out),
            client=client,
            domain=domain,
            skip_ppt=False,
            skip_metadatos=False,
        )

        # 5. Return result page
        total = sum(counts.values())
        summary_rows = ""
        for sheet, n in sorted(counts.items()):
            if n > 0:
                summary_rows += f"<tr><td>{sheet}</td><td style='text-align:right'>{n:,}</td></tr>"

        result_html = (BASE_DIR / "templates" / "result.html").read_text(encoding="utf-8")
        result_html = result_html.replace("{{CLIENT}}", client)
        result_html = result_html.replace("{{MES}}", month)
        result_html = result_html.replace("{{AÑO}}", year)
        result_html = result_html.replace("{{JOB_ID}}", job_id)
        result_html = result_html.replace("{{TOTAL}}", f"{total:,}")
        result_html = result_html.replace("{{SUMMARY_ROWS}}", summary_rows)
        result_html = result_html.replace("{{XLSX_NAME}}", os.path.basename(xlsx_out))
        result_html = result_html.replace("{{PPTX_NAME}}", os.path.basename(pptx_out))

        return HTMLResponse(content=result_html)

    except Exception as e:
        errors.append(str(e))
        return HTMLResponse(content=f"<h3>Error procesando la auditoría</h3><pre>{e}</pre>", status_code=500)


@app.get("/download/{job_id}/{file_type}")
async def download_file(job_id: str, file_type: str):
    job_dir = OUTPUT_DIR / job_id
    if not job_dir.exists():
        return HTMLResponse(content="<h3>Archivo no encontrado. El job puede haber expirado.</h3>", status_code=404)

    for f in job_dir.iterdir():
        if file_type == "xlsx" and f.suffix == ".xlsx":
            return FileResponse(f, filename=os.path.basename(f))
        if file_type == "pptx" and f.suffix == ".pptx":
            return FileResponse(f, filename=os.path.basename(f))

    return HTMLResponse(content="<h3>Archivo no encontrado</h3>", status_code=404)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
