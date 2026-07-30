#!/usr/bin/env python3
"""
FastAPI web app for RICOPA SEO Audit generation.
v2: Auto-detects CSV/NDJSON files by column signatures.
    Supports mapping.json for training custom file→sheet assignments.
"""

import os, shutil, zipfile, uuid, json, tempfile
from pathlib import Path
from datetime import datetime

from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

from build_audit import build_audit, _get_columns, _hint_from_filename, SIGNATURES

app = FastAPI(title="Auditoría SEO RICOPA")

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "_plantillas"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


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
    mapping_file: Optional[UploadFile] = File(None),
):
    if not files.filename or not files.filename.endswith(".zip"):
        return HTMLResponse(content="<h3>Error: el archivo debe ser un .zip con los archivos CSV/NDJSON</h3>", status_code=400)

    job_id = uuid.uuid4().hex[:8]
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    try:
        # 1. Extract ZIP
        zip_path = job_dir / "upload.zip"
        with open(zip_path, "wb") as f:
            shutil.copyfileobj(files.file, f)

        data_dir = job_dir / "data"
        data_dir.mkdir(exist_ok=True)
        with zipfile.ZipFile(zip_path) as zf:
            for member in zf.namelist():
                base = os.path.basename(member)
                if base.startswith(".") or base.startswith("__") or not base:
                    continue
                if member.endswith("/"):
                    continue
                # Extract to data_dir root
                target = data_dir / base
                with zf.open(member) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)

        # 2. Load mapping.json if provided
        custom_mapping = {}
        mapping_path = data_dir / "mapping.json"
        if mapping_path.exists():
            with open(mapping_path, encoding="utf-8") as f:
                custom_mapping = json.load(f)
        elif mapping_file and mapping_file.filename:
            raw = await mapping_file.read()
            custom_mapping = json.loads(raw)

        # 3. Copy templates
        xlsx_out = job_dir / f"{client} - Auditoria Seo - {month} {year}.xlsx"
        pptx_out = job_dir / f"{client} - Auditoria Seo - {month} {year}.pptx"
        shutil.copy2(TEMPLATES_DIR / "UNA SEO - AUDITORIA - plantilla.xlsx", xlsx_out)
        shutil.copy2(TEMPLATES_DIR / "UNA SEO - AUDITORIA - plantilla.pptx", pptx_out)

        # 4. Guess domain
        domain = _guess_domain(data_dir, client)

        # 5. Pre-scan: detect unmatched files
        pre_assignments, pre_unmatched = _prescan(data_dir, custom_mapping)

        # 6. If unmatched files exist and no mapping, show help screen
        if pre_unmatched and not custom_mapping:
            return _render_help_screen(client, month, year, pre_unmatched)

        # 7. Build audit
        result = build_audit(
            data_dir=str(data_dir),
            xlsx_path=str(xlsx_out),
            pptx_path=str(pptx_out),
            client=client,
            domain=domain,
            skip_ppt=False,
            skip_metadatos=False,
            mapping=custom_mapping,
        )
        counts = result["counts"]
        unmatched = result["unmatched"]
        total = result["total"]

        # 8. Save mapping.json
        mapping_out = {}
        for sheet_name in sorted(counts.keys()):
            for fname in os.listdir(data_dir):
                fpath = data_dir / fname
                if fname.startswith("."):
                    continue
                cols_orig = _get_columns(fpath)
                if not cols_orig:
                    continue
                # Find if this file → this sheet
                for sig, entries in SIGNATURES.items():
                    for e in entries:
                        if e["sheet"] == sheet_name and e["sheet"] is not None:
                            mapping_out[fname] = sheet_name
                            break
        # Add custom overrides
        mapping_out.update(custom_mapping)
        mapping_json = job_dir / "mapping.json"
        with open(mapping_json, "w", encoding="utf-8") as f:
            json.dump(mapping_out, f, ensure_ascii=False, indent=2)

        # 9. Render result
        summary_rows = ""
        for sheet, n in sorted(counts.items()):
            if n > 0:
                summary_rows += f"<tr><td>{sheet}</td><td style='text-align:right'>{n:,}</td></tr>"

        unmatched_html = ""
        if unmatched:
            unmatched_html = "<div style='margin-top:16px;background:#1e293b;border-radius:8px;padding:12px'><strong style='color:#f59e0b'>⚠ Archivos no reconocidos:</strong><ul style='color:#94a3b8;font-size:12px'>"
            for u in unmatched:
                unmatched_html += f"<li>{u['file']} — columnas: {', '.join(u['columns'][:5])}</li>"
            unmatched_html += "</ul></div>"

        result_html = (BASE_DIR / "templates" / "result.html").read_text(encoding="utf-8")
        result_html = result_html.replace("{{CLIENT}}", client)
        result_html = result_html.replace("{{MES}}", month)
        result_html = result_html.replace("{{AÑO}}", year)
        result_html = result_html.replace("{{JOB_ID}}", job_id)
        result_html = result_html.replace("{{TOTAL}}", f"{total:,}")
        result_html = result_html.replace("{{SUMMARY_ROWS}}", summary_rows)
        result_html = result_html.replace("{{XLSX_NAME}}", os.path.basename(xlsx_out))
        result_html = result_html.replace("{{PPTX_NAME}}", os.path.basename(pptx_out))
        result_html = result_html.replace("{{UNMATCHED}}", unmatched_html)
        result_html = result_html.replace("{{MAPPING_NAME}}", f"mapping_{client.lower().replace(' ','_')}.json")

        return HTMLResponse(content=result_html)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return HTMLResponse(content=f"<h3>Error procesando la auditoría</h3><pre>{e}</pre>", status_code=500)




def _render_help_screen(client, month, year, unmatched):
    """Show which files couldn't be auto-detected so the user can create a mapping.json."""
    rows = ""
    for i, u in enumerate(unmatched):
        hint = u.get("hint") or ""
        rows += f"""
        <tr>
            <td style='font-size:12px;max-width:300px;word-break:break-all'>{u['file']}</td>
            <td style='font-size:11px;color:#94a3b8'>{', '.join(u['columns'][:4])}</td>
            <td style='font-size:11px;color:#60a5fa'>{hint or '—'}</td>
        </tr>"""
    
    # Generate mapping template JSON
    mapping_template = {}
    for u in unmatched:
        hint = u.get("hint") or "AQUI_NOMBRE_DE_LA_HOJA"
        mapping_template[u["file"]] = hint
    mapping_json = json.dumps(mapping_template, indent=2, ensure_ascii=False)
    
    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Archivos no reconocidos — {client}</title>
<style>
    :root {{ --bg: #0f172a; --card: #1e293b; --border: #334155; --text: #e2e8f0; --muted: #94a3b8; --accent: #3b82f6; }}
    * {{ margin:0; padding:0; box-sizing:border-box; }}
    body {{ font-family: -apple-system, sans-serif; background: var(--bg); color: var(--text); padding: 20px; }}
    .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 24px; max-width: 700px; margin: 0 auto; }}
    h2 {{ margin-bottom: 8px; }} h2 span {{ color: #f59e0b; }}
    .muted {{ color: var(--muted); font-size: 13px; margin-bottom: 20px; }}
    table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
    th, td {{ padding: 8px; border-bottom: 1px solid var(--border); font-size: 13px; }}
    th {{ color: var(--muted); text-align: left; }}
    code {{ background: #1e293b; padding: 2px 6px; border-radius: 4px; font-size: 11px; color: #e2e8f0; }}
    .step {{ background: #1e3a5f; border-radius: 8px; padding: 14px; margin-bottom: 12px; font-size: 12px; }}
    .step strong {{ color: #60a5fa; }}
    a {{ color: var(--accent); }}
</style></head>
<body>
<div class="card">
    <h2>Archivos sin <span>identificar</span></h2>
    <p class="muted">{client} — {month} {year} · {len(unmatched)} archivos no pudieron asignarse automáticamente</p>
    
    <div class="step">
        <strong>1. Copiá este JSON en un archivo llamado <code>mapping.json</code>:</strong><br>
        <textarea readonly style='width:100%;height:120px;margin-top:8px;background:#0f172a;color:#e2e8f0;border:1px solid #334155;border-radius:4px;padding:8px;font-size:11px;font-family:monospace'>""" + mapping_json + """</textarea>
    </div>
    
    <div class="step">
        <strong>2. Editá cada valor con el nombre de la hoja correcta</strong> (ej: "Metatítulo — Falta", "PS Minificar JS").<br>
        Las hojas válidas son: <span style="color:#94a3b8">Errores 404, Redirecciones 3xx, Metatítulo — Falta, Metadescripción — Falta, H1 — Falta, PS Minificar JS, ...</span>
    </div>
    
    <div class="step">
        <strong>3. Volvé a subir el ZIP incluyendo el <code>mapping.json</code></strong> y la app lo reconocerá automáticamente.
    </div>

    <table>
        <tr><th>Archivo</th><th>Columnas detectadas</th><th>Posible hoja</th></tr>
        {rows}
    </table>

    <p style="color:var(--muted);font-size:12px">Mientras tanto, estos archivos se omitirán del Excel. <a href="/">← Volver</a></p>
</div>
</body></html>"""
    return HTMLResponse(content=html)


def _prescan(data_dir: Path, mapping: dict):
    """Pre-scan to detect file→sheet assignments without building."""
    assignments = {}
    unmatched = []
    data_dir = Path(data_dir)
    
    # Apply mapping first
    fname_to_sheet = {}
    for fname, sheet in (mapping or {}).items():
        fname_to_sheet[fname.strip().lower()] = sheet
    
    for fpath in sorted(data_dir.glob("*")):
        if not fpath.is_file():
            continue
        fname = fpath.name
        if fname.startswith("."):
            continue
        ext = fpath.suffix.lower()
        if ext not in (".csv", ".ndjson", ".json"):
            continue
        
        cols = _get_columns(fpath)
        if not cols:
            continue
        
        assigned = None
        
        # 1. Mapping override
        for key, sheet in fname_to_sheet.items():
            if key in fname.lower():
                for sig, entries in SIGNATURES.items():
                    for e in entries:
                        if e["sheet"] == sheet:
                            assigned = e["sheet"]
                            break
                    if assigned:
                        break
                if assigned:
                    break
        
        # 2. Signature detection (unique)
        if not assigned:
            entries = SIGNATURES.get(cols, [])
            unique = [e for e in entries if e["sheet"] is not None]
            if len(unique) == 1:
                assigned = unique[0]["sheet"]
        
        # 3. Signature + filename hint
        if not assigned:
            entries = SIGNATURES.get(cols, [])
            valid = [e for e in entries if e["sheet"] is not None]
            if len(valid) > 1:
                hint = _hint_from_filename(fname)
                if hint and any(e["sheet"] == hint for e in valid):
                    assigned = hint
        
        # 4. Pure filename hint
        if not assigned:
            hint = _hint_from_filename(fname)
            if hint:
                assigned = hint
        
        if assigned:
            assignments[assigned] = assignments.get(assigned, 0) + 1
        else:
            unmatched.append({
                "file": fname,
                "columns": sorted(cols),
                "hint": _hint_from_filename(fname),
            })
    
    return assignments, unmatched


def _guess_domain(data_dir: Path, client: str) -> str:
    """Try to guess the client domain from data files."""
    try:
        for fpath in data_dir.glob("*.ndjson"):
            with open(fpath, encoding="utf-8") as f:
                for line in f:
                    if '"Dirección"' in line or '"Fuente"' in line:
                        d = json.loads(line.strip())
                        url = d.get("Dirección") or d.get("Fuente") or ""
                        from urllib.parse import urlparse
                        host = urlparse(url).hostname or ""
                        if host:
                            parts = host.split(".")
                            if len(parts) >= 2:
                                return ".".join(parts[-3:]) if len(parts) >= 3 and parts[-3] != "www" else ".".join(parts[-2:])
                        break
            break
    except Exception:
        pass
    return client.lower().replace(" ", "") + ".com"


@app.get("/download/{job_id}/{file_type}")
async def download_file(job_id: str, file_type: str):
    job_dir = OUTPUT_DIR / job_id
    if not job_dir.exists():
        return HTMLResponse(content="<h3>Archivo no encontrado</h3>", status_code=404)
    
    if file_type == "mapping":
        mapping = job_dir / "mapping.json"
        if mapping.exists():
            return FileResponse(mapping, filename=os.path.basename(mapping), media_type="application/json")

    for f in job_dir.iterdir():
        if file_type == "xlsx" and f.suffix == ".xlsx":
            return FileResponse(f, filename=os.path.basename(f))
        if file_type == "pptx" and f.suffix == ".pptx":
            return FileResponse(f, filename=os.path.basename(f))

    return HTMLResponse(content="<h3>Archivo no encontrado</h3>", status_code=404)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
