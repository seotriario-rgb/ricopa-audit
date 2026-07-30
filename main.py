#!/usr/bin/env python3
"""
FastAPI web app for RICOPA SEO Audit generation.
v4: Data sampling + dropdown-based manual assignment when auto-detection fails.
"""

import os, shutil, zipfile, uuid, json
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse, FileResponse

from build_audit import build_audit, _detect_assignments, _hint_from_filename, _sample_detect, _get_columns, SHEET_ORDER

app = FastAPI(title="Auditoria SEO RICOPA")

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "_plantillas"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

ALL_SHEETS = [s for s in SHEET_ORDER if s not in ("Matriz de entendimiento", "Resumen")]

# Categorize sheets by RICOPA for grouped dropdown
SHEET_CATEGORIES = {
    "Rastreo": ["Errores 404", "Detalle Errores 404", "Redirecciones 3xx", "Bucles y cadenas redirección",
                 "URLs HTTP (no HTTPS)", "URLs no ASCII", "PS Minificar CSS", "PS Minificar JS"],
    "Indexación": ["Canónica — Falta", "Canónica — Múltiple", "Canónica — Errores",
                    "Directivas — Noindex", "Directivas — Nofollow",
                    "PS Visualización fuentes", "PS Solicitudes LCP", "PageSpeed — CLS"],
    "On Page": ["Metatítulo — Falta", "Metatítulo — Duplicado", "Metatítulo — Múltiple",
                 "Metatítulo — Largo", "Metatítulo — Corto", "Metatítulo = H1",
                 "Metadescripción — Falta", "Metadescripción — Duplicada",
                 "Metadescripción — Larga", "Metadescripción — Corta", "Metadescripción — Múltiple",
                 "Metadatos propuestos"],
    "Contenido": ["H1 — Falta", "H1 — Múltiple", "H1 — Duplicado", "H1 — Largo (+70)",
                   "Poco contenido", "Imágenes +100kb", "Detalle Imágenes +100kb",
                   "Imágenes sin ALT text", "Detalle Imágenes sin ALT",
                   "Imágenes sin atributo tamaño", "Detalle Imágenes sin size attr",
                   "Imágenes dim incorrectas", "Detalle Img dim incorrectas"],
    "PS (resto)": ["PS Mejorar entrega imágenes"],
}
# Flatten: all categorized sheets
_categorized = set()
for sheets in SHEET_CATEGORIES.values():
    _categorized.update(sheets)
# Add any remaining from ALL_SHEETS not categorized
for s in ALL_SHEETS:
    if s not in _categorized:
        if "PS" in s or "Pagespeed" in s or "PageSpeed" in s:
            SHEET_CATEGORIES.setdefault("Indexación", []).append(s)
        else:
            SHEET_CATEGORIES.setdefault("Contenido", []).append(s)

# Grouped options for the dropdown: special "Todo" derivations + individual sheets
DROPDOWN_OPTIONS = [
    ("", "-- Seleccionar hoja --", ""),
    ("__derive_title", "📋 Títulos (Todo) → derivar Falta, Largo, Corto, Duplicado", "Metatítulo — Falta"),
    ("__derive_meta", "📋 Meta descripción (Todo) → derivar Falta, Larga, Corta, Duplicada", "Metadescripción — Falta"),
    ("__derive_h1", "📋 H1 (Todo) → derivar Falta, Largo, Duplicado", "H1 — Falta"),
    ("__derive_images", "📋 Imágenes (Todo) → derivar +100kb, sin ALT, sin size", "Imágenes +100kb"),
    ("", "─── Hojas individuales ───", ""),
] + [(s, s, s) for s in ALL_SHEETS]


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
        return HTMLResponse(content="<h3>Error: el archivo debe ser un .zip</h3>", status_code=400)

    job_id = uuid.uuid4().hex[:8]
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Extract ZIP
        data_dir = job_dir / "data"
        data_dir.mkdir(exist_ok=True)
        zip_path = job_dir / "upload.zip"
        with open(zip_path, "wb") as f:
            shutil.copyfileobj(files.file, f)

        # Also save a copy of the raw upload for re-processing
        raw_zip = job_dir / "raw.zip"
        shutil.copy2(zip_path, raw_zip)

        with zipfile.ZipFile(zip_path) as zf:
            for member in zf.namelist():
                base = os.path.basename(member)
                if base.startswith(".") or base.startswith("__") or not base:
                    continue
                if member.endswith("/"):
                    continue
                target = data_dir / base
                with zf.open(member) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)

        # Load mapping
        custom_mapping = {}
        mapping_path = data_dir / "mapping.json"
        if mapping_path.exists():
            with open(mapping_path, encoding="utf-8") as f:
                custom_mapping = json.load(f)
        elif mapping_file and mapping_file.filename:
            raw = await mapping_file.read()
            custom_mapping = json.loads(raw)

        # Detect
        assignments, unmatched = _detect_assignments(str(data_dir), custom_mapping)

        # Save detection info for assign endpoint
        with open(job_dir / "detection.json", "w", encoding="utf-8") as f:
            json.dump({"assignments": {s: p for s, (h, cm, p) in assignments.items()},
                       "unmatched": unmatched, "client": client, "month": month, "year": year}, f)

        # If unmatched and no mapping → show dropdown help
        if unmatched and not custom_mapping:
            return _render_help_screen(job_id, client, month, year, unmatched)

        # All matched → build directly
        return await _build_and_render(job_id, client, month, year, custom_mapping)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return HTMLResponse(content=f"<h3>Error: {e}</h3>", status_code=500)


@app.post("/audit/assign/{job_id}", response_class=HTMLResponse)
async def process_assign(request: Request, job_id: str):
    """Receive manual file->sheet assignments from the dropdown help screen."""
    job_dir = OUTPUT_DIR / job_id
    if not job_dir.exists():
        return HTMLResponse(content="<h3>Job expirado — volve a subir el ZIP</h3>", status_code=404)

    # Parse form data
    form = await request.form()
    client = form.get("client", "")
    month = form.get("month", "")
    year = form.get("year", "")

    # Build custom mapping from form fields
    mapping = {}
    detection_path = job_dir / "detection.json"
    if detection_path.exists():
        with open(detection_path, encoding="utf-8") as f:
            det = json.load(f)
            client = det.get("client", client)
            month = det.get("month", month)
            year = det.get("year", year)
            # Pre-fill with auto-detected
            for sheet, path in det.get("assignments", {}).items():
                fname = os.path.basename(path)
                mapping[fname] = sheet

    for key, value in form.items():
        if key.startswith("assign__") and value.strip():
            fname = key[len("assign__"):]
            # Handle __derive_* special options
            derive_map = {
                "__derive_title": "title",
                "__derive_meta": "meta",
                "__derive_h1": "h1",
                "__derive_images": "images",
            }
            if value.strip().startswith("__derive_"):
                elem = derive_map.get(value.strip())
                if elem:
                    from build_audit import ELEMENT_FINGERPRINTS, _detect_element_type, _read_file, _derive_subsheets, _filter_by_domain, _dedup_by_key
                    # Mark this as a Todo file for derivation
                    mapping[fname] = f"__derive_{elem}"
            else:
                mapping[fname] = value.strip()

    if not mapping:
        return HTMLResponse(content="<h3>No se seleccionaron hojas — volve e intentalo de nuevo</h3>", status_code=400)

    return await _build_and_render(job_id, client, month, year, mapping)


async def _build_and_render(job_id, client, month, year, mapping):
    """Internal: build Excel + PPT from saved data + mapping, return result HTML."""
    job_dir = OUTPUT_DIR / job_id
    data_dir = job_dir / "data"

    # Re-extract if needed (first extraction might have gotten partial)
    raw_zip = job_dir / "raw.zip"
    if raw_zip.exists() and not list(data_dir.glob("*")):
        raw_zip_path = job_dir / "raw.zip"
        with zipfile.ZipFile(raw_zip_path) as zf:
            for member in zf.namelist():
                base = os.path.basename(member)
                if base.startswith(".") or not base:
                    continue
                if member.endswith("/"):
                    continue
                target = data_dir / base
                with zf.open(member) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)

    # Save mapping for this attempt
    # Pre-process __derive_* entries: read files, derive sub-sheets, add to mapping
    derive_files = {k: v for k, v in mapping.items() if isinstance(v, str) and v.startswith("__derive_")}
    for fname, derive_type in derive_files.items():
        fpath = data_dir / fname
        if not fpath.exists():
            continue
        from build_audit import _read_file, _derive_subsheets, _filter_by_domain, _dedup_by_key
        elem = derive_type.replace("__derive_", "")
        domain = _guess_domain(data_dir, client)
        rows = _read_file(fpath)
        if rows:
            rows = _dedup_by_key(rows, "Dirección")
            derived = _derive_subsheets(rows, elem)
            for sheet_name, (headers, rows_out) in derived.items():
                # Write derived rows to temporary NDJSON
                tmp_f = job_dir / f"__derived__{sheet_name}.ndjson"
                with open(tmp_f, "w", encoding="utf-8") as tf:
                    for row in rows_out:
                        d = {h: row[i] if i < len(row) else "" for i, h in enumerate(headers)}
                        json.dump(d, tf, ensure_ascii=False)
                        tf.write("\n")
                # Point mapping to temp file
                mapping[tmp_f.name] = sheet_name
        # Remove the __derive_ placeholder
        del mapping[fname]
    
    with open(data_dir / "mapping.json", "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    # Copy templates
    xlsx_out = job_dir / f"{client} - Auditoria Seo - {month} {year}.xlsx"
    pptx_out = job_dir / f"{client} - Auditoria Seo - {month} {year}.pptx"
    if not xlsx_out.exists():
        shutil.copy2(TEMPLATES_DIR / "UNA SEO - AUDITORIA - plantilla.xlsx", xlsx_out)
    if not pptx_out.exists():
        shutil.copy2(TEMPLATES_DIR / "UNA SEO - AUDITORIA - plantilla.pptx", pptx_out)

    # Guess domain
    domain = _guess_domain(data_dir, client)

    # Build
    result = build_audit(
        data_dir=str(data_dir),
        xlsx_path=str(xlsx_out),
        pptx_path=str(pptx_out),
        client=client,
        domain=domain,
        skip_ppt=False,
        skip_metadatos=False,
        mapping=mapping,
    )

    counts = result["counts"]
    unmatched = result["unmatched"]
    total = result["total"]

    # Render
    summary_rows = ""
    for sheet, n in sorted(counts.items()):
        if n > 0:
            summary_rows += f"<tr><td>{sheet}</td><td style='text-align:right'>{n:,}</td></tr>"

    unmatched_html = ""
    if unmatched:
        unmatched_html = "<div style='margin-top:16px;background:#1e293b;border-radius:8px;padding:12px'><strong style='color:#f59e0b'>Archivos no reconocidos:</strong><ul style='color:#94a3b8;font-size:12px'>"
        for u in unmatched:
            unmatched_html += f"<li>{u['file']} — columnas: {', '.join(u['columns'][:5])}</li>"
        unmatched_html += "</ul><p style='color:#94a3b8;font-size:11px'>Volve a subir con un mapping.json</p></div>"

    result_html = (BASE_DIR / "templates" / "result.html").read_text(encoding="utf-8")
    result_html = result_html.replace("{{CLIENT}}", client)
    result_html = result_html.replace("{{MES}}", month)
    result_html = result_html.replace("{{ANO}}", year)
    result_html = result_html.replace("{{AÑO}}", year)
    result_html = result_html.replace("{{JOB_ID}}", job_id)
    result_html = result_html.replace("{{TOTAL}}", f"{total:,}")
    result_html = result_html.replace("{{SUMMARY_ROWS}}", summary_rows)
    result_html = result_html.replace("{{XLSX_NAME}}", os.path.basename(xlsx_out))
    result_html = result_html.replace("{{PPTX_NAME}}", os.path.basename(pptx_out))
    result_html = result_html.replace("{{UNMATCHED}}", unmatched_html)
    result_html = result_html.replace("{{MAPPING_NAME}}", f"mapping_{client.lower().replace(' ','_')}.json")

    return HTMLResponse(content=result_html)


def _render_help_screen(job_id, client, month, year, unmatched):
    """Show unmatched files with dropdowns for manual sheet assignment."""
    # Build grouped dropdown options
    options_html = '<option value="">-- Seleccionar hoja --</option>'
    options_html += '<optgroup label="📋 Archivos Todo (derivan varias hojas)">'
    for value, label in [("__derive_title", "Títulos → Falta, Largo, Corto, Duplicado"),
                          ("__derive_meta", "Meta descripción → Falta, Larga, Corta, Duplicada"),
                          ("__derive_h1", "H1 → Falta, Largo, Duplicado"),
                          ("__derive_images", "Imágenes → +100kb, sin ALT, sin size")]:
        options_html += f'<option value="{value}">{label}</option>'
    options_html += '</optgroup>'
    
    for cat, sheets in SHEET_CATEGORIES.items():
        options_html += f'<optgroup label="── {cat} ──">'
        for sheet in sheets:
            options_html += f'<option value="{sheet}">{sheet}</option>'
        options_html += '</optgroup>'
    
    rows = ""
    for i, u in enumerate(unmatched):
        hint = u.get("hint") or ""
        hint_selected = f'<option value="{hint}" selected style="display:none">{hint} (sugerido)</option>' if hint else ""
        pre_options = f'<option value="">-- Seleccionar hoja --</option>{hint_selected}' if hint else ''
        rows += f"""
        <tr>
            <td style='font-size:12px;word-break:break-all;max-width:400px'>📄 {u['file']}</td>
            <td>
                <select name="assign__{u['file']}" style='padding:6px 10px;border-radius:6px;background:#0f172a;color:#e2e8f0;border:1px solid #334155;width:100%;min-width:260px'>
                    {options_html}
                </select>
            </td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Identificar archivos — {client}</title>
<style>
    :root {{ --bg: #0f172a; --card: #1e293b; --border: #334155; --text: #e2e8f0; --muted: #94a3b8; --accent: #3b82f6; }}
    * {{ margin:0; padding:0; box-sizing:border-box; }}
    body {{ font-family: -apple-system, sans-serif; background: var(--bg); color: var(--text); padding: 20px; }}
    .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 28px; max-width: 850px; margin: 0 auto; }}
    h2 {{ font-size: 20px; margin-bottom: 6px; }} h2 span {{ color: #f59e0b; }}
    .muted {{ color: var(--muted); font-size: 13px; margin-bottom: 20px; }}
    table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
    th, td {{ padding: 10px 8px; border-bottom: 1px solid var(--border); font-size: 13px; }}
    th {{ color: var(--muted); text-align: left; }}
    button {{ padding: 12px 28px; background: var(--accent); color: white; border: none; border-radius: 8px; font-weight: 600; cursor: pointer; font-size: 14px; }}
    button:hover {{ opacity: 0.9; }}
    .info {{ background: #1e3a5f; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-size: 12px; color: #60a5fa; }}
</style></head>
<body>
<div class="card">
    <h2><span>{len(unmatched)} archivos</span> necesitan ayuda</h2>
    <p class="muted">{client} — {month} {year} • Selecciona a que hoja pertenece cada archivo</p>
    <div class="info">Las hojas estan agrupadas por RICOPA. Usa "Archivos Todo" para CSVs que contienen todos los datos de un elemento y necesitan derivarse en varias hojas. Varios archivos pueden ir a la misma hoja (ej: Falta texto ALT + Falta atributo ALT → Imagenes sin ALT text).</div>
    <form action="/audit/assign/{job_id}" method="POST">
        <input type="hidden" name="client" value="{client}">
        <input type="hidden" name="month" value="{month}">
        <input type="hidden" name="year" value="{year}">
        <table>
            <tr><th>Archivo</th><th>Asignar a hoja</th></tr>
            {rows}
        </table>
        <button type="submit">Confirmar y generar auditoria</button>
    </form>
</div>
</body></html>"""
    return HTMLResponse(content=html)


def _guess_domain(data_dir: Path, client: str) -> str:
    try:
        for fpath in data_dir.glob("*.ndjson"):
            with open(fpath, encoding="utf-8") as f:
                for line in f:
                    if '"Dirección"' in line:
                        d = json.loads(line.strip())
                        url = d.get("Dirección") or ""
                        from urllib.parse import urlparse
                        host = urlparse(url).hostname or ""
                        if host:
                            parts = host.split(".")
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
        # Return mapping.json from data dir
        data_dir = job_dir / "data"
        mapping = data_dir / "mapping.json"
        if mapping.exists():
            return FileResponse(mapping, filename="mapping.json", media_type="application/json")
        # Legacy: check job dir
        mapping2 = job_dir / "mapping.json"
        if mapping2.exists():
            return FileResponse(mapping2, filename="mapping.json", media_type="application/json")

    for f in job_dir.iterdir():
        if file_type == "xlsx" and f.suffix == ".xlsx":
            return FileResponse(f, filename=os.path.basename(f))
        if file_type == "pptx" and f.suffix == ".pptx":
            return FileResponse(f, filename=os.path.basename(f))

    return HTMLResponse(content="<h3>Archivo no encontrado</h3>", status_code=404)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
