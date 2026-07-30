#!/usr/bin/env python3
"""
Core logic for building RICOPA SEO audit deliverables (Excel + PowerPoint).
Importable module — no CLI dependency. Called by main.py (FastAPI).

Usage:
    from build_audit import build_audit
    summary = build_audit(data_dir="path/to/ndjsons",
                          xlsx="output.xlsx",
                          pptx="output.pptx",
                          client="Avafin",
                          domain="avafin.mx")
"""

from __future__ import annotations
import sys, os, json, re, shutil
from pathlib import Path
from urllib.parse import urlparse
from datetime import date

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from pptx import Presentation as PptxPresentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

BRAND_FONT = "Gabarito"
METADATA_FILL_A = PatternFill("solid", fgColor="FFF7E6")
METADATA_FILL_B = PatternFill("solid", fgColor="E6F4FF")


# ---------- Excel helpers ----------

def _load_wb(path: str) -> openpyxl.Workbook:
    return openpyxl.load_workbook(path)

def _save_wb(wb, path: str) -> None:
    wb.save(path)

def _ensure_sheet(wb, name: str, headers: list[str]):
    if name in wb.sheetnames:
        ws = wb[name]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(title=name)
    ws.append(headers)
    for col_idx, _ in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 22
    return ws

def _fill_sheet(wb, name: str, headers: list[str], rows: list[list]) -> int:
    ws = _ensure_sheet(wb, name, headers)
    n = 0
    for r in rows:
        ws.append(r)
        n += 1
    if n == 0:
        ws.append(["Sin hallazgos"] + [""] * (len(headers) - 1))
    return n


# ---------- Metadatos propuestos ----------

def _fill_metadatos_propuestos(wb, rows: list[list]) -> None:
    name = "Metadatos propuestos"
    headers = ["URL", "Factor a corregir", "Texto actual", "Propuesta"]
    ws = _ensure_sheet(wb, name, headers)
    last_url = None
    toggle = False
    for r in rows:
        url = r[0]
        if url != last_url:
            toggle = not toggle
            last_url = url
        fill = METADATA_FILL_A if toggle else METADATA_FILL_B
        ws.append(r)
        row_idx = ws.max_row
        for col_idx in range(1, 5):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 60


# ---------- PowerPoint helpers ----------

def _set_ppt_hallazgo(pptx_path: str, slide_index_1based: int, shape_substr: str,
                      count: int, excel_filename: str, sheet_name: str) -> int:
    prs = PptxPresentation(pptx_path)
    slide = prs.slides[slide_index_1based - 1]
    n_edited = 0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if shape_substr.lower() not in shape.name.lower():
            continue
        tf = shape.text_frame
        edited_here = 0
        for para in tf.paragraphs:
            for run in list(para.runs):
                t = run.text or ""
                if re.search(r"Hay \d+", t):
                    run.text = re.sub(r"Hay \d+", f"Hay {count}", t)
                    run.text = re.sub(r"\s*VER\s*$", "", run.text)
                    n_edited += 1
                    edited_here += 1
        if not edited_here:
            tf.clear()
            p = tf.paragraphs[0]
            p.text = f"Hay {count} hallazgos."
            n_edited += 1
            edited_here += 1
        last_para = tf.paragraphs[-1]
        ver_run = last_para.add_run()
        ver_run.text = "  VER"
        ver_run.font.name = BRAND_FONT
        ver_run.font.size = Pt(14)
        try:
            ver_run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
        except Exception:
            pass
        try:
            ver_run.hyperlink.address = f"{excel_filename}#'{sheet_name}'!A1"
        except Exception:
            pass
    prs.save(pptx_path)
    return n_edited


# ---------- Data processing ----------

def _read_ndjson(path: Path) -> list[dict]:
    rows = []
    if not path.exists():
        return rows
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return rows

def _filter_by_domain(rows: list[dict], domain: str, key: str = "Dirección") -> list[dict]:
    if not domain:
        return rows
    result = []
    for r in rows:
        url = r.get(key, "")
        if not url:
            continue
        host = urlparse(url).hostname or ""
        if host.endswith(domain) or host == domain:
            result.append(r)
    return result

def _dedup_by_key(rows: list[dict], key: str = "Dirección") -> list[dict]:
    seen = set()
    result = []
    for r in rows:
        val = r.get(key, "")
        if val in seen:
            continue
        seen.add(val)
        result.append(r)
    return result


# ---------- Sheet specs ----------

SHEETS_SPEC = [
    ("Errores 404", "errors_4xx.ndjson",
     ["URL", "Código", "Tipo Contenido", "URL de redirección", "Tipo redirección"],
     ["Dirección", "Respuesta", "Tipo de contenido", "URL de redirección", "Tipo de redirección"]),
    ("URLs no ASCII", "urls_no_ascii.ndjson",
     ["URL", "Indexabilidad"],
     ["Dirección", "Indexabilidad"]),
    ("Metatítulo — Falta", "title_falta.ndjson",
     ["URL", "Indexabilidad"],
     ["Dirección", "Indexabilidad"]),
    ("Metatítulo — Duplicado", "title_dup.ndjson",
     ["URL", "Título", "Repeticiones"],
     ["Dirección", "Título 1", "Repeticiones"]),
    ("Metatítulo — Múltiple", "title_multiple.ndjson",
     ["URL", "Títulos"],
     ["Dirección", "Título 1"]),
    ("Metatítulo — Largo", "title_largo.ndjson",
     ["URL", "Título", "Caracteres", "Píxeles"],
     ["Dirección", "Título 1", "Longitud del título 1", "Ancho de píxeles del título 1"]),
    ("Metatítulo — Corto", "title_corto.ndjson",
     ["URL", "Título", "Caracteres", "Píxeles"],
     ["Dirección", "Título 1", "Longitud del título 1", "Ancho de píxeles del título 1"]),
    ("Metadescripción — Falta", "meta_falta.ndjson",
     ["URL", "Indexabilidad"],
     ["Dirección", "Indexabilidad"]),
    ("Metadescripción — Duplicada", "meta_dup.ndjson",
     ["URL", "Metadescripción", "Repeticiones"],
     ["Dirección", "Meta description 1", "Repeticiones"]),
    ("Metadescripción — Larga", "meta_larga.ndjson",
     ["URL", "Metadescripción", "Caracteres", "Píxeles"],
     ["Dirección", "Meta description 1", "Longitud de la meta description 1", "Ancho de píxeles de la meta description 1"]),
    ("Metadescripción — Corta", "meta_corta.ndjson",
     ["URL", "Metadescripción", "Caracteres", "Píxeles"],
     ["Dirección", "Meta description 1", "Longitud de la meta description 1", "Ancho de píxeles de la meta description 1"]),
    ("Metadescripción — Múltiple", "meta_multiple.ndjson",
     ["URL", "Metadescripciones"],
     ["Dirección", "Meta description 1"]),
    ("H1 — Falta", "h1_falta.ndjson",
     ["URL", "Indexabilidad"],
     ["Dirección", "Indexabilidad"]),
    ("H1 — Múltiple", "h1_multiple.ndjson",
     ["URL", "H1s"],
     ["Dirección", "H1-1"]),
    ("H1 — Duplicado", "h1_dup.ndjson",
     ["URL", "H1", "Repeticiones"],
     ["Dirección", "H1-1", "Repeticiones"]),
    ("H1 — Largo (+70)", "h1_largo.ndjson",
     ["URL", "H1"],
     ["Dirección", "H1-1"]),
    ("Poco contenido", "poco_contenido.ndjson",
     ["URL", "Indexabilidad"],
     ["Dirección", "Indexabilidad"]),
    ("Imágenes +100kb", "img_over_100kb.ndjson",
     ["URL imagen", "Tipo", "Tamaño (kB)", "Dimensiones", "# enlaces entrada"],
     ["Dirección", "Tipo de contenido", "Tamaño (Bytes)", "Dimensiones", "Enlaces de entrada de imágenes"]),
    ("Imágenes sin ALT text", "img_sin_alt.ndjson",
     ["URL imagen", "Tipo", "# páginas donde aparece"],
     ["Dirección", "Tipo de contenido", "Enlaces de entrada de imágenes"]),
    ("Imágenes sin atributo tamaño", "img_sin_size.ndjson",
     ["URL imagen", "Tipo", "Dimensiones"],
     ["Dirección", "Tipo de contenido", "Dimensiones"]),
    ("Imágenes dim incorrectas", "img_dim_incorrectas.ndjson",
     ["URL imagen", "Dimensiones declaradas", "Tipo"],
     ["Dirección", "Dimensiones", "Tipo de contenido"]),
    ("Canónica — Falta", "can_falta.ndjson",
     ["URL", "Indexabilidad"],
     ["Dirección", "Indexabilidad"]),
    ("Canónica — Múltiple", "can_multiple.ndjson",
     ["URL", "Canónicas"],
     ["Dirección", "Elemento de enlace canónico 1"]),
    ("Redirecciones 3xx", "redirects_3xx.ndjson",
     ["URL", "Código", "URL destino", "Tipo", "Enlaces internos"],
     ["Dirección", "Respuesta", "URL de redirección", "Tipo de redirección", "Enlaces internos"]),
    ("URLs HTTP (no HTTPS)", "urls_http.ndjson",
     ["URL", "Tipo"],
     ["Dirección", None]),
]

RESUMEN_SPEC = [
    ("¿Hay URLs con respuesta 4XX?", "Errores 404", "Rastreo"),
    ("¿Hay redirecciones 3XX por auditar?", "Redirecciones 3xx", "Rastreo"),
    ("¿Hay recursos HTTP (no HTTPS)?", "URLs HTTP (no HTTPS)", "Rastreo"),
    ("¿Hay bucles o cadenas de redirección?", "Bucles y cadenas redirección", "Rastreo"),
    ("URLs con caracteres no ASCII", "URLs no ASCII", "Rastreo"),
    ("Oportunidad minificar CSS", "PS Minificar CSS", "Rastreo"),
    ("Oportunidad minificar JS", "PS Minificar JS", "Rastreo"),
    ("Páginas sobredimensionadas (imágenes)", "Imágenes +100kb", "Rastreo"),
    ("Oportunidad WebP / nueva entrega imágenes", "PS Mejorar entrega imágenes", "Rastreo"),
    ("URLs sin canónica", "Canónica — Falta", "Indexación"),
    ("URLs con múltiples canónicas", "Canónica — Múltiple", "Indexación"),
    ("URLs con errores de canónica", "Canónica — Errores", "Indexación"),
    ("URLs con directiva Noindex", "Directivas — Noindex", "Indexación"),
    ("URLs con directiva Nofollow", "Directivas — Nofollow", "Indexación"),
    ("Metatítulo ausente", "Metatítulo — Falta", "On page"),
    ("Metatítulo duplicado", "Metatítulo — Duplicado", "On page"),
    ("Metatítulo múltiple", "Metatítulo — Múltiple", "On page"),
    ("Metatítulo muy largo (+60 car / +561 px)", "Metatítulo — Largo", "On page"),
    ("Metatítulo muy corto (−30 car / −200 px)", "Metatítulo — Corto", "On page"),
    ("Metadescripción ausente", "Metadescripción — Falta", "On page"),
    ("Metadescripción duplicada", "Metadescripción — Duplicada", "On page"),
    ("Metadescripción larga (+155 car / +985 px)", "Metadescripción — Larga", "On page"),
    ("Metadescripción corta (−70 car / −400 px)", "Metadescripción — Corta", "On page"),
    ("Metadescripción múltiple", "Metadescripción — Múltiple", "On page"),
    ("URLs sin H1", "H1 — Falta", "Contenido"),
    ("URLs con múltiples H1", "H1 — Múltiple", "Contenido"),
    ("URLs con H1 duplicado", "H1 — Duplicado", "Contenido"),
    ("URLs con H1 muy largo (+70 car)", "H1 — Largo (+70)", "Contenido"),
    ("URLs con poco contenido", "Poco contenido", "Contenido"),
    ("Imágenes +100kb", "Imágenes +100kb", "Contenido"),
    ("Imágenes sin ALT text", "Imágenes sin ALT text", "Contenido"),
    ("Imágenes sin atributo de tamaño", "Imágenes sin atributo tamaño", "Contenido"),
    ("Imágenes con dimensiones incorrectas", "Imágenes dim incorrectas", "Contenido"),
    ("PageSpeed − visualización de fuentes", "PS Visualización fuentes", "Indexación"),
    ("PageSpeed — LCP", "PS Solicitudes LCP", "Indexación"),
    ("PageSpeed — CLS", "PageSpeed — CLS", "Indexación"),
    ("URLs con metadatos propuestos", "Metadatos propuestos", "On page"),
]

SHEET_ORDER = [
    "Errores 404", "Detalle Errores 404",
    "Redirecciones 3xx", "Bucles y cadenas redirección",
    "URLs HTTP (no HTTPS)", "URLs no ASCII",
    "PS Minificar CSS", "PS Minificar JS",
    "PS Visualización fuentes", "PS Mejorar entrega imágenes",
    "PS Solicitudes LCP", "PageSpeed — CLS",
    "Canónica — Falta", "Canónica — Múltiple", "Canónica — Errores",
    "Directivas — Noindex", "Directivas — Nofollow",
    "Metatítulo — Falta", "Metatítulo — Duplicado", "Metatítulo — Múltiple",
    "Metatítulo — Largo", "Metatítulo — Corto", "Metatítulo = H1",
    "Metadescripción — Falta", "Metadescripción — Duplicada",
    "Metadescripción — Larga", "Metadescripción — Corta", "Metadescripción — Múltiple",
    "H1 — Falta", "H1 — Múltiple", "H1 — Duplicado", "H1 — Largo (+70)",
    "Poco contenido",
    "Imágenes +100kb", "Detalle Imágenes +100kb",
    "Imágenes sin ALT text", "Detalle Imágenes sin ALT",
    "Imágenes sin atributo tamaño", "Detalle Imágenes sin size attr",
    "Imágenes dim incorrectas", "Detalle Img dim incorrectas",
    "Metadatos propuestos",
]

PPT_HALLAZGOS = [
    (10, "Marcador de texto 2", "Errores 404"),
    (10, "Marcador de texto 4", "Redirecciones 3xx"),
    (13, "Marcador de texto 4", "Canónica — Errores"),
    (18, "Marcador de texto 4", "PS Minificar CSS"),
    (18, "Marcador de texto 6", "PS Minificar JS"),
    (19, "Marcador de texto 2", "Imágenes +100kb"),
    (22, "Marcador de texto 2", "Imágenes +100kb"),
    (22, "Marcador de texto 4", "Imágenes sin ALT text"),
    (23, "Marcador de texto 4", "Metadescripción — Duplicada"),
    (23, "Marcador de texto 6", "Poco contenido"),
    (26, "Marcador de texto 2", "H1 — Falta"),
    (26, "Marcador de texto 4", "H1 — Duplicado"),
    (26, "Marcador de texto 6", "H1 — Largo (+70)"),
    (27, "Marcador de texto 4", "Metatítulo — Corto"),
    (27, "Marcador de texto 6", "Metatítulo — Largo"),
    (28, "Marcador de texto 4", "Metadescripción — Corta"),
    (28, "Marcador de texto 6", "Metadescripción — Larga"),
]


def _reorder_sheets(wb):
    fixed = ["Matriz de entendimiento", "Resumen"]
    desired = fixed + [s for s in SHEET_ORDER if s in wb.sheetnames and s not in fixed]
    current = wb.sheetnames
    if current == desired:
        return
    for i, name in enumerate(desired):
        if i >= len(current):
            break
        if current[i] != name and name in current:
            idx = wb.sheetnames.index(name)
            wb.move_sheet(name, offset=i - idx)


# ---------- Main entry point ----------

def build_audit(data_dir: str, xlsx_path: str, pptx_path: str = None,
                client: str = "", domain: str = "",
                skip_ppt: bool = False, skip_metadatos: bool = True) -> dict:
    """
    Main function: reads NDJSONs from data_dir, builds Excel + PPT.
    Returns dict with sheet_name → count for all sheets.
    """
    data_dir = Path(data_dir)
    wb = _load_wb(xlsx_path)
    counts = {}

    # 1. Populate main sheets from SEO element NDJSONs
    for sheet_name, ndjson_file, headers, col_map in SHEETS_SPEC:
        ndjson_path = data_dir / ndjson_file
        rows_raw = _read_ndjson(ndjson_path)
        if sheet_name not in ("URLs HTTP (no HTTPS)",):
            rows_raw = _filter_by_domain(rows_raw, domain)
        rows_raw = _dedup_by_key(rows_raw, "Dirección")
        rows_out = []
        for r in rows_raw:
            row = []
            for key in col_map:
                if key is None:
                    row.append("")
                else:
                    val = r.get(key, "")
                    row.append(str(val) if val is not None else "")
            rows_out.append(row)
        _fill_sheet(wb, sheet_name, headers, rows_out)
        n = len(rows_out)
        counts[sheet_name] = n

    # 2. Detail sheets from bulk exports
    detail_specs = [
        ("Detalle Imágenes +100kb", "img_over_100kb_detalle.ndjson",
         ["URL imagen", "Tamaño (kB)", "URL donde aparece", "Título de la página"],
         ["Destino", "Tamaño (bytes)", "Fuente", None]),
        ("Detalle Imágenes sin ALT", "img_sin_alt_detalle.ndjson",
         ["URL imagen", "URL página", "Texto ancla"],
         ["Destino", "Fuente", "Ancla"]),
        ("Detalle Imágenes sin size attr", "img_sin_size_detalle.ndjson",
         ["URL imagen", "URL página"],
         ["Destino", "Fuente"]),
    ]
    for sheet_name, ndjson_file, headers, col_map in detail_specs:
        ndjson_path = data_dir / ndjson_file
        rows_raw = _read_ndjson(ndjson_path)
        rows_raw = _filter_by_domain(rows_raw, domain, key="Fuente")
        rows_out = [[str(r.get(k, "") or "") for k in col_map] for r in rows_raw]
        _fill_sheet(wb, sheet_name, headers, rows_out)
        counts[sheet_name] = len(rows_out)

    # 3. Detalle Errores 404 (special key "URL 404")
    d404_ndjson = data_dir / "detalle_errores_404.ndjson"
    if d404_ndjson.exists():
        rows_raw = _read_ndjson(d404_ndjson)
        rows_out = [[r.get("URL 404", ""), r.get("URL Origen", ""), r.get("Texto ancla", "") or ""] for r in rows_raw]
        _fill_sheet(wb, "Detalle Errores 404", ["URL 404", "URL Origen", "Texto ancla"], rows_out)
        counts["Detalle Errores 404"] = len(rows_out)
    else:
        _fill_sheet(wb, "Detalle Errores 404", ["URL 404", "URL Origen", "Texto ancla"], [])
        counts["Detalle Errores 404"] = 0

    # 4. Canónica sheets (bulk export with Dirección key)
    for cname, cheaders, cfile in [
        ("Canónica — Múltiple", ["URL", "Canónicas"], "canónica_múltiple.ndjson"),
        ("Canónica — Errores", ["URL", "Tipo problema", "Canónica", "Estado"], "canónica_errores.ndjson"),
    ]:
        ndjson_path = data_dir / cfile
        if ndjson_path.exists():
            rows_raw = _read_ndjson(ndjson_path)
            rows_raw = _dedup_by_key(rows_raw, "Dirección")
            rows_out = [[r.get("Dirección", "")] + [""] * (len(cheaders) - 1) for r in rows_raw]
            _fill_sheet(wb, cname, cheaders, rows_out)
            counts[cname] = len(rows_out)
        else:
            _fill_sheet(wb, cname, cheaders, [])
            counts[cname] = 0

    # 5. PageSpeed sheets from sf_generate_report (file-level detail)
    ps_report_sheets = [
        ("PS Minificar JS", "ps_report_minificar_js.ndjson",
         ["URL Página", "Archivo JS a minificar", "Tamaño (bytes)", "Ahorro estimado (bytes)"],
         ["Página fuente", "URL", "Tamaño (bytes)", "Posible ahorro (bytes)"]),
        ("PS Minificar CSS", "ps_report_minificar_css.ndjson",
         ["URL Página", "Archivo CSS a minificar", "Tamaño (bytes)", "Ahorro estimado (bytes)"],
         ["Página fuente", "URL", "Tamaño (bytes)", "Posible ahorro (bytes)"]),
        ("PS Visualización fuentes", "ps_report_fuentes.ndjson",
         ["URL Página", "Archivo de fuente", "Ahorro (ms)"],
         ["Página fuente", "URL", "Posible ahorro (ms)"]),
        ("PS Mejorar entrega imágenes", "ps_report_mejorar_img.ndjson",
         ["URL Página", "Imagen a optimizar", "Tamaño (bytes)", "Ahorro (bytes)", "Motivo"],
         ["Página fuente", "URL", "Tamaño (bytes)", "Posible ahorro (bytes)", "Motivo"]),
        ("PS Solicitudes LCP", "ps_report_lcp.ndjson",
         ["URL Página", "Selector LCP", "Carga diferida", "fetchpriority=high", "Detectable inicial"],
         ["Página fuente", "Selector", "Carga diferida no aplicada", 'Se ha de aplicar "fetchpriority=high"', "La solicitud es detectable en el documento inicial"]),
        ("PageSpeed — CLS", "ps_report_cls.ndjson",
         ["URL Página", "Elemento causante", "Contribución CLS", "Fragmento HTML"],
         ["Página fuente", "Etiqueta", "Contribución de CLS", "Recortes"]),
    ]
    for sheet_name, ndjson_file, headers, col_map in ps_report_sheets:
        ndjson_path = data_dir / ndjson_file
        rows_raw = _read_ndjson(ndjson_path)
        if not rows_raw:
            _fill_sheet(wb, sheet_name, headers, [])
            counts[sheet_name] = 0
        else:
            rows_out = [[str(r.get(k, "") or "") for k in col_map] for r in rows_raw]
            _fill_sheet(wb, sheet_name, headers, rows_out)
            counts[sheet_name] = len(rows_out)

    # 6. Remaining empty sheets
    empty_sheets = [
        ("Bucles y cadenas redirección", ["URL", "Tipo (Bucle/Cadena)", "Código", "URL destino"]),
        ("Metatítulo = H1", ["URL", "Título", "H1"]),
        ("Detalle Img dim incorrectas", ["URL imagen", "URL página", "Dimensión declarada"]),
    ]
    for sheet_name, headers in empty_sheets:
        _fill_sheet(wb, sheet_name, headers, [])
        counts[sheet_name] = 0

    # 7. Directivas (from bulk export NDJSONs)
    for sheet_name, ndjson_file, headers in [
        ("Directivas — Noindex", "noindex.ndjson", ["URL", "Indexabilidad", "Estado de indexabilidad"]),
        ("Directivas — Nofollow", "nofollow.ndjson", ["URL", "Directivas"]),
    ]:
        ndjson_path = data_dir / ndjson_file
        rows_raw = _read_ndjson(ndjson_path)
        seen = set()
        rows_out = []
        for r in rows_raw:
            url = r.get("Fuente", "")
            if url in seen:
                continue
            seen.add(url)
            rows_out.append([url, "", ""])
        _fill_sheet(wb, sheet_name, headers, rows_out)
        counts[sheet_name] = len(rows_out)

    # 8. Metadatos propuestos (from rows_metadatos.json in data_dir)
    if not skip_metadatos:
        metadatos_json = data_dir / "rows_metadatos.json"
        if metadatos_json.exists():
            with open(metadatos_json, encoding="utf-8") as f:
                met_rows = json.load(f)
            _fill_metadatos_propuestos(wb, met_rows)
            counts["Metadatos propuestos"] = len(met_rows)
        else:
            _fill_sheet(wb, "Metadatos propuestos",
                        ["URL", "Factor a corregir", "Texto actual", "Propuesta"], [])
            counts["Metadatos propuestos"] = 0
    else:
        _fill_sheet(wb, "Metadatos propuestos",
                    ["URL", "Factor a corregir", "Texto actual", "Propuesta"], [])
        counts["Metadatos propuestos"] = 0

    # 9. Rebuild Resumen
    ws_resumen = wb["Resumen"] if "Resumen" in wb.sheetnames else wb.create_sheet("Resumen")
    ws_resumen.delete_rows(1, ws_resumen.max_row)
    ws_resumen.append(["RICOPA", "FACTORES", "DESCRIPCIÓN DE FACTORES", "CANTIDAD", "VER"])
    ws_resumen.column_dimensions["A"].width = 10
    ws_resumen.column_dimensions["B"].width = 40
    ws_resumen.column_dimensions["C"].width = 50
    ws_resumen.column_dimensions["D"].width = 12
    ws_resumen.column_dimensions["E"].width = 8

    current_cat = None
    for factor_text, sheet_name, cat in RESUMEN_SPEC:
        count = counts.get(sheet_name, 0)
        row_idx = ws_resumen.max_row + 1
        ws_resumen.cell(row=row_idx, column=1, value=cat if cat != current_cat else "")
        ws_resumen.cell(row=row_idx, column=2, value=factor_text)
        ws_resumen.cell(row=row_idx, column=3, value=factor_text)
        ws_resumen.cell(row=row_idx, column=4, value=count)
        ver_cell = ws_resumen.cell(row=row_idx, column=5, value="VER")
        ver_cell.hyperlink = f"#'{sheet_name}'!A1"
        ver_cell.font = Font(color="0563C1", underline="single")
        current_cat = cat

    # 10. Reorder & save
    _reorder_sheets(wb)
    _save_wb(wb, xlsx_path)

    # 11. PowerPoint
    if not skip_ppt and pptx_path:
        excel_filename = os.path.basename(xlsx_path)
        prs = PptxPresentation(pptx_path)
        today_str = date.today().isoformat()

        for shape in prs.slides[0].shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if "Tivit" in (run.text or ""):
                            run.text = run.text.replace("Tivit", client)

        for shape in prs.slides[1].shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if re.search(r"\d{4}-\d{2}-\d{2}", run.text or ""):
                            run.text = re.sub(r"\d{4}-\d{2}-\d{2}", today_str, run.text or "")

        for slide_idx, shape_substr, sheet_name in PPT_HALLAZGOS:
            count = counts.get(sheet_name, 0)
            slide = prs.slides[slide_idx - 1]
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                if shape_substr.lower() not in shape.name.lower():
                    continue
                tf = shape.text_frame
                edited_here = False
                for para in tf.paragraphs:
                    for run in list(para.runs):
                        t = run.text or ""
                        if re.search(r"Hay \d+", t):
                            run.text = re.sub(r"Hay \d+", f"Hay {count}", t)
                            run.text = re.sub(r"\s*VER\s*$", "", run.text)
                            edited_here = True
                if not edited_here:
                    tf.clear()
                    p = tf.paragraphs[0]
                    p.text = f"Hay {count} hallazgos."
                last_para = tf.paragraphs[-1]
                ver_run = last_para.add_run()
                ver_run.text = "  VER"
                ver_run.font.name = BRAND_FONT
                ver_run.font.size = Pt(14)
                try:
                    ver_run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
                except Exception:
                    pass
                try:
                    ver_run.hyperlink.address = f"{excel_filename}#'{sheet_name}'!A1"
                except Exception:
                    pass

        prs.save(pptx_path)

    return counts
