# Auditoría SEO RICOPA — Web App

Genera entregables Excel + PowerPoint para auditorías SEO siguiendo la metodología RICOPA.

## ¿Qué hace?

1. Recibe un ZIP con archivos NDJSON exportados de Screaming Frog
2. Genera el Excel con 44 hojas (hallazgos + resumen + metadatos propuestos)
3. Genera el PowerPoint con los hallazgos, conteos y links VER
4. Devuelve ambos archivos para descargar

## Uso local

```bash
# 1. Crear entorno virtual
python3 -m venv .venv
source .venv/bin/activate

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Iniciar servidor
uvicorn main:app --reload --port 8000

# 4. Abrir http://localhost:8000
```

## ¿Qué meter en el ZIP?

Zipea todos los archivos `.ndjson` del crawl. Ejemplo para un cliente `Avafin`:

```bash
cd /Users/wilman/seo_spider_mcp_server
zip -r avafin.zip audits/avafin/*.ndjson
```

Incluye también `rows_metadatos.json` si tienes metadatos propuestos generados.

Archivos esperados (todos `.ndjson`):
- `errors_4xx.ndjson`, `redirects_3xx.ndjson`
- `title_*.ndjson`, `meta_*.ndjson`, `h1_*.ndjson`
- `img_over_100kb.ndjson`, `img_sin_alt.ndjson`, `img_sin_size.ndjson`
- `can_falta.ndjson`, `canónica_múltiple.ndjson`, `canónica_errores.ndjson`
- `noindex.ndjson`, `nofollow.ndjson`
- `ps_report_minificar_js.ndjson`, `ps_report_fuentes.ndjson`, etc.
- `detalle_errores_404.ndjson`
- `img_over_100kb_detalle.ndjson`, `img_sin_alt_detalle.ndjson`, `img_sin_size_detalle.ndjson`

## Deploy en Render

1. Crear repositorio en GitHub con este proyecto
2. Ir a [render.com](https://render.com) → New Web Service
3. Conectar el repo de GitHub
4. Configurar:
   - **Environment**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
5. Deploy → obtienes URL tipo `https://auditoria-seo.onrender.com`

Render auto-deploya en cada push a main.

## Estructura del proyecto

```
auditoria-web/
├── main.py              # FastAPI app (upload, process, download)
├── build_audit.py       # Core logic (copied from skill)
├── requirements.txt
├── _plantillas/         # Excel + PPT templates (committed to repo)
├── templates/           # HTML templates (upload.html, result.html)
├── static/              # Static files (CSS, JS)
├── outputs/             # Generated files (gitignored)
└── .gitignore
```
