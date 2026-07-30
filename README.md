# Auditoría SEO RICOPA — Web App

Genera entregables Excel + PowerPoint para auditorías SEO siguiendo la metodología RICOPA.

## ¿Qué hace?

1. Recibe un ZIP con archivos CSV/NDJSON exportados de Screaming Frog
2. **Detecta automáticamente** a qué hoja pertenece cada archivo por sus columnas (sin importar el nombre)
3. Deriva hojas desde export "Todo" — un solo CSV de Títulos produce Metatítulo Falta/Largo/Corto/Duplicado
4. Genera el Excel con ~44 hojas (hallazgos + resumen + metadatos propuestos)
5. Genera el PowerPoint con conteos actualizados y links `VER`
6. Devuelve ambos archivos + `mapping.json` para reutilizar en futuros crawls

## V4 — Detección inteligente (columnas + data sampling)

El sistema identifica archivos por su **firma de columnas**:

```
CSV con Dirección, Título 1, Longitud... → "es un archivo de títulos"
CSV con Página fuente, URL, Posible ahorro... → "es un reporte de PageSpeed"
CSV con Dirección, Indexabilidad → "es un archivo 'Falta', ver nombre"
```

**Compatibilidad total SF Desktop → MCP**: los CSV exportados desde el desktop de SF usan nombres de columna distintos (`Desde`, `Hasta`, `Código de respuesta`). El sistema los normaliza automáticamente para que coincidan con las firmas.

Si no puede determinarlo, **no bloquea** — muestra una pantalla con dropdowns agrupados por RICOPA (Rastreo/Indexación/On Page/Contenido) para que selecciones manualmente. El `mapping.json` generado al finalizar guarda tus elecciones para el siguiente crawl.

### Data sampling

Lee las primeras 10 filas del CSV y decide por los valores reales:
- Todas las filas con `Longitud > 60` → Metatítulo — Largo
- Ningún `Título 1` presente → Metatítulo — Falta
- Columnas `Desde`/`Hasta`/`Código de respuesta` → Detalle Errores 404 (normalizado a Fuente/Destino/Código de estado)

### Derivación desde Todo

Si subís un CSV "Todo" de un elemento, el dropdown ofrece opciones especiales:
- `📋 Títulos → derivar hojas` (genera Falta, Largo, Corto, Duplicado)
- `📋 Meta descripción → derivar hojas`
- `📋 H1 → derivar hojas`
- `📋 Imágenes → derivar hojas` (+100kb, sin ALT, sin size)

## Uso local

```bash
# 1. entorno + dependencias
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 2. iniciar
uvicorn main:app --reload --port 8000

# 3. abrir http://localhost:8000
```

## ¿Qué exportar de Screaming Frog?

**Mínimo (por archivos individuales):**
- Códigos de respuesta: Error de cliente (4xx), Redirección (3xx)
- Títulos de página: Falta, Duplicado, Múltiple, Más de 60 car, Menos de 30 car, Más de 561 px, Menos de 200 px
- Meta description: Falta, Duplicado, Múltiple, Más de 155 car, Menos de 70 car
- H1: Falta, Múltiple, Duplicado, Más de 70 car
- Canonicals: Falta, Múltiple, errores varios
- Imágenes: Por encima de 100 kB, Falta texto ALT, Atributos de tamaño ausentes
- Directivas: Noindex, Nofollow (bulk export)
- PageSpeed: Informes > Minimizar JavaScript/CSS, Visualización fuentes, Mejorar entrega, LCP, CLS

**Recomendado (más eficiente):**
- **1 export "Todo"** por elemento: Títulos → Todo, Meta → Todo, H1 → Todo
- El script deriva automáticamente las sub-hojas por reglas de longitud/caracteres
- Más los bulk exports e informes PageSpeed

> **Nombrá los archivos como quieras** — el sistema solo mira las columnas adentro.

## Deploy en Render

1. Crear repo GitHub con este proyecto
2. [render.com](https://render.com) → New Web Service → conectar repo
3. Config: Python 3, `pip install -r requirements.txt`, `uvicorn main:app --host 0.0.0.0 --port $PORT`
4. Deploy — obtienes URL gratis tipo `https://miauditoria.onrender.com`

## Estructura

```
auditoria-web/
├── main.py              # FastAPI (upload, process, download)
├── build_audit.py       # Core logic: detection + derivation + Excel + PPT
├── requirements.txt     # fastapi, openpyxl, python-pptx, uvicorn
├── _plantillas/         # Templates Excel + PPT
├── templates/           # HTML (upload.html, result.html)
└── outputs/             # Generated files (gitignored)
```

## mapping.json

Al finalizar una auditoría, la app genera un `mapping.json` que asigna cada archivo → hoja. 
Inclúyelo en futuros ZIPs del mismo cliente para que la detección sea 100% automática sin intervención.

Formato:
```json
{
  "title_falta.csv": "Metatítulo — Falta",
  "h1_duplicado.csv": "H1 — Duplicado",
  "noindex_bulk.csv": "Directivas — Noindex"
}
```
