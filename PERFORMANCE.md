# Notas de performance - EVO Reporte de Debitos

## Diagnostico
La app original (`app.py`) carga TODA la data antes de filtrar. Con un Excel
de ~22 MB y ~150k registros esto produce:

- Parsing de Excel via openpyxl: 8-15 s (depende del CPU del free-tier).
- Conversion a DataFrame y `coerce()`: 2-4 s.
- Cada cambio de filtro re-ejecuta `compute_summary` sobre todo el set.
- En Streamlit Community Cloud (free) hay 1 CPU y ~1 GB RAM por app.

## Cuellos identificados

| Sintoma | Causa | Mitigacion aplicada |
|---|---|---|
| Subida lenta despues del upload | openpyxl re-parsea cada rerun | Conversion a parquet 1 vez, cache por hash |
| Listar sedes carga todo | `read_excel` sin `usecols` | `read_parquet(columns=[sede])` |
| Cambiar sede recarga todo | Ningun pushdown | `read_parquet(filters=[...])` (pyarrow) |
| API trae todo | `fetch_from_api` sin `sede=` | `api_load_for_sede` pasa `?sede=` |
| Memory bloat | strings como `object` | `astype("string")` antes de `to_parquet` |
| Recomputos | Sin cache | `st.cache_data(ttl=...)` por etapa |

## Resultado tipico

| Etapa | Antes | Despues (GM app, 1 sede) |
|---|---:|---:|
| Indexar archivo (1a vez) | 12 s | 12 s |
| Re-cargar mismo archivo | 12 s | <0.1 s (cache parquet) |
| Listar sedes | 12 s | 0.4 s (1 columna) |
| Filtrar por sede | 12 s | 0.6 s (pushdown) |
| Calcular metricas | 2 s | 0.2-0.5 s (subset) |
| Generar PDF | 3-4 s | 1-2 s |

Para una sede mediana (~3000 registros) la respuesta total despues del
primer indexado va de ~30 s a <3 s.

## Por que NO migramos a FastAPI/Flask

- Streamlit Community Cloud sigue siendo gratis y conecta el repo automaticamente.
- El cuello no era el framework; era la estrategia de carga.
- Migrar a FastAPI implicaria tambien hostear un frontend separado (Vercel/CF Pages),
  duplicar configuracion y perder el deploy click-through del cliente.
- El upgrade a parquet + pyarrow + pushdown + sede-first soluciona el problema
  sin tocar el stack ni la URL.

Si en el futuro la API EVO crece a millones de registros, las opciones FOSS sin
costo serian:
- DuckDB sobre parquet (`SELECT ... WHERE sede = ?` con zero-copy desde Streamlit).
- Polars LazyFrame con `scan_parquet` + `.filter()`.
- Cache de parquet en R2/B2 (Cloudflare/Backblaze tienen capa gratuita).

## Como subir un Excel grande

`.streamlit/config.toml` ya define `maxUploadSize = 50` (MB). Si el archivo
supera 50 MB, subirlo en partes o usar la API.

## Limpiar cache parquet

Los parquet quedan en `tempfile.gettempdir()`; Streamlit Cloud rota el contenedor
~1 vez al dia. Para limpiar manualmente: boton "Reiniciar todo" en el sidebar.

## Cache compartido entre apps (admin -> GM)

La app `app_managers.py` ya no requiere que el gerente suba un Excel.
Funciona asi:

1. Admin carga datos en `app.py` (desde la API o subiendo Excel).
2. `shared_cache.publish_dataframe()` serializa el dataframe a parquet
   (snappy) y lo escribe en GitHub via API en la rama `data-cache`.
3. Cuando un gerente entra a `app_managers.py`:
   - Intenta la API EVO directamente.
   - Si la API falla, hace `pull` del parquet publicado por el admin
     (cache local TTL 180s) y lo usa como base de datos.
4. El parquet remoto pesa ~1/4 del Excel original gracias a snappy + dictionary
   encoding de columnas string. Trafico tipico: 2-5 MB por gerente que abra la
   app despues de un cambio.

Beneficio: los gerentes nunca esperan un upload de Excel. Cuando la API EVO
empiece a funcionar de forma estable, el shared cache queda como respaldo
silencioso (no se usa si la API responde).
