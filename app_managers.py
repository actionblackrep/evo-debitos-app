# -*- coding: utf-8 -*-
"""
EVO - Reporte de Debitos - APP DE GERENTES GENERALES

Auto-fuente:
  1. Intenta la API EVO (filtra en servidor).
  2. Si la API falla, usa los datos publicados por el admin en el shared cache
     (parquet en una rama del repo de GitHub).
  3. Si tampoco hay shared cache, ofrece subida manual de Excel como ultimo recurso.

Despues de tener fuente, pide la sede (obligatorio) antes de cargar/procesar.
Genera el mismo PDF que app.py.
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import urllib.error
import urllib.parse
import urllib.request
import tempfile
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

import generate_report as gr
import shared_cache

st.set_page_config(
    page_title="EVO - Reporte de Debitos | Gerentes Generales",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.main .block-container {padding-top: 1.6rem; padding-bottom: 2rem;}
h1, h2, h3 {color: #0F2A4A;}
.metric-card {background: #F4F6FA; border-left: 4px solid #0F2A4A;
              border-radius: 6px; padding: 14px 16px;}
.metric-card.good {border-left-color: #1F8A4C;}
.metric-card.bad {border-left-color: #C0392B;}
.metric-card.gold {border-left-color: #D4AF37;}
.metric-value {font-size: 28px; font-weight: 700; color: #0F2A4A; line-height: 1;}
.metric-label {font-size: 11px; color: #7A8597; text-transform: uppercase;
               letter-spacing: 0.5px; margin-top: 4px;}
.metric-sub {font-size: 12px; color: #7A8597; margin-top: 2px;}
hr {margin: 0.8rem 0;}
.stDownloadButton button {background: #0F2A4A; color: white; border-radius: 6px;
                          padding: 0.6rem 1.4rem; font-weight: 600;}
.sede-banner {background: linear-gradient(90deg, #0F2A4A 0%, #1B3A66 100%);
              color: white; padding: 14px 18px; border-radius: 8px;
              margin-bottom: 14px;}
.sede-banner b {color: #D4AF37;}
.src-pill {display: inline-block; padding: 2px 10px; border-radius: 999px;
           font-size: 11px; font-weight: 600; letter-spacing: 0.4px;
           text-transform: uppercase; margin-right: 6px;}
.src-pill.api {background: #1F8A4C; color: white;}
.src-pill.shared {background: #D4AF37; color: #0F2A4A;}
.src-pill.file {background: #7A8597; color: white;}
</style>
""", unsafe_allow_html=True)


# =========================================================
# Bogota timezone helpers (UTC-5, sin DST)
# =========================================================
from datetime import timezone as _tz

BOGOTA_TZ = _tz(timedelta(hours=-5))
ES_MONTHS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul",
             "Ago", "Sep", "Oct", "Nov", "Dic"]


def format_publish_date_bogota(iso_str: str | None) -> str:
    """Convert UTC ISO8601 to '09 May 2026 - 11:47 PM Bogota'."""
    if not iso_str:
        return ""
    try:
        s = iso_str.strip()
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=_tz.utc)
        bg = dt.astimezone(BOGOTA_TZ)
        h12 = bg.hour % 12 or 12
        ampm = "PM" if bg.hour >= 12 else "AM"
        return f"{bg.day:02d} {ES_MONTHS[bg.month-1]} {bg.year} - {h12}:{bg.minute:02d} {ampm} Bogota"
    except Exception:
        return iso_str


SEDE_COL_CANDIDATES = ["Sede/club", "Sede", "Club", "sede", "club", "sede_club", "branch"]


# =========================================================
# CACHE LAYER
# =========================================================
@st.cache_data(ttl=1800, show_spinner=False)
def excel_bytes_to_parquet(file_bytes: bytes, name: str) -> str:
    suffix = os.path.splitext(name)[1].lower() or ".xlsx"
    digest = hashlib.md5(file_bytes).hexdigest()[:16]
    pq_path = os.path.join(tempfile.gettempdir(), f"evo_debits_{digest}.parquet")
    if os.path.exists(pq_path) and os.path.getsize(pq_path) > 0:
        return pq_path
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as f:
        f.write(file_bytes)
        xpath = f.name
    try:
        df = gr.load_data([xpath])
        _vectorize_canonicalize(df)
        df.to_parquet(pq_path, index=False, compression="snappy")
    finally:
        try:
            os.unlink(xpath)
        except Exception:
            pass
    return pq_path


@st.cache_data(ttl=1800, show_spinner=False)
def parquet_sede_column(pq_path: str):
    import pyarrow.parquet as pq
    schema_names = pq.ParquetFile(pq_path).schema_arrow.names
    for c in SEDE_COL_CANDIDATES:
        if c in schema_names:
            return c
    return None


@st.cache_data(ttl=1800, show_spinner=False)
def parquet_list_sedes(pq_path: str):
    sede_col = parquet_sede_column(pq_path)
    if not sede_col:
        return None, []
    df = pd.read_parquet(pq_path, columns=[sede_col])
    sedes = (
        df[sede_col]
        .dropna()
        .astype(str)
        .map(str.strip)
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    return sede_col, sorted(sedes)


@st.cache_data(ttl=900, show_spinner=False)
def parquet_filter_by_sede(pq_path: str, sede_col: str, sede: str) -> pd.DataFrame:
    """Filter parquet by sede.

    1) Try exact-match via pyarrow predicate pushdown (fast).
    2) If 0 rows, fallback to full scan with whitespace-stripped comparison
       to handle parquets that were not canonicalized at write time.
    """
    df = pd.read_parquet(pq_path, filters=[(sede_col, "=", sede)])
    if not df.empty:
        return df
    # Fallback path: read full and compare stripped strings
    full = pd.read_parquet(pq_path)
    target = (sede or "").strip()
    mask = full[sede_col].astype(str).str.strip() == target
    return full.loc[mask].copy()


@st.cache_data(ttl=300, show_spinner=False)
def api_list_sedes(api_url: str, api_key: str):
    df = gr.fetch_from_api(api_url, api_key)
    cols = gr.resolve_columns(df)
    if "sede" not in cols:
        return []
    return sorted(df[cols["sede"]].dropna().astype(str).unique().tolist())


@st.cache_data(ttl=300, show_spinner=False)
def api_load_for_sede(api_url: str, api_key: str, sede: str) -> pd.DataFrame:
    try:
        df = gr.fetch_from_api(api_url, api_key, sede=sede)
    except Exception:
        df = gr.fetch_from_api(api_url, api_key)
    cols = gr.resolve_columns(df)
    if "sede" in cols and len(df):
        target = gr.normalize_match(sede)
        mask = df[cols["sede"]].astype(str).map(gr.normalize_match) == target
        if mask.sum() > 0:
            df = df.loc[mask].copy()
    return df


@st.cache_data(ttl=180, show_spinner=False)
def shared_pull_path() -> str | None:
    """Cached pull from GitHub shared cache. Returns local parquet path or None."""
    return shared_cache.pull_parquet()



# =========================================================
# API v2 (manual connect, multi-header, month param)
# =========================================================
# Hard cap to protect Streamlit free-tier (1GB RAM). Adjust if endpoint grows.
MAX_RECORDS = 300_000


def _vectorize_canonicalize(df: pd.DataFrame) -> None:
    """Strip whitespace on all string-like columns. In-place. Vectorized.

    Replaces the previous .map(lambda) approach which was O(rows*cols) Python
    overhead and OOM-prone on 100k+ rows.
    """
    for c in df.columns:
        s = df[c]
        try:
            if s.dtype == object or str(s.dtype) == "string":
                df[c] = s.astype("string").str.strip()
        except Exception:
            continue


def _fetch_debitos_v2(url: str, token: str, month: str | None,
                      page_size: int = 10000, timeout: int = 60) -> pd.DataFrame:
    """Fetch debitos with `month=YYYY-MM`, trying multiple auth header variants.

    Token never reaches the UI. Records cap at MAX_RECORDS to prevent OOM.
    Uses pd.DataFrame() (much faster than json_normalize on flat records).
    """
    if not month:
        month = datetime.now().strftime("%Y-%m")

    base_qs = {"month": month}
    header_variants = [
        {"Authorization": f"Bearer {token}"},
        {"Authorization": token},
        {"x-api-key": token},
        {"X-API-Token": token},
        {"api-token": token},
    ]
    last_err = "no se intento ninguna llamada"
    for hv in header_variants:
        headers = dict(hv)
        headers["Accept"] = "application/json"
        headers["User-Agent"] = "evo-debitos-gm/2.1"
        rows: list = []
        ok = False
        page = 1
        max_pages = (MAX_RECORDS // page_size) + 1
        while True:
            qs = dict(base_qs)
            qs["page"] = page
            qs["size"] = page_size
            full = url + "?" + urllib.parse.urlencode(qs)
            req = urllib.request.Request(full, headers=headers, method="GET")
            try:
                with urllib.request.urlopen(req, timeout=timeout) as r:
                    raw = r.read()
                    payload = json.loads(raw)
                    ok = True
            except urllib.error.HTTPError as e:
                last_err = f"HTTP {e.code} con header {list(hv)[0]}"
                ok = False
                break
            except Exception as e:
                last_err = f"{type(e).__name__}: {e}"
                ok = False
                break

            if isinstance(payload, list):
                recs = payload
            elif isinstance(payload, dict):
                recs = next(
                    (payload[k] for k in ("data", "records", "items", "results", "rows")
                     if isinstance(payload.get(k), list)),
                    [],
                )
            else:
                recs = []

            if not recs:
                break
            rows.extend(recs)
            if len(rows) >= MAX_RECORDS:
                # Hard stop. Better to return partial than crash.
                break
            if len(recs) < page_size:
                break
            page += 1
            if page > max_pages:
                break

        if ok and rows:
            # Use pd.DataFrame on flat records (much faster than json_normalize).
            try:
                df = pd.DataFrame(rows)
            except Exception:
                df = pd.json_normalize(rows)
            df["__source_file__"] = f"api:{url}?month={month}"
            return df

    raise RuntimeError(f"No se pudo conectar a la API. Detalle: {last_err}")


@st.cache_data(ttl=600, show_spinner=False)
def cached_fetch_api_range(url: str, from_str: str, to_str: str):
    """Fetch via /api/debitos?from=&to=. Auto-split por mes.

    Returns (parquet_path, chunks_log).
    Token leido internamente desde gr.DEFAULT_API_KEY; nunca en los args.
    """
    from datetime import date as _dt
    f = _dt.fromisoformat(from_str)
    t = _dt.fromisoformat(to_str)
    chunks_log = []
    df = gr.fetch_debitos_range(
        url, gr.DEFAULT_API_KEY, f, t,
        logger=lambda m: chunks_log.append(m),
    )
    _vectorize_canonicalize(df)
    h = hashlib.md5(f"{url}::{from_str}::{to_str}".encode()).hexdigest()[:16]
    pq_path = os.path.join(tempfile.gettempdir(), f"evo_debits_api_{h}.parquet")
    df.to_parquet(pq_path, index=False, compression="snappy")
    return pq_path, chunks_log


def available_months(n: int = 12) -> list[str]:
    out = []
    today = date.today()
    y, m = today.year, today.month
    for _ in range(n):
        out.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return out


# =========================================================
# Pipeline
# =========================================================
ID_CLIENT_CANDIDATES = [
    "id_client", "id_cliente", "client_id", "Cliente_id",
    "id", "ID", "Id", "user_id", "userId", "userid",
    "id_usuario", "ID_usuario", "idUsuario",
]


def detect_id_client_column(df) -> str | None:
    for c in ID_CLIENT_CANDIDATES:
        if c in df.columns:
            return c
    return None


def run_pipeline_for_sede(df, sede_label, since, until):
    cols = gr.resolve_columns(df)
    df = gr.coerce(df, cols)
    def _to_naive(s):
        # Handle tz-aware ISO8601 strings from the new API
        try:
            x = pd.to_datetime(s, errors="coerce", utc=True)
            if getattr(x.dt, "tz", None) is not None:
                x = x.dt.tz_convert("UTC").dt.tz_localize(None)
            return x
        except Exception:
            return pd.to_datetime(s, errors="coerce")

    if "intento" in cols and (since or until):
        d = _to_naive(df[cols["intento"]])
        if since:
            df = df[d >= pd.to_datetime(since)]
            d = _to_naive(df[cols["intento"]])
        if until:
            df = df[d <= pd.to_datetime(until) + pd.Timedelta(days=1)]
    summary, ap, dn = gr.compute_summary(df, cols)
    # Nota: compute_summary ya computa clients_total/approved/denied de forma
    # mutuamente exclusiva (clients_approved = sin fallos; clients_denied = con
    # >= 1 fallo). NO redefinir aqui.

    sedes = gr.by_sede(df, cols, ap, dn)
    motivos = gr.by_motivo(df, cols, dn)
    tipos = gr.by_tipo(df, cols, dn)
    marcas = gr.by_marca(df, cols, ap, dn)
    daily = gr.by_day(df, cols, ap, dn)
    return df, cols, summary, sedes, motivos, tipos, marcas, daily


def build_pdf_bytes(df, cols, summary, sedes, motivos, tipos, marcas, daily, sede_label,
                    since=None, until=None):
    """Genera el PDF.

    El header del PDF muestra EXACTAMENTE el rango since/until elegido por el
    usuario (no se infiere del dataset). Esto garantiza que el PDF, las cards,
    la tabla y el Excel siempre reporten el mismo periodo.
    """
    if since is not None and until is not None:
        summary = dict(summary)
        summary["date_min"] = pd.Timestamp(since)
        summary["date_max"] = pd.Timestamp(until)
    with tempfile.TemporaryDirectory() as tmp:
        chart_dir = os.path.join(tmp, "charts")
        os.makedirs(chart_dir, exist_ok=True)
        gr.setup_mpl()
        chart_paths = {
            "donut": os.path.join(chart_dir, "donut.png"),
            "trend": os.path.join(chart_dir, "trend.png"),
            "motivos": None,
            "sedes_fail": None,
            "sedes_rate": None,
        }
        gr.chart_donut_status(summary, chart_paths["donut"])
        if not daily.empty and len(daily) > 1:
            gr.chart_trend(daily, chart_paths["trend"])
        else:
            chart_paths["trend"] = None
        pdf_path = os.path.join(tmp, "reporte.pdf")
        gr.build_pdf(pdf_path, df, cols, summary, sedes, motivos, tipos, marcas, daily,
                     chart_paths, sede_label=sede_label)
        with open(pdf_path, "rb") as f:
            return f.read()


def metric_card(label, value, sub=None, kind="default"):
    klass = {"good": "metric-card good", "bad": "metric-card bad",
             "gold": "metric-card gold"}.get(kind, "metric-card")
    sub_html = f"<div class='metric-sub'>{sub}</div>" if sub else ""
    st.markdown(
        f"""<div class="{klass}">
              <div class="metric-value">{value}</div>
              <div class="metric-label">{label}</div>
              {sub_html}
            </div>""",
        unsafe_allow_html=True,
    )


def reset_state():
    keys = ["source", "sedes_api", "sedes_shared", "shared_pq_path", "shared_sede_col",
            "api_error", "shared_error", "used_fallback", "sede_pick", "upl",
            "api_from", "api_to", "api_chunks_log", "api_month"]
    for k in keys:
        st.session_state.pop(k, None)


# =========================================================
# Sidebar
# =========================================================
# Endpoint config (token NO se expone en la UI; viene de st.secrets / env)
api_url = gr.DEFAULT_API_URL

with st.sidebar:
    st.title("Configuracion")
    st.caption("Token y URL gestionados por el backend.")
    st.code(api_url, language=None)
    if st.button("Recargar datos", type="primary", use_container_width=True,
                 help="Limpia todos los caches locales y reinicia la fuente."):
        reset_state()
        st.cache_data.clear()
        st.rerun()
    if st.button("Reiniciar fuente", use_container_width=True):
        reset_state()
        st.rerun()
    st.markdown("---")
    st.caption("Fuente actual:")
    src = st.session_state.get("source")
    if src == "api":
        st.markdown("<span class='src-pill api'>API EVO</span>", unsafe_allow_html=True)
    elif src == "shared":
        st.markdown("<span class='src-pill shared'>Datos del admin</span>", unsafe_allow_html=True)
    elif src == "file":
        st.markdown("<span class='src-pill file'>Excel manual</span>", unsafe_allow_html=True)
    else:
        st.caption("(sin fuente todavia)")
    st.markdown("---")
    st.markdown("**Cache compartido (admin -> GM)**")
    if shared_cache.is_configured():
        cfg = shared_cache.cache_config()
        st.caption(f"Repo: `{cfg['repo']}` · Rama: `{cfg['branch']}`")
        meta = shared_cache.remote_meta()
        if meta:
            st.success(f"Ultima publicacion:\n{format_publish_date_bogota(meta['date'])}")
            st.caption(f"Commit: `{meta['sha']}`")
        else:
            st.warning("Configurado pero sin datos publicados todavia. "
                       "Pide al admin que cargue datos en `evo-debitos`.")
    else:
        st.error("NO configurado. Faltan secrets `GITHUB_TOKEN` y `GITHUB_REPO`.")
    st.markdown("---")
    st.caption("Vista de Gerentes Generales. La sede es obligatoria.")


# =========================================================
# Header
# =========================================================
st.markdown("# 🏢 EVO - Reporte de Debitos (Gerentes Generales)")
st.caption("La fuente se selecciona automaticamente: API si responde, "
           "si no, los datos publicados por el admin.")


# =========================================================
# Source picker (manual). NO se llama a la API hasta que el usuario clickee.
# =========================================================
if "source" not in st.session_state:
    st.markdown("### 1. Selecciona la fuente de datos")
    st.caption("La API NO se consulta automaticamente. Elige una opcion para cargar.")

    col_shared, col_api = st.columns(2, gap="large")

    # --- Datos del admin (primero, mas comun) ---
    with col_shared:
        st.markdown("#### 🛰  Datos del admin")
        meta = shared_cache.remote_meta() if shared_cache.is_configured() else None
        if meta:
            pretty = format_publish_date_bogota(meta["date"])
            st.markdown(f"**Última publicación:**  \n{pretty}")
            st.caption(f"Commit: `{meta['sha']}`")
        elif shared_cache.is_configured():
            st.caption("Cache configurado pero el admin no ha publicado aun.")
        else:
            st.caption("Cache NO configurado. Pide secrets `GITHUB_TOKEN` y `GITHUB_REPO`.")
        if st.button("Usar datos del admin", type="primary", use_container_width=True,
                     key="btn_connect_shared",
                     disabled=not shared_cache.is_configured()):
            with st.spinner("Descargando dataset del admin..."):
                try:
                    pq_path_sh = shared_pull_path()
                    if not pq_path_sh:
                        st.error("No hay datos publicados todavia.")
                    else:
                        sede_col, sedes_list = parquet_list_sedes(pq_path_sh)
                        if not sedes_list:
                            st.error("El parquet del admin no tiene columna de sede.")
                        else:
                            st.session_state["source"] = "shared"
                            st.session_state["shared_pq_path"] = pq_path_sh
                            st.session_state["shared_sede_col"] = sede_col
                            st.session_state["sedes_shared"] = sedes_list
                            st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    # --- API EVO ---
    with col_api:
        st.markdown("#### 🌐 API EVO")
        st.caption("Rango por dias (max 5). `Hasta` es exclusivo.")
        _today = date.today()
        cda, cdb = st.columns(2)
        with cda:
            gm_api_from = st.date_input("Desde", value=_today - timedelta(days=1),
                                        max_value=_today, key="gm_api_from")
        with cdb:
            gm_api_to = st.date_input("Hasta (exclusivo)", value=_today,
                                      max_value=_today + timedelta(days=1),
                                      min_value=gm_api_from + timedelta(days=1),
                                      key="gm_api_to")
        _rng, _err = gr.validate_date_range(gm_api_from.isoformat(), gm_api_to.isoformat())
        if _err:
            st.caption(f"⚠ {_err}")
        else:
            _chunks = gr.split_date_ranges_by_month(_rng[0], _rng[1])
            if len(_chunks) > 1:
                st.caption(f"ℹ Rango cruza mes: {len(_chunks)} llamadas + concat.")
        if st.button("Conectar", use_container_width=True, key="btn_connect_api",
                     disabled=bool(_err)):
            with st.spinner(f"Consultando API {gm_api_from} .. {gm_api_to}..."):
                try:
                    pq_path_api, chunks_log = cached_fetch_api_range(
                        api_url, gm_api_from.isoformat(), gm_api_to.isoformat()
                    )
                    sede_col, sedes_list = parquet_list_sedes(pq_path_api)
                    if not sedes_list:
                        st.error("La API respondio pero sin columna de sede.")
                    else:
                        st.session_state["source"] = "api"
                        st.session_state["api_from"] = gm_api_from.isoformat()
                        st.session_state["api_to"] = gm_api_to.isoformat()
                        st.session_state["api_chunks_log"] = chunks_log
                        st.session_state["shared_pq_path"] = pq_path_api
                        st.session_state["shared_sede_col"] = sede_col
                        st.session_state["sedes_shared"] = sedes_list
                        st.rerun()
                except Exception as e:
                    st.error(f"No se pudo cargar la API: {e}")
    st.stop()


# =========================================================
# Banner de fuente
# =========================================================
source = st.session_state["source"]
if source == "api":
    f_label = st.session_state.get("api_from", "?")
    t_label = st.session_state.get("api_to", "?")
    st.success(f"Fuente: **API EVO** ({f_label} a {t_label}, exclusivo).")
    if st.session_state.get("api_chunks_log"):
        with st.expander("Detalle de llamadas a la API"):
            for line in st.session_state["api_chunks_log"]:
                st.text(line)
elif source == "shared":
    meta = shared_cache.remote_meta()
    msg = "Fuente: **Datos publicados por el admin**."
    if meta:
        msg += f"  ·  Última publicación: {format_publish_date_bogota(meta['date'])} (`{meta['sha']}`)"
    st.info(msg)
elif source == "file":
    st.info("Fuente: **Excel local** (manual).")


# =========================================================
# Step - sede picker
# =========================================================
sedes_options = st.session_state.get("sedes_shared", [])

if not sedes_options:
    st.error("No se obtuvieron sedes desde la fuente activa.")
    if st.button("Reintentar"):
        reset_state()
        st.cache_data.clear()
        st.rerun()
    st.stop()

st.markdown("### Selecciona la sede")
sede_sel = st.selectbox(
    "Sede / Club (obligatorio)",
    ["-- Selecciona una sede --"] + sedes_options,
    index=0,
    key="sede_pick",
)

if sede_sel == "-- Selecciona una sede --":
    st.info(f"Hay {len(sedes_options)} sedes disponibles. Elige una para continuar.")
    if st.button("Cambiar fuente"):
        reset_state()
        st.rerun()
    st.stop()


# =========================================================
# Carga filtrada
# =========================================================
with st.spinner(f"Cargando registros de '{sede_sel}'..."):
    pq_path = st.session_state["shared_pq_path"]
    sede_col = st.session_state["shared_sede_col"]
    df = parquet_filter_by_sede(pq_path, sede_col, sede_sel)

if df is None or df.empty:
    st.warning(f"No hay registros para la sede '{sede_sel}'.")
    if st.button("Cambiar sede"):
        st.session_state.pop("sede_pick", None)
        st.rerun()
    st.stop()

count_str = f"{len(df):,}".replace(",", ".")
st.markdown(
    f"<div class='sede-banner'>Sede activa: <b>{sede_sel}</b> &middot; "
    f"Registros cargados: <b>{count_str}</b></div>",
    unsafe_allow_html=True,
)


# =========================================================
# Date filter + pipeline
# =========================================================
cols = gr.resolve_columns(df)
date_col = cols.get("intento")
if date_col:
    try:
        dts = pd.to_datetime(df[date_col], errors="coerce", utc=True)
        if getattr(dts.dt, "tz", None) is not None:
            dts = dts.dt.tz_convert("UTC").dt.tz_localize(None)
        dts = dts.dropna()
    except Exception:
        dts = pd.to_datetime(df[date_col], errors="coerce").dropna()
    if len(dts):
        fmin = dts.min().date()
        fmax = dts.max().date()
    else:
        fmin = date.today() - timedelta(days=30)
        fmax = date.today()
else:
    fmin = fmax = date.today()

# Banner con rango real de los datos cargados + fecha de publicacion del admin
range_msg = f"Rango disponible en los datos: **{fmin}** a **{fmax}**"
if source == "shared":
    meta = shared_cache.remote_meta()
    if meta:
        range_msg += f"  ·  Publicado por admin: {format_publish_date_bogota(meta['date'])}"
elif source == "api":
    range_msg += "  ·  Origen: API EVO en vivo"
elif source == "file":
    range_msg += "  ·  Origen: Excel local manual"
range_msg += "  ·  Si no es el rango esperado, presiona **Recargar datos** en la sidebar."
st.caption(range_msg)

c1, c2 = st.columns(2)
with c1:
    since = st.date_input("Desde", value=fmin, min_value=fmin, max_value=fmax)
with c2:
    until = st.date_input("Hasta", value=fmax, min_value=fmin, max_value=fmax)

with st.spinner("Calculando metricas..."):
    df_f, cols, summary, sedes, motivos, tipos, marcas, daily = run_pipeline_for_sede(
        df, sede_sel, since.isoformat(), until.isoformat())

if df_f.empty or summary.get("total", 0) == 0:
    st.warning(f"No hay registros para '{sede_sel}' entre {since} y {until}. "
               "Amplia el rango de fechas para continuar.")
    st.stop()


# =========================================================
# KPIs
# =========================================================
st.markdown("### Resultados")

def _has_clients() -> bool:
    return summary.get("clients_total") is not None

k1, k2, k3, k4 = st.columns(4)
with k1:
    if _has_clients():
        metric_card(
            "Intentos",
            gr.fmt_int(summary["clients_total"]),
            sub=f"usuarios unicos<br>{gr.fmt_int(summary['total'])} operaciones",
        )
    else:
        metric_card("Intentos", gr.fmt_int(summary["total"]), sub=sede_sel)
with k2:
    if _has_clients():
        ct = max(summary.get("clients_total") or 1, 1)
        client_succ = (summary["clients_approved"] / ct)
        metric_card(
            "Aprobados",
            gr.fmt_int(summary["clients_approved"]),
            sub=(f"{gr.fmt_pct(client_succ)} de usuarios unicos"
                 f"<br>{gr.fmt_int(summary['approved'])} operaciones "
                 f"({gr.fmt_pct(summary['success_rate'])})"),
            kind="good",
        )
    else:
        metric_card("Aprobados", gr.fmt_int(summary["approved"]),
                    sub=gr.fmt_pct(summary["success_rate"]), kind="good")
with k3:
    if _has_clients():
        ct = max(summary.get("clients_total") or 1, 1)
        client_fail = (summary["clients_denied"] / ct)
        metric_card(
            "Negados",
            gr.fmt_int(summary["clients_denied"]),
            sub=(f"{gr.fmt_pct(client_fail)} de usuarios unicos"
                 f"<br>{gr.fmt_int(summary['denied'])} operaciones "
                 f"({gr.fmt_pct(summary['fail_rate'])})"),
            kind="bad",
        )
    else:
        metric_card("Negados", gr.fmt_int(summary["denied"]),
                    sub=gr.fmt_pct(summary["fail_rate"]), kind="bad")
with k4:
    if "amount_approved" in summary:
        metric_card("Recuperado", gr.fmt_money(summary["amount_approved"]),
                    sub=f"En riesgo: {gr.fmt_money(summary.get('amount_denied', 0))}",
                    kind="gold")


# ---------- Usuarios nunca aprobados (nuevo, additivo) ----------
if summary.get("users_never_approved") is not None and summary.get("clients_total"):
    nv = summary["users_never_approved"]
    ev = summary["users_ever_approved"]
    tot = summary["clients_total"]
    nv_pct = (nv / tot) * 100 if tot else 0
    ev_pct = (ev / tot) * 100 if tot else 0
    st.markdown(
        f"""<div style='background:#FFF8E1; border-left:4px solid #E67E22;
        padding:10px 14px; border-radius:6px; margin-top:10px; margin-bottom:6px;'>
        <b>Usuarios nunca aprobados:</b> {gr.fmt_int(nv)} de {gr.fmt_int(tot)}
        ({nv_pct:.1f}%) &middot;
        <span style='color:#7A8597;'>Alguna vez aprobados: {gr.fmt_int(ev)} ({ev_pct:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )

# =========================================================
# Donut + tendencia
# =========================================================
left, right = st.columns([1.2, 2])
with left:
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(4.5, 3.0))
    pie_total = (summary.get("approved", 0) or 0) + (summary.get("denied", 0) or 0)
    if pie_total <= 0:
        ax.text(0.5, 0.5, "Sin intentos clasificados", ha="center", va="center",
                fontsize=10, color="#7A8597", transform=ax.transAxes)
        ax.set_axis_off()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        ax.pie([summary["approved"], summary["denied"]],
               colors=["#1F8A4C", "#C0392B"], startangle=90,
               wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
        ax.text(0, 0.05, f"{summary['success_rate']*100:.1f}%",
                ha="center", va="center", fontsize=22, fontweight="bold", color="#0F2A4A")
        ax.text(0, -0.18, "Tasa de exito", ha="center", va="center",
                fontsize=9, color="#7A8597")
        ax.legend(["Aprobado", "Negado"], loc="lower center",
                  bbox_to_anchor=(0.5, -0.05), ncol=2, frameon=False)
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
with right:
    if not daily.empty and len(daily) > 1:
        import matplotlib.pyplot as plt
        fig, ax1 = plt.subplots(figsize=(8, 3))
        ax1.bar(daily["Dia"], daily["Aprobado"], color="#1F8A4C", label="Aprobado", alpha=0.85)
        ax1.bar(daily["Dia"], daily["Negado"], bottom=daily["Aprobado"],
                color="#C0392B", label="Negado", alpha=0.85)
        ax1.set_ylabel("Intentos")
        ax2 = ax1.twinx()
        ax2.plot(daily["Dia"], daily["TasaExito"]*100,
                 color="#0F2A4A", linewidth=2, marker="o", markersize=3)
        ax2.set_ylabel("Tasa exito (%)")
        ax2.set_ylim(0, 100)
        ax1.set_title("Tendencia diaria")
        fig.autofmt_xdate()
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        st.info("No hay suficientes dias para mostrar tendencia.")


# =========================================================
# Motivos
# =========================================================
st.markdown("### Motivos de rechazo")
if not motivos.empty:
    table = motivos.head(15)[
        ["Ranking", "MotivoES", "Veces", "Pct", "Responsable", "Accion"]
    ].copy()
    table.columns = ["#", "Motivo (ES)", "Veces", "% del total negado",
                     "Responsable", "Accion sugerida"]
    table["Veces"] = table["Veces"].apply(gr.fmt_int)
    table["% del total negado"] = (motivos.head(15)["Pct"] * 100).round(2).astype(str) + "%"
    st.dataframe(table, use_container_width=True, hide_index=True)
else:
    st.info("No hay motivos negados en este corte.")


# =========================================================
# Descargas
# =========================================================
st.markdown("### Descargar")
col_a, col_b = st.columns(2)
with col_a:
    if st.button("Generar PDF", type="primary", use_container_width=True):
        with st.spinner("Generando PDF..."):
            pdf_bytes = build_pdf_bytes(df_f, cols, summary, sedes, motivos,
                                        tipos, marcas, daily, sede_sel,
                                        since=since, until=until)
            tag = "_" + gr.normalize_match(sede_sel).upper().replace(" ", "_")[:30]
            fname = f"Reporte_Debitos{tag}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            st.download_button("Descargar PDF", data=pdf_bytes, file_name=fname,
                               mime="application/pdf", use_container_width=True)
with col_b:
    # Excel de negados para la sede activa
    status_col = cols.get("status")
    if status_col:
        s = df_f[status_col].astype(str).str.lower()
        denied_mask = s.str.startswith("nega") | s.str.startswith("deni") | s.str.startswith("rech")
        denied = df_f.loc[denied_mask].copy()
    else:
        denied = df_f.iloc[0:0]

    # Build working frame with the raw columns we need
    work = pd.DataFrame()
    id_candidates = ["id", "ID", "Id", "user_id", "userId", "userid",
                     "id_usuario", "ID_usuario", "idUsuario", "Cliente_id", "id_cliente"]
    id_col = next((c for c in id_candidates if c in denied.columns), None)
    if id_col:
        work["id"] = denied[id_col].astype(str).values
    else:
        work["id"] = [str(i) for i in range(1, len(denied) + 1)]
    work["Cliente"] = denied[cols["cliente"]].astype(str).values if "cliente" in cols else ""

    # Detectar columnas de contacto que el endpoint ahora retorna al final
    EMAIL_CANDS = ["email", "Email", "EMAIL", "correo", "Correo", "mail", "e-mail"]
    PHONE_CANDS = ["phone", "Phone", "PHONE", "telefono", "Telefono", "TELEFONO",
                   "celular", "Celular", "movil", "Movil", "tel"]
    email_col = next((c for c in EMAIL_CANDS if c in denied.columns), None)
    phone_col = next((c for c in PHONE_CANDS if c in denied.columns), None)
    work["email"] = denied[email_col].astype(str).values if email_col else ""
    work["phone"] = denied[phone_col].astype(str).values if phone_col else ""

    if "intento" in cols:
        work["Intento"] = pd.to_datetime(denied[cols["intento"]], errors="coerce").values
    else:
        work["Intento"] = pd.NaT
    work["Motivo del rechazo"] = denied[cols["motivo"]].astype(str).values if "motivo" in cols else ""

    # Helper para "first non-null/non-empty" en columnas de contacto
    def _first_non_empty(series: pd.Series):
        s = series.astype(str).replace({"nan": "", "None": "", "NaN": ""}).str.strip()
        for v in s:
            if v:
                return v
        return ""

    # Aggregate: one row per (id, Motivo del rechazo).
    # Total intentos = numero de veces que ese usuario tuvo ese motivo en el rango activo.
    # Cliente / email / phone: primer valor no-vacio del grupo. Intento: max del grupo.
    if len(work) > 0:
        out_df = (
            work.groupby(["id", "Motivo del rechazo"], dropna=False, sort=False)
            .agg(
                Cliente=("Cliente", "first"),
                email=("email", _first_non_empty),
                phone=("phone", _first_non_empty),
                Intento=("Intento", "max"),
                Total_intentos=("Motivo del rechazo", "size"),
            )
            .reset_index()
            .rename(columns={"Total_intentos": "Total intentos"})
        )
        out_df = out_df[["id", "Cliente", "email", "phone", "Intento",
                         "Motivo del rechazo", "Total intentos"]]
        out_df = out_df.sort_values(
            ["Total intentos", "id"], ascending=[False, True]
        ).reset_index(drop=True)
    else:
        out_df = pd.DataFrame(
            columns=["id", "Cliente", "email", "phone", "Intento",
                     "Motivo del rechazo", "Total intentos"]
        )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Negados")
        ws = writer.sheets["Negados"]
        # Auto-width simple
        for col_idx, col_name in enumerate(out_df.columns, start=1):
            max_len = max(
                [len(str(col_name))]
                + [len(str(v)) for v in out_df[col_name].head(200).tolist()]
            )
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 50)
    xlsx_bytes = buf.getvalue()

    sede_tag = gr.normalize_match(sede_sel).upper().replace(" ", "_")[:30]
    fname_xlsx = f"Negados_{sede_tag}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        f"Descargar negados (Excel) - {len(out_df):,} filas".replace(",", "."),
        data=xlsx_bytes,
        file_name=fname_xlsx,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown(
    "<div style='font-size:12px; color:#7A8597; margin-top:6px;'>"
    "ⓘ El Excel puede contener usuarios duplicados porque tuvieron mas de un motivo de rechazo."
    "</div>",
    unsafe_allow_html=True,
)
st.caption(f"Sede: {sede_sel}  ·  Periodo: {since} a {until}  ·  "
           f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
