# -*- coding: utf-8 -*-
"""
EVO - Reporte de Debitos - APP DE GERENTES GENERALES
Flujo:
  1. Elegir fuente (API o Excel).
  2. Listar sedes (lectura ligera de la columna Sede/club).
  3. Elegir sede (obligatorio).
  4. Cargar SOLO los datos de esa sede.
  5. Generar el mismo reporte/PDF que app.py.

Optimizaciones clave:
  - Excel se convierte a parquet una unica vez (cacheado por hash de bytes).
  - El parquet se filtra por sede usando pyarrow (lazy, predicate pushdown).
  - La columna de sede se lee de forma aislada para listar opciones rapido.
  - st.cache_data reusa resultados entre interacciones del mismo usuario.

Deploy: en Streamlit Community Cloud, configurar
"Main file path" = `app_managers.py`. Pueden coexistir las dos apps.
"""
from __future__ import annotations

import hashlib
import os
import tempfile
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

import generate_report as gr

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
</style>
""", unsafe_allow_html=True)


SEDE_COL_CANDIDATES = ["Sede/club", "Sede", "Club", "sede", "club", "sede_club", "branch"]


# =========================================================
# CACHE LAYER
# =========================================================
@st.cache_data(ttl=1800, show_spinner=False)
def excel_bytes_to_parquet(file_bytes: bytes, name: str) -> str:
    """One-shot Excel -> parquet conversion. Returns parquet path on disk.

    Hash of bytes short-circuits re-uploads of the same file.
    Subsequent reads use pyarrow column/predicate pushdown.
    """
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
        for c in df.select_dtypes(include="object").columns:
            try:
                df[c] = df[c].astype("string")
            except Exception:
                pass
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
    return pd.read_parquet(pq_path, filters=[(sede_col, "=", sede)])


@st.cache_data(ttl=300, show_spinner=False)
def api_list_sedes(api_url: str, api_key: str):
    df = gr.fetch_from_api(api_url, api_key)
    cols = gr.resolve_columns(df)
    if "sede" not in cols:
        return []
    return sorted(df[cols["sede"]].dropna().astype(str).unique().tolist())


@st.cache_data(ttl=300, show_spinner=False)
def api_load_for_sede(api_url: str, api_key: str, sede: str) -> pd.DataFrame:
    """Try to push sede filter to API; fallback to client-side filter if needed."""
    try:
        df = gr.fetch_from_api(api_url, api_key, sede=sede)
    except Exception:
        df = gr.fetch_from_api(api_url, api_key)
    cols = gr.resolve_columns(df)
    if "sede" in cols and len(df):
        from generate_report import normalize_match
        target = normalize_match(sede)
        mask = df[cols["sede"]].astype(str).map(normalize_match) == target
        if mask.sum() > 0:
            df = df.loc[mask].copy()
    return df


# =========================================================
# Pipeline
# =========================================================
def run_pipeline_for_sede(df, sede_label, since, until):
    cols = gr.resolve_columns(df)
    df = gr.coerce(df, cols)
    if "intento" in cols and (since or until):
        d = pd.to_datetime(df[cols["intento"]], errors="coerce")
        if since:
            df = df[d >= pd.to_datetime(since)]
            d = pd.to_datetime(df[cols["intento"]], errors="coerce")
        if until:
            df = df[d <= pd.to_datetime(until) + pd.Timedelta(days=1)]
    summary, ap, dn = gr.compute_summary(df, cols)
    sedes = gr.by_sede(df, cols, ap, dn)
    motivos = gr.by_motivo(df, cols, dn)
    tipos = gr.by_tipo(df, cols, dn)
    marcas = gr.by_marca(df, cols, ap, dn)
    daily = gr.by_day(df, cols, ap, dn)
    return df, cols, summary, sedes, motivos, tipos, marcas, daily


def build_pdf_bytes(df, cols, summary, sedes, motivos, tipos, marcas, daily, sede_label):
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
    for k in list(st.session_state.keys()):
        st.session_state.pop(k, None)


# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.title("Configuracion")
    st.caption("Solo administradores.")
    api_url = st.text_input("URL del endpoint", value=gr.DEFAULT_API_URL)
    api_key = st.text_input("API Key", value=gr.DEFAULT_API_KEY, type="password")
    if st.button("Reiniciar todo", use_container_width=True):
        reset_state()
        st.rerun()
    st.markdown("---")
    st.caption("Vista de Gerentes Generales. La sede es obligatoria; "
               "esto reduce el tiempo de carga.")


# =========================================================
# Header
# =========================================================
st.markdown("# 🏢 EVO - Reporte de Debitos (Gerentes Generales)")
st.caption("Selecciona la sede ANTES de cargar. Solo se procesan los datos de esa sede.")


# =========================================================
# Step 1 - source picker
# =========================================================
if "source" not in st.session_state:
    st.markdown("### 1. ¿De donde tomamos los datos?")
    col_a, col_b = st.columns(2, gap="large")
    with col_a:
        st.markdown("#### 🌐 API EVO")
        st.caption("Recomendado. Filtra en el servidor.")
        if st.button("Usar API", type="primary", use_container_width=True, key="src_api"):
            st.session_state["source"] = "api"
            st.rerun()
    with col_b:
        st.markdown("#### 📄 Excel")
        st.caption("Sube .xlsx/.xlsm con hoja `data`.")
        if st.button("Subir Excel", use_container_width=True, key="src_xls"):
            st.session_state["source"] = "file"
            st.rerun()
    st.stop()

source = st.session_state["source"]


# =========================================================
# Step 2 - list sedes
# =========================================================
sedes_options = []
sede_col_name = None
pq_path = None

if source == "api":
    if "sedes_api" not in st.session_state:
        with st.spinner("Listando sedes desde la API..."):
            try:
                st.session_state["sedes_api"] = api_list_sedes(api_url, api_key)
            except Exception as e:
                st.error(f"Error consultando la API: {e}")
                if st.button("Volver"):
                    reset_state()
                    st.rerun()
                st.stop()
    sedes_options = st.session_state.get("sedes_api", [])

elif source == "file":
    uploaded = st.file_uploader("Archivo .xlsx", type=["xlsx", "xlsm"],
                                key="upl",
                                label_visibility="collapsed")
    if uploaded is None:
        st.info("Sube el archivo de debitos para continuar.")
        if st.button("Volver"):
            reset_state()
            st.rerun()
        st.stop()

    with st.spinner("Indexando archivo (una sola vez por archivo)..."):
        pq_path = excel_bytes_to_parquet(uploaded.getvalue(), uploaded.name)
        sede_col_name, sedes_options = parquet_list_sedes(pq_path)

    if not sede_col_name:
        st.error("No se encontro columna de sede en el archivo. "
                 "Columnas esperadas: Sede/club, Sede, Club.")
        if st.button("Volver"):
            reset_state()
            st.rerun()
        st.stop()


# =========================================================
# Step 3 - sede picker (BLOCKS until chosen)
# =========================================================
if not sedes_options:
    st.error("No se obtuvieron sedes. Verifica la fuente de datos.")
    if st.button("Volver"):
        reset_state()
        st.rerun()
    st.stop()

st.markdown("### 2. Selecciona la sede")
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
# Step 4 - load filtered data
# =========================================================
with st.spinner(f"Cargando registros de '{sede_sel}'..."):
    if source == "api":
        df = api_load_for_sede(api_url, api_key, sede_sel)
    else:
        df = parquet_filter_by_sede(pq_path, sede_col_name, sede_sel)

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
# Step 5 - date filter + pipeline
# =========================================================
cols = gr.resolve_columns(df)
date_col = cols.get("intento")
if date_col:
    dts = pd.to_datetime(df[date_col], errors="coerce").dropna()
    if len(dts):
        fmin = dts.min().date()
        fmax = dts.max().date()
    else:
        fmin = date.today() - timedelta(days=30)
        fmax = date.today()
else:
    fmin = fmax = date.today()

c1, c2 = st.columns(2)
with c1:
    since = st.date_input("Desde", value=fmin, min_value=fmin, max_value=fmax)
with c2:
    until = st.date_input("Hasta", value=fmax, min_value=fmin, max_value=fmax)

with st.spinner("Calculando metricas..."):
    df_f, cols, summary, sedes, motivos, tipos, marcas, daily = run_pipeline_for_sede(
        df, sede_sel, since.isoformat(), until.isoformat())


# =========================================================
# KPIs
# =========================================================
st.markdown("### Resultados")
k1, k2, k3, k4 = st.columns(4)
with k1:
    metric_card("Intentos", gr.fmt_int(summary["total"]), sub=sede_sel)
with k2:
    metric_card("Aprobados", gr.fmt_int(summary["approved"]),
                sub=gr.fmt_pct(summary["success_rate"]), kind="good")
with k3:
    metric_card("Negados", gr.fmt_int(summary["denied"]),
                sub=gr.fmt_pct(summary["fail_rate"]), kind="bad")
with k4:
    if "amount_approved" in summary:
        metric_card("Recuperado", gr.fmt_money(summary["amount_approved"]),
                    sub=f"En riesgo: {gr.fmt_money(summary.get('amount_denied', 0))}",
                    kind="gold")


# =========================================================
# Donut + tendencia
# =========================================================
left, right = st.columns([1.2, 2])
with left:
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(4.5, 3.0))
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
st.markdown("### Motivos de rechazo (en espanol, con accion sugerida)")
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
                                        tipos, marcas, daily, sede_sel)
            tag = "_" + gr.normalize_match(sede_sel).upper().replace(" ", "_")[:30]
            fname = f"Reporte_Debitos{tag}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            st.download_button("Descargar PDF", data=pdf_bytes, file_name=fname,
                               mime="application/pdf", use_container_width=True)
with col_b:
    md_lines = ["# Motivos de rechazo - acciones", "",
                f"Reporte para: **{sede_sel}**", "",
                "| # | Motivo (ES) | Veces | % | Responsable | Accion |",
                "|---|---|---:|---:|---|---|"]
    for _, r in motivos.head(20).iterrows():
        accion = (r["Accion"] or "").replace("|", "/")
        es = (r["MotivoES"] or "").replace("|", "/")
        md_lines.append(f"| {r['Ranking']} | {es} | {gr.fmt_int(r['Veces'])} | "
                        f"{r['Pct']*100:.1f}% | {r['Responsable']} | {accion} |")
    md = "\n".join(md_lines).encode("utf-8")
    st.download_button("Descargar catalogo Motivos.md", data=md,
                       file_name="Motivos_Acciones.md", mime="text/markdown",
                       use_container_width=True)

st.caption(f"Sede: {sede_sel}  ·  Periodo: {since} a {until}  ·  "
           f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
