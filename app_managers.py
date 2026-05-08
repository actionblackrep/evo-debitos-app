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
import os
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
    keys = ["source", "sedes_api", "sedes_shared", "shared_pq_path", "shared_sede_col",
            "api_error", "shared_error", "used_fallback", "sede_pick", "upl"]
    for k in keys:
        st.session_state.pop(k, None)


# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.title("Configuracion")
    st.caption("Solo administradores.")
    api_url = st.text_input("URL del endpoint", value=gr.DEFAULT_API_URL)
    api_key = st.text_input("API Key", value=gr.DEFAULT_API_KEY, type="password")
    if st.button("Recargar datos del admin", type="primary", use_container_width=True,
                 help="Borra todos los caches locales y vuelve a leer la API y el cache compartido."):
        reset_state()
        st.cache_data.clear()
        try:
            shared_pull_path.clear()
        except Exception:
            pass
        st.rerun()
    if st.button("Forzar reconexion API", use_container_width=True):
        reset_state()
        st.cache_data.clear()
        st.rerun()
    if st.button("Reiniciar todo", use_container_width=True):
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
            st.success(f"Ultima publicacion: {meta['date'][:16]}Z")
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
# Auto-detectar fuente
# =========================================================
if "source" not in st.session_state:
    api_err = None
    with st.spinner("Conectando a la API EVO..."):
        try:
            sedes_api = api_list_sedes(api_url, api_key)
            if sedes_api:
                st.session_state["source"] = "api"
                st.session_state["sedes_api"] = sedes_api
                st.rerun()
        except Exception as e:
            api_err = str(e)[:300]
            st.session_state["api_error"] = api_err

    with st.spinner("Buscando datos publicados por el admin..."):
        try:
            pq_path = shared_pull_path()
            if pq_path:
                sede_col, sedes_shared = parquet_list_sedes(pq_path)
                if sedes_shared:
                    st.session_state["source"] = "shared"
                    st.session_state["shared_pq_path"] = pq_path
                    st.session_state["shared_sede_col"] = sede_col
                    st.session_state["sedes_shared"] = sedes_shared
                    if api_err:
                        st.session_state["used_fallback"] = True
                    st.rerun()
        except Exception as e:
            st.session_state["shared_error"] = str(e)[:300]

    # Manual upload as last resort
    st.warning(
        "No hay datos disponibles automaticamente. "
        "La API no respondio y el admin aun no ha publicado un archivo."
    )
    if st.session_state.get("api_error"):
        with st.expander("Detalle del error de API"):
            st.code(st.session_state["api_error"])
    if not shared_cache.is_configured():
        st.info(
            "El cache compartido NO esta configurado. Pide al admin que agregue "
            "los secrets `GITHUB_TOKEN` y `GITHUB_REPO` en Streamlit Community Cloud."
        )
    st.markdown("#### Subir un Excel manualmente")
    uploaded = st.file_uploader("Archivo .xlsx", type=["xlsx", "xlsm"],
                                key="upl_fallback",
                                label_visibility="collapsed")
    if uploaded is not None:
        with st.spinner("Indexando archivo..."):
            pq_path = excel_bytes_to_parquet(uploaded.getvalue(), uploaded.name)
            sede_col, sedes_local = parquet_list_sedes(pq_path)
        if sedes_local:
            st.session_state["source"] = "file"
            st.session_state["shared_pq_path"] = pq_path
            st.session_state["shared_sede_col"] = sede_col
            st.session_state["sedes_shared"] = sedes_local
            st.rerun()
        else:
            st.error("No se detecto columna de sede en el archivo.")
    st.stop()


# =========================================================
# Banner de fuente
# =========================================================
source = st.session_state["source"]
if source == "api":
    st.success("Fuente: **API EVO** (datos en vivo).")
elif source == "shared":
    meta = shared_cache.remote_meta()
    if st.session_state.get("used_fallback"):
        msg = "API no disponible. Usando **datos publicados por el admin**."
    else:
        msg = "Usando **datos publicados por el admin**."
    if meta:
        msg += f"  ·  Ultima publicacion: `{meta['date'][:16]}Z` ({meta['sha']})"
    st.info(msg)
elif source == "file":
    st.info("Fuente: **Excel local** (subido manualmente, no compartido).")


# =========================================================
# Step - sede picker
# =========================================================
if source == "api":
    sedes_options = st.session_state.get("sedes_api", [])
else:
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
    if source == "api":
        df = api_load_for_sede(api_url, api_key, sede_sel)
    else:
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
        range_msg += f"  ·  Publicado por admin: `{meta['date'][:16]}Z`"
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
    # Excel de negados para la sede activa
    status_col = cols.get("status")
    if status_col:
        s = df_f[status_col].astype(str).str.lower()
        denied_mask = s.str.startswith("nega") | s.str.startswith("deni") | s.str.startswith("rech")
        denied = df_f.loc[denied_mask].copy()
    else:
        denied = df_f.iloc[0:0]

    out_df = pd.DataFrame()
    # Buscar columna user-id (no la ID de la venta)
    id_candidates = ["id", "ID", "Id", "user_id", "userId", "userid",
                     "id_usuario", "ID_usuario", "idUsuario", "Cliente_id", "id_cliente"]
    id_col = next((c for c in id_candidates if c in denied.columns), None)
    if id_col:
        out_df["id"] = denied[id_col].astype(str).values
    else:
        out_df["id"] = range(1, len(denied) + 1)
    out_df["Cliente"] = denied[cols["cliente"]].astype(str).values if "cliente" in cols else ""
    if "intento" in cols:
        out_df["Intento"] = pd.to_datetime(denied[cols["intento"]], errors="coerce").values
    else:
        out_df["Intento"] = ""
    out_df["Motivo del rechazo"] = denied[cols["motivo"]].astype(str).values if "motivo" in cols else ""

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

st.caption(f"Sede: {sede_sel}  ·  Periodo: {since} a {until}  ·  "
           f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
