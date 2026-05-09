# -*- coding: utf-8 -*-
"""
EVO - Reporte de Debitos (web)
App Streamlit que envuelve generate_report.py.
Quien la usa solo abre la URL, escoge sede y fecha, y descarga el PDF.

Deploy en Streamlit Community Cloud (gratis):
  1. Subir este repo a GitHub.
  2. https://share.streamlit.io > New app > seleccionar repo > main file: app.py
  3. Listo. Cualquier persona con la URL puede usarlo desde Windows o macOS.
"""
from __future__ import annotations
import io
import os
import tempfile
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

import generate_report as gr
from motivos import MOTIVO_CATALOG, lookup_motivo, translate_motivo
import shared_cache


st.set_page_config(
    page_title="EVO - Reporte de Debitos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Estilos ----------
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
</style>
""", unsafe_allow_html=True)


# ---------- Helpers ----------
@st.cache_data(ttl=600, show_spinner=False)
def load_from_api(api_url, month):
    """Fetch via v2 endpoint with month + multi-header. Token from env."""
    df = gr.fetch_from_api_v2(api_url, gr.DEFAULT_API_KEY, month=month)
    # Vectorized strip on string columns (avoid pyarrow byte-mismatch downstream)
    for c in df.columns:
        try:
            s = df[c]
            if s.dtype == object or str(s.dtype) == "string":
                df[c] = s.astype("string").str.strip()
        except Exception:
            continue
    return df


@st.cache_data(ttl=600, show_spinner=False)
def load_from_upload(file_bytes, name):
    with tempfile.NamedTemporaryFile(suffix=os.path.splitext(name)[1] or ".xlsx", delete=False) as f:
        f.write(file_bytes)
        path = f.name
    df = gr.load_data([path])
    return df


def metric_card(label, value, sub=None, kind="default"):
    klass = {"good": "metric-card good", "bad": "metric-card bad", "gold": "metric-card gold"}.get(kind, "metric-card")
    sub_html = f"<div class='metric-sub'>{sub}</div>" if sub else ""
    st.markdown(
        f"""
        <div class="{klass}">
          <div class="metric-value">{value}</div>
          <div class="metric-label">{label}</div>
          {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def run_pipeline(df, args_sede, args_since, args_until):
    cols = gr.resolve_columns(df)
    df, sede_label = gr.filter_by_sede(df, cols, args_sede)
    df = gr.coerce(df, cols)
    if args_since or args_until and "intento" in cols:
        d = pd.to_datetime(df[cols["intento"]], errors="coerce")
        if args_since:
            df = df[d >= pd.to_datetime(args_since)]
            d = pd.to_datetime(df[cols["intento"]], errors="coerce")
        if args_until:
            df = df[d <= pd.to_datetime(args_until) + pd.Timedelta(days=1)]
    summary, ap, dn = gr.compute_summary(df, cols)
    sedes = gr.by_sede(df, cols, ap, dn)
    motivos = gr.by_motivo(df, cols, dn)
    tipos = gr.by_tipo(df, cols, dn)
    marcas = gr.by_marca(df, cols, ap, dn)
    daily = gr.by_day(df, cols, ap, dn)
    return df, cols, summary, sedes, motivos, tipos, marcas, daily, sede_label


def robust_period(df, cols):
    """Returns (date_min, date_max) ignoring 1% of outliers on each tail."""
    if "intento" not in cols:
        return None, None
    dts = pd.to_datetime(df[cols["intento"]], errors="coerce").dropna()
    if dts.empty:
        return None, None
    if len(dts) >= 50:
        lo = dts.quantile(0.01)
        hi = dts.quantile(0.99)
    else:
        lo, hi = dts.min(), dts.max()
    return lo, hi


def build_pdf_bytes(df, cols, summary, sedes, motivos, tipos, marcas, daily, sede_label):
    # Periodo robusto en el header: ignora fechas aisladas fuera del rango real
    lo, hi = robust_period(df, cols)
    if lo is not None:
        summary = dict(summary)
        summary["date_min"] = lo
        summary["date_max"] = hi
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


# ---------- Sidebar (configuracion avanzada) ----------
with st.sidebar:
    st.title("⚙ Configuracion avanzada")
    st.caption("Solo administradores. Token gestionado por el backend (env / st.secrets).")
    api_url = st.text_input("URL del endpoint",
                            value=gr.DEFAULT_API_URL,
                            help="Endpoint que devuelve la data de debitos en JSON.")
    if st.button("Limpiar fuente y empezar de nuevo", use_container_width=True):
        for k in ("data_df", "data_source", "data_label", "data_err",
                  "publish_msg", "publish_ok"):
            st.session_state.pop(k, None)
        st.rerun()


# ---------- Header ----------
st.markdown("# 📊 EVO - Reporte de Debitos")
st.caption("Reporte gerencial de descuentos automaticos en una pagina. Filtra por sede y por fecha.")

# ---------- Source picker ----------
df = st.session_state.get("data_df")
src_label = st.session_state.get("data_label")

if df is None:
    st.markdown("### ¿De donde tomamos los datos?")
    st.write("Elige una de las dos opciones para empezar.")

    col_api, col_xls = st.columns(2, gap="large")

    with col_api:
        st.markdown("#### 🌐 Cargar desde la API EVO")
        st.caption(f"Endpoint configurado: `{api_url}`")
        months = gr.available_months(12)
        api_month = st.selectbox(
            "Mes a consultar",
            months,
            index=0,
            key="admin_api_month",
            help="Periodo que se pedira al endpoint (`?month=YYYY-MM`).",
        )
        if st.button("Conectar", type="primary", use_container_width=True, key="btn_api"):
            with st.spinner(f"Consultando API ({api_month})..."):
                try:
                    df_loaded = load_from_api(api_url, api_month)
                    st.session_state["data_df"] = df_loaded
                    st.session_state["data_source"] = "api"
                    st.session_state["data_label"] = f"API mes {api_month} ({len(df_loaded):,} registros)"
                    st.session_state.pop("data_err", None)
                    try:
                        ok, msg = shared_cache.publish_dataframe(
                            df_loaded.copy(), source_label=f"API mes {api_month}"
                        )
                        st.session_state["publish_msg"] = msg
                        st.session_state["publish_ok"] = ok
                    except Exception as _e:
                        st.session_state["publish_msg"] = f"Error publicando: {_e}"
                        st.session_state["publish_ok"] = False
                    st.rerun()
                except Exception as e:
                    st.session_state["data_err"] = str(e)
                    st.rerun()
        if st.session_state.get("data_err"):
            st.error(f"No fue posible cargar los datos desde la API:\n\n{st.session_state['data_err']}\n\n"
                     "Usa la opcion de la derecha para subir un archivo Excel.")

    with col_xls:
        st.markdown("#### 📄 Subir un archivo Excel")
        st.caption("Acepta .xlsx o .xlsm con hoja `data`.")
        uploaded = st.file_uploader("Archivo .xlsx", type=["xlsx", "xlsm"],
                                    label_visibility="collapsed")
        if uploaded is not None:
            with st.spinner("Procesando archivo..."):
                df_loaded = load_from_upload(uploaded.read(), uploaded.name)
                st.session_state["data_df"] = df_loaded
                st.session_state["data_source"] = "file"
                st.session_state["data_label"] = f"Archivo: {uploaded.name} ({len(df_loaded):,} registros)"
                st.session_state.pop("data_err", None)
                try:
                    ok, msg = shared_cache.publish_dataframe(df_loaded.copy(), source_label=f"Excel:{uploaded.name}")
                    st.session_state["publish_msg"] = msg
                    st.session_state["publish_ok"] = ok
                except Exception as _e:
                    st.session_state["publish_msg"] = f"Error publicando: {_e}"
                    st.session_state["publish_ok"] = False
                st.rerun()

    st.stop()

# Tenemos data: muestra la fuente y opcion de cambiar
src_col1, src_col2 = st.columns([5, 1])
with src_col1:
    src_kind = st.session_state.get("data_source", "")
    icon = "🌐" if src_kind == "api" else "📄"
    st.success(f"{icon}  **Fuente activa:** {src_label}")
with src_col2:
    if st.button("Cambiar fuente", use_container_width=True):
        for k in ("data_df", "data_source", "data_label", "data_err", "publish_msg", "publish_ok"):
            st.session_state.pop(k, None)
        st.rerun()
if st.session_state.get("publish_msg"):
    if st.session_state.get("publish_ok"):
        st.success("🛰  " + st.session_state["publish_msg"])
    else:
        st.error("⚠  Cache compartido no funciona: " + st.session_state["publish_msg"]
                 + "  ·  Revisa los secrets `GITHUB_TOKEN` y `GITHUB_REPO` en Streamlit -> Settings -> Secrets.")

cols = gr.resolve_columns(df)
if "sede" not in cols or "status" not in cols:
    st.error(f"La data no contiene las columnas esperadas. Detectadas: {list(df.columns)[:20]}")
    st.stop()

# ---------- Filtros ----------
sedes_disponibles = ["Todas las sedes"] + sorted(df[cols["sede"]].dropna().astype(str).unique().tolist())
date_col = cols.get("intento")
if date_col:
    dts = pd.to_datetime(df[date_col], errors="coerce").dropna()
    fmin = dts.min().date() if len(dts) else date.today() - timedelta(days=30)
    fmax = dts.max().date() if len(dts) else date.today()
else:
    fmin = fmax = date.today()

c1, c2, c3 = st.columns([2.4, 1.2, 1.2])
with c1:
    sede_sel = st.selectbox("Sede", sedes_disponibles, index=0)
with c2:
    since = st.date_input("Desde", value=fmin, min_value=fmin, max_value=fmax)
with c3:
    until = st.date_input("Hasta", value=fmax, min_value=fmin, max_value=fmax)

sede_arg = None if sede_sel == "Todas las sedes" else sede_sel

# ---------- Procesar ----------
with st.spinner("Calculando metricas..."):
    df_f, cols, summary, sedes, motivos, tipos, marcas, daily, sede_label = run_pipeline(
        df, sede_arg, since.isoformat(), until.isoformat())

# ---------- KPIs ----------
st.markdown("### Resultados")
k1, k2, k3, k4 = st.columns(4)
with k1:
    metric_card("Intentos", gr.fmt_int(summary["total"]),
                sub=f"{df_f[cols['sede']].nunique() if 'sede' in cols else 0} sedes" if not sede_arg else sede_label)
with k2:
    metric_card("Aprobados", gr.fmt_int(summary["approved"]),
                sub=gr.fmt_pct(summary["success_rate"]), kind="good")
with k3:
    metric_card("Negados", gr.fmt_int(summary["denied"]),
                sub=gr.fmt_pct(summary["fail_rate"]), kind="bad")
with k4:
    if "amount_approved" in summary:
        metric_card("Recuperado", gr.fmt_money(summary["amount_approved"]),
                    sub=f"En riesgo: {gr.fmt_money(summary.get('amount_denied', 0))}", kind="gold")

# ---------- Donut + Tendencia ----------
left, right = st.columns([1.2, 2])
with left:
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(4.5, 3.0))
    ax.pie([summary["approved"], summary["denied"]], colors=["#1F8A4C", "#C0392B"],
           startangle=90, wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    ax.text(0, 0.05, f"{summary['success_rate']*100:.1f}%", ha="center", va="center",
            fontsize=22, fontweight="bold", color="#0F2A4A")
    ax.text(0, -0.18, "Tasa de exito", ha="center", va="center", fontsize=9, color="#7A8597")
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
        ax1.bar(daily["Dia"], daily["Negado"], bottom=daily["Aprobado"], color="#C0392B", label="Negado", alpha=0.85)
        ax1.set_ylabel("Intentos")
        ax2 = ax1.twinx()
        ax2.plot(daily["Dia"], daily["TasaExito"]*100, color="#0F2A4A", linewidth=2, marker="o", markersize=3)
        ax2.set_ylabel("Tasa exito (%)")
        ax2.set_ylim(0, 100)
        ax1.set_title("Tendencia diaria")
        fig.autofmt_xdate()
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        st.info("No hay suficientes dias para mostrar tendencia.")

# ---------- Motivos con explicacion ----------
st.markdown("### 🔎 Motivos de rechazo (en espanol, con accion sugerida)")
if not motivos.empty:
    table = motivos.head(15)[["Ranking", "MotivoES", "Veces", "Pct", "Responsable", "Accion"]].copy()
    table.columns = ["#", "Motivo (ES)", "Veces", "% del total negado", "Responsable", "Accion sugerida"]
    table["Veces"] = table["Veces"].apply(gr.fmt_int)
    table["% del total negado"] = (table["% del total negado"]).apply(lambda x: f"{x*100:.2f}%") if False else (motivos.head(15)["Pct"]*100).round(2).astype(str) + "%"
    st.dataframe(table, use_container_width=True, hide_index=True)
else:
    st.info("No hay motivos negados en este corte.")

# ---------- Sedes (solo si todas) ----------
if not sede_arg and not sedes.empty:
    st.markdown("### 🏢 Desempeno por sede")
    s = sedes.copy()
    s["Tasa exito"] = (s["TasaExito"] * 100).round(1).astype(str) + "%"
    for c in ("Total", "Aprobado", "Negado"):
        s[c] = s[c].apply(gr.fmt_int)
    st.dataframe(s[["Sede", "Total", "Aprobado", "Negado", "Tasa exito"]],
                 use_container_width=True, hide_index=True)

# ---------- Descargas ----------
st.markdown("### 📥 Descargar")
col_a, col_b = st.columns([1, 1])
with col_a:
    if st.button("Generar PDF", type="primary", use_container_width=True):
        with st.spinner("Generando PDF..."):
            pdf_bytes = build_pdf_bytes(df_f, cols, summary, sedes, motivos, tipos, marcas, daily, sede_label)
            tag = ""
            if sede_arg:
                tag = "_" + gr.normalize_match(sede_label).upper().replace(" ", "_")[:30]
            fname = f"Reporte_Debitos{tag}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            st.download_button("Descargar PDF", data=pdf_bytes, file_name=fname,
                               mime="application/pdf", use_container_width=True)
with col_b:
    md_lines = ["# Motivos de rechazo - acciones", "",
                f"Reporte para: **{sede_label}**", "",
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

st.caption(f"Periodo: {since} a {until}  ·  Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
