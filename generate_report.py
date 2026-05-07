#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de Reporte de Debitos - EVO
Lee los archivos .xlsx en ./input/ y produce un PDF gerencial en ./output/.

Uso:
    python generate_report.py
    python generate_report.py --input "ruta/al/archivo.xlsx"
    python generate_report.py --input ./input --output ./output

Requisitos: ver requirements.txt
"""
from __future__ import annotations
import argparse
import glob
import json
import os
import re
import sys
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import pandas as pd

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Image, KeepTogether,
    PageBreak, Paragraph, Spacer, Table, TableStyle, NextPageTemplate,
)

from motivos import MOTIVO_CATALOG, lookup_motivo, translate_motivo

# ------------------ API config ------------------
DEFAULT_API_URL = os.environ.get(
    "EVO_DEBITS_URL",
    "https://action-branches-api.vercel.app/api/debitos",
)
DEFAULT_API_KEY = os.environ.get(
    "EVO_DEBITS_API_KEY",
    "dLjaU5u4LfycyRpbBTU7EMcXDBL2zFrOiX6fBWO6b-s",
)

# Brand palette
PRIMARY = colors.HexColor("#0F2A4A")
ACCENT = colors.HexColor("#D4AF37")
GOOD = colors.HexColor("#1F8A4C")
BAD = colors.HexColor("#C0392B")
WARN = colors.HexColor("#E67E22")
LIGHT = colors.HexColor("#F4F6FA")
GREY = colors.HexColor("#7A8597")

CHART_DIR = None  # set per run


def fmt_int(n):
    try:
        return f"{int(n):,}".replace(",", ".")
    except Exception:
        return str(n)


def fmt_pct(p, decimals=1):
    return f"{p*100:.{decimals}f}%"


def fmt_money(v):
    try:
        return f"${int(round(v)):,}".replace(",", ".")
    except Exception:
        return str(v)


def normalize_text(s):
    if not isinstance(s, str):
        return s
    s = unicodedata.normalize("NFKC", s)
    return s.strip()


def find_input(input_arg):
    """Resolve input path to a list of xlsx files."""
    if input_arg and os.path.isfile(input_arg):
        return [input_arg]
    candidates = []
    if input_arg and os.path.isdir(input_arg):
        candidates.append(input_arg)
    here = os.path.dirname(os.path.abspath(__file__))
    candidates += [
        os.path.join(here, "input"),
        os.path.join(here, "agent_report_debits"),
        os.path.join(os.path.dirname(here), "agent_report_debits"),
        here,
    ]
    for c in candidates:
        if os.path.isdir(c):
            files = sorted(glob.glob(os.path.join(c, "*.xlsx")) +
                           glob.glob(os.path.join(c, "*.xlsm")))
            files = [f for f in files if not os.path.basename(f).startswith("~$")]
            if files:
                return files
    return []


def load_data(xlsx_files):
    frames = []
    for fp in xlsx_files:
        xl = pd.ExcelFile(fp)
        sheet = "data" if "data" in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(fp, sheet_name=sheet)
        df["__source_file__"] = os.path.basename(fp)
        frames.append(df)
    full = pd.concat(frames, ignore_index=True)
    return full


def fetch_from_api(url, api_key, sede=None, since=None, until=None, page_size=5000, timeout=60):
    """Fetch debits from the configured API.

    Sends x-api-key header, supports optional ?sede=, ?from=, ?to=, ?page=&size= params.
    Tries to follow standard pagination: keeps requesting `page` until empty.
    Accepts response shapes: list, {"data":[...]}, {"records":[...]}, {"items":[...]}.
    """
    headers = {"Accept": "application/json"}
    if api_key:
        headers["x-api-key"] = api_key
        headers["Authorization"] = f"Bearer {api_key}"

    def _req(qs):
        full = url + ("?" + urllib.parse.urlencode({k: v for k, v in qs.items() if v}) if qs else "")
        req = urllib.request.Request(full, headers=headers, method="GET")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read()
            try:
                return json.loads(raw)
            except json.JSONDecodeError as e:
                raise RuntimeError(f"Respuesta no es JSON ({len(raw)} bytes): {raw[:200]!r}") from e

    def _records(payload):
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            for k in ("data", "records", "items", "results", "rows"):
                if isinstance(payload.get(k), list):
                    return payload[k]
        return []

    base_qs = {}
    if sede:
        base_qs["sede"] = sede
    if since:
        base_qs["from"] = since
    if until:
        base_qs["to"] = until

    # Try paginated; if API ignores page, single response is returned and loop breaks.
    rows = []
    page = 1
    while True:
        qs = dict(base_qs)
        qs["page"] = page
        qs["size"] = page_size
        try:
            payload = _req(qs)
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")[:500]
            raise RuntimeError(f"HTTP {e.code} al consultar la API: {body}") from e
        recs = _records(payload)
        if not recs:
            break
        rows.extend(recs)
        if len(recs) < page_size:
            break
        page += 1
        if page > 50:  # safety net
            break

    if not rows:
        raise RuntimeError(
            f"La API no devolvio registros. URL probada: {url}. "
            f"Verifica el endpoint y la API key."
        )
    df = pd.json_normalize(rows)
    df["__source_file__"] = "api:" + url
    return df


COL_MAP = {
    "intento": ["Intento", "Fecha", "Date", "intento", "fecha", "date", "attempt_date"],
    "valor": ["Valor", "Monto", "Amount", "valor", "amount", "monto"],
    "status": ["Status", "Estado", "status", "estado"],
    "motivo": ["Motivo del rechazo", "Motivo", "Reason", "motivo_rechazo", "reason", "motivo"],
    "tipo": ["Tipo de rechazo", "Tipo", "tipo_rechazo", "tipo"],
    "sede": ["Sede/club", "Sede", "Club", "sede", "club", "sede_club", "branch"],
    "marca": ["Marca de la tarjeta", "Franquicia", "Marca", "marca_tarjeta", "brand", "franquicia"],
    "operador": ["Operador de tarjeta", "Operador", "operador_tarjeta"],
    "cliente": ["Cliente", "cliente", "customer"],
    "venta": ["ID de la venta", "Venta", "sale_id", "id_venta"],
}


def resolve_columns(df):
    out = {}
    for key, options in COL_MAP.items():
        for c in options:
            if c in df.columns:
                out[key] = c
                break
    return out


def coerce(df, cols):
    df = df.copy()
    if "intento" in cols:
        df[cols["intento"]] = pd.to_datetime(df[cols["intento"]], errors="coerce")
    for k in ("status", "motivo", "tipo", "sede", "marca"):
        if k in cols:
            df[cols[k]] = df[cols[k]].apply(normalize_text)
    if "marca" in cols:
        df[cols["marca"]] = df[cols["marca"]].astype(str).str.upper().replace({"NAN": np.nan})
    return df


def compute_summary(df, cols):
    total = len(df)
    s = df[cols["status"]].astype(str).str.lower()
    approved_mask = s.str.startswith("aprob") | s.str.startswith("approv")
    denied_mask = s.str.startswith("nega") | s.str.startswith("deni") | s.str.startswith("rech")
    approved = int(approved_mask.sum())
    denied = int(denied_mask.sum())
    other = total - approved - denied
    rate = approved / total if total else 0
    fail_rate = denied / total if total else 0
    res = {
        "total": total,
        "approved": approved,
        "denied": denied,
        "other": other,
        "success_rate": rate,
        "fail_rate": fail_rate,
    }
    if "valor" in cols:
        res["amount_total"] = float(pd.to_numeric(df[cols["valor"]], errors="coerce").fillna(0).sum())
        res["amount_approved"] = float(pd.to_numeric(df.loc[approved_mask, cols["valor"]], errors="coerce").fillna(0).sum())
        res["amount_denied"] = float(pd.to_numeric(df.loc[denied_mask, cols["valor"]], errors="coerce").fillna(0).sum())
    if "intento" in cols:
        dts = df[cols["intento"]].dropna()
        if len(dts):
            res["date_min"] = dts.min()
            res["date_max"] = dts.max()
    return res, approved_mask, denied_mask


def by_sede(df, cols, approved_mask, denied_mask):
    if "sede" not in cols:
        return pd.DataFrame()
    g = df.groupby(cols["sede"], dropna=False)
    out = pd.DataFrame({
        "Sede": g.size().index,
        "Total": g.size().values,
    })
    appr = df[approved_mask].groupby(cols["sede"]).size().reindex(out["Sede"]).fillna(0).astype(int).values
    den = df[denied_mask].groupby(cols["sede"]).size().reindex(out["Sede"]).fillna(0).astype(int).values
    out["Aprobado"] = appr
    out["Negado"] = den
    out["TasaExito"] = np.where(out["Total"] > 0, out["Aprobado"] / out["Total"], 0)
    out["TasaFallo"] = 1 - out["TasaExito"]
    out = out.sort_values(["Total"], ascending=False).reset_index(drop=True)
    return out


def by_motivo(df, cols, denied_mask):
    if "motivo" not in cols:
        return pd.DataFrame()
    sub = df.loc[denied_mask, cols["motivo"]].fillna("Sin motivo registrado")
    counts = sub.value_counts()
    total = counts.sum()
    out = pd.DataFrame({
        "Motivo": counts.index,
        "Veces": counts.values,
        "Pct": counts.values / total if total else 0,
    })
    out["Ranking"] = np.arange(1, len(out) + 1)
    out["MotivoES"] = out["Motivo"].apply(translate_motivo)
    out["Accion"] = out["Motivo"].apply(lambda m: (lookup_motivo(m) or {}).get("accion", "Revisar caso a caso"))
    out["Responsable"] = out["Motivo"].apply(lambda m: (lookup_motivo(m) or {}).get("responsable", "Operacion"))
    return out


def by_tipo(df, cols, denied_mask):
    if "tipo" not in cols:
        return pd.DataFrame()
    sub = df.loc[denied_mask, cols["tipo"]].fillna("Sin clasificacion")
    c = sub.value_counts()
    return pd.DataFrame({"Tipo": c.index, "Veces": c.values})


def by_marca(df, cols, approved_mask, denied_mask):
    if "marca" not in cols:
        return pd.DataFrame()
    g = df[cols["marca"]].fillna("DESCONOCIDA")
    appr = g[approved_mask].value_counts()
    den = g[denied_mask].value_counts()
    idx = sorted(set(appr.index) | set(den.index))
    out = pd.DataFrame({
        "Franquicia": idx,
        "Aprobado": [int(appr.get(i, 0)) for i in idx],
        "Negado": [int(den.get(i, 0)) for i in idx],
    })
    out["Total"] = out["Aprobado"] + out["Negado"]
    out["TasaExito"] = np.where(out["Total"] > 0, out["Aprobado"] / out["Total"], 0)
    out = out.sort_values("Total", ascending=False).reset_index(drop=True)
    return out


def by_day(df, cols, approved_mask, denied_mask):
    if "intento" not in cols:
        return pd.DataFrame()
    d = df[cols["intento"]].dt.normalize()
    out = pd.DataFrame({"Dia": d, "Aprobado": approved_mask.astype(int).values, "Negado": denied_mask.astype(int).values})
    g = out.groupby("Dia").sum().reset_index()
    g["Total"] = g["Aprobado"] + g["Negado"]
    g["TasaExito"] = np.where(g["Total"] > 0, g["Aprobado"] / g["Total"], 0)
    return g


# ------------------ Charts ------------------

def setup_mpl():
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "font.size": 9,
        "axes.titlesize": 11,
        "axes.titleweight": "bold",
        "axes.labelsize": 9,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.edgecolor": "#7A8597",
        "axes.titlecolor": "#0F2A4A",
        "axes.labelcolor": "#0F2A4A",
        "xtick.color": "#0F2A4A",
        "ytick.color": "#0F2A4A",
        "figure.facecolor": "white",
    })


def chart_donut_status(s, path):
    fig, ax = plt.subplots(figsize=(4.0, 2.6), dpi=180)
    vals = [s["approved"], s["denied"]]
    labels = ["Aprobado", "Negado"]
    colors_ = ["#1F8A4C", "#C0392B"]
    wedges, _ = ax.pie(vals, colors=colors_, startangle=90,
                       wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    rate = s["success_rate"] * 100
    ax.text(0, 0.08, f"{rate:.1f}%", ha="center", va="center",
            fontsize=22, fontweight="bold", color="#0F2A4A")
    ax.text(0, -0.18, "Exito de descuentos", ha="center", va="center",
            fontsize=8, color="#7A8597")
    ax.legend(wedges, [f"{l}: {fmt_int(v)}" for l, v in zip(labels, vals)],
              loc="lower center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False, fontsize=8)
    ax.set_title("Resultado global de intentos")
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def chart_top_sedes_failures(sedes, path, n=15):
    if sedes.empty:
        return None
    sub = sedes.sort_values("Negado", ascending=False).head(n).iloc[::-1]
    fig, ax = plt.subplots(figsize=(7.2, 0.32 * len(sub) + 1.4), dpi=160)
    ax.barh(sub["Sede"], sub["Negado"], color="#C0392B", alpha=0.9)
    ax.barh(sub["Sede"], sub["Aprobado"], left=sub["Negado"], color="#1F8A4C", alpha=0.85)
    for i, (neg, ap) in enumerate(zip(sub["Negado"], sub["Aprobado"])):
        ax.text(neg / 2, i, fmt_int(neg), va="center", ha="center", color="white", fontsize=7)
    ax.set_xlabel("Intentos")
    ax.set_title(f"Top {n} sedes con mas fallos (Negado vs Aprobado)")
    ax.xaxis.set_major_formatter(mtick.FuncFormatter(lambda x, _: fmt_int(x)))
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def chart_top_sedes_failrate(sedes, path, n=15, min_total=200):
    if sedes.empty:
        return None
    sub = sedes[sedes["Total"] >= min_total].sort_values("TasaFallo", ascending=False).head(n).iloc[::-1]
    if sub.empty:
        return None
    fig, ax = plt.subplots(figsize=(7.2, 0.32 * len(sub) + 1.4), dpi=160)
    bars = ax.barh(sub["Sede"], sub["TasaFallo"] * 100, color="#E67E22")
    for b, v, t in zip(bars, sub["TasaFallo"] * 100, sub["Total"]):
        ax.text(v + 0.5, b.get_y() + b.get_height() / 2,
                f"{v:.1f}% ({fmt_int(t)})", va="center", fontsize=7, color="#0F2A4A")
    ax.set_xlim(0, max(100, sub["TasaFallo"].max() * 100 * 1.15))
    ax.set_xlabel("Tasa de fallo (%)")
    ax.set_title(f"Sedes con peor tasa de fallo (>= {min_total} intentos)")
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def chart_motivos_pie(motivos, path, n=8):
    if motivos.empty:
        return None
    top = motivos.head(n).copy()
    others = motivos.iloc[n:]["Veces"].sum()
    if others > 0:
        top = pd.concat([top, pd.DataFrame([{"Motivo": "Otros", "Veces": others, "Pct": others / motivos["Veces"].sum()}])], ignore_index=True)
    fig, ax = plt.subplots(figsize=(7.4, 4.4), dpi=160)
    palette = ["#0F2A4A", "#C0392B", "#E67E22", "#D4AF37", "#1F8A4C",
               "#7A8597", "#5D9CEC", "#9B59B6", "#34495E", "#16A085"]
    wedges, _ = ax.pie(top["Veces"], startangle=90, colors=palette[:len(top)],
                       wedgeprops=dict(edgecolor="white", linewidth=1.5))
    labels = []
    for m, v, p in zip(top["Motivo"], top["Veces"], top["Veces"] / top["Veces"].sum()):
        m_short = (m[:55] + "...") if isinstance(m, str) and len(m) > 58 else m
        labels.append(f"{m_short} - {p*100:.1f}% ({fmt_int(v)})")
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.02, 0.5),
              frameon=False, fontsize=8)
    ax.set_title("Distribucion de motivos de rechazo (Top)")
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def chart_trend(daily, path):
    if daily.empty:
        return None
    fig, ax1 = plt.subplots(figsize=(8.4, 2.6), dpi=180)
    ax1.bar(daily["Dia"], daily["Aprobado"], color="#1F8A4C", label="Aprobado", alpha=0.85)
    ax1.bar(daily["Dia"], daily["Negado"], bottom=daily["Aprobado"], color="#C0392B", label="Negado", alpha=0.85)
    ax1.set_ylabel("Intentos por dia")
    ax1.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x, _: fmt_int(x)))
    ax2 = ax1.twinx()
    ax2.plot(daily["Dia"], daily["TasaExito"] * 100, color="#0F2A4A", linewidth=2, marker="o", markersize=3, label="Tasa exito")
    ax2.set_ylabel("Tasa de exito (%)")
    ax2.set_ylim(0, 100)
    ax2.spines["top"].set_visible(False)
    ax2.grid(False)
    fig.autofmt_xdate()
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, loc="upper center", ncol=3, frameon=False, fontsize=8, bbox_to_anchor=(0.5, -0.18))
    ax1.set_title("Tendencia diaria de descuentos")
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


# ------------------ PDF ------------------

def make_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleBig", parent=styles["Title"], fontName="Helvetica-Bold",
                              fontSize=28, leading=32, textColor=PRIMARY, alignment=0))
    styles.add(ParagraphStyle(name="Subtitle", parent=styles["Normal"], fontName="Helvetica",
                              fontSize=13, leading=16, textColor=GREY))
    styles.add(ParagraphStyle(name="H1", parent=styles["Heading1"], fontName="Helvetica-Bold",
                              fontSize=16, leading=20, textColor=PRIMARY, spaceBefore=8, spaceAfter=8))
    styles.add(ParagraphStyle(name="H2", parent=styles["Heading2"], fontName="Helvetica-Bold",
                              fontSize=12, leading=16, textColor=PRIMARY, spaceBefore=6, spaceAfter=4))
    styles.add(ParagraphStyle(name="Body", parent=styles["BodyText"], fontName="Helvetica",
                              fontSize=10, leading=14, textColor=colors.HexColor("#1B2535")))
    styles.add(ParagraphStyle(name="Small", parent=styles["BodyText"], fontName="Helvetica",
                              fontSize=8.5, leading=11, textColor=GREY))
    styles.add(ParagraphStyle(name="KPIValue", parent=styles["Title"], fontName="Helvetica-Bold",
                              fontSize=15, leading=17, textColor=PRIMARY, alignment=1))
    styles.add(ParagraphStyle(name="KPILabel", parent=styles["BodyText"], fontName="Helvetica",
                              fontSize=7, leading=9, textColor=GREY, alignment=1))
    styles.add(ParagraphStyle(name="MyBullet", parent=styles["BodyText"], fontName="Helvetica",
                              fontSize=10, leading=14, leftIndent=12, bulletIndent=0, textColor=colors.HexColor("#1B2535")))
    return styles


def cover_page(canv, doc, summary, period_str, generated_str):
    canv.saveState()
    w, h = A4
    canv.setFillColor(PRIMARY)
    canv.rect(0, 0, w, h, fill=1, stroke=0)
    canv.setFillColor(ACCENT)
    canv.rect(0, h - 0.7 * cm, w, 0.7 * cm, fill=1, stroke=0)
    canv.rect(0, 0, w, 0.4 * cm, fill=1, stroke=0)
    canv.setFillColor(colors.white)
    canv.setFont("Helvetica-Bold", 11)
    canv.drawString(2 * cm, h - 1.6 * cm, "EVO  ·  REPORTE GERENCIAL")
    canv.setFont("Helvetica-Bold", 36)
    canv.drawString(2 * cm, h - 7 * cm, "Reporte de Debitos")
    canv.setFont("Helvetica", 18)
    canv.setFillColor(ACCENT)
    canv.drawString(2 * cm, h - 8.2 * cm, "Analisis de descuentos automaticos")
    canv.setFillColor(colors.white)
    canv.setFont("Helvetica", 12)
    canv.drawString(2 * cm, h - 10.2 * cm, f"Periodo analizado: {period_str}")
    canv.drawString(2 * cm, h - 10.9 * cm, f"Total intentos: {fmt_int(summary['total'])}")
    canv.drawString(2 * cm, h - 11.6 * cm, f"Tasa de exito global: {fmt_pct(summary['success_rate'])}")
    canv.setFont("Helvetica", 9)
    canv.setFillColor(colors.HexColor("#9FB3C8"))
    canv.drawString(2 * cm, 2 * cm, f"Generado: {generated_str}")
    canv.drawRightString(w - 2 * cm, 2 * cm, "Confidencial - Uso interno")
    canv.restoreState()


def content_page(canv, doc):
    canv.saveState()
    w, h = A4
    canv.setFillColor(PRIMARY)
    canv.rect(0, h - 1.0 * cm, w, 1.0 * cm, fill=1, stroke=0)
    canv.setFillColor(colors.white)
    canv.setFont("Helvetica-Bold", 9)
    canv.drawString(1.5 * cm, h - 0.6 * cm, "EVO  ·  Reporte de Debitos")
    canv.drawRightString(w - 1.5 * cm, h - 0.6 * cm, datetime.now().strftime("%Y-%m-%d"))
    canv.setFillColor(GREY)
    canv.setFont("Helvetica", 8)
    canv.drawCentredString(w / 2, 1.2 * cm, f"Pagina {doc.page}  ·  Confidencial")
    canv.setFillColor(ACCENT)
    canv.rect(0, h - 1.05 * cm, w, 0.05 * cm, fill=1, stroke=0)
    canv.restoreState()


def kpi_box(label, value, sub=None, color=PRIMARY, styles=None):
    val = Paragraph(f'<font color="{color.hexval()}">{value}</font>', styles["KPIValue"])
    lbl = Paragraph(label.upper(), styles["KPILabel"])
    cells = [[val], [lbl]]
    if sub:
        cells.append([Paragraph(f'<font color="{GREY.hexval()}">{sub}</font>', styles["KPILabel"])])
    t = Table(cells, colWidths=[4.4 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#DCE2EC")),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEABOVE", (0, 0), (-1, 0), 2, color),
    ]))
    return t


def df_to_table(df, col_widths=None, header_bg=PRIMARY, zebra=True, max_rows=None,
                fmt_map=None, align_map=None, font_size=8.5):
    if max_rows:
        df = df.head(max_rows)
    headers = list(df.columns)
    rows = [headers]
    for _, r in df.iterrows():
        row = []
        for c in headers:
            v = r[c]
            if fmt_map and c in fmt_map:
                v = fmt_map[c](v)
            elif isinstance(v, float) and not pd.isna(v):
                v = f"{v:,.2f}"
            row.append(str(v) if not pd.isna(v) else "")
        rows.append(row)
    t = Table(rows, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), font_size + 0.5),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), font_size),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, 0), 1, ACCENT),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#DCE2EC")),
    ]
    if zebra:
        for i in range(1, len(rows)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    if align_map:
        for col, alignment in align_map.items():
            if col in headers:
                ci = headers.index(col)
                style.append(("ALIGN", (ci, 1), (ci, -1), alignment))
    t.setStyle(TableStyle(style))
    return t


def build_pdf(out_path, df, cols, summary, sedes, motivos, tipos, marcas, daily, chart_paths,
              sede_label="Todas las sedes"):
    """Single-page A4 management report."""
    doc = BaseDocTemplate(out_path, pagesize=A4,
                          leftMargin=1.0 * cm, rightMargin=1.0 * cm,
                          topMargin=1.0 * cm, bottomMargin=0.8 * cm,
                          title="Reporte de Debitos", author="EVO Analytics")
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  A4[0] - doc.leftMargin - doc.rightMargin,
                  A4[1] - doc.topMargin - doc.bottomMargin,
                  id="content", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    period_str = "—"
    if "date_min" in summary and "date_max" in summary:
        period_str = f"{summary['date_min'].strftime('%d %b %Y')} - {summary['date_max'].strftime('%d %b %Y')}"
    generated_str = datetime.now().strftime("%d %b %Y, %H:%M")

    def page_decor(canv, d):
        canv.saveState()
        w, h = A4
        canv.setFillColor(PRIMARY)
        canv.rect(0, h - 1.0 * cm, w, 1.0 * cm, fill=1, stroke=0)
        canv.setFillColor(ACCENT)
        canv.rect(0, h - 1.05 * cm, w, 0.05 * cm, fill=1, stroke=0)
        canv.setFillColor(colors.white)
        canv.setFont("Helvetica-Bold", 12)
        canv.drawString(1.0 * cm, h - 0.65 * cm,
                        f"EVO  ·  REPORTE GERENCIAL DE DEBITOS  ·  {sede_label.upper()}")
        canv.setFont("Helvetica", 8.5)
        canv.drawRightString(w - 1.0 * cm, h - 0.65 * cm, f"{period_str}  ·  {generated_str}")
        canv.setFillColor(GREY)
        canv.setFont("Helvetica", 7)
        canv.drawCentredString(w / 2, 0.4 * cm, "Confidencial - Uso interno")
        canv.restoreState()

    doc.addPageTemplates([PageTemplate(id="content", frames=[frame], onPage=page_decor)])
    styles = make_styles()
    story = []
    story.append(Spacer(1, 0.35 * cm))

    n_sedes = df[cols["sede"]].nunique() if "sede" in cols else 0

    # --- KPI ROW ---
    kpis = [
        kpi_box("Intentos", fmt_int(summary["total"]), sub=f"{n_sedes} sedes", styles=styles),
        kpi_box("Aprobados", fmt_int(summary["approved"]),
                sub=fmt_pct(summary["success_rate"]), color=GOOD, styles=styles),
        kpi_box("Negados", fmt_int(summary["denied"]),
                sub=fmt_pct(summary["fail_rate"]), color=BAD, styles=styles),
    ]
    if "amount_total" in summary:
        kpis.append(kpi_box("Recuperado", fmt_money(summary.get("amount_approved", 0)),
                            sub=f"Riesgo: {fmt_money(summary.get('amount_denied', 0))}",
                            color=ACCENT, styles=styles))
    kpi_w = (A4[0] - 2 * cm) / len(kpis)
    kpi_table = Table([kpis], colWidths=[kpi_w] * len(kpis))
    kpi_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 6))

    # --- LEFT: donut  | RIGHT: top motivos ---
    left_blocks = []
    if chart_paths.get("donut"):
        left_blocks.append(Image(chart_paths["donut"], width=7.6 * cm, height=4.4 * cm))

    right_blocks = [Paragraph("<b>Top 5 motivos de rechazo (en espanol)</b>", styles["Small"])]
    if not motivos.empty:
        m2 = motivos.head(5).copy()
        m2["Pct"] = (m2["Pct"] * 100).round(1).astype(str) + "%"
        m2["MotivoES"] = m2["MotivoES"].apply(lambda x: (x[:55] + "...") if isinstance(x, str) and len(x) > 58 else x)
        m2["Veces"] = m2["Veces"].apply(fmt_int)
        m2 = m2[["Ranking", "MotivoES", "Veces", "Pct"]]
        m2.columns = ["#", "Motivo", "Veces", "%"]
        right_blocks.append(df_to_table(m2,
                                        col_widths=[0.6*cm, 7.5*cm, 1.6*cm, 1.3*cm],
                                        font_size=7.0,
                                        align_map={"Veces": "RIGHT", "%": "RIGHT", "#": "CENTER"}))
    if not tipos.empty:
        ttot = tipos["Veces"].sum()
        rev = int(tipos[tipos["Tipo"].astype(str).str.lower() == "reversible"]["Veces"].sum())
        irr = int(tipos[tipos["Tipo"].astype(str).str.lower() == "irreversible"]["Veces"].sum())
        right_blocks.append(Spacer(1, 2))
        right_blocks.append(Paragraph(
            f"<font size=7 color='{GREY.hexval()}'>Reversibles: <b>{fmt_int(rev)}</b> ({rev/ttot*100:.1f}%)  ·  "
            f"Irreversibles: <b>{fmt_int(irr)}</b> ({irr/ttot*100:.1f}%)</font>",
            styles["Small"]))

    row_table = Table([[left_blocks, right_blocks]],
                      colWidths=[(A4[0]-2*cm)*0.42, (A4[0]-2*cm)*0.58])
    row_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(row_table)
    story.append(Spacer(1, 3))

    # --- Tendencia ---
    if chart_paths.get("trend"):
        story.append(Image(chart_paths["trend"], width=A4[0]-2.2*cm, height=3.6 * cm))
        story.append(Spacer(1, 3))

    # --- Sedes peor / mas fallos (solo cuando hay >1 sede) ---
    avg = summary["success_rate"]
    multi_sede = len(sedes) > 1
    if multi_sede:
        worst = sedes[sedes["Total"] >= 200].sort_values("TasaExito").head(5).copy()
        worst["TasaExito"] = (worst["TasaExito"] * 100).round(1).astype(str) + "%"
        worst["Total"] = worst["Total"].apply(fmt_int)
        worst = worst[["Sede", "Total", "TasaExito"]]
        worst.columns = ["Sede", "Total", "Tasa"]

        big = sedes.sort_values("Negado", ascending=False).head(5).copy()
        big["TasaExito"] = (big["TasaExito"] * 100).round(1).astype(str) + "%"
        for c in ("Negado", "Total"):
            big[c] = big[c].apply(fmt_int)
        big = big[["Sede", "Negado", "TasaExito"]]
        big.columns = ["Sede", "Negados", "Tasa"]

        sedes_left = [Paragraph("<b>Top 5 sedes con peor tasa de exito (vol >=200)</b>", styles["Small"]),
                      df_to_table(worst,
                                  col_widths=[5.6*cm, 1.5*cm, 1.4*cm],
                                  font_size=7.0,
                                  align_map={"Total": "RIGHT", "Tasa": "RIGHT"})]
        sedes_right = [Paragraph("<b>Top 5 sedes con mas fallos absolutos</b>", styles["Small"]),
                       df_to_table(big,
                                   col_widths=[5.6*cm, 1.5*cm, 1.4*cm],
                                   font_size=7.0,
                                   align_map={"Negados": "RIGHT", "Tasa": "RIGHT"})]
        s_table = Table([[sedes_left, sedes_right]],
                        colWidths=[(A4[0]-2*cm)*0.5, (A4[0]-2*cm)*0.5])
        s_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(s_table)
        story.append(Spacer(1, 3))
    elif not motivos.empty:
        # Reemplazo: tabla de motivos con accion sugerida
        story.append(Paragraph("<b>Acciones por motivo (top 5)</b>", styles["Small"]))
        ac = motivos.head(5).copy()
        ac["MotivoES"] = ac["MotivoES"].apply(lambda x: (x[:38] + "...") if isinstance(x, str) and len(x) > 40 else x)
        ac["Accion"] = ac["Accion"].apply(lambda x: (x[:90] + "...") if isinstance(x, str) and len(x) > 92 else x)
        ac = ac[["MotivoES", "Responsable", "Accion"]]
        ac.columns = ["Motivo (ES)", "Resp.", "Accion sugerida"]
        story.append(df_to_table(ac,
                                 col_widths=[4.4*cm, 1.8*cm, 12.6*cm],
                                 font_size=7.0))
        story.append(Spacer(1, 3))

    # --- Conclusiones + Recomendaciones (compact, side by side) ---
    insights = build_insights(summary, sedes, motivos, tipos)
    recs = build_recommendations(summary, sedes, motivos, tipos)

    concl_style = ParagraphStyle("ConclTxt", parent=styles["Body"], fontSize=7.5, leading=9.5,
                                 textColor=colors.HexColor("#1B2535"))
    h_style = ParagraphStyle("h", parent=styles["Small"], fontSize=8.5, textColor=PRIMARY, leading=10)
    concl_left = [Paragraph("<b>Conclusiones clave</b>", h_style)]
    for it in insights[:3]:
        concl_left.append(Paragraph(f"• {it}", concl_style))

    concl_right = [Paragraph("<b>Recomendaciones</b>", h_style)]
    for r in recs[:3]:
        concl_right.append(Paragraph(f"• {r}", concl_style))

    c_table = Table([[concl_left, concl_right]],
                    colWidths=[(A4[0]-2*cm)*0.5, (A4[0]-2*cm)*0.5])
    c_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#DCE2EC")),
        ("LINEABOVE", (0, 0), (-1, 0), 2, ACCENT),
    ]))
    story.append(c_table)

    doc.build(story)


def build_insights(summary, sedes, motivos, tipos):
    out = []
    rate = summary["success_rate"] * 100
    out.append(f"Se procesaron <b>{fmt_int(summary['total'])}</b> intentos de debito; "
               f"<b>{fmt_int(summary['approved'])}</b> exitosos ({rate:.1f}%) y "
               f"<b>{fmt_int(summary['denied'])}</b> negados ({summary['fail_rate']*100:.1f}%).")
    if "amount_total" in summary:
        recovered = summary.get("amount_approved", 0)
        lost = summary.get("amount_denied", 0)
        out.append(f"Monto total cursado: <b>{fmt_money(summary['amount_total'])}</b>. "
                   f"Recuperado: <b>{fmt_money(recovered)}</b>. "
                   f"En riesgo (negado): <b>{fmt_money(lost)}</b>.")
    if not motivos.empty:
        top3 = motivos.head(3)
        parts = [f"<b>{m[:55]}{'...' if len(m)>55 else ''}</b> ({p*100:.1f}%)"
                 for m, p in zip(top3["MotivoES"], top3["Pct"])]
        out.append("Motivos dominantes de rechazo: " + "; ".join(parts) + ".")
    if not tipos.empty:
        rev = tipos[tipos["Tipo"].astype(str).str.lower() == "reversible"]["Veces"].sum()
        irr = tipos[tipos["Tipo"].astype(str).str.lower() == "irreversible"]["Veces"].sum()
        if rev + irr > 0:
            pct_rev = rev / (rev + irr) * 100
            out.append(f"<b>{pct_rev:.1f}%</b> de los rechazos son reversibles "
                       f"(recuperables con reintento o gestion); el resto requiere actualizacion del medio de pago.")
    if not sedes.empty:
        avg = summary["success_rate"]
        bad = sedes[(sedes["Total"] >= 200) & (sedes["TasaExito"] < avg * 0.85)].sort_values("TasaExito").head(3)
        if not bad.empty:
            parts = [f"<b>{s}</b> ({r*100:.1f}%)" for s, r in zip(bad["Sede"], bad["TasaExito"])]
            out.append("Sedes criticas con tasa de exito muy por debajo del promedio: " + "; ".join(parts) + ".")
        big_loss = sedes.sort_values("Negado", ascending=False).head(3)
        parts = [f"<b>{s}</b> ({fmt_int(n)})" for s, n in zip(big_loss["Sede"], big_loss["Negado"])]
        out.append("Sedes con mayor volumen absoluto de fallos: " + "; ".join(parts) + ".")
    return out


def build_recommendations(summary, sedes, motivos, tipos):
    recs = []
    if not motivos.empty:
        top1 = motivos.iloc[0]
        recs.append(f"Atacar el motivo #1 (<b>{top1['MotivoES']}</b>, "
                    f"{top1['Pct']*100:.1f}% del total negado): {top1['Accion']}")
    if not tipos.empty:
        rev = tipos[tipos["Tipo"].astype(str).str.lower() == "reversible"]["Veces"].sum()
        if rev > 0:
            recs.append(f"Activar pipeline de reintentos automaticos sobre los <b>{fmt_int(rev)}</b> rechazos reversibles "
                        "(probar reintento a 24h y a 72h con monitoreo de exito incremental).")
    if not sedes.empty:
        avg = summary["success_rate"]
        bad = sedes[(sedes["Total"] >= 200) & (sedes["TasaExito"] < avg * 0.85)]
        if not bad.empty:
            recs.append(f"Auditoria operativa en {len(bad)} sede(s) con tasa de exito inferior al 85% del promedio: "
                        "validar calidad del registro inicial de medio de pago y procesos de actualizacion en recepcion.")
    recs.append("Establecer un tablero diario con tasa de exito por sede; alertar cuando una sede caiga 5 pp por debajo del promedio dos dias seguidos.")
    recs.append("Definir SLA mensual de tasa de exito por sede y vincular a evaluacion de desempeno del responsable comercial.")
    recs.append("Cruzar motivos de \"profile not found\" / \"tarjeta no registrada\" con la base de afiliacion: hay datos faltantes en el alta.")
    return recs


def write_motivos_md(out_dir, motivos):
    """Companion file: catalogo de motivos con explicacion y accion."""
    if motivos.empty:
        return None
    path = os.path.join(out_dir, "Motivos_Acciones.md")
    lines = ["# Motivos de rechazo - Catalogo y acciones",
             "", "Texto traducido al espanol con explicacion y accion sugerida.",
             "", "| # | Motivo (ES) | Veces | % | Responsable | Accion |",
             "|---|---|---:|---:|---|---|"]
    for _, r in motivos.head(20).iterrows():
        accion = (r["Accion"] or "").replace("|", "/")
        es = (r["MotivoES"] or "").replace("|", "/")
        lines.append(f"| {r['Ranking']} | {es} | {fmt_int(r['Veces'])} | "
                     f"{r['Pct']*100:.1f}% | {r['Responsable']} | {accion} |")
    lines.append("")
    lines.append("## Texto original del gateway")
    lines.append("")
    for _, r in motivos.head(20).iterrows():
        lines.append(f"- **{r['MotivoES']}** -- original: \"{r['Motivo']}\"")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path


def normalize_match(s):
    if not isinstance(s, str):
        return ""
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower().strip()


def filter_by_sede(df, cols, sede):
    if not sede or sede.upper() in ("ALL", "TODAS", "TODOS", "*"):
        return df, "Todas las sedes"
    if "sede" not in cols:
        raise SystemExit("La data no tiene columna de sede; no se puede filtrar.")
    target = normalize_match(sede)
    matched_col = df[cols["sede"]].astype(str).apply(normalize_match)
    mask = matched_col == target
    if mask.sum() == 0:
        # try contains
        mask = matched_col.str.contains(target, na=False)
    if mask.sum() == 0:
        opts = sorted(df[cols["sede"]].dropna().astype(str).unique().tolist())
        sample = ", ".join(opts[:10])
        raise SystemExit(
            f"No se encontro la sede '{sede}'. Hay {len(opts)} sedes. "
            f"Ejemplos: {sample}. Usa --list-sedes para ver todas."
        )
    df2 = df.loc[mask].copy()
    actual = df2[cols["sede"]].dropna().astype(str).iloc[0]
    return df2, actual


def main():
    parser = argparse.ArgumentParser(
        description="Generador de Reporte de Debitos EVO. Por defecto lee desde la API.")
    src = parser.add_mutually_exclusive_group()
    src.add_argument("--api", action="store_true", default=True,
                     help="(Default) Cargar desde la API EVO_DEBITS_URL.")
    src.add_argument("--file", help="Cargar desde archivo .xlsx en lugar de la API.", default=None)
    src.add_argument("--input-dir", help="Carpeta con .xlsx (modo legacy).", default=None)

    parser.add_argument("--api-url", help="Endpoint de la API (override).", default=DEFAULT_API_URL)
    parser.add_argument("--api-key", help="API key (override).", default=DEFAULT_API_KEY)
    parser.add_argument("--sede", help="Filtrar por una sede especifica. Use 'ALL' o omita para todas.",
                        default=None)
    parser.add_argument("--list-sedes", action="store_true",
                        help="Imprime la lista de sedes disponibles y termina.")
    parser.add_argument("--since", help="Fecha desde (YYYY-MM-DD).", default=None)
    parser.add_argument("--until", help="Fecha hasta (YYYY-MM-DD).", default=None)
    parser.add_argument("--output", help="Carpeta de salida.", default=None)
    args = parser.parse_args()

    # --- Source selection ---
    use_api = not (args.file or args.input_dir)
    if use_api:
        print(f"[1/5] Consultando API: {args.api_url}")
        try:
            df = fetch_from_api(args.api_url, args.api_key,
                                sede=args.sede if args.sede and args.sede.upper() != "ALL" else None,
                                since=args.since, until=args.until)
            print(f"      OK - {len(df):,} registros recibidos")
        except Exception as e:
            print(f"      ERROR API: {e}")
            print("      Sugerencia: usa --file <ruta.xlsx> mientras se publica el endpoint.")
            sys.exit(2)
    else:
        files = find_input(args.file or args.input_dir)
        if not files:
            print("ERROR: no se encontraron archivos .xlsx. Coloca el archivo en ./input/ "
                  "o pasa --file <ruta> o --api.")
            sys.exit(1)
        print(f"[1/5] Cargando {len(files)} archivo(s)...")
        for f in files:
            print(f"      - {os.path.basename(f)}")
        df = load_data(files)

    cols = resolve_columns(df)

    if args.list_sedes:
        if "sede" not in cols:
            print("La data no tiene columna de sede.")
            sys.exit(1)
        sedes = sorted(df[cols["sede"]].dropna().astype(str).unique().tolist())
        print(f"Sedes disponibles ({len(sedes)}):")
        for s in sedes:
            print(f"  - {s}")
        sys.exit(0)

    df, sede_label = filter_by_sede(df, cols, args.sede)
    if args.sede:
        print(f"      Filtrado por sede: {sede_label} ({len(df):,} registros)")
    missing = [k for k in ("status", "sede", "motivo") if k not in cols]
    if missing:
        print(f"ADVERTENCIA: faltan columnas esperadas: {missing}. El reporte usara lo disponible.")
    df = coerce(df, cols)

    print("[2/5] Calculando metricas...")
    summary, ap_mask, dn_mask = compute_summary(df, cols)
    sedes = by_sede(df, cols, ap_mask, dn_mask)
    motivos = by_motivo(df, cols, dn_mask)
    tipos = by_tipo(df, cols, dn_mask)
    marcas = by_marca(df, cols, ap_mask, dn_mask)
    daily = by_day(df, cols, ap_mask, dn_mask)

    here = os.path.dirname(os.path.abspath(__file__))
    out_dir = args.output or os.path.join(here, "output")
    os.makedirs(out_dir, exist_ok=True)
    chart_dir = os.path.join(out_dir, ".charts")
    os.makedirs(chart_dir, exist_ok=True)

    print("[3/5] Generando graficos...")
    setup_mpl()
    chart_paths = {
        "donut": os.path.join(chart_dir, "donut_status.png"),
        "motivos": os.path.join(chart_dir, "motivos_pie.png"),
        "sedes_fail": os.path.join(chart_dir, "sedes_fail.png"),
        "sedes_rate": os.path.join(chart_dir, "sedes_rate.png"),
        "trend": os.path.join(chart_dir, "trend.png"),
    }
    chart_donut_status(summary, chart_paths["donut"])
    if not motivos.empty:
        chart_motivos_pie(motivos, chart_paths["motivos"])
    else:
        chart_paths["motivos"] = None
    if not sedes.empty:
        chart_top_sedes_failures(sedes, chart_paths["sedes_fail"])
        chart_top_sedes_failrate(sedes, chart_paths["sedes_rate"])
    else:
        chart_paths["sedes_fail"] = chart_paths["sedes_rate"] = None
    if not daily.empty and len(daily) > 1:
        chart_trend(daily, chart_paths["trend"])
    else:
        chart_paths["trend"] = None

    print("[4/5] Construyendo PDF...")
    period_tag = ""
    if "date_min" in summary:
        period_tag = "_" + summary["date_min"].strftime("%Y%m") + "-" + summary["date_max"].strftime("%Y%m")
    sede_tag = ""
    if args.sede and args.sede.upper() != "ALL":
        sede_tag = "_" + re.sub(r"[^A-Za-z0-9_-]+", "_",
                                normalize_match(sede_label).upper())[:40]
    pdf_name = f"Reporte_Debitos{sede_tag}{period_tag}.pdf"
    pdf_path = os.path.join(out_dir, pdf_name)
    build_pdf(pdf_path, df, cols, summary, sedes, motivos, tipos, marcas, daily,
              chart_paths, sede_label=sede_label)

    md_path = write_motivos_md(out_dir, motivos)

    print("[5/5] Listo.")
    print(f"      PDF: {pdf_path}")
    if md_path:
        print(f"      Motivos detallados: {md_path}")
    print(f"      Sedes: {len(sedes)}  ·  Motivos distintos: {len(motivos)}  ·  "
          f"Tasa exito: {summary['success_rate']*100:.2f}%")
    return pdf_path


if __name__ == "__main__":
    main()
