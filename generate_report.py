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
# Bumped when API helpers change. UI lee esto para mostrar version cargada.
API_FEATURE_VERSION = "v4.1-pdf-top15-franq"

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


def parse_currency_to_float(v):
    """Parse one cell to float. Returns NaN for unparseable.

    Handles:
      - numbers (int/float) -> as-is
      - dict with "value"/"amount"/"valor"/"monto"
      - strings: "$50,000.50" (US), "$50.000,50" (Colombiano),
        "50.000.000" (CO miles), "50,000" (US miles), "50,5" (CO decimal),
        "50.5" (decimal), with optional $/USD/COP/spaces.
    """
    import math
    if v is None:
        return float("nan")
    if isinstance(v, bool):
        return float("nan")
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return float("nan")
        return float(v)
    if isinstance(v, dict):
        for k in ("value", "amount", "valor", "monto", "total"):
            if k in v:
                return parse_currency_to_float(v[k])
        return float("nan")
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null", "n/a", "-"):
        return float("nan")
    import re
    # Strip symbols, currency codes and whitespace
    s = re.sub(r"[\$€₡£¥]", "", s)
    s = re.sub(r"(?<![A-Za-z])(?:USD|COP|EUR|MXN)(?![A-Za-z])", "", s, flags=re.I)
    s = re.sub(r"[\s  _]+", "", s)
    if not s or s == "-":
        return float("nan")
    sign = 1
    if s.startswith("-"):
        sign = -1
        s = s[1:]
    if s.startswith("(") and s.endswith(")"):
        sign = -1
        s = s[1:-1]
    has_comma = "," in s
    has_dot = "." in s
    if has_comma and has_dot:
        # The last occurring separator is the decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        parts = s.split(",")
        if len(parts[-1]) == 3 and all(p.isdigit() for p in parts):
            s = s.replace(",", "")  # miles
        else:
            s = s.replace(",", ".")  # decimal
    elif has_dot:
        parts = s.split(".")
        if len(parts) > 1 and len(parts[-1]) == 3 and all(p.isdigit() for p in parts):
            s = s.replace(".", "")  # miles
        # else trato el punto como decimal
    try:
        return float(s) * sign
    except (ValueError, TypeError):
        return float("nan")


def parse_currency_series(series):
    """Vectorized-ish parser for a pandas Series. Returns float Series."""
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    return series.map(parse_currency_to_float).astype(float)


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


def validate_date_range(from_str, to_str, max_days=5):
    """Validate a YYYY-MM-DD date range.

    Returns ((from_date, to_date), None) on success
            or (None, error_message) on failure.

    Rules:
      - format YYYY-MM-DD strict
      - to > from (strictly)
      - (to - from).days <= max_days  (5 by default)
    """
    from datetime import date
    try:
        f = date.fromisoformat(str(from_str))
        t = date.fromisoformat(str(to_str))
    except Exception:
        return None, "Formato invalido. Usa YYYY-MM-DD para ambas fechas."
    if t <= f:
        return None, "`to` (hasta) debe ser estrictamente mayor que `from` (desde)."
    days = (t - f).days
    if days > max_days:
        return None, f"Rango maximo permitido: {max_days} dias. Seleccionaste {days}."
    return (f, t), None


def split_date_ranges_by_month(from_date, to_date):
    """Split [from_date, to_date) into sub-ranges that never cross a month boundary.

    Devuelve lista de tuplas (sub_from, sub_to). `to_date` exclusivo.
    Si el rango no cruza meses, devuelve [(from_date, to_date)].
    """
    from datetime import date
    out = []
    current = from_date
    while current < to_date:
        if current.month == 12:
            next_month_start = date(current.year + 1, 1, 1)
        else:
            next_month_start = date(current.year, current.month + 1, 1)
        chunk_end = min(next_month_start, to_date)
        out.append((current, chunk_end))
        current = chunk_end
    return out


def _extract_records(payload):
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for k in ("data", "records", "items", "results", "rows"):
            if isinstance(payload.get(k), list):
                return payload[k]
    return []


# (label, kind, name, value_template). kind: "header" o "qs".
_AUTH_VARIANTS = [
    ("x-api-key header",     "header", "x-api-key",     "{t}"),
    ("?key= query param",    "qs",     "key",           "{t}"),
    ("?api_key= query param","qs",     "api_key",       "{t}"),
    ("Authorization Bearer", "header", "Authorization", "Bearer {t}"),
    ("Authorization raw",    "header", "Authorization", "{t}"),
    ("X-API-Token header",   "header", "X-API-Token",   "{t}"),
    ("api-token header",     "header", "api-token",     "{t}"),
]


def _fetch_single_range(url, token, from_str, to_str, page_size=10000, timeout=60):
    """One API call for [from, to). Tries header AND query-string auth variants."""
    import urllib.parse, urllib.request, urllib.error, json as _json
    errors = []
    for label, kind, name, tmpl in _AUTH_VARIANTS:
        val = tmpl.format(t=token)
        rows = []
        page = 1
        ok = False
        max_pages = 30
        while page <= max_pages:
            qs = {"from": from_str, "to": to_str, "page": page, "size": page_size}
            headers = {"Accept": "application/json", "User-Agent": "evo-debitos/3.1"}
            if kind == "header":
                headers[name] = val
            else:
                qs[name] = val
            full = url + "?" + urllib.parse.urlencode(qs)
            req = urllib.request.Request(full, headers=headers, method="GET")
            try:
                with urllib.request.urlopen(req, timeout=timeout) as r:
                    raw = r.read()
                    payload = _json.loads(raw)
                    ok = True
            except urllib.error.HTTPError as e:
                body = e.read().decode("utf-8", "replace")[:120].replace("\n", " ")
                errors.append(f"{label}: HTTP {e.code} {body}")
                ok = False
                break
            except Exception as e:
                errors.append(f"{label}: {type(e).__name__}: {e}")
                ok = False
                break
            recs = _extract_records(payload)
            if not recs:
                break
            rows.extend(recs)
            # Si el API devolvio MAS de lo que pedimos en page_size, esta
            # ignorando la paginacion y nos esta dando el set completo cada vez.
            # Romper para no acumular duplicados (sintoma: card Recuperado x30).
            if len(recs) > page_size:
                break
            if len(recs) < page_size:
                break
            page += 1
        if ok:
            if rows:
                try:
                    df = pd.DataFrame(rows)
                except Exception:
                    df = pd.json_normalize(rows)
                df["__source_file__"] = f"api:{url}?from={from_str}&to={to_str}"
                # Dedupe SOLO de filas 100% identicas (defensivo contra paginacion
                # rota). NUNCA por subset=[Id] porque `Id` no es row-unique en
                # esta API: es algo cercano a id_cliente, y dedup por ahi
                # colapsa retries legitimos.
                before = len(df)
                df = df.drop_duplicates().reset_index(drop=True)
                if len(df) != before:
                    df.attrs["dedup_dropped"] = before - len(df)
                return df
            return pd.DataFrame()
    raise RuntimeError(
        "No se pudo conectar a la API. Detalle de cada intento:\n  - "
        + "\n  - ".join(errors)
    )


def fetch_debitos_range(url, token, from_date, to_date, page_size=10000, timeout=60,
                         logger=None):
    """Fetch debitos for [from_date, to_date).

    La API es un snapshot MENSUAL: cuando el rango cruza meses, splittea
    automaticamente en sub-rangos por mes y concatena. Llamadas:
      - (2026-04-25, 2026-05-01)
      - (2026-05-01, 2026-05-06)
    Luego concat + dedup por (id, Intento, Status, Valor) si esas columnas existen.

    logger: callable opcional que recibe strings de progreso.
    """
    if isinstance(from_date, str):
        from datetime import date
        from_date = date.fromisoformat(from_date)
    if isinstance(to_date, str):
        from datetime import date
        to_date = date.fromisoformat(to_date)

    chunks = split_date_ranges_by_month(from_date, to_date)
    if logger:
        logger(f"Rango {from_date}..{to_date} dividido en {len(chunks)} llamada(s).")

    dfs = []
    for f, t in chunks:
        if logger:
            logger(f"  Llamando from={f} to={t}")
        df_chunk = _fetch_single_range(url, token, f.isoformat(), t.isoformat(),
                                        page_size=page_size, timeout=timeout)
        if df_chunk is not None and len(df_chunk):
            if logger:
                logger(f"    -> {len(df_chunk)} registros")
            dfs.append(df_chunk)
        else:
            if logger:
                logger("    -> 0 registros")

    if not dfs:
        return pd.DataFrame()

    result = pd.concat(dfs, ignore_index=True)
    dedup_cols = [c for c in ("id", "Intento", "Status", "Valor",
                              "ID de la venta", "Cliente") if c in result.columns]
    if len(dedup_cols) >= 2 and len(dfs) > 1:
        before = len(result)
        result = result.drop_duplicates(subset=dedup_cols)
        if logger and len(result) != before:
            logger(f"  Dedup: {before} -> {len(result)} filas")
    return result


def fetch_from_api_v2(url, token, month=None, page_size=10000, timeout=60):
    """Fetch debitos with `month=YYYY-MM` and multi-header authentication.

    Used by both apps (app.py admin and app_managers.py GM) for consistency.
    Tries 5 header variants in order until one returns 2xx. Hard-caps at
    300k records to protect Streamlit free-tier RAM.
    """
    MAX_RECORDS = 300_000
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
        headers["User-Agent"] = "evo-debitos/2.1"
        rows = []
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
                break
            if len(recs) < page_size:
                break
            page += 1
            if page > max_pages:
                break

        if ok and rows:
            try:
                df = pd.DataFrame(rows)
            except Exception:
                df = pd.json_normalize(rows)
            df["__source_file__"] = f"api:{url}?month={month}"
            return df

    raise RuntimeError(f"No se pudo conectar a la API. Detalle: {last_err}")


def available_months(n=12):
    out = []
    today = datetime.now().date()
    y, m = today.year, today.month
    for _ in range(n):
        out.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return out


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


ID_CLIENT_CANDIDATES = [
    "id_client", "id_cliente", "client_id", "Cliente_id",
    "id", "ID", "Id", "user_id", "userId", "userid",
    "id_usuario", "ID_usuario", "idUsuario",
]


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
    "prevision": ["Previsión de pagos", "Prevision de pagos", "prevision_pagos",
                  "PrevisionDePagos", "PrevisionPagos", "Prevision"],
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
        # Parse fechas. Reglas:
        # - API devuelve ISO 8601 ("2026-05-01T00:00:58Z"). dayfirst=True
        #   con ISO da bug: interpreta "2026-05-01" como YYYY-DD-MM (Jan 5).
        # - Excel colombiano trae "D/M/YYYY" ("1/5/2026" = 1 de mayo).
        #   dayfirst=True es necesario para no parsear "1/5/2026" como Ene 5.
        # Solucion: detectar formato y aplicar dayfirst SOLO si NO es ISO.
        s_raw = df[cols["intento"]]
        sample = s_raw.dropna().astype(str).head(20)
        # ISO si empieza por YYYY- (4 digitos + guion)
        is_iso = sample.str.match(r"^\d{4}-").any() if len(sample) else False
        def _parse_dates(series, iso):
            kw = {"errors": "coerce", "utc": True}
            if iso:
                for fmt in ("ISO8601", "mixed"):
                    try:
                        return pd.to_datetime(series, format=fmt, **kw)
                    except (TypeError, ValueError):
                        continue
                return pd.to_datetime(series, **kw)
            return pd.to_datetime(series, dayfirst=True, **kw)
        try:
            s = _parse_dates(s_raw, is_iso)
            if getattr(s.dt, "tz", None) is not None:
                # Convert UTC -> America/Bogota (UTC-5) so day grouping coincide
                # con el calendario local que usa el Excel. Sin esto, records
                # entre 00:00-04:59 UTC quedan asignados al dia siguiente en UTC
                # y aparecen como duplicados o desplazados vs el Excel.
                if is_iso:
                    s = s.dt.tz_convert("America/Bogota").dt.tz_localize(None)
                else:
                    s = s.dt.tz_convert("UTC").dt.tz_localize(None)
        except Exception:
            s = pd.to_datetime(s_raw, errors="coerce", dayfirst=not is_iso)
        df[cols["intento"]] = s
    for k in ("status", "motivo", "tipo", "sede", "marca"):
        if k in cols:
            df[cols[k]] = df[cols[k]].apply(normalize_text)
    if "marca" in cols:
        df[cols["marca"]] = df[cols["marca"]].astype(str).str.upper().replace({"NAN": np.nan})
    # Excluir intentos one-shot (sin "Prevision de pagos"). Estos son cobros
    # no recurrentes que el Excel del cliente filtra antes de reportar, asi
    # las metricas de la app coinciden con el ground truth operativo.
    if "prevision" in cols:
        prev_raw = df[cols["prevision"]]
        # Normalizar a string para detectar "", "nan", "None"
        prev_str = prev_raw.astype(str).str.strip()
        keep = prev_raw.notna() & ~prev_str.isin(("", "nan", "NaT", "None", "null", "<NA>"))
        df = df.loc[keep].copy()
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
    # Unique clients per metric (additive; no impact on existing calculations)
    ID_CLIENT_CANDIDATES = [
        "id_client", "id_cliente", "client_id", "Cliente_id",
        "id", "ID", "Id", "user_id", "userId", "userid",
        "id_usuario", "ID_usuario", "idUsuario",
    ]
    id_col_local = next((c for c in ID_CLIENT_CANDIDATES if c in df.columns), None)
    if id_col_local:
        try:
            # Mutually exclusive partition por usuario:
            #   clients_denied   = usuarios con AL MENOS 1 intento negado
            #   clients_approved = usuarios con 0 intentos negados (sin fallos)
            # Garantiza clients_approved + clients_denied == clients_total
            # (un usuario que mezcla aprobados y negados va a "denied").
            all_users = set(df[id_col_local].dropna().astype(str).unique())
            users_denied = set(
                df.loc[denied_mask, id_col_local].dropna().astype(str).unique()
            )
            users_approved_ever = set(
                df.loc[approved_mask, id_col_local].dropna().astype(str).unique()
            )
            res["clients_total"] = len(all_users)
            # clients_approved/denied = usuarios con AL MENOS 1 intento de ese
            # status. Coincide con la definicion del negocio y con groupby
            # del Excel. Los buckets pueden traslapar (un usuario con 4 aprob
            # + 1 negado cuenta en ambos), asi suma > clients_total es normal.
            res["clients_approved"] = len(users_approved_ever)
            res["clients_denied"] = len(users_denied)
            # Estos dos SI son mutuamente exclusivos (any-aprobado vs zero-aprobado).
            res["users_ever_approved"] = len(users_approved_ever)
            res["users_never_approved"] = len(all_users - users_approved_ever)
        except Exception:
            pass
    if "valor" in cols:
        # Parser robusto (Excel da numero, API puede dar "50.000" o "$50,000.00")
        valor_clean = parse_currency_series(df[cols["valor"]])
        res["amount_total"] = float(valor_clean.fillna(0).sum())
        # Per-intento (los reintentos del mismo cobro inflan estas cifras)
        res["amount_approved_attempts"] = float(valor_clean.loc[approved_mask].fillna(0).sum())
        res["amount_denied_attempts"] = float(valor_clean.loc[denied_mask].fillna(0).sum())

        # Per-venta unica (cada cobro cuenta UNA vez, independiente de reintentos).
        # Recuperado = ventas con >=1 aprobado.
        # En riesgo  = ventas sin ningun aprobado.
        # Esto refleja el dinero real, no la suma de reintentos.
        if "venta" in cols:
            sale_col = cols["venta"]
            tmp = pd.DataFrame({
                "_sale": df[sale_col].values,
                "_valor": valor_clean.values,
                "_ap": approved_mask.values,
            })
            # Para cada venta: cualquier intento aprobado, y un Valor representativo
            ever_approved = tmp.groupby("_sale", dropna=False)["_ap"].any()
            # Tomamos el max del Valor por venta (defensivo si hay diffs)
            sale_valor = tmp.groupby("_sale", dropna=False)["_valor"].max()
            res["amount_approved"] = float(sale_valor[ever_approved].fillna(0).sum())
            res["amount_denied"] = float(sale_valor[~ever_approved].fillna(0).sum())
            res["sales_total"] = int(len(sale_valor))
            res["sales_recovered"] = int(ever_approved.sum())
            res["sales_at_risk"] = int((~ever_approved).sum())
        else:
            # Sin ID de la venta no se puede agrupar; usar per-intento como fallback
            res["amount_approved"] = res["amount_approved_attempts"]
            res["amount_denied"] = res["amount_denied_attempts"]
    if "intento" in cols:
        dts = df[cols["intento"]].dropna()
        if len(dts):
            res["date_min"] = dts.min()
            res["date_max"] = dts.max()
    return res, approved_mask, denied_mask


def by_sede(df, cols, approved_mask, denied_mask):
    """Aggregate by sede.

    Total/Aprobado/Negado ahora cuentan USUARIOS UNICOS (no intentos):
      - Total    = usuarios distintos en la sede
      - Aprobado = usuarios con >= 1 intento aprobado
      - Negado   = usuarios con >= 1 intento negado
    Aprobado + Negado puede exceder Total cuando un mismo usuario tiene ambos.
    TasaExito = Aprobado / Total (fraccion de usuarios que en algun momento
    fueron aprobados en esa sede). Sin id de cliente cae al modo per-intento.
    """
    if "sede" not in cols:
        return pd.DataFrame()
    id_col = next((c for c in ID_CLIENT_CANDIDATES if c in df.columns), None)
    if id_col:
        g = df.groupby(cols["sede"], dropna=False)
        total = g[id_col].nunique()
        appr = (df.loc[approved_mask].groupby(cols["sede"], dropna=False)[id_col]
                 .nunique().reindex(total.index).fillna(0).astype(int))
        den = (df.loc[denied_mask].groupby(cols["sede"], dropna=False)[id_col]
                 .nunique().reindex(total.index).fillna(0).astype(int))
        out = pd.DataFrame({
            "Sede": total.index,
            "Total": total.values,
            "Aprobado": appr.values,
            "Negado": den.values,
        })
    else:
        g = df.groupby(cols["sede"], dropna=False)
        out = pd.DataFrame({"Sede": g.size().index, "Total": g.size().values})
        out["Aprobado"] = (df[approved_mask].groupby(cols["sede"]).size()
                            .reindex(out["Sede"]).fillna(0).astype(int).values)
        out["Negado"] = (df[denied_mask].groupby(cols["sede"]).size()
                          .reindex(out["Sede"]).fillna(0).astype(int).values)
    out["TasaExito"] = np.where(out["Total"] > 0, out["Aprobado"] / out["Total"], 0)
    out["TasaFallo"] = np.where(out["Total"] > 0, out["Negado"] / out["Total"], 0)
    out = out.sort_values(["Total"], ascending=False).reset_index(drop=True)
    return out


def by_motivo(df, cols, denied_mask):
    """Aggregate denied attempts by motivo.

    Cuenta USUARIOS UNICOS por motivo (no intentos). Si la columna `Veces`
    aparece en consumidores aguas abajo (PDF, tablas), su semantica ahora es
    "usuarios unicos con al menos un intento negado por este motivo". El
    porcentaje (`Pct`) se calcula contra el total de usuarios unicos negados.
    """
    if "motivo" not in cols:
        return pd.DataFrame()
    sub = df.loc[denied_mask].copy()
    sub[cols["motivo"]] = sub[cols["motivo"]].fillna("Sin motivo registrado")

    id_col = next((c for c in ID_CLIENT_CANDIDATES if c in sub.columns), None)
    if id_col:
        usuarios = sub.groupby(cols["motivo"], dropna=False)[id_col].nunique()
    else:
        usuarios = sub[cols["motivo"]].value_counts()
    usuarios = usuarios.sort_values(ascending=False)
    intentos = sub.groupby(cols["motivo"], dropna=False).size().reindex(usuarios.index).fillna(0).astype(int)

    total_users = int(sub[id_col].dropna().nunique()) if id_col else int(usuarios.sum())
    out = pd.DataFrame({
        "Motivo": usuarios.index,
        "Veces": usuarios.values,            # ahora = usuarios unicos
        "Intentos": intentos.values,         # cuenta cruda de intentos (referencia)
        "Pct": (usuarios.values / total_users) if total_users else 0,
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
    """Aggregate by franquicia (Marca de tarjeta). Cuenta USUARIOS UNICOS.

    Aprobado = usuarios con >=1 aprobado en esa franquicia.
    Negado   = usuarios con >=1 negado en esa franquicia.
    """
    if "marca" not in cols:
        return pd.DataFrame()
    id_col = next((c for c in ID_CLIENT_CANDIDATES if c in df.columns), None)
    marca_norm = df[cols["marca"]].astype(str).fillna("DESCONOCIDA")
    if id_col:
        # Per-franquicia: unique users
        tmp = pd.DataFrame({
            "_marca": marca_norm.values,
            "_id": df[id_col].astype(str).values,
            "_ap": approved_mask.values,
            "_dn": denied_mask.values,
        })
        total = tmp.groupby("_marca")["_id"].nunique()
        appr = tmp.loc[tmp["_ap"]].groupby("_marca")["_id"].nunique().reindex(total.index).fillna(0).astype(int)
        den = tmp.loc[tmp["_dn"]].groupby("_marca")["_id"].nunique().reindex(total.index).fillna(0).astype(int)
        out = pd.DataFrame({
            "Franquicia": total.index,
            "Aprobado": appr.values,
            "Negado": den.values,
            "Total": total.values,
        })
    else:
        # Fallback: count attempts (no id column available)
        appr_counts = marca_norm[approved_mask].value_counts()
        den_counts = marca_norm[denied_mask].value_counts()
        idx = sorted(set(appr_counts.index) | set(den_counts.index))
        out = pd.DataFrame({
            "Franquicia": idx,
            "Aprobado": [int(appr_counts.get(i, 0)) for i in idx],
            "Negado": [int(den_counts.get(i, 0)) for i in idx],
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
    vals = [s.get("approved", 0) or 0, s.get("denied", 0) or 0]
    labels = ["Aprobado", "Negado"]
    colors_ = ["#1F8A4C", "#C0392B"]
    if sum(vals) <= 0:
        ax.text(0.5, 0.5, "Sin intentos clasificados", ha="center", va="center",
                fontsize=11, color="#7A8597", transform=ax.transAxes)
        ax.set_axis_off()
        ax.set_title("Resultado global de intentos")
        fig.tight_layout()
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        return
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
                          topMargin=1.4 * cm, bottomMargin=0.8 * cm,
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
        # Banner de 1.3 cm. Lado izquierdo: titulo + sede. Lado derecho: dos
        # lineas con labels claros (rango de la data y fecha de generacion).
        canv.setFillColor(PRIMARY)
        canv.rect(0, h - 1.3 * cm, w, 1.3 * cm, fill=1, stroke=0)
        canv.setFillColor(ACCENT)
        canv.rect(0, h - 1.35 * cm, w, 0.05 * cm, fill=1, stroke=0)
        canv.setFillColor(colors.white)

        # Linea 1: marca + titulo
        canv.setFont("Helvetica-Bold", 11.5)
        canv.drawString(1.0 * cm, h - 0.50 * cm, "REPORTE DEBITOS - EVO")

        # Lado derecho (dos lineas con etiqueta explicita)
        right_top = f"Fecha de generacion de datos: {period_str}"
        right_bot = f"Generado el: {generated_str}"
        canv.setFont("Helvetica", 7.5)
        canv.drawRightString(w - 1.0 * cm, h - 0.85 * cm, right_top)
        canv.setFillColor(colors.HexColor("#C9D6E2"))
        canv.setFont("Helvetica", 7)
        canv.drawRightString(w - 1.0 * cm, h - 1.18 * cm, right_bot)

        # Sede izquierda con auto-shrink defensivo
        canv.setFillColor(colors.white)
        sede_text = f"Sede: {sede_label.upper()}"
        sede_font = "Helvetica-Bold"
        sede_size = 9.0
        # Ancho disponible: hasta donde empieza el texto derecho mas largo
        right_w = max(
            canv.stringWidth(right_top, "Helvetica", 7.5),
            canv.stringWidth(right_bot, "Helvetica", 7),
        )
        avail = (w - 2.0 * cm) - right_w - (0.6 * cm)
        while sede_size > 7.5 and canv.stringWidth(sede_text, sede_font, sede_size) > avail:
            sede_size -= 0.5
        while canv.stringWidth(sede_text, sede_font, sede_size) > avail and len(sede_text) > 12:
            sede_text = sede_text[:-4] + "..."
        canv.setFont(sede_font, sede_size)
        canv.drawString(1.0 * cm, h - 1.05 * cm, sede_text)

        # Footer
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

    right_blocks = [Paragraph("<b>Top 15 motivos de rechazo (en espanol)</b>", styles["Small"])]
    if not motivos.empty:
        m2 = motivos.head(15).copy()
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
        # Tabla de motivos con accion sugerida (texto con wrap, no truncado)
        story.append(Paragraph("<b>Acciones por motivo (top 15)</b>", styles["Small"]))
        ac = motivos.head(15).copy()

        cell_style = ParagraphStyle(
            "ActCell", parent=styles["Body"],
            fontName="Helvetica", fontSize=7.5, leading=9.5,
            textColor=colors.HexColor("#1B2535"), alignment=0,
        )
        rows = [["Motivo (ES)", "Resp.", "Accion sugerida"]]
        for _, r in ac.iterrows():
            rows.append([
                Paragraph(str(r["MotivoES"] or "").replace("&", "&amp;"), cell_style),
                Paragraph(str(r["Responsable"] or "").replace("&", "&amp;"), cell_style),
                Paragraph(str(r["Accion"] or "").replace("&", "&amp;"), cell_style),
            ])
        action_table = Table(rows,
                             colWidths=[5.4 * cm, 1.8 * cm, 11.6 * cm],
                             repeatRows=1)
        ts = [
            ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7.5),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LINEBELOW", (0, 0), (-1, 0), 1, ACCENT),
            ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#DCE2EC")),
        ]
        for i in range(1, len(rows)):
            if i % 2 == 0:
                ts.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
        action_table.setStyle(TableStyle(ts))
        story.append(action_table)
        story.append(Spacer(1, 3))

    # --- Franquicias (usuarios unicos) ---
    if not marcas.empty:
        franq_rows = [["Franquicia", "Aprobado", "Negado"]]
        # Mostrar hasta 8 para no romper la pagina
        for _, r in marcas.head(8).iterrows():
            franq_rows.append([str(r["Franquicia"]), fmt_int(r["Aprobado"]), fmt_int(r["Negado"])])
        franq_table = Table(
            franq_rows,
            colWidths=[6.5 * cm, 2.5 * cm, 2.5 * cm],
            repeatRows=1,
        )
        franq_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7.5),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, -1), 7.5),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LINEBELOW", (0, 0), (-1, 0), 1, ACCENT),
            ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#DCE2EC")),
            *[("BACKGROUND", (0, i), (-1, i), LIGHT)
              for i in range(1, len(franq_rows)) if i % 2 == 0],
        ]))
        story.append(Paragraph(
            "<b>Franquicias (usuarios unicos)</b>", styles["Small"]))
        story.append(franq_table)
        story.append(Spacer(1, 3))

    # --- Conclusiones + Recomendaciones (compact, side by side) ---
    insights = build_insights(summary, sedes, motivos, tipos)
    recs = build_recommendations(summary, sedes, motivos, tipos)

    concl_style = ParagraphStyle("ConclTxt", parent=styles["Body"], fontSize=7.5, leading=9.5,
                                 textColor=colors.HexColor("#1B2535"))
    h_style = ParagraphStyle("h", parent=styles["Small"], fontSize=8.5, textColor=PRIMARY, leading=10)
    concl_left = [Paragraph("<b>Conclusiones clave</b>", h_style)]
    for it in insights[:5]:
        concl_left.append(Paragraph(f"• {it}", concl_style))

    concl_right = [Paragraph("<b>Recomendaciones</b>", h_style)]
    for r in recs[:5]:
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
    if "users_never_approved" in summary and summary.get("clients_total"):
        nv = summary["users_never_approved"]
        tot = summary["clients_total"]
        pct = (nv / tot) * 100 if tot else 0
        out.append(f"<b>{fmt_int(nv)}</b> usuarios distintos ({pct:.1f}%) "
                   f"nunca tuvieron un debito aprobado en el periodo.")
    if "amount_total" in summary:
        recovered = summary.get("amount_approved", 0)
        lost = summary.get("amount_denied", 0)
        clients_total = summary.get("clients_total") or 0
        ever_ap = summary.get("users_ever_approved") or 0
        never_ap = summary.get("users_never_approved") or 0
        sales_total = summary.get("sales_total") or 0
        sales_rec = summary.get("sales_recovered") or 0
        sales_at_risk = summary.get("sales_at_risk") or 0
        pct_rec_users = (ever_ap / clients_total * 100) if clients_total else 0
        pct_risk_users = (never_ap / clients_total * 100) if clients_total else 0
        out.append(
            f"Monto total cursado: <b>{fmt_money(summary['amount_total'])}</b> "
            f"({fmt_int(clients_total)} usuarios distintos)."
        )
        out.append(
            f"Recuperado: <b>{fmt_money(recovered)}</b> "
            f"({pct_rec_users:.1f}% de usuarios = {fmt_int(ever_ap)} con aprobado; "
            f"{fmt_int(sales_rec)} de {fmt_int(sales_total)} ventas)."
        )
        out.append(
            f"En riesgo: <b>{fmt_money(lost)}</b> "
            f"({pct_risk_users:.1f}% de usuarios = {fmt_int(never_ap)} nunca aprobados; "
            f"{fmt_int(sales_at_risk)} ventas)."
        )
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
        recs.append(f"<b>Motivo #1</b> ({top1['MotivoES']}, {top1['Pct']*100:.1f}% usuarios): {top1['Accion']}")
    if not tipos.empty:
        rev = int(tipos[tipos["Tipo"].astype(str).str.lower() == "reversible"]["Veces"].sum())
        if rev > 0:
            recs.append(f"Reintentos automaticos a 24h y 72h sobre <b>{fmt_int(rev)}</b> rechazos reversibles.")
    if not sedes.empty:
        avg = summary["success_rate"]
        bad = sedes[(sedes["Total"] >= 200) & (sedes["TasaExito"] < avg * 0.85)]
        if not bad.empty:
            recs.append(f"Auditoria operativa en <b>{len(bad)} sede(s)</b> por debajo del 85% del promedio.")
    never_ap = summary.get("users_never_approved") or 0
    if never_ap > 0:
        recs.append(
            f"Contactar a los <b>{fmt_int(never_ap)}</b> usuarios nunca aprobados (SMS + email) "
            f"para actualizar medio de pago."
        )
    sales_risk = summary.get("sales_at_risk") or 0
    risk_money = summary.get("amount_denied", 0)
    if sales_risk > 0:
        recs.append(
            f"Gestion manual sobre <b>{fmt_int(sales_risk)}</b> ventas en riesgo "
            f"(<b>{fmt_money(risk_money)}</b>), priorizando >= 3 intentos negados."
        )
    recs.append("Tablero diario por sede; alerta si baja 5pp del promedio 2 dias seguidos.")
    recs.append("SLA mensual de tasa de exito por sede ligado al responsable.")
    recs.append("Cruce 'profile not found' / 'tarjeta no registrada' con afiliacion (datos faltantes).")
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
