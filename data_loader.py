# -*- coding: utf-8 -*-
"""
Capa de carga optimizada para EVO Reporte de Debitos.
Reusa generate_report.py para no romper la logica/calculo del reporte.

Optimizaciones:
- Excel -> parquet snappy una sola vez (ahorra ~10x en relecturas).
- read_parquet con columns= (lectura por columna).
- read_parquet con filters= (predicate pushdown via pyarrow).
- Cache TTL via st.cache_data cuando se importa desde Streamlit.

Sin Streamlit tambien funciona (los decoradores se reemplazan por no-op).
"""
from __future__ import annotations

import hashlib
import os
import tempfile
from typing import Optional

import pandas as pd

import generate_report as gr

try:
    import streamlit as st  # type: ignore
    cache_data = st.cache_data
except Exception:
    def cache_data(*dargs, **dkwargs):
        def deco(f):
            return f
        return deco


SEDE_COL_CANDIDATES = ["Sede/club", "Sede", "Club", "sede", "club", "sede_club", "branch"]


@cache_data(ttl=1800, show_spinner=False)
def excel_bytes_to_parquet(file_bytes: bytes, name: str) -> str:
    """Convierte Excel a parquet una sola vez (cache por hash de bytes)."""
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


@cache_data(ttl=1800, show_spinner=False)
def parquet_sede_column(pq_path: str) -> Optional[str]:
    import pyarrow.parquet as pq
    schema_names = pq.ParquetFile(pq_path).schema_arrow.names
    for c in SEDE_COL_CANDIDATES:
        if c in schema_names:
            return c
    return None


@cache_data(ttl=1800, show_spinner=False)
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


@cache_data(ttl=900, show_spinner=False)
def parquet_filter_by_sede(pq_path: str, sede_col: str, sede: str) -> pd.DataFrame:
    """Predicate pushdown: pyarrow lee solo las filas que matchean."""
    return pd.read_parquet(pq_path, filters=[(sede_col, "=", sede)])


@cache_data(ttl=300, show_spinner=False)
def api_list_sedes(api_url: str, api_key: str):
    df = gr.fetch_from_api(api_url, api_key)
    cols = gr.resolve_columns(df)
    if "sede" not in cols:
        return []
    return sorted(df[cols["sede"]].dropna().astype(str).unique().tolist())


@cache_data(ttl=300, show_spinner=False)
def api_load_for_sede(api_url: str, api_key: str, sede: str) -> pd.DataFrame:
    """Intenta filtrar en el servidor; si la API ignora ?sede= filtra cliente."""
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
