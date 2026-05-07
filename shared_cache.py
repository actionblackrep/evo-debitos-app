# -*- coding: utf-8 -*-
"""
Cache compartido entre app.py (admin) y app_managers.py (Gerentes Generales).

Estrategia:
- El admin publica el ultimo dataset como parquet en una rama dedicada del
  repo de GitHub (`data-cache` por defecto). Esa rama NO se despliega.
- La app de GM hace pull desde esa rama via API de GitHub cuando la API EVO no
  responde.

Secrets requeridos en Streamlit Community Cloud (mismos para ambas apps):
  GITHUB_TOKEN          = "<fine-grained PAT con Contents: Read and write>"
  GITHUB_REPO           = "USUARIO/evo-debitos-app"
  GITHUB_DATA_BRANCH    = "data-cache"          # opcional, default data-cache
  GITHUB_DATA_PATH      = "data/latest.parquet" # opcional, default data/latest.parquet

Si los secrets NO estan presentes las funciones devuelven None / False
silenciosamente: la app sigue operando con su flujo local.

Solo stdlib (urllib + json + base64). No requiere requests.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import tempfile
import urllib.error
import urllib.request
from datetime import datetime, timezone
from typing import Optional

GITHUB_API = "https://api.github.com"
DEFAULT_BRANCH = "data-cache"
DEFAULT_PATH = "data/latest.parquet"


# ---------- secrets ----------
def _get_secret(key: str, default: str = "") -> str:
    """Read from env or st.secrets (when running in Streamlit) without crashing."""
    val = os.environ.get(key)
    if val:
        return val
    try:
        import streamlit as st  # type: ignore
        try:
            return st.secrets[key]  # type: ignore[index]
        except Exception:
            return default
    except Exception:
        return default


def cache_config() -> dict:
    return {
        "token": _get_secret("GITHUB_TOKEN"),
        "repo": _get_secret("GITHUB_REPO"),
        "branch": _get_secret("GITHUB_DATA_BRANCH", DEFAULT_BRANCH),
        "path": _get_secret("GITHUB_DATA_PATH", DEFAULT_PATH),
    }


def is_configured() -> bool:
    c = cache_config()
    return bool(c["token"] and c["repo"])


# ---------- low level GH ----------
def _gh(method: str, url: str, token: str, payload=None,
        accept="application/vnd.github+json", timeout: int = 30):
    headers = {
        "Authorization": f"token {token}",
        "Accept": accept,
        "User-Agent": "evo-debitos-cache/1.0",
    }
    data = None
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    return urllib.request.urlopen(req, timeout=timeout)


def _default_branch(c) -> Optional[str]:
    try:
        with _gh("GET", f"{GITHUB_API}/repos/{c['repo']}", c["token"]) as r:
            return json.loads(r.read()).get("default_branch")
    except Exception:
        return None


def _branch_sha(c, branch: str) -> Optional[str]:
    try:
        with _gh("GET", f"{GITHUB_API}/repos/{c['repo']}/git/refs/heads/{branch}",
                 c["token"]) as r:
            return json.loads(r.read())["object"]["sha"]
    except Exception:
        return None


def ensure_branch(c=None) -> bool:
    """Create the data branch from default branch if it does not exist."""
    c = c or cache_config()
    if not c["token"] or not c["repo"]:
        return False
    # Check if branch exists
    try:
        with _gh("GET", f"{GITHUB_API}/repos/{c['repo']}/branches/{c['branch']}",
                 c["token"]) as r:
            return True
    except urllib.error.HTTPError as e:
        if e.code != 404:
            return False
    except Exception:
        return False
    # Create from default branch
    default = _default_branch(c)
    if not default:
        return False
    sha = _branch_sha(c, default)
    if not sha:
        return False
    try:
        payload = {"ref": f"refs/heads/{c['branch']}", "sha": sha}
        with _gh("POST", f"{GITHUB_API}/repos/{c['repo']}/git/refs",
                 c["token"], payload) as r:
            return r.status in (200, 201)
    except Exception:
        return False


def _remote_sha(c) -> Optional[str]:
    url = f"{GITHUB_API}/repos/{c['repo']}/contents/{c['path']}?ref={c['branch']}"
    try:
        with _gh("GET", url, c["token"]) as r:
            return json.loads(r.read()).get("sha")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        return None
    except Exception:
        return None


# ---------- public API ----------
def push_parquet(parquet_path: str) -> bool:
    """Upload (create or replace) parquet in the shared cache. Idempotent."""
    if not is_configured():
        return False
    c = cache_config()
    ensure_branch(c)
    try:
        with open(parquet_path, "rb") as f:
            raw = f.read()
    except OSError:
        return False
    sig = hashlib.sha256(raw).hexdigest()[:12]
    payload = {
        "message": f"data sync {datetime.now(timezone.utc).isoformat(timespec='seconds')} [{sig}]",
        "content": base64.b64encode(raw).decode(),
        "branch": c["branch"],
    }
    sha = _remote_sha(c)
    if sha:
        payload["sha"] = sha
    url = f"{GITHUB_API}/repos/{c['repo']}/contents/{c['path']}"
    try:
        with _gh("PUT", url, c["token"], payload, timeout=60) as r:
            return r.status in (200, 201)
    except Exception:
        return False


def pull_parquet() -> Optional[str]:
    """Download latest parquet from cache. Returns local path or None."""
    if not is_configured():
        return None
    c = cache_config()
    url = f"{GITHUB_API}/repos/{c['repo']}/contents/{c['path']}?ref={c['branch']}"
    try:
        with _gh("GET", url, c["token"], accept="application/vnd.github.v3.raw",
                 timeout=60) as r:
            content = r.read()
    except Exception:
        return None
    if not content:
        return None
    digest = hashlib.md5(content).hexdigest()[:16]
    out = os.path.join(tempfile.gettempdir(), f"evo_debits_shared_{digest}.parquet")
    if not (os.path.exists(out) and os.path.getsize(out) == len(content)):
        with open(out, "wb") as f:
            f.write(content)
    return out


def remote_meta() -> Optional[dict]:
    """Returns {sha, date, message} of the last commit that touched the cache file."""
    if not is_configured():
        return None
    c = cache_config()
    url = (f"{GITHUB_API}/repos/{c['repo']}/commits"
           f"?path={c['path']}&sha={c['branch']}&per_page=1")
    try:
        with _gh("GET", url, c["token"]) as r:
            commits = json.loads(r.read())
            if not commits:
                return None
            commit = commits[0]
            return {
                "sha": commit["sha"][:8],
                "date": commit["commit"]["committer"]["date"],
                "message": commit["commit"]["message"],
            }
    except Exception:
        return None


def publish_dataframe(df, source_label: str = "") -> tuple[bool, str]:
    """Save df to a temp parquet and push to cache. Returns (ok, message)."""
    if not is_configured():
        return False, "Cache compartido no configurado (faltan GITHUB_TOKEN / GITHUB_REPO)."
    pq_path = os.path.join(tempfile.gettempdir(), "evo_debits_publish.parquet")
    try:
        # Reduce object dtype bloat before serializing
        try:
            for col in df.select_dtypes(include="object").columns:
                df[col] = df[col].astype("string")
        except Exception:
            pass
        df.to_parquet(pq_path, index=False, compression="snappy")
    except Exception as e:
        return False, f"No se pudo serializar a parquet: {e}"
    ok = push_parquet(pq_path)
    if ok:
        rows_str = f"{len(df):,}".replace(",", ".")
        suffix = f" ({source_label})" if source_label else ""
        return True, f"Publicado para Gerentes Generales: {rows_str} filas{suffix}."
    return False, "Falla al publicar (verifica GITHUB_TOKEN/GITHUB_REPO o conexion)."
