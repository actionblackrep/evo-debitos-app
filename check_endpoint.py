#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_endpoint.py - prueba directa de /api/debitos con from/to

Uso:
    python check_endpoint.py                            # ayer -> hoy
    python check_endpoint.py 2026-05-05 2026-05-10
    python check_endpoint.py 2026-04-29 2026-05-03 --raw
"""
from __future__ import annotations
import argparse, json, os, sys, time, urllib.error, urllib.parse, urllib.request
from collections import Counter
from datetime import datetime, date, timedelta

DEFAULT_URL = os.environ.get("EVO_DEBITS_URL",
    "https://action-branches-api.vercel.app/api/debitos")
DEFAULT_TOKEN = os.environ.get("EVO_DEBITS_API_KEY",
    "dLjaU5u4LfycyRpbBTU7EMcXDBL2zFrOiX6fBWO6b-s")

VARIANTS = [
    ("x-api-key header",    "header", "x-api-key",      "{t}"),
    ("?key= query param",   "qs",     "key",            "{t}"),
    ("X-API-Token header",  "header", "X-API-Token",    "{t}"),
    ("api-token header",    "header", "api-token",      "{t}"),
    ("Authorization Bearer","header", "Authorization",  "Bearer {t}"),
]


def http_get(url, headers, timeout=45):
    req = urllib.request.Request(url, headers=headers, method="GET")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def extract_records(payload):
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for k in ("data", "records", "items", "results", "rows"):
            if isinstance(payload.get(k), list):
                return payload[k]
    return []


def split_by_month(f, t):
    out = []
    cur = f
    while cur < t:
        nxt = date(cur.year + (1 if cur.month == 12 else 0),
                   (cur.month % 12) + 1, 1)
        end = min(nxt, t)
        out.append((cur, end))
        cur = end
    return out


def fetch_one(url, token, from_s, to_s, page_size=1000):
    last_err = ""
    for label, kind, name, tmpl in VARIANTS:
        val = tmpl.format(t=token)
        rows = []
        page = 1
        ok = False
        while page <= 30:
            qs = {"from": from_s, "to": to_s, "page": page, "size": page_size}
            headers = {"Accept": "application/json", "User-Agent": "evo-check/3.0"}
            if kind == "header":
                headers[name] = val
            else:
                qs[name] = val
            full = url + "?" + urllib.parse.urlencode(qs)
            t0 = time.time()
            try:
                raw = http_get(full, headers)
                payload = json.loads(raw)
                ok = True
            except urllib.error.HTTPError as e:
                last_err = f"HTTP {e.code} con {label}: {e.read()[:120].decode('utf-8','replace')}"
                ok = False
                break
            except Exception as e:
                last_err = f"{type(e).__name__} con {label}: {e}"
                ok = False
                break
            recs = extract_records(payload)
            dt = time.time() - t0
            print(f"     [{label}] pagina {page}: {len(recs):>6} recs ({dt:.1f}s)")
            if not recs:
                break
            rows.extend(recs)
            if len(recs) < page_size:
                break
            page += 1
        if ok:
            return rows, label
        print(f"   -> {last_err}")
    raise RuntimeError("Ninguna variante de auth funciono")


def parse_iso(s):
    if not isinstance(s, str):
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None


def main():
    p = argparse.ArgumentParser()
    today = date.today()
    p.add_argument("from_date", nargs="?", default=(today - timedelta(days=1)).isoformat())
    p.add_argument("to_date",   nargs="?", default=today.isoformat())
    p.add_argument("--url", default=DEFAULT_URL)
    p.add_argument("--token", default=DEFAULT_TOKEN)
    p.add_argument("--raw", action="store_true")
    args = p.parse_args()

    f = date.fromisoformat(args.from_date)
    t = date.fromisoformat(args.to_date)
    if t <= f:
        print("ERROR: to debe ser > from")
        sys.exit(2)
    if (t - f).days > 5:
        print(f"WARN: rango {(t-f).days} dias (>5)")

    chunks = split_by_month(f, t)
    print(f"URL:    {args.url}")
    print(f"Rango:  {f} -> {t}   ({(t-f).days} dias, {len(chunks)} llamada(s))")
    print()

    all_rows = []
    used = None
    for cf, ct in chunks:
        print(f"-- Chunk: {cf} -> {ct}")
        rows, label = fetch_one(args.url, args.token, cf.isoformat(), ct.isoformat())
        used = used or label
        all_rows.extend(rows)
        print(f"   subtotal: {len(rows):,}\n")

    print(f">> OK con `{used}`  ·  total recibidos: {len(all_rows):,}")

    if not all_rows:
        return

    if args.raw:
        print("\n>> Primer record:")
        print(json.dumps(all_rows[0], indent=2, ensure_ascii=False, default=str)[:2000])

    cols = sorted({k for r in all_rows[:2000] for k in r.keys()})
    print(f"\n>> Columnas ({len(cols)}):")
    for c in cols:
        print(f"   - {c}")

    intento_key = next((c for c in cols if c.lower() in ("intento","fecha","date","attempt_date")), None)
    if intento_key:
        dts = sorted([d for d in (parse_iso(r.get(intento_key)) for r in all_rows) if d])
        if dts:
            print(f"\n>> {intento_key}: min={dts[0].isoformat()}  max={dts[-1].isoformat()}")
            by_day = Counter(d.date() for d in dts)
            print(">> Conteo por dia (top 10):")
            for day, n in sorted(by_day.items(), reverse=True)[:10]:
                print(f"   {day}: {n:,}")


if __name__ == "__main__":
    main()
