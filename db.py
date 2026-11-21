"""Database layer for the TGFC Buy/Sell app.

Uses the same Postgres database (Supabase) as the Freight Desk when DATABASE_URL
is set in Streamlit secrets or the environment, so the team signs in with the
same accounts. Falls back to a local SQLite file for testing.

It only ADDS three tables of its own (bs_access, buy_sells, buy_sell_lines).
The `users` table below is the Freight Desk's, written exactly the same way;
CREATE TABLE IF NOT EXISTS leaves the existing one untouched.
"""
from __future__ import annotations

import os
import sqlite3
from datetime import datetime, timezone
from decimal import Decimal

try:
    import streamlit as st
except Exception:  # pragma: no cover - allows import outside Streamlit
    st = None


def _database_url() -> str | None:
    url = os.environ.get("DATABASE_URL")
    if not url and st is not None:
        try:
            url = st.secrets.get("DATABASE_URL")
        except Exception:
            url = None
    return url


IS_PG = bool(_database_url())

SCHEMA_PG = """
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    email TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('sales','freight','admin')),
    password_hash TEXT NOT NULL,
    must_change BOOLEAN NOT NULL DEFAULT TRUE,
    active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS bs_access (
    user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
    buy_side BOOLEAN NOT NULL DEFAULT FALSE,
    updated_by INTEGER REFERENCES users(id),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS buy_sells (
    id SERIAL PRIMARY KEY,
    po TEXT UNIQUE NOT NULL,
    order_date DATE,
    pu_date DATE,
    vendor TEXT NOT NULL DEFAULT '',
    contact TEXT NOT NULL DEFAULT '',
    release_no TEXT NOT NULL DEFAULT '',
    pu_location TEXT NOT NULL DEFAULT '',
    customer TEXT NOT NULL DEFAULT '',
    del_po TEXT NOT NULL DEFAULT '',
    delivery_date DATE,
    del_location TEXT NOT NULL DEFAULT '',
    pp_discount_pct NUMERIC,
    program_fee NUMERIC,
    pua NUMERIC,
    freight NUMERIC,
    other_exp NUMERIC,
    commissions NUMERIC,
    split_pct NUMERIC,
    notes TEXT NOT NULL DEFAULT '',
    created_by INTEGER REFERENCES users(id),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_by INTEGER REFERENCES users(id),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS buy_sell_lines (
    id SERIAL PRIMARY KEY,
    buy_sell_id INTEGER NOT NULL REFERENCES buy_sells(id) ON DELETE CASCADE,
    line_no INTEGER NOT NULL DEFAULT 0,
    item_code TEXT NOT NULL DEFAULT '',
    description TEXT NOT NULL DEFAULT '',
    cases NUMERIC,
    pounds NUMERIC,
    cost NUMERIC,
    sale_price NUMERIC
);
"""

SCHEMA_SQLITE = (
    SCHEMA_PG.replace("SERIAL PRIMARY KEY", "INTEGER PRIMARY KEY AUTOINCREMENT")
    .replace("TIMESTAMPTZ NOT NULL DEFAULT NOW()", "TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))")
    .replace("TIMESTAMPTZ", "TEXT")
    .replace("BOOLEAN NOT NULL DEFAULT TRUE", "INTEGER NOT NULL DEFAULT 1")
    .replace("BOOLEAN NOT NULL DEFAULT FALSE", "INTEGER NOT NULL DEFAULT 0")
    .replace("NUMERIC", "REAL")
    .replace(" DATE,", " TEXT,")
)


def _connect():
    url = _database_url()
    if url:
        import psycopg2
        import psycopg2.extras

        conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor, connect_timeout=10)
        conn.autocommit = True
        return conn
    path = os.environ.get("BUY_SELL_SQLITE", "buy_sell.sqlite3")
    conn = sqlite3.connect(path, check_same_thread=False, isolation_level=None)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


_conn = None


def conn():
    global _conn
    if _conn is None:
        _conn = _connect()
        init_schema(_conn)
    else:
        # make sure a dropped Postgres connection is re-opened
        try:
            if IS_PG and _conn.closed:
                _conn = _connect()
        except Exception:
            _conn = _connect()
    return _conn


def init_schema(c):
    if IS_PG:
        with c.cursor() as cur:
            cur.execute(SCHEMA_PG)
    else:
        c.executescript(SCHEMA_SQLITE)


def _q(sql: str) -> str:
    return sql if IS_PG else sql.replace("%s", "?")


def _fix_params(params):
    if IS_PG:
        return params
    # SQLite has no native booleans/dates; store as ints / ISO strings
    out = []
    for p in params:
        if isinstance(p, bool):
            out.append(1 if p else 0)
        elif hasattr(p, "isoformat"):
            out.append(p.isoformat())
        else:
            out.append(p)
    return tuple(out)


def _clean(row: dict) -> dict:
    # Postgres NUMERIC comes back as Decimal; the app works in floats
    return {k: (float(v) if isinstance(v, Decimal) else v) for k, v in row.items()}


def run(sql: str, params=()):
    """Execute a write statement. Returns the new row id when RETURNING id is used."""
    c = conn()
    sql = _q(sql)
    params = _fix_params(params)
    try:
        if IS_PG:
            with c.cursor() as cur:
                cur.execute(sql, params)
                if cur.description:
                    row = cur.fetchone()
                    return row["id"] if row and "id" in row else row
                return None
        cur = c.execute(sql, params)
        if cur.description:
            row = cur.fetchone()
            return row["id"] if row and "id" in row.keys() else row
        return cur.lastrowid
    except Exception:
        if IS_PG:
            try:
                c.rollback()
            except Exception:
                pass
        raise


def fetchall(sql: str, params=()):
    c = conn()
    sql = _q(sql)
    params = _fix_params(params)
    if IS_PG:
        with c.cursor() as cur:
            cur.execute(sql, params)
            return [_clean(dict(r)) for r in cur.fetchall()]
    cur = c.execute(sql, params)
    return [_clean(dict(r)) for r in cur.fetchall()]


def fetchone(sql: str, params=()):
    rows = fetchall(sql, params)
    return rows[0] if rows else None


def now():
    return datetime.now(timezone.utc)


def as_bool(v) -> bool:
    return bool(v) and v not in (0, "0", "false", "False")
