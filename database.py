"""
database.py — SQLite persistence for SHOWLOG / DataHub.

All DB access goes through this module so the storage backend can be swapped
(e.g. to PostgreSQL) without touching the rest of the app.

Tables
------
files        : one row per uploaded source file (dedupe by filename)
shows        : one row per transformed show (the Zee-format output)
channel_map  : button label -> raw CHNLNAME value (the 7-channel mapper)
"""

import sqlite3
import json
import os
from contextlib import contextmanager

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "showlog.db")

CHANNEL_BUTTONS = ["Zee", "Star Plus", "Colors", "Set", "Sab",
                   "Star Bharat", "&TV"]


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")  # better concurrency on a server
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """Create tables if they do not exist."""
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS files (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                filename    TEXT UNIQUE NOT NULL,
                uploaded_at TEXT NOT NULL,
                row_count   INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS shows (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id   INTEGER NOT NULL,
                channel   TEXT,
                show_date TEXT,
                sort_st   INTEGER,
                data      TEXT NOT NULL,
                FOREIGN KEY (file_id) REFERENCES files(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_shows_file    ON shows(file_id);
            CREATE INDEX IF NOT EXISTS idx_shows_channel ON shows(channel);
            CREATE INDEX IF NOT EXISTS idx_shows_date    ON shows(show_date);

            CREATE TABLE IF NOT EXISTS channel_map (
                button TEXT PRIMARY KEY,
                value  TEXT
            );
            """
        )


# ---------------------------------------------------------------------------
# Files + shows
# ---------------------------------------------------------------------------

def add_file(filename, uploaded_at, show_rows):
    """
    Insert (or replace) a file and its transformed show rows.
    Re-adding a file with the same name replaces its rows (no double-counting).
    Returns the file id.
    """
    with get_conn() as conn:
        # Remove any existing file of the same name (cascade drops its shows).
        conn.execute("DELETE FROM files WHERE filename = ?", (filename,))
        cur = conn.execute(
            "INSERT INTO files (filename, uploaded_at, row_count) "
            "VALUES (?, ?, ?)",
            (filename, uploaded_at, len(show_rows)),
        )
        file_id = cur.lastrowid

        payload = []
        for r in show_rows:
            # _date is a date object; serialise sortable fields separately.
            d = r.get("_date")
            show_date = d.isoformat() if d else None
            sort_st = r.get("_prog_st")
            clean = {k: v for k, v in r.items() if not k.startswith("_")}
            payload.append((
                file_id,
                r.get("_channel") or "",
                show_date,
                sort_st if sort_st is not None else 0,
                json.dumps(clean),
            ))
        conn.executemany(
            "INSERT INTO shows (file_id, channel, show_date, sort_st, data) "
            "VALUES (?, ?, ?, ?, ?)",
            payload,
        )
        return file_id


def list_files():
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT id, filename, uploaded_at, row_count "
            "FROM files ORDER BY uploaded_at DESC, id DESC"
        ).fetchall()
        return [dict(r) for r in rows]


def delete_file(file_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM files WHERE id = ?", (file_id,))


def clear_all():
    with get_conn() as conn:
        conn.execute("DELETE FROM shows")
        conn.execute("DELETE FROM files")


def get_all_shows():
    """
    Return all show rows as a list of dicts, sorted by date then start time.
    Each dict includes the parsed Zee-format columns plus _channel /
    _show_date / _sort_st (private keys used for filtering and grouping).
    """
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT channel, show_date, sort_st, data FROM shows "
            "ORDER BY show_date IS NULL, show_date ASC, sort_st ASC"
        ).fetchall()
    result = []
    for r in rows:
        d = json.loads(r["data"])
        d["_channel"] = r["channel"]
        d["_show_date"] = r["show_date"]
        d["_sort_st"] = r["sort_st"]
        result.append(d)
    return result


# ---------------------------------------------------------------------------
# Channel mapper
# ---------------------------------------------------------------------------

def get_channel_map():
    """Return {button: value} for all 7 buttons (missing -> '')."""
    with get_conn() as conn:
        rows = conn.execute("SELECT button, value FROM channel_map").fetchall()
    stored = {r["button"]: r["value"] for r in rows}
    return {b: stored.get(b, "") for b in CHANNEL_BUTTONS}


def set_channel_map(mapping):
    """Persist the {button: value} mapping."""
    with get_conn() as conn:
        for button in CHANNEL_BUTTONS:
            value = mapping.get(button, "") or ""
            conn.execute(
                "INSERT INTO channel_map (button, value) VALUES (?, ?) "
                "ON CONFLICT(button) DO UPDATE SET value = excluded.value",
                (button, value),
            )


def distinct_channels():
    """All distinct raw CHNLNAME values present in the data."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT DISTINCT channel FROM shows "
            "WHERE channel IS NOT NULL AND channel != '' "
            "ORDER BY channel"
        ).fetchall()
    return [r["channel"] for r in rows]
