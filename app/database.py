"""
lightweight sqlite wrapper for storing processed complaints.
keeps history across uploads so you can track trends later.
"""

import sqlite3
import json
from datetime import datetime
from typing import List, Optional

DB_PATH = "data/complaints.db"


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """
    creates tables if they don't exist yet.
    safe to call every time the app starts.
    """
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            uploaded_at TEXT NOT NULL,
            total_rows INTEGER DEFAULT 0,
            early INTEGER DEFAULT 0,
            delayed INTEGER DEFAULT 0,
            on_time INTEGER DEFAULT 0,
            pending INTEGER DEFAULT 0,
            total_penalty REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS complaints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id INTEGER NOT NULL,
            complaint_id INTEGER,
            ro_code INTEGER,
            ro_name TEXT,
            vendor_code TEXT,
            assignment_time TEXT,
            due_time TEXT,
            close_time TEXT,
            duration_text TEXT,
            delay_hours REAL,
            status TEXT,
            penalty INTEGER,
            sla_hours REAL,
            is_auto_closed INTEGER DEFAULT 0,
            FOREIGN KEY (upload_id) REFERENCES uploads(id)
        );
    """)
    conn.commit()
    conn.close()


def save_upload(filename: str, records: List[dict], summary: dict) -> int:
    """
    stores an upload session and all its complaint records.
    returns the upload id.
    """
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO uploads (filename, uploaded_at, total_rows, early, delayed, on_time, pending, total_penalty)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        filename,
        datetime.now().isoformat(),
        summary["total"],
        summary["early"],
        summary["delayed"],
        summary["on_time"],
        summary["pending"],
        summary["total_penalty"],
    ))
    upload_id = cur.lastrowid

    for rec in records:
        cur.execute("""
            INSERT INTO complaints (
                upload_id, complaint_id, ro_code, ro_name, vendor_code,
                assignment_time, due_time, close_time, duration_text,
                delay_hours, status, penalty, sla_hours, is_auto_closed
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            upload_id,
            rec["complaint_id"],
            rec["ro_code"],
            rec["ro_name"],
            rec["vendor_code"],
            rec["assignment_time"].isoformat() if rec["assignment_time"] else None,
            rec["due_time"].isoformat() if rec["due_time"] else None,
            rec["close_time"].isoformat() if rec["close_time"] else None,
            rec["duration_text"],
            rec["delay_hours"],
            rec["status"],
            rec["penalty"],
            rec["sla_hours"],
            1 if rec["is_auto_closed"] else 0,
        ))

    conn.commit()
    conn.close()
    return upload_id


def get_upload_history(limit: int = 20) -> List[dict]:
    """
    returns recent uploads for the sidebar / history view.
    """
    conn = get_conn()
    rows = conn.execute("""
        SELECT * FROM uploads ORDER BY uploaded_at DESC LIMIT ?
    """, (limit,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_complaints_by_upload(upload_id: int) -> List[dict]:
    conn = get_conn()
    rows = conn.execute("""
        SELECT * FROM complaints WHERE upload_id = ? ORDER BY complaint_id
    """, (upload_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]
