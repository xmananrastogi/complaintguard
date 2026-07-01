"""
tests for the processor — makes sure the logic matches
what the vba macro was doing (but way easier to debug).
"""

import os
import sys
import tempfile
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pandas as pd
from app.processor import (
    parse_sla,
    parse_datetime,
    format_duration,
    check_auto_close,
    process_excel,
)


def test_parse_sla():
    assert parse_sla("48 hours") == 48
    assert parse_sla("24 hours") == 24
    assert parse_sla("24 HOURS") == 24
    assert parse_sla("NA") is None
    assert parse_sla(None) is None
    assert parse_sla("") is None


def test_parse_datetime():
    dt = parse_datetime("30-Apr-26 09:39:01 PM")
    assert dt is not None
    assert dt.year == 2026
    assert dt.month == 4
    assert dt.day == 30
    assert dt.hour == 21

    assert parse_datetime(None) is None
    assert parse_datetime("NA") is None
    assert parse_datetime("PENDING") is None


def test_format_duration():
    assert format_duration(0) == "0 Hours"
    assert format_duration(1.5) == "+1 Day 12 Hours"
    assert format_duration(-0.5) == "-12 Hours"
    assert format_duration(-2.0) == "-2 Day(s)"
    assert format_duration(0.75) == "+18 Hours"


def test_check_auto_close():
    assert check_auto_close("Auto closed after 24 Hrs") is True
    assert check_auto_close("auto close") is True
    assert check_auto_close("Completed By GVR") is False
    assert check_auto_close(None) is False


def test_process_excel():
    """
    create a mini excel file with known data and verify
    the processor handles it correctly.
    """
    data = {
        "Complaint ID": [1001, 1002, 1003],
        "Complaint Resolution Time": ["48 hours", "24 hours", "NA"],
        "Complaint DateTime": [
            "01-May-26 10:00:00 AM",
            "02-May-26 02:00:00 PM",
            "03-May-26 08:00:00 AM",
        ],
        "Vendor Close DateTime": [
            "02-May-26 10:00:00 AM",  # 24h early from 48h sla → Early
            "04-May-26 06:00:00 PM",  # 52h late (24h sla) → 2 days → ₹2,000
            "NA",                      # pending
        ],
        "RO Code": [101, 102, 103],
        "RO Name": ["Dealer A", "Dealer B", "Dealer C"],
        "Vendor Code": ["VEND1", "VEND2", "VEND3"],
        "Vendor Remarks": ["ok", "Auto closed after 24 Hrs", None],
    }

    df = pd.DataFrame(data)

    # write to temp file
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        df.to_excel(f.name, index=False)
        fpath = f.name

    records, summary = process_excel(fpath)
    os.unlink(fpath)

    assert summary["total"] == 3
    assert summary["early"] == 1   # 1001
    assert summary["delayed"] == 1  # 1002
    assert summary["pending"] == 1  # 1003
    assert summary["total_penalty"] == 1000  # 28h late → 1 full day → ₹1,000

    # check auto-close flag on row 2 (1002)
    r2 = [r for r in records if r["complaint_id"] == 1002][0]
    assert r2["is_auto_closed"] is True

    # 1001 should be early
    r1 = [r for r in records if r["complaint_id"] == 1001][0]
    assert r1["status"] == "Early"
    assert r1["penalty"] == 0

    print("all tests passed!")
