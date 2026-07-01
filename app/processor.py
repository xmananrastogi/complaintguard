"""
core engine — replaces the VBA macro entirely.
takes an excel file, spits out analysed records.

column mapping is for the specific IOCL export format
we worked with. if the columns change, just update the
dict at the top.
"""

import math
import re
from datetime import datetime, timedelta
from typing import List, Optional, Tuple

import pandas as pd

# --------------------------------------------------------------
# column names as they appear in the excel header row.
# change these if the export format is different.
# --------------------------------------------------------------
COL_COMPLAINT_ID = "Complaint ID"
COL_SLA = "Complaint Resolution Time"
COL_COMPLAINT_DT = "Complaint DateTime"
COL_VENDOR_CLOSE = "Vendor Close DateTime"
COL_RO_CODE = "RO Code"
COL_RO_NAME = "RO Name"
COL_VENDOR_CODE = "Vendor Code"
COL_VENDOR_REMARKS = "Vendor Remarks"

# --------------------------------------------------------------
# sla is stored as text like "48 hours" or "24 hours".
# this pulls the number out.
# --------------------------------------------------------------
SLA_PATTERN = re.compile(r"(\d+)\s*hours?", re.IGNORECASE)

# --------------------------------------------------------------
# if the delay is less than this many hours, we call it "on time"
# even if technically a few minutes late. stops the "+0 Hours"
# nonsense we saw in the vba version.
# --------------------------------------------------------------
ON_TIME_THRESHOLD_HOURS = 1.0


def parse_sla(text: str) -> Optional[float]:
    """
    extracts the number of hours from strings like "48 hours".
    returns None if it can't figure it out.
    """
    if pd.isna(text) or str(text).strip().upper() == "NA":
        return None
    match = SLA_PATTERN.search(str(text))
    if match:
        return float(match.group(1))
    return None


def parse_datetime(val) -> Optional[datetime]:
    """
    tries really hard to turn whatever excel gives us into
    a proper python datetime. handles both the string format
    from the export and the excel serial number format.
    """
    if pd.isna(val) or str(val).strip().upper() in ("NA", "PENDING", ""):
        return None

    # if pandas already parsed it (datetime64), we're good
    if isinstance(val, (datetime, pd.Timestamp)):
        return pd.Timestamp(val).to_pydatetime()

    # try parsing the string format: "30-Apr-26 09:39:01 PM"
    str_val = str(val).strip()
    for fmt in [
        "%d-%b-%y %I:%M:%S %p",
        "%d-%b-%y %I:%M:%S",
        "%d-%b-%Y %I:%M:%S %p",
        "%Y-%m-%d %H:%M:%S",
    ]:
        try:
            return datetime.strptime(str_val, fmt)
        except ValueError:
            continue
    return None


def check_auto_close(remarks: str) -> bool:
    """
    some complaints are auto-closed by the system after 24 hours
    even if the vendor never actually visited. we flag these
    separately so they don't get counted as properly resolved.
    """
    if pd.isna(remarks):
        return False
    return "auto close" in str(remarks).lower() or "auto closed" in str(remarks).lower()


def format_duration(delay_days: float) -> str:
    """
    turns a fractional day difference into a human string
    like "-2 Day 3 Hours" or "+0 Hours".
    """
    if abs(delay_days) < 0.00001:
        return "0 Hours"

    sign = "-" if delay_days < 0 else "+"
    ad = abs(delay_days)
    days = int(ad)
    hours = int((ad - days) * 24 + 0.0001)

    if days >= 1 and hours >= 1:
        return f"{sign}{days} Day {hours} Hours"
    elif days >= 1:
        return f"{sign}{days} Day(s)"
    else:
        return f"{sign}{hours} Hours"


def process_excel(filepath: str) -> Tuple[List[dict], dict]:
    """
    main function. reads the excel, runs the same logic as the
    vba macro, returns a list of records + a summary dict.

    returns (records, summary)
    """
    df = pd.read_excel(filepath, engine="openpyxl")

    records = []
    early = delayed = on_time = pending = 0
    total_penalty = 0.0

    for _, row in df.iterrows():
        cid = row.get(COL_COMPLAINT_ID)
        if pd.isna(cid):
            continue

        sla_text = row.get(COL_SLA, "")
        assign_raw = row.get(COL_COMPLAINT_DT)
        close_raw = row.get(COL_VENDOR_CLOSE)
        ro_code = row.get(COL_RO_CODE)
        ro_name = row.get(COL_RO_NAME)
        vendor = row.get(COL_VENDOR_CODE)
        remarks = row.get(COL_VENDOR_REMARKS, "")

        close_dt = parse_datetime(close_raw)
        assign_dt = parse_datetime(assign_raw)
        sla_hours = parse_sla(sla_text)

        is_pending = close_dt is None
        is_auto = check_auto_close(remarks)

        due_time = None
        delay_days = 0.0
        delay_hours = 0.0
        status = "Pending"
        dur_text = "Pending"
        penalty = 0

        if is_pending:
            # still open — we can still calculate the due time
            # from assign + sla if we have both
            if assign_dt and sla_hours:
                due_time = assign_dt + timedelta(hours=sla_hours)
            status = "Pending"
            pending += 1

        else:
            # closed complaint — full analysis
            if not assign_dt or not sla_hours or sla_hours <= 0:
                # can't calculate without assign time or sla
                continue

            due_time = assign_dt + timedelta(hours=sla_hours)
            delay_days = (close_dt - due_time).total_seconds() / 86400.0
            delay_hours = delay_days * 24

            # build the text representation
            dur_text = format_duration(delay_days)

            # decide status
            if delay_hours < -ON_TIME_THRESHOLD_HOURS:
                status = "Early"
                early += 1
            elif delay_hours > ON_TIME_THRESHOLD_HOURS:
                status = "Delayed"
                delayed += 1
                penalty = int(delay_hours / 24) * 1000
                total_penalty += penalty
            else:
                status = "On Time"
                on_time += 1

        records.append({
            "complaint_id": cid,
            "ro_code": ro_code if not pd.isna(ro_code) else None,
            "ro_name": ro_name if not pd.isna(ro_name) else None,
            "vendor_code": vendor if not pd.isna(vendor) else None,
            "assignment_time": assign_dt,
            "due_time": due_time,
            "close_time": close_dt,
            "duration_text": dur_text,
            "delay_hours": round(delay_hours, 2),
            "status": status,
            "penalty": penalty,
            "sla_hours": sla_hours or 0,
            "is_auto_closed": is_auto,
        })

    summary = {
        "total": len(records),
        "early": early,
        "delayed": delayed,
        "on_time": on_time,
        "pending": pending,
        "total_penalty": total_penalty,
    }

    return records, summary
