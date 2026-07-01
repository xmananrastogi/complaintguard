"""
core engine — replaces the VBA macro entirely.
takes an excel file, spits out analysed records + aggregations.

column mapping is for the specific IOCL export format
we worked with. if the columns change, just update the
dict at the top.
"""

import math
import re
from collections import Counter, defaultdict
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
COL_ENGINEER = "Engineer Name"
COL_NATURE = "Nature of Complaint"
COL_DU_SERIAL = "DU serial No"
COL_COMP_MODE = "Comp Mode"

# --------------------------------------------------------------
# sla is stored as text like "48 hours" or "24 hours".
# this pulls the number out.
# --------------------------------------------------------------
SLA_PATTERN = re.compile(r"(\d+)\s*hours?", re.IGNORECASE)

# --------------------------------------------------------------
# thresholds matching the vba macro exactly:
#  - any negative diff = early
#  - >= 1 hour late = delayed
#  - everything between = on time
# --------------------------------------------------------------
EARLY_THRESHOLD_DAYS = -0.00001
DELAYED_THRESHOLD_HOURS = 1.0


def parse_sla(text) -> Optional[float]:
    if pd.isna(text) or str(text).strip().upper() == "NA":
        return None
    match = SLA_PATTERN.search(str(text))
    if match:
        return float(match.group(1))
    return None


def parse_datetime(val) -> Optional[datetime]:
    if pd.isna(val) or str(val).strip().upper() in ("NA", "PENDING", ""):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return pd.Timestamp(val).to_pydatetime()
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


def check_auto_close(remarks) -> bool:
    if pd.isna(remarks):
        return False
    return "auto close" in str(remarks).lower() or "auto closed" in str(remarks).lower()


def format_duration(delay_days: float) -> str:
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
        engineer = row.get(COL_ENGINEER, "")
        nature = row.get(COL_NATURE, "")
        du_serial = row.get(COL_DU_SERIAL, "")
        comp_mode = row.get(COL_COMP_MODE, "")

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
            if assign_dt and sla_hours:
                due_time = assign_dt + timedelta(hours=sla_hours)
            status = "Pending"
            pending += 1
        else:
            if not assign_dt or not sla_hours or sla_hours <= 0:
                continue
            due_time = assign_dt + timedelta(hours=sla_hours)
            delay_days = (close_dt - due_time).total_seconds() / 86400.0
            delay_hours = delay_days * 24
            dur_text = format_duration(delay_days)

            if delay_days < EARLY_THRESHOLD_DAYS:
                status = "Early"
                early += 1
            elif delay_hours >= DELAYED_THRESHOLD_HOURS:
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
            "engineer_name": engineer if not pd.isna(engineer) else None,
            "nature": nature if not pd.isna(nature) else None,
            "du_serial": str(du_serial).strip() if not pd.isna(du_serial) else None,
            "comp_mode": comp_mode if not pd.isna(comp_mode) else None,
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


# --------------------------------------------------------------
# aggregation helpers — run after process_excel
# --------------------------------------------------------------

def analyse_natures(records: List[dict]) -> List[dict]:
    """
    groups by nature of complaint, returns stats per type.
    helps spot which issues cause the most delays.
    """
    by_nature = defaultdict(lambda: {"total": 0, "delayed": 0, "early": 0, "penalty": 0, "hours": []})
    for r in records:
        n = r.get("nature") or "Unknown"
        by_nature[n]["total"] += 1
        by_nature[n]["hours"].append(r["delay_hours"])
        by_nature[n]["penalty"] += r["penalty"]
        if r["status"] == "Delayed":
            by_nature[n]["delayed"] += 1
        elif r["status"] == "Early":
            by_nature[n]["early"] += 1

    result = []
    for nature, d in sorted(by_nature.items(), key=lambda x: x[1]["penalty"], reverse=True):
        avg_delay = sum(d["hours"]) / len(d["hours"]) if d["hours"] else 0
        result.append({
            "nature": nature,
            "total": d["total"],
            "delayed": d["delayed"],
            "early": d["early"],
            "delay_rate": round(d["delayed"] / d["total"] * 100, 1) if d["total"] else 0,
            "avg_delay_hours": round(avg_delay, 1),
            "penalty": d["penalty"],
        })
    return result


def analyse_engineers(records: List[dict]) -> List[dict]:
    """
    per-engineer stats. who resolves fast, who has delays.
    """
    by_eng = defaultdict(lambda: {"total": 0, "delayed": 0, "early": 0, "ontime": 0, "penalty": 0, "hours": []})
    for r in records:
        e = r.get("engineer_name")
        if not e or str(e).strip().upper() in ("NA", "NONE", ""):
            continue
        by_eng[e]["total"] += 1
        by_eng[e]["hours"].append(r["delay_hours"])
        by_eng[e]["penalty"] += r["penalty"]
        if r["status"] == "Delayed":
            by_eng[e]["delayed"] += 1
        elif r["status"] == "Early":
            by_eng[e]["early"] += 1
        else:
            by_eng[e]["ontime"] += 1

    result = []
    for eng, d in sorted(by_eng.items(), key=lambda x: x[1]["delayed"], reverse=True):
        avg_d = sum(d["hours"]) / len(d["hours"]) if d["hours"] else 0
        result.append({
            "engineer": eng,
            "total": d["total"],
            "delayed": d["delayed"],
            "early": d["early"],
            "ontime": d["ontime"],
            "compliance_rate": round((d["early"] + d["ontime"]) / d["total"] * 100, 1) if d["total"] else 0,
            "avg_delay_hours": round(avg_d, 1),
            "penalty": d["penalty"],
        })
    return sorted(result, key=lambda x: x["compliance_rate"])


def analyse_du_revisits(records: List[dict], window_days: int = 30) -> List[dict]:
    """
    finds DU serials that appear multiple times within window_days.
    if the same equipment breaks again quickly, the fix was incomplete.
    """
    by_du = defaultdict(list)
    for r in records:
        du = r.get("du_serial")
        if du and r.get("assignment_time"):
            by_du[du].append(r)

    flagged = []
    for du, complaints in by_du.items():
        if len(complaints) < 2:
            continue
        times = [(c["assignment_time"], c) for c in complaints if c["assignment_time"]]
        times.sort(key=lambda x: x[0])
        revisits = []
        for i in range(1, len(times)):
            gap = (times[i][0] - times[i - 1][0]).days
            if gap <= window_days:
                revisits.append({
                    "prev_id": times[i - 1][1]["complaint_id"],
                    "next_id": times[i][1]["complaint_id"],
                    "gap_days": gap,
                    "prev_status": times[i - 1][1]["status"],
                    "next_status": times[i][1]["status"],
                    "prev_close": times[i - 1][1]["close_time"],
                })
        if revisits:
            flagged.append({
                "du_serial": du,
                "total_complaints": len(complaints),
                "revisit_count": len(revisits),
                "revisits": revisits,
                "vendor": complaints[0].get("vendor_code"),
                "ro_name": complaints[0].get("ro_name"),
            })

    return sorted(flagged, key=lambda x: x["revisit_count"], reverse=True)


def analyse_modes(records: List[dict]) -> List[dict]:
    """
    compares WEB vs SYSTEM logged complaints.
    """
    by_mode = defaultdict(lambda: {"total": 0, "delayed": 0, "penalty": 0})
    for r in records:
        m = r.get("comp_mode") or "Unknown"
        by_mode[m]["total"] += 1
        by_mode[m]["penalty"] += r["penalty"]
        if r["status"] == "Delayed":
            by_mode[m]["delayed"] += 1

    result = []
    for mode, d in sorted(by_mode.items(), key=lambda x: x[1]["total"], reverse=True):
        result.append({
            "mode": mode,
            "total": d["total"],
            "delayed": d["delayed"],
            "delay_rate": round(d["delayed"] / d["total"] * 100, 1) if d["total"] else 0,
            "penalty": d["penalty"],
        })
    return result


def generate_report_html(records: List[dict], summary: dict) -> str:
    """
    builds a clean html report string for pdf / print.
    """
    natures = analyse_natures(records)
    engineers = analyse_engineers(records)
    revisits = analyse_du_revisits(records)
    modes = analyse_modes(records)

    lines = []
    lines.append("<h1>ComplaintGuard — SLA Compliance Report</h1>")
    lines.append(f"<p>Generated: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}</p>")

    lines.append("<h2>Summary</h2>")
    lines.append(f"<p>Total: {summary['total']} | Early: {summary['early']} | "
                 f"Delayed: {summary['delayed']} | On Time: {summary['on_time']} | "
                 f"Pending: {summary['pending']} | Total Penalty: ₹{summary['total_penalty']:,.0f}</p>")

    lines.append("<h2>Top Complaint Types</h2><table border=1 cellpadding=4>")
    lines.append("<tr><th>Nature</th><th>Total</th><th>Delayed</th><th>Delay%</th><th>Avg Delay</th><th>Penalty</th></tr>")
    for n in natures[:10]:
        lines.append(f"<tr><td>{n['nature']}</td><td>{n['total']}</td><td>{n['delayed']}</td>"
                     f"<td>{n['delay_rate']}%</td><td>{n['avg_delay_hours']}h</td><td>₹{n['penalty']:,}</td></tr>")
    lines.append("</table>")

    lines.append("<h2>Bottom Engineers</h2><table border=1 cellpadding=4>")
    lines.append("<tr><th>Engineer</th><th>Total</th><th>Delayed</th><th>Compliance%</th><th>Penalty</th></tr>")
    for e in engineers[:10]:
        lines.append(f"<tr><td>{e['engineer']}</td><td>{e['total']}</td><td>{e['delayed']}</td>"
                     f"<td>{e['compliance_rate']}%</td><td>₹{e['penalty']:,}</td></tr>")
    lines.append("</table>")

    if revisits:
        lines.append("<h2>DU Revisits (flag)</h2><table border=1 cellpadding=4>")
        lines.append("<tr><th>DU Serial</th><th>Complaints</th><th>Revisits</th><th>Vendor</th></tr>")
        for rv in revisits[:10]:
            lines.append(f"<tr><td>{rv['du_serial']}</td><td>{rv['total_complaints']}</td>"
                         f"<td>{rv['revisit_count']}</td><td>{rv['vendor']}</td></tr>")
        lines.append("</table>")

    return "\n".join(lines)
