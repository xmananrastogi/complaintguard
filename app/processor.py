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

from app.config_loader import column_aliases, revisit_window as get_revisit_window

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
# column aliases — what the user's file might call each column.
# the detector tries each alias in order until it finds a match.
# add more aliases here if new export formats show up.
# --------------------------------------------------------------
COLUMN_ALIASES = column_aliases()


def detect_columns(df: pd.DataFrame) -> dict:
    """
    figures out which column in the uploaded file maps to which field
    we need. tries aliases in order, case-insensitive partial match.

    returns a dict like {"complaint_id": "Complaint ID", "sla": "SLA", ...}
    any field that couldnt be matched gets None.
    """
    actual_cols = [str(c).strip().lower() for c in df.columns]
    actual_map = {str(c).strip().lower(): str(c).strip() for c in df.columns}

    result = {}
    for field, aliases in COLUMN_ALIASES.items():
        matched = None
        for alias in aliases:
            alias_lower = alias.strip().lower()
            # exact match
            if alias_lower in actual_map:
                matched = actual_map[alias_lower]
                break
            # partial match — alias is a substring of column name
            for actual_lower, actual_orig in actual_map.items():
                if alias_lower in actual_lower or actual_lower in alias_lower:
                    matched = actual_orig
                    break
            if matched:
                break
        result[field] = matched
    return result


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


def process_excel(filepath: str, penalty_per_block: float = 1000.0) -> Tuple[List[dict], dict, dict]:
    """
    main function. reads the excel, runs the same logic as the
    vba macro, returns (records, summary, column_mapping).

    penalty_per_block: ₹ charged per full 24h delay block (default ₹1,000).
    column_mapping shows which columns were matched from the file.
    """
    df = pd.read_excel(filepath, engine="openpyxl")

    # detect columns
    col_map = detect_columns(df)
    missing = [k for k, v in col_map.items() if v is None]

    cid_col = col_map.get("complaint_id")
    sla_col = col_map.get("sla")
    dt_col = col_map.get("complaint_dt")
    close_col = col_map.get("close_dt")
    ro_code_col = col_map.get("ro_code")
    ro_name_col = col_map.get("ro_name")
    vendor_col = col_map.get("vendor")
    remarks_col = col_map.get("vendor_remarks")
    eng_col = col_map.get("engineer")
    nature_col = col_map.get("nature")
    du_col = col_map.get("du_serial")
    mode_col = col_map.get("comp_mode")

    # if complaint id is missing, nothing we can do
    if not cid_col:
        return [], {"error": "Could not find Complaint ID column. Check file headers."}, col_map

    records = []
    early = delayed = on_time = pending = 0
    total_penalty = 0.0

    for _, row in df.iterrows():
        cid = row.get(cid_col) if cid_col else None
        if pd.isna(cid):
            continue

        sla_text = row.get(sla_col, "") if sla_col else ""
        assign_raw = row.get(dt_col) if dt_col else None
        close_raw = row.get(close_col) if close_col else None
        ro_code = row.get(ro_code_col) if ro_code_col else None
        ro_name = row.get(ro_name_col) if ro_name_col else None
        vendor = row.get(vendor_col) if vendor_col else None
        remarks = row.get(remarks_col, "") if remarks_col else ""
        engineer = row.get(eng_col, "") if eng_col else ""
        nature = row.get(nature_col, "") if nature_col else ""
        du_serial = row.get(du_col, "") if du_col else ""
        comp_mode = row.get(mode_col, "") if mode_col else ""

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
                penalty = int(delay_hours / 24) * penalty_per_block
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

    return records, summary, col_map


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
    
    window_days defaults to config/settings.yaml → revisit.window_days.
    """
    if window_days is None:
        window_days = get_revisit_window()
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


def analyse_vendors(records: List[dict]) -> List[dict]:
    """
    per-vendor stats — penalty, delays, complaint volume.
    """
    by_v = defaultdict(lambda: {"total": 0, "delayed": 0, "early": 0, "ontime": 0, "pending": 0, "penalty": 0})
    for r in records:
        v = r.get("vendor_code")
        if not v or str(v).strip().upper() in ("NA", "NONE", ""):
            continue
        by_v[v]["total"] += 1
        by_v[v]["penalty"] += r["penalty"]
        if r["status"] == "Delayed":
            by_v[v]["delayed"] += 1
        elif r["status"] == "Early":
            by_v[v]["early"] += 1
        elif r["status"] == "On Time":
            by_v[v]["ontime"] += 1
        else:
            by_v[v]["pending"] += 1

    result = []
    for v, d in sorted(by_v.items(), key=lambda x: x[1]["penalty"], reverse=True):
        result.append({
            "vendor": v,
            "total": d["total"],
            "delayed": d["delayed"],
            "early": d["early"],
            "ontime": d["ontime"],
            "pending": d["pending"],
            "compliance_rate": round((d["early"] + d["ontime"]) / d["total"] * 100, 1) if d["total"] else 0,
            "penalty": d["penalty"],
        })
    return result


def analyse_ros(records: List[dict]) -> List[dict]:
    """
    per-ro stats — which retail outlets have the most issues.
    """
    by_ro = defaultdict(lambda: {"ro_name": "", "total": 0, "delayed": 0, "penalty": 0})
    for r in records:
        ro = r.get("ro_code")
        if not ro or str(ro).strip().upper() in ("NA", "NONE", ""):
            continue
        by_ro[ro]["ro_name"] = r.get("ro_name", "")
        by_ro[ro]["total"] += 1
        by_ro[ro]["penalty"] += r["penalty"]
        if r["status"] == "Delayed":
            by_ro[ro]["delayed"] += 1

    result = []
    for ro, d in sorted(by_ro.items(), key=lambda x: x[1]["penalty"], reverse=True):
        result.append({
            "ro_code": ro,
            "ro_name": d["ro_name"],
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
    revisits = analyse_du_revisits(records, window_days=None)  # None = read from config
    modes = analyse_modes(records)
    vendors = analyse_vendors(records)
    ros = analyse_ros(records)

    lines = []
    lines.append("""<!DOCTYPE html><html><head><meta charset='utf-8'>
<style>
body { font-family: Arial, sans-serif; background: #fff; color: #222; padding: 30px; color-scheme: only light; }
h1 { color: #1a5276; border-bottom: 2px solid #1a5276; padding-bottom: 8px; }
h2 { color: #2c3e50; margin-top: 30px; }
table { border-collapse: collapse; width: 100%; margin: 10px 0 20px 0; }
th { background: #1a5276; color: #fff; padding: 10px; text-align: left; font-size: 14px; }
td { padding: 8px 10px; border-bottom: 1px solid #ddd; font-size: 13px; }
tr:nth-child(even) { background: #f8f9fa; }
tr:hover { background: #eaf2f8; }
.summary { background: #eaf2f8; padding: 15px; border-radius: 6px; font-size: 15px; line-height: 1.6; }
.status-early { color: #155724; background: #d4edda; padding: 2px 8px; border-radius: 4px; }
.status-delayed { color: #721c24; background: #f8d7da; padding: 2px 8px; border-radius: 4px; }
.status-pending { color: #856404; background: #fff3cd; padding: 2px 8px; border-radius: 4px; }
.footer { margin-top: 40px; font-size: 12px; color: #888; border-top: 1px solid #ddd; padding-top: 10px; }
</style></head><body>""")

    lines.append("<h1>ComplaintGuard SLA Compliance Report</h1>")
    lines.append(f"<p>Generated: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}</p>")

    lines.append("<div class='summary'>")
    lines.append(f"<b>Total:</b> {summary['total']} &nbsp;|&nbsp; "
                 f"<b>Early:</b> {summary['early']} &nbsp;|&nbsp; "
                 f"<b>Delayed:</b> {summary['delayed']} &nbsp;|&nbsp; "
                 f"<b>On Time:</b> {summary['on_time']} &nbsp;|&nbsp; "
                 f"<b>Pending:</b> {summary['pending']} &nbsp;|&nbsp; "
                 f"<b>Total Penalty:</b> ₹{summary['total_penalty']:,.0f}")
    lines.append("</div>")

    lines.append("<h2>Top Complaint Types</h2><table>")
    lines.append("<tr><th>Nature</th><th>Total</th><th>Delayed</th><th>Delay%</th><th>Avg Delay</th><th>Penalty</th></tr>")
    for n in natures[:10]:
        lines.append(f"<tr><td>{n['nature']}</td><td>{n['total']}</td><td>{n['delayed']}</td>"
                     f"<td>{n['delay_rate']}%</td><td>{n['avg_delay_hours']}h</td><td>₹{n['penalty']:,}</td></tr>")
    lines.append("</table>")

    lines.append("<h2>Engineer Performance (worst first)</h2><table>")
    lines.append("<tr><th>Engineer</th><th>Total</th><th>Delayed</th><th>Compliance%</th><th>Avg Delay</th><th>Penalty</th></tr>")
    for e in engineers[:15]:
        lines.append(f"<tr><td>{e['engineer']}</td><td>{e['total']}</td><td>{e['delayed']}</td>"
                     f"<td>{e['compliance_rate']}%</td><td>{e['avg_delay_hours']}h</td><td>₹{e['penalty']:,}</td></tr>")
    lines.append("</table>")

    if revisits:
        lines.append("<h2>DU Revisits (flagged)</h2><table>")
        lines.append("<tr><th>DU Serial</th><th>Complaints</th><th>Revisits</th><th>Vendor</th><th>RO</th></tr>")
        for rv in revisits[:15]:
            lines.append(f"<tr><td>{rv['du_serial']}</td><td>{rv['total_complaints']}</td>"
                         f"<td>{rv['revisit_count']}</td><td>{rv.get('vendor', '')}</td><td>{rv.get('ro_name', '')}</td></tr>")
        lines.append("</table>")

    lines.append("<h2>Complaint Mode</h2><table>")
    lines.append("<tr><th>Mode</th><th>Total</th><th>Delayed</th><th>Delay%</th><th>Penalty</th></tr>")
    for m in modes:
        lines.append(f"<tr><td>{m['mode']}</td><td>{m['total']}</td><td>{m['delayed']}</td>"
                     f"<td>{m['delay_rate']}%</td><td>₹{m['penalty']:,}</td></tr>")
    lines.append("</table>")

    lines.append("<h2>Vendor Performance</h2><table>")
    lines.append("<tr><th>Vendor</th><th>Total</th><th>Delayed</th><th>Compliance%</th><th>Penalty</th></tr>")
    for v in vendors[:20]:
        lines.append(f"<tr><td>{v['vendor']}</td><td>{v['total']}</td><td>{v['delayed']}</td>"
                     f"<td>{v['compliance_rate']}%</td><td>₹{v['penalty']:,}</td></tr>")
    lines.append("</table>")

    lines.append("<h2>Retail Outlet Issues</h2><table>")
    lines.append("<tr><th>RO Code</th><th>RO Name</th><th>Total</th><th>Delayed</th><th>Delay%</th><th>Penalty</th></tr>")
    for r in ros[:20]:
        lines.append(f"<tr><td>{r['ro_code']}</td><td>{r['ro_name']}</td><td>{r['total']}</td>"
                     f"<td>{r['delayed']}</td><td>{r['delay_rate']}%</td><td>₹{r['penalty']:,}</td></tr>")
    lines.append("</table>")

    lines.append("<div class='footer'>Generated by ComplaintGuard • IOCL SLA Compliance Tool</div>")
    lines.append("</body></html>")

    return "\n".join(lines)
