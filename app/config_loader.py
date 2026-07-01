"""
reads settings.yaml and makes values available everywhere.
falls back to hardcoded defaults if the file is missing or broken.
"""

import os
from typing import Any, Dict

try:
    import yaml
except ImportError:
    yaml = None

_DEFAULTS = {
    "penalty": {"per_block": 1000, "currency": "₹"},
    "revisit": {"window_days": 30},
    "columns": {
        "complaint_id": ["Complaint ID", "Ticket ID", "Complaint No", "ID", "Ticket Number", "Issue ID"],
        "sla": ["Complaint Resolution Time", "SLA", "Resolution Time", "SLA Hours", "Response Time", "SLA (Hours)", "TAT"],
        "complaint_dt": ["Complaint DateTime", "Complaint Date", "Logged Date", "Created Date", "Complaint Date Time", "Ticket Date"],
        "close_dt": ["Vendor Close DateTime", "Close DateTime", "Closed Date", "Resolution Date", "Vendor Closed Date", "Close Date"],
        "ro_code": ["RO Code", "Dealer Code", "Retail Outlet Code", "Outlet Code"],
        "ro_name": ["RO Name", "Dealer Name", "Retail Outlet Name", "Outlet Name"],
        "vendor": ["Vendor Code", "Vendor", "Vendor Name", "Supplier"],
        "vendor_remarks": ["Vendor Remarks", "Remarks", "Closure Remarks", "Work Done", "Technician Remarks"],
        "engineer": ["Engineer Name", "Engineer", "Technician", "Assigned To", "Engineer Assigned"],
        "nature": ["Nature of Complaint", "Nature", "Complaint Type", "Issue Type", "Problem Category"],
        "du_serial": ["DU serial No", "DU Serial", "Serial No", "Equipment Serial", "DU Number", "Device ID"],
        "comp_mode": ["Comp Mode", "Complaint Mode", "Mode", "Source", "Channel"],
    },
}

_CONFIG: Dict[str, Any] | None = None


def _find_file() -> str | None:
    """walk up from this file to find config/settings.yaml."""
    if yaml is None:
        return None
    here = os.path.dirname(os.path.abspath(__file__))
    # try: app/../config/settings.yaml (normal layout)
    candidate = os.path.join(here, os.pardir, "config", "settings.yaml")
    if os.path.isfile(candidate):
        return candidate
    # try: cwd/config/settings.yaml (streamlit cloud might set cwd to repo root)
    candidate = os.path.join(os.getcwd(), "config", "settings.yaml")
    if os.path.isfile(candidate):
        return candidate
    return None


def load() -> Dict[str, Any]:
    global _CONFIG
    if _CONFIG is not None:
        return _CONFIG
    path = _find_file()
    if path and yaml:
        try:
            with open(path, "r") as f:
                _CONFIG = yaml.safe_load(f)
            return _CONFIG
        except Exception:
            pass
    _CONFIG = _DEFAULTS
    return _CONFIG


def get(key: str, default: Any = None) -> Any:
    return load().get(key, default)


def penalty_rate() -> int:
    return get("penalty", {}).get("per_block", 1000)


def revisit_window() -> int:
    return get("revisit", {}).get("window_days", 30)


def column_aliases() -> Dict[str, list]:
    return get("columns", {})


def reload():
    """force re-read on next access (call after editing the file)."""
    global _CONFIG
    _CONFIG = None
