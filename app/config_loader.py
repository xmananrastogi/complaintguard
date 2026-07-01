"""
reads settings.yaml and makes values available everywhere.
"""

import os
from typing import Any, Dict

import yaml

_CONFIG: Dict[str, Any] | None = None


def _path() -> str:
    return os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "settings.yaml")


def load() -> Dict[str, Any]:
    global _CONFIG
    if _CONFIG is None:
        with open(_path(), "r") as f:
            _CONFIG = yaml.safe_load(f)
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
