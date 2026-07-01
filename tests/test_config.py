import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from app.config_loader import penalty_rate, revisit_window, column_aliases, reload


def test_defaults_match():
    """config values should match the shipped settings.yaml."""
    reload()
    assert penalty_rate() == 1000
    assert revisit_window() == 30
    aliases = column_aliases()
    assert "Complaint ID" in aliases["complaint_id"]
    assert "48 hours" not in str(aliases)  # just check it's loading from yaml, not hardcoded

    print("config tests passed!")
