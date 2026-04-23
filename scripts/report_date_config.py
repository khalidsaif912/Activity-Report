"""Single place for the report roster/dispatch date (YYYY-MM-DD)."""

from __future__ import annotations

import os
import re
from datetime import datetime, timedelta

_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _default_report_date() -> str:
    """
    Operational date for report generation.
    Night shift runs 21:00-06:00, so 00:00-05:59 belongs to previous date.
    """
    now = datetime.now()
    if now.hour < 6:
        now = now - timedelta(days=1)
    return now.strftime("%Y-%m-%d")


def get_report_date() -> str:
    default_date = _default_report_date()
    raw = (os.environ.get("ACTIVITY_REPORT_DATE") or default_date).strip()
    return raw if _RE.match(raw) else default_date
