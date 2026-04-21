"""Single place for the report roster/dispatch date (YYYY-MM-DD)."""

from __future__ import annotations

import os
import re
from datetime import datetime

_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def get_report_date() -> str:
    today = datetime.now().strftime("%Y-%m-%d")
    raw = (os.environ.get("ACTIVITY_REPORT_DATE") or today).strip()
    return raw if _RE.match(raw) else today
