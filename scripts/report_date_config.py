"""Single place for the report roster/dispatch date (YYYY-MM-DD)."""

from __future__ import annotations

import os
import re

_DEFAULT = "2026-04-20"
_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def get_report_date() -> str:
    raw = (os.environ.get("ACTIVITY_REPORT_DATE") or _DEFAULT).strip()
    return raw if _RE.match(raw) else _DEFAULT
