"""Read the live roster-site schedules JSON (source of truth for names).

HTML date pages change markup over time; this JSON is what the homepage
already uses as a fallback and stays stable.
"""

from __future__ import annotations

import json
from typing import Any
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.request import Request, urlopen

SCHEDULES_INDEX_URLS = (
    "https://khalidsaif912.github.io/roster-site/schedules/index.json",
    "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/schedules/index.json",
)
SCHEDULE_EMPLOYEE_URLS = (
    "https://khalidsaif912.github.io/roster-site/schedules/{emp_id}.json",
    "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/schedules/{emp_id}.json",
)

DUTY_SHIFT_GROUPS = {
    "morning": "morning",
    "afternoon": "afternoon",
    "night": "night",
}

EXPORT_DEPTS = {
    "Supervisors",
    "Load Control",
    "Export Checker",
    "Export Operators",
}


def _fetch_json(url: str, timeout: int = 30) -> Any:
    req = Request(url, headers={"User-Agent": "Activity-Report/1.0"})
    with urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8")
    return json.loads(raw)


def fetch_json_any(urls: list[str] | tuple[str, ...]) -> Any:
    last_err: Exception | None = None
    for url in urls:
        try:
            return _fetch_json(url)
        except Exception as exc:  # noqa: BLE001 — try next mirror
            last_err = exc
    if last_err:
        raise last_err
    raise RuntimeError("No URLs to fetch")


def load_schedules_index() -> dict:
    data = fetch_json_any(SCHEDULES_INDEX_URLS)
    if not isinstance(data, dict):
        return {"employees": []}
    if not isinstance(data.get("employees"), list):
        data["employees"] = []
    return data


def load_employee_schedule(emp_id: str) -> dict | None:
    emp_id = str(emp_id or "").strip()
    if not emp_id:
        return None
    urls = [u.format(emp_id=emp_id) for u in SCHEDULE_EMPLOYEE_URLS]
    try:
        data = fetch_json_any(urls)
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def format_employee_name(emp_id: str, name: str) -> str:
    sn = str(emp_id or "").strip()
    label = str(name or "").strip()
    if sn and not sn.upper().startswith("SN"):
        sn = f"SN{sn}"
    return f"{sn} {label}".strip()


def employee_directory(index: dict | None = None) -> list[str]:
    """Full autocomplete list: SN{id} Name from schedules/index.json."""
    idx = index if isinstance(index, dict) else load_schedules_index()
    names: list[str] = []
    seen: set[str] = set()
    for emp in idx.get("employees") or []:
        if not isinstance(emp, dict):
            continue
        line = format_employee_name(emp.get("id", ""), emp.get("name", ""))
        key = line.upper()
        if not line or key in seen:
            continue
        seen.add(key)
        names.append(line)
    return names


def dept_section_title(raw: str) -> str:
    d = str(raw or "").strip()
    if d.lower() == "supervisors":
        return "Supervisor"
    return d or "Other"


def staff_rows_for_date(iso_date: str, index: dict | None = None) -> list[dict]:
    """On-duty export-warehouse staff for one calendar date."""
    date = str(iso_date or "").strip()
    if not date:
        return []
    month_key = date[:7]
    idx = index if isinstance(index, dict) else load_schedules_index()
    emps = [
        e
        for e in (idx.get("employees") or [])
        if isinstance(e, dict)
        and str(e.get("department") or "").strip() in EXPORT_DEPTS
        and month_key in (e.get("months") or [])
    ]
    rows: list[dict] = []
    with ThreadPoolExecutor(max_workers=12) as pool:
        future_map = {
            pool.submit(load_employee_schedule, str(emp.get("id") or "")): emp for emp in emps
        }
        for fut in as_completed(future_map):
            emp = future_map[fut]
            try:
                schedule = fut.result()
            except Exception:
                continue
            month_rows = []
            if schedule and isinstance(schedule.get("schedules"), dict):
                month_rows = schedule["schedules"].get(month_key) or []
            if not isinstance(month_rows, list):
                continue
            row = next(
                (r for r in month_rows if isinstance(r, dict) and str(r.get("date") or "") == date),
                None,
            )
            if not row:
                continue
            grp = str(row.get("shift_group") or "").strip().lower()
            shift_key = DUTY_SHIFT_GROUPS.get(grp)
            if not shift_key:
                continue
            rows.append(
                {
                    "sn": f"SN{str(emp.get('id') or '').strip()}",
                    "name": str(emp.get("name") or "").strip(),
                    "department": str(emp.get("department") or "").strip(),
                    "shift": str(row.get("shift_code") or "").strip(),
                    "shift_key": shift_key,
                }
            )
    return rows
