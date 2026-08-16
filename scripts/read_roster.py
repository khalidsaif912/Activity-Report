import json
from datetime import datetime
from pathlib import Path
from typing import Iterable

import requests
from bs4 import BeautifulSoup

from manpower_layout import ALL_MANPOWER_TITLES, finalize_manpower_sections
from report_date_config import get_report_date
from roster_schedules import employee_directory, staff_rows_for_date

SHIFTS = {
    "morning": {"start": "06:00", "end": "15:00", "code": "MN06"},
    "afternoon": {"start": "13:00", "end": "22:00", "code": "AN13"},
    "night": {"start": "21:00", "end": "06:00", "code": "NN21"},
}

SHIFT_FILTER_CODES: dict[str, tuple[str, ...]] = {
    "morning": ("MN06",),
    "afternoon": ("AN13",),
    "night": ("NN21",),
}

EXPORT_DEPTS = {
    "Supervisors",
    "Load Control",
    "Export Checker",
    "Export Operators",
}


def html_urls_for_date(date: str) -> list[str]:
    return [
        f"https://khalidsaif912.github.io/new/docs/date/{date}/",
        f"https://khalidsaif912.github.io/new/docs/date/{date}/index.html",
        f"https://raw.githubusercontent.com/khalidsaif912/new/main/docs/date/{date}/index.html",
        f"https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/date/{date}/index.html",
        f"https://khalidsaif912.github.io/roster-site/date/{date}/",
        f"https://khalidsaif912.github.io/roster-site/date/{date}/index.html",
    ]


def get_current_shift() -> str:
    now = datetime.now().time()

    def in_range(start_str: str, end_str: str) -> bool:
        start = datetime.strptime(start_str, "%H:%M").time()
        end = datetime.strptime(end_str, "%H:%M").time()
        if start < end:
            return start <= now <= end
        return now >= start or now <= end

    for shift_name, cfg in SHIFTS.items():
        if in_range(cfg["start"], cfg["end"]):
            return shift_name

    return "morning"


def fetch_html_any(urls: list[str]) -> str:
    last_err: Exception | None = None
    for url in urls:
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            text = response.text or ""
            if "<html" in text.lower() or "deptCard" in text or "empName" in text:
                print(f"Fetched roster HTML: {url}")
                return text
        except Exception as exc:  # noqa: BLE001 — try next mirror
            last_err = exc
    if last_err:
        print(f"Roster HTML not available ({last_err})")
    return ""


def _dept_name_from_card(dept_card) -> str:
    raw = (dept_card.get("data-dept") or "").strip()
    if raw:
        return raw
    title_tag = dept_card.find(class_="deptTitle") or dept_card.find(class_="dept-title")
    return title_tag.get_text(" ", strip=True) if title_tag else ""


def parse_roster(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    staff = []

    dept_cards = soup.find_all("div", class_="deptCard")
    if not dept_cards:
        dept_cards = soup.find_all("div", class_="dept-card")

    for dept_card in dept_cards:
        dept_name = _dept_name_from_card(dept_card)
        if dept_name not in EXPORT_DEPTS:
            continue

        shift_cards = dept_card.find_all("details", class_="shiftCard")
        if not shift_cards:
            shift_cards = [dept_card]

        for shift_card in shift_cards:
            employees = shift_card.find_all("div", class_="empRow")
            if not employees:
                employees = shift_card.find_all("div", class_="emp-row")

            for emp in employees:
                name_tag = emp.find("span", class_="empName") or emp.find("span", class_="emp-name")
                status_tag = emp.find("span", class_="empStatus") or emp.find("span", class_="emp-code")

                if not name_tag or not status_tag:
                    continue

                full_text = name_tag.get_text(" ", strip=True)
                shift_code = status_tag.get_text(" ", strip=True)

                if " - " in full_text:
                    name, sn = full_text.rsplit(" - ", 1)
                else:
                    name, sn = full_text, ""

                sn = sn.strip()
                if sn and not sn.startswith("SN"):
                    sn = f"SN{sn}"

                staff.append(
                    {
                        "sn": sn,
                        "name": name.strip(),
                        "department": dept_name,
                        "shift": shift_code.strip(),
                    }
                )

    return staff


def filter_by_shift(staff: list[dict], shift_name: str) -> list[dict]:
    codes = SHIFT_FILTER_CODES[shift_name]
    keyed = [row for row in staff if row.get("shift_key") == shift_name]
    if keyed:
        return keyed
    return [row for row in staff if any(c in row["shift"] for c in codes)]


def shift_meta(shift_name: str, date: str) -> dict:
    return {
        "key": shift_name,
        "title": f"{shift_name.capitalize()} Shift",
        "date": date,
        "time": f'{SHIFTS[shift_name]["start"]} - {SHIFTS[shift_name]["end"]}',
    }


def build_manpower_sections(staff: list[dict]) -> list[dict]:
    grouped = {name: [] for name in ALL_MANPOWER_TITLES}

    department_map = {
        "Supervisors": "Supervisor",
        "Load Control": "Load Control",
        "Export Checker": "Export Checker",
        "Export Operators": "Export Operators",
        "Officers": "Officers",
    }

    for row in staff:
        dept = department_map.get(row["department"])
        if dept:
            grouped.setdefault(dept, []).append(f'{row["sn"]} {row["name"]}'.strip())

    return finalize_manpower_sections(grouped, [])


def load_existing_employees(path: Path) -> list[str]:
    if not path.is_file():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(raw, list):
        return []
    return [str(x).strip() for x in raw if str(x).strip()]


def export_employee_directory(
    directory: list[str],
    shift_payloads: Iterable[dict],
    output_dir: Path,
) -> None:
    employees: list[str] = []
    seen: set[str] = set()

    def add(value: str) -> None:
        line = str(value or "").strip()
        key = line.upper()
        if not line or key in seen:
            return
        seen.add(key)
        employees.append(line)

    for item in directory:
        add(item)
    for payload in shift_payloads:
        for section in payload.get("manpowerSections", []):
            for item in section.get("items", []):
                add(item)

    employees_file = output_dir / "employees.json"
    if not employees:
        existing = load_existing_employees(employees_file)
        for item in existing:
            add(item)

    with employees_file.open("w", encoding="utf-8") as f:
        json.dump(employees, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"employees.json count={len(employees)}")


def load_staff(date: str) -> tuple[list[dict], list[str]]:
    directory: list[str] = []
    staff: list[dict] = []

    try:
        directory = employee_directory()
        staff = staff_rows_for_date(date)
        print(f"Schedules JSON: directory={len(directory)} on-duty={len(staff)}")
    except Exception as exc:  # noqa: BLE001 — HTML fallback still available
        print(f"Schedules JSON unavailable ({exc}); falling back to HTML")

    if staff:
        return staff, directory

    print("Fetching roster HTML...")
    html = fetch_html_any(html_urls_for_date(date))
    if not html:
        return staff, directory
    parsed = parse_roster(html)
    print(f"HTML roster parsed staff={len(parsed)}")
    return parsed, directory


def main() -> None:
    date = get_report_date()
    print(f"Report date: {date}")
    staff, directory = load_staff(date)

    shift_name = get_current_shift()
    print(f"Detected shift: {shift_name}")

    shifts: dict[str, dict] = {}
    for sn in ("morning", "afternoon", "night"):
        filtered = filter_by_shift(staff, sn)
        shifts[sn] = {
            "shiftMeta": shift_meta(sn, date),
            "manpowerSections": build_manpower_sections(filtered),
        }

    current = shifts[shift_name]
    data = {
        "shiftMeta": current["shiftMeta"],
        "manpowerSections": current["manpowerSections"],
        "shifts": shifts,
        "defaultShift": shift_name,
    }

    base_dir = Path(__file__).resolve().parent.parent

    roster_dir = base_dir / "data" / "roster"
    roster_dir.mkdir(parents=True, exist_ok=True)

    report_dir = base_dir / "data" / "report"
    report_dir.mkdir(parents=True, exist_ok=True)

    roster_file = roster_dir / "latest.json"
    with roster_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    by_date_roster = roster_dir / "by-date" / date
    by_date_roster.mkdir(parents=True, exist_ok=True)
    roster_dated_file = by_date_roster / "latest.json"
    with roster_dated_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    export_employee_directory(directory, shifts.values(), report_dir)

    print(f"Done [ok] {roster_file} created")
    print(f"Done [ok] {roster_dated_file} created")
    print(f"Done [ok] {report_dir / 'employees.json'} created")


if __name__ == "__main__":
    main()
