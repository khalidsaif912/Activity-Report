import json
from datetime import datetime
from pathlib import Path
from typing import Iterable

import requests
from bs4 import BeautifulSoup

from manpower_layout import ALL_MANPOWER_TITLES, finalize_manpower_sections
from report_date_config import get_report_date


def raw_url_for_date(date: str) -> str:
    return (
        "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/date/"
        f"{date}/index.html"
    )

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


def fetch_html(url: str) -> str:
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return response.text


def parse_roster(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    staff = []

    dept_cards = soup.find_all("div", class_="deptCard")

    for dept_card in dept_cards:
        dept_name = dept_card.get("data-dept", "").strip()
        if dept_name not in {
            "Supervisors",
            "Load Control",
            "Export Checker",
            "Export Operators",
        }:
            continue

        shift_cards = dept_card.find_all("details", class_="shiftCard")

        for shift_card in shift_cards:
            employees = shift_card.find_all("div", class_="empRow")

            for emp in employees:
                name_tag = emp.find("span", class_="empName")
                status_tag = emp.find("span", class_="empStatus")

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
    }

    for row in staff:
        dept = department_map.get(row["department"])
        if dept:
            grouped[dept].append(f'{row["sn"]} {row["name"]}'.strip())

    return finalize_manpower_sections(grouped, [])


def main() -> None:
    date = get_report_date()
    url = raw_url_for_date(date)
    print(f"Report date: {date}")
    print("Fetching roster...")
    html = fetch_html(url)

    print("Parsing...")
    staff = parse_roster(html)

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

    export_employee_directory_union(shifts.values(), report_dir)

    print(f"Done [ok] {roster_file} created")
    print(f"Done [ok] {roster_dated_file} created")
    print(f"Done [ok] {report_dir / 'employees.json'} created")


def export_employee_directory_union(
    shift_payloads: Iterable[dict],
    output_dir: Path,
) -> None:
    employees: list[str] = []
    for payload in shift_payloads:
        for section in payload.get("manpowerSections", []):
            for item in section.get("items", []):
                value = item.strip()
                if value and value not in employees:
                    employees.append(value)

    employees_file = output_dir / "employees.json"
    with employees_file.open("w", encoding="utf-8") as f:
        json.dump(employees, f, ensure_ascii=False, indent=2)






if __name__ == "__main__":
    main()