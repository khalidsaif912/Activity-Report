import json
from pathlib import Path

import requests
from bs4 import BeautifulSoup


DATE = "2026-04-18"

EXPORT_ROSTER_URL = f"https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/date/{DATE}/index.html"
IMPORT_DISPATCH_URL = f"https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/import/{DATE}/index.html"


def fetch_html(url: str) -> str:
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return response.text


def parse_export_roster_all(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    employees = []

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

        for emp in dept_card.find_all("div", class_="empRow"):
            name_tag = emp.find("span", class_="empName")
            if not name_tag:
                continue

            full_text = name_tag.get_text(" ", strip=True)

            if " - " in full_text:
                name, sn = full_text.rsplit(" - ", 1)
                sn = sn.strip()
                if sn and not sn.startswith("SN"):
                    sn = f"SN{sn}"
                employees.append(f"{sn} {name.strip()}".strip())
            else:
                employees.append(full_text.strip())

    return employees


def parse_flight_dispatch_all(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    employees = []

    for card in soup.find_all("div", class_="dept-card"):
        title_tag = card.find("div", class_="dept-title")
        if not title_tag:
            continue

        title = title_tag.get_text(" ", strip=True)
        if title not in {"Flight Dispatch (Export)", "Flight Dispatch (Import)"}:
            continue

        for row in card.find_all("div", class_="emp-row"):
            name_tag = row.find("span", class_="emp-name")
            if not name_tag:
                continue

            text = name_tag.get_text(" ", strip=True)

            if "·" in text:
                name, sn = text.split("·", 1)
                name = name.strip()
                sn = sn.strip()
                employees.append(f"SN{sn} {name}")
            else:
                employees.append(text.strip())

    return employees


def unique_keep_order(items: list[str]) -> list[str]:
    seen = set()
    result = []

    for item in items:
        value = item.strip()
        if value and value not in seen:
            seen.add(value)
            result.append(value)

    return result


def main() -> None:
    print("Fetching export roster...")
    export_html = fetch_html(EXPORT_ROSTER_URL)

    print("Fetching flight dispatch roster...")
    dispatch_html = fetch_html(IMPORT_DISPATCH_URL)

    print("Parsing all employees...")
    export_employees = parse_export_roster_all(export_html)
    dispatch_employees = parse_flight_dispatch_all(dispatch_html)

    all_employees = unique_keep_order(export_employees + dispatch_employees)

    base_dir = Path(__file__).resolve().parent.parent
    output_dir = base_dir / "data" / "report"
    output_dir.mkdir(parents=True, exist_ok=True)

    output_file = output_dir / "employees.json"
    with output_file.open("w", encoding="utf-8") as f:
        json.dump(all_employees, f, ensure_ascii=False, indent=2)

    print(f"Done ✔ {output_file} created")
    print(f"Total employees: {len(all_employees)}")


if __name__ == "__main__":
    main()