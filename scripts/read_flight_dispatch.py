import json
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from report_date_config import get_report_date


SHIFT_CODE = "AN13"   # MN06 / AN13 / NN21


def page_url_for_date(date: str) -> str:
    return (
        "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/import/"
        f"{date}/index.html"
    )


def fetch_html(url: str) -> str:
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return response.text


def parse_dispatch(html: str, shift_code: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    items = []

    for card in soup.find_all("div", class_="dept-card"):
        title_tag = card.find("div", class_="dept-title")
        if not title_tag:
            continue

        title = title_tag.get_text(" ", strip=True)
        if title not in ["Flight Dispatch (Export)", "Flight Dispatch (Import)"]:
            continue

        for row in card.find_all("div", class_="emp-row"):
            name_tag = row.find("span", class_="emp-name")
            code_tag = row.find("span", class_="emp-code")

            if not name_tag or not code_tag:
                continue

            emp_shift = code_tag.get_text(" ", strip=True)
            if emp_shift != shift_code:
                continue

            text = name_tag.get_text(" ", strip=True)

            if "·" in text:
                name, sn = text.split("·", 1)
                name = name.strip()
                sn = sn.strip()
                items.append(f"SN{sn} {name}")
            else:
                items.append(text.strip())

    seen = set()
    result = []
    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)

    return result


def main() -> None:
    date = get_report_date()
    print(f"Report date: {date}")
    print("Fetching flight dispatch...")
    html = fetch_html(page_url_for_date(date))

    print("Parsing...")
    dispatch_items = parse_dispatch(html, SHIFT_CODE)

    by_shift: dict[str, dict] = {}
    for key, code in (
        ("morning", "MN06"),
        ("afternoon", "AN13"),
        ("night", "NN21"),
    ):
        by_shift[key] = {
            "title": "Flight Dispatch",
            "items": parse_dispatch(html, code),
        }

    code_to_shift = {"MN06": "morning", "AN13": "afternoon", "NN21": "night"}

    data = {
        "shiftMeta": {
            "date": date,
            "shiftCode": SHIFT_CODE,
        },
        "flightDispatch": {
            "title": "Flight Dispatch",
            "items": dispatch_items,
        },
        "byShift": by_shift,
        "defaultShift": code_to_shift.get(SHIFT_CODE, "afternoon"),
    }

    base_dir = Path(__file__).resolve().parent.parent
    output_dir = base_dir / "data" / "flight_dispatch"
    output_dir.mkdir(parents=True, exist_ok=True)

    output_file = output_dir / "latest.json"
    with output_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    by_date_dir = output_dir / "by-date" / date
    by_date_dir.mkdir(parents=True, exist_ok=True)
    dated_file = by_date_dir / "latest.json"
    with dated_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Done [ok] {output_file} created")
    print(f"Done [ok] {dated_file} created")


if __name__ == "__main__":
    main()