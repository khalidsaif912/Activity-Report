import json
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from report_date_config import get_report_date


SHIFT_CODE = "AN13"   # MN06 / AN13 / NN21
DISPATCH_TITLES = {"Flight Dispatch (Export)", "Flight Dispatch (Import)"}


def page_urls_for_date(date: str) -> list[str]:
    return [
        f"https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/import/date/{date}/index.html",
        f"https://khalidsaif912.github.io/roster-site/import/date/{date}/",
        f"https://khalidsaif912.github.io/roster-site/import/date/{date}/index.html",
        f"https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/import/{date}/index.html",
    ]


def fetch_html_candidates(urls: list[str]) -> list[str]:
    pages: list[str] = []
    for url in urls:
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            text = response.text or ""
            if "<html" in text.lower() or "Flight Dispatch" in text or "deptCard" in text:
                print(f"Fetched dispatch HTML: {url} ({len(text)} bytes)")
                pages.append(text)
        except Exception as exc:  # noqa: BLE001 — try next mirror
            print(f"Skip {url} ({exc})")
    return pages


def _line_from_name(text: str) -> str:
    full = (text or "").strip()
    if "·" in full:
        name, sn = full.split("·", 1)
        name, sn = name.strip(), sn.strip()
        if sn and not sn.upper().startswith("SN"):
            sn = f"SN{sn}"
        return f"{sn} {name}".strip()
    if " - " in full:
        name, sn = full.rsplit(" - ", 1)
        name, sn = name.strip(), sn.strip()
        if sn and not sn.upper().startswith("SN"):
            sn = f"SN{sn}"
        return f"{sn} {name}".strip()
    return full


def parse_dispatch(html: str, shift_code: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    items = []

    dept_cards = soup.find_all("div", class_="deptCard")
    if not dept_cards:
        dept_cards = soup.find_all("div", class_="dept-card")

    for card in dept_cards:
        title_tag = card.find(class_="deptTitle") or card.find(class_="dept-title")
        if not title_tag:
            continue

        title = title_tag.get_text(" ", strip=True)
        if title not in DISPATCH_TITLES:
            continue

        for row in card.find_all("div", class_=["empRow", "emp-row"]):
            name_tag = row.find("span", class_="empName") or row.find("span", class_="emp-name")
            code_tag = row.find("span", class_="empStatus") or row.find("span", class_="emp-code")
            if not name_tag:
                continue

            emp_shift = code_tag.get_text(" ", strip=True) if code_tag else ""
            if shift_code and shift_code not in emp_shift:
                continue

            items.append(_line_from_name(name_tag.get_text(" ", strip=True)))

    seen = set()
    result = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)

    return result


def main() -> None:
    date = get_report_date()
    print(f"Report date: {date}")
    print("Fetching flight dispatch...")
    pages = fetch_html_candidates(page_urls_for_date(date))
    html = ""
    best_count = -1
    for page in pages:
        n = sum(len(parse_dispatch(page, code)) for code in ("MN06", "AN13", "NN21"))
        if n > best_count:
            best_count = n
            html = page
    if not html:
        print("No dispatch HTML; writing empty payload")

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
    print(
        "Dispatch counts: "
        f"morning={len(by_shift['morning']['items'])} "
        f"afternoon={len(by_shift['afternoon']['items'])} "
        f"night={len(by_shift['night']['items'])}"
    )


if __name__ == "__main__":
    main()
