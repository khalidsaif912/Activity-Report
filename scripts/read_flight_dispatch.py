import json
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from report_date_config import get_report_date

SHIFT_CODE = "AN13"
DUTY_SHIFT_KEYS = {
    "morning": "morning",
    "afternoon": "afternoon",
    "night": "night",
}
SHIFT_CODE_TO_KEY = {
    "MN06": "morning",
    "ME06": "morning",
    "ME07": "morning",
    "ME12": "morning",
    "AN13": "afternoon",
    "AE14": "afternoon",
    "NN21": "night",
    "NE22": "night",
}


def page_urls_for_date(date: str) -> list[str]:
    return [
        f"https://khalidsaif912.github.io/new/docs/import/date/{date}/",
        f"https://khalidsaif912.github.io/new/docs/import/date/{date}/index.html",
        f"https://raw.githubusercontent.com/khalidsaif912/new/main/docs/import/date/{date}/index.html",
        f"https://khalidsaif912.github.io/new/docs/import/now/",
        f"https://khalidsaif912.github.io/new/docs/import/",
        f"https://raw.githubusercontent.com/khalidsaif912/new/main/docs/import/now/index.html",
        f"https://khalidsaif912.github.io/roster-site/import/date/{date}/",
        f"https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/import/date/{date}/index.html",
    ]


def fetch_html_candidates(urls: list[str]) -> list[str]:
    pages: list[str] = []
    for url in urls:
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            text = response.text or ""
            if "Flight Dispatch" not in text:
                print(f"Skip (no Flight Dispatch): {url}")
                continue
            print(f"Fetched import dispatch HTML: {url} ({len(text)} bytes)")
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


def is_dispatch_title(title: str) -> bool:
    t = str(title or "").strip().lower()
    return t.startswith("flight dispatch")


def shift_key_from_card(shift_card, emp_shift: str) -> str | None:
    raw = str(shift_card.get("data-shift") or "").strip().lower()
    if raw in DUTY_SHIFT_KEYS:
        return DUTY_SHIFT_KEYS[raw]
    code = ""
    for token in str(emp_shift or "").upper().replace("(", " ").split():
        if token in SHIFT_CODE_TO_KEY:
            code = token
            break
    return SHIFT_CODE_TO_KEY.get(code)


def unique_add(bucket: list[str], seen: set[str], line: str) -> None:
    value = str(line or "").strip()
    key = value.upper()
    if not value or key in seen:
        return
    seen.add(key)
    bucket.append(value)


def parse_dispatch_by_shift(html: str) -> dict[str, list[str]]:
    """Merge Flight Dispatch (Export) + Flight Dispatch (Import) into one list per duty shift."""
    by_shift: dict[str, list[str]] = {"morning": [], "afternoon": [], "night": []}
    seen: dict[str, set[str]] = {k: set() for k in by_shift}
    if not html:
        return by_shift

    soup = BeautifulSoup(html, "html.parser")
    dept_cards = soup.find_all("div", class_="deptCard")
    if not dept_cards:
        dept_cards = soup.find_all("div", class_="dept-card")

    for card in dept_cards:
        title_tag = card.find(class_="deptTitle") or card.find(class_="dept-title")
        title = title_tag.get_text(" ", strip=True) if title_tag else ""
        if not is_dispatch_title(title):
            continue

        shift_cards = card.find_all("details", class_="shiftCard")
        if not shift_cards:
            shift_cards = [card]

        for shift_card in shift_cards:
            rows = shift_card.find_all("div", class_=["empRow", "emp-row"])
            for row in rows:
                name_tag = row.find("span", class_="empName") or row.find("span", class_="emp-name")
                code_tag = row.find("span", class_="empStatus") or row.find("span", class_="emp-code")
                if not name_tag:
                    continue
                emp_shift = code_tag.get_text(" ", strip=True) if code_tag else ""
                key = shift_key_from_card(shift_card, emp_shift)
                if not key:
                    continue
                unique_add(by_shift[key], seen[key], _line_from_name(name_tag.get_text(" ", strip=True)))

    return by_shift


def union_into_employees_json(names: list[str], report_dir: Path) -> None:
    path = report_dir / "employees.json"
    existing: list[str] = []
    if path.is_file():
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(raw, list):
                existing = [str(x).strip() for x in raw if str(x).strip()]
        except (OSError, json.JSONDecodeError):
            existing = []
    seen = {x.upper() for x in existing}
    out = existing[:]
    for name in names:
        line = str(name or "").strip()
        key = line.upper()
        if not line or key in seen:
            continue
        seen.add(key)
        out.append(line)
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"employees.json count={len(out)} (after dispatch union)")


def main() -> None:
    date = get_report_date()
    print(f"Report date: {date}")
    print("Fetching Flight Dispatch from import roster...")
    pages = fetch_html_candidates(page_urls_for_date(date))
    best: dict[str, list[str]] = {"morning": [], "afternoon": [], "night": []}
    best_count = -1
    for page in pages:
        parsed = parse_dispatch_by_shift(page)
        n = sum(len(v) for v in parsed.values())
        if n > best_count:
            best_count = n
            best = parsed
    if best_count < 0:
        print("No import dispatch HTML; writing empty payload")
        best_count = 0

    by_shift = {
        key: {"title": "Flight Dispatch", "items": best.get(key, [])}
        for key in ("morning", "afternoon", "night")
    }
    all_names: list[str] = []
    seen_all: set[str] = set()
    for key in ("morning", "afternoon", "night"):
        for item in by_shift[key]["items"]:
            unique_add(all_names, seen_all, item)

    data = {
        "shiftMeta": {
            "date": date,
            "shiftCode": SHIFT_CODE,
            "source": "import-roster",
        },
        "flightDispatch": {
            "title": "Flight Dispatch",
            "items": all_names,
        },
        "byShift": by_shift,
        "defaultShift": "afternoon",
    }

    base_dir = Path(__file__).resolve().parent.parent
    output_dir = base_dir / "data" / "flight_dispatch"
    output_dir.mkdir(parents=True, exist_ok=True)

    output_file = output_dir / "latest.json"
    with output_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")

    by_date_dir = output_dir / "by-date" / date
    by_date_dir.mkdir(parents=True, exist_ok=True)
    dated_file = by_date_dir / "latest.json"
    with dated_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")

    union_into_employees_json(all_names, base_dir / "data" / "report")

    print(f"Done [ok] {output_file} created")
    print(f"Done [ok] {dated_file} created")
    print(
        "Dispatch counts: "
        f"morning={len(by_shift['morning']['items'])} "
        f"afternoon={len(by_shift['afternoon']['items'])} "
        f"night={len(by_shift['night']['items'])} "
        f"unique={len(all_names)}"
    )


if __name__ == "__main__":
    main()
