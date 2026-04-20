from pathlib import Path
import json
import re


def norm(text: str) -> str:
    return " ".join(str(text).split()).strip()


def find_value_after_label(row, label):
    label = label.upper()
    for i, cell in enumerate(row):
        if norm(cell).upper() == label:
            for j in range(i + 1, len(row)):
                value = norm(row[j])
                if value:
                    return value
    return ""


_MONTH_ABBREVS = (
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
)


def normalize_offload_date(value: str) -> str:
    """Turn YYYY-MMM-DD or YYYY-MM-DD into DD-MMM-YYYY (e.g. 19-APR-2026)."""
    s = norm(value)
    if not s:
        return ""
    m = re.match(r"^(\d{4})-([A-Za-z]{3})-(\d{1,2})$", s)
    if m:
        year, mon, day = m.groups()
        return f"{int(day):02d}-{mon.upper()}-{year}"
    m2 = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
    if m2:
        year, mo, day = m2.groups()
        mi = int(mo)
        if 1 <= mi <= 12:
            mon = _MONTH_ABBREVS[mi - 1]
            return f"{int(day):02d}-{mon}-{year}"
    return s


def build_offload(rows):
    result = {
        "flight": "",
        "date": "",
        "destination": "",
        "items": []
    }

    awb_header_index = -1

    for i, row in enumerate(rows):
        upper_row = [norm(x).upper() for x in row]

        if "FLIGHT #" in upper_row:
            result["flight"] = find_value_after_label(row, "FLIGHT #")
            result["date"] = normalize_offload_date(find_value_after_label(row, "DATE"))
            result["destination"] = find_value_after_label(row, "DESTINATION")

        if "AWB" in upper_row and "PCS" in upper_row and "KGS" in upper_row:
            awb_header_index = i
            break

    if awb_header_index == -1:
        return result

    for row in rows[awb_header_index + 1:]:
        cleaned = [norm(x) for x in row]

        if not any(cleaned):
            continue

        first = cleaned[0].upper() if len(cleaned) > 0 else ""
        if first in {"RGDS,", "LC-ZAINAB"}:
            continue

        awb = cleaned[0] if len(cleaned) > 0 else ""
        pcs = cleaned[1] if len(cleaned) > 1 else ""
        kgs = cleaned[2] if len(cleaned) > 2 else ""
        desc = cleaned[3] if len(cleaned) > 3 else ""
        reason = cleaned[4] if len(cleaned) > 4 else ""

        if not awb:
            continue

        result["items"].append({
            "awb": awb,
            "pcs": pcs,
            "kgs": kgs,
            "desc": desc,
            "reason": reason
        })

    return result


def main():
    base_dir = Path(__file__).resolve().parent
    parsed_file = base_dir / "parsed" / "latest.json"
    report_dir = base_dir / "report"
    report_dir.mkdir(parents=True, exist_ok=True)

    with parsed_file.open("r", encoding="utf-8") as f:
        data = json.load(f)

    tables = data.get("tables", [])
    if not tables:
        print("No tables found [ERROR]")
        return

    rows = tables[0]["rows"]
    offload = build_offload(rows)

    output_file = report_dir / "latest.json"
    with output_file.open("w", encoding="utf-8") as f:
        json.dump(offload, f, ensure_ascii=False, indent=2)

    print(f"Done [OK] {output_file}")


if __name__ == "__main__":
    main()