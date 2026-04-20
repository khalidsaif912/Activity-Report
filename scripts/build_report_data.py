import json
from pathlib import Path

from manpower_layout import finalize_manpower_sections
from report_date_config import get_report_date


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def roster_sections_from_list(manpower_sections: list) -> dict[str, list]:
    return {
        section["title"]: section.get("items", [])
        for section in manpower_sections
    }


def write_dates_index(report_dir: Path, active_date: str) -> None:
    by_date = report_dir / "by-date"
    dates: list[str] = []
    if by_date.is_dir():
        for p in sorted(by_date.iterdir()):
            if p.is_dir() and (p / "latest.json").is_file():
                dates.append(p.name)
    if active_date not in dates:
        dates.append(active_date)
        dates.sort()
    idx = {"dates": dates, "default": active_date}
    (report_dir / "dates_index.json").write_text(
        json.dumps(idx, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    date = get_report_date()
    print(f"Report date: {date}")

    roster_dated = base_dir / "data" / "roster" / "by-date" / date / "latest.json"
    roster_file = roster_dated if roster_dated.is_file() else base_dir / "data" / "roster" / "latest.json"

    dispatch_dated = base_dir / "data" / "flight_dispatch" / "by-date" / date / "latest.json"
    dispatch_file = dispatch_dated if dispatch_dated.is_file() else base_dir / "data" / "flight_dispatch" / "latest.json"

    report_dir = base_dir / "data" / "report"
    report_dir.mkdir(parents=True, exist_ok=True)
    report_file = report_dir / "latest.json"

    roster_data = load_json(roster_file)
    dispatch_data = load_json(dispatch_file)

    shifts_in = roster_data.get("shifts")
    by_shift_dispatch = dispatch_data.get("byShift")

    if isinstance(shifts_in, dict) and shifts_in:
        if not isinstance(by_shift_dispatch, dict):
            single_fd = dispatch_data.get("flightDispatch", {}).get("items", [])
            by_shift_dispatch = {
                k: {"items": list(single_fd)}
                for k in ("morning", "afternoon", "night")
                if k in shifts_in
            }

        shifts_out: dict[str, dict] = {}
        for key in ("morning", "afternoon", "night"):
            if key not in shifts_in:
                continue
            rs = roster_sections_from_list(shifts_in[key].get("manpowerSections", []))
            fd_items = (by_shift_dispatch.get(key) or {}).get("items", [])
            shifts_out[key] = {
                "shiftMeta": shifts_in[key].get("shiftMeta", {}),
                "manpowerSections": finalize_manpower_sections(rs, fd_items),
            }

        default_key = roster_data.get("defaultShift") or dispatch_data.get("defaultShift") or "morning"
        if default_key not in shifts_out:
            default_key = next(iter(shifts_out))

        merged = shifts_out[default_key]
        final_data = {
            "shiftMeta": merged["shiftMeta"],
            "manpowerSections": merged["manpowerSections"],
            "shifts": shifts_out,
            "defaultShift": default_key,
        }
    else:
        roster_sections = roster_sections_from_list(roster_data.get("manpowerSections", []))
        flight_dispatch_items = dispatch_data.get("flightDispatch", {}).get("items", [])
        merged_sections = finalize_manpower_sections(roster_sections, flight_dispatch_items)
        final_data = {
            "shiftMeta": roster_data.get("shiftMeta", {}),
            "manpowerSections": merged_sections,
        }

    text = json.dumps(final_data, ensure_ascii=False, indent=2)
    report_file.write_text(text, encoding="utf-8")

    by_date_dir = report_dir / "by-date" / date
    by_date_dir.mkdir(parents=True, exist_ok=True)
    (by_date_dir / "latest.json").write_text(text, encoding="utf-8")

    write_dates_index(report_dir, date)

    print(f"Done [OK] {report_file}")
    print(f"Done [OK] {by_date_dir / 'latest.json'}")
    print(f"Done [OK] {report_dir / 'dates_index.json'}")


if __name__ == "__main__":
    main()