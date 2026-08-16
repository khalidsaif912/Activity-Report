import json
from pathlib import Path

from roster_schedules import employee_directory


def unique_keep_order(items: list[str]) -> list[str]:
    seen = set()
    result = []
    for item in items:
        value = item.strip()
        key = value.upper()
        if value and key not in seen:
            seen.add(key)
            result.append(value)
    return result


def main() -> None:
    print("Fetching employee directory from roster-site schedules...")
    all_employees = unique_keep_order(employee_directory())

    base_dir = Path(__file__).resolve().parent.parent
    output_dir = base_dir / "data" / "report"
    output_dir.mkdir(parents=True, exist_ok=True)

    output_file = output_dir / "employees.json"
    if not all_employees and output_file.is_file():
        print("Schedules returned 0 names; keeping existing employees.json")
        return

    with output_file.open("w", encoding="utf-8") as f:
        json.dump(all_employees, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print(f"Done {output_file}")
    print(f"Total employees: {len(all_employees)}")


if __name__ == "__main__":
    main()
