from pathlib import Path
import json
from bs4 import BeautifulSoup


def clean_text(text: str) -> str:
    return " ".join(str(text).replace("\xa0", " ").split()).strip()


def parse_tables(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    tables_data = []

    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            values = [clean_text(cell.get_text(" ", strip=True)) for cell in cells]
            if any(values):
                rows.append(values)

        if rows:
            tables_data.append({"rows": rows})

    return tables_data


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    source_file = base_dir / "source" / "latest.html"
    parsed_dir = base_dir / "parsed"
    parsed_dir.mkdir(parents=True, exist_ok=True)

    if not source_file.exists():
        print("Source HTML not found ❌")
        return

    print("Reading HTML...")
    html = source_file.read_text(encoding="utf-8", errors="ignore")

    print("Parsing offload tables...")
    data = {
        "tables": parse_tables(html)
    }

    output_file = parsed_dir / "latest.json"
    with output_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Done [OK] {output_file}")


if __name__ == "__main__":
    main()