from pathlib import Path
from datetime import datetime
import shutil


def main() -> None:
    base_dir = Path(__file__).resolve().parent

    source_dir = base_dir / "source"
    archive_dir = base_dir / "archive"

    archive_dir.mkdir(parents=True, exist_ok=True)

    latest_file = source_dir / "latest.html"
    latest_hash_file = source_dir / "latest.hash"
    last_hash_file = archive_dir / "last.hash"

    if not latest_file.exists() or not latest_hash_file.exists():
        print("No source file found ❌")
        return

    current_hash = latest_hash_file.read_text(encoding="utf-8").strip()

    if last_hash_file.exists():
        last_hash = last_hash_file.read_text(encoding="utf-8").strip()

        if current_hash == last_hash:
            print("No changes detected [OK]")
            return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    archive_file = archive_dir / f"{timestamp}.html"

    shutil.copy2(latest_file, archive_file)
    last_hash_file.write_text(current_hash, encoding="utf-8")

    print(f"Archived ✔ {archive_file}")


if __name__ == "__main__":
    main()