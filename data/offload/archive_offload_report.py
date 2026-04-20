from pathlib import Path
from datetime import datetime
import hashlib
import shutil


def file_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    base_dir = Path(__file__).resolve().parent

    report_dir = base_dir / "report"
    archive_dir = base_dir / "report_archive"

    archive_dir.mkdir(parents=True, exist_ok=True)

    latest_file = report_dir / "latest.json"
    latest_hash_file = report_dir / "latest.hash"
    last_hash_file = archive_dir / "last.hash"

    if not latest_file.exists():
        print("No report file found ❌")
        return

    current_hash = file_sha256(latest_file)
    latest_hash_file.write_text(current_hash, encoding="utf-8")

    if last_hash_file.exists():
        last_hash = last_hash_file.read_text(encoding="utf-8").strip()
        if current_hash == last_hash:
            print("No report changes detected [OK]")
            return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    archive_file = archive_dir / f"{timestamp}.json"

    shutil.copy2(latest_file, archive_file)
    last_hash_file.write_text(current_hash, encoding="utf-8")

    print(f"Archived report [OK] {archive_file}")


if __name__ == "__main__":
    main()