import subprocess
import sys
from pathlib import Path


def run(script_name: str):
    print(f"\n[RUNNING] {script_name}")
    result = subprocess.run(
        [sys.executable, script_name],
        capture_output=True,
        text=True
    )

    print(result.stdout)
    if result.stderr:
        print("Error:", result.stderr)


def main():
    base_dir = Path(__file__).resolve().parent

    scripts = [
        "fetch_offload_source.py",
        "archive_offload_snapshot.py",
        "parse_offload_to_json.py",
        "build_offload_latest.py",
        "archive_offload_report.py",
    ]

    for script in scripts:
        run(str(base_dir / script))


if __name__ == "__main__":
    main()