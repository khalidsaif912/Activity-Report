import subprocess
import sys
from pathlib import Path


def run(script_path: Path) -> bool:
    print(f"\n[RUN] {script_path}")
    result = subprocess.run(
        [sys.executable, str(script_path)],
        capture_output=True,
        text=True
    )

    if result.stdout:
        print(result.stdout)

    if result.stderr:
        print("Error:", result.stderr)

    if result.returncode != 0:
        print(f"[FAIL] {script_path} exited with code {result.returncode}")
        return False
    print(f"[OK] {script_path}")
    return True


def main():
    base_dir = Path(__file__).resolve().parent

    scripts = [
        base_dir / "scripts" / "read_roster.py",
        base_dir / "scripts" / "read_flight_dispatch.py",
        base_dir / "scripts" / "build_report_data.py",
        base_dir / "data" / "offload" / "run_pipeline.py",
    ]

    ok = True
    for script in scripts:
        ok = run(script) and ok
    if not ok:
        raise SystemExit(1)


if __name__ == "__main__":
    main()