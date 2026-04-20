import hashlib
from pathlib import Path
import requests


OFFLOAD_URL = "https://omanair-my.sharepoint.com/:u:/p/8715_hq/IQAhl_w3GIpTR5hhL-y1Gb9lAbYN-eMz6VSIH3tkJcf-3A4?e=hQ27fH&download=1"


def download_file(url: str) -> bytes:
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    return response.content


def sha256_bytes(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    source_dir = base_dir / "source"
    source_dir.mkdir(parents=True, exist_ok=True)

    print("Downloading offload source...")
    content = download_file(OFFLOAD_URL)

    file_path = source_dir / "latest.html"
    hash_path = source_dir / "latest.hash"

    file_path.write_bytes(content)
    hash_path.write_text(sha256_bytes(content), encoding="utf-8")

    print(f"Done [OK] {file_path}")
    print(f"Done [OK] {hash_path}")


if __name__ == "__main__":
    main()