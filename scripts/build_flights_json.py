"""
Build / refresh data/report/flights.json from a live provider (Aviationstack).

This keeps GitHub Pages (static) flight suggestions up-to-date without running a backend.

Env:
  AVIATIONSTACK_ACCESS_KEY   (required)
  LIVE_FLIGHTS_AIRLINES      default: "WY,OV"
  LIVE_FLIGHTS_DEP_IATA      default: "MCT"
  LIVE_FLIGHTS_TIMEOUT_SEC   default: "20"
"""

from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path

import requests


def _fmt_ddmon(dt_like: str) -> str:
    s = str(dt_like or "").strip()
    if not s:
        return ""
    try:
        s_norm = s.replace("Z", "+00:00")
        dt = datetime.fromisoformat(s_norm)
        return dt.strftime("%d%b").upper()
    except ValueError:
        pass
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        try:
            dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return dt.strftime("%d%b").upper()
        except ValueError:
            return ""
    return ""


def _fmt_hhmm(dt_like: str) -> str:
    s = str(dt_like or "").strip()
    if not s:
        return ""
    try:
        s_norm = s.replace("Z", "+00:00")
        dt = datetime.fromisoformat(s_norm)
        return dt.strftime("%H%M")
    except ValueError:
        pass
    m = re.search(r"\b(\d{2}):(\d{2})", s)
    if m:
        return f"{m.group(1)}{m.group(2)}"
    return ""


def _norm_iata3(value: str) -> str:
    s = str(value or "").strip().upper()
    m = re.match(r"^([A-Z]{3})$", s)
    return m.group(1) if m else ""


def _norm_flight_code(value: str) -> str:
    s = str(value or "").strip().upper().replace(" ", "")
    m = re.match(r"^([A-Z]{2}\d{1,4})$", s)
    return m.group(1) if m else ""


def _build_std_etd(std_raw: str, etd_raw: str) -> str:
    std = _fmt_hhmm(std_raw)
    etd = _fmt_hhmm(etd_raw)
    if std and etd and std != etd:
        return f"{std}/{etd}"
    return std or etd


def fetch_aviationstack(access_key: str, airlines: list[str], dep_iata: str, timeout_sec: int) -> list[dict]:
    out_by_key: dict[str, dict] = {}
    scopes = airlines[:] if airlines else [""]
    for airline in scopes:
        params = {"access_key": access_key}
        if airline:
            params["airline_iata"] = airline
        if dep_iata:
            params["dep_iata"] = dep_iata
        try:
            r = requests.get(
                "http://api.aviationstack.com/v1/flights",
                params=params,
                timeout=timeout_sec,
            )
            if not r.ok:
                continue
            payload = r.json() if r.text else {}
            rows = payload.get("data")
            if not isinstance(rows, list):
                continue
        except (requests.RequestException, ValueError):
            continue

        for row in rows:
            if not isinstance(row, dict):
                continue
            flight = row.get("flight") if isinstance(row.get("flight"), dict) else {}
            dep = row.get("departure") if isinstance(row.get("departure"), dict) else {}
            arr = row.get("arrival") if isinstance(row.get("arrival"), dict) else {}

            code = _norm_flight_code(flight.get("iata"))
            date = _fmt_ddmon(dep.get("scheduled") or dep.get("estimated") or dep.get("actual"))
            dest = _norm_iata3(arr.get("iata"))
            std_etd = _build_std_etd(dep.get("scheduled"), dep.get("estimated") or dep.get("actual"))
            if not code or not date or not dest:
                continue
            key = f"{code}|{date}"
            out_by_key[key] = {"code": code, "date": date, "destination": dest, "stdEtd": std_etd}

    out = list(out_by_key.values())
    out.sort(key=lambda x: (x.get("code", ""), x.get("date", "")))
    return out


def main() -> None:
    access_key = (os.environ.get("AVIATIONSTACK_ACCESS_KEY") or "").strip()
    if not access_key:
        print("AVIATIONSTACK_ACCESS_KEY is not set; keeping existing flights.json.")
        return

    airlines_raw = (os.environ.get("LIVE_FLIGHTS_AIRLINES") or "WY,OV").strip()
    if airlines_raw.upper() in {"ALL", "ANY", "*"}:
        airlines = []
    else:
        airlines = [x.strip().upper() for x in airlines_raw.split(",") if x.strip()]

    dep_raw = (os.environ.get("LIVE_FLIGHTS_DEP_IATA") or "MCT").strip().upper()
    dep_iata = "" if dep_raw in {"ALL", "ANY", "*"} else dep_raw
    timeout_sec = int((os.environ.get("LIVE_FLIGHTS_TIMEOUT_SEC") or "20").strip() or "20")

    print(
        f"Flight source config: airlines={'ALL' if not airlines else ','.join(airlines)} "
        f"dep_iata={'ALL' if not dep_iata else dep_iata}"
    )

    flights = fetch_aviationstack(access_key, airlines or ["WY", "OV"], dep_iata, timeout_sec)
    if not flights:
        print("No live flights returned; keeping existing flights.json.")
        return

    base_dir = Path(__file__).resolve().parent.parent
    report_dir = base_dir / "data" / "report"
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / "flights.json"
    out_path.write_text(json.dumps(flights, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Done [OK] {out_path} ({len(flights)} flights)")


if __name__ == "__main__":
    main()

