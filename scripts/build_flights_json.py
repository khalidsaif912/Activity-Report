"""
Build / refresh data/report/flights.json from live sources.

Primary source:
  - Muscat Airport departures page (official MCT departures list)

Fallback source:
  - Aviationstack API

This keeps GitHub Pages (static) flight suggestions up-to-date without running a backend.

Env:
  AVIATIONSTACK_ACCESS_KEY   (optional, used as fallback/enrichment)
  LIVE_FLIGHTS_AIRLINES      default: "WY,OV" (or "ALL")
  LIVE_FLIGHTS_DEP_IATA      default: "MCT"   (or "ALL")
  LIVE_FLIGHTS_TIMEOUT_SEC   default: "20"
"""

from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

DESTINATION_NAME_TO_IATA = {
    "ABHA": "AHB",
    "ABU DHABI": "AUH",
    "ADDIS ABABA": "ADD",
    "AMMAN": "AMM",
    "AMSTERDAM": "AMS",
    "BAHRAIN": "BAH",
    "BANGKOK": "BKK",
    "BENGALURU": "BLR",
    "BOMBAY": "BOM",
    "CAIRO": "CAI",
    "CALICUT": "CCJ",
    "CHENNAI": "MAA",
    "CHITTAGONG": "CGP",
    "COCHIN": "COK",
    "COLOMBO": "CMB",
    "DAMMAM": "DMM",
    "DELHI": "DEL",
    "DHAKA": "DAC",
    "DOHA": "DOH",
    "DUQM": "DQM",
    "DUBAI": "DXB",
    "FAHUD": "FAU",
    "FUJAIRAH": "FJR",
    "HYDERABAD": "HYD",
    "ISLAMABAD": "ISB",
    "ISTANBUL": "IST",
    "JEDDAH": "JED",
    "KARACHI": "KHI",
    "KUALA LUMPUR": "KUL",
    "KUWAIT": "KWI",
    "LAHORE": "LHE",
    "LONDON": "LHR",
    "LUCKNOW": "LKO",
    "MEDINA": "MED",
    "MILANO": "MXP",
    "MOSCOW": "SVO",
    "MUMBAI": "BOM",
    "MULTAN": "MUX",
    "MUSCAT": "MCT",
    "MUKHAIZNA": "UKH",
    "MUSCAT (MCT)": "MCT",
    "MUNICH": "MUC",
    "PHUKET": "HKT",
    "PORT SUDAN": "PZU",
    "RIYADH": "RUH",
    "SALALAH": "SLL",
    "SHARJAH": "SHJ",
    "SIALKOT": "SKT",
    "THIRUVANANTHAPURAM": "TRV",
    "TRIVANDRUM": "TRV",
    "TEHRAN": "IKA",
    "ZURICH": "ZRH",
}


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


def _dest_fallback_code(destination_name: str) -> str:
    src = re.sub(r"[^A-Za-z0-9 ]+", " ", str(destination_name or "")).upper()
    src = re.sub(r"\s+", " ", src).strip()
    if src in DESTINATION_NAME_TO_IATA:
        return DESTINATION_NAME_TO_IATA[src]
    tokens = [x for x in src.split() if x]
    if not tokens:
        return ""
    first = tokens[0]
    if re.fullmatch(r"[A-Z]{3}", first):
        return first
    return (first[:3] if len(first) >= 3 else first).upper()


def _parse_muscat_sched(value: str) -> tuple[str, str]:
    s = str(value or "").strip()
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2})", s)
    if not m:
        return "", ""
    dd, mm, yyyy, hh, mi = m.groups()
    try:
        dt = datetime(int(yyyy), int(mm), int(dd))
    except ValueError:
        return "", ""
    return dt.strftime("%d%b").upper(), f"{hh}{mi}"


def load_existing_dest_maps(report_dir: Path) -> tuple[dict[str, str], dict[str, str]]:
    by_code_date: dict[str, str] = {}
    by_code: dict[str, str] = {}
    src = report_dir / "flights.json"
    if not src.is_file():
        return by_code_date, by_code
    try:
        rows = json.loads(src.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return by_code_date, by_code
    if not isinstance(rows, list):
        return by_code_date, by_code
    for row in rows:
        if not isinstance(row, dict):
            continue
        code = _norm_flight_code(row.get("code"))
        date = str(row.get("date") or "").strip().upper()
        dest = _norm_iata3(row.get("destination")) or _dest_fallback_code(row.get("destination"))
        if not code or not dest:
            continue
        if date:
            by_code_date[f"{code}|{date}"] = dest
        by_code[code] = dest
    return by_code_date, by_code


def fetch_muscat_departures(timeout_sec: int) -> list[dict]:
    url = (
        "https://www.muscatairport.co.om/flightstatusframe"
        "?type=2&airline=&from=&to=&flight_name=&date=&condition=&date_type="
    )
    try:
        r = requests.get(url, timeout=timeout_sec)
        if not r.ok:
            return []
        html = r.text
    except requests.RequestException:
        return []

    soup = BeautifulSoup(html, "html.parser")
    rows = soup.select("table tbody tr")
    out_by_key: dict[str, dict] = {}
    for tr in rows:
        cells = tr.find_all("td")
        if len(cells) < 4:
            continue
        destination_name = cells[1].get_text(" ", strip=True)
        flight_text = cells[2].get_text(" ", strip=True).replace(" ", "")
        sched_text = cells[3].get_text(" ", strip=True)

        code = _norm_flight_code(flight_text)
        date, std = _parse_muscat_sched(sched_text)
        if not code or not date:
            continue
        key = f"{code}|{date}"
        out_by_key[key] = {
            "code": code,
            "date": date,
            "destination_name": destination_name,
            "stdEtd": std,
        }

    out = list(out_by_key.values())
    out.sort(key=lambda x: (x.get("code", ""), x.get("date", "")))
    return out


def fetch_aviationstack(access_key: str, airlines: list[str], dep_iata: str, timeout_sec: int) -> list[dict]:
    out_by_key: dict[str, dict] = {}
    scopes = airlines[:] if airlines else [""]
    for airline in scopes:
        params = {"access_key": access_key}
        if airline:
            params["airline_iata"] = airline
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
            dep_code = _norm_iata3(dep.get("iata"))
            if dep_iata and dep_code != dep_iata:
                continue

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

    base_dir = Path(__file__).resolve().parent.parent
    report_dir = base_dir / "data" / "report"
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / "flights.json"
    by_code_date, by_code = load_existing_dest_maps(report_dir)
    if access_key:
        enrich = fetch_aviationstack(access_key, airlines or ["WY", "OV"], dep_iata, timeout_sec)
        for row in enrich:
            code = _norm_flight_code(row.get("code"))
            date = str(row.get("date") or "").strip().upper()
            dest = _norm_iata3(row.get("destination"))
            if not code or not dest:
                continue
            if date:
                by_code_date[f"{code}|{date}"] = dest
            by_code[code] = dest

    muscat_rows = fetch_muscat_departures(timeout_sec)
    if muscat_rows:
        flights = []
        for row in muscat_rows:
            code = row.get("code", "")
            date = row.get("date", "")
            key = f"{code}|{date}"
            dest_from_name = _dest_fallback_code(row.get("destination_name", ""))
            dest = dest_from_name or by_code_date.get(key) or by_code.get(code)
            if not dest:
                continue
            flights.append(
                {
                    "code": code,
                    "date": date,
                    "destination": dest,
                    "stdEtd": row.get("stdEtd", ""),
                }
            )
        if flights:
            out_path.write_text(json.dumps(flights, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"Done [OK] {out_path} ({len(flights)} flights) source=MUSCAT")
            return

    if not access_key:
        print("Muscat source empty and AVIATIONSTACK_ACCESS_KEY is not set; keeping existing flights.json.")
        return

    flights = fetch_aviationstack(access_key, airlines or ["WY", "OV"], dep_iata, timeout_sec)
    if not flights:
        print("No live flights returned from fallback source; keeping existing flights.json.")
        return
    out_path.write_text(json.dumps(flights, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Done [OK] {out_path} ({len(flights)} flights) source=AVIATIONSTACK")


if __name__ == "__main__":
    main()

