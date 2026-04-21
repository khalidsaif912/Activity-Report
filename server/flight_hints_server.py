"""
Lightweight API + static file server for flight STD/DEST hints (shared team learning).

Run from repo root:
  pip install -r requirements.txt
  python server/flight_hints_server.py

Open: http://127.0.0.1:5050/offload_report.html

Hints are stored in data/report/flight-hints.json (single JSON object, keys like "2026-04-19|WY223").
CSD route usage counts: data/report/csd-route-hints.json (keys like "FRA-MNL", integer values).
Phrase usage counts: data/report/phrase-usage.json (nested by key, e.g. loadPlan/advanceLoading).
Manpower role usage: data/report/manpower-role-hints.json (per-name role counts).
Recipients (To/Cc/Bcc): data/report/recipients.json.

Optional env FLIGHT_HINTS_TOKEN: if set, clients must send header:
  X-Flight-Hints-Token: <token>
for GET and POST.
"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import re
import secrets
from pathlib import Path
from urllib.parse import urlencode
from datetime import datetime

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from flask import Flask, abort, jsonify, request, send_file, send_from_directory
import requests

ROOT = Path(__file__).resolve().parent.parent
REPORT_DIR = ROOT / "data" / "report"
HINTS_FILE = REPORT_DIR / "flight-hints.json"
CSD_ROUTES_FILE = REPORT_DIR / "csd-route-hints.json"
PHRASE_USAGE_FILE = REPORT_DIR / "phrase-usage.json"
MANPOWER_ROLE_HINTS_FILE = REPORT_DIR / "manpower-role-hints.json"
RECIPIENTS_FILE = REPORT_DIR / "recipients.json"
GMAIL_TOKEN_FILE = REPORT_DIR / "gmail-token.json"

GMAIL_CLIENT_ID = (os.environ.get("GMAIL_CLIENT_ID") or "").strip()
GMAIL_CLIENT_SECRET = (os.environ.get("GMAIL_CLIENT_SECRET") or "").strip()
GMAIL_REDIRECT_URI = (os.environ.get("GMAIL_REDIRECT_URI") or "http://127.0.0.1:5050/api/gmail/oauth/callback").strip()
GMAIL_SCOPE = "https://www.googleapis.com/auth/gmail.send"

LIVE_FLIGHTS_PROVIDER = (os.environ.get("LIVE_FLIGHTS_PROVIDER") or "aviationstack").strip().lower()
AVIATIONSTACK_ACCESS_KEY = (os.environ.get("AVIATIONSTACK_ACCESS_KEY") or "").strip()
LIVE_FLIGHTS_AIRLINES = [
    x.strip().upper()
    for x in (os.environ.get("LIVE_FLIGHTS_AIRLINES") or "WY,OV").split(",")
    if x.strip()
]
LIVE_FLIGHTS_DEP_IATA = (os.environ.get("LIVE_FLIGHTS_DEP_IATA") or "MCT").strip().upper()
LIVE_FLIGHTS_TIMEOUT_SEC = int((os.environ.get("LIVE_FLIGHTS_TIMEOUT_SEC") or "20").strip() or "20")

_gmail_oauth_state: str | None = None

app = Flask(__name__)

TOKEN = (os.environ.get("FLIGHT_HINTS_TOKEN") or "").strip()

_PREFIX_JS = "js/"
_PREFIX_DATA_REPORT = "data/report/"
_PREFIX_DATA_OFFLOAD = "data/offload/"


def _check_token() -> bool:
    if not TOKEN:
        return True
    got = (request.headers.get("X-Flight-Hints-Token") or "").strip()
    return got == TOKEN


def _cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Flight-Hints-Token"
    return resp


@app.after_request
def _after(resp):
    return _cors(resp)


def _read_hints() -> dict:
    if not HINTS_FILE.is_file():
        return {}
    try:
        return json.loads(HINTS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _atomic_write(data: dict) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = HINTS_FILE.with_suffix(".json.tmp")
    text = json.dumps(data, ensure_ascii=False, indent=2)
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(HINTS_FILE)


def _read_csd_routes() -> dict:
    if not CSD_ROUTES_FILE.is_file():
        return {}
    try:
        return json.loads(CSD_ROUTES_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _atomic_write_csd(data: dict) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = CSD_ROUTES_FILE.with_suffix(".json.tmp")
    text = json.dumps(data, ensure_ascii=False, indent=2)
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(CSD_ROUTES_FILE)


def _read_phrase_usage() -> dict:
    if not PHRASE_USAGE_FILE.is_file():
        return {
            "loadPlan": {},
            "advanceLoading": {},
            "handoverDetails": {},
            "offloadReason": {},
            "offloadRemarks": {},
            "other": {},
            "specialHO": {},
        }
    try:
        raw = json.loads(PHRASE_USAGE_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return {
                "loadPlan": {},
                "advanceLoading": {},
                "handoverDetails": {},
                "offloadReason": {},
                "offloadRemarks": {},
                "other": {},
                "specialHO": {},
            }
        out = {
            "loadPlan": {},
            "advanceLoading": {},
            "handoverDetails": {},
            "offloadReason": {},
            "offloadRemarks": {},
            "other": {},
            "specialHO": {},
        }
        for key in ("loadPlan", "advanceLoading", "handoverDetails", "offloadReason", "offloadRemarks", "other", "specialHO"):
            bucket = raw.get(key)
            if not isinstance(bucket, dict):
                continue
            for phrase, n in bucket.items():
                p = str(phrase or "").strip().upper()
                try:
                    c = int(n)
                except (TypeError, ValueError):
                    continue
                if p and c > 0:
                    out[key][p] = c
        return out
    except (json.JSONDecodeError, OSError):
        return {
            "loadPlan": {},
            "advanceLoading": {},
            "handoverDetails": {},
            "offloadReason": {},
            "offloadRemarks": {},
            "other": {},
            "specialHO": {},
        }


def _atomic_write_phrase_usage(data: dict) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = PHRASE_USAGE_FILE.with_suffix(".json.tmp")
    text = json.dumps(data, ensure_ascii=False, indent=2)
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(PHRASE_USAGE_FILE)


def _read_manpower_role_hints() -> dict:
    if not MANPOWER_ROLE_HINTS_FILE.is_file():
        return {}
    try:
        raw = json.loads(MANPOWER_ROLE_HINTS_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return {}
        out = {}
        for name_key, entry in raw.items():
            if not isinstance(entry, dict):
                continue
            display = str(entry.get("display") or name_key).strip()
            roles_src = entry.get("roles")
            if not display or not isinstance(roles_src, dict):
                continue
            roles = {}
            for role_key, role_entry in roles_src.items():
                if not isinstance(role_entry, dict):
                    continue
                label = str(role_entry.get("label") or role_key).strip()
                try:
                    count = int(role_entry.get("count"))
                except (TypeError, ValueError):
                    continue
                if not label or count <= 0:
                    continue
                roles[str(role_key).strip().upper()] = {"label": label, "count": count}
            if roles:
                out[str(name_key).strip().upper()] = {"display": display, "roles": roles}
        return out
    except (json.JSONDecodeError, OSError):
        return {}


def _atomic_write_manpower_role_hints(data: dict) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = MANPOWER_ROLE_HINTS_FILE.with_suffix(".json.tmp")
    text = json.dumps(data, ensure_ascii=False, indent=2)
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(MANPOWER_ROLE_HINTS_FILE)


def _norm_email_list(arr) -> list[str]:
    if not isinstance(arr, list):
        return []
    out = []
    seen = set()
    for x in arr:
        e = str(x or "").strip().lower()
        if not e or e in seen:
            continue
        seen.add(e)
        out.append(e)
    return out


def _read_recipients() -> dict:
    if not RECIPIENTS_FILE.is_file():
        return {"to": [], "cc": [], "bcc": []}
    try:
        raw = json.loads(RECIPIENTS_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return {"to": [], "cc": [], "bcc": []}
        return {
            "to": _norm_email_list(raw.get("to")),
            "cc": _norm_email_list(raw.get("cc")),
            "bcc": _norm_email_list(raw.get("bcc")),
        }
    except (json.JSONDecodeError, OSError):
        return {"to": [], "cc": [], "bcc": []}


def _atomic_write_recipients(data: dict) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = RECIPIENTS_FILE.with_suffix(".json.tmp")
    clean = {
        "to": _norm_email_list(data.get("to") if isinstance(data, dict) else []),
        "cc": _norm_email_list(data.get("cc") if isinstance(data, dict) else []),
        "bcc": _norm_email_list(data.get("bcc") if isinstance(data, dict) else []),
    }
    text = json.dumps(clean, ensure_ascii=False, indent=2)
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(RECIPIENTS_FILE)


def _read_gmail_token() -> dict:
    if not GMAIL_TOKEN_FILE.is_file():
        return {}
    try:
        raw = json.loads(GMAIL_TOKEN_FILE.read_text(encoding="utf-8"))
        return raw if isinstance(raw, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def _atomic_write_gmail_token(data: dict) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = GMAIL_TOKEN_FILE.with_suffix(".json.tmp")
    text = json.dumps(data if isinstance(data, dict) else {}, ensure_ascii=False, indent=2)
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(GMAIL_TOKEN_FILE)


def _gmail_configured() -> bool:
    return bool(GMAIL_CLIENT_ID and GMAIL_CLIENT_SECRET and GMAIL_REDIRECT_URI)


def _gmail_access_token() -> str:
    token = _read_gmail_token()
    access_token = str(token.get("access_token") or "").strip()
    if access_token:
        return access_token
    refresh_token = str(token.get("refresh_token") or "").strip()
    if not refresh_token or not _gmail_configured():
        return ""

    payload = {
        "client_id": GMAIL_CLIENT_ID,
        "client_secret": GMAIL_CLIENT_SECRET,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    }
    try:
        r = requests.post("https://oauth2.googleapis.com/token", data=payload, timeout=20)
        if not r.ok:
            return ""
        out = r.json()
        new_access = str(out.get("access_token") or "").strip()
        if not new_access:
            return ""
        token["access_token"] = new_access
        _atomic_write_gmail_token(token)
        return new_access
    except requests.RequestException:
        return ""


def _norm_email(value: str) -> str:
    return str(value or "").strip()


def _norm_email_list_relaxed(arr) -> list[str]:
    if not isinstance(arr, list):
        return []
    out = []
    seen = set()
    for x in arr:
        e = _norm_email(x)
        if not e:
            continue
        k = e.lower()
        if k in seen:
            continue
        seen.add(k)
        out.append(e)
    return out


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


def _fetch_live_flights_aviationstack() -> list[dict]:
    if not AVIATIONSTACK_ACCESS_KEY:
        return []

    out_by_key: dict[str, dict] = {}
    for airline in LIVE_FLIGHTS_AIRLINES or ["WY", "OV"]:
        params = {
            "access_key": AVIATIONSTACK_ACCESS_KEY,
            "airline_iata": airline,
        }
        if LIVE_FLIGHTS_DEP_IATA:
            params["dep_iata"] = LIVE_FLIGHTS_DEP_IATA
        try:
            r = requests.get(
                "http://api.aviationstack.com/v1/flights",
                params=params,
                timeout=LIVE_FLIGHTS_TIMEOUT_SEC,
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
            out_by_key[key] = {
                "code": code,
                "date": date,
                "destination": dest,
                "stdEtd": std_etd,
            }

    out = list(out_by_key.values())
    out.sort(key=lambda x: (x.get("code", ""), x.get("date", "")))
    return out


def _fetch_live_flights() -> list[dict]:
    if LIVE_FLIGHTS_PROVIDER == "aviationstack":
        return _fetch_live_flights_aviationstack()
    return []


@app.route("/api/server-info")
def server_info():
    """Use this URL to confirm the browser hits THIS Flask app (not another process on :5050)."""
    emp = REPORT_DIR / "employees.json"
    return jsonify(
        {
            "ok": True,
            "report_dir": str(REPORT_DIR),
            "employees_json_exists": emp.is_file(),
            "employees_url_should_work": emp.is_file(),
            "live_flights_provider": LIVE_FLIGHTS_PROVIDER,
            "live_flights_configured": bool(AVIATIONSTACK_ACCESS_KEY) if LIVE_FLIGHTS_PROVIDER == "aviationstack" else False,
        }
    )


@app.route("/api/live-flights", methods=["GET"])
def live_flights_get():
    """
    Return live flight list in the same shape as data/report/flights.json:
    [{ code, date, destination, stdEtd }]
    """
    flights = _fetch_live_flights()
    if flights:
        return jsonify(flights)
    return jsonify({"error": "live_flights_unavailable", "provider": LIVE_FLIGHTS_PROVIDER}), 503


@app.route("/api/flight-hints", methods=["OPTIONS"])
def hints_options():
    return ("", 204)


@app.route("/api/flight-hints", methods=["GET"])
def hints_get():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(_read_hints())


@app.route("/api/csd-route-hints", methods=["OPTIONS"])
def csd_routes_options():
    return ("", 204)


@app.route("/api/csd-route-hints", methods=["GET"])
def csd_routes_get():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(_read_csd_routes())


@app.route("/api/csd-route-hints", methods=["POST"])
def csd_routes_post():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    body = request.get_json(force=True, silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": "expected JSON object"}), 400
    merge = body.get("merge")
    if merge is None and body:
        merge = body
    if not isinstance(merge, dict):
        return jsonify({"error": "missing merge"}), 400

    current = _read_csd_routes()
    n = 0
    for key, inc in merge.items():
        if not isinstance(key, str):
            continue
        k = key.strip().upper()
        if len(k) != 7 or k[3] != "-":
            continue
        a, b = k[:3], k[4:]
        if not (a.isalpha() and b.isalpha()):
            continue
        try:
            delta = int(inc)
        except (TypeError, ValueError):
            continue
        if delta <= 0:
            continue
        current[k] = int(current.get(k, 0)) + delta
        n += 1
    _atomic_write_csd(current)
    return jsonify({"ok": True, "updated": n})


@app.route("/api/phrase-usage", methods=["OPTIONS"])
def phrase_usage_options():
    return ("", 204)


@app.route("/api/phrase-usage", methods=["GET"])
def phrase_usage_get():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(_read_phrase_usage())


@app.route("/api/phrase-usage", methods=["POST"])
def phrase_usage_post():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    body = request.get_json(force=True, silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": "expected JSON object"}), 400

    allowed_keys = ("loadPlan", "advanceLoading", "handoverDetails", "offloadReason", "offloadRemarks", "other", "specialHO")
    current = _read_phrase_usage()
    n = 0

    remove = body.get("remove")
    if isinstance(remove, dict):
        for key, phrases in remove.items():
            if key not in allowed_keys:
                continue
            if not isinstance(phrases, list):
                continue
            for raw in phrases:
                p = str(raw or "").strip().upper()
                if not p or key not in current:
                    continue
                if p in current[key]:
                    current[key].pop(p, None)
                    n += 1

    merge = body.get("merge")
    if merge is None and body and "remove" not in body:
        merge = body
    if isinstance(merge, dict):
        for key, bucket in merge.items():
            if key in ("merge", "remove"):
                continue
            if key not in allowed_keys or not isinstance(bucket, dict):
                continue
            for phrase, inc in bucket.items():
                p = str(phrase or "").strip().upper()
                if not p:
                    continue
                try:
                    delta = int(inc)
                except (TypeError, ValueError):
                    continue
                if delta == 0:
                    continue
                if delta < 0:
                    cur = int(current[key].get(p, 0))
                    newv = cur + delta
                    if newv <= 0:
                        current[key].pop(p, None)
                    else:
                        current[key][p] = newv
                    n += 1
                else:
                    current[key][p] = int(current[key].get(p, 0)) + delta
                    n += 1

    _atomic_write_phrase_usage(current)
    return jsonify({"ok": True, "updated": n})


@app.route("/api/manpower-role-hints", methods=["OPTIONS"])
def manpower_role_hints_options():
    return ("", 204)


@app.route("/api/manpower-role-hints", methods=["GET"])
def manpower_role_hints_get():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(_read_manpower_role_hints())


@app.route("/api/manpower-role-hints", methods=["POST"])
def manpower_role_hints_post():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    body = request.get_json(force=True, silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": "expected JSON object"}), 400
    merge = body.get("merge")
    if merge is None and body:
        merge = body
    if not isinstance(merge, dict):
        return jsonify({"error": "missing merge"}), 400

    current = _read_manpower_role_hints()
    n = 0
    for name_raw, roles in merge.items():
        if not isinstance(name_raw, str) or not isinstance(roles, dict):
            continue
        name_label = " ".join(name_raw.strip().split())
        if not name_label:
            continue
        name_key = name_label.upper()
        if name_key not in current:
            current[name_key] = {"display": name_label, "roles": {}}
        else:
            current[name_key]["display"] = name_label

        roles_cur = current[name_key].get("roles")
        if not isinstance(roles_cur, dict):
            roles_cur = {}
            current[name_key]["roles"] = roles_cur

        for role_raw, inc in roles.items():
            if not isinstance(role_raw, str):
                continue
            role_label = " ".join(role_raw.strip().split())
            if not role_label:
                continue
            role_key = role_label.upper()
            try:
                delta = int(inc)
            except (TypeError, ValueError):
                continue
            if delta <= 0:
                continue
            cur = roles_cur.get(role_key)
            if not isinstance(cur, dict):
                cur = {"label": role_label, "count": 0}
            cur["label"] = role_label
            cur["count"] = int(cur.get("count", 0)) + delta
            roles_cur[role_key] = cur
            n += 1
    _atomic_write_manpower_role_hints(current)
    return jsonify({"ok": True, "updated": n})


@app.route("/api/recipients", methods=["OPTIONS"])
def recipients_options():
    return ("", 204)


@app.route("/api/recipients", methods=["GET"])
def recipients_get():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(_read_recipients())


@app.route("/api/recipients", methods=["POST"])
def recipients_post():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    body = request.get_json(force=True, silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": "expected JSON object"}), 400
    clean = {
        "to": _norm_email_list(body.get("to")),
        "cc": _norm_email_list(body.get("cc")),
        "bcc": _norm_email_list(body.get("bcc")),
    }
    _atomic_write_recipients(clean)
    return jsonify({"ok": True, "counts": {k: len(v) for k, v in clean.items()}})


@app.route("/api/gmail/status", methods=["GET"])
def gmail_status():
    token = _read_gmail_token()
    return jsonify(
        {
            "configured": _gmail_configured(),
            "authorized": bool(str(token.get("refresh_token") or "").strip()),
            "redirect_uri": GMAIL_REDIRECT_URI,
        }
    )


@app.route("/api/gmail/auth-url", methods=["GET"])
def gmail_auth_url():
    global _gmail_oauth_state
    if not _gmail_configured():
        return jsonify({"error": "gmail_not_configured"}), 400
    _gmail_oauth_state = secrets.token_urlsafe(24)
    params = {
        "client_id": GMAIL_CLIENT_ID,
        "redirect_uri": GMAIL_REDIRECT_URI,
        "response_type": "code",
        "scope": GMAIL_SCOPE,
        "access_type": "offline",
        "prompt": "consent",
        "state": _gmail_oauth_state,
    }
    return jsonify({"url": f"https://accounts.google.com/o/oauth2/v2/auth?{urlencode(params)}"})


@app.route("/api/gmail/oauth/callback", methods=["GET"])
def gmail_oauth_callback():
    global _gmail_oauth_state
    if not _gmail_configured():
        return jsonify({"error": "gmail_not_configured"}), 400
    code = (request.args.get("code") or "").strip()
    state = (request.args.get("state") or "").strip()
    if not code:
        return jsonify({"error": "missing_code"}), 400
    if _gmail_oauth_state and state != _gmail_oauth_state:
        return jsonify({"error": "invalid_state"}), 400
    _gmail_oauth_state = None

    payload = {
        "client_id": GMAIL_CLIENT_ID,
        "client_secret": GMAIL_CLIENT_SECRET,
        "code": code,
        "grant_type": "authorization_code",
        "redirect_uri": GMAIL_REDIRECT_URI,
    }
    try:
        r = requests.post("https://oauth2.googleapis.com/token", data=payload, timeout=20)
        if not r.ok:
            return jsonify({"error": "token_exchange_failed", "status": r.status_code, "body": r.text[:500]}), 400
        out = r.json()
    except requests.RequestException as e:
        return jsonify({"error": "token_exchange_network_error", "detail": str(e)}), 502

    keep = _read_gmail_token()
    token = {
        "access_token": out.get("access_token", ""),
        "refresh_token": out.get("refresh_token") or keep.get("refresh_token", ""),
        "scope": out.get("scope", ""),
        "token_type": out.get("token_type", ""),
    }
    _atomic_write_gmail_token(token)
    return (
        "<html><body style='font-family:Arial,sans-serif;padding:24px;'>"
        "<h2>Gmail authorization complete.</h2>"
        "<p>You can close this tab and return to the report page.</p>"
        "</body></html>"
    )


@app.route("/api/gmail/send", methods=["POST"])
def gmail_send():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    body = request.get_json(force=True, silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": "expected JSON object"}), 400

    to_list = _norm_email_list_relaxed(body.get("to"))
    cc_list = _norm_email_list_relaxed(body.get("cc"))
    bcc_list = _norm_email_list_relaxed(body.get("bcc"))
    if not to_list and not cc_list and not bcc_list:
        return jsonify({"error": "missing_recipients"}), 400

    subject = str(body.get("subject") or "Export Warehouse Activity Report").strip()
    plain = str(body.get("plain") or "").strip()
    html = str(body.get("html") or "").strip()

    access_token = _gmail_access_token()
    if not access_token:
        return jsonify({"error": "gmail_not_authorized"}), 401

    msg = MIMEMultipart("alternative")
    if to_list:
        msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    if bcc_list:
        msg["Bcc"] = ", ".join(bcc_list)
    msg["Subject"] = subject
    msg.attach(MIMEText(plain or "(No text body)", "plain", "utf-8"))
    if html:
        msg.attach(MIMEText(html, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    payload = {"raw": raw}
    try:
        r = requests.post(
            "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            data=json.dumps(payload),
            timeout=25,
        )
        if not r.ok:
            return jsonify({"error": "gmail_send_failed", "status": r.status_code, "body": r.text[:500]}), 502
        out = r.json() if r.text else {}
        return jsonify({"ok": True, "id": out.get("id")})
    except requests.RequestException as e:
        return jsonify({"error": "gmail_send_network_error", "detail": str(e)}), 502


@app.route("/api/flight-hints", methods=["POST"])
def hints_post():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    body = request.get_json(force=True, silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": "expected JSON object"}), 400
    merge = body.get("merge")
    if merge is None and body:
        merge = body
    if not isinstance(merge, dict):
        return jsonify({"error": "missing merge"}), 400

    current = _read_hints()
    n = 0
    for key, v in merge.items():
        if not isinstance(key, str) or not isinstance(v, dict):
            continue
        std = (v.get("std") or "").strip()
        dest = (v.get("destination") or "").strip()
        if not std or not dest:
            continue
        current[key] = {"std": std, "destination": dest}
        n += 1
    _atomic_write(current)
    return jsonify({"ok": True, "updated": n})


def _safe_send_file(base: Path, rel: str):
    """Serve a file under base using send_file (avoids send_from_directory quirks on some Windows setups)."""
    rel = (rel or "").replace("\\", "/").lstrip("/")
    if not rel or ".." in Path(rel).parts:
        abort(404)
    base_res = base.resolve()
    full = (base_res / rel).resolve()
    try:
        full.relative_to(base_res)
    except ValueError:
        abort(404)
    if not full.is_file():
        abort(404)
    mt, _ = mimetypes.guess_type(full.name)
    return send_file(
        os.fspath(full),
        mimetype=mt or "application/octet-stream",
        conditional=True,
        max_age=0,
    )


@app.route("/")
def index():
    return send_from_directory(os.fspath(REPORT_DIR.resolve()), "offload_report.html")


@app.route("/js/<path:rel>")
def serve_js(rel: str):
    """Explicit prefix so /js/... always resolves (avoids catch-all quirks on some setups)."""
    return _safe_send_file(ROOT / "js", rel)


@app.route("/data/report/<path:rel>")
def serve_data_report(rel: str):
    """JSON/HTML under data/report (employees.json, phrases.json, by-date/..., etc.)."""
    return _safe_send_file(REPORT_DIR, rel)


@app.route("/data/offload/<path:rel>")
def serve_data_offload(rel: str):
    """Offload pipeline JSON/HTML under data/offload."""
    return _safe_send_file(ROOT / "data" / "offload", rel)


@app.route("/<path:filename>")
def static_files(filename: str):
    """Top-level report files: offload_report.html, latest.json, dates_index.json, …"""
    filename = (filename or "").replace("\\", "/").lstrip("/")
    if filename.startswith("api/"):
        abort(404)
    # Prefer dedicated routes above for these prefixes
    if filename.startswith(_PREFIX_JS) or filename.startswith(_PREFIX_DATA_REPORT) or filename.startswith(
        _PREFIX_DATA_OFFLOAD
    ):
        abort(404)
    return _safe_send_file(REPORT_DIR, filename)


def main():
    port = int(os.environ.get("PORT", "5050"))
    host = os.environ.get("HOST", "127.0.0.1")
    sample = REPORT_DIR / "employees.json"
    print(f"[flight_hints_server] REPORT_DIR={REPORT_DIR}")
    print(f"[flight_hints_server] employees.json exists={sample.is_file()} (restart server after git pull)")
    print(f"[flight_hints_server] Test: http://{host}:{port}/api/server-info")
    print(f"[flight_hints_server] Test: http://{host}:{port}/data/report/employees.json")
    app.run(host=host, port=port, debug=os.environ.get("FLASK_DEBUG") == "1")


if __name__ == "__main__":
    main()
