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

Optional env FLIGHT_HINTS_TOKEN: if set, clients must send header:
  X-Flight-Hints-Token: <token>
for GET and POST.
"""

from __future__ import annotations

import json
import mimetypes
import os
from pathlib import Path

from flask import Flask, abort, jsonify, request, send_file, send_from_directory

ROOT = Path(__file__).resolve().parent.parent
REPORT_DIR = ROOT / "data" / "report"
HINTS_FILE = REPORT_DIR / "flight-hints.json"
CSD_ROUTES_FILE = REPORT_DIR / "csd-route-hints.json"
PHRASE_USAGE_FILE = REPORT_DIR / "phrase-usage.json"
MANPOWER_ROLE_HINTS_FILE = REPORT_DIR / "manpower-role-hints.json"

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
        return {"loadPlan": {}, "advanceLoading": {}, "offloadReason": {}, "offloadRemarks": {}, "other": {}, "specialHO": {}}
    try:
        raw = json.loads(PHRASE_USAGE_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return {"loadPlan": {}, "advanceLoading": {}, "offloadReason": {}, "offloadRemarks": {}, "other": {}, "specialHO": {}}
        out = {"loadPlan": {}, "advanceLoading": {}, "offloadReason": {}, "offloadRemarks": {}, "other": {}, "specialHO": {}}
        for key in ("loadPlan", "advanceLoading", "offloadReason", "offloadRemarks", "other", "specialHO"):
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
        return {"loadPlan": {}, "advanceLoading": {}, "offloadReason": {}, "offloadRemarks": {}, "other": {}, "specialHO": {}}


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
        }
    )


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
    merge = body.get("merge")
    if merge is None and body:
        merge = body
    if not isinstance(merge, dict):
        return jsonify({"error": "missing merge"}), 400

    current = _read_phrase_usage()
    n = 0
    allowed_keys = ("loadPlan", "advanceLoading", "offloadReason", "offloadRemarks", "other", "specialHO")
    for key, bucket in merge.items():
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
            if delta <= 0:
                continue
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
