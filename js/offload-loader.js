window.offloadLoader = {
  _monthAbbrevs: ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"],
  _monthIndex: {
    JAN: 0,
    FEB: 1,
    MAR: 2,
    APR: 3,
    MAY: 4,
    JUN: 5,
    JUL: 6,
    AUG: 7,
    SEP: 8,
    OCT: 9,
    NOV: 10,
    DEC: 11
  },

  normalizeOffloadDate(value) {
    if (value == null || typeof value !== "string") return value || "";
    const s = value.trim().replace(/\s+/g, " ");
    const alpha = s.match(/^(\d{4})-([A-Za-z]{3})-(\d{1,2})$/);
    if (alpha) {
      const [, year, mon, day] = alpha;
      const d = String(parseInt(day, 10)).padStart(2, "0");
      return `${d}-${mon.toUpperCase()}-${year}`;
    }
    const num = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (num) {
      const [, year, mo, day] = num;
      const mi = parseInt(mo, 10);
      if (mi >= 1 && mi <= 12) {
        const mon = this._monthAbbrevs[mi - 1];
        const d = String(parseInt(day, 10)).padStart(2, "0");
        return `${d}-${mon}-${year}`;
      }
    }
    return s;
  },

  /** Convert report ISO date (YYYY-MM-DD) to keys like 19APR for flights.json */
  isoToFlightDateKey(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((iso || "").trim());
    if (!m) return "";
    const day = parseInt(m[3], 10);
    const mi = parseInt(m[2], 10);
    if (mi < 1 || mi > 12) return "";
    return `${day}${this._monthAbbrevs[mi - 1]}`;
  },

  /**
   * Parse offload header date from SharePoint JSON into YYYY-MM-DD when possible.
   * Handles DD-MMM-YYYY, YYYY-MM-DD, DD.MMM, DD/MMM (year from report or current year).
   */
  parseOffloadSourceDateToIso(raw, reportIso) {
    if (raw == null || typeof raw !== "string") return null;
    const s = raw.trim().replace(/\s+/g, " ");
    if (!s) return null;

    const iso = (reportIso || "").trim();
    let y =
      iso && /^(\d{4})-\d{2}-\d{2}$/.test(iso) ? parseInt(iso.slice(0, 4), 10) : new Date().getFullYear();

    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;

    m = s.match(/^(\d{1,2})[-/]([A-Za-z]{3})[-/](\d{4})$/i);
    if (m) {
      const mon = m[2].toUpperCase();
      const mi = this._monthIndex[mon];
      if (mi == null) return null;
      const d = String(parseInt(m[1], 10)).padStart(2, "0");
      return `${m[3]}-${String(mi + 1).padStart(2, "0")}-${d}`;
    }

    m = s.match(/^(\d{1,2})[./]([A-Za-z]{3})$/i);
    if (m) {
      const mon = m[2].toUpperCase();
      const mi = this._monthIndex[mon];
      if (mi == null) return null;
      const d = String(parseInt(m[1], 10)).padStart(2, "0");
      return `${y}-${String(mi + 1).padStart(2, "0")}-${d}`;
    }

    m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/i);
    if (m) {
      const mon = m[2].toUpperCase();
      const mi = this._monthIndex[mon];
      if (mi == null) return null;
      const d = String(parseInt(m[1], 10)).padStart(2, "0");
      return `${m[3]}-${String(mi + 1).padStart(2, "0")}-${d}`;
    }

    const norm = this.normalizeOffloadDate(s);
    m = /^(\d{2})-([A-Z]{3})-(\d{4})$/.exec(norm);
    if (m) {
      const mi = this._monthIndex[m[2]];
      if (mi == null) return null;
      return `${m[3]}-${String(mi + 1).padStart(2, "0")}-${m[1]}`;
    }

    return null;
  },

  offloadDateMatchesReport(offloadDateRaw, reportIso) {
    if (!reportIso || !/^\d{4}-\d{2}-\d{2}$/.test(reportIso.trim())) return true;
    const parsed = this.parseOffloadSourceDateToIso(offloadDateRaw, reportIso.trim());
    if (!parsed) return false;
    return parsed === reportIso.trim();
  },

  /** STD/ETD first segment → minutes from midnight; null if unknown */
  parseStdMinutes(stdEtd) {
    if (stdEtd == null || typeof stdEtd !== "string") return null;
    const first = stdEtd.split("/")[0].trim().replace(/\s/g, "");
    if (!first) return null;
    const digits = first.replace(/\D/g, "");
    if (digits.length < 3) return null;
    const pad = digits.padStart(4, "0").slice(-4);
    const hh = parseInt(pad.slice(0, 2), 10);
    const mm = parseInt(pad.slice(2, 4), 10);
    if (hh > 23 || mm > 59) return null;
    return hh * 60 + mm;
  },

  /**
   * Aligns with roster shift windows: morning 06–15, afternoon 13–22, night 21–06.
   * Afternoon excludes departures before 13:00 so morning flights are not shown on afternoon.
   */
  flightStdMatchesShift(stdMinutes, activeShift) {
    if (stdMinutes == null) return false;
    const t = stdMinutes;
    if (activeShift === "morning") {
      return t >= 6 * 60 && t < 15 * 60;
    }
    if (activeShift === "afternoon") {
      return t >= 13 * 60 && t < 22 * 60;
    }
    if (activeShift === "night") {
      return t >= 21 * 60 || t < 6 * 60;
    }
    return true;
  },

  findFlightForReport(flightCode, reportIso) {
    const fa = window.flightAutocomplete;
    if (!fa || !Array.isArray(fa.flights) || !fa.flights.length) return null;
    const code = fa.normalizeCode(flightCode);
    if (!code) return null;
    const key = this.isoToFlightDateKey(reportIso || "");
    const normalizedKey =
      typeof fa.normalizeDateKey === "function" ? fa.normalizeDateKey(key || "") : String(key || "").toUpperCase().replace(/\s/g, "");
    const sameDay = fa.flights.filter(
      (f) =>
        fa.normalizeCode(f.code) === code &&
        (typeof fa.normalizeDateKey === "function"
          ? fa.normalizeDateKey(f.date || "") === normalizedKey
          : (f.date || "").toUpperCase().replace(/\s/g, "") === normalizedKey)
    );
    if (sameDay.length === 1) return sameDay[0];
    if (sameDay.length > 1) return sameDay[0];
    return fa.findByCode(code);
  },

  offloadFlightMatchesActiveShift(data, state) {
    const shift = state.activeShift;
    if (!shift || !state.shiftsFromServer) return true;

    const reportIso = (state.activeDate || state.shiftMeta.date || "").trim();
    const flightCode = (data.flight || "").trim();
    // Do not block the table when the source flight code is missing.
    // The row can still be useful for manual completion.
    if (!flightCode) return true;

    const flight = this.findFlightForReport(flightCode, reportIso);
    // Fail-open: if flights cache is missing/outdated, keep offload rows visible.
    if (!flight) return true;

    const mins = this.parseStdMinutes(flight.stdEtd || "");
    // If STD is unavailable, do not hide rows.
    if (mins == null) return true;
    return this.flightStdMatchesShift(mins, shift);
  },

  resetOffloadsBlank(state) {
    const dateStr = this.normalizeOffloadDate(state.shiftMeta.date || state.activeDate || "");
    state.offloads = [
      {
        item: 1,
        awb: "",
        date: dateStr,
        flight: "",
        std: "",
        destination: "",
        emailTime: "",
        rampReceived: "",
        trolley: "",
        cmsCompleted: "",
        piecesVerification: "",
        reason: "",
        remarks: ""
      }
    ];
  },

  isMeaningfulOffloadRow(row) {
    if (!row || typeof row !== "object") return false;
    return [
      row.awb,
      row.flight,
      row.std,
      row.destination,
      row.emailTime,
      row.rampReceived,
      row.trolley,
      row.cmsCompleted,
      row.piecesVerification,
      row.reason,
      row.remarks
    ].some((value) => String(value || "").trim());
  },

  normalizeOffloadRow(row, defaultDate) {
    const out = {
      item: 0,
      awb: String((row && row.awb) || "").trim(),
      date: this.normalizeOffloadDate((row && row.date) || defaultDate || ""),
      flight: String((row && row.flight) || "").trim().toUpperCase(),
      std: String((row && row.std) || "").trim().toUpperCase(),
      destination: String((row && row.destination) || "").trim().toUpperCase(),
      emailTime: String((row && row.emailTime) || "").trim(),
      rampReceived: String((row && row.rampReceived) || "").trim(),
      trolley: String((row && row.trolley) || "").trim(),
      cmsCompleted: String((row && row.cmsCompleted) || "").trim(),
      piecesVerification: String((row && row.piecesVerification) || "").trim(),
      reason: String((row && row.reason) || "").trim(),
      remarks: String((row && row.remarks) || "").trim()
    };
    return out;
  },

  /**
   * Parses "13 PCS / 271 KGS" style strings (spacing/case tolerant).
   */
  parsePiecesVerification(value) {
    const s = String(value ?? "").trim();
    if (!s) return null;
    const m = s.match(/(\d+)\s*PCS\s*\/\s*(\d+)\s*KGS/i);
    if (!m) return null;
    const pcs = parseInt(m[1], 10);
    const kgs = parseInt(m[2], 10);
    if (Number.isNaN(pcs) || Number.isNaN(kgs)) return null;
    return { pcs, kgs };
  },

  formatPiecesVerification(pcs, kgs) {
    return `${pcs} PCS / ${kgs} KGS`;
  },

  mergePiecesVerification(a, b) {
    const pa = this.parsePiecesVerification(a);
    const pb = this.parsePiecesVerification(b);
    if (pa && pb) return this.formatPiecesVerification(pa.pcs + pb.pcs, pa.kgs + pb.kgs);
    if (pa && !String(b ?? "").trim()) return a;
    if (pb && !String(a ?? "").trim()) return b;
    const ta = String(a ?? "").trim();
    const tb = String(b ?? "").trim();
    if (!ta) return tb;
    if (!tb) return ta;
    if (ta === tb) return ta;
    return `${ta} + ${tb}`;
  },

  /**
   * Dedupe key: same flight leg (date + flight + destination). AWB/reason can differ
   * across SharePoint rows for one flight; those rows should merge into one table row.
   */
  offloadRowKey(row, defaultDate) {
    const normalized = this.normalizeOffloadRow(row, defaultDate);
    const flight = normalized.flight;
    const date = normalized.date;
    const destination = normalized.destination;
    if (!flight || !date) return "";
    if (destination) return `${flight}|${date}|${destination}`;
    return `${flight}|${date}`;
  },

  _mergeOffloadInto(existing, incoming) {
    existing.awb = (() => {
      const a = String(existing.awb || "").trim();
      const b = String(incoming.awb || "").trim();
      if (!b) return a;
      if (!a) return b;
      if (a.toUpperCase() === b.toUpperCase()) return a;
      return `${a}, ${b}`;
    })();
    existing.date = incoming.date || existing.date;
    existing.flight = incoming.flight || existing.flight;
    existing.std = existing.std || incoming.std;
    existing.destination = incoming.destination || existing.destination;
    existing.piecesVerification = this.mergePiecesVerification(existing.piecesVerification, incoming.piecesVerification);
    existing.reason = (() => {
      const a = String(existing.reason || "").trim();
      const b = String(incoming.reason || "").trim();
      if (!b) return a;
      if (!a) return b;
      if (a.toUpperCase() === b.toUpperCase()) return a;
      return `${a} / ${b}`;
    })();
    if (!existing.emailTime && incoming.emailTime) existing.emailTime = incoming.emailTime;
    if (!existing.rampReceived && incoming.rampReceived) existing.rampReceived = incoming.rampReceived;
    if (!existing.trolley && incoming.trolley) existing.trolley = incoming.trolley;
    if (!existing.cmsCompleted && incoming.cmsCompleted) existing.cmsCompleted = incoming.cmsCompleted;
    if (incoming.remarks) {
      const er = String(existing.remarks || "").trim();
      const ir = String(incoming.remarks || "").trim();
      if (!er) existing.remarks = ir;
      else if (ir && er.toUpperCase() !== ir.toUpperCase()) existing.remarks = `${er} | ${ir}`;
    }
  },

  resequenceRows(rows) {
    return (Array.isArray(rows) ? rows : []).map((row, index) => ({
      ...row,
      item: index + 1
    }));
  },

  mergedOffloadRows(existingRows, incomingRows, defaultDate) {
    const out = [];
    const indexByKey = new Map();

    const upsertRow = (normalized) => {
      const key = this.offloadRowKey(normalized, defaultDate);
      if (!key) {
        out.push({ ...normalized });
        return;
      }
      const idx = indexByKey.get(key);
      if (idx == null) {
        indexByKey.set(key, out.length);
        out.push({ ...normalized });
        return;
      }
      this._mergeOffloadInto(out[idx], normalized);
    };

    const addExisting = (rows) => {
      (Array.isArray(rows) ? rows : []).forEach((row) => {
        if (!this.isMeaningfulOffloadRow(row)) return;
        upsertRow(this.normalizeOffloadRow(row, defaultDate));
      });
    };

    const mergeIncoming = (rows) => {
      (Array.isArray(rows) ? rows : []).forEach((row) => {
        if (!this.isMeaningfulOffloadRow(row)) return;
        upsertRow(this.normalizeOffloadRow(row, defaultDate));
      });
    };

    addExisting(existingRows);
    mergeIncoming(incomingRows);

    if (!out.length) {
      return this.resequenceRows([
        {
          item: 1,
          awb: "",
          date: this.normalizeOffloadDate(defaultDate || ""),
          flight: "",
          std: "",
          destination: "",
          emailTime: "",
          rampReceived: "",
          trolley: "",
          cmsCompleted: "",
          piecesVerification: "",
          reason: "",
          remarks: ""
        }
      ]);
    }

    return this.resequenceRows(out);
  },

  normalizeOffloadRows(state) {
    if (!state || !Array.isArray(state.offloads)) return;
    state.offloads.forEach((row) => {
      row.date = this.normalizeOffloadDate(row.date || "");
    });
  },

  async load(url = "../../data/offload/report/latest.json") {
    try {
      const res = await fetch(url);
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (text.trimStart().startsWith("<")) {
        throw new Error(`Expected JSON but got HTML (wrong URL?). ${url}`);
      }
      return JSON.parse(text);
    } catch (err) {
      console.error("Failed to load offload json", url, err);
      return null;
    }
  },

  applyToState(data, state) {
    if (!data) return;

    const items = Array.isArray(data.items) ? data.items : [];
    const cleanItems = items.filter(item => (item.awb || "").toUpperCase() !== "TOTAL");

    if (!cleanItems.length) return;

    const reportIso = (state.activeDate || state.shiftMeta.date || "").trim();
    /*
     * Keep source-of-truth loading from ShareFolder, but enforce:
     * 1) report date match
     * 2) active shift period match
     */
    if (reportIso && !this.offloadDateMatchesReport(data.date || "", reportIso)) {
      // Keep loading rows even when source date does not match selected report date.
      // This avoids hard-empty table states when ShareFolder latest file is stale.
      console.warn("[offload] Source date mismatch; using source rows with selected report context.", {
        sourceDate: data.date || "",
        reportDate: reportIso
      });
    }

    if (!this.offloadFlightMatchesActiveShift(data, state)) {
      console.warn("[offload] Flight did not match active shift; skipping source import for this shift.");
      return;
    }

    const headerDate = this.normalizeOffloadDate(state.shiftMeta.date || state.activeDate || data.date || "");
    const flightRow = this.findFlightForReport(data.flight || "", reportIso);
    const stdStr =
      flightRow && flightRow.stdEtd
        ? String(flightRow.stdEtd)
            .split("/")[0]
            .trim()
        : "";

    const importedRows = cleanItems.map((item, index) => ({
      item: index + 1,
      awb: String(item.awb || "").trim(),
      date: headerDate,
      flight: data.flight || "",
      std: stdStr,
      destination: data.destination || "",
      emailTime: "",
      rampReceived: "",
      trolley: "",
      cmsCompleted: "",
      piecesVerification: `${item.pcs || ""} PCS / ${item.kgs || ""} KGS`,
      reason: item.reason || "",
      remarks: ""
    }));

    state.offloads = this.mergedOffloadRows(state.offloads, importedRows, headerDate);
  }
};
