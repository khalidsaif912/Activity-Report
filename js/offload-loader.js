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
    if (!flightCode) return false;

    const flight = this.findFlightForReport(flightCode, reportIso);
    if (!flight) return false;

    const mins = this.parseStdMinutes(flight.stdEtd || "");
    return this.flightStdMatchesShift(mins, shift);
  },

  resetOffloadsBlank(state) {
    const dateStr = this.normalizeOffloadDate(state.shiftMeta.date || state.activeDate || "");
    state.offloads = [
      {
        item: 1,
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
      this.resetOffloadsBlank(state);
      return;
    }

    if (!this.offloadFlightMatchesActiveShift(data, state)) {
      this.resetOffloadsBlank(state);
      return;
    }

    const headerDate = this.normalizeOffloadDate(data.date || state.shiftMeta.date || "");
    const flightRow = this.findFlightForReport(data.flight || "", reportIso);
    const stdStr =
      flightRow && flightRow.stdEtd
        ? String(flightRow.stdEtd)
            .split("/")[0]
            .trim()
        : "";

    state.offloads = cleanItems.map((item, index) => ({
      item: index + 1,
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
  }
};
