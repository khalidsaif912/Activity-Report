window.flightAutocomplete = {
  flights: [],
  activeBox: null,
  activeInput: null,
  activeMatches: [],
  activeIndex: -1,
  activeOnPick: null,
  activeMode: "replace",
  _monthIndex: {
    JAN: 1,
    FEB: 2,
    MAR: 3,
    APR: 4,
    MAY: 5,
    JUN: 6,
    JUL: 7,
    AUG: 8,
    SEP: 9,
    OCT: 10,
    NOV: 11,
    DEC: 12
  },
  _monthAbbrevs: ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"],

  async load(url = "../../data/report/flights.json", options = {}) {
    const { silent = false } = options || {};
    try {
      const res = await fetch(url);
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (text.trimStart().startsWith("<")) {
        throw new Error(`Expected JSON but got HTML (wrong URL?). ${url}`);
      }
      const parsed = JSON.parse(text);
      this.flights = Array.isArray(parsed)
        ? parsed
            .map((flight) => ({
              ...flight,
              code: this.normalizeCode(flight && flight.code),
              date: this.normalizeDateKey((flight && flight.date) || ""),
              destination: this.normalizeDestination(flight && flight.destination),
              stdEtd: String((flight && flight.stdEtd) || "").trim().toUpperCase()
            }))
            .filter((flight) => flight.code)
        : [];
    } catch (err) {
      if (!silent) {
        console.error("Failed to load flights.json", url, err);
      }
      this.flights = [];
    }
  },

  normalizeCode(value) {
    return (value || "").trim().toUpperCase();
  },

  normalizeDestination(value) {
    const src = String(value || "").trim().toUpperCase();
    if (!src) return "";
    if (/^[A-Z]{3}$/.test(src)) return src;

    const fromParens = [...src.matchAll(/\(([A-Z]{3})\)/g)].map((m) => m[1]);
    if (fromParens.length) return fromParens[fromParens.length - 1];

    const tokenCodes = [...src.matchAll(/(?:^|[^A-Z])([A-Z]{3})(?=$|[^A-Z])/g)].map((m) => m[1]);
    if (tokenCodes.length) return tokenCodes[tokenCodes.length - 1];

    const lettersOnly = src.replace(/[^A-Z]/g, "");
    return lettersOnly.slice(0, 3);
  },

  formatFlight(flight) {
    const code = this.normalizeCode(flight.code);
    const date = this.normalizeDateKey(flight.date || "");
    const destination = this.normalizeDestination(flight.destination);
    return [code, date, destination].filter(Boolean).join("/");
  },

  isoToDateKey(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || "").trim());
    if (!m) return "";
    const day = parseInt(m[3], 10);
    const month = parseInt(m[2], 10);
    if (day < 1 || day > 31 || month < 1 || month > 12) return "";
    return `${day}${this._monthAbbrevs[month - 1]}`;
  },

  normalizeDateKey(value) {
    const raw = String(value || "").trim().toUpperCase().replace(/\s+/g, "");
    if (!raw) return "";

    let m = /^(\d{1,2})([A-Z]{3})$/.exec(raw);
    if (m && this._monthIndex[m[2]]) return `${parseInt(m[1], 10)}${m[2]}`;

    m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(raw);
    if (m) return this.isoToDateKey(`${m[1]}-${m[2]}-${m[3]}`);

    m = /^(\d{1,2})[-/.]([A-Z]{3})[-/.](\d{4})$/.exec(raw);
    if (m && this._monthIndex[m[2]]) return `${parseInt(m[1], 10)}${m[2]}`;

    m = /^(\d{4})[-/.]([A-Z]{3})[-/.](\d{1,2})$/.exec(raw);
    if (m && this._monthIndex[m[2]]) return `${parseInt(m[3], 10)}${m[2]}`;

    return raw;
  },

  _todayIso() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  },

  _resolveReportIso(reportIso) {
    const src = String(reportIso || "").trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(src)) return src;
    const byWindow = String(window.__flightSuggestIsoDate || "").trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(byWindow)) return byWindow;
    return this._todayIso();
  },

  _buildHintFlightsForDate(reportIso) {
    const out = [];
    const cache = window.flightHintCache;
    const memory = cache && cache._memory && typeof cache._memory === "object" ? cache._memory : null;
    if (!memory) return out;
    const prefix = `${reportIso}|`;
    const dateKey = this.isoToDateKey(reportIso);
    Object.keys(memory).forEach((k) => {
      if (!k.startsWith(prefix)) return;
      const code = this.normalizeCode(k.slice(prefix.length));
      if (!code) return;
      const meta = memory[k] || {};
      out.push({
        code,
        date: dateKey,
        destination: this.normalizeDestination(meta.destination || ""),
        stdEtd: String(meta.std || "").trim().toUpperCase()
      });
    });
    return out;
  },

  _poolForReportDate(reportIso) {
    const normalizedIso = this._resolveReportIso(reportIso);
    const reportKey = this.isoToDateKey(normalizedIso);
    const all = Array.isArray(this.flights) ? this.flights : [];

    const byDate = new Map();
    all.forEach((f) => {
      const key = this.normalizeDateKey(f.date || "");
      if (!key) return;
      if (!byDate.has(key)) byDate.set(key, []);
      byDate.get(key).push({ ...f, date: key });
    });

    // Start with exact report date.
    const pool = reportKey && byDate.has(reportKey) ? byDate.get(reportKey).slice() : [];

    const hintFlights = this._buildHintFlightsForDate(normalizedIso);
    if (hintFlights.length) {
      const seen = new Set(pool.map((f) => this.normalizeCode(f.code)));
      hintFlights.forEach((f) => {
        const c = this.normalizeCode(f.code);
        if (!c || seen.has(c)) return;
        seen.add(c);
        pool.push(f);
      });
    }

    // If report-day data is sparse/missing, augment by unique flight codes from all dates,
    // but pin displayed/stored date to reportKey to avoid showing stale calendar dates.
    if (reportKey && pool.length < 20 && all.length) {
      const seen = new Set(pool.map((f) => this.normalizeCode(f.code)));
      for (let i = 0; i < all.length; i += 1) {
        const f = all[i] || {};
        const code = this.normalizeCode(f.code);
        if (!code || seen.has(code)) continue;
        seen.add(code);
        pool.push({
          ...f,
          code,
          date: reportKey
        });
      }
    }

    return pool;
  },

  findByCode(code) {
    const q = this.normalizeCode(code);
    return this.flights.find(f => this.normalizeCode(f.code) === q) || null;
  },

  findMatches(query, options = {}) {
    const q = this.normalizeCode(query);
    const reportIso = options && options.reportIso ? String(options.reportIso) : "";
    const listAllWhenEmpty = !!(options && options.listAllWhenEmpty);
    const pool = this._poolForReportDate(reportIso);
    if (!q) {
      return listAllWhenEmpty ? pool.slice(0, 14) : [];
    }
    const digitsOnly = /^\d{1,4}$/.test(q);
    return pool
      .filter(f => {
        const code = this.normalizeCode(f.code);
        if (digitsOnly) {
          const num = code.replace(/^[A-Z]{2}/, "");
          return num.startsWith(q) || code.includes(q);
        }
        return code.startsWith(q);
      })
      .slice(0, 8);
  },

  ensureBox() {
    if (this.activeBox) return this.activeBox;

    const box = document.createElement("div");
    box.id = "flight-suggest-box";
    box.style.position = "absolute";
    box.style.zIndex = "9999";
    box.style.background = "#fff";
    box.style.border = "1px solid #cbd5e1";
    box.style.boxShadow = "0 8px 24px rgba(0,0,0,0.12)";
    box.style.maxHeight = "220px";
    box.style.overflowY = "auto";
    box.style.fontSize = "14px";
    box.style.display = "none";
    document.body.appendChild(box);

    this.activeBox = box;
    return box;
  },

  hideBox() {
    const box = this.ensureBox();
    box.style.display = "none";
    box.innerHTML = "";
    this.activeMatches = [];
    this.activeIndex = -1;
    this.activeInput = null;
    this.activeOnPick = null;
    this.activeMode = "replace";
  },

  positionBox(input) {
    const box = this.ensureBox();
    const rect = input.getBoundingClientRect();
    const desiredWidth = Math.max(rect.width, 260);
    const maxLeft = window.scrollX + window.innerWidth - desiredWidth - 8;
    const left = Math.max(window.scrollX + 8, Math.min(window.scrollX + rect.left, maxLeft));
    box.style.left = `${left}px`;
    box.style.top = `${window.scrollY + rect.bottom + 4}px`;
    box.style.width = `${desiredWidth}px`;
  },

  renderBox(input, matches, onPick, mode = "replace") {
    const box = this.ensureBox();

    if (!matches.length) {
      this.hideBox();
      return;
    }

    this.activeInput = input;
    this.activeMatches = matches;
    this.activeIndex = 0;
    this.activeOnPick = onPick;
    this.activeMode = mode;

    this.positionBox(input);
    box.innerHTML = "";

    matches.forEach((flight, index) => {
      const item = document.createElement("div");
      item.style.padding = "8px 10px";
      item.style.cursor = "pointer";
      item.style.borderBottom = "1px solid #e2e8f0";
      item.style.background = index === 0 ? "#eff6ff" : "#fff";
      item.style.whiteSpace = "nowrap";
      item.textContent = this.formatFlight(flight);

      item.addEventListener("mouseenter", () => {
        this.activeIndex = index;
        this.refreshHighlight();
      });

      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        this.pickActive(index);
      });

      box.appendChild(item);
    });

    box.style.display = "block";
  },

  refreshHighlight() {
    const box = this.ensureBox();
    [...box.children].forEach((child, index) => {
      child.style.background = index === this.activeIndex ? "#eff6ff" : "#fff";
    });
  },

  pickActive(index = null) {
    if (!this.activeMatches.length || !this.activeInput || !this.activeOnPick) return;
    const pickedIndex = index === null ? this.activeIndex : index;
    const picked = this.activeMatches[pickedIndex];
    if (!picked) return;

    this.activeOnPick(picked, this.activeMode);
    this.hideBox();
  },

  /** Flight-ish token at line end; supports code+number (WY223) and number-only (223). */
  lastFlightChunkRe: /(?:^|[^A-Z0-9])([A-Z]{2}\d{0,4}|\d{1,4})$/i,

  replaceLastToken(text, replacement) {
    const t = (text || "").toUpperCase();
    const m = this.lastFlightChunkRe.exec(t);
    if (m) {
      const token = m[1];
      const keepLen = m[0].length - token.length;
      return t.slice(0, m.index + keepLen) + replacement;
    }
    const trimmed = t.trimEnd();
    if (!trimmed) return replacement;
    return `${trimmed} ${replacement}`;
  },

  extractLastToken(text) {
    const m = this.lastFlightChunkRe.exec((text || "").toUpperCase());
    return m ? m[1].toUpperCase() : "";
  },

  attach(input, key, onPick) {
    input.addEventListener("input", () => {
      input.value = this.normalizeCode(input.value);
      const matches = this.findMatches(input.value, { reportIso: window.__flightSuggestIsoDate });
      this.renderBox(input, matches, (picked) => {
        onPick(picked);
      }, "replace");
    });

    input.addEventListener(
      "keydown",
      (e) => {
        if (this.activeInput !== input || !this.activeMatches.length) return;

        if (e.key === "ArrowDown") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.min(this.activeIndex + 1, this.activeMatches.length - 1);
          this.refreshHighlight();
        }

        if (e.key === "ArrowUp") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.max(this.activeIndex - 1, 0);
          this.refreshHighlight();
        }

        if (e.key === "Enter") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.pickActive();
        }
      },
      true
    );

    input.addEventListener("blur", () => {
      setTimeout(() => {
        if (this.activeInput === input) this.hideBox();
      }, 150);
    });
  },

  attachInlineInput(input, onSave) {
    if (input.dataset.flightInlineBound === "1") return;
    input.dataset.flightInlineBound = "1";

    input.addEventListener("input", () => {
      input.value = input.value.toUpperCase();
      if (onSave) onSave(input.value);

      const token = this.extractLastToken(input.value);
      const matches = this.findMatches(token, { reportIso: window.__flightSuggestIsoDate });

      this.renderBox(input, matches, (picked) => {
        const value = this.replaceLastToken(input.value.toUpperCase(), this.formatFlight(picked));
        input.value = value;
        if (onSave) onSave(value);
      }, "inline");
    });

    input.addEventListener(
      "keydown",
      (e) => {
        if (this.activeInput !== input || !this.activeMatches.length) return;

        if (e.key === "ArrowDown") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.min(this.activeIndex + 1, this.activeMatches.length - 1);
          this.refreshHighlight();
        }

        if (e.key === "ArrowUp") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.max(this.activeIndex - 1, 0);
          this.refreshHighlight();
        }

        if (e.key === "Enter") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.pickActive();
        }
      },
      true
    );

    input.addEventListener("blur", () => {
      setTimeout(() => {
        if (this.activeInput === input) this.hideBox();
      }, 150);
    });
  },

  attachInlineTextarea(textarea, onSave) {
    if (!textarea || textarea.dataset.flightInlineBound === "1") return;
    textarea.dataset.flightInlineBound = "1";

    textarea.addEventListener("input", () => {
      textarea.value = textarea.value.toUpperCase();
      if (onSave) onSave(textarea.value);

      const beforeCursor = textarea.value.slice(0, textarea.selectionStart);
      const token = this.extractLastToken(beforeCursor);
      const matches = this.findMatches(token, { reportIso: window.__flightSuggestIsoDate });

      this.renderBox(textarea, matches, (picked) => {
        const start = textarea.selectionStart;
        const before = textarea.value.slice(0, start);
        const after = textarea.value.slice(start);
        const replacedBefore = this.replaceLastToken(before.toUpperCase(), this.formatFlight(picked));
        textarea.value = replacedBefore + after;
        const newPos = replacedBefore.length;
        textarea.selectionStart = newPos;
        textarea.selectionEnd = newPos;
        if (onSave) onSave(textarea.value);
      }, "inline");
    });

    textarea.addEventListener(
      "keydown",
      (e) => {
        if (this.activeInput !== textarea || !this.activeMatches.length) return;

        if (e.key === "ArrowDown") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.min(this.activeIndex + 1, this.activeMatches.length - 1);
          this.refreshHighlight();
        }

        if (e.key === "ArrowUp") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.max(this.activeIndex - 1, 0);
          this.refreshHighlight();
        }

        if (e.key === "Enter") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.pickActive();
        }
      },
      true
    );

    textarea.addEventListener("blur", () => {
      setTimeout(() => {
        if (this.activeInput === textarea) this.hideBox();
      }, 150);
    });
  }
};

window.addEventListener("resize", () => {
  if (window.flightAutocomplete && window.flightAutocomplete.activeInput && window.flightAutocomplete.activeMatches.length) {
    window.flightAutocomplete.positionBox(window.flightAutocomplete.activeInput);
  }
});

window.addEventListener("scroll", () => {
  if (window.flightAutocomplete && window.flightAutocomplete.activeInput && window.flightAutocomplete.activeMatches.length) {
    window.flightAutocomplete.positionBox(window.flightAutocomplete.activeInput);
  }
}, true);