window.phraseAutocomplete = {
  phrases: {},
  /** IATA-oriented lines for CSD Rescreening (Oman Air WY + Salam OV); editable JSON. */
  csdDestinationHints: [],
  activeBox: null,
  activeInput: null,
  activeItems: [],
  activeIndex: -1,
  activeOnPick: null,

  async load(url = "../../data/report/phrases.json") {
    try {
      const res = await fetch(url);
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (text.trimStart().startsWith("<")) {
        throw new Error(`Expected JSON but got HTML (wrong URL?). ${url}`);
      }
      this.phrases = JSON.parse(text);
    } catch (err) {
      console.error("Failed to load phrases.json", url, err);
      this.phrases = {};
    }
    this._advancePhraseExcludeSetCache = null;
  },

  async loadCsdDestinationHints(url = "../../data/report/csd-wy-ov-destinations.json") {
    try {
      const res = await fetch(url);
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (text.trimStart().startsWith("<")) {
        throw new Error(`Expected JSON but got HTML (wrong URL?). ${url}`);
      }
      const data = JSON.parse(text);
      this.csdDestinationHints = Array.isArray(data)
        ? data
            .map((x) => {
              const s = String(x || "").trim().toUpperCase();
              const m = s.match(/^([A-Z]{3})\b/);
              return m ? m[1] : null;
            })
            .filter(Boolean)
        : [];
    } catch (err) {
      console.error("Failed to load csd-wy-ov-destinations.json", url, err);
      this.csdDestinationHints = [];
    }
  },

  /** Today’s flights.json — 3-letter destination codes only (deduped). Empty query = all (up to limit). */
  _csdFlightDestinationHints(query) {
    const q = (query || "").trim().toUpperCase();
    try {
      const flights = window.flightAutocomplete && window.flightAutocomplete.flights;
      const normalizeDest =
        window.flightAutocomplete && typeof window.flightAutocomplete.normalizeDestination === "function"
          ? window.flightAutocomplete.normalizeDestination.bind(window.flightAutocomplete)
          : (v) => String(v || "").trim().toUpperCase();
      if (!Array.isArray(flights)) return [];
      const out = [];
      const seen = new Set();
      for (const f of flights) {
        const code3 = normalizeDest(f.destination);
        if (!code3) continue;
        if (q && !code3.includes(q)) continue;
        if (seen.has(code3)) continue;
        seen.add(code3);
        out.push(code3);
      }
      return out.slice(0, 8);
    } catch {
      return [];
    }
  },

  _advancePhraseExcludeSet() {
    if (!this._advancePhraseExcludeSetCache) {
      this._advancePhraseExcludeSetCache = new Set(
        this.getList("advanceLoading").map((s) => String(s || "").toUpperCase().trim())
      );
    }
    return this._advancePhraseExcludeSetCache;
  },

  /** CSD-only phrase lines — never surface Advance Loading boilerplate here. */
  _csdRescreeningPhrases(list) {
    const ex = this._advancePhraseExcludeSet();
    return (list || []).filter((item) => !ex.has(String(item || "").toUpperCase().trim()));
  },

  /**
   * Policy / AWB block (digits + hyphens) then space → remainder drives destination suggestions.
   * Until a space after the number exists, suggestions stay hidden while the cursor is still in the number.
   * Lines without a leading policy block (e.g. only "ZRH-CGK") still get suggestions on the full text.
   */
  csdTailAfterPolicy(value) {
    const v = String(value ?? "");
    const trimmed = v.trim();
    const m = v.match(/^(\d[\d\-]*)\s+(.*)$/);
    if (m) return { ready: true, tail: (m[2] || "").trim().toUpperCase() };
    if (trimmed && /^\d[\d\-]*$/.test(trimmed)) return { ready: false, tail: "" };
    return { ready: true, tail: trimmed.toUpperCase() };
  },

  mergePhrasePickCsd(fullValue, pickedPhrase) {
    const v = String(fullValue ?? "");
    const m = v.match(/^(\d[\d\-]*)(\s+)(.*)$/);
    if (m) {
      return m[1] + m[2] + (pickedPhrase || "").toUpperCase();
    }
    return this.mergePhrasePick(fullValue, pickedPhrase);
  },

  getList(key) {
    return Array.isArray(this.phrases[key]) ? this.phrases[key] : [];
  },

  _isLearnedPhraseKey(key) {
    return (
      key === "loadPlan" ||
      key === "advanceLoading" ||
      key === "handoverDetails" ||
      key === "offloadReason" ||
      key === "offloadRemarks" ||
      key === "other" ||
      key === "specialHO"
    );
  },

  _isFlightAwarePhraseKey(key) {
    return (
      key === "loadPlan" ||
      key === "advanceLoading" ||
      key === "handoverDetails" ||
      key === "other" ||
      key === "specialHO"
    );
  },

  _isLineScopedTextareaKey(key) {
    return (
      key === "handoverDetails" ||
      key === "other" ||
      key === "specialHO" ||
      key === "offloadReason" ||
      key === "offloadRemarks"
    );
  },

  /** Last token as flight code prefix (e.g. WY, WY1, WY101). */
  _flightPrefixQuery(value) {
    const parts = String(value || "")
      .trim()
      .toUpperCase()
      .split(/\s+/);
    const tok = parts.length ? parts[parts.length - 1] : "";
    if (!/^(?:[A-Z]{2}\d{0,4}|\d{1,4})$/.test(tok)) return "";
    return tok;
  },

  _normPick(picked) {
    if (picked == null) return "";
    if (typeof picked === "string") return picked;
    if (typeof picked === "object" && picked.text != null) return String(picked.text);
    return String(picked);
  },

  _mergeFlightPickIntoValue(value, pickedPhrase) {
    const v = String(value || "").toUpperCase();
    const p = String(pickedPhrase || "").toUpperCase();
    const fq = this._flightPrefixQuery(v);
    if (!fq || !p) return null;
    const i = v.lastIndexOf(fq);
    if (i < 0) return null;
    return v.slice(0, i) + p + v.slice(i + fq.length);
  },

  _getTextareaActiveLine(textarea) {
    const value = String(textarea && textarea.value ? textarea.value : "");
    const caret = typeof textarea.selectionStart === "number" ? textarea.selectionStart : value.length;
    const start = value.lastIndexOf("\n", Math.max(0, caret - 1)) + 1;
    const endIdx = value.indexOf("\n", caret);
    const end = endIdx < 0 ? value.length : endIdx;
    return {
      value,
      start,
      end,
      line: value.slice(start, end)
    };
  },

  _replaceTextareaActiveLine(textarea, nextLine) {
    const ctx = this._getTextareaActiveLine(textarea);
    const left = ctx.value.slice(0, ctx.start);
    const right = ctx.value.slice(ctx.end);
    const line = String(nextLine || "").toUpperCase();
    const next = left + line + right;
    textarea.value = next;
    const pos = left.length + line.length;
    try {
      textarea.selectionStart = textarea.selectionEnd = pos;
    } catch (_) {}
    return next;
  },

  _tailTokenForFlightAwareLine(line) {
    const u = String(line || "").toUpperCase();
    if (!u.trim()) return "";
    if (/\s$/.test(u)) return "";
    const parts = u.trimEnd().split(/\s+/);
    return parts.length ? parts[parts.length - 1] : "";
  },

  _mergeFlightAwareLine(line, picked) {
    const src = String(line || "").toUpperCase();
    const p = String(picked || "").toUpperCase();
    if (!src.trim()) return p;
    if (/\s$/.test(src)) return src + p;
    return src.replace(/\S+$/, p);
  },

  /**
   * @param {string} flightQuery — prefix of flight code (e.g. WY1), or "".
   * @param {boolean} [listAllWhenNoFlightToken] — if true and flightQuery is "", return all of today’s flights (after space / new segment).
   */
  _flightSuggestionStrings(flightQuery, listAllWhenNoFlightToken) {
    const q = (flightQuery || "").trim().toUpperCase();
    const fa = window.flightAutocomplete;
    if (!fa || !Array.isArray(fa.flights) || !fa.flights.length) return [];

    let iso = "";
    try {
      const w = window.__flightSuggestIsoDate;
      if (w && /^\d{4}-\d{2}-\d{2}$/.test(String(w))) iso = String(w).trim();
    } catch (_) {}
    if (!iso) {
      const d = new Date();
      iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    }

    let dateKey = "";
    if (window.offloadLoader && typeof window.offloadLoader.isoToFlightDateKey === "function") {
      dateKey = String(window.offloadLoader.isoToFlightDateKey(iso) || "")
        .toUpperCase()
        .replace(/\s/g, "");
    }

    const poolAll = fa.flights.slice();
    let pool = poolAll;
    if (dateKey) {
      const sameDay = poolAll.filter((f) => String(f.date || "").toUpperCase().replace(/\s/g, "") === dateKey);
      if (sameDay.length) {
        pool = sameDay;
      } else if (window.offloadLoader && typeof window.offloadLoader.isoToFlightDateKey === "function") {
        /*
         * If report date has no flights, fall back to browser "today" flights so
         * Load Plan remains useful during date-index/report-date lag.
         */
        const d = new Date();
        const todayIso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
        const todayKey = String(window.offloadLoader.isoToFlightDateKey(todayIso) || "")
          .toUpperCase()
          .replace(/\s/g, "");
        pool = todayKey
          ? poolAll.filter((f) => String(f.date || "").toUpperCase().replace(/\s/g, "") === todayKey)
          : [];
      } else {
        pool = [];
      }
    }

    const fmt = (f) =>
      typeof fa.formatFlight === "function"
        ? fa.formatFlight(f)
        : [f.code, f.date, f.destination].filter(Boolean).join("/");

    if (!q) {
      if (listAllWhenNoFlightToken) {
        return pool.slice(0, 14).map(fmt);
      }
      return [];
    }

    const digitsOnly = /^\d{1,4}$/.test(q);
    pool = pool.filter((f) => {
      const code = fa.normalizeCode(f.code);
      if (digitsOnly) {
        const num = code.replace(/^[A-Z]{2}/, "");
        return num.startsWith(q) || code.includes(q);
      }
      return code.startsWith(q);
    });

    return pool.slice(0, 14).map(fmt);
  },

  /**
   * Split value into text before the last full flight segment (CODE/DATE/DEST) and the phrase tail.
   * Phrase matching uses only tailQuery so suggestions appear after a flight is chosen.
   */
  phraseCompositeParts(value) {
    const v = (value || "").toUpperCase();
    const flightRe = /[A-Z]{2}\d{2,4}\/[0-9A-Z]+\/[A-Z0-9]{2,4}/g;
    let lastEnd = -1;
    let m;
    while ((m = flightRe.exec(v)) !== null) {
      lastEnd = m.index + m[0].length;
    }
    if (lastEnd < 0) {
      return { prefix: "", tailQuery: v.trim() };
    }
    const tailQuery = v.slice(lastEnd).trim();
    const prefix = v.slice(0, lastEnd).trimEnd();
    return { prefix, tailQuery };
  },

  mergePhrasePick(fullValue, pickedPhrase) {
    const { prefix } = this.phraseCompositeParts(fullValue);
    const p = (pickedPhrase || "").toUpperCase();
    if (!prefix) return p;
    return `${prefix} ${p}`;
  },

  findMatches(key, phraseQuery, flightQuery) {
    const q = (phraseQuery || "").trim().toUpperCase();
    const fq =
      arguments.length >= 3 ? String(flightQuery ?? "").trim().toUpperCase() : this._isFlightAwarePhraseKey(key) ? "" : q;
    const list = this.getList(key);

    if (key !== "csdRescreening") {
      const learnedRaw =
        this._isLearnedPhraseKey(key) &&
        window.phraseUsageCache &&
        typeof window.phraseUsageCache.getSortedMatches === "function"
          ? window.phraseUsageCache.getSortedMatches(key, q, 14)
          : [];
      const fixedRaw = (list || [])
        .map((item) => String(item || "").toUpperCase().trim())
        .filter((item) => item && (!q || item.startsWith(q) || (q.length >= 2 && item.includes(q))));

      const flightsRaw = this._isFlightAwarePhraseKey(key)
        ? fq
          ? this._flightSuggestionStrings(fq, false)
          : this._flightSuggestionStrings("", true)
        : [];

      /* Priority: learned → fixed → flights, with guaranteed flight slots for flight-aware keys. */
      const merged = [];
      const seen = new Set();
      const cap = this._isFlightAwarePhraseKey(key) ? 32 : 12;
      const flightAware = this._isFlightAwarePhraseKey(key);
      const minFlightSlots = flightAware ? 10 : 0;
      const phraseCap = flightAware ? Math.max(0, cap - minFlightSlots) : cap;

      const pushKind = (text, kind) => {
        const t = String(text || "").trim();
        if (!t) return;
        const norm = t.toUpperCase();
        if (seen.has(norm)) return;
        seen.add(norm);
        merged.push({ text: t, kind });
        return merged.length >= cap;
      };

      for (const x of learnedRaw) {
        if (merged.length >= phraseCap) break;
        pushKind(x, "learned");
      }
      for (const x of fixedRaw) {
        if (merged.length >= phraseCap) break;
        pushKind(x, "fixed");
      }
      for (const x of flightsRaw) {
        if (pushKind(x, "flight")) break;
      }

      return merged;
    }

    const csdPhrases = this._csdRescreeningPhrases(list);

    /* CSD: learned routes (by frequency) + phrases + 3-letter codes (static + today’s flights). */
    const learned =
      window.csdRouteHintCache && typeof window.csdRouteHintCache.getSortedRoutesMatching === "function"
        ? window.csdRouteHintCache.getSortedRoutesMatching(q, 14)
        : [];

    /* After origin dash (e.g. MXP-), suggest route completions immediately. */
    const routePrefix = q.match(/^([A-Z]{3})-([A-Z]{0,3})$/);
    if (routePrefix) {
      const origin = routePrefix[1];
      const destQuery = routePrefix[2] || "";
      const learnedFromOrigin =
        window.csdRouteHintCache && typeof window.csdRouteHintCache.getSortedRoutesMatching === "function"
          ? window.csdRouteHintCache.getSortedRoutesMatching(`${origin}-`, 30)
          : [];

      const destPool = [];
      const poolSeen = new Set();
      const addDest = (code) => {
        const c = String(code || "").trim().toUpperCase();
        if (!/^[A-Z]{3}$/.test(c)) return;
        if (poolSeen.has(c)) return;
        poolSeen.add(c);
        destPool.push(c);
      };

      (this.csdDestinationHints || []).forEach(addDest);
      this._csdFlightDestinationHints("").forEach(addDest);

      const generatedRoutes = destPool
        .filter((d) => d.includes(destQuery))
        .map((d) => `${origin}-${d}`);

      const merged = [];
      const seen = new Set();
      for (const x of [...learnedFromOrigin, ...generatedRoutes]) {
        const norm = x.toUpperCase();
        if (!norm.startsWith(`${origin}-`)) continue;
        if (seen.has(norm)) continue;
        seen.add(norm);
        merged.push(x);
        if (merged.length >= 16) break;
      }
      return merged;
    }

    if (!q) {
      const phrasesFirst = csdPhrases.slice(0, 6);
      const flightTop = this._csdFlightDestinationHints("");
      const staticTop = (this.csdDestinationHints || []).slice(0, 10);
      const merged = [];
      const seen = new Set();
      for (const x of [...learned.slice(0, 10), ...staticTop, ...flightTop, ...phrasesFirst]) {
        const norm = x.toUpperCase();
        if (seen.has(norm)) continue;
        seen.add(norm);
        merged.push(x);
        if (merged.length >= 16) break;
      }
      return merged;
    }

    const phraseHits = csdPhrases
      .filter((item) => item.toUpperCase().includes(q))
      .slice(0, 6);
    const flightHits = this._csdFlightDestinationHints(q);
    const staticHits = (this.csdDestinationHints || [])
      .filter((h) => h.toUpperCase().includes(q))
      .slice(0, 10);

    const merged = [];
    const seen = new Set();
    for (const x of [...learned, ...phraseHits, ...flightHits, ...staticHits]) {
      const norm = x.toUpperCase();
      if (seen.has(norm)) continue;
      seen.add(norm);
      merged.push(x);
      if (merged.length >= 16) break;
    }
    return merged;
  },

  ensureBox() {
    if (this.activeBox) return this.activeBox;

    const box = document.createElement("div");
    box.id = "phrase-suggest-box";
    box.style.position = "absolute";
    box.style.zIndex = "10050";
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
    this.activeInput = null;
    this.activeItems = [];
    this.activeIndex = -1;
    this.activeOnPick = null;
  },

  positionBox(input) {
    const box = this.ensureBox();
    const rect = input.getBoundingClientRect();
    box.style.left = `${window.scrollX + rect.left}px`;
    box.style.top = `${window.scrollY + rect.bottom + 4}px`;
    box.style.width = `${rect.width}px`;
  },

  refreshHighlight() {
    const box = this.ensureBox();
    [...box.children].forEach((child, index) => {
      child.style.background = index === this.activeIndex ? "#eff6ff" : "#fff";
    });
  },

  /**
   * @param {Array<string|{text:string,kind?:string}>} items
   * @param {string} [phraseKey] — for deleting learned rows
   * @param {() => void} [afterDelete] — refresh suggestions
   */
  renderBox(input, items, onPick, phraseKey, afterDelete) {
    const box = this.ensureBox();

    const rows = (items || []).map((x) =>
      typeof x === "string" ? { text: x, kind: "fixed" } : { text: String(x.text || ""), kind: x.kind || "fixed" }
    );

    if (!rows.length) {
      this.hideBox();
      return;
    }

    this.activeInput = input;
    this.activeItems = rows;
    this.activeIndex = 0;
    this.activeOnPick = onPick;
    this._activePhraseKey = phraseKey || "";

    this.positionBox(input);
    box.innerHTML = "";

    rows.forEach((row, index) => {
      const item = document.createElement("div");
      item.style.padding = "6px 8px";
      item.style.cursor = "pointer";
      item.style.borderBottom = "1px solid #e2e8f0";
      item.style.background = index === 0 ? "#eff6ff" : "#fff";
      item.style.display = "flex";
      item.style.alignItems = "center";
      item.style.gap = "6px";

      const label = document.createElement("span");
      label.style.flex = "1";
      label.style.minWidth = "0";
      label.style.wordBreak = "break-word";
      label.textContent = row.text;
      item.appendChild(label);

      if (row.kind === "learned" && phraseKey && window.phraseUsageCache && typeof window.phraseUsageCache.removePhrase === "function") {
        const del = document.createElement("button");
        del.type = "button";
        del.setAttribute("aria-label", "Remove phrase");
        del.title = "Remove from learned list";
        del.textContent = "❌";
        del.style.flex = "0 0 auto";
        del.style.cursor = "pointer";
        del.style.border = "none";
        del.style.background = "transparent";
        del.style.padding = "2px 4px";
        del.style.fontSize = "12px";
        del.style.lineHeight = "1";
        del.addEventListener("mousedown", (e) => {
          e.preventDefault();
          e.stopPropagation();
          window.phraseUsageCache.removePhrase(phraseKey, row.text);
          if (typeof afterDelete === "function") afterDelete();
        });
        item.appendChild(del);
      }

      item.addEventListener("mouseenter", () => {
        this.activeIndex = index;
        this.refreshHighlight();
      });

      item.addEventListener("mousedown", (e) => {
        if (e.target && e.target.closest && e.target.closest("button")) return;
        e.preventDefault();
        if (this.activeOnPick) this.activeOnPick(row);
        this.hideBox();
      });

      box.appendChild(item);
    });

    box.style.display = "block";
  },

  pickActive() {
    if (!this.activeItems.length || !this.activeOnPick) return;
    this.activeOnPick(this.activeItems[this.activeIndex]);
    this.hideBox();
  },

  attachTextarea(textarea, key, onSave) {
    const row = textarea && textarea.dataset ? textarea.dataset.row : "";
    const field = textarea && textarea.dataset ? textarea.dataset.field : "";
    const mark = `phraseTa:${key}:${textarea.id || "noid"}:${row || "norow"}:${field || "nofield"}`;
    if (!textarea || textarea.dataset.phraseAttachMark === mark) return;
    textarea.dataset.phraseAttachMark = mark;

    const showSuggest = () => {
      let phraseTail = "";
      let flightQ = "";
      if (this._isLineScopedTextareaKey(key)) {
        const ctx = this._getTextareaActiveLine(textarea);
        if (this._isFlightAwarePhraseKey(key)) {
          const rawTail = this.phraseCompositeParts(ctx.line).tailQuery;
          phraseTail = rawTail.trim().toUpperCase();
          flightQ = this._flightPrefixQuery(ctx.line);
        } else {
          phraseTail = ctx.line.trim().toUpperCase();
        }
      } else {
        const parts = this.phraseCompositeParts(textarea.value);
        const rawTail = parts.tailQuery;
        phraseTail = rawTail.trim().toUpperCase();
        if (this._isFlightAwarePhraseKey(key)) {
          flightQ = this._flightPrefixQuery(textarea.value);
        }
      }
      const matches = this.findMatches(key, phraseTail, flightQ);
      const useDel = this._isLearnedPhraseKey(key);
      this.renderBox(
        textarea,
        matches,
        (picked) => {
          const t = this._normPick(picked);
          let merged;
          if (this._isLineScopedTextareaKey(key)) {
            if (this._isFlightAwarePhraseKey(key)) {
              const ctx = this._getTextareaActiveLine(textarea);
              const nextLine = this._mergeFlightAwareLine(ctx.line, t);
              merged = this._replaceTextareaActiveLine(textarea, nextLine);
            } else {
              merged = this._replaceTextareaActiveLine(textarea, t);
            }
          } else {
            merged = this.mergePhrasePick(textarea.value, t);
            textarea.value = merged;
            textarea.selectionStart = textarea.selectionEnd = merged.length;
          }
          if (onSave) onSave(merged);
        },
        useDel ? key : "",
        useDel ? showSuggest : null
      );
    };

    textarea.addEventListener("input", () => {
      const value = textarea.value.toUpperCase();
      textarea.value = value;
      if (onSave) onSave(value);
      showSuggest();
    });

    textarea.addEventListener("focus", showSuggest);

    textarea.addEventListener("keydown", (e) => {
      if (e.key === " " || e.code === "Space") {
        requestAnimationFrame(() => showSuggest());
      }
    });

    textarea.addEventListener(
      "keydown",
      (e) => {
        if (this.activeInput !== textarea || !this.activeItems.length) return;

        if (e.key === "ArrowDown") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.min(this.activeIndex + 1, this.activeItems.length - 1);
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

        if (e.key === "Escape") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.hideBox();
        }
      },
      true
    );

    textarea.addEventListener("blur", () => {
      if (this._isLearnedPhraseKey(key) && window.phraseUsageCache && typeof window.phraseUsageCache.recordPhrase === "function") {
        window.phraseUsageCache.recordPhrase(key, textarea.value);
      }
      setTimeout(() => {
        if (this.activeInput === textarea) this.hideBox();
      }, 150);
    });
  },

  attachInput(input, key, onSave) {
    const g = input.dataset.group ?? "x";
    const i = input.dataset.index ?? "x";
    const mark = `phraseInp:${key}:${g}:${i}`;
    if (!input || input.dataset.phraseAttachMark === mark) return;
    input.dataset.phraseAttachMark = mark;

    const showSuggest = () => {
      let tailQuery;
      if (key === "csdRescreening") {
        const { ready, tail } = this.csdTailAfterPolicy(input.value);
        if (!ready) {
          this.hideBox();
          return;
        }
        tailQuery = tail;
        const matches = this.findMatches(key, tailQuery);
        this.renderBox(input, matches, (picked) => {
          const t = this._normPick(picked);
          const merged = this.mergePhrasePickCsd(input.value, t);
          input.value = merged;
          if (onSave) onSave(merged);
        });
        return;
      }
      const parts = this.phraseCompositeParts(input.value);
      const rawTail = parts.tailQuery;
      tailQuery = rawTail.trim().toUpperCase();
      const flightQ = this._isFlightAwarePhraseKey(key) ? this._flightPrefixQuery(input.value) : "";
      const matches = this.findMatches(key, tailQuery, flightQ);
      const useDel = this._isLearnedPhraseKey(key);
      this.renderBox(
        input,
        matches,
        (picked) => {
          const t = this._normPick(picked);
          let merged;
          if (this._isFlightAwarePhraseKey(key)) {
            const via = this._mergeFlightPickIntoValue(input.value, t);
            merged = via != null ? via : this.mergePhrasePick(input.value, t);
          } else {
            merged = this.mergePhrasePick(input.value, t);
          }
          input.value = merged;
          if (onSave) onSave(merged);
        },
        useDel ? key : "",
        useDel ? showSuggest : null
      );
    };

    input.addEventListener("input", () => {
      const value = input.value.toUpperCase();
      input.value = value;
      if (onSave) onSave(value);
      showSuggest();
    });

    input.addEventListener("focus", showSuggest);

    input.addEventListener("keydown", (e) => {
      if (e.key === " " || e.code === "Space") {
        requestAnimationFrame(() => showSuggest());
      }
    });

    input.addEventListener(
      "keydown",
      (e) => {
        if (this.activeInput !== input || !this.activeItems.length) return;

        if (e.key === "ArrowDown") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.activeIndex = Math.min(this.activeIndex + 1, this.activeItems.length - 1);
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

        if (e.key === "Escape") {
          e.preventDefault();
          e.stopImmediatePropagation();
          this.hideBox();
        }
      },
      true
    );

    input.addEventListener("blur", () => {
      if (this._isLearnedPhraseKey(key) && window.phraseUsageCache && typeof window.phraseUsageCache.recordPhrase === "function") {
        window.phraseUsageCache.recordPhrase(key, input.value);
      }
      setTimeout(() => {
        if (this.activeInput === input) this.hideBox();
      }, 150);
    });
  }
};

window.addEventListener("resize", () => {
  if (window.phraseAutocomplete && window.phraseAutocomplete.activeInput && window.phraseAutocomplete.activeItems.length) {
    window.phraseAutocomplete.positionBox(window.phraseAutocomplete.activeInput);
  }
});

window.addEventListener("scroll", () => {
  if (window.phraseAutocomplete && window.phraseAutocomplete.activeInput && window.phraseAutocomplete.activeItems.length) {
    window.phraseAutocomplete.positionBox(window.phraseAutocomplete.activeInput);
  }
}, true);