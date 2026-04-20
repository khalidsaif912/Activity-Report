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
      if (!Array.isArray(flights)) return [];
      const out = [];
      const seen = new Set();
      for (const f of flights) {
        const dest = String(f.destination || "")
          .trim()
          .toUpperCase()
          .match(/^([A-Z]{3})/);
        const code3 = dest ? dest[1] : "";
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
      key === "offloadReason" ||
      key === "offloadRemarks" ||
      key === "other" ||
      key === "specialHO"
    );
  },

  _isFlightAwarePhraseKey(key) {
    return key === "other" || key === "specialHO";
  },

  _isLineScopedTextareaKey(key) {
    return key === "other" || key === "specialHO" || key === "offloadReason" || key === "offloadRemarks";
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

  _flightSuggestionStrings(query) {
    const q = (query || "").trim().toUpperCase();
    if (!window.flightAutocomplete || typeof window.flightAutocomplete.findMatches !== "function") return [];
    let matches = [];
    if (!q) {
      matches = Array.isArray(window.flightAutocomplete.flights) ? window.flightAutocomplete.flights.slice(0, 8) : [];
    } else if (/^[A-Z]{1,2}\d{0,4}$/.test(q)) {
      matches = window.flightAutocomplete.findMatches(q);
    } else {
      matches = [];
    }
    if (!Array.isArray(matches) || !matches.length) return [];
    return matches.map((f) =>
      window.flightAutocomplete && typeof window.flightAutocomplete.formatFlight === "function"
        ? window.flightAutocomplete.formatFlight(f)
        : [f.code, f.date, f.destination].filter(Boolean).join("/")
    );
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

  findMatches(key, query) {
    const q = (query || "").trim().toUpperCase();
    const list = this.getList(key);

    if (key !== "csdRescreening") {
      const learned =
        this._isLearnedPhraseKey(key) &&
        window.phraseUsageCache &&
        typeof window.phraseUsageCache.getSortedPrefixMatches === "function"
          ? window.phraseUsageCache.getSortedPrefixMatches(key, q, 12)
          : [];
      const fixed = (list || [])
        .map((item) => String(item || "").toUpperCase().trim())
        .filter((item) => item && (!q || item.startsWith(q)));
      const flights = this._isFlightAwarePhraseKey(key) ? this._flightSuggestionStrings(q) : [];

      const merged = [];
      const seen = new Set();
      for (const x of [...flights, ...learned, ...fixed]) {
        const norm = x.toUpperCase();
        if (seen.has(norm)) continue;
        seen.add(norm);
        merged.push(x);
        if (merged.length >= 12) break;
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

  renderBox(input, items, onPick) {
    const box = this.ensureBox();

    if (!items.length) {
      this.hideBox();
      return;
    }

    this.activeInput = input;
    this.activeItems = items;
    this.activeIndex = 0;
    this.activeOnPick = onPick;

    this.positionBox(input);
    box.innerHTML = "";

    items.forEach((text, index) => {
      const item = document.createElement("div");
      item.style.padding = "8px 10px";
      item.style.cursor = "pointer";
      item.style.borderBottom = "1px solid #e2e8f0";
      item.style.background = index === 0 ? "#eff6ff" : "#fff";
      item.textContent = text;

      item.addEventListener("mouseenter", () => {
        this.activeIndex = index;
        this.refreshHighlight();
      });

      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        if (this.activeOnPick) this.activeOnPick(text);
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
      let tailQuery;
      if (this._isLineScopedTextareaKey(key)) {
        const ctx = this._getTextareaActiveLine(textarea);
        tailQuery = this._isFlightAwarePhraseKey(key) ? this._tailTokenForFlightAwareLine(ctx.line) : ctx.line.trim().toUpperCase();
      } else {
        const parts = this.phraseCompositeParts(textarea.value);
        tailQuery = parts.tailQuery;
      }
      const matches = this.findMatches(key, tailQuery);
      this.renderBox(textarea, matches, (picked) => {
        let merged;
        if (this._isLineScopedTextareaKey(key)) {
          if (this._isFlightAwarePhraseKey(key)) {
            const ctx = this._getTextareaActiveLine(textarea);
            const nextLine = this._mergeFlightAwareLine(ctx.line, picked);
            merged = this._replaceTextareaActiveLine(textarea, nextLine);
          } else {
            merged = this._replaceTextareaActiveLine(textarea, picked);
          }
        } else {
          merged = this.mergePhrasePick(textarea.value, picked);
          textarea.value = merged;
          textarea.selectionStart = textarea.selectionEnd = merged.length;
        }
        if (onSave) onSave(merged);
      });
    };

    textarea.addEventListener("input", () => {
      const value = textarea.value.toUpperCase();
      textarea.value = value;
      if (onSave) onSave(value);
      showSuggest();
    });

    textarea.addEventListener("focus", showSuggest);

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
      } else {
        tailQuery = this.phraseCompositeParts(input.value).tailQuery;
      }
      const matches = this.findMatches(key, tailQuery);
      this.renderBox(input, matches, (picked) => {
        const merged =
          key === "csdRescreening" ? this.mergePhrasePickCsd(input.value, picked) : this.mergePhrasePick(input.value, picked);
        input.value = merged;
        if (onSave) onSave(merged);
      });
    };

    input.addEventListener("input", () => {
      const value = input.value.toUpperCase();
      input.value = value;
      if (onSave) onSave(value);
      showSuggest();
    });

    input.addEventListener("focus", showSuggest);

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