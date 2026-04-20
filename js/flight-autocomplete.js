window.flightAutocomplete = {
  flights: [],
  activeBox: null,
  activeInput: null,
  activeMatches: [],
  activeIndex: -1,
  activeOnPick: null,
  activeMode: "replace",

  async load(url = "../../data/report/flights.json") {
    try {
      const res = await fetch(url);
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (text.trimStart().startsWith("<")) {
        throw new Error(`Expected JSON but got HTML (wrong URL?). ${url}`);
      }
      this.flights = JSON.parse(text);
    } catch (err) {
      console.error("Failed to load flights.json", url, err);
      this.flights = [];
    }
  },

  normalizeCode(value) {
    return (value || "").trim().toUpperCase();
  },

  formatFlight(flight) {
    const code = this.normalizeCode(flight.code);
    const date = (flight.date || "").trim().toUpperCase();
    const destination = (flight.destination || "").trim().toUpperCase();
    return [code, date, destination].filter(Boolean).join("/");
  },

  findByCode(code) {
    const q = this.normalizeCode(code);
    return this.flights.find(f => this.normalizeCode(f.code) === q) || null;
  },

  findMatches(query) {
    const q = this.normalizeCode(query);
    if (!q) return [];
    return this.flights
      .filter(f => this.normalizeCode(f.code).startsWith(q))
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

  /** Flight-ish token at line end; prefix must be start or non-alnum so "HELLO" does not match "LO". */
  lastFlightChunkRe: /(?:^|[^A-Z0-9])([A-Z]{2}\d{0,4})$/i,

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
      const matches = this.findMatches(input.value);
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
      const matches = this.findMatches(token);

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
      const matches = this.findMatches(token);

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