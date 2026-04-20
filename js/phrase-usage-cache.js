/**
 * Learned usage counts for phrase suggestions.
 * Same-origin GET/POST /api/phrase-usage when using server/flight_hints_server.py.
 * Fallback: data/report/phrase-usage.json + localStorage mirror.
 */
window.phraseUsageCache = {
  _memory: {},
  _apiBase: "",
  LS_KEY: "activityReport_phraseUsage_v1",
  KEYS: ["loadPlan", "advanceLoading", "offloadReason", "offloadRemarks", "other", "specialHO"],

  _hintsUrl() {
    const b = (this._apiBase || "").replace(/\/$/, "");
    const path = "/api/phrase-usage";
    if (!b) return path;
    return `${b}${path}`;
  },

  _authHeaders() {
    const meta = document.querySelector('meta[name="flight-hints-token"]');
    const token = meta && meta.getAttribute("content");
    const t =
      (token && token.trim()) || (typeof window.FLIGHT_HINTS_TOKEN === "string" ? window.FLIGHT_HINTS_TOKEN.trim() : "");
    return t ? { "X-Flight-Hints-Token": t } : {};
  },

  _replaceMemory(obj) {
    this._memory = {};
    if (!obj || typeof obj !== "object") return;
    this.KEYS.forEach((k) => {
      const src = obj[k];
      if (!src || typeof src !== "object") return;
      const inner = {};
      Object.keys(src).forEach((phrase) => {
        const p = String(phrase || "").trim().toUpperCase();
        const c = Number(src[phrase]);
        if (!p || !Number.isFinite(c) || c <= 0) return;
        inner[p] = Math.floor(c);
      });
      this._memory[k] = inner;
    });
  },

  _persistLocal() {
    try {
      localStorage.setItem(this.LS_KEY, JSON.stringify(this._memory));
    } catch (_) {}
  },

  _loadLocalStorageOnly() {
    try {
      const raw = localStorage.getItem(this.LS_KEY);
      if (!raw) {
        this._replaceMemory({});
        return;
      }
      this._replaceMemory(JSON.parse(raw));
    } catch {
      this._replaceMemory({});
    }
  },

  /**
   * @param {{ apiBase?: string, fallbackUrl?: string } | string | undefined} options
   */
  async hydrate(options) {
    let apiBase;
    let fallbackUrl = null;

    if (typeof options === "string") {
      apiBase = options;
    } else if (options && typeof options === "object") {
      if (options.apiBase != null) apiBase = String(options.apiBase);
      if (options.fallbackUrl != null) fallbackUrl = String(options.fallbackUrl);
    }

    if (apiBase === undefined) {
      const meta = document.querySelector('meta[name="flight-hints-api-base"]');
      if (meta && meta.getAttribute("content") != null) {
        apiBase = String(meta.getAttribute("content") || "");
      } else if (typeof window.FLIGHT_HINTS_API_BASE === "string") {
        apiBase = window.FLIGHT_HINTS_API_BASE;
      } else {
        apiBase = "";
      }
    }

    if (fallbackUrl == null && typeof window.PHRASE_USAGE_FALLBACK_URL === "string") {
      fallbackUrl = window.PHRASE_USAGE_FALLBACK_URL;
    }

    this._apiBase = apiBase;

    try {
      const r = await fetch(this._hintsUrl(), {
        cache: "no-store",
        headers: { ...this._authHeaders() }
      });
      if (r.ok) {
        this._replaceMemory(await r.json());
        this._persistLocal();
        return true;
      }
    } catch (e) {
      console.warn("Phrase usage API unavailable", e);
    }

    if (fallbackUrl) {
      try {
        const r2 = await fetch(fallbackUrl);
        if (r2.ok) {
          this._replaceMemory(await r2.json());
          this._persistLocal();
          return true;
        }
      } catch (e) {
        console.warn("Phrase usage fallback file failed", e);
      }
    }

    this._loadLocalStorageOnly();
    return false;
  },

  /** Prefix-only matches sorted by usage desc. */
  getSortedPrefixMatches(key, query, limit) {
    const k = String(key || "").trim();
    const q = String(query || "").trim().toUpperCase();
    const bucket = this._memory[k] && typeof this._memory[k] === "object" ? this._memory[k] : {};
    const limitN = Math.min(Number(limit) || 12, 40);
    return Object.entries(bucket)
      .filter(([phrase, c]) => Number(c) > 0 && (!q || phrase.startsWith(q)))
      .sort((a, b) => Number(b[1]) - Number(a[1]))
      .map(([phrase]) => phrase)
      .slice(0, limitN);
  },

  recordPhrase(key, text) {
    const k = String(key || "").trim();
    if (!this.KEYS.includes(k)) return;
    const phrase = String(text || "").trim().toUpperCase();
    if (!phrase) return;
    this.pushIncrementMerge({ [k]: { [phrase]: 1 } });
  },

  pushIncrementMerge(merge) {
    if (!merge || typeof merge !== "object") return;
    Object.keys(merge).forEach((k) => {
      if (!this.KEYS.includes(k)) return;
      const bucket = merge[k];
      if (!bucket || typeof bucket !== "object") return;
      if (!this._memory[k] || typeof this._memory[k] !== "object") this._memory[k] = {};
      Object.keys(bucket).forEach((phrase) => {
        const p = String(phrase || "").trim().toUpperCase();
        const inc = Number(bucket[phrase]);
        if (!p || !Number.isFinite(inc) || inc <= 0) return;
        this._memory[k][p] = (this._memory[k][p] || 0) + inc;
      });
    });
    this._persistLocal();

    try {
      fetch(this._hintsUrl(), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...this._authHeaders()
        },
        body: JSON.stringify({ merge })
      }).then((r) => {
        if (!r.ok) console.warn("phrase usage push failed", r.status);
      });
    } catch (e) {
      console.warn("phrase usage push network error", e);
    }
  }
};
