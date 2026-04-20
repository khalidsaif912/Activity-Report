/**
 * CSD Rescreening: learned route pairs (e.g. MXP-KUL, FRA-MNL) with usage counts.
 * Same-origin GET/POST /api/csd-route-hints when using server/flight_hints_server.py.
 * Configure base URL via <meta name="flight-hints-api-base" content="http://127.0.0.1:5050"> (shared with flight hints).
 *
 * Fallback: data/report/csd-route-hints.json + localStorage mirror.
 */
window.csdRouteHintCache = {
  _memory: {},
  _apiBase: "",
  LS_KEY: "activityReport_csdRouteCounts_v1",

  _hintsUrl() {
    const b = (this._apiBase || "").replace(/\/$/, "");
    const path = "/api/csd-route-hints";
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
    Object.keys(obj).forEach((k) => {
      const norm = String(k || "")
        .trim()
        .toUpperCase();
      if (!/^[A-Z]{3}-[A-Z]{3}$/.test(norm)) return;
      const c = Number(obj[k]);
      if (!Number.isFinite(c) || c <= 0) return;
      this._memory[norm] = Math.floor(c);
    });
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

    if (fallbackUrl == null && typeof window.CSD_ROUTE_HINTS_FALLBACK_URL === "string") {
      fallbackUrl = window.CSD_ROUTE_HINTS_FALLBACK_URL;
    }

    this._apiBase = apiBase;

    try {
      const r = await fetch(this._hintsUrl(), {
        cache: "no-store",
        headers: { ...this._authHeaders() }
      });
      if (r.ok) {
        const o = await r.json();
        this._replaceMemory(o);
        this._persistLocal();
        return true;
      }
    } catch (e) {
      console.warn("CSD route hints API unavailable", e);
    }

    if (fallbackUrl) {
      try {
        const r2 = await fetch(fallbackUrl);
        if (r2.ok) {
          const o = await r2.json();
          this._replaceMemory(o);
          this._persistLocal();
          return true;
        }
      } catch (e) {
        console.warn("CSD route hints fallback file failed", e);
      }
    }

    this._loadLocalStorageOnly();
    return false;
  },

  _loadLocalStorageOnly() {
    try {
      const raw = localStorage.getItem(this.LS_KEY);
      if (!raw) {
        this._replaceMemory({});
        return;
      }
      const o = JSON.parse(raw);
      this._replaceMemory(o);
    } catch {
      this._replaceMemory({});
    }
  },

  _persistLocal() {
    try {
      localStorage.setItem(this.LS_KEY, JSON.stringify(this._memory));
    } catch (_) {}
  },

  /**
   * Extract XXX-YYY and glued XXXYYY (e.g. MXPBKK → MXP-BKK). Ignores numeric AWB prefixes.
   * @returns {string[]}
   */
  extractRouteTokensFromText(text) {
    const u = String(text || "").toUpperCase();
    const found = new Set();
    const reDash = /\b([A-Z]{3})-([A-Z]{3})\b/g;
    let m;
    while ((m = reDash.exec(u)) !== null) {
      found.add(`${m[1]}-${m[2]}`);
    }
    const dejunk = u.replace(/[0-9]+/g, " ");
    const re6 = /\b([A-Z]{3})([A-Z]{3})\b/g;
    while ((m = re6.exec(dejunk)) !== null) {
      found.add(`${m[1]}-${m[2]}`);
    }
    return [...found];
  },

  /** Sorted by count desc, optional query filter (substring match). */
  getSortedRoutesMatching(query, limit) {
    const q = (query || "").trim().toUpperCase();
    const limitN = Math.min(Number(limit) || 16, 40);
    const entries = Object.entries(this._memory)
      .filter(([k, v]) => /^[A-Z]{3}-[A-Z]{3}$/.test(k) && Number(v) > 0)
      .map(([k, v]) => [k, Number(v)])
      .sort((a, b) => b[1] - a[1]);
    let list = entries.map(([k]) => k);
    if (q) list = list.filter((k) => k.includes(q));
    return list.slice(0, limitN);
  },

  /** Increment counts for routes found in text; POST merge to server (best-effort). */
  recordFromText(text) {
    const routes = this.extractRouteTokensFromText(text);
    if (!routes.length) return;
    const merge = {};
    routes.forEach((r) => {
      merge[r] = 1;
    });
    this.pushIncrementMerge(merge);
  },

  pushIncrementMerge(merge) {
    if (!merge || typeof merge !== "object") return;
    Object.keys(merge).forEach((k) => {
      if (!/^[A-Z]{3}-[A-Z]{3}$/.test(k)) return;
      const inc = Number(merge[k]);
      if (!Number.isFinite(inc) || inc <= 0) return;
      this._memory[k] = (this._memory[k] || 0) + inc;
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
        if (!r.ok) console.warn("csd route hints push failed", r.status);
      });
    } catch (e) {
      console.warn("csd route hints push network error", e);
    }
  }
};
