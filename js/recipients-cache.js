/**
 * Shared recipients (To/Cc/Bcc) storage.
 * API:
 *   GET /api/recipients
 *   POST /api/recipients  { to:[], cc:[], bcc:[] }
 */
window.recipientsCache = {
  _memory: { to: [], cc: [], bcc: [] },
  _apiBase: "",
  LS_KEY: "activityReport_recipients_v1",

  _canUseApi() {
    const base = String(this._apiBase || "").trim();
    if (base) return true;
    if (location.protocol !== "http:" && location.protocol !== "https:") return false;
    const host = String(location.hostname || "").trim().toLowerCase();
    if (!host) return false;
    return !host.endsWith("github.io");
  },

  _hintsUrl() {
    const b = (this._apiBase || "").replace(/\/$/, "");
    const path = "/api/recipients";
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

  _normList(arr) {
    const seen = new Set();
    const out = [];
    (Array.isArray(arr) ? arr : []).forEach((x) => {
      const e = String(x || "").trim().toLowerCase();
      if (!e || seen.has(e)) return;
      seen.add(e);
      out.push(e);
    });
    return out;
  },

  _replaceMemory(obj) {
    const src = obj && typeof obj === "object" ? obj : {};
    this._memory = {
      to: this._normList(src.to),
      cc: this._normList(src.cc),
      bcc: this._normList(src.bcc),
    };
  },

  _persistLocal() {
    try {
      localStorage.setItem(this.LS_KEY, JSON.stringify(this._memory));
    } catch (_) {}
  },

  _loadLocalStorageOnly() {
    try {
      const raw = localStorage.getItem(this.LS_KEY);
      this._replaceMemory(raw ? JSON.parse(raw) : {});
    } catch {
      this._replaceMemory({});
    }
  },

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
      apiBase = meta && meta.getAttribute("content") != null ? String(meta.getAttribute("content") || "") : "";
    }
    this._apiBase = apiBase;

    if (this._canUseApi()) {
      try {
        const r = await fetch(this._hintsUrl(), { cache: "no-store", headers: { ...this._authHeaders() } });
        if (r.ok) {
          this._replaceMemory(await r.json());
          this._persistLocal();
          return true;
        }
      } catch (e) {
        console.warn("Recipients API unavailable", e);
      }
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
        console.warn("Recipients fallback file failed", e);
      }
    }
    this._loadLocalStorageOnly();
    return false;
  },

  getAll() {
    return JSON.parse(JSON.stringify(this._memory));
  },

  async replaceAll(next) {
    this._replaceMemory(next);
    this._persistLocal();
    if (this._canUseApi()) {
      try {
        const r = await fetch(this._hintsUrl(), {
          method: "POST",
          headers: { "Content-Type": "application/json", ...this._authHeaders() },
          body: JSON.stringify(this._memory),
        });
        if (!r.ok) console.warn("recipients push failed", r.status);
      } catch (e) {
        console.warn("recipients push network error", e);
      }
    }
  },
};
