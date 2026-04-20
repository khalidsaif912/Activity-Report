/**
 * Flight STD/DEST hints: loaded from server API (shared), not localStorage.
 * Configure base URL via <meta name="flight-hints-api-base" content="http://127.0.0.1:5050">
 * (empty = same origin, e.g. when using server/flight_hints_server.py).
 *
 * GET/POST /api/flight-hints — see server/flight_hints_server.py
 */
window.flightHintCache = {
  _memory: {},
  _apiBase: "",

  cacheKey(iso, code) {
    const c = String(code || "")
      .trim()
      .toUpperCase()
      .replace(/\s/g, "");
    const d = String(iso || "")
      .trim()
      .match(/^\d{4}-\d{2}-\d{2}$/)
      ? String(iso).trim()
      : "";
    return d && c ? `${d}|${c}` : "";
  },

  _hintsUrl() {
    const b = (this._apiBase || "").replace(/\/$/, "");
    const path = "/api/flight-hints";
    if (!b) return path;
    return `${b}${path}`;
  },

  _authHeaders() {
    const meta = document.querySelector('meta[name="flight-hints-token"]');
    const token = meta && meta.getAttribute("content");
    const t = (token && token.trim()) || (typeof window.FLIGHT_HINTS_TOKEN === "string" ? window.FLIGHT_HINTS_TOKEN.trim() : "");
    return t ? { "X-Flight-Hints-Token": t } : {};
  },

  _replaceMemory(obj) {
    this._memory = obj && typeof obj === "object" ? { ...obj } : {};
  },

  /**
   * Load hints from API; on failure optionally merge static JSON (bundled file).
   * @param {{ apiBase?: string, fallbackUrl?: string } | string | undefined} options
   */
  async hydrate(options) {
    let apiBase;
    let fallbackUrl = null;

    if (typeof options === "string") {
      apiBase = options;
    } else if (options && typeof options === "object") {
      if (options.apiBase != null) apiBase = String(options.apiBase);
      if (options.fallbackUrl != null) fallbackUrl = options.fallbackUrl;
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

    if (fallbackUrl == null && typeof window.FLIGHT_HINTS_FALLBACK_URL === "string") {
      fallbackUrl = window.FLIGHT_HINTS_FALLBACK_URL;
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
        return true;
      }
    } catch (e) {
      console.warn("Flight hints API unavailable", e);
    }

    if (fallbackUrl) {
      try {
        const r2 = await fetch(fallbackUrl);
        if (r2.ok) {
          this._replaceMemory(await r2.json());
          return true;
        }
      } catch (e) {
        console.warn("Flight hints fallback file failed", e);
      }
    }

    this._replaceMemory({});
    return false;
  },

  get(iso, code) {
    const k = this.cacheKey(iso, code);
    if (!k) return null;
    const v = this._memory[k];
    if (!v || typeof v !== "object") return null;
    return { std: v.std || "", destination: v.destination || "" };
  },

  /** Batch-save to server (merge). Updates memory immediately; POST is best-effort. */
  async pushMerge(merge) {
    if (!merge || typeof merge !== "object") return;
    Object.keys(merge).forEach((k) => {
      const v = merge[k];
      if (v && typeof v === "object" && (v.std || "").trim() && (v.destination || "").trim()) {
        this._memory[k] = {
          std: String(v.std).trim(),
          destination: String(v.destination).trim()
        };
      }
    });
    try {
      const r = await fetch(this._hintsUrl(), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...this._authHeaders()
        },
        body: JSON.stringify({ merge })
      });
      if (!r.ok) {
        console.warn("flight hints push failed", r.status);
      }
    } catch (e) {
      console.warn("flight hints push network error", e);
    }
  },

  exportJson() {
    return JSON.stringify(this._memory, null, 2);
  },

  downloadExport() {
    const blob = new Blob([this.exportJson()], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "flight-hints.json";
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
};
