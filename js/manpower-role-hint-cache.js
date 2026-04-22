/**
 * Learned role hints per manpower name (e.g. "SN82547 ... -> W/B Checker").
 * Same-origin GET/POST /api/manpower-role-hints when using server/flight_hints_server.py.
 * Fallback: data/report/manpower-role-hints.json + localStorage mirror.
 */
window.manpowerRoleHintCache = {
  _memory: {},
  _defaultRoles: [],
  _apiBase: "",
  LS_KEY: "activityReport_manpowerRoleHints_v1",

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
    const path = "/api/manpower-role-hints";
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

  _normalizeName(name) {
    return String(name || "")
      .trim()
      .replace(/\s+/g, " ")
      .toUpperCase();
  },

  _normalizeRole(role) {
    return String(role || "")
      .trim()
      .replace(/\s+/g, " ")
      .toUpperCase();
  },

  _cleanLabel(s) {
    return String(s || "").trim().replace(/\s+/g, " ");
  },

  async loadRoleOptions(url = "../../data/report/manpower-role-options.json") {
    try {
      const res = await fetch(url);
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (text.trimStart().startsWith("<")) throw new Error(`Expected JSON but got HTML (wrong URL?). ${url}`);
      const arr = JSON.parse(text);
      if (!Array.isArray(arr)) throw new Error("Expected array");
      this._defaultRoles = arr.map((x) => this._cleanLabel(x)).filter(Boolean);
    } catch (e) {
      console.warn("Manpower role options load failed", e);
      this._defaultRoles = [];
    }
  },

  _parseLine(line) {
    const v = this._cleanLabel(line);
    const m = v.match(/^(.+?)\s*-\s*(.+)$/);
    if (!m) return null;
    const name = this._cleanLabel(m[1]);
    const role = this._cleanLabel(m[2]);
    if (!name || !role) return null;
    return { name, role };
  },

  _replaceMemory(obj) {
    this._memory = {};
    if (!obj || typeof obj !== "object") return;
    Object.keys(obj).forEach((nameKeyRaw) => {
      const entry = obj[nameKeyRaw];
      if (!entry || typeof entry !== "object") return;
      const nameKey = this._normalizeName(nameKeyRaw);
      if (!nameKey) return;
      const display = this._cleanLabel(entry.display || nameKeyRaw);
      const rolesSrc = entry.roles;
      if (!rolesSrc || typeof rolesSrc !== "object") return;
      const roles = {};
      Object.keys(rolesSrc).forEach((roleKeyRaw) => {
        const roleEntry = rolesSrc[roleKeyRaw];
        if (!roleEntry || typeof roleEntry !== "object") return;
        const roleKey = this._normalizeRole(roleKeyRaw);
        const label = this._cleanLabel(roleEntry.label || roleKeyRaw);
        const count = Number(roleEntry.count);
        if (!roleKey || !label || !Number.isFinite(count) || count <= 0) return;
        roles[roleKey] = { label, count: Math.floor(count) };
      });
      if (Object.keys(roles).length) {
        this._memory[nameKey] = { display, roles };
      }
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

    if (fallbackUrl == null && typeof window.MANPOWER_ROLE_HINTS_FALLBACK_URL === "string") {
      fallbackUrl = window.MANPOWER_ROLE_HINTS_FALLBACK_URL;
    }

    this._apiBase = apiBase;

    if (this._canUseApi()) {
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
        console.warn("Manpower role hints API unavailable", e);
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
        console.warn("Manpower role hints fallback file failed", e);
      }
    }

    this._loadLocalStorageOnly();
    return false;
  },

  getRolesForName(name, query, limit) {
    const nameKey = this._normalizeName(name);
    const q = this._normalizeRole(query || "");
    const limitN = Math.min(Number(limit) || 8, 20);
    const learned = !nameKey || !this._memory[nameKey]
      ? []
      : Object.values(this._memory[nameKey].roles)
      .filter((r) => !q || this._normalizeRole(r.label).startsWith(q))
      .sort((a, b) => Number(b.count) - Number(a.count))
      .map((r) => r.label);
    const defaults = (this._defaultRoles || []).filter((r) => !q || this._normalizeRole(r).startsWith(q));
    const merged = [];
    const seen = new Set();
    for (const x of [...learned, ...defaults]) {
      const norm = this._normalizeRole(x);
      if (!norm || seen.has(norm)) continue;
      seen.add(norm);
      merged.push(x);
      if (merged.length >= limitN) break;
    }
    return merged;
  },

  getTopRoleForName(name) {
    const roles = this.getRolesForName(name, "", 1);
    return roles.length ? roles[0] : "";
  },

  /** Learned-only top role (ignores default role options). */
  getTopLearnedRoleForName(name, minCount) {
    const nameKey = this._normalizeName(name);
    if (!nameKey || !this._memory[nameKey] || !this._memory[nameKey].roles) return "";
    const minN = Math.max(1, Number(minCount) || 1);
    const top = Object.values(this._memory[nameKey].roles)
      .filter((r) => Number(r.count) >= minN)
      .sort((a, b) => Number(b.count) - Number(a.count))[0];
    return top && top.label ? top.label : "";
  },

  recordFromLine(line) {
    const parsed = this._parseLine(line);
    if (!parsed) return;
    this.pushIncrementMerge({ [parsed.name]: { [parsed.role]: 1 } });
  },

  pushIncrementMerge(merge) {
    if (!merge || typeof merge !== "object") return;
    const outMerge = {};

    Object.keys(merge).forEach((nameRaw) => {
      const roles = merge[nameRaw];
      if (!roles || typeof roles !== "object") return;
      const nameLabel = this._cleanLabel(nameRaw);
      const nameKey = this._normalizeName(nameLabel);
      if (!nameKey) return;
      if (!this._memory[nameKey]) this._memory[nameKey] = { display: nameLabel, roles: {} };
      if (!outMerge[nameLabel]) outMerge[nameLabel] = {};
      this._memory[nameKey].display = nameLabel || this._memory[nameKey].display;

      Object.keys(roles).forEach((roleRaw) => {
        const inc = Number(roles[roleRaw]);
        const roleLabel = this._cleanLabel(roleRaw);
        const roleKey = this._normalizeRole(roleLabel);
        if (!roleKey || !roleLabel || !Number.isFinite(inc) || inc <= 0) return;

        const cur = this._memory[nameKey].roles[roleKey];
        const count = (cur ? Number(cur.count) : 0) + inc;
        this._memory[nameKey].roles[roleKey] = { label: roleLabel, count };
        outMerge[nameLabel][roleLabel] = (outMerge[nameLabel][roleLabel] || 0) + inc;
      });
    });

    this._persistLocal();

    if (!Object.keys(outMerge).length) return;
    if (this._canUseApi()) {
      try {
        fetch(this._hintsUrl(), {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            ...this._authHeaders()
          },
          body: JSON.stringify({ merge: outMerge })
        }).then((r) => {
          if (!r.ok) console.warn("manpower role hints push failed", r.status);
        });
      } catch (e) {
        console.warn("manpower role hints push network error", e);
      }
    }
  }
};
