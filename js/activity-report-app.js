/**
 * Export Warehouse Activity Report — page logic (kept out of HTML for maintainability).
 * Expects DOM ids used by data/report/offload_report.html.
 */
(function () {
  console.info(
    "[activity-report] client bundle v33 — Phrase delete + merged flight/phrase suggestions."
  );
  /**
   * Single-folder fallback when meta/window overrides are unset.
   * - file: ../../data/report/
   * - http(s): same directory as this HTML, then site-root /data/report/ (Flask).
   */
  function reportAssetsRoot() {
    const meta = document.querySelector('meta[name="activity-report-assets-root"]');
    if (meta && String(meta.content || "").trim()) {
      return String(meta.content).trim().replace(/\/?$/, "/");
    }
    if (typeof window.ACTIVITY_REPORT_ASSETS_ROOT === "string" && window.ACTIVITY_REPORT_ASSETS_ROOT.trim()) {
      return window.ACTIVITY_REPORT_ASSETS_ROOT.trim().replace(/\/?$/, "/");
    }
    if (location.protocol === "http:" || location.protocol === "https:") {
      return new URL("/data/report/", window.location.origin).href;
    }
    return "../../data/report/";
  }

  /** Ordered URL bases to try for employees.json / flights.json / … (IDE preview, Flask, Live Server). */
  function buildReportAssetBaseCandidates() {
    const out = [];
    const add = (s) => {
      if (s == null || s === "") return;
      const x = String(s).trim().replace(/\/?$/, "/");
      if (out.indexOf(x) === -1) out.push(x);
    };

    const meta = document.querySelector('meta[name="activity-report-assets-root"]');
    if (meta && String(meta.content || "").trim()) {
      add(meta.content.trim());
      return out;
    }
    if (typeof window.ACTIVITY_REPORT_ASSETS_ROOT === "string" && window.ACTIVITY_REPORT_ASSETS_ROOT.trim()) {
      add(window.ACTIVITY_REPORT_ASSETS_ROOT);
      return out;
    }

    if (location.protocol === "file:") {
      add("../../data/report/");
      return out;
    }

    if (location.protocol !== "http:" && location.protocol !== "https:") {
      add(reportAssetsRoot());
      return out;
    }

    const href = window.location.href;
    const origin = window.location.origin;
    add(new URL(".", href).href);
    add(new URL("/data/report/", origin).href);
    return out;
  }

  /**
   * Find a base that serves employees.json; fills employeeAutocomplete when found.
   * @returns {string} working base ending with /
   */
  async function resolveReportAssetBase() {
    const bases = buildReportAssetBaseCandidates();
    for (let i = 0; i < bases.length; i++) {
      const b = bases[i];
      try {
        const r = await fetch(b + "employees.json", { cache: "no-store" });
        if (!r.ok) continue;
        const t = await r.text();
        if (t.trimStart().startsWith("<")) continue;
        const data = JSON.parse(t);
        if (window.employeeAutocomplete) {
          window.employeeAutocomplete.employees = filterActivityReportEmployeeList(Array.isArray(data) ? data : []);
          try {
            const ctuRes = await fetch(b + "ctu_staff_suggestions.json", { cache: "no-store" });
            if (ctuRes.ok) {
              const ctuText = await ctuRes.text();
              if (!ctuText.trimStart().startsWith("<")) {
                const ctuData = JSON.parse(ctuText);
                window.employeeAutocomplete.ctuSuggestions = Array.isArray(ctuData) ? ctuData.slice() : [];
              }
            }
          } catch (_) {
            window.employeeAutocomplete.ctuSuggestions = [];
          }
        }
        return b;
      } catch (_) {
        /* try next */
      }
    }
    return null;
  }

  /** Try several URLs for offload latest.json (preview vs Flask). */
  async function resolveOffloadReportJsonUrl() {
    const meta = document.querySelector('meta[name="activity-report-offload-json"]');
    if (meta && String(meta.content || "").trim()) return String(meta.content).trim();
    if (typeof window.ACTIVITY_REPORT_OFFLOAD_JSON === "string" && window.ACTIVITY_REPORT_OFFLOAD_JSON.trim()) {
      return window.ACTIVITY_REPORT_OFFLOAD_JSON.trim();
    }
    if (location.protocol === "file:") return "../../data/offload/report/latest.json";

    const urls = [];
    urls.push(new URL("../offload/report/latest.json", window.location.href).href);
    urls.push(new URL("/data/offload/report/latest.json", window.location.origin).href);

    for (let i = 0; i < urls.length; i++) {
      const u = urls[i];
      try {
        const r = await fetch(u, { cache: "no-store", method: "GET" });
        if (!r.ok) continue;
        const t = await r.text();
        if (t.trimStart().startsWith("<")) continue;
        JSON.parse(t);
        return u;
      } catch (_) {
        /* next */
      }
    }
    return urls[0];
  }

  function filterActivityReportEmployeeList(arr) {
    if (!Array.isArray(arr)) return [];
    return arr.slice();
  }

  function canUseSameOriginApi() {
    if (location.protocol !== "http:" && location.protocol !== "https:") return false;
    const host = String(location.hostname || "").trim().toLowerCase();
    if (!host) return false;
    return !host.endsWith("github.io");
  }

  function configuredApiBase() {
    const meta = document.querySelector('meta[name="flight-hints-api-base"]');
    if (meta && String(meta.content || "").trim()) {
      return String(meta.content).trim().replace(/\/$/, "");
    }
    if (typeof window.FLIGHT_HINTS_API_BASE === "string" && window.FLIGHT_HINTS_API_BASE.trim()) {
      return window.FLIGHT_HINTS_API_BASE.trim().replace(/\/$/, "");
    }
    return "";
  }

  function apiUrl(path) {
    const base = configuredApiBase();
    if (!base) return path;
    return `${base}${path}`;
  }

  function canUseServerGmailApi() {
    return canUseSameOriginApi() || !!configuredApiBase();
  }

  function stripExcludedEmployeesFromManpower() {
    /* Roster lines come from JSON; SN990737 appears under Inventory when on shift (server rules). */
  }

  /**
   * Saved drafts replace all of manpowerSections and can drop roster-driven rows (e.g. SN990737 → Inventory).
   * Always copy Inventory + Support Team from the active server shift pack so those sections stay accurate.
   */
  function applyServerInventorySupportTailSections() {
    if (!state.shiftsFromServer || !state.activeShift || !Array.isArray(state.manpowerSections)) return;
    const pack = state.shiftsFromServer[state.activeShift];
    if (!pack || !Array.isArray(pack.manpowerSections)) return;
    const byTitle = {};
    pack.manpowerSections.forEach((s) => {
      const t = String(s.title || "").trim();
      if (t === "Inventory" || t === "Support Team") {
        byTitle[t] = Array.isArray(s.items) ? deepClone(s.items) : [];
      }
    });
    state.manpowerSections.forEach((sec) => {
      const t = String(sec.title || "").trim();
      if (Object.prototype.hasOwnProperty.call(byTitle, t)) {
        sec.items = byTitle[t];
        ensureManpowerRowForEditing(sec);
      }
    });
  }

  /** Display name for signature: drop SN prefix from roster lines like "SN12345 Name". */
  function displayNameFromRosterEntry(entry) {
    const s = String(entry || "").trim();
    if (!s) return "";
    const noBullet = s.replace(/^[\s\u2022\-*]+/, "");
    const noSn = noBullet.replace(/^SN\d+\s+/i, "").trim() || noBullet;
    const noRoleTail = noSn.replace(/\s*-\s*(supervisor|duty supervisor)\b.*$/i, "").trim();
    return noRoleTail || noSn;
  }

  /** First non-empty Supervisor section line. */
  function getDutySupervisorDisplayName() {
    const pickFromSections = (sections) => {
      const list = Array.isArray(sections) ? sections : [];
      const sup = list.find((sec) => /supervisor/i.test(String((sec && sec.title) || "").trim()));
      if (!sup || !Array.isArray(sup.items)) return "";
      const line = sup.items.map((x) => String(x || "").trim()).find((x) => x);
      return line ? displayNameFromRosterEntry(line) : "";
    };

    const pickFromShiftMap = (shiftMap, preferredKey) => {
      if (!shiftMap || typeof shiftMap !== "object") return "";
      const order = [];
      if (preferredKey && shiftMap[preferredKey]) order.push(preferredKey);
      ["morning", "afternoon", "night"].forEach((k) => {
        if (shiftMap[k] && !order.includes(k)) order.push(k);
      });
      for (const k of order) {
        const pack = shiftMap[k];
        const v = pickFromSections(pack && pack.manpowerSections);
        if (v) return v;
      }
      return "";
    };

    // 1) Current editable manpower list.
    const direct = pickFromSections(state.manpowerSections);
    if (direct) return direct;

    // 2) Active shift from server payload.
    const activePack =
      state.shiftsFromServer && state.activeShift ? state.shiftsFromServer[state.activeShift] : null;
    const fromActiveShift = pickFromSections(activePack && activePack.manpowerSections);
    if (fromActiveShift) return fromActiveShift;

    // 3) Any available shift as last fallback.
    const fromShiftMap = pickFromShiftMap(state.shiftsFromServer, state.activeShift);
    if (fromShiftMap) return fromShiftMap;

    // 4) Raw fetched payload (before any draft overrides).
    const raw = state._fetchedReportJson && typeof state._fetchedReportJson === "object" ? state._fetchedReportJson : null;
    if (raw) {
      const fromTop = pickFromSections(raw.manpowerSections);
      if (fromTop) return fromTop;
      const fromRawShifts = pickFromShiftMap(raw.shifts, state.activeShift || raw.defaultShift);
      if (fromRawShifts) return fromRawShifts;
    }

    // 5) Reset baseline snapshot (captured right after server data apply).
    if (state._resetBaseline && typeof state._resetBaseline === "object") {
      const b = state._resetBaseline;
      const fromBaselineTop = pickFromSections(b.manpowerSections);
      if (fromBaselineTop) return fromBaselineTop;
      const fromBaselineShifts = pickFromShiftMap(b.shiftsFromServer, b.activeShift);
      if (fromBaselineShifts) return fromBaselineShifts;
    }

    // 6) Live DOM fallback (when UI has loaded rows but state is stale).
    try {
      const sectionNodes = Array.from(document.querySelectorAll(".manpower-section"));
      for (const sec of sectionNodes) {
        const titleEl = sec.querySelector(".manpower-section-title");
        const title = String((titleEl && (titleEl.value || titleEl.textContent)) || "").trim();
        if (!/supervisor/i.test(title)) continue;
        const lineEls = Array.from(sec.querySelectorAll(".manpower-line"));
        for (const inp of lineEls) {
          const v = String((inp && inp.value) || "").trim();
          if (!v) continue;
          const name = displayNameFromRosterEntry(v);
          if (name) return name;
        }
      }
    } catch (_) {
      /* ignore DOM fallback errors */
    }
    return "";
  }

  const state = {
    shiftMeta: { title: "", date: "", time: "" },
    operationalActivities: [
      { title: "Load Plan", items: [""] },
      { title: "Advance Loading", items: [""] },
      { title: "CSD Rescreening", items: [""] },
    ],
    briefings: [
      "Safety toolbox conducted.",
      "ULD and net serviceability checked.",
      "Staff reminded about punctuality, proper cargo loading/counting, and no mobile phone use while driving.",
    ],
    flightPerformance: "ALL FLIGHTS DEPARTED ON TIME; NO DELAY RELATED TO CARGO.",
    operationalNotes: [
      "ALL FLIGHTS DEPARTED ON TIME AS PER RDM MR. SALEH.",
      "DG EMBARGO STATION CHECK COMPLETED.",
      "PIGEONHOLE CHECK DONE FOR ANY PENDING DOCUMENTS.",
    ],
    checksCompliance: "DG EMBARGO STATION CHECK DONE. AWB LEFT BEHIND: NIL.",
    offloads: [
      {
        item: 1,
        date: "",
        flight: "",
        std: "",
        destination: "",
        emailTime: "",
        rampReceived: "",
        trolley: "",
        cmsCompleted: "",
        piecesVerification: "",
        reason: "",
        remarks: "",
      },
    ],
    safety: "SAFETY BRIEFING CONDUCTED TO ALL STAFF, DRIVERS AND PORTERS.",
    manpowerSections: [],
    equipmentStatus: "ALL EQUIPMENT ARE OK.",
    handoverDetails: "READ AND SIGN. SHELL & AL-MAHA CARD FUEL. DIP MAIL CAGE KEYS.",
    otherText: "",
    specialHO: "",
    recipients: { to: ["ops@company.com", "supervisor@company.com"], cc: [], bcc: [] },
    scheduledSendAt: "",
    scheduledSendEnabled: false,
    _scheduledSendLastFiredAt: "",
    shiftsFromServer: null,
    activeShift: null,
    datesList: [],
    availableDates: [],
    activeDate: null,
    loadErrorMessage: "",
    noDataMode: false,
    _fetchedReportJson: null,
    _resetBaseline: null,
  };

  const offloadFieldOrder = [
    "date",
    "flight",
    "std",
    "destination",
    "emailTime",
    "rampReceived",
    "trolley",
    "cmsCompleted",
    "piecesVerification",
    "reason",
    "remarks",
  ];

  const SHIFT_TAB_LABELS = [
    { key: "morning", label: "Morning" },
    { key: "afternoon", label: "Afternoon" },
    { key: "night", label: "Night" },
  ];

  let _scheduledSendTimer = null;
  let _autoTodayWatcherTimer = null;
  let _autoTodayLastIso = "";
  let _gmailStatus = { configured: false, authorized: false };
  const EMAIL_BUTTON_DEFAULT_TEXT = "Send Report";

  /** Bump when draft shape changes; old keys are ignored so stale roster lists are not restored. */
  const DRAFT_STORAGE_PREFIX = "activity-report-draft-v4";

  function el(id) {
    return document.getElementById(id);
  }

  function isFlightSuggestOpenFor(target) {
    return (
      window.flightAutocomplete &&
      window.flightAutocomplete.activeInput === target &&
      Array.isArray(window.flightAutocomplete.activeMatches) &&
      window.flightAutocomplete.activeMatches.length > 0
    );
  }

  function isPhraseSuggestOpenFor(target) {
    return (
      window.phraseAutocomplete &&
      window.phraseAutocomplete.activeInput === target &&
      Array.isArray(window.phraseAutocomplete.activeItems) &&
      window.phraseAutocomplete.activeItems.length > 0
    );
  }

  function isAnySuggestOpenFor(target) {
    return isFlightSuggestOpenFor(target) || isPhraseSuggestOpenFor(target);
  }

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function captureResetSnapshot() {
    return {
      operationalActivities: deepClone(state.operationalActivities),
      briefings: deepClone(state.briefings),
      flightPerformance: state.flightPerformance,
      operationalNotes: deepClone(state.operationalNotes),
      checksCompliance: state.checksCompliance,
      offloads: deepClone(state.offloads),
      safety: state.safety,
      manpowerSections: deepClone(state.manpowerSections),
      equipmentStatus: state.equipmentStatus,
      handoverDetails: state.handoverDetails,
      otherText: state.otherText,
      specialHO: state.specialHO,
      recipients: deepClone(state.recipients),
      scheduledSendAt: state.scheduledSendAt,
      scheduledSendEnabled: state.scheduledSendEnabled,
      _scheduledSendLastFiredAt: state._scheduledSendLastFiredAt,
      shiftMeta: { ...state.shiftMeta },
      activeShift: state.activeShift,
    };
  }

  function restoreResetSnapshot() {
    const b = state._resetBaseline;
    if (!b) return;
    state.operationalActivities = deepClone(b.operationalActivities);
    state.briefings = deepClone(b.briefings);
    state.flightPerformance = b.flightPerformance;
    state.operationalNotes = deepClone(b.operationalNotes);
    state.checksCompliance = b.checksCompliance;
    state.offloads = deepClone(b.offloads);
    state.safety = b.safety;
    state.manpowerSections = deepClone(b.manpowerSections);
    state.equipmentStatus = b.equipmentStatus;
    state.handoverDetails = b.handoverDetails;
    state.otherText = b.otherText;
    state.specialHO = b.specialHO;
    state.recipients = deepClone(b.recipients);
    state.scheduledSendAt = b.scheduledSendAt || "";
    state.scheduledSendEnabled = !!b.scheduledSendEnabled;
    state._scheduledSendLastFiredAt = b._scheduledSendLastFiredAt || "";
    state.shiftMeta = { ...b.shiftMeta };
    if (b.activeShift != null) state.activeShift = b.activeShift;
  }

  function toIsoDateLocal(n) {
    const y = n.getFullYear();
    const mo = String(n.getMonth() + 1).padStart(2, "0");
    const da = String(n.getDate()).padStart(2, "0");
    return `${y}-${mo}-${da}`;
  }

  /**
   * Operational report date YYYY-MM-DD.
   * Night shift is 21:00-06:00, so 00:00-05:59 belongs to the previous day.
   */
  function getReportDateIsoLocal() {
    const n = new Date();
    if (n.getHours() < 6) {
      const prev = new Date(n);
      prev.setDate(prev.getDate() - 1);
      return toIsoDateLocal(prev);
    }
    return toIsoDateLocal(n);
  }

  function isIsoDate(value) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
  }

  /** Keep today's date first, with the rest newest-to-oldest. */
  function buildDatesListTodayFirst(sourceDates, todayIso) {
    const today = isIsoDate(todayIso) ? String(todayIso).trim() : getReportDateIsoLocal();
    const rest = [];
    const seen = new Set([today]);
    (Array.isArray(sourceDates) ? sourceDates : []).forEach((raw) => {
      const iso = String(raw || "").trim();
      if (!isIsoDate(iso) || seen.has(iso)) return;
      seen.add(iso);
      rest.push(iso);
    });
    rest.sort((a, b) => (a < b ? 1 : a > b ? -1 : 0));
    return [today, ...rest];
  }

  /**
   * Same windows and order as read_roster.SHIFTS / get_current_shift():
   * morning 06–15, afternoon 13–22, night 21–06 (first match wins).
   */
  function getCurrentShiftKey() {
    const now = new Date();
    const mins = now.getHours() * 60 + now.getMinutes() + now.getSeconds() / 60;

    function inRange(startStr, endStr) {
      const parse = (s) => {
        const [h, m] = s.split(":").map(Number);
        return h * 60 + m;
      };
      const start = parse(startStr);
      const end = parse(endStr);
      if (start <= end) {
        return mins >= start && mins <= end;
      }
      return mins >= start || mins <= end;
    }

    const ordered = [
      ["morning", "06:00", "15:00"],
      ["afternoon", "13:00", "22:00"],
      ["night", "21:00", "06:00"]
    ];
    for (let i = 0; i < ordered.length; i++) {
      const name = ordered[i][0];
      if (inRange(ordered[i][1], ordered[i][2])) return name;
    }
    return "morning";
  }

  function getUseAutoToday() {
    const meta = document.querySelector('meta[name="activity-report-auto-today"]');
    if (meta && String(meta.getAttribute("content") || "").trim() === "0") return false;
    return true;
  }

  function parseInitialLaunchContext() {
    try {
      const params = new URLSearchParams(window.location.search || "");
      const dateRaw = String(params.get("date") || "").trim();
      const shiftRaw = String(params.get("shift") || "").trim().toLowerCase();
      const date = /^\d{4}-\d{2}-\d{2}$/.test(dateRaw) ? dateRaw : "";
      const shift = shiftRaw === "morning" || shiftRaw === "afternoon" || shiftRaw === "night" ? shiftRaw : "";
      return { date, shift };
    } catch (_) {
      return { date: "", shift: "" };
    }
  }

  /** Align JSON dates with the selected report day (usually today). */
  function syncFetchedContentDates(data) {
    const d = state.activeDate;
    if (!d || !data) return;
    if (data.shiftMeta) data.shiftMeta.date = d;
    if (data.shifts && typeof data.shifts === "object") {
      Object.keys(data.shifts).forEach((k) => {
        const pack = data.shifts[k];
        if (pack && pack.shiftMeta) pack.shiftMeta.date = d;
      });
    }
  }

  function formatDisplayDate(dateStr) {
    if (!dateStr) return "";
    const m = String(dateStr).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const d = parseInt(m[3], 10);
      const mo = months[parseInt(m[2], 10) - 1] || m[2];
      return `${d} ${mo} ${m[1]}`;
    }
    return dateStr;
  }

  function dateTabLabel(iso) {
    return formatDisplayDate(iso) || iso;
  }

  function metaLineText() {
    const d = formatDisplayDate(state.shiftMeta.date || "");
    const t = state.shiftMeta.time || "";
    const title = state.shiftMeta.title || "";
    return `${d}  |  ${t}  |  ${title}`;
  }

  function draftStorageKey() {
    const d = state.activeDate;
    const s = state.activeShift;
    if (d && s) return `${DRAFT_STORAGE_PREFIX}-${d}-${s}`;
    if (d) return `${DRAFT_STORAGE_PREFIX}-${d}`;
    if (s) return `${DRAFT_STORAGE_PREFIX}-${s}`;
    return DRAFT_STORAGE_PREFIX;
  }

  /**
   * Prefer per-day roster under by-date/ (avoids silently using root latest.json from another day).
   * Do not gate on datesList — if today is missing from the index we still load by-date/today/.
   */
  function reportJsonUrl() {
    const d = state.activeDate && String(state.activeDate).trim();
    if (d && /^\d{4}-\d{2}-\d{2}$/.test(d)) {
      return `by-date/${d}/latest.json`;
    }
    return "latest.json";
  }

  /** Calendar day embedded in fetched JSON (before syncFetchedContentDates rewrites labels). */
  function getPayloadCalendarDate(data) {
    if (!data || typeof data !== "object") return "";
    const top = data.shiftMeta && data.shiftMeta.date;
    if (top && /^\d{4}-\d{2}-\d{2}$/.test(String(top).trim())) return String(top).trim();
    const keys = ["morning", "afternoon", "night"];
    for (let i = 0; i < keys.length; i++) {
      const pack = data.shifts && data.shifts[keys[i]];
      const sd = pack && pack.shiftMeta && pack.shiftMeta.date;
      if (sd && /^\d{4}-\d{2}-\d{2}$/.test(String(sd).trim())) return String(sd).trim();
    }
    return "";
  }

  function applyShiftFromServer(key) {
    const pack = state.shiftsFromServer && state.shiftsFromServer[key];
    if (!pack) return;
    state.shiftMeta = { ...(pack.shiftMeta || {}), key };
    state.manpowerSections = JSON.parse(JSON.stringify(pack.manpowerSections || []));
    stripExcludedEmployeesFromManpower();
  }

  function syncManpowerFromServerShifts() {
    if (!state.shiftsFromServer || !state.activeShift) return;
    const pack = state.shiftsFromServer[state.activeShift];
    if (!pack || !Array.isArray(pack.manpowerSections)) return;
    state.manpowerSections = JSON.parse(JSON.stringify(pack.manpowerSections));
    stripExcludedEmployeesFromManpower();
  }

  async function switchShift(key) {
    if (!state.shiftsFromServer || key === state.activeShift) return;
    if (state.noDataMode) {
      state.activeShift = key;
      applyShiftFromServer(key);
      if (state.offloads[0]) {
        state.offloads[0].date = window.offloadLoader
          ? window.offloadLoader.normalizeOffloadDate(state.shiftMeta.date || "")
          : state.shiftMeta.date || "";
      }
      saveDraft();
      renderAll();
      return;
    }
    saveDraft();
    state.activeShift = key;
    if (state._fetchedReportJson) {
      await finalizeAfterServerData(state._fetchedReportJson);
    } else {
      applyShiftFromServer(key);
      await applyOffloadFromServer();
      state._resetBaseline = captureResetSnapshot();
      const applied = loadDraft();
      if (state.shiftsFromServer && state.activeShift && !applied) syncManpowerFromServerShifts();
      if (window.offloadLoader) window.offloadLoader.normalizeOffloadRows(state);
    }
    renderAll();
  }

  async function switchDate(dateStr) {
    if (!dateStr || dateStr === state.activeDate) return;
    if (state.datesList.length && !state.datesList.includes(dateStr)) return;
    const prevDate = state.activeDate;
    saveDraft();
    state.activeDate = dateStr;
    state.activeShift = null;
    state._fetchedReportJson = null;
    try {
      await loadReportPayload();
      renderAll();
    } catch (err) {
      const msg = err && err.message ? String(err.message) : "Failed to load selected date";
      state.activeDate = prevDate || state.activeDate;
      if (state.activeDate) {
        try {
          await loadReportPayload();
          renderAll();
          return;
        } catch (_) {
          /* fallback to missing view below */
        }
      }
      applyMissingDateView(msg);
    }
  }

  function startAutoTodayWatcher() {
    if (_autoTodayWatcherTimer) return;
    _autoTodayLastIso = getReportDateIsoLocal();
    _autoTodayWatcherTimer = window.setInterval(() => {
      const nowIso = getReportDateIsoLocal();
      if (nowIso === _autoTodayLastIso) return;
      _autoTodayLastIso = nowIso;
      state.datesList = buildDatesListTodayFirst(state.availableDates, nowIso);
      if (!getUseAutoToday()) {
        renderDateTabs();
        syncFlightSuggestIsoDate();
        return;
      }
      if (state.activeDate === nowIso) {
        renderDateTabs();
        syncFlightSuggestIsoDate();
        return;
      }
      switchDate(nowIso).catch((err) => {
        const msg = err && err.message ? String(err.message) : "Failed to load selected date";
        state.activeDate = nowIso;
        applyMissingDateView(msg);
      });
    }, 60000);
  }

  function renderShiftTabs() {
    const wrap = el("shiftTabs");
    if (!state.shiftsFromServer) {
      wrap.hidden = true;
      wrap.innerHTML = "";
      return;
    }
    wrap.hidden = false;
    wrap.innerHTML = "";
    SHIFT_TAB_LABELS.forEach(({ key, label }) => {
      if (!state.shiftsFromServer[key]) return;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "shift-tab" + (state.activeShift === key ? " active" : "");
      btn.textContent = label;
      btn.addEventListener("click", () => {
        switchShift(key).catch(console.error);
      });
      wrap.appendChild(btn);
    });
  }

  function renderDateTabs() {
    const wrap = el("dateTabs");
    if (!state.datesList || state.datesList.length <= 1) {
      wrap.hidden = true;
      wrap.innerHTML = "";
      return;
    }
    wrap.hidden = false;
    wrap.innerHTML = "";
    state.datesList.forEach((iso) => {
      const isAvailable =
        !Array.isArray(state.availableDates) ||
        state.availableDates.length === 0 ||
        state.availableDates.includes(iso);
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "date-tab" + (state.activeDate === iso ? " active" : "");
      btn.textContent = dateTabLabel(iso);
      btn.title = isAvailable ? iso : `${iso} (no data yet; will seed from latest available day)`;
      if (!isAvailable) btn.style.opacity = "0.82";
      btn.addEventListener("click", () => switchDate(iso));
      wrap.appendChild(btn);
    });
  }

  async function loadDatesIndex() {
    const candidates = [];
    if (location.protocol === "http:" || location.protocol === "https:") {
      candidates.push(new URL("dates_index.json", window.location.href).href);
      candidates.push(new URL("/data/report/dates_index.json", window.location.origin).href);
    } else {
      candidates.push("dates_index.json");
    }
    for (let i = 0; i < candidates.length; i++) {
      try {
        const r = await fetch(candidates[i], { cache: "no-store" });
        if (!r.ok) continue;
        return await r.json();
      } catch {
        /* next */
      }
    }
    return null;
  }

  async function applyOffloadFromServer() {
    if (state.offloads[0]) {
      state.offloads[0].date = window.offloadLoader
        ? window.offloadLoader.normalizeOffloadDate(state.shiftMeta.date || "")
        : state.shiftMeta.date || "";
    }
    if (window.offloadLoader) {
      const offloadUrl = await resolveOffloadReportJsonUrl();
      const offloadData = await window.offloadLoader.load(offloadUrl);
      window.offloadLoader.applyToState(offloadData, state);
    }
  }

  async function finalizeAfterServerData(data) {
    state.loadErrorMessage = "";
    state.noDataMode = false;
    syncFetchedContentDates(data);
    state._fetchedReportJson = deepClone(data);

    if (data.shifts && typeof data.shifts === "object" && Object.keys(data.shifts).length) {
      state.shiftsFromServer = data.shifts;
      if (!state.activeShift || !state.shiftsFromServer[state.activeShift]) {
        let def = data.defaultShift || "morning";
        const todayIso = getReportDateIsoLocal();
        const reportDay =
          (state.activeDate && String(state.activeDate).trim()) ||
          String((data.shiftMeta && data.shiftMeta.date) || "").trim();
        if (reportDay === todayIso) {
          const clock = getCurrentShiftKey();
          if (data.shifts[clock]) def = clock;
        }
        if (!state.shiftsFromServer[def]) def = Object.keys(state.shiftsFromServer)[0];
        state.activeShift = def;
      }
      applyShiftFromServer(state.activeShift);
    } else {
      state.shiftsFromServer = null;
      state.activeShift = null;
      if (data.shiftMeta) state.shiftMeta = data.shiftMeta;
      if (Array.isArray(data.manpowerSections)) state.manpowerSections = data.manpowerSections;
    }

    await applyOffloadFromServer();

    state._resetBaseline = captureResetSnapshot();
    const appliedManpower = loadDraft();
    if (state.shiftsFromServer && state.activeShift && !appliedManpower) {
      syncManpowerFromServerShifts();
    }
    applyServerInventorySupportTailSections();
    if (window.offloadLoader) window.offloadLoader.normalizeOffloadRows(state);
    stripExcludedEmployeesFromManpower();
  }

  async function loadReportPayload() {
    const want = state.activeDate && String(state.activeDate).trim();
    const wantMissingFromIndex =
      !!want &&
      Array.isArray(state.availableDates) &&
      state.availableDates.length > 0 &&
      !state.availableDates.includes(want);

    if (wantMissingFromIndex) {
      throw new Error(
        `No report data exists for ${want} yet. Generate today's report files first (by-date/${want}/latest.json), then reload.`
      );
    }

    const primaryUrl = reportJsonUrl();
    let url = primaryUrl;
    let r = await fetch(url, { cache: "no-store" });
    let usedRootLatestFallback = false;
    if (!r.ok) {
      usedRootLatestFallback = true;
      url = "latest.json";
      r = await fetch(url, { cache: "no-store" });
    }
    if (!r.ok) throw new Error(`Report data not found (${url})`);
    const data = await r.json();

    const got = getPayloadCalendarDate(data);
    if (want && /^\d{4}-\d{2}-\d{2}$/.test(want) && got && got !== want) {
      const hint = usedRootLatestFallback
        ? ` Folder by-date/${want}/ is missing or unreachable; root latest.json is for ${got}.`
        : "";
      throw new Error(
        `Report JSON is for ${got}, but the selected day is ${want}.${hint} Rebuild data for ${want} or pick date ${got} in the date tabs.`
      );
    }

    await finalizeAfterServerData(data);
  }

  async function loadData() {
    try {
      if (String(location.protocol || "").toLowerCase() === "file:") {
        try {
          const nextUrl = `http://localhost:8000/data/report/offload_report.html${location.search || ""}${location.hash || ""}`;
          window.location.replace(nextUrl);
          return;
        } catch (_) {}
        applyMissingDateView(
          "This report cannot load data over file:// due to browser security. Run start-local-server.bat and open http://localhost:8000/data/report/offload_report.html"
        );
        return;
      }
      const launch = parseInitialLaunchContext();
      const foundBase = await resolveReportAssetBase();
      const assetBase = foundBase || reportAssetsRoot();
      if (foundBase === null && window.employeeAutocomplete) {
        await window.employeeAutocomplete.load(assetBase + "employees.json");
      }

      if (window.flightAutocomplete) {
        if (canUseSameOriginApi()) {
          await window.flightAutocomplete.load("/api/live-flights", { silent: true });
        }
        if (!Array.isArray(window.flightAutocomplete.flights) || !window.flightAutocomplete.flights.length) {
          await window.flightAutocomplete.load(assetBase + "flights.json");
        }
      }
      if (window.phraseAutocomplete) {
        await window.phraseAutocomplete.load(assetBase + "phrases.json");
        await window.phraseAutocomplete.loadCsdDestinationHints(assetBase + "csd-wy-ov-destinations.json");
      }

      if (window.flightHintCache) {
        await window.flightHintCache.hydrate({
          fallbackUrl: assetBase + "flight-hints.json"
        });
      }
      if (window.csdRouteHintCache) {
        await window.csdRouteHintCache.hydrate({
          fallbackUrl: assetBase + "csd-route-hints.json"
        });
      }
      if (window.phraseUsageCache) {
        await window.phraseUsageCache.hydrate({
          fallbackUrl: assetBase + "phrase-usage.json"
        });
      }
      if (window.manpowerRoleHintCache) {
        await window.manpowerRoleHintCache.hydrate({
          fallbackUrl: assetBase + "manpower-role-hints.json"
        });
        await window.manpowerRoleHintCache.loadRoleOptions(assetBase + "manpower-role-options.json");
      }
      if (window.recipientsCache) {
        await window.recipientsCache.hydrate({
          fallbackUrl: assetBase + "recipients.json"
        });
        state.recipients = normalizeRecipientsShape(window.recipientsCache.getAll());
      }

      const idx = await loadDatesIndex();
      const todayIso = getReportDateIsoLocal();
      state.availableDates = idx && Array.isArray(idx.dates) ? idx.dates.slice().sort() : [];
      state.datesList = buildDatesListTodayFirst(state.availableDates, todayIso);
      if (launch.date) {
        state.activeDate = launch.date;
      } else if (getUseAutoToday()) {
        state.activeDate = todayIso;
      } else if (idx && Array.isArray(idx.dates) && idx.dates.length) {
        const def = idx.default && state.datesList.includes(idx.default) ? idx.default : state.datesList[0];
        state.activeDate = def;
      } else {
        state.activeDate = state.datesList[0] || null;
      }
      if (launch.shift) state.activeShift = launch.shift;

      // If today's date has no generated report yet, fallback to the latest available date
      // so manpower names remain visible instead of entering empty no-data mode.
      if (
        !launch.date &&
        state.activeDate &&
        Array.isArray(state.availableDates) &&
        state.availableDates.length &&
        !state.availableDates.includes(state.activeDate)
      ) {
        const fallbackDate =
          (idx && idx.default && state.availableDates.includes(idx.default) && idx.default) ||
          state.availableDates[state.availableDates.length - 1];
        if (fallbackDate) state.activeDate = fallbackDate;
      }

      await loadReportPayload();
      renderAll();
      ensureSignatureBadgesDataUriForClipboard().catch(() => {});
    } catch (err) {
      console.error(err);
      const msg = err && err.message ? String(err.message) : "Failed to load data";
      applyMissingDateView(msg.length > 220 ? msg.slice(0, 217) + "…" : msg);
    }
  }

  function blankOffloadRow(item, dateText) {
    return {
      item,
      date: window.offloadLoader ? window.offloadLoader.normalizeOffloadDate(dateText || "") : dateText || "",
      flight: "",
      std: "",
      destination: "",
      emailTime: "",
      rampReceived: "",
      trolley: "",
      cmsCompleted: "",
      piecesVerification: "",
      reason: "",
      remarks: "",
    };
  }

  function buildEmptyShiftsForDate(dateText) {
    const date = String(dateText || "").trim();
    return {
      morning: {
        shiftMeta: { key: "morning", title: "Morning Shift", date, time: "06:00 - 15:00" },
        manpowerSections: []
      },
      afternoon: {
        shiftMeta: { key: "afternoon", title: "Afternoon Shift", date, time: "13:00 - 22:00" },
        manpowerSections: []
      },
      night: {
        shiftMeta: { key: "night", title: "Night Shift", date, time: "21:00 - 06:00" },
        manpowerSections: []
      }
    };
  }

  function applyMissingDateView(message) {
    const dateText = state.activeDate || getReportDateIsoLocal();
    state.loadErrorMessage = String(message || "No report data found for selected date.");
    state.noDataMode = true;
    const hasExistingShiftData =
      !!state.shiftsFromServer &&
      Object.values(state.shiftsFromServer).some((pack) => {
        const sections = Array.isArray(pack && pack.manpowerSections) ? pack.manpowerSections : [];
        return sections.some((sec) => Array.isArray(sec && sec.items) && sec.items.some((x) => String(x || "").trim()));
      });

    // Keep last valid manpower visible when selected date has no generated files yet.
    if (!hasExistingShiftData) {
      state.shiftsFromServer = buildEmptyShiftsForDate(dateText);
      state.activeShift = state.shiftsFromServer[getCurrentShiftKey()] ? getCurrentShiftKey() : "morning";
      applyShiftFromServer(state.activeShift);
      state.manpowerSections = [];
    }
    state.offloads = [blankOffloadRow(1, dateText)];
    renderAll();
  }

  function renderWebSignaturePreview() {
    const nameEl = el("webSignatureSupervisorName");
    if (!nameEl) return;
    const name = getDutySupervisorDisplayName();
    if (name) {
      nameEl.textContent = name;
      nameEl.classList.remove("web-sig-name-empty");
    } else {
      nameEl.textContent = "(Add name under Manpower → Supervisor)";
      nameEl.classList.add("web-sig-name-empty");
    }
  }

  function renderAll() {
    state.flightPerformance = normalizeIndentedBullets(toSentenceCaseText(state.flightPerformance));
    state.checksCompliance = normalizeIndentedBullets(toSentenceCaseText(state.checksCompliance));
    state.equipmentStatus = normalizeIndentedBullets(toSentenceCaseText(state.equipmentStatus));
    state.handoverDetails = normalizeIndentedBullets(toSentenceCaseText(state.handoverDetails));
    state.specialHO = normalizeIndentedBullets(toSentenceCaseText(state.specialHO));
    state.otherText = normalizeIndentedBullets(toSentenceCaseText(state.otherText));
    renderMeta();
    renderOperationalActivities();
    renderBriefings();
    renderOperationalNotes();
    renderOffloads();
    renderManpower();
    renderRecipients();
    renderScheduledSendStatus();

    el("flightPerformance").value = state.flightPerformance;
    el("checksCompliance").value = state.checksCompliance;
    el("safety").value = state.safety;
    el("equipmentStatus").value = state.equipmentStatus;
    el("handoverDetails").value = state.handoverDetails;
    el("otherText").value = state.otherText;
    el("specialHO").value = state.specialHO;

    bindAutocompleteWiring();
    scheduleSendTimerFromState();
    refreshGmailStatus().catch(() => {});
  }

  function syncFlightSuggestIsoDate() {
    try {
      const iso = (state.activeDate || state.shiftMeta.date || "").trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
        window.__flightSuggestIsoDate = iso;
        return;
      }
      const d = new Date();
      window.__flightSuggestIsoDate = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
        d.getDate()
      ).padStart(2, "0")}`;
    } catch {
      window.__flightSuggestIsoDate = "";
    }
  }

  function renderMeta() {
    syncFlightSuggestIsoDate();
    if (state.loadErrorMessage) {
      el("shiftMetaText").textContent = state.loadErrorMessage;
      renderDateTabs();
      renderShiftTabs();
      return;
    }
    const d = formatDisplayDate(state.shiftMeta.date || "");
    const t = state.shiftMeta.time || "";
    const title = state.shiftMeta.title || "";
    el("shiftMetaText").innerHTML = `Shift Date: <span class="shift-meta-highlight">${escapeHtml(
      d
    )}</span> &nbsp;|&nbsp; Time: <span class="shift-meta-highlight">${escapeHtml(
      t
    )}</span> &nbsp;|&nbsp; <span class="shift-meta-highlight">${escapeHtml(title)}</span>`;
    renderDateTabs();
    renderShiftTabs();
  }

  function makeEditableRow(value, onInput, onKeyDown, onDelete) {
    const row = document.createElement("div");
    row.className = "line-item";

    const input = document.createElement("input");
    input.type = "text";
    input.value = value;
    input.oninput = (e) => onInput(e.target.value);
    input.onkeydown = onKeyDown;

    const del = document.createElement("button");
    del.type = "button";
    del.className = "delete-btn hidden-print";
    del.textContent = "✕";
    del.onclick = onDelete;

    row.appendChild(input);
    row.appendChild(del);
    return { row, input };
  }

  function splitOperationalTwoFieldLine(raw) {
    const line = String(raw || "").trim();
    if (!line) return { flight: "", phrase: "" };
    const dashSplit = line.match(/^(.+?)\s-\s(.+)$/);
    if (dashSplit) {
      return { flight: String(dashSplit[1] || "").trim(), phrase: String(dashSplit[2] || "").trim() };
    }
    const slashPattern = /^(\S+\/\d{1,2}[A-Za-z]{3}\/\S+)\s+(.*)$/;
    const slashMatch = line.match(slashPattern);
    if (slashMatch) {
      return { flight: String(slashMatch[1] || "").trim(), phrase: String(slashMatch[2] || "").trim() };
    }
    const hardGap = line.match(/^(.+?)\s{2,}(.+)$/);
    if (hardGap) {
      return { flight: String(hardGap[1] || "").trim(), phrase: String(hardGap[2] || "").trim() };
    }
    return { flight: line, phrase: "" };
  }

  function composeOperationalTwoFieldLine(flight, phrase) {
    const f = String(flight || "").trim();
    const p = String(phrase || "").trim();
    if (f && p) return `${f} - ${p}`;
    return f || p;
  }

  function toSentenceCaseText(value) {
    const src = String(value || "");
    if (!src.trim()) return src;
    const acronyms = new Set(["AWB", "DG", "ULD", "CTU", "STC", "CSD", "IATA", "ISO", "GDP", "RA3", "MHS", "ETA", "ETD", "STD", "WY", "OV"]);
    const lower = src.toLowerCase();
    let out = "";
    let capitalizeNext = true;
    for (let i = 0; i < lower.length; i += 1) {
      const ch = lower[i];
      if (capitalizeNext && /[a-z]/.test(ch)) {
        out += ch.toUpperCase();
        capitalizeNext = false;
      } else {
        out += ch;
      }
      if (/[.!?]/.test(ch)) capitalizeNext = true;
      if (ch === "\n") capitalizeNext = true;
    }
    out = out.replace(/\b[A-Za-z]{2,5}\b/g, (word) => {
      const up = word.toUpperCase();
      return acronyms.has(up) ? up : word;
    });
    return out;
  }

  function makeDualEditableRow(parts, onInput, onKeyDown, onDelete) {
    const row = document.createElement("div");
    row.className = "line-item";
    row.style.display = "flex";
    row.style.alignItems = "center";
    row.style.gap = "8px";

    const flightInput = document.createElement("input");
    flightInput.type = "text";
    flightInput.value = parts.flight || "";
    flightInput.placeholder = "WY101/24APR/LHR";
    flightInput.style.flex = "0 0 260px";
    flightInput.oninput = () => onInput(flightInput.value, phraseInput.value);
    flightInput.onkeydown = onKeyDown;

    const phraseInput = document.createElement("input");
    phraseInput.type = "text";
    phraseInput.value = parts.phrase || "";
    phraseInput.placeholder = "ALL ULD ALLOCATIONS CONFIRMED.";
    phraseInput.style.flex = "1";
    phraseInput.oninput = () => onInput(flightInput.value, phraseInput.value);
    phraseInput.onkeydown = onKeyDown;

    const separator = document.createElement("span");
    separator.className = "opact-separator";
    separator.textContent = "-";
    separator.setAttribute("aria-hidden", "true");

    const del = document.createElement("button");
    del.type = "button";
    del.className = "delete-btn hidden-print";
    del.textContent = "✕";
    del.tabIndex = -1;
    del.onclick = onDelete;

    row.appendChild(flightInput);
    row.appendChild(separator);
    row.appendChild(phraseInput);
    row.appendChild(del);
    return { row, flightInput, phraseInput };
  }

  function handleEnterBackspace(list, index, onInsert, onDelete) {
    return function (e) {
      if (e.key === "Enter") {
        if (isAnySuggestOpenFor(e.target)) return;
        e.preventDefault();
        onInsert(index + 1);
      }
      if (e.key === "Backspace" && !e.target.value && list.length > 1) {
        e.preventDefault();
        onDelete(index);
      }
    };
  }

  function renderOperationalActivities() {
    const wrap = el("operationalActivities");
    wrap.innerHTML = "";

    const opPhraseKeysStatic = ["loadPlan", "advanceLoading", "csdRescreening"];

    state.operationalActivities.forEach((group, groupIndex) => {
      const box = document.createElement("div");
      box.className = "line-group";

      const title = document.createElement("div");
      title.className = "line-group-title";
      title.textContent = `${group.title}:`;
      box.appendChild(title);

      group.items.forEach((item, itemIndex) => {
        const useTwoFields = groupIndex === 0 || groupIndex === 1;
        if (useTwoFields) {
          const parts = splitOperationalTwoFieldLine(item);
          const editable = makeDualEditableRow(
            parts,
            (flightValue, phraseValue) => {
              state.operationalActivities[groupIndex].items[itemIndex] = composeOperationalTwoFieldLine(
                flightValue,
                phraseValue
              );
              saveDraft();
            },
            function (e) {
              if (e.key === "Enter") {
                if (isAnySuggestOpenFor(e.target)) return;
                e.preventDefault();
                state.operationalActivities[groupIndex].items.splice(itemIndex + 1, 0, "");
                saveDraft();
                renderOperationalActivities();
                focusOperational(groupIndex, itemIndex + 1, "flight");
              }
              if (
                e.key === "Backspace" &&
                !editable.flightInput.value.trim() &&
                !editable.phraseInput.value.trim() &&
                state.operationalActivities[groupIndex].items.length > 1
              ) {
                e.preventDefault();
                state.operationalActivities[groupIndex].items.splice(itemIndex, 1);
                saveDraft();
                renderOperationalActivities();
                focusOperational(groupIndex, Math.max(0, itemIndex - 1));
              }
            },
            () => {
              if (state.operationalActivities[groupIndex].items.length > 1) {
                state.operationalActivities[groupIndex].items.splice(itemIndex, 1);
                saveDraft();
                renderOperationalActivities();
              }
            }
          );
          editable.flightInput.dataset.group = groupIndex;
          editable.flightInput.dataset.index = itemIndex;
          editable.flightInput.classList.add("opact-input");
          editable.flightInput.classList.add("opact-flight-input");
          editable.flightInput.dataset.phraseKey = "";
          editable.flightInput.dataset.segment = "flight";

          editable.phraseInput.dataset.group = groupIndex;
          editable.phraseInput.dataset.index = itemIndex;
          editable.phraseInput.classList.add("opact-input");
          editable.phraseInput.classList.add("opact-phrase-input");
          const phraseOnlyKey =
            groupIndex === 0 ? "loadPlanPhraseOnly" : groupIndex === 1 ? "advanceLoadingPhraseOnly" : "";
          editable.phraseInput.dataset.phraseKey = phraseOnlyKey || opPhraseKeysStatic[groupIndex] || "";
          editable.phraseInput.dataset.segment = "phrase";
          editable.row.classList.add("opact-line-item");
          const bullet = document.createElement("span");
          bullet.className = "opact-bullet";
          bullet.setAttribute("aria-hidden", "true");
          bullet.textContent = "\u2022";
          bullet.style.marginLeft = "24px";
          editable.row.insertBefore(bullet, editable.row.firstChild);
          box.appendChild(editable.row);
          return;
        }

        const editable = makeEditableRow(
          item,
          (value) => {
            state.operationalActivities[groupIndex].items[itemIndex] = value;
            saveDraft();
          },
          function (e) {
            if (e.key === "Enter") {
              if (isAnySuggestOpenFor(e.target)) return;
              e.preventDefault();
              state.operationalActivities[groupIndex].items.splice(itemIndex + 1, 0, "");
              saveDraft();
              renderOperationalActivities();
              focusOperational(groupIndex, itemIndex + 1);
            }
            if (e.key === "Backspace" && !e.target.value && state.operationalActivities[groupIndex].items.length > 1) {
              e.preventDefault();
              state.operationalActivities[groupIndex].items.splice(itemIndex, 1);
              saveDraft();
              renderOperationalActivities();
              focusOperational(groupIndex, Math.max(0, itemIndex - 1));
            }
          },
          () => {
            if (state.operationalActivities[groupIndex].items.length > 1) {
              state.operationalActivities[groupIndex].items.splice(itemIndex, 1);
              saveDraft();
              renderOperationalActivities();
            }
          }
        );
        editable.input.dataset.group = groupIndex;
        editable.input.dataset.index = itemIndex;
        editable.input.classList.add("opact-input");
        editable.input.dataset.phraseKey = opPhraseKeysStatic[groupIndex] || "";
        editable.input.dataset.segment = "full";
        if (groupIndex === 2) {
          editable.input.title =
            "After the policy number, press Space — then destination suggestions (codes & routes). Advance Loading phrases never appear here.";
        }
        editable.row.classList.add("opact-line-item");
        const bullet = document.createElement("span");
        bullet.className = "opact-bullet";
        bullet.setAttribute("aria-hidden", "true");
        bullet.textContent = "\u2022";
        bullet.style.marginLeft = "24px";
        editable.row.insertBefore(bullet, editable.row.firstChild);
        box.appendChild(editable.row);
      });

      wrap.appendChild(box);
    });

    bindAutocompleteWiring();
  }

  /** Learn FRA-MNL style routes from CSD lines (blur) → server / localStorage. */
  function attachCsdRouteLearning() {
    if (!window.csdRouteHintCache) return;
    document.querySelectorAll('.opact-input[data-group="2"]').forEach((inp) => {
      if (inp.dataset.csdRouteLearn === "1") return;
      inp.dataset.csdRouteLearn = "1";
      inp.addEventListener("blur", () => {
        window.csdRouteHintCache.recordFromText(inp.value);
      });
    });
  }

  /** Flight + phrase + manpower helpers; call after any DOM that includes .opact-input or static textareas. */
  function bindAutocompleteWiring() {
    attachFlightExpansionHelpers();
    attachPhraseHelpers();
    attachCsdRouteLearning();
  }

  function focusOperational(groupIndex, itemIndex, segment) {
    setTimeout(() => {
      const inputs = Array.from(document.querySelectorAll(".opact-input"));
      const matches = inputs.filter(
        (inp) => +inp.dataset.group === groupIndex && +inp.dataset.index === itemIndex
      );
      if (!matches.length) return;

      const requested = segment ? matches.find((inp) => (inp.dataset.segment || "") === segment) : null;
      const preferred =
        requested ||
        matches.find((inp) => (inp.dataset.segment || "") === "flight") ||
        matches[0];
      preferred.focus();
    }, 0);
  }

  function renderBriefings() {
    const wrap = el("briefings");
    wrap.innerHTML = "";

    state.briefings.forEach((item, index) => {
      const editable = makeEditableRow(
        item,
        (value) => {
          state.briefings[index] = value;
          saveDraft();
        },
        handleEnterBackspace(
          state.briefings,
          index,
          (pos) => {
            state.briefings.splice(pos, 0, "");
            saveDraft();
            renderBriefings();
            focusBriefing(pos);
          },
          (i) => {
            state.briefings.splice(i, 1);
            saveDraft();
            renderBriefings();
            focusBriefing(Math.max(0, i - 1));
          }
        ),
        () => {
          if (state.briefings.length > 1) {
            state.briefings.splice(index, 1);
            saveDraft();
            renderBriefings();
          }
        }
      );
      editable.input.classList.add("briefing-input");
      editable.input.dataset.index = index;
      editable.row.classList.add("bullet-line-item");
      const bullet = document.createElement("span");
      bullet.className = "opact-bullet";
      bullet.setAttribute("aria-hidden", "true");
      bullet.textContent = "\u2022";
      bullet.style.marginLeft = "24px";
      editable.row.insertBefore(bullet, editable.row.firstChild);
      wrap.appendChild(editable.row);
    });
  }

  function focusBriefing(index) {
    setTimeout(() => {
      const input = document.querySelector(`.briefing-input[data-index="${index}"]`);
      if (input) input.focus();
    }, 0);
  }

  function renderOperationalNotes() {
    const wrap = el("operationalNotes");
    wrap.innerHTML = "";

    state.operationalNotes.forEach((item, index) => {
      const editable = makeEditableRow(
        item,
        (value) => {
          state.operationalNotes[index] = value;
          saveDraft();
        },
        handleEnterBackspace(
          state.operationalNotes,
          index,
          (pos) => {
            state.operationalNotes.splice(pos, 0, "");
            saveDraft();
            renderOperationalNotes();
            focusOperationalNote(pos);
          },
          (i) => {
            state.operationalNotes.splice(i, 1);
            saveDraft();
            renderOperationalNotes();
            focusOperationalNote(Math.max(0, i - 1));
          }
        ),
        () => {
          if (state.operationalNotes.length > 1) {
            state.operationalNotes.splice(index, 1);
            saveDraft();
            renderOperationalNotes();
          }
        }
      );
      editable.input.classList.add("opnote-input");
      editable.input.dataset.index = index;
      editable.row.classList.add("bullet-line-item");
      const bullet = document.createElement("span");
      bullet.className = "opact-bullet";
      bullet.setAttribute("aria-hidden", "true");
      bullet.textContent = "\u2022";
      bullet.style.marginLeft = "24px";
      editable.row.insertBefore(bullet, editable.row.firstChild);
      wrap.appendChild(editable.row);
    });
  }

  function focusOperationalNote(index) {
    setTimeout(() => {
      const input = document.querySelector(`.opnote-input[data-index="${index}"]`);
      if (input) input.focus();
    }, 0);
  }

  function applyFlightToOffloadRow(rowIndex, flightData) {
    if (!flightData || !state.offloads[rowIndex]) return;
    state.offloads[rowIndex].flight = (flightData.code || "").toUpperCase();
    state.offloads[rowIndex].destination = (flightData.destination || "").toUpperCase();
    state.offloads[rowIndex].std = (flightData.stdEtd || "").toUpperCase();
    if (!state.offloads[rowIndex].date) {
      const raw = (flightData.date || "").toUpperCase();
      state.offloads[rowIndex].date = window.offloadLoader
        ? window.offloadLoader.normalizeOffloadDate(raw)
        : raw;
    }
    saveDraft();
    renderOffloads();
  }

  function reportIsoForOffloads() {
    return (state.activeDate || state.shiftMeta.date || "").trim();
  }

  function compactOffloadDate(value) {
    const raw = String(value || "").trim().toUpperCase();
    if (!raw) return "";
    if (/^\d{1,2}[A-Z]{3}$/.test(raw)) return raw;

    const normalized =
      window.offloadLoader && typeof window.offloadLoader.normalizeOffloadDate === "function"
        ? String(window.offloadLoader.normalizeOffloadDate(raw) || "").trim().toUpperCase()
        : raw;

    const m = normalized.match(/^(\d{1,2})-([A-Z]{3})(?:-(\d{4}))?$/);
    if (m) return `${parseInt(m[1], 10)}${m[2]}`;
    return normalized.replace(/[\s\-\/.]/g, "");
  }

  function looksLikeFlightCode(s) {
    return /^[A-Z]{2}\d{1,4}$/.test(String(s || "").trim().toUpperCase());
  }

  /**
   * After typing a flight number (without picking from the list), fill STD/ETD and DEST from
   * flights.json when possible, else from learned hints (localStorage + optional flight-hints.json).
   */
  function tryAutofillOffloadFlight(rowIndex) {
    const row = state.offloads[rowIndex];
    if (!row || !looksLikeFlightCode(row.flight)) return;

    const code = row.flight.trim().toUpperCase();
    const iso = reportIsoForOffloads();

    let fromCatalog = null;
    if (window.offloadLoader && typeof window.offloadLoader.findFlightForReport === "function") {
      fromCatalog = window.offloadLoader.findFlightForReport(code, iso);
    } else if (window.flightAutocomplete) {
      fromCatalog = window.flightAutocomplete.findByCode(code);
    }

    if (fromCatalog) {
      row.flight = code;
      row.std = (fromCatalog.stdEtd || "").toUpperCase().trim();
      row.destination = (fromCatalog.destination || "").toUpperCase().trim();
      saveDraft();
      renderOffloads();
      return;
    }

    if (!window.flightHintCache || !iso) return;
    const hint = window.flightHintCache.get(iso, code);
    if (!hint) return;

    let changed = false;
    if (!row.std && hint.std) {
      row.std = String(hint.std).toUpperCase().trim();
      changed = true;
    }
    if (!row.destination && hint.destination) {
      row.destination = String(hint.destination).toUpperCase().trim();
      changed = true;
    }
    if (changed) {
      saveDraft();
      renderOffloads();
    }
  }

  function persistOffloadFlightHints() {
    if (!window.flightHintCache) return;
    const iso = reportIsoForOffloads();
    if (!iso) return;
    const merge = {};
    state.offloads.forEach((row) => {
      const code = (row.flight || "").trim().toUpperCase();
      if (!looksLikeFlightCode(code)) return;
      const std = (row.std || "").trim();
      const dest = (row.destination || "").trim();
      if (!std || !dest) return;
      const k = window.flightHintCache.cacheKey(iso, code);
      if (k) merge[k] = { std, destination: dest };
    });
    if (Object.keys(merge).length) {
      window.flightHintCache.pushMerge(merge).catch(() => {});
    }
  }

  function renderOffloads() {
    const body = el("offloadTableBody");
    body.innerHTML = "";

    if (!state.offloads.length) {
      const tr = document.createElement("tr");
      tr.innerHTML =
        '<td colspan="12" style="text-align:center;color:#64748b">NIL — No offload data recorded for this shift.</td>';
      body.appendChild(tr);
      return;
    }

    state.offloads.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");

      const itemTd = document.createElement("td");
      itemTd.className = "offload-item-col";
      itemTd.textContent = row.item;
      itemTd.style.textAlign = "center";
      tr.appendChild(itemTd);

      offloadFieldOrder.forEach((field) => {
        const td = document.createElement("td");
        const isCompactSingleLine =
          field === "date" || field === "flight" || field === "std" || field === "destination";
        const cellInput = isCompactSingleLine ? document.createElement("input") : document.createElement("textarea");
        if (isCompactSingleLine) {
          cellInput.type = "text";
          cellInput.className = "offload-cell-compact";
        } else {
          cellInput.className = "offload-cell";
          cellInput.rows = 2;
        }
        cellInput.value = field === "date" ? compactOffloadDate(row[field] || "") : row[field] || "";
        cellInput.dataset.row = String(rowIndex);
        cellInput.dataset.field = field;
        cellInput.oninput = (e) => {
          let value = field === "flight" ? e.target.value.toUpperCase() : e.target.value;
          if (field === "date") value = compactOffloadDate(value);
          e.target.value = value;
          state.offloads[rowIndex][field] = value;
          saveDraft();
        };
        cellInput.onkeydown = (e) => handleOffloadKeydown(e, rowIndex, field);

        if (field === "remarks") {
          const wrap = document.createElement("div");
          wrap.className = "remarks-cell-wrap";
          const del = document.createElement("button");
          del.type = "button";
          del.className = "delete-btn hidden-print";
          del.textContent = "✕";
          del.onclick = () => {
            if (state.offloads.length > 1) {
              state.offloads.splice(rowIndex, 1);
              resequenceOffloads();
              saveDraft();
              renderOffloads();
            }
          };
          wrap.appendChild(cellInput);
          wrap.appendChild(del);
          td.appendChild(wrap);
          if (window.phraseAutocomplete) {
            window.phraseAutocomplete.attachTextarea(cellInput, "offloadRemarks", (value) => {
              state.offloads[rowIndex][field] = value;
              saveDraft();
            });
          }
        } else {
          if (field === "flight" && window.flightAutocomplete) {
            window.flightAutocomplete.attach(cellInput, `offload-${rowIndex}`, (picked) =>
              applyFlightToOffloadRow(rowIndex, picked)
            );
            cellInput.addEventListener("blur", () => {
              setTimeout(() => {
                if (
                  window.flightAutocomplete &&
                  window.flightAutocomplete.activeInput === cellInput &&
                  window.flightAutocomplete.activeMatches &&
                  window.flightAutocomplete.activeMatches.length
                ) {
                  return;
                }
                tryAutofillOffloadFlight(rowIndex);
              }, 220);
            });
          }
          if (field === "reason" && window.phraseAutocomplete) {
            window.phraseAutocomplete.attachTextarea(cellInput, "offloadReason", (value) => {
              state.offloads[rowIndex][field] = value;
              saveDraft();
            });
          }
          td.appendChild(cellInput);
        }

        tr.appendChild(td);
      });

      body.appendChild(tr);
    });
  }

  function handleOffloadKeydown(e, rowIndex, field) {
    const fieldIndex = offloadFieldOrder.indexOf(field);
    const isLastField = fieldIndex === offloadFieldOrder.length - 1;
    const isLastRow = rowIndex === state.offloads.length - 1;

    if (e.key === "Enter") {
      if (isAnySuggestOpenFor(e.target)) return;
      /* Enter inserts a new line inside the cell (textarea); flight autocomplete handles Enter when the list is open. */
      return;
    }

    if (e.key === "Backspace" && !e.target.value && state.offloads.length > 1) {
      e.preventDefault();
      const prevField = fieldIndex > 0 ? offloadFieldOrder[fieldIndex - 1] : offloadFieldOrder[offloadFieldOrder.length - 1];
      const prevRow = fieldIndex > 0 ? rowIndex : Math.max(0, rowIndex - 1);
      state.offloads.splice(rowIndex, 1);
      resequenceOffloads();
      saveDraft();
      renderOffloads();
      focusOffload(prevRow, prevField);
      return;
    }

    if (e.key === "Tab" && !e.shiftKey) {
      e.preventDefault();
      if (isLastField && isLastRow) {
        insertOffloadRow(state.offloads.length);
        focusOffload(state.offloads.length - 1, "date");
        return;
      }
      const nextField = offloadFieldOrder[(fieldIndex + 1) % offloadFieldOrder.length];
      const nextRow = isLastField ? Math.min(rowIndex + 1, state.offloads.length - 1) : rowIndex;
      focusOffload(nextRow, nextField);
    }
  }

  function insertOffloadRow(index) {
    const defaultDate = window.offloadLoader
      ? window.offloadLoader.normalizeOffloadDate(state.shiftMeta.date || "")
      : state.shiftMeta.date || "";
    const blank = {
      item: index + 1,
      date: defaultDate,
      flight: "",
      std: "",
      destination: "",
      emailTime: "",
      rampReceived: "",
      trolley: "",
      cmsCompleted: "",
      piecesVerification: "",
      reason: "",
      remarks: "",
    };
    state.offloads.splice(index, 0, blank);
    resequenceOffloads();
    saveDraft();
    renderOffloads();
  }

  function resequenceOffloads() {
    state.offloads = state.offloads.map((r, i) => ({ ...r, item: i + 1 }));
  }

  function focusOffload(rowIndex, field) {
    setTimeout(() => {
      const input = document.querySelector(`[data-row="${rowIndex}"][data-field="${field}"]`);
      if (input) input.focus();
    }, 0);
  }

  function ensureManpowerRowForEditing(section) {
    if (!section.items || section.items.length === 0) section.items = [""];
  }

  function splitManpowerNameRole(line) {
    const v = String(line || "").trim();
    if (!v) return { name: "", role: "" };
    const m = v.match(/^(.+?)\s*-\s*(.+)$/);
    if (!m) return { name: v, role: "" };
    return { name: String(m[1] || "").trim(), role: String(m[2] || "").trim() };
  }

  function normalizeManpowerEmployeeName(line) {
    const parsed = splitManpowerNameRole(line);
    return String(parsed.name || "")
      .trim()
      .replace(/\s+/g, " ")
      .toUpperCase();
  }

  function dedupeManpowerEmployeeNames() {
    const seen = new Set();
    let changed = false;
    (state.manpowerSections || []).forEach((section) => {
      if (!section || !Array.isArray(section.items)) return;
      section.items = section.items.map((raw) => {
        const txt = String(raw || "");
        const key = normalizeManpowerEmployeeName(txt);
        if (!key) return txt;
        if (seen.has(key)) {
          changed = true;
          return "";
        }
        seen.add(key);
        return txt;
      });
      ensureManpowerRowForEditing(section);
    });
    return changed;
  }

  function sectionAllowsAutoRoleHint(sectionTitle) {
    const t = String(sectionTitle || "").trim();
    return /^(export checker|export operators|flight dispatch)$/i.test(t);
  }

  function isExportCheckerSection(sectionTitle) {
    return String(sectionTitle || "").trim().toLowerCase() === "export checker";
  }

  function sortExportCheckerSectionByRole(sectionIndex) {
    const section = state.manpowerSections[sectionIndex];
    if (!section || !Array.isArray(section.items) || !isExportCheckerSection(section.title)) return false;

    const rows = section.items.map((raw, idx) => {
      const text = String(raw || "").trim();
      const parsed = splitManpowerNameRole(text);
      return {
        idx,
        raw: raw == null ? "" : String(raw),
        text,
        name: parsed.name,
        role: parsed.role
      };
    });
    const filled = rows.filter((r) => r.text);
    if (!filled.length) return false;
    const roleCounts = new Map();
    filled.forEach((r) => {
      const key = String(r.role || "").trim().toLowerCase();
      if (!key) return;
      roleCounts.set(key, (roleCounts.get(key) || 0) + 1);
    });

    const sortedFilled = filled.slice().sort((a, b) => {
      const aRole = String(a.role || "").trim();
      const bRole = String(b.role || "").trim();
      const aHasRole = Boolean(aRole);
      const bHasRole = Boolean(bRole);
      if (aHasRole !== bHasRole) return aHasRole ? -1 : 1;
      if (aHasRole && bHasRole) {
        const aCount = roleCounts.get(aRole.toLowerCase()) || 0;
        const bCount = roleCounts.get(bRole.toLowerCase()) || 0;
        if (aCount !== bCount) return bCount - aCount;
        const roleCmp = aRole.localeCompare(bRole, undefined, { sensitivity: "base" });
        if (roleCmp !== 0) return roleCmp;
      }
      const nameCmp = String(a.name || a.text).localeCompare(String(b.name || b.text), undefined, { sensitivity: "base" });
      if (nameCmp !== 0) return nameCmp;
      return a.idx - b.idx;
    });

    const emptyCount = rows.length - filled.length;
    const nextItems = sortedFilled.map((r) => r.raw);
    const padCount = Math.max(1, emptyCount);
    for (let i = 0; i < padCount; i += 1) nextItems.push("");

    const changed =
      nextItems.length !== section.items.length ||
      nextItems.some((v, i) => v !== section.items[i]);
    if (!changed) return false;
    section.items = nextItems;
    return true;
  }

  function syncManpowerSectionFromDom(sectionIndex) {
    const nodes = Array.from(document.querySelectorAll(`input.manpower-line[data-section="${sectionIndex}"]`));
    if (!nodes.length || !state.manpowerSections[sectionIndex]) return;
    state.manpowerSections[sectionIndex].items = nodes.map((n) => String((n && n.value) || ""));
  }

  /**
   * If a name is frequently paired with a role, auto-fill "NAME - ROLE"
   * in selected sections so supervisors see the likely role on open.
   */
  function applyLearnedRolesToManpowerSections() {
    if (!window.manpowerRoleHintCache || typeof window.manpowerRoleHintCache.getTopLearnedRoleForName !== "function") return false;
    let changed = false;
    (state.manpowerSections || []).forEach((section) => {
      if (!section || !Array.isArray(section.items) || !sectionAllowsAutoRoleHint(section.title)) return;
      section.items = section.items.map((raw) => {
        const { name, role } = splitManpowerNameRole(raw);
        if (!name || role) return raw;
        const topRole = window.manpowerRoleHintCache.getTopLearnedRoleForName(name, 4);
        if (!topRole) return raw;
        const next = `${name} - ${topRole}`;
        if (next !== raw) changed = true;
        return next;
      });
    });
    return changed;
  }

  function renderManpower() {
    const wrap = el("manpowerWrap");
    wrap.innerHTML = "";
    if (dedupeManpowerEmployeeNames()) {
      saveDraft();
    }
    if (applyLearnedRolesToManpowerSections()) {
      saveDraft();
    }

    state.manpowerSections.forEach((section, sectionIndex) => {
      ensureManpowerRowForEditing(state.manpowerSections[sectionIndex]);

      const sectionBox = document.createElement("div");
      sectionBox.className = "manpower-section";

      const title = document.createElement("input");
      title.type = "text";
      title.value = section.title;
      title.className = "manpower-section-title";
      title.oninput = (e) => {
        state.manpowerSections[sectionIndex].title = e.target.value;
        saveDraft();
        renderWebSignaturePreview();
      };
      sectionBox.appendChild(title);

      const list = state.manpowerSections[sectionIndex].items;
      list.forEach((item, itemIndex) => {
        const row = document.createElement("div");
        row.className = "manpower-item";

        const bullet = document.createElement("span");
        bullet.className = "manpower-bullet";
        bullet.setAttribute("aria-hidden", "true");
        bullet.textContent = "\u2022";
        bullet.style.marginLeft = "24px";

        const input = document.createElement("input");
        input.type = "text";
        input.className = "manpower-line";
        input.placeholder = "Add name…";
        input.value = item;
        input.dataset.section = sectionIndex;
        input.dataset.index = itemIndex;
        input.oninput = (e) => {
          state.manpowerSections[sectionIndex].items[itemIndex] = e.target.value;
          saveDraft();
          renderWebSignaturePreview();
        };
        const maybeSortExportChecker = () => {
          syncManpowerSectionFromDom(sectionIndex);
          if (dedupeManpowerEmployeeNames()) {
            saveDraft();
            renderManpower();
            return;
          }
          if (!sortExportCheckerSectionByRole(sectionIndex)) return;
          saveDraft();
          renderManpower();
        };
        input.onchange = maybeSortExportChecker;
        input.onkeydown = (e) => handleManpowerKeydown(e, sectionIndex, itemIndex);
        if (window.employeeAutocomplete) {
          window.employeeAutocomplete.attach(input, `${sectionIndex}-${itemIndex}`, {
            sectionTitle: section.title || ""
          });
        }

        const del = document.createElement("button");
        del.type = "button";
        del.className = "delete-btn hidden-print";
        del.textContent = "✕";
        del.onclick = () => {
          const cur = state.manpowerSections[sectionIndex].items;
          if (cur.length > 1) {
            cur.splice(itemIndex, 1);
            ensureManpowerRowForEditing(state.manpowerSections[sectionIndex]);
            saveDraft();
            renderManpower();
          } else {
            cur[0] = "";
            saveDraft();
            renderManpower();
          }
        };

        row.appendChild(bullet);
        row.appendChild(input);
        row.appendChild(del);
        sectionBox.appendChild(row);
      });

      wrap.appendChild(sectionBox);
    });
    renderWebSignaturePreview();
  }

  function handleManpowerKeydown(e, sectionIndex, itemIndex) {
    const list = state.manpowerSections[sectionIndex].items;
    if (e.key === "Enter") {
      e.preventDefault();
      const inputEl = e.target;
      const typedValue = String((inputEl && inputEl.value) || "");
      // Prevent browser datalist from auto-committing first suggestion on blur.
      if (inputEl && typeof inputEl.removeAttribute === "function" && inputEl.hasAttribute("list")) {
        inputEl.removeAttribute("list");
      }
      list[itemIndex] = typedValue;
      list.splice(itemIndex + 1, 0, "");
      saveDraft();
      renderManpower();
      focusManpower(sectionIndex, itemIndex + 1);
      return;
    }
    if (e.key === "Backspace" && !e.target.value) {
      if (list.length > 1) {
        e.preventDefault();
        list.splice(itemIndex, 1);
        ensureManpowerRowForEditing(state.manpowerSections[sectionIndex]);
        saveDraft();
        renderManpower();
        focusManpower(sectionIndex, Math.max(0, itemIndex - 1));
      }
    }
  }

  function focusManpower(sectionIndex, itemIndex) {
    setTimeout(() => {
      const input = document.querySelector(`input[data-section="${sectionIndex}"][data-index="${itemIndex}"]`);
      if (input) input.focus();
    }, 0);
  }

  function wireSpecialHoTextarea() {
    const area = el("specialHO");
    if (!area) return;
    if (area.dataset.wiredSpecialHo === "1") return;
    area.dataset.wiredSpecialHo = "1";

    const syncStateFromSpecialHo = () => {
      state.specialHO = normalizeIndentedBullets(toSentenceCaseText(area.value));
      saveDraft();
    };

    area.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const start = area.selectionStart;
      const end = area.selectionEnd;
      const value = String(area.value || "");
      const insert = "\n    \u2022 ";
      area.value = value.slice(0, start) + insert + value.slice(end);
      const pos = start + insert.length;
      area.selectionStart = pos;
      area.selectionEnd = pos;
      syncStateFromSpecialHo();
    });

    area.addEventListener("input", syncStateFromSpecialHo);
    area.addEventListener("blur", () => {
      const normalized = normalizeIndentedBullets(toSentenceCaseText(area.value));
      if (area.value !== normalized) area.value = normalized;
      syncStateFromSpecialHo();
    });
  }

  function normalizeRecipientsShape(payload) {
    const src = Array.isArray(payload) ? { to: payload, cc: [], bcc: [] } : payload && typeof payload === "object" ? payload : {};
    const uniq = (arr) => {
      const seen = new Set();
      const out = [];
      (Array.isArray(arr) ? arr : []).forEach((x) => {
        const v = String(x || "").trim().toLowerCase();
        if (!v || seen.has(v)) return;
        seen.add(v);
        out.push(v);
      });
      return out;
    };
    return {
      to: uniq(src.to),
      cc: uniq(src.cc),
      bcc: uniq(src.bcc),
    };
  }

  function recipientsAnyCount() {
    const r = normalizeRecipientsShape(state.recipients);
    return r.to.length + r.cc.length + r.bcc.length;
  }

  function persistRecipientsToServer() {
    if (window.recipientsCache && typeof window.recipientsCache.replaceAll === "function") {
      window.recipientsCache.replaceAll(state.recipients).catch((e) => {
        console.warn("Recipients sync failed", e);
      });
    }
  }

  function renderRecipients() {
    const map = [
      ["to", "recipientTagsTo"],
      ["cc", "recipientTagsCc"],
      ["bcc", "recipientTagsBcc"],
    ];
    const recipients = normalizeRecipientsShape(state.recipients);
    state.recipients = recipients;
    map.forEach(([kind, wrapId]) => {
      const wrap = el(wrapId);
      if (!wrap) return;
      wrap.innerHTML = "";
      recipients[kind].forEach((emailAddr) => {
        const tag = document.createElement("div");
        tag.className = "recipient-tag";
        tag.textContent = `${emailAddr} ×`;
        tag.onclick = () => {
          state.recipients[kind] = state.recipients[kind].filter((x) => x !== emailAddr);
          saveDraft();
          renderRecipients();
          scheduleSendTimerFromState();
          persistRecipientsToServer();
        };
        wrap.appendChild(tag);
      });
    });
  }

  function parseRecipientInput(raw) {
    return String(raw || "")
      .split(/[,\n;]+/)
      .map((x) => x.trim())
      .filter((x) => x);
  }

  function isValidEmailBasic(v) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || "").trim());
  }

  function recipientsMailtoValue(kind) {
    const r = normalizeRecipientsShape(state.recipients);
    return r[kind].join(",");
  }

  function openEmailComposer(options = {}) {
    const { includePlainBody = true, bodyOverride = "" } = options || {};
    const to = recipientsMailtoValue("to");
    const cc = recipientsMailtoValue("cc");
    const bcc = recipientsMailtoValue("bcc");
    if (!to && !cc && !bcc) return false;
    const subject = "Export Warehouse Activity Report";
    const parts = [`subject=${encodeURIComponent(subject)}`];
    if (includePlainBody) {
      const bodyText = String(bodyOverride || "").trim() || buildReportPlainBody();
      const body = encodeURIComponent(bodyText);
      parts.push(`body=${body}`);
    }
    if (cc) parts.push(`cc=${encodeURIComponent(cc)}`);
    if (bcc) parts.push(`bcc=${encodeURIComponent(bcc)}`);
    window.location.href = `mailto:${to}?${parts.join("&")}`;
    return true;
  }

  /**
   * mailto cannot carry full HTML reliably in Outlook desktop.
   * We pre-copy the rich report so the user can paste it in compose.
   */
  async function openEmailComposerWithClipboardHint() {
    const statusEl = el("gmailStatus");
    let copied = false;
    try {
      await copyReportToClipboard();
      copied = true;
    } catch (err) {
      console.warn("Could not pre-copy rich report before mailto fallback", err);
    }
    const opened = copied
      ? openEmailComposer({
          includePlainBody: true,
          bodyOverride:
            "\u200EFormatted report is copied to clipboard.\n\u200EPaste inside the email body with Ctrl+V."
        })
      : openEmailComposer({ includePlainBody: true });
    if (statusEl && opened) {
      statusEl.textContent = copied
        ? "Mail draft opened with text body. The formatted report is copied — paste with Ctrl+V in Outlook body."
        : "Mail draft opened. If the body appears plain in Outlook, use Copy then paste with Ctrl+V.";
    }
    return opened;
  }

  async function refreshGmailStatus() {
    const statusEl = el("gmailStatus");
    if (!canUseServerGmailApi()) {
      _gmailStatus = { configured: false, authorized: false };
      if (statusEl) {
        statusEl.textContent = "Gmail API: unavailable on static host (will use mailto fallback).";
      }
      return;
    }
    try {
      const r = await fetch(apiUrl("/api/gmail/status"), { cache: "no-store" });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const st = await r.json();
      _gmailStatus = {
        configured: !!st.configured,
        authorized: !!st.authorized,
      };
    } catch {
      _gmailStatus = { configured: false, authorized: false };
    }
    if (statusEl) {
      if (_gmailStatus.configured && _gmailStatus.authorized) {
        statusEl.textContent = "Gmail API: connected.";
      } else if (_gmailStatus.configured && !_gmailStatus.authorized) {
        statusEl.textContent = "Gmail API: configured but not authorized (click Connect Gmail).";
      } else {
        statusEl.textContent = "Gmail API: not configured on server (will use mailto fallback).";
      }
    }
  }

  async function connectGmailFlow() {
    const statusEl = el("gmailStatus");
    if (!canUseServerGmailApi()) {
      if (statusEl) statusEl.textContent = "Gmail connect is not available on static host. Using mailto fallback.";
      return;
    }
    try {
      const r = await fetch(apiUrl("/api/gmail/auth-url"), { cache: "no-store" });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const out = await r.json();
      if (!out.url) throw new Error("Missing auth URL");
      window.open(out.url, "_blank", "noopener");
      if (statusEl) statusEl.textContent = "Gmail auth tab opened. Complete login, then come back here.";
      setTimeout(() => refreshGmailStatus(), 2000);
    } catch (e) {
      if (statusEl) statusEl.textContent = `Gmail connect failed: ${e && e.message ? e.message : "unknown error"}`;
    }
  }

  async function sendEmailNow() {
    const recipients = normalizeRecipientsShape(state.recipients);
    if (!recipients.to.length && !recipients.cc.length && !recipients.bcc.length) return false;

    if (!canUseServerGmailApi()) {
      const statusEl = el("gmailStatus");
      if (statusEl) {
        statusEl.textContent = "Direct send is unavailable. Start the local API server and connect Gmail.";
      }
      return false;
    }

    const subject = "Export Warehouse Activity Report";
    const plain = buildReportPlainBody();
    const html = await buildReportHtmlForClipboard();

    try {
      const r = await fetch(apiUrl("/api/gmail/send"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          to: recipients.to,
          cc: recipients.cc,
          bcc: recipients.bcc,
          subject,
          plain,
          html,
        }),
      });
      if (r.ok) {
        await refreshGmailStatus();
        return true;
      }
    } catch (_) {
      /* fallback below */
    }
    const statusEl = el("gmailStatus");
    if (statusEl) {
      statusEl.textContent = "Direct send failed. Check Gmail connection and API server.";
    }
    return false;
  }

  function setEmailButtonState(mode) {
    const btn = el("emailBtn");
    if (!btn) return;
    if (mode === "sending") {
      btn.disabled = true;
      btn.classList.remove("send-success");
      btn.textContent = "Sending...";
      return;
    }
    btn.disabled = false;
    if (mode === "success") {
      btn.classList.add("send-success");
      btn.textContent = "Sent via Gmail";
      return;
    }
    btn.classList.remove("send-success");
    btn.textContent = EMAIL_BUTTON_DEFAULT_TEXT;
  }

  function clearScheduledSendTimer() {
    if (_scheduledSendTimer) {
      clearTimeout(_scheduledSendTimer);
      _scheduledSendTimer = null;
    }
  }

  function parseScheduleDate(value) {
    const v = String(value || "").trim();
    if (!v) return null;
    const d = new Date(v);
    if (Number.isNaN(d.getTime())) return null;
    return d;
  }

  function renderScheduledSendStatus() {
    const statusEl = el("scheduleStatus");
    const inputEl = el("autoSendAt");
    if (!statusEl || !inputEl) return;
    inputEl.value = state.scheduledSendAt || "";
    if (!state.scheduledSendEnabled || !state.scheduledSendAt) {
      statusEl.textContent = "No scheduled send.";
      return;
    }
    const d = parseScheduleDate(state.scheduledSendAt);
    if (!d) {
      statusEl.textContent = "Scheduled time is invalid.";
      return;
    }
    statusEl.textContent = `Scheduled for ${d.toLocaleString()} (page must stay open).`;
  }

  async function runScheduledSendIfDue() {
    if (!state.scheduledSendEnabled || !state.scheduledSendAt) return;
    const due = parseScheduleDate(state.scheduledSendAt);
    if (!due) return;
    if (Date.now() < due.getTime()) return;
    if (state._scheduledSendLastFiredAt === state.scheduledSendAt) return;
    const ok = await sendEmailNow();
    state._scheduledSendLastFiredAt = state.scheduledSendAt;
    state.scheduledSendEnabled = false;
    renderScheduledSendStatus();
    saveDraft();
    if (!ok) {
      const statusEl = el("scheduleStatus");
      if (statusEl) statusEl.textContent = "Scheduled time reached, but recipients list is empty.";
    }
  }

  function scheduleSendTimerFromState() {
    clearScheduledSendTimer();
    renderScheduledSendStatus();
    if (!state.scheduledSendEnabled || !state.scheduledSendAt) return;
    const due = parseScheduleDate(state.scheduledSendAt);
    if (!due) return;
    const delay = due.getTime() - Date.now();
    if (delay <= 0) {
      runScheduledSendIfDue().catch(console.error);
      return;
    }
    _scheduledSendTimer = setTimeout(() => {
      _scheduledSendTimer = null;
      runScheduledSendIfDue().catch(console.error);
    }, delay);
  }

  /** Keep offload headers readable in clipboard HTML (avoid per-letter wrapping). */
  function normalizeOffloadTableForClipboard(root) {
    const table = root.querySelector("table.offload-table");
    if (!table) return;

    // Outlook-safe fixed column plan (prevents header letter breaking).
    const colPx = [28, 80, 88, 70, 52, 70, 96, 120, 96, 86, 76, 138];

    const existingTableStyle = table.getAttribute("style") || "";
    const stripped = existingTableStyle.replace(/\b(min-)?width\s*:\s*[^;]+;?/gi, "");
    table.setAttribute("width", "1000");
    table.setAttribute(
      "style",
      "width:1000px;max-width:1000px;min-width:1000px;table-layout:fixed;border-collapse:collapse;box-sizing:border-box;margin:0;" +
        stripped
    );

    const headerCells = table.querySelectorAll("thead tr th");
    headerCells.forEach((th, i) => {
      const px = colPx[i] || 110;
      let prev = th.getAttribute("style") || "";
      prev = prev.replace(/\bwidth\s*:\s*[^;]+;?/gi, "");
      th.setAttribute(
        "style",
        `width:${px}px;max-width:${px}px;min-width:${px}px;background:rgba(238,243,252,0.85);border:1px solid #d0d5e8;white-space:normal;word-wrap:normal;overflow-wrap:normal;word-break:keep-all;hyphens:none;line-height:1.35;font-size:9.2pt;mso-ansi-font-size:9.2pt;padding:10px 8px;text-align:center;vertical-align:middle;box-sizing:border-box;${prev}`
      );
    });

    table.querySelectorAll("tbody tr").forEach((tr) => {
      let colIndex = 0;
      tr.querySelectorAll("td").forEach((td) => {
        const span = parseInt(td.getAttribute("colspan") || "1", 10);
        let prev = td.getAttribute("style") || "";
        prev = prev.replace(/\bwidth\s*:\s*[^;]+;?/gi, "");
        if (span > 1) {
          td.setAttribute(
            "style",
            "width:100%;max-width:100%;min-width:0;word-wrap:normal;overflow-wrap:normal;word-break:normal;box-sizing:border-box;" + prev
          );
          return;
        }
        const px = colPx[colIndex] || 110;
        colIndex += 1;
        td.setAttribute(
          "style",
          `width:${px}px;max-width:${px}px;min-width:${px}px;border:1px solid #d0d5e8;word-wrap:normal;overflow-wrap:normal;word-break:normal;font-size:11.5pt;mso-ansi-font-size:11.5pt;padding:12px;box-sizing:border-box;line-height:1.9;${prev}`
        );
        // DATE column (2nd visible column, index 1 after item): smaller text for cleaner fit.
        if (colIndex - 1 === 1) {
          const s = td.getAttribute("style") || "";
          td.setAttribute(
            "style",
            `font-size:9.8pt;mso-ansi-font-size:9.8pt;line-height:1.5;word-break:normal;overflow-wrap:normal;` + s
          );
        }
      });
    });
  }

  /** Outlook-friendly alternative: convert wide offload table into stacked record cards. */
  function convertOffloadTableToRecordCardsForClipboard(root) {
    const table = root.querySelector("table.offload-table");
    if (!table) return false;
    const bodyRows = Array.from(table.querySelectorAll("tbody tr"));
    if (!bodyRows.length) return false;

    const labels = [
      "ITEM",
      "DATE",
      "FLIGHT",
      "STD/ETD",
      "DEST",
      "Email Received Time",
      "Physical Cargo Received from Ramp",
      "Trolley / ULD Number",
      "Offloading Process Completed in CMS",
      "Offloading Pieces Verification",
      "Offloading Reason",
      "Remarks / Additional Information",
    ];

    const cardsWrap = document.createElement("div");
    cardsWrap.className = "offload-cards-wrap";
    cardsWrap.setAttribute(
      "style",
      "width:100%;max-width:100%;box-sizing:border-box;margin-top:8px;"
    );

    bodyRows.forEach((tr) => {
      const cells = Array.from(tr.querySelectorAll("td"));
      if (!cells.length) return;

      const values = cells.map((td) => String(td.textContent || "").replace(/\s+/g, " ").trim());
      const card = document.createElement("table");
      card.setAttribute("role", "presentation");
      card.setAttribute("cellpadding", "0");
      card.setAttribute("cellspacing", "0");
      card.setAttribute("border", "0");
      card.setAttribute("width", "100%");
      card.setAttribute(
        "style",
        "border-collapse:collapse;width:100%;table-layout:fixed;margin:0 0 10px 0;border:1px solid #cbd5e1;border-radius:8px;background:#f8fafc;"
      );

      for (let i = 0; i < Math.min(labels.length, values.length); i += 1) {
        const row = document.createElement("tr");
        const tdLabel = document.createElement("td");
        tdLabel.setAttribute(
          "style",
          "width:230px;min-width:230px;max-width:230px;vertical-align:top;padding:6px 8px;border-bottom:1px solid #e2e8f0;font-family:Arial,sans-serif;font-size:10.0pt;mso-ansi-font-size:10.0pt;font-weight:700;color:#1e3a8a;background:#eff6ff;"
        );
        tdLabel.textContent = `${labels[i]}:`;

        const tdVal = document.createElement("td");
        tdVal.setAttribute(
          "style",
          "vertical-align:top;padding:6px 8px;border-bottom:1px solid #e2e8f0;font-family:Arial,sans-serif;font-size:10.5pt;mso-ansi-font-size:10.5pt;color:#000000;background:#ffffff;"
        );
        tdVal.textContent = values[i] || "";

        row.appendChild(tdLabel);
        row.appendChild(tdVal);
        card.appendChild(row);
      }

      cardsWrap.appendChild(card);
    });

    const offloadWrap = root.querySelector(".offload-wrap");
    if (offloadWrap) {
      offloadWrap.innerHTML = "";
      offloadWrap.appendChild(cardsWrap);
    } else {
      table.replaceWith(cardsWrap);
    }
    return true;
  }

  /** Word respects pt + mso-* on spans more than px on table cells. */
  function wordStyledSpan(text, css) {
    const span = document.createElement("span");
    span.setAttribute("lang", "EN-US");
    span.setAttribute("style", css);
    span.appendChild(document.createTextNode(text));
    return span;
  }

  const WORD_CLIPBOARD = {
    body:
      "mso-ansi-font-size:12.0pt;mso-bidi-font-size:12.0pt;font-size:12.0pt;font-family:Calibri,Arial,sans-serif;color:#000000;line-height:190%;mso-line-height-rule:at-least;",
    bullet:
      "mso-ansi-font-size:12.0pt;font-size:12.0pt;font-family:Calibri,Arial,sans-serif;mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;font-weight:bold;color:#000000;",
    sectionTitle:
      "mso-ansi-font-size:14.0pt;mso-bidi-font-size:14.0pt;font-size:14.0pt;font-family:Calibri,Arial,sans-serif;mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;font-weight:bold;letter-spacing:0.02em;color:#000000;",
    bannerMajor:
      "mso-ansi-font-size:14.0pt;mso-bidi-font-size:14.0pt;font-size:14.0pt;font-family:Calibri,Arial,sans-serif;mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;font-weight:bold;letter-spacing:0.02em;color:#000000;line-height:1.9;",
    bannerDefault:
      "mso-ansi-font-size:12.0pt;mso-bidi-font-size:12.0pt;font-size:12.0pt;font-family:Calibri,Arial,sans-serif;font-weight:bold;color:#000000;line-height:1.9;",
  };

  /** Keep section titles as simple divs with inline styles matching web exactly. */
  function convertSectionTitlesToOutlookShadeTables(root) {
    root.querySelectorAll(".section-title").forEach((el) => {
      // Don't replace with table — just apply inline styles to match web CSS exactly.
      const isMajor = el.classList && el.classList.contains("manpower-block-heading");
      const fontSize = isMajor ? "14px" : "14px";
      const inlineStyle =
        "display:block;width:100%;box-sizing:border-box;" +
        "margin-top:14px;margin-bottom:0;" +
        "background:#eef3fc !important;" +
        "padding:8px 12px !important;" +
        "font-weight:700;" +
        "border:1px solid #e4e9f5 !important;" +
        "border-top:1px solid #e4e9f5 !important;" +
        "border-left:4px solid #0b3a78 !important;" +
        "border-radius:0 !important;" +
        "line-height:1.3;" +
        "color:#111827 !important;" +
        "text-transform:uppercase;" +
        "letter-spacing:0.06em;" +
        "font-size:" + fontSize + ";" +
        "font-family:Calibri,Arial,sans-serif;" +
        "mso-line-height-rule:exactly;";
      el.setAttribute("style", inlineStyle);
    });
  }

  /**
   * Word/Outlook paste: flex + display:block on name spans stacks bullet above text.
   * Replace each manpower row with a 2-column table (bullet | text) — classic Word-friendly layout.
   */
  function restructureManpowerItemsForWordPaste(root) {
    root.querySelectorAll(".manpower-item").forEach((row) => {
      const bulletEl = row.querySelector(".manpower-bullet");
      const valEl = row.querySelector(".export-val.manpower-line") || row.querySelector(".manpower-line");
      const bulletChar =
        bulletEl && String(bulletEl.textContent || "").trim() ? String(bulletEl.textContent).trim() : "\u2022";
      const text = valEl ? String(valEl.textContent || "") : "";
      const table = document.createElement("table");
      table.setAttribute("cellpadding", "0");
      table.setAttribute("cellspacing", "0");
      table.setAttribute("border", "0");
      table.setAttribute("width", "100%");
      table.className = "manpower-paste-row";
      table.setAttribute(
        "style",
        "border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;width:100%;margin:0 0 6px 0;"
      );
      const tr = document.createElement("tr");
      const tdBullet = document.createElement("td");
      tdBullet.setAttribute(
        "style",
        "vertical-align:top;width:18px;padding:2px 2px 0 14px;white-space:nowrap;mso-line-height-rule:exactly;border:none;"
      );
      tdBullet.appendChild(wordStyledSpan(`\u00A0\u00A0\u00A0\u00A0${bulletChar}`, WORD_CLIPBOARD.bullet));
      const tdText = document.createElement("td");
      tdText.setAttribute(
        "style",
        "vertical-align:top;padding:0 0 0 1px;width:99%;mso-line-height-rule:exactly;border:none;"
      );
      tdText.appendChild(wordStyledSpan(text, WORD_CLIPBOARD.body));
      tr.appendChild(tdBullet);
      tr.appendChild(tdText);
      table.appendChild(tr);
      row.replaceWith(table);
    });
  }

  /**
   * Outlook/Word may reflow flex/span rows unexpectedly.
   * Convert generic line rows to single-cell tables to preserve one-line output.
   */
  function restructureLineItemsForWordPaste(root) {
    root.querySelectorAll(".line-item").forEach((row) => {
      const textParts = Array.from(row.querySelectorAll(".export-val"))
        .map((el) => String(el.textContent || "").trim())
        .filter(Boolean);
      const needsBullet = row.classList.contains("opact-line-item") || row.classList.contains("bullet-line-item");
      if (!textParts.length && !needsBullet) return;
      const lineCore = textParts.join(" ");
      const text = needsBullet ? lineCore : lineCore;
      const table = document.createElement("table");
      table.setAttribute("cellpadding", "0");
      table.setAttribute("cellspacing", "0");
      table.setAttribute("border", "0");
      table.setAttribute("width", "100%");
      table.className = "line-item-paste-row";
      table.setAttribute(
        "style",
        "border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;width:100%;margin:0 0 6px 0;"
      );
      const tr = document.createElement("tr");
      if (needsBullet) {
        const tdIndent = document.createElement("td");
        tdIndent.setAttribute(
          "style",
          "vertical-align:top;width:22px;padding:2px 0 0 0;white-space:nowrap;mso-line-height-rule:exactly;border:none;"
        );
        tdIndent.appendChild(wordStyledSpan("\u00A0\u00A0\u00A0\u00A0\u2022", WORD_CLIPBOARD.bullet));

        const tdText = document.createElement("td");
        tdText.setAttribute(
          "style",
          "vertical-align:top;padding:2px 0 0 8px;mso-line-height-rule:exactly;border:none;"
        );
        tdText.appendChild(wordStyledSpan(text, WORD_CLIPBOARD.body));
        tr.appendChild(tdIndent);
        tr.appendChild(tdText);
      } else {
        const td = document.createElement("td");
        td.setAttribute("style", "vertical-align:top;padding:2px 0;mso-line-height-rule:exactly;");
        td.appendChild(wordStyledSpan(text, WORD_CLIPBOARD.body));
        tr.appendChild(td);
      }
      table.appendChild(tr);
      row.replaceWith(table);
    });
  }

  /** Outlook uses Word HTML; it often strips &lt;style&gt; — inline CSS + bgcolor where needed. */
  function applyOutlookInlineClipboardStyles(root) {
    const baseFont =
      "font-family:Calibri,Arial,sans-serif;font-size:12.0pt;mso-ansi-font-size:12.0pt;color:#000000;line-height:190%;";
    root.setAttribute(
      "style",
      "width:1000px;max-width:1000px;min-width:1000px;box-sizing:border-box;margin:0;padding:0;" +
        baseFont +
        (root.getAttribute("style") || "")
    );

    root.querySelectorAll(".block").forEach((el) => {
      el.setAttribute("style", "padding:12px 4px 0;" + (el.getAttribute("style") || ""));
    });
    root.querySelectorAll(".line-group").forEach((el) => {
      el.setAttribute("style", "margin-bottom:14px;" + (el.getAttribute("style") || ""));
    });
    root.querySelectorAll(".line-group-title").forEach((el) => {
      el.setAttribute(
        "style",
        "font-weight:700;color:#000000;margin-bottom:6px;mso-ansi-font-size:12.0pt;font-size:12.0pt;font-family:'Arial',sans-serif;" +
          (el.getAttribute("style") || "")
      );
    });
    // .section-title styles already applied by convertSectionTitlesToOutlookShadeTables — skip.
    root.querySelectorAll(".line-item").forEach((el) => {
      el.setAttribute(
        "style",
          "display:block;margin-bottom:6px;padding:2px 0;font-size:12.0pt;mso-ansi-font-size:12.0pt;font-family:'Arial',sans-serif;" +
          (el.getAttribute("style") || "")
      );
    });

    root.querySelectorAll(".manpower-wrap").forEach((el) => {
      el.setAttribute("style", "padding-top:10px;" + (el.getAttribute("style") || ""));
    });
    root.querySelectorAll(".manpower-section").forEach((el) => {
      el.setAttribute("style", "margin-bottom:14px;" + (el.getAttribute("style") || ""));
    });
    root.querySelectorAll(".manpower-section-title").forEach((el) => {
      el.setAttribute(
        "style",
          "display:block;width:100%;box-sizing:border-box;border:none;padding:6px 2px 8px 2px;margin-bottom:8px;" +
          WORD_CLIPBOARD.sectionTitle +
          (el.getAttribute("style") || "")
      );
    });
    root.querySelectorAll(".manpower-bullet").forEach((el) => {
      el.setAttribute(
        "style",
        "display:inline-block;min-width:1em;font-weight:700;font-size:15px;line-height:1.9;color:#000000;padding-top:2px;font-family:Calibri,Arial,sans-serif;" +
          (el.getAttribute("style") || "")
      );
    });
    root.querySelectorAll(".manpower-item").forEach((el) => {
      el.setAttribute(
        "style",
        "display:flex;flex-direction:row;align-items:flex-start;gap:10px;margin-bottom:8px;" + (el.getAttribute("style") || "")
      );
    });
    root.querySelectorAll(".manpower-item .export-val.manpower-line").forEach((el) => {
      el.setAttribute(
        "style",
        "flex:1;min-width:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.45;color:#000000;" +
          (el.getAttribute("style") || "")
      );
    });

    root.querySelectorAll(".export-val").forEach((el) => {
      if (el.classList.contains("manpower-line")) return;
      const prev = el.getAttribute("style") || "";
      const wordBody =
        "font-size:12.0pt;mso-ansi-font-size:12.0pt;font-family:'Arial',sans-serif;color:#000000;text-decoration:none;border:none;outline:none;background:transparent;";
      if (!/min-height/i.test(prev)) {
        el.setAttribute("style", "display:block;min-height:1.2em;" + wordBody + prev);
      } else {
        el.setAttribute("style", "display:block;" + wordBody + prev);
      }
    });
    root.querySelectorAll(".line-item-paste-row, .line-item-paste-row tr, .line-item-paste-row td").forEach((el) => {
      el.setAttribute("style", "border:none !important;outline:none !important;" + (el.getAttribute("style") || ""));
    });
    // Keep bullet rows (including Load Plan / Advance Loading) on one visual line in Outlook.
    root.querySelectorAll(".line-item > .export-val").forEach((el) => {
      el.setAttribute(
        "style",
        (el.getAttribute("style") || "") +
          ";display:inline !important;min-height:0 !important;line-height:115%;vertical-align:baseline;"
      );
    });
    root.querySelectorAll(".export-multiline").forEach((el) => {
      el.setAttribute("style", "white-space:pre-wrap;" + (el.getAttribute("style") || ""));
    });
    root.querySelectorAll("td .export-multiline").forEach((el) => {
      const prev = el.getAttribute("style") || "";
      el.setAttribute("style", "min-height:2.6em;line-height:1.35;" + prev);
    });

    root.querySelectorAll(".remarks-cell-wrap").forEach((el) => {
      el.setAttribute("style", "display:block;margin:0;padding:0;" + (el.getAttribute("style") || ""));
    });

    root.querySelectorAll(".offload-wrap").forEach((el) => {
      el.setAttribute(
        "style",
        "max-width:100%;width:100%;min-width:0;box-sizing:border-box;overflow:visible;margin:10px 0 0 0;" +
          (el.getAttribute("style") || "")
      );
    });

    root.querySelectorAll("table").forEach((t) => {
      if (t.classList && t.classList.contains("line-item-paste-row")) {
        t.setAttribute("border", "0");
        t.setAttribute("cellspacing", "0");
        t.setAttribute("cellpadding", "0");
        t.setAttribute(
          "style",
          "width:100%;border-collapse:collapse;table-layout:fixed;mso-table-lspace:0pt;mso-table-rspace:0pt;border:none;" +
            (t.getAttribute("style") || "")
        );
        return;
      }
      t.setAttribute("border", "1");
      t.setAttribute("cellspacing", "0");
      t.setAttribute("cellpadding", "9");
      t.setAttribute(
        "style",
          "width:100%;max-width:100%;min-width:0;border-collapse:collapse;table-layout:fixed;font-size:12.0pt;mso-ansi-font-size:12.0pt;font-family:Calibri,Arial,sans-serif;mso-table-lspace:0pt;mso-table-rspace:0pt;box-sizing:border-box;line-height:190%;mso-line-height-rule:at-least;" +
          (t.getAttribute("style") || "")
      );
    });
    root.querySelectorAll("th, td").forEach((cell) => {
      if (cell.closest && cell.closest("table.offload-table")) return;
      if (cell.closest && cell.closest("table.line-item-paste-row")) {
        cell.setAttribute(
          "style",
          "border:none;padding:2px 0;vertical-align:top;line-height:1.35;background-color:transparent;" +
            (cell.getAttribute("style") || "")
        );
        return;
      }
      cell.setAttribute(
        "style",
        "border:1px solid #d0d5e8;padding:10px;vertical-align:top;word-wrap:break-word;overflow-wrap:break-word;min-width:0;line-height:190%;mso-line-height-rule:at-least;background-color:#ffffff;" +
          (cell.getAttribute("style") || "")
      );
    });
    root.querySelectorAll("th").forEach((cell) => {
      cell.setAttribute("bgcolor", "#eef3fc");
      cell.setAttribute(
        "style",
        "background-color:#eef3fc;font-weight:bold;text-align:center;border:1px solid #d0d5e8;padding:10px;vertical-align:top;line-height:190%;mso-line-height-rule:at-least;" +
          (cell.getAttribute("style") || "")
      );
    });

    /* Final pass for offload table: preserve words in headers and stabilize Outlook wrapping. */
    root.querySelectorAll("table.offload-table").forEach((t) => {
      t.setAttribute(
        "style",
        "width:100%;max-width:100%;min-width:0;border-collapse:collapse;table-layout:fixed;font-size:11.5pt;mso-ansi-font-size:11.5pt;font-family:Calibri,Arial,sans-serif;mso-table-lspace:0pt;mso-table-rspace:0pt;box-sizing:border-box;line-height:190%;mso-line-height-rule:at-least;" +
          (t.getAttribute("style") || "")
      );
      t.querySelectorAll("thead th").forEach((th) => {
        th.setAttribute(
          "style",
          `background-color:#eef3fc;font-weight:bold;text-align:center;border:1px solid #d0d5e8;padding:10px 8px;vertical-align:middle;white-space:normal;word-break:keep-all;overflow-wrap:normal;word-wrap:normal;hyphens:none;line-height:1.35;font-size:9.2pt;mso-ansi-font-size:9.2pt;mso-line-height-rule:at-least;` +
            (th.getAttribute("style") || "")
        );
      });
      t.querySelectorAll("tbody td").forEach((td) => {
        const idx = Array.prototype.indexOf.call(td.parentNode ? td.parentNode.children : [], td);
        td.setAttribute(
          "style",
          "border:1px solid #d0d5e8;padding:12px;vertical-align:top;word-break:normal !important;overflow-wrap:normal !important;word-wrap:normal !important;line-height:1.9;mso-line-height-rule:at-least;background-color:#ffffff;" +
            (td.getAttribute("style") || "")
        );
        if (idx === 1) {
          const s = td.getAttribute("style") || "";
          td.setAttribute("style", "font-size:9.0pt;mso-ansi-font-size:9.0pt;line-height:1.5;" + s);
        }
        // Keep short identifiers on one line (DATE/FLIGHT/STD/DEST).
        if (idx >= 1 && idx <= 4) {
          const s = td.getAttribute("style") || "";
          td.setAttribute("style", "white-space:nowrap;word-break:keep-all !important;overflow-wrap:normal !important;" + s);
        }
        // Keep ULD entries multiline when needed.
        if (idx === 7) {
          const s = td.getAttribute("style") || "";
          td.setAttribute("style", "white-space:pre-wrap;word-break:normal !important;overflow-wrap:normal !important;" + s);
        }
      });
    });
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /**
   * Report-friendly bullets for multiline notes (used in copy/email output).
   * Each non-empty line becomes "• <text>".
   */
  function toBulletedLines(raw) {
    const lines = String(raw || "")
      .split(/\r?\n/)
      .map((line) => String(line || "").trim())
      .filter((line) => line);
    if (!lines.length) return "";
    return lines
      .map((line) => `    \u2022 ${line.replace(/^[\u2022\-*]\s*/, "")}`)
      .join("\n");
  }

  /** For copy/paste output: keep one visible bullet even when section is empty. */
  function toBulletedLinesWithFallback(raw) {
    const bullets = toBulletedLines(raw);
    return bullets || "    \u2022";
  }

  function normalizeIndentedBullets(raw) {
    const lines = String(raw || "").split(/\r?\n/);
    return lines
      .map((line) => {
        const text = String(line || "").trim();
        if (!text) return "";
        return `    \u2022 ${text.replace(/^[\u2022\-*]\s*/, "")}`;
      })
      .join("\n");
  }

  function buildWordClipboardHeadHtml() {
    return [
      '<meta charset="utf-8">',
      '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
      '<meta name="ProgId" content="Word.Document">',
      '<meta name="Generator" content="Microsoft Word 15">',
      '<meta name="Originator" content="Microsoft Word 15">',
      "<style>",
      "body,table,td,th,div,p,span,li{font-family:Calibri,Arial,sans-serif;mso-ansi-font-family:Calibri;mso-hansi-font-family:Calibri;font-size:12.0pt;mso-ansi-font-size:12.0pt;color:#000000;line-height:190%;mso-line-height-rule:at-least;}",
      "table{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;}",
      "p{margin:0 0 8px 0;}",
      "</style>",
    ].join("");
  }

  function wrapWordClipboardDocument(innerHtml) {
    const head = buildWordClipboardHeadHtml();
    const bodyStyle =
      "margin:0;padding:0;font-family:Calibri,Arial,sans-serif;mso-ansi-font-family:Calibri;mso-hansi-font-family:Calibri;font-size:12.0pt;mso-ansi-font-size:12.0pt;color:#000000;line-height:190%;mso-line-height-rule:at-least;";
    return `<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head>${head}</head><body style="${bodyStyle}">${innerHtml}</body></html>`;
  }

  function wrapOutlookFixedReportContainer(innerHtml) {
    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="1000" style="width:1000px;max-width:1000px;min-width:1000px;border-collapse:collapse;table-layout:fixed;mso-table-lspace:0pt;mso-table-rspace:0pt;margin:0;"><tr><td style="width:1000px;max-width:1000px;min-width:1000px;padding:0;margin:0;vertical-align:top;">${innerHtml}</td></tr></table>`;
  }

  /** Outlook can ignore head-level CSS; force font family/size inline on every node. */
  function hardInlineWordFonts(root) {
    const ensureStyleToken = (style, token, re) => (re.test(style) ? style : token + style);
    root.querySelectorAll("*").forEach((el) => {
      let s = el.getAttribute("style") || "";
      s = ensureStyleToken(s, "font-family:Calibri,Arial,sans-serif;", /\bfont-family\s*:/i);
      s = ensureStyleToken(s, "mso-ansi-font-family:Calibri;", /\bmso-ansi-font-family\s*:/i);
      s = ensureStyleToken(s, "mso-hansi-font-family:Calibri;", /\bmso-hansi-font-family\s*:/i);
      s = ensureStyleToken(s, "mso-bidi-font-family:Calibri;", /\bmso-bidi-font-family\s*:/i);
      s = ensureStyleToken(s, "font-size:12.0pt;", /\bfont-size\s*:/i);
      s = ensureStyleToken(s, "mso-ansi-font-size:12.0pt;", /\bmso-ansi-font-size\s*:/i);
      s = ensureStyleToken(s, "mso-bidi-font-size:12.0pt;", /\bmso-bidi-font-size\s*:/i);
      s = ensureStyleToken(s, "line-height:190%;", /\bline-height\s*:/i);
      s = ensureStyleToken(s, "mso-line-height-rule:at-least;", /\bmso-line-height-rule\s*:/i);
      el.setAttribute("style", s);
    });
  }

  /** Same folder as offload_report.html — served with the page for Copy → Outlook. */
  function reportSignatureBadgesImageUrl() {
    try {
      return new URL("signature-badges.png", window.location.href).href;
    } catch {
      return "signature-badges.png";
    }
  }

  let _signatureBadgesDataUriCache = null;

  /**
   * Outlook often drops &lt;img src="http..."&gt; on paste; embedding PNG as data: keeps the badge strip visible.
   * @returns {Promise<string|null>} data URL or null if fetch fails (e.g. file:// without server)
   */
  async function ensureSignatureBadgesDataUriForClipboard() {
    if (_signatureBadgesDataUriCache) return _signatureBadgesDataUriCache;
    const url = reportSignatureBadgesImageUrl();
    try {
      const r = await fetch(url, { cache: "force-cache" });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const blob = await r.blob();
      const dataUrl = await new Promise((resolve, reject) => {
        const rd = new FileReader();
        rd.onload = () => resolve(rd.result);
        rd.onerror = () => reject(rd.error);
        rd.readAsDataURL(blob);
      });
      _signatureBadgesDataUriCache = dataUrl;
      return dataUrl;
    } catch (err) {
      console.warn("Could not inline signature-badges.png for clipboard", err);
      return null;
    }
  }

  /** Outlook paste header tuned to match homepage visual identity. */
  function buildOutlookClipboardHeaderHtml() {
    const d = formatDisplayDate(state.shiftMeta.date || "");
    const t = state.shiftMeta.time || "";
    const title = state.shiftMeta.title || "";
    const dE = escapeHtml(d);
    const tE = escapeHtml(t);
    const titleE = escapeHtml(title);
    const bannerBg = "#1e5799";
    const ink = "#ffffff";
    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;margin-bottom:14px;max-width:100%;border:1px solid #d0d5e8;table-layout:fixed;"><tr><td width="3" bgcolor="#0b3a78" style="width:3px;min-width:3px;max-width:3px;background-color:#0b3a78;font-size:0;line-height:0;mso-line-height-rule:exactly;">&nbsp;</td><td width="1" bgcolor="#eef3fc" style="width:1px;min-width:1px;max-width:1px;background-color:#eef3fc;font-size:0;line-height:0;mso-line-height-rule:exactly;">&nbsp;</td><td bgcolor="${bannerBg}" style="background-color:${bannerBg};padding:18px 20px;vertical-align:top;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;"><tr><td style="vertical-align:top;font-family:Calibri,Arial,sans-serif;color:${ink};"><div style="mso-ansi-font-size:16.5pt;mso-bidi-font-size:16.5pt;font-size:16.5pt;font-weight:bold;line-height:190%;margin:0 0 10px 0;color:${ink};mso-line-height-rule:at-least;"><span style="mso-ansi-font-size:15.0pt;font-size:15.0pt;margin-right:6px;">✈️</span> Export Warehouse Activity Report</div><div style="mso-ansi-font-size:11.0pt;mso-bidi-font-size:11.0pt;font-size:11.0pt;line-height:190%;color:${ink};mso-line-height-rule:at-least;">Shift Date: <span style="color:#fde68a;font-weight:bold;">${dE}</span> &nbsp;|&nbsp; Time: <span style="color:#fde68a;font-weight:bold;">${tE}</span> &nbsp;|&nbsp; <span style="color:#fde68a;font-weight:bold;">${titleE}</span></div></td><td style="vertical-align:top;text-align:right;font-family:Calibri,Arial,sans-serif;mso-ansi-font-size:10.0pt;font-size:10.0pt;color:${ink};width:170px;"><div style="color:${ink};">Transom Cargo LLC.</div><div style="font-weight:bold;margin-top:4px;color:${ink};">Export Operations</div></td></tr></table></td></tr></table>`;
  }

  function buildOutlookClipboardSignatureHtml(badgeSrcResolved) {
    const nameRaw = getDutySupervisorDisplayName();
    const nameE = escapeHtml(nameRaw);
    const nameBlock = nameRaw
      ? `<span style="font-family:'Segoe Script','Brush Script MT','Lucida Handwriting',cursive;font-size:18px;font-weight:bold;color:#111111;line-height:1.1;">${nameE}</span>`
      : `<span style="font-family:Arial,Helvetica,sans-serif;font-size:9.0pt;color:#000000;">(Add name under Manpower → Supervisor)</span>`;
    const rawSrc = badgeSrcResolved || reportSignatureBadgesImageUrl();
    const badgeSrc = typeof rawSrc === "string" && rawSrc.startsWith("data:") ? rawSrc : escapeHtml(rawSrc);
    const badgesRow = `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin-top:7px;"><tr><td style="padding:0;"><img src="${badgeSrc}" alt="IATA CEIV PHARMA, IATA CEIV FRESH, ISO 9001:2015, ISO 45001:2018, HACCP, GDP, RA3" width="470" style="width:470px;max-width:470px;height:auto;border:0;display:block;-ms-interpolation-mode:bicubic;" /></td></tr></table>`;
    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin-top:22px;max-width:100%;font-family:Arial,Helvetica,sans-serif;background:#ffffff;"><tr><td style="padding:0;"><div style="font-size:11px;color:#6b7280;line-height:1.3;margin:0 0 4px 0;">Best Regards,</div><div style="margin:0 0 4px 0;">${nameBlock}</div><div style="font-size:11.5px;color:#1f2937;line-height:1.3;margin:0 0 7px 0;">Duty Supervisor – Export Operation</div><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;table-layout:fixed;"><tr><td valign="top" style="width:4px;min-width:4px;background:#c1121f;font-size:0;line-height:0;">&nbsp;</td><td valign="top" style="width:150px;padding:3px 7px 3px 6px;background:#fff4e6;"><div style="line-height:1;margin:0;"><span style="display:block;font-size:24px;font-weight:700;color:#c1121f;letter-spacing:0.2px;">TRANSOM</span><span style="display:block;font-size:11px;font-weight:700;color:#4b5563;letter-spacing:1.3px;text-transform:uppercase;margin-top:3px;">CARGO</span></div></td><td valign="top" style="width:2px;min-width:2px;background:#c1121f;font-size:0;line-height:0;">&nbsp;</td><td valign="top" style="padding:0 0 0 7px;"><div style="font-size:13px;line-height:1.38;color:#111111;"><strong style="font-weight:700;">Transom Cargo LLC.</strong><br/>P.O. Box: 618, P.C: 111<br/>Sultanate of Oman<br/>Phone: 97297474<br/><a href="http://www.transomcargo.com" style="color:#c1121f;text-decoration:none;">www.transomcargo.com</a></div></td></tr></table>${badgesRow}</td></tr></table>`;
  }

  /** Plain text for mailto / fallback: sections 1–10 only (no page title, no header meta, no toolbar). */
  function buildReportPlainBody() {
    const plainHeader = [
      "Export Warehouse Activity Report",
      "Transom Cargo LLC. | Export Operations",
      `Shift Date: ${formatDisplayDate(state.shiftMeta.date || "")} | Time: ${state.shiftMeta.time || ""} | ${state.shiftMeta.title || ""}`,
    ].join("\n");
    const plainSig = [
      "",
      "—",
      "Best Regards,",
      getDutySupervisorDisplayName() || "(Duty Supervisor — set under Manpower → Supervisor)",
      "Duty Supervisor – Export Operation",
      "",
      "Transom Cargo LLC.",
      "P.O. Box: 618, P.C: 111",
      "Sultanate of Oman",
      "Phone No. 97297474",
      "www.transomcargo.com",
    ].join("\n");
    const groupedText = state.operationalActivities
      .map((group) =>
        `${group.title}:\n${
          group.items.length ? group.items.map((i) => `    \u2022 ${String(i || "").trim()}`).join("\n") : "NIL"
        }`
      )
      .join("\n\n");
    const manpowerText = state.manpowerSections
      .map((section) => {
        const lines = (section.items || []).map((i) => String(i).trim()).filter((i) => i);
        return `${section.title}:\n${lines.length ? lines.map((i) => `    \u2022 ${i}`).join("\n") : "NIL"}`;
      })
      .join("\n\n");
    const offloadLine = (v) => String(v ?? "").replace(/\r?\n/g, " ").trim();
    const offloadText = state.offloads.length
      ? state.offloads
          .map(
            (row) =>
              `${row.item}. ${offloadLine(row.date)} | ${offloadLine(row.flight)} | ${offloadLine(row.std)} | ${offloadLine(row.destination)} | ${offloadLine(row.emailTime)} | ${offloadLine(row.reason)} | ${offloadLine(row.remarks)}`
          )
          .join("\n")
      : "NIL — No offload data recorded for this shift.";

    const core = [
      "1. OPERATIONAL ACTIVITIES",
      groupedText,
      "",
      "2. BRIEFINGS CONDUCTED",
      state.briefings.map((i) => `    \u2022 ${String(i || "").trim()}`).join("\n"),
      "",
      "3. FLIGHT PERFORMANCE",
      toBulletedLinesWithFallback(state.flightPerformance),
      "",
      "OPERATIONAL NOTES",
      state.operationalNotes.map((i) => `    \u2022 ${String(i || "").trim()}`).join("\n"),
      "",
      "4. CHECKS & COMPLIANCE",
      toBulletedLinesWithFallback(state.checksCompliance),
      "",
      "OFFLOADING CARGO",
      offloadText,
      "",
      "5. SAFETY",
      toBulletedLinesWithFallback(state.safety),
      "",
      "6. MANPOWER",
      manpowerText,
      "",
      "7. EQUIPMENT STATUS",
      toBulletedLinesWithFallback(state.equipmentStatus),
      "",
      "8. HANDOVER DETAILS",
      toBulletedLines(state.handoverDetails),
      "",
      "9. SPECIAL H/O",
      toBulletedLinesWithFallback(state.specialHO),
      "",
      "10. OTHER",
      toBulletedLinesWithFallback(state.otherText),
    ].join("\n");
    return [plainHeader, "", core, plainSig].join("\n");
  }

  function cloneReportExportInnerHtml() {
    const root = el("reportExportRoot");
    if (!root) return "";

    // cloneNode(true) copies HTML *attributes*, not the live .value/.checked properties.
    // Snapshot all values from the live DOM first, then apply them to the clone by position.
    const liveInputs = Array.from(root.querySelectorAll("input"));
    const liveTextareas = Array.from(root.querySelectorAll("textarea"));
    const inputVals = liveInputs.map((inp) => inp.value);
    const textareaVals = liveTextareas.map((ta) => ta.value);

    const clone = root.cloneNode(true);

    // Restore values on the clone (same tree order, so index positions match).
    Array.from(clone.querySelectorAll("input")).forEach((inp, i) => {
      if (i < inputVals.length) inp.value = inputVals[i];
    });
    Array.from(clone.querySelectorAll("textarea")).forEach((ta, i) => {
      if (i < textareaVals.length) ta.value = textareaVals[i];
    });

    clone.querySelectorAll(".delete-btn").forEach((n) => n.remove());
    clone.querySelectorAll("input").forEach((inp) => {
      const span = document.createElement("span");
      const isManpowerLine = inp.classList && inp.classList.contains("manpower-line");
      const isManpowerSectionTitle = inp.classList && inp.classList.contains("manpower-section-title");
      if (isManpowerLine) {
        span.className = "export-val manpower-line";
      } else if (isManpowerSectionTitle) {
        span.className = "export-val manpower-section-title";
      } else {
        span.className = "export-val";
      }
      span.dataset.segment = String(inp.dataset.segment || "").trim();
      span.textContent = inp.value;
      inp.replaceWith(span);
    });
    clone.querySelectorAll(".line-item").forEach((row) => {
      const flight = row.querySelector('span.export-val[data-segment="flight"]');
      const phrase = row.querySelector('span.export-val[data-segment="phrase"]');
      if (!flight && !phrase) return;
      const f = String((flight && flight.textContent) || "").trim();
      const p = String((phrase && phrase.textContent) || "").trim();
      const joined = f && p ? `${f} - ${p}` : f || p;
      const one = document.createElement("span");
      one.className = "export-val";
      one.textContent = joined;
      [flight, phrase].forEach((n) => {
        if (n && n.parentNode === row) n.remove();
      });
      row.insertBefore(one, row.firstChild);
    });
    clone.querySelectorAll("textarea").forEach((ta) => {
      const div = document.createElement("div");
      const keepClasses = String(ta.className || "")
        .split(/\s+/)
        .filter((c) => c && c !== "offload-cell");
      div.className = ["export-val", "export-multiline", ...keepClasses].join(" ").trim();
      const id = String(ta.id || "").trim();
      const shouldBullet =
        id === "flightPerformance" ||
        id === "checksCompliance" ||
        id === "safety" ||
        id === "handoverDetails" ||
        id === "specialHO" ||
        id === "otherText";
      const shouldBulletWithFallback =
        id === "flightPerformance" ||
        id === "checksCompliance" ||
        id === "safety" ||
        id === "equipmentStatus" ||
        id === "handoverDetails" ||
        id === "specialHO" ||
        id === "otherText";
      if (shouldBulletWithFallback) {
        div.textContent = toBulletedLinesWithFallback(ta.value);
      } else if (shouldBullet) {
        div.textContent = toBulletedLines(ta.value);
      } else {
        div.textContent = ta.value;
      }
      ta.replaceWith(div);
    });
    return clone.innerHTML;
  }

  async function buildReportHtmlForClipboard() {
    const badgeDataUri = await ensureSignatureBadgesDataUriForClipboard();
    const inner = cloneReportExportInnerHtml();
    const headerHtml = buildOutlookClipboardHeaderHtml();
    const sigHtml = buildOutlookClipboardSignatureHtml(badgeDataUri);
    if (!inner.trim()) {
      return wrapWordClipboardDocument(
        `${wrapOutlookFixedReportContainer(
          `${headerHtml}<p style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">${escapeHtml(
          "(Report body is empty.)"
          )}</p>${sigHtml}`
        )}`
      );
    }
    const wrap = document.createElement("div");
    wrap.className = "report-fragment";
    wrap.innerHTML = inner;
    // Force unified section-title look in clipboard HTML every time.
    convertSectionTitlesToOutlookShadeTables(wrap);
    normalizeOffloadTableForClipboard(wrap);
    restructureLineItemsForWordPaste(wrap);
    applyOutlookInlineClipboardStyles(wrap);
    restructureManpowerItemsForWordPaste(wrap);
    hardInlineWordFonts(wrap);
    const fragmentHtml = wrap.outerHTML;
    return wrapWordClipboardDocument(wrapOutlookFixedReportContainer(`${headerHtml}${fragmentHtml}${sigHtml}`));
  }

  async function copyReportAsEmail() {
    const inner = cloneReportExportInnerHtml();
    const headerHtml = buildOutlookClipboardHeaderHtml();
    // Keep direct image URLs for better email-client fetch behavior.
    const sigHtml = buildOutlookClipboardSignatureHtml(reportSignatureBadgesImageUrl());
    const wrap = document.createElement("div");
    wrap.className = "report-fragment";
    wrap.innerHTML = inner;

    convertSectionTitlesToOutlookShadeTables(wrap);
    normalizeOffloadTableForClipboard(wrap);
    restructureLineItemsForWordPaste(wrap);
    applyOutlookInlineClipboardStyles(wrap);
    restructureManpowerItemsForWordPaste(wrap);
    hardInlineWordFonts(wrap);

    // Remove interactive/web-only leftovers.
    wrap.querySelectorAll("button, script, .hidden-print, .delete-btn, .toolbar, .bottom-tools, .modal-backdrop").forEach((n) =>
      n.remove()
    );
    wrap.querySelectorAll("[contenteditable]").forEach((n) => n.removeAttribute("contenteditable"));

    const fragmentHtml = wrap.outerHTML;
    const html = wrapWordClipboardDocument(
      wrapOutlookFixedReportContainer(`${headerHtml}${fragmentHtml}${sigHtml}`)
    );
    const htmlBlob = new Blob([html], { type: "text/html;charset=utf-8" });

    if (navigator.clipboard && typeof ClipboardItem !== "undefined") {
      await navigator.clipboard.write([new ClipboardItem({ "text/html": htmlBlob })]);
      return;
    }

    let copiedViaHtmlEvent = false;
    const onCopy = (ev) => {
      if (!ev.clipboardData) return;
      ev.clipboardData.setData("text/html", html);
      ev.preventDefault();
      copiedViaHtmlEvent = true;
    };
    document.addEventListener("copy", onCopy);
    const ok = document.execCommand("copy");
    document.removeEventListener("copy", onCopy);
    if (!ok || !copiedViaHtmlEvent) {
      throw new Error("HTML clipboard copy failed");
    }
  }

  async function copyReportToClipboard() {
    await copyReportAsEmail();
  }

  function saveDraft() {
    const draft = {
      _reportDate: state.activeDate || "",
      _reportShift: state.activeShift != null ? state.activeShift : "",
      shiftMeta: state.shiftMeta,
      operationalActivities: state.operationalActivities,
      briefings: state.briefings,
      flightPerformance: state.flightPerformance,
      operationalNotes: state.operationalNotes,
      checksCompliance: state.checksCompliance,
      offloads: state.offloads,
      safety: state.safety,
      manpowerSections: state.manpowerSections,
      equipmentStatus: state.equipmentStatus,
      handoverDetails: state.handoverDetails,
      otherText: state.otherText,
      specialHO: state.specialHO,
      recipients: state.recipients,
      scheduledSendAt: state.scheduledSendAt || "",
      scheduledSendEnabled: !!state.scheduledSendEnabled,
      _scheduledSendLastFiredAt: state._scheduledSendLastFiredAt || "",
    };
    localStorage.setItem(draftStorageKey(), JSON.stringify(draft));
    persistOffloadFlightHints();
  }

  function draftMatchesCurrentReport(draft) {
    if (!draft || typeof draft !== "object") return false;
    const d = state.activeDate;
    const s = state.activeShift;
    if (draft._reportDate && d && draft._reportDate !== d) return false;
    if (!draft._reportDate && draft.shiftMeta && draft.shiftMeta.date && d && draft.shiftMeta.date !== d) {
      return false;
    }
    if (draft._reportShift !== undefined && draft._reportShift !== "" && s != null && draft._reportShift !== s) {
      return false;
    }
    return true;
  }

  /**
   * When shift tabs exist, only restore roster/manpower lists from a draft that was saved for the
   * same shift. Legacy drafts without _reportShift must not overwrite morning/afternoon/night names.
   */
  function shouldApplyDraftManpower(draft) {
    const countNonEmptyNames = (sections) =>
      (Array.isArray(sections) ? sections : []).reduce((sum, sec) => {
        const items = Array.isArray(sec && sec.items) ? sec.items : [];
        return sum + items.filter((x) => String(x || "").trim()).length;
      }, 0);

    if (!draft || typeof draft !== "object") return false;
    if (!Array.isArray(draft.manpowerSections) || !draft.manpowerSections.length) return false;
    // Safety: ignore broken local drafts that accidentally saved empty manpower lists.
    const draftNameCount = countNonEmptyNames(draft.manpowerSections);
    if (draftNameCount === 0) return false;
    if (!state.shiftsFromServer || !state.activeShift) return true;
    const serverPack = state.shiftsFromServer[state.activeShift];
    const serverNameCount = countNonEmptyNames(serverPack && serverPack.manpowerSections);
    if (serverNameCount > 0 && draftNameCount === 0) return false;
    const w = String(draft._reportShift || "").trim();
    if (!w) return false;
    return w === state.activeShift;
  }

  function restoreMissingManpowerSectionsFromServer() {
    if (!state.shiftsFromServer || !state.activeShift || !Array.isArray(state.manpowerSections)) return false;
    const pack = state.shiftsFromServer[state.activeShift];
    const serverSections = Array.isArray(pack && pack.manpowerSections) ? pack.manpowerSections : [];
    if (!serverSections.length) return false;

    const countNonEmpty = (items) => (Array.isArray(items) ? items.filter((x) => String(x || "").trim()).length : 0);
    const byTitle = new Map();
    serverSections.forEach((sec) => {
      const t = String((sec && sec.title) || "").trim().toLowerCase();
      if (!t) return;
      byTitle.set(t, sec);
    });

    let changed = false;
    let currentTotal = 0;
    state.manpowerSections.forEach((sec) => {
      currentTotal += countNonEmpty(sec && sec.items);
    });

    // If draft wiped all names, fully recover from server roster for the active shift.
    if (currentTotal === 0) {
      state.manpowerSections = deepClone(serverSections);
      state.manpowerSections.forEach((sec) => ensureManpowerRowForEditing(sec));
      return true;
    }

    // If only some sections are empty, recover those sections by title.
    state.manpowerSections.forEach((sec) => {
      const key = String((sec && sec.title) || "").trim().toLowerCase();
      if (!key || countNonEmpty(sec && sec.items) > 0) return;
      const serverSec = byTitle.get(key);
      if (!serverSec || countNonEmpty(serverSec.items) === 0) return;
      sec.items = deepClone(serverSec.items);
      ensureManpowerRowForEditing(sec);
      changed = true;
    });
    return changed;
  }

  /** @returns {boolean} true if draft contained manpower lists (so server roster should not overwrite). */
  function loadDraft() {
    try {
      const raw = localStorage.getItem(draftStorageKey());
      if (!raw) return false;
      const draft = JSON.parse(raw);

      if (!draftMatchesCurrentReport(draft)) {
        return false;
      }

      let appliedManpower = false;

      state.operationalActivities = draft.operationalActivities || state.operationalActivities;
      state.briefings = draft.briefings || state.briefings;
      state.flightPerformance = draft.flightPerformance || state.flightPerformance;
      state.operationalNotes = draft.operationalNotes || state.operationalNotes;
      state.checksCompliance = draft.checksCompliance || state.checksCompliance;
      state.offloads = draft.offloads || state.offloads;
      state.safety = draft.safety || state.safety;

      if (shouldApplyDraftManpower(draft)) {
        if (state.shiftsFromServer) {
          state.manpowerSections = deepClone(draft.manpowerSections);
          appliedManpower = true;
        } else {
          state.manpowerSections = draft.manpowerSections;
          appliedManpower = true;
        }
      }
      if (restoreMissingManpowerSectionsFromServer()) {
        appliedManpower = true;
      }

      state.equipmentStatus = draft.equipmentStatus || state.equipmentStatus;
      state.handoverDetails = draft.handoverDetails || state.handoverDetails;
      state.otherText = draft.otherText || "";
      state.specialHO = draft.specialHO || "";
      state.recipients = normalizeRecipientsShape(draft.recipients || state.recipients);
      state.scheduledSendAt = draft.scheduledSendAt || "";
      state.scheduledSendEnabled = !!draft.scheduledSendEnabled;
      state._scheduledSendLastFiredAt = draft._scheduledSendLastFiredAt || "";
      stripExcludedEmployeesFromManpower();
      return appliedManpower;
    } catch (err) {
      console.error("Failed to load draft", err);
      return false;
    }
  }

  function attachFlightExpansionHelpers() {
    if (!window.flightAutocomplete) return;
    document.querySelectorAll('.opact-flight-input[data-segment="flight"]').forEach((inp) => {
      const g = +inp.dataset.group;
      const ii = +inp.dataset.index;
      window.flightAutocomplete.attach(inp, `opact-flight-${g}-${ii}`, (picked) => {
        const pickedText = window.flightAutocomplete.formatFlight(picked);
        inp.value = pickedText;
        if (state.operationalActivities[g] && state.operationalActivities[g].items[ii] !== undefined) {
          const current = splitOperationalTwoFieldLine(state.operationalActivities[g].items[ii]);
          state.operationalActivities[g].items[ii] = composeOperationalTwoFieldLine(pickedText, current.phrase);
          saveDraft();
        }
      });
    });
  }

  function attachPhraseHelpers() {
    if (!window.phraseAutocomplete) return;

    const opPhraseKeys = ["loadPlan", "advanceLoading", "csdRescreening"];
    document.querySelectorAll(".opact-input").forEach((inp) => {
      const gi = +inp.dataset.group;
      const segment = (inp.dataset.segment || "full").trim();
      if (segment === "flight") return;
      const phraseKey = (inp.dataset.phraseKey || "").trim() || (segment === "full" ? opPhraseKeys[gi] : "");
      if (!phraseKey) return;
      window.phraseAutocomplete.attachInput(
        inp,
        phraseKey,
        (value) => {
          const g = +inp.dataset.group;
          const ii = +inp.dataset.index;
          if (state.operationalActivities[g] && state.operationalActivities[g].items[ii] !== undefined) {
            if (segment === "phrase") {
              const current = splitOperationalTwoFieldLine(state.operationalActivities[g].items[ii]);
              const normalized = toSentenceCaseText(value);
              inp.value = normalized;
              state.operationalActivities[g].items[ii] = composeOperationalTwoFieldLine(current.flight, normalized);
            } else if (segment === "flight") {
              const current = splitOperationalTwoFieldLine(state.operationalActivities[g].items[ii]);
              state.operationalActivities[g].items[ii] = composeOperationalTwoFieldLine(value, current.phrase);
            } else {
              const normalized = phraseKey === "csdRescreening" ? String(value || "").toUpperCase() : toSentenceCaseText(value);
              inp.value = normalized;
              state.operationalActivities[g].items[ii] = normalized;
            }
            saveDraft();
          }
        },
        { preserveCase: segment === "phrase" }
      );
    });

    window.phraseAutocomplete.attachTextarea(el("handoverDetails"), "handoverDetails", (value) => {
      state.handoverDetails = normalizeIndentedBullets(toSentenceCaseText(value));
      saveDraft();
    }, { preserveCase: true, pickOnEnter: false });

    window.phraseAutocomplete.attachTextarea(el("otherText"), "other", (value) => {
      state.otherText = normalizeIndentedBullets(toSentenceCaseText(value));
      saveDraft();
    }, { preserveCase: true, pickOnEnter: false });

    window.phraseAutocomplete.attachTextarea(el("specialHO"), "specialHO", (value) => {
      state.specialHO = normalizeIndentedBullets(toSentenceCaseText(value));
      saveDraft();
    }, { preserveCase: true, pickOnEnter: false });
  }

  function bindStaticEvents() {
    el("flightPerformance").addEventListener("input", (e) => {
      state.flightPerformance = normalizeIndentedBullets(toSentenceCaseText(e.target.value));
      e.target.value = state.flightPerformance;
      saveDraft();
    });
    el("checksCompliance").addEventListener("input", (e) => {
      state.checksCompliance = normalizeIndentedBullets(toSentenceCaseText(e.target.value));
      e.target.value = state.checksCompliance;
      saveDraft();
    });
    el("safety").addEventListener("input", (e) => {
      state.safety = normalizeIndentedBullets(toSentenceCaseText(e.target.value));
      e.target.value = state.safety;
      saveDraft();
    });
    el("equipmentStatus").addEventListener("input", (e) => {
      state.equipmentStatus = normalizeIndentedBullets(toSentenceCaseText(e.target.value));
      e.target.value = state.equipmentStatus;
      saveDraft();
    });
    el("equipmentStatus").addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const area = e.target;
      const value = String(area.value || "");
      const start = typeof area.selectionStart === "number" ? area.selectionStart : value.length;
      const end = typeof area.selectionEnd === "number" ? area.selectionEnd : start;
      const left = value.slice(0, start);
      const right = value.slice(end);
      const joiner = left.endsWith("\n") || !left.length ? "    \u2022 " : "\n    \u2022 ";
      const next = left + joiner + right;
      area.value = next;
      try {
        const pos = (left + joiner).length;
        area.selectionStart = area.selectionEnd = pos;
      } catch (_) {}
      state.equipmentStatus = normalizeIndentedBullets(toSentenceCaseText(area.value));
      area.value = state.equipmentStatus;
      saveDraft();
    });
    const forceBulletedTextarea = (id, stateKey) => {
      const node = el(id);
      if (!node) return;
      node.addEventListener("input", (e) => {
        const normalized = normalizeIndentedBullets(toSentenceCaseText(e.target.value));
        state[stateKey] = normalized;
        e.target.value = normalized;
        saveDraft();
      });
    };
    forceBulletedTextarea("handoverDetails", "handoverDetails");
    forceBulletedTextarea("specialHO", "specialHO");
    forceBulletedTextarea("otherText", "otherText");
    if (!window.phraseAutocomplete) {
      wireSpecialHoTextarea();
    }
    /* handoverDetails, otherText, specialHO: uppercase + save via phrase attachTextarea (see attachPhraseHelpers). */

    const recipientInputs = [
      ["to", "newRecipientTo", "addRecipientToBtn"],
      ["cc", "newRecipientCc", "addRecipientCcBtn"],
      ["bcc", "newRecipientBcc", "addRecipientBccBtn"],
    ];
    recipientInputs.forEach(([kind, inputId, btnId]) => {
      const input = el(inputId);
      const btn = el(btnId);
      if (btn) btn.addEventListener("click", () => addRecipientTo(kind));
      if (input) {
        input.addEventListener("keydown", (e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            addRecipientTo(kind);
          }
        });
      }
    });
    const copyBtn = el("copyBtn");
    if (copyBtn) {
      copyBtn.addEventListener("click", async () => {
        await copyReportToClipboard();
      });
    }
    const emailBtn = el("emailBtn");
    if (emailBtn) {
      emailBtn.addEventListener("click", async () => {
        setEmailButtonState("sending");
        const ok = await sendEmailNow();
        const statusEl = el("gmailStatus");
        if (ok) {
          setEmailButtonState("success");
          if (statusEl) statusEl.textContent = "Email sent successfully via Gmail.";
        } else {
          setEmailButtonState("idle");
          if (statusEl) statusEl.textContent = "Send failed. Please connect Gmail or check recipients.";
        }
      });
    }
    const connectGmailBtn = el("connectGmailBtn");
    if (connectGmailBtn) {
      connectGmailBtn.addEventListener("click", () => {
        connectGmailFlow().catch(console.error);
      });
    }
    const autoSendAtEl = el("autoSendAt");
    if (autoSendAtEl) {
      autoSendAtEl.addEventListener("change", (e) => {
        state.scheduledSendAt = String(e.target.value || "");
        saveDraft();
        scheduleSendTimerFromState();
      });
    }
    const scheduleBtn = el("scheduleSendBtn");
    if (scheduleBtn) {
      scheduleBtn.addEventListener("click", () => {
        const when = autoSendAtEl ? String(autoSendAtEl.value || "") : "";
        if (!when) {
          const statusEl = el("scheduleStatus");
          if (statusEl) statusEl.textContent = "Please select a date/time first.";
          return;
        }
        state.scheduledSendAt = when;
        state.scheduledSendEnabled = true;
        state._scheduledSendLastFiredAt = "";
        saveDraft();
        scheduleSendTimerFromState();
      });
    }
    const clearScheduleBtn = el("clearScheduleBtn");
    if (clearScheduleBtn) {
      clearScheduleBtn.addEventListener("click", () => {
        state.scheduledSendEnabled = false;
        state.scheduledSendAt = "";
        state._scheduledSendLastFiredAt = "";
        if (autoSendAtEl) autoSendAtEl.value = "";
        saveDraft();
        scheduleSendTimerFromState();
      });
    }
    const printBtn = el("printBtn");
    if (printBtn) {
      printBtn.addEventListener("click", () => window.print());
    }

    const exportHintsBtn = el("exportFlightHintsBtn");
    if (exportHintsBtn && window.flightHintCache) {
      exportHintsBtn.addEventListener("click", () => window.flightHintCache.downloadExport());
    }

    const backdrop = el("resetConfirmBackdrop");
    const resetBtn = el("resetReportBtn");
    if (backdrop && resetBtn) {
      const openResetModal = () => {
        if (!state._resetBaseline) return;
        backdrop.hidden = false;
      };
      const closeResetModal = () => {
        backdrop.hidden = true;
      };
      resetBtn.addEventListener("click", openResetModal);
      const cancelBtn = el("resetCancelBtn");
      const confirmBtn = el("resetConfirmBtn");
      if (cancelBtn) cancelBtn.addEventListener("click", closeResetModal);
      backdrop.addEventListener("click", (e) => {
        if (e.target === backdrop) closeResetModal();
      });
      if (confirmBtn) {
        confirmBtn.addEventListener("click", () => {
          restoreResetSnapshot();
          try {
            localStorage.removeItem(draftStorageKey());
          } catch (e) {
            /* ignore */
          }
          closeResetModal();
          renderAll();
          saveDraft();
        });
      }
      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && !backdrop.hidden) closeResetModal();
      });
    }

    window.addEventListener("beforeunload", () => {
      try {
        saveDraft();
      } catch (e) {
        /* ignore */
      }
    });
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden") {
        try {
          saveDraft();
        } catch (e) {
          /* ignore */
        }
      } else {
        runScheduledSendIfDue().catch(console.error);
        scheduleSendTimerFromState();
        refreshGmailStatus().catch(() => {});
      }
    });
  }

  function addRecipientTo(kind) {
    const map = {
      to: "newRecipientTo",
      cc: "newRecipientCc",
      bcc: "newRecipientBcc",
    };
    if (!Object.prototype.hasOwnProperty.call(map, kind)) return;
    const input = el(map[kind]);
    if (!input) return;
    const values = parseRecipientInput(input.value);
    if (!values.length) return;
    const current = normalizeRecipientsShape(state.recipients);
    if (!Array.isArray(current[kind])) current[kind] = [];
    values.forEach((valueRaw) => {
      const value = valueRaw.trim();
      if (!isValidEmailBasic(value)) return;
      const v = value.toLowerCase();
      if (!current[kind].includes(v)) current[kind].push(v);
    });
    state.recipients = current;
    input.value = "";
    saveDraft();
    renderRecipients();
    scheduleSendTimerFromState();
    persistRecipientsToServer();
  }

  bindStaticEvents();
  startAutoTodayWatcher();
  loadData();
})();
