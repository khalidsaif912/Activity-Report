window.employeeAutocomplete = {
  employees: [],

  async load(url = "../../data/report/employees.json") {
    try {
      const res = await fetch(url);
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (text.trimStart().startsWith("<")) {
        throw new Error(`Expected JSON but got HTML (wrong URL?). ${url}`);
      }
      const raw = JSON.parse(text);
      this.employees = Array.isArray(raw) ? raw.slice() : [];
    } catch (err) {
      console.error("Failed to load employees.json", url, err);
      this.employees = [];
    }
  },

  _normalize(value) {
    return String(value || "").trim().toUpperCase();
  },

  _splitNameRole(value) {
    const v = String(value || "").trim();
    const m = v.match(/^(.+?)\s*-\s*(.*)$/);
    if (!m) return { hasDash: false, namePart: v, rolePart: "" };
    return { hasDash: true, namePart: String(m[1] || "").trim(), rolePart: String(m[2] || "").trim() };
  },

  _employeeMatchesByName(query) {
    const q = this._normalize(query);
    let matches = this.employees;
    if (q) {
      const qNoSn = q.replace(/^SN/, "");
      matches = this.employees.filter((item) => {
        const v = this._normalize(item);
        const vNoSn = v.replace(/^SN/, "");
        return v.startsWith(q) || v.includes(q) || vNoSn.startsWith(qNoSn);
      });
    }
    return matches.slice(0, 20);
  },

  _mergeUnique(list, max) {
    const out = [];
    const seen = new Set();
    list.forEach((x) => {
      const v = String(x || "").trim();
      if (!v) return;
      const k = v.toUpperCase();
      if (seen.has(k)) return;
      seen.add(k);
      out.push(v);
    });
    return out.slice(0, max);
  },

  attach(input, key) {
    const mark = `emp:${key}`;
    if (input.dataset.empAttachMark === mark) return;
    input.dataset.empAttachMark = mark;

    const listId = `employee-list-${key}`;
    input.setAttribute("list", listId);

    let datalist = document.getElementById(listId);
    if (!datalist) {
      datalist = document.createElement("datalist");
      datalist.id = listId;
      document.body.appendChild(datalist);
    }

    const refresh = () => {
      const { hasDash, namePart, rolePart } = this._splitNameRole(input.value);
      const names = this._employeeMatchesByName(namePart);
      let matches = [];

      if (hasDash) {
        names.forEach((name) => {
          const roles =
            window.manpowerRoleHintCache && typeof window.manpowerRoleHintCache.getRolesForName === "function"
              ? window.manpowerRoleHintCache.getRolesForName(name, rolePart, 6)
              : [];
          roles.forEach((role) => matches.push(`${name} - ${role}`));
        });
      } else {
        names.forEach((name) => {
          const topRole =
            window.manpowerRoleHintCache && typeof window.manpowerRoleHintCache.getTopRoleForName === "function"
              ? window.manpowerRoleHintCache.getTopRoleForName(name)
              : "";
          if (topRole) matches.push(`${name} - ${topRole}`);
          matches.push(name);
        });
      }

      const finalMatches = this._mergeUnique(matches, 20);
      datalist.innerHTML = "";
      finalMatches.forEach((item) => {
        const option = document.createElement("option");
        option.value = item;
        datalist.appendChild(option);
      });
    };

    input.addEventListener("input", refresh);
    input.addEventListener("focus", refresh);
    input.addEventListener("blur", () => {
      if (window.manpowerRoleHintCache && typeof window.manpowerRoleHintCache.recordFromLine === "function") {
        window.manpowerRoleHintCache.recordFromLine(input.value);
      }
    });
    refresh();
  }
};