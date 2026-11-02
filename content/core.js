/* Docent Tools voor Teams: kern.
 * Modulesysteem, instellingen, en de brug naar de API-calls van de pagina (zie netlog.js).
 * Modules registreren zich via TT.register({ id, init, destroy }).
 */
(() => {
  if (window.TT) return;

  const DEFAULTS = globalThis.TT_DEFAULTS;
  const modules = new Map();
  let settings = structuredClone(DEFAULTS);
  const listeners = new Set();
  const GRAPH = "https://graph.microsoft.com/v1.0";

  const TT = {
    version: chrome.runtime.getManifest().version,
    isTop: window === window.top,
    GRAPH,
    get settings() { return settings; },

    register(mod) {
      if (!mod || !mod.id) throw new Error("Module zonder id");
      modules.set(mod.id, { ...mod, active: false });
    },

    async loadSettings() {
      const stored = (await chrome.storage.sync.get("settings")).settings || {};
      settings = deepMerge(structuredClone(DEFAULTS), stored);
      return settings;
    },
    async saveSettings(patch) {
      settings = deepMerge(settings, patch);
      await chrome.storage.sync.set({ settings });
      listeners.forEach((fn) => fn(settings));
    },
    onSettingsChange(fn) { listeners.add(fn); return () => listeners.delete(fn); },
    notify() { listeners.forEach((fn) => fn(settings)); },

    startAll() {
      for (const mod of modules.values()) {
        const enabled = settings[mod.id] ? settings[mod.id].enabled !== false : true;
        if (enabled && !mod.active) {
          try { mod.init(TT); mod.active = true; }
          catch (e) { console.warn(`[TT] module ${mod.id} faalde`, e); }
        } else if (!enabled && mod.active) {
          try { mod.destroy && mod.destroy(TT); } catch (e) {}
          mod.active = false;
        }
      }
    },
    module(id) { return modules.get(id); },

    /* ---------- DOM-helpers ---------- */
    el(tag, attrs = {}, children = []) {
      const node = document.createElement(tag);
      for (const [k, v] of Object.entries(attrs)) {
        if (k === "class") node.className = v;
        else if (k === "text") node.textContent = v;
        else if (k === "html") node.innerHTML = v;
        else if (k.startsWith("on") && typeof v === "function") node.addEventListener(k.slice(2), v);
        else if (v !== null && v !== undefined && v !== false) node.setAttribute(k, v === true ? "" : v);
      }
      for (const c of [].concat(children)) {
        if (c === null || c === undefined || c === false) continue;
        node.append(typeof c === "string" || typeof c === "number" ? document.createTextNode(String(c)) : c);
      }
      return node;
    },
    toast(msg, kind = "info", ms = 2800) {
      let host = document.getElementById("tt-toasts");
      if (!host) { host = TT.el("div", { id: "tt-toasts" }); document.body.append(host); }
      const t = TT.el("div", { class: `tt-toast tt-toast-${kind}`, text: msg });
      host.append(t);
      requestAnimationFrame(() => t.classList.add("tt-show"));
      setTimeout(() => { t.classList.remove("tt-show"); setTimeout(() => t.remove(), 250); }, ms);
    },
    async copy(text) {
      try { await navigator.clipboard.writeText(text); TT.toast("Gekopieerd naar klembord", "ok"); return true; }
      catch { TT.toast("Kopiëren mislukt", "err"); return false; }
    },
    download(filename, content, mime = "text/plain;charset=utf-8") {
      const blob = content instanceof Blob ? content : new Blob([content], { type: mime });
      const a = TT.el("a", { href: URL.createObjectURL(blob), download: filename });
      document.body.append(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    },
    /* Donker of licht: instelling, anders het systeem. Teams zelf zet geen betrouwbare klasse. */
    dark() {
      const t = settings.thema || "auto";
      if (t === "donker") return true;
      if (t === "licht") return false;
      return matchMedia("(prefers-color-scheme: dark)").matches;
    },
    fmtDate(d, opts = { day: "numeric", month: "short" }) {
      if (!d) return "";
      const x = d instanceof Date ? d : new Date(d);
      return isNaN(x) ? "" : x.toLocaleDateString("nl-NL", opts);
    },
    fmtDateTime(d) {
      if (!d) return "";
      const x = new Date(d);
      return isNaN(x) ? "" : x.toLocaleDateString("nl-NL", { day: "numeric", month: "short" }) + " " + x.toLocaleTimeString("nl-NL", { hour: "2-digit", minute: "2-digit" });
    },
    pct(n) { return n == null || isNaN(n) ? "" : Math.round(n * 100) + "%"; },

    /* ---------- Frames ---------- */
    /* Alle vensters: dit venster en alle (ook cross-origin) frames eronder; alleen postMessage wordt gebruikt. */
    frames() {
      const out = [];
      const walk = (w, depth) => {
        out.push(w);
        if (depth > 6) return;
        let n = 0; try { n = w.frames.length; } catch { return; }
        for (let i = 0; i < n; i++) { try { walk(w.frames[i], depth + 1); } catch {} }
      };
      try { walk(window.top, 0); } catch { walk(window, 0); }
      return out;
    },

    /* ---------- Brug naar de pagina (netlog.js in de hoofdwereld) ---------- */
    /* Vraagt een venster een API-call te doen met het token dat de pagina zelf gebruikt.
     * Antwoord: { status, json?, error?, noToken? }. Alleen lezen (GET, of POST naar $batch). */
    api(url, { win = window, prefer = "", timeout = 20000, method = "GET", body = null } = {}) {
      return new Promise((resolve) => {
        const id = Math.random().toString(36).slice(2);
        const h = (e) => { if (e.data?.type === "tt:api:result" && e.data.id === id) { window.removeEventListener("message", h); resolve(e.data); } };
        window.addEventListener("message", h);
        try { win.postMessage({ type: "tt:api:get", id, url, prefer, method, body }, "*"); } catch { resolve({ status: 0, error: "frame onbereikbaar" }); }
        setTimeout(() => { window.removeEventListener("message", h); resolve({ status: 0, error: "time-out" }); }, timeout);
      });
    },

    /* Welke frames hebben welke tokens (alleen host, scopes en vervaltijd; nooit het token zelf). */
    tokensInfo(wait = 700) {
      return new Promise((resolve) => {
        const out = [];
        const h = (e) => { if (e.data?.type === "tt:tokens:data") out.push({ frame: e.data.frame, tokens: e.data.tokens || [] }); };
        window.addEventListener("message", h);
        for (const w of TT.frames()) { try { w.postMessage({ type: "tt:tokens:info" }, "*"); } catch {} }
        setTimeout(() => { window.removeEventListener("message", h); resolve(out); }, wait);
      });
    },
    netlog(wait = 700) {
      return new Promise((resolve) => {
        const out = [];
        const h = (e) => { if (e.data?.type === "tt:netlog:data") out.push({ frame: e.data.frame, log: e.data.log || [] }); };
        window.addEventListener("message", h);
        for (const w of TT.frames()) { try { w.postMessage({ type: "tt:netlog:get" }, "*"); } catch {} }
        setTimeout(() => { window.removeEventListener("message", h); resolve(out); }, wait);
      });
    },

    /* Venster dat een Graph-token heeft (bij voorkeur met de gevraagde scopes). Onthouden per sessie. */
    _graphWin: null,
    async graphWindow(prefer = "EduAssignments|EduRoster") {
      if (TT._graphWin) return TT._graphWin;
      const ws = TT.frames();
      // Eerst dit venster zelf proberen: één snelle call.
      const probe = async (w) => { const r = await TT.api(`${GRAPH}/me?$select=id`, { win: w, prefer, timeout: 6000 }); return r.status === 200; };
      for (const w of [window, ...ws.filter((x) => x !== window)]) { if (await probe(w)) { TT._graphWin = w; return w; } }
      return null;
    },

    /* Graph-aanroep met paginering. Geeft { status, value, json, error }. */
    async graph(path, { prefer = "EduAssignments|EduRoster", pages = true, win = null, method = "GET", body = null } = {}) {
      const w = win || (await TT.graphWindow(prefer));
      if (!w) return { status: 0, error: "geen Graph-token gevonden in deze Teams-pagina", noToken: true };
      let url = /^https?:/.test(path) ? path : GRAPH + path;
      const value = []; let json = null; let status = 0; let guard = 0;
      while (url && guard++ < 50) {
        let r = await TT.api(url, { win: w, prefer, method, body });
        // Graph beperkt het tempo (429) of is even weg (503): wachten en opnieuw.
        for (let poging = 1; (r.status === 429 || r.status === 503 || r.status === 504) && poging <= 3; poging++) {
          await new Promise((res) => setTimeout(res, 2000 * poging));
          r = await TT.api(url, { win: w, prefer, method, body });
        }
        status = r.status;
        if (r.status !== 200) return { status: r.status, error: r.error || r.json?.error?.message || `HTTP ${r.status}`, json: r.json, code: r.json?.error?.code };
        json = r.json;
        if (Array.isArray(json?.value)) value.push(...json.value);
        url = pages && json?.["@odata.nextLink"] ? json["@odata.nextLink"] : null;
      }
      return { status, value, json };
    },

    /* Meerdere GET's in één keer via $batch (max 20 per batch). Geeft een Map id → { status, body }. */
    async graphBatch(requests, { prefer = "EduAssignments|EduRoster" } = {}) {
      const out = new Map();
      const w = await TT.graphWindow(prefer);
      if (!w) return out;
      for (let i = 0; i < requests.length; i += 20) {
        const part = requests.slice(i, i + 20).map((r, j) => ({ id: String(i + j), method: "GET", url: r.url, headers: { Accept: "application/json" } }));
        const r = await TT.api(`${GRAPH}/$batch`, { win: w, prefer, method: "POST", body: { requests: part } });
        for (const resp of r.json?.responses || []) out.set(requests[+resp.id]?.key ?? resp.id, { status: resp.status, body: resp.body });
      }
      return out;
    },

    /* Beperkte gelijktijdigheid voor veel kleine calls. */
    async pool(items, limit, fn, onProgress) {
      const results = new Array(items.length); let next = 0, done = 0;
      const worker = async () => { while (next < items.length) { const i = next++; results[i] = await fn(items[i], i); done++; onProgress && onProgress(done, items.length); } };
      await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
      return results;
    },

    /* Reageert op DOM-wijzigingen (met periodieke scan, voor shadow DOM). */
    observe(callback, { debounce = 300, poll = 2000 } = {}) {
      let timer = null;
      if (!document.body) return () => {};
      const run = () => { clearTimeout(timer); timer = setTimeout(callback, debounce); };
      const obs = new MutationObserver(run);
      obs.observe(document.body, { childList: true, subtree: true });
      const iv = poll ? setInterval(run, poll) : null;
      callback();
      return () => { obs.disconnect(); clearInterval(iv); };
    },
    onRoute(callback) {
      let last = location.href;
      const check = () => { if (location.href !== last) { last = location.href; callback(location.href); } };
      const push = history.pushState, replace = history.replaceState;
      history.pushState = function () { push.apply(this, arguments); check(); };
      history.replaceState = function () { replace.apply(this, arguments); check(); };
      window.addEventListener("popstate", check);
      callback(location.href);
    },
    deepQueryAll(selector, root = document) {
      const out = [];
      const walk = (r) => {
        try { out.push(...r.querySelectorAll(selector)); } catch {}
        r.querySelectorAll("*").forEach((e) => { if (e.shadowRoot) walk(e.shadowRoot); });
      };
      walk(root);
      return out;
    },
    deepQuery(selector, root = document) { return TT.deepQueryAll(selector, root)[0] || null; },

    /* Excel-export: entries = [[bladnaam, rijen(2D)]]. */
    toXlsx(entries, basis = "teams") {
      if (typeof XLSX === "undefined") return TT.toast("Excel-bibliotheek niet geladen", "err");
      const wb = XLSX.utils.book_new(); const used = new Set();
      const coerce = (v) => { if (v instanceof Date) return v; if (typeof v === "number") return v; const s = String(v ?? ""); if (/^-?\d+([.,]\d+)?$/.test(s) && s.length < 12) return parseFloat(s.replace(",", ".")); return s; };
      for (const [name, rows] of entries) {
        let sheet = String(name).replace(/[\\/?*[\]:]/g, " ").slice(0, 28) || "Blad"; let n = 1;
        while (used.has(sheet)) sheet = sheet.slice(0, 25) + " " + (++n);
        used.add(sheet);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows.map((r) => r.map(coerce))), sheet);
      }
      XLSX.writeFile(wb, `${basis}-${new Date().toISOString().slice(0, 10)}.xlsx`);
    }
  };

  function deepMerge(base, patch) {
    for (const [k, v] of Object.entries(patch || {})) {
      if (v && typeof v === "object" && !Array.isArray(v) && base[k] && typeof base[k] === "object" && !Array.isArray(base[k])) deepMerge(base[k], v);
      else base[k] = v;
    }
    return base;
  }

  window.TT = TT;
})();
