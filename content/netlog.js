/* Netwerkbrug (draait in de pagina zelf, in elk frame van Teams, niet in de extensie-sandbox).
 * 1. Registreert welke API-calls Teams doet: URL-patroon, methode, status, duur en de namen van
 *    velden. Geen waarden, geen inhoud. Zichtbaar in de Inspector.
 * 2. Onthoudt per host het Bearer-token dat de pagina zelf meestuurt (alleen in het geheugen van
 *    deze pagina; het verlaat de pagina nooit). Op verzoek van de extensie doet dit script met dat
 *    token één leesaanroep (GET, of POST naar Graph $batch) en geeft het antwoord terug.
 * 3. Valt terug op de MSAL-cache van Teams in localStorage/sessionStorage als er voor een host nog
 *    geen call is gezien (Teams haalt Graph-tokens vaak al bij het opstarten op).
 */
(() => {
  if (window.__ttNetlog) return;
  const log = []; window.__ttNetlog = log;
  const MAX = 300;
  const INTERESTING = /(^|\.)(microsoft\.com|office\.com|office\.net|onenote\.com|skype\.com|sharepoint\.com|cloud\.microsoft|microsoftonline\.com|live\.com)$/i;
  const ALLOWED_ORIGIN = /^https:\/\/([a-z0-9-]+\.)*(teams\.microsoft\.com|cloud\.microsoft|onenote\.com|microsoft\.com)$/i;
  const anon = (u) => String(u).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, "{guid}").replace(/\d{3,}/g, "{id}").replace(/[A-Za-z0-9_-]{40,}/g, "{token}").replace(/[\w.+-]+@[\w-]+\.[\w.]+/g, "{mail}").slice(0, 400);
  const frameName = () => (window === window.top ? "hoofdpagina" : location.host + anon(location.pathname));
  const hostOf = (u) => { try { return new URL(u, location.href).host; } catch { return ""; } };

  /* ---- JWT uitlezen (alleen aud, scp, exp; voor de status, niet het token) ---- */
  const decodeJwt = (t) => {
    try {
      const p = String(t).split("."); if (p.length < 3) return null;
      let b = p[1].replace(/-/g, "+").replace(/_/g, "/"); b += "=".repeat((4 - (b.length % 4)) % 4);
      const bytes = Uint8Array.from(atob(b), (c) => c.charCodeAt(0));
      return JSON.parse(new TextDecoder().decode(bytes));
    } catch { return null; }
  };

  /* ---- Tokens per host, uit de headers die de pagina zelf meestuurt ---- */
  const tokens = new Map(); // host -> { token, exp, scp, aud, t }
  const noteToken = (url, auth) => {
    try {
      if (!auth || !/^Bearer\s+\S+/i.test(auth)) return;
      const host = hostOf(url); if (!host || !INTERESTING.test(host)) return;
      const token = auth.replace(/^Bearer\s+/i, "").trim();
      const j = decodeJwt(token);
      const exp = j?.exp ? j.exp * 1000 : Date.now() + 30 * 60e3;
      const cur = tokens.get(host);
      // Nieuwer of met meer scopes wint.
      if (!cur || cur.token !== token && (exp >= cur.exp || String(j?.scp || "").length > String(cur.scp || "").length)) tokens.set(host, { token, exp, scp: j?.scp || "", aud: j?.aud || "", t: Date.now() });
    } catch {}
  };

  /* ---- MSAL-cache van Teams (localStorage/sessionStorage) ---- */
  const msalTokens = () => {
    const out = [];
    for (const store of [localStorage, sessionStorage]) {
      try {
        for (let i = 0; i < store.length; i++) {
          const k = store.key(i); if (!k) continue;
          const v = store.getItem(k); if (!v || v[0] !== "{" || !/accesstoken/i.test(k + v.slice(0, 200))) continue;
          let o; try { o = JSON.parse(v); } catch { continue; }
          if (!o || !o.secret || !/^AccessToken/i.test(String(o.credentialType || ""))) continue;
          const exp = Number(o.expiresOn) * 1000;
          if (!exp || exp < Date.now() + 60e3) continue;
          const target = String(o.target || "");
          const hosts = new Set();
          for (const s of target.split(/\s+/)) { const m = /^https?:\/\/([^/]+)\//.exec(s); if (m) hosts.add(m[1].toLowerCase()); }
          if (!hosts.size && /graph/i.test(target)) hosts.add("graph.microsoft.com");
          out.push({ token: o.secret, exp, scp: target, hosts: [...hosts], source: "msal" });
        }
      } catch {}
    }
    return out;
  };

  /* Beste token voor een host; prefer is een regex op de scopes (bv. EduAssignments). */
  const pickToken = (host, prefer) => {
    const re = prefer ? new RegExp(prefer, "i") : null;
    const cands = [];
    const h = tokens.get(host); if (h && h.exp > Date.now()) cands.push({ ...h, source: "header" });
    for (const m of msalTokens()) if (m.hosts.includes(host.toLowerCase())) cands.push(m);
    if (!cands.length) return null;
    cands.sort((a, b) => ((re && re.test(b.scp)) - (re && re.test(a.scp))) || (String(b.scp).length - String(a.scp).length));
    return cands[0];
  };

  /* ---- Logboek van API-calls ---- */
  const shape = (data, depth = 0) => {
    if (depth > 2 || data === null || typeof data !== "object") return typeof data;
    if (Array.isArray(data)) return data.length ? [shape(data[0], depth + 1)] : [];
    const out = {}; let n = 0;
    for (const k of Object.keys(data)) { if (n++ > 60) { out["…"] = "meer velden"; break; } out[k] = shape(data[k], depth + 1); }
    return out;
  };
  const record = (method, url, status, text, ms, error) => {
    const host = hostOf(url); if (!host || !INTERESTING.test(host)) return;
    if (/login\.microsoftonline\.com|login\.live\.com/.test(host)) text = null; // tokens nooit parsen
    if (log.length >= MAX) log.shift();
    let fields = null;
    try { if (text && text.length < 400000 && /^[\[{]/.test(text.trim())) fields = shape(JSON.parse(text)); } catch {}
    log.push({ t: Date.now(), method, url: anon(url), host, status, fields, ms: ms == null ? undefined : Math.round(ms), size: text ? text.length : undefined, error });
  };
  const headerOf = (init, input, name) => {
    try {
      const h = init?.headers || input?.headers; if (!h) return null;
      if (h instanceof Headers) return h.get(name);
      if (Array.isArray(h)) { const p = h.find(([k]) => k.toLowerCase() === name.toLowerCase()); return p ? p[1] : null; }
      for (const k of Object.keys(h)) if (k.toLowerCase() === name.toLowerCase()) return h[k];
    } catch {}
    return null;
  };

  const of = window.fetch;
  window.fetch = async function (input, init) {
    const t0 = performance.now();
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input?.url;
    try { noteToken(url, headerOf(init, input, "authorization")); } catch {}
    let res;
    try { res = await of.apply(this, arguments); }
    catch (err) { try { record((init && init.method) || input?.method || "GET", url, 0, null, performance.now() - t0, String(err)); } catch {} throw err; }
    try {
      const ct = res.headers.get("content-type") || "";
      if (/json/i.test(ct) && res.status !== 204) res.clone().text().then((t) => record((init && init.method) || input?.method || "GET", url, res.status, t, performance.now() - t0)).catch(() => {});
      else record((init && init.method) || input?.method || "GET", url, res.status, null, performance.now() - t0);
    } catch {}
    return res;
  };
  const oo = XMLHttpRequest.prototype.open, os = XMLHttpRequest.prototype.send, osh = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.open = function (m, u) { this.__tt = { m, u: String(u) }; return oo.apply(this, arguments); };
  XMLHttpRequest.prototype.setRequestHeader = function (k, v) { if (this.__tt && /^authorization$/i.test(k)) noteToken(this.__tt.u, v); return osh.apply(this, arguments); };
  XMLHttpRequest.prototype.send = function () {
    const t0 = performance.now();
    this.addEventListener("loadend", () => {
      try {
        if (!this.__tt) return;
        const text = (this.responseType === "" || this.responseType === "text") ? this.responseText : null;
        record(this.__tt.m, this.__tt.u, this.status, text, performance.now() - t0, this.status === 0 ? "mislukt" : undefined);
      } catch {}
    });
    return os.apply(this, arguments);
  };

  /* ---- Verzoeken van de extensie (uit dit of een ander Teams-frame) ---- */
  const trusted = (e) => { try { return e.source === window || ALLOWED_ORIGIN.test(e.origin || ""); } catch { return false; } };
  const reply = (e, msg) => { try { e.source.postMessage(msg, e.origin && e.origin !== "null" ? e.origin : "*"); } catch {} };

  window.addEventListener("message", async (e) => {
    if (!(e.data && e.data.type === "tt:api:get" && trusted(e))) return;
    const { id, url, prefer, method = "GET", body } = e.data;
    try {
      const host = hostOf(url);
      const isBatch = /^https:\/\/graph\.microsoft\.com\/(v1\.0|beta)\/\$batch$/.test(url);
      if (!(method === "GET" || (method === "POST" && isBatch))) throw new Error("alleen lezen toegestaan");
      const tok = pickToken(host, prefer);
      if (!tok) return reply(e, { type: "tt:api:result", id, status: 0, error: `geen token voor ${host} in dit frame`, noToken: true });
      const t0 = performance.now();
      const res = await of(url, { method, headers: { Authorization: `Bearer ${tok.token}`, Accept: "application/json", Prefer: "include-unknown-enum-members", ...(body ? { "Content-Type": "application/json" } : {}) }, body: body ? JSON.stringify(body) : undefined });
      let json = null, text = ""; try { text = await res.clone().text(); json = JSON.parse(text); } catch {}
      try { record(`${method} (Docent Tools)`, url, res.status, text, performance.now() - t0); } catch {}
      reply(e, { type: "tt:api:result", id, status: res.status, json, source: tok.source, scp: tok.scp || "" });
    } catch (err) { reply(e, { type: "tt:api:result", id, status: 0, error: String(err && err.message || err) }); }
  });

  window.addEventListener("message", (e) => {
    if (!(e.data && e.data.type === "tt:tokens:info" && trusted(e))) return;
    const list = [];
    for (const [host, t] of tokens) list.push({ host, source: "header", aud: anon(t.aud), scp: t.scp, minuten: Math.round((t.exp - Date.now()) / 60e3) });
    for (const m of msalTokens()) list.push({ host: m.hosts.join(", ") || "?", source: "msal", scp: m.scp, minuten: Math.round((m.exp - Date.now()) / 60e3) });
    reply(e, { type: "tt:tokens:data", frame: frameName(), origin: location.origin, tokens: list });
  });

  window.addEventListener("message", (e) => {
    if (!(e.data && e.data.type === "tt:netlog:get" && trusted(e))) return;
    reply(e, { type: "tt:netlog:data", frame: frameName(), log });
  });
})();
