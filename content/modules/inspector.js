/* Inspector (ontwikkelaarsmodus): welke frames heeft deze Teams-pagina, welke tokens (host, scopes,
 * resterende minuten; nooit het token zelf) en welke API-calls Teams doet (patroon, status, duur,
 * veldnamen). Het rapport is geanonimiseerd en te kopiëren om samen een nieuwe analyse te bouwen. */
TT.register({
  id: "inspector",
  _overlay: null,
  init(TT) { if (!TT.isTop) return; },
  destroy() { this._overlay?.remove(); this._overlay = null; },

  async open(TT) {
    if (!this._overlay) {
      this._overlay = TT.el("div", { id: "tt-overlay", role: "dialog", "aria-label": "Inspector", onclick: (e) => { if (e.target === this._overlay) this._overlay.hidden = true; } });
      document.body.append(this._overlay);
      window.addEventListener("keydown", (e) => { if (e.key === "Escape" && this._overlay && !this._overlay.hidden) this._overlay.hidden = true; }, true);
    }
    this._overlay.hidden = false;
    this._overlay.replaceChildren(TT.el("div", { class: "tt-sheet" }, [TT.el("div", { class: "tt-sheet-head" }, [TT.el("h2", { text: "Inspector" })]), TT.el("div", { class: "tt-empty", text: "Frames en logboek ophalen…" })]));
    const [tokens, logs] = await Promise.all([TT.tokensInfo(900), TT.netlog(900)]);
    this.render(TT, tokens, logs);
  },

  render(TT, tokens, logs, filter = "") {
    const q = filter.toLowerCase();
    const tokenRows = tokens.flatMap((f) => f.tokens.map((t) => ({ frame: f.frame, ...t })));
    const calls = logs.flatMap((f) => f.log.map((c) => ({ frame: f.frame, ...c }))).sort((a, b) => b.t - a.t).filter((c) => !q || `${c.url} ${c.host} ${c.frame}`.toLowerCase().includes(q));
    const perHost = {}; for (const c of calls) perHost[c.host] = (perHost[c.host] || 0) + 1;
    const rapport = () => JSON.stringify({ versie: TT.version, pagina: location.host, frames: logs.map((f) => f.frame), tokens: tokenRows.map((t) => ({ frame: t.frame, host: t.host, bron: t.source, minuten: t.minuten, scopes: String(t.scp || "").split(/\s+/).filter(Boolean) })), calls: calls.slice(0, 150).map((c) => ({ frame: c.frame, method: c.method, url: c.url, status: c.status, ms: c.ms, size: c.size, fields: c.fields, error: c.error })) }, null, 1);
    const zoek = TT.el("input", { class: "tt-input", type: "search", placeholder: "Filter op URL, host of frame…", value: filter, oninput: (e) => this.render(TT, tokens, logs, e.target.value) });
    const head = TT.el("div", { class: "tt-sheet-head" }, [
      TT.el("h2", { text: "Inspector" }), zoek,
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Vernieuwen", onclick: () => this.open(TT) }),
      TT.el("button", { class: "tt-btn tt-btn-sm tt-btn-primary", text: "Rapport kopiëren", title: "Geanonimiseerd: geen tokens, geen waarden, id's en e-mailadressen vervangen", onclick: () => TT.copy(rapport()) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: () => { this._overlay.hidden = true; } })
    ]);
    const edu = tokenRows.filter((t) => /EduAssignments|EduRoster/i.test(t.scp));
    const body = TT.el("div", { class: "tt-sheet-body" }, [
      TT.el("div", { class: "tt-cards" }, [["Frames", logs.length], ["Tokens", tokenRows.length], ["Met education-scopes", edu.length], ["API-calls gelogd", calls.length]].map(([l, v]) => TT.el("div", { class: "tt-card" }, [TT.el("b", { text: String(v) }), TT.el("small", { text: l })]))),
      TT.el("h3", { style: "font-size:12px;margin:8px 0", text: "Tokens per frame (alleen host, bron, scopes en resterende tijd)" }),
      tokenRows.length ? TT.el("table", { class: "tt-table" }, [TT.el("thead", {}, TT.el("tr", {}, ["Frame", "Host", "Bron", "Min.", "Scopes"].map((h) => TT.el("th", { text: h })))),
        TT.el("tbody", {}, tokenRows.map((t) => TT.el("tr", {}, [TT.el("td", { text: t.frame }), TT.el("td", { text: t.host }), TT.el("td", { text: t.source }), TT.el("td", { class: "tt-num", text: String(t.minuten) }), TT.el("td", { style: "white-space:normal;max-width:600px", class: /EduAssignments/i.test(t.scp) ? "tt-good" : "", text: String(t.scp || "").split(/\s+/).map((s) => s.replace(/^https?:\/\/[^/]+\//, "")).join(" ") })])))])
        : TT.el("p", { class: "tt-muted", text: "Nog geen tokens gezien. Teams stuurt ze pas mee bij API-calls; klik iets aan in Teams (bijvoorbeeld een team of Opdrachten) en vernieuw." }),
      TT.el("h3", { style: "font-size:12px;margin:14px 0 8px", text: `API-calls (nieuwste eerst) · ${Object.entries(perHost).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([h, n]) => `${h}: ${n}`).join(" · ")}` }),
      TT.el("table", { class: "tt-table" }, [TT.el("thead", {}, TT.el("tr", {}, ["Frame", "Methode", "URL", "Status", "ms", "Grootte", "Velden"].map((h) => TT.el("th", { text: h })))),
        TT.el("tbody", {}, calls.slice(0, 300).map((c) => TT.el("tr", {}, [TT.el("td", { text: c.frame }), TT.el("td", { text: c.method }), TT.el("td", { style: "white-space:normal;max-width:480px;word-break:break-all", text: c.url }), TT.el("td", { class: c.status >= 400 || !c.status ? "tt-bad" : "", text: String(c.status) + (c.error ? " " + c.error : "") }), TT.el("td", { class: "tt-num", text: c.ms == null ? "" : String(c.ms) }), TT.el("td", { class: "tt-num", text: c.size == null ? "" : String(c.size) }), TT.el("td", { style: "white-space:normal;max-width:360px", class: "tt-muted", text: c.fields ? JSON.stringify(c.fields).slice(0, 300) : "" })])))])
    ]);
    this._overlay.replaceChildren(TT.el("div", { class: "tt-sheet" }, [head, body]));
  }
});
