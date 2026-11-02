/* Opdrachten over alle teams heen. Haalt via Microsoft Graph (education-API) de klassen van de
 * ingelogde docent op, per klas de opdrachten en leden, per opdracht de inleveringen met cijfers,
 * en toont dat in één overzicht: per klas, per leerling (team-overstijgend), per opdracht,
 * deadlines, nakijkwerk en signalen. Export naar Excel. Gegevens blijven in chrome.storage.local.
 * De berekeningen staan in content/analyse.js. */
TT.register({
  id: "opdrachten",
  _data: null, _overlay: null, _busy: false, _view: "klassen", _sort: {}, _open: new Set(), _bron: null,
  _filter: { periode: null, klassen: null, zoek: "" },

  init(TT) {
    if (!TT.isTop) return;
    this._filter.periode = TT.settings.opdrachten.periode || "schooljaar";
    chrome.storage.local.get("ttData").then(({ ttData }) => { if (ttData?.classes) { this._data = ttData; this.renderWidget(TT); } });
  },
  destroy() { this._overlay?.remove(); this._overlay = null; },

  /* ---------- Ophalen ---------- */
  async laden(TT, voortgang = () => {}) {
    if (this._busy) return;
    this._busy = true;
    try {
      voortgang("Klassen ophalen…", 0);
      const kl = await TT.graph("/education/me/classes?$select=id,displayName,classCode,externalName,description");
      if (kl.status !== 200) throw new Error(kl.noToken ? "Geen Graph-token in deze pagina. Open eerst een team of de app Opdrachten in Teams, en probeer opnieuw." : `Klassen ophalen mislukt (HTTP ${kl.status}${kl.code ? ", " + kl.code : ""}): ${kl.error || ""}`);
      const classes = kl.value.map((c) => ({ id: c.id, displayName: c.displayName, classCode: c.classCode || "", externalName: c.externalName || "" }));
      if (!classes.length) throw new Error("Graph geeft geen klassen terug. Ben je eigenaar van een klasteam (teamtype Klas)?");
      const members = {}, assignments = [];
      let n = 0;
      await TT.pool(classes, 3, async (c) => {
        const [a, m] = await Promise.all([
          TT.graph(`/education/classes/${c.id}/assignments`),
          TT.graph(`/education/classes/${c.id}/members?$select=id,displayName,primaryRole`)
        ]);
        if (a.status === 200) for (const x of a.value) assignments.push({ id: x.id, classId: c.id, displayName: x.displayName, status: x.status, dueDateTime: x.dueDateTime, assignedDateTime: x.assignedDateTime, closeDateTime: x.closeDateTime, createdDateTime: x.createdDateTime, allowLateSubmissions: x.allowLateSubmissions, grading: x.grading ? { "@odata.type": x.grading["@odata.type"], maxPoints: x.grading.maxPoints } : null, webUrl: x.webUrl });
        else c.fout = `opdrachten: HTTP ${a.status}`;
        members[c.id] = m.status === 200 ? m.value.filter((u) => (u.primaryRole || "student") === "student").map((u) => ({ id: u.id, displayName: u.displayName, primaryRole: u.primaryRole })) : [];
        voortgang(`Opdrachten van ${++n}/${classes.length} klassen…`, 0.1 + 0.2 * (n / classes.length));
      });
      const actief = assignments.filter((a) => ["published", "assigned", "inactive"].includes(a.status));
      const submissions = {}; let k = 0;
      await TT.pool(actief, 4, async (a) => {
        const s = await TT.graph(`/education/classes/${a.classId}/assignments/${a.id}/submissions?$expand=outcomes`);
        submissions[a.id] = s.status === 200 ? s.value.map((x) => ({ id: x.id, status: x.status, submittedDateTime: x.submittedDateTime, returnedDateTime: x.returnedDateTime, reassignedDateTime: x.reassignedDateTime, excusedDateTime: x.excusedDateTime, recipient: { userId: x.recipient?.userId || x.recipient?.id || null },
          outcomes: (x.outcomes || []).filter((o) => /pointsOutcome/i.test(o["@odata.type"] || "")).map((o) => ({ "@odata.type": o["@odata.type"], points: o.points ? { points: o.points.points } : null, publishedPoints: o.publishedPoints ? { points: o.publishedPoints.points } : null })) })) : [];
        voortgang(`Inleveringen van ${++k}/${actief.length} opdrachten…`, 0.3 + 0.6 * (k / actief.length));
      });
      // Namen van leerlingen die niet (meer) lid zijn: via $batch op /users.
      const bekend = new Set(Object.values(members).flat().map((u) => u.id));
      const onbekend = [...new Set(Object.values(submissions).flat().map((s) => s.recipient.userId).filter((u) => u && !bekend.has(u)))];
      const namen = {};
      if (onbekend.length) {
        voortgang(`Namen ophalen (${onbekend.length})…`, 0.92);
        const res = await TT.graphBatch(onbekend.map((id) => ({ key: id, url: `/users/${id}?$select=id,displayName` })), { prefer: "User.Read" });
        for (const [id, r] of res) if (r.status === 200 && r.body?.displayName) namen[id] = r.body.displayName;
      }
      const data = { at: Date.now(), classes, members, assignments, submissions, namen, fouten: classes.filter((c) => c.fout).map((c) => `${c.displayName}: ${c.fout}`) };
      this._data = data;
      try { await chrome.storage.local.set({ ttData: data }); } catch (e) { console.warn("[TT] opslaan mislukt", e); }
      voortgang("Klaar", 1);
      this.renderWidget(TT);
      return data;
    } finally { this._busy = false; }
  },

  analyse(TT) {
    const o = TT.settings.opdrachten;
    return TT_ANALYSE.bouw(this._data, { periode: this._filter.periode, klassen: this._filter.klassen, drempelOntbreekt: o.drempelOntbreekt, drempelTeLaat: o.drempelTeLaat });
  },

  /* ---------- Widget in het zijpaneel ---------- */
  renderWidget(TT) {
    const host = document.getElementById("tt-opdrachten-widget"); if (!host) return;
    if (!this._data) return host.replaceChildren(TT.el("p", { class: "tt-muted", text: "Nog niets opgehaald. Klik op Overzicht openen; de eerste keer duurt het even (één aanroep per opdracht)." }));
    const r = this.analyse(TT);
    host.replaceChildren(
      TT.el("div", { class: "tt-stat" }, [["Klassen", r.totaal.klassen], ["Opdrachten", r.totaal.opdrachten], ["Leerlingen", r.totaal.leerlingen], ["Na te kijken", r.totaal.nakijken], ["Ontbreekt", r.totaal.ontbreekt], ["Signalen", r.signalen.length]].map(([l, v]) => TT.el("div", {}, [TT.el("b", { text: String(v) }), TT.el("small", { text: l })]))),
      TT.el("div", { class: "tt-muted", style: "margin-top:6px", text: `Opgehaald ${TT.fmtDateTime(this._data.at)} · periode: ${this.periodeLabel()}` })
    );
  },
  periodeLabel() { return { schooljaar: "dit schooljaar", 30: "laatste 30 dagen", 90: "laatste 90 dagen", alles: "alles" }[this._filter.periode] || "dit schooljaar"; },

  /* ---------- Overzicht ---------- */
  async open(TT, { vernieuw = false } = {}) {
    if (!this._overlay) {
      this._overlay = TT.el("div", { id: "tt-overlay", role: "dialog", "aria-label": "Opdrachten over alle teams", onclick: (e) => { if (e.target === this._overlay) this.close(); } });
      document.body.append(this._overlay);
      this._esc = (e) => { if (e.key === "Escape" && !this._overlay.hidden) { e.stopPropagation(); this.close(); } };
      window.addEventListener("keydown", this._esc, true);
    }
    this._overlay.hidden = false;
    const stale = this._data && Date.now() - this._data.at > (TT.settings.opdrachten.cacheUren || 6) * 3600e3;
    if (!this._data || vernieuw || stale) {
      const prog = TT.el("div", { class: "tt-sheet", style: "max-width:520px;align-self:center;height:auto" }, [
        TT.el("div", { class: "tt-sheet-head" }, [TT.el("h2", { text: "Opdrachten ophalen" }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: () => this.close() })]),
        TT.el("div", { style: "padding:18px" }, [TT.el("p", { id: "tt-prog-txt", text: stale && !vernieuw ? "Gegevens zijn ouder dan de cache-tijd, vernieuwen…" : "Verbinden met Microsoft Graph…" }), TT.el("div", { class: "tt-progress" }, TT.el("i")), TT.el("p", { class: "tt-muted", style: "margin-top:10px", text: "Blijft binnen je eigen Teams-sessie; er gaat niets naar een eigen server." })])
      ]);
      this._overlay.replaceChildren(prog);
      try {
        await this.laden(TT, (txt, f) => { const t = prog.querySelector("#tt-prog-txt"), b = prog.querySelector(".tt-progress i"); if (t) t.textContent = txt; if (b) b.style.width = Math.round(f * 100) + "%"; });
      } catch (e) {
        if (!this._data) { prog.querySelector("#tt-prog-txt").textContent = `Mislukt: ${e.message || e}`; prog.querySelector(".tt-progress").remove(); prog.querySelector("div[style]").append(TT.el("div", { class: "tt-row", style: "margin-top:10px" }, [TT.el("button", { class: "tt-btn tt-btn-primary", text: "Opnieuw proberen", onclick: () => this.open(TT, { vernieuw: true }) }), TT.el("button", { class: "tt-btn", text: "Inspector", onclick: () => { this.close(); TT.module("inspector")?.open(TT); } })])); return; }
        TT.toast(`Vernieuwen mislukt, oude gegevens getoond: ${e.message || e}`, "err", 5000);
      }
    }
    this.render(TT);
  },
  close() { if (this._overlay) this._overlay.hidden = true; },

  render(TT) {
    const r = this.analyse(TT);
    const views = [["klassen", "Klassen", r.klassen.length], ["leerlingen", "Leerlingen", r.leerlingen.length], ["opdrachten", "Opdrachten", r.opdrachten.length], ["deadlines", "Deadlines", r.deadlines.length], ["nakijken", "Nakijken", r.nakijken.length], ["signalen", "Signalen", r.signalen.length]];
    const tabs = TT.el("div", { class: "tt-tabs", role: "tablist" }, views.map(([id, l, n]) => TT.el("button", { class: "tt-tab" + (this._view === id ? " is-on" : ""), role: "tab", "aria-selected": String(this._view === id), onclick: () => { this._view = id; this.render(TT); } }, [l, TT.el("span", { class: "tt-n", text: String(n) })])));
    const periode = TT.el("select", { class: "tt-select", "aria-label": "Periode", onchange: (e) => { this._filter.periode = e.target.value; this.render(TT); this.renderWidget(TT); } }, [["schooljaar", "Dit schooljaar"], ["90", "Laatste 90 dagen"], ["30", "Laatste 30 dagen"], ["alles", "Alles"]].map(([v, l]) => TT.el("option", { value: v, text: l })));
    periode.value = this._filter.periode;
    const zoek = TT.el("input", { class: "tt-input", type: "search", placeholder: "Zoek leerling of opdracht…", "aria-label": "Zoeken", value: this._filter.zoek, oninput: (e) => { this._filter.zoek = e.target.value; this.renderBody(TT); } });
    const klassen = TT.el("div", { class: "tt-filters" }, this._data.classes.map((c) => {
      const on = !this._filter.klassen || this._filter.klassen.has(c.id);
      return TT.el("button", { class: "tt-chip" + (on ? "" : " is-off"), "aria-pressed": String(on), title: on ? "Klas verbergen" : "Klas tonen", onclick: (e) => {
        const all = new Set(this._data.classes.map((x) => x.id));
        let s = this._filter.klassen ? new Set(this._filter.klassen) : all;
        if (e.altKey || e.metaKey) s = new Set([c.id]); else if (s.has(c.id)) s.delete(c.id); else s.add(c.id);
        this._filter.klassen = s.size === all.size ? null : s; this.render(TT); this.renderWidget(TT);
      } }, c.displayName);
    }));
    const head = TT.el("div", { class: "tt-sheet-head" }, [
      TT.el("h2", { text: "Opdrachten over alle teams" }), tabs,
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Dit tabblad naar Excel", onclick: () => this.exportTab(TT) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Alles naar Excel", onclick: () => TT.toXlsx(TT_ANALYSE.naarTabellen(this.analyse(TT)), "teams-opdrachten") }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Vernieuwen", title: `Opgehaald ${TT.fmtDateTime(this._data.at)}`, onclick: () => this.open(TT, { vernieuw: true }) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: () => this.close() })
    ]);
    const tools = TT.el("div", { class: "tt-sheet-tools" }, [periode, zoek, klassen, TT.el("span", { class: "tt-muted", style: "margin-left:auto", text: `Opgehaald ${TT.fmtDateTime(this._data.at)}${this._data.fouten?.length ? " · " + this._data.fouten.length + " klas(sen) met fout" : ""}` })]);
    this._body = TT.el("div", { class: "tt-sheet-body" });
    this._overlay.replaceChildren(TT.el("div", { class: "tt-sheet" }, [head, tools, this._body]));
    this.renderBody(TT);
  },

  renderBody(TT) {
    const r = this.analyse(TT);
    const q = this._filter.zoek.trim().toLowerCase();
    const hit = (s) => !q || String(s || "").toLowerCase().includes(q);
    const D = TT.fmtDate, P = TT.pct;
    const bar = (f, cls) => TT.el("span", {}, [TT.el("span", { class: "tt-bar" }, TT.el("i", { style: `width:${Math.round((f || 0) * 100)}%` })), TT.el("span", { class: cls || "", text: P(f) })]);
    const num = (v, cls) => ({ v, text: String(v ?? ""), cls: cls && v > 0 ? cls : "" });
    const cards = (items) => TT.el("div", { class: "tt-cards" }, items.map(([l, v, sub]) => TT.el("div", { class: "tt-card" }, [TT.el("b", { text: String(v) }), TT.el("small", { text: l }), sub ? TT.el("span", { class: "tt-sub", text: sub }) : null])));
    let out;
    if (this._view === "klassen") {
      out = [cards([["Klassen", r.totaal.klassen], ["Opdrachten", r.totaal.opdrachten], ["Leerlingen", r.totaal.leerlingen], ["Ingeleverd", P(r.totaal.pctIngeleverd), `${r.totaal.ingeleverd} van ${r.totaal.toegewezen}`], ["Te laat", P(r.totaal.pctTeLaat), "van wat is ingeleverd"], ["Ontbreekt", r.totaal.ontbreekt], ["Na te kijken", r.totaal.nakijken], ["Gemiddeld", P(r.totaal.gemiddeld), "van de max. punten"]]),
        this.table(TT, "klassen", [
          { k: "naam", l: "Klas", cls: "tt-name", f: (x) => [x.naam, x.code ? TT.el("span", { class: "tt-sub", text: x.code }) : null] },
          { k: "leerlingen", l: "Leerl.", n: 1 }, { k: "opdrachten", l: "Opdr.", n: 1 }, { k: "toegewezen", l: "Toegew.", n: 1 },
          { k: "pctIngeleverd", l: "Ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) },
          { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") },
          { k: "open", l: "Open", n: 1 }, { k: "nakijken", l: "Na te kijken", n: 1, f: (x) => num(x.nakijken, "tt-warnc") }, { k: "nagekeken", l: "Nagekeken", n: 1 },
          { k: "gemiddeld", l: "Gemiddeld", n: 1, f: (x) => P(x.gemiddeld) }
        ], r.klassen.filter((x) => hit(x.naam) || hit(x.code)))];
    } else if (this._view === "leerlingen") {
      out = [this.table(TT, "leerlingen", [
        { k: "naam", l: "Leerling", cls: "tt-name", f: (x) => [x.naam, TT.el("span", { class: "tt-sub", text: x.klassen.join(", ") })] },
        { k: "aantalKlassen", l: "Klassen", n: 1 }, { k: "toegewezen", l: "Toegew.", n: 1 },
        { k: "pctIngeleverd", l: "Ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) },
        { k: "opTijd", l: "Op tijd", n: 1 }, { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") },
        { k: "open", l: "Open", n: 1 }, { k: "nagekeken", l: "Nagekeken", n: 1 }, { k: "gemiddeld", l: "Gemiddeld", n: 1, f: (x) => P(x.gemiddeld) }, { k: "laatste", l: "Laatste inlevering", f: (x) => D(x.laatste) }
      ], r.leerlingen.filter((x) => hit(x.naam) || x.klassen.some(hit)), (x) => this.leerlingDetail(TT, x))];
    } else if (this._view === "opdrachten") {
      out = [this.table(TT, "opdrachten", [
        { k: "naam", l: "Opdracht", cls: "tt-name", f: (x) => [x.webUrl ? TT.el("a", { href: x.webUrl, target: "_blank", rel: "noopener", text: x.naam }) : x.naam, TT.el("span", { class: "tt-sub", text: x.klas })] },
        { k: "due", l: "Deadline", f: (x) => D(x.due) }, { k: "toegewezen", l: "Toegew.", n: 1 },
        { k: "pctIngeleverd", l: "Ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) },
        { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") }, { k: "open", l: "Open", n: 1 },
        { k: "nakijken", l: "Na te kijken", n: 1, f: (x) => num(x.nakijken, "tt-warnc") }, { k: "nagekeken", l: "Nagekeken", n: 1 }, { k: "max", l: "Max", n: 1 }, { k: "gemiddeld", l: "Gemiddeld", n: 1, f: (x) => P(x.gemiddeld) }
      ], r.opdrachten.filter((x) => hit(x.naam) || hit(x.klas)))];
    } else if (this._view === "deadlines") {
      out = [TT.el("p", { class: "tt-muted", text: "Opdrachten met een deadline vanaf gisteren, over alle klassen, op volgorde van datum. Zo zie je waar leerlingen die in meerdere van je klassen zitten tegelijk mee te maken krijgen." }),
        this.table(TT, "deadlines", [
          { k: "due", l: "Deadline", f: (x) => [D(x.due, { weekday: "short", day: "numeric", month: "short" }), TT.el("span", { class: "tt-sub", text: x.dagen <= 0 ? "vandaag" : x.dagen === 1 ? "morgen" : `over ${x.dagen} dagen` })] },
          { k: "naam", l: "Opdracht", cls: "tt-name", f: (x) => [x.webUrl ? TT.el("a", { href: x.webUrl, target: "_blank", rel: "noopener", text: x.naam }) : x.naam, TT.el("span", { class: "tt-sub", text: x.klas })] },
          { k: "toegewezen", l: "Toegew.", n: 1 }, { k: "pctIngeleverd", l: "Al ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) }, { k: "open", l: "Nog open", n: 1 }
        ], r.deadlines.filter((x) => hit(x.naam) || hit(x.klas)))];
    } else if (this._view === "nakijken") {
      out = [TT.el("p", { class: "tt-muted", text: "Ingeleverd maar nog niet teruggegeven, over alle klassen. Langst wachtende bovenaan." }),
        this.table(TT, "nakijken", [
          { k: "naam", l: "Opdracht", cls: "tt-name", f: (x) => [x.webUrl ? TT.el("a", { href: x.webUrl, target: "_blank", rel: "noopener", text: x.naam }) : x.naam, TT.el("span", { class: "tt-sub", text: x.klas })] },
          { k: "nakijken", l: "Wachten", n: 1, f: (x) => num(x.nakijken, "tt-warnc") }, { k: "wachtDagen", l: "Langste wachttijd", n: 1, f: (x) => ({ v: x.wachtDagen, text: `${x.wachtDagen} dagen`, cls: x.wachtDagen >= 14 ? "tt-bad" : x.wachtDagen >= 7 ? "tt-warnc" : "" }) },
          { k: "due", l: "Deadline", f: (x) => D(x.due) }, { k: "nagekeken", l: "Al nagekeken", n: 1 }, { k: "toegewezen", l: "Toegew.", n: 1 }
        ], r.nakijken.filter((x) => hit(x.naam) || hit(x.klas)))];
    } else {
      const o = TT.settings.opdrachten;
      out = [TT.el("p", { class: "tt-muted", text: `Leerlingen met minstens ${o.drempelOntbreekt}× ontbrekend werk of ${o.drempelTeLaat}× te laat, opgeteld over al je klassen. Drempels staan in de instellingen.` }),
        this.table(TT, "signalen", [
          { k: "naam", l: "Leerling", cls: "tt-name", f: (x) => [x.naam, TT.el("span", { class: "tt-sub", text: x.klassen.join(", ") })] },
          { k: "redenen", l: "Signaal", f: (x) => x.redenen.join(", ") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") }, { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") },
          { k: "klassenMetProbleem", l: "In klassen", f: (x) => x.klassenMetProbleem.join(", ") }, { k: "aantalKlassen", l: "Klassen", n: 1 }, { k: "laatste", l: "Laatste inlevering", f: (x) => D(x.laatste) }
        ], r.signalen.filter((x) => hit(x.naam)), (x) => this.leerlingDetail(TT, x))];
    }
    this._body.replaceChildren(...out);
  },

  leerlingDetail(TT, l) {
    const D = TT.fmtDate, P = TT.pct;
    const kleur = { ontbreekt: "tt-bad", ingeleverd: "tt-warnc", nagekeken: "tt-good", teruggegeven: "tt-warnc" };
    return TT.el("div", { class: "tt-detail-grid" }, [
      TT.el("div", {}, [TT.el("h4", { style: "margin:0 0 6px;font-size:12px", text: "Per klas" }), TT.el("table", { class: "tt-mini" }, [
        TT.el("tr", {}, ["Klas", "Toegew.", "Ingel.", "Te laat", "Ontbr.", "Open", "Gem."].map((h) => TT.el("th", { text: h }))),
        ...l.perKlas.map((k) => TT.el("tr", {}, [k.klas, k.toegewezen, k.ingeleverd, k.teLaat, k.ontbreekt, k.open, P(k.gemiddeld)].map((v, i) => TT.el("td", { class: i === 4 && v > 0 ? "tt-bad" : i === 3 && v > 0 ? "tt-warnc" : "", text: String(v ?? "") }))))
      ])]),
      TT.el("div", {}, [TT.el("h4", { style: "margin:0 0 6px;font-size:12px", text: "Opdrachten" }), TT.el("table", { class: "tt-mini" }, [
        TT.el("tr", {}, ["Opdracht", "Klas", "Deadline", "Status", "Ingeleverd", "Punten"].map((h) => TT.el("th", { text: h }))),
        ...l.opdrachten.map((o) => TT.el("tr", {}, [
          TT.el("td", {}, o.webUrl ? TT.el("a", { href: o.webUrl, target: "_blank", rel: "noopener", text: o.opdracht }) : o.opdracht), TT.el("td", { text: o.klas }), TT.el("td", { text: D(o.due) }),
          TT.el("td", { class: kleur[o.klasse] || "", text: o.klasse + (o.teLaat ? " (te laat)" : "") }), TT.el("td", { text: D(o.ingeleverdOp) }), TT.el("td", { text: o.punten != null ? `${o.punten}${o.max ? " / " + o.max : ""}` : "" })
        ]))
      ])])
    ]);
  },

  /* Sorteerbare tabel. cols: { k, l, n (numeriek), cls, f(row) → tekst | node | [nodes] | { v, text, cls } } */
  table(TT, id, cols, rows, detail) {
    if (!rows.length) return TT.el("div", { class: "tt-empty", text: "Niets gevonden voor deze filters." });
    const s = this._sort[id] || { k: cols[0].k, asc: !cols[0].n };
    const val = (x, c) => { const v = x[c.k]; return v == null ? (c.n ? -Infinity : "") : Array.isArray(v) ? v.join(", ") : v; };
    const sorted = [...rows].sort((a, b) => { const c = cols.find((c) => c.k === s.k) || cols[0]; const va = val(a, c), vb = val(b, c); const r = typeof va === "number" || typeof vb === "number" ? (va === -Infinity ? -1 : vb === -Infinity ? 1 : va - vb) : String(va).localeCompare(String(vb), "nl"); return s.asc ? r : -r; });
    const thead = TT.el("tr", {}, cols.map((c) => TT.el("th", { class: (c.n ? "tt-num " : "") + (s.k === c.k ? "is-sort" + (s.asc ? " is-asc" : "") : ""), onclick: () => { this._sort[id] = { k: c.k, asc: s.k === c.k ? !s.asc : !c.n }; this.renderBody(TT); } }, c.l)));
    const body = [];
    for (const x of sorted) {
      const key = id + ":" + (x.id || x.naam);
      const tr = TT.el("tr", { class: detail ? "is-click" : "", onclick: detail ? () => { this._open.has(key) ? this._open.delete(key) : this._open.add(key); this.renderBody(TT); } : null }, cols.map((c) => {
        let v = c.f ? c.f(x) : x[c.k]; let cls = (c.n ? "tt-num " : "") + (c.cls || "");
        if (v && typeof v === "object" && !Array.isArray(v) && !(v instanceof Node)) { cls += " " + (v.cls || ""); v = v.text; }
        return TT.el("td", { class: cls.trim() }, v == null ? "" : Array.isArray(v) ? v : v instanceof Node ? v : String(v));
      }));
      body.push(tr);
      if (detail && this._open.has(key)) body.push(TT.el("tr", { class: "tt-detail" }, TT.el("td", { colspan: String(cols.length) }, detail(x))));
    }
    return TT.el("table", { class: "tt-table" }, [TT.el("thead", {}, thead), TT.el("tbody", {}, body)]);
  },

  exportTab(TT) {
    const tabs = TT_ANALYSE.naarTabellen(this.analyse(TT));
    const map = { klassen: "Klassen", leerlingen: "Leerlingen", opdrachten: "Opdrachten", signalen: "Signalen" };
    const r = this.analyse(TT);
    const f = (t) => (t ? new Date(t).toISOString().slice(0, 10) : "");
    const extra = { deadlines: [["Deadline", "Over (dagen)", "Klas", "Opdracht", "Toegewezen", "Ingeleverd", "Open"], ...r.deadlines.map((o) => [f(o.due), o.dagen, o.klas, o.naam, o.toegewezen, o.ingeleverd, o.open])],
      nakijken: [["Klas", "Opdracht", "Wachten", "Langste wachttijd (dagen)", "Deadline", "Nagekeken", "Toegewezen"], ...r.nakijken.map((o) => [o.klas, o.naam, o.nakijken, o.wachtDagen, f(o.due), o.nagekeken, o.toegewezen])] };
    const entry = map[this._view] ? tabs.find((t) => t[0] === map[this._view]) : [this._view[0].toUpperCase() + this._view.slice(1), extra[this._view]];
    TT.toXlsx([entry], `teams-${this._view}`);
  }
});
