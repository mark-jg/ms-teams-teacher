/* Opdrachten over alle teams heen. Haalt via Microsoft Graph (education-API) de klassen van de
 * ingelogde docent op, per klas de opdrachten en leden, per opdracht de inleveringen met cijfers,
 * en toont dat in één overzicht: per klas, per leerling (team-overstijgend), per opdracht,
 * deadlines, nakijkwerk en signalen. Export naar Excel. Gegevens blijven in chrome.storage.local.
 * De berekeningen staan in content/analyse.js. */
TT.register({
  id: "opdrachten",
  _data: null, _overlay: null, _view: "klassen", _sort: {}, _open: new Set(), _bron: null,
  /* auto: klassen automatisch bepalen (eigen, actieve klasteams van dit schooljaar); anders de onthouden keuze. */
  _filter: { periode: null, klassen: null, auto: true, problemen: false, klas: null, zoek: "" },

  init(TT) {
    if (!TT.isTop) return;
    this._filter.periode = TT.settings.opdrachten.periode || "schooljaar";
    chrome.storage.local.get(["ttData", "ttFilter"]).then(({ ttData, ttFilter }) => {
      if (ttFilter) { Object.assign(this._filter, ttFilter, { klassen: ttFilter.klassen ? new Set(ttFilter.klassen) : null, zoek: "" }); if (ttFilter.view) this._view = ttFilter.view; }
      if (ttData?.classes) { this._data = this.migreer(ttData); this.renderWidget(TT); }
    });
  },
  /* Gegevens van een oudere versie bruikbaar maken; ontbrekende velden krijgen een neutrale waarde.
   * Zo hoeft er na een update van de extensie niets opnieuw opgehaald te worden. */
  migreer(d) {
    if (!d || !d.classes) return d;
    d.schema = 3;
    if (!("me" in d)) d.me = null;
    d.volledigAt ??= d.at;
    for (const c of d.classes) { c.isArchived ??= false; c.specialization ??= null; c.eigenaar ??= null; c.createdDateTime ??= null; c.opgehaald ??= true; }
    d.members ??= {}; d.submissions ??= {}; d.namen ??= {}; d.fouten ??= [];
    return d;
  },
  saveFilter() {
    const f = this._filter;
    // Mag nooit een klik blokkeren: zonder extensiecontext gaat de keuze alleen niet mee naar de volgende keer.
    try { chrome.storage.local.set({ ttFilter: { periode: f.periode, klassen: f.klassen ? [...f.klassen] : null, auto: f.auto, problemen: f.problemen, klas: f.klas, view: this._view } }).catch(() => {}); }
    catch { TT.meldDood(); }
  },

  /* ---------- Welke klassen tellen mee ----------
   * Vier signalen uit Graph: klasteam (specialization educationClass), gearchiveerd, eigenaar (jij staat bij
   * de eigenaren) en schooljaar (team aangemaakt sinds juni vóór dit schooljaar, of opdrachten met een datum
   * in dit schooljaar). Standaard tellen alleen eigen, actieve klasteams van dit jaar. */
  schooljaarStart(nu = new Date()) { return new Date(nu.getMonth() >= 7 ? nu.getFullYear() : nu.getFullYear() - 1, 7, 1); },
  huidigJaar(c) {
    const start = this.schooljaarStart(), grens = new Date(start.getFullYear(), 5, 1);
    if (c.createdDateTime && new Date(c.createdDateTime) >= grens) return true;
    // Eén keer per dataset bepalen welke teams dit schooljaar opdrachten hebben.
    if (this._jaarCache?.at !== this._data?.at) {
      const set = new Set();
      for (const a of this._data?.assignments || []) if ([a.dueDateTime, a.assignedDateTime, a.createdDateTime].some((d) => d && new Date(d) >= start)) set.add(a.classId);
      this._jaarCache = { at: this._data?.at, set };
    }
    return this._jaarCache.set.has(c.id);
  },
  klasRedenen(c) {
    const r = [];
    if (c.opgehaald === false) r.push("niet opgehaald");
    if (c.isArchived) r.push("gearchiveerd");
    if (c.specialization && c.specialization !== "educationClass") r.push("geen klasteam");
    if (c.eigenaar === false) r.push("geen eigenaar");
    if (!this.huidigJaar(c)) r.push("vorig schooljaar");
    return r;
  },
  autoKlassen() {
    if (this._autoCache?.at !== this._data?.at) this._autoCache = { at: this._data?.at, set: new Set((this._data?.classes || []).filter((c) => !this.klasRedenen(c).length).map((c) => c.id)) };
    return this._autoCache.set;
  },
  /* Geselecteerde klassen als Set, of null voor alle. */
  selectie() {
    if (!this._data) return null;
    if (this._filter.auto) { const s = this.autoKlassen(); return s.size ? s : null; }
    return this._filter.klassen;
  },
  destroy() { this._overlay?.remove(); this._overlay = null; },

  /* ---------- Ophalen ----------
   * Het Graph-token met EduAssignments zit alleen in het frame van de app Opdrachten en mist EduRoster.
   * Daarom: klassen via /education/me/classes als dat mag, anders via de teams-lijst (/me/joinedTeams);
   * leden via /education/classes/{id}/members als dat mag, anders via groepsleden minus eigenaren.
   * Alles wat per team of per opdracht nodig is gaat gebundeld via $batch (TT.graphMany), en teams
   * waarvan je geen eigenaar bent worden overgeslagen (instelling alleTeams), net als gearchiveerde. */
  _lopend: null,
  laden(TT, voortgang = () => {}, { volledig = false } = {}) {
    // Loopt er al een ronde, dan haakt deze aanroep daarop aan: zelfde voortgang, zelfde resultaat.
    if (this._lopend) { this._lopend.cbs.push(voortgang); if (this._lopend.laatste) { try { voortgang(...this._lopend.laatste); } catch {} } return this._lopend.promise; }
    const job = { cbs: [voortgang], laatste: null };
    const melden = (txt, f) => { job.laatste = [txt, f]; for (const cb of job.cbs) { try { cb(txt, f); } catch {} } };
    job.promise = this._laden(TT, melden, { volledig }).finally(() => { this._lopend = null; });
    this._lopend = job;
    return job.promise;
  },
  async _laden(TT, voortgang, { volledig }) {
    const inst = TT.settings.opdrachten;
    // Incrementeel: oude inleveringen hergebruiken voor opdrachten waarvan de deadline langer dan 45 dagen
    // voorbij is en de opdracht zelf niet is gewijzigd. Eens per volledigDagen toch alles.
    const oud = !volledig && this._data?.classes && Date.now() - (this._data.volledigAt || 0) < (inst.volledigDagen ?? 7) * 864e5 ? this._data : null;
    const oudA = new Map((oud?.assignments || []).map((a) => [a.id, a]));
    let hergebruikt = 0;
    voortgang("Zoeken naar de app Opdrachten in Teams…", 0);
    const w = await TT.graphWindow("EduAssignments", { wait: 90000, onWait: (sec) => voortgang(`Open de app Opdrachten in Teams (linkerbalk). Daarin zit het token dat de extensie nodig heeft. Wachten… ${sec}s`, 0) });
    if (!w) throw new Error("Geen Graph-token met EduAssignments gevonden. Open in Teams de app Opdrachten (linkerbalk, eventueel via ⋯ Meer apps) en probeer opnieuw.");
    voortgang("Klassen ophalen…", 0.02);
    const me = (await TT.graph("/me?$select=id")).json?.id || null;
    let classes = [], bron = "education";
    const kl = await TT.graph("/education/me/classes?$select=id,displayName,classCode,externalName");
    if (kl.status === 200) classes = kl.value.map((c) => ({ id: c.id, displayName: c.displayName, classCode: c.classCode || "", externalName: c.externalName || "" }));
    else {
      bron = "teams";
      const jt = await TT.graph("/me/joinedTeams?$select=id,displayName,isArchived");
      if (jt.status !== 200) throw new Error(`Klassen ophalen mislukt: education-API gaf HTTP ${kl.status}${kl.code ? " (" + kl.code + ")" : ""}, teams-lijst gaf HTTP ${jt.status}${jt.error ? ": " + jt.error : ""}`);
      classes = jt.value.map((t) => ({ id: t.id, displayName: t.displayName, classCode: "", externalName: "", isArchived: !!t.isArchived }));
    }
    if (!classes.length) throw new Error("Geen klassen of teams gevonden.");
    const totaalTeams = classes.length;

    // 1. Teamgegevens (type, archief, aanmaakdatum) en eigenaren, gebundeld.
    voortgang(`Teamgegevens ophalen (${classes.length} teams)…`, 0.04);
    const meta = await TT.graphMany(classes.map((c) => ({ key: c.id, url: `/teams/${c.id}?$select=id,specialization,isArchived,createdDateTime` })));
    for (const c of classes) { const m = meta.get(c.id); if (m?.status === 200 && m.json) { c.specialization = m.json.specialization || null; c.isArchived = !!m.json.isArchived; c.createdDateTime = m.json.createdDateTime || null; } }
    if (bron === "teams") classes = classes.filter((c) => !c.specialization || c.specialization === "educationClass" || c.specialization === "none");
    if (!inst.gearchiveerd) classes = classes.filter((c) => !c.isArchived);
    if (!classes.length) throw new Error("Alleen gearchiveerde of niet-klasteams gevonden. Zet in de instellingen 'ook gearchiveerde teams' aan als je die wilt zien.");
    voortgang(`Eigenaren ophalen (${classes.length} teams)…`, 0.08);
    const own = await TT.graphMany(classes.map((c) => ({ key: c.id, url: `/groups/${c.id}/owners?$select=id` })));
    const ownersVan = new Map();
    for (const c of classes) { const o = own.get(c.id); const set = new Set((o?.value || []).map((u) => u.id)); ownersVan.set(c.id, set); c.eigenaar = o?.status === 200 && me ? set.has(me) : null; }
    // Alleen eigen teams doorzoeken, tenzij de instelling anders zegt of een team handmatig is gekozen.
    const handmatig = this._filter.klassen && !this._filter.auto ? this._filter.klassen : new Set();
    for (const c of classes) c.opgehaald = inst.alleTeams || c.eigenaar !== false || handmatig.has(c.id);
    const doen = classes.filter((c) => c.opgehaald);
    if (!doen.length) throw new Error(`Van de ${classes.length} teams ben je nergens eigenaar. Zet in de instellingen 'ook teams waarvan ik geen eigenaar ben' aan.`);

    // 2. Opdrachten per klas, gebundeld.
    voortgang(`Opdrachten ophalen (${doen.length} teams)…`, 0.12);
    const assignments = [];
    const asg = await TT.graphMany(doen.map((c) => ({ key: c.id, url: `/education/classes/${c.id}/assignments` })), { onProgress: (n, t) => voortgang(`Opdrachten van ${n}/${t} teams…`, 0.12 + 0.1 * (n / t)) });
    for (const c of doen) {
      const a = asg.get(c.id);
      if (a?.status === 200) {
        c.klas = true;
        for (const x of a.value) assignments.push({ id: x.id, classId: c.id, displayName: x.displayName, status: x.status, dueDateTime: x.dueDateTime, assignedDateTime: x.assignedDateTime, closeDateTime: x.closeDateTime, createdDateTime: x.createdDateTime, lastModifiedDateTime: x.lastModifiedDateTime, allowLateSubmissions: x.allowLateSubmissions, grading: x.grading ? { "@odata.type": x.grading["@odata.type"], maxPoints: x.grading.maxPoints } : null, webUrl: x.webUrl });
      } else { c.klas = bron === "education"; c.fout = `opdrachten: HTTP ${a?.status}${a?.code ? " " + a.code : ""}`; }
    }
    // Bij de teams-lijst vallen teams zonder werkende opdrachten-API (geen klasteam) weg.
    const eerder = classes.length;
    classes = classes.filter((c) => !c.opgehaald || c.klas);
    if (!classes.some((c) => c.klas)) throw new Error(bron === "teams" ? "Geen enkel team gaf opdrachten terug. Ben je eigenaar van een klasteam (teamtype Klas)?" : "Graph geeft geen klassen terug.");

    // 3. Leden: education-rooster als dat mag, anders groepsleden minus eigenaren.
    const klassen = classes.filter((c) => c.klas);
    voortgang(`Leden ophalen (${klassen.length} klassen)…`, 0.24);
    const members = {};
    const proef = await TT.graph(`/education/classes/${klassen[0].id}/members?$select=id,displayName,primaryRole`);
    const rosterOk = proef.status === 200;
    const mem = await TT.graphMany(klassen.map((c) => ({ key: c.id, url: rosterOk ? `/education/classes/${c.id}/members?$select=id,displayName,primaryRole` : `/groups/${c.id}/members?$select=id,displayName` })));
    for (const c of klassen) {
      const m = mem.get(c.id);
      if (m?.status !== 200) { members[c.id] = oud?.members?.[c.id] || []; continue; }
      const owners = ownersVan.get(c.id) || new Set();
      members[c.id] = rosterOk
        ? m.value.filter((u) => (u.primaryRole || "student") === "student").map((u) => ({ id: u.id, displayName: u.displayName, primaryRole: u.primaryRole }))
        : m.value.filter((u) => !owners.has(u.id) && /user$/i.test(u["@odata.type"] || "#microsoft.graph.user")).map((u) => ({ id: u.id, displayName: u.displayName, primaryRole: "student" }));
    }

    // 4. Inleveringen, alleen voor recente, actieve opdrachten; oude inleveringen hergebruiken.
    const jaren = Math.min(6, Math.max(1, inst.schooljaren || 2));
    const oudGrens = new Date(this.schooljaarStart().getFullYear() - (jaren - 1), 7, 1);
    const datum = (a) => new Date(a.dueDateTime || a.assignedDateTime || a.createdDateTime || 0);
    const weggelaten = assignments.filter((a) => datum(a) < oudGrens).length;
    const recent = assignments.filter((a) => datum(a) >= oudGrens);
    assignments.length = 0; assignments.push(...recent);
    const actief = assignments.filter((a) => ["published", "assigned", "inactive"].includes(a.status));
    const submissions = {};
    const grens45 = Date.now() - 45 * 864e5;
    const herbruikbaar = (a) => {
      const o = oudA.get(a.id), subs = oud?.submissions?.[a.id];
      if (!o || !subs) return false;
      const due = a.dueDateTime ? new Date(a.dueDateTime).getTime() : null;
      return due != null && due < grens45 && (!o.lastModifiedDateTime || o.lastModifiedDateTime === a.lastModifiedDateTime);
    };
    const ophalen = [];
    for (const a of actief) { if (herbruikbaar(a)) { submissions[a.id] = oud.submissions[a.id]; hergebruikt++; } else ophalen.push(a); }
    voortgang(`Inleveringen ophalen (${ophalen.length} opdrachten${hergebruikt ? `, ${hergebruikt} hergebruikt` : ""})…`, 0.3);
    const subs = await TT.graphMany(ophalen.map((a) => ({ key: a.id, url: `/education/classes/${a.classId}/assignments/${a.id}/submissions?$expand=outcomes` })), { onProgress: (n, t) => voortgang(`Inleveringen van ${n}/${t} opdrachten…${hergebruikt ? ` (${hergebruikt} hergebruikt)` : ""}`, 0.3 + 0.6 * (n / t)) });
    for (const a of ophalen) {
      const r = subs.get(a.id);
      if (r?.status !== 200) { a.fout = `inleveringen: HTTP ${r?.status}`; submissions[a.id] = oud?.submissions?.[a.id] || []; continue; }
      submissions[a.id] = r.value.map((x) => ({ id: x.id, status: x.status, submittedDateTime: x.submittedDateTime, returnedDateTime: x.returnedDateTime, reassignedDateTime: x.reassignedDateTime, excusedDateTime: x.excusedDateTime, recipient: { userId: x.recipient?.userId || x.recipient?.id || null },
        outcomes: (x.outcomes || []).filter((o) => /pointsOutcome/i.test(o["@odata.type"] || "")).map((o) => ({ "@odata.type": o["@odata.type"], points: o.points ? { points: o.points.points } : null, publishedPoints: o.publishedPoints ? { points: o.publishedPoints.points } : null })) }));
    }

    // 5. Namen van leerlingen die niet (meer) lid zijn.
    const bekend = new Set(Object.values(members).flat().map((u) => u.id));
    const onbekend = [...new Set(Object.values(submissions).flat().map((s) => s.recipient.userId).filter((u) => u && !bekend.has(u)))];
    const namen = { ...(oud?.namen || {}) };
    const nogOp = onbekend.filter((u) => !namen[u]);
    if (nogOp.length) {
      voortgang(`Namen ophalen (${nogOp.length})…`, 0.92);
      const res = await TT.graphMany(nogOp.map((id) => ({ key: id, url: `/users/${id}?$select=id,displayName` })), { prefer: "User.Read" });
      for (const [id, r] of res) if (r.status === 200 && r.json?.displayName) namen[id] = r.json.displayName;
    }
    const fouten = [...classes.filter((c) => c.fout).map((c) => `${c.displayName}: ${c.fout}`), ...actief.filter((a) => a.fout).map((a) => `${a.displayName}: ${a.fout}`)];
    const data = { schema: 3, at: Date.now(), volledigAt: oud ? oud.volledigAt || oud.at : Date.now(), hergebruikt, bron, rosterOk, me, weggelaten, totaalTeams, nietKlas: eerder - classes.length,
      classes: classes.map(({ id, displayName, classCode, externalName, isArchived, specialization, eigenaar, createdDateTime, opgehaald }) => ({ id, displayName, classCode, externalName, isArchived: !!isArchived, specialization: specialization || null, eigenaar: eigenaar ?? null, createdDateTime: createdDateTime || null, opgehaald: !!opgehaald })),
      members, assignments, submissions, namen, fouten };
    this._data = data;
    voortgang("Opslaan…", 0.98);
    try { await chrome.storage.local.set({ ttData: data }); }
    catch (e) { console.warn("[TT] opslaan mislukt", e); TT.toast(`Opslaan van de gegevens mislukt (${e?.message || e}); het overzicht werkt wel, maar na herladen moet alles opnieuw worden opgehaald.`, "err", 8000); }
    voortgang("Klaar", 1);
    this.renderWidget(TT);
    return data;
  },

  /* Klassen afgeleid uit lidmaatschap en teamnaam (zie analyse.js), binnen de geselecteerde teams. */
  klassen(TT) {
    const sel = this.selectie();
    const sleutel = JSON.stringify([this._data?.at, sel ? [...sel].sort() : null, TT.settings.opdrachten.klascodePatroon || ""]);
    if (this._klassen?.sleutel !== sleutel) this._klassen = { sleutel, ...TT_ANALYSE.klassen(this._data, { teams: sel, patroon: TT.settings.opdrachten.klascodePatroon || "" }) };
    return this._klassen;
  },
  /* Analyse wordt één keer per combinatie van gegevens en filters berekend; tabwissels zijn dan gratis. */
  analyse(TT) {
    const o = TT.settings.opdrachten;
    const K = this.klassen(TT);
    const sel = this.selectie();
    const sleutel = JSON.stringify([this._data?.at, this._filter.periode, sel ? [...sel].sort() : null, K.sleutel, o.drempelOntbreekt, o.drempelTeLaat]);
    if (this._analyseCache?.sleutel !== sleutel) this._analyseCache = { sleutel, r: TT_ANALYSE.bouw(this._data, { periode: this._filter.periode, klassen: sel, klasVan: K.klasVan, drempelOntbreekt: o.drempelOntbreekt, drempelTeLaat: o.drempelTeLaat }) };
    return this._analyseCache.r;
  },
  /* Per afgeleide klas de totalen van haar leerlingen. */
  klasOverzicht(TT, r) {
    if (this._klasOverzichtCache?.r === r) return this._klasOverzichtCache.uit;
    const uit = this._klasOverzicht(TT, r);
    this._klasOverzichtCache = { r, uit };
    return uit;
  },
  _klasOverzicht(TT, r) {
    const K = this.klassen(TT);
    const perId = new Map(r.leerlingen.map((l) => [l.id, l]));
    return K.klassen.map((k) => {
      const ls = k.leerlingen.map((x) => perId.get(x.id)).filter(Boolean);
      const som = (f) => ls.reduce((a, l) => a + (l[f] || 0), 0);
      const pcts = ls.map((l) => l.gemiddeld).filter((x) => x != null);
      return { id: k.id, naam: k.naam, teams: k.teams.map((t) => t.naam), aantalTeams: k.teams.length, leerlingen: k.leerlingen.length, toegewezen: som("toegewezen"), ingeleverd: som("ingeleverd"), teLaat: som("teLaat"), ontbreekt: som("ontbreekt"), open: som("open"), nakijken: som("nakijken"), nagekeken: som("nagekeken"),
        pctIngeleverd: som("toegewezen") ? som("ingeleverd") / som("toegewezen") : null, gemiddeld: pcts.length ? pcts.reduce((a, b) => a + b, 0) / pcts.length : null, metSignaal: ls.filter((l) => r.signalen.some((x) => x.id === l.id)).length, lijst: k.leerlingen.map((x) => ({ ...x, ...(perId.get(x.id) || {}) })) };
    });
  },

  /* ---------- Widget in het zijpaneel ---------- */
  renderWidget(TT) {
    const host = document.getElementById("tt-opdrachten-widget"); if (!host) return;
    if (!this._data) return host.replaceChildren(TT.el("p", { class: "tt-muted", text: "Nog niets opgehaald. Klik op Overzicht openen; de eerste keer duurt het even." }), TT.el("div", { class: "tt-muted tt-bijwerk", "aria-live": "polite", text: this._lopend?.laatste ? `Bezig: ${this._lopend.laatste[0]}` : "" }));
    const r = this.analyse(TT);
    host.replaceChildren(
      TT.el("div", { class: "tt-stat" }, [["Klassen", this.klassen(TT).klassen.length], ["Teams", r.totaal.klassen], ["Opdrachten", r.totaal.opdrachten], ["Leerlingen", r.totaal.leerlingen], ["Na te kijken", r.totaal.nakijken], ["Signalen", r.signalen.length]].map(([l, v]) => TT.el("div", {}, [TT.el("b", { text: String(v) }), TT.el("small", { text: l })]))),
      TT.el("div", { class: "tt-muted", style: "margin-top:6px", text: `Opgehaald ${TT.fmtDateTime(this._data.at)} · ${this.selectieLabel()} · ${this.periodeLabel()}` }),
      TT.el("div", { class: "tt-muted tt-bijwerk", "aria-live": "polite" })
    );
  },
  selectieLabel() { const s = this.selectie(), n = this._data?.classes.length || 0; return s ? `${s.size} van ${n} klassen` : `alle ${n} klassen`; },
  periodeOpties(TT) {
    const jaren = Math.min(6, Math.max(1, TT.settings.opdrachten.schooljaren || 2));
    const uit = [["schooljaar", "Dit schooljaar"]];
    for (let n = 2; n <= jaren; n++) uit.push([`sj${n}`, `Laatste ${n} schooljaren`]);
    uit.push(["90", "Laatste 90 dagen"], ["30", "Laatste 30 dagen"], ["alles", "Alles wat is opgehaald"]);
    return uit;
  },
  periodeLabel() { const p = this._filter.periode; const m = /^sj(\d)$/.exec(p || ""); return m ? `laatste ${m[1]} schooljaren` : { schooljaar: "dit schooljaar", 30: "laatste 30 dagen", 90: "laatste 90 dagen", alles: "alles" }[p] || "dit schooljaar"; },

  /* ---------- Overzicht ---------- */
  async open(TT, { vernieuw = false, volledig = false } = {}) {
    if (!TT.alive()) TT.meldDood();
    if (!this._overlay) {
      this._overlay = TT.el("div", { id: "tt-overlay", role: "dialog", "aria-label": "Opdrachten over alle teams", onclick: (e) => { if (e.target === this._overlay) this.close(); } });
      document.body.append(this._overlay);
      this._esc = (e) => { if (e.key === "Escape" && !this._overlay.hidden) { e.stopPropagation(); this.close(); } };
      window.addEventListener("keydown", this._esc, true);
    }
    this._overlay.hidden = false;
    const stale = this._data && Date.now() - this._data.at > (TT.settings.opdrachten.cacheUren || 6) * 3600e3;
    if (this._data && !volledig) {
      // Oude gegevens meteen tonen; bijwerken gebeurt op de achtergrond.
      this.render(TT);
      if (vernieuw || stale) this.bijwerken(TT);
      return;
    }
    const prog = TT.el("div", { class: "tt-sheet", style: "max-width:520px;align-self:center;height:auto" }, [
      TT.el("div", { class: "tt-sheet-head" }, [TT.el("h2", { text: "Opdrachten ophalen" }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: () => this.close() })]),
      TT.el("div", { style: "padding:18px" }, [TT.el("p", { id: "tt-prog-txt", text: "Verbinden met Microsoft Graph…" }), TT.el("div", { class: "tt-progress" }, TT.el("i")), TT.el("p", { class: "tt-muted", style: "margin-top:10px", text: "Blijft binnen je eigen Teams-sessie; er gaat niets naar een eigen server." })])
    ]);
    this._overlay.replaceChildren(prog);
    try {
      await this.laden(TT, (txt, f) => { const t = prog.querySelector("#tt-prog-txt"), b = prog.querySelector(".tt-progress i"); if (t) t.textContent = txt; if (b) b.style.width = Math.round(f * 100) + "%"; }, { volledig: true });
    } catch (e) {
      if (!this._data) { prog.querySelector("#tt-prog-txt").textContent = `Mislukt: ${e.message || e}`; prog.querySelector(".tt-progress").remove(); prog.querySelector("div[style]").append(TT.el("div", { class: "tt-row", style: "margin-top:10px" }, [TT.el("button", { class: "tt-btn tt-btn-primary", text: "Opnieuw proberen", onclick: () => this.open(TT, { volledig: true }) }), TT.el("button", { class: "tt-btn", text: "Inspector", onclick: () => { this.close(); TT.module("inspector")?.open(TT); } })])); return; }
      TT.toast(`Ophalen mislukt, oude gegevens getoond: ${e.message || e}`, "err", 5000);
    }
    this.render(TT);
  },
  /* Bijwerken op de achtergrond terwijl het overzicht open blijft; de werkbalk toont de voortgang. */
  async bijwerken(TT, { volledig = false } = {}) {
    const status = (txt) => { const el = this._overlay?.querySelector("#tt-bijwerk-status"); if (el) el.textContent = txt; const w = document.querySelector("#tt-opdrachten-widget .tt-bijwerk"); if (w) w.textContent = txt; };
    try {
      await this.laden(TT, (txt) => status(`Bijwerken: ${txt}`), { volledig });
      status("");
      if (this._overlay && !this._overlay.hidden) this.render(TT);
      TT.toast(volledig ? "Alles opnieuw opgehaald" : `Bijgewerkt${this._data?.hergebruikt ? ` (${this._data.hergebruikt} opdrachten hergebruikt)` : ""}`, "ok");
    } catch (e) { status(""); TT.toast(`Bijwerken mislukt: ${e.message || e}`, "err", 5000); }
  },
  close() { if (this._overlay) this._overlay.hidden = true; },

  render(TT) {
    this._limiet = null;
    const r = this.analyse(TT);
    const K = this.klassen(TT);
    const views = [["groepen", "Klassen", K.klassen.length], ["klassen", "Teams", r.klassen.length], ["leerlingen", "Leerlingen", r.leerlingen.length], ["opdrachten", "Opdrachten", r.opdrachten.length], ["deadlines", "Deadlines", r.deadlines.length], ["nakijken", "Nakijken", r.nakijken.length], ["signalen", "Signalen", r.signalen.length]];
    const tabs = TT.el("div", { class: "tt-tabs", role: "tablist" }, views.map(([id, l, n]) => TT.el("button", { class: "tt-tab" + (this._view === id ? " is-on" : ""), role: "tab", "aria-selected": String(this._view === id), "data-view": id, onclick: () => TT.guard(() => {
      this._view = id; this._limiet = null;
      for (const b of tabs.querySelectorAll(".tt-tab")) { const on = b.dataset.view === id; b.classList.toggle("is-on", on); b.setAttribute("aria-selected", String(on)); }
      this.renderBody(TT); this.saveFilter();
    }, "Tabblad") }, [l, TT.el("span", { class: "tt-n", text: String(n) })])));
    const periode = TT.el("select", { class: "tt-select", "aria-label": "Periode", onchange: (e) => TT.guard(() => { this._filter.periode = e.target.value; this.render(TT); this.renderWidget(TT); this.saveFilter(); }, "Periode") }, this.periodeOpties(TT).map(([v, l]) => TT.el("option", { value: v, text: l })));
    if (!this.periodeOpties(TT).some(([v]) => v === this._filter.periode)) this._filter.periode = "schooljaar";
    periode.value = this._filter.periode;
    const zoek = TT.el("input", { class: "tt-input", type: "search", placeholder: "Zoek leerling of opdracht…", "aria-label": "Zoeken", value: this._filter.zoek, oninput: (e) => TT.guard(() => { this._filter.zoek = e.target.value; this._limiet = null; this.renderBody(TT); }, "Zoeken") });
    const sel = this.selectie(), totaal = this._data.classes.length, nSel = sel ? sel.size : totaal;
    const kiezer = TT.el("button", { class: "tt-btn", id: "tt-kiezer-knop", "aria-haspopup": "dialog", text: `Klassen: ${nSel} van ${totaal} ▾`, onclick: (e) => TT.guard(() => this.kiezer(TT, e.currentTarget), "Klassenkiezer") });
    if (this._filter.klas && !K.klassen.some((k) => k.naam === this._filter.klas)) this._filter.klas = null;
    const klasSel = TT.el("select", { class: "tt-select", "aria-label": "Klas", onchange: (e) => TT.guard(() => { this._filter.klas = e.target.value || null; this.render(TT); this.saveFilter(); }, "Klasfilter") }, [TT.el("option", { value: "", text: "Alle klassen" }), ...K.klassen.map((k) => TT.el("option", { value: k.naam, text: `${k.naam} (${k.leerlingen.length})` }))]);
    klasSel.value = this._filter.klas || "";
    const prob = TT.el("button", { class: "tt-chip tt-chip-btn" + (this._filter.problemen ? "" : " is-off"), "aria-pressed": String(this._filter.problemen), title: "Alleen rijen met ontbrekend werk, te laat of nakijkwerk", text: "Alleen problemen", onclick: () => TT.guard(() => { this._filter.problemen = !this._filter.problemen; this.render(TT); this.saveFilter(); }, "Alleen problemen") });
    const verborgen = this._filter.auto && sel ? this._data.classes.filter((c) => !sel.has(c.id)) : [];
    const redenTel = {}; for (const c of verborgen) for (const r of this.klasRedenen(c)) redenTel[r] = (redenTel[r] || 0) + 1;
    const verborgenTxt = verborgen.length ? `${verborgen.length} verborgen (${Object.entries(redenTel).map(([r, n]) => `${n} ${r}`).join(", ")})` : "";
    const head = TT.el("div", { class: "tt-sheet-head" }, [
      TT.el("h2", { text: "Opdrachten over alle teams" }), tabs,
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Dit tabblad naar Excel", onclick: () => this.exportTab(TT) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Alles naar Excel", onclick: () => this.exportAlles(TT) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Bijwerken", title: `Opgehaald ${TT.fmtDateTime(this._data.at)}. Haalt alleen op wat kan veranderen; Alt+klik haalt alles opnieuw op.`, onclick: (e) => this.bijwerken(TT, { volledig: e.altKey }) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: () => this.close() })
    ]);
    const tools = TT.el("div", { class: "tt-sheet-tools" }, [kiezer, klasSel, periode, prob, zoek, TT.el("span", { id: "tt-bijwerk-status", class: "tt-muted", style: "margin-left:auto", "aria-live": "polite" }), TT.el("span", { class: "tt-muted", text: [verborgenTxt, `opgehaald ${TT.fmtDateTime(this._data.at)}`, this._data.fouten?.length ? `${this._data.fouten.length} met fout` : ""].filter(Boolean).join(" · ") })]);
    this._body = TT.el("div", { class: "tt-sheet-body" });
    this._overlay.replaceChildren(TT.el("div", { class: "tt-sheet" }, [head, tools, this._body]));
    this.renderBody(TT);
  },

  /* Popover met zoekveld en vinkjes per klas; "Automatisch" is de slimme standaard. */
  kiezer(TT, anchor) {
    document.getElementById("tt-kiezer")?.remove();
    const box = TT.el("div", { id: "tt-kiezer", role: "dialog", "aria-label": "Klassen kiezen" });
    const r = anchor.getBoundingClientRect();
    box.style.top = `${r.bottom + 6}px`; box.style.left = `${Math.max(8, Math.min(r.left, innerWidth - 380))}px`;
    const sluit = () => { box.remove(); document.removeEventListener("mousedown", buiten, true); document.removeEventListener("keydown", esc, true); };
    const buiten = (e) => { if (!box.contains(e.target) && e.target !== anchor) sluit(); };
    const esc = (e) => { if (e.key === "Escape") { e.stopPropagation(); sluit(); anchor.focus(); } };
    document.addEventListener("mousedown", buiten, true); document.addEventListener("keydown", esc, true);
    let q = "";
    const lijst = TT.el("div", { class: "tt-kiezer-lijst" });
    const alle = () => new Set(this._data.classes.map((c) => c.id));
    const bouw = () => {
      const sel = this.selectie();
      const kls = [...this._data.classes].sort((a, b) => (this.klasRedenen(a).length - this.klasRedenen(b).length) || a.displayName.localeCompare(b.displayName, "nl"));
      lijst.replaceChildren(...kls.filter((c) => !q || c.displayName.toLowerCase().includes(q)).map((c) => {
        const on = !sel || sel.has(c.id);
        const red = this.klasRedenen(c);
        const cb = TT.el("input", { type: "checkbox", onchange: () => {
          const s = new Set(sel || alle()); cb.checked ? s.add(c.id) : s.delete(c.id);
          TT.guard(() => { this._filter.auto = false; this._filter.klassen = s.size === alle().size ? null : s; this.render(TT); this.renderWidget(TT); bouw(); kop(); this.saveFilter(); }, "Klas kiezen");
        } });
        cb.checked = on;
        const kl = this._klassen?.teamKlas?.[c.id];
        return TT.el("label", { class: "tt-kiezer-rij" + (red.length ? " is-dim" : "") }, [cb, TT.el("span", {}, [c.displayName, red.length || kl ? TT.el("small", { text: [kl ? "klas " + kl : "", ...red].filter(Boolean).join(" · ") }) : null])]);
      }));
      if (!lijst.children.length) lijst.append(TT.el("div", { class: "tt-muted", style: "padding:8px", text: "Geen klas gevonden." }));
    };
    const zoek = TT.el("input", { class: "tt-input", type: "search", placeholder: "Zoek klas…", "aria-label": "Zoek klas", oninput: (e) => { q = e.target.value.trim().toLowerCase(); bouw(); } });
    const knop = (txt, title, fn) => TT.el("button", { class: "tt-btn tt-btn-sm", text: txt, title, onclick: () => TT.guard(() => { fn(); this.render(TT); this.renderWidget(TT); bouw(); kop(); this.saveFilter(); }, "Klassen") });
    const kopEl = TT.el("div", { class: "tt-muted", style: "padding:4px 2px 0" });
    const kop = () => { const sel = this.selectie(); kopEl.textContent = this._filter.auto ? `Automatisch: ${sel ? sel.size : this._data.classes.length} eigen, actieve klasteams van dit schooljaar.` : `Handmatige keuze: ${sel ? sel.size : this._data.classes.length} van ${this._data.classes.length}.`; };
    box.append(zoek, TT.el("div", { class: "tt-row", style: "margin:8px 0 2px" }, [
      knop("Automatisch", "Eigen, actieve klasteams van dit schooljaar", () => { this._filter.auto = true; this._filter.klassen = null; }),
      knop("Alles", "", () => { this._filter.auto = false; this._filter.klassen = null; }),
      knop("Niets", "", () => { this._filter.auto = false; this._filter.klassen = new Set(); })
    ]), kopEl, lijst);
    bouw(); kop();
    document.body.append(box);
    zoek.focus();
  },

  renderBody(TT) {
    try { this._renderBody(TT); }
    catch (e) {
      console.error("[Docent Tools] overzicht", e);
      const details = `Docent Tools v${TT.version} · tabblad ${this._view}\n${e?.stack || e}`;
      this._body?.replaceChildren(TT.el("div", { class: "tt-empty" }, [
        TT.el("p", { text: `Dit tabblad kon niet worden opgebouwd: ${e?.message || e}` }),
        TT.el("div", { class: "tt-row", style: "justify-content:center" }, [
          TT.el("button", { class: "tt-btn tt-btn-sm", text: "Foutdetails kopiëren", onclick: () => TT.copy(details) }),
          TT.el("button", { class: "tt-btn tt-btn-sm", text: "Opnieuw proberen", onclick: () => TT.guard(() => this.render(TT), "Overzicht") })
        ])
      ]));
      if (!TT.alive()) TT.meldDood();
    }
  },
  _renderBody(TT) {
    const r = this.analyse(TT);
    const q = this._filter.zoek.trim().toLowerCase();
    const hit = (s) => !q || String(s || "").toLowerCase().includes(q);
    const P_ = this._filter.problemen;
    const KL = this._filter.klas;
    const inKlas = (x) => !KL || x.klas === KL;
    const teamInKlas = (x) => !KL || K.teamKlas[x.klasId || x.id] === KL;
    const K = this.klassen(TT);
    const D = TT.fmtDate, P = TT.pct;
    const bar = (f, cls) => TT.el("span", {}, [TT.el("span", { class: "tt-bar" }, TT.el("i", { style: `width:${Math.round((f || 0) * 100)}%` })), TT.el("span", { class: cls || "", text: P(f) })]);
    const num = (v, cls) => ({ v, text: String(v ?? ""), cls: cls && v > 0 ? cls : "" });
    const cards = (items) => TT.el("div", { class: "tt-cards" }, items.map(([l, v, sub]) => TT.el("div", { class: "tt-card" }, [TT.el("b", { text: String(v) }), TT.el("small", { text: l }), sub ? TT.el("span", { class: "tt-sub", text: sub }) : null])));
    let out;
    if (this._view === "groepen") {
      const rows = this.klasOverzicht(TT, r).filter((k) => (!KL || k.naam === KL) && (hit(k.naam) || k.teams.some(hit)) && (!P_ || k.ontbreekt > 0 || k.nakijken > 0));
      out = [TT.el("p", { class: "tt-muted", text: `Klassen afgeleid uit wie samen in dezelfde teams zit${TT.settings.opdrachten.klascodePatroon ? " en uit de klascode in de teamnaam" : ""}; elke leerling hoort bij de klas waarvan die de meeste teams volgt. Klik op een rij voor de klaslijst. Klopt een indeling niet, dan helpt een klascodepatroon in de instellingen.` }),
        this.table(TT, "groepen", [
          { k: "naam", l: "Klas", cls: "tt-name", f: (x) => [x.naam, TT.el("span", { class: "tt-sub", text: x.teams.join(" · ") })] },
          { k: "leerlingen", l: "Leerl.", n: 1 }, { k: "aantalTeams", l: "Teams", n: 1 }, { k: "toegewezen", l: "Toegew.", n: 1 },
          { k: "pctIngeleverd", l: "Ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) },
          { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") },
          { k: "nakijken", l: "Na te kijken", n: 1, f: (x) => num(x.nakijken, "tt-warnc") }, { k: "metSignaal", l: "Met signaal", n: 1, f: (x) => num(x.metSignaal, "tt-bad") }, { k: "gemiddeld", l: "Gemiddeld", n: 1, f: (x) => P(x.gemiddeld) }
        ], rows, (x) => this.klasDetail(TT, x), { k: "naam", asc: true })];
    } else if (this._view === "klassen") {
      out = [cards([["Teams", r.totaal.klassen], ["Opdrachten", r.totaal.opdrachten], ["Leerlingen", r.totaal.leerlingen], ["Ingeleverd", P(r.totaal.pctIngeleverd), `${r.totaal.ingeleverd} van ${r.totaal.toegewezen}`], ["Te laat", P(r.totaal.pctTeLaat), "van wat is ingeleverd"], ["Ontbreekt", r.totaal.ontbreekt], ["Na te kijken", r.totaal.nakijken], ["Gemiddeld", P(r.totaal.gemiddeld), "van de max. punten"]]),
        this.table(TT, "klassen", [
          { k: "naam", l: "Team", cls: "tt-name", f: (x) => [x.naam, [x.code, K.teamKlas[x.id] ? "klas " + K.teamKlas[x.id] : ""].filter(Boolean).length ? TT.el("span", { class: "tt-sub", text: [x.code, K.teamKlas[x.id] ? "klas " + K.teamKlas[x.id] : ""].filter(Boolean).join(" · ") }) : null] },
          { k: "leerlingen", l: "Leerl.", n: 1 }, { k: "opdrachten", l: "Opdr.", n: 1 }, { k: "toegewezen", l: "Toegew.", n: 1 },
          { k: "pctIngeleverd", l: "Ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) },
          { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") },
          { k: "open", l: "Open", n: 1 }, { k: "nakijken", l: "Na te kijken", n: 1, f: (x) => num(x.nakijken, "tt-warnc") }, { k: "nagekeken", l: "Nagekeken", n: 1 },
          { k: "gemiddeld", l: "Gemiddeld", n: 1, f: (x) => P(x.gemiddeld) }
        ], r.klassen.filter((x) => teamInKlas(x) && (hit(x.naam) || hit(x.code)) && (!P_ || x.ontbreekt > 0 || x.nakijken > 0)), null, { k: "naam", asc: true })];
    } else if (this._view === "leerlingen") {
      out = [this.table(TT, "leerlingen", [
        { k: "naam", l: "Leerling", cls: "tt-name", f: (x) => [x.naam, TT.el("span", { class: "tt-sub", text: x.klassen.join(", ") })] },
        { k: "klas", l: "Klas" }, { k: "aantalKlassen", l: "Teams", n: 1 }, { k: "toegewezen", l: "Toegew.", n: 1 },
        { k: "pctIngeleverd", l: "Ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) },
        { k: "opTijd", l: "Op tijd", n: 1 }, { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") },
        { k: "open", l: "Open", n: 1 }, { k: "nagekeken", l: "Nagekeken", n: 1 }, { k: "gemiddeld", l: "Gemiddeld", n: 1, f: (x) => P(x.gemiddeld) }, { k: "laatste", l: "Laatste inlevering", f: (x) => D(x.laatste) }
      ], r.leerlingen.filter((x) => inKlas(x) && (hit(x.naam) || hit(x.klas) || x.klassen.some(hit)) && (!P_ || x.ontbreekt > 0 || x.teLaat > 0)), (x) => this.leerlingDetail(TT, x), { k: "ontbreekt", asc: false })];
    } else if (this._view === "opdrachten") {
      out = [this.table(TT, "opdrachten", [
        { k: "naam", l: "Opdracht", cls: "tt-name", f: (x) => [x.webUrl ? TT.el("a", { href: x.webUrl, target: "_blank", rel: "noopener", text: x.naam }) : x.naam, TT.el("span", { class: "tt-sub", text: x.klas })] },
        { k: "due", l: "Deadline", f: (x) => D(x.due) }, { k: "toegewezen", l: "Toegew.", n: 1 },
        { k: "pctIngeleverd", l: "Ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) },
        { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") }, { k: "open", l: "Open", n: 1 },
        { k: "nakijken", l: "Na te kijken", n: 1, f: (x) => num(x.nakijken, "tt-warnc") }, { k: "nagekeken", l: "Nagekeken", n: 1 }, { k: "max", l: "Max", n: 1 }, { k: "gemiddeld", l: "Gemiddeld", n: 1, f: (x) => P(x.gemiddeld) }
      ], r.opdrachten.filter((x) => teamInKlas(x) && (hit(x.naam) || hit(x.klas)) && (!P_ || x.ontbreekt > 0 || x.nakijken > 0)), null, { k: "due", asc: false })];
    } else if (this._view === "deadlines") {
      out = [TT.el("p", { class: "tt-muted", text: "Opdrachten met een deadline vanaf gisteren, over alle klassen, op volgorde van datum. Zo zie je waar leerlingen die in meerdere van je klassen zitten tegelijk mee te maken krijgen." }),
        this.table(TT, "deadlines", [
          { k: "due", l: "Deadline", f: (x) => [D(x.due, { weekday: "short", day: "numeric", month: "short" }), TT.el("span", { class: "tt-sub", text: x.dagen <= 0 ? "vandaag" : x.dagen === 1 ? "morgen" : `over ${x.dagen} dagen` })] },
          { k: "naam", l: "Opdracht", cls: "tt-name", f: (x) => [x.webUrl ? TT.el("a", { href: x.webUrl, target: "_blank", rel: "noopener", text: x.naam }) : x.naam, TT.el("span", { class: "tt-sub", text: x.klas })] },
          { k: "toegewezen", l: "Toegew.", n: 1 }, { k: "pctIngeleverd", l: "Al ingeleverd", n: 1, f: (x) => bar(x.pctIngeleverd) }, { k: "open", l: "Nog open", n: 1 }
        ], r.deadlines.filter((x) => teamInKlas(x) && (hit(x.naam) || hit(x.klas)) && (!P_ || x.open > 0)), null, { k: "due", asc: true })];
    } else if (this._view === "nakijken") {
      out = [TT.el("p", { class: "tt-muted", text: "Ingeleverd maar nog niet teruggegeven, over alle klassen. Langst wachtende bovenaan." }),
        this.table(TT, "nakijken", [
          { k: "naam", l: "Opdracht", cls: "tt-name", f: (x) => [x.webUrl ? TT.el("a", { href: x.webUrl, target: "_blank", rel: "noopener", text: x.naam }) : x.naam, TT.el("span", { class: "tt-sub", text: x.klas })] },
          { k: "nakijken", l: "Wachten", n: 1, f: (x) => num(x.nakijken, "tt-warnc") }, { k: "wachtDagen", l: "Langste wachttijd", n: 1, f: (x) => ({ v: x.wachtDagen, text: `${x.wachtDagen} dagen`, cls: x.wachtDagen >= 14 ? "tt-bad" : x.wachtDagen >= 7 ? "tt-warnc" : "" }) },
          { k: "due", l: "Deadline", f: (x) => D(x.due) }, { k: "nagekeken", l: "Al nagekeken", n: 1 }, { k: "toegewezen", l: "Toegew.", n: 1 }
        ], r.nakijken.filter((x) => teamInKlas(x) && (hit(x.naam) || hit(x.klas))), null, { k: "wachtDagen", asc: false })];
    } else {
      const o = TT.settings.opdrachten;
      out = [TT.el("p", { class: "tt-muted", text: `Leerlingen met minstens ${o.drempelOntbreekt}× ontbrekend werk of ${o.drempelTeLaat}× te laat, opgeteld over al je klassen. Drempels staan in de instellingen.` }),
        this.table(TT, "signalen", [
          { k: "naam", l: "Leerling", cls: "tt-name", f: (x) => [x.naam, TT.el("span", { class: "tt-sub", text: x.klassen.join(", ") })] },
          { k: "klas", l: "Klas" }, { k: "redenen", l: "Signaal", f: (x) => x.redenen.join(", ") }, { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (x) => num(x.ontbreekt, "tt-bad") }, { k: "teLaat", l: "Te laat", n: 1, f: (x) => num(x.teLaat, "tt-warnc") },
          { k: "klassenMetProbleem", l: "In teams", f: (x) => x.klassenMetProbleem.join(", ") }, { k: "aantalKlassen", l: "Teams", n: 1 }, { k: "laatste", l: "Laatste inlevering", f: (x) => D(x.laatste) }
        ], r.signalen.filter((x) => inKlas(x) && (hit(x.naam) || hit(x.klas))), (x) => this.leerlingDetail(TT, x), { k: "ontbreekt", asc: false })];
    }
    this._body.replaceChildren(...out);
  },

  klasDetail(TT, k) {
    const P = TT.pct, D = TT.fmtDate;
    return TT.el("div", {}, [TT.el("h4", { style: "margin:0 0 6px;font-size:12px", text: `Klaslijst ${k.naam} (${k.lijst.length})` }), TT.el("table", { class: "tt-mini" }, [
      TT.el("tr", {}, ["Leerling", "Teams", "Toegew.", "Ingel.", "Te laat", "Ontbr.", "Open", "Gem.", "Laatste inlevering"].map((h) => TT.el("th", { text: h }))),
      ...k.lijst.map((l) => TT.el("tr", {}, [l.naam, l.teams, l.toegewezen ?? 0, l.ingeleverd ?? 0, l.teLaat ?? 0, l.ontbreekt ?? 0, l.open ?? 0, P(l.gemiddeld), D(l.laatste)].map((v, i) => TT.el("td", { class: i === 5 && v > 0 ? "tt-bad" : i === 4 && v > 0 ? "tt-warnc" : "", text: String(v ?? "") }))))
    ])]);
  },

  leerlingDetail(TT, l) {
    const D = TT.fmtDate, P = TT.pct;
    const kleur = { ontbreekt: "tt-bad", ingeleverd: "tt-warnc", nagekeken: "tt-good", teruggegeven: "tt-warnc" };
    return TT.el("div", { class: "tt-detail-grid" }, [
      TT.el("div", {}, [TT.el("h4", { style: "margin:0 0 6px;font-size:12px", text: "Per team" }), TT.el("table", { class: "tt-mini" }, [
        TT.el("tr", {}, ["Team", "Toegew.", "Ingel.", "Te laat", "Ontbr.", "Open", "Gem."].map((h) => TT.el("th", { text: h }))),
        ...l.perKlas.map((k) => TT.el("tr", {}, [k.klas, k.toegewezen, k.ingeleverd, k.teLaat, k.ontbreekt, k.open, P(k.gemiddeld)].map((v, i) => TT.el("td", { class: i === 4 && v > 0 ? "tt-bad" : i === 3 && v > 0 ? "tt-warnc" : "", text: String(v ?? "") }))))
      ])]),
      TT.el("div", {}, [TT.el("h4", { style: "margin:0 0 6px;font-size:12px", text: "Opdrachten" }), TT.el("table", { class: "tt-mini" }, [
        TT.el("tr", {}, ["Opdracht", "Team", "Deadline", "Status", "Ingeleverd", "Punten"].map((h) => TT.el("th", { text: h }))),
        ...l.opdrachten.map((o) => TT.el("tr", {}, [
          TT.el("td", {}, o.webUrl ? TT.el("a", { href: o.webUrl, target: "_blank", rel: "noopener", text: o.opdracht }) : o.opdracht), TT.el("td", { text: o.klas }), TT.el("td", { text: D(o.due) }),
          TT.el("td", { class: kleur[o.klasse] || "", text: o.klasse + (o.teLaat ? " (te laat)" : "") }), TT.el("td", { text: D(o.ingeleverdOp) }), TT.el("td", { text: o.punten != null ? `${o.punten}${o.max ? " / " + o.max : ""}` : "" })
        ]))
      ])])
    ]);
  },

  /* Sorteerbare tabel. cols: { k, l, n (numeriek), cls, f(row) → tekst | node | [nodes] | { v, text, cls } } */
  _limiet: null,
  table(TT, id, cols, rows, detail, standaard) {
    if (!rows.length) return TT.el("div", { class: "tt-empty", text: this._filter.problemen ? "Geen problemen binnen deze filters." : "Niets gevonden voor deze filters." });
    const s = this._sort[id] || standaard || { k: cols[0].k, asc: !cols[0].n };
    // Grote tabellen in porties: sorteren gebeurt over alles, tonen eerst 200 rijen.
    const limiet = this._limiet ?? 200;
    const totaalRijen = rows.length;
    const val = (x, c) => { const v = x[c.k]; return v == null ? (c.n ? -Infinity : "") : Array.isArray(v) ? v.join(", ") : v; };
    const sorted = [...rows].sort((a, b) => { const c = cols.find((c) => c.k === s.k) || cols[0]; const va = val(a, c), vb = val(b, c); const r = typeof va === "number" || typeof vb === "number" ? (va === -Infinity ? -1 : vb === -Infinity ? 1 : va - vb) : String(va).localeCompare(String(vb), "nl"); return s.asc ? r : -r; });
    const thead = TT.el("tr", {}, cols.map((c) => TT.el("th", { class: (c.n ? "tt-num " : "") + (s.k === c.k ? "is-sort" + (s.asc ? " is-asc" : "") : ""), onclick: () => { this._sort[id] = { k: c.k, asc: s.k === c.k ? !s.asc : !c.n }; this.renderBody(TT); } }, c.l)));
    const body = [];
    for (const x of sorted.slice(0, limiet)) {
      const key = id + ":" + (x.id || x.naam);
      const tr = TT.el("tr", { class: detail ? "is-click" : "", onclick: detail ? () => { this._open.has(key) ? this._open.delete(key) : this._open.add(key); this.renderBody(TT); } : null }, cols.map((c) => {
        let v = c.f ? c.f(x) : x[c.k]; let cls = (c.n ? "tt-num " : "") + (c.cls || "");
        if (v && typeof v === "object" && !Array.isArray(v) && !(v instanceof Node)) { cls += " " + (v.cls || ""); v = v.text; }
        return TT.el("td", { class: cls.trim() }, v == null ? "" : Array.isArray(v) ? v : v instanceof Node ? v : String(v));
      }));
      body.push(tr);
      if (detail && this._open.has(key)) body.push(TT.el("tr", { class: "tt-detail" }, TT.el("td", { colspan: String(cols.length) }, detail(x))));
    }
    const tabel = TT.el("table", { class: "tt-table" }, [TT.el("thead", {}, thead), TT.el("tbody", {}, body)]);
    if (totaalRijen <= limiet) return tabel;
    return TT.el("div", {}, [tabel, TT.el("div", { class: "tt-row", style: "justify-content:center;padding:12px" }, [
      TT.el("span", { class: "tt-muted", text: `${limiet} van ${totaalRijen} rijen` }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Nog 500", onclick: () => { this._limiet = limiet + 500; this.renderBody(TT); } }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: `Alles (${totaalRijen})`, onclick: () => { this._limiet = Infinity; this.renderBody(TT); } })
    ])]);
  },

  klasTabellen(TT, r) {
    const f = (t) => (t ? new Date(t).toISOString().slice(0, 10) : ""), p = (x) => (x == null ? "" : Math.round(x * 100) + "%");
    const ko = this.klasOverzicht(TT, r);
    return [["Klassen", [["Klas", "Teams", "Leerlingen", "Toegewezen", "Ingeleverd", "% ingeleverd", "Te laat", "Ontbreekt", "Open", "Na te kijken", "Met signaal", "Gemiddeld"], ...ko.map((k) => [k.naam, k.teams.join(", "), k.leerlingen, k.toegewezen, k.ingeleverd, p(k.pctIngeleverd), k.teLaat, k.ontbreekt, k.open, k.nakijken, k.metSignaal, p(k.gemiddeld)])]],
      ["Klasindeling", [["Klas", "Leerling", "Teams in klas", "Toegewezen", "Ingeleverd", "Te laat", "Ontbreekt", "Gemiddeld", "Laatste inlevering"], ...ko.flatMap((k) => k.lijst.map((l) => [k.naam, l.naam, l.teams, l.toegewezen ?? 0, l.ingeleverd ?? 0, l.teLaat ?? 0, l.ontbreekt ?? 0, p(l.gemiddeld), f(l.laatste)]))]]];
  },
  exportAlles(TT) {
    const r = this.analyse(TT);
    TT.toXlsx([...this.klasTabellen(TT, r), ...TT_ANALYSE.naarTabellen(r)], "teams-opdrachten");
  },
  exportTab(TT) {
    const r = this.analyse(TT);
    if (this._view === "groepen") return TT.toXlsx(this.klasTabellen(TT, r), "teams-klassen");
    const tabs = TT_ANALYSE.naarTabellen(r);
    const map = { klassen: "Teams", leerlingen: "Leerlingen", opdrachten: "Opdrachten", signalen: "Signalen" };
    const f = (t) => (t ? new Date(t).toISOString().slice(0, 10) : "");
    const extra = { deadlines: [["Deadline", "Over (dagen)", "Klas", "Opdracht", "Toegewezen", "Ingeleverd", "Open"], ...r.deadlines.map((o) => [f(o.due), o.dagen, o.klas, o.naam, o.toegewezen, o.ingeleverd, o.open])],
      nakijken: [["Klas", "Opdracht", "Wachten", "Langste wachttijd (dagen)", "Deadline", "Nagekeken", "Toegewezen"], ...r.nakijken.map((o) => [o.klas, o.naam, o.nakijken, o.wachtDagen, f(o.due), o.nagekeken, o.toegewezen])] };
    const entry = map[this._view] ? tabs.find((t) => t[0] === map[this._view]) : [this._view[0].toUpperCase() + this._view.slice(1), extra[this._view]];
    TT.toXlsx([entry], `teams-${this._view}`);
  }
});
