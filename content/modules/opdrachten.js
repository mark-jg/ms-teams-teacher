/* Opdrachten over alle teams heen. Haalt via Microsoft Graph (education-API) de klassen van de
 * ingelogde docent op, per klas de opdrachten en leden, per opdracht de inleveringen met cijfers,
 * en toont dat in één overzicht: per klas, per leerling (team-overstijgend), per opdracht,
 * deadlines, nakijkwerk en signalen. Export naar Excel. Gegevens blijven in chrome.storage.local.
 * De berekeningen staan in content/analyse.js. */
TT.register({
  id: "opdrachten",
  _data: null, _overlay: null, _sort: {}, _open: new Set(), _bron: null,
  /* auto: klassen automatisch bepalen (eigen, actieve klasteams van dit schooljaar); anders de onthouden keuze. */
  _filter: { periode: null, klassen: null, auto: true, problemen: false, klas: null, zoek: "" },

  init(TT) {
    if (!TT.isTop) return;
    this._filter.periode = TT.settings.opdrachten.periode || "schooljaar";
    chrome.storage.local.get(["ttData", "ttFilter"]).then(({ ttData, ttFilter }) => {
      if (ttFilter) Object.assign(this._filter, { periode: ttFilter.periode || this._filter.periode, klassen: ttFilter.klassen ? new Set(ttFilter.klassen) : null, auto: ttFilter.auto !== false });
      chrome.storage.local.get(["ttSorteer", "ttVerborgen", "ttVerborgenLijst"]).then(({ ttSorteer, ttVerborgen, ttVerborgenLijst }) => { if (ttVerborgenLijst) this._verborgenLijst = ttVerborgenLijst; if (ttSorteer) this._sorteer = ttSorteer; if (ttVerborgen) this._verborgen = { ...this._verborgen, ...ttVerborgen }; this._verborgenVersie++; if (this._data) this.renderWidget(TT); }).catch(() => {});
      if (ttData?.classes) { this._data = this.migreer(ttData); this.renderWidget(TT); }
    });
  },
  /* Gegevens van een oudere versie bruikbaar maken; ontbrekende velden krijgen een neutrale waarde.
   * Zo hoeft er na een update van de extensie niets opnieuw opgehaald te worden. */
  migreer(d) {
    if (!d || !d.classes) return d;
    d.schema ??= 3;
    if (!("me" in d)) d.me = null;
    d.volledigAt ??= d.at;
    for (const c of d.classes) { c.isArchived ??= false; c.specialization ??= null; c.eigenaar ??= null; c.createdDateTime ??= null; c.opgehaald ??= true; }
    d.members ??= {}; d.submissions ??= {}; d.namen ??= {}; d.fouten ??= []; d.categorieen ??= {}; d.rubrics ??= {}; d.profielen ??= {};
    return d;
  },
  saveFilter() {
    const f = this._filter;
    // Mag nooit een klik blokkeren: zonder extensiecontext gaat de keuze alleen niet mee naar de volgende keer.
    try { chrome.storage.local.set({ ttFilter: { periode: f.periode, klassen: f.klassen ? [...f.klassen] : null, auto: f.auto } }).catch(() => {}); }
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
  /* Basiskeuze van teams als Set, of null voor alle: automatisch of handmatig. */
  selectieBasis() {
    if (!this._data) return null;
    if (this._filter.auto) { const s = this.autoKlassen(); return s.size ? s : null; }
    return this._filter.klassen;
  },
  /* Werkelijke selectie: de basiskeuze zonder verborgen teams en zonder de teams van verborgen klassen. */
  selectie() {
    const basis = this.selectieBasis();
    if (!this._data) return null;
    const vt = this._verborgen.teams, vk = this._verborgen.klassen;
    if (!Object.keys(vt).length && !Object.keys(vk).length) return basis;
    const sleutel = JSON.stringify([this._data.at, basis ? [...basis].sort() : null, this._verborgenVersie, this.patroonCache]);
    if (this._selectieCache?.sleutel === sleutel) return this._selectieCache.set;
    let teamKlas = {};
    if (Object.keys(vk).length) teamKlas = TT_ANALYSE.klassen(this._data, { teams: basis, patroon: this.patroonCache || "" }).teamKlas;
    const set = new Set((this._data.classes || []).filter((c) => (!basis || basis.has(c.id)) && !vt[c.id] && !vk[teamKlas[c.id]]).map((c) => c.id));
    this._selectieCache = { sleutel, set };
    return set;
  },

  /* ---------- Verbergen ----------
   * Leerlingen, klassen, vakken (teams) en opdrachten die je uit het overzicht haalt. Ze tellen nergens
   * meer mee (ook niet in de startkaarten en rapporten) tot je ze terugzet. Bewaard in chrome.storage.local. */
  _verborgen: { leerlingen: {}, klassen: {}, teams: {}, opdrachten: {} }, _verborgenVersie: 0, patroonCache: "",
  /* Alleen in één lijst verborgen (telt overal mee): { [lijstId]: { lijstLabel, items: { [id]: label } } } */
  _verborgenLijst: {},
  verborgenAantal() { return Object.values(this._verborgen).reduce((a, m) => a + Object.keys(m).length, 0) + Object.values(this._verborgenLijst).reduce((a, l) => a + Object.keys(l.items || {}).length, 0); },
  bewaarVerborgen() { try { chrome.storage.local.set({ ttVerborgen: this._verborgen, ttVerborgenLijst: this._verborgenLijst }).catch(() => {}); } catch {} },
  verbergInLijst(TT, lijst, lijstLabel, id, label) {
    const l = (this._verborgenLijst[lijst] ||= { lijstLabel, items: {} }); l.lijstLabel = lijstLabel; l.items[id] = label || String(id);
    this.bewaarVerborgen();
    TT.toast(`${label || "Item"} alleen hier verborgen; telt wel mee.`, "ok");
    this.render(TT);
  },
  toonInLijst(TT, lijst, id) {
    const l = this._verborgenLijst[lijst]; if (l) { delete l.items[id]; if (!Object.keys(l.items).length) delete this._verborgenLijst[lijst]; }
    this.bewaarVerborgen(); this.render(TT);
  },
  /* Filtert een lijst op lokaal verborgen items; idOf geeft het id van een item. */
  zonderLokaal(lijst, items, idOf) { const l = this._verborgenLijst[lijst]; if (!l) return items; return items.filter((x) => !l.items[idOf(x)]); },
  /* Keuzemenu bij het kruisje: alleen hier, of overal en niet meetellen. */
  verbergMenu(TT, anchor, v) {
    document.getElementById("tt-verbergmenu")?.remove();
    const box = TT.el("div", { id: "tt-verbergmenu", role: "menu" });
    const r = anchor.getBoundingClientRect();
    box.style.top = `${r.bottom + 4}px`; box.style.left = `${Math.max(8, Math.min(r.right - 260, innerWidth - 270))}px`;
    const sluit = () => { box.remove(); document.removeEventListener("mousedown", buiten, true); document.removeEventListener("keydown", esc, true); };
    const buiten = (e) => { if (!box.contains(e.target)) sluit(); };
    const esc = (e) => { if (e.key === "Escape") { e.stopPropagation(); sluit(); anchor.focus(); } };
    document.addEventListener("mousedown", buiten, true); document.addEventListener("keydown", esc, true);
    const knop = (titel, sub, fn) => TT.el("button", { class: "tt-verbergmenu-knop", role: "menuitem", onclick: () => { sluit(); TT.guard(fn, "Verbergen"); } }, [TT.el("b", { text: titel }), TT.el("small", { text: sub })]);
    box.append(TT.el("div", { class: "tt-muted", style: "padding:4px 8px 6px;font-size:11px", text: v.label }),
      v.lijst ? knop("Alleen in deze lijst verbergen", `Blijft overal meetellen; verdwijnt uit ${v.lijstLabel || "deze lijst"}.`, () => this.verbergInLijst(TT, v.lijst, v.lijstLabel, v.id, v.label)) : null,
      knop("Verbergen en niet meetellen", "Uit het hele overzicht, de startkaarten en de rapporten.", () => this.verberg(TT, v.type, v.id, v.label)));
    document.body.append(box);
    box.querySelector("button")?.focus();
  },
  verberg(TT, type, id, label) {
    this._verborgen[type][id] = label || String(id);
    this._verborgenVersie++; this.bewaarVerborgen();
    TT.toast(`${label || "Item"} verborgen. Terugzetten via de knop Verborgen.`, "ok");
    this.render(TT); this.renderWidget(TT);
  },
  toon(TT, type, id) {
    delete this._verborgen[type][id];
    this._verborgenVersie++; this.bewaarVerborgen();
    this.render(TT); this.renderWidget(TT);
  },
  /* Gegevens zonder verborgen leerlingen en opdrachten; teams en klassen gaan via de selectie. */
  zichtbareData() {
    const d = this._data; if (!d) return d;
    const vl = this._verborgen.leerlingen, vo = this._verborgen.opdrachten;
    if (!Object.keys(vl).length && !Object.keys(vo).length) return d;
    const sleutel = JSON.stringify([d.at, this._verborgenVersie]);
    if (this._zichtbaarCache?.sleutel === sleutel) return this._zichtbaarCache.d;
    const members = {}; for (const [k, ms] of Object.entries(d.members || {})) members[k] = ms.filter((m) => !vl[m.id]);
    const assignments = (d.assignments || []).filter((a) => !vo[a.id]);
    const submissions = {}; for (const a of assignments) submissions[a.id] = (d.submissions?.[a.id] || []).filter((x) => !vl[x.recipient?.userId]);
    const uit = { ...d, members, assignments, submissions };
    this._zichtbaarCache = { sleutel, d: uit };
    return uit;
  },
  /* Beheer: alles wat verborgen is, per soort, met terugzetten. */
  verborgenBeheer(TT, anchor) {
    document.getElementById("tt-kiezer")?.remove();
    const box = TT.el("div", { id: "tt-kiezer", role: "dialog", "aria-label": "Verborgen items" });
    const r = anchor.getBoundingClientRect();
    box.style.top = `${r.bottom + 6}px`; box.style.left = `${Math.max(8, Math.min(r.left, innerWidth - 380))}px`;
    const sluit = () => { box.remove(); document.removeEventListener("mousedown", buiten, true); document.removeEventListener("keydown", esc, true); };
    const buiten = (e) => { if (!box.contains(e.target) && e.target !== anchor) sluit(); };
    const esc = (e) => { if (e.key === "Escape") { e.stopPropagation(); sluit(); anchor.focus(); } };
    document.addEventListener("mousedown", buiten, true); document.addEventListener("keydown", esc, true);
    const soorten = [["klassen", "Klassen"], ["teams", "Vakken (teams)"], ["leerlingen", "Leerlingen"], ["opdrachten", "Opdrachten"]];
    const lijst = TT.el("div", { class: "tt-kiezer-lijst" });
    const kopje = (t) => TT.el("div", { class: "tt-muted", style: "padding:6px 2px 2px;font-size:10px;text-transform:uppercase;letter-spacing:.04em", text: t });
    const teken = () => {
      const globaal = soorten.filter(([t]) => Object.keys(this._verborgen[t]).length).flatMap(([t, l]) => [kopje(l),
        ...Object.entries(this._verborgen[t]).sort((a, b) => String(a[1]).localeCompare(String(b[1]), "nl")).map(([id, label]) => TT.el("div", { class: "tt-kiezer-rij", style: "justify-content:space-between" }, [TT.el("span", { text: String(label) }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Terugzetten", onclick: () => { this.toon(TT, t, id); teken(); } })]))]);
      const lokaal = Object.entries(this._verborgenLijst).flatMap(([lijstId, l]) => [kopje(l.lijstLabel || lijstId),
        ...Object.entries(l.items || {}).sort((a, b) => String(a[1]).localeCompare(String(b[1]), "nl")).map(([id, label]) => TT.el("div", { class: "tt-kiezer-rij", style: "justify-content:space-between" }, [TT.el("span", { text: String(label) }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Terugzetten", onclick: () => { this.toonInLijst(TT, lijstId, id); teken(); } })]))]);
      lijst.replaceChildren(
        TT.el("div", { class: "tt-verborgen-kop", text: "Verborgen en telt niet mee" }), ...(globaal.length ? globaal : [TT.el("div", { class: "tt-muted", style: "padding:4px 8px 8px", text: "Niets." })]),
        TT.el("div", { class: "tt-verborgen-kop", style: "margin-top:10px", text: "Alleen in een lijst verborgen (telt mee)" }), ...(lokaal.length ? lokaal : [TT.el("div", { class: "tt-muted", style: "padding:4px 8px 8px", text: "Niets." })])
      );
    };
    box.append(TT.el("div", { class: "tt-row", style: "justify-content:space-between" }, [TT.el("b", { text: "Verborgen" }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Alles terugzetten", onclick: () => { for (const [t] of soorten) this._verborgen[t] = {}; this._verborgenLijst = {}; this._verborgenVersie++; this.bewaarVerborgen(); this.render(TT); this.renderWidget(TT); teken(); } })]),
      TT.el("p", { class: "tt-muted", style: "margin:4px 0 6px", text: "Het kruisje op een rij of kaart geeft twee keuzes: alleen in die lijst verbergen, of overal en niet meetellen." }), lijst);
    teken();
    document.body.append(box);
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
    // Gegevens van vóór schema 4 missen categorieën, rubrics en rubric-niveaus: dan één keer alles.
    const oud = !volledig && this._data?.classes && (this._data.schema || 0) >= 4 && Date.now() - (this._data.volledigAt || 0) < (inst.volledigDagen ?? 7) * 864e5 ? this._data : null;
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
    // Met de cijfercategorie erbij ($expand); lukt dat voor een klas niet, dan zonder.
    const asg = await TT.graphMany(doen.map((c) => ({ key: c.id, url: `/education/classes/${c.id}/assignments?$expand=gradingCategory` })), { onProgress: (n, t) => voortgang(`Opdrachten van ${n}/${t} teams…`, 0.12 + 0.08 * (n / t)) });
    const zonder = doen.filter((c) => asg.get(c.id)?.status === 400);
    if (zonder.length) for (const [k, v] of await TT.graphMany(zonder.map((c) => ({ key: c.id, url: `/education/classes/${c.id}/assignments` })))) asg.set(k, v);
    for (const c of doen) {
      const a = asg.get(c.id);
      if (a?.status === 200) {
        c.klas = true;
        for (const x of a.value) assignments.push({ id: x.id, classId: c.id, displayName: x.displayName, status: x.status, dueDateTime: x.dueDateTime, assignedDateTime: x.assignedDateTime, closeDateTime: x.closeDateTime, createdDateTime: x.createdDateTime, lastModifiedDateTime: x.lastModifiedDateTime, allowLateSubmissions: x.allowLateSubmissions, grading: x.grading ? { "@odata.type": x.grading["@odata.type"], maxPoints: x.grading.maxPoints } : null, webUrl: x.webUrl,
          categorie: x.gradingCategory ? { id: x.gradingCategory.id, naam: x.gradingCategory.displayName, gewicht: x.gradingCategory.percentageWeight ?? null } : null });
      } else { c.klas = bron === "education"; c.fout = `opdrachten: HTTP ${a?.status}${a?.code ? " " + a.code : ""}`; }
    }
    // Cijfercategorieën met weging per klas (instellingen van de klas).
    voortgang("Cijfercategorieën ophalen…", 0.21);
    const categorieen = {};
    const cat = await TT.graphMany(doen.filter((c) => c.klas).map((c) => ({ key: c.id, url: `/education/classes/${c.id}/assignmentSettings?$expand=gradingCategories` })));
    for (const c of doen) { const r = cat.get(c.id); categorieen[c.id] = r?.status === 200 ? (r.json?.gradingCategories || []).map((g) => ({ id: g.id, naam: g.displayName, gewicht: g.percentageWeight ?? null })) : (oud?.categorieen?.[c.id] || []); }
    // Bij de teams-lijst vallen teams zonder werkende opdrachten-API (geen klasteam) weg.
    const eerder = classes.length;
    classes = classes.filter((c) => !c.opgehaald || c.klas);
    if (!classes.some((c) => c.klas)) throw new Error(bron === "teams" ? "Geen enkel team gaf opdrachten terug. Ben je eigenaar van een klasteam (teamtype Klas)?" : "Graph geeft geen klassen terug.");

    // 3. Leden: education-rooster als dat mag, anders groepsleden minus eigenaren.
    const klassen = classes.filter((c) => c.klas);
    voortgang(`Leden ophalen (${klassen.length} klassen)…`, 0.24);
    const members = {};
    const velden = "id,displayName,mail,userPrincipalName,givenName,surname";
    const proef = await TT.graph(`/education/classes/${klassen[0].id}/members?$select=${velden},primaryRole`);
    const rosterOk = proef.status === 200;
    const mem = await TT.graphMany(klassen.map((c) => ({ key: c.id, url: rosterOk ? `/education/classes/${c.id}/members?$select=${velden},primaryRole` : `/groups/${c.id}/members?$select=${velden}` })));
    const lid = (u, rol) => ({ id: u.id, displayName: u.displayName, primaryRole: rol, mail: u.mail || "", userPrincipalName: u.userPrincipalName || "", givenName: u.givenName || "", surname: u.surname || "" });
    for (const c of klassen) {
      const m = mem.get(c.id);
      if (m?.status !== 200) { members[c.id] = oud?.members?.[c.id] || []; continue; }
      const owners = ownersVan.get(c.id) || new Set();
      members[c.id] = rosterOk
        ? m.value.filter((u) => (u.primaryRole || "student") === "student").map((u) => lid(u, u.primaryRole))
        : m.value.filter((u) => !owners.has(u.id) && /user$/i.test(u["@odata.type"] || "#microsoft.graph.user")).map((u) => lid(u, "student"));
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
        outcomes: (x.outcomes || []).filter((o) => /pointsOutcome|rubricOutcome/i.test(o["@odata.type"] || "")).map((o) => /rubricOutcome/i.test(o["@odata.type"]) 
          ? { "@odata.type": o["@odata.type"], niveaus: (o.publishedRubricQualitySelectedLevels?.length ? o.publishedRubricQualitySelectedLevels : o.rubricQualitySelectedLevels || []).map((l) => ({ qualityId: l.qualityId, columnId: l.columnId })) }
          : { "@odata.type": o["@odata.type"], points: o.points ? { points: o.points.points } : null, publishedPoints: o.publishedPoints ? { points: o.publishedPoints.points } : null }) }));
    }

    // 4b. Rubrics per opdracht: rijen (qualities), niveaus en punten. 404 = geen rubric.
    const rubrics = {};
    const rubricNodig = [];
    for (const a of actief) { if (herbruikbaar(a) && a.id in (oud?.rubrics || {})) rubrics[a.id] = oud.rubrics[a.id]; else rubricNodig.push(a); }
    if (rubricNodig.length) {
      voortgang(`Rubrics ophalen (${rubricNodig.length} opdrachten)…`, 0.9);
      const rb = await TT.graphMany(rubricNodig.map((a) => ({ key: a.id, url: `/education/classes/${a.classId}/assignments/${a.id}/rubric` })), { onProgress: (n, t) => voortgang(`Rubrics van ${n}/${t} opdrachten…`, 0.9 + 0.05 * (n / t)) });
      const tekst = (b) => (b && typeof b === "object" ? String(b.content || "") : String(b || "")).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      for (const a of rubricNodig) {
        const r = rb.get(a.id);
        if (r?.status !== 200 || !r.json) { rubrics[a.id] = r?.status === 404 ? null : (oud?.rubrics?.[a.id] ?? null); continue; }
        const j = r.json;
        rubrics[a.id] = { id: j.id, naam: j.displayName || "", beschrijving: tekst(j.description),
          niveaus: (j.levels || []).map((l) => ({ id: l.levelId, naam: l.displayName || "", beschrijving: tekst(l.description), punten: typeof l.grading?.maxPoints === "number" ? l.grading.maxPoints : null })),
          rijen: (j.qualities || []).map((q) => ({ id: q.qualityId, tekst: tekst(q.description), gewicht: typeof q.weight === "number" ? q.weight : null, criteria: (q.criteria || []).map((c) => ({ id: c.criterionId, tekst: tekst(c.description) })) })) };
      }
    }

    // 5. Namen van leerlingen die niet (meer) lid zijn.
    const bekend = new Set(Object.values(members).flat().map((u) => u.id));
    const onbekend = [...new Set(Object.values(submissions).flat().map((s) => s.recipient.userId).filter((u) => u && !bekend.has(u)))];
    const namen = { ...(oud?.namen || {}) }, profielen = { ...(oud?.profielen || {}) };
    const nogOp = onbekend.filter((u) => !namen[u] || !profielen[u]);
    if (nogOp.length) {
      voortgang(`Namen ophalen (${nogOp.length})…`, 0.96);
      const res = await TT.graphMany(nogOp.map((id) => ({ key: id, url: `/users/${id}?$select=${velden}` })), { prefer: "User.Read" });
      for (const [id, r] of res) if (r.status === 200 && r.json?.displayName) { namen[id] = r.json.displayName; profielen[id] = lid(r.json, "student"); }
    }
    const fouten = [...classes.filter((c) => c.fout).map((c) => `${c.displayName}: ${c.fout}`), ...actief.filter((a) => a.fout).map((a) => `${a.displayName}: ${a.fout}`)];
    const data = { schema: 4, at: Date.now(), volledigAt: oud ? oud.volledigAt || oud.at : Date.now(), hergebruikt, bron, rosterOk, me, weggelaten, totaalTeams, nietKlas: eerder - classes.length,
      classes: classes.map(({ id, displayName, classCode, externalName, isArchived, specialization, eigenaar, createdDateTime, opgehaald }) => ({ id, displayName, classCode, externalName, isArchived: !!isArchived, specialization: specialization || null, eigenaar: eigenaar ?? null, createdDateTime: createdDateTime || null, opgehaald: !!opgehaald })),
      members, assignments, submissions, categorieen, rubrics, namen, profielen, fouten };
    this._data = data;
    voortgang("Opslaan…", 0.98);
    try { await chrome.storage.local.set({ ttData: data }); }
    catch (e) { console.warn("[TT] opslaan mislukt", e); TT.toast(`Opslaan van de gegevens mislukt (${e?.message || e}); het overzicht werkt wel, maar na herladen moet alles opnieuw worden opgehaald.`, "err", 8000); }
    voortgang("Klaar", 1);
    this.renderWidget(TT);
    return data;
  },

  /* Klassen afgeleid uit lidmaatschap en teamnaam (zie analyse.js), binnen de geselecteerde teams. */
  patroon(TT) { const p = (TT.settings.opdrachten.klascodePatroon || "").trim() || globalThis.TT_KLASCODE || ""; this.patroonCache = p; return p; },
  klassen(TT) {
    const patroon = this.patroon(TT); // vóór selectie(): die gebruikt het patroon voor verborgen klassen
    const sel = this.selectie();
    const sleutel = JSON.stringify([this._data?.at, sel ? [...sel].sort() : null, patroon, this._verborgenVersie]);
    if (this._klassen?.sleutel !== sleutel) this._klassen = { sleutel, ...TT_ANALYSE.klassen(this.zichtbareData(), { teams: sel, patroon }) };
    return this._klassen;
  },
  /* Analyse wordt één keer per combinatie van gegevens en filters berekend; tabwissels zijn dan gratis. */
  analyse(TT) {
    const o = TT.settings.opdrachten;
    const K = this.klassen(TT);
    const sel = this.selectie();
    const sleutel = JSON.stringify([this._data?.at, this._filter.periode, sel ? [...sel].sort() : null, K.sleutel, o.drempelOntbreekt, o.drempelTeLaat]);
    if (this._analyseCache?.sleutel !== sleutel) this._analyseCache = { sleutel, r: TT_ANALYSE.bouw(this.zichtbareData(), { periode: this._filter.periode, klassen: sel, klasVan: K.klasVan, drempelOntbreekt: o.drempelOntbreekt, drempelTeLaat: o.drempelTeLaat }) };
    return this._analyseCache.r;
  },
  cijfers(TT) {
    const r = this.analyse(TT); // zorgt voor dezelfde sleutel
    if (this._cijfersCache?.r !== r) this._cijfersCache = { r, c: TT_ANALYSE.cijfers(this.zichtbareData(), { periode: this._filter.periode, klassen: this.selectie(), klasVan: this.klassen(TT).klasVan }) };
    return this._cijfersCache.c;
  },
  rubrics(TT) {
    const r = this.analyse(TT);
    if (this._rubricsCache?.r !== r) this._rubricsCache = { r, u: TT_ANALYSE.rubrics(this.zichtbareData(), { periode: this._filter.periode, klassen: this.selectie() }) };
    return this._rubricsCache.u;
  },
  /* ---------- Rapporten ---------- */
  /* Instellingen voor rapporten; lege velden vallen terug op het standaardsjabloon (TT_RAPPORT_STANDAARD),
   * vakgroepen op de naamheuristiek (AVO/BUR herkend op vaknaam, de rest praktijk). */
  rapportParse(r) {
    const regels = (t) => String(t || "").split("\n").map((x) => x.trim()).filter(Boolean);
    const periodes = regels(r.periodes).map((x) => x.split("|").map((y) => y.trim())).filter((p) => p.length >= 2).map(([naam, begin, einde]) => ({ naam, begin, einde: einde || "" }));
    const opleidingen = {}; for (const x of regels(r.opleidingen)) { const m = /^([A-Za-z]+)\s*[=:]\s*(.+)$/.exec(x); if (m) opleidingen[m[1].toUpperCase()] = m[2].trim(); }
    const vakgroepen = regels(r.vakgroepen).map((x) => { const i = x.indexOf(":"); if (i < 0) return null; return { naam: x.slice(0, i).trim(), vakken: x.slice(i + 1).split(",").map((v) => v.trim()).filter(Boolean) }; }).filter(Boolean);
    return { periodes, opleidingen, vakgroepen };
  },
  rapportInstellingen(TT) {
    const r = TT.settings.rapport || {}, S = globalThis.TT_RAPPORT_STANDAARD || {};
    const eigen = this.rapportParse(r), std = this.rapportParse(S);
    return {
      periodes: eigen.periodes.length ? eigen.periodes : std.periodes,
      opleidingen: { ...std.opleidingen, ...eigen.opleidingen },
      vakgroepen: eigen.vakgroepen.length ? eigen.vakgroepen : this.vakgroepenStandaard(TT),
      cesuur: Math.min(0.95, Math.max(0.05, (parseFloat(r.cesuur) || 55) / 100)), afronding: r.afronding === "decimaal" ? "decimaal" : "heel",
      toelichting: (r.toelichting || "").trim() || S.toelichting || "", titel: r.titel || "Resultaten jaar {jaar}", ingesteld: !!r.ingesteld,
      thema: (globalThis.TT_RAPPORT_THEMAS || []).some(([id]) => id === r.thema) ? r.thema : "magenta", school: r.school ?? "Mediacollege Amsterdam"
    };
  },
  /* Vakken (teamnaam zonder klascode) in de geselecteerde teams. */
  vakkenInData(TT) {
    let re = null; try { re = new RegExp(this.patroon(TT), "i"); } catch {}
    const sel = this.selectie();
    const uit = new Set();
    for (const c of this._data?.classes || []) { if (sel && !sel.has(c.id)) continue; const n = c.displayName || ""; const m = re && re.exec(n); const rest = m ? n.slice(m.index + m[0].length).trim() : n; if (rest) uit.add(rest); }
    return [...uit].sort((a, b) => a.localeCompare(b, "nl"));
  },
  vakgroepenStandaard(TT) {
    const S = globalThis.TT_RAPPORT_STANDAARD || {}, g = S.groepen || ["Praktijkvakken", "AVO-vakken en BUR"];
    const vakken = this.vakkenInData(TT);
    return [{ naam: g[0], vakken: vakken.filter((v) => !S.avo?.test(v)) }, { naam: g[1], vakken: vakken.filter((v) => S.avo?.test(v)) }].filter((x) => x.vakken.length);
  },
  opleidingCodes(TT) {
    const uit = new Set();
    for (const k of this.klassen(TT).klassen) { const m = /^([A-Za-z]+)\d{2}/.exec(k.naam); if (m) uit.add(m[1].toUpperCase()); }
    return [...uit].sort();
  },

  /* ---------- Wizard: in vijf stappen een goed rapportsjabloon, vooraf ingevuld uit de teams ---------- */
  rapportWizard(TT) {
    document.getElementById("tt-rapport")?.remove();
    const inst = this.rapportInstellingen(TT), S = globalThis.TT_RAPPORT_STANDAARD || {};
    const vakken = this.vakkenInData(TT);
    const codes = [...new Set([...this.opleidingCodes(TT), ...Object.keys(inst.opleidingen)])].sort();
    const groepNamen = inst.vakgroepen.length ? inst.vakgroepen.map((g) => g.naam) : (S.groepen || ["Praktijkvakken", "AVO-vakken en BUR"]);
    const st = {
      periodes: inst.periodes.map((p) => ({ ...p })),
      opleidingen: Object.fromEntries(codes.map((c) => [c, inst.opleidingen[c] || ""])),
      groepen: [...groepNamen],
      vakGroep: Object.fromEntries(vakken.map((v) => { const i = inst.vakgroepen.findIndex((g) => g.vakken.some((x) => x.toLowerCase() === v.toLowerCase())); return [v, i >= 0 ? i : (S.avo?.test(v) ? 1 : 0)]; })),
      cesuur: Math.round(inst.cesuur * 100), afronding: inst.afronding, titel: inst.titel, toelichting: inst.toelichting, thema: inst.thema, school: inst.school
    };
    let stap = 0;
    const box = TT.el("div", { id: "tt-rapport", class: "tt-wizard", role: "dialog", "aria-label": "Rapport instellen" });
    const sluit = () => { box.remove(); document.removeEventListener("keydown", esc, true); };
    const esc = (e) => { if (e.key === "Escape") { e.stopPropagation(); sluit(); } };
    document.addEventListener("keydown", esc, true);
    const input = (val, attrs, on) => { const i = TT.el("input", { class: "tt-input", value: val ?? "", ...attrs }); i.addEventListener("input", () => on(i.value)); return i; };
    const stappen = [
      ["Periodes", () => {
        const lijst = TT.el("div", { class: "tt-wiz-lijst" });
        const teken = () => lijst.replaceChildren(TT.el("div", { class: "tt-wiz-rij tt-wiz-kop" }, ["Naam", "Begin", "Einde", ""].map((h) => TT.el("span", { text: h }))),
          ...st.periodes.map((p, i) => TT.el("div", { class: "tt-wiz-rij" }, [
            input(p.naam, { "aria-label": "Naam" }, (v) => { p.naam = v; }), input(p.begin, { placeholder: "MM-DD", "aria-label": "Begin" }, (v) => { p.begin = v; }), input(p.einde, { placeholder: "MM-DD", "aria-label": "Einde" }, (v) => { p.einde = v; }),
            TT.el("button", { class: "tt-btn tt-btn-sm", text: "×", title: "Verwijderen", onclick: () => { st.periodes.splice(i, 1); teken(); } })
          ])));
        teken();
        return [TT.el("p", { class: "tt-muted", text: "Per periode de naam en de begin- en einddatum. Datums als MM-DD volgen elk schooljaar vanzelf; {jaar} wordt het leerjaar van de leerling (Periode 1.1 voor een eerstejaars, 2.1 voor een tweedejaars). Alleen periodes die begonnen zijn komen op het rapport." }), lijst,
          TT.el("div", { class: "tt-row" }, [TT.el("button", { class: "tt-btn tt-btn-sm", text: "Periode toevoegen", onclick: () => { st.periodes.push({ naam: `Periode {jaar}.${st.periodes.length + 1}`, begin: "", einde: "" }); teken(); } }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Standaard (4 periodes)", onclick: () => { st.periodes = this.rapportParse(S).periodes; teken(); } })])];
      }],
      ["Opleidingen", () => [TT.el("p", { class: "tt-muted", text: `Gevonden in je klascodes: ${codes.join(", ") || "geen"}. Vul per code de naam van de opleiding in zoals die op het rapport moet staan.` }),
        TT.el("div", { class: "tt-wiz-lijst" }, codes.map((c) => TT.el("div", { class: "tt-wiz-rij tt-wiz-2" }, [TT.el("b", { text: c }), input(st.opleidingen[c], { placeholder: "Naam van de opleiding" }, (v) => { st.opleidingen[c] = v; })]))),
        TT.el("div", { class: "tt-row" }, [TT.el("button", { class: "tt-btn tt-btn-sm", text: "Code toevoegen", onclick: () => { const c = prompt("Opleidingscode (zoals in de klascode, bijvoorbeeld MEP):"); if (c) { st.opleidingen[c.toUpperCase()] = ""; codes.push(c.toUpperCase()); teken(); } } })])]],
      ["Vakken", () => [TT.el("p", { class: "tt-muted", text: `${vakken.length} vakken gevonden in je teams (teamnaam zonder klascode). Zet elk vak in een groep; de voorinvulling herkent AVO-vakken en BUR op naam.` }),
        TT.el("div", { class: "tt-wiz-rij tt-wiz-2", style: "margin-bottom:6px" }, st.groepen.map((g, i) => input(g, { "aria-label": `Naam groep ${i + 1}` }, (v) => { st.groepen[i] = v; }))),
        TT.el("div", { class: "tt-wiz-lijst" }, vakken.map((v) => { const sel = TT.el("select", { class: "tt-select" }, [...st.groepen.map((g, i) => TT.el("option", { value: String(i), text: g || `Groep ${i + 1}` })), TT.el("option", { value: "-1", text: "Overige vakken" })]); sel.value = String(st.vakGroep[v]); sel.addEventListener("change", () => { st.vakGroep[v] = parseInt(sel.value, 10); }); return TT.el("div", { class: "tt-wiz-rij tt-wiz-2" }, [TT.el("span", { text: v }), sel]); }))]],
      ["Cijfers en tekst", () => {
        const afr = TT.el("select", { class: "tt-select" }, [["heel", "Hele cijfers"], ["decimaal", "Eén decimaal"]].map(([v, l]) => TT.el("option", { value: v, text: l }))); afr.value = st.afronding; afr.addEventListener("change", () => { st.afronding = afr.value; });
        const toel = TT.el("textarea", { class: "tt-input", style: "width:100%;min-height:110px;box-sizing:border-box", "aria-label": "Toelichting" }); toel.value = st.toelichting; toel.addEventListener("input", () => { st.toelichting = toel.value; });
        return [TT.el("div", { class: "tt-wiz-lijst" }, [
          TT.el("label", { class: "tt-wiz-rij tt-wiz-2" }, ["Cesuur (% voor een 5,5)", input(String(st.cesuur), { type: "number", min: "5", max: "95" }, (v) => { st.cesuur = parseInt(v, 10) || 55; })]),
          TT.el("label", { class: "tt-wiz-rij tt-wiz-2" }, ["Afronding", afr]),
          TT.el("label", { class: "tt-wiz-rij tt-wiz-2" }, ["Titel", input(st.titel, { placeholder: "Resultaten jaar {jaar}" }, (v) => { st.titel = v; })])
        ]), TT.el("p", { class: "tt-muted", text: "Toelichting onder de cijfers ({jaar}, {klas}, {cohort} en {schooljaar} worden ingevuld):" }), toel];
      }],
      ["Thema", () => {
        const themas = globalThis.TT_RAPPORT_THEMAS || [];
        const lijst = TT.el("div", { class: "tt-kiezer-lijst" }, themas.map(([id, naam, oms]) => { const rb = TT.el("input", { type: "radio", name: "tt-thema", value: id }); rb.checked = st.thema === id; rb.addEventListener("change", () => { st.thema = id; }); return TT.el("label", { class: "tt-kiezer-rij" }, [rb, TT.el("span", {}, [naam, TT.el("small", { text: oms })])]); }));
        return [TT.el("p", { class: "tt-muted", text: "Zwart/magenta is de huisstijl van Mediacollege Amsterdam, met het logo. De algemene thema's tonen de schoolnaam in plaats van het logo." }), lijst,
          TT.el("label", { class: "tt-wiz-rij tt-wiz-2", style: "margin-top:8px" }, ["Schoolnaam (bij algemene thema's)", input(st.school, { placeholder: "Naam van de school" }, (v) => { st.school = v; })])];
      }]
    ];
    const kop = TT.el("div", { class: "tt-row", style: "justify-content:space-between" }, [TT.el("b", { id: "tt-wiz-titel" }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: sluit })]);
    const stappenBalk = TT.el("div", { class: "tt-wiz-stappen" });
    const inhoud = TT.el("div", { class: "tt-wiz-inhoud" });
    const vorige = TT.el("button", { class: "tt-btn tt-btn-sm", text: "Vorige", onclick: () => { stap--; teken(); } });
    const volgende = TT.el("button", { class: "tt-btn tt-btn-sm tt-btn-primary", onclick: () => TT.guard(() => { if (stap < stappen.length - 1) { stap++; teken(); } else opslaan(); }, "Rapport instellen") });
    const opslaan = async () => {
      const rapport = {
        ingesteld: true,
        periodes: st.periodes.filter((p) => p.naam && p.begin).map((p) => `${p.naam} | ${p.begin} | ${p.einde}`).join("\n"),
        opleidingen: Object.entries(st.opleidingen).filter(([, n]) => n.trim()).map(([c, n]) => `${c} = ${n.trim()}`).join("\n"),
        vakgroepen: st.groepen.map((g, i) => `${g || `Groep ${i + 1}`}: ${vakken.filter((v) => st.vakGroep[v] === i).join(", ")}`).join("\n"),
        cesuur: st.cesuur, afronding: st.afronding, titel: st.titel || "Resultaten jaar {jaar}", toelichting: st.toelichting, thema: st.thema, school: st.school
      };
      await TT.saveSettings({ rapport });
      TT.toast("Rapportsjabloon opgeslagen", "ok");
      sluit();
      this.rapportDialoog(TT);
    };
    const teken = () => {
      kop.querySelector("#tt-wiz-titel").textContent = `Rapport instellen · stap ${stap + 1} van ${stappen.length}: ${stappen[stap][0]}`;
      stappenBalk.replaceChildren(...stappen.map(([n], i) => TT.el("span", { class: "tt-wiz-stap" + (i === stap ? " is-on" : i < stap ? " is-klaar" : ""), text: n })));
      inhoud.replaceChildren(...stappen[stap][1]());
      vorige.disabled = stap === 0;
      volgende.textContent = stap === stappen.length - 1 ? "Opslaan en rapporten maken" : "Volgende";
    };
    box.append(kop, stappenBalk, inhoud, TT.el("div", { class: "tt-row", style: "justify-content:space-between;margin-top:10px" }, [vorige, TT.el("span", { class: "tt-muted", text: "Alles is later aan te passen bij Instellingen › Rapporten." }), volgende]));
    document.body.append(box);
    teken();
  },

  /* Keuzevenster: klassen en periodes, dan de afdrukbare rapporten openen of naar Excel. */
  rapportDialoog(TT, voorKlas = null) {
    document.getElementById("tt-rapport")?.remove();
    const inst = this.rapportInstellingen(TT);
    if (!inst.ingesteld) return this.rapportWizard(TT);
    const K = this.klassen(TT);
    const klassen = K.klassen.filter((k) => k.leerlingen.length);
    const box = TT.el("div", { id: "tt-rapport", role: "dialog", "aria-label": "Rapporten" });
    const sluit = () => { box.remove(); document.removeEventListener("keydown", esc, true); };
    const esc = (e) => { if (e.key === "Escape") { e.stopPropagation(); sluit(); } };
    document.addEventListener("keydown", esc, true);
    const kl = klassen.map((k) => { const cb = TT.el("input", { type: "checkbox" }); cb.checked = !voorKlas || voorKlas === k.naam; return { k, cb, row: TT.el("label", { class: "tt-kiezer-rij" }, [cb, TT.el("span", {}, [k.naam, TT.el("small", { text: `${k.leerlingen.length} leerlingen · ${k.teams.length} teams` })])]) }; });
    const nu = Date.now();
    const pr = inst.periodes.map((p) => { const cb = TT.el("input", { type: "checkbox" }); cb.checked = new Date(p.begin).getTime() <= nu; return { p, cb, row: TT.el("label", { class: "tt-kiezer-rij" }, [cb, TT.el("span", {}, [p.naam, TT.el("small", { text: `${p.begin} t/m ${p.einde || "…"}` })])]) }; });
    const afr = TT.el("select", { class: "tt-select" }, [["heel", "Hele cijfers"], ["decimaal", "Eén decimaal"]].map(([v, l]) => TT.el("option", { value: v, text: l }))); afr.value = inst.afronding;
    const thema = TT.el("select", { class: "tt-select", "aria-label": "Thema" }, (globalThis.TT_RAPPORT_THEMAS || []).map(([id, naam]) => TT.el("option", { value: id, text: naam }))); thema.value = inst.thema;
    const status = TT.el("div", { class: "tt-muted", style: "margin-top:8px" });
    const maak = () => {
      const alleenKlassen = new Set(kl.filter((x) => x.cb.checked).map((x) => x.k.naam));
      const periodes = pr.filter((x) => x.cb.checked).map((x) => x.p);
      if (!alleenKlassen.size) { status.textContent = "Kies minstens één klas."; return null; }
      if (!periodes.length) { status.textContent = "Kies minstens één periode (instellen bij Instellingen › Rapporten)."; return null; }
      const r = TT_ANALYSE.rapport(this.zichtbareData(), { periodes, klasVan: K.klasVan, teamKlas: K.teamKlas, klassen: this.selectie(), patroon: this.patroon(TT), opleidingen: inst.opleidingen, vakgroepen: inst.vakgroepen, cesuur: inst.cesuur, afronding: afr.value, alleenKlassen, schooljaarStart: this.schooljaarStart() });
      if (!r.leerlingen.length) { status.textContent = "Geen leerlingen in deze klassen."; return null; }
      const sj = this.schooljaarStart().getFullYear();
      return { ...r, gemaakt: Date.now(), titel: inst.titel, toelichting: inst.toelichting, cesuur: inst.cesuur, afronding: afr.value, thema: thema.value, school: inst.school, schooljaar: `${sj}-${sj + 1}` };
    };
    const open = TT.el("button", { class: "tt-btn tt-btn-primary", text: "Rapporten openen (afdrukbaar)", onclick: () => TT.guard(async () => {
      const r = maak(); if (!r) return;
      await chrome.storage.local.set({ ttRapport: r });
      window.open(chrome.runtime.getURL("rapport/rapport.html"), "_blank");
      sluit();
    }, "Rapporten") });
    const excel = TT.el("button", { class: "tt-btn", text: "Naar Excel", onclick: () => TT.guard(() => {
      const r = maak(); if (!r) return;
      const f = (c) => (c == null ? "" : c);
      const rows = [["Studentnummer", "Roepnaam", "Achternaam", "Klas", "Opleiding", "Cohort", "Leerjaar", "Groep", "Vak", "Team", ...r.periodes.flatMap((p) => [p, `${p} (aantal)`])],
        ...r.leerlingen.flatMap((l) => l.vakken.map((v) => [l.studentnummer, l.roepnaam, l.achternaam, l.klas, l.opleiding, f(l.cohort), f(l.jaar), v.groep, v.vak, v.team, ...v.periodes.flatMap((p) => [f(p.cijfer), p.n])]))];
      const overz = [["Studentnummer", "Roepnaam", "Achternaam", "Klas", "Opleiding", "Cohort", "Leerjaar", "E-mail", "Aantal cijfers"], ...r.leerlingen.map((l) => [l.studentnummer, l.roepnaam, l.achternaam, l.klas, l.opleiding, f(l.cohort), f(l.jaar), l.mail, l.aantalCijfers])];
      TT.toXlsx([["Rapportcijfers", rows], ["Leerlingen", overz]], "rapporten");
    }, "Rapporten naar Excel") });
    box.append(
      TT.el("div", { class: "tt-row", style: "justify-content:space-between;margin-bottom:6px" }, [TT.el("b", { text: "Rapporten" }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: sluit })]),
      TT.el("div", { class: "tt-muted", text: "Klassen" }), TT.el("div", { class: "tt-kiezer-lijst", style: "max-height:30vh" }, kl.map((x) => x.row)),
      TT.el("div", { class: "tt-muted", style: "margin-top:8px", text: "Periodes" }), pr.length ? TT.el("div", { class: "tt-kiezer-lijst", style: "max-height:22vh" }, pr.map((x) => x.row)) : TT.el("p", { class: "tt-muted", text: "Nog geen periodes. Open de wizard om ze in te stellen." }),
      TT.el("div", { class: "tt-row", style: "margin-top:8px" }, [thema, afr, TT.el("span", { class: "tt-muted", text: `cesuur ${Math.round(inst.cesuur * 100)}%` }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Wizard", title: "Sjabloon in vijf stappen aanpassen", onclick: () => this.rapportWizard(TT) }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Instellingen", onclick: () => chrome.runtime.sendMessage({ type: "openOptions" }) })]),
      TT.el("div", { class: "tt-row", style: "margin-top:10px" }, [open, excel]), status
    );
    document.body.append(box);
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
  async open(TT, { vernieuw = false, volledig = false, vol = null } = {}) {
    if (!TT.alive()) TT.meldDood();
    // Weergave: automatisch volgt de route (icoon in de balk = native personal app, Alt+O en zijpaneel =
    // zwevend venster), of altijd overlay, of altijd native (instelling).
    const w = TT.settings.opdrachten.weergave || "auto";
    if (vol != null) this._vol = w === "overlay" ? false : w === "native" ? true : vol;
    else if (this._vol == null) this._vol = w === "native";
    if (!this._overlay) {
      this._overlay = TT.el("div", { id: "tt-overlay", role: "dialog", "aria-label": "Opdrachten over alle teams", onclick: (e) => { if (e.target === this._overlay) this.close(); } });
      document.body.append(this._overlay);
      this._esc = (e) => { if (e.key === "Escape" && !this._overlay.hidden) { e.stopPropagation(); this.close(); } };
      window.addEventListener("keydown", this._esc, true);
    }
    this._overlay.hidden = false;
    this._overlay.classList.toggle("tt-vol", !!this._vol);
    this._overlay.classList.toggle("tt-native", !!this._vol);
    this._overlay.classList.toggle("tt-native-dark", !!this._vol && this.teamsDonker());
    TT.module("appbar")?.actief?.(true);
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
  close() { if (this._overlay) this._overlay.hidden = true; TT.module("appbar")?.actief?.(false); },
  /* Donker of licht thema van Teams zelf: gemeten aan de achtergrondkleur van de pagina. */
  teamsDonker() {
    try {
      for (const el of [document.body, document.documentElement, document.querySelector("#app, main, [role=main]")]) {
        if (!el) continue;
        const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(getComputedStyle(el).backgroundColor || "");
        if (!m || (m[4] !== undefined && +m[4] < 0.5)) continue;
        return (0.2126 * m[1] + 0.7152 * m[2] + 0.0722 * m[3]) / 255 < 0.5;
      }
    } catch {}
    return matchMedia("(prefers-color-scheme: dark)").matches;
  },

  /* ---------- Pagina's: start › klas › team › leerling ----------
   * Geen tabbladen vol getallen, maar doorklikken: van klassen naar een klas, van daar naar een vak (team)
   * of een leerling. Elke pagina laat weinig zien en verwijst door. */
  _pagina: { type: "home" }, _geschiedenis: [],
  ga(TT, pagina) { this._geschiedenis.push(this._pagina); if (this._geschiedenis.length > 30) this._geschiedenis.shift(); this._pagina = pagina; this._limiet = null; this.renderBody(TT); this.renderTabs(TT); this._body?.scrollTo?.(0, 0); },
  terug(TT) { this._pagina = this._geschiedenis.pop() || { type: "home" }; this._limiet = null; this.renderBody(TT); this.renderTabs(TT); },

  render(TT) {
    this._limiet = null;
    const periode = TT.el("select", { class: "tt-select", "aria-label": "Periode", onchange: (e) => TT.guard(() => { this._filter.periode = e.target.value; this.render(TT); this.renderWidget(TT); this.saveFilter(); }, "Periode") }, this.periodeOpties(TT).map(([v, l]) => TT.el("option", { value: v, text: l })));
    if (!this.periodeOpties(TT).some(([v]) => v === this._filter.periode)) this._filter.periode = "schooljaar";
    periode.value = this._filter.periode;
    const zoekBox = TT.el("div", { class: "tt-zoek" });
    const zoek = TT.el("input", { class: "tt-input", type: "search", placeholder: "Zoek leerling, vak of klas…", "aria-label": "Zoeken", autocomplete: "off" });
    const resultaten = TT.el("div", { class: "tt-zoek-lijst", hidden: true });
    zoek.addEventListener("input", () => TT.guard(() => this.zoekLijst(TT, zoek.value, resultaten), "Zoeken"));
    zoek.addEventListener("focus", () => { if (zoek.value) this.zoekLijst(TT, zoek.value, resultaten); });
    zoek.addEventListener("blur", () => setTimeout(() => { resultaten.hidden = true; }, 150));
    zoekBox.append(zoek, resultaten);
    const sel = this.selectie(), totaal = this._data.classes.length, nSel = sel ? sel.size : totaal;
    const kiezer = TT.el("button", { class: "tt-btn tt-btn-sm", id: "tt-kiezer-knop", "aria-haspopup": "dialog", text: `Teams: ${nSel} van ${totaal} ▾`, title: "Welke teams meetellen", onclick: (e) => TT.guard(() => this.kiezer(TT, e.currentTarget), "Teams kiezen") });
    const acties = [
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Rapporten", title: "Rapport per leerling, afdrukbaar", onclick: () => TT.guard(() => this.rapportDialoog(TT), "Rapporten") }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Excel", title: "Alle overzichten in één werkmap", onclick: () => this.exportAlles(TT) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Bijwerken", title: `Opgehaald ${TT.fmtDateTime(this._data.at)}. Alt+klik haalt alles opnieuw op.`, onclick: (e) => this.bijwerken(TT, { volledig: e.altKey }) }),
      TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", onclick: () => this.close() })
    ];
    // Native: kop als personal app (icoon, naam, tabs voor de hoofdpagina's, acties rechts).
    const hoofdTabs = [["home", "Overzicht"], ["nakijken", "Na te kijken"], ["deadlines", "Deadlines"], ["signalen", "Signalen"], ["rubrics", "Rubrics"]];
    const tabKnoppen = () => hoofdTabs.map(([type, naam]) => TT.el("button", { class: "tt-app-tab" + (this._pagina.type === type || (type === "home" && ["klas", "team", "leerling"].includes(this._pagina.type)) ? " is-on" : ""), role: "tab", "data-type": type, text: naam, onclick: () => TT.guard(() => { this._geschiedenis = []; this._pagina = { type }; this._limiet = null; this.renderBody(TT); this.renderTabs(TT); }, "Tabblad") }));
    this._tabRij = TT.el("div", { class: "tt-app-tabs", role: "tablist" }, tabKnoppen());
    this._tabKnoppen = tabKnoppen;
    const head = this._vol
      ? TT.el("div", { class: "tt-sheet-head" }, [TT.el("div", { class: "tt-app" }, [TT.el("img", { src: chrome.runtime.getURL("icons/icon32.png"), alt: "" }), TT.el("h2", { text: "Docent Tools" })]), this._tabRij, TT.el("div", { class: "tt-app-acties" }, acties)])
      : TT.el("div", { class: "tt-sheet-head" }, [TT.el("h2", { text: "Docent Tools" }), ...acties]);
    const nV = this.verborgenAantal();
    const verborgen = TT.el("button", { class: "tt-btn tt-btn-sm" + (nV ? " tt-btn-accent" : ""), text: nV ? `Verborgen (${nV})` : "Verborgen", title: "Verborgen leerlingen, klassen, vakken en opdrachten terugzetten", onclick: (e) => TT.guard(() => this.verborgenBeheer(TT, e.currentTarget), "Verborgen") });
    const tools = TT.el("div", { class: "tt-sheet-tools" }, [zoekBox, periode, kiezer, verborgen, TT.el("span", { id: "tt-bijwerk-status", class: "tt-muted", style: "margin-left:auto", "aria-live": "polite" }), TT.el("span", { class: "tt-muted", text: `opgehaald ${TT.fmtDateTime(this._data.at)}` })]);
    this._body = TT.el("div", { class: "tt-sheet-body" });
    this._overlay.replaceChildren(TT.el("div", { class: "tt-sheet" }, [head, tools, this._body]));
    this.renderBody(TT);
  },

  renderTabs(TT) { if (this._vol && this._tabRij?.isConnected && this._tabKnoppen) this._tabRij.replaceChildren(...this._tabKnoppen()); },
  /* Zoeken: leerlingen, vakken (teams) en klassen; klikken springt naar de pagina. */
  zoekLijst(TT, q, host) {
    q = q.trim().toLowerCase();
    if (!q) { host.hidden = true; return; }
    const r = this.analyse(TT), K = this.klassen(TT);
    const hit = (s) => String(s || "").toLowerCase().includes(q);
    const uit = [
      ...K.klassen.filter((k) => k.leerlingen.length && hit(k.naam)).slice(0, 3).map((k) => ({ l: k.naam, sub: "klas", p: { type: "klas", id: k.naam } })),
      ...r.klassen.filter((t) => hit(t.naam)).slice(0, 4).map((t) => ({ l: t.naam, sub: "vak", p: { type: "team", id: t.id } })),
      ...r.leerlingen.filter((x) => hit(x.naam)).slice(0, 6).map((x) => ({ l: x.naam, sub: x.klas || "leerling", p: { type: "leerling", id: x.id } }))
    ];
    host.replaceChildren(...(uit.length ? uit.map((x) => TT.el("button", { class: "tt-zoek-rij", onmousedown: (e) => { e.preventDefault(); host.hidden = true; this.ga(TT, x.p); } }, [TT.el("span", { text: x.l }), TT.el("small", { text: x.sub })])) : [TT.el("div", { class: "tt-muted", style: "padding:8px", text: "Niets gevonden." })]));
    host.hidden = false;
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
      const sel = this.selectieBasis();
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
    const kop = () => { const sel = this.selectieBasis(); kopEl.textContent = this._filter.auto ? `Automatisch: ${sel ? sel.size : this._data.classes.length} eigen, actieve klasteams van dit schooljaar.` : `Handmatige keuze: ${sel ? sel.size : this._data.classes.length} van ${this._data.classes.length}.`; };
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
      const details = `Docent Tools v${TT.version} · pagina ${JSON.stringify(this._pagina)}\n${e?.stack || e}`;
      this._body?.replaceChildren(TT.el("div", { class: "tt-empty" }, [
        TT.el("p", { text: `Deze pagina kon niet worden opgebouwd: ${e?.message || e}` }),
        TT.el("div", { class: "tt-row", style: "justify-content:center" }, [
          TT.el("button", { class: "tt-btn tt-btn-sm", text: "Foutdetails kopiëren", onclick: () => TT.copy(details) }),
          TT.el("button", { class: "tt-btn tt-btn-sm", text: "Naar het begin", onclick: () => TT.guard(() => { this._pagina = { type: "home" }; this.render(TT); }, "Overzicht") })
        ])
      ]));
      if (!TT.alive()) TT.meldDood();
    }
  },

  /* ---- Bouwstenen ---- */
  ui(TT) {
    const P = TT.pct, D = TT.fmtDate;
    const chip = (tekst, kind) => TT.el("span", { class: "tt-chip" + (kind ? " is-" + kind : ""), text: tekst });
    const chips = (x) => [x.ontbreekt > 0 ? chip(`${x.ontbreekt} ontbreekt`, "bad") : null, x.teLaat > 0 ? chip(`${x.teLaat} te laat`, "warn") : null, x.nakijken > 0 ? chip(`${x.nakijken} na te kijken`, "warn") : null].filter(Boolean);
    const balk = (f, label) => TT.el("div", { class: "tt-balk" }, [TT.el("span", { class: "tt-bar" }, TT.el("i", { style: `width:${Math.round((f || 0) * 100)}%` })), TT.el("span", { class: "tt-muted", text: `${P(f) || "–"} ${label || ""}` })]);
    const kop = (titel, sub, extra) => TT.el("div", { class: "tt-pagina-kop" }, [TT.el("div", {}, [TT.el("h3", { text: titel }), sub ? TT.el("div", { class: "tt-muted", text: sub }) : null]), extra ? TT.el("div", { class: "tt-row" }, extra) : null]);
    const sectie = (titel, inhoud, hint) => TT.el("section", { class: "tt-sectie" }, [TT.el("h4", { text: titel }), hint ? TT.el("p", { class: "tt-muted", text: hint }) : null, inhoud]);
    const toets = (fn) => (e) => { if (e.target === e.currentTarget && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); fn(); } };
    const kruis = (v) => v ? TT.el("button", { class: "tt-verberg", title: `${v.label} verbergen…`, "aria-label": `${v.label} verbergen`, "aria-haspopup": "menu", text: "×", onclick: (e) => { e.stopPropagation(); TT.guard(() => this.verbergMenu(TT, e.currentTarget, v), "Verbergen"); }, onkeydown: (e) => { e.stopPropagation(); } }) : null;
    const rij = (hoofd, sub, rechts, onclick, verbergen) => TT.el("div", { class: "tt-rij", role: "button", tabindex: "0", onclick, onkeydown: toets(onclick) }, [TT.el("span", { class: "tt-rij-tekst" }, [TT.el("b", { text: hoofd }), sub ? TT.el("small", { text: sub }) : null]), TT.el("span", { class: "tt-rij-rechts" }, rechts), kruis(verbergen), TT.el("span", { class: "tt-rij-pijl", text: "›" })]);
    const kaart = (kids, onclick, verbergen, cls = "") => TT.el("div", { class: "tt-kaart " + cls, role: "button", tabindex: "0", onclick, onkeydown: toets(onclick) }, [kruis(verbergen), ...kids]);
    const leeg = (t) => TT.el("div", { class: "tt-empty", text: t });
    const kruimels = (delen) => TT.el("nav", { class: "tt-kruimels", "aria-label": "Pad" }, delen.flatMap((x, i) => [i ? TT.el("span", { class: "tt-muted", text: "›" }) : null, x.p ? TT.el("button", { class: "tt-kruimel", text: x.l, onclick: () => this.ga(TT, x.p) }) : TT.el("span", { class: "tt-kruimel is-hier", text: x.l })]).filter(Boolean));
    const gemChip = (v) => v == null ? chip("geen cijfers", "") : chip(`gem. ${P(v)}`, v < 0.55 ? "bad" : v < 0.7 ? "warn" : "ok");
    return { P, D, chip, chips, balk, kop, sectie, rij, kaart, leeg, kruimels, gemChip };
  },
  /* Sorteerbalk boven een lijst. opties: [{ k, l, f(x) → waarde, desc }]; de keuze wordt per lijst onthouden. */
  _sorteer: {},
  sorteer(TT, id, items, opties, standaard) {
    const keuze = this._sorteer[id] || standaard || opties[0].k;
    const o = opties.find((x) => x.k === keuze) || opties[0];
    const val = (x) => { const v = o.f(x); return v == null || v === "" ? null : v; };
    const gesorteerd = [...items].sort((a, b) => {
      const va = val(a), vb = val(b);
      if (va == null && vb == null) return 0; if (va == null) return 1; if (vb == null) return -1;
      const r = typeof va === "number" ? va - vb : String(va).localeCompare(String(vb), "nl", { numeric: true });
      return o.desc ? -r : r;
    });
    const balk = TT.el("div", { class: "tt-sorteer", role: "group", "aria-label": "Sorteren" }, [TT.el("span", { class: "tt-muted", text: "Sorteer op" }), ...opties.map((x) => TT.el("button", { class: "tt-sorteer-knop" + (x.k === o.k ? " is-on" : ""), "aria-pressed": String(x.k === o.k), text: x.l, onclick: () => TT.guard(() => { this._sorteer[id] = x.k; this.renderBody(TT); try { chrome.storage.local.set({ ttSorteer: this._sorteer }); } catch {} }, "Sorteren") }))]);
    return { balk, items: gesorteerd };
  },
  vakVan(TT, teamNaam) { let re = null; try { re = new RegExp(this.patroon(TT), "i"); } catch {} const m = re && re.exec(teamNaam || ""); const rest = m ? teamNaam.slice(m.index + m[0].length).trim() : teamNaam; return rest || teamNaam; },

  _renderBody(TT) {
    const r = this.analyse(TT), K = this.klassen(TT), C = this.cijfers(TT), U = this.ui(TT);
    const ko = this.klasOverzicht(TT, r);
    const pg = this._pagina;
    const gemVan = (uid, tid) => { const c = C.leerlingen.find((x) => x.id === uid); return tid ? c?.teams[tid]?.eind ?? null : c?.totaal ?? null; };
    // Sorteeropties voor leerlingenlijsten (x.l = leerling, x.k = de teller die telt: totaal of per team).
    const lnOpties = (tid) => [
      { k: "naam", l: "Naam", f: (x) => x.l.naam }, { k: "activiteit", l: "Activiteit", f: (x) => x.k.laatste || 0, desc: true },
      { k: "gemiddeld", l: "Gemiddelde", f: (x) => gemVan(x.l.id, tid), desc: true }, { k: "ontbreekt", l: "Ontbreekt", f: (x) => x.k.ontbreekt + x.k.teLaat * 0.01, desc: true }, { k: "teLaat", l: "Te laat", f: (x) => x.k.teLaat, desc: true }
    ];
    const start = { l: "Overzicht", p: { type: "home" } };
    let out = [];

    if (pg.type === "home") {
      const kaarten = this.zonderLokaal("start:klassen", ko.filter((k) => k.leerlingen > 0), (k) => k.naam).map((k) => U.kaart([
        TT.el("b", { text: k.naam }), TT.el("small", { text: `${k.leerlingen} leerlingen · ${k.aantalTeams} ${k.aantalTeams === 1 ? "vak" : "vakken"}` }),
        U.balk(k.pctIngeleverd, "ingeleverd"),
        TT.el("div", { class: "tt-row" }, [...U.chips(k), k.metSignaal > 0 ? U.chip(`${k.metSignaal} met signaal`, "bad") : null, !k.ontbreekt && !k.nakijken && !k.metSignaal ? U.chip("alles in orde", "ok") : null].filter(Boolean))
      ], () => this.ga(TT, { type: "klas", id: k.naam }), { type: "klassen", id: k.naam, label: k.naam, lijst: "start:klassen", lijstLabel: "Klassen op het startscherm" }));
      const nak = r.totaal.nakijken, dl = r.deadlines.filter((o) => o.dagen <= 14).length;
      out = [
        U.sectie("Mijn werk", TT.el("div", { class: "tt-kaarten" }, [
          U.kaart([TT.el("b", { text: String(nak) }), TT.el("small", { text: nak === 1 ? "inlevering wacht op nakijken" : "inleveringen wachten op nakijken" }), r.nakijken.length ? TT.el("span", { class: "tt-muted", text: `in ${r.nakijken.length} opdrachten · langste wachttijd ${r.nakijken[0].wachtDagen} dagen` }) : null], () => this.ga(TT, { type: "nakijken" }), null, "tt-kaart-actie"),
          U.kaart([TT.el("b", { text: String(dl) }), TT.el("small", { text: dl === 1 ? "deadline in de komende 14 dagen" : "deadlines in de komende 14 dagen" }), r.deadlines[0] ? TT.el("span", { class: "tt-muted", text: `eerstvolgende: ${r.deadlines[0].naam} (${U.D(r.deadlines[0].due)})` }) : null], () => this.ga(TT, { type: "deadlines" }), null, "tt-kaart-actie"),
          U.kaart([TT.el("b", { text: String(r.signalen.length) }), TT.el("small", { text: r.signalen.length === 1 ? "leerling met een signaal" : "leerlingen met een signaal" }), TT.el("span", { class: "tt-muted", text: "ontbrekend werk of te laat, over alle vakken opgeteld" })], () => this.ga(TT, { type: "signalen" }), null, "tt-kaart-actie")
        ])),
        U.sectie("Klassen", kaarten.length ? TT.el("div", { class: "tt-kaarten" }, kaarten) : U.leeg("Geen klassen met leerlingen in de gekozen teams."), "Klik op een klas voor de leerlingen en vakken."),
        this.rubrics(TT).rijen.some((x) => x.gedeeld) ? U.sectie("Rubrics", U.rij("Gedeelde rubric-rijen", `${this.rubrics(TT).rijen.filter((x) => x.gedeeld).length} rijen komen in meerdere vakken of opdrachten voor`, [], () => this.ga(TT, { type: "rubrics" }))) : null
      ];
    } else if (pg.type === "klas") {
      const k = ko.find((x) => x.naam === pg.id), g = K.klassen.find((x) => x.naam === pg.id);
      if (!k || !g) out = [U.kruimels([start, { l: pg.id }]), U.leeg("Deze klas staat niet (meer) in de selectie.")];
      else {
        const lijstL = `klas:${pg.id}:leerlingen`, lijstV = `klas:${pg.id}:vakken`;
        const lS = this.sorteer(TT, "klas-leerlingen", this.zonderLokaal(lijstL, r.leerlingen.filter((l) => l.klas === pg.id), (l) => l.id).map((l) => ({ l, k: l })), lnOpties(null), "naam");
        const tS = this.sorteer(TT, "klas-vakken", this.zonderLokaal(lijstV, g.teams.map((t) => r.klassen.find((x) => x.id === t.id)).filter(Boolean), (t) => t.id), [
          { k: "naam", l: "Naam", f: (t) => this.vakVan(TT, t.naam) }, { k: "nakijken", l: "Na te kijken", f: (t) => t.nakijken, desc: true }, { k: "ontbreekt", l: "Ontbreekt", f: (t) => t.ontbreekt, desc: true }, { k: "ingeleverd", l: "Ingeleverd", f: (t) => t.pctIngeleverd }, { k: "gemiddeld", l: "Gemiddelde", f: (t) => t.gemiddeld, desc: true }
        ], "naam");
        const lln = lS.items.map((x) => x.l), teams = tS.items;
        const teamIds = new Set(teams.map((t) => t.id));
        const dls = r.deadlines.filter((o) => teamIds.has(o.klasId)).slice(0, 5);
        out = [U.kruimels([start, { l: pg.id }]),
          U.kop(pg.id, `${k.leerlingen} leerlingen · ${teams.length} vakken · ${k.toegewezen} opdrachten toegewezen`, [U.gemChip(k.gemiddeld), ...U.chips(k), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Rapporten", onclick: () => TT.guard(() => this.rapportDialoog(TT, pg.id), "Rapporten") })]),
          TT.el("div", { class: "tt-twee" }, [
            U.sectie("Leerlingen", TT.el("div", {}, [lS.balk, lln.length ? TT.el("div", { class: "tt-lijst" }, lln.map((l) => U.rij(l.naam, `${l.ingeleverd} van ${l.toegewezen} ingeleverd${l.laatste ? " · laatst actief " + U.D(l.laatste) : ""}`, [U.gemChip(gemVan(l.id)), ...U.chips(l).slice(0, 2)], () => this.ga(TT, { type: "leerling", id: l.id }), { type: "leerlingen", id: l.id, label: l.naam, lijst: lijstL, lijstLabel: `Leerlingen van ${pg.id}` }))) : U.leeg("Geen leerlingen.")])),
            TT.el("div", {}, [
              U.sectie("Vakken", TT.el("div", {}, [tS.balk, TT.el("div", { class: "tt-lijst" }, teams.map((t) => U.rij(this.vakVan(TT, t.naam), `${t.opdrachten} opdrachten · ${U.P(t.pctIngeleverd) || "–"} ingeleverd`, U.chips(t).slice(0, 2), () => this.ga(TT, { type: "team", id: t.id }), { type: "teams", id: t.id, label: t.naam, lijst: lijstV, lijstLabel: `Vakken van ${pg.id}` })))])),
              dls.length ? U.sectie("Komende deadlines", TT.el("div", { class: "tt-lijst" }, dls.map((o) => U.rij(o.naam, `${this.vakVan(TT, o.klas)} · ${U.D(o.due, { weekday: "short", day: "numeric", month: "short" })} · ${o.dagen <= 0 ? "vandaag" : o.dagen === 1 ? "morgen" : "over " + o.dagen + " dagen"}`, [U.chip(`${o.open} open`, o.open ? "warn" : "ok")], () => this.ga(TT, { type: "team", id: o.klasId }))))) : null
            ])
          ])];
      }
    } else if (pg.type === "team") {
      const t = r.klassen.find((x) => x.id === pg.id);
      const c = this._data.classes.find((x) => x.id === pg.id);
      const klasNaam = K.teamKlas[pg.id];
      if (!t) out = [U.kruimels([start, { l: c?.displayName || "Vak" }]), U.leeg("Dit team heeft geen opdrachten in de gekozen periode.")];
      else {
        const vak = this.vakVan(TT, t.naam);
        const lijstO = `team:${pg.id}:opdrachten`, lijstTL = `team:${pg.id}:leerlingen`;
        const oS = this.sorteer(TT, "team-opdrachten", this.zonderLokaal(lijstO, r.opdrachten.filter((o) => o.klasId === pg.id), (o) => o.id), [
          { k: "deadline", l: "Deadline", f: (o) => o.due || 0, desc: true }, { k: "naam", l: "Naam", f: (o) => o.naam }, { k: "nakijken", l: "Na te kijken", f: (o) => o.nakijken, desc: true }, { k: "ontbreekt", l: "Ontbreekt", f: (o) => o.ontbreekt, desc: true }, { k: "ingeleverd", l: "Ingeleverd", f: (o) => o.pctIngeleverd }
        ], "deadline");
        const lS = this.sorteer(TT, "team-leerlingen", this.zonderLokaal(lijstTL, r.leerlingen.filter((l) => l.perKlas.some((k) => k.klasId === pg.id)), (l) => l.id).map((l) => ({ l, k: l.perKlas.find((k) => k.klasId === pg.id) })), lnOpties(pg.id), "naam");
        const ops = oS.items, lln = lS.items;
        const open = new Set();
        const opRij = (o) => {
          const key = "op:" + o.id;
          const knop = U.rij(o.naam, `${o.due ? "deadline " + U.D(o.due, { weekday: "short", day: "numeric", month: "short" }) : "zonder deadline"} · ${o.ingeleverd} van ${o.toegewezen} ingeleverd`, [...U.chips(o), o.gemiddeld != null ? U.gemChip(o.gemiddeld) : null].filter(Boolean), () => { this._open.has(key) ? this._open.delete(key) : this._open.add(key); this.renderBody(TT); }, { type: "opdrachten", id: o.id, label: o.naam, lijst: lijstO, lijstLabel: `Opdrachten van ${vak}` });
          if (!this._open.has(key)) return knop;
          const stat = r.leerlingen.map((l) => ({ l, o: l.opdrachten.find((x) => x.opdrachtId === o.id) })).filter((x) => x.o).sort((a, b) => a.l.naam.localeCompare(b.l.naam, "nl"));
          const kleur = { ontbreekt: "bad", ingeleverd: "warn", nagekeken: "ok", teruggegeven: "warn" };
          return TT.el("div", {}, [knop, TT.el("div", { class: "tt-detail" }, [
            TT.el("div", { class: "tt-row", style: "margin-bottom:6px" }, [o.webUrl ? TT.el("a", { class: "tt-btn tt-btn-sm", href: o.webUrl, target: "_blank", rel: "noopener", text: "Openen in Teams" }) : null]),
            TT.el("div", { class: "tt-lijst tt-lijst-compact" }, stat.map((x) => U.rij(x.l.naam, x.o.ingeleverdOp ? `ingeleverd ${U.D(x.o.ingeleverdOp)}${x.o.teLaat ? " (te laat)" : ""}` : "", [U.chip(x.o.klasse, kleur[x.o.klasse] || ""), x.o.punten != null ? U.chip(`${x.o.punten}${x.o.max ? " / " + x.o.max : ""}`, "") : null].filter(Boolean), () => this.ga(TT, { type: "leerling", id: x.l.id }))))
          ])]);
        };
        out = [U.kruimels([start, klasNaam ? { l: klasNaam, p: { type: "klas", id: klasNaam } } : null, { l: vak }].filter(Boolean)),
          U.kop(vak, `${t.naam} · ${t.leerlingen} leerlingen · ${t.opdrachten} opdrachten`, [U.gemChip(t.gemiddeld), ...U.chips(t), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Verbergen", title: "Dit vak uit het overzicht halen", onclick: () => TT.guard(() => { this.verberg(TT, "teams", t.id, t.naam); this.terug(TT); }, "Verbergen") })]),
          TT.el("div", { class: "tt-twee" }, [
            U.sectie("Opdrachten", TT.el("div", {}, [oS.balk, ops.length ? TT.el("div", { class: "tt-lijst" }, ops.map(opRij)) : U.leeg("Geen opdrachten in deze periode.")]), "Klik op een opdracht voor de stand per leerling."),
            U.sectie("Leerlingen", TT.el("div", {}, [lS.balk, TT.el("div", { class: "tt-lijst" }, lln.map(({ l, k }) => U.rij(l.naam, `${k.ingeleverd} van ${k.toegewezen} ingeleverd${k.laatste ? " · laatst actief " + U.D(k.laatste) : ""}`, [U.gemChip(gemVan(l.id, pg.id)), ...U.chips(k).slice(0, 2)], () => this.ga(TT, { type: "leerling", id: l.id }), { type: "leerlingen", id: l.id, label: l.naam, lijst: lijstTL, lijstLabel: `Leerlingen van ${vak}` })))]))
          ])];
      }
    } else if (pg.type === "leerling") {
      const l = r.leerlingen.find((x) => x.id === pg.id), c = C.leerlingen.find((x) => x.id === pg.id);
      if (!l) out = [U.kruimels([start, { l: "Leerling" }]), U.leeg("Deze leerling heeft geen opdrachten in de gekozen periode.")];
      else {
        const prof = this._data.profielen?.[l.id] || Object.values(this._data.members || {}).flat().find((m) => m.id === l.id) || {};
        const nr = (/^(\d{4,10})@/.exec(prof.mail || prof.userPrincipalName || "") || [])[1] || "";
        const sig = r.signalen.find((x) => x.id === l.id);
        const kleur = { ontbreekt: "bad", ingeleverd: "warn", nagekeken: "ok", teruggegeven: "warn" };
        const vakRijen = l.perKlas.map((k) => ({ ...k, id: k.klasId, vak: this.vakVan(TT, k.klas), gem: c?.teams[k.klasId] || null, eind: c?.teams[k.klasId]?.eind ?? null, nCijfers: c?.teams[k.klasId]?.n || 0 }));
        const vakDetail = (k) => {
          const ops = l.opdrachten.filter((o) => o.klasId === k.klasId);
          return TT.el("div", {}, [
            TT.el("div", { class: "tt-row", style: "margin-bottom:6px" }, [
              k.gem?.categorieen?.length ? TT.el("span", { class: "tt-muted", text: "Categorieën: " + k.gem.categorieen.map((x) => `${x.naam}${x.gewicht != null ? " " + x.gewicht + "%" : ""}: ${U.P(x.gemiddeld)} (${x.n})`).join(" · ") }) : null,
              TT.el("button", { class: "tt-btn tt-btn-sm", text: "Naar het vak", onclick: () => this.ga(TT, { type: "team", id: k.klasId }) })
            ]),
            TT.el("div", { class: "tt-lijst tt-lijst-compact" }, ops.map((o) => U.rij(o.opdracht, `${o.due ? "deadline " + U.D(o.due) : ""}${o.ingeleverdOp ? " · ingeleverd " + U.D(o.ingeleverdOp) : ""}${o.teLaat ? " (te laat)" : ""}`, [U.chip(o.klasse, kleur[o.klasse] || ""), o.punten != null ? U.chip(`${o.punten}${o.max ? " / " + o.max : ""}`, o.pct != null && o.pct < 0.55 ? "bad" : "") : null].filter(Boolean), () => this.ga(TT, { type: "team", id: o.klasId }))))
          ]);
        };
        const cel = (v) => (v == null ? { v: null, text: "–", cls: "tt-muted" } : { v, text: U.P(v), cls: v < 0.55 ? "tt-bad" : v < 0.7 ? "tt-warnc" : "" });
        const vakTabel = this.table(TT, "leerling-vakken", [
          { k: "vak", l: "Vak", cls: "tt-name", f: (k) => [k.vak, k.klas !== k.vak ? TT.el("span", { class: "tt-sub", text: k.klas }) : null] },
          { k: "eind", l: "Gemiddelde", n: 1, f: (k) => cel(k.eind) }, { k: "nCijfers", l: "Cijfers", n: 1 },
          { k: "ingeleverd", l: "Ingeleverd", n: 1, f: (k) => ({ v: k.ingeleverd, text: `${k.ingeleverd} / ${k.toegewezen}`, cls: "" }) },
          { k: "ontbreekt", l: "Ontbreekt", n: 1, f: (k) => ({ v: k.ontbreekt, text: k.ontbreekt ? String(k.ontbreekt) : "", cls: k.ontbreekt ? "tt-bad" : "" }) },
          { k: "teLaat", l: "Te laat", n: 1, f: (k) => ({ v: k.teLaat, text: k.teLaat ? String(k.teLaat) : "", cls: k.teLaat ? "tt-warnc" : "" }) },
          { k: "nakijken", l: "Na te kijken", n: 1, f: (k) => ({ v: k.nakijken, text: k.nakijken ? String(k.nakijken) : "", cls: k.nakijken ? "tt-warnc" : "" }) },
          { k: "laatste", l: "Laatste activiteit", f: (k) => U.D(k.laatste) }
        ], vakRijen, vakDetail, { k: "vak", asc: true });
        const gedeeldeSleutels = new Set(this.rubrics(TT).rijen.filter((x) => x.gedeeld).map((x) => x.sleutel));
        const LR = TT_ANALYSE.leerlingRubrics(this.zichtbareData(), l.id, { periode: this._filter.periode, klassen: this.selectie(), gedeeldeSleutels });
        const tab = this._leerlingTab || "vakken";
        const tabs = TT.el("div", { class: "tt-tabs", role: "tablist" }, [["vakken", "Vakken", l.aantalKlassen], ["rubrics", "Rubrics", LR.gedeeld.length + LR.perOpdracht.length]].map(([id, naam, n]) => TT.el("button", { class: "tt-tab" + (tab === id ? " is-on" : ""), role: "tab", "aria-selected": String(tab === id), onclick: () => { this._leerlingTab = id; this.renderBody(TT); } }, [naam, TT.el("span", { class: "tt-n", text: String(n) })])));
        out = [U.kruimels([start, l.klas ? { l: l.klas, p: { type: "klas", id: l.klas } } : null, { l: l.naam }].filter(Boolean)),
          U.kop(l.naam, [l.klas, nr ? "studentnummer " + nr : "", `${l.aantalKlassen} vakken`].filter(Boolean).join(" · "), [U.gemChip(c?.totaal ?? null), ...U.chips(l), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Rapport", onclick: () => TT.guard(() => this.rapportDialoog(TT, l.klas), "Rapport") }), TT.el("button", { class: "tt-btn tt-btn-sm", text: "Verbergen", title: "Deze leerling uit het overzicht halen", onclick: () => TT.guard(() => { this.verberg(TT, "leerlingen", l.id, l.naam); this.terug(TT); }, "Verbergen") })]),
          sig ? TT.el("div", { class: "tt-signaal" }, [TT.el("b", { text: "Signaal: " }), `${sig.redenen.join(", ")}${sig.klassenMetProbleem.length ? " · in " + sig.klassenMetProbleem.map((x) => this.vakVan(TT, x)).join(", ") : ""}`]) : null,
          TT.el("div", { style: "margin:0 0 12px" }, tabs),
          ...(tab === "vakken" ? [U.sectie("Per vak", vakTabel, "Klik op een kolomkop om te sorteren, op een rij voor de opdrachten en categorieën.")] : [
            (() => { const M = this.rubricMatrix(TT, LR); return M ? U.sectie("Rubric-rijen per vak", M, "Rijen die in meerdere opdrachten voorkomen; per vak het laatst gekozen niveau. Laagste niveau rood, hoogste groen. Klik op een rij voor het raster met alle niveaus en omschrijvingen.") : null; })(),
            LR.perOpdracht.length ? U.sectie("Per opdracht", TT.el("div", { class: "tt-lijst" }, LR.perOpdracht.map((o) => this.rubricOpdrachtRaster(TT, o))), "Klik op een opdracht voor het hele raster.") : null,
            !LR.gedeeld.length && !LR.perOpdracht.length ? U.leeg("Geen rubrics in de gekozen periode.") : null
          ])
        ];
      }
    } else if (pg.type === "nakijken") {
      out = [U.kruimels([start, { l: "Nakijken" }]), U.kop("Na te kijken", `${r.totaal.nakijken} inleveringen in ${r.nakijken.length} opdrachten, langst wachtende bovenaan`),
        ...(() => { const S_ = this.sorteer(TT, "nakijken", this.zonderLokaal("nakijken", r.nakijken, (o) => o.id), [{ k: "wacht", l: "Wachttijd", f: (o) => o.wachtDagen, desc: true }, { k: "aantal", l: "Aantal", f: (o) => o.nakijken, desc: true }, { k: "naam", l: "Naam", f: (o) => o.naam }, { k: "vak", l: "Vak", f: (o) => this.vakVan(TT, o.klas) }], "wacht"); return [S_.balk, S_.items.length ? TT.el("div", { class: "tt-lijst" }, S_.items.map((o) => U.rij(o.naam, `${this.vakVan(TT, o.klas)} · ${K.teamKlas[o.klasId] || ""}`, [U.chip(`${o.nakijken} wachten`, "warn"), U.chip(`${o.wachtDagen} dagen`, o.wachtDagen >= 14 ? "bad" : o.wachtDagen >= 7 ? "warn" : "")], () => { this._open.add("op:" + o.id); this.ga(TT, { type: "team", id: o.klasId }); }, { type: "opdrachten", id: o.id, label: o.naam, lijst: "nakijken", lijstLabel: "Na te kijken" }))) : U.leeg("Niets te nakijken.")]; })()];
    } else if (pg.type === "deadlines") {
      const dls = r.deadlines.filter((o) => o.dagen <= 30);
      out = [U.kruimels([start, { l: "Deadlines" }]), U.kop("Deadlines", "Komende 30 dagen, over alle klassen, op datum"),
        dls.length ? TT.el("div", { class: "tt-lijst" }, dls.map((o) => U.rij(o.naam, `${this.vakVan(TT, o.klas)} · ${K.teamKlas[o.klasId] || ""} · ${U.D(o.due, { weekday: "long", day: "numeric", month: "long" })}`, [U.chip(o.dagen <= 0 ? "vandaag" : o.dagen === 1 ? "morgen" : `over ${o.dagen} dagen`, o.dagen <= 1 ? "bad" : o.dagen <= 3 ? "warn" : ""), U.chip(`${o.open} open`, "")], () => this.ga(TT, { type: "team", id: o.klasId })))) : U.leeg("Geen deadlines in de komende 30 dagen.")];
    } else if (pg.type === "signalen") {
      const o = TT.settings.opdrachten;
      out = [U.kruimels([start, { l: "Signalen" }]), U.kop("Signalen", `Minstens ${o.drempelOntbreekt}× ontbrekend werk of ${o.drempelTeLaat}× te laat, over alle vakken opgeteld (drempels in de instellingen)`),
        ...(() => { const S_ = this.sorteer(TT, "signalen", this.zonderLokaal("signalen", r.signalen, (l) => l.id), [{ k: "zwaarte", l: "Zwaarte", f: (l) => l.ontbreekt + l.teLaat, desc: true }, { k: "naam", l: "Naam", f: (l) => l.naam }, { k: "klas", l: "Klas", f: (l) => l.klas }, { k: "activiteit", l: "Activiteit", f: (l) => l.laatste || 0, desc: true }], "zwaarte"); return [S_.balk, S_.items.length ? TT.el("div", { class: "tt-lijst" }, S_.items.map((l) => U.rij(l.naam, `${l.klas || ""} · ${l.redenen.join(", ")} · in ${l.klassenMetProbleem.map((x) => this.vakVan(TT, x)).join(", ")}`, U.chips(l).slice(0, 2), () => this.ga(TT, { type: "leerling", id: l.id }), { type: "leerlingen", id: l.id, label: l.naam, lijst: "signalen", lijstLabel: "Signalen" }))) : U.leeg("Geen signalen.")]; })()];
    } else if (pg.type === "rubrics") {
      const rijen = this.rubrics(TT).rijen.filter((x) => x.gedeeld);
      out = [U.kruimels([start, { l: "Rubrics" }]), U.kop("Gedeelde rubric-rijen", `${rijen.length} rijen komen in meerdere opdrachten of vakken voor (herkend op tekst)`),
        rijen.length ? TT.el("div", { class: "tt-lijst" }, rijen.map((rij) => { const key = "rb:" + rij.sleutel; const knop = U.rij(rij.tekst, `${rij.n} opdrachten · ${rij.teams.map((x) => this.vakVan(TT, x)).join(", ")}`, [U.chip(`${rij.aantalTeams} vakken`, rij.aantalTeams > 1 ? "ok" : "")], () => { this._open.has(key) ? this._open.delete(key) : this._open.add(key); this.renderBody(TT); }); return this._open.has(key) ? TT.el("div", {}, [knop, TT.el("div", { class: "tt-detail" }, this.rubricDetail(TT, rij))]) : knop; })) : U.leeg("Geen gedeelde rubric-rijen in deze periode.")];
    }
    this._body.replaceChildren(...out.filter(Boolean));
  },

  /* Raster van één gedeelde rij voor één leerling: kolommen = niveaus, rij 1 = omschrijving, daarna per
   * opdracht een stip op het gekozen niveau. Groepen met andere niveaunamen krijgen elk een eigen raster. */
  rubricRijRaster(TT, rij) {
    const U = this.ui(TT);
    const kop = TT.el("div", { class: "tt-rubric-kop" }, [
      TT.el("div", {}, [TT.el("b", { text: this.hoofdletter(rij.tekst) }), TT.el("small", { class: "tt-muted", text: ` · ${rij.aantalVakken} ${rij.aantalVakken === 1 ? "vak" : "vakken"} · ${rij.beoordeeld} van ${rij.voorkomens.length} beoordeeld` })]),
      rij.gemiddeld != null ? U.chip(`gem. ${U.P(rij.gemiddeld)}`, rij.gemiddeld < 0.5 ? "bad" : rij.gemiddeld < 0.7 ? "warn" : "ok") : null
    ]);
    const rasters = rij.groepen.map((g) => {
      const n = g.niveaus.length;
      const tabel = TT.el("table", { class: "tt-raster" }, [
        TT.el("thead", {}, TT.el("tr", {}, [TT.el("th", { text: "" }), ...g.niveaus.map((niv) => TT.el("th", {}, [TT.el("b", { text: niv.naam || "niveau" }), niv.punten != null ? TT.el("small", { text: `${niv.punten} pt` }) : null]))])),
        TT.el("tbody", {}, [
          g.criteria.some(Boolean) ? TT.el("tr", { class: "tt-raster-criteria" }, [TT.el("td", { text: "Omschrijving" }), ...g.criteria.slice(0, n).map((c) => TT.el("td", { text: c || "" }))]) : null,
          ...g.voorkomens.map((v) => TT.el("tr", { class: "is-klik", onclick: () => this.ga(TT, { type: "team", id: v.classId }) }, [
            TT.el("td", { class: "tt-raster-label" }, [TT.el("b", { text: this.vakVan(TT, v.klas) }), TT.el("small", { text: `${v.opdracht}${v.due ? " · " + U.D(v.due) : ""}` })]),
            ...g.niveaus.map((niv, j) => TT.el("td", { class: "tt-raster-cel" + (v.index === j ? " is-gekozen is-" + (this.niveauKleur({ ...v, niveaus: g.niveaus }) || "neutraal") : "") + (v.index == null ? " is-leeg" : "") }, v.index === j ? TT.el("span", { class: "tt-stip", title: `${niv.naam}${niv.punten != null ? " · " + niv.punten + " punten" : ""}` }) : null))
          ]))
        ].filter(Boolean))
      ]);
      return TT.el("div", { class: "tt-raster-wrap" }, tabel);
    });
    return TT.el("div", { class: "tt-rubric is-ingebed" }, [kop, ...rasters]);
  },
  /* Kleur van een gekozen niveau op positie: laagste rood, hoogste groen, ertussen neutraal; namen als
   * "onvoldoende" en "goed" wegen mee. */
  niveauKleur(v) {
    if (v.index == null) return "";
    const naam = String(v.niveau || "").toLowerCase(), n = v.niveaus?.length || 0;
    if (/onvoldoende|insufficient|slecht|niet/.test(naam)) return "bad";
    if (/uitstekend|excellent|goed|good|expert/.test(naam)) return "ok";
    if (v.index === 0) return "bad";
    if (n && v.index === n - 1) return "ok";
    return "";
  },
  hoofdletter(t) { t = String(t || "").trim(); return t.charAt(0).toUpperCase() + t.slice(1); },
  /* Matrix: rijen = rubric-rijen die in meerdere opdrachten voorkomen, kolommen = vakken, cel = laatst
   * gekozen niveau (met telletje bij meerdere opdrachten in dat vak). Klik op een rij voor het raster. */
  rubricMatrix(TT, LR) {
    const U = this.ui(TT);
    const rijen = LR.gedeeld.filter((r) => r.beoordeeld > 0 && r.voorkomens.length > 1);
    if (!rijen.length) return null;
    const vakken = [...new Set(rijen.flatMap((r) => r.voorkomens.map((v) => v.classId)))].map((id) => ({ id, naam: this.vakVan(TT, rijen.flatMap((r) => r.voorkomens).find((v) => v.classId === id).klas) })).sort((a, b) => a.naam.localeCompare(b.naam, "nl"));
    const data = rijen.map((r) => {
      const per = {};
      for (const vk of vakken) {
        const vs = r.voorkomens.filter((v) => v.classId === vk.id && v.index != null).sort((a, b) => (b.due || 0) - (a.due || 0));
        per[vk.id] = vs.length ? { laatste: vs[0], n: vs.length, alle: vs } : null;
      }
      return { id: r.sleutel, rij: r, tekst: this.hoofdletter(r.tekst), per, gemiddeld: r.gemiddeld, aantalVakken: r.aantalVakken, ...Object.fromEntries(vakken.map((vk) => ["v:" + vk.id, per[vk.id] ? per[vk.id].laatste.index : null])) };
    });
    const cel = (x, vk) => {
      const c = x.per[vk.id]; if (!c) return { v: null, text: "", cls: "tt-muted" };
      const kleur = this.niveauKleur(c.laatste);
      const node = TT.el("span", { class: "tt-niveau is-" + (kleur || "neutraal"), title: c.alle.map((v) => `${v.opdracht}${v.due ? " (" + U.D(v.due) + ")" : ""}: ${v.niveau}`).join("\n") }, [TT.el("span", { class: "tt-stip" }), TT.el("span", { text: c.laatste.niveau || "niveau " + (c.laatste.index + 1) }), c.n > 1 ? TT.el("small", { text: `×${c.n}` }) : null]);
      return node;
    };
    return this.table(TT, "leerling-rubricmatrix", [
      { k: "tekst", l: "Rubric-rij", cls: "tt-name", f: (x) => [x.tekst, TT.el("span", { class: "tt-sub", text: `${x.rij.voorkomens.length} opdrachten · ${x.aantalVakken} ${x.aantalVakken === 1 ? "vak" : "vakken"}` })] },
      ...vakken.map((vk) => ({ k: "v:" + vk.id, l: vk.naam, n: 1, f: (x) => cel(x, vk) })),
      { k: "gemiddeld", l: "Gem.", n: 1, f: (x) => (x.gemiddeld == null ? "" : { v: x.gemiddeld, text: U.P(x.gemiddeld), cls: x.gemiddeld < 0.5 ? "tt-bad" : x.gemiddeld < 0.7 ? "tt-warnc" : "" }) }
    ], data, (x) => this.rubricRijRaster(TT, x.rij), { k: "aantalVakken", asc: false });
  },
  /* Eén regel per gedeelde rij: titel, gemiddelde en per vak een label met het gekozen niveau; klik opent het raster. */
  rubricRijCompact(TT, rij) {
    const U = this.ui(TT);
    const key = "rbr:" + rij.sleutel;
    const labels = rij.voorkomens.map((v) => U.chip(`${this.vakVan(TT, v.klas)}: ${v.index == null ? "–" : v.niveau || "niveau " + (v.index + 1)}`, v.index == null ? "" : v.pct != null && v.pct < 0.5 ? "bad" : v.pct != null && v.pct < 0.7 ? "warn" : "ok"));
    const knop = TT.el("div", { class: "tt-rij tt-rij-rubric", role: "button", tabindex: "0", onclick: () => { this._open.has(key) ? this._open.delete(key) : this._open.add(key); this.renderBody(TT); }, onkeydown: (e) => { if (e.target === e.currentTarget && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); e.currentTarget.click(); } } }, [
      TT.el("span", { class: "tt-rij-tekst" }, [TT.el("b", { text: rij.tekst }), TT.el("small", { text: `${rij.aantalVakken} ${rij.aantalVakken === 1 ? "vak" : "vakken"} · ${rij.beoordeeld} van ${rij.voorkomens.length} beoordeeld` })]),
      TT.el("span", { class: "tt-rij-rechts" }, [...labels, rij.gemiddeld != null ? U.chip(`gem. ${U.P(rij.gemiddeld)}`, rij.gemiddeld < 0.5 ? "bad" : rij.gemiddeld < 0.7 ? "warn" : "ok") : null].filter(Boolean)),
      TT.el("span", { class: "tt-rij-pijl", text: this._open.has(key) ? "⌃" : "⌄" })
    ]);
    if (!this._open.has(key)) return knop;
    return TT.el("div", {}, [knop, TT.el("div", { class: "tt-detail" }, this.rubricRijRaster(TT, rij))]);
  },
  /* Volledig raster van één opdracht voor één leerling (ingeklapt tot je klikt). */
  rubricOpdrachtRaster(TT, o) {
    const U = this.ui(TT);
    const key = "rbo:" + o.aid;
    const chips = [o.beoordeeld ? U.chip(`${o.beoordeeld} van ${o.rijen.length} rijen beoordeeld`, o.beoordeeld === o.rijen.length ? "ok" : "warn") : U.chip("nog niet beoordeeld", ""), o.totaalPunten != null ? U.chip(`${o.totaalPunten} / ${o.maxPunten} pt`, o.maxPunten && o.totaalPunten / o.maxPunten < 0.5 ? "bad" : "") : null, o.gemiddeld != null ? U.chip(`gem. ${U.P(o.gemiddeld)}`, o.gemiddeld < 0.5 ? "bad" : o.gemiddeld < 0.7 ? "warn" : "ok") : null].filter(Boolean);
    const knop = U.rij(o.opdracht, `${this.vakVan(TT, o.klas)} · ${o.rubric || "rubric"}${o.due ? " · " + U.D(o.due) : ""} · ${o.rijen.length} rijen, ${o.niveaus.length} niveaus`, chips, () => { this._open.has(key) ? this._open.delete(key) : this._open.add(key); this.renderBody(TT); });
    if (!this._open.has(key)) return knop;
    const tabel = TT.el("table", { class: "tt-raster tt-raster-vol" }, [
      TT.el("thead", {}, TT.el("tr", {}, [TT.el("th", { text: "Rij" }), ...o.niveaus.map((niv) => TT.el("th", {}, [TT.el("b", { text: niv.naam || "niveau" }), niv.punten != null ? TT.el("small", { text: `${niv.punten} pt` }) : null, niv.beschrijving ? TT.el("small", { class: "tt-muted", text: niv.beschrijving }) : null])), TT.el("th", { class: "tt-num", text: "Punten" })])),
      TT.el("tbody", {}, o.rijen.map((r) => TT.el("tr", {}, [
        TT.el("td", { class: "tt-raster-label" }, [TT.el("b", { text: this.hoofdletter(r.tekst) }), r.gewicht != null ? TT.el("small", { text: `weging ${r.gewicht}%` }) : null]),
        ...o.niveaus.map((niv, j) => TT.el("td", { class: "tt-raster-cel tt-raster-tekst" + (r.index === j ? " is-gekozen is-" + (this.niveauKleur({ index: r.index, niveau: r.niveau, niveaus: o.niveaus }) || "neutraal") : "") }, [r.index === j ? TT.el("span", { class: "tt-stip" }) : null, TT.el("span", { text: r.criteria[j] || "" })])),
        TT.el("td", { class: "tt-num", text: r.punten != null ? String(r.punten) : r.index == null ? "–" : "" })
      ])))
    ]);
    return TT.el("div", {}, [knop, TT.el("div", { class: "tt-detail" }, [
      TT.el("div", { class: "tt-row", style: "margin-bottom:6px" }, [o.webUrl ? TT.el("a", { class: "tt-btn tt-btn-sm", href: o.webUrl, target: "_blank", rel: "noopener", text: "Openen in Teams" }) : null, TT.el("button", { class: "tt-btn tt-btn-sm", text: "Naar het vak", onclick: () => this.ga(TT, { type: "team", id: o.classId }) })]),
      TT.el("div", { class: "tt-raster-wrap" }, tabel),
      o.totaalPunten != null ? TT.el("div", { class: "tt-muted", style: "margin-top:6px", text: `Totaal ${o.totaalPunten} van ${o.maxPunten} punten` }) : null
    ])]);
  },
  rubricDetail(TT, rij) {
    const P = TT.pct, D = TT.fmtDate;
    const N = TT_ANALYSE.rijNiveaus(this.zichtbareData(), rij, { klasVan: this.klassen(TT).klasVan });
    return TT.el("div", {}, [
      N.length ? TT.el("table", { class: "tt-mini" }, [TT.el("tr", {}, ["Leerling", "Klas", "Gemiddeld", "Beoordelingen"].map((h) => TT.el("th", { text: h }))),
        ...N.map((x) => TT.el("tr", {}, [TT.el("td", {}, TT.el("a", { href: "#", text: x.naam, onclick: (e) => { e.preventDefault(); this.ga(TT, { type: "leerling", id: x.id }); } })), TT.el("td", { text: x.klas }), TT.el("td", { class: x.gemiddeld != null && x.gemiddeld < 0.5 ? "tt-bad" : "", text: P(x.gemiddeld) }), TT.el("td", { style: "white-space:normal", text: x.beoordelingen.map((b) => `${this.vakVan(TT, b.klas)} › ${b.opdracht} (${D(b.due)}): ${b.niveau}`).join(" · ") })]))])
        : TT.el("p", { class: "tt-muted", text: "Nog geen beoordelingen op deze rij." })
    ]);
  },

  /* Sorteerbare tabel. cols: { k, l, n (numeriek), cls, f(row) → tekst | node | [nodes] | { v, text, cls } } */
  _limiet: null,
  table(TT, id, cols, rows, detail, standaard) {
    if (!rows.length) return TT.el("div", { class: "tt-empty", text: "Niets gevonden voor deze filters." });
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
  cijferTabellen(TT) {
    const C = this.cijfers(TT), p = (x) => (x == null ? "" : Math.round(x * 100) + "%");
    const M = [["Leerling", "Klas", "Totaal", "Cijfers", ...C.teams.map((t) => t.naam)], ...C.leerlingen.map((l) => [l.naam, l.klas, p(l.totaal), l.n, ...C.teams.map((t) => p(l.teams[t.id]?.eind))])];
    const K = [["Leerling", "Klas", "Team", "Categorie", "Weging", "Cijfers", "Gemiddeld", "Gewogen team", "Gewoon team"], ...C.leerlingen.flatMap((l) => Object.values(l.teams).flatMap((t) => t.categorieen.map((c) => [l.naam, l.klas, t.klas, c.naam, c.gewicht ?? "", c.n, p(c.gemiddeld), p(t.gewogen), p(t.gemiddeld)])))];
    return [["Cijfers", M], ["Cijfers per categorie", K]];
  },
  rubricTabellen(TT) {
    const U = this.rubrics(TT), p = (x) => (x == null ? "" : Math.round(x * 100) + "%");
    const R = [["Rubric-rij", "Opdrachten", "Teams", "In teams", "Rubrics"], ...U.rijen.map((r) => [r.tekst, r.n, r.aantalTeams, r.teams.join(", "), r.rubrics.join(", ")])];
    const kv = this.klassen(TT).klasVan;
    const L = [["Rubric-rij", "Leerling", "Klas", "Gemiddeld", "Team", "Opdracht", "Niveau", "Punten"], ...U.rijen.filter((r) => r.gedeeld).flatMap((r) => TT_ANALYSE.rijNiveaus(this.zichtbareData(), r, { klasVan: kv }).flatMap((x) => x.beoordelingen.map((b) => [r.tekst, x.naam, x.klas, p(x.gemiddeld), b.klas, b.opdracht, b.niveau, b.punten ?? ""])))];
    return [["Rubric-rijen", R], ["Rubric-niveaus per leerling", L]];
  },
  exportAlles(TT) {
    const r = this.analyse(TT);
    TT.toXlsx([...this.klasTabellen(TT, r), ...TT_ANALYSE.naarTabellen(r), ...this.cijferTabellen(TT), ...this.rubricTabellen(TT)], "teams-opdrachten");
  },
});
