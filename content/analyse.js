/* Analyse van opdrachten over alle klassen (teams) heen. Pure functies, zonder DOM, zodat ze ook
 * in node te testen zijn (test/analyse.test.js).
 *
 * Invoer (zoals opgehaald uit Microsoft Graph, education-API):
 *   classes:     [{ id, displayName, classCode }]
 *   members:     { [classId]: [{ id, displayName, primaryRole }] }     (studenten; mag ontbreken)
 *   assignments: [{ id, classId, displayName, status, dueDateTime, assignedDateTime, closeDateTime,
 *                   allowLateSubmissions, grading: { "@odata.type", maxPoints }, webUrl }]
 *   submissions: { [assignmentId]: [{ id, status, submittedDateTime, returnedDateTime, reassignedDateTime,
 *                   excusedDateTime, recipient: { userId }, outcomes: [{ "@odata.type", points: { points }, publishedPoints }] }] }
 *   namen:       { [userId]: displayName }                               (aanvulling voor leerlingen zonder member-record)
 */
(() => {
  const DAG = 864e5;
  const ts = (s) => { if (!s) return null; const t = new Date(s).getTime(); return isNaN(t) ? null : t; };

  /* Periode-keuze → [begin, einde] in ms. Schooljaar loopt van 1 augustus tot 31 juli. */
  function periodeBereik(keuze, nu = Date.now()) {
    const d = new Date(nu);
    if (keuze === "30") return [nu - 30 * DAG, Infinity];
    if (keuze === "90") return [nu - 90 * DAG, Infinity];
    if (keuze === "alles") return [-Infinity, Infinity];
    const sj = /^sj(\d)$/.exec(String(keuze || ""));
    if (keuze && typeof keuze === "object") return [ts(keuze.begin) ?? -Infinity, keuze.einde ? ts(keuze.einde) + DAG - 1 : Infinity];
    const startJaar = d.getMonth() >= 7 ? d.getFullYear() : d.getFullYear() - 1;
    return [new Date(startJaar - (sj ? +sj[1] - 1 : 0), 7, 1).getTime(), new Date(startJaar + 1, 6, 31, 23, 59, 59).getTime()];
  }

  /* Status van één inlevering ten opzichte van de deadline. */
  function beoordeel(sub, asg, nu) {
    const due = ts(asg.dueDateTime);
    const sd = ts(sub.submittedDateTime);
    const st = sub.status || "working";
    const ingeleverd = st === "submitted" || st === "returned" || (st === "reassigned" && !!sd) || (!!sd && st !== "working");
    const teLaat = ingeleverd && due != null && sd != null && sd > due;
    let klasse;
    if (st === "excused") klasse = "vrijgesteld";
    else if (st === "returned") klasse = "nagekeken";
    else if (st === "submitted") klasse = "ingeleverd";
    else if (st === "reassigned") klasse = "teruggegeven";
    else if (due != null && nu > due) klasse = "ontbreekt";
    else klasse = "open";
    let punten = null, max = null, pct = null;
    const maxPoints = asg.grading && typeof asg.grading.maxPoints === "number" ? asg.grading.maxPoints : null;
    for (const o of sub.outcomes || []) {
      if (/pointsOutcome/i.test(o["@odata.type"] || "")) {
        const p = o.publishedPoints?.points ?? o.points?.points;
        if (typeof p === "number") { punten = p; max = maxPoints; pct = maxPoints ? p / maxPoints : null; }
      }
    }
    return { klasse, ingeleverd, teLaat, punten, max, pct, due, sd, returned: ts(sub.returnedDateTime) };
  }

  function leeg() { return { toegewezen: 0, ingeleverd: 0, opTijd: 0, teLaat: 0, ontbreekt: 0, open: 0, nagekeken: 0, teruggegeven: 0, vrijgesteld: 0, nakijken: 0, pctSom: 0, pctN: 0, laatste: null }; }
  function tel(acc, b) {
    acc.toegewezen++;
    if (b.ingeleverd) { acc.ingeleverd++; if (b.teLaat) acc.teLaat++; else acc.opTijd++; }
    // "ingeleverd" als klasse (wacht op nakijken) telt onder nakijken; de teller ingeleverd is alles wat is ingeleverd.
    if (b.klasse === "ingeleverd") acc.nakijken++; else acc[b.klasse] = (acc[b.klasse] || 0) + 1;
    if (b.pct != null) { acc.pctSom += b.pct; acc.pctN++; }
    if (b.sd && (!acc.laatste || b.sd > acc.laatste)) acc.laatste = b.sd;
  }
  function afronden(acc) {
    const n = acc.toegewezen || 0;
    return { ...acc, gemiddeld: acc.pctN ? acc.pctSom / acc.pctN : null, pctIngeleverd: n ? acc.ingeleverd / n : null, pctTeLaat: acc.ingeleverd ? acc.teLaat / acc.ingeleverd : null, pctOntbreekt: n ? acc.ontbreekt / n : null };
  }

  /* Actieve opdrachten binnen periode en klassenkeuze (gedeeld door bouw, cijfers en rubrics). */
  function actieveOpdrachten(data, opties = {}) {
    const nu = opties.nu || Date.now();
    const [van, tot] = periodeBereik(opties.periode || "schooljaar", nu);
    const actief = (a) => ["published", "assigned", "inactive"].includes(a.status || "assigned");
    const inPeriode = (a) => { const t = ts(a.dueDateTime) ?? ts(a.assignedDateTime) ?? ts(a.createdDateTime); return t == null || (t >= van && t <= tot); };
    const inKlassen = (a) => !opties.klassen || opties.klassen.has(a.classId);
    return (data.assignments || []).filter((a) => actief(a) && inPeriode(a) && inKlassen(a) && (ts(a.assignedDateTime) == null || ts(a.assignedDateTime) <= nu || a.status !== "scheduled"));
  }
  const gem = (xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null);

  /* ---------- Cijfers per leerling per team, met gewogen categorieën ----------
   * Per team: gemiddelde per cijfercategorie (percentage van de max. punten) en het gewogen gemiddelde
   * volgens de wegingen van de klas (data.categorieen[classId]), genormaliseerd over categorieën waarin
   * de leerling cijfers heeft. Zonder wegingen is gewogen = gewoon gemiddelde. */
  /* Gemiddelde van een reeks cijfers (items met pct en catId) volgens de categorieën van de klas. */
  function teamGemiddelde(items, cats = []) {
    const perCat = new Map();
    for (const it of items) { const k = it.catId || ""; if (!perCat.has(k)) perCat.set(k, []); perCat.get(k).push(it.pct); }
    const categorieen = [...perCat.entries()].map(([id, ps]) => { const c = cats.find((x) => x.id === id); return { id, naam: c?.naam || items.find((x) => x.catId === id)?.catNaam || (id ? "Onbekend" : "Zonder categorie"), gewicht: c?.gewicht ?? null, n: ps.length, gemiddeld: gem(ps) }; }).sort((a, b) => (b.gewicht || 0) - (a.gewicht || 0) || a.naam.localeCompare(b.naam, "nl"));
    const gewogenCats = categorieen.filter((c) => c.gewicht > 0);
    const somG = gewogenCats.reduce((a, c) => a + c.gewicht, 0);
    const gewogen = somG > 0 ? gewogenCats.reduce((a, c) => a + c.gemiddeld * c.gewicht, 0) / somG : null;
    const gemiddeld = gem(items.map((x) => x.pct));
    return { n: items.length, gemiddeld, gewogen, eind: gewogen ?? gemiddeld, categorieen };
  }
  /* Percentage → cijfer 1–10 met cesuur (standaard 55% = 5,5), lineair aan beide kanten. */
  function naarCijfer(pct, { cesuur = 0.55, afronding = "heel" } = {}) {
    if (pct == null || isNaN(pct)) return null;
    const c = Math.min(0.95, Math.max(0.05, cesuur));
    let x = pct >= c ? 5.5 + 4.5 * (pct - c) / (1 - c) : 1 + 4.5 * pct / c;
    x = Math.min(10, Math.max(1, x));
    return afronding === "decimaal" ? Math.round(x * 10) / 10 : Math.round(x);
  }

  function cijfers(data, opties = {}) {
    const nu = opties.nu || Date.now();
    const klassen = new Map((data.classes || []).map((c) => [c.id, c]));
    const memberNaam = new Map(Object.values(data.members || {}).flat().map((m) => [m.id, m.displayName]));
    const naam = (uid) => data.namen?.[uid] || memberNaam.get(uid) || uid;
    const per = new Map(); // uid -> Map(classId -> [{ aid, opdracht, catId, pct, punten, max, due }])
    const teamsMet = new Set();
    for (const a of actieveOpdrachten(data, opties)) {
      for (const sub of data.submissions?.[a.id] || []) {
        const b = beoordeel(sub, a, nu);
        if (b.pct == null) continue;
        const uid = sub.recipient?.userId || "?";
        if (!per.has(uid)) per.set(uid, new Map());
        const m = per.get(uid); if (!m.has(a.classId)) m.set(a.classId, []);
        m.get(a.classId).push({ aid: a.id, opdracht: a.displayName, catId: a.categorie?.id || null, catNaam: a.categorie?.naam || "", pct: b.pct, punten: b.punten, max: b.max, due: b.due });
        teamsMet.add(a.classId);
      }
    }
    const teams = [...teamsMet].map((id) => ({ id, naam: klassen.get(id)?.displayName || id })).sort((a, b) => a.naam.localeCompare(b.naam, "nl"));
    const leerlingen = [...per.entries()].map(([uid, m]) => {
      const perTeam = {};
      for (const [classId, items] of m) {
        perTeam[classId] = { klas: klassen.get(classId)?.displayName || classId, ...teamGemiddelde(items, data.categorieen?.[classId] || []), opdrachten: items.sort((a, b) => (b.due || 0) - (a.due || 0)) };
      }
      const eind = Object.values(perTeam).map((t) => t.eind);
      return { id: uid, naam: naam(uid), klas: opties.klasVan?.[uid] || "", teams: perTeam, aantalTeams: eind.length, totaal: gem(eind), n: Object.values(perTeam).reduce((a, t) => a + t.n, 0) };
    }).sort((a, b) => a.naam.localeCompare(b.naam, "nl"));
    return { teams, leerlingen };
  }

  /* ---------- Rubrics: gelijke rijen over teams en opdrachten heen ----------
   * Een rij (quality) wordt herkend aan de genormaliseerde tekst. Rijen die in meer dan één opdracht of
   * team voorkomen zijn "gedeeld". Niveaus en punten per voorkomen blijven beschikbaar voor rijNiveaus. */
  const normTekst = (t) => String(t || "").toLowerCase().replace(/[^\p{L}\p{N}\s]/gu, " ").replace(/\s+/g, " ").trim();
  function rubrics(data, opties = {}) {
    const klassen = new Map((data.classes || []).map((c) => [c.id, c]));
    const rijen = new Map(), perRubric = new Map();
    for (const a of actieveOpdrachten(data, opties)) {
      const rb = data.rubrics?.[a.id]; if (!rb) continue;
      const klasNaam = klassen.get(a.classId)?.displayName || a.classId;
      const rk = normTekst(rb.naam) || "(zonder naam)";
      if (!perRubric.has(rk)) perRubric.set(rk, { naam: rb.naam || "(zonder naam)", opdrachten: 0, teams: new Set(), rijen: rb.rijen.length });
      const pr = perRubric.get(rk); pr.opdrachten++; pr.teams.add(klasNaam);
      for (const q of rb.rijen) {
        const key = normTekst(q.tekst); if (!key) continue;
        if (!rijen.has(key)) rijen.set(key, { sleutel: key, tekst: q.tekst, voorkomens: [] });
        rijen.get(key).voorkomens.push({ aid: a.id, classId: a.classId, klas: klasNaam, opdracht: a.displayName, rubric: rb.naam || "", qualityId: q.id, gewicht: q.gewicht, niveaus: rb.niveaus, due: ts(a.dueDateTime) });
      }
    }
    const uit = [...rijen.values()].map((r) => {
      const teams = [...new Set(r.voorkomens.map((v) => v.klas))].sort((a, b) => a.localeCompare(b, "nl"));
      const rbs = [...new Set(r.voorkomens.map((v) => v.rubric).filter(Boolean))].sort();
      return { ...r, n: r.voorkomens.length, teams, aantalTeams: teams.length, rubrics: rbs, gedeeld: r.voorkomens.length > 1 };
    }).sort((a, b) => b.aantalTeams - a.aantalTeams || b.n - a.n || a.tekst.localeCompare(b.tekst, "nl"));
    const rubricsLijst = [...perRubric.values()].map((r) => ({ ...r, teams: [...r.teams].sort(), aantalTeams: r.teams.size })).sort((a, b) => b.opdrachten - a.opdrachten || a.naam.localeCompare(b.naam, "nl"));
    return { rijen: uit, rubrics: rubricsLijst };
  }
  /* Per leerling het gekozen niveau voor één rubric-rij, over alle voorkomens. pct = punten van het niveau
   * gedeeld door de hoogste punten van de rubric; zonder punten telt de positie van het niveau. */
  function rijNiveaus(data, rij, opties = {}) {
    const memberNaam = new Map(Object.values(data.members || {}).flat().map((m) => [m.id, m.displayName]));
    const naam = (uid) => data.namen?.[uid] || memberNaam.get(uid) || uid;
    const per = new Map();
    for (const v of rij.voorkomens) {
      const nivs = v.niveaus || [];
      const maxP = Math.max(...nivs.map((n) => n.punten ?? -Infinity));
      const metPunten = isFinite(maxP) && maxP > 0;
      for (const sub of data.submissions?.[v.aid] || []) {
        const uid = sub.recipient?.userId || "?";
        for (const o of sub.outcomes || []) {
          if (!/rubricOutcome/i.test(o["@odata.type"] || "")) continue;
          const sel = (o.niveaus || []).find((x) => x.qualityId === v.qualityId); if (!sel) continue;
          const idx = nivs.findIndex((n) => n.id === sel.columnId); if (idx < 0) continue;
          const niv = nivs[idx];
          const pct = metPunten ? (niv.punten ?? 0) / maxP : nivs.length > 1 ? idx / (nivs.length - 1) : null;
          if (!per.has(uid)) per.set(uid, []);
          per.get(uid).push({ aid: v.aid, opdracht: v.opdracht, klas: v.klas, niveau: niv.naam || `niveau ${idx + 1}`, index: idx, punten: niv.punten, max: metPunten ? maxP : null, pct, due: v.due });
        }
      }
    }
    return [...per.entries()].map(([uid, bs]) => ({ id: uid, naam: naam(uid), klas: opties.klasVan?.[uid] || "", beoordelingen: bs.sort((a, b) => (b.due || 0) - (a.due || 0)), n: bs.length, gemiddeld: gem(bs.map((b) => b.pct).filter((x) => x != null)) })).sort((a, b) => a.naam.localeCompare(b.naam, "nl"));
  }

  /**
   * Bouwt alle overzichten. opties: { periode, klassen (Set van classId of null), nu, drempelOntbreekt, drempelTeLaat }
   */
  function bouw(data, opties = {}) {
    const nu = opties.nu || Date.now();
    const [van, tot] = periodeBereik(opties.periode || "schooljaar", nu);
    const klassen = new Map((data.classes || []).map((c) => [c.id, c]));
    const memberNaam = new Map(Object.values(data.members || {}).flat().map((m) => [m.id, m.displayName]));
    const naam = (uid) => data.namen?.[uid] || memberNaam.get(uid) || uid;
    const actief = (a) => ["published", "assigned", "inactive"].includes(a.status || "assigned");
    const inPeriode = (a) => { const t = ts(a.dueDateTime) ?? ts(a.assignedDateTime) ?? ts(a.createdDateTime); return t == null || (t >= van && t <= tot); };
    const inKlassen = (a) => !opties.klassen || opties.klassen.has(a.classId);
    const asgs = (data.assignments || []).filter((a) => actief(a) && inPeriode(a) && inKlassen(a) && (ts(a.assignedDateTime) == null || ts(a.assignedDateTime) <= nu || a.status !== "scheduled"));

    const perKlas = new Map(), perLeerling = new Map(), perOpdracht = [];
    const leerlingKlas = new Map(); // uid -> Map(classId -> acc)
    const details = new Map(); // uid -> [{ klas, opdracht, klasse, ... }]
    for (const a of asgs) {
      const subs = data.submissions?.[a.id] || [];
      const k = klassen.get(a.classId) || { id: a.classId, displayName: a.classId };
      if (!perKlas.has(k.id)) perKlas.set(k.id, { id: k.id, naam: k.displayName, code: k.classCode || "", opdrachten: 0, leerlingen: new Set(), ...leeg() });
      const kAcc = perKlas.get(k.id); kAcc.opdrachten++;
      const oAcc = { id: a.id, klasId: k.id, klas: k.displayName, naam: a.displayName, due: ts(a.dueDateTime), toegewezenOp: ts(a.assignedDateTime), status: a.status, max: a.grading?.maxPoints ?? null, webUrl: a.webUrl || "", oudsteWacht: null, ...leeg() };
      for (const s of subs) {
        const uid = s.recipient?.userId || s.recipient?.id || "?";
        const b = beoordeel(s, a, nu);
        tel(kAcc, b); tel(oAcc, b); kAcc.leerlingen.add(uid);
        if (b.klasse === "ingeleverd" && b.sd && (!oAcc.oudsteWacht || b.sd < oAcc.oudsteWacht)) oAcc.oudsteWacht = b.sd;
        if (!perLeerling.has(uid)) perLeerling.set(uid, { id: uid, naam: naam(uid), klassen: new Set(), ...leeg() });
        const l = perLeerling.get(uid); tel(l, b); l.klassen.add(k.displayName);
        if (!leerlingKlas.has(uid)) leerlingKlas.set(uid, new Map());
        const lk = leerlingKlas.get(uid); if (!lk.has(k.id)) lk.set(k.id, { klasId: k.id, klas: k.displayName, ...leeg() });
        tel(lk.get(k.id), b);
        if (!details.has(uid)) details.set(uid, []);
        details.get(uid).push({ klas: k.displayName, klasId: k.id, opdracht: a.displayName, opdrachtId: a.id, klasse: b.klasse, teLaat: b.teLaat, due: b.due, ingeleverdOp: b.sd, punten: b.punten, max: b.max, pct: b.pct, webUrl: a.webUrl || "" });
      }
      perOpdracht.push(afronden(oAcc));
    }

    const klassenUit = [...perKlas.values()].map((k) => ({ ...afronden(k), leerlingen: k.leerlingen.size })).sort((a, b) => a.naam.localeCompare(b.naam, "nl"));
    const leerlingenUit = [...perLeerling.values()].map((l) => ({
      ...afronden(l), klas: opties.klasVan?.[l.id] || "", klassen: [...l.klassen].sort((a, b) => a.localeCompare(b, "nl")), aantalKlassen: l.klassen.size,
      perKlas: [...(leerlingKlas.get(l.id)?.values() || [])].map(afronden).sort((a, b) => a.klas.localeCompare(b.klas, "nl")),
      opdrachten: (details.get(l.id) || []).sort((a, b) => (b.due || 0) - (a.due || 0))
    })).sort((a, b) => a.naam.localeCompare(b.naam, "nl"));

    const deadlines = perOpdracht.filter((o) => o.due != null && o.due >= nu - DAG).sort((a, b) => a.due - b.due).map((o) => ({ ...o, dagen: Math.ceil((o.due - nu) / DAG) }));
    const nakijken = perOpdracht.filter((o) => o.nakijken > 0).map((o) => ({ ...o, wachtDagen: o.oudsteWacht ? Math.floor((nu - o.oudsteWacht) / DAG) : 0 })).sort((a, b) => b.wachtDagen - a.wachtDagen || b.nakijken - a.nakijken);

    const dO = opties.drempelOntbreekt ?? 2, dL = opties.drempelTeLaat ?? 3;
    const signalen = leerlingenUit.filter((l) => l.ontbreekt >= dO || l.teLaat >= dL).map((l) => ({
      ...l, redenen: [l.ontbreekt >= dO ? `${l.ontbreekt}× ontbreekt` : null, l.teLaat >= dL ? `${l.teLaat}× te laat` : null].filter(Boolean),
      // Meerdere klassen betrokken: pas dán is het team-overstijgend.
      klassenMetProbleem: l.perKlas.filter((k) => k.ontbreekt > 0 || k.teLaat > 0).map((k) => k.klas)
    })).sort((a, b) => (b.ontbreekt + b.teLaat) - (a.ontbreekt + a.teLaat));

    const totaal = afronden(klassenUit.reduce((acc, k) => { for (const f of Object.keys(leeg())) if (typeof acc[f] === "number") acc[f] += k[f] || 0; return acc; }, leeg()));
    return { periode: [van, tot], klassen: klassenUit, leerlingen: leerlingenUit, opdrachten: perOpdracht, deadlines, nakijken, signalen, totaal: { ...totaal, klassen: klassenUit.length, opdrachten: asgs.length, leerlingen: leerlingenUit.length } };
  }

  /* Tabellen voor Excel: [[bladnaam, rijen]] */
  function naarTabellen(r, fmt = { d: (t) => (t ? new Date(t).toISOString().slice(0, 10) : ""), p: (x) => (x == null ? "" : Math.round(x * 100) + "%") }) {
    const K = [["Team", "Code", "Leerlingen", "Opdrachten", "Toegewezen", "Ingeleverd", "% ingeleverd", "Op tijd", "Te laat", "Ontbreekt", "Open", "Nagekeken", "Na te kijken", "Gemiddeld"],
      ...r.klassen.map((k) => [k.naam, k.code, k.leerlingen, k.opdrachten, k.toegewezen, k.ingeleverd, fmt.p(k.pctIngeleverd), k.opTijd, k.teLaat, k.ontbreekt, k.open, k.nagekeken, k.nakijken, fmt.p(k.gemiddeld)])];
    const L = [["Leerling", "Klas", "Teams", "Aantal teams", "Toegewezen", "Ingeleverd", "% ingeleverd", "Op tijd", "Te laat", "Ontbreekt", "Open", "Nagekeken", "Gemiddeld", "Laatste inlevering"],
      ...r.leerlingen.map((l) => [l.naam, l.klas, l.klassen.join(", "), l.aantalKlassen, l.toegewezen, l.ingeleverd, fmt.p(l.pctIngeleverd), l.opTijd, l.teLaat, l.ontbreekt, l.open, l.nagekeken, fmt.p(l.gemiddeld), fmt.d(l.laatste)])];
    const LK = [["Leerling", "Team", "Toegewezen", "Ingeleverd", "Te laat", "Ontbreekt", "Open", "Nagekeken", "Gemiddeld"],
      ...r.leerlingen.flatMap((l) => l.perKlas.map((k) => [l.naam, k.klas, k.toegewezen, k.ingeleverd, k.teLaat, k.ontbreekt, k.open, k.nagekeken, fmt.p(k.gemiddeld)]))];
    const O = [["Team", "Opdracht", "Deadline", "Toegewezen", "Ingeleverd", "Te laat", "Ontbreekt", "Open", "Nagekeken", "Na te kijken", "Max. punten", "Gemiddeld"],
      ...r.opdrachten.map((o) => [o.klas, o.naam, fmt.d(o.due), o.toegewezen, o.ingeleverd, o.teLaat, o.ontbreekt, o.open, o.nagekeken, o.nakijken, o.max ?? "", fmt.p(o.gemiddeld)])];
    const D = [["Leerling", "Team", "Opdracht", "Deadline", "Status", "Te laat", "Ingeleverd op", "Punten", "Max", "Percentage"],
      ...r.leerlingen.flatMap((l) => l.opdrachten.map((o) => [l.naam, o.klas, o.opdracht, fmt.d(o.due), o.klasse, o.teLaat ? "ja" : "", fmt.d(o.ingeleverdOp), o.punten ?? "", o.max ?? "", fmt.p(o.pct)]))];
    const S = [["Leerling", "Klas", "Teams", "Ontbreekt", "Te laat", "Signaal", "Teams met probleem"],
      ...r.signalen.map((s) => [s.naam, s.klas, s.klassen.join(", "), s.ontbreekt, s.teLaat, s.redenen.join(", "), s.klassenMetProbleem.join(", ")])];
    return [["Teams", K], ["Leerlingen", L], ["Leerling per team", LK], ["Opdrachten", O], ["Signalen", S], ["Alle inleveringen", D]];
  }

  /* ---------- Klassen afleiden ----------
   * Op scholen met een team per vak is de klas niet één team maar een groep leerlingen die samen in
   * dezelfde teams zitten. Twee signalen:
   *  1. klascode in de teamnaam (opties.patroon, een regex met één groep, bv. "\\b(\\d[A-Z]{1,3}\\d?)\\b");
   *  2. overlap in lidmaatschap: twee teams horen bij dezelfde klas als het kleinste team voor minstens
   *     opties.drempel (0.5) uit dezelfde leerlingen bestaat, met minstens opties.minLeerlingen (3) gemeenschappelijk.
   * Elke leerling krijgt de klas waarvan die de meeste teams volgt.
   * Geeft { klassen: [{ id, naam, teams: [{ id, naam }], leerlingen: [{ id, naam, teams }] }], klasVan: { [uid]: klasnaam }, teamKlas: { [teamId]: klasnaam } } */
  function klassen(data, opties = {}) {
    const drempel = opties.drempel ?? 0.5, minL = opties.minLeerlingen ?? 3;
    const teams = (data.classes || []).filter((c) => !opties.teams || opties.teams.has(c.id));
    const naam = (uid) => data.namen?.[uid] || Object.values(data.members || {}).flat().find((m) => m.id === uid)?.displayName || uid;
    // Leerlingen per team: leden plus iedereen met een inlevering in een opdracht van dat team.
    const leden = new Map();
    for (const t of teams) leden.set(t.id, new Set((data.members?.[t.id] || []).map((m) => m.id)));
    for (const a of data.assignments || []) { const set = leden.get(a.classId); if (!set) continue; for (const s of data.submissions?.[a.id] || []) if (s.recipient?.userId) set.add(s.recipient.userId); }
    // Union-find over teams.
    const ouder = new Map(teams.map((t) => [t.id, t.id]));
    const vind = (x) => { while (ouder.get(x) !== x) { ouder.set(x, ouder.get(ouder.get(x))); x = ouder.get(x); } return x; };
    const voeg = (a, b) => { a = vind(a); b = vind(b); if (a !== b) ouder.set(a, b); };
    let re = null; try { if (opties.patroon) re = new RegExp(opties.patroon, "i"); } catch {}
    const code = (t) => { if (!re) return null; const m = re.exec(t.displayName || ""); return m ? (m[1] || m[0]).toUpperCase() : null; };
    // 1. Zelfde klascode in de naam → zelfde klas.
    const perCode = new Map();
    for (const t of teams) { const c = code(t); if (c) { if (perCode.has(c)) voeg(t.id, perCode.get(c)); else perCode.set(c, t.id); } }
    // 2. Overlap in lidmaatschap: een team zonder herkende code sluit aan bij het team (met of zonder code)
    //    waarmee het de meeste leerlingen deelt, als dat genoeg is.
    const overlap = (A, B) => { let n = 0; for (const x of A) if (B.has(x)) n++; return n; };
    for (const t of teams) {
      if (code(t)) continue;
      const A = leden.get(t.id); if (!A.size) continue;
      let best = null, bestN = 0;
      for (const u of teams) {
        if (u === t) continue;
        const B = leden.get(u.id); if (!B.size) continue;
        const n = overlap(A, B);
        if (n >= minL && n / Math.min(A.size, B.size) >= drempel && n > bestN) { best = u; bestN = n; }
      }
      if (best) voeg(t.id, best.id);
    }
    // Groepen samenstellen.
    const groepen = new Map();
    for (const t of teams) { const r = vind(t.id); if (!groepen.has(r)) groepen.set(r, []); groepen.get(r).push(t); }
    const STOP = /^(team|klas|class|groep|vak|les|en|de|het|van|\d{4}|\d{2}-\d{2}|\d{4}-\d{4}|\d{4}-\d{2})$/i;
    const tokens = (n) => String(n || "").split(/[^\p{L}\p{N}]+/u).filter((x) => x.length >= 2 && !STOP.test(x));
    const benoem = (ts) => {
      // Heeft een team in de groep een klascode, dan is de meest voorkomende code de naam.
      const tel = new Map(); for (const t of ts) { const c = code(t); if (c) tel.set(c, (tel.get(c) || 0) + 1); }
      if (tel.size) return [...tel.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0][0];
      if (ts.length === 1) return ts[0].displayName;
      const telW = new Map();
      for (const t of ts) for (const x of new Set(tokens(t.displayName).map((y) => y.toUpperCase()))) telW.set(x, (telW.get(x) || 0) + 1);
      const gemeen = [...telW.entries()].filter(([, n]) => n >= Math.max(2, Math.ceil(ts.length * 0.6))).map(([x]) => x);
      const metCijfer = gemeen.filter((x) => /\d/.test(x));
      const keus = (metCijfer.length ? metCijfer : gemeen).sort((a, b) => a.length - b.length || a.localeCompare(b))[0];
      return keus || ts.map((t) => t.displayName).sort()[0];
    };
    const uit = [];
    for (const ts of groepen.values()) {
      const telPer = new Map();
      for (const t of ts) for (const u of leden.get(t.id)) telPer.set(u, (telPer.get(u) || 0) + 1);
      uit.push({ teams: ts.map((t) => ({ id: t.id, naam: t.displayName })).sort((a, b) => a.naam.localeCompare(b.naam, "nl")), telPer, naam: benoem(ts) });
    }
    // Dubbele namen uniek maken.
    const gezien = new Map();
    for (const g of uit) { const n = gezien.get(g.naam) || 0; gezien.set(g.naam, n + 1); if (n) g.naam = `${g.naam} (${n + 1})`; }
    uit.sort((a, b) => a.naam.localeCompare(b.naam, "nl", { numeric: true }));
    uit.forEach((g, i) => { g.id = "k" + i; });
    // Leerling → klas met de meeste teams (bij gelijk: de grootste klas).
    const beste = new Map();
    for (const g of uit) for (const [u, n] of g.telPer) { const b = beste.get(u); if (!b || n > b.n || (n === b.n && g.telPer.size > b.g.telPer.size)) beste.set(u, { n, g }); }
    const klasVan = {}, teamKlas = {};
    for (const g of uit) { g.leerlingen = []; for (const t of g.teams) teamKlas[t.id] = g.naam; }
    for (const [u, b] of beste) { klasVan[u] = b.g.naam; b.g.leerlingen.push({ id: u, naam: naam(u), teams: b.n }); }
    for (const g of uit) { g.leerlingen.sort((a, b) => a.naam.localeCompare(b.naam, "nl")); delete g.telPer; }
    return { klassen: uit, klasVan, teamKlas };
  }

  /* ---------- Rapport per leerling ----------
   * opties: { periodes: [{ naam, begin: "JJJJ-MM-DD", einde }], klasVan, teamKlas, klassen (Set van teamIds),
   *   patroon (klascode-regex), opleidingen: { MR: "Mediaredactiemedewerker" }, vakgroepen: [{ naam, vakken: [] }],
   *   cesuur, afronding, schooljaarStart (Date), alleenKlassen (Set van klasnamen) }
   * Vak = teamnaam zonder klascode. Cijfer per vak per periode = gewogen gemiddelde van de opdrachten met een
   * deadline in die periode. Studentnummer = cijfers vóór de @ van het e-mailadres. Cohort en leerjaar uit de
   * klascode (MIP26A: cohort 2026, leerjaar = schooljaar − cohort + 1). */
  function rapport(data, opties = {}) {
    const nu = opties.nu || Date.now();
    const klassen = new Map((data.classes || []).map((c) => [c.id, c]));
    const members = Object.values(data.members || {}).flat();
    const memberNaam = new Map(members.map((m) => [m.id, m.displayName]));
    const profiel = (uid) => data.profielen?.[uid] || members.find((m) => m.id === uid) || {};
    const naam = (uid) => data.namen?.[uid] || memberNaam.get(uid) || uid;
    let re = null; try { if (opties.patroon) re = new RegExp(opties.patroon, "i"); } catch {}
    const vakNaam = (t) => { const n = t.displayName || ""; const m = re && re.exec(n); if (!m) return n; const rest = n.slice(m.index + m[0].length).trim(); return rest || n; };
    const sjStart = opties.schooljaarStart ? new Date(opties.schooljaarStart) : new Date(new Date(nu).getMonth() >= 7 ? new Date(nu).getFullYear() : new Date(nu).getFullYear() - 1, 7, 1);
    // Datum "MM-DD" krijgt het jaar van het lopende schooljaar: augustus t/m december het startjaar, daarna het jaar erna.
    const datum = (d) => { if (!d) return null; const m = /^(\d{2})-(\d{2})$/.exec(String(d).trim()); if (!m) return ts(d); const mnd = +m[1]; return new Date(mnd >= 8 ? sjStart.getFullYear() : sjStart.getFullYear() + 1, mnd - 1, +m[2]).getTime(); };
    const periodes = (opties.periodes || []).map((p) => ({ naam: p.naam, begin: datum(p.begin) ?? -Infinity, einde: p.einde ? datum(p.einde) + DAG - 1 : Infinity })).filter((p) => p.begin <= nu).sort((a, b) => a.begin - b.begin);
    const periodeNaam = (naam, jaar) => String(naam).replace(/\{jaar\}/g, jaar ?? "");
    const inPeriode = (a, p) => { const t = ts(a.dueDateTime) ?? ts(a.assignedDateTime); return t != null && t >= p.begin && t <= p.einde; };
    const klasInfo = (klas) => {
      const m = /^([A-Za-z]+)(\d{2})([A-Za-z]?)/.exec(klas || "");
      if (!m) return { opleidingCode: "", opleiding: "", cohort: null, jaar: null };
      const cohort = 2000 + parseInt(m[2], 10);
      const jaar = Math.max(1, sjStart.getFullYear() - cohort + 1);
      const code = m[1].toUpperCase();
      return { opleidingCode: code, opleiding: opties.opleidingen?.[code] || code, cohort, jaar };
    };
    const groepVan = (vak) => { const v = vak.toLowerCase(); for (const g of opties.vakgroepen || []) if (g.vakken.some((x) => x.toLowerCase() === v)) return g.naam; return (opties.vakgroepen || []).length ? "Overige vakken" : "Vakken"; };
    // Cijfers per leerling per team per periode.
    const per = new Map(); // uid -> Map(classId -> Map(periodeIndex -> items))
    for (const a of actieveOpdrachten(data, { ...opties, periode: "alles" })) {
      const pi = periodes.findIndex((p) => inPeriode(a, p)); if (pi < 0) continue;
      for (const sub of data.submissions?.[a.id] || []) {
        const b = beoordeel(sub, a, nu); if (b.pct == null) continue;
        const uid = sub.recipient?.userId || "?";
        if (!per.has(uid)) per.set(uid, new Map());
        const m = per.get(uid); if (!m.has(a.classId)) m.set(a.classId, new Map());
        const mm = m.get(a.classId); if (!mm.has(pi)) mm.set(pi, []);
        mm.get(pi).push({ aid: a.id, opdracht: a.displayName, catId: a.categorie?.id || null, catNaam: a.categorie?.naam || "", pct: b.pct });
      }
    }
    // Welke leerlingen: alle leerlingen van de geselecteerde teams (ook zonder cijfers).
    const uids = new Set();
    for (const [classId, ms] of Object.entries(data.members || {})) if (!opties.klassen || opties.klassen.has(classId)) for (const m of ms) uids.add(m.id);
    for (const uid of per.keys()) uids.add(uid);
    const leerlingen = [...uids].map((uid) => {
      const klas = opties.klasVan?.[uid] || "";
      if (opties.alleenKlassen && !opties.alleenKlassen.has(klas)) return null;
      const p = profiel(uid);
      const mail = p.mail || p.userPrincipalName || "";
      const nr = (/^(\d{4,10})@/.exec(mail) || [])[1] || "";
      const volledig = naam(uid);
      const roepnaam = p.givenName || volledig.split(" ")[0] || "";
      const achternaam = p.surname || volledig.split(" ").slice(1).join(" ") || "";
      // Teams van de klas van deze leerling (plus teams waar de leerling cijfers heeft).
      const teamIds = new Set([...(data.classes || []).filter((c) => opties.teamKlas?.[c.id] === klas && (!opties.klassen || opties.klassen.has(c.id)) && (data.members?.[c.id] || []).some((m) => m.id === uid)).map((c) => c.id), ...(per.get(uid)?.keys() || [])]);
      const info = klasInfo(klas);
      const vakken = [...teamIds].map((tid) => {
        const t = klassen.get(tid); if (!t) return null;
        const vak = vakNaam(t);
        const perP = periodes.map((pr, i) => { const items = per.get(uid)?.get(tid)?.get(i) || []; const g = teamGemiddelde(items, data.categorieen?.[tid] || []); return { naam: periodeNaam(pr.naam, info.jaar), n: items.length, pct: g.eind, cijfer: naarCijfer(g.eind, opties) }; });
        return { vak, team: t.displayName, teamId: tid, groep: groepVan(vak), periodes: perP, n: perP.reduce((a, x) => a + x.n, 0) };
      }).filter(Boolean).sort((a, b) => a.groep.localeCompare(b.groep, "nl") || a.vak.localeCompare(b.vak, "nl"));
      const groepen = [];
      for (const v of vakken) { let g = groepen.find((x) => x.naam === v.groep); if (!g) { g = { naam: v.groep, vakken: [] }; groepen.push(g); } g.vakken.push(v); }
      const volgorde = (opties.vakgroepen || []).map((g) => g.naam);
      groepen.sort((a, b) => { const ia = volgorde.indexOf(a.naam), ib = volgorde.indexOf(b.naam); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.naam.localeCompare(b.naam, "nl"); });
      return { id: uid, naam: volledig, roepnaam, achternaam, studentnummer: nr, mail, klas, ...info, periodeNamen: periodes.map((p) => periodeNaam(p.naam, info.jaar)), vakken, groepen, aantalCijfers: vakken.reduce((a, v) => a + v.n, 0) };
    }).filter(Boolean).sort((a, b) => a.klas.localeCompare(b.klas, "nl", { numeric: true }) || a.achternaam.localeCompare(b.achternaam, "nl") || a.naam.localeCompare(b.naam, "nl"));
    return { periodes: periodes.map((p) => p.naam), periodeKolommen: periodes.map((p) => String(p.naam).replace(/\{jaar\}\.?/g, "").trim()), leerlingen };
  }

  const API = { bouw, beoordeel, periodeBereik, naarTabellen, klassen, cijfers, rubrics, rijNiveaus, actieveOpdrachten, rapport, naarCijfer, teamGemiddelde };
  globalThis.TT_ANALYSE = API;
  if (typeof module !== "undefined" && module.exports) module.exports = API;
})();
