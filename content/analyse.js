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
    // 2. Overlap in lidmaatschap (alleen teams zonder herkende code, of als er geen patroon is).
    const los = teams.filter((t) => !code(t));
    for (let i = 0; i < los.length; i++) for (let j = i + 1; j < los.length; j++) {
      const A = leden.get(los[i].id), B = leden.get(los[j].id);
      if (!A.size || !B.size) continue;
      let n = 0; for (const x of A) if (B.has(x)) n++;
      if (n >= minL && n / Math.min(A.size, B.size) >= drempel) voeg(los[i].id, los[j].id);
    }
    // Groepen samenstellen.
    const groepen = new Map();
    for (const t of teams) { const r = vind(t.id); if (!groepen.has(r)) groepen.set(r, []); groepen.get(r).push(t); }
    const STOP = /^(team|klas|class|groep|vak|les|en|de|het|van|\d{4}|\d{2}-\d{2}|\d{4}-\d{4}|\d{4}-\d{2})$/i;
    const tokens = (n) => String(n || "").split(/[^\p{L}\p{N}]+/u).filter((x) => x.length >= 2 && !STOP.test(x));
    const benoem = (ts) => {
      const c = code(ts[0]); if (c && ts.every((t) => code(t) === c)) return c;
      if (ts.length === 1) return ts[0].displayName;
      const tel = new Map();
      for (const t of ts) for (const x of new Set(tokens(t.displayName).map((y) => y.toUpperCase()))) tel.set(x, (tel.get(x) || 0) + 1);
      const gemeen = [...tel.entries()].filter(([, n]) => n >= Math.max(2, Math.ceil(ts.length * 0.6))).map(([x]) => x);
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

  const API = { bouw, beoordeel, periodeBereik, naarTabellen, klassen };
  globalThis.TT_ANALYSE = API;
  if (typeof module !== "undefined" && module.exports) module.exports = API;
})();
