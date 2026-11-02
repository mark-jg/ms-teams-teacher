/* node test/analyse.test.js */
const assert = require("node:assert/strict");
const A = require("../content/analyse.js");

const nu = new Date("2026-11-02T10:00:00Z").getTime();
const d = (s) => new Date(s).toISOString();
const data = {
  classes: [{ id: "k1", displayName: "4A Engels", classCode: "EN4A" }, { id: "k2", displayName: "4B Engels", classCode: "EN4B" }],
  members: { k1: [{ id: "u1", displayName: "Aisha" }, { id: "u2", displayName: "Bram" }], k2: [{ id: "u1", displayName: "Aisha" }, { id: "u3", displayName: "Chris" }] },
  assignments: [
    { id: "a1", classId: "k1", displayName: "Essay 1", status: "assigned", dueDateTime: d("2026-10-20T22:00:00Z"), assignedDateTime: d("2026-10-01T08:00:00Z"), grading: { "@odata.type": "#microsoft.graph.educationAssignmentPointsGradeType", maxPoints: 10 } },
    { id: "a2", classId: "k1", displayName: "Essay 2", status: "assigned", dueDateTime: d("2026-11-10T22:00:00Z"), assignedDateTime: d("2026-10-25T08:00:00Z"), grading: { "@odata.type": "#microsoft.graph.educationAssignmentPointsGradeType", maxPoints: 20 } },
    { id: "a3", classId: "k2", displayName: "Listening", status: "assigned", dueDateTime: d("2026-10-28T22:00:00Z"), assignedDateTime: d("2026-10-10T08:00:00Z"), grading: null },
    { id: "a4", classId: "k2", displayName: "Concept", status: "draft", dueDateTime: d("2026-12-01T22:00:00Z") },
    { id: "a5", classId: "k1", displayName: "Vorig jaar", status: "assigned", dueDateTime: d("2025-05-01T22:00:00Z"), assignedDateTime: d("2025-04-01T08:00:00Z") }
  ],
  submissions: {
    a1: [
      { id: "s1", status: "returned", submittedDateTime: d("2026-10-19T12:00:00Z"), returnedDateTime: d("2026-10-22T12:00:00Z"), recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 8 }, publishedPoints: { points: 8 } }] },
      { id: "s2", status: "submitted", submittedDateTime: d("2026-10-23T12:00:00Z"), recipient: { userId: "u2" }, outcomes: [] }
    ],
    a2: [
      { id: "s3", status: "working", recipient: { userId: "u1" } },
      { id: "s4", status: "working", recipient: { userId: "u2" } }
    ],
    a3: [
      { id: "s5", status: "working", recipient: { userId: "u1" } },
      { id: "s6", status: "excused", excusedDateTime: d("2026-10-27T12:00:00Z"), recipient: { userId: "u3" } }
    ],
    a5: [{ id: "s7", status: "working", recipient: { userId: "u2" } }]
  }
};

const r = A.bouw(data, { periode: "schooljaar", nu, drempelOntbreekt: 1, drempelTeLaat: 1 });

// Concept (draft) en de opdracht van vorig schooljaar vallen weg.
assert.equal(r.totaal.opdrachten, 3);
assert.deepEqual(r.klassen.map((k) => k.naam), ["4A Engels", "4B Engels"]);

const k1 = r.klassen[0];
assert.equal(k1.leerlingen, 2);
assert.equal(k1.toegewezen, 4);
assert.equal(k1.ingeleverd, 2);
assert.equal(k1.teLaat, 1);          // Bram leverde Essay 1 na de deadline in
assert.equal(k1.nakijken, 1);        // Bram wacht op nakijken
assert.equal(k1.open, 2);            // Essay 2 nog niet over tijd
assert.equal(k1.ontbreekt, 0);
assert.equal(k1.gemiddeld, 0.8);     // alleen Aisha heeft punten: 8/10

const k2 = r.klassen[1];
assert.equal(k2.ontbreekt, 1);       // Aisha, Listening
assert.equal(k2.vrijgesteld, 1);     // Chris

const aisha = r.leerlingen.find((l) => l.naam === "Aisha");
assert.deepEqual(aisha.klassen, ["4A Engels", "4B Engels"]);
assert.equal(aisha.aantalKlassen, 2);
assert.equal(aisha.toegewezen, 3);
assert.equal(aisha.ontbreekt, 1);
assert.equal(aisha.open, 1);
assert.equal(aisha.nagekeken, 1);
assert.equal(aisha.perKlas.length, 2);

// Signalen: Aisha (ontbreekt) en Bram (te laat), beide boven drempel 1.
assert.deepEqual(r.signalen.map((s) => s.naam).sort(), ["Aisha", "Bram"]);
assert.deepEqual(r.signalen.find((s) => s.naam === "Aisha").klassenMetProbleem, ["4B Engels"]);

// Deadlines: alleen Essay 2 ligt in de toekomst (binnen 1 dag terugkijken).
assert.deepEqual(r.deadlines.map((o) => o.naam), ["Essay 2"]);
assert.equal(r.deadlines[0].dagen, 9);

// Nakijkwerk: Essay 1 met 1 wachtende, al 9 volle dagen (23 okt 12:00 tot 2 nov 10:00).
assert.equal(r.nakijken.length, 1);
assert.equal(r.nakijken[0].nakijken, 1);
assert.equal(r.nakijken[0].wachtDagen, 9);

// Periodefilter "alles" neemt de oude opdracht mee; klasfilter beperkt tot k2.
assert.equal(A.bouw(data, { periode: "alles", nu }).totaal.opdrachten, 4);
assert.equal(A.bouw(data, { periode: "alles", nu, klassen: new Set(["k2"]) }).totaal.opdrachten, 1);

// Tabellen: zes bladen met kopregel en rijen.
const tabs = A.naarTabellen(r);
assert.equal(tabs.length, 6);
assert.equal(tabs[0][1].length, 3);
assert.equal(tabs[1][1][0][0], "Leerling");

// Schooljaarbereik: op 2 nov 2026 loopt het van 1 aug 2026 tot 31 jul 2027.
const [van, tot] = A.periodeBereik("schooljaar", nu);
assert.equal(new Date(van).getMonth(), 7);
assert.equal(new Date(tot).getFullYear(), 2027);

// Klassen afleiden: team per vak, klas = leerlingen die samen in dezelfde teams zitten.
const sch = {
  classes: [
    { id: "t1", displayName: "Engels 4A 2526" }, { id: "t2", displayName: "Wiskunde 4A" }, { id: "t3", displayName: "Nederlands 4A" },
    { id: "t4", displayName: "Engels 4B" }, { id: "t5", displayName: "Wiskunde 4B" }, { id: "t6", displayName: "Keuzevak Film" }
  ],
  members: {
    t1: ["a", "b", "c", "d"].map((x) => ({ id: x, displayName: x.toUpperCase() })),
    t2: ["a", "b", "c", "d", "z"].map((x) => ({ id: x, displayName: x.toUpperCase() })),
    t3: ["a", "b", "c"].map((x) => ({ id: x, displayName: x.toUpperCase() })),
    t4: ["e", "f", "g", "h"].map((x) => ({ id: x, displayName: x.toUpperCase() })),
    t5: ["e", "f", "g", "h", "a"].map((x) => ({ id: x, displayName: x.toUpperCase() })),
    t6: ["a", "e", "q"].map((x) => ({ id: x, displayName: x.toUpperCase() }))
  },
  assignments: [], submissions: {}, namen: {}
};
const K = A.klassen(sch);
assert.deepEqual(K.klassen.map((k) => k.naam), ["4A", "4B", "Keuzevak Film"]);
assert.deepEqual(K.klassen[0].teams.map((t) => t.id), ["t1", "t3", "t2"]); // op naam: Engels, Nederlands, Wiskunde
assert.deepEqual(K.klassen[1].teams.map((t) => t.id), ["t4", "t5"]);
assert.equal(K.klasVan.a, "4A");            // a zit in 3 teams van 4A en 1 van 4B
assert.equal(K.klasVan.e, "4B");
assert.equal(K.klasVan.z, "4A");
assert.equal(K.klasVan.q, "Keuzevak Film");
assert.equal(K.teamKlas.t5, "4B");
assert.equal(K.klassen[0].leerlingen.length, 5); // a, b, c, d, z
// Met een patroon voor de klascode gaat de naam en de groepering op de code.
const K2 = A.klassen({ ...sch, classes: [{ id: "t1", displayName: "EN-4A" }, { id: "t2", displayName: "WI-4A" }, { id: "t4", displayName: "EN-4B" }] }, { patroon: "\\b(\\d[A-Z])\\b" });
assert.deepEqual(K2.klassen.map((k) => [k.naam, k.teams.length]), [["4A", 2], ["4B", 1]]);
// Klas komt terug in de leerlingrijen van bouw().
const r2 = A.bouw(data, { periode: "schooljaar", nu, klasVan: { u1: "4A" } });
assert.equal(r2.leerlingen.find((l) => l.id === "u1").klas, "4A");

// Laatste 2 schooljaren: op 2 nov 2026 van 1 aug 2025 tot 31 jul 2027; de oude opdracht (mei 2025) valt er net buiten.
const [van2] = A.periodeBereik("sj2", nu);
assert.equal(new Date(van2).getFullYear(), 2025);
assert.equal(A.bouw(data, { periode: "sj2", nu }).totaal.opdrachten, 3);
assert.equal(new Date(A.periodeBereik("sj3", nu)[0]).getFullYear(), 2024);

// Naamgeving met klascode vóór het streepje (team per vak), zoals op het Mediacollege.
const PAT = "^\\s*([A-Za-z]{1,8}\\d{2,4}[A-Za-z]{0,2})\\s*(?:[-–:]|$)";
const mc = {
  classes: [
    { id: "a", displayName: "MIP26A - Burgerschap" }, { id: "b", displayName: "MIP26A - Rekenen" }, { id: "c", displayName: "MIP26A" },
    { id: "d", displayName: "MIP26B - Burgerschap" }, { id: "e", displayName: "MIP26B" },
    { id: "f", displayName: "KD25 - Creatief Ondernemerschap" }, { id: "g", displayName: "KD - Verdieping Beeldgebruik" },
    { id: "h", displayName: "MEPMR25 - KD Basis Fotografie" }
  ],
  members: {
    a: [], // nog geen leden bekend: hoort toch bij MIP26A via de code
    b: ["p", "q", "r", "s"].map((x) => ({ id: x, displayName: x })), c: ["p", "q", "r", "s", "t"].map((x) => ({ id: x, displayName: x })),
    d: ["u", "v", "w"].map((x) => ({ id: x, displayName: x })), e: ["u", "v", "w", "x"].map((x) => ({ id: x, displayName: x })),
    f: ["k", "l", "m", "n"].map((x) => ({ id: x, displayName: x })), g: ["k", "l", "m"].map((x) => ({ id: x, displayName: x })), // geen code, sluit via overlap aan bij KD25
    h: ["p", "u", "k"].map((x) => ({ id: x, displayName: x }))
  }, assignments: [], submissions: {}, namen: {}
};
const M = A.klassen(mc, { patroon: PAT });
assert.deepEqual(M.klassen.map((k) => [k.naam, k.teams.length, k.leerlingen.length]), [["KD25", 2, 4], ["MEPMR25", 1, 0], ["MIP26A", 3, 5], ["MIP26B", 2, 4]]);
assert.equal(M.teamKlas.a, "MIP26A");
assert.equal(M.teamKlas.g, "KD25");
assert.equal(M.klasVan.p, "MIP26A");

// Cijfers met gewogen categorieën: toetsen 70%, huiswerk 30%.
const cd = {
  classes: [{ id: "k1", displayName: "MIP26A - Engels" }, { id: "k2", displayName: "MIP26A - Rekenen" }],
  members: { k1: [{ id: "u1", displayName: "Aisha" }], k2: [{ id: "u1", displayName: "Aisha" }] },
  categorieen: { k1: [{ id: "c1", naam: "Toetsen", gewicht: 70 }, { id: "c2", naam: "Huiswerk", gewicht: 30 }], k2: [] },
  assignments: [
    { id: "a1", classId: "k1", displayName: "Toets 1", status: "assigned", dueDateTime: d("2026-10-01T00:00:00Z"), grading: { maxPoints: 10 }, categorie: { id: "c1", naam: "Toetsen", gewicht: 70 } },
    { id: "a2", classId: "k1", displayName: "Huiswerk 1", status: "assigned", dueDateTime: d("2026-10-02T00:00:00Z"), grading: { maxPoints: 10 }, categorie: { id: "c2", naam: "Huiswerk", gewicht: 30 } },
    { id: "a3", classId: "k1", displayName: "Los", status: "assigned", dueDateTime: d("2026-10-03T00:00:00Z"), grading: { maxPoints: 10 }, categorie: null },
    { id: "a4", classId: "k2", displayName: "Rekentoets", status: "assigned", dueDateTime: d("2026-10-04T00:00:00Z"), grading: { maxPoints: 20 }, categorie: null }
  ],
  submissions: {
    a1: [{ id: "s1", status: "returned", submittedDateTime: d("2026-09-30T00:00:00Z"), recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 6 } }] }],
    a2: [{ id: "s2", status: "returned", submittedDateTime: d("2026-10-01T00:00:00Z"), recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 10 } }] }],
    a3: [{ id: "s3", status: "returned", submittedDateTime: d("2026-10-02T00:00:00Z"), recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 2 } }] }],
    a4: [{ id: "s4", status: "returned", submittedDateTime: d("2026-10-03T00:00:00Z"), recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 15 } }] }]
  }, namen: {}
};
const C = A.cijfers(cd, { periode: "schooljaar", nu });
const ai = C.leerlingen[0];
assert.equal(C.teams.length, 2);
assert.equal(ai.teams.k1.n, 3);
assert.ok(Math.abs(ai.teams.k1.gemiddeld - 0.6) < 1e-9);       // (0.6 + 1.0 + 0.2) / 3
assert.ok(Math.abs(ai.teams.k1.gewogen - 0.72) < 1e-9);         // 0.7 × 0.6 + 0.3 × 1.0; "Los" telt niet mee in de weging
assert.equal(ai.teams.k1.categorieen[0].naam, "Toetsen");
assert.equal(ai.teams.k2.gewogen, null);                        // geen wegingen → alleen gewoon gemiddelde
assert.equal(ai.teams.k2.eind, 0.75);
assert.ok(Math.abs(ai.totaal - (0.72 + 0.75) / 2) < 1e-9);

// Rubrics: rij "Samenwerken" komt in twee teams voor; niveaus met punten.
const niv = [{ id: "l1", naam: "Onvoldoende", punten: 1 }, { id: "l2", naam: "Voldoende", punten: 2 }, { id: "l3", naam: "Goed", punten: 4 }];
const rd = {
  ...cd,
  rubrics: {
    a1: { id: "r1", naam: "Rubric Engels", niveaus: niv, rijen: [{ id: "q1", tekst: "Samenwerken", gewicht: null, criteria: [] }, { id: "q2", tekst: "Spelling", gewicht: null, criteria: [] }] },
    a4: { id: "r2", naam: "Rubric Rekenen", niveaus: niv, rijen: [{ id: "q9", tekst: "samenwerken.", gewicht: null, criteria: [] }] }
  },
  submissions: {
    a1: [{ id: "s1", status: "returned", recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationRubricOutcome", niveaus: [{ qualityId: "q1", columnId: "l3" }, { qualityId: "q2", columnId: "l1" }] }] }],
    a4: [{ id: "s4", status: "returned", recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationRubricOutcome", niveaus: [{ qualityId: "q9", columnId: "l2" }] }] }]
  }
};
const R = A.rubrics(rd, { periode: "schooljaar", nu });
assert.deepEqual(R.rijen.map((r) => [r.tekst, r.n, r.aantalTeams, r.gedeeld]), [["Samenwerken", 2, 2, true], ["Spelling", 1, 1, false]]);
assert.equal(R.rubrics.length, 2);
const N = A.rijNiveaus(rd, R.rijen[0]);
assert.equal(N[0].naam, "Aisha");
assert.deepEqual(N[0].beoordelingen.map((b) => [b.klas, b.niveau, b.pct]), [["MIP26A - Rekenen", "Voldoende", 0.5], ["MIP26A - Engels", "Goed", 1]]);
assert.equal(N[0].gemiddeld, 0.75);

// Cijfer uit percentage: cesuur 55% = 5,5, lineair.
assert.equal(A.naarCijfer(0.55), 6);                                  // heel: 5,5 → 6
assert.equal(A.naarCijfer(0.55, { afronding: "decimaal" }), 5.5);
assert.equal(A.naarCijfer(1, { afronding: "decimaal" }), 10);
assert.equal(A.naarCijfer(0, { afronding: "decimaal" }), 1);
assert.equal(A.naarCijfer(0.275, { afronding: "decimaal" }), 3.3);     // halverwege onder de cesuur: 1 + 2,25
assert.equal(A.naarCijfer(null), null);

// Rapport: klas MR26A met twee vakteams, twee periodes, studentnummer uit e-mail, cohort/leerjaar uit de klascode.
const rp = {
  classes: [{ id: "t1", displayName: "MR26A - Journalistiek" }, { id: "t2", displayName: "MR26A - Tekst" }, { id: "t3", displayName: "MR25A - Tekst" }],
  members: { t1: [{ id: "u1", displayName: "Julia Ypey", mail: "40887@student.ma-web.nl", givenName: "Julia", surname: "Ypey" }], t2: [{ id: "u1", displayName: "Julia Ypey", mail: "40887@student.ma-web.nl" }], t3: [{ id: "u2", displayName: "Bram Bos", mail: "40001@student.ma-web.nl" }] },
  categorieen: { t1: [{ id: "c1", naam: "Toetsen", gewicht: 100 }] },
  assignments: [
    { id: "a1", classId: "t1", displayName: "Artikel 1", status: "assigned", dueDateTime: d("2026-09-15T00:00:00Z"), grading: { maxPoints: 10 }, categorie: { id: "c1" } },
    { id: "a2", classId: "t1", displayName: "Artikel 2", status: "assigned", dueDateTime: d("2026-09-20T00:00:00Z"), grading: { maxPoints: 10 }, categorie: null },
    { id: "a3", classId: "t2", displayName: "Tekst 1", status: "assigned", dueDateTime: d("2026-11-01T00:00:00Z"), grading: { maxPoints: 20 } }
  ],
  submissions: {
    a1: [{ id: "s1", status: "returned", recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 8 } }] }],
    a2: [{ id: "s2", status: "returned", recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 2 } }] }],
    a3: [{ id: "s3", status: "returned", recipient: { userId: "u1" }, outcomes: [{ "@odata.type": "#microsoft.graph.educationPointsOutcome", points: { points: 11 } }] }]
  }, namen: {}
};
const Kr = A.klassen(rp, { patroon: PAT });
const R2 = A.rapport(rp, { nu, periodes: [{ naam: "Periode 1.1", begin: "2026-08-24", einde: "2026-10-31" }, { naam: "Periode 1.2", begin: "2026-11-01", einde: "2027-01-31" }, { naam: "Periode 1.3", begin: "2027-02-01", einde: "2027-04-15" }],
  klasVan: Kr.klasVan, teamKlas: Kr.teamKlas, patroon: PAT, opleidingen: { MR: "Mediaredactiemedewerker" }, vakgroepen: [{ naam: "Praktijkvakken", vakken: ["Journalistiek", "Tekst"] }], afronding: "decimaal" });
assert.deepEqual(R2.periodes, ["Periode 1.1", "Periode 1.2"]);        // periode 1.3 is nog niet begonnen
const ju = R2.leerlingen.find((l) => l.id === "u1");
assert.equal(ju.studentnummer, "40887");
assert.equal(ju.roepnaam, "Julia"); assert.equal(ju.achternaam, "Ypey");
assert.equal(ju.klas, "MR26A"); assert.equal(ju.opleiding, "Mediaredactiemedewerker"); assert.equal(ju.cohort, 2026); assert.equal(ju.jaar, 1);
assert.equal(R2.leerlingen.find((l) => l.id === "u2").jaar, 2);
assert.deepEqual(ju.groepen.map((g) => g.naam), ["Praktijkvakken"]);
const jour = ju.vakken.find((v) => v.vak === "Journalistiek");
assert.equal(jour.periodes[0].cijfer, 8);                              // alleen de gewogen categorie telt: 8/10 → 0,8 → 5,5 + 4,5 × 0,25/0,45 = 8,0
assert.equal(jour.periodes[1].cijfer, null);
assert.equal(ju.vakken.find((v) => v.vak === "Tekst").periodes[1].cijfer, 5.5);  // 11/20 = 55% → cesuur

console.log("analyse.test.js: alles in orde");
