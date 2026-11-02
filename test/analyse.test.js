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

console.log("analyse.test.js: alles in orde");
