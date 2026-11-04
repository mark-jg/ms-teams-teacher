/* Standaardinstellingen, gedeeld door core.js (content), background.js, de popup en de instellingenpagina.
 * Eén plek bewerken is genoeg. */
globalThis.TT_DEFAULTS = {
  panel: { enabled: true, position: "right", handle: true },
  opdrachten: { enabled: true, periode: "schooljaar", drempelOntbreekt: 2, drempelTeLaat: 3, cacheUren: 24, volledigDagen: 7, schooljaren: 2, weergave: "auto", gearchiveerd: false, alleTeams: false, klascodePatroon: "" }, // leeg = standaardpatroon (TT_KLASCODE), klascode vóór het streepje
  /* Rapporten: periodes "naam | begin | einde" per regel; opleidingen "CODE = Naam"; vakgroepen "Groep: vak, vak". */
  rapport: { ingesteld: false, thema: "magenta", school: "Mediacollege Amsterdam", periodes: "", opleidingen: "", vakgroepen: "", cesuur: 55, afronding: "heel", titel: "Resultaten jaar {jaar}", toelichting: "" },
  appbar: { enabled: true },
  inspector: { enabled: true },
  developer: { enabled: false },
  updates: { check: true },
  thema: "auto" // auto | licht | donker
};

/* Modules met een aan/uit-schakelaar in het zijpaneel. */
globalThis.TT_MODULES_LIST = [
  ["opdrachten", "Opdrachten", "Over alle teams heen", "#5B5FC7"],
  ["appbar", "Icoon in de linkerbalk", "Docent Tools naast Opdrachten", "#007AFF"],
  ["inspector", "Inspector", "Frames, tokens, API-calls", "#8E8E93"]
];

/* Standaardpatroon voor de klascode in een teamnaam: letters + jaar/cohort + optionele letter, vóór een
 * streepje of aan het eind. Herkent MIP26A, MR26B, MEP25A, KD25, MEPMR25; "KD - Verdieping" krijgt geen code
 * en sluit via overlap in lidmaatschap bij een klas aan. Leeg laten in de instellingen = dit patroon. */
globalThis.TT_KLASCODE = "^\\s*([A-Za-z]{1,8}\\d{2,4}[A-Za-z]{0,2})\\s*(?:[-–:]|$)";

/* Standaardsjabloon voor rapporten: geldt zolang een veld leeg is. Periodes volgen het schooljaar
 * (datums zonder jaar: MM-DD, de extensie vult het lopende schooljaar in); {jaar} is het leerjaar van
 * de leerling, zodat een eerstejaars "Periode 1.1" ziet en een tweedejaars "Periode 2.1". */
globalThis.TT_RAPPORT_STANDAARD = {
  periodes: "Periode {jaar}.1 | 08-01 | 10-31\nPeriode {jaar}.2 | 11-01 | 01-31\nPeriode {jaar}.3 | 02-01 | 04-15\nPeriode {jaar}.4 | 04-16 | 07-31",
  opleidingen: "MR = Mediaredactiemedewerker\nMIP = Music Industry Professional",
  groepen: ["Praktijkvakken", "AVO-vakken en BUR"],
  /* Vakken die op naam als AVO of BUR herkend worden; de rest is praktijk. */
  avo: /rekenen|engels|nederlands|burgerschap|\bbur\b|loopbaan|\bslb\b|duits|frans|spaans|wiskunde|keuzedeel|lob/i,
  toelichting: "De student krijgt in periode 2 een positief studieadvies als: minimaal vier van de vijf praktijkvakken zijn behaald met een voldoende, minimaal vier van de vijf AVO-vakken/BUR zijn behaald met een voldoende en de aanwezigheid 80% of meer is. Als de student niet voldoet aan deze voorwaarden, dan wordt het studieadvies door het docententeam bepaald."
};
/* Thema's voor de rapporten: [id, naam, omschrijving]. */
globalThis.TT_RAPPORT_THEMAS = [
  ["magenta", "Zwart/magenta", "Huisstijl Mediacollege Amsterdam, met logo"],
  ["blauw", "Blauw", "Algemeen, rustig blauw"],
  ["groen", "Groen", "Algemeen, groen"],
  ["grijs", "Grijs", "Algemeen, neutraal zonder accentkleur"],
  ["warm", "Warm", "Algemeen, terracotta"]
];
