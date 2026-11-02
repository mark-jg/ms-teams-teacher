/* Standaardinstellingen, gedeeld door core.js (content), background.js, de popup en de instellingenpagina.
 * Eén plek bewerken is genoeg. */
globalThis.TT_DEFAULTS = {
  panel: { enabled: true, position: "right", handle: true },
  opdrachten: { enabled: true, periode: "schooljaar", drempelOntbreekt: 2, drempelTeLaat: 3, cacheUren: 24, volledigDagen: 7, schooljaren: 2, gearchiveerd: false, alleTeams: false, klascodePatroon: "" }, // leeg = standaardpatroon (TT_KLASCODE), klascode vóór het streepje
  inspector: { enabled: true },
  developer: { enabled: false },
  updates: { check: true },
  thema: "auto" // auto | licht | donker
};

/* Modules met een aan/uit-schakelaar in het zijpaneel. */
globalThis.TT_MODULES_LIST = [
  ["opdrachten", "Opdrachten", "Over alle teams heen", "#5B5FC7"],
  ["inspector", "Inspector", "Frames, tokens, API-calls", "#8E8E93"]
];

/* Standaardpatroon voor de klascode in een teamnaam: letters + jaar/cohort + optionele letter, vóór een
 * streepje of aan het eind. Herkent MIP26A, MR26B, MEP25A, KD25, MEPMR25; "KD - Verdieping" krijgt geen code
 * en sluit via overlap in lidmaatschap bij een klas aan. Leeg laten in de instellingen = dit patroon. */
globalThis.TT_KLASCODE = "^\\s*([A-Za-z]{1,8}\\d{2,4}[A-Za-z]{0,2})\\s*(?:[-–:]|$)";
