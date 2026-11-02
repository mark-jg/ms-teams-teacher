/* Afdrukbare rapporten. Leest het rapportmodel dat het overzicht in Teams heeft klaargezet
 * (chrome.storage.local.ttRapport) en de opmerkingen van de docent (ttOpmerkingen, per leerling en
 * periode bewaard). Eén A4-pagina per leerling; afdrukken via de browser (ook naar PDF). */
const $ = (id) => document.getElementById(id);
const el = (tag, attrs = {}, kids = []) => { const n = document.createElement(tag); for (const [k, v] of Object.entries(attrs)) { if (k === "class") n.className = v; else if (k === "text") n.textContent = v; else if (v != null && v !== false) n.setAttribute(k, v === true ? "" : v); } for (const c of [].concat(kids)) if (c != null) n.append(typeof c === "string" ? document.createTextNode(c) : c); return n; };
const fmtCijfer = (c, afronding) => (c == null ? "" : afronding === "decimaal" ? c.toLocaleString("nl-NL", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : String(Math.round(c)));
let model = null, opmerkingen = {}, timer = null, logo = null;
/* Logo: eigen upload uit de instellingen (lokaal bewaard), anders rapport/ma-logo.jpg in de extensiemap,
 * anders het woordmerk als tekst. Het officiële logo is een merkbestand en zit niet in de repo. */
function logoEl() {
  const img = el("img", { src: logo || "ma-logo.jpg", alt: "Mediacollege Amsterdam" });
  img.addEventListener("error", () => img.replaceWith(el("div", { class: "woordmerk", text: "Mediacollege Amsterdam" })));
  return img;
}

function sleutel(l, periode) { return `${l.studentnummer || l.id}|${periode}`; }

function render() {
  const host = $("paginas"); host.replaceChildren();
  if (!model?.leerlingen?.length) { $("leeg").hidden = false; return; }
  $("leeg").hidden = true;
  const klas = $("klasFilter").value, toonLeeg = $("toonLeeg").checked, toonOpm = $("toonOpm").checked;
  let n = 0;
  for (const l of model.leerlingen) {
    if (klas && l.klas !== klas) continue;
    n++;
    const periodes = l.periodeNamen || model.periodes;
    const titel = (model.titel || "Resultaten jaar {jaar}").replace("{jaar}", l.jaar ?? "").replace("{klas}", l.klas).replace("{cohort}", l.cohort ?? "").replace("{schooljaar}", model.schooljaar || "");
    const pagina = el("article", { class: "pagina" }, [
      el("div", { class: "kop" }, [
        logoEl(),
        el("div", { class: "kop-rechts" }, [
          el("h1", { class: "opleiding" }, [el("small", { text: "Opleiding" }), l.opleiding || l.klas || ""]),
          el("div", { class: "kader", text: titel })
        ])
      ]),
      el("div", { class: "gegevens" }, [
        ["Studentnummer", l.studentnummer || "–"], ["Roepnaam", l.roepnaam], ["Achternaam", l.achternaam],
        ["Klas", l.klas || "–"], ["Cohort", l.cohort ?? "–"], ["Leerjaar", l.jaar ?? "–"]
      ].map(([k, v]) => el("div", {}, [el("small", { text: k }), el("b", { text: String(v ?? "") })]))),
      ...l.groepen.map((g) => {
        const vakken = g.vakken.filter((v) => toonLeeg || v.n > 0);
        if (!vakken.length) return null;
        return el("section", { class: "groep" }, [
          el("h2", {}, el("span", { class: "kader", text: g.naam })),
          el("table", { class: "cijfers" }, [
            el("thead", {}, el("tr", {}, [el("th", { text: "Vak" }), ...periodes.map((p) => el("th", { class: "p", text: p }))])),
            el("tbody", {}, vakken.map((v) => el("tr", {}, [
              el("td", { class: "vak" }, [v.vak, v.team !== v.vak ? el("small", { text: v.team }) : null]),
              ...v.periodes.map((p) => {
                if (p.cijfer == null) return el("td", { class: "p geen", text: "–" });
                return el("td", { class: "p" + (p.cijfer < 5.5 ? " onv" : "") }, [fmtCijfer(p.cijfer, model.afronding), el("span", { class: "n", text: `${p.n} ${p.n === 1 ? "cijfer" : "cijfers"}` })]);
              })
            ])))
          ])
        ]);
      }),
      model.toelichting ? el("div", { class: "toelichting", text: model.toelichting }) : null,
      toonOpm ? el("div", { class: "opmerkingen" }, periodes.map((p) => {
        const box = el("div", { contenteditable: "true", "data-placeholder": `Opmerkingen bij ${p}…`, text: opmerkingen[sleutel(l, p)] || "" });
        box.addEventListener("input", () => { opmerkingen[sleutel(l, p)] = box.innerText.trim(); clearTimeout(timer); timer = setTimeout(() => chrome.storage.local.set({ ttOpmerkingen: opmerkingen }), 400); });
        return el("div", { class: "opm" }, [el("h3", { text: `Opmerkingen bij ${p}` }), box]);
      })) : null,
      el("div", { class: "voet" }, [el("span", { text: `${l.naam} · ${l.klas}` }), el("span", { text: `Gemaakt ${new Date(model.gemaakt).toLocaleDateString("nl-NL", { day: "numeric", month: "long", year: "numeric" })} · ${l.aantalCijfers} cijfers` })])
    ]);
    host.append(pagina);
  }
  $("info").textContent = `${n} ${n === 1 ? "rapport" : "rapporten"} · periodes: ${(model.periodeKolommen || model.periodes).join(", ") || "geen"} · cesuur ${Math.round((model.cesuur || 0.55) * 100)}%`;
  document.title = `Rapporten${klas ? " " + klas : ""} · ${n}`;
}

chrome.storage.local.get(["ttRapport", "ttOpmerkingen", "ttLogo"]).then(({ ttRapport, ttOpmerkingen, ttLogo }) => {
  model = ttRapport || null; opmerkingen = ttOpmerkingen || {}; logo = ttLogo || null;
  const sel = $("klasFilter");
  for (const k of [...new Set((model?.leerlingen || []).map((l) => l.klas))].sort((a, b) => a.localeCompare(b, "nl", { numeric: true }))) sel.append(el("option", { value: k, text: k }));
  render();
});
for (const id of ["klasFilter", "toonLeeg", "toonOpm"]) $(id).addEventListener("change", render);
$("afdrukken").addEventListener("click", () => print());
