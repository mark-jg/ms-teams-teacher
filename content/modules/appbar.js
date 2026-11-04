/* Icoon "Docent Tools" in de linkerbalk van Teams, naast Activity, Chat, Teams en Opdrachten.
 * De balk wordt niet op klassenamen gezocht (die veranderen met elke Teams-versie) maar op de knoppen
 * die er zeker in staan: de kleinste gemeenschappelijke voorouder van minstens drie bekende knoppen.
 * Klikken opent het overzicht als volledige pagina naast de balk. Lukt het niet, dan blijft het
 * handvat aan de rechterrand en Alt+O gewoon werken. */
TT.register({
  id: "appbar",
  _stop: null, _knop: null, _item: null, _balk: null, _gemeld: false,

  init(TT) {
    if (!TT.isTop) return;
    this._stop = TT.observe(() => this.mount(TT), { debounce: 600, poll: 3000 });
  },
  destroy() { this._stop && this._stop(); this._item?.remove(); this._knop = this._item = this._balk = null; },

  /* Bekende knoppen in de app-balk (Engels en Nederlands). */
  vindBalk(TT) {
    const namen = /^(activity|activiteit|chat|teams|assignments|opdrachten|calendar|agenda|calls|oproepen|onedrive|insights|apps|meer|more)\b/i;
    const tekst = (el) => (el.getAttribute("aria-label") || el.getAttribute("title") || el.textContent || "").trim().replace(/\s+/g, " ");
    const kandidaten = TT.deepQueryAll('button, a[href], [role="tab"], [role="button"]').filter((el) => {
      const r = el.getBoundingClientRect();
      return r.width > 24 && r.width < 140 && r.left < 160 && namen.test(tekst(el)) && !el.closest("#tt-overlay, #tt-panel");
    });
    if (kandidaten.length < 3) return null;
    // Kleinste voorouder die minstens drie kandidaten bevat.
    let node = kandidaten[0];
    while (node && node !== document.body) {
      const n = kandidaten.filter((k) => node.contains(k)).length;
      if (n >= 3) return { balk: node, knoppen: kandidaten.filter((k) => node.contains(k)) };
      node = node.parentElement;
    }
    return null;
  },

  mount(TT) {
    if (this._item?.isConnected) { this.meetBreedte(); return; }
    const gevonden = this.vindBalk(TT);
    if (!gevonden) { if (!this._gemeld) { this._gemeld = true; console.info("[Docent Tools] app-balk nog niet gevonden; het handvat en Alt+O blijven werken"); } return; }
    const { balk, knoppen } = gevonden;
    this._balk = balk;
    // Het item is de directe nakomeling van de balk waarin de laatste bekende knop staat; ons item wordt
    // een broer daarvan, zodat het dezelfde ruimte krijgt.
    const laatste = knoppen[knoppen.length - 1];
    let voorbeeld = laatste; while (voorbeeld.parentElement && voorbeeld.parentElement !== balk) voorbeeld = voorbeeld.parentElement;
    const knop = TT.el("button", { class: "tt-appbar-knop", type: "button", "aria-label": "Docent Tools", title: "Docent Tools: overzicht over alle teams (Alt+O)", onclick: (e) => { e.preventDefault(); e.stopPropagation(); TT.guard(() => TT.module("opdrachten")?.open(TT, { vol: true }), "Docent Tools"); } });
    knop.append(TT.el("img", { src: chrome.runtime.getURL("icons/icon32.png"), alt: "", width: "24", height: "24", draggable: "false" }), TT.el("span", { text: "Docent Tools" }));
    const item = voorbeeld === laatste ? knop : TT.el(voorbeeld.tagName.toLowerCase() === "li" ? "li" : "div", { class: "tt-appbar-item", role: voorbeeld.getAttribute("role") || null }, knop);
    voorbeeld.after(item);
    this._knop = knop; this._item = item;
    this.meetBreedte();
    console.info("[Docent Tools] icoon in de app-balk geplaatst");
  },

  /* Breedte van de balk, zodat de volledige pagina er precies naast begint. */
  meetBreedte() {
    try { const r = this._balk?.getBoundingClientRect(); if (r && r.width > 0 && r.width < 200) document.documentElement.style.setProperty("--tt-appbar-w", `${Math.round(r.right)}px`); } catch {}
  },
  actief(aan) { this._knop?.classList.toggle("is-on", !!aan); this._knop?.setAttribute("aria-pressed", String(!!aan)); }
});
