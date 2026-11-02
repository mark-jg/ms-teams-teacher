/* Zijpaneel: handvat aan de rand van Teams, Alt+K opent/sluit. Bovenin de status van de koppeling
 * (is er een Graph-token met de education-scopes?), daaronder de modules met hun acties. */
TT.register({
  id: "panel",
  _panel: null, _handle: null, _unsub: null, _key: null,

  init(TT) {
    if (!TT.isTop) return;
    const pos = TT.settings.panel.position === "left" ? "tt-left" : "";
    this._handle = TT.el("div", { id: "tt-handle", class: pos, text: "Docent Tools", title: "Openen (Alt+K)", role: "button", tabindex: "0", onclick: () => this.toggle(), onkeydown: (e) => { if (e.key === "Enter" || e.key === " ") this.toggle(); } });
    this._panel = TT.el("div", { id: "tt-panel", class: pos, role: "complementary", "aria-label": "Docent Tools paneel" });
    document.body.append(this._handle, this._panel);
    this._handle.hidden = TT.settings.panel.handle === false;
    this.applyTheme(TT);
    this.render(TT);
    this._unsub = TT.onSettingsChange(() => { this._handle.hidden = TT.settings.panel.handle === false; this.applyTheme(TT); this.render(TT); });
    this._key = (e) => {
      if (e.altKey && !e.ctrlKey && !e.metaKey && (e.code === "KeyK")) { e.preventDefault(); this.toggle(); }
      if (e.altKey && !e.ctrlKey && !e.metaKey && (e.code === "KeyO")) { e.preventDefault(); TT.module("opdrachten")?.open(TT); }
      if (e.key === "Escape" && this._panel.classList.contains("tt-open")) this.toggle(false);
    };
    window.addEventListener("keydown", this._key, true);
    document.addEventListener("tt:command", (e) => { if (e.detail === "toggle-panel") this.toggle(); if (e.detail === "open-overzicht") TT.module("opdrachten")?.open(TT); });
    matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => this.applyTheme(TT));
  },

  destroy() {
    this._panel?.remove(); this._handle?.remove();
    this._unsub && this._unsub();
    window.removeEventListener("keydown", this._key, true);
  },

  applyTheme(TT) { document.documentElement.classList.toggle("tt-dark", TT.dark()); },

  toggle(force) {
    const open = force ?? !this._panel.classList.contains("tt-open");
    this._panel.classList.toggle("tt-open", open);
    if (open) { this.refreshStatus(); this._panel.querySelector("button, a")?.focus(); }
  },

  /* Status van de koppeling. Het bruikbare Graph-token zit in het frame van de app Opdrachten
   * (assignments.edu.cloud.microsoft) en heeft EduAssignments maar geen EduRoster. */
  async refreshStatus() {
    const box = this._panel?.querySelector("#tt-status"); if (!box) return;
    const set = (cls, txt) => { box.className = `tt-status is-${cls}`; box.querySelector("span:last-child").textContent = txt; };
    set("", "Koppeling controleren…");
    const info = await TT.tokensInfo();
    const graph = info.flatMap((f) => f.tokens.filter((t) => /graph\.microsoft\.com/.test(t.host)).map((t) => ({ ...t, frame: f.frame })));
    const edu = graph.filter((t) => /EduAssignments/i.test(t.scp));
    if (!edu.length) return set(graph.length ? "warn" : "err", `${graph.length ? "Wel een Graph-token, maar zonder EduAssignments." : "Nog geen Graph-token gezien."} Open in Teams de app Opdrachten (linkerbalk, eventueel via ⋯ Meer apps); daarin zit het token dat de extensie gebruikt. Controleer daarna opnieuw.`);
    const w = await TT.graphWindow("EduAssignments");
    if (!w) return set("warn", "Token met EduAssignments gezien, maar het frame antwoordt niet. Herlaad het Teams-tabblad en open Opdrachten opnieuw.");
    const kl = await TT.graph("/education/me/classes?$select=id,displayName&$top=1");
    if (kl.status === 200) return set("ok", "Koppeling werkt: education-API met klassenlijst en leden.");
    const jt = await TT.graph("/me/joinedTeams?$select=id,displayName");
    if (jt.status === 200) return set("ok", `Koppeling werkt via de teams-lijst (${jt.value.length} teams): opdrachten en inleveringen via de education-API, klassen en leden via Teams-groepen (het token mist EduRoster, HTTP ${kl.status}).`);
    set("warn", `Token gevonden, maar zowel de klassenlijst (HTTP ${kl.status}) als de teams-lijst (HTTP ${jt.status}) faalt${jt.error ? ": " + jt.error : ""}.`);
  },

  render(TT) {
    const M = globalThis.TT_MODULES_LIST || [];
    const widget = (title, content, hint) => TT.el("section", { class: "tt-widget" }, [TT.el("h3", { text: title }), content, hint ? TT.el("div", { class: "tt-muted", style: "margin-top:6px", text: hint }) : null]);
    const status = TT.el("div", { id: "tt-status", class: "tt-status" }, [TT.el("span", { class: "tt-dot" }), TT.el("span", { text: "Nog niet gecontroleerd." })]);
    const tiles = TT.el("div", { class: "tt-tiles" }, M.map(([id, label, sub, k]) => {
      const on = TT.settings[id]?.enabled !== false;
      const b = TT.el("button", { class: "tt-tile" + (on ? " is-on" : ""), style: `--k:${k}`, "aria-pressed": String(on), title: (on ? "Uitzetten: " : "Aanzetten: ") + label, onclick: () => TT.saveSettings({ [id]: { enabled: !on } }) });
      b.innerHTML = `<span class="tt-tile-ico">${TT_ICONS[id] || ""}</span><span><b>${label}</b><small>${sub}</small></span><span class="tt-knob"></span>`;
      return b;
    }));
    const thema = TT.el("select", { class: "tt-select", "aria-label": "Thema", onchange: (e) => TT.saveSettings({ thema: e.target.value }) }, [["auto", "Volg systeem"], ["licht", "Licht"], ["donker", "Donker"]].map(([v, l]) => TT.el("option", { value: v, text: l })));
    thema.value = TT.settings.thema || "auto";

    const body = TT.el("div", { class: "tt-panel-body" }, [
      widget("Koppeling", TT.el("div", {}, [status, TT.el("div", { class: "tt-row", style: "margin-top:8px" }, [
        TT.el("button", { class: "tt-btn tt-btn-sm", text: "Opnieuw controleren", onclick: () => this.refreshStatus() })
      ])]), "De extensie gebruikt het token dat Teams zelf al heeft; er gaat niets naar een eigen server."),
      widget("Opdrachten over alle teams", TT.el("div", {}, [
        TT.el("div", { id: "tt-opdrachten-widget" }),
        TT.el("div", { class: "tt-row", style: "margin-top:8px" }, [
          TT.el("button", { class: "tt-btn tt-btn-primary", text: "Overzicht openen", onclick: () => TT.module("opdrachten")?.open(TT) }),
          TT.el("button", { class: "tt-btn", text: "Bijwerken", title: "Haalt alleen op wat kan veranderen; Alt+klik haalt alles opnieuw op", onclick: (e) => TT.module("opdrachten")?.bijwerken(TT, { volledig: e.altKey }) }),
          TT.el("span", { class: "tt-kbd", text: "Alt+O" })
        ])
      ]), "Klassen, opdrachten, inleveringen en cijfers via Microsoft Graph. Blijft lokaal bewaard, zodat het overzicht de volgende keer direct opent."),
      widget("Modules", tiles),
      widget("Weergave", TT.el("div", { class: "tt-row" }, [thema, TT.el("button", { class: "tt-btn tt-btn-sm", text: "Instellingen", onclick: () => chrome.runtime.sendMessage({ type: "openOptions" }) })])),
      TT.settings.developer?.enabled ? widget("Ontwikkelaar", TT.el("div", { class: "tt-row" }, [
        TT.el("button", { class: "tt-btn tt-btn-sm", text: "Inspector openen", onclick: () => { this.toggle(false); TT.module("inspector")?.open(TT); } })
      ]), "Frames, tokens (alleen host en scopes) en API-calls van Teams.") : null
    ]);
    this._panel.replaceChildren(
      TT.el("div", { class: "tt-panel-head" }, [
        TT.el("h2", {}, ["Docent Tools", TT.el("small", { text: `v${TT.version} · Teams` })]),
        TT.el("button", { class: "tt-btn tt-btn-sm", text: "Sluiten", "aria-label": "Paneel sluiten", onclick: () => this.toggle(false) })
      ]),
      body
    );
    TT.module("opdrachten")?.renderWidget?.(TT);
  }
});

const TT_ICONS = Object.fromEntries(Object.entries({
  opdrachten: `<rect x="4" y="4" width="16" height="16" rx="3"/><path d="M8 12l3 3 5-6"/>`,
  inspector: `<path d="M4 4h6v2H6v4H4zM14 4h6v6h-2V6h-4zM4 14h2v4h4v2H4zM18 14h2v6h-6v-2h4z"/>`
}).map(([k, d]) => [k, `<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`]));
