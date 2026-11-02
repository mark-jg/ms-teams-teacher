const DEFAULTS = globalThis.TT_DEFAULTS;
let settings;
const get = (obj, path) => path.split(".").reduce((o, k) => (o ? o[k] : undefined), obj);
const set = (obj, path, val) => { const ks = path.split("."); const last = ks.pop(); const o = ks.reduce((o, k) => (o[k] ??= {}), obj); o[last] = val; };
const merge = (base, patch) => { for (const [k, v] of Object.entries(patch || {})) { if (v && typeof v === "object" && !Array.isArray(v) && base[k] && typeof base[k] === "object" && !Array.isArray(base[k])) merge(base[k], v); else base[k] = v; } return base; };

async function load() {
  const stored = (await chrome.storage.sync.get("settings")).settings || {};
  settings = merge(structuredClone(DEFAULTS), stored);
  for (const el of document.querySelectorAll("[data-path]")) {
    const v = get(settings, el.dataset.path);
    if (el.type === "checkbox") el.checked = v === true || (v !== false && !["opdrachten.gearchiveerd", "opdrachten.alleTeams"].includes(el.dataset.path));
    else el.value = v ?? "";
  }
  dataInfo();
}
async function dataInfo() {
  const { ttData } = await chrome.storage.local.get("ttData");
  const out = document.getElementById("dataInfo");
  if (!ttData) return (out.textContent = "Nog geen gegevens opgehaald.");
  const n = Object.values(ttData.submissions || {}).reduce((a, s) => a + s.length, 0);
  out.textContent = `${ttData.classes?.length || 0} klassen, ${ttData.assignments?.length || 0} opdrachten, ${n} inleveringen; opgehaald ${new Date(ttData.at).toLocaleString("nl-NL")}.`;
}
let t;
async function save() {
  await chrome.storage.sync.set({ settings });
  const s = document.getElementById("status");
  s.classList.add("show"); clearTimeout(t); t = setTimeout(() => s.classList.remove("show"), 1200);
}
document.addEventListener("input", (e) => {
  const el = e.target; if (!el.dataset?.path) return;
  let v;
  if (el.type === "checkbox") v = el.checked;
  else if (el.dataset.type === "int") { v = parseInt(el.value, 10); if (isNaN(v)) return; }
  else v = el.value;
  set(settings, el.dataset.path, v);
  save();
});
document.getElementById("wisData").onclick = async () => {
  if (!confirm("Lokaal opgeslagen klassen, opdrachten en inleveringen wissen? Ze worden bij het volgende openen opnieuw opgehaald.")) return;
  await chrome.storage.local.remove("ttData"); dataInfo();
};
document.getElementById("reset").onclick = async () => {
  if (!confirm("Alle instellingen terugzetten naar standaard?")) return;
  settings = structuredClone(DEFAULTS); await save(); load();
};
document.getElementById("versionLine").textContent = `Versie ${chrome.runtime.getManifest().version}`;

function showUpdateStatus(info) {
  const out = document.getElementById("updateStatus");
  if (!info) out.textContent = "Controleren mislukt.";
  else if (info.error) out.textContent = `Kon GitHub niet bereiken: ${info.error}.`;
  else if (info.available) out.textContent = `Versie ${info.latest} beschikbaar.`;
  else if (!info.latest) out.textContent = "Toegang in orde, er is nog geen release gepubliceerd.";
  else out.textContent = `Je hebt de nieuwste versie (${info.current}).`;
}
document.getElementById("checkUpdate").onclick = async () => {
  document.getElementById("updateStatus").textContent = "Controleren…";
  showUpdateStatus(await chrome.runtime.sendMessage({ type: "checkUpdate" }).catch(() => null));
};
const tokenField = document.getElementById("ghToken");
chrome.storage.local.get("ghToken").then(({ ghToken }) => { tokenField.value = ghToken || ""; });
let tokenTimer;
tokenField.oninput = () => {
  clearTimeout(tokenTimer);
  tokenTimer = setTimeout(async () => {
    document.getElementById("updateStatus").textContent = "Token opgeslagen, controleren…";
    showUpdateStatus(await chrome.runtime.sendMessage({ type: "setToken", token: tokenField.value }).catch(() => null));
  }, 600);
};
document.getElementById("ghTokenClear").onclick = async () => {
  tokenField.value = "";
  document.getElementById("updateStatus").textContent = "Token gewist.";
  await chrome.runtime.sendMessage({ type: "setToken", token: "" }).catch(() => null);
};
load();

/* Logo voor het rapport: als data-URL in chrome.storage.local (alleen dit apparaat). */
const logoVeld = document.getElementById("logoBestand"), logoStatus = document.getElementById("logoStatus");
chrome.storage.local.get("ttLogo").then(({ ttLogo }) => { logoStatus.textContent = ttLogo ? "Eigen logo ingesteld." : "Geen eigen logo gekozen."; });
logoVeld.onchange = () => {
  const f = logoVeld.files?.[0]; if (!f) return;
  if (f.size > 600000) { logoStatus.textContent = "Bestand is groter dan 600 kB; kies een kleinere versie."; return; }
  const r = new FileReader();
  r.onload = async () => { await chrome.storage.local.set({ ttLogo: String(r.result) }); logoStatus.textContent = `Logo opgeslagen (${f.name}).`; };
  r.readAsDataURL(f);
};
document.getElementById("logoWissen").onclick = async () => { await chrome.storage.local.remove("ttLogo"); logoVeld.value = ""; logoStatus.textContent = "Eigen logo gewist."; };
