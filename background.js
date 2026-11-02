importScripts("content/defaults.js");

/* ---------- Updates via GitHub Releases ----------
 * Een uitgepakte extensie kan zichzelf niet vervangen. Wat wél kan: periodiek de nieuwste
 * release op GitHub opvragen (alleen het versienummer en de downloadlink, er gaat niets
 * over de gebruiker mee), een badge op het icoon zetten en in de popup een downloadknop
 * en een herlaadknop tonen. Uit te zetten in de instellingen (updates.check).
 */
const REPO = "mark-jg/ms-teams-teacher";
const RELEASES_API = `https://api.github.com/repos/${REPO}/releases/latest`;
const RELEASES_PAGE = `https://github.com/${REPO}/releases`;
const REPO_API = `https://api.github.com/repos/${REPO}`;
const ALARM = "tt-update-check";
const CHECK_INTERVAL_MIN = 6 * 60;

/* De repo is openbaar: controleren en bijwerken werkt zonder token. Een fine-grained token
 * (alleen Contents: lezen) is optioneel, bijvoorbeeld tegen de limiet van GitHub voor anonieme
 * aanroepen; het staat in chrome.storage.local (dit apparaat) en gaat alleen naar api.github.com. */
async function getToken() {
  return ((await chrome.storage.local.get("ghToken")).ghToken || "").trim();
}

function ghHeaders(token, accept = "application/vnd.github+json") {
  const h = { Accept: accept, "X-GitHub-Api-Version": "2022-11-28" };
  if (token) h.Authorization = `Bearer ${token}`;
  return h;
}

function cmpVersion(a, b) {
  const pa = String(a).split(".").map((n) => parseInt(n, 10) || 0);
  const pb = String(b).split(".").map((n) => parseInt(n, 10) || 0);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (pa[i] || 0) - (pb[i] || 0);
    if (d) return d;
  }
  return 0;
}

async function updatesEnabled() {
  const { settings } = await chrome.storage.sync.get("settings");
  return settings?.updates?.check !== false;
}

function scheduleUpdateCheck() {
  chrome.alarms.create(ALARM, { delayInMinutes: 1, periodInMinutes: CHECK_INTERVAL_MIN });
}

async function setBadge(on) {
  await chrome.action.setBadgeText({ text: on ? "1" : "" });
  if (on) await chrome.action.setBadgeBackgroundColor({ color: "#E8B84A" });
}

async function checkForUpdate(force = false) {
  const current = chrome.runtime.getManifest().version;
  if (!force && !(await updatesEnabled())) {
    await setBadge(false);
    return { current, disabled: true };
  }
  const prev = (await chrome.storage.local.get("update")).update || {};
  const token = await getToken();
  let info;
  try {
    const res = await fetch(RELEASES_API, { headers: ghHeaders(token), cache: "no-store" });
    if (res.status === 401) {
      throw new Error("het GitHub-token is ongeldig of verlopen");
    } else if (res.status === 403 || res.status === 429) {
      throw new Error("GitHub-limiet bereikt, probeer het later opnieuw");
    } else if (res.status === 404) {
      // Óf er is nog geen release, óf we hebben geen toegang tot de (privé) repo.
      const repo = await fetch(REPO_API, { headers: ghHeaders(token), cache: "no-store" });
      if (!repo.ok) {
        throw new Error(token
          ? "het token geeft geen toegang tot deze repo (controleer repo en rechten)"
          : "GitHub gaf geen toegang tot de repo (tijdelijke limiet of de repo is niet bereikbaar); probeer later opnieuw of vul een GitHub-token in bij Instellingen > Updates");
      }
      info = { current, latest: null, available: false, page: RELEASES_PAGE, zip: null, zipApi: null, zipName: null, hasToken: Boolean(token), checkedAt: Date.now(), error: null };
    } else if (!res.ok) {
      throw new Error(`GitHub antwoordde met ${res.status}`);
    } else {
      const rel = await res.json();
      const latest = String(rel.tag_name || "").replace(/^v/i, "");
      const asset = (rel.assets || []).find((a) => /\.zip$/i.test(a.name));
      info = {
        current,
        latest,
        available: Boolean(latest) && cmpVersion(latest, current) > 0,
        page: rel.html_url || RELEASES_PAGE,
        zip: asset?.browser_download_url || null,
        zipApi: asset?.url || null,
        zipName: asset?.name || null,
        hasToken: Boolean(token),
        notes: String(rel.body || "").slice(0, 1500),
        publishedAt: rel.published_at || null,
        checkedAt: Date.now(),
        error: null
      };
    }
  } catch (e) {
    info = { ...prev, current, available: Boolean(prev.latest) && cmpVersion(prev.latest, current) > 0, hasToken: Boolean(token), checkedAt: Date.now(), error: String(e?.message || e) };
  }
  await chrome.storage.local.set({ update: info });
  await setBadge(Boolean(info.available));
  return info;
}

chrome.alarms.onAlarm.addListener((alarm) => { if (alarm.name === ALARM) checkForUpdate(); });
chrome.runtime.onStartup.addListener(() => { scheduleUpdateCheck(); checkForUpdate(); });

/* ---------- Sneltoetsen ---------- */
// Stuurt sneltoetsen door naar het actieve Teams-tabblad.
chrome.commands.onCommand.addListener(async (command) => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !tab.url || !/teams\.microsoft\.com|teams\.cloud\.microsoft/.test(tab.url)) return;
  chrome.tabs.sendMessage(tab.id, { type: "command", command }).catch(() => {});
});

/* ---------- Installatie en herladen ---------- */
chrome.runtime.onInstalled.addListener(async ({ reason, previousVersion }) => {
  scheduleUpdateCheck();
  checkForUpdate();
  // Na een (zelf)update: één keer "Bijgewerkt van X naar Y" in het menu tonen.
  const current = chrome.runtime.getManifest().version;
  if (reason === "update" && previousVersion && previousVersion !== current) {
    await chrome.storage.local.set({ justUpdated: { from: previousVersion, to: current, at: Date.now() } });
  }
  if (reason !== "install") return;
  const { settings } = await chrome.storage.sync.get("settings");
  if (settings) return;
  await chrome.storage.sync.set({ settings: structuredClone(globalThis.TT_DEFAULTS) });
});

/* ---------- Berichten van popup, instellingen en content ---------- */
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  switch (msg?.type) {
    case "openPopup":
      (chrome.action.openPopup ? chrome.action.openPopup() : Promise.reject()).catch(() => chrome.runtime.openOptionsPage());
      break;
    case "openOptions":
      chrome.runtime.openOptionsPage();
      return false;
    case "checkUpdate":
      checkForUpdate(true).then(sendResponse);
      return true; // asynchroon antwoord
    case "setToken":
      chrome.storage.local.set({ ghToken: String(msg.token || "").trim() }).then(() => checkForUpdate(true)).then(sendResponse);
      return true;
    case "getDownloadHeaders":
      // Voor de popup: headers om de zip via de API op te halen (privé repo).
      getToken().then((t) => sendResponse(ghHeaders(t, "application/octet-stream")));
      return true;
    case "reloadExtension":
      // Laadt de (inmiddels vervangen) bestanden op schijf opnieuw.
      setTimeout(() => chrome.runtime.reload(), 150);
      return false;
    case "openExtensionsPage": {
      const scheme = /Edg\//.test(navigator.userAgent) ? "edge" : "chrome";
      chrome.tabs.create({ url: `${scheme}://extensions/?id=${chrome.runtime.id}` });
      return false;
    }
  }
  return false;
});
