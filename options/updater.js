/* Zelf bijwerken.
 * Een uitgepakte extensie kan zijn eigen map niet zomaar beschrijven. Via de File System
 * Access API kan het wél, als de gebruiker de extensiemap één keer zelf aanwijst. De map-
 * handle bewaren we in IndexedDB; bij elke update vraagt Chrome opnieuw om toestemming
 * (één klik). Daarna: zip ophalen met het GitHub-token, uitpakken met de ingebouwde
 * DecompressionStream, bestanden schrijven, extensie herladen. Er wordt niets verwijderd.
 */

/* ---------- Zip lezen (alleen store en deflate, geen zip64) ---------- */
async function dtUnzip(buffer) {
  const u8 = new Uint8Array(buffer);
  const dv = new DataView(buffer);
  // End of central directory: achteraan zoeken naar 50 4b 05 06.
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x05 && u8[i + 3] === 0x06) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("geen geldig zip-bestand");
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  const files = [];
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("zip: centrale map beschadigd");
    const method = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const size = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nameLen));
    p += 46 + nameLen + extraLen + commentLen;
    if (name.endsWith("/")) continue; // map
    if (dv.getUint32(local, true) !== 0x04034b50) throw new Error("zip: bestandskop beschadigd");
    const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
    const raw = u8.slice(start, start + compSize);
    let data;
    if (method === 0) data = raw;
    else if (method === 8) data = new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).arrayBuffer());
    else throw new Error(`zip: compressie ${method} niet ondersteund`);
    if (data.length !== size) throw new Error(`zip: ${name} heeft een onverwachte grootte`);
    files.push({ name, data });
  }
  return files;
}
globalThis.dtUnzip = dtUnzip;

if (typeof document !== "undefined" && typeof chrome !== "undefined" && chrome.storage) (() => {
  const $ = (id) => document.getElementById(id);
  const manifest = chrome.runtime.getManifest();
  const progress = (t) => { $("updateProgress").textContent = t; };

  /* ---------- Map-handle bewaren in IndexedDB ---------- */
  const idb = () => new Promise((res, rej) => {
    const r = indexedDB.open("tt-updater", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
  const kvGet = async (k) => { const db = await idb(); return new Promise((res, rej) => { const r = db.transaction("kv").objectStore("kv").get(k); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); };
  const kvSet = async (k, v) => { const db = await idb(); return new Promise((res, rej) => { const tx = db.transaction("kv", "readwrite"); tx.objectStore("kv").put(v, k); tx.oncomplete = res; tx.onerror = () => rej(tx.error); }); };
  const kvDel = async (k) => { const db = await idb(); return new Promise((res, rej) => { const tx = db.transaction("kv", "readwrite"); tx.objectStore("kv").delete(k); tx.oncomplete = res; tx.onerror = () => rej(tx.error); }); };

  /* Controleert dat de gekozen map echt deze extensie bevat. */
  async function verifyFolder(dir) {
    let text;
    try { text = await (await (await dir.getFileHandle("manifest.json")).getFile()).text(); }
    catch { throw new Error("deze map bevat geen manifest.json; kies de map waarin je de extensie hebt uitgepakt"); }
    const m = JSON.parse(text);
    if (m.name !== manifest.name) throw new Error(`deze map bevat een andere extensie (${m.name})`);
    return m;
  }

  async function ensurePermission(dir) {
    if ((await dir.queryPermission({ mode: "readwrite" })) === "granted") return true;
    return (await dir.requestPermission({ mode: "readwrite" })) === "granted";
  }

  async function writeFile(dir, path, data) {
    const parts = path.split("/").filter(Boolean);
    const file = parts.pop();
    let d = dir;
    for (const seg of parts) d = await d.getDirectoryHandle(seg, { create: true });
    const fh = await d.getFileHandle(file, { create: true });
    const w = await fh.createWritable();
    await w.write(data); await w.close();
  }

  /* ---------- UI-status ---------- */
  let update = null, dir = null;

  async function refreshUi() {
    update = (await chrome.storage.local.get("update")).update || null;
    dir = (await kvGet("dir").catch(() => null)) || null;
    const supported = typeof window.showDirectoryPicker === "function";
    $("folderName").textContent = dir ? `Extensiemap: ${dir.name}` : "Nog geen extensiemap gekozen.";
    $("pickFolder").textContent = dir ? "Andere map kiezen" : "Extensiemap kiezen";
    $("forgetFolder").hidden = !dir;
    const inPopup = !$("updates");
    if (inPopup) { $("pickFolder").hidden = !!dir; $("folderName").hidden = !dir; }
    const can = supported && dir && update?.available && (update.zipApi || update.zip);
    $("doUpdate").disabled = !can;
    $("doUpdate").textContent = update?.available ? `Nu bijwerken naar ${update.latest}` : "Nu bijwerken";
    if (!supported) progress("Deze browser ondersteunt het kiezen van een map niet; werk handmatig bij (zip downloaden).");
    else if (!dir) progress(update?.available ? "Kies eerst één keer de map waarin de extensie staat (de map met manifest.json)." : "");
    else if (!update?.available) progress(update?.latest ? `Je hebt de nieuwste versie (${manifest.version}).` : "");
    else progress(`Versie ${update.latest} staat klaar. Chrome vraagt bij het bijwerken om toestemming voor de map.`);
  }

  $("pickFolder").onclick = async () => {
    try {
      const picked = await window.showDirectoryPicker({ mode: "readwrite", id: "tt-extension-folder" });
      const m = await verifyFolder(picked);
      await kvSet("dir", picked);
      progress(`Map gekozen (${picked.name}, versie ${m.version}).`);
    } catch (e) {
      if (e?.name === "AbortError") { /* geannuleerd */ }
      else if (!$("updates")) { progress("Kies de map via de instellingenpagina, die opent nu…"); chrome.tabs.create({ url: chrome.runtime.getURL("options/options.html#update") }); }
      else progress(`Map niet gebruikt: ${e.message || e}.`);
    }
    refreshUi();
  };
  $("forgetFolder").onclick = async () => { await kvDel("dir").catch(() => {}); refreshUi(); };

  /* ---------- Bijwerken ---------- */
  async function fetchZip() {
    const headers = await chrome.runtime.sendMessage({ type: "getDownloadHeaders" });
    if (update.zipApi) {
      const res = await fetch(update.zipApi, { headers, cache: "no-store" });
      if (res.ok) return res.arrayBuffer();
      progress(`Download via API gaf ${res.status}, probeer de gewone downloadlink…`);
    }
    if (update.zip) {
      const res = await fetch(update.zip, { credentials: "include", cache: "no-store" });
      if (res.ok) return res.arrayBuffer();
      throw new Error(`download mislukt (${res.status}); controleer het token of log in bij GitHub`);
    }
    throw new Error("geen downloadlink in de release");
  }

  $("doUpdate").onclick = async () => {
    const btn = $("doUpdate");
    btn.disabled = true;
    try {
      // Toestemming eerst, zolang de klik nog "vers" is.
      if (!(await ensurePermission(dir))) throw new Error("geen toestemming voor de map");
      await verifyFolder(dir);
      progress("Downloaden…");
      const buf = await fetchZip();
      progress("Uitpakken…");
      const files = await dtUnzip(buf);
      // Bovenste map in de zip (teams-docent-tools/) weglaten.
      const strip = (n) => n.includes("/") ? n.slice(n.indexOf("/") + 1) : n;
      const entries = files.map((f) => ({ path: strip(f.name), data: f.data })).filter((f) => f.path && !f.path.startsWith(".") && !f.path.includes("/."));
      const mf = entries.find((f) => f.path === "manifest.json");
      if (!mf) throw new Error("zip bevat geen manifest.json");
      const nm = JSON.parse(new TextDecoder().decode(mf.data));
      if (nm.name !== manifest.name) throw new Error(`zip bevat een andere extensie (${nm.name})`);
      if (nm.version !== update.latest) throw new Error(`zip is versie ${nm.version}, verwacht ${update.latest}`);
      // Manifest als laatste, zodat een afgebroken update geen nieuw manifest met oude bestanden achterlaat.
      const ordered = [...entries.filter((f) => f.path !== "manifest.json"), mf];
      for (let i = 0; i < ordered.length; i++) {
        progress(`Bestanden schrijven (${i + 1}/${ordered.length}): ${ordered[i].path}`);
        await writeFile(dir, ordered[i].path, ordered[i].data);
      }
      progress(`Klaar: ${ordered.length} bestanden geschreven. De extensie wordt herladen…`);
      await chrome.storage.local.set({ update: { ...update, available: false, current: nm.version } });
      setTimeout(() => chrome.runtime.reload(), 800);
    } catch (e) {
      progress(`Bijwerken mislukt: ${e.message || e}. Je kunt ook handmatig bijwerken (zip downloaden).`);
      btn.disabled = false;
    }
  };

  $("manualZip").onclick = (e) => { e.preventDefault(); if (update?.zip || update?.page) chrome.tabs.create({ url: update.zip || update.page }); };

  chrome.storage.onChanged.addListener((ch, area) => { if (area === "local" && ch.update) refreshUi(); });
  refreshUi();
  if (location.hash === "#update" && $("updates")) setTimeout(() => $("updates").scrollIntoView({ behavior: "smooth" }), 100);
})();
