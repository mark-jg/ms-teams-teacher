/* Opstarten: instellingen laden, modules starten, live reageren op wijzigingen uit de instellingen. */
(async () => {
  if (!document.body || !document.head) return;
  await TT.loadSettings();
  TT.startAll();
  chrome.storage.onChanged.addListener(async (changes, area) => {
    if (area !== "sync" || !changes.settings) return;
    await TT.loadSettings();
    TT.startAll();
    TT.notify();
  });
  chrome.runtime.onMessage.addListener((msg) => {
    if (msg?.type === "command" && TT.isTop) document.dispatchEvent(new CustomEvent("tt:command", { detail: msg.command }));
  });
  if (TT.isTop) console.info(`[Docent Tools voor Teams] v${TT.version} actief op ${location.host}`);
})();
