/* global Office, Excel, QRCode */
(() => {
  "use strict";

  const DEFAULT_SERVER = "https://ntfy.sh";
  const storageKey = "scan2cell.session.v1";
  let session;
  let socket;

  const $ = (id) => document.getElementById(id);
  const statusEl = $("status");
  const errorEl = $("error");

  function showError(message) {
    errorEl.hidden = false;
    errorEl.textContent = String(message);
  }

  function setStatus(text, css) {
    statusEl.textContent = text;
    statusEl.className = `status ${css}`;
  }

  function bytesToBase64Url(bytes) {
    let binary = "";
    bytes.forEach((b) => { binary += String.fromCharCode(b); });
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  function base64UrlToBytes(value) {
    const padded = value.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((value.length + 3) % 4);
    const binary = atob(padded);
    return Uint8Array.from(binary, (c) => c.charCodeAt(0));
  }

  function randomToken(bytes) {
    const data = new Uint8Array(bytes);
    crypto.getRandomValues(data);
    return bytesToBase64Url(data);
  }

  function createSession() {
    return {
      server: DEFAULT_SERVER,
      topic: `s2c-${randomToken(24)}`,
      key: randomToken(32),
      createdAt: Date.now()
    };
  }

  function saveSession() {
    localStorage.setItem(storageKey, JSON.stringify(session));
  }

  function loadSession() {
    try {
      const stored = JSON.parse(localStorage.getItem(storageKey));
      if (stored?.server && stored?.topic && stored?.key) return stored;
    } catch (_) { /* ignore invalid storage */ }
    return createSession();
  }

  function pairingUri() {
    const params = new URLSearchParams({
      server: session.server,
      topic: session.topic,
      key: session.key
    });
    return `scan2cell://pair?${params.toString()}`;
  }

  async function renderQr() {
    const target = $("pairQr");
    target.replaceChildren();
    new QRCode(target, {
      text: pairingUri(),
      width: 260,
      height: 260,
      colorDark: "#111827",
      colorLight: "#ffffff",
      correctLevel: QRCode.CorrectLevel.M
    });
  }

  function websocketUrl() {
    const base = session.server.replace(/^https:/, "wss:").replace(/^http:/, "ws:").replace(/\/$/, "");
    return `${base}/${encodeURIComponent(session.topic)}/ws`;
  }

  async function decryptMessage(encoded) {
    const envelope = JSON.parse(encoded);
    if (envelope.v !== 1 || !envelope.iv || !envelope.data) throw new Error("Unsupported encrypted message.");
    const key = await crypto.subtle.importKey("raw", base64UrlToBytes(session.key), { name: "AES-GCM" }, false, ["decrypt"]);
    const plain = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: base64UrlToBytes(envelope.iv) },
      key,
      base64UrlToBytes(envelope.data)
    );
    return JSON.parse(new TextDecoder().decode(plain));
  }

  async function insertIntoExcel(value) {
    const moveMode = $("moveMode").value;
    const asText = $("asText").checked;

    await Excel.run(async (context) => {
      const selected = context.workbook.getSelectedRange();
      const cell = selected.getCell(0, 0);
      cell.load(["address", "rowIndex", "columnIndex"]);
      await context.sync();

      if (asText) cell.numberFormat = [["@"]];
      cell.values = [[String(value)]];

      let next = cell;
      if (moveMode === "down") next = cell.getOffsetRange(1, 0);
      if (moveMode === "right") next = cell.getOffsetRange(0, 1);
      next.select();
      await context.sync();

      $("lastInsert").textContent = `${cell.address}: ${String(value).slice(0, 26)}`;
    });
  }

  async function handleNtfyEvent(raw) {
    const event = JSON.parse(raw);
    if (event.event !== "message" || !event.message) return;
    const message = await decryptMessage(event.message);
    if (message.type === "paired") {
      setStatus("Phone paired", "paired");
      return;
    }
    if (message.type === "scan" && typeof message.text === "string") {
      await insertIntoExcel(message.text);
      setStatus("Connected", "online");
    }
  }

  function connect() {
    if (socket) socket.close();
    setStatus("Connecting…", "offline");
    socket = new WebSocket(websocketUrl());
    socket.addEventListener("open", () => setStatus("Waiting for phone", "online"));
    socket.addEventListener("message", (event) => {
      handleNtfyEvent(event.data).catch((error) => showError(error.stack || error));
    });
    socket.addEventListener("close", () => {
      setStatus("Reconnecting…", "offline");
      setTimeout(connect, 2500);
    });
    socket.addEventListener("error", () => setStatus("Connection error", "offline"));
  }

  async function resetPairing() {
    session = createSession();
    saveSession();
    await renderQr();
    connect();
  }

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) {
      showError("This add-in must be opened in Excel.");
      return;
    }
    session = loadSession();
    saveSession();
    $("moveMode").value = localStorage.getItem("scan2cell.move") || "down";
    $("asText").checked = localStorage.getItem("scan2cell.asText") !== "false";
    $("moveMode").addEventListener("change", (e) => localStorage.setItem("scan2cell.move", e.target.value));
    $("asText").addEventListener("change", (e) => localStorage.setItem("scan2cell.asText", String(e.target.checked)));
    $("newPairing").addEventListener("click", () => resetPairing().catch(showError));
    await renderQr();
    connect();
  });
})();
