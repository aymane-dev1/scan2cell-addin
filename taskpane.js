/* global Office, Excel */
(() => {
  "use strict";

  const DEFAULT_SERVER = "https://ntfy.sh";
  const STORAGE_KEY = "scan2cell.session.v2";
  const PAIRING_TTL_MS = 5 * 60 * 1000;
  const PAIRING_REPUBLISH_MS = 30 * 1000;

  let session;
  let socket;
  let pairingPublishTimer;
  let countdownTimer;
  let resetting = false;

  const $ = (id) => document.getElementById(id);
  const statusEl = $("status");
  const errorEl = $("error");

  function showError(message) {
    errorEl.hidden = false;
    errorEl.textContent = String(message);
  }

  function clearError() {
    errorEl.hidden = true;
    errorEl.textContent = "";
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

  function randomSixDigitCode() {
    const value = new Uint32Array(1);
    crypto.getRandomValues(value);
    return String(100000 + (value[0] % 900000));
  }

  function createSession() {
    return {
      server: DEFAULT_SERVER,
      topic: `s2c-${randomToken(24)}`,
      key: randomToken(32),
      pairCode: randomSixDigitCode(),
      pairExpiresAt: Date.now() + PAIRING_TTL_MS,
      createdAt: Date.now(),
      paired: false
    };
  }

  function saveSession() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(session));
  }

  function loadSession() {
    try {
      const stored = JSON.parse(localStorage.getItem(STORAGE_KEY));
      if (stored?.server && stored?.topic && stored?.key) {
        if (!stored.paired && (!stored.pairCode || Number(stored.pairExpiresAt) <= Date.now())) {
          return createSession();
        }
        return stored;
      }
    } catch (_) { /* ignore invalid storage */ }
    return createSession();
  }

  async function sha256Bytes(text) {
    return new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)));
  }

  function bytesToHex(bytes) {
    return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
  }

  async function pairingCoordinates(code) {
    const topicHash = await sha256Bytes(`topic|scan2cell|v1|${code}`);
    const keyBytes = await sha256Bytes(`key|scan2cell|v1|${code}`);
    return {
      topic: `s2c-pair-${bytesToHex(topicHash).slice(0, 40)}`,
      key: bytesToBase64Url(keyBytes)
    };
  }

  async function encryptWithKey(payload, base64UrlKey) {
    const iv = new Uint8Array(12);
    crypto.getRandomValues(iv);
    const key = await crypto.subtle.importKey(
      "raw",
      base64UrlToBytes(base64UrlKey),
      { name: "AES-GCM" },
      false,
      ["encrypt"]
    );
    const encrypted = await crypto.subtle.encrypt(
      { name: "AES-GCM", iv },
      key,
      new TextEncoder().encode(JSON.stringify(payload))
    );
    return JSON.stringify({
      v: 1,
      iv: bytesToBase64Url(iv),
      data: bytesToBase64Url(new Uint8Array(encrypted))
    });
  }

  async function publishPairingPackage() {
    if (session.paired || Date.now() >= Number(session.pairExpiresAt)) return;
    const rendezvous = await pairingCoordinates(session.pairCode);
    const packageBody = await encryptWithKey({
      v: 1,
      server: session.server,
      topic: session.topic,
      key: session.key,
      expires: Number(session.pairExpiresAt)
    }, rendezvous.key);

    const response = await fetch(`${session.server.replace(/\/$/, "")}/${encodeURIComponent(rendezvous.topic)}`, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: packageBody
    });
    if (!response.ok) throw new Error(`Pairing relay returned HTTP ${response.status}`);
  }

  function renderPairingCode() {
    $("pairCode").textContent = session.paired ? "PAIRED" : session.pairCode;
    updateCountdown();
  }

  function updateCountdown() {
    const expiry = $("pairExpiry");
    if (session.paired) {
      expiry.textContent = "This phone is connected. Generate a new code to pair another phone.";
      return;
    }
    const remaining = Number(session.pairExpiresAt) - Date.now();
    if (remaining <= 0) {
      expiry.textContent = "Code expired. Generating a new code…";
      if (!resetting) resetPairing().catch(showError);
      return;
    }
    const minutes = Math.floor(remaining / 60000);
    const seconds = Math.floor((remaining % 60000) / 1000);
    expiry.textContent = `Expires in ${minutes}:${String(seconds).padStart(2, "0")}`;
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
      session.paired = true;
      saveSession();
      renderPairingCode();
      setStatus("Phone paired", "paired");
      stopPairingPublisher();
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
    socket.addEventListener("open", () => {
      setStatus(session.paired ? "Ready for phone" : "Waiting for phone", "online");
    });
    socket.addEventListener("message", (event) => {
      handleNtfyEvent(event.data).catch((error) => showError(error.stack || error));
    });
    socket.addEventListener("close", () => {
      setStatus("Reconnecting…", "offline");
      setTimeout(connect, 2500);
    });
    socket.addEventListener("error", () => setStatus("Connection error", "offline"));
  }

  function stopPairingPublisher() {
    if (pairingPublishTimer) {
      clearInterval(pairingPublishTimer);
      pairingPublishTimer = null;
    }
  }

  async function startPairingPublisher() {
    stopPairingPublisher();
    if (session.paired) return;
    await publishPairingPackage();
    pairingPublishTimer = setInterval(() => {
      publishPairingPackage().catch(showError);
    }, PAIRING_REPUBLISH_MS);
  }

  async function resetPairing() {
    if (resetting) return;
    resetting = true;
    try {
      clearError();
      stopPairingPublisher();
      session = createSession();
      saveSession();
      renderPairingCode();
      connect();
      await startPairingPublisher();
    } finally {
      resetting = false;
    }
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

    renderPairingCode();
    connect();
    await startPairingPublisher();

    if (countdownTimer) clearInterval(countdownTimer);
    countdownTimer = setInterval(updateCountdown, 1000);
  });
})();
