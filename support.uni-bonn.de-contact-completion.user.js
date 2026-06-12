// ==UserScript==
// @name        Znuny: Contact Autocomplete (multi-source)
// @namespace   github.com/olifre/userstyles
// @match       https://support.uni-bonn.de/*
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-contact-completion.user.js
// @downloadURL https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-contact-completion.user.js
// @icon        https://olifre.github.io/favicon.ico
// @version     1.6.1
// @grant       GM_xmlhttpRequest
// @grant       GM.xmlHttpRequest
// @connect     jira.team.uni-bonn.de
// @connect     grp_phy.gitlab-pages.uni-bonn.de
// @description Autocomplete for Znuny contacts based on multiple sources (priority on collisions)
// @author      Oliver Freyermuth <o.freyermuth@googlemail.com> (https://olifre.github.io/)
// @license     Unlicense
// ==/UserScript==

(function () {
  'use strict';

  // Filter on relevant pages.
  if (!(
    /\bAction=AgentTicketCompose\b/.test(location.search) ||
    /\bAction=AgentTicketEmail\b/.test(location.search) ||
    /\bAction=AgentTicketEmailOutbound\b/.test(location.search) ||
    /\bAction=AgentTicketPhoneOutbound\b/.test(location.search) ||
    /\bAction=AgentTicketPhoneInbound\b/.test(location.search) ||
    /\bAction=AgentTicketPhone\b/.test(location.search) ||
    /\bAction=AgentTicketForward\b/.test(location.search)
  )) return;

  // ---------------------------
  // Persisted settings via localStorage
  // ---------------------------
  const LS_CUSTOMTEXT_URLS = "znuny_contact_autocomplete_customtext_urls";
  const LS_CUSTOMTEXT_UI_ENABLED = "znuny_contact_autocomplete_customtext_ui_enabled";
  const LS_STATUSBOX_COLLAPSED = "znuny_contact_autocomplete_statusbox_collapsed";

  function lsGet(key, def) {
    try {
      const v = localStorage.getItem(key);
      return v === null ? def : v;
    } catch {
      return def;
    }
  }

  function lsSet(key, val) {
    try {
      localStorage.setItem(key, String(val ?? ""));
    } catch {}
  }

  let customTextHasUrls = false;

  async function refreshCustomTextHasUrlsCache() {
    const uiVal = lsGet(LS_CUSTOMTEXT_URLS, "");
    customTextHasUrls = !!String(uiVal || "").trim();
  }

  // ---------------------------
  // In-memory state
  // ---------------------------
  const perSourceProgress = {};
  const perSourceError = {};
  const perSourceHoldoffUntil = {};
  const HOLDOFF_MS = 60 * 1000;
  const perSourceHoldoffTimer = {};
  const perSourceRefreshing = {};
  const perSourceLeaseBlockedBy = {};

  const BC_NAME = "znuny_contacts_bc_v2";
  const bc = new BroadcastChannel(BC_NAME);

  const OLD_CACHE_KEY = "znuny_contacts_cache_v1";
  const DB_NAME = "ZnunyContactDB";
  const DB_VERSION = 4;

  let contacts = [];
  let perSourceState = {};
  let perSourceCounts = {};
  let autocompleteInstances = [];
  let refreshInFlightLocal = false;
  let statusBox = null;

  function parseCustomTextLines(text) {
    const lines = String(text || "")
      .split(/\r?\n/)
      .map(l => l.trim())
      .filter(Boolean);

    const out = [];
    for (const line of lines) {
      const m = line.match(/^(\S+)\s+(.+)$/);
      if (!m) continue;

      const emailRaw = m[1];
      const nameRaw = m[2];

      const email = String(emailRaw).toLowerCase().trim();
      if (!email) continue;
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) continue;

      const display = String(nameRaw).replace(/\s+/g, " ").trim();
      const textOut = `"${display}" <${email}>`.replace(/\s+/g, " ").trim();

      out.push({
        key: `${email}|customtext`,
        email,
        text: textOut,
        search: textOut.toLowerCase(),
        source: "customtext",
        syncID: 0
      });
    }
    return out;
  }

  async function fetchText(url) {
    const req =
      (typeof GM_xmlhttpRequest === "function" && GM_xmlhttpRequest) ||
      (typeof GM !== "undefined" && GM?.xmlHttpRequest) ||
      null;
    if (!req) throw new Error("No userscript XHR API available");

    return new Promise((resolve, reject) => {
      req({
        method: "GET",
        url,
        withCredentials: false,
        headers: { "Accept": "text/plain, */*" },
        onload: (r) => {
          const status = r.status || 0;
          if (status < 200 || status >= 300) {
            const err = new Error("HTTP " + status);
            err._gm = r;
            return reject(err);
          }
          resolve(r.responseText || "");
        },
        onerror: (r) => {
          const err = new Error("network error");
          err._gm = r;
          reject(err);
        },
        ontimeout: (r) => {
          const err = new Error("network timeout");
          err._gm = r;
          reject(err);
        }
      });
    });
  }

  // ---------------------------
  // Sources
  // ---------------------------
  const SOURCES = {
    phy: {
      id: "phy",
      priority: 100,
      cacheTtlMs: 12 * 60 * 60 * 1000,
      refreshLeaseTtlMs: 2 * 60 * 1000,
      url: "https://grp_phy.gitlab-pages.uni-bonn.de/it/web/vcard_generator/contacts.json",

      async fetch() {
        const data = await xFetchJson(this.url);
        return data;
      },

      transform(data, syncID) {
        const entries = Object.values(data);
        return entries.map(e => {
          const email = (e.email || "").toLowerCase();
          const text = `"${e.degree ? e.degree + " " : ""}${e.firstname || ""} ${e.lastname || ""}" <${email}>`
            .replace(/\s+/g, " ")
            .trim();
          return {
            key: `${email}|${this.id}`,
            email,
            text,
            search: text.toLowerCase(),
            source: this.id,
            syncID
          };
        }).filter(x => x.email);
      }
    },

    // Jira (cursor-based) data source
    jira: {
      id: "jira",
      priority: 50,
      cacheTtlMs: 24 * 60 * 60 * 1000,
      refreshLeaseTtlMs: 2 * 60 * 1000,
      url: "https://jira.team.uni-bonn.de/rest/api/2/user/list",
      batchSize: 2000,

      async fetch() {
        let cursor = 0;
        const sleep = (ms) => new Promise(r => setTimeout(r, ms));

        const results = [];
        let totalFetched = 0;
        let isDone = false;

        perSourceRefreshing[this.id] = true;
        updateStatusBox();

        while (!isDone) {
          const u = new URL(this.url);
          u.searchParams.set("cursor", cursor);
          u.searchParams.set("maxResults", this.batchSize);

          const json = await xFetchJson(u.toString(), { withCredentials: true });

          const batch = Array.isArray(json.values) ? json.values : [];
          results.push(...batch);

          totalFetched += batch.length;
          perSourceProgress[this.id] = { fetched: totalFetched };
          updateStatusBox();

          const nc = json.nextCursor;
          const nextCursorNum = (typeof nc === "number") ? nc : (typeof nc === "string" ? Number(nc) : NaN);
          if (!Number.isNaN(nextCursorNum)) cursor = nextCursorNum;

          if (json.isLast) isDone = true;
          await sleep(200);
        }

        perSourceProgress[this.id] = { fetched: totalFetched };
        updateStatusBox();

        return results;
      },

      transform(data, syncID) {
        const entries = Array.isArray(data) ? data : [];
        return entries.map(u => {
          const email = ((u.emailAddress || "")).toLowerCase();
          const displayName = u.displayName || u.name || "";
          const text = `"${displayName}" <${email}>`.replace(/\s+/g, " ").trim();
          return {
            key: `${email}|${this.id}`,
            email,
            text,
            search: text.toLowerCase(),
            source: this.id,
            syncID
          };
        }).filter(x => x.email);
      }
    },

    customtext: {
      id: "customtext",
      priority: 150,
      cacheTtlMs: 6 * 60 * 60 * 1000,
      refreshLeaseTtlMs: 2 * 60 * 1000,

      async fetch() {
        if (!customTextHasUrls) return { urls: [], texts: [] };

        const ui = lsGet(LS_CUSTOMTEXT_URLS, "");
        const urls = String(ui || "")
          .split(/\r?\n/)
          .map(s => s.trim())
          .filter(Boolean);

        if (!urls.length) return { urls: [], texts: [] };

        const texts = [];
        for (const url of urls) {
          const t = await fetchText(url);
          texts.push(t);
        }
        return { urls, texts };
      },

      transform(data, syncID) {
        const texts = data?.texts || [];
        if (!texts.length) return [];

        const all = [];
        for (const t of texts) all.push(...parseCustomTextLines(t));

        return all.map(x => ({
          ...x,
          key: `${x.email}|${this.id}`,
          source: this.id,
          syncID
        }));
      }
    }
  };

  // ---------------------------
  // Status UI
  // ---------------------------
  function ensureStatusBox() {
    if (statusBox) return statusBox;

    const box = document.createElement("div");
    statusBox = box;
    box.id = "znunyContactSourceStatus";

    Object.assign(box.style, {
      position: "fixed",
      top: "8px",
      left: "50%",
      transform: "translateX(-50%)",
      zIndex: 100000,
      background: "rgba(255,255,255,0.97)",
      border: "1px solid #bbb",
      borderRadius: "6px",
      padding: "6px 10px",
      boxShadow: "0 4px 10px rgba(0,0,0,0.15)",
      fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
      fontSize: "12px",
      color: "#111",
      minWidth: "360px",
      maxWidth: "900px",
      pointerEvents: "auto",
      whiteSpace: "normal"
    });

    document.body.appendChild(box);
    return box;
  }

  function fmtIso(ts) {
    if (!ts) return "never";
    try { return new Date(ts).toISOString(); } catch { return "invalid"; }
  }

  function fmtRemaining(ms) {
    const s = Math.ceil(ms / 1000);
    return `${s}s`;
  }

  function holdoffRemainingMs(sourceId) {
    const until = perSourceHoldoffUntil[sourceId] || 0;
    return Math.max(0, until - Date.now());
  }

  function setSourceHoldoff(sourceId, ms = HOLDOFF_MS) {
    const until = Date.now() + ms;
    perSourceHoldoffUntil[sourceId] = until;

    if (perSourceHoldoffTimer[sourceId]) clearTimeout(perSourceHoldoffTimer[sourceId]);
    perSourceHoldoffTimer[sourceId] = setTimeout(() => {
      perSourceHoldoffUntil[sourceId] = 0;
      clearSourceError(sourceId);
      updateStatusBox();
      backgroundRefreshIfNeeded();
    }, Math.max(0, until - Date.now()) + 50);

    updateStatusBox();
  }

  function clearSourceHoldoff(sourceId) {
    if (perSourceHoldoffUntil[sourceId]) {
      perSourceHoldoffUntil[sourceId] = 0;
      if (perSourceHoldoffTimer[sourceId]) {
        clearTimeout(perSourceHoldoffTimer[sourceId]);
        perSourceHoldoffTimer[sourceId] = 0;
      }
      updateStatusBox();
    }
  }

  function setSourceError(sourceId, message) {
    perSourceError[sourceId] = { message: String(message || ""), ts: Date.now() };
    updateStatusBox();
  }

  function clearSourceError(sourceId) {
    if (perSourceError[sourceId]?.message) {
      perSourceError[sourceId] = { message: "", ts: 0 };
      updateStatusBox();
    }
  }

  function classifyError(e) {
    const msg = String(e?.message || e || "");
    const m = msg.match(/^HTTP\s+(\d{3})/);
    if (m) {
      const code = Number(m[1]);
      if (code === 401 || code === 403) return "authentication required";
      return `HTTP ${code}`;
    }
    if (/network timeout|failed to fetch|networkerror|load failed|fetch/i.test(msg)) return "network error";
    if (/network error/i.test(msg)) return "network error";
    return msg || "unknown error";
  }

  function isAuthError(e) {
    const msg = String(e?.message || e || "");
    const m = msg.match(/^HTTP\s+(\d{3})/);
    if (!m) return false;
    const code = Number(m[1]);
    return code === 401 || code === 403;
  }

  function isNetworkError(e) {
    const msg = String(e?.message || e || "");
    if (/^HTTP\s+\d{3}\b/.test(msg)) return false;
    return /network error|network timeout|failed to fetch|networkerror|load failed|fetch/i.test(msg);
  }

  function statusLine(sourceId) {
    const s = perSourceState[sourceId] || { hasAny: false, isStale: true, lastUpdate: 0 };
    const n = perSourceCounts[sourceId] ?? 0;
    const p = perSourceProgress[sourceId] || { fetched: 0 };
    const freshness = s.hasAny ? (s.isStale ? "stale" : "fresh") : "empty";

    const holdMs = holdoffRemainingMs(sourceId);
    const onHoldoff = holdMs > 0;

    const refreshing = !!perSourceRefreshing[sourceId] && !onHoldoff;
    const fetchingText = refreshing ? " (refreshing…)" : "";
    const fetchPart = refreshing ? ` | fetch: ${p.fetched ?? 0}` : "";

    const holdPart = onHoldoff ? ` | hold-off: ${fmtRemaining(holdMs)}` : "";

    const err = perSourceError[sourceId]?.message;
    const errPart = err ? ` | error: ${err}` : "";

    const blockedBy = perSourceLeaseBlockedBy[sourceId];
    const lockPart = blockedBy ? ` | locked (other tab)` : "";

    const label = (sourceId === "customtext") ? "Custom Text sources" : sourceId;
    return `${label}: ${freshness}${fetchingText}${fetchPart} | contacts: ${n} | last fetch: ${fmtIso(s.lastUpdate)}${holdPart}${lockPart}${errPart}`;
  }

  function updateStatusBox() {
    const box = ensureStatusBox();
    box.innerHTML = "";

    const collapsed = lsGet(LS_STATUSBOX_COLLAPSED, "0") === "1";

    Object.assign(box.style, collapsed ? {
      padding: "4px 8px",
      minWidth: "0",
      maxWidth: "none",
      width: "auto",
      borderRadius: "999px",
      whiteSpace: "nowrap"
    } : {
      padding: "6px 10px",
      minWidth: "360px",
      maxWidth: "900px",
      width: "auto",
      borderRadius: "6px",
      whiteSpace: "normal"
    });

    if (collapsed) {
      const open = document.createElement("a");
      open.href = "#";
      open.textContent = "👥";
      open.title = "Show contact autocomplete status";
      open.style.textDecoration = "none";
      open.style.fontWeight = "700";
      open.style.color = "#0645ad";
      open.style.pointerEvents = "auto";

      open.addEventListener("click", (ev) => {
        ev.preventDefault();
        lsSet(LS_STATUSBOX_COLLAPSED, "0");
        updateStatusBox();
      });

      box.appendChild(open);
      return;
    }

    const enabledEntries = Object.keys(SOURCES).filter(sid => {
      if (sid === "customtext") {
        if (!customTextHasUrls) return false;
      }
      return true;
    });

    for (const sid of enabledEntries) {
      const row = document.createElement("div");
      row.style.display = "flex";
      row.style.gap = "8px";
      row.style.alignItems = "baseline";

      const line = document.createElement("span");
      line.textContent = statusLine(sid);
      line.style.whiteSpace = "pre";
      line.style.flex = "1 1 auto";
      line.style.minWidth = "0";

      const a = document.createElement("a");
      a.href = "#";
      a.textContent = "⟳";
      a.title = `Force refresh ${sid}`;
      a.style.flex = "0 0 auto";
      a.style.textDecoration = "none";
      a.style.fontWeight = "600";
      a.style.color = "#0645ad";
      a.style.pointerEvents = "auto";

      a.addEventListener("click", async (ev) => {
        ev.preventDefault();
        const old = a.textContent;
        a.textContent = "…";
        try { await forceStaleAndRefresh(sid); }
        finally { a.textContent = old; }
      });

      row.appendChild(line);
      row.appendChild(a);
      box.appendChild(row);
    }

    const gearRow = document.createElement("div");
    gearRow.style.display = "flex";
    gearRow.style.gap = "8px";
    gearRow.style.alignItems = "baseline";

    const gearLabel = document.createElement("span");
    gearLabel.textContent = "Settings";
    gearLabel.style.whiteSpace = "pre";
    gearLabel.style.flex = "1 1 auto";
    gearLabel.style.minWidth = "0";

    const collapse = document.createElement("a");
    collapse.href = "#";
    collapse.textContent = "−";
    collapse.title = "Collapse status";
    collapse.style.flex = "0 0 auto";
    collapse.style.textDecoration = "none";
    collapse.style.fontWeight = "700";
    collapse.style.color = "#666";
    collapse.style.pointerEvents = "auto";

    collapse.addEventListener("click", (ev) => {
      ev.preventDefault();
      lsSet(LS_STATUSBOX_COLLAPSED, "1");
      updateStatusBox();
    });

    const gear = document.createElement("a");
    gear.href = "#";
    gear.textContent = "⚙";
    gear.title = "Configure custom text URLs";
    gear.style.flex = "0 0 auto";
    gear.style.textDecoration = "none";
    gear.style.fontWeight = "700";
    gear.style.color = "#666";
    gear.style.pointerEvents = "auto";

    gear.addEventListener("click", async (ev) => {
      ev.preventDefault();

      const existing = document.getElementById("znunyCustomTextUrlsConfig");
      if (existing) {
        lsSet(LS_CUSTOMTEXT_UI_ENABLED, "0");
        existing.remove();
        await refreshCustomTextHasUrlsCache();
        updateStatusBox();
        return;
      }

      lsSet(LS_CUSTOMTEXT_UI_ENABLED, "1");
      ensureCustomTextUrlsConfigUI();

      const uiVal = lsGet(LS_CUSTOMTEXT_URLS, "");
      const area = document.getElementById("znunyCustomTextUrlsArea");
      if (area) area.value = uiVal;

      await refreshCustomTextHasUrlsCache();
      updateStatusBox();
    });

    gearRow.appendChild(gearLabel);
    gearRow.appendChild(collapse);
    gearRow.appendChild(gear);
    box.appendChild(gearRow);
  }

  setInterval(() => {
    const anyHoldOrJustEnded = Object.keys(SOURCES).some(sid => {
      const rem = holdoffRemainingMs(sid);
      return rem > 0 || (perSourceHoldoffUntil[sid] && rem === 0);
    });
    if (anyHoldOrJustEnded) updateStatusBox();
  }, 1000);

  // ---------------------------
  // XHR + Storage helpers
  // ---------------------------
  function cleanupOldCache() {
    if (localStorage.getItem(OLD_CACHE_KEY)) {
      localStorage.removeItem(OLD_CACHE_KEY);
      console.log("[Autocomplete] Old localStorage cache cleared.");
    }
  }

  function xFetchJson(url, { withCredentials = false } = {}) {
    const req =
      (typeof GM_xmlhttpRequest === "function" && GM_xmlhttpRequest) ||
      (typeof GM !== "undefined" && GM?.xmlHttpRequest) ||
      null;
    if (!req) throw new Error("No userscript XHR API available");

    return new Promise((resolve, reject) => {
      req({
        method: "GET",
        url,
        withCredentials,
        headers: { "Accept": "application/json" },
        onload: (r) => {
          const status = r.status || 0;
          if (status < 200 || status >= 300) {
            const err = new Error("HTTP " + status);
            err._gm = r;
            return reject(err);
          }
          try { resolve(JSON.parse(r.responseText)); }
          catch (e) { e._gm = r; reject(e); }
        },
        onerror: (r) => {
          const err = new Error("network error");
          err._gm = r;
          reject(err);
        },
        ontimeout: (r) => {
          const err = new Error("network timeout");
          err._gm = r;
          reject(err);
        }
      });
    });
  }

  function metaKey(sourceId, field) {
    return `meta:source:${sourceId}:${field}`;
  }

  function openDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (event) => {
        const db = event.target.result;

        if (db.objectStoreNames.contains("contacts")) {
          db.deleteObjectStore("contacts");
        }
        db.createObjectStore("contacts", { keyPath: "key" });

        if (!db.objectStoreNames.contains("metadata")) {
          db.createObjectStore("metadata");
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function getMeta(key) {
    const db = await openDB();
    const tx = db.transaction("metadata", "readonly");
    const store = tx.objectStore("metadata");
    return new Promise((resolve) => {
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(undefined);
    });
  }

  async function setMeta(key, value) {
    const db = await openDB();
    const tx = db.transaction("metadata", "readwrite");
    tx.objectStore("metadata").put(value, key);
    return new Promise((resolve) => { tx.oncomplete = () => resolve(); });
  }

  async function tryAcquireRefreshLease(sourceId) {
    const now = Date.now();
    const owner = `${now}-${Math.random().toString(16).slice(2)}`;
    const leaseKey = metaKey(sourceId, "lease");
    const leaseTtl = (SOURCES[sourceId]?.refreshLeaseTtlMs ?? (2 * 60 * 1000));

    const db = await openDB();
    const tx = db.transaction("metadata", "readwrite");
    const store = tx.objectStore("metadata");

    const current = await new Promise((resolve) => {
      const req = store.get(leaseKey);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    });

    if (current && current.expiresAt && current.expiresAt > now) {
      return { ok: false, owner: current.owner };
    }

    store.put({ owner, expiresAt: now + leaseTtl }, leaseKey);

    return await new Promise((resolve) => {
      tx.oncomplete = async () => {
        const verify = await getMeta(leaseKey);
        resolve({ ok: !!(verify && verify.owner === owner), owner });
      };
      tx.onerror = () => resolve({ ok: false, owner: null });
      tx.onabort = () => resolve({ ok: false, owner: null });
    });
  }

  async function releaseRefreshLease(sourceId, owner) {
    const leaseKey = metaKey(sourceId, "lease");
    const cur = await getMeta(leaseKey);
    if (cur && cur.owner === owner) {
      await setMeta(leaseKey, { owner: "", expiresAt: 0 });
    }
  }

  async function saveSourceToIndexedDB(sourceId, data) {
    const db = await openDB();
    const tx = db.transaction(["contacts", "metadata"], "readwrite");
    const contactStore = tx.objectStore("contacts");
    const metaStore = tx.objectStore("metadata");

    data.forEach(item => contactStore.put(item));
    metaStore.put(Date.now(), metaKey(sourceId, "lastUpdate"));

    perSourceCounts[sourceId] = (perSourceCounts[sourceId] || 0) + data.length;

    return new Promise((resolve) => { tx.oncomplete = () => resolve(); });
  }

  async function cleanupOldRecords(sourceId, validSyncID) {
    const db = await openDB();
    const tx = db.transaction("contacts", "readwrite");
    const store = tx.objectStore("contacts");
    const request = store.openCursor();

    request.onsuccess = (event) => {
      const cursor = event.target.result;
      if (!cursor) return;

      if (cursor.value.source === sourceId && cursor.value.syncID !== validSyncID) {
        cursor.delete();
      }
      cursor.continue();
    };
  }

  function buildEffectiveContacts(allRows) {
    const bestByEmail = new Map();
    for (const row of allRows) {
      if (!row?.email) continue;
      const cur = bestByEmail.get(row.email);
      if (!cur) {
        bestByEmail.set(row.email, row);
        continue;
      }
      const pNew = SOURCES[row.source]?.priority ?? 0;
      const pCur = SOURCES[cur.source]?.priority ?? 0;
      if (pNew > pCur) bestByEmail.set(row.email, row);
    }
    return [...bestByEmail.values()];
  }

  async function loadFromIndexedDBWithMeta() {
    try {
      const db = await openDB();
      const tx = db.transaction(["contacts", "metadata"], "readonly");
      const contactStore = tx.objectStore("contacts");
      const metaStore = tx.objectStore("metadata");

      const allContacts = await new Promise((resolve) => {
        const req = contactStore.getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => resolve([]);
      });

      const state = {};
      const counts = {};
      await Promise.all(Object.keys(SOURCES).map(async (sourceId) => {
        const lastUpdate = await new Promise(res => {
          const req = metaStore.get(metaKey(sourceId, "lastUpdate"));
          req.onsuccess = () => res(req.result || 0);
          req.onerror = () => res(0);
        });

        const n = allContacts.reduce((acc, c) => acc + (c.source === sourceId ? 1 : 0), 0);
        counts[sourceId] = n;
        const hasAny = n > 0;
        const ttl = SOURCES[sourceId].cacheTtlMs;
        const isStale = !lastUpdate || (Date.now() - lastUpdate > ttl);

        state[sourceId] = { hasAny, isStale, lastUpdate };

        perSourceProgress[sourceId] = perSourceProgress[sourceId] || { fetched: 0 };
        perSourceRefreshing[sourceId] = perSourceRefreshing[sourceId] || false;
        perSourceHoldoffUntil[sourceId] = perSourceHoldoffUntil[sourceId] || 0;
        perSourceLeaseBlockedBy[sourceId] = perSourceLeaseBlockedBy[sourceId] || null;
      }));

      return { contacts: allContacts, state, counts };
    } catch (e) {
      console.error("[Autocomplete] IDB Error:", e);
      const state = {};
      const counts = {};
      for (const sid of Object.keys(SOURCES)) state[sid] = { hasAny: false, isStale: true, lastUpdate: 0 };
      for (const sid of Object.keys(SOURCES)) counts[sid] = 0;

      Object.keys(SOURCES).forEach(sid => {
        perSourceProgress[sid] = { fetched: 0 };
        perSourceRefreshing[sid] = false;
        perSourceHoldoffUntil[sid] = perSourceHoldoffUntil[sid] || 0;
        perSourceLeaseBlockedBy[sid] = perSourceLeaseBlockedBy[sid] || null;
      });
      return { contacts: [], state, counts };
    }
  }

  async function reloadContactsFromIDB() {
    const { contacts: allRows, state, counts } = await loadFromIndexedDBWithMeta();
    perSourceState = state;
    perSourceCounts = counts;
    contacts = buildEffectiveContacts(allRows);

    Object.keys(SOURCES).forEach(sid => {
      perSourceProgress[sid] = perSourceProgress[sid] || { fetched: 0 };
      perSourceRefreshing[sid] = perSourceRefreshing[sid] || false;
      perSourceHoldoffUntil[sid] = perSourceHoldoffUntil[sid] || 0;
      perSourceLeaseBlockedBy[sid] = perSourceLeaseBlockedBy[sid] || null;
    });

    updateStatusBox();
    console.log("[Autocomplete] IDB cache loaded:", contacts.length, "effective entries");
  }

  function broadcastCacheUpdated() {
    bc.postMessage({ type: "cacheUpdated", ts: Date.now() });
  }

  // ---------------------------
  // Refresh logic
  // ---------------------------
  function anySourceNeedsRefresh() {
    return Object.keys(SOURCES).some(sid => {
      if (sid === "customtext" && !customTextHasUrls) return false;
      const s = perSourceState[sid];
      return !s || !s.hasAny || s.isStale;
    });
  }

  async function forceStaleAndRefresh(sourceId) {
    if (sourceId === "customtext") {
      await refreshCustomTextHasUrlsCache();
      if (!customTextHasUrls) return;
    }

    await setMeta(metaKey(sourceId, "lastUpdate"), 0);
    clearSourceHoldoff(sourceId);
    clearSourceError(sourceId);
    await reloadContactsFromIDB();
    backgroundRefreshIfNeeded();
  }

  async function refreshSourceIfNeeded(sourceId) {
    const src = SOURCES[sourceId];
    const st = perSourceState[sourceId] || { hasAny: false, isStale: true, lastUpdate: 0 };

    if (sourceId === "customtext" && !customTextHasUrls) return;
    if (st.hasAny && !st.isStale) return;

    if (holdoffRemainingMs(sourceId) > 0) {
      perSourceRefreshing[sourceId] = false;
      updateStatusBox();
      return;
    }

    perSourceLeaseBlockedBy[sourceId] = null;

    const lease = await tryAcquireRefreshLease(sourceId);
    if (!lease.ok) {
      perSourceRefreshing[sourceId] = false;
      perSourceLeaseBlockedBy[sourceId] = lease.owner || "other";
      updateStatusBox();
      return;
    }

    perSourceRefreshing[sourceId] = true;
    perSourceProgress[sourceId] = perSourceProgress[sourceId] || { fetched: 0 };
    updateStatusBox();

    try {
      await reloadContactsFromIDB();
      const st2 = perSourceState[sourceId] || { hasAny: false, isStale: true, lastUpdate: 0 };
      if (st2.hasAny && !st2.isStale) return;

      const syncID = Date.now();
      const raw = await src.fetch();
      const fresh = src.transform(raw, syncID);

      if (sourceId === "customtext" && !customTextHasUrls) return;

      await saveSourceToIndexedDB(sourceId, fresh);
      await cleanupOldRecords(sourceId, syncID);

      clearSourceError(sourceId);
      clearSourceHoldoff(sourceId);
      console.log(`[Autocomplete] ${sourceId} refreshed:`, fresh.length);
    } catch (e) {
      setSourceError(sourceId, classifyError(e));
      console.error(`[Autocomplete] Error refreshing ${sourceId} (full):`, e, e?._gm);

      if (isNetworkError(e) || isAuthError(e)) {
        setSourceHoldoff(sourceId, HOLDOFF_MS);
      }
    } finally {
      perSourceRefreshing[sourceId] = false;
      perSourceLeaseBlockedBy[sourceId] = null;
      updateStatusBox();
      await releaseRefreshLease(sourceId, lease.owner);
    }
  }

  async function backgroundRefreshIfNeeded() {
    if (refreshInFlightLocal) return;
    if (!anySourceNeedsRefresh()) return;

    refreshInFlightLocal = true;
    updateStatusBox();

    try {
      for (const sid of Object.keys(SOURCES)) {
        await refreshSourceIfNeeded(sid);
        broadcastCacheUpdated();
      }
      await reloadContactsFromIDB();
      updateAllDropdownsLoadingHint();
    } finally {
      refreshInFlightLocal = false;
      updateStatusBox();
    }
  }

  // ---------------------------
  // Autocomplete UI
  // ---------------------------
  function getQueryWords(value) {
    return String(value || "")
      .toLowerCase()
      .split(/\s+/)
      .map(w => w.trim())
      .filter(Boolean);
  }

  function replaceFullInput(input, replacement) {
    input.value = replacement;
    const newPos = replacement.length;
    input.setSelectionRange(newPos, newPos);
  }

  function createDropdown(input) {
    const box = document.createElement("div");
    Object.assign(box.style, {
      position: "absolute",
      border: "1px solid #ccc",
      background: "#fff",
      zIndex: 9999,
      maxHeight: "200px",
      overflowY: "auto",
      fontSize: "14px",
      display: "none",
      boxShadow: "0px 4px 6px rgba(0,0,0,0.1)"
    });

    document.body.appendChild(box);
    let selectedIndex = -1;

    function position() {
      const r = input.getBoundingClientRect();
      box.style.left = r.left + window.scrollX + "px";
      box.style.top = r.bottom + window.scrollY + "px";
      box.style.width = r.width + "px";
    }

    function render(matches) {
      box.innerHTML = "";
      selectedIndex = -1;

      matches.forEach(m => {
        const el = document.createElement("div");
        el.dataset.value = m.text;
        el.style.display = "flex";
        el.style.alignItems = "baseline";
        el.style.gap = "8px";

        const main = document.createElement("span");
        main.textContent = m.text;
        main.style.flex = "1 1 auto";
        main.style.minWidth = "0";
        el.appendChild(main);

        if (m.source) {
          const src = document.createElement("span");
          src.textContent = m.source;
          src.style.flex = "0 0 auto";
          src.style.marginLeft = "auto";
          src.style.fontStyle = "italic";
          src.style.color = "#666";
          el.appendChild(src);
        }

        el.style.padding = "4px";
        el.style.cursor = m.disabled ? "default" : "pointer";
        el.style.color = m.disabled ? "#666" : "#000";
        el.style.background = m.disabled ? "#f7f7f7" : "#fff";

        if (!m.disabled) {
          el.addEventListener("mousedown", e => {
            e.preventDefault();
            replaceFullInput(input, m.text);
            hide();
          });
        }
        box.appendChild(el);
      });

      position();
      box.style.display = matches.length ? "block" : "none";
    }

    function move(delta) {
      const count = box.children.length;
      if (!count) return;
      selectedIndex = (selectedIndex + delta + count) % count;
      [...box.children].forEach((el, i) => {
        el.style.background = i === selectedIndex ? "#def" : (el.style.cursor === "pointer" ? "#fff" : "#f7f7f7");
      });
    }

    function choose() {
      if (selectedIndex >= 0) {
        const el = box.children[selectedIndex];
        if (el && el.style.cursor === "pointer") {
          replaceFullInput(input, el.dataset.value || el.textContent);
          hide();
        }
      }
    }

    function hide() {
      box.style.display = "none";
      selectedIndex = -1;
    }

    return { render, move, choose, hide };
  }

  function queryMatchesFromValue(inputValue) {
    const words = getQueryWords(inputValue);
    if (words.length === 0) return [];

    const matches = contacts
      .filter(c => words.every(w => c.search.includes(w)))
      .slice(0, 10)
      .map(c => ({ text: c.text, source: c.source }));

    const anyMissing = Object.keys(SOURCES).some(sid => {
      if (sid === "customtext" && !customTextHasUrls) return false;
      return !(perSourceState[sid]?.hasAny);
    });

    if (anyMissing && refreshInFlightLocal) {
      matches.unshift({ text: "Loading contacts… (cache is being filled)", disabled: true });
    }
    return matches;
  }

  function updateAllDropdownsLoadingHint() {
    for (const inst of autocompleteInstances) {
      const input = inst.input;
      const dropdown = inst.dropdown;
      if (!input) continue;
      const words = getQueryWords(input.value);
      if (words.join(" ").length >= 2) dropdown.render(queryMatchesFromValue(input.value));
    }
  }

  function attachAutocomplete(input) {
    if (!input) return;
    const dropdown = createDropdown(input);
    autocompleteInstances.push({ input, dropdown });

    input.addEventListener("input", () => {
      const words = getQueryWords(input.value);
      if (words.join(" ").length < 2) {
        dropdown.hide();
        return;
      }
      dropdown.render(queryMatchesFromValue(input.value));
    });

    input.addEventListener("keydown", e => {
      if (e.key === "ArrowDown") {
        dropdown.move(1);
        e.preventDefault();
      } else if (e.key === "ArrowUp") {
        dropdown.move(-1);
        e.preventDefault();
      } else if (e.key === "Enter") {
        if (document.querySelector('[style*="display: block"]')) {
          dropdown.choose();
          e.preventDefault();
        }
      } else if (e.key === "Escape") {
        dropdown.hide();
      }
    });

    input.addEventListener("blur", () => {
      setTimeout(() => dropdown.hide(), 200);
    });
  }

  function setupCrossTabListeners() {
    bc.onmessage = async (ev) => {
      if (ev?.data?.type === "cacheUpdated") {
        Object.keys(SOURCES).forEach(sid => { perSourceLeaseBlockedBy[sid] = null; });
        await reloadContactsFromIDB();
        updateAllDropdownsLoadingHint();
        backgroundRefreshIfNeeded();
      }
    };
  }

  // ---------------------------
  // Custom text config UI (top near status box)
  // ---------------------------
  function ensureCustomTextUrlsConfigUI() {
    const existing = document.getElementById("znunyCustomTextUrlsConfig");
    if (existing) return;

    const box = document.createElement("div");
    box.id = "znunyCustomTextUrlsConfig";
    Object.assign(box.style, {
      position: "fixed",
      top: "54px",
      right: "12px",
      zIndex: 100000,
      background: "rgba(255,255,255,0.97)",
      border: "1px solid #bbb",
      borderRadius: "6px",
      padding: "10px",
      boxShadow: "0 4px 10px rgba(0,0,0,0.15)",
      fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
      fontSize: "12px",
      color: "#111",
      width: "420px",
      maxWidth: "90vw"
    });

    box.innerHTML = `
      <div style="display:flex; align-items:center; justify-content:space-between; gap:10px;">
        <div style="font-weight:700;">Custom text sources</div>
        <button type="button" id="znunyCustomTextUrlsClose" style="cursor:pointer;">✕</button>
      </div>
      <div style="margin-top:6px; opacity:0.9;">
        URLs (one URL per line), URL must return plain text.
      </div>
      <div style="margin-top:6px; opacity:0.9;">
        Format of text files: <code>mail@example.com Full Name</code>
      </div>
      <textarea id="znunyCustomTextUrlsArea" style="width:100%; height:120px; margin-top:8px; font-family:monospace; font-size:12px;"></textarea>
      <div style="display:flex; gap:8px; margin-top:8px;">
        <button type="button" id="znunyCustomTextUrlsSave" style="cursor:pointer; font-weight:600;">Save</button>
      </div>
      <div id="znunyCustomTextUrlsHint" style="margin-top:8px; opacity:0.75;"></div>
    `;

    document.body.appendChild(box);

    const hint = box.querySelector("#znunyCustomTextUrlsHint");

    box.querySelector("#znunyCustomTextUrlsClose").addEventListener("click", async () => {
      lsSet(LS_CUSTOMTEXT_UI_ENABLED, "0");
      box.remove();
      await refreshCustomTextHasUrlsCache();
      updateStatusBox();
    });

    box.querySelector("#znunyCustomTextUrlsSave").addEventListener("click", async () => {
      const area = box.querySelector("#znunyCustomTextUrlsArea");

      // store trimmed non-empty lines only
      const lines = String(area.value || "")
        .split(/\r?\n/)
        .map(l => l.trim())
        .filter(Boolean);

      const toStore = lines.join("\n");
      lsSet(LS_CUSTOMTEXT_URLS, toStore);

      await refreshCustomTextHasUrlsCache();
      updateStatusBox();

      const ok = !!toStore.trim();
      hint.textContent = `Saved. URLs active: ${ok ? "yes" : "no"}.`;
    });
  }

  // ---------------------------
  // Init
  // ---------------------------
  async function init() {
    cleanupOldCache();
    ensureStatusBox();
    setupCrossTabListeners();

    await refreshCustomTextHasUrlsCache();

    const uiEnabled = lsGet(LS_CUSTOMTEXT_UI_ENABLED, "0") === "1";
    updateStatusBox();

    if (uiEnabled) {
      ensureCustomTextUrlsConfigUI();
      const uiVal = lsGet(LS_CUSTOMTEXT_URLS, "");
      const area = document.getElementById("znunyCustomTextUrlsArea");
      if (area) area.value = uiVal;
      await refreshCustomTextHasUrlsCache();
      updateStatusBox();
    }

    await reloadContactsFromIDB();

    ["FromCustomer", "ToCustomer", "CcCustomer", "BccCustomer"].forEach(id => {
      attachAutocomplete(document.getElementById(id));
    });

    backgroundRefreshIfNeeded();
  }

  init();
})();
