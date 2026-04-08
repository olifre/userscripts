// ==UserScript==
// @name        Znuny: Contact Autocomplete (multi-source)
// @namespace   github.com/olifre/userstyles
// @match       https://support.uni-bonn.de/*
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-contact-completion.user.js
// @downloadURL https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-contact-completion.user.js
// @version     1.5.0
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
  )) {
    return;
  }

  // ---------------------------
  // Sources (extendable)
  // ---------------------------

  // In-memory progress tracking (per source)
  const perSourceProgress = {}; // { [sourceId]: { fetched } }

  // In-memory last error (per source)
  const perSourceError = {}; // { [sourceId]: { message: string, ts: number } }

  // In-memory hold-off after network/auth errors (per source, per tab)
  const perSourceHoldoffUntil = {}; // { [sourceId]: number (ts ms) }
  const HOLDOFF_MS = 60 * 1000;

  // Timer to wake up exactly when hold-off ends (per source)
  const perSourceHoldoffTimer = {}; // { [sourceId]: number }

  // In-memory: whether this tab is actively refreshing a given source
  const perSourceRefreshing = {}; // { [sourceId]: boolean }

  // Broadcast channel and helper
  const BC_NAME = "znuny_contacts_bc_v2";
  const bc = new BroadcastChannel(BC_NAME);

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
      priority: 50, // lower than phy
      cacheTtlMs: 24 * 60 * 60 * 1000,
      refreshLeaseTtlMs: 2 * 60 * 1000,
      url: "https://jira.team.uni-bonn.de/rest/api/2/user/list",
      batchSize: 2000,

      async fetch() {        // Start from scratch for Jira (do not persist resume cursors)
        let cursor = 0;

        const sleep = (ms) => new Promise(r => setTimeout(r, ms));

        const results = [];
        let totalFetched = 0;
        let isDone = false;

        // Make sure the UI shows "refreshing" while we are inside this loop.
        perSourceRefreshing[this.id] = true;
        updateStatusBox();

        while (!isDone) {
          const u = new URL(this.url);
          u.searchParams.set("cursor", cursor);
          u.searchParams.set("maxResults", this.batchSize);

          const json = await xFetchJson(u.toString(), { withCredentials: true });

          const batch = Array.isArray(json.values) ? json.values : [];
          results.push(...batch);

          // Update every successful batch (= every 2000, except maybe the last one)
          totalFetched += batch.length;
          perSourceProgress[this.id] = { fetched: totalFetched };
          updateStatusBox();

          // Determine next cursor, but do not persist it
          const nc = json.nextCursor;
          const nextCursorNum = (typeof nc === "number") ? nc : (typeof nc === "string" ? Number(nc) : NaN);
          if (!Number.isNaN(nextCursorNum)) {
            cursor = nextCursorNum;
          }

          if (json.isLast) isDone = true;

          // Sleep a bit to be gentle between requests (as requested)
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
    }
  };

  // ---------------------------
  // Storage / IPC constants
  // ---------------------------
  const OLD_CACHE_KEY = "znuny_contacts_cache_v1";

  const DB_NAME = "ZnunyContactDB";
  const DB_VERSION = 4;

  // ---------------------------
  // In-memory state
  // ---------------------------
  let contacts = []; // effective list: unique by email, priority wins
  let perSourceState = {}; // { [sourceId]: { hasAny, isStale, lastUpdate } }
  let perSourceCounts = {}; // { [sourceId]: number of raw rows in IDB for that source }
  let autocompleteInstances = []; // { input, dropdown }
  let refreshInFlightLocal = false;

  // ---------------------------
  // Status UI
  // ---------------------------
  let statusBox = null;

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

    // schedule exactly one wake-up when hold-off expires
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

    // Show "refreshing" whenever this tab is actively refreshing the source.
    const refreshing = !!perSourceRefreshing[sourceId] && !onHoldoff;
    const fetchingText = refreshing ? " (refreshing…)" : "";
    const fetchPart = refreshing ? ` | fetch: ${p.fetched ?? 0}` : "";

    const holdPart = onHoldoff ? ` | hold-off: ${fmtRemaining(holdMs)}` : "";

    const err = perSourceError[sourceId]?.message;
    const errPart = err ? ` | error: ${err}` : "";

    return `${sourceId}: ${freshness}${fetchingText}${fetchPart} | contacts: ${n} | last fetch: ${fmtIso(s.lastUpdate)}${holdPart}${errPart}`;
  }

  async function forceStaleAndRefresh(sourceId) {
    // Mark stale
    await setMeta(metaKey(sourceId, "lastUpdate"), 0);

    // Clear "blocked" states immediately
    clearSourceHoldoff(sourceId);
    clearSourceError(sourceId);

    await reloadContactsFromIDB();
    backgroundRefreshIfNeeded();
  }

  function updateStatusBox() {
    const box = ensureStatusBox();
    box.innerHTML = "";

    for (const sid of Object.keys(SOURCES)) {
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
  }

  // Update countdown once per second (also do a final update when a hold-off just ended)
  setInterval(() => {
    const anyHoldOrJustEnded = Object.keys(SOURCES).some(sid => {
      const rem = holdoffRemainingMs(sid);
      return rem > 0 || (perSourceHoldoffUntil[sid] && rem === 0);
    });
    if (anyHoldOrJustEnded) updateStatusBox();
  }, 1000);

  // ---------------------------
  // Helpers
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
          catch (e) {
            e._gm = r;
            reject(e);
          }
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
    return `meta:source:${sourceId}:${field}`; // field: lastUpdate | lease
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

  // Per-source refresh lease (cross-tab lock).
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

    // Update simple per-source count for UI (not strictly required)
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
    const bestByEmail = new Map(); // email -> row

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
      const s = perSourceState[sid];
      return !s || !s.hasAny || s.isStale;
    });
  }

  async function refreshSourceIfNeeded(sourceId) {
    const src = SOURCES[sourceId];
    const st = perSourceState[sourceId] || { hasAny: false, isStale: true, lastUpdate: 0 };

    if (st.hasAny && !st.isStale) return;

    // Hold-off: do not try to refresh while on hold-off
    if (holdoffRemainingMs(sourceId) > 0) {
      perSourceRefreshing[sourceId] = false;
      updateStatusBox();
      return;
    }

    const lease = await tryAcquireRefreshLease(sourceId);
    if (!lease.ok) {
      perSourceRefreshing[sourceId] = false;
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

      await saveSourceToIndexedDB(sourceId, fresh);
      await cleanupOldRecords(sourceId, syncID);

      clearSourceError(sourceId);
      clearSourceHoldoff(sourceId);
      console.log(`[Autocomplete] ${sourceId} refreshed:`, fresh.length);
    } catch (e) {
      setSourceError(sourceId, classifyError(e));

      // Log full error (and GM response details if present)
      console.error(`[Autocomplete] Error refreshing ${sourceId} (full):`, e, e?._gm);

      if (isNetworkError(e) || isAuthError(e)) {
        setSourceHoldoff(sourceId, HOLDOFF_MS);
      }
    } finally {
      perSourceRefreshing[sourceId] = false;
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

  function getCurrentToken(value, cursorPos) {
    return value.slice(0, cursorPos).split(",").pop().trim();
  }

  function replaceCurrentToken(input, replacement) {
    const pos = input.selectionStart;
    const left = input.value.slice(0, pos);
    const right = input.value.slice(pos);

    const parts = left.split(",");
    parts[parts.length - 1] = " " + replacement;

    const newLeft = parts.join(",").replace(/^,/, "");
    input.value = newLeft + ", " + right.replace(/^,?\s*/, "");

    const newPos = newLeft.length + 2;
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
            replaceCurrentToken(input, m.text);
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
          replaceCurrentToken(input, el.dataset.value || el.textContent);
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

  function queryMatches(tokenLower) {
    const matches = contacts
      .filter(c => c.search.includes(tokenLower))
      .slice(0, 10)
      .map(c => ({ text: c.text, source: c.source }));

    const anyMissing = Object.keys(SOURCES).some(sid => !(perSourceState[sid]?.hasAny));
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
      const token = getCurrentToken(input.value, input.selectionStart).toLowerCase();
      if (token.length >= 2) dropdown.render(queryMatches(token));
    }
  }

  function attachAutocomplete(input) {
    if (!input) return;
    const dropdown = createDropdown(input);
    autocompleteInstances.push({ input, dropdown });

    input.addEventListener("input", () => {
      const token = getCurrentToken(input.value, input.selectionStart).toLowerCase();
      if (token.length < 2) {
        dropdown.hide();
        return;
      }
      dropdown.render(queryMatches(token));
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
        await reloadContactsFromIDB();
        updateAllDropdownsLoadingHint();
        backgroundRefreshIfNeeded();
      }
    };
  }

  async function init() {
    cleanupOldCache();
    ensureStatusBox();
    updateStatusBox();
    setupCrossTabListeners();

    await reloadContactsFromIDB();

    ["FromCustomer", "ToCustomer", "CcCustomer", "BccCustomer"].forEach(id => {
      attachAutocomplete(document.getElementById(id));
    });

    backgroundRefreshIfNeeded();
  }

  init();
})();
