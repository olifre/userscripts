// ==UserScript==
// @name        Znuny: Contact Autocomplete (multi-source)
// @namespace   github.com/olifre/userstyles
// @match       https://support.uni-bonn.de/*
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-contact-completion.user.js
// @version     1.4.2
// @grant       none
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
        const res = await fetch(this.url);
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.json();
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
      cacheTtlMs: 6 * 60 * 60 * 1000,
      refreshLeaseTtlMs: 2 * 60 * 1000,
      url: "https://jira.team.uni-bonn.de/rest/api/2/user/list",
      batchSize: 2000,

      async fetch() {
        // Sleep helper
        const sleep = (ms) => new Promise(r => setTimeout(r, ms));

        // Start at last saved cursor (default 0)
        const lastCursorVal = await getMeta(metaKey(this.id, "cursor"));
        let cursor = (typeof lastCursorVal === "number" ? lastCursorVal : 0);

        const results = [];
        let totalFetched = 0;
        let isDone = false;

        while (!isDone) {
          const u = new URL(this.url);
          u.searchParams.set("cursor", cursor);
          u.searchParams.set("maxResults", this.batchSize);

          const res = await fetch(u.toString(), { credentials: "include" });
          if (!res.ok) throw new Error("HTTP " + res.status);
          const json = await res.json();

          const batch = Array.isArray(json.values) ? json.values : [];
          results.push(...batch);

          // Incremental progress for status box
          totalFetched += batch.length;
          perSourceProgress[this.id] = { fetched: totalFetched };
          updateStatusBox();

          // Persist next cursor for resuming
          // handle nextCursor which may be a string
          const nc = json.nextCursor;
          const nextCursorNum = (typeof nc === "number") ? nc : (typeof nc === "string" ? Number(nc) : NaN);
          if (!Number.isNaN(nextCursorNum)) {
            cursor = nextCursorNum;
            await setMeta(metaKey(this.id, "cursor"), cursor);
          }

          if (json.isLast) isDone = true;

          // Sleep a bit to be gentle between requests (as requested)
          await sleep(200);
        }

        // After all pages fetched, broadcast once
        if (results.length > 0) {
          broadcastCacheUpdated();
        }

        // Final progress
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
      maxWidth: "720px",
      pointerEvents: "none",
      whiteSpace: "pre"
    });

    document.body.appendChild(box);
    return box;
  }

  function fmtIso(ts) {
    if (!ts) return "never";
    try { return new Date(ts).toISOString(); } catch { return "invalid"; }
  }

  function statusLine(sourceId) {
    const s = perSourceState[sourceId] || { hasAny: false, isStale: true, lastUpdate: 0 };
    const n = perSourceCounts[sourceId] ?? 0;
    const p = perSourceProgress[sourceId] || { fetched: 0 };
    const freshness = s.hasAny ? (s.isStale ? "stale" : "fresh") : "empty";

    // Show fetch progress only while actively refreshing
    const refreshing = refreshInFlightLocal && (s.isStale || !s.hasAny);
    const fetchingText = refreshing ? " (refreshing…)" : "";
    const fetchPart = refreshing ? ` | fetch: ${p.fetched ?? 0}` : "";

    return `${sourceId}: ${freshness}${fetchingText}${fetchPart} | contacts: ${n} | last fetch: ${fmtIso(s.lastUpdate)}`;
  }

  function updateStatusBox() {
    const box = ensureStatusBox();
    const lines = Object.keys(SOURCES).map(statusLine);
    box.textContent = lines.join("\n");
  }

  // ---------------------------
  // Helpers
  // ---------------------------
  function cleanupOldCache() {
    if (localStorage.getItem(OLD_CACHE_KEY)) {
      localStorage.removeItem(OLD_CACHE_KEY);
      console.log("[Autocomplete] Old localStorage cache cleared.");
    }
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
        // initialize progress tracking
        perSourceProgress[sourceId] = perSourceProgress[sourceId] || { fetched: 0 };
      }));

      return { contacts: allContacts, state, counts };
    } catch (e) {
      console.error("[Autocomplete] IDB Error:", e);
      const state = {};
      const counts = {};
      for (const sid of Object.keys(SOURCES)) state[sid] = { hasAny: false, isStale: true, lastUpdate: 0 };
      for (const sid of Object.keys(SOURCES)) counts[sid] = 0;
      // initialize progress placeholders
      Object.keys(SOURCES).forEach(sid => perSourceProgress[sid] = { fetched: 0 });
      return { contacts: [], state, counts };
    }
  }

  async function reloadContactsFromIDB() {
    const { contacts: allRows, state, counts } = await loadFromIndexedDBWithMeta();
    perSourceState = state;
    perSourceCounts = counts;
    contacts = buildEffectiveContacts(allRows);
    // reset per-source progress for fresh UI state
    Object.keys(SOURCES).forEach(sid => {
      perSourceProgress[sid] = perSourceProgress[sid] || { fetched: 0 };
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

    const lease = await tryAcquireRefreshLease(sourceId);
    if (!lease.ok) return; // another tab is refreshing this source

    try {
      await reloadContactsFromIDB();
      const st2 = perSourceState[sourceId] || { hasAny: false, isStale: true, lastUpdate: 0 };
      if (st2.hasAny && !st2.isStale) return;

      const syncID = Date.now();

      const raw = await src.fetch();
      const fresh = src.transform(raw, syncID);

      await saveSourceToIndexedDB(sourceId, fresh);
      await cleanupOldRecords(sourceId, syncID);

      console.log(`[Autocomplete] ${sourceId} refreshed:`, fresh.length);
    } catch (e) {
      console.error(`[Autocomplete] Error refreshing ${sourceId}:`, e);
    } finally {
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
      }
      // Broadcast only after all sources finished refreshing
      broadcastCacheUpdated();
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
