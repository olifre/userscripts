// ==UserScript==
// @name        Znuny: Contact Autocomplete for Physics Institute Members
// @namespace   github.com/olifre/userstyles
// @match       https://support.uni-bonn.de/*
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-pi-members.user.js
// @version     1.2.0
// @grant       none
// @description Autocomplete for Znuny contacts based on public Physics institute member data
// @author      Oliver Freyermuth <o.freyermuth@googlemail.com> (https://olifre.github.io/)
// @license     Unlicense
// ==/UserScript==

(function () {
  'use strict';

  // Filter on relevant pages.
  if (!(/\bAction=AgentTicketCompose\b/.test(location.search) ||
      /\bAction=AgentTicketEmail\b/.test(location.search) ||
      /\bAction=AgentTicketEmailOutbound\b/.test(location.search) ||
      /\bAction=AgentTicketPhoneOutbound\b/.test(location.search) ||
      /\bAction=AgentTicketPhoneInbound\b/.test(location.search) ||
      /\bAction=AgentTicketPhone\b/.test(location.search) ||
      /\bAction=AgentTicketForward\b/.test(location.search))) {
    return;
  }

  const PHY_DATA_URL = "https://grp_phy.gitlab-pages.uni-bonn.de/it/web/vcard_generator/contacts.json";
  const OLD_CACHE_KEY = "znuny_contacts_cache_v1";
  const DB_NAME = "ZnunyContactDB";
  const DB_VERSION = 2;
  const PHY_CACHE_TTL = 12 * 60 * 60 * 1000;

  // IDB-based refresh lease (cross-tab lock)
  const PHY_REFRESH_LEASE_KEY = "phyRefreshLease";
  const PHY_REFRESH_LEASE_TTL = 2 * 60 * 1000;

  // Cross-tab broadcast (modern browsers)
  const BC_NAME = "znuny_contacts_bc_v1";
  const bc = new BroadcastChannel(BC_NAME);

  let contacts = [];
  let cacheState = { hasAny: false, isStale: true, lastUpdate: 0 };
  let autocompleteInstances = []; // { input, dropdown }
  let refreshInFlightLocal = false;

  function cleanupOldCache() {
    if (localStorage.getItem(OLD_CACHE_KEY)) {
      localStorage.removeItem(OLD_CACHE_KEY);
      console.log("[Autocomplete] Old localStorage cache cleared.");
    }
  }

  function openDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        if (db.objectStoreNames.contains("contacts")) {
          db.deleteObjectStore("contacts");
        }
        db.createObjectStore("contacts", { keyPath: "email" });

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
    const store = tx.objectStore("metadata");
    store.put(value, key);
    return new Promise((resolve) => { tx.oncomplete = () => resolve(); });
  }

  // One-readwrite-transaction lease acquisition to serialize refreshes across tabs.
  async function tryAcquireRefreshLease() {
    const now = Date.now();
    const owner = `${now}-${Math.random().toString(16).slice(2)}`;

    const db = await openDB();
    const tx = db.transaction("metadata", "readwrite");
    const store = tx.objectStore("metadata");

    const current = await new Promise((resolve) => {
      const req = store.get(PHY_REFRESH_LEASE_KEY);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    });

    if (current && current.expiresAt && current.expiresAt > now) {
      return { ok: false, owner: current.owner };
    }

    const lease = { owner, expiresAt: now + PHY_REFRESH_LEASE_TTL };
    store.put(lease, PHY_REFRESH_LEASE_KEY);

    return await new Promise((resolve) => {
      tx.oncomplete = async () => {
        const verify = await getMeta(PHY_REFRESH_LEASE_KEY);
        resolve({ ok: !!(verify && verify.owner === owner), owner });
      };
      tx.onerror = () => resolve({ ok: false, owner: null });
      tx.onabort = () => resolve({ ok: false, owner: null });
    });
  }

  async function releaseRefreshLease(owner) {
    const cur = await getMeta(PHY_REFRESH_LEASE_KEY);
    if (cur && cur.owner === owner) {
      await setMeta(PHY_REFRESH_LEASE_KEY, { owner: "", expiresAt: 0 });
    }
  }

  async function saveToIndexedDB(data) {
    const db = await openDB();
    const tx = db.transaction(["contacts", "metadata"], "readwrite");
    const contactStore = tx.objectStore("contacts");
    const metaStore = tx.objectStore("metadata");

    // Use put to handle updates via email key, live incremental updates. Purge happens after sync.
    data.forEach(item => contactStore.put(item));
    metaStore.put(Date.now(), "lastPhyUpdate");

    return new Promise((resolve) => {
      tx.oncomplete = () => resolve();
    });
  }

  // Actual cleanup of older entries after full sync.
  async function cleanupOldRecords(source, validSyncID) {
    const db = await openDB();
    const tx = db.transaction("contacts", "readwrite");
    const store = tx.objectStore("contacts");
    const request = store.openCursor();

    request.onsuccess = (event) => {
      const cursor = event.target.result;
      if (cursor) {
        if (cursor.value.source === source && cursor.value.syncID !== validSyncID) {
          cursor.delete();
        }
        cursor.continue();
      }
    };
  }

  // Cache-first: always return whatever IDB has (even if stale), plus metadata about staleness.
  async function loadFromIndexedDBWithMeta() {
    try {
      const db = await openDB();
      const tx = db.transaction(["contacts", "metadata"], "readonly");
      const contactStore = tx.objectStore("contacts");
      const metaStore = tx.objectStore("metadata");

      const lastPhyUpdate = await new Promise(res => {
        const req = metaStore.get("lastPhyUpdate");
        req.onsuccess = () => res(req.result || 0);
        req.onerror = () => res(0);
      });

      const allContacts = await new Promise((resolve) => {
        const req = contactStore.getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => resolve([]);
      });

      const hasAny = allContacts.length > 0;
      const isStale = !lastPhyUpdate || (Date.now() - lastPhyUpdate > PHY_CACHE_TTL);
      return { contacts: allContacts, meta: { lastPhyUpdate, hasAny, isStale } };
    } catch (e) {
      console.error("[Autocomplete] IDB Error:", e);
      return { contacts: [], meta: { lastPhyUpdate: 0, hasAny: false, isStale: true } };
    }
  }

  async function reloadContactsFromIDB() {
    const { contacts: c, meta } = await loadFromIndexedDBWithMeta();
    contacts = c;
    cacheState = { hasAny: meta.hasAny, isStale: meta.isStale, lastUpdate: meta.lastPhyUpdate };
    console.log("[Autocomplete] IDB cache loaded:", contacts.length, "stale=", cacheState.isStale);
  }

  function broadcastCacheUpdated() {
    bc.postMessage({ type: "cacheUpdated", ts: Date.now() });
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

  async function fetchPhyData() {
    const res = await fetch(PHY_DATA_URL);
    if (!res.ok) throw new Error("HTTP " + res.status);
    return res.json();
  }

  function transformPhy(data, syncID) {
    const entries = Object.values(data);
    return entries.map(e => {
      const text = `"${e.degree ? e.degree + " " : ""}${e.firstname || ""} ${e.lastname || ""}" <${e.email || ""}>`
        .replace(/\s+/g, ' ')
        .trim();
      const search = text.toLowerCase();
      return {
        email: (e.email || "").toLowerCase(),
        text,
        search,
        source: 'phy',
        syncID: syncID
      };
    }).filter(x => x.email); // keyPath must be present
  }

  async function backgroundRefreshIfNeeded() {
    if (refreshInFlightLocal) return;
    if (cacheState.hasAny && !cacheState.isStale) return;

    // If cache empty, let UI show "Loading..." while refresh is pending.
    if (!cacheState.hasAny) updateAllDropdownsLoadingHint();

    const lease = await tryAcquireRefreshLease();
    if (!lease.ok) return; // another tab is refreshing

    refreshInFlightLocal = true;
    if (!cacheState.hasAny) updateAllDropdownsLoadingHint();

    try {
      // Re-check after lease acquisition: another tab may have refreshed just before we got the lease.
      await reloadContactsFromIDB();
      if (cacheState.hasAny && !cacheState.isStale) {
        updateAllDropdownsLoadingHint();
        return;
      }

      const syncID = Date.now();
      const data = await fetchPhyData();
      const fresh = transformPhy(data, syncID);
      await saveToIndexedDB(fresh);
      await cleanupOldRecords('phy', syncID);
      console.log("[Autocomplete] PHY Data refreshed and stored in IDB");

      broadcastCacheUpdated();
      await reloadContactsFromIDB(); // update current tab too
      updateAllDropdownsLoadingHint();
    } catch (e) {
      console.error("[Autocomplete] Error refreshing data:", e);
    } finally {
      refreshInFlightLocal = false;
      await releaseRefreshLease(lease.owner);
    }
  }

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
        el.textContent = m.text;
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
          replaceCurrentToken(input, el.textContent);
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
      .map(c => ({ text: c.text }));

    // Show dropdown entry while cache is empty and being filled.
    if (!cacheState.hasAny && (refreshInFlightLocal || cacheState.isStale)) {
      matches.unshift({ text: "Loading contacts… (cache is being filled)", disabled: true });
    }
    return matches;
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
        // Ensures this tab doesn't decide to refresh based on outdated local state.
        backgroundRefreshIfNeeded();
      }
    };
  }

  async function init() {
    cleanupOldCache();
    setupCrossTabListeners();

    // Cache-first: load whatever is available immediately (even if stale).
    await reloadContactsFromIDB();

    ["FromCustomer", "ToCustomer", "CcCustomer", "BccCustomer"].forEach(id => {
      attachAutocomplete(document.getElementById(id));
    });

    // Background refresh if cache missing or stale (serialized via IDB refresh lease).
    backgroundRefreshIfNeeded();
  }

  init();
})();
