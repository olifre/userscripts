// ==UserScript==
// @name        Znuny: Contact Autocomplete for Physics Institute Members
// @namespace   github.com/olifre/userstyles
// @match       https://support.uni-bonn.de/*
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-pi-members.user.js
// @version     1.1.0
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

  const DATA_URL = "https://grp_phy.gitlab-pages.uni-bonn.de/it/web/vcard_generator/contacts.json";
  const OLD_CACHE_KEY = "znuny_contacts_cache_v1";
  const DB_NAME = "ZnunyContactDB";
  const DB_VERSION = 1;
  const CACHE_TTL = 12 * 60 * 60 * 1000;

  let contacts = [];

  function cleanupOldCache() {
    if (localStorage.getItem(OLD_CACHE_KEY)) {
      localStorage.removeItem(OLD_CACHE_KEY);
      console.log("[Autocomplete] Old localStorage cache cleared.");
    }
  }

  // Create DB.
  function openDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        if (!db.objectStoreNames.contains("contacts")) {
          db.createObjectStore("contacts", { autoIncrement: true });
        }
        if (!db.objectStoreNames.contains("metadata")) {
          db.createObjectStore("metadata");
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function saveToIndexedDB(data) {
    const db = await openDB();
    const tx = db.transaction(["contacts", "metadata"], "readwrite");
    const contactStore = tx.objectStore("contacts");
    const metaStore = tx.objectStore("metadata");

    contactStore.clear();
    data.forEach(item => contactStore.add(item));
    metaStore.put(Date.now(), "lastUpdate");

    return new Promise((resolve) => {
      tx.oncomplete = () => resolve();
    });
  }

  async function loadFromIndexedDB() {
    try {
      const db = await openDB();
      const tx = db.transaction(["contacts", "metadata"], "readonly");
      const contactStore = tx.objectStore("contacts");
      const metaStore = tx.objectStore("metadata");

      const lastUpdate = await new Promise(res => {
        const req = metaStore.get("lastUpdate");
        req.onsuccess = () => res(req.result);
      });

      if (!lastUpdate || (Date.now() - lastUpdate > CACHE_TTL)) {
        return null;
      }

      return new Promise((resolve) => {
        const req = contactStore.getAll();
        req.onsuccess = () => {
          console.log("[Autocomplete] IDB Cache used:", req.result.length, "contacts");
          resolve(req.result);
        };
      });
    } catch (e) {
      console.error("[Autocomplete] IDB Error:", e);
      return null;
    }
  }

  async function fetchData() {
    const res = await fetch(DATA_URL);
    if (!res.ok) throw new Error("HTTP " + res.status);
    return res.json();
  }

  function transform(data) {
    const entries = Object.values(data);
    return entries.map(e => {
      const text = `"${e.degree ? e.degree + " " : ""}${e.firstname || ""} ${e.lastname || ""}" <${e.email || ""}>`
        .replace(/\s+/g, ' ')
        .trim();
      const search = text.toLowerCase();
      return { text, search };
    });
  }

  async function loadData() {
    cleanupOldCache();

    const cached = await loadFromIndexedDB();

    if (cached && cached.length > 0) {
      contacts = cached;
      return;
    }

    try {
      const data = await fetchData();
      contacts = transform(data);
      await saveToIndexedDB(contacts);
      console.log("[Autocomplete] Data reloaded and stored in IDB");
    } catch (e) {
      console.error("[Autocomplete] Error loading data:", e);
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
        el.style.cursor = "pointer";
        el.addEventListener("mousedown", e => {
          e.preventDefault();
          replaceCurrentToken(input, m.text);
          hide();
        });
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
        el.style.background = i === selectedIndex ? "#def" : "#fff";
      });
    }

    function choose() {
      if (selectedIndex >= 0) {
        replaceCurrentToken(input, box.children[selectedIndex].textContent);
        hide();
      }
    }

    function hide() {
      box.style.display = "none";
      selectedIndex = -1;
    }

    return { render, move, choose, hide };
  }

  function attachAutocomplete(input) {
    if (!input) return;
    const dropdown = createDropdown(input);

    input.addEventListener("input", () => {
      const token = getCurrentToken(input.value, input.selectionStart).toLowerCase();
      if (token.length < 2) {
        dropdown.hide();
        return;
      }
      const matches = contacts
        .filter(c => c.search.includes(token))
        .slice(0, 10);
      dropdown.render(matches);
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

  async function init() {
    await loadData();
    ["FromCustomer", "ToCustomer", "CcCustomer", "BccCustomer"].forEach(id => {
      attachAutocomplete(document.getElementById(id));
    });
  }

  init();
})();
