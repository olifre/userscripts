// ==UserScript==
// @name        Znuny: Contact Autocomplete for Physics Institute Members
// @namespace   github.com/olifre/userstyles
// @match       https://support.uni-bonn.de/*
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de-pi-members.user.js
// @version     1.0
// @grant       none
// @description Autocomplete for Znuny contacts based on public Physics institute member data
// @author      Oliver Freyermuth <o.freyermuth@googlemail.com> (https://olifre.github.io/)
// @license     Unlicense
// ==/UserScript==

(function () {
  'use strict';

  // Filter on relevant pages.
  if (!(/\bAction=AgentTicketCompose\b/.test (location.search) ||
      /\bAction=AgentTicketEmail\b/.test (location.search) ||
      /\bAction=AgentTicketEmailOutbound\b/.test (location.search) ||
      /\bAction=AgentTicketPhoneOutbound\b/.test (location.search) ||
      /\bAction=AgentTicketPhoneInbound\b/.test (location.search) ||
      /\bAction=AgentTicketForward\b/.test (location.search) )) {
      	return;
  }

  const DATA_URL = "https://grp_phy.gitlab-pages.uni-bonn.de/it/web/vcard_generator/contacts.json";

  const CACHE_KEY = "znuny_contacts_cache_v1";
  const CACHE_TTL = 24 * 60 * 60 * 1000;

  let contacts = [];

  function saveCache(data) {
    localStorage.setItem(CACHE_KEY, JSON.stringify({
      timestamp: Date.now(),
      data
    }));
  }

  function loadCache() {
    try {
      const raw = localStorage.getItem(CACHE_KEY);
      if (!raw) return null;

      const parsed = JSON.parse(raw);

      if (Date.now() - parsed.timestamp > CACHE_TTL) {
        return null;
      }

      console.log("[Autocomplete] Cache used:", parsed.data.length, "contacts");
      return parsed.data;

    } catch {
      return null;
    }
  }

  async function fetchData() {
    const res = await fetch(DATA_URL);
    if (!res.ok) throw new Error("HTTP " + res.status);
    return res.json();
  }

  function transform(data) {
    // Objekt → Werte extrahieren
    const entries = Object.values(data);

    // Map zu Objekten mit text und search
    return entries.map(e => {
      const text = `${e.degree || ""} ${e.firstname || ""} ${e.lastname || ""} <${e.email || ""}>`
      .replace(/\s+/g, ' ')
      .trim();

      const search = `${text} ${e.function_tags || ""} ${e.group_tags || ""}`.toLowerCase();

      return { text, search };
    });
  }

  async function loadData() {
    const cached = loadCache();

    if (cached) {
      contacts = cached;
      return;
    }

    try {
      const data = await fetchData();
      contacts = transform(data);
      saveCache(contacts);

      console.log("[Autocomplete] Data reloaded");

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
      display: "none"
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
        dropdown.choose();
        e.preventDefault();
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

    ["ToCustomer", "CcCustomer", "BccCustomer"].forEach(id => {
      attachAutocomplete(document.getElementById(id));
    });
  }

  init();

})();
