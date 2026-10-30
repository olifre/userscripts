// ==UserScript==
// @name        Phrasenschwein
// @description Upload selected text into a NextCloud directory (public upload-only share)
// @namespace   github.com/olifre/userstyles
// @icon        https://olifre.github.io/favicon.ico
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/phrasenschwein.user.js
// @downloadURL https://raw.githubusercontent.com/olifre/userscripts/main/phrasenschwein.user.js
// @version     1.0.0
// @match       *://*/*
// @run-at      document-idle
// @grant       GM_xmlhttpRequest
// @grant       GM_registerMenuCommand
// @grant       GM_getValue
// @grant       GM_setValue
// @connect     *
// @author      Oliver Freyermuth <o.freyermuth@googlemail.com> (https://olifre.github.io/)
// @license     Unlicense
// ==/UserScript==

(() => {
    'use strict';

    /*
     * Kritische Browserfunktionen früh sichern. Für Event-Listener wird
     * absichtlich die normale Methode verwendet, da ein Aufruf über
     * EventTarget.prototype in Userscript-Sandboxes Probleme machen kann.
     */
    const NATIVE = {
        attachShadow: Element.prototype.attachShadow,
        getSelection: window.getSelection,
        matchMedia: window.matchMedia
    };

    function on(target, type, listener, options) {
        target.addEventListener(type, listener, options);
    }

    /* =========================================================
     * Grundeinstellungen
     * ========================================================= */

    const DEFAULTS = {
        server: 'https://uni-bonn.sciebo.de',
        token: 'AbCdEf123',
        password: 'Beispiel-Passwort',
        hotkey: 'Alt+Shift+P'
    };

    const STORAGE_PREFIX = 'phrasenschwein.';
    const HOST_ID = 'phrasenschwein-userscript';

    function getConfig(key) {
        return GM_getValue(
            STORAGE_PREFIX + key,
            DEFAULTS[key]
        );
    }

    function setConfig(key, value) {
        GM_setValue(STORAGE_PREFIX + key, value);
    }

    /* =========================================================
     * Konfigurationsmenü
     * ========================================================= */

    function promptConfig(key, label, { secret = false } = {}) {
        const current = String(getConfig(key) ?? '');

        const message = secret
            ? `${label} eingeben.\n\n` +
              'Leer lassen, um den bisherigen Wert beizubehalten.'
            : `${label} eingeben:`;

        const value = window.prompt(
            message,
            secret ? '' : current
        );

        if (value === null || (secret && value === '')) {
            return;
        }

        const normalized = key === 'password'
            ? value
            : value.trim();

        if (!normalized && key !== 'password') {
            window.alert(`${label} darf nicht leer sein.`);
            return;
        }

        setConfig(key, normalized);
        window.alert(`${label} wurde gespeichert.`);
    }

    /* =========================================================
     * Hilfsfunktionen
     * ========================================================= */

    function createBasicAuth(username, password) {
        const bytes = new TextEncoder().encode(
            `${username}:${password}`
        );

        let binary = '';

        for (const byte of bytes) {
            binary += String.fromCharCode(byte);
        }

        return btoa(binary);
    }

    /*
     * Erzeugt zum Beispiel:
     *
     * https://uni-bonn.sciebo.de/public.php/dav/files/SHAREID/DATEI.txt
     */
    function getUploadUrl(filename) {
        const server = String(getConfig('server'))
            .trim()
            .replace(/\/+$/, '');

        const shareId = encodeURIComponent(
            String(getConfig('token')).trim()
        );

        return (
            `${server}/public.php/dav/files/` +
            `${shareId}/${encodeURIComponent(filename)}`
        );
    }

    /*
     * Ermittelt markierten Text aus:
     *
     * 1. INPUT- und TEXTAREA-Elementen
     * 2. offenen Shadow DOMs
     * 3. normaler Textauswahl im Dokument
     */
    function getSelectedText() {
        try {
            let element = document.activeElement;

            while (element) {
                let next = null;

                try {
                    next = element.shadowRoot?.activeElement || null;
                } catch {
                    next = null;
                }

                if (!next) break;
                element = next;
            }

            if (
                element &&
                typeof element.value === 'string' &&
                typeof element.selectionStart === 'number' &&
                typeof element.selectionEnd === 'number' &&
                element.selectionEnd > element.selectionStart
            ) {
                return element.value.substring(
                    element.selectionStart,
                    element.selectionEnd
                );
            }

            const selection =
                typeof document.getSelection === 'function'
                    ? document.getSelection()
                    : typeof window.getSelection === 'function'
                        ? window.getSelection()
                        : null;

            return selection ? `${selection}` : '';
        } catch (error) {
            console.warn(
                'Phrasenschwein: Auswahl konnte nicht gelesen werden:',
                error
            );
            return '';
        }
    }

    function browserPrefersDarkMode() {
        if (typeof NATIVE.matchMedia !== 'function') {
            return false;
        }

        try {
            return Boolean(
                NATIVE.matchMedia.call(
                    window,
                    '(prefers-color-scheme: dark)'
                ).matches
            );
        } catch {
            return false;
        }
    }

    /*
     * Erkennt beispielsweise:
     *
     * Alt+Shift+P
     * Ctrl+Alt+P
     * Ctrl+Shift+F8
     * Meta+Shift+P
     */
    function hotkeyMatches(event, definition) {
        if (typeof definition !== 'string') {
            return false;
        }

        const parts = definition
            .toLowerCase()
            .split('+')
            .map(part => part.trim())
            .filter(Boolean);

        const has = name => parts.includes(name);

        const modifierNames = new Set([
            'alt',
            'shift',
            'ctrl',
            'control',
            'meta',
            'cmd',
            'command'
        ]);

        const wantedKey = parts.find(
            part => !modifierNames.has(part)
        );

        if (!wantedKey) {
            return false;
        }

        const modifiersMatch =
            event.altKey === has('alt') &&
            event.shiftKey === has('shift') &&
            event.ctrlKey === (
                has('ctrl') || has('control')
            ) &&
            event.metaKey === (
                has('meta') ||
                has('cmd') ||
                has('command')
            );

        if (!modifiersMatch) {
            return false;
        }

        const eventKey = String(event.key || '').toLowerCase();
        const eventCode = String(event.code || '').toLowerCase();

        /*
         * event.key berücksichtigt das Tastaturlayout.
         * event.code erkennt zusätzlich die physische Buchstabentaste.
         */
        return (
            eventKey === wantedKey ||
            eventCode === wantedKey ||
            eventCode === `key${wantedKey}` ||
            eventCode === `digit${wantedKey}`
        );
    }

    function appendHost(host) {
        const target =
            document.body ||
            document.documentElement;

        if (target) {
            target.appendChild(host);
            return;
        }

        /*
         * Falls das Tastenkürzel extrem früh ausgelöst wird, bevor ein
         * Dokumentelement vorhanden ist.
         */
        on(document, 'DOMContentLoaded', () => {
            const delayedTarget =
                document.body ||
                document.documentElement;

            delayedTarget?.appendChild(host);
        }, { once: true });
    }

    /* =========================================================
     * Floating Dialog
     * ========================================================= */

    function openPhrasenschwein() {
        const existing = document.getElementById(HOST_ID);

        if (existing) {
            existing.shadowRoot
                ?.querySelector('textarea')
                ?.focus();

            return;
        }

        if (typeof NATIVE.attachShadow !== 'function') {
            window.alert(
                'Phrasenschwein konnte nicht geöffnet werden: ' +
                'Shadow DOM ist nicht verfügbar.'
            );
            return;
        }

        const selectedText = getSelectedText();
        const host = document.createElement('div');

        host.id = HOST_ID;

        Object.assign(host.style, {
            position: 'fixed',
            top: '20px',
            right: '20px',
            left: 'auto',
            zIndex: '2147483647'
        });

        /*
         * Beim erneuten Öffnen wird immer wieder die Browserpräferenz
         * verwendet. Die manuelle Umschaltung gilt nur für diese Box.
         */
        let darkMode = browserPrefersDarkMode();

        host.setAttribute(
            'data-theme',
            darkMode ? 'dark' : 'light'
        );

        let root;

        try {
            root = NATIVE.attachShadow.call(
                host,
                { mode: 'open' }
            );
        } catch (error) {
            window.alert(
                'Phrasenschwein konnte nicht geöffnet werden: ' +
                (error?.message || String(error))
            );
            return;
        }

        root.innerHTML = `
            <style>
                :host {
                    --bg: #ffffff;
                    --fg: #202124;
                    --field: #ffffff;
                    --border: #a8adb3;
                    --header: #f5f6f7;
                    --primary: #0675b8;
                    --primary-hover: #055f95;
                    --danger: #c62828;
                    --danger-hover: #f7dddd;
                    --success: #087830;
                    --focus: rgb(6 117 184 / 25%);

                    display: block;
                    width: min(370px, calc(100vw - 40px));
                    color-scheme: light;
                    font: 14px system-ui, -apple-system,
                          BlinkMacSystemFont, "Segoe UI", sans-serif;
                }

                :host([data-theme="dark"]) {
                    --bg: #202124;
                    --fg: #f1f3f4;
                    --field: #292a2d;
                    --border: #686d73;
                    --header: #292a2d;
                    --primary: #168dcc;
                    --primary-hover: #35a4df;
                    --danger: #ff7770;
                    --danger-hover: #543438;
                    --success: #72d993;
                    --focus: rgb(22 141 204 / 30%);

                    color-scheme: dark;
                }

                * {
                    box-sizing: border-box;
                }

                .box {
                    overflow: hidden;
                    padding: 0 16px 16px;
                    color: var(--fg);
                    background: var(--bg);
                    border: 1px solid var(--border);
                    border-radius: 10px;
                    box-shadow: 0 5px 25px rgb(0 0 0 / 45%);
                }

                .header {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    margin: 0 -16px 12px;
                    padding: 9px 9px 8px 16px;
                    background: var(--header);
                    border-bottom: 1px solid var(--border);
                    cursor: move;
                    user-select: none;
                    touch-action: none;
                }

                h2 {
                    flex: 1;
                    margin: 0;
                    color: var(--fg);
                    font-size: 20px;
                    line-height: 1.3;
                }

                button {
                    font: inherit;
                }

                .icon-button {
                    min-width: 32px;
                    height: 32px;
                    padding: 0 7px;
                    color: var(--fg);
                    background: transparent;
                    border: 0;
                    border-radius: 5px;
                    cursor: pointer;
                }

                .icon-button:hover {
                    background: rgb(127 127 127 / 18%);
                }

                .theme {
                    font-size: 18px;
                    line-height: 1;
                }

                .close {
                    color: var(--danger);
                    font-size: 27px;
                    line-height: 1;
                }

                .close:hover {
                    background: var(--danger-hover);
                }

                textarea {
                    display: block;
                    width: 100%;
                    height: 150px;
                    padding: 9px;
                    resize: vertical;
                    color: var(--fg);
                    background: var(--field);
                    border: 1px solid var(--border);
                    border-radius: 5px;
                    outline: none;
                    font: inherit;
                    line-height: 1.4;
                }

                textarea:focus {
                    border-color: var(--primary);
                    box-shadow: 0 0 0 2px var(--focus);
                }

                .send {
                    margin-top: 10px;
                    padding: 8px 14px;
                    color: #ffffff;
                    background: var(--primary);
                    border: 0;
                    border-radius: 5px;
                    font-weight: 600;
                    cursor: pointer;
                }

                .send:hover {
                    background: var(--primary-hover);
                }

                .send:disabled {
                    opacity: .6;
                    cursor: wait;
                }

                .status {
                    min-height: 18px;
                    margin-top: 8px;
                    color: var(--fg);
                    font-size: 13px;
                    line-height: 1.35;
                    word-break: break-word;
                }

                .status.error {
                    color: var(--danger);
                }

                .status.success {
                    color: var(--success);
                }
            </style>

            <section
                class="box"
                role="dialog"
                aria-modal="false"
                aria-labelledby="phrasenschwein-title"
            >
                <header class="header">
                    <h2 id="phrasenschwein-title">
                        Phrasenschwein
                    </h2>

                    <button
                        class="icon-button theme"
                        type="button"
                        aria-label="Farbschema wechseln"
                    ></button>

                    <button
                        class="icon-button close"
                        type="button"
                        title="Schließen"
                        aria-label="Schließen"
                    >×</button>
                </header>

                <textarea
                    aria-label="Zu übertragender Text"
                    placeholder="Text eingeben …"
                ></textarea>

                <button class="send" type="button">
                    Übertragen
                </button>

                <div
                    class="status"
                    aria-live="polite"
                ></div>
            </section>
        `;

        appendHost(host);

        const $ = selector => root.querySelector(selector);

        const header = $('.header');
        const textarea = $('textarea');
        const sendButton = $('.send');
        const themeButton = $('.theme');
        const closeButton = $('.close');
        const status = $('.status');

        textarea.value = selectedText;

        /* -----------------------------------------------------
         * Theme
         * ----------------------------------------------------- */

        function updateThemeButton() {
            themeButton.textContent = darkMode ? '☀' : '☾';

            themeButton.title = darkMode
                ? 'Helles Farbschema verwenden'
                : 'Dunkles Farbschema verwenden';

            themeButton.setAttribute(
                'aria-pressed',
                String(darkMode)
            );
        }

        updateThemeButton();

        on(themeButton, 'click', () => {
            darkMode = !darkMode;

            host.setAttribute(
                'data-theme',
                darkMode ? 'dark' : 'light'
            );

            updateThemeButton();
            textarea.focus();
        });

        on(closeButton, 'click', () => {
            host.remove();
        });

        /* -----------------------------------------------------
         * Box verschieben
         * ----------------------------------------------------- */

        let drag = null;

        on(header, 'pointerdown', event => {
            const clickedButton =
                event.target instanceof Element &&
                event.target.closest('button');

            if (event.button !== 0 || clickedButton) {
                return;
            }

            const rect = host.getBoundingClientRect();

            drag = {
                startX: event.clientX,
                startY: event.clientY,
                left: rect.left,
                top: rect.top
            };

            host.style.right = 'auto';
            host.style.left = `${rect.left}px`;
            host.style.top = `${rect.top}px`;

            if (typeof header.setPointerCapture === 'function') {
                header.setPointerCapture(event.pointerId);
            }

            event.preventDefault();
        });

        on(header, 'pointermove', event => {
            if (!drag) {
                return;
            }

            const maxLeft = Math.max(
                0,
                window.innerWidth - host.offsetWidth
            );

            const maxTop = Math.max(
                0,
                window.innerHeight - host.offsetHeight
            );

            const left = Math.max(
                0,
                Math.min(
                    maxLeft,
                    drag.left + event.clientX - drag.startX
                )
            );

            const top = Math.max(
                0,
                Math.min(
                    maxTop,
                    drag.top + event.clientY - drag.startY
                )
            );

            host.style.left = `${left}px`;
            host.style.top = `${top}px`;
        });

        const stopDragging = () => {
            drag = null;
        };

        on(header, 'pointerup', stopDragging);
        on(header, 'pointercancel', stopDragging);
        on(header, 'lostpointercapture', stopDragging);

        /* -----------------------------------------------------
         * Upload
         * ----------------------------------------------------- */

        on(sendButton, 'click', () => {
            const text = textarea.value;

            if (!text.trim()) {
                status.className = 'status error';
                status.textContent =
                    'Bitte zuerst Text eingeben.';
                return;
            }

            const timestamp = Math.floor(Date.now() / 1000);
            const randomPart = Math.random()
                .toString(36)
                .slice(2, 7);

            const filename =
                `${timestamp}-${randomPart}.txt`;

            sendButton.disabled = true;
            status.className = 'status';
            status.textContent = 'Wird hochgeladen …';

            GM_xmlhttpRequest({
                method: 'PUT',
                url: getUploadUrl(filename),
                timeout: 30000,

                headers: {
                    Authorization: `Basic ${createBasicAuth(
                        String(getConfig('token')).trim(),
                        String(getConfig('password'))
                    )}`,
                    'Content-Type':
                        'text/plain; charset=UTF-8'
                },

                data: text,

                onload(response) {
                    sendButton.disabled = false;

                    if (
                        response.status >= 200 &&
                        response.status < 300
                    ) {
                        status.className = 'status success';
                        status.textContent =
                            `Erfolgreich übertragen: ${filename}`;

                        textarea.value = '';
                        return;
                    }

                    const detail = String(
                        response.responseText || ''
                    )
                        .replace(/<[^>]+>/g, ' ')
                        .replace(/\s+/g, ' ')
                        .trim()
                        .slice(0, 250);

                    status.className = 'status error';
                    status.textContent =
                        `Upload fehlgeschlagen: HTTP ` +
                        `${response.status}` +
                        (response.statusText
                            ? ` ${response.statusText}`
                            : '') +
                        (detail ? ` – ${detail}` : '');
                },

                onerror() {
                    sendButton.disabled = false;
                    status.className = 'status error';
                    status.textContent =
                        'Upload fehlgeschlagen: Netzwerkfehler.';
                },

                ontimeout() {
                    sendButton.disabled = false;
                    status.className = 'status error';
                    status.textContent =
                        'Upload fehlgeschlagen: Zeitüberschreitung.';
                }
            });
        });

        /*
         * Strg/Cmd + Enter überträgt den Inhalt.
         */
        on(textarea, 'keydown', event => {
            if (
                event.key === 'Enter' &&
                (event.ctrlKey || event.metaKey)
            ) {
                event.preventDefault();
                sendButton.click();
            }
        });

        /*
         * Escape schließt den Dialog.
         */
        on(root, 'keydown', event => {
            if (event.key === 'Escape') {
                host.remove();
            }
        });

        /*
         * focus() erst nach dem Einfügen in das Dokument ausführen.
         */
        window.requestAnimationFrame(() => {
            textarea.focus();
        });
    }

    /* =========================================================
     * Userscript-Menü
     * ========================================================= */

    GM_registerMenuCommand(
        '🐷 Phrasenschwein öffnen',
        openPhrasenschwein
    );

    GM_registerMenuCommand(
        '⚙ Serveradresse einstellen',
        () => {
            promptConfig(
                'server',
                'Serveradresse ohne WebDAV-Pfad'
            );
        }
    );

    GM_registerMenuCommand(
        '⚙ Share-ID einstellen',
        () => {
            promptConfig(
                'token',
                'Share-ID beziehungsweise Share-Token'
            );
        }
    );

    GM_registerMenuCommand(
        '⚙ Passwort einstellen',
        () => {
            promptConfig(
                'password',
                'Passwort',
                { secret: true }
            );
        }
    );

    GM_registerMenuCommand(
        '⌨ Tastenkombination einstellen',
        () => {
            promptConfig(
                'hotkey',
                'Tastenkombination, z. B. ' +
                'Alt+Shift+P oder Ctrl+Alt+P'
            );
        }
    );

    GM_registerMenuCommand(
        '↺ Einstellungen zurücksetzen',
        () => {
            const confirmed = window.confirm(
                'Alle Phrasenschwein-Einstellungen zurücksetzen?'
            );

            if (!confirmed) {
                return;
            }

            for (const [key, value] of Object.entries(DEFAULTS)) {
                setConfig(key, value);
            }

            window.alert(
                'Die Einstellungen wurden zurückgesetzt.'
            );
        }
    );

    /* =========================================================
     * Globale Tastenkombination
     * ========================================================= */

    window.addEventListener('keydown', event => {
        const configuredHotkey = String(
            getConfig('hotkey') || DEFAULTS.hotkey
        );

        if (
            event.repeat ||
            !hotkeyMatches(event, configuredHotkey)
        ) {
            return;
        }

        event.preventDefault();
        event.stopImmediatePropagation();

        openPhrasenschwein();
    }, { capture: true });
})();
