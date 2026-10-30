// ==UserScript==
// @name        Phrasenschwein
// @description Upload selected text into a NextCloud directory (public upload-only share)
// @namespace   github.com/olifre/userstyles
// @icon        https://olifre.github.io/favicon.ico
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/phrasenschwein.user.js
// @downloadURL https://raw.githubusercontent.com/olifre/userscripts/main/phrasenschwein.user.js
// @version     1.0.2
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

    const NATIVE = {
        attachShadow: Element.prototype.attachShadow,
        matchMedia: window.matchMedia
    };

    const DEFAULTS = {
        server: 'https://uni-bonn.sciebo.de',
        token: 'AbCdEf123',
        password: 'Beispiel-Passwort',
        hotkey: 'Alt+Shift+P'
    };

    const PREFIX = 'phrasenschwein.';
    const HOST_ID = 'phrasenschwein-userscript';

    const getConfig = key =>
        GM_getValue(PREFIX + key, DEFAULTS[key]);

    const setConfig = (key, value) =>
        GM_setValue(PREFIX + key, value);

    const on = (target, type, listener, options) =>
        target?.addEventListener?.(type, listener, options);

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
     * Liest Auswahlen aus normalen Dokumenten, Eingabefeldern
     * und offenen Shadow DOMs. Das Userscript läuft außerdem
     * auch direkt in Frames.
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

                if (!next) {
                    break;
                }

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

            let selection = null;

            if (typeof document.getSelection === 'function') {
                selection = document.getSelection();
            } else if (typeof window.getSelection === 'function') {
                selection = window.getSelection();
            }

            return selection ? String(selection) : '';
        } catch (error) {
            console.warn(
                'Phrasenschwein: Auswahl konnte nicht gelesen werden:',
                error
            );

            return '';
        }
    }

    function browserPrefersDarkMode() {
        try {
            return typeof NATIVE.matchMedia === 'function' &&
                NATIVE.matchMedia.call(
                    window,
                    '(prefers-color-scheme: dark)'
                ).matches;
        } catch {
            return false;
        }
    }

    function hotkeyMatches(event, definition) {
        const parts = String(definition || '')
            .toLowerCase()
            .split('+')
            .map(part => part.trim())
            .filter(Boolean);

        const has = name => parts.includes(name);

        const modifiers = new Set([
            'alt',
            'shift',
            'ctrl',
            'control',
            'meta',
            'cmd',
            'command'
        ]);

        const wantedKey = parts.find(
            part => !modifiers.has(part)
        );

        if (
            !wantedKey ||
            event.altKey !== has('alt') ||
            event.shiftKey !== has('shift') ||
            event.ctrlKey !== (
                has('ctrl') || has('control')
            ) ||
            event.metaKey !== (
                has('meta') ||
                has('cmd') ||
                has('command')
            )
        ) {
            return false;
        }

        const key = String(event.key || '').toLowerCase();
        const code = String(event.code || '').toLowerCase();

        return (
            key === wantedKey ||
            code === wantedKey ||
            code === `key${wantedKey}` ||
            code === `digit${wantedKey}`
        );
    }

    function appendHost(host, callback) {
        const target = document.body || document.documentElement;

        if (target) {
            target.appendChild(host);
            callback();
            return;
        }

        on(document, 'DOMContentLoaded', () => {
            const delayedTarget =
                document.body || document.documentElement;

            if (delayedTarget && !host.isConnected) {
                delayedTarget.appendChild(host);
                callback();
            }
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
         * Beim Öffnen wird immer die Browserpräferenz verwendet.
         * Eine manuelle Änderung gilt nur für diese Dialoginstanz.
         */
        let darkMode = browserPrefersDarkMode();

        host.dataset.theme = darkMode ? 'dark' : 'light';

        appendHost(host, initializeDialog);

        function initializeDialog() {
            let root;

            try {
                root = NATIVE.attachShadow.call(
                    host,
                    { mode: 'open' }
                );
            } catch (error) {
                host.remove();

                window.alert(
                    'Phrasenschwein konnte nicht geöffnet werden: ' +
                    (error?.message || String(error))
                );

                return;
            }

            root.innerHTML = `
                <style>
                    :host {
                        --bg: #fff;
                        --fg: #202124;
                        --field: #fff;
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
                        gap: 5px;
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

                    button,
                    input {
                        font: inherit;
                    }

                    .icon-button {
                        min-width: 32px;
                        height: 32px;
                        padding: 0 6px;
                        color: var(--fg);
                        background: transparent;
                        border: 0;
                        border-radius: 5px;
                        cursor: pointer;
                    }

                    .icon-button:hover {
                        background: rgb(127 127 127 / 18%);
                    }

                    .settings {
                        font-size: 18px;
                        line-height: 1;
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

                    .main-panel[hidden],
                    .settings-panel[hidden] {
                        display: none;
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

                    textarea:focus,
                    input:focus {
                        border-color: var(--primary);
                        box-shadow: 0 0 0 2px var(--focus);
                    }

                    .send,
                    .save-settings,
                    .cancel-settings {
                        margin-top: 10px;
                        padding: 8px 14px;
                        border: 0;
                        border-radius: 5px;
                        font-weight: 600;
                        cursor: pointer;
                    }

                    .send,
                    .save-settings {
                        color: #fff;
                        background: var(--primary);
                    }

                    .send:hover,
                    .save-settings:hover {
                        background: var(--primary-hover);
                    }

                    .send:disabled {
                        opacity: .6;
                        cursor: wait;
                    }

                    .status,
                    .settings-message {
                        min-height: 18px;
                        margin-top: 8px;
                        color: var(--fg);
                        font-size: 13px;
                        line-height: 1.35;
                        word-break: break-word;
                    }

                    .error {
                        color: var(--danger);
                    }

                    .success {
                        color: var(--success);
                    }

                    .settings-panel label {
                        display: block;
                        margin-bottom: 10px;
                        color: var(--fg);
                        font-size: 13px;
                        font-weight: 600;
                    }

                    .settings-panel input {
                        display: block;
                        width: 100%;
                        margin-top: 4px;
                        padding: 8px 9px;
                        color: var(--fg);
                        background: var(--field);
                        border: 1px solid var(--border);
                        border-radius: 5px;
                        outline: none;
                    }

                    .settings-actions {
                        display: flex;
                        gap: 8px;
                    }

                    .cancel-settings {
                        color: var(--fg);
                        background: rgb(127 127 127 / 20%);
                    }

                    .cancel-settings:hover {
                        background: rgb(127 127 127 / 30%);
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
                            class="icon-button settings"
                            type="button"
                            title="Einstellungen"
                            aria-label="Einstellungen"
                        >⚙</button>

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

                    <div class="main-panel">
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
                    </div>

                    <div class="settings-panel" hidden>
                        <label>
                            Serveradresse
                            <input
                                class="config-server"
                                type="url"
                                placeholder="https://uni-bonn.sciebo.de"
                            >
                        </label>

                        <label>
                            Share-ID
                            <input
                                class="config-token"
                                type="text"
                                autocomplete="off"
                                spellcheck="false"
                            >
                        </label>

                        <label>
                            Passwort
                            <input
                                class="config-password"
                                type="password"
                                autocomplete="off"
                            >
                        </label>

                        <label>
                            Tastenkombination
                            <input
                                class="config-hotkey"
                                type="text"
                                autocomplete="off"
                                spellcheck="false"
                                placeholder="Alt+Shift+P"
                            >
                        </label>

                        <div class="settings-actions">
                            <button
                                class="save-settings"
                                type="button"
                            >
                                Speichern
                            </button>

                            <button
                                class="cancel-settings"
                                type="button"
                            >
                                Abbrechen
                            </button>
                        </div>

                        <div
                            class="settings-message"
                            aria-live="polite"
                        ></div>
                    </div>
                </section>
            `;

            const $ = selector => root.querySelector(selector);

            const header = $('.header');
            const textarea = $('textarea');
            const sendButton = $('.send');
            const status = $('.status');
            const themeButton = $('.theme');
            const closeButton = $('.close');

            /*
             * Neue Elemente für die eingebetteten Einstellungen.
             */
            const mainPanel = $('.main-panel');
            const settingsPanel = $('.settings-panel');
            const settingsButton = $('.settings');
            const serverInput = $('.config-server');
            const tokenInput = $('.config-token');
            const passwordInput = $('.config-password');
            const hotkeyInput = $('.config-hotkey');
            const saveSettingsButton = $('.save-settings');
            const cancelSettingsButton = $('.cancel-settings');
            const settingsMessage = $('.settings-message');

            textarea.value = selectedText;

            /* =================================================
             * Theme
             * ================================================= */

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

            on(themeButton, 'click', event => {
                event.stopPropagation();

                darkMode = !darkMode;
                host.dataset.theme =
                    darkMode ? 'dark' : 'light';

                updateThemeButton();
                textarea.focus();
            });

            /* =================================================
             * Einstellungen im Dialog
             * ================================================= */

            function openSettings() {
                serverInput.value = String(
                    getConfig('server') ?? ''
                );

                tokenInput.value = String(
                    getConfig('token') ?? ''
                );

                passwordInput.value = String(
                    getConfig('password') ?? ''
                );

                hotkeyInput.value = String(
                    getConfig('hotkey') ?? DEFAULTS.hotkey
                );

                settingsMessage.className =
                    'settings-message';

                settingsMessage.textContent = '';
                mainPanel.hidden = true;
                settingsPanel.hidden = false;
                serverInput.focus();
            }

            function closeSettings() {
                settingsPanel.hidden = true;
                mainPanel.hidden = false;
                textarea.focus();
            }

            on(settingsButton, 'click', event => {
                event.stopPropagation();

                if (settingsPanel.hidden) {
                    openSettings();
                } else {
                    closeSettings();
                }
            });

            on(cancelSettingsButton, 'click', closeSettings);

            on(saveSettingsButton, 'click', () => {
                const server = serverInput.value
                    .trim()
                    .replace(/\/+$/, '');

                const token = tokenInput.value.trim();
                const password = passwordInput.value;
                const hotkey = hotkeyInput.value.trim();

                if (!server || !token || !hotkey) {
                    settingsMessage.className =
                        'settings-message error';

                    settingsMessage.textContent =
                        'Serveradresse, Share-ID und ' +
                        'Tastenkombination sind erforderlich.';

                    return;
                }

                try {
                    const parsed = new URL(server);

                    if (
                        parsed.protocol !== 'https:' &&
                        parsed.protocol !== 'http:'
                    ) {
                        throw new Error('Ungültiges Protokoll');
                    }
                } catch {
                    settingsMessage.className =
                        'settings-message error';

                    settingsMessage.textContent =
                        'Bitte eine gültige Serveradresse eingeben.';

                    return;
                }

                setConfig('server', server);
                setConfig('token', token);
                setConfig('password', password);
                setConfig('hotkey', hotkey);

                settingsMessage.className =
                    'settings-message success';

                settingsMessage.textContent =
                    'Einstellungen gespeichert.';

                window.setTimeout(closeSettings, 400);
            });

            for (const input of [
                serverInput,
                tokenInput,
                passwordInput,
                hotkeyInput
            ]) {
                on(input, 'keydown', event => {
                    if (event.key === 'Enter') {
                        event.preventDefault();
                        saveSettingsButton.click();
                    }
                });
            }

            /* =================================================
             * Schließen
             * ================================================= */

            on(closeButton, 'click', () => {
                host.remove();
            });

            /* =================================================
             * Box verschieben
             * ================================================= */

            let drag = null;

            on(header, 'pointerdown', event => {
                let clickedButton = false;

                try {
                    clickedButton = Boolean(
                        event.target?.closest?.('button')
                    );
                } catch {
                    clickedButton = false;
                }

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

                try {
                    header.setPointerCapture?.(event.pointerId);
                } catch {
                    // Pointer Capture ist optional.
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
                        drag.left +
                        event.clientX -
                        drag.startX
                    )
                );

                const top = Math.max(
                    0,
                    Math.min(
                        maxTop,
                        drag.top +
                        event.clientY -
                        drag.startY
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

            /* =================================================
             * Upload
             * ================================================= */

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

                let authorization;

                try {
                    authorization = createBasicAuth(
                        String(getConfig('token')).trim(),
                        String(getConfig('password'))
                    );
                } catch (error) {
                    sendButton.disabled = false;
                    status.className = 'status error';
                    status.textContent =
                        'Zugangsdaten konnten nicht verarbeitet ' +
                        'werden: ' +
                        (error?.message || String(error));

                    return;
                }

                GM_xmlhttpRequest({
                    method: 'PUT',
                    url: getUploadUrl(filename),
                    timeout: 30000,

                    headers: {
                        Authorization: `Basic ${authorization}`,
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
                            status.className =
                                'status success';

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
                            'Upload fehlgeschlagen: ' +
                            'Zeitüberschreitung.';
                    }
                });
            });

            /*
             * Strg/Cmd + Enter überträgt den Text.
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
             * Escape schließt zuerst die Einstellungen,
             * ansonsten den vollständigen Dialog.
             */
            on(root, 'keydown', event => {
                if (event.key !== 'Escape') {
                    return;
                }

                if (!settingsPanel.hidden) {
                    closeSettings();
                } else {
                    host.remove();
                }
            });

            window.setTimeout(() => {
                textarea.focus();
            }, 0);
        }
    }

    /* =========================================================
     * Vorhandenes Userscript-Menü
     * ========================================================= */

    function promptConfig(key, label, secret = false) {
        const current = String(getConfig(key) ?? '');

        const value = window.prompt(
            secret
                ? `${label} eingeben.\n\n` +
                  'Leer lassen, um den bisherigen Wert beizubehalten.'
                : `${label} eingeben:`,
            secret ? '' : current
        );

        if (
            value === null ||
            (secret && value === '')
        ) {
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

    /*
     * Menübefehle nur einmal im Hauptdokument registrieren.
     * Das eigentliche Script läuft trotzdem auch in Frames.
     */
    if (window.top === window) {
        GM_registerMenuCommand(
            '🐷 Phrasenschwein öffnen',
            openPhrasenschwein
        );

        GM_registerMenuCommand(
            '⚙ Serveradresse einstellen',
            () => promptConfig(
                'server',
                'Serveradresse ohne WebDAV-Pfad'
            )
        );

        GM_registerMenuCommand(
            '⚙ Share-ID einstellen',
            () => promptConfig(
                'token',
                'Share-ID beziehungsweise Share-Token'
            )
        );

        GM_registerMenuCommand(
            '⚙ Passwort einstellen',
            () => promptConfig(
                'password',
                'Passwort',
                true
            )
        );

        GM_registerMenuCommand(
            '⌨ Tastenkombination einstellen',
            () => promptConfig(
                'hotkey',
                'Tastenkombination, z. B. Alt+Shift+P'
            )
        );

        GM_registerMenuCommand(
            '↺ Einstellungen zurücksetzen',
            () => {
                if (!window.confirm(
                    'Alle Phrasenschwein-Einstellungen zurücksetzen?'
                )) {
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
    }

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
