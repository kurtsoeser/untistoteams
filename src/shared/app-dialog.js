/**
 * Zentrierte App-Dialoge (ersetzt native alert / confirm / prompt).
 * Lädt app.css (modal-overlay / modal-box) mit.
 * @file
 */
(function () {
    'use strict';

    let root = null;
    let titleEl;
    let msgEl;
    let promptWrap;
    let inputLabelTextEl;
    let inputEl;
    let okBtn;
    let cancelBtn;
    /** @type {'alert'|'confirm'|'prompt'} */
    let mode = 'alert';
    let resolver = null;

    function cacheRefs() {
        titleEl = root.querySelector('[data-app-dialog-title]');
        msgEl = root.querySelector('[data-app-dialog-message]');
        promptWrap = root.querySelector('[data-app-dialog-prompt]');
        inputLabelTextEl = root.querySelector('[data-app-dialog-input-label-text]');
        inputEl = root.querySelector('[data-app-dialog-input]');
        okBtn = root.querySelector('[data-app-dialog-ok]');
        cancelBtn = root.querySelector('[data-app-dialog-cancel]');
    }

    function teardown(result) {
        if (!resolver) return;
        document.removeEventListener('keydown', onDocKey, true);
        if (root) root.classList.remove('open');
        const fn = resolver;
        resolver = null;
        fn(result);
    }

    function onDocKey(ev) {
        if (!root || !root.classList.contains('open')) return;
        if (ev.key === 'Escape') {
            ev.preventDefault();
            ev.stopPropagation();
            if (mode === 'confirm') teardown(false);
            else if (mode === 'prompt') teardown(null);
            else teardown();
        } else if (ev.key === 'Enter' && mode === 'prompt' && !ev.shiftKey) {
            ev.preventDefault();
            teardown(inputEl.value);
        } else if (ev.key === 'Enter' && mode === 'confirm' && !ev.shiftKey) {
            ev.preventDefault();
            teardown(true);
        }
    }

    function ensure() {
        if (root) return root;
        root = document.createElement('div');
        root.id = 'ms365AppDialog';
        root.className = 'modal-overlay ms365-app-dialog-overlay';
        root.setAttribute('role', 'dialog');
        root.setAttribute('aria-modal', 'true');
        root.innerHTML =
            '<div class="modal-box ms365-app-dialog-box" tabindex="-1">' +
            '<h3 class="app-dialog-title" data-app-dialog-title></h3>' +
            '<p class="app-dialog-message" data-app-dialog-message></p>' +
            '<div class="app-dialog-prompt" data-app-dialog-prompt style="display:none">' +
            '<label class="app-dialog-input-label">' +
            '<span data-app-dialog-input-label-text></span>' +
            '<input type="text" class="app-dialog-input" data-app-dialog-input autocomplete="off" spellcheck="false" />' +
            '</label></div>' +
            '<div class="modal-actions">' +
            '<button type="button" class="btn" data-app-dialog-cancel>Abbrechen</button>' +
            '<button type="button" class="btn btn-success" data-app-dialog-ok>OK</button>' +
            '</div></div>';
        document.body.appendChild(root);
        cacheRefs();
        root.addEventListener('click', function (e) {
            if (e.target !== root) return;
            if (mode === 'confirm') teardown(false);
            else if (mode === 'prompt') teardown(null);
            else teardown();
        });
        cancelBtn.addEventListener('click', function () {
            if (mode === 'confirm') teardown(false);
            else if (mode === 'prompt') teardown(null);
            else teardown();
        });
        okBtn.addEventListener('click', function () {
            if (mode === 'prompt') teardown(inputEl.value);
            else if (mode === 'confirm') teardown(true);
            else teardown();
        });
        return root;
    }

    /**
     * @param {string} message
     * @param {{ title?: string, okText?: string }} [options]
     * @returns {Promise<void>}
     */
    function ms365AppDialogAlert(message, options) {
        ensure();
        mode = 'alert';
        titleEl.textContent = (options && options.title) || 'Hinweis';
        msgEl.textContent = String(message ?? '');
        promptWrap.style.display = 'none';
        cancelBtn.style.display = 'none';
        okBtn.textContent = (options && options.okText) || 'OK';
        okBtn.className = 'btn btn-success';
        return new Promise(function (resolve) {
            resolver = function () {
                resolve();
            };
            root.classList.add('open');
            document.addEventListener('keydown', onDocKey, true);
            requestAnimationFrame(function () {
                okBtn.focus();
            });
        });
    }

    /**
     * @param {string} message
     * @param {{ title?: string, okText?: string, cancelText?: string, danger?: boolean }} [options]
     * @returns {Promise<boolean>}
     */
    function ms365AppDialogConfirm(message, options) {
        ensure();
        mode = 'confirm';
        titleEl.textContent = (options && options.title) || 'Bestätigung';
        msgEl.textContent = String(message ?? '');
        promptWrap.style.display = 'none';
        cancelBtn.style.display = '';
        cancelBtn.textContent = (options && options.cancelText) || 'Abbrechen';
        const danger = options && options.danger;
        okBtn.textContent = (options && options.okText) || 'OK';
        okBtn.className = danger ? 'btn btn-danger' : 'btn btn-success';
        return new Promise(function (resolve) {
            resolver = function (v) {
                resolve(!!v);
            };
            root.classList.add('open');
            document.addEventListener('keydown', onDocKey, true);
            requestAnimationFrame(function () {
                if (danger) cancelBtn.focus();
                else okBtn.focus();
            });
        });
    }

    /**
     * @param {string} message
     * @param {string} [defaultValue]
     * @param {{ title?: string, inputLabel?: string, okText?: string, cancelText?: string }} [options]
     * @returns {Promise<string|null>} null bei Abbrechen
     */
    function ms365AppDialogPrompt(message, defaultValue, options) {
        ensure();
        mode = 'prompt';
        titleEl.textContent = (options && options.title) || 'Eingabe';
        msgEl.textContent = String(message ?? '');
        if (inputLabelTextEl) {
            inputLabelTextEl.textContent = (options && options.inputLabel) || 'Eingabe';
        }
        inputEl.value = defaultValue != null ? String(defaultValue) : '';
        promptWrap.style.display = 'block';
        cancelBtn.style.display = '';
        cancelBtn.textContent = (options && options.cancelText) || 'Abbrechen';
        okBtn.textContent = (options && options.okText) || 'OK';
        okBtn.className = 'btn btn-success';
        return new Promise(function (resolve) {
            resolver = function (v) {
                if (v === null || v === undefined) resolve(null);
                else resolve(String(v));
            };
            root.classList.add('open');
            document.addEventListener('keydown', onDocKey, true);
            requestAnimationFrame(function () {
                inputEl.focus();
                try {
                    inputEl.select();
                } catch {
                    // ignore
                }
            });
        });
    }

    window.ms365AppDialogAlert = ms365AppDialogAlert;
    window.ms365AppDialogConfirm = ms365AppDialogConfirm;
    window.ms365AppDialogPrompt = ms365AppDialogPrompt;

    /**
     * Kurzinfo: zuerst Toast (falls vorhanden), sonst modal „Hinweis“.
     * @param {string} msg
     * @param {{ title?: string }} [opts]
     */
    function ms365ToastOrAlert(msg, opts) {
        if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
            return;
        }
        void ms365AppDialogAlert(msg, Object.assign({ title: 'Hinweis' }, opts || {}));
    }
    window.ms365ToastOrAlert = ms365ToastOrAlert;
})();
