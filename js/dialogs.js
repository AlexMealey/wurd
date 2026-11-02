// Modal dialogs (native <dialog>: the rest of the page is inert while one is open, and Esc closes it):
// Export As, the password prompt, the keyboard shortcuts list and Settings
(function (Wurd) {
  'use strict';

  const $ = (selector) => document.querySelector(selector);
  const MIN_PASSWORD = 8;

  // Opens a dialog and puts focus back where it was when the dialog closes. Dialogs that edit
  // the document pass restoreFocus: false; the app puts the caret back itself.
  const show = (dialog, focusTarget, { restoreFocus = true } = {}) => {
    const opener = document.activeElement;
    if (restoreFocus) dialog.addEventListener('close', () => opener?.focus?.({ preventScroll: true }), { once: true });
    dialog.showModal();
    focusTarget?.focus();
  };

  const setError = (input, errorEl, message) => {
    errorEl.textContent = message || '';
    errorEl.hidden = !message;
    if (message) input.setAttribute('aria-invalid', 'true');
    else input.removeAttribute('aria-invalid');
  };

  const resetPasswordToggle = (button) => {
    const input = document.getElementById(button.getAttribute('aria-controls'));
    input.type = 'password';
    button.setAttribute('aria-pressed', 'false');
  };

  document.querySelectorAll('.password__toggle').forEach((button) => {
    button.addEventListener('click', () => {
      const input = document.getElementById(button.getAttribute('aria-controls'));
      const reveal = input.type === 'password';
      input.type = reveal ? 'text' : 'password';
      button.setAttribute('aria-pressed', String(reveal));
    });
  });

  // Runs one dialog session: resolves with whatever finish() is given (null when cancelled)
  const session = (dialog, { onSubmit, cancelButton }) => new Promise((resolve) => {
    const form = dialog.querySelector('form');
    const finish = (value) => {
      form.removeEventListener('submit', submit);
      dialog.removeEventListener('cancel', cancel);
      cancelButton?.removeEventListener('click', cancelClick);
      if (dialog.open) dialog.close();
      resolve(value);
    };
    const submit = (e) => {
      e.preventDefault();
      onSubmit(finish);
    };
    const cancel = (e) => {
      e.preventDefault();
      finish(null);
    };
    const cancelClick = () => finish(null);
    form.addEventListener('submit', submit);
    dialog.addEventListener('cancel', cancel);
    cancelButton?.addEventListener('click', cancelClick);
  });

  // Asks for a document's password. Resolves with the password, or null if cancelled.
  const askPassword = ({ fileName, error } = {}) => {
    const dialog = $('#password-dialog');
    const input = $('#password-input');
    const errorEl = $('#password-error');
    $('#password-message').textContent = `“${fileName}” is protected with a password.`;
    input.value = '';
    resetPasswordToggle(dialog.querySelector('.password__toggle'));
    setError(input, errorEl, error);
    const result = session(dialog, {
      cancelButton: $('#password-cancel'),
      onSubmit: (finish) => {
        if (!input.value) {
          setError(input, errorEl, 'Enter the password.');
          input.focus();
          return;
        }
        const password = input.value;
        input.value = '';
        finish(password);
      },
    });
    show(dialog, input);
    return result;
  };

  // The Export As dialog. Resolves with { title, format, password, destination } or null if cancelled.
  // destination ('download' or 'server') is offered only when one is given, i.e. a server is connected.
  const exportAs = ({ title, format, destination = null }) => {
    const dialog = $('#export-dialog');
    const name = $('#export-name');
    const nameError = $('#export-name-error');
    const hint = $('#export-format-hint');
    const protect = $('#export-protect');
    const protectField = $('#export-protect-field');
    const passwordFields = $('#export-password-fields');
    const password = $('#export-password');
    const passwordError = $('#export-password-error');
    const confirm = $('#export-confirm');
    const confirmError = $('#export-confirm-error');
    const cryptoNote = $('#export-crypto-note');
    const radios = Array.from(dialog.querySelectorAll('input[name="export-format"]'));
    const destinationField = $('#export-destination');
    const destinations = Array.from(dialog.querySelectorAll('input[name="export-destination"]'));
    const cryptoOk = Wurd.cryptoAvailable();
    const formats = Object.fromEntries(Wurd.formats.EXPORTS.map((f) => [f.id, f]));

    const selected = () => radios.find((r) => r.checked)?.value;
    const needsPassword = () => {
      const f = formats[selected()];
      return Boolean(f && (f.encrypted || (f.encryptable && protect.checked)));
    };

    const update = () => {
      const f = formats[selected()];
      hint.textContent = f ? f.hint : '';
      protectField.hidden = !f?.encryptable;
      passwordFields.hidden = !needsPassword();
      cryptoNote.hidden = cryptoOk;
    };

    name.value = title;
    protect.checked = false;
    protect.disabled = !cryptoOk;
    radios.forEach((r) => {
      r.disabled = r.value === 'wurd' && !cryptoOk;
      r.checked = r.value === (formats[format] && !(format === 'wurd' && !cryptoOk) ? format : 'docx');
    });
    destinationField.hidden = !destination;
    destinations.forEach((r) => {
      r.checked = r.value === (destination || 'download');
    });
    [password, confirm].forEach((input) => {
      input.value = '';
    });
    dialog.querySelectorAll('.password__toggle').forEach(resetPasswordToggle);
    setError(name, nameError, '');
    setError(password, passwordError, '');
    setError(confirm, confirmError, '');
    update();

    const onChange = () => update();
    dialog.addEventListener('change', onChange);

    const result = session(dialog, {
      cancelButton: $('#export-cancel'),
      onSubmit: (finish) => {
        const fileTitle = name.value.trim();
        setError(name, nameError, fileTitle ? '' : 'Enter a file name.');
        let invalid = fileTitle ? null : name;
        if (needsPassword()) {
          const tooShort = password.value.length < MIN_PASSWORD;
          setError(password, passwordError, tooShort ? `Use at least ${MIN_PASSWORD} characters.` : '');
          const mismatch = !tooShort && confirm.value !== password.value;
          setError(confirm, confirmError, mismatch ? "The passwords don't match." : '');
          invalid = invalid || (tooShort ? password : mismatch ? confirm : null);
        } else {
          setError(password, passwordError, '');
          setError(confirm, confirmError, '');
        }
        if (invalid) {
          invalid.focus();
          return;
        }
        const value = {
          title: fileTitle,
          format: selected(),
          password: needsPassword() ? password.value : null,
          destination: destination ? destinations.find((r) => r.checked).value : 'download',
        };
        password.value = '';
        confirm.value = '';
        finish(value);
      },
    });
    result.then(() => dialog.removeEventListener('change', onChange));
    show(dialog, name);
    return result;
  };

  // The keyboard shortcuts list, grouped, for the current platform
  const showShortcuts = () => {
    const dialog = $('#shortcuts-dialog');
    const body = $('#shortcuts-body');
    body.replaceChildren();
    let group = null;
    const rows = [...Wurd.shortcuts.list(), { id: 'escape', label: 'Close the find bar or a dialog', group: 'View', labels: ['Esc'] }];
    for (const shortcut of rows) {
      if (shortcut.group !== group) {
        group = shortcut.group;
        const headRow = document.createElement('tr');
        const head = document.createElement('th');
        head.scope = 'rowgroup';
        head.colSpan = 2;
        head.className = 'shortcuts__group';
        head.textContent = group;
        headRow.append(head);
        body.append(headRow);
      }
      const row = document.createElement('tr');
      const name = document.createElement('th');
      name.scope = 'row';
      name.textContent = shortcut.label;
      const keys = document.createElement('td');
      shortcut.labels.forEach((text, i) => {
        if (i > 0) keys.append(' or ');
        const kbd = document.createElement('kbd');
        kbd.textContent = text;
        keys.append(kbd);
      });
      row.append(name, keys);
      body.append(row);
    }
    const close = $('#shortcuts-close');
    const done = () => dialog.close();
    close.addEventListener('click', done, { once: true });
    show(dialog, close);
  };

  // Settings take effect as they change; onChange runs after a draft setting changes,
  // onPageSize after the page size does
  const showSettings = ({ onChange, onPageSize }) => {
    const dialog = $('#settings-dialog');
    const lifetime = $('#settings-draft');
    const status = $('#settings-draft-status');
    const pageSize = $('#settings-page-size');
    const change = () => {
      Wurd.drafts.setMinutes(Number(lifetime.value));
      onChange();
      status.textContent = Wurd.drafts.describe();
    };
    const changeSize = () => {
      Wurd.setPageSize(pageSize.value);
      onPageSize?.(Wurd.pageSize());
    };
    lifetime.value = String(Wurd.drafts.minutes());
    pageSize.value = Wurd.pageSize();
    status.textContent = Wurd.drafts.describe();
    lifetime.addEventListener('change', change);
    pageSize.addEventListener('change', changeSize);
    dialog.addEventListener('close', () => {
      lifetime.removeEventListener('change', change);
      pageSize.removeEventListener('change', changeSize);
    }, { once: true });
    show(dialog, pageSize);
  };

  // Insert, change or remove a link. Resolves with { href, text }, { remove: true }, or null if cancelled.
  // text is offered when inserting at a caret or editing a link; editing also offers Remove.
  const editLink = ({ href = '', text = '', showText = true, editing = false } = {}) => {
    const dialog = $('#link-dialog');
    const title = $('#link-title');
    const textField = $('#link-text-field');
    const textInput = $('#link-text');
    const url = $('#link-url');
    const help = $('#link-url-help');
    const error = $('#link-url-error');
    const remove = $('#link-remove');
    const submit = $('#link-submit');
    title.textContent = editing ? 'Edit link' : 'Insert link';
    submit.textContent = editing ? 'Save' : 'Insert';
    textField.hidden = !showText;
    textInput.value = text;
    url.value = href.replace(/^mailto:/i, '');
    remove.hidden = !editing;
    help.textContent = `A web or email address. ${Wurd.shortcuts.isMac() ? '⌘' : 'Ctrl'}+click a link in the document to open it.`;
    setError(url, error, '');
    let removed = false;
    const onRemove = () => {
      removed = true;
      dialog.querySelector('form').requestSubmit();
    };
    remove.addEventListener('click', onRemove);
    const result = session(dialog, {
      cancelButton: $('#link-cancel'),
      onSubmit: (finish) => {
        if (removed) {
          finish({ remove: true });
          return;
        }
        const target = Wurd.insert.normalizeUrl(url.value);
        if (!target) {
          setError(url, error, url.value.trim() ? 'Use a web address (like example.com) or an email address.' : 'Enter an address.');
          url.focus();
          return;
        }
        finish({ href: target, text: showText ? textInput.value.trim() : '' });
      },
    });
    result.then(() => remove.removeEventListener('click', onRemove));
    show(dialog, url, { restoreFocus: false });
    return result;
  };

  // The special characters most recently inserted, newest first (kept on this device)
  const RECENT_KEY = 'wurd-recent-symbols';
  const MAX_RECENT = 8;
  const recentSymbols = () => {
    try {
      const saved = JSON.parse(localStorage.getItem(RECENT_KEY));
      return Array.isArray(saved) ? saved.filter((c) => typeof c === 'string' && c.length <= 4).slice(0, MAX_RECENT) : [];
    } catch (e) {
      return [];
    }
  };
  const rememberSymbol = (ch) => {
    try {
      localStorage.setItem(RECENT_KEY, JSON.stringify([ch, ...recentSymbols().filter((c) => c !== ch)].slice(0, MAX_RECENT)));
    } catch (e) {
      // Storage unavailable: nothing is remembered
    }
  };

  const codeOf = (ch) => `U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`;

  // The special characters dialog. Resolves with the character picked, or null.
  const pickSymbol = () => {
    const dialog = $('#symbols-dialog');
    const body = $('#symbols-body');
    const names = new Map(Wurd.insert.SYMBOLS.flatMap((g) => g.chars));
    const groups = [...Wurd.insert.SYMBOLS];
    const recent = recentSymbols().filter((c) => names.has(c));
    if (recent.length) groups.unshift({ group: 'Recently used', chars: recent.map((c) => [c, names.get(c)]) });
    let picked = null;
    body.replaceChildren(...groups.map((g, i) => {
      const section = document.createElement('section');
      section.className = 'symbols__group';
      const heading = document.createElement('h3');
      heading.className = 'symbols__heading';
      heading.id = `symbols-group-${i}`;
      heading.textContent = g.group;
      const grid = document.createElement('div');
      grid.className = 'symbols__grid';
      grid.setAttribute('role', 'group');
      grid.setAttribute('aria-labelledby', heading.id);
      grid.append(...g.chars.map(([ch, name]) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'symbol';
        button.textContent = ch;
        button.setAttribute('aria-label', name);
        button.title = `${name} (${codeOf(ch)})`;
        button.addEventListener('click', () => {
          picked = ch;
          dialog.querySelector('form').requestSubmit();
        });
        return button;
      }));
      section.append(heading, grid);
      return section;
    }));
    const result = session(dialog, {
      cancelButton: $('#symbols-cancel'),
      onSubmit: (finish) => {
        if (picked) rememberSymbol(picked);
        finish(picked);
      },
    });
    show(dialog, body.querySelector('.symbol'), { restoreFocus: false });
    return result;
  };

  Wurd.dialogs = {
    askPassword,
    exportAs,
    showShortcuts,
    showSettings,
    editLink,
    pickSymbol,
    isOpen: () => Boolean(document.querySelector('dialog[open]')),
  };
})(window.Wurd = window.Wurd || {});
