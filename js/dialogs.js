// Modal dialogs (native <dialog>: the rest of the page is inert while one is open, and Esc closes it):
// Export As, the password prompt and the keyboard shortcuts list
(function (Wurd) {
  'use strict';

  const $ = (selector) => document.querySelector(selector);
  const MIN_PASSWORD = 8;

  // Opens a dialog and puts focus back where it was when the dialog closes
  const show = (dialog, focusTarget) => {
    const opener = document.activeElement;
    dialog.addEventListener('close', () => opener?.focus?.({ preventScroll: true }), { once: true });
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

  // The Export As dialog. Resolves with { fileName, format, password } or null if cancelled.
  const exportAs = ({ title, format }) => {
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
        const value = { title: fileTitle, format: selected(), password: needsPassword() ? password.value : null };
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

  Wurd.dialogs = {
    askPassword,
    exportAs,
    showShortcuts,
    isOpen: () => Boolean(document.querySelector('dialog[open]')),
  };
})(window.Wurd = window.Wurd || {});
