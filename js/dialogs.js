// Modal dialogs (native <dialog>: the rest of the page is inert while one is open, and Esc closes it):
// Export As, the password prompt, the keyboard shortcuts list and Settings
(function (Wurd) {
  'use strict';

  const $ = (selector) => document.querySelector(selector);

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

  const showPassword = (button, shown) => {
    const input = document.getElementById(button.getAttribute('aria-controls'));
    input.type = shown ? 'text' : 'password';
    button.setAttribute('aria-pressed', String(shown));
  };

  const resetPasswordToggle = (button) => showPassword(button, false);

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
    const strengthNote = $('#export-password-strength');
    const suggestButton = $('#export-suggest');
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
    strengthNote.hidden = true;
    update();

    const onChange = () => update();
    dialog.addEventListener('change', onChange);

    // Why a password isn't accepted, or '' if it is. The file name counts as something an attacker would try.
    const MIN_BITS = Wurd.passwords.MIN_BITS;
    const passwordProblem = async () => {
      if (!password.value) return 'Enter a password.';
      try {
        const { bits, warning } = await Wurd.passwords.strength(password.value, [name.value]);
        if (bits >= MIN_BITS) return '';
        return `Too easy to guess (about ${bits} bits of strength; it needs ${MIN_BITS}).${warning ? ` ${warning}.` : ''} Add more unrelated words, or use Suggest a password.`;
      } catch (err) {
        // Without the checker, saving waits rather than accepting a password nobody checked
        return err instanceof Wurd.UserError ? err.message : "Couldn't check the password's strength.";
      }
    };

    // A live estimate while typing; the check that counts runs again on submit
    let estimate = 0;
    const showStrength = async () => {
      const run = ++estimate;
      if (!password.value) {
        strengthNote.hidden = true;
        return;
      }
      let text;
      try {
        const { bits } = await Wurd.passwords.strength(password.value, [name.value]);
        text = bits >= MIN_BITS ? `Strong enough: about ${bits} bits.` : `Too easy to guess: about ${bits} bits, needs ${MIN_BITS}.`;
      } catch (err) {
        text = err instanceof Wurd.UserError ? err.message : '';
      }
      if (run !== estimate) return;
      strengthNote.textContent = text;
      strengthNote.hidden = !text;
    };
    password.addEventListener('input', showStrength);

    const suggest = async () => {
      suggestButton.disabled = true;
      try {
        const value = await Wurd.passwords.suggest([name.value]);
        password.value = value;
        confirm.value = value;
        // Shown, so it can be copied into a password manager before saving
        dialog.querySelectorAll('.password__toggle').forEach((button) => showPassword(button, true));
        setError(password, passwordError, '');
        setError(confirm, confirmError, '');
        await showStrength();
        strengthNote.textContent += ' Copy it somewhere safe, like a password manager, before you save.';
      } catch (err) {
        setError(password, passwordError, err instanceof Wurd.UserError ? err.message : "Couldn't suggest a password.");
      } finally {
        suggestButton.disabled = false;
      }
    };
    suggestButton.addEventListener('click', suggest);

    let checking = false;
    const result = session(dialog, {
      cancelButton: $('#export-cancel'),
      onSubmit: async (finish) => {
        if (checking) return;
        const fileTitle = name.value.trim();
        setError(name, nameError, fileTitle ? '' : 'Enter a file name.');
        let invalid = fileTitle ? null : name;
        if (needsPassword()) {
          const checked = password.value;
          checking = true;
          const problem = await passwordProblem();
          checking = false;
          // Cancelled, or the password changed, while it was being checked
          if (!dialog.open || password.value !== checked) return;
          setError(password, passwordError, problem);
          const mismatch = !problem && confirm.value !== password.value;
          setError(confirm, confirmError, mismatch ? "The passwords don't match." : '');
          invalid = invalid || (problem ? password : mismatch ? confirm : null);
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
    result.then(() => {
      dialog.removeEventListener('change', onChange);
      password.removeEventListener('input', showStrength);
      suggestButton.removeEventListener('click', suggest);
      // An estimate still running is dropped, and none is left for the next time the dialog opens
      estimate += 1;
      strengthNote.textContent = '';
    });
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

  // Settings take effect as they change; onChange runs after each change
  const showSettings = ({ onChange }) => {
    const dialog = $('#settings-dialog');
    const lifetime = $('#settings-draft');
    const status = $('#settings-draft-status');
    const change = () => {
      Wurd.drafts.setMinutes(Number(lifetime.value));
      onChange();
      status.textContent = Wurd.drafts.describe();
    };
    lifetime.value = String(Wurd.drafts.minutes());
    status.textContent = Wurd.drafts.describe();
    lifetime.addEventListener('change', change);
    dialog.addEventListener('close', () => lifetime.removeEventListener('change', change), { once: true });
    show(dialog, lifetime);
  };

  Wurd.dialogs = {
    askPassword,
    exportAs,
    showShortcuts,
    showSettings,
    isOpen: () => Boolean(document.querySelector('dialog[open]')),
  };
})(window.Wurd = window.Wurd || {});
