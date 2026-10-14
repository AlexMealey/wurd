// Keyboard shortcuts: one table drives the key handling, tooltips, aria-keyshortcuts and the help dialog.
// "Mod" is Ctrl on Windows/Linux and Cmd on macOS. Keys are matched by physical key (event.code)
// so they work on any keyboard layout.
(function (Wurd) {
  'use strict';

  const detectMac = () => /Mac|iPhone|iPad|iPod/i.test(navigator.userAgentData?.platform || navigator.platform || '');
  let mac = detectMac();

  // scope 'editor': document commands, ignored while typing in a text field (the title, find box, source view).
  // scope 'app': work everywhere outside dialogs.
  const SHORTCUTS = [
    { id: 'undo', label: 'Undo', group: 'Editing', scope: 'editor', keys: ['Mod+KeyZ'] },
    { id: 'redo', label: 'Redo', group: 'Editing', scope: 'editor', keys: { pc: ['Mod+KeyY', 'Mod+Shift+KeyZ'], mac: ['Mod+Shift+KeyZ', 'Mod+KeyY'] } },
    { id: 'bold', label: 'Bold', group: 'Text', scope: 'editor', keys: ['Mod+KeyB'] },
    { id: 'italic', label: 'Italic', group: 'Text', scope: 'editor', keys: ['Mod+KeyI'] },
    { id: 'underline', label: 'Underline', group: 'Text', scope: 'editor', keys: ['Mod+KeyU'] },
    { id: 'strikeThrough', label: 'Strikethrough', group: 'Text', scope: 'editor', keys: { pc: ['Alt+Shift+Digit5'], mac: ['Mod+Shift+KeyX'] } },
    { id: 'p', label: 'Normal text', group: 'Paragraph', scope: 'editor', keys: ['Mod+Alt+Digit0'] },
    { id: 'h1', label: 'Heading 1', group: 'Paragraph', scope: 'editor', keys: ['Mod+Alt+Digit1'] },
    { id: 'h2', label: 'Heading 2', group: 'Paragraph', scope: 'editor', keys: ['Mod+Alt+Digit2'] },
    { id: 'h3', label: 'Heading 3', group: 'Paragraph', scope: 'editor', keys: ['Mod+Alt+Digit3'] },
    { id: 'justifyLeft', label: 'Align left', group: 'Paragraph', scope: 'editor', keys: ['Mod+Shift+KeyL'] },
    { id: 'justifyCenter', label: 'Center', group: 'Paragraph', scope: 'editor', keys: ['Mod+Shift+KeyE'] },
    { id: 'justifyRight', label: 'Align right', group: 'Paragraph', scope: 'editor', keys: ['Mod+Shift+KeyR'] },
    { id: 'justifyFull', label: 'Justify', group: 'Paragraph', scope: 'editor', keys: ['Mod+Shift+KeyJ'] },
    { id: 'insertOrderedList', label: 'Numbered list', group: 'Paragraph', scope: 'editor', keys: ['Mod+Shift+Digit7'] },
    { id: 'insertUnorderedList', label: 'Bulleted list', group: 'Paragraph', scope: 'editor', keys: ['Mod+Shift+Digit8'] },
    // Browsers reserve Ctrl+N for a new window, so New uses Ctrl+Alt+N
    { id: 'new', label: 'New document', group: 'File', scope: 'app', keys: ['Mod+Alt+KeyN'] },
    { id: 'open', label: 'Open', group: 'File', scope: 'app', keys: ['Mod+KeyO'] },
    { id: 'save', label: 'Save', group: 'File', scope: 'app', keys: ['Mod+KeyS'] },
    { id: 'exportAs', label: 'Export As…', group: 'File', scope: 'app', keys: ['Mod+Shift+KeyS'] },
    { id: 'find', label: 'Find & replace', group: 'View', scope: 'app', keys: { pc: ['Mod+KeyF', 'Mod+KeyH'], mac: ['Mod+KeyF', 'Mod+Shift+KeyH'] } },
    { id: 'source', label: 'Toggle source view', group: 'View', scope: 'app', keys: ['Mod+Alt+KeyU'] },
    { id: 'help', label: 'Keyboard shortcuts', group: 'View', scope: 'app', keys: ['Mod+Slash'] },
    { id: 'settings', label: 'Settings', group: 'View', scope: 'app', keys: ['Mod+Comma'] },
  ];

  const keysFor = (shortcut) => (Array.isArray(shortcut.keys) ? shortcut.keys : shortcut.keys[mac ? 'mac' : 'pc']);

  const parse = (spec) => {
    const parts = spec.split('+');
    return {
      code: parts.pop(),
      mod: parts.includes('Mod'),
      alt: parts.includes('Alt'),
      shift: parts.includes('Shift'),
    };
  };

  const KEY_NAMES = { Slash: '/', Comma: ',', Escape: 'Esc' };
  const keyName = (code) => KEY_NAMES[code] || code.replace(/^Key|^Digit/, '');

  // Human label: "Ctrl+Shift+Z" on Windows/Linux, "⇧⌘Z" on macOS
  const label = (spec) => {
    const k = parse(spec);
    if (mac) return `${k.alt ? '⌥' : ''}${k.shift ? '⇧' : ''}${k.mod ? '⌘' : ''}${keyName(k.code)}`;
    return [k.mod && 'Ctrl', k.alt && 'Alt', k.shift && 'Shift', keyName(k.code)].filter(Boolean).join('+');
  };

  // aria-keyshortcuts value, e.g. "Control+Shift+Z"
  const ariaLabel = (spec) => {
    const k = parse(spec);
    return [k.mod && (mac ? 'Meta' : 'Control'), k.alt && 'Alt', k.shift && 'Shift', keyName(k.code)].filter(Boolean).join('+');
  };

  const matches = (spec, e) => {
    const k = parse(spec);
    const codeMatches = e.code === k.code || (k.code === 'Slash' && e.key === '/');
    if (!codeMatches) return false;
    const modDown = mac ? e.metaKey : e.ctrlKey;
    const otherMod = mac ? e.ctrlKey : e.metaKey;
    return modDown === k.mod && !otherMod && e.altKey === k.alt && e.shiftKey === k.shift;
  };

  const isTextField = (el) => Boolean(el && (el.matches?.('input, textarea, select') || (el.isContentEditable && !el.closest('#editor'))));

  let handlers = {};

  const onKeyDown = (e) => {
    if (e.defaultPrevented || e.isComposing) return;
    if (document.querySelector('dialog[open]')) return;
    // AltGr arrives as Ctrl+Alt on Windows; it types characters, so never treat it as a shortcut
    if (e.getModifierState?.('AltGraph')) return;
    const shortcut = SHORTCUTS.find((s) => keysFor(s).some((spec) => matches(spec, e)));
    if (!shortcut || !handlers[shortcut.id]) return;
    if (shortcut.scope === 'editor' && isTextField(e.target)) return;
    e.preventDefault();
    // Holding a key down shouldn't open several Save or Open prompts
    if (e.repeat && shortcut.scope === 'app') return;
    handlers[shortcut.id](e);
  };

  // Adds the shortcut to the tooltip and aria-keyshortcuts of every [data-shortcut] control
  const decorate = (root = document) => {
    root.querySelectorAll('[data-shortcut]').forEach((el) => {
      const shortcut = SHORTCUTS.find((s) => s.id === el.dataset.shortcut);
      if (!shortcut) return;
      const specs = keysFor(shortcut);
      const base = el.dataset.title || el.getAttribute('title') || shortcut.label;
      el.dataset.title = base;
      el.setAttribute('title', `${base} (${label(specs[0])})`);
      el.setAttribute('aria-keyshortcuts', specs.map(ariaLabel).join(' '));
    });
  };

  Wurd.shortcuts = {
    list: () => SHORTCUTS.map((s) => ({ ...s, labels: keysFor(s).map(label) })),
    labelFor: (id) => {
      const shortcut = SHORTCUTS.find((s) => s.id === id);
      return shortcut ? label(keysFor(shortcut)[0]) : '';
    },
    isMac: () => mac,
    // For testing the macOS bindings on another platform
    setPlatform: (platform) => {
      mac = platform === 'mac';
      decorate();
    },
    init: (actions) => {
      handlers = actions;
      document.addEventListener('keydown', onKeyDown);
      decorate();
    },
    decorate,
  };
})(window.Wurd = window.Wurd || {});
