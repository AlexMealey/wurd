// Wurd editor behaviour. Markup lives in index.html, styles in css/app.css, file formats in js/formats/.
(function (Wurd) {
  'use strict';

  const $ = (selector) => document.querySelector(selector);

  const root = document.documentElement;
  const editor = $('#editor');
  const source = $('#source');
  const page = $('#page');
  const docTitle = $('#doc-title');
  const docFormat = $('#doc-format');
  const fileInput = $('#file-input');
  const toast = $('#toast');
  const toastMessage = $('#toast-message');
  const announcer = $('#announcer');
  const outline = $('#outline');
  const outlineEmpty = $('#outline-empty');
  const wordCount = $('#word-count');
  const charCount = $('#char-count');
  const readTime = $('#read-time');
  const ribbon = $('#ribbon');
  const undoButton = $('#undo');
  const redoButton = $('#redo');
  const blockSelect = $('#block-format');
  const fontSelect = $('#font-family');
  const findbar = $('#findbar');
  const findToggle = $('#find-toggle');
  const findInput = $('#find-input');
  const replaceInput = $('#replace-input');
  const sourceToggle = $('#source-toggle');
  const sourceToggleLabel = $('#source-toggle-label');
  const zoomLevel = $('#zoom-level');
  const themeToggle = $('#theme-toggle');
  const saveButton = $('#save-doc');

  const BLANK_DOCUMENT = '<h1>Untitled Document</h1><p>Start typing your document here...</p>';
  const SAVE_FORMATS = ['docx', 'odt', 'doc', 'html', 'txt', 'wurd'];
  const ICON_ATTRS = { 'aria-hidden': 'true', focusable: 'false' };

  let isSourceMode = false;
  let zoom = 100;
  let toastTimer = null;
  // The open document's lock (see Wurd.formats.lock), kept in memory only so Save can encrypt it again.
  // The password itself is never kept.
  let documentLock = null;
  // The Wurd server copy the open document belongs to ({ id, rev }), or null. Save writes back to it,
  // or, while connected, creates one in the vault.
  let serverFile = null;
  // The document as last opened or saved; null when unknown, so it counts as changed
  let cleanHtml = editor.innerHTML;
  let draftTimer = 0;
  // The sample the page opens with needs no draft: a reload shows it anyway
  const SAMPLE = { html: editor.innerHTML, title: docTitle.value };

  // ---- Notifications ------------------------------------------------------

  const notify = (msg) => {
    toastMessage.textContent = msg;
    toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      toast.hidden = true;
    }, 3000);
    // Clearing first makes screen readers announce a repeated message again
    announcer.textContent = '';
    setTimeout(() => {
      announcer.textContent = msg;
    }, 50);
  };

  // ---- Document state -----------------------------------------------------

  const setFormat = (format) => {
    docFormat.dataset.format = format;
    docFormat.textContent = `.${format}`;
  };

  // The page-wide font, kept in step with the font menu
  const setPageFont = (font) => {
    fontSelect.value = font;
    page.style.fontFamily = font;
  };

  const updateWindowTitle = () => {
    document.title = `${docTitle.value.trim() || 'Untitled'} – Wurd`;
  };

  // Moves the caret to the start of an element in the editor and scrolls it into view
  const placeCaretAt = (el) => {
    editor.focus({ preventScroll: true });
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(true);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    el.scrollIntoView({ block: 'start', behavior: reduceMotion ? 'auto' : 'smooth' });
  };

  const updateStatsAndOutline = () => {
    // The editor is hidden in source mode; stats and outline refresh when switching back
    if (isSourceMode) return;
    const text = editor.innerText || '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    wordCount.textContent = `${words} words`;
    charCount.textContent = `${text.length} chars`;
    readTime.textContent = `~${Math.max(1, Math.ceil(words / 200))} min read`;

    // Extract headings; each outline entry jumps to its heading
    const headings = Array.from(editor.querySelectorAll('h1, h2, h3'));
    outline.replaceChildren(...headings.map((el, i) => {
      const item = document.createElement('li');
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `outline__item outline__item--${el.tagName.toLowerCase()}`;
      button.textContent = el.innerText || `Untitled Heading ${i + 1}`;
      button.addEventListener('click', () => placeCaretAt(el));
      item.append(button);
      return item;
    }));
    outline.hidden = headings.length === 0;
    outlineEmpty.hidden = headings.length > 0;
  };

  // Pressed states for the formatting buttons and the current paragraph style, from the selection
  const updateToolStates = () => {
    if (isSourceMode) return;
    const selection = window.getSelection();
    if (!selection.rangeCount || !editor.contains(selection.anchorNode)) return;
    document.querySelectorAll('[data-command]').forEach((button) => {
      let state = false;
      try {
        state = document.queryCommandState(button.dataset.command);
      } catch (e) {
        // Unsupported command state: leave it unpressed
      }
      button.setAttribute('aria-pressed', String(state));
    });
    const block = (document.queryCommandValue('formatBlock') || '').toLowerCase();
    blockSelect.value = ['h1', 'h2', 'h3'].includes(block) ? block : 'p';
  };

  const history = Wurd.createHistory({
    editor,
    limit: 100,
    onChange: () => {
      undoButton.disabled = !history.canUndo();
      redoButton.disabled = !history.canRedo();
      ensureToolbarStop();
      updateStatsAndOutline();
      updateToolStates();
      scheduleDraft();
    },
  });

  // Puts a document into the editor (or the source view). Everything loaded is sanitized here.
  const loadDocument = ({ html, title, format, font, lock = null }) => {
    const clean = Wurd.sanitizeHtml(html);
    editor.innerHTML = clean;
    if (isSourceMode) source.value = format === 'html' ? html : clean;
    docTitle.value = title;
    updateWindowTitle();
    setFormat(format);
    setPageFont(font || Wurd.DEFAULT_FONT);
    documentLock = lock;
    cleanHtml = editor.innerHTML;
    history.reset();
  };

  // Chrome's list commands leave the new list inside the old paragraph (<p><ol>...</ol></p>), which isn't valid HTML
  const unwrapListParagraphs = () => {
    editor.querySelectorAll('p').forEach((p) => {
      const only = p.children.length === 1 && p.firstElementChild;
      if (only && /^(UL|OL)$/.test(only.tagName) && !p.textContent.replace(only.textContent, '').trim()) {
        // Ranges are live and would collapse when the list moves, so keep plain node/offset pairs
        const selection = window.getSelection();
        const r = selection.rangeCount ? selection.getRangeAt(0) : null;
        const saved = r && [r.startContainer, r.startOffset, r.endContainer, r.endOffset];
        p.replaceWith(only);
        if (saved && only.contains(saved[0]) && only.contains(saved[2])) {
          const range = document.createRange();
          range.setStart(saved[0], saved[1]);
          range.setEnd(saved[2], saved[3]);
          selection.removeAllRanges();
          selection.addRange(range);
        }
      }
    });
  };

  const exec = (command, value = null) => {
    if (isSourceMode) return;
    history.batch(command, () => {
      document.execCommand(command, false, value);
      if (command === 'insertOrderedList' || command === 'insertUnorderedList') unwrapListParagraphs();
    });
    editor.focus();
    updateStatsAndOutline();
    updateToolStates();
  };

  // ---- Files --------------------------------------------------------------

  const updateSaveTitle = () => {
    saveButton.title = Wurd.server.connected() ? 'Save to your vault' : 'Save Document';
  };

  const setServerFile = (value) => {
    serverFile = value;
    updateSaveTitle();
    scheduleDraft();
  };

  // Opens a File. fromServer is { id, rev } when it came from the Wurd server.
  const openFile = async (file, fromServer = null) => {
    try {
      const doc = await Wurd.formats.open(file, {
        askPassword: (options) => Wurd.dialogs.askPassword({ fileName: file.name, ...options }),
      });
      loadDocument(doc);
      setServerFile(fromServer);
      notify(doc.notice || `Loaded "${file.name}"${fromServer ? ' from the server' : ''}`);
    } catch (err) {
      if (err instanceof Wurd.Cancelled) return;
      if (!(err instanceof Wurd.UserError)) console.error(err);
      notify(err instanceof Wurd.UserError ? err.message : `Couldn't open "${file.name}". It may be damaged or in an unsupported format.`);
    }
  };

  const documentHtml = () => (isSourceMode ? Wurd.sanitizeHtml(source.value) : editor.innerHTML);

  // to: 'download', 'server' (the open document's server copy) or 'new-on-server'. A new password (from
  // Export As) makes a new lock; Save passes the document's existing one.
  const writeFile = async ({ title, format, password = null, lock = null, to = 'download' }) => {
    saveButton.disabled = true;
    try {
      const html = documentHtml();
      const fileLock = password ? await Wurd.formats.lock(format, password) : lock;
      const blob = await Wurd.formats.export(format, html, { title, pageFont: fontSelect.value, lock: fileLock });
      const fileName = Wurd.formats.fileName(title, format);
      const withPassword = fileLock ? ' with a password' : '';
      if (to === 'download') {
        Wurd.downloadBlob(blob, fileName);
        cleanHtml = html;
        notify(`Saved "${fileName}"${withPassword}`);
        return;
      }
      const saved = to === 'server'
        ? await Wurd.server.update(serverFile.id, serverFile.rev, fileName, blob)
        : await Wurd.server.create(fileName, blob);
      // The document is now that server file, so Save keeps its name, format and lock
      setServerFile({ id: saved.id, rev: saved.rev });
      docTitle.value = title;
      updateWindowTitle();
      setFormat(format);
      documentLock = fileLock;
      cleanHtml = html;
      refreshServerList();
      notify(`Saved "${fileName}" to your vault${withPassword}`);
    } catch (err) {
      if (err instanceof Wurd.server.ServerError && err.status === 412) {
        notify('This document was changed somewhere else since you opened it. To keep your version, use Export As to save it to the server as a new document.');
        return;
      }
      if (err instanceof Wurd.server.ServerError && err.status === 404 && to === 'server') {
        setServerFile(null);
        notify('This document is no longer in your vault. Save again to keep it there as a new document.');
        return;
      }
      if (!(err instanceof Wurd.UserError)) console.error(err);
      notify(err instanceof Wurd.UserError ? err.message : `Couldn't save as .${format}: ${err.message}`);
    } finally {
      saveButton.disabled = false;
    }
  };

  // toServer preselects the vault, for a Save that needs the dialog (an event passed by a click leaves it unset)
  const exportAs = async ({ toServer = false } = {}) => {
    const choice = await Wurd.dialogs.exportAs({
      title: docTitle.value.trim() || 'Untitled',
      format: docFormat.dataset.format,
      // Offered only with a server connected; preselected when the document already lives there
      destination: Wurd.server.connected() ? (serverFile || toServer ? 'server' : 'download') : null,
    });
    if (choice) await writeFile({ ...choice, to: choice.destination === 'server' ? 'new-on-server' : 'download' });
  };

  // Save keeps the document's own format (and its lock, if it was opened with a password). While connected
  // to a Wurd server it goes to the vault: back to the document's copy there, or as a new one.
  const save = () => {
    const connected = Wurd.server.connected();
    const format = SAVE_FORMATS.includes(docFormat.dataset.format) ? docFormat.dataset.format : 'docx';
    const lock = documentLock?.format === format ? documentLock : null;
    if (format === 'wurd' && !lock) return exportAs({ toServer: connected });
    return writeFile({
      title: docTitle.value.trim() || 'Untitled',
      format,
      lock,
      to: !connected ? 'download' : serverFile ? 'server' : 'new-on-server',
    });
  };

  const newDocument = () => {
    loadDocument({ html: BLANK_DOCUMENT, title: 'Untitled_Document', format: docFormat.dataset.format, font: fontSelect.value });
    setServerFile(null);
    notify('New document created');
  };

  // ---- Drafts -------------------------------------------------------------

  // Keeps the open document for a reload (js/drafts.js). True if a reload will bring it back.
  const keepDraft = () => {
    clearTimeout(draftTimer);
    const html = documentHtml();
    if (html === SAMPLE.html && docTitle.value === SAMPLE.title && !serverFile) {
      Wurd.drafts.clear();
      return true;
    }
    return Wurd.drafts.save({
      html,
      title: docTitle.value,
      format: docFormat.dataset.format,
      font: fontSelect.value,
      server: serverFile && Wurd.server.connected() ? { url: Wurd.server.url(), ...serverFile } : null,
      protected: Boolean(documentLock) || docFormat.dataset.format === 'wurd',
    });
  };

  // Kept a second after edits stop, and straight away when the page is left
  function scheduleDraft() {
    clearTimeout(draftTimer);
    draftTimer = setTimeout(keepDraft, 1000);
  }

  source.addEventListener('input', scheduleDraft);

  // When the document can't be kept (password-protected, too large, drafts off), the browser asks
  // before a reload or close loses unsaved changes
  window.addEventListener('beforeunload', (e) => {
    if (!keepDraft() && documentHtml() !== cleanHtml) e.preventDefault();
  });
  // Mobile browsers can skip beforeunload
  window.addEventListener('pagehide', keepDraft);

  const openSettings = () => {
    keepDraft();
    Wurd.dialogs.showSettings({ onChange: keepDraft });
  };

  // ---- Header -------------------------------------------------------------

  $('#new-doc').addEventListener('click', newDocument);
  $('#open-doc').addEventListener('click', () => fileInput.click());
  $('#export-doc').addEventListener('click', exportAs);
  saveButton.addEventListener('click', save);
  $('#shortcuts-help').addEventListener('click', () => Wurd.dialogs.showShortcuts());
  $('#settings-open').addEventListener('click', openSettings);

  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    // Clear the input so picking the same file again still fires change
    fileInput.value = '';
    if (file) openFile(file);
  });

  docTitle.addEventListener('input', () => {
    updateWindowTitle();
    scheduleDraft();
  });

  const syncThemeToggle = () => themeToggle.setAttribute('aria-pressed', String(root.dataset.theme === 'dark'));
  themeToggle.addEventListener('click', () => {
    const theme = root.dataset.theme === 'dark' ? 'light' : 'dark';
    root.dataset.theme = theme;
    syncThemeToggle();
    try {
      localStorage.setItem('wurd-theme', theme);
    } catch (e) {
      // Storage unavailable: the toggle still works for this visit
    }
  });

  // ---- Ribbon toolbar -----------------------------------------------------

  document.querySelectorAll('[data-command]').forEach((button) => {
    // Keep focus (and the selection) in the editor when clicking a formatting button
    button.addEventListener('mousedown', (e) => e.preventDefault());
    button.addEventListener('click', () => exec(button.dataset.command));
  });

  undoButton.addEventListener('click', () => {
    if (!isSourceMode) history.undo();
  });
  redoButton.addEventListener('click', () => {
    if (!isSourceMode) history.redo();
  });

  blockSelect.addEventListener('change', (e) => exec('formatBlock', e.target.value));

  fontSelect.addEventListener('change', (e) => {
    // Apply to the selection before the page font changes, or the browser sees the selection
    // as already being in that font and doesn't mark it
    exec('fontName', e.target.value);
    setPageFont(e.target.value);
    scheduleDraft();
  });

  // The toolbar is one Tab stop; arrow keys, Home and End move between its controls
  const toolbarItems = () => Array.from(ribbon.querySelectorAll('button, select')).filter((el) => !el.disabled && el.getClientRects().length > 0);

  const setToolbarStop = (item) => {
    ribbon.querySelectorAll('button, select').forEach((el) => {
      el.tabIndex = el === item ? 0 : -1;
    });
  };

  // Keeps exactly one Tab stop, moving it if its control became disabled or hidden
  // (Undo at the start of history, the font menu on narrow screens)
  function ensureToolbarStop() {
    const items = toolbarItems();
    setToolbarStop(items.find((el) => el.tabIndex === 0) || items[0]);
  }

  ribbon.addEventListener('focusin', (e) => {
    if (e.target.matches('button, select')) setToolbarStop(e.target);
  });

  ribbon.addEventListener('keydown', (e) => {
    const items = toolbarItems();
    const index = items.indexOf(document.activeElement);
    if (index < 0) return;
    const next = {
      ArrowRight: items[(index + 1) % items.length],
      ArrowLeft: items[(index - 1 + items.length) % items.length],
      Home: items[0],
      End: items[items.length - 1],
    }[e.key];
    if (!next) return;
    e.preventDefault();
    setToolbarStop(next);
    next.focus();
  });

  window.addEventListener('resize', ensureToolbarStop);

  // ---- Find & replace -----------------------------------------------------

  let findOpener = null;

  const setFindOpen = (open) => {
    if (open && findbar.hidden) findOpener = document.activeElement;
    findbar.hidden = !open;
    findToggle.classList.toggle('is-active', open);
    findToggle.setAttribute('aria-expanded', String(open));
    if (open) {
      findInput.focus();
      findInput.select();
    } else if (findbar.contains(document.activeElement) || document.activeElement === document.body) {
      (findOpener && findOpener.isConnected ? findOpener : editor).focus();
    }
  };

  findToggle.addEventListener('click', () => setFindOpen(findbar.hidden));
  $('#find-close').addEventListener('click', () => setFindOpen(false));

  const replaceAll = () => {
    const searchTerm = findInput.value;
    const replaceTerm = replaceInput.value;
    if (!searchTerm || isSourceMode) return;
    // Match the search term literally, and only inside text nodes so tags and attributes are never touched
    const regex = new RegExp(searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    let count = 0;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const text = node.nodeValue;
      // A replacer function stops `$&`, `$1` etc. in the replacement from being expanded
      const replaced = text.replace(regex, () => {
        count++;
        return replaceTerm;
      });
      if (replaced !== text) node.nodeValue = replaced;
    }
    if (count === 0) {
      notify('No occurrences found');
      return;
    }
    history.record('replace');
    updateStatsAndOutline();
    notify(`Replaced ${count} occurrence${count === 1 ? '' : 's'} of "${searchTerm}"`);
  };

  $('#replace-all').addEventListener('click', replaceAll);

  findbar.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      setFindOpen(false);
    } else if (e.key === 'Enter' && e.target.matches('input')) {
      e.preventDefault();
      replaceAll();
    }
  });

  // ---- Source view and zoom -----------------------------------------------

  const toggleSource = () => {
    const hadFocus = document.activeElement === (isSourceMode ? source : editor);
    if (!isSourceMode) {
      source.value = editor.innerHTML;
    } else {
      editor.innerHTML = Wurd.sanitizeHtml(source.value);
    }
    isSourceMode = !isSourceMode;
    editor.hidden = isSourceMode;
    source.hidden = !isSourceMode;
    sourceToggle.classList.toggle('is-active', isSourceMode);
    sourceToggleLabel.textContent = isSourceMode ? 'WYSIWYG' : 'Source';
    if (!isSourceMode) history.record('source');
    updateStatsAndOutline();
    if (hadFocus) (isSourceMode ? source : editor).focus();
  };

  sourceToggle.addEventListener('click', toggleSource);

  const setZoom = (level) => {
    zoom = level;
    page.style.transform = `scale(${zoom / 100})`;
    zoomLevel.textContent = `${zoom}%`;
  };

  $('#zoom-out').addEventListener('click', () => setZoom(Math.max(50, zoom - 10)));
  $('#zoom-in').addEventListener('click', () => setZoom(Math.min(200, zoom + 10)));

  // ---- Sidebar ------------------------------------------------------------

  const tabs = Array.from(document.querySelectorAll('[role="tab"]'));

  const selectTab = (tab, focus = false) => {
    tabs.forEach((t) => {
      const active = t === tab;
      t.classList.toggle('is-active', active);
      t.setAttribute('aria-selected', String(active));
      t.tabIndex = active ? 0 : -1;
    });
    document.querySelectorAll('[data-panel]').forEach((panel) => {
      panel.hidden = panel.dataset.panel !== tab.dataset.tab;
    });
    if (focus) tab.focus();
    if (tab.dataset.tab === 'security') refreshServerList();
  };

  tabs.forEach((tab, i) => {
    tab.addEventListener('click', () => selectTab(tab));
    tab.addEventListener('keydown', (e) => {
      const next = {
        ArrowRight: tabs[(i + 1) % tabs.length],
        ArrowLeft: tabs[(i - 1 + tabs.length) % tabs.length],
        Home: tabs[0],
        End: tabs[tabs.length - 1],
      }[e.key];
      if (!next) return;
      e.preventDefault();
      selectTab(next, true);
    });
  });

  // ---- Wurd server --------------------------------------------------------

  const vaultPanel = $('#panel-security');
  const vaultList = $('#vault-list');
  const vaultEmpty = $('#vault-empty');
  const serverStatus = $('#server-status');
  const serverForm = $('#server-form');
  const serverUrl = $('#server-url');
  const serverToken = $('#server-token');
  const serverRememberUrl = $('#server-remember-url');
  const serverKeepToken = $('#server-keep-token');
  const serverError = $('#server-error');
  const serverConnect = $('#server-connect');
  const serverRequest = $('#server-request');
  const serverDisconnect = $('#server-disconnect');
  const tokenReveal = $('#token-reveal');
  const tokenNew = $('#token-new');
  const tokenCopy = $('#token-copy');

  const relativeTime = new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' });
  const TIME_UNITS = [['year', 31536000], ['month', 2592000], ['week', 604800], ['day', 86400], ['hour', 3600], ['minute', 60]];
  const timeAgo = (iso) => {
    const seconds = (Date.parse(iso) - Date.now()) / 1000;
    const [unit, size] = TIME_UNITS.find(([, s]) => Math.abs(seconds) >= s) || [];
    return unit ? relativeTime.format(Math.round(seconds / size), unit) : 'just now';
  };

  const span = (className, text) => {
    const el = document.createElement('span');
    el.className = className;
    el.textContent = text;
    return el;
  };

  const showVaultMessage = (text) => {
    vaultEmpty.textContent = text;
    vaultEmpty.hidden = false;
    vaultList.hidden = true;
  };

  const openServerFile = async (file) => {
    try {
      const { file: download, rev } = await Wurd.server.download(file);
      await openFile(download, { id: file.id, rev });
    } catch (err) {
      if (!(err instanceof Wurd.UserError)) console.error(err);
      notify(err instanceof Wurd.UserError ? err.message : `Couldn't open "${file.name}" from the server.`);
    }
  };

  const deleteServerFile = async (file) => {
    if (!window.confirm(`Delete "${file.name}" from the server? This can't be undone.`)) return;
    try {
      await Wurd.server.remove(file.id, file.rev);
      if (serverFile?.id === file.id) setServerFile(null);
      notify(`Deleted "${file.name}" from the server`);
    } catch (err) {
      if (!(err instanceof Wurd.UserError)) console.error(err);
      notify(err.status === 412 ? `"${file.name}" was changed somewhere else, so it wasn't deleted.` : err instanceof Wurd.UserError ? err.message : `Couldn't delete "${file.name}".`);
    }
    // The button is gone once the list refreshes, so keep focus in the panel
    vaultPanel.focus();
    refreshServerList();
  };

  const renderServerFile = (file) => {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'recent-card';
    if (serverFile?.id === file.id) card.setAttribute('aria-current', 'true');
    const row = span('recent-card__row', '');
    row.append(span('recent-card__title', file.name.replace(/\.[^.]+$/, '')), span('recent-card__format', `.${Wurd.formats.extensionOf(file.name)}`));
    const updated = span('recent-card__updated', timeAgo(file.updatedAt));
    updated.title = new Date(file.updatedAt).toLocaleString();
    card.append(row, updated);
    card.addEventListener('click', () => openServerFile(file));

    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'recent-item__delete';
    remove.title = 'Delete from the server';
    remove.setAttribute('aria-label', `Delete ${file.name} from the server`);
    const icon = document.createElement('i');
    icon.dataset.lucide = 'trash-2';
    icon.className = 'icon icon--sm';
    remove.append(icon);
    remove.addEventListener('click', () => deleteServerFile(file));

    const item = document.createElement('li');
    item.className = 'recent-item';
    item.append(card, remove);
    return item;
  };

  // Lists your documents on the server in the Vault tab, newest first
  async function refreshServerList() {
    if (!Wurd.server.connected()) {
      showVaultMessage('Connect to a Wurd server above to keep your documents there.');
      return;
    }
    let files;
    try {
      files = await Wurd.server.list();
    } catch (err) {
      if (!(err instanceof Wurd.UserError)) console.error(err);
      showVaultMessage(err instanceof Wurd.UserError ? err.message : "Couldn't load your documents from the server.");
      return;
    }
    if (files.length === 0) {
      showVaultMessage('No documents in your vault yet. Press Save to keep the open one here.');
      return;
    }
    vaultList.replaceChildren(...files.map(renderServerFile));
    lucide.createIcons({ root: vaultList, attrs: ICON_ATTRS });
    vaultEmpty.hidden = true;
    vaultList.hidden = false;
  }

  const setServerError = (message) => {
    serverError.textContent = message;
    serverError.hidden = !message;
    if (message) serverToken.setAttribute('aria-invalid', 'true');
    else serverToken.removeAttribute('aria-invalid');
  };

  const syncServerPanel = () => {
    const connected = Wurd.server.connected();
    serverStatus.textContent = connected
      ? `Connected to ${Wurd.server.url()}. Save now keeps documents in your vault.`
      : 'Not connected. Documents are saved to this device only.';
    serverForm.hidden = connected;
    serverDisconnect.hidden = !connected;
    updateSaveTitle();
    serverRememberUrl.checked = Boolean(Wurd.server.savedUrl());
    if (connected) serverKeepToken.checked = Wurd.server.keptInTab();
    if (!connected && !serverUrl.value) {
      // Served by a Wurd server, the API is on the same address
      serverUrl.value = Wurd.server.savedUrl() || (/^https?:$/.test(location.protocol) ? location.origin : '');
    }
  };

  serverForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    serverConnect.disabled = true;
    try {
      await Wurd.server.connect(serverUrl.value, serverToken.value, {
        rememberUrl: serverRememberUrl.checked,
        keepInTab: serverKeepToken.checked,
      });
      serverToken.value = '';
      setServerError('');
      syncServerPanel();
      serverDisconnect.focus();
      refreshServerList();
      notify(`Connected to ${Wurd.server.url()}`);
    } catch (err) {
      if (!(err instanceof Wurd.UserError)) console.error(err);
      setServerError(err instanceof Wurd.UserError ? err.message : `Couldn't connect: ${err.message}`);
      serverToken.focus();
    } finally {
      serverConnect.disabled = false;
    }
  });

  // A new token opens a new, empty vault. It's shown once and filled in, ready to connect with.
  serverRequest.addEventListener('click', async () => {
    serverRequest.disabled = true;
    try {
      const token = await Wurd.server.requestToken(serverUrl.value);
      tokenNew.value = token;
      tokenReveal.hidden = false;
      serverToken.value = token;
      setServerError('');
      tokenNew.focus();
      tokenNew.select();
    } catch (err) {
      if (!(err instanceof Wurd.UserError)) console.error(err);
      setServerError(err instanceof Wurd.UserError ? err.message : `Couldn't request a token: ${err.message}`);
    } finally {
      serverRequest.disabled = false;
    }
  });

  tokenCopy.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(tokenNew.value);
      notify('Token copied. Keep it somewhere safe.');
    } catch (e) {
      // No clipboard access (e.g. not a secure context): leave it selected to copy by hand
      tokenNew.focus();
      tokenNew.select();
      notify('Select and copy the token by hand.');
    }
  });

  serverDisconnect.addEventListener('click', () => {
    Wurd.server.disconnect();
    tokenReveal.hidden = true;
    tokenNew.value = '';
    setServerFile(null);
    syncServerPanel();
    serverUrl.focus();
    refreshServerList();
    notify('Disconnected from the Wurd server');
  });

  // ---- Editor -------------------------------------------------------------

  editor.addEventListener('input', updateStatsAndOutline);
  editor.addEventListener('keyup', updateStatsAndOutline);

  let toolStateFrame = 0;
  document.addEventListener('selectionchange', () => {
    cancelAnimationFrame(toolStateFrame);
    toolStateFrame = requestAnimationFrame(updateToolStates);
  });

  $('#skip-link').addEventListener('click', (e) => {
    e.preventDefault();
    const target = isSourceMode ? source : editor;
    target.focus();
    target.scrollIntoView({ block: 'start' });
  });

  // ---- Keyboard shortcuts -------------------------------------------------

  const formatBlock = (tag) => () => exec('formatBlock', tag);
  Wurd.shortcuts.init({
    undo: () => !isSourceMode && history.undo(),
    redo: () => !isSourceMode && history.redo(),
    bold: () => exec('bold'),
    italic: () => exec('italic'),
    underline: () => exec('underline'),
    strikeThrough: () => exec('strikeThrough'),
    p: formatBlock('p'),
    h1: formatBlock('h1'),
    h2: formatBlock('h2'),
    h3: formatBlock('h3'),
    justifyLeft: () => exec('justifyLeft'),
    justifyCenter: () => exec('justifyCenter'),
    justifyRight: () => exec('justifyRight'),
    justifyFull: () => exec('justifyFull'),
    insertOrderedList: () => exec('insertOrderedList'),
    insertUnorderedList: () => exec('insertUnorderedList'),
    new: newDocument,
    open: () => fileInput.click(),
    save,
    exportAs,
    find: () => setFindOpen(true),
    source: toggleSource,
    help: () => Wurd.dialogs.showShortcuts(),
    settings: openSettings,
  });

  // ---- Start --------------------------------------------------------------

  lucide.createIcons({ attrs: ICON_ATTRS });
  syncThemeToggle();
  updateWindowTitle();
  setZoom(zoom);
  setPageFont(Wurd.DEFAULT_FONT);
  history.reset();
  // A reload brings back the document that was open
  const draft = Wurd.drafts.restore();
  if (draft) {
    loadDocument(draft);
    // Whether it had unsaved changes isn't known, so it counts as changed
    cleanHtml = null;
    setServerFile(draft.server?.url === Wurd.server.url() ? { id: draft.server.id, rev: draft.server.rev } : null);
    notify(`Restored "${draft.title || 'Untitled'}" after the reload`);
  }
  syncServerPanel();
  refreshServerList();

  // For automated tests and debugging
  Wurd.app = { history, loadDocument, save, exportAs, openFile, isSourceMode: () => isSourceMode, serverFile: () => serverFile };
})(window.Wurd = window.Wurd || {});
