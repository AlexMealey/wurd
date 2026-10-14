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

  let isSourceMode = false;
  let zoom = 100;
  let toastTimer = null;
  // The open document's password, kept in memory only so Save can re-encrypt it
  let documentPassword = null;

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
    },
  });

  // Puts a document into the editor (or the source view). Everything loaded is sanitized here.
  const loadDocument = ({ html, title, format, font, password = null }) => {
    const clean = Wurd.sanitizeHtml(html);
    editor.innerHTML = clean;
    if (isSourceMode) source.value = format === 'html' ? html : clean;
    docTitle.value = title;
    updateWindowTitle();
    setFormat(format);
    setPageFont(font || Wurd.DEFAULT_FONT);
    documentPassword = password;
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

  const openFile = async (file) => {
    try {
      const doc = await Wurd.formats.open(file, {
        askPassword: (options) => Wurd.dialogs.askPassword({ fileName: file.name, ...options }),
      });
      loadDocument(doc);
      notify(doc.notice || `Loaded "${file.name}"`);
    } catch (err) {
      if (err instanceof Wurd.Cancelled) return;
      if (!(err instanceof Wurd.UserError)) console.error(err);
      notify(err instanceof Wurd.UserError ? err.message : `Couldn't open "${file.name}". It may be damaged or in an unsupported format.`);
    }
  };

  const documentHtml = () => (isSourceMode ? Wurd.sanitizeHtml(source.value) : editor.innerHTML);

  const writeFile = async ({ title, format, password = null }) => {
    saveButton.disabled = true;
    try {
      const blob = await Wurd.formats.export(format, documentHtml(), { title, pageFont: fontSelect.value, password });
      const fileName = Wurd.formats.fileName(title, format);
      Wurd.downloadBlob(blob, fileName);
      notify(`Saved "${fileName}"${password ? ' with a password' : ''}`);
    } catch (err) {
      if (!(err instanceof Wurd.UserError)) console.error(err);
      notify(err instanceof Wurd.UserError ? err.message : `Couldn't save as .${format}: ${err.message}`);
    } finally {
      saveButton.disabled = false;
    }
  };

  const exportAs = async () => {
    const choice = await Wurd.dialogs.exportAs({ title: docTitle.value.trim() || 'Untitled', format: docFormat.dataset.format });
    if (choice) await writeFile(choice);
  };

  // Save keeps the document's own format (and its password, if it was opened with one)
  const save = () => {
    const format = SAVE_FORMATS.includes(docFormat.dataset.format) ? docFormat.dataset.format : 'docx';
    if (format === 'wurd' && !documentPassword) return exportAs();
    return writeFile({
      title: docTitle.value.trim() || 'Untitled',
      format,
      password: format === 'odt' || format === 'wurd' ? documentPassword : null,
    });
  };

  const newDocument = () => {
    loadDocument({ html: BLANK_DOCUMENT, title: 'Untitled_Document', format: docFormat.dataset.format, font: fontSelect.value });
    notify('New document created');
  };

  // ---- Header -------------------------------------------------------------

  $('#new-doc').addEventListener('click', newDocument);
  $('#open-doc').addEventListener('click', () => fileInput.click());
  $('#export-doc').addEventListener('click', exportAs);
  saveButton.addEventListener('click', save);
  $('#shortcuts-help').addEventListener('click', () => Wurd.dialogs.showShortcuts());

  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    // Clear the input so picking the same file again still fires change
    fileInput.value = '';
    if (file) openFile(file);
  });

  docTitle.addEventListener('input', updateWindowTitle);

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

  document.querySelectorAll('.recent-card').forEach((card) => {
    card.addEventListener('click', () => {
      const { title, format } = card.dataset;
      docTitle.value = title;
      updateWindowTitle();
      setFormat(format);
      notify(`Loaded ${title}.${format}`);
    });
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
  });

  // ---- Start --------------------------------------------------------------

  lucide.createIcons({ attrs: { 'aria-hidden': 'true', focusable: 'false' } });
  syncThemeToggle();
  updateWindowTitle();
  setZoom(zoom);
  setPageFont(Wurd.DEFAULT_FONT);
  history.reset();

  // For automated tests and debugging
  Wurd.app = { history, loadDocument, save, exportAs, openFile, isSourceMode: () => isSourceMode };
})(window.Wurd = window.Wurd || {});
