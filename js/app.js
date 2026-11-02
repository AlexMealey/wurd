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
  const sizeSelect = $('#font-size');
  const textBoxButton = $('#insert-textbox');
  const imageInput = $('#image-input');
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
  const SAVE_FORMATS = ['docx', 'odt', 'odg', 'doc', 'html', 'txt', 'wurd'];
  // Formats that keep the password a document was opened with
  const PROTECTABLE = ['odt', 'odg', 'wurd'];
  const ICON_ATTRS = { 'aria-hidden': 'true', focusable: 'false' };

  let isSourceMode = false;
  let zoom = 100;
  // The document's default font (the font menu shows the font at the caret, which may differ)
  let pageFont = Wurd.DEFAULT_FONT;
  // The editor selection last seen, so controls that take focus (menus, palettes, dialogs) still act on it
  let lastRange = null;
  let toastTimer = null;
  // The open document's password, kept in memory only so Save can re-encrypt it
  let documentPassword = null;
  // The Wurd server copy the open document belongs to ({ id, rev }), or null. Save writes back to it,
  // or, while connected, creates one in the vault.
  let serverFile = null;
  // The document as last opened or saved; null when unknown, so it counts as changed
  let cleanHtml = editor.innerHTML;
  let draftTimer = 0;
  // The sample the page opens with needs no draft: a reload shows it anyway
  const SAMPLE = { html: editor.innerHTML, title: docTitle.value };

  // ---- Notifications ------------------------------------------------------

  // Clearing first makes screen readers announce a repeated message again
  const announce = (msg) => {
    announcer.textContent = '';
    setTimeout(() => {
      announcer.textContent = msg;
    }, 50);
  };

  const notify = (msg) => {
    toastMessage.textContent = msg;
    toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      toast.hidden = true;
    }, 3000);
    announce(msg);
  };

  // ---- Document state -----------------------------------------------------

  const setFormat = (format) => {
    docFormat.dataset.format = format;
    docFormat.textContent = `.${format}`;
  };

  // The document's default font. The font menu changes only the selection.
  const setPageFont = (font) => {
    pageFont = font;
    fontSelect.value = font;
    page.style.fontFamily = font;
  };

  // The on-screen page follows the page size setting (A4 or Letter)
  const applyPageSize = (size) => {
    page.dataset.size = size;
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
    // Typing can change a text box's size
    Wurd.textboxes.position();
  };

  // Pressed states for the formatting buttons and the current paragraph style, from the selection
  const updateToolStates = () => {
    // The move and resize handles follow the text box the caret is in
    Wurd.textboxes.update();
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
    const style = Wurd.formatting.currentStyle();
    if (style) {
      fontSelect.value = style.font;
      showFontSize(style.size);
    }
    textBoxButton.setAttribute('aria-pressed', String(Wurd.insert.inTextBox(selection.anchorNode)));
  };

  // Shows a size in the size menu, adding it for the moment if it isn't one of the menu's sizes
  const showFontSize = (size) => {
    const value = String(size);
    sizeSelect.querySelector('option[data-extra]')?.remove();
    if (!Array.from(sizeSelect.options).some((o) => o.value === value)) {
      const option = new Option(value, value);
      option.dataset.extra = '';
      const after = Array.from(sizeSelect.options).find((o) => Number(o.value) > size);
      sizeSelect.add(option, after || null);
    }
    sizeSelect.value = value;
  };

  // Before the history: typed text picks up a pending font, size or colour before it's recorded
  Wurd.formatting.init(editor);
  Wurd.insert.init(editor);

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

  // Text boxes: moving, resizing, free or in line. Handles are drawn on the (zoomed) page.
  Wurd.textboxes.init({ editor, history, getScale: () => zoom / 100 });

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
    cleanHtml = editor.innerHTML;
    lastRange = null;
    Wurd.formatting.clearPending();
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

  // to: 'download', 'server' (the open document's server copy) or 'new-on-server'
  const writeFile = async ({ title, format, password = null, to = 'download' }) => {
    saveButton.disabled = true;
    try {
      const html = documentHtml();
      // A drawing places free text boxes on its pages, which needs the editor's layout (not there in source view)
      const boxLayout = isSourceMode ? null : Wurd.textboxes.layout();
      const blob = await Wurd.formats.export(format, html, { title, pageFont, pageSize: Wurd.pageSize(), password, boxLayout });
      const fileName = Wurd.formats.fileName(title, format);
      const withPassword = password ? ' with a password' : '';
      if (to === 'download') {
        Wurd.downloadBlob(blob, fileName);
        cleanHtml = html;
        notify(`Saved "${fileName}"${withPassword}`);
        return;
      }
      const saved = to === 'server'
        ? await Wurd.server.update(serverFile.id, serverFile.rev, fileName, blob)
        : await Wurd.server.create(fileName, blob);
      // The document is now that server file, so Save keeps its name, format and password
      setServerFile({ id: saved.id, rev: saved.rev });
      docTitle.value = title;
      updateWindowTitle();
      setFormat(format);
      documentPassword = password;
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

  // Save keeps the document's own format (and its password, if it was opened with one). While connected
  // to a Wurd server it goes to the vault: back to the document's copy there, or as a new one.
  const save = () => {
    const connected = Wurd.server.connected();
    const format = SAVE_FORMATS.includes(docFormat.dataset.format) ? docFormat.dataset.format : 'docx';
    if (format === 'wurd' && !documentPassword) return exportAs({ toServer: connected });
    return writeFile({
      title: docTitle.value.trim() || 'Untitled',
      format,
      password: PROTECTABLE.includes(format) ? documentPassword : null,
      to: !connected ? 'download' : serverFile ? 'server' : 'new-on-server',
    });
  };

  const newDocument = () => {
    loadDocument({ html: BLANK_DOCUMENT, title: 'Untitled_Document', format: docFormat.dataset.format, font: pageFont });
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
      font: pageFont,
      server: serverFile && Wurd.server.connected() ? { url: Wurd.server.url(), ...serverFile } : null,
      protected: Boolean(documentPassword) || docFormat.dataset.format === 'wurd',
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
    Wurd.dialogs.showSettings({
      onChange: keepDraft,
      onPageSize: (size) => {
        applyPageSize(size);
        notify(`Page size set to ${Wurd.PAGE_SIZES[size].label}`);
      },
    });
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

  const undo = () => {
    if (isSourceMode) return;
    Wurd.formatting.clearPending();
    history.undo();
  };
  const redo = () => {
    if (isSourceMode) return;
    Wurd.formatting.clearPending();
    history.redo();
  };
  undoButton.addEventListener('click', undo);
  redoButton.addEventListener('click', redo);

  blockSelect.addEventListener('change', (e) => exec('formatBlock', e.target.value));

  // ---- Editing helpers ------------------------------------------------------

  const rememberRange = () => {
    const selection = window.getSelection();
    if (selection.rangeCount && editor.contains(selection.getRangeAt(0).commonAncestorContainer)) {
      lastRange = selection.getRangeAt(0).cloneRange();
    }
  };

  const inEditor = (range) => Boolean(range && editor.contains(range.startContainer) && editor.contains(range.endContainer));

  // Focuses the editor and puts the selection where it should act: `saved` if given (a range kept
  // across a dialog or a drop), else the selection if it's still in the editor (a menu or palette
  // doesn't move it), else the editor selection last seen. True if the selection is in the editor.
  const restoreRange = (saved = null) => {
    const selection = window.getSelection();
    const current = selection.rangeCount ? selection.getRangeAt(0) : null;
    const target = inEditor(saved) ? saved : inEditor(current) ? null : inEditor(lastRange) ? lastRange : null;
    editor.focus({ preventScroll: true });
    if (target) {
      selection.removeAllRanges();
      selection.addRange(target);
    }
    return selection.rangeCount > 0 && editor.contains(selection.getRangeAt(0).commonAncestorContainer);
  };

  // The editor's selection now, if it's there
  const editorRange = () => {
    const selection = window.getSelection();
    const range = selection.rangeCount ? selection.getRangeAt(0) : null;
    return inEditor(range) ? range.cloneRange() : lastRange;
  };

  // Runs one change to the document at the editor's selection (or at `range`) as one undo step.
  // Returns what fn returned.
  const edit = (kind, fn, range = null) => {
    if (isSourceMode || !restoreRange(range)) return undefined;
    let result;
    history.batch(kind, () => {
      result = fn(window.getSelection().getRangeAt(0));
    });
    updateStatsAndOutline();
    updateToolStates();
    return result;
  };

  // Font, size, colour or highlight for the selection, or for what's typed next at a caret.
  // value null takes the property away (Automatic colour, No highlight, the document's own font).
  const applyInline = (prop, value) => {
    if (isSourceMode || !restoreRange()) return;
    const selection = window.getSelection();
    if (selection.isCollapsed) {
      Wurd.formatting.setPending(prop, value);
      updateToolStates();
      return;
    }
    edit(prop, (range) => {
      const styled = Wurd.formatting.applyStyle(range, prop, value);
      if (styled) {
        selection.removeAllRanges();
        selection.addRange(styled);
      }
    });
  };

  fontSelect.addEventListener('change', () => {
    // The document's own font needs no mark of its own
    applyInline('font-family', fontSelect.value === pageFont ? null : fontSelect.value);
  });

  sizeSelect.addEventListener('change', () => applyInline('font-size', `${sizeSelect.value}pt`));

  // Colour palettes: positioned under their button, and each choice applies straight away
  document.querySelectorAll('.palette').forEach((palette) => {
    const prop = palette.dataset.property;
    const button = document.querySelector(`[popovertarget="${palette.id}"]`);
    const swatch = button.querySelector('.tool__swatch');
    const picker = palette.querySelector('.palette__picker');
    button.addEventListener('mousedown', (e) => e.preventDefault());
    palette.addEventListener('beforetoggle', (e) => {
      if (e.newState !== 'open') return;
      rememberRange();
      const rect = button.getBoundingClientRect();
      palette.style.top = `${Math.round(rect.bottom + 4)}px`;
      palette.style.left = `${Math.round(Math.max(8, Math.min(rect.left, window.innerWidth - palette.offsetWidth - 8, window.innerWidth - 240)))}px`;
    });
    palette.addEventListener('toggle', (e) => {
      if (e.newState === 'open') palette.querySelector('.palette__none').focus();
    });
    // Clicking a swatch shouldn't take the selection out of the editor
    palette.addEventListener('mousedown', (e) => {
      if (!e.target.closest('label, input')) e.preventDefault();
    });
    const choose = (color) => {
      palette.hidePopover();
      if (color) swatch.style.setProperty('--swatch', color);
      applyInline(prop, color || null);
    };
    palette.addEventListener('click', (e) => {
      const option = e.target.closest('[data-color]');
      if (option) choose(option.dataset.color);
    });
    picker.addEventListener('change', () => choose(picker.value.toUpperCase()));
  });

  const indentBy = (direction) => edit(direction > 0 ? 'indent' : 'outdent', (range) => Wurd.formatting.indent(range, direction));

  // ---- Inserting ------------------------------------------------------------

  // Opens the link dialog for the link at the caret (to change or remove it) or for a new one
  const editLink = async () => {
    if (isSourceMode || !restoreRange()) return;
    const range = window.getSelection().getRangeAt(0).cloneRange();
    const link = Wurd.insert.linkAt(range.startContainer) || Wurd.insert.linkAt(range.endContainer);
    const selected = range.toString().trim();
    const choice = await Wurd.dialogs.editLink({
      href: link ? link.getAttribute('href') : (/^(https?:\/\/|www\.)\S+$|^\S+@\S+\.\S+$/i.test(selected) ? selected : ''),
      text: link ? link.textContent : '',
      showText: Boolean(link) || range.collapsed,
      editing: Boolean(link),
    });
    if (!choice) {
      restoreRange(range);
      return;
    }
    edit('link', () => Wurd.insert.setLink({ link, href: choice.remove ? null : choice.href, text: choice.text }), range);
    notify(choice.remove ? 'Link removed' : link ? 'Link updated' : 'Link inserted');
  };

  // Inserts image files at `range` (where the caret was when they were picked, pasted or dropped)
  const insertImageFiles = async (files, range = editorRange()) => {
    const images = [];
    for (const file of files) {
      try {
        images.push(await Wurd.insert.imageFromFile(file));
      } catch (err) {
        if (!(err instanceof Wurd.UserError)) console.error(err);
        notify(err instanceof Wurd.UserError ? err.message : `Couldn't insert "${file.name}".`);
      }
    }
    if (images.length && edit('image', () => Wurd.insert.insertImages(images), range)) {
      notify(images.length === 1 ? 'Image inserted' : `${images.length} images inserted`);
    }
  };

  const insertSymbol = async () => {
    if (isSourceMode) return;
    const range = editorRange();
    const ch = await Wurd.dialogs.pickSymbol();
    if (ch) edit('symbol', () => Wurd.insert.insertText(ch), range);
    else restoreRange(range);
  };

  const insertPageBreak = () => edit('page-break', () => Wurd.insert.insertPageBreak());
  const toggleTextBox = () => edit('text-box', () => Wurd.insert.toggleTextBox());

  // Ribbon buttons that act on the editor keep its selection when clicked
  ['#indent', '#outdent', '#insert-link', '#insert-image', '#insert-textbox', '#insert-symbol', '#insert-page-break'].forEach((selector) => {
    $(selector).addEventListener('mousedown', (e) => e.preventDefault());
  });
  $('#indent').addEventListener('click', () => indentBy(1));
  $('#outdent').addEventListener('click', () => indentBy(-1));
  $('#insert-link').addEventListener('click', editLink);
  // Where the caret was when the file picker opened
  let imageRange = null;
  $('#insert-image').addEventListener('click', () => {
    if (isSourceMode) return;
    imageRange = editorRange();
    imageInput.click();
  });
  imageInput.addEventListener('change', () => {
    const files = Array.from(imageInput.files || []);
    imageInput.value = '';
    if (files.length) insertImageFiles(files, imageRange || editorRange());
  });
  textBoxButton.addEventListener('click', toggleTextBox);
  $('#insert-symbol').addEventListener('click', insertSymbol);
  $('#insert-page-break').addEventListener('click', insertPageBreak);

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
    Wurd.textboxes.update();
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
    rememberRange();
    cancelAnimationFrame(toolStateFrame);
    toolStateFrame = requestAnimationFrame(updateToolStates);
  });

  // Pasting a picture (a screenshot, an image copied from a web page) inserts it. Anything with
  // text, including a copy from Word that also carries a picture of itself, pastes as text.
  editor.addEventListener('paste', (e) => {
    if (isSourceMode || !e.clipboardData) return;
    const files = Array.from(e.clipboardData.files || []).filter(Wurd.insert.isImageFile);
    if (!files.length || e.clipboardData.getData('text/plain').trim()) return;
    e.preventDefault();
    insertImageFiles(files, editorRange());
  });

  const hasImageFiles = (transfer) => Array.from(transfer?.items || []).some((item) => item.kind === 'file' && /^image\//.test(item.type));

  editor.addEventListener('dragover', (e) => {
    if (isSourceMode || !hasImageFiles(e.dataTransfer)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  });

  // Dropped images go where they're dropped
  editor.addEventListener('drop', (e) => {
    const files = Array.from(e.dataTransfer?.files || []).filter(Wurd.insert.isImageFile);
    if (isSourceMode || !files.length) return;
    e.preventDefault();
    let range = null;
    if (document.caretPositionFromPoint) {
      const position = document.caretPositionFromPoint(e.clientX, e.clientY);
      if (position) {
        range = document.createRange();
        range.setStart(position.offsetNode, position.offset);
      }
    } else if (document.caretRangeFromPoint) {
      range = document.caretRangeFromPoint(e.clientX, e.clientY);
    }
    if (range && editor.contains(range.startContainer)) range.collapse(true);
    insertImageFiles(files, inEditor(range) ? range : editorRange());
  });

  // Links can't be followed by a plain click while editing; Ctrl+click (Cmd+click on a Mac) opens them
  editor.addEventListener('click', (e) => {
    if (!(Wurd.shortcuts.isMac() ? e.metaKey : e.ctrlKey)) return;
    const link = Wurd.insert.linkAt(e.target);
    if (!link || !Wurd.insert.openable(link.getAttribute('href'))) return;
    e.preventDefault();
    window.open(link.href, '_blank', 'noopener,noreferrer');
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
    undo,
    redo,
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
    indent: () => indentBy(1),
    outdent: () => indentBy(-1),
    link: editLink,
    pageBreak: insertPageBreak,
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
  applyPageSize(Wurd.pageSize());
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

  // For js/pwa.js, automated tests and debugging
  Wurd.app = {
    history, loadDocument, save, exportAs, openFile, notify, announce, isSourceMode: () => isSourceMode, serverFile: () => serverFile,
    pageFont: () => pageFont, applyInline, indentBy, insertPageBreak, toggleTextBox, insertImageFiles,
  };
})(window.Wurd = window.Wurd || {});
