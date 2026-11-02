// Putting things into the document: images, page breaks, text boxes, links and special characters.
// Each works on the current selection in the editor; the app wraps calls in an undo step.
(function (Wurd) {
  'use strict';

  let editor = null;

  const isElement = (node) => node && node.nodeType === Node.ELEMENT_NODE;
  const BLOCKISH = /^(P|DIV|H[1-6]|UL|OL|LI|BLOCKQUOTE|PRE|TABLE|THEAD|TBODY|TFOOT|TR|TD|TH)$/;

  const selectionRange = () => {
    const selection = window.getSelection();
    return selection.rangeCount && editor.contains(selection.getRangeAt(0).commonAncestorContainer) ? selection.getRangeAt(0) : null;
  };

  const setCaret = (range) => {
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  };

  const caretAfter = (node) => {
    const range = document.createRange();
    range.setStartAfter(node);
    range.collapse(true);
    setCaret(range);
  };

  // The caret at the start of a block's first line, descending into boxes and lists
  const caretInto = (el) => {
    let target = el;
    while (isElement(target.firstElementChild) && BLOCKISH.test(target.firstElementChild.tagName) && !target.firstChild.nodeValue?.trim()) {
      target = target.firstElementChild;
    }
    const range = document.createRange();
    range.selectNodeContents(target);
    range.collapse(true);
    setCaret(range);
    (isElement(target) ? target : target.parentElement)?.scrollIntoView?.({ block: 'nearest' });
  };

  // The child of the editor a node is in (null for the editor itself)
  const topLevelOf = (node) => {
    if (!node || node === editor || !editor.contains(node)) return null;
    while (node.parentNode !== editor) node = node.parentNode;
    return node;
  };

  const emptyParagraph = () => {
    const p = document.createElement('p');
    p.append(document.createElement('br'));
    return p;
  };

  const isEmptyBlock = (el) => !el.textContent.trim() && !(isElement(el) && el.querySelector('img, hr, table'));

  // ---- Images ---------------------------------------------------------------

  const IMAGE_TYPES = /^image\/(png|jpeg|gif|webp|bmp|svg\+xml)$/;
  // Kept as they are when small enough; the rest are redrawn, since Word and LibreOffice can't all show them
  const KEEP_TYPES = /^image\/(png|jpeg|gif)$/;
  const MAX_FILE_BYTES = 20 * 1024 * 1024;
  const MAX_SIDE = 2000;

  const readAsDataUrl = (blob) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });

  const decode = (src, name) => new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Wurd.UserError(`Couldn't read the image "${name}". It may be damaged.`));
    img.src = src;
  });

  const altFrom = (name) => (name || '').replace(/\.[^.]+$/, '').replace(/[_-]+/g, ' ').trim();

  const isImageFile = (file) => Boolean(file && IMAGE_TYPES.test(file.type));

  // An image file -> { src (a data: URL), width, height, alt }, sized to fit the page
  const imageFromFile = async (file) => {
    if (!isImageFile(file)) throw new Wurd.UserError(`"${file.name}" isn't an image Wurd can insert. Use a PNG, JPEG, GIF, WebP, BMP or SVG file.`);
    if (file.size > MAX_FILE_BYTES) throw new Wurd.UserError(`"${file.name}" is too large to insert (the limit is 20 MB).`);
    const original = await readAsDataUrl(file);
    const img = await decode(original, file.name);
    // An SVG without its own size gets the browser's default replaced-element size
    let width = img.naturalWidth || 300;
    let height = img.naturalHeight || 150;
    let src = original;
    const longest = Math.max(width, height);
    if (!KEEP_TYPES.test(file.type) || longest > MAX_SIDE) {
      const scale = Math.min(1, MAX_SIDE / longest);
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(width * scale));
      canvas.height = Math.max(1, Math.round(height * scale));
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      src = file.type === 'image/jpeg' ? canvas.toDataURL('image/jpeg', 0.92) : canvas.toDataURL('image/png');
      width = canvas.width;
      height = canvas.height;
    }
    const shown = Math.min(width, Wurd.contentWidthPx());
    return { src, width: shown, height: Math.max(1, Math.round((height * shown) / width)), alt: altFrom(file.name) };
  };

  // Inserts images (from imageFromFile) at the selection, replacing what's selected
  const insertImages = (images) => {
    if (!selectionRange()) return false;
    if (!window.getSelection().isCollapsed) document.execCommand('delete');
    let range = selectionRange();
    let last = null;
    for (const image of images) {
      const img = document.createElement('img');
      img.setAttribute('src', image.src);
      img.setAttribute('alt', image.alt || '');
      img.setAttribute('width', String(image.width));
      img.setAttribute('height', String(image.height));
      // Straight into the editor (between blocks), an image gets its own paragraph
      if (range.startContainer === editor) {
        const p = document.createElement('p');
        p.append(img);
        range.insertNode(p);
      } else {
        range.insertNode(img);
      }
      last = img;
      caretAfter(img);
      range = selectionRange();
    }
    return Boolean(last);
  };

  // ---- Page breaks ----------------------------------------------------------

  // A page break at the caret. In a top-level paragraph or heading it splits it there; anywhere
  // else (a list, table or text box) it goes after the top-level block. The caret goes after it.
  const insertPageBreak = () => {
    if (!selectionRange()) return false;
    if (!window.getSelection().isCollapsed) document.execCommand('delete');
    const range = selectionRange();
    const hr = document.createElement('hr');
    hr.className = 'wurd-page-break';
    const top = topLevelOf(range.startContainer);
    let next;
    if (!top) {
      editor.insertBefore(hr, editor.childNodes[range.startOffset] || null);
      next = hr.nextElementSibling;
    } else if (isElement(top) && /^(P|H[1-6])$/.test(top.tagName)) {
      const tail = document.createRange();
      tail.setStart(range.startContainer, range.startOffset);
      tail.setEnd(top, top.childNodes.length);
      let second = top.cloneNode(false);
      second.removeAttribute('id');
      second.append(tail.extractContents());
      if (isEmptyBlock(second)) {
        // A heading split at its end continues as a normal paragraph
        if (/^H/.test(top.tagName)) second = emptyParagraph();
        else second.replaceChildren(document.createElement('br'));
      }
      if (isEmptyBlock(top)) top.replaceWith(hr, second);
      else top.after(hr, second);
      next = second;
    } else {
      top.after(hr);
      next = hr.nextElementSibling;
    }
    if (!next || Wurd.isPageBreak(next)) {
      next = emptyParagraph();
      hr.after(next);
    }
    caretInto(next);
    return true;
  };

  // ---- Text boxes -----------------------------------------------------------

  const boxAt = (node) => {
    const el = isElement(node) ? node : node?.parentElement;
    const box = el?.closest?.('.wurd-textbox');
    return box && editor.contains(box) ? box : null;
  };

  // The top-level blocks a selection touches, first to last
  const topLevelBlocksIn = (range) => {
    const first = range.startContainer === editor ? editor.childNodes[range.startOffset] : topLevelOf(range.startContainer);
    let last = range.endContainer === editor ? editor.childNodes[range.endOffset - 1] : topLevelOf(range.endContainer);
    // A selection that ends at the very start of a block (a triple-click) doesn't include it
    if (last && last !== first && range.endContainer !== editor && range.endOffset === 0) last = last.previousSibling;
    const blocks = [];
    for (let node = first; node; node = node.nextSibling) {
      blocks.push(node);
      if (node === last) break;
    }
    return blocks;
  };

  // So there's always somewhere to type after a box
  const ensureParagraphAfter = (box) => {
    const next = box.nextElementSibling;
    if (!next || Wurd.isTextBox(next)) box.after(emptyParagraph());
  };

  // Puts the selected blocks in a text box, or takes the box the caret is in away.
  // Returns 'added', 'removed' or null.
  const toggleTextBox = () => {
    const range = selectionRange();
    if (!range) return null;
    const existing = boxAt(range.startContainer) || boxAt(range.endContainer);
    if (existing) {
      Wurd.formatting.keepSelection(() => existing.replaceWith(...existing.childNodes));
      return 'removed';
    }
    const box = document.createElement('div');
    box.className = 'wurd-textbox';
    const top = topLevelOf(range.startContainer);
    if (range.collapsed && !(isElement(top) && top.tagName === 'P' && isEmptyBlock(top))) {
      // A new, empty box after the paragraph the caret is in
      box.append(emptyParagraph());
      if (top) top.after(box);
      else editor.insertBefore(box, editor.childNodes[range.startOffset] || null);
      ensureParagraphAfter(box);
      caretInto(box);
      return 'added';
    }
    const blocks = topLevelBlocksIn(range).filter((node) => !(isElement(node) && Wurd.isPageBreak(node)));
    if (!blocks.length) return null;
    Wurd.formatting.keepSelection(() => {
      blocks[0].before(box);
      for (const block of blocks) {
        // Boxes don't nest: a box inside the selection gives up its contents
        if (isElement(block) && Wurd.isTextBox(block)) box.append(...block.childNodes);
        else box.append(block);
        if (Wurd.isTextBox(block)) block.remove();
      }
      ensureParagraphAfter(box);
    });
    return 'added';
  };

  const inTextBox = (node) => Boolean(boxAt(node));

  // ---- Links ----------------------------------------------------------------

  const linkAt = (node) => {
    const el = isElement(node) ? node : node?.parentElement;
    const link = el?.closest?.('a[href]');
    return link && editor.contains(link) ? link : null;
  };

  // What someone typed as a link address -> an href, or null if Wurd won't link to it.
  // Email addresses get mailto:, bare addresses get https://; other schemes (javascript:, file:...) are refused.
  const normalizeUrl = (input) => {
    const value = String(input || '').trim();
    // Spaces, and paths or anchors that only mean something on a web server
    if (!value || /\s/.test(value) || /^[#/.?]/.test(value)) return null;
    if (/^(https?:\/\/|mailto:|tel:)/i.test(value)) return /^https?:\/\/[^/?#]/i.test(value) || !/^https?:/i.test(value) ? value : null;
    if (/^[^@/:]+@[^@/:]+\.[^@/:]+$/.test(value)) return `mailto:${value}`;
    // host:port looks like a scheme but isn't one
    if (/^[a-z][a-z0-9+.-]*:(?!\d)/i.test(value)) return null;
    return `https://${value.replace(/^\/+/, '')}`;
  };

  // Whether a link may be opened from the editor
  const openable = (href) => /^(https?:|mailto:|tel:)/i.test(href || '');

  // Adds, changes or removes a link. `link` is the existing <a> being edited, if any.
  // With text, a new link at a caret shows that text; with a selection, the selection becomes the link.
  const setLink = ({ link = null, href = null, text = '' }) => {
    if (link) {
      if (!href) {
        Wurd.formatting.keepSelection(() => link.replaceWith(...link.childNodes));
        return true;
      }
      link.setAttribute('href', href);
      if (text && text !== link.textContent) link.textContent = text;
      caretAfter(link);
      return true;
    }
    const range = selectionRange();
    if (!range || !href) return false;
    if (!range.collapsed) {
      document.execCommand('createLink', false, href);
      return true;
    }
    const a = document.createElement('a');
    a.setAttribute('href', href);
    a.textContent = text || href.replace(/^mailto:|^tel:/i, '');
    range.insertNode(a);
    caretAfter(a);
    return true;
  };

  // ---- Special characters ---------------------------------------------------

  const COMBINING = { '̀': 'grave', '́': 'acute', '̂': 'circumflex', '̃': 'tilde', '̈': 'diaeresis', '̊': 'ring above', '̧': 'cedilla', '̌': 'caron' };
  // Accented letters are named from their Unicode decomposition ("Latin small letter e with acute")
  const latin = (chars, named = {}) => Array.from(chars).map((ch) => {
    if (named[ch]) return [ch, named[ch]];
    const [base, mark] = Array.from(ch.normalize('NFD'));
    const kind = base === base.toLowerCase() ? 'small' : 'capital';
    return [ch, `Latin ${kind} letter ${base.toLowerCase()}${mark && COMBINING[mark] ? ` with ${COMBINING[mark]}` : ''}`];
  });
  const greek = (chars, kind) => {
    const NAMES = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta', 'iota', 'kappa', 'lambda', 'mu', 'nu', 'xi', 'omicron', 'pi', 'rho', 'final sigma', 'sigma', 'tau', 'upsilon', 'phi', 'chi', 'psi', 'omega'];
    const start = kind === 'small' ? 0x3b1 : 0x391;
    return Array.from(chars).map((ch) => [ch, `Greek ${kind} letter ${NAMES[ch.codePointAt(0) - start]}`]);
  };

  const SYMBOLS = [
    { group: 'Common', chars: [
      ['©', 'Copyright sign'], ['®', 'Registered sign'], ['™', 'Trade mark sign'], ['§', 'Section sign'], ['¶', 'Pilcrow sign'],
      ['†', 'Dagger'], ['‡', 'Double dagger'], ['•', 'Bullet'], ['·', 'Middle dot'], ['…', 'Horizontal ellipsis'], ['°', 'Degree sign'],
      ['№', 'Numero sign'], ['℃', 'Degree Celsius'], ['‰', 'Per mille sign'], ['′', 'Prime'], ['″', 'Double prime'],
    ] },
    { group: 'Punctuation', chars: [
      ['–', 'En dash'], ['—', 'Em dash'], ['‘', 'Left single quotation mark'], ['’', 'Right single quotation mark'],
      ['“', 'Left double quotation mark'], ['”', 'Right double quotation mark'], ['‚', 'Single low-9 quotation mark'], ['„', 'Double low-9 quotation mark'],
      ['«', 'Left-pointing double angle quotation mark'], ['»', 'Right-pointing double angle quotation mark'],
      ['‹', 'Single left-pointing angle quotation mark'], ['›', 'Single right-pointing angle quotation mark'],
      ['¡', 'Inverted exclamation mark'], ['¿', 'Inverted question mark'], ['‑', 'Non-breaking hyphen'],
    ] },
    { group: 'Currency', chars: [
      ['€', 'Euro sign'], ['£', 'Pound sign'], ['¥', 'Yen sign'], ['¢', 'Cent sign'], ['$', 'Dollar sign'], ['₹', 'Indian rupee sign'],
      ['₽', 'Ruble sign'], ['₩', 'Won sign'], ['₪', 'New sheqel sign'], ['₺', 'Turkish lira sign'], ['₦', 'Naira sign'], ['₫', 'Dong sign'],
      ['฿', 'Thai currency symbol baht'], ['₿', 'Bitcoin sign'], ['¤', 'Currency sign'],
    ] },
    { group: 'Maths', chars: [
      ['±', 'Plus-minus sign'], ['×', 'Multiplication sign'], ['÷', 'Division sign'], ['−', 'Minus sign'], ['≠', 'Not equal to'],
      ['≈', 'Almost equal to'], ['≡', 'Identical to'], ['≤', 'Less-than or equal to'], ['≥', 'Greater-than or equal to'], ['∞', 'Infinity'],
      ['√', 'Square root'], ['∑', 'N-ary summation'], ['∏', 'N-ary product'], ['∫', 'Integral'], ['∂', 'Partial differential'],
      ['∆', 'Increment'], ['∇', 'Nabla'], ['∈', 'Element of'], ['∉', 'Not an element of'], ['∩', 'Intersection'], ['∪', 'Union'],
      ['⊂', 'Subset of'], ['∀', 'For all'], ['∃', 'There exists'], ['¬', 'Not sign'], ['∝', 'Proportional to'], ['µ', 'Micro sign'],
      ['½', 'Vulgar fraction one half'], ['⅓', 'Vulgar fraction one third'], ['⅔', 'Vulgar fraction two thirds'], ['¼', 'Vulgar fraction one quarter'],
      ['¾', 'Vulgar fraction three quarters'], ['⅛', 'Vulgar fraction one eighth'], ['¹', 'Superscript one'], ['²', 'Superscript two'],
      ['³', 'Superscript three'], ['₁', 'Subscript one'], ['₂', 'Subscript two'],
    ] },
    { group: 'Arrows', chars: [
      ['←', 'Leftwards arrow'], ['→', 'Rightwards arrow'], ['↑', 'Upwards arrow'], ['↓', 'Downwards arrow'], ['↔', 'Left right arrow'],
      ['↕', 'Up down arrow'], ['⇐', 'Leftwards double arrow'], ['⇒', 'Rightwards double arrow'], ['⇔', 'Left right double arrow'],
      ['↵', 'Downwards arrow with corner leftwards'], ['↺', 'Anticlockwise open circle arrow'], ['↻', 'Clockwise open circle arrow'],
    ] },
    { group: 'Greek', chars: [...greek('αβγδεζηθικλμνξπρστυφχψω', 'small'), ...greek('ΓΔΘΛΞΠΣΦΨΩ', 'capital')] },
    { group: 'Latin', chars: latin('àáâãäåæçèéêëìíîïñòóôõöøœùúûüýÿßðþÀÁÂÄÅÆÇÈÉÊËÌÍÎÏÑÒÓÔÖØŒÙÚÛÜÝ', {
      æ: 'Latin small letter ae', Æ: 'Latin capital letter AE', ø: 'Latin small letter o with stroke', Ø: 'Latin capital letter O with stroke',
      œ: 'Latin small ligature oe', Œ: 'Latin capital ligature OE', ß: 'Latin small letter sharp s', ð: 'Latin small letter eth', þ: 'Latin small letter thorn',
    }) },
    { group: 'Symbols', chars: [
      ['✓', 'Check mark'], ['✔', 'Heavy check mark'], ['✗', 'Ballot X'], ['✘', 'Heavy ballot X'], ['★', 'Black star'], ['☆', 'White star'],
      ['☐', 'Ballot box'], ['☑', 'Ballot box with check'], ['☒', 'Ballot box with X'], ['♠', 'Black spade suit'], ['♣', 'Black club suit'],
      ['♥', 'Black heart suit'], ['♦', 'Black diamond suit'], ['♪', 'Eighth note'], ['♫', 'Beamed eighth notes'], ['☺', 'White smiling face'],
      ['☎', 'Black telephone'], ['✉', 'Envelope'], ['⚠', 'Warning sign'], ['☀', 'Black sun with rays'], ['☁', 'Cloud'], ['♀', 'Female sign'],
      ['♂', 'Male sign'], ['■', 'Black square'], ['□', 'White square'], ['▲', 'Black up-pointing triangle'], ['▼', 'Black down-pointing triangle'],
      ['●', 'Black circle'], ['○', 'White circle'], ['◆', 'Black diamond'], ['◇', 'White diamond'], ['⌘', 'Place of interest sign'],
    ] },
  ];

  const insertText = (text) => {
    if (!selectionRange()) return false;
    document.execCommand('insertText', false, text);
    return true;
  };

  const init = (editorEl) => {
    editor = editorEl;
  };

  Wurd.insert = {
    init,
    isImageFile,
    imageFromFile,
    insertImages,
    insertPageBreak,
    toggleTextBox,
    inTextBox,
    linkAt,
    normalizeUrl,
    openable,
    setLink,
    insertText,
    caretAfter,
    ensureParagraphAfter,
    SYMBOLS,
  };
})(window.Wurd = window.Wurd || {});
