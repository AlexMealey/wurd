// Word .docx: import with mammoth, export with the docx library (both loaded on first use)
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});
  const NBSP = Wurd.NBSP;

  // ---- Import ---------------------------------------------------------------

  // Word paragraph alignment (w:jc) -> CSS text-align
  const ALIGN_IN = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };

  // [Word style name, HTML path prefix, tag]
  const BLOCK_STYLES = [
    ['heading 1', '', 'h1'], ['heading 2', '', 'h2'], ['heading 3', '', 'h3'],
    ['heading 4', '', 'h4'], ['heading 5', '', 'h4'], ['heading 6', '', 'h4'],
    ['Title', '', 'h1'], ['Subtitle', '', 'h2'],
    ['Quote', 'blockquote > ', 'p'], ['Intense Quote', 'blockquote > ', 'p'],
    // LibreOffice's names for its quotation style, older and newer
    ['Quotations', 'blockquote > ', 'p'], ['Block Quotation', 'blockquote > ', 'p'],
  ];
  const BLOCK_STYLE_NAMES = new Set(BLOCK_STYLES.map(([name]) => name.toLowerCase()));

  const ALIGNS = ['left', 'center', 'right', 'justify'];

  // The style attribute for a paragraph tag (alignment and indent steps)
  const paragraphAttr = (align, steps) => {
    const parts = [];
    if (align !== 'left') parts.push(`text-align: ${align}`);
    if (steps) parts.push(`margin-left: ${steps * Wurd.INDENT_PT}pt`);
    return `[style='${parts.join('; ')}']`;
  };

  // mammoth drops alignment and indentation, so such paragraphs are tagged with a style id
  // (see tagParagraphs) and these rules, which come first, carry them through as a style attribute.
  // runRules come from annotateRuns and do the same for run colours, sizes and fonts.
  const buildStyleMap = (runRules = []) => {
    const rules = [];
    for (const align of ALIGNS) {
      for (let steps = 0; steps <= Wurd.MAX_INDENT_STEPS; steps++) {
        if (align === 'left' && !steps) continue;
        const id = `WurdPara-${align}-${steps}`;
        const attr = paragraphAttr(align, steps);
        for (const [name, prefix, tag] of BLOCK_STYLES) {
          rules.push(`p.${id}[style-name='${name}'] => ${prefix}${tag}${attr}:fresh`);
        }
        rules.push(`p.${id} => p${attr}:fresh`);
      }
    }
    for (const [name, prefix, tag] of BLOCK_STYLES) rules.push(`p[style-name='${name}'] => ${prefix}${tag}:fresh`);
    rules.push(
      ...runRules,
      // Word's and LibreOffice's built-in character styles for emphasis
      "r[style-name='Strong'] => strong",
      "r[style-name='Strong Emphasis'] => strong",
      "r[style-name='Emphasis'] => em",
      // Page breaks come out inside paragraphs; tidyImported moves them between blocks
      "br[type='page'] => br.wurd-page-break",
      // Text boxes Wurd wrote (one-cell tables); tidyImported turns them back into boxes
      'table.WurdTextBox => table.wurd-textbox:fresh',
      "p[style-name='HTML Preformatted'] => pre:separator('\\n')",
      "p[style-name='Preformatted Text'] => pre:separator('\\n')",
      "p[style-name='Horizontal Line'] => hr:fresh",
      'u => u',
      'strike => s',
      "r[style-name='HTML Code'] => code",
      "r[style-name='Source Text'] => code",
    );
    return rules;
  };

  // Keeps tabs and runs of spaces visible, as the .txt import does
  const preserveRuns = (text) => text
    .replace(/\t/g, NBSP.repeat(4))
    .replace(/ {2,}/g, (run) => NBSP.repeat(run.length - 1) + ' ');

  // A page break inside a paragraph becomes one between paragraphs, splitting it if it's mid-text
  const hoistPageBreak = (root, br) => {
    const hr = document.createElement('hr');
    hr.className = 'wurd-page-break';
    let top = br;
    while (top.parentNode !== root) top = top.parentNode;
    if (top === br) {
      br.replaceWith(hr);
      return;
    }
    const empty = (el) => !el.textContent.trim() && !el.querySelector('img');
    if (/^(P|H[1-6])$/.test(top.tagName)) {
      const rest = document.createRange();
      rest.setStartAfter(br);
      rest.setEnd(top, top.childNodes.length);
      const after = top.cloneNode(false);
      after.append(rest.extractContents());
      br.remove();
      top.after(hr);
      if (!empty(after)) hr.after(after);
      if (empty(top)) top.remove();
      return;
    }
    // Inside a list or table: the break goes after it
    br.remove();
    top.after(hr);
  };

  const isEmpty = (el) => !el.textContent.trim() && !el.querySelector('img, hr, table, .wurd-textbox');

  const tidyImported = (html, boxes = []) => {
    const root = Wurd.documentRoot(html);
    root.querySelectorAll('br.wurd-page-break').forEach((br) => hoistPageBreak(root, br));
    // Text boxes: Wurd's in-line ones come back from one-cell tables, Word's from their paragraphs
    root.querySelectorAll('table.wurd-textbox, table[data-wurd-box]').forEach((table) => {
      const cells = table.querySelectorAll('td, th');
      if (cells.length !== 1) {
        table.removeAttribute('class');
        table.removeAttribute('data-wurd-box');
        return;
      }
      const box = document.createElement('div');
      box.className = 'wurd-textbox';
      box.append(...cells[0].childNodes);
      Wurd.setBoxGeometry(box, boxes[Number(table.dataset.wurdBox) - 1]);
      table.replaceWith(box);
    });
    root.querySelectorAll('div[data-wurd-box]').forEach((div) => {
      const info = boxes[Number(div.dataset.wurdBox) - 1] || {};
      const box = document.createElement('div');
      box.className = 'wurd-textbox';
      box.append(...div.childNodes);
      Wurd.setBoxGeometry(box, info);
      div.replaceWith(box);
      // mammoth puts a box's text after the paragraph it's anchored in; a free box goes before it,
      // and an anchor paragraph with nothing else in it (as Wurd writes them) goes
      const anchor = box.previousElementSibling;
      if (info.floating && anchor && box.parentNode === root && !Wurd.isTextBox(anchor)) {
        anchor.before(box);
        if (isEmpty(anchor)) anchor.remove();
      }
    });
    root.querySelectorAll('h5, h6').forEach((heading) => {
      const h4 = document.createElement('h4');
      if (heading.getAttribute('style')) h4.setAttribute('style', heading.getAttribute('style'));
      h4.append(...heading.childNodes);
      heading.replaceWith(h4);
    });
    root.querySelectorAll('img').forEach((img) => {
      if (!img.getAttribute('src')) {
        const placeholder = document.createElement('em');
        placeholder.textContent = '[image]';
        img.replaceWith(placeholder);
      }
    });
    // Line breaks inside preformatted text are newlines
    root.querySelectorAll('pre br').forEach((br) => br.replaceWith('\n'));
    Wurd.formats.unwrapHeaderBold(root);
    // Bookmark anchors mammoth leaves behind
    root.querySelectorAll('a[id]:not([href])').forEach((a) => a.replaceWith(...a.childNodes));
    // Empty paragraphs are blank lines in Word; keep them visible
    root.querySelectorAll('p').forEach((p) => {
      if (!p.textContent.trim() && !p.querySelector('img, br')) p.innerHTML = '<br>';
    });
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (/\t| {2}/.test(node.data) && !node.parentElement.closest('pre')) node.data = preserveRuns(node.data);
    }
    return root.innerHTML;
  };

  // Aligned or indented paragraphs get a style id the style map turns into a style attribute. The
  // style name is kept, so heading and quote rules still match; list items keep their numbering.
  const tagParagraphs = (document) => mammoth.transforms.paragraph((paragraph) => {
    // Paragraphs in a text box carry the box's id
    if (/^WurdBox-/.test(paragraph.styleId || '')) return paragraph;
    const align = ALIGN_IN[paragraph.alignment] || 'left';
    const steps = Math.min(Math.max(Math.round(Number(paragraph.indent?.start) / 720) || 0, 0), Wurd.MAX_INDENT_STEPS);
    // Numbered paragraphs are list items and must stay lists, except headings: LibreOffice's heading
    // styles carry (unnumbered) outline numbering
    const listItem = paragraph.numbering && !BLOCK_STYLE_NAMES.has((paragraph.styleName || '').toLowerCase());
    if (listItem || (align === 'left' && !steps)) return paragraph;
    return { ...paragraph, styleId: `WurdPara-${align}-${steps}` };
  })(document);

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const childW = (el, name) => Array.from(el.children).find((c) => c.namespaceURI === W_NS && c.localName === name);
  const attrW = (el, name) => el?.getAttributeNS(W_NS, name) || el?.getAttribute(`w:${name}`) || '';
  const HEX = /^[0-9a-f]{6}$/i;

  // A run's own font, size, colour and highlight as CSS ('' when it has none of them)
  const runCss = (rPr) => {
    const parts = [];
    const fonts = childW(rPr, 'rFonts');
    const fontName = attrW(fonts, 'ascii') || attrW(fonts, 'hAnsi');
    if (fontName) {
      const font = Wurd.classifyFont(fontName);
      if (font !== Wurd.DEFAULT_FONT) parts.push(`font-family: ${font}`);
    }
    const size = Number(attrW(childW(rPr, 'sz'), 'val'));
    if (size > 0 && size <= 3276) parts.push(`font-size: ${size / 2}pt`);
    const color = attrW(childW(rPr, 'color'), 'val');
    // Plain black is left as automatic, so it still shows on the dark theme's page
    if (HEX.test(color) && color !== '000000') parts.push(`color: #${color.toUpperCase()}`);
    const highlight = Wurd.HIGHLIGHTS[attrW(childW(rPr, 'highlight'), 'val')];
    const fill = attrW(childW(rPr, 'shd'), 'fill');
    if (highlight) parts.push(`background-color: #${highlight}`);
    else if (HEX.test(fill) && fill.toUpperCase() !== 'FFFFFF') parts.push(`background-color: #${fill.toUpperCase()}`);
    return parts.join('; ');
  };

  // mammoth doesn't read run colours, and keeps only one style per run, so before it runs each
  // run with its own font, size, colour or highlight gets a style id naming that combination.
  // Returns the file to convert and the style map rules that turn those ids into <span style>.
  // Runs that already have a character style keep it.
  const VML_NS = 'urn:schemas-microsoft-com:vml';
  const childNS = (el, ns, name) => Array.from(el.children).find((c) => c.namespaceURI === ns && c.localName === name);

  // Gives a paragraph a style id (creating its w:pPr and w:pStyle if need be)
  const setParagraphStyle = (xml, p, id) => {
    let pPr = childW(p, 'pPr');
    if (!pPr) {
      pPr = xml.createElementNS(W_NS, 'w:pPr');
      p.insertBefore(pPr, p.firstChild);
    }
    let pStyle = childW(pPr, 'pStyle');
    if (!pStyle) {
      pStyle = xml.createElementNS(W_NS, 'w:pStyle');
      pPr.insertBefore(pStyle, pPr.firstChild);
    }
    pStyle.setAttributeNS(W_NS, 'w:val', id);
  };

  const vmlStyle = (style) => Object.fromEntries(String(style || '').split(';')
    .map((part) => part.split(':').map((s) => s.trim().toLowerCase()))
    .filter(([key, value]) => key && value));
  const ptOrZero = (value) => {
    const pt = Wurd.cssLengthToPt(value);
    return Number.isFinite(pt) ? pt : 0;
  };

  // Text boxes: Word's (as the VML that mammoth reads, Word's own fallback for its text boxes) and
  // Wurd's in-line ones (one-cell WurdTextBox tables). Each gets a numbered style id so it comes out
  // as its own element, and its size and place are kept in `boxes` for tidyImported.
  const annotateBoxes = (xml, rules, boxes) => {
    const sections = xml.getElementsByTagNameNS(W_NS, 'sectPr');
    const section = sections[sections.length - 1];
    const twipsToPt = (el, name, fallback) => (Number(attrW(el, name)) / 20) || fallback;
    const pageWidthPt = twipsToPt(section && childW(section, 'pgSz'), 'w', 595.3);
    const margins = section && childW(section, 'pgMar');
    const marginLeftPt = twipsToPt(margins, 'left', 72);
    const textWidthPt = pageWidthPt - marginLeftPt - twipsToPt(margins, 'right', 72);
    for (const shape of Array.from(xml.getElementsByTagNameNS(VML_NS, 'shape'))) {
      const textbox = childNS(shape, VML_NS, 'textbox');
      const content = textbox && textbox.getElementsByTagNameNS(W_NS, 'txbxContent')[0];
      if (!content) continue;
      const s = vmlStyle(shape.getAttribute('style'));
      const widthPt = Wurd.cssLengthToPt(s.width);
      // stroked="f": no outline
      const info = { minHeightPt: Wurd.cssLengthToPt(s.height), plain: /^(f|false)$/i.test(shape.getAttribute('stroked') || '') };
      if (s.position === 'absolute') {
        // Left from the page edge; vertical offsets are taken from the anchor paragraph
        const fromPage = /^page/.test(s['mso-position-horizontal-relative'] || '');
        info.floating = true;
        info.leftPct = ((fromPage ? 0 : marginLeftPt) + ptOrZero(s['margin-left']) + ptOrZero(s.left)) / pageWidthPt * 100;
        info.topPt = ptOrZero(s['margin-top']) + ptOrZero(s.top);
        info.widthPct = widthPt / pageWidthPt * 100;
      } else {
        info.widthPct = widthPt / textWidthPt * 100;
      }
      boxes.push(info);
      const id = `WurdBox-${boxes.length}`;
      for (const p of Array.from(content.getElementsByTagNameNS(W_NS, 'p'))) setParagraphStyle(xml, p, id);
      rules.push(`p.${id} => div[data-wurd-box='${boxes.length}'] > p:fresh`);
    }
    for (const table of Array.from(xml.getElementsByTagNameNS(W_NS, 'tbl'))) {
      const tblPr = childW(table, 'tblPr');
      const style = tblPr && childW(tblPr, 'tblStyle');
      if (!style || attrW(style, 'val') !== 'WurdTextBox') continue;
      const width = childW(tblPr, 'tblW');
      const rowHeight = table.getElementsByTagNameNS(W_NS, 'trHeight')[0];
      boxes.push({
        widthPct: attrW(width, 'type') === 'dxa' ? (Number(attrW(width, 'w')) / 20) / textWidthPt * 100 : NaN,
        minHeightPt: Number(attrW(rowHeight, 'val')) / 20,
      });
      style.setAttributeNS(W_NS, 'w:val', `WurdTextBox-${boxes.length}`);
      rules.push(`table.WurdTextBox-${boxes.length} => table[data-wurd-box='${boxes.length}']:fresh`);
    }
  };

  const annotateRuns = async (bytes) => {
    const unchanged = { bytes, rules: [], boxes: [] };
    try {
      await Wurd.loadScript('lib/jszip.min.js');
      const zip = await JSZip.loadAsync(bytes);
      const part = zip.file('word/document.xml');
      if (!part) return unchanged;
      const xml = new DOMParser().parseFromString(await part.async('string'), 'application/xml');
      if (xml.getElementsByTagName('parsererror').length) return unchanged;
      const ids = new Map();
      for (const run of Array.from(xml.getElementsByTagNameNS(W_NS, 'r'))) {
        const rPr = childW(run, 'rPr');
        if (!rPr || childW(rPr, 'rStyle')) continue;
        const css = runCss(rPr);
        if (!css) continue;
        if (!ids.has(css)) ids.set(css, `WurdRun-${ids.size + 1}`);
        const rStyle = xml.createElementNS(W_NS, 'w:rStyle');
        rStyle.setAttributeNS(W_NS, 'w:val', ids.get(css));
        rPr.insertBefore(rStyle, rPr.firstChild);
      }
      const boxRules = [];
      const boxes = [];
      annotateBoxes(xml, boxRules, boxes);
      if (!ids.size && !boxes.length) return unchanged;
      zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
      return {
        bytes: await zip.generateAsync({ type: 'uint8array', compression: 'STORE' }),
        rules: [...boxRules, ...Array.from(ids, ([css, id]) => `r.${id} => span[style='${css}']`)],
        boxes,
      };
    } catch (err) {
      // mammoth reports anything really wrong with the file
      console.warn('[docx] Run formatting left out:', err);
      return unchanged;
    }
  };

  const importDocx = async (original) => {
    await Wurd.loadScript('lib/mammoth.browser.min.js');
    const { bytes, rules, boxes } = await annotateRuns(original);
    const options = {
      styleMap: buildStyleMap(rules),
      ignoreEmptyParagraphs: false,
      transformDocument: tagParagraphs,
      convertImage: mammoth.images.imgElement(async (image) => {
        if (!/^image\/(png|jpeg|gif|bmp|webp|svg\+xml)$/.test(image.contentType)) return { src: '' };
        return { src: `data:${image.contentType};base64,${await image.read('base64')}` };
      }),
    };
    let result;
    try {
      const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
      result = await mammoth.convertToHtml({ arrayBuffer: buffer }, options);
    } catch (err) {
      console.warn('[mammoth]', err);
      throw new Wurd.UserError("This file isn't a readable Word document.");
    }
    result.messages.forEach((m) => console.warn('[mammoth]', m.message));
    const problems = result.messages.filter((m) => m.type === 'error' || /image/i.test(m.message)).length;
    return {
      html: tidyImported(result.value, boxes),
      font: Wurd.DEFAULT_FONT,
      notice: problems ? `Loaded, but ${problems} item${problems === 1 ? '' : 's'} couldn't be converted` : undefined,
    };
  };

  // ---- Export ---------------------------------------------------------------

  // A4 with 1 inch margins: 9026 twips of text width; 1 CSS px = 15 twips
  // Page size from Settings (A4 or Letter) with 1in margins
  const MARGIN_TWIPS = 1440;
  const pageFor = (size) => {
    const page = Wurd.PAGE_SIZES[size] || Wurd.PAGE_SIZES.a4;
    return { width: page.widthTwips, height: page.heightTwips, margin: MARGIN_TWIPS, content: page.widthTwips - 2 * MARGIN_TWIPS };
  };
  // Word's highlight names by colour; other background colours become shading
  const HIGHLIGHT_NAMES = Object.fromEntries(Object.entries(Wurd.HIGHLIGHTS).map(([name, hex]) => [hex, name]));

  // Word can't fall back from fonts it doesn't have, so map to fonts every Word install has
  const WORD_FONTS = {
    inter: 'Arial', 'sans-serif': 'Arial', 'system-ui': 'Arial', 'ui-sans-serif': 'Arial',
    georgia: 'Georgia', serif: 'Times New Roman',
    'courier new': 'Courier New', courier: 'Courier New', monospace: 'Courier New',
  };
  const wordFont = (css) => {
    const name = Wurd.firstFontFamily(css);
    return WORD_FONTS[name.toLowerCase()] || name || 'Arial';
  };

  const BLOCK = /^(P|DIV|H[1-6]|UL|OL|LI|BLOCKQUOTE|PRE|HR|TABLE|SECTION|ARTICLE|HEADER|FOOTER|FIGURE|FIGCAPTION|ADDRESS|DL|DT|DD)$/;
  const FONT_SIZES_PT = [8, 10, 12, 14, 18, 24, 36]; // <font size="1..7">

  const toHex = Wurd.colorToHex;

  // A background colour: one of Word's highlights if it matches exactly, otherwise shading
  const setBackground = (fmt, color) => {
    const hex = toHex(color);
    if (!hex) return;
    fmt.highlight = HIGHLIGHT_NAMES[hex];
    fmt.shading = fmt.highlight ? undefined : hex;
  };

  // Inline CSS that matters for Word runs
  const applyStyle = (style, fmt) => {
    if (!style) return;
    if (style.fontWeight === 'bold' || Number(style.fontWeight) >= 600) fmt.bold = true;
    if (style.fontStyle === 'italic' || style.fontStyle === 'oblique') fmt.italics = true;
    const decoration = `${style.textDecorationLine || ''} ${style.textDecoration || ''}`;
    if (/underline/.test(decoration)) fmt.underline = true;
    if (/line-through/.test(decoration)) fmt.strike = true;
    if (style.fontFamily) fmt.font = wordFont(style.fontFamily);
    if (style.color) fmt.color = toHex(style.color) || fmt.color;
    if (style.backgroundColor) setBackground(fmt, style.backgroundColor);
    const size = /^([\d.]+)(px|pt)$/.exec(style.fontSize || '');
    if (size) fmt.size = Math.round(size[2] === 'pt' ? Number(size[1]) * 2 : Number(size[1]) * 1.5);
  };

  const alignOf = (el, D) => {
    const value = (el.style?.textAlign || el.getAttribute?.('align') || '').toLowerCase();
    return { center: D.AlignmentType.CENTER, right: D.AlignmentType.RIGHT, end: D.AlignmentType.RIGHT, justify: D.AlignmentType.JUSTIFIED }[value];
  };

  // Flattens inline content into runs: { kind: 'text'|'break'|'image', text, fmt, href, image }
  const collectInline = (node, ctx, fmt, runs, href) => {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.data) runs.push({ kind: 'text', text: Wurd.collapseSpaces(node.data), fmt, href });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const tag = node.tagName;
    if (tag === 'BR') {
      runs.push({ kind: 'break' });
      return;
    }
    if (tag === 'IMG') {
      const image = ctx.images.get(node);
      if (image) runs.push({ kind: 'image', image, href });
      else if (node.alt) runs.push({ kind: 'text', text: `[${node.alt}]`, fmt, href });
      return;
    }
    const next = { ...fmt };
    let link = href;
    switch (tag) {
      case 'B': case 'STRONG': next.bold = true; break;
      case 'I': case 'EM': case 'CITE': case 'DFN': case 'VAR': next.italics = true; break;
      case 'U': case 'INS': next.underline = true; break;
      case 'S': case 'STRIKE': case 'DEL': next.strike = true; break;
      case 'SUB': next.sub = true; break;
      case 'SUP': next.sup = true; break;
      case 'CODE': case 'KBD': case 'TT': case 'SAMP': next.code = true; break;
      case 'MARK': setBackground(next, 'yellow'); break;
      case 'FONT': {
        if (node.getAttribute('face')) next.font = wordFont(node.getAttribute('face'));
        if (node.getAttribute('color')) next.color = toHex(node.getAttribute('color')) || next.color;
        const size = Number(node.getAttribute('size'));
        if (size >= 1 && size <= 7) next.size = FONT_SIZES_PT[size - 1] * 2;
        break;
      }
      case 'A': {
        const target = node.getAttribute('href') || '';
        if (/^(https?:|mailto:|tel:)/i.test(target)) link = target;
        break;
      }
      default:
        break;
    }
    applyStyle(node.style, next);
    // A block inside inline content starts a new line
    if (BLOCK.test(tag) && runs.length) runs.push({ kind: 'break' });
    node.childNodes.forEach((child) => collectInline(child, ctx, next, runs, link));
  };

  // Collapses spaces across runs and trims them at line edges, the way HTML renders them
  const normalizeRuns = (runs) => {
    // contentEditable ends blocks with a placeholder <br> that isn't a real line
    if (runs.length && runs[runs.length - 1].kind === 'break') runs.pop();
    const trimEnd = (end) => {
      for (let i = end - 1; i >= 0 && runs[i].kind === 'text'; i--) {
        runs[i].text = runs[i].text.replace(/ +$/, '');
        if (runs[i].text) break;
      }
    };
    let atLineStart = true;
    let afterSpace = false;
    runs.forEach((run, i) => {
      if (run.kind === 'break') {
        trimEnd(i);
        atLineStart = true;
        afterSpace = false;
      } else if (run.kind === 'image') {
        atLineStart = false;
        afterSpace = false;
      } else {
        if (atLineStart || afterSpace) run.text = run.text.replace(/^ +/, '');
        if (run.text) {
          atLineStart = false;
          afterSpace = run.text.endsWith(' ');
        }
      }
    });
    trimEnd(runs.length);
    return runs.filter((run) => run.kind !== 'text' || run.text);
  };

  const makeRun = (run, D) => {
    if (run.kind === 'break') return new D.TextRun({ break: 1 });
    if (run.kind === 'image') return run.image;
    const f = run.fmt;
    const options = { text: run.text };
    if (f.bold) options.bold = true;
    if (f.italics) options.italics = true;
    if (f.underline) options.underline = {};
    if (f.strike) options.strike = true;
    if (f.sub) options.subScript = true;
    if (f.sup) options.superScript = true;
    if (f.font) options.font = f.font;
    if (f.size) options.size = f.size;
    if (f.color) options.color = f.color;
    if (f.highlight) options.highlight = f.highlight;
    else if (f.shading) options.shading = { type: D.ShadingType.CLEAR, color: 'auto', fill: f.shading };
    if (f.code) options.style = 'HTMLCode';
    else if (run.href) options.style = 'Hyperlink';
    return new D.TextRun(options);
  };

  // Runs -> paragraph children, grouping consecutive runs with the same link into one hyperlink
  const toChildren = (runs, D) => {
    const children = [];
    let link = null;
    let linkRuns = [];
    const flushLink = () => {
      if (link) children.push(new D.ExternalHyperlink({ link, children: linkRuns }));
      link = null;
      linkRuns = [];
    };
    for (const run of runs) {
      const docRun = makeRun(run, D);
      if (run.href) {
        if (run.href !== link) flushLink();
        link = run.href;
        linkRuns.push(docRun);
      } else {
        flushLink();
        children.push(docRun);
      }
    }
    flushLink();
    return children;
  };

  // A paragraph from inline nodes. Loose whitespace between blocks makes no paragraph; an empty <p> does.
  const paragraphFrom = (nodes, ctx, state, options = {}) => {
    const { D } = ctx;
    const raw = [];
    nodes.forEach((node) => collectInline(node, ctx, state.fmt, raw, null));
    const runs = normalizeRuns(raw);
    if (!runs.length && !options.keepEmpty) return null;
    const props = { children: toChildren(runs, D) };
    const alignment = options.align ?? state.align;
    if (alignment) props.alignment = alignment;
    if (options.heading) props.heading = options.heading;
    if (options.numbering) props.numbering = options.numbering;
    if (options.indent) props.indent = options.indent;
    if (state.quote) {
      props.style = 'Quote';
      if (state.quote > 1) props.indent = { left: 720 * state.quote };
    }
    // margin-left on the paragraph (or a container around it), on top of any quote indent
    const blockIndent = (state.indent || 0) + (options.blockIndent || 0);
    if (blockIndent) props.indent = { ...props.indent, left: (props.indent?.left ?? (state.quote ? 720 : 0)) + blockIndent };
    return new D.Paragraph(props);
  };

  const preParagraph = (el, ctx) => {
    const { D } = ctx;
    const children = [];
    el.textContent.replace(/\n$/, '').split('\n').forEach((line, i) => {
      if (i > 0) children.push(new D.TextRun({ break: 1 }));
      line.split('\t').forEach((part, j) => {
        if (j > 0) children.push(new D.TextRun({ children: [new D.Tab()] }));
        if (part) children.push(new D.TextRun({ text: part }));
      });
    });
    return new D.Paragraph({ style: 'HTMLPreformatted', children });
  };

  const listBlocks = (list, ctx, state) => {
    const level = state.list ? Math.min(state.list.level + 1, 8) : 0;
    // A new instance restarts numbering for every list
    const instance = ++ctx.listInstance;
    const reference = list.tagName === 'OL' ? 'wurd-number' : 'wurd-bullet';
    // Items are indented by their numbering level, not by margins around the list
    const itemState = { ...state, list: { level, instance }, indent: 0 };
    const out = [];
    for (const child of list.children) {
      // execCommand can nest a list directly inside a list rather than inside an item
      if (child.tagName === 'UL' || child.tagName === 'OL') {
        out.push(...listBlocks(child, ctx, itemState));
        continue;
      }
      if (child.tagName !== 'LI') continue;
      let numbered = false;
      let inline = [];
      const emit = (nodes, extra = {}) => {
        const paragraph = paragraphFrom(nodes, ctx, itemState, {
          ...extra,
          keepEmpty: !numbered,
          numbering: numbered ? undefined : { reference, level, instance },
          indent: numbered ? { left: 720 * (level + 1) } : undefined,
        });
        if (paragraph) {
          out.push(paragraph);
          numbered = true;
        }
      };
      const flush = () => {
        if (inline.length) emit(inline);
        inline = [];
      };
      for (const node of child.childNodes) {
        if (node.nodeType === Node.ELEMENT_NODE && (node.tagName === 'UL' || node.tagName === 'OL')) {
          flush();
          if (!numbered) emit([]);
          out.push(...listBlocks(node, ctx, itemState));
        } else if (node.nodeType === Node.ELEMENT_NODE && /^(P|DIV|H[1-6])$/.test(node.tagName)) {
          flush();
          emit(Array.from(node.childNodes), { align: alignOf(node, ctx.D) });
        } else if (node.nodeType === Node.ELEMENT_NODE && BLOCK.test(node.tagName)) {
          flush();
          if (!numbered) emit([]);
          out.push(...blockFrom(node, ctx, itemState));
        } else {
          inline.push(node);
        }
      }
      flush();
      if (!numbered) emit([]);
    }
    return out;
  };

  const tableFrom = (table, ctx, state) => {
    const { D } = ctx;
    const rows = Array.from(table.rows).filter((row) => row.cells.length);
    if (!rows.length) return [];
    const columns = Math.max(...rows.map((row) => Array.from(row.cells).reduce((n, cell) => n + (cell.colSpan || 1), 0)));
    const columnWidth = Math.floor((state.width || ctx.contentTwips) / columns);
    const docRows = rows.map((row) => {
      const header = row.parentElement?.tagName === 'THEAD' || Array.from(row.cells).every((cell) => cell.tagName === 'TH');
      // Only set when true: docx writes <w:tblHeader w:val="off"/> for false, which readers treat as a header row
      return new D.TableRow({
        ...(header ? { tableHeader: true } : {}),
        children: Array.from(row.cells).map((cell) => {
          const span = cell.colSpan || 1;
          const cellState = {
            ...state,
            list: null,
            quote: 0,
            indent: 0,
            align: alignOf(cell, D),
            width: columnWidth * span,
            fmt: cell.tagName === 'TH' ? { ...state.fmt, bold: true } : state.fmt,
          };
          const children = blocksFrom(cell, ctx, cellState);
          // Word requires every cell to end with a paragraph
          if (!children.length || children[children.length - 1] instanceof D.Table) children.push(new D.Paragraph(''));
          const options = { children, width: { size: columnWidth * span, type: D.WidthType.DXA } };
          if (span > 1) options.columnSpan = span;
          if (cell.rowSpan > 1) options.rowSpan = cell.rowSpan;
          return new D.TableCell(options);
        }),
      });
    });
    // Explicit column widths: LibreOffice squashes the columns without them
    return [new D.Table({
      rows: docRows,
      width: { size: columnWidth * columns, type: D.WidthType.DXA },
      columnWidths: Array(columns).fill(columnWidth),
    })];
  };

  const HEADINGS = { H1: 'HEADING_1', H2: 'HEADING_2', H3: 'HEADING_3', H4: 'HEADING_4', H5: 'HEADING_4', H6: 'HEADING_4' };

  const percentOf = (value, fallback) => {
    const pct = parseFloat(value);
    return Number.isFinite(pct) && /%$/.test(value || '') ? Math.min(Math.max(pct, 1), 100) / 100 : fallback;
  };
  const ptOf = (value) => {
    const pt = Wurd.cssLengthToPt(value);
    return Number.isFinite(pt) ? pt : 0;
  };

  // A free text box is a Word text box in front of the text, positioned from the page's left edge
  // and from its anchor: an exactly 1pt paragraph where the box sits among the blocks.
  const floatingBoxFrom = (box, ctx, state) => {
    const { D } = ctx;
    const pageWidthPt = ctx.pageWidthTwips / 20;
    const widthPt = Math.round(percentOf(box.style.width, 0.4) * pageWidthPt * 10) / 10;
    const leftPt = Math.round(percentOf(box.style.left, 0) * pageWidthPt * 10) / 10;
    const children = blocksFrom(box, ctx, { ...state, list: null, quote: 0, indent: 0, width: Math.round(widthPt * 20) - 288 });
    if (!children.length) children.push(new D.Paragraph(''));
    ctx.floatCount = (ctx.floatCount || 0) + 1;
    return [new D.Textbox({
      children,
      style: {
        position: 'absolute',
        marginLeft: `${leftPt}pt`,
        marginTop: `${ptOf(box.style.marginTop)}pt`,
        width: `${widthPt}pt`,
        height: `${Math.max(ptOf(box.style.minHeight), 18)}pt`,
        positionHorizontalRelative: 'page',
        positionVerticalRelative: 'text',
        zIndex: 251659264 + ctx.floatCount,
      },
      spacing: { before: 0, after: 0, line: 20, lineRule: D.LineRuleType.EXACT },
    })];
  };

  // An in-line text box is a one-cell table with a border. Its style id marks it, so it imports back as a text box.
  const textBoxFrom = (box, ctx, state) => {
    const { D } = ctx;
    if (Wurd.isFloatingBox(box)) return floatingBoxFrom(box, ctx, state);
    const width = Math.round((state.width || ctx.contentTwips) * percentOf(box.style.width, 1));
    const minHeight = Math.round(ptOf(box.style.minHeight) * 20);
    const plain = box.classList.contains('wurd-textbox--plain');
    const none = { style: D.BorderStyle.NONE, size: 0, color: 'auto' };
    const border = plain ? none : { style: D.BorderStyle.SINGLE, size: 6, color: '808080' };
    const children = blocksFrom(box, ctx, { ...state, list: null, quote: 0, indent: 0, width: width - 288 });
    // Word requires every cell to end with a paragraph
    if (!children.length || children[children.length - 1] instanceof D.Table) children.push(new D.Paragraph(''));
    return [new D.Table({
      style: 'WurdTextBox',
      width: { size: width, type: D.WidthType.DXA },
      columnWidths: [width],
      borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: none, insideVertical: none },
      margins: { top: 115, bottom: 115, left: 144, right: 144 },
      rows: [new D.TableRow({
        ...(minHeight ? { height: { value: minHeight, rule: D.HeightRule.ATLEAST } } : {}),
        children: [new D.TableCell({ children, width: { size: width, type: D.WidthType.DXA } })],
      })],
    })];
  };

  // margin-left in twips
  const indentTwips = (el) => Math.round(Wurd.indentOf(el) * 20);

  const blockFrom = (el, ctx, state) => {
    const { D } = ctx;
    const tag = el.tagName;
    const align = alignOf(el, D) ?? state.align;
    if (HEADINGS[tag]) {
      const paragraph = paragraphFrom(Array.from(el.childNodes), ctx, { ...state, quote: 0 }, { heading: D.HeadingLevel[HEADINGS[tag]], align, keepEmpty: true, blockIndent: indentTwips(el) });
      return paragraph ? [paragraph] : [];
    }
    if (tag === 'UL' || tag === 'OL') return listBlocks(el, ctx, state);
    if (tag === 'BLOCKQUOTE') return blocksFrom(el, ctx, { ...state, quote: state.quote + 1, align });
    if (tag === 'PRE') return [preParagraph(el, ctx)];
    if (tag === 'HR') {
      if (Wurd.isPageBreak(el)) return [new D.Paragraph({ children: [new D.PageBreak()] })];
      return [new D.Paragraph({ style: 'HorizontalLine', children: [] })];
    }
    if (tag === 'TABLE') return tableFrom(el, ctx, state);
    if (Wurd.isTextBox(el)) return textBoxFrom(el, ctx, state);
    // P, DIV, LI outside a list and the like: a container if it holds blocks, otherwise one paragraph
    const holdsBlocks = Array.from(el.children).some((child) => BLOCK.test(child.tagName));
    if (holdsBlocks) return blocksFrom(el, ctx, { ...state, align, indent: (state.indent || 0) + indentTwips(el) });
    const paragraph = paragraphFrom(Array.from(el.childNodes), ctx, state, { align, keepEmpty: tag === 'P', blockIndent: indentTwips(el) });
    return paragraph ? [paragraph] : [];
  };

  function blocksFrom(container, ctx, state) {
    const out = [];
    let inline = [];
    const flush = () => {
      if (inline.length) {
        const paragraph = paragraphFrom(inline, ctx, state);
        if (paragraph) out.push(paragraph);
      }
      inline = [];
    };
    for (const node of container.childNodes) {
      if (node.nodeType === Node.ELEMENT_NODE && BLOCK.test(node.tagName)) {
        flush();
        const blocks = blockFrom(node, ctx, state);
        // Word joins tables that touch into one, so keep a paragraph between them
        if (blocks[0] instanceof ctx.D.Table && out[out.length - 1] instanceof ctx.D.Table) out.push(new ctx.D.Paragraph(''));
        out.push(...blocks);
      } else if (node.nodeType === Node.TEXT_NODE || node.nodeType === Node.ELEMENT_NODE) {
        inline.push(node);
      }
    }
    flush();
    return out;
  }

  // Images must be decoded before the (synchronous) walk. Only embedded data: images can be exported offline.
  const prepareImages = async (root, D, maxImagePx) => {
    const images = new Map();
    for (const img of root.querySelectorAll('img')) {
      const match = /^data:image\/(png|jpe?g|gif|bmp|webp|svg\+xml);base64,(.*)$/i.exec(img.getAttribute('src') || '');
      if (!match) continue;
      try {
        const decoded = new Image();
        decoded.src = img.getAttribute('src');
        await decoded.decode();
        let width = Number(img.getAttribute('width')) || decoded.naturalWidth;
        let height = Number(img.getAttribute('height')) || decoded.naturalHeight;
        if (!width || !height) continue;
        if (width > maxImagePx) {
          height = Math.round((height * maxImagePx) / width);
          width = maxImagePx;
        }
        let type = match[1].toLowerCase().replace('jpeg', 'jpg');
        let data;
        if (type === 'webp' || type === 'svg+xml') {
          // Word can't show these, so redraw them as PNG
          const canvas = document.createElement('canvas');
          canvas.width = decoded.naturalWidth || width;
          canvas.height = decoded.naturalHeight || height;
          canvas.getContext('2d').drawImage(decoded, 0, 0, canvas.width, canvas.height);
          data = Wurd.base64ToBytes(canvas.toDataURL('image/png').split(',')[1]);
          type = 'png';
        } else {
          data = Wurd.base64ToBytes(match[2]);
        }
        const alt = img.getAttribute('alt') || '';
        images.set(img, new D.ImageRun({ type, data, transformation: { width, height }, altText: { name: 'image', description: alt, title: alt } }));
      } catch (err) {
        console.warn('Image left out of the .docx:', err);
      }
    }
    return images;
  };

  const exportDocx = async (root, { title, pageFont, pageSize }) => {
    await Wurd.loadScript('lib/docx.iife.js');
    const D = window.docx;
    const page = pageFor(pageSize);
    const ctx = { D, images: await prepareImages(root, D, Math.floor(page.content / 15)), listInstance: 0, contentTwips: page.content, pageWidthTwips: page.width };
    const children = blocksFrom(root, ctx, { list: null, quote: 0, align: undefined, fmt: {}, indent: 0 });
    if (!children.length || children[children.length - 1] instanceof D.Table) children.push(new D.Paragraph(''));

    const levels = (format, text) => Array.from({ length: 9 }, (_, level) => ({
      level,
      format,
      text: text(level),
      alignment: D.AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
    }));

    const doc = new D.Document({
      creator: 'Wurd',
      lastModifiedBy: 'Wurd',
      title,
      styles: {
        default: {
          document: { run: { font: wordFont(pageFont), size: 24 }, paragraph: { spacing: { after: 160 } } },
          // outlineLevel puts headings in LibreOffice's navigator and tables of contents (Word goes by style)
          heading1: { run: { size: 54, bold: true }, paragraph: { spacing: { after: 240 }, keepNext: true, outlineLevel: 0 } },
          heading2: { run: { size: 36, bold: true }, paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 1 } },
          heading3: { run: { size: 30, bold: true }, paragraph: { spacing: { before: 280, after: 120 }, keepNext: true, outlineLevel: 2 } },
          heading4: { run: { size: 24, bold: true }, paragraph: { spacing: { before: 240, after: 80 }, keepNext: true, outlineLevel: 3 } },
        },
        // Names match the mammoth style map above, so an exported file imports back the same way
        paragraphStyles: [
          {
            id: 'Quote', name: 'Quote', basedOn: 'Normal', next: 'Normal',
            run: { italics: true, color: '4B5563' },
            paragraph: { indent: { left: 720 }, border: { left: { style: D.BorderStyle.SINGLE, size: 18, color: 'D1D5DB', space: 8 } } },
          },
          {
            id: 'HTMLPreformatted', name: 'HTML Preformatted', basedOn: 'Normal', next: 'Normal',
            run: { font: 'Courier New', size: 20 },
            paragraph: { shading: { type: D.ShadingType.CLEAR, color: 'auto', fill: 'F3F4F6' } },
          },
          {
            id: 'HorizontalLine', name: 'Horizontal Line', basedOn: 'Normal', next: 'Normal',
            paragraph: { border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: 'auto', space: 1 } } },
          },
        ],
        characterStyles: [{ id: 'HTMLCode', name: 'HTML Code', run: { font: 'Courier New', size: 20 } }],
      },
      numbering: {
        config: [
          { reference: 'wurd-bullet', levels: levels(D.LevelFormat.BULLET, (level) => ['•', '◦', '▪'][level % 3]) },
          { reference: 'wurd-number', levels: levels(D.LevelFormat.DECIMAL, (level) => `%${level + 1}.`) },
        ],
      },
      sections: [{
        properties: {
          page: {
            size: { width: page.width, height: page.height },
            margin: { top: page.margin, right: page.margin, bottom: page.margin, left: page.margin },
          },
        },
        children,
      }],
    });
    return D.Packer.toBlob(doc);
  };

  F.docx = { import: importDocx, export: exportDocx, wordFont };
})(window.Wurd = window.Wurd || {});
