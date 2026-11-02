// OpenDocument text (.odt): converts content.xml/styles.xml to editor HTML and back.
// The zip, manifest and password encryption live in odf-package.js.
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});
  const NBSP = Wurd.NBSP;

  const NS = {
    office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    xlink: 'http://www.w3.org/1999/xlink',
    dc: 'http://purl.org/dc/elements/1.1/',
    meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  };

  const is = (node, ns, localName) => node && node.nodeType === Node.ELEMENT_NODE && node.namespaceURI === NS[ns] && node.localName === localName;
  const at = (el, ns, name) => (el ? el.getAttributeNS(NS[ns], name) : null);
  const elementsOf = (el) => (el ? Array.from(el.children) : []);
  const findChild = (el, ns, localName) => elementsOf(el).find((c) => is(c, ns, localName));
  const unquote = (name) => (name || '').trim().replace(/^['"]|['"]$/g, '');

  const parseXml = (xml) => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Wurd.UserError('This OpenDocument file is damaged.');
    return doc;
  };

  // ---- Import: styles -------------------------------------------------------

  const DIRECT_PROPS = ['color', 'bg', 'size', 'indent', 'breakBefore', 'breakAfter'];

  const buildStyleIndex = (contentDoc, stylesDoc) => {
    const families = new Map();
    const listStyles = new Map();
    const defaults = new Map();
    const fontFaces = new Map();
    // Automatic styles are direct formatting; only they carry colours, sizes, indents and page breaks
    // through (a named style's are its design, like a heading's size)
    const automatic = new Set();
    const officeChild = (doc, name) => doc && findChild(doc.documentElement, 'office', name);

    const collect = (container, isAutomatic = false) => {
      for (const el of elementsOf(container)) {
        if (isAutomatic) automatic.add(el);
        if (is(el, 'style', 'style')) {
          const family = at(el, 'style', 'family');
          if (!families.has(family)) families.set(family, new Map());
          const byName = families.get(family);
          if (!byName.has(at(el, 'style', 'name'))) byName.set(at(el, 'style', 'name'), el);
        } else if (is(el, 'text', 'list-style')) {
          if (!listStyles.has(at(el, 'style', 'name'))) listStyles.set(at(el, 'style', 'name'), el);
        } else if (is(el, 'style', 'default-style')) {
          if (!defaults.has(at(el, 'style', 'family'))) defaults.set(at(el, 'style', 'family'), el);
        }
      }
    };
    // Lookup order: the document's automatic styles, then named styles, then styles.xml's automatic styles
    collect(officeChild(contentDoc, 'automatic-styles'), true);
    collect(officeChild(stylesDoc, 'styles'));
    collect(officeChild(stylesDoc, 'automatic-styles'), true);
    for (const doc of [contentDoc, stylesDoc]) {
      for (const face of elementsOf(officeChild(doc, 'font-face-decls'))) {
        if (!is(face, 'style', 'font-face')) continue;
        fontFaces.set(at(face, 'style', 'name'), {
          family: unquote(at(face, 'svg', 'font-family') || at(face, 'style', 'name')),
          generic: at(face, 'style', 'font-family-generic'),
          pitch: at(face, 'style', 'font-pitch'),
        });
      }
    }

    const ownProps = (el) => {
      const props = {};
      const tp = findChild(el, 'style', 'text-properties');
      if (tp) {
        const weight = at(tp, 'fo', 'font-weight');
        if (weight) props.b = weight === 'bold' || Number(weight) >= 600;
        const fontStyle = at(tp, 'fo', 'font-style');
        if (fontStyle) props.i = fontStyle === 'italic' || fontStyle === 'oblique';
        const underline = at(tp, 'style', 'text-underline-style');
        if (underline) props.u = underline !== 'none';
        const strikeStyle = at(tp, 'style', 'text-line-through-style');
        const strikeType = at(tp, 'style', 'text-line-through-type');
        if (strikeStyle || strikeType) props.s = Boolean((strikeStyle && strikeStyle !== 'none') || (strikeType && strikeType !== 'none'));
        const fontName = at(tp, 'style', 'font-name');
        const fontFamily = at(tp, 'fo', 'font-family');
        if (fontName && fontFaces.has(fontName)) {
          const face = fontFaces.get(fontName);
          props.font = Wurd.classifyFont(face.family, face);
        } else if (fontName || fontFamily) {
          props.font = Wurd.classifyFont(unquote(fontFamily || fontName));
        }
        // Black counts as automatic, so it still shows on the dark theme's page; '' overrides a parent's colour
        const color = at(tp, 'fo', 'color');
        if (color) props.color = /^#[0-9a-f]{6}$/i.test(color) && color !== '#000000' ? color.slice(1).toUpperCase() : '';
        const background = at(tp, 'fo', 'background-color');
        if (background) props.bg = /^#[0-9a-f]{6}$/i.test(background) ? background.slice(1).toUpperCase() : '';
        const size = Wurd.cssLengthToPt(at(tp, 'fo', 'font-size'));
        if (size > 0) props.size = Math.round(size * 2) / 2;
      }
      const pp = findChild(el, 'style', 'paragraph-properties');
      const align = pp && at(pp, 'fo', 'text-align');
      if (align) props.align = { center: 'center', end: 'right', right: 'right', justify: 'justify' }[align] || '';
      const margin = pp && (at(pp, 'fo', 'margin-left') || at(pp, 'fo', 'margin-start'));
      if (margin) {
        const indent = Wurd.cssLengthToPt(margin);
        if (Number.isFinite(indent)) props.indent = indent >= 1 ? Math.round(indent) : 0;
      }
      const breakBefore = pp && at(pp, 'fo', 'break-before');
      if (breakBefore) props.breakBefore = breakBefore === 'page';
      const breakAfter = pp && at(pp, 'fo', 'break-after');
      if (breakAfter) props.breakAfter = breakAfter === 'page';
      const listStyle = at(el, 'style', 'list-style-name');
      if (listStyle) props.listStyle = listStyle;
      // For text boxes: where a frame is measured from, and a box table's width and row height
      const gp = findChild(el, 'style', 'graphic-properties');
      if (gp && at(gp, 'style', 'horizontal-rel')) props.hRel = at(gp, 'style', 'horizontal-rel');
      if (gp && at(gp, 'draw', 'stroke')) props.stroke = at(gp, 'draw', 'stroke');
      const tableProps = findChild(el, 'style', 'table-properties');
      if (tableProps) {
        props.tableWidth = at(tableProps, 'style', 'width');
        props.tableAlign = at(tableProps, 'table', 'align');
      }
      const rowProps = findChild(el, 'style', 'table-row-properties');
      if (rowProps) props.minRowHeight = at(rowProps, 'style', 'min-row-height');
      return props;
    };

    const cache = new Map();
    // Resolved properties of a style, following parent-style-name; chain lists the style names involved
    const resolve = (family, name, depth = 0) => {
      const key = `${family}|${name}`;
      if (cache.has(key)) return cache.get(key);
      const el = name ? families.get(family)?.get(name) : null;
      let props;
      if (!el || depth > 20) {
        props = { ...(defaults.has(family) && !name ? ownProps(defaults.get(family)) : {}), chain: [] };
        if (name && !el) props = { ...resolve(family, null, depth + 1), chain: [name] };
      } else {
        const parent = resolve(family, at(el, 'style', 'parent-style-name'), depth + 1);
        const own = ownProps(el);
        if (!automatic.has(el)) DIRECT_PROPS.forEach((k) => delete own[k]);
        props = { ...parent, ...own, chain: [name, ...parent.chain] };
      }
      cache.set(key, props);
      return props;
    };

    // ul or ol for a list style at a nesting level
    const listKind = (styleName, level) => {
      const style = listStyles.get(styleName);
      if (!style) return 'ul';
      let best = null;
      for (const levelStyle of elementsOf(style)) {
        const n = Number(at(levelStyle, 'text', 'level')) || 1;
        if (n <= level && (!best || n > best.n)) best = { n, el: levelStyle };
      }
      if (!best) return 'ul';
      return is(best.el, 'text', 'list-level-style-number') && at(best.el, 'style', 'num-format') ? 'ol' : 'ul';
    };

    const standard = resolve('paragraph', 'Standard');
    // Drawings keep their font on the default graphic style
    const pageFont = standard.font || resolve('paragraph', null).font || resolve('graphic', null).font || Wurd.DEFAULT_FONT;
    return { resolve, listKind, pageFont };
  };

  // What a paragraph style means, from its own and its parents' names
  const semantic = (props) => {
    for (const name of props.chain) {
      if (/^Heading_20_\d+$/.test(name) || name === 'Heading') return 'heading';
      if (name === 'Title') return 'title';
      if (name === 'Quotations' || name === 'Quote') return 'quote';
      if (name === 'Preformatted_20_Text') return 'pre';
      if (name === 'Horizontal_20_Line') return 'hr';
    }
    return null;
  };

  // ---- Import: content --------------------------------------------------------

  const lengthToPx = (value) => {
    const m = /^([\d.]+)(cm|mm|in|pt|pc|px)$/.exec(value || '');
    if (!m) return null;
    const n = Number(m[1]);
    return Math.round({ cm: n * 96 / 2.54, mm: n * 96 / 25.4, in: n * 96, pt: n * 96 / 72, pc: n * 16, px: n }[m[2]]);
  };

  const imageType = (bytes) => {
    if (bytes[0] === 0x89 && bytes[1] === 0x50) return 'image/png';
    if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg';
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'image/gif';
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return 'image/bmp';
    if (bytes[0] === 0x52 && bytes[8] === 0x57 && bytes[9] === 0x45) return 'image/webp';
    if (/<svg[\s>]/.test(new TextDecoder().decode(bytes.subarray(0, 512)))) return 'image/svg+xml';
    return null; // EMF, WMF, SVM and the like can't be shown in a browser
  };

  // Images are read up front so the conversion itself can be synchronous
  const loadImages = async (contentDoc, pkg) => {
    const images = new Map();
    for (const image of contentDoc.getElementsByTagNameNS(NS.draw, 'image')) {
      const href = (at(image, 'xlink', 'href') || '').replace(/^\.\//, '');
      let bytes = null;
      if (href && !/^[a-z]+:/i.test(href)) bytes = await pkg.readBytes(href);
      const binary = findChild(image, 'office', 'binary-data');
      if (!bytes && binary) bytes = Wurd.base64ToBytes(binary.textContent.replace(/\s+/g, ''));
      const type = bytes && imageType(bytes);
      if (type) images.set(image, `data:${type};base64,${Wurd.bytesToBase64(bytes)}`);
    }
    return images;
  };

  const pt = (value) => {
    const n = Wurd.cssLengthToPt(value);
    return Number.isFinite(n) ? n : 0;
  };

  // A text box frame's size and place. Frames anchored to a paragraph (or character) with a position
  // are free boxes; the rest sit in line with the text.
  const frameBox = (frame, textBox, ctx) => {
    const page = ctx.page;
    const widthPt = pt(at(frame, 'svg', 'width'));
    const minHeightPt = pt(at(textBox, 'fo', 'min-height') || at(frame, 'fo', 'min-height') || at(frame, 'svg', 'height'));
    const anchor = at(frame, 'text', 'anchor-type');
    const positioned = at(frame, 'svg', 'x') || at(frame, 'svg', 'y');
    if (!positioned || !/^(paragraph|char)$/.test(anchor || '')) return { widthPct: (widthPt / page.textWidthPt) * 100, minHeightPt };
    const graphic = ctx.styles.resolve('graphic', at(frame, 'draw', 'style-name'));
    const fromPage = /^page/.test(graphic.hRel || '');
    return {
      floating: true,
      leftPct: (((fromPage ? 0 : page.marginLeftPt) + pt(at(frame, 'svg', 'x'))) / page.widthPt) * 100,
      topPt: pt(at(frame, 'svg', 'y')),
      widthPct: (widthPt / page.widthPt) * 100,
      minHeightPt,
    };
  };

  // Text nodes collapse whitespace (ODF 1.2 part 1, 6.1.2); text:s and text:tab carry real spaces
  const collectRuns = (node, props, ctx, runs, href) => {
    for (const child of node.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        runs.push({ text: child.data.replace(/[ \t\r\n]+/g, ' '), props, href });
      } else if (child.nodeType !== Node.ELEMENT_NODE) {
        continue;
      } else if (is(child, 'text', 's')) {
        runs.push({ text: NBSP.repeat(Math.min(Number(at(child, 'text', 'c')) || 1, 1000)), props, href, keep: true });
      } else if (is(child, 'text', 'tab')) {
        runs.push({ text: NBSP.repeat(4), props, href, keep: true });
      } else if (is(child, 'text', 'line-break')) {
        runs.push({ br: true });
      } else if (is(child, 'text', 'span')) {
        const span = ctx.styles.resolve('text', at(child, 'text', 'style-name'));
        const code = span.chain.some((n) => n === 'Source_20_Text' || n === 'Teletype');
        collectRuns(child, { ...props, ...strip(span), code: props.code || code }, ctx, runs, href);
      } else if (is(child, 'text', 'a')) {
        // The link's own style (Internet link) only adds the usual underline, so it's not applied
        collectRuns(child, props, ctx, runs, at(child, 'xlink', 'href') || href);
      } else if (is(child, 'draw', 'frame') || is(child, 'draw', 'a')) {
        for (const item of elementsOf(child)) {
          if (is(item, 'draw', 'image') && ctx.images.has(item)) {
            runs.push({ img: { src: ctx.images.get(item), width: lengthToPx(at(child, 'svg', 'width')), height: lengthToPx(at(child, 'svg', 'height')) }, href });
            break;
          }
          if (is(item, 'draw', 'text-box')) ctx.pending.push({ nodes: elementsOf(item), box: frameBox(child, item, ctx) });
          if (is(item, 'draw', 'frame')) collectRuns(child, props, ctx, runs, href);
        }
      } else if (/^(note|bookmark|bookmark-start|bookmark-end|bookmark-ref|reference-mark|reference-mark-start|reference-mark-end|toc-mark|toc-mark-start|toc-mark-end|alphabetical-index-mark|alphabetical-index-mark-start|alphabetical-index-mark-end|user-index-mark|user-index-mark-start|user-index-mark-end|soft-page-break|change|change-start|change-end|number|hidden-text|ruby-text)$/.test(child.localName) && child.namespaceURI === NS.text) {
        // Footnotes, bookmarks, index marks, tracked-change markers and cached list numbers aren't document text here
      } else if (child.namespaceURI === NS.office && /^annotation(-end)?$/.test(child.localName)) {
        // Comments are left out
      } else if (child.namespaceURI === NS.draw) {
        // Shapes other than images
      } else {
        // Fields (dates, page numbers, titles...) and unknown elements: keep their displayed text
        collectRuns(child, props, ctx, runs, href);
      }
    }
  };

  // Only the character formatting a run can carry
  const strip = (p) => {
    const out = {};
    for (const k of ['b', 'i', 'u', 's', 'font', 'color', 'bg', 'size']) if (p[k] !== undefined) out[k] = p[k];
    return out;
  };

  // Trims collapsible spaces at line edges and merges a space that follows another
  const normalizeOdfRuns = (runs) => {
    let lineStart = true;
    let afterSpace = false;
    for (const run of runs) {
      if (run.br) {
        lineStart = true;
        afterSpace = false;
      } else if (run.img) {
        lineStart = false;
        afterSpace = false;
      } else if (!run.keep) {
        if (lineStart || afterSpace) run.text = run.text.replace(/^ /, '');
        if (run.text) {
          lineStart = false;
          afterSpace = run.text.endsWith(' ');
        }
      } else if (run.text) {
        lineStart = false;
        afterSpace = false;
      }
    }
    for (let i = runs.length - 1; i >= 0; i--) {
      if (runs[i].br || runs[i].img || runs[i].keep) break;
      runs[i].text = runs[i].text.replace(/ +$/, '');
      if (runs[i].text) break;
    }
    return runs.filter((run) => run.br || run.img || run.text);
  };

  const VALUE_PROPS = ['font', 'color', 'bg', 'size'];
  const sameFormat = (a, b) => a.href === b.href
    && ['b', 'i', 'u', 's', 'code'].every((k) => Boolean(a.props?.[k]) === Boolean(b.props?.[k]))
    && VALUE_PROPS.every((k) => (a.props?.[k] || '') === (b.props?.[k] || ''));

  // Writes runs into an HTML element, wrapping a > font > strong > em > u > s > code
  const emitRuns = (runs, target, base, ctx) => {
    const merged = [];
    for (const run of runs) {
      const last = merged[merged.length - 1];
      if (last && !last.br && !last.img && !run.br && !run.img && sameFormat(last, run)) last.text += run.text;
      else merged.push({ ...run });
    }
    let linkEl = null;
    let linkHref = null;
    for (const run of merged) {
      let node;
      if (run.br) {
        node = document.createElement('br');
      } else if (run.img) {
        node = document.createElement('img');
        node.setAttribute('src', run.img.src);
        node.setAttribute('alt', '');
        if (run.img.width) node.setAttribute('width', run.img.width);
        if (run.img.height) node.setAttribute('height', run.img.height);
      } else {
        node = document.createTextNode(run.text);
        const p = run.props || {};
        const wrap = (tag, attrs = {}) => {
          const el = document.createElement(tag);
          for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
          el.append(node);
          node = el;
        };
        if (p.code) wrap('code');
        if (p.s && !base.s) wrap('s');
        if (p.u && !base.u) wrap('u');
        if (p.i && !base.i) wrap('em');
        if (p.b && !base.b) wrap('strong');
        // Code is already monospace, so it doesn't also need a font
        if (p.font && !p.code && p.font !== ctx.pageFont && p.font !== base.font) wrap('font', { face: p.font });
        const css = [p.color && `color: #${p.color}`, p.bg && `background-color: #${p.bg}`, p.size && `font-size: ${p.size}pt`].filter(Boolean).join('; ');
        if (css) wrap('span', { style: css });
      }
      if (run.href) {
        if (run.href !== linkHref) {
          linkEl = document.createElement('a');
          linkEl.setAttribute('href', run.href);
          linkHref = run.href;
          target.append(linkEl);
        }
        linkEl.append(node);
      } else {
        linkHref = null;
        target.append(node);
      }
    }
  };

  // Formatting a paragraph style applies to all its text becomes the block's own (bold headings aren't <strong>)
  const paragraphInto = (el, block, ctx, { heading = false, inList = false } = {}) => {
    const props = ctx.styles.resolve('paragraph', at(el, 'text', 'style-name'));
    const runs = [];
    collectRuns(el, strip(props), ctx, runs, null);
    const base = heading ? { b: true } : {};
    emitRuns(normalizeOdfRuns(runs), block, base, ctx);
    // A list's indent comes from its nesting, not from its paragraphs' margins
    const css = [props.align && `text-align: ${props.align}`, !inList && props.indent && `margin-left: ${props.indent}pt`].filter(Boolean).join('; ');
    if (css) block.setAttribute('style', css);
    if (!block.hasChildNodes()) block.append(document.createElement('br'));
    return props;
  };

  // Text boxes anchored in the paragraph just converted go after it, each as a Wurd text box
  // (one inside another box just adds its contents, as boxes don't nest)
  // Free boxes go before the paragraph they're anchored in, which goes if it held nothing else
  const flushPending = (ctx, parent) => {
    if (!ctx.pending.length) return;
    const anchor = parent.lastElementChild;
    let floated = false;
    for (const { nodes, box: geometry } of ctx.pending.splice(0)) {
      if (Wurd.isTextBox(parent) || parent.closest?.('.wurd-textbox')) {
        convertBlocks(nodes, parent, ctx);
        continue;
      }
      const box = document.createElement('div');
      box.className = 'wurd-textbox';
      convertBlocks(nodes, box, ctx);
      if (!box.hasChildNodes()) continue;
      Wurd.setBoxGeometry(box, geometry);
      if (geometry?.floating && anchor && !Wurd.isTextBox(anchor)) {
        anchor.before(box);
        floated = true;
      } else {
        parent.append(box);
      }
    }
    if (floated && !anchor.textContent.trim() && !anchor.querySelector('img, hr, table, .wurd-textbox')) anchor.remove();
  };

  const pageBreak = () => {
    const hr = document.createElement('hr');
    hr.className = 'wurd-page-break';
    return hr;
  };

  const MAX_TABLE_ROWS = 1000;
  const MAX_TABLE_CELLS = 64;

  const convertTable = (el, ctx) => {
    const tableEl = document.createElement('table');
    const tbody = document.createElement('tbody');
    let thead = null;
    let rows = 0;
    const addRows = (container, section, header) => {
      for (const child of elementsOf(container)) {
        if (is(child, 'table', 'table-header-rows')) {
          thead = thead || document.createElement('thead');
          addRows(child, thead, true);
        } else if (is(child, 'table', 'table-rows') || is(child, 'table', 'table-row-group')) {
          addRows(child, section, header);
        } else if (is(child, 'table', 'table-row')) {
          const repeat = Math.max(1, Number(at(child, 'table', 'number-rows-repeated')) || 1);
          for (let r = 0; r < repeat && rows < MAX_TABLE_ROWS; r++, rows++) {
            const tr = document.createElement('tr');
            let cells = 0;
            for (const cell of elementsOf(child)) {
              if (!is(cell, 'table', 'table-cell')) continue; // covered cells are part of a span
              const cellRepeat = Math.max(1, Number(at(cell, 'table', 'number-columns-repeated')) || 1);
              for (let c = 0; c < cellRepeat && cells < MAX_TABLE_CELLS; c++, cells++) {
                const td = document.createElement(header ? 'th' : 'td');
                const colspan = Number(at(cell, 'table', 'number-columns-spanned')) || 1;
                const rowspan = Number(at(cell, 'table', 'number-rows-spanned')) || 1;
                if (colspan > 1) td.setAttribute('colspan', Math.min(colspan, MAX_TABLE_CELLS));
                if (rowspan > 1) td.setAttribute('rowspan', Math.min(rowspan, MAX_TABLE_ROWS));
                convertBlocks(elementsOf(cell), td, ctx);
                // A cell with one plain paragraph holds its text directly
                const only = td.children.length === 1 && td.firstElementChild;
                if (only && only.tagName === 'P' && !only.hasAttribute('style') && td.childNodes.length === 1) only.replaceWith(...only.childNodes);
                if (td.childNodes.length === 1 && td.firstChild.nodeName === 'BR') td.firstChild.remove();
                tr.append(td);
              }
            }
            if (tr.childNodes.length) section.append(tr);
          }
        }
      }
    };
    addRows(el, tbody, false);
    if (thead) tableEl.append(thead);
    tableEl.append(tbody);
    Wurd.formats.unwrapHeaderBold(tableEl);
    return tableEl;
  };

  const convertList = (el, ctx, level, inheritedStyle) => {
    let styleName = at(el, 'text', 'style-name') || inheritedStyle;
    if (!styleName) {
      // Lists can also take their style from the first item's paragraph style
      const firstParagraph = el.getElementsByTagNameNS(NS.text, 'p')[0];
      if (firstParagraph) styleName = ctx.styles.resolve('paragraph', at(firstParagraph, 'text', 'style-name')).listStyle;
    }
    const list = document.createElement(ctx.styles.listKind(styleName, level));
    let first = true;
    for (const item of elementsOf(el)) {
      if (!is(item, 'text', 'list-item') && !is(item, 'text', 'list-header')) continue;
      const start = at(item, 'text', 'start-value');
      if (first && start && list.tagName === 'OL') list.setAttribute('start', start);
      first = false;
      const li = document.createElement('li');
      const paragraphs = elementsOf(item).filter((c) => is(c, 'text', 'p') || is(c, 'text', 'h'));
      for (const child of elementsOf(item)) {
        if (is(child, 'text', 'list')) {
          li.append(convertList(child, ctx, level + 1, styleName));
        } else if (is(child, 'text', 'p') || is(child, 'text', 'h')) {
          if (paragraphs.length === 1) {
            paragraphInto(child, li, ctx, { inList: true });
            if (li.childNodes.length === 1 && li.firstChild.nodeName === 'BR' && elementsOf(item).length > 1) li.firstChild.remove();
          } else {
            const p = document.createElement('p');
            paragraphInto(child, p, ctx, { inList: true });
            li.append(p);
          }
        } else if (!is(child, 'text', 'number')) {
          convertBlocks([child], li, ctx);
        }
      }
      list.append(li);
    }
    return list;
  };

  function convertBlocks(elements, parent, ctx) {
    let pre = null;
    let quote = null;
    for (const el of elements) {
      if (is(el, 'text', 'p') || is(el, 'text', 'h')) {
        const props = ctx.styles.resolve('paragraph', at(el, 'text', 'style-name'));
        const kind = is(el, 'text', 'h') ? 'heading' : semantic(props);
        // Page breaks are kept between top-level blocks; a paragraph that's only there to carry one
        // (as Wurd writes them) becomes just the break
        const breaks = parent === ctx.root && kind !== 'pre';
        if (breaks && props.breakBefore) parent.append(pageBreak());
        if (breaks && (props.breakBefore || props.breakAfter) && !el.textContent.trim() && !el.getElementsByTagNameNS(NS.draw, 'frame').length) {
          if (props.breakAfter) parent.append(pageBreak());
          pre = null;
          quote = null;
          continue;
        }
        if (kind === 'pre') {
          // Consecutive preformatted paragraphs become one <pre>
          if (!pre) {
            pre = document.createElement('pre');
            parent.append(pre);
          } else {
            pre.append('\n');
          }
          const runs = [];
          collectRuns(el, {}, ctx, runs, null);
          pre.append(runs.map((r) => (r.br ? '\n' : r.text || '')).join('').split(NBSP).join(' '));
          quote = null;
          continue;
        }
        pre = null;
        if (kind === 'hr' && !el.textContent.trim()) {
          parent.append(document.createElement('hr'));
          quote = null;
          continue;
        }
        if (kind === 'heading' || kind === 'title') {
          // Level from the outline level, or from a "Heading N" style on a plain paragraph
          const styleLevel = Number(/^Heading_20_(\d+)$/.exec(props.chain.find((n) => /^Heading_20_\d+$/.test(n)) || '')?.[1]);
          const level = kind === 'title' ? 1 : Math.min(Math.max(Number(at(el, 'text', 'outline-level')) || styleLevel || 1, 1), 4);
          const h = document.createElement(`h${level}`);
          paragraphInto(el, h, ctx, { heading: true });
          parent.append(h);
          quote = null;
        } else {
          const p = document.createElement('p');
          paragraphInto(el, p, ctx);
          if (kind === 'quote') {
            if (!quote) {
              quote = document.createElement('blockquote');
              parent.append(quote);
            }
            quote.append(p);
          } else {
            parent.append(p);
            quote = null;
          }
        }
        flushPending(ctx, parent);
        if (breaks && props.breakAfter) {
          parent.append(pageBreak());
          quote = null;
        }
        continue;
      }
      pre = null;
      quote = null;
      if (is(el, 'text', 'list')) {
        parent.append(convertList(el, ctx, 1, null));
      } else if (is(el, 'table', 'table')) {
        // Wurd writes text boxes as one-cell tables named WurdTextBox<n>
        const cells = el.getElementsByTagNameNS(NS.table, 'table-cell');
        if (/^WurdTextBox/.test(at(el, 'table', 'name') || '') && cells.length === 1 && !(Wurd.isTextBox(parent) || parent.closest?.('.wurd-textbox'))) {
          const box = document.createElement('div');
          box.className = 'wurd-textbox';
          convertBlocks(elementsOf(cells[0]), box, ctx);
          if (!box.hasChildNodes()) box.append(document.createElement('p'));
          const tableStyle = ctx.styles.resolve('table', at(el, 'table', 'style-name'));
          const row = el.getElementsByTagNameNS(NS.table, 'table-row')[0];
          const rowStyle = row ? ctx.styles.resolve('table-row', at(row, 'table', 'style-name')) : {};
          Wurd.setBoxGeometry(box, {
            widthPct: tableStyle.tableAlign !== 'margins' ? (pt(tableStyle.tableWidth) / ctx.page.textWidthPt) * 100 : NaN,
            minHeightPt: pt(rowStyle.minRowHeight),
          });
          parent.append(box);
        } else {
          parent.append(convertTable(el, ctx));
        }
      } else if (is(el, 'text', 'section') || is(el, 'text', 'index-body') || is(el, 'text', 'index-title')) {
        convertBlocks(elementsOf(el), parent, ctx);
      } else if (el.namespaceURI === NS.text && /(^|-)index$|^table-of-content$|^bibliography$/.test(el.localName)) {
        const body = findChild(el, 'text', 'index-body');
        if (body) convertBlocks(elementsOf(body), parent, ctx);
      }
      // Declarations, tracked changes, forms and the like have no document text
    }
  }

  // The page's width and left margin, in points, from the first page layout (A4 with 2cm margins if none)
  const pageGeometry = (stylesDoc) => {
    const props = stylesDoc?.getElementsByTagNameNS(NS.style, 'page-layout-properties')[0];
    const read = (name, fallback) => {
      const n = Wurd.cssLengthToPt(props && at(props, 'fo', name));
      return Number.isFinite(n) && n > 0 ? n : fallback;
    };
    const widthPt = read('page-width', 595.3);
    const marginLeftPt = read('margin-left', 56.7);
    return { widthPt, heightPt: read('page-height', 841.9), marginLeftPt, textWidthPt: widthPt - marginLeftPt - read('margin-right', 56.7) };
  };

  const importOdt = async (bytes, { askPassword } = {}) => {
    const pkg = await F.odfPackage.read(bytes, { askPassword });
    const contentXml = await pkg.readText('content.xml');
    if (!contentXml) throw new Wurd.UserError('This OpenDocument file has no content.');
    const stylesXml = await pkg.readText('styles.xml');
    const contentDoc = parseXml(contentXml);
    const stylesDoc = stylesXml ? parseXml(stylesXml) : null;
    const styles = buildStyleIndex(contentDoc, stylesDoc);
    const ctx = { styles, pageFont: styles.pageFont, page: pageGeometry(stylesDoc), images: await loadImages(contentDoc, pkg), pending: [] };
    const body = contentDoc.getElementsByTagNameNS(NS.office, 'text')[0];
    if (!body) throw new Wurd.UserError("This OpenDocument file isn't a text document.");
    const container = document.createElement('div');
    ctx.root = container;
    convertBlocks(elementsOf(body), container, ctx);
    flushPending(ctx, container);
    return { html: container.innerHTML || '<p><br></p>', font: styles.pageFont, password: pkg.password };
  };

  // ---- Export ---------------------------------------------------------------

  // Wurd's fonts as ODF font faces
  const FONT_FACES = {
    [Wurd.DEFAULT_FONT]: { name: 'Inter', family: 'Inter', generic: 'swiss', pitch: 'variable' },
    [Wurd.SERIF_FONT]: { name: 'Georgia', family: 'Georgia', generic: 'roman', pitch: 'variable' },
    [Wurd.MONO_FONT]: { name: 'Courier New', family: "'Courier New'", generic: 'modern', pitch: 'fixed' },
  };
  const faceFor = (font) => FONT_FACES[Wurd.classifyFont(Wurd.firstFontFamily(font))] || FONT_FACES[Wurd.DEFAULT_FONT];

  const NAMESPACES = Object.entries({ office: NS.office, style: NS.style, text: NS.text, table: NS.table, draw: NS.draw, fo: NS.fo, svg: NS.svg, xlink: NS.xlink })
    .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`).join(' ');

  const fontFaceDecls = () => `<office:font-face-decls>${Object.values(FONT_FACES).map((f) =>
    `<style:font-face style:name="${f.name}" svg:font-family="${f.family}" style:font-family-generic="${f.generic}" style:font-pitch="${f.pitch}"/>`).join('')}</office:font-face-decls>`;

  const listLevels = (kind) => Array.from({ length: 10 }, (_, i) => {
    const level = i + 1;
    const position = (0.635 * level).toFixed(3);
    const props = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${position}cm" fo:text-indent="-0.635cm" fo:margin-left="${position}cm"/></style:list-level-properties>`;
    return kind === 'bullet'
      ? `<text:list-level-style-bullet text:level="${level}" text:bullet-char="•">${props}</text:list-level-style-bullet>`
      : `<text:list-level-style-number text:level="${level}" style:num-suffix="." style:num-format="1">${props}</text:list-level-style-number>`;
  }).join('');

  const bold = 'fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"';

  // Page size from Settings (A4 or Letter) with 2cm margins
  const MARGIN_CM = 2;
  const pageFor = (size) => {
    const page = Wurd.PAGE_SIZES[size] || Wurd.PAGE_SIZES.a4;
    return { width: page.widthCm, height: page.heightCm, textWidthCm: parseFloat(page.widthCm) - 2 * MARGIN_CM };
  };

  // The named styles .odt and .odg files both carry
  const namedStylesXml = (face) => {
    const heading = (n, size) => `<style:style style:name="Heading_20_${n}" style:display-name="Heading ${n}" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="${n}" style:class="text"><style:text-properties fo:font-size="${size}pt" ${bold}/></style:style>`;
    return `<style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="1.25cm"/><style:text-properties style:font-name="${face.name}" fo:font-size="12pt" fo:language="en" fo:country="GB"/></style:default-style>
<style:style style:name="Standard" style:family="paragraph" style:class="text"/>
<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text"><style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:keep-with-next="always"/></style:style>
${heading(1, 27)}${heading(2, 18)}${heading(3, 15)}${heading(4, 12)}
<style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.247cm"/></style:style>
<style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-left="1cm" fo:margin-right="1cm" fo:margin-top="0cm" fo:margin-bottom="0.283cm"/></style:style>
<style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/><style:text-properties style:font-name="Courier New" fo:font-size="10pt"/></style:style>
<style:style style:name="Horizontal_20_Line" style:display-name="Horizontal Line" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.5cm" fo:padding="0cm" fo:border-bottom="0.5pt solid #808080"/><style:text-properties fo:font-size="6pt"/></style:style>
<style:style style:name="Table_20_Contents" style:display-name="Table Contents" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"/>
<style:style style:name="Table_20_Heading" style:display-name="Table Heading" style:family="paragraph" style:parent-style-name="Table_20_Contents" style:class="extra"><style:text-properties ${bold}/></style:style>
<style:style style:name="Source_20_Text" style:display-name="Source Text" style:family="text"><style:text-properties style:font-name="Courier New"/></style:style>
<style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text"><style:text-properties fo:color="#000080" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
<text:outline-style style:name="Outline">${Array.from({ length: 10 }, (_, i) => `<text:outline-level-style text:level="${i + 1}" style:num-format=""/>`).join('')}</text:outline-style>
<text:list-style style:name="Wurd_20_Bullets" style:display-name="Wurd Bullets">${listLevels('bullet')}</text:list-style>
<text:list-style style:name="Wurd_20_Numbers" style:display-name="Wurd Numbers">${listLevels('number')}</text:list-style>`;
  };

  const stylesXml = (pageFont, page) => `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NAMESPACES} office:version="${F.odfPackage.ODF_VERSION}">
${fontFaceDecls()}
<office:styles>
${namedStylesXml(faceFor(pageFont))}
</office:styles>
<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="${page.width}" fo:page-height="${page.height}" style:print-orientation="portrait" fo:margin-top="${MARGIN_CM}cm" fo:margin-bottom="${MARGIN_CM}cm" fo:margin-left="${MARGIN_CM}cm" fo:margin-right="${MARGIN_CM}cm" style:writing-mode="lr-tb"/></style:page-layout></office:automatic-styles>
<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>
</office:document-styles>
`;

  // A drawing's pages have no margins: every frame is placed from the paper's top-left corner.
  // Text in frames takes its font from the default graphic style.
  const drawingStylesXml = (pageFont, page) => {
    const face = faceFor(pageFont);
    return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NAMESPACES} office:version="${F.odfPackage.ODF_VERSION}">
${fontFaceDecls()}
<office:styles>
<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#808080" draw:fill-color="#ffffff"/><style:text-properties style:font-name="${face.name}" fo:font-size="12pt" fo:language="en" fo:country="GB"/></style:default-style>
${namedStylesXml(face)}
</office:styles>
<office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="${page.width}" fo:page-height="${page.height}" style:print-orientation="portrait"/></style:page-layout><style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="none"/></style:style></office:automatic-styles>
<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp1"/></office:master-styles>
</office:document-styles>
`;
  };

  const metaXml = (title) => {
    const now = new Date().toISOString().replace(/\.\d+Z$/, '');
    return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="${NS.office}" xmlns:meta="${NS.meta}" xmlns:dc="${NS.dc}" office:version="${F.odfPackage.ODF_VERSION}"><office:meta><meta:generator>Wurd</meta:generator><dc:title>${Wurd.escapeXml(title)}</dc:title><meta:creation-date>${now}</meta:creation-date><dc:date>${now}</dc:date></office:meta></office:document-meta>
`;
  };

  const BLOCK = /^(P|DIV|H[1-6]|UL|OL|LI|BLOCKQUOTE|PRE|HR|TABLE|SECTION|ARTICLE|HEADER|FOOTER|FIGURE|FIGCAPTION|ADDRESS|DL|DT|DD)$/;
  const FONT_SIZES_PT = [8, 10, 12, 14, 18, 24, 36]; // <font size="1..7">
  const ODF_ALIGN = { left: 'start', start: 'start', right: 'end', end: 'end', center: 'center', justify: 'justify' };

  // Automatic styles, numbered in first-use order so the output is deterministic
  const textStyle = (ctx, f) => {
    const font = f.code ? FONT_FACES[Wurd.MONO_FONT].name : f.font && f.font !== ctx.pageFace ? f.font : null;
    const plain = !f.b && !f.i && !f.u && !f.s && !f.color && !f.bg && !f.size;
    if (plain && !font) return null;
    if (plain && f.code && !f.font) return 'Source_20_Text';
    const key = `${+!!f.b}${+!!f.i}${+!!f.u}${+!!f.s}|${font || ''}|${f.color || ''}|${f.bg || ''}|${f.size || ''}`;
    if (!ctx.textStyles.has(key)) {
      const props = [
        f.b && bold,
        f.i && 'fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"',
        f.u && 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"',
        f.s && 'style:text-line-through-style="solid" style:text-line-through-type="single"',
        font && `style:font-name="${font}"`,
        f.color && `fo:color="#${f.color}"`,
        f.bg && `fo:background-color="#${f.bg}"`,
        f.size && `fo:font-size="${f.size}pt" style:font-size-asian="${f.size}pt" style:font-size-complex="${f.size}pt"`,
      ].filter(Boolean).join(' ');
      ctx.textStyles.set(key, { name: `T${ctx.textStyles.size + 1}`, xml: props });
    }
    return ctx.textStyles.get(key).name;
  };

  // In a drawing, LibreOffice Draw may not apply a text document's named paragraph styles, so each
  // automatic style also spells out how its named style looks
  const HEADING_LOOK = (size) => ({ para: { 'fo:margin-top': '0.423cm', 'fo:margin-bottom': '0.212cm' }, text: `fo:font-size="${size}pt" ${bold}` });
  const DRAWING_LOOK = {
    Text_20_body: { para: { 'fo:margin-top': '0cm', 'fo:margin-bottom': '0.247cm' } },
    Heading_20_1: HEADING_LOOK(27),
    Heading_20_2: HEADING_LOOK(18),
    Heading_20_3: HEADING_LOOK(15),
    Heading_20_4: HEADING_LOOK(12),
    Quotations: { para: { 'fo:margin-left': '1cm', 'fo:margin-right': '1cm', 'fo:margin-top': '0cm', 'fo:margin-bottom': '0.283cm' } },
    Table_20_Heading: { para: {}, text: bold },
  };

  // An automatic paragraph style for alignment, a left indent (points) and a page break after
  const paragraphStyle = (ctx, parent, align, { indentPt = 0, breakAfter = false } = {}) => {
    const odfAlign = ODF_ALIGN[align] === 'start' ? null : ODF_ALIGN[align];
    const indent = Math.round(indentPt * 100) / 100;
    const look = ctx.drawing ? DRAWING_LOOK[parent] : null;
    if (!odfAlign && !indent && !breakAfter && !look) return parent;
    const key = `${parent}|${odfAlign || ''}|${indent}|${+breakAfter}`;
    if (!ctx.paragraphStyles.has(key)) {
      const name = `P${ctx.paragraphStyles.size + 1}`;
      const props = { ...(look?.para || {}) };
      if (odfAlign) props['fo:text-align'] = odfAlign;
      if (odfAlign === 'justify') props['style:justify-single-word'] = 'false';
      if (indent) props['fo:margin-left'] = `${indent}pt`;
      if (breakAfter) props['fo:break-after'] = 'page';
      const attributes = Object.entries(props).map(([k, v]) => `${k}="${v}"`).join(' ');
      ctx.paragraphStyles.set(key, {
        name,
        xml: `<style:style style:name="${name}" style:family="paragraph" style:parent-style-name="${parent}"><style:paragraph-properties ${attributes}/>${look?.text ? `<style:text-properties ${look.text}/>` : ''}</style:style>`,
      });
    }
    return ctx.paragraphStyles.get(key).name;
  };

  const alignOf = (el) => (el.style?.textAlign || el.getAttribute?.('align') || '').toLowerCase();

  // Inline HTML -> runs { text, f, href } | { br } | { img }
  const inlineRuns = (nodes, ctx, f, href, runs) => {
    for (const node of nodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        runs.push({ text: Wurd.collapseSpaces(node.data), f, href });
        continue;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) continue;
      const tag = node.tagName;
      if (tag === 'BR') {
        runs.push({ br: true });
        continue;
      }
      if (tag === 'IMG') {
        const image = ctx.images.get(node);
        if (image) runs.push({ img: image, href });
        continue;
      }
      const next = { ...f };
      let link = href;
      if (tag === 'B' || tag === 'STRONG') next.b = true;
      else if (tag === 'I' || tag === 'EM' || tag === 'CITE') next.i = true;
      else if (tag === 'U' || tag === 'INS') next.u = true;
      else if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') next.s = true;
      else if (tag === 'CODE' || tag === 'KBD' || tag === 'TT' || tag === 'SAMP') next.code = true;
      else if (tag === 'MARK') next.bg = Wurd.HIGHLIGHTS.yellow;
      else if (tag === 'FONT') {
        if (node.getAttribute('face')) next.font = faceFor(node.getAttribute('face')).name;
        if (node.getAttribute('color')) next.color = Wurd.colorToHex(node.getAttribute('color')) || next.color;
        const size = Number(node.getAttribute('size'));
        if (size >= 1 && size <= 7) next.size = FONT_SIZES_PT[size - 1];
      } else if (tag === 'A' && /^(https?:|mailto:|tel:|#)/i.test(node.getAttribute('href') || '')) link = node.getAttribute('href');
      const style = node.style;
      if (style) {
        if (style.fontWeight === 'bold' || Number(style.fontWeight) >= 600) next.b = true;
        if (style.fontStyle === 'italic') next.i = true;
        if (/underline/.test(style.textDecorationLine || style.textDecoration || '')) next.u = true;
        if (/line-through/.test(style.textDecorationLine || style.textDecoration || '')) next.s = true;
        if (style.fontFamily) next.font = faceFor(style.fontFamily).name;
        if (style.color) next.color = Wurd.colorToHex(style.color) || next.color;
        if (style.backgroundColor) next.bg = Wurd.colorToHex(style.backgroundColor) || next.bg;
        const size = Wurd.cssLengthToPt(style.fontSize);
        if (size > 0 && /(pt|px)$/.test(style.fontSize)) next.size = Math.round(size * 2) / 2;
      }
      if (BLOCK.test(tag) && runs.length) runs.push({ br: true });
      inlineRuns(node.childNodes, ctx, next, link, runs);
    }
    return runs;
  };

  // Text with ODF's space rules: a space after a space (or at a line start) must be written as text:s
  const encodeText = (text, state) => {
    let out = '';
    let spaces = 0;
    const flushSpaces = () => {
      if (!spaces) return;
      const literal = state.lineStart || state.afterSpace ? 0 : 1;
      if (literal) out += ' ';
      if (spaces - literal > 0) out += spaces - literal === 1 ? '<text:s/>' : `<text:s text:c="${spaces - literal}"/>`;
      state.afterSpace = true;
      state.lineStart = false;
      spaces = 0;
    };
    for (const ch of text.split(NBSP).join(' ')) {
      if (ch === ' ') {
        spaces++;
      } else if (ch === '\t') {
        flushSpaces();
        out += '<text:tab/>';
        state.afterSpace = false;
        state.lineStart = false;
      } else {
        flushSpaces();
        out += Wurd.escapeXml(ch);
        state.afterSpace = false;
        state.lineStart = false;
      }
    }
    flushSpaces();
    return out;
  };

  // Runs -> paragraph XML content (spans for formatting, text:a for links)
  const runsXml = (runs, ctx) => {
    // Drop the placeholder <br> contentEditable leaves at the end of a block, then trim collapsible spaces
    if (runs.length && runs[runs.length - 1].br) runs.pop();
    let lineStart = true;
    let afterSpace = false;
    for (const run of runs) {
      if (run.br) {
        lineStart = true;
        afterSpace = false;
      } else if (run.img) {
        lineStart = false;
        afterSpace = false;
      } else {
        if (lineStart || afterSpace) run.text = run.text.replace(/^ /, '');
        if (run.text) {
          lineStart = false;
          afterSpace = run.text.endsWith(' ');
        }
      }
    }
    for (let i = runs.length - 1; i >= 0 && !runs[i].br && !runs[i].img; i--) {
      runs[i].text = runs[i].text.replace(/ +$/, '');
      if (runs[i].text) break;
    }
    const state = { lineStart: true, afterSpace: false };
    let xml = '';
    let openHref = null;
    for (const run of runs.filter((r) => r.br || r.img || r.text)) {
      const href = run.href || null;
      if (href !== openHref) {
        if (openHref) xml += '</text:a>';
        if (href) xml += `<text:a xlink:type="simple" xlink:href="${Wurd.escapeXml(href)}" text:style-name="Internet_20_link">`;
        openHref = href;
      }
      if (run.br) {
        xml += '<text:line-break/>';
        state.lineStart = true;
        state.afterSpace = false;
      } else if (run.img) {
        const img = run.img;
        xml += `<draw:frame draw:style-name="fr1" draw:name="Image${img.n}" text:anchor-type="as-char" svg:width="${img.widthCm}cm" svg:height="${img.heightCm}cm" draw:z-index="0"><draw:image xlink:href="${img.path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
        state.lineStart = false;
        state.afterSpace = false;
      } else {
        const style = textStyle(ctx, run.f);
        const text = encodeText(run.text, state);
        xml += style ? `<text:span text:style-name="${style}">${text}</text:span>` : text;
      }
    }
    if (openHref) xml += '</text:a>';
    return xml;
  };

  const paragraphXml = (nodes, ctx, styleName, { keepEmpty = false, heading = 0, base = {} } = {}) => {
    const runs = inlineRuns(nodes, ctx, base, null, []);
    const content = runsXml(runs, ctx);
    if (!content && !keepEmpty) return '';
    if (heading) return `<text:h text:style-name="${styleName}" text:outline-level="${heading}">${content}</text:h>`;
    return content ? `<text:p text:style-name="${styleName}">${content}</text:p>` : `<text:p text:style-name="${styleName}"/>`;
  };

  // One Preformatted Text paragraph per line, with every space kept
  const preXml = (el) => el.textContent.replace(/\n$/, '').split('\n').map((line) => {
    const text = encodeText(line, { lineStart: true, afterSpace: false });
    return text ? `<text:p text:style-name="Preformatted_20_Text">${text}</text:p>` : '<text:p text:style-name="Preformatted_20_Text"/>';
  }).join('');

  const listXml = (list, ctx, state) => {
    const style = ctx.drawing ? (list.tagName === 'OL' ? 'WurdL2' : 'WurdL1') : list.tagName === 'OL' ? 'Wurd_20_Numbers' : 'Wurd_20_Bullets';
    let xml = `<text:list text:style-name="${style}">`;
    let first = true;
    const itemState = { ...state, parent: 'Standard', inList: true, indentPt: 0 };
    for (const child of list.children) {
      if (child.tagName === 'UL' || child.tagName === 'OL') {
        xml += `<text:list-item>${listXml(child, ctx, itemState)}</text:list-item>`;
        continue;
      }
      if (child.tagName !== 'LI') continue;
      const start = first && list.tagName === 'OL' && Number(list.getAttribute('start')) > 1 ? ` text:start-value="${Number(list.getAttribute('start'))}"` : '';
      first = false;
      let content = '';
      let inline = [];
      const flush = () => {
        if (inline.length) content += paragraphXml(inline, ctx, paragraphStyle(ctx, 'Standard', alignOf(child)));
        inline = [];
      };
      for (const node of child.childNodes) {
        if (node.nodeType === Node.ELEMENT_NODE && (node.tagName === 'UL' || node.tagName === 'OL')) {
          flush();
          content += listXml(node, ctx, itemState);
        } else if (node.nodeType === Node.ELEMENT_NODE && BLOCK.test(node.tagName)) {
          flush();
          content += blockXml(node, ctx, itemState);
        } else {
          inline.push(node);
        }
      }
      flush();
      // A list item needs at least one paragraph or list
      xml += `<text:list-item${start}>${content || '<text:p text:style-name="Standard"/>'}</text:list-item>`;
    }
    return `${xml}</text:list>`;
  };

  const tableXml = (table, ctx, state) => {
    const rows = Array.from(table.rows).filter((row) => row.cells.length);
    if (!rows.length) return '';
    // ODF tables can't sit inside list items; flatten them to paragraphs there
    if (state.inList) {
      return rows.map((row) => `<text:p text:style-name="Standard">${encodeText(Array.from(row.cells, (c) => Wurd.collapseSpaces(c.textContent).trim()).join('\t'), { lineStart: true, afterSpace: false })}</text:p>`).join('');
    }
    const n = ++ctx.tableCount;
    const columns = Math.max(1, ...rows.map((row) => Array.from(row.cells).reduce((sum, cell) => sum + (cell.colSpan || 1), 0)));
    ctx.columnStyles.add(columns);
    const occupied = new Set(); // "row,col" slots covered by a span from above
    let xml = `<table:table table:name="Table${n}" table:style-name="Tbl${n}"><table:table-column table:style-name="Wcol${columns}" table:number-columns-repeated="${columns}"/>`;
    let inHeader = false;
    rows.forEach((row, r) => {
      const header = row.parentElement?.tagName === 'THEAD' || Array.from(row.cells).every((cell) => cell.tagName === 'TH');
      // Only leading rows can repeat as table headings
      const asHeader = header && (r === 0 || inHeader);
      if (asHeader && !inHeader) xml += '<table:table-header-rows>';
      if (!asHeader && inHeader) xml += '</table:table-header-rows>';
      inHeader = asHeader;
      xml += '<table:table-row>';
      const cells = Array.from(row.cells);
      let c = 0;
      for (let col = 0; col < columns;) {
        if (occupied.has(`${r},${col}`)) {
          xml += '<table:covered-table-cell/>';
          col++;
          continue;
        }
        const cell = cells[c++];
        if (!cell) {
          xml += '<table:table-cell table:style-name="Wcell" office:value-type="string"><text:p text:style-name="Table_20_Contents"/></table:table-cell>';
          col++;
          continue;
        }
        const colspan = Math.min(cell.colSpan || 1, columns - col);
        const rowspan = Math.min(cell.rowSpan || 1, rows.length - r);
        for (let dr = 0; dr < rowspan; dr++) {
          for (let dc = 0; dc < colspan; dc++) if (dr || dc) occupied.add(`${r + dr},${col + dc}`);
        }
        const parent = cell.tagName === 'TH' ? 'Table_20_Heading' : 'Table_20_Contents';
        const content = blocksXml(cell.childNodes, ctx, { ...state, parent, align: alignOf(cell), inList: false, indentPt: 0 });
        const spans = `${colspan > 1 ? ` table:number-columns-spanned="${colspan}"` : ''}${rowspan > 1 ? ` table:number-rows-spanned="${rowspan}"` : ''}`;
        xml += `<table:table-cell table:style-name="Wcell" office:value-type="string"${spans}>${content || `<text:p text:style-name="${parent}"/>`}</table:table-cell>`;
        for (let dc = 1; dc < colspan; dc++) xml += '<table:covered-table-cell/>';
        col += colspan;
      }
      xml += '</table:table-row>';
    });
    if (inHeader) xml += '</table:table-header-rows>';
    return `${xml}</table:table>`;
  };

  const HEADING_LEVEL = { H1: 1, H2: 2, H3: 3, H4: 4, H5: 4, H6: 4 };

  const percentOf = (value, fallback) => {
    const pct = parseFloat(value);
    return Number.isFinite(pct) && /%$/.test(value || '') ? Math.min(Math.max(pct, 1), 100) / 100 : fallback;
  };
  const cmOf = (cssLength) => {
    const pt = Wurd.cssLengthToPt(cssLength);
    return Number.isFinite(pt) ? (pt / 72) * 2.54 : 0;
  };
  const cm = (n) => `${n.toFixed(3)}cm`;

  // A free text box is a frame in front of the text, anchored to a tiny paragraph where the box sits
  // among the blocks: left from the page edge, top from that paragraph
  const floatingBoxXml = (box, ctx, state) => {
    const content = blocksXml(box.childNodes, ctx, { ...state, inList: false, indentPt: 0 });
    if (state.inList) return content;
    const n = ++ctx.boxCount;
    ctx.floats = true;
    const pageWidth = parseFloat(ctx.page.width);
    const minHeight = Math.max(cmOf(box.style.minHeight), 0.5);
    const plain = box.classList.contains('wurd-textbox--plain');
    if (plain) ctx.plainFloats = true;
    return `<text:p text:style-name="WurdAnchor"><draw:frame draw:style-name="${plain ? 'WurdFloatPlain' : 'WurdFloat'}" draw:name="WurdTextBox${n}" text:anchor-type="paragraph" svg:x="${cm(percentOf(box.style.left, 0) * pageWidth)}" svg:y="${cm(cmOf(box.style.marginTop))}" svg:width="${cm(percentOf(box.style.width, 0.4) * pageWidth)}" draw:z-index="${n}"><draw:text-box fo:min-height="${cm(minHeight)}">${content || '<text:p text:style-name="Standard"/>'}</draw:text-box></draw:frame></text:p>`;
  };

  // An in-line text box is a one-cell bordered table named WurdTextBox<n>, so it imports back as a
  // text box. ODF tables can't sit inside list items, so there it's just its contents.
  const textBoxXml = (box, ctx, state) => {
    if (Wurd.isFloatingBox(box)) return floatingBoxXml(box, ctx, state);
    const content = blocksXml(box.childNodes, ctx, { ...state, inList: false, indentPt: 0 });
    if (state.inList) return content;
    const n = ++ctx.boxCount;
    const share = percentOf(box.style.width, 1);
    const width = cm(ctx.page.textWidthCm * share);
    const minHeight = cmOf(box.style.minHeight);
    ctx.boxStyles.push(
      `<style:style style:name="WurdBox${n}" style:family="table"><style:table-properties style:width="${width}" table:align="${share < 1 ? 'left' : 'margins'}" fo:margin-top="0.2cm" fo:margin-bottom="0.2cm"/></style:style>`,
      `<style:style style:name="WurdBoxCol${n}" style:family="table-column"><style:table-column-properties style:column-width="${width}"/></style:style>`,
      minHeight ? `<style:style style:name="WurdBoxRow${n}" style:family="table-row"><style:table-row-properties style:min-row-height="${cm(minHeight)}"/></style:style>` : '',
    );
    const row = minHeight ? ` table:style-name="WurdBoxRow${n}"` : '';
    const cell = box.classList.contains('wurd-textbox--plain') ? 'WurdBoxCellPlain' : 'WurdBoxCell';
    return `<table:table table:name="WurdTextBox${n}" table:style-name="WurdBox${n}"><table:table-column table:style-name="WurdBoxCol${n}"/><table:table-row${row}><table:table-cell table:style-name="${cell}" office:value-type="string">${content || '<text:p text:style-name="Standard"/>'}</table:table-cell></table:table-row></table:table>`;
  };

  function blockXml(el, ctx, state) {
    const tag = el.tagName;
    const align = alignOf(el) || state.align;
    const indentPt = (state.indentPt || 0) + Wurd.indentOf(el);
    if (HEADING_LEVEL[tag]) {
      const level = HEADING_LEVEL[tag];
      if (state.inList) return paragraphXml(Array.from(el.childNodes), ctx, paragraphStyle(ctx, state.parent, align), { keepEmpty: true, base: { b: true } });
      return paragraphXml(Array.from(el.childNodes), ctx, paragraphStyle(ctx, `Heading_20_${level}`, align, { indentPt }), { keepEmpty: true, heading: level });
    }
    if (tag === 'UL' || tag === 'OL') return listXml(el, ctx, state);
    if (tag === 'BLOCKQUOTE') return blocksXml(el.childNodes, ctx, { ...state, parent: 'Quotations', align });
    if (tag === 'PRE') return preXml(el);
    if (tag === 'HR') {
      // An empty paragraph with a page break after it
      if (Wurd.isPageBreak(el)) return `<text:p text:style-name="${paragraphStyle(ctx, 'Standard', '', { breakAfter: true })}"/>`;
      return '<text:p text:style-name="Horizontal_20_Line"/>';
    }
    if (tag === 'TABLE') return tableXml(el, ctx, state);
    if (Wurd.isTextBox(el)) return textBoxXml(el, ctx, state);
    if (Array.from(el.children).some((c) => BLOCK.test(c.tagName))) return blocksXml(el.childNodes, ctx, { ...state, align, indentPt });
    return paragraphXml(Array.from(el.childNodes), ctx, paragraphStyle(ctx, state.parent, align, { indentPt: state.inList ? 0 : indentPt }), { keepEmpty: tag === 'P' });
  }

  function blocksXml(nodes, ctx, state) {
    let xml = '';
    let inline = [];
    const flush = () => {
      if (inline.length) xml += paragraphXml(inline, ctx, paragraphStyle(ctx, state.parent, state.align, { indentPt: state.inList ? 0 : state.indentPt || 0 }));
      inline = [];
    };
    for (const node of nodes) {
      if (node.nodeType === Node.ELEMENT_NODE && BLOCK.test(node.tagName)) {
        flush();
        xml += blockXml(node, ctx, state);
      } else if (node.nodeType === Node.TEXT_NODE || node.nodeType === Node.ELEMENT_NODE) {
        inline.push(node);
      }
    }
    flush();
    return xml;
  }

  // Embedded data: images become Pictures/ entries (sizes need decoding first)
  const prepareImages = async (root, textWidthCm) => {
    const images = new Map();
    let n = 0;
    for (const img of root.querySelectorAll('img')) {
      const match = /^data:(image\/(png|jpe?g|gif|bmp|webp|svg\+xml));base64,(.*)$/i.exec(img.getAttribute('src') || '');
      if (!match) continue;
      try {
        const decoded = new Image();
        decoded.src = img.getAttribute('src');
        await decoded.decode();
        let width = Number(img.getAttribute('width')) || decoded.naturalWidth;
        let height = Number(img.getAttribute('height')) || decoded.naturalHeight;
        if (!width || !height) continue;
        let widthCm = width * 0.026458;
        let heightCm = height * 0.026458;
        if (widthCm > textWidthCm) {
          heightCm *= textWidthCm / widthCm;
          widthCm = textWidthCm;
        }
        n++;
        const ext = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/svg+xml': 'svg' }[match[1].toLowerCase()];
        images.set(img, {
          n,
          path: `Pictures/image${n}.${ext}`,
          mediaType: match[1].toLowerCase().replace('image/jpg', 'image/jpeg'),
          bytes: Wurd.base64ToBytes(match[3]),
          widthCm: widthCm.toFixed(3),
          heightCm: heightCm.toFixed(3),
        });
      } catch (err) {
        console.warn('Image left out of the .odt:', err);
      }
    }
    return images;
  };

  const exportOdt = async (root, { title, pageFont, pageSize, password }) => {
    const page = pageFor(pageSize);
    const ctx = await newExportContext(root, pageFont, page);
    const body = blocksXml(root.childNodes, ctx, { parent: 'Text_20_body', align: '', inList: false, indentPt: 0 });
    const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NAMESPACES} office:version="${F.odfPackage.ODF_VERSION}">
${fontFaceDecls()}
<office:automatic-styles>${automaticStylesXml(ctx)}</office:automatic-styles>
<office:body><office:text>${body || '<text:p text:style-name="Standard"/>'}</office:text></office:body>
</office:document-content>
`;
    const parts = [
      { path: 'content.xml', data: content, mediaType: 'text/xml' },
      { path: 'styles.xml', data: stylesXml(pageFont, page), mediaType: 'text/xml' },
      { path: 'meta.xml', data: metaXml(title), mediaType: 'text/xml' },
      ...Array.from(ctx.images.values(), (img) => ({ path: img.path, data: img.bytes, mediaType: img.mediaType })),
    ];
    return F.odfPackage.write(parts, { password });
  };

  const newExportContext = async (root, pageFont, page) => ({
    pageFace: faceFor(pageFont).name,
    textStyles: new Map(),
    paragraphStyles: new Map(),
    columnStyles: new Set(),
    tableCount: 0,
    boxCount: 0,
    boxStyles: [],
    floats: false,
    page,
    images: await prepareImages(root, page.textWidthCm),
  });

  // The automatic styles the body used, for content.xml
  const automaticStylesXml = (ctx) => {
    const { page } = ctx;
    const textWidth = page.textWidthCm.toFixed(3);
    return [
      ...Array.from(ctx.textStyles.values(), (s) => `<style:style style:name="${s.name}" style:family="text"><style:text-properties ${s.xml}/></style:style>`),
      ...Array.from(ctx.paragraphStyles.values(), (s) => s.xml),
      ...Array.from({ length: ctx.tableCount }, (_, i) => `<style:style style:name="Tbl${i + 1}" style:family="table"><style:table-properties style:width="${textWidth}cm" table:align="margins"/></style:style>`),
      ...Array.from(ctx.columnStyles, (n) => `<style:style style:name="Wcol${n}" style:family="table-column"><style:table-column-properties style:column-width="${(page.textWidthCm / n).toFixed(3)}cm"/></style:style>`),
      ...ctx.boxStyles,
      ctx.boxStyles.length ? '<style:style style:name="WurdBoxCell" style:family="table-cell"><style:table-cell-properties fo:padding="0.2cm" fo:border="0.5pt solid #808080"/></style:style><style:style style:name="WurdBoxCellPlain" style:family="table-cell"><style:table-cell-properties fo:padding="0.1cm" fo:border="none"/></style:style>' : '',
      ctx.floats ? '<style:style style:name="WurdAnchor" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:line-height="0.035cm"/><style:text-properties fo:font-size="1pt"/></style:style>' : '',
      ctx.floats ? '<style:style style:name="WurdFloat" style:family="graphic"><style:graphic-properties style:wrap="run-through" style:run-through="foreground" style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="page" fo:padding="0.15cm" fo:border="0.5pt solid #808080" fo:background-color="#ffffff" draw:fill="solid" draw:fill-color="#ffffff"/></style:style>' : '',
      ctx.plainFloats ? '<style:style style:name="WurdFloatPlain" style:family="graphic"><style:graphic-properties style:wrap="run-through" style:run-through="foreground" style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="page" fo:padding="0.1cm" fo:border="none" draw:fill="none"/></style:style>' : '',
      ctx.tableCount ? '<style:style style:name="Wcell" style:family="table-cell"><style:table-cell-properties fo:padding="0.097cm" fo:border="0.5pt solid #000000"/></style:style>' : '',
      ctx.images.size ? '<style:style style:name="fr1" style:family="graphic"><style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline"/></style:style>' : '',
    ].join('');
  };

  // ---- OpenDocument drawing (.odg) --------------------------------------------
  // A drawing is pages of positioned shapes. Opening one, each page becomes a Wurd page (pages are
  // separated by page breaks) whose text frames, shapes with text, pictures and tables become free
  // text boxes where they were. A page-high paragraph after them holds the page open. Saving, each
  // Wurd page becomes a drawing page: its text in a frame inside the margins, each free box a frame.

  const ODG_MIME = 'application/vnd.oasis.opendocument.graphics';
  const ODG_TYPES = /^application\/vnd\.oasis\.opendocument\.graphics(-template)?$/;
  // Wurd's page margin (1in): a drawing page's text frame sits inside it, and free boxes are measured from it
  const DRAWING_MARGIN_PT = 72;
  const SHAPES = /^(rect|ellipse|circle|custom-shape|polygon|polyline|path|regular-polygon|caption|measure|line|connector)$/;
  const ptToCm = (n) => `${((n / 72) * 2.54).toFixed(3)}cm`;

  // Position and size of a shape, from svg:x/y or, for transformed shapes, the transform's translate()
  const shapeGeometry = (el) => {
    let x = pt(at(el, 'svg', 'x'));
    let y = pt(at(el, 'svg', 'y'));
    const translate = /translate\(\s*([-\d.]+[a-z]*)[\s,]+([-\d.]+[a-z]*)\s*\)/i.exec(at(el, 'draw', 'transform') || '');
    if (translate) {
      x += pt(translate[1]);
      y += pt(translate[2]);
    }
    return { x, y, w: pt(at(el, 'svg', 'width')), h: pt(at(el, 'svg', 'height')) };
  };

  // The things on a drawing page Wurd can keep, in drawing order (later ones are on top)
  const collectShapes = (parent, ctx, out) => {
    for (const el of elementsOf(parent)) {
      if (is(el, 'draw', 'g')) {
        collectShapes(el, ctx, out);
        continue;
      }
      const stroke = ctx.styles.resolve('graphic', at(el, 'draw', 'style-name')).stroke;
      if (is(el, 'draw', 'frame')) {
        const textBox = findChild(el, 'draw', 'text-box');
        const image = elementsOf(el).find((c) => is(c, 'draw', 'image') && ctx.images.has(c));
        const table = findChild(el, 'table', 'table');
        // Text frames have no outline unless their style gives them one
        if (textBox) out.push({ ...shapeGeometry(el), nodes: elementsOf(textBox), plain: !stroke || stroke === 'none' });
        else if (image) out.push({ ...shapeGeometry(el), image: ctx.images.get(image), plain: true });
        else if (table) out.push({ ...shapeGeometry(el), nodes: [table], plain: true });
        else ctx.skipped++;
      } else if (el.namespaceURI === NS.draw && SHAPES.test(el.localName)) {
        // A shape keeps its text; the shape itself (a rectangle, an arrow...) can't be drawn in Wurd
        const nodes = elementsOf(el).filter((c) => is(c, 'text', 'p') || is(c, 'text', 'h') || is(c, 'text', 'list'));
        if (nodes.some((n) => n.textContent.trim())) out.push({ ...shapeGeometry(el), nodes, plain: stroke === 'none' });
        else ctx.skipped++;
      }
      // Notes, forms and the like have nothing to show
    }
    return out;
  };

  const importOdg = async (bytes, { askPassword } = {}) => {
    const pkg = await F.odfPackage.read(bytes, { askPassword, mimePattern: ODG_TYPES, kind: 'drawing' });
    const contentXml = await pkg.readText('content.xml');
    if (!contentXml) throw new Wurd.UserError('This OpenDocument file has no content.');
    const stylesXmlText = await pkg.readText('styles.xml');
    const contentDoc = parseXml(contentXml);
    const stylesDoc = stylesXmlText ? parseXml(stylesXmlText) : null;
    const styles = buildStyleIndex(contentDoc, stylesDoc);
    const page = pageGeometry(stylesDoc);
    const ctx = { styles, pageFont: styles.pageFont, page, images: await loadImages(contentDoc, pkg), pending: [], skipped: 0 };
    const drawing = contentDoc.getElementsByTagNameNS(NS.office, 'drawing')[0];
    if (!drawing) throw new Wurd.UserError("This file isn't an OpenDocument drawing.");
    const container = document.createElement('div');
    ctx.root = container;
    const pages = elementsOf(drawing).filter((el) => is(el, 'draw', 'page'));
    pages.forEach((drawPage, i) => {
      if (i) container.append(pageBreak());
      for (const shape of collectShapes(drawPage, ctx, [])) {
        const box = document.createElement('div');
        box.className = 'wurd-textbox';
        if (shape.image) {
          const p = document.createElement('p');
          const img = document.createElement('img');
          img.setAttribute('src', shape.image);
          img.setAttribute('alt', '');
          if (shape.w) img.setAttribute('width', String(Math.round((shape.w * 4) / 3)));
          if (shape.h) img.setAttribute('height', String(Math.round((shape.h * 4) / 3)));
          p.append(img);
          box.append(p);
        } else {
          convertBlocks(shape.nodes, box, ctx);
          flushPending(ctx, box);
        }
        if (!box.hasChildNodes()) box.append(document.createElement('p'));
        Wurd.setBoxGeometry(box, {
          floating: true,
          plain: shape.plain,
          leftPct: (shape.x / page.widthPt) * 100,
          topPt: shape.y - DRAWING_MARGIN_PT,
          widthPct: (shape.w / page.widthPt) * 100,
          minHeightPt: shape.image ? 0 : shape.h,
        });
        container.append(box);
      }
      // The page's own text goes here; it keeps the page as tall as the drawing's
      const body = document.createElement('p');
      body.setAttribute('style', `margin-top: 0; min-height: ${Math.round(Math.max(page.heightPt - 2 * DRAWING_MARGIN_PT, 72))}pt`);
      body.append(document.createElement('br'));
      container.append(body);
    });
    const left = ctx.skipped;
    return {
      html: container.innerHTML || '<p><br></p>',
      font: styles.pageFont,
      password: pkg.password,
      notice: left ? `Loaded. ${left} shape${left === 1 ? '' : 's'} without text ${left === 1 ? 'was' : 'were'} left out.` : undefined,
    };
  };

  // A free box as a frame on a drawing page. topPt comes from the editor's layout when saving from
  // the page (see Wurd.textboxes.layout); without it, the box's own offset is used.
  const drawingBoxXml = (box, ctx, topPt) => {
    const n = ++ctx.boxCount;
    const pageWidthPt = Wurd.PAGE_SIZES[ctx.pageSizeId].widthPt;
    const plain = box.classList.contains('wurd-textbox--plain');
    const content = blocksXml(box.childNodes, ctx, { parent: 'Standard', align: '', inList: false, indentPt: 0 });
    const y = (Number.isFinite(topPt) ? topPt : pt(box.style.marginTop)) + DRAWING_MARGIN_PT;
    const minHeight = Math.max(pt(box.style.minHeight), 14);
    return `<draw:frame draw:style-name="${plain ? 'WurdPlain' : 'WurdBox'}" draw:name="WurdTextBox${n}" draw:z-index="${n}" svg:x="${ptToCm(percentOf(box.style.left, 0) * pageWidthPt)}" svg:y="${ptToCm(y)}" svg:width="${ptToCm(percentOf(box.style.width, 0.4) * pageWidthPt)}" svg:height="${ptToCm(minHeight)}"><draw:text-box fo:min-height="${ptToCm(minHeight)}">${content || '<text:p/>'}</draw:text-box></draw:frame>`;
  };

  const exportOdg = async (root, { title, pageFont, pageSize, password, boxLayout = null }) => {
    const size = Wurd.PAGE_SIZES[pageSize] ? pageSize : 'a4';
    const sheet = Wurd.PAGE_SIZES[size];
    const margin = DRAWING_MARGIN_PT;
    const textWidthPt = sheet.widthPt - 2 * margin;
    const page = { width: sheet.widthCm, height: sheet.heightCm, textWidthCm: (textWidthPt / 72) * 2.54 };
    const ctx = { ...(await newExportContext(root, pageFont, page)), drawing: true, pageSizeId: size };
    // Pages: split at the page breaks between top-level blocks
    const sections = [{ flow: [], boxes: [] }];
    for (const node of root.childNodes) {
      if (node.nodeType === Node.ELEMENT_NODE && Wurd.isPageBreak(node)) sections.push({ flow: [], boxes: [] });
      else if (node.nodeType === Node.ELEMENT_NODE && Wurd.isFloatingBox(node)) sections[sections.length - 1].boxes.push(node);
      else sections[sections.length - 1].flow.push(node);
    }
    let boxIndex = 0;
    const pages = sections.map((section, i) => {
      const text = blocksXml(section.flow, ctx, { parent: 'Text_20_body', align: '', inList: false, indentPt: 0 });
      const frames = section.boxes.map((box) => drawingBoxXml(box, ctx, boxLayout?.[boxIndex++]?.topPt)).join('');
      const body = `<draw:frame draw:style-name="WurdBody" draw:name="Page text ${i + 1}" draw:z-index="0" svg:x="${ptToCm(margin)}" svg:y="${ptToCm(margin)}" svg:width="${ptToCm(textWidthPt)}" svg:height="${ptToCm(sheet.heightPt - 2 * margin)}"><draw:text-box>${text || '<text:p/>'}</draw:text-box></draw:frame>`;
      return `<draw:page draw:name="page${i + 1}" draw:style-name="dp1" draw:master-page-name="Default">${body}${frames}</draw:page>`;
    });
    const graphic = (name, props) => `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties draw:textarea-vertical-align="top" ${props}/></style:style>`;
    const automatic = [
      automaticStylesXml(ctx),
      '<style:style style:name="dp1" style:family="drawing-page"/>',
      graphic('WurdBody', 'draw:stroke="none" draw:fill="none" draw:auto-grow-height="false" fo:padding-top="0cm" fo:padding-bottom="0cm" fo:padding-left="0cm" fo:padding-right="0cm"'),
      graphic('WurdBox', 'draw:stroke="solid" svg:stroke-color="#808080" svg:stroke-width="0.018cm" draw:fill="solid" draw:fill-color="#ffffff" draw:auto-grow-height="true" fo:padding-top="0.2cm" fo:padding-bottom="0.2cm" fo:padding-left="0.25cm" fo:padding-right="0.25cm"'),
      graphic('WurdPlain', 'draw:stroke="none" draw:fill="none" draw:auto-grow-height="true" fo:padding-top="0.1cm" fo:padding-bottom="0.1cm" fo:padding-left="0.1cm" fo:padding-right="0.1cm"'),
      `<text:list-style style:name="WurdL1">${listLevels('bullet')}</text:list-style>`,
      `<text:list-style style:name="WurdL2">${listLevels('number')}</text:list-style>`,
    ].join('');
    const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NAMESPACES} office:version="${F.odfPackage.ODF_VERSION}">
${fontFaceDecls()}
<office:automatic-styles>${automatic}</office:automatic-styles>
<office:body><office:drawing>${pages.join('')}</office:drawing></office:body>
</office:document-content>
`;
    const parts = [
      { path: 'content.xml', data: content, mediaType: 'text/xml' },
      { path: 'styles.xml', data: drawingStylesXml(pageFont, page), mediaType: 'text/xml' },
      { path: 'meta.xml', data: metaXml(title), mediaType: 'text/xml' },
      ...Array.from(ctx.images.values(), (img) => ({ path: img.path, data: img.bytes, mediaType: img.mediaType })),
    ];
    return F.odfPackage.write(parts, { password, mimeType: ODG_MIME });
  };

  F.odt = { import: importOdt, export: exportOdt };
  F.odg = { import: importOdg, export: exportOdg };
})(window.Wurd = window.Wurd || {});
