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

  // mammoth drops alignment, so aligned paragraphs are tagged with a style id (see transformDocument)
  // and these rules, which come first, carry the alignment through as a style attribute
  const buildStyleMap = () => {
    const rules = [];
    for (const align of ['center', 'right', 'justify']) {
      const attr = `[style='text-align: ${align}']`;
      for (const [name, prefix, tag] of BLOCK_STYLES) {
        rules.push(`p.WurdAlign-${align}[style-name='${name}'] => ${prefix}${tag}${attr}:fresh`);
      }
      rules.push(`p.WurdAlign-${align} => p${attr}:fresh`);
    }
    for (const [name, prefix, tag] of BLOCK_STYLES) rules.push(`p[style-name='${name}'] => ${prefix}${tag}:fresh`);
    rules.push(
      // Word's and LibreOffice's built-in character styles for emphasis
      "r[style-name='Strong'] => strong",
      "r[style-name='Strong Emphasis'] => strong",
      "r[style-name='Emphasis'] => em",
      "r.WurdFont-serif => font[face='Georgia, serif']",
      "r.WurdFont-mono => font[face='Courier New, monospace']",
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

  const tidyImported = (html) => {
    const root = Wurd.documentRoot(html);
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

  // Aligned paragraphs get a style id the style map turns into text-align. The style name is kept,
  // so heading and quote rules still match; list items keep their numbering.
  const tagAlignment = (document) => mammoth.transforms.paragraph((paragraph) => {
    const align = ALIGN_IN[paragraph.alignment];
    // Numbered paragraphs are list items and must stay lists, except headings: LibreOffice's heading
    // styles carry (unnumbered) outline numbering
    const listItem = paragraph.numbering && !BLOCK_STYLE_NAMES.has((paragraph.styleName || '').toLowerCase());
    return align && !listItem ? { ...paragraph, styleId: `WurdAlign-${align}` } : paragraph;
  })(document);

  // Runs in a serif or monospace font come back as Wurd's Georgia or Courier fonts
  const tagFonts = (document) => mammoth.transforms.run((run) => {
    if (!run.font || run.styleId) return run;
    const font = Wurd.classifyFont(run.font);
    if (font === Wurd.SERIF_FONT) return { ...run, styleId: 'WurdFont-serif' };
    if (font === Wurd.MONO_FONT) return { ...run, styleId: 'WurdFont-mono' };
    return run;
  })(document);

  const importDocx = async (bytes) => {
    await Wurd.loadScript('lib/mammoth.browser.min.js');
    const options = {
      styleMap: buildStyleMap(),
      ignoreEmptyParagraphs: false,
      transformDocument: (document) => tagFonts(tagAlignment(document)),
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
      html: tidyImported(result.value),
      font: Wurd.DEFAULT_FONT,
      notice: problems ? `Loaded, but ${problems} item${problems === 1 ? '' : 's'} couldn't be converted` : undefined,
    };
  };

  // ---- Export ---------------------------------------------------------------

  // A4 with 1 inch margins: 9026 twips of text width; 1 CSS px = 15 twips
  const PAGE = { width: 11906, height: 16838, margin: 1440 };
  const CONTENT_TWIPS = PAGE.width - 2 * PAGE.margin;
  const MAX_IMAGE_PX = Math.floor(CONTENT_TWIPS / 15);

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

  const toHex = (color) => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(color || '');
    if (m) return m.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
    const hex = /^#([0-9a-f]{6})$/i.exec(color || '');
    return hex ? hex[1].toUpperCase() : undefined;
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
    const itemState = { ...state, list: { level, instance } };
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
    const columnWidth = Math.floor((state.width || CONTENT_TWIPS) / columns);
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

  const blockFrom = (el, ctx, state) => {
    const { D } = ctx;
    const tag = el.tagName;
    const align = alignOf(el, D) ?? state.align;
    if (HEADINGS[tag]) {
      const paragraph = paragraphFrom(Array.from(el.childNodes), ctx, { ...state, quote: 0 }, { heading: D.HeadingLevel[HEADINGS[tag]], align, keepEmpty: true });
      return paragraph ? [paragraph] : [];
    }
    if (tag === 'UL' || tag === 'OL') return listBlocks(el, ctx, state);
    if (tag === 'BLOCKQUOTE') return blocksFrom(el, ctx, { ...state, quote: state.quote + 1, align });
    if (tag === 'PRE') return [preParagraph(el, ctx)];
    if (tag === 'HR') return [new D.Paragraph({ style: 'HorizontalLine', children: [] })];
    if (tag === 'TABLE') return tableFrom(el, ctx, state);
    // P, DIV, LI outside a list and the like: a container if it holds blocks, otherwise one paragraph
    const holdsBlocks = Array.from(el.children).some((child) => BLOCK.test(child.tagName));
    if (holdsBlocks) return blocksFrom(el, ctx, { ...state, align });
    const paragraph = paragraphFrom(Array.from(el.childNodes), ctx, state, { align, keepEmpty: tag === 'P' });
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
        out.push(...blockFrom(node, ctx, state));
      } else if (node.nodeType === Node.TEXT_NODE || node.nodeType === Node.ELEMENT_NODE) {
        inline.push(node);
      }
    }
    flush();
    return out;
  }

  // Images must be decoded before the (synchronous) walk. Only embedded data: images can be exported offline.
  const prepareImages = async (root, D) => {
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
        if (width > MAX_IMAGE_PX) {
          height = Math.round((height * MAX_IMAGE_PX) / width);
          width = MAX_IMAGE_PX;
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

  const exportDocx = async (root, { title, pageFont }) => {
    await Wurd.loadScript('lib/docx.iife.js');
    const D = window.docx;
    const ctx = { D, images: await prepareImages(root, D), listInstance: 0 };
    const children = blocksFrom(root, ctx, { list: null, quote: 0, align: undefined, fmt: {} });
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
            size: { width: PAGE.width, height: PAGE.height },
            margin: { top: PAGE.margin, right: PAGE.margin, bottom: PAGE.margin, left: PAGE.margin },
          },
        },
        children,
      }],
    });
    return D.Packer.toBlob(doc);
  };

  F.docx = { import: importDocx, export: exportDocx, wordFont };
})(window.Wurd = window.Wurd || {});
