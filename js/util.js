// Shared helpers. Every Wurd script is a classic script that attaches to window.Wurd
// (ES modules don't load from file:// pages, so the app avoids them).
(function (Wurd) {
  'use strict';

  Wurd.NBSP = String.fromCharCode(0xa0);

  // Must match the font <option> values in index.html
  Wurd.DEFAULT_FONT = 'Inter, sans-serif';
  Wurd.SERIF_FONT = 'Georgia, serif';
  Wurd.MONO_FONT = 'Courier New, monospace';

  // Page sizes. Margins are 1in on screen and in .docx/.doc; .odt keeps its 2cm margins.
  Wurd.PAGE_SIZES = {
    a4: { label: 'A4', css: 'A4', widthMm: 210, heightMm: 297, widthTwips: 11906, heightTwips: 16838, widthCm: '21.001cm', heightCm: '29.7cm', widthPt: 595.3, heightPt: 841.9 },
    letter: { label: 'Letter', css: 'letter', widthMm: 215.9, heightMm: 279.4, widthTwips: 12240, heightTwips: 15840, widthCm: '21.59cm', heightCm: '27.94cm', widthPt: 612, heightPt: 792 },
  };
  const PAGE_SIZE_KEY = 'wurd-page-size';
  let pageSize = 'a4';
  try {
    const saved = localStorage.getItem(PAGE_SIZE_KEY);
    if (Wurd.PAGE_SIZES[saved]) pageSize = saved;
  } catch (e) {
    // Storage unavailable: A4
  }
  Wurd.pageSize = () => pageSize;
  Wurd.setPageSize = (size) => {
    pageSize = Wurd.PAGE_SIZES[size] ? size : 'a4';
    try {
      localStorage.setItem(PAGE_SIZE_KEY, pageSize);
    } catch (e) {
      // Storage unavailable: the setting lasts for this visit
    }
  };
  // The page's text width in CSS pixels (1in margins), which caps inserted images
  Wurd.contentWidthPx = (size = pageSize) => Math.floor(((Wurd.PAGE_SIZES[size] || Wurd.PAGE_SIZES.a4).widthMm / 25.4 - 2) * 96);

  // Word's highlight colours (w:highlight names) as hex
  Wurd.HIGHLIGHTS = {
    yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
    darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
    darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
  };

  // CSS colour (#rgb, #rrggbb, rgb()/rgba(), or a highlight name) -> 'RRGGBB'; undefined for transparent or unknown
  Wurd.colorToHex = (color) => {
    const value = String(color || '').trim();
    const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/i.exec(value);
    if (rgb) {
      if (rgb[4] !== undefined && Number(rgb[4]) === 0) return undefined;
      return rgb.slice(1, 4).map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0')).join('').toUpperCase();
    }
    const hex = /^#?([0-9a-f]{6})$/i.exec(value);
    if (hex) return hex[1].toUpperCase();
    const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(value);
    if (short) return short.slice(1, 4).map((c) => c + c).join('').toUpperCase();
    const named = Object.keys(Wurd.HIGHLIGHTS).find((name) => name.toLowerCase() === value.toLowerCase());
    return named ? Wurd.HIGHLIGHTS[named] : undefined;
  };

  // A CSS length (pt, px, in, cm, mm, pc, em) -> points; NaN when it can't be read
  Wurd.cssLengthToPt = (length) => {
    const m = /^(-?[\d.]+)(pt|px|in|cm|mm|pc|em|rem)?$/i.exec(String(length || '').trim());
    if (!m) return NaN;
    const n = Number(m[1]);
    const factor = { pt: 1, px: 0.75, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12, em: 12, rem: 12 }[(m[2] || 'px').toLowerCase()];
    return n * factor;
  };

  // Indentation step for the indent buttons and imports (0.5in, Word's default)
  Wurd.INDENT_PT = 36;
  Wurd.MAX_INDENT_STEPS = 8;

  // Points of left indent on a block (margin-left / margin-inline-start), 0 when none
  Wurd.indentOf = (el) => {
    const value = el?.style?.marginLeft || el?.style?.marginInlineStart || '';
    const pt = Wurd.cssLengthToPt(value);
    return Number.isFinite(pt) && pt > 0 ? pt : 0;
  };

  // A page break and a text box, as the editor writes them
  Wurd.isPageBreak = (el) => Boolean(el && el.tagName === 'HR' && (el.classList.contains('wurd-page-break')
    || /page/.test(`${el.style?.breakAfter || ''} ${el.style?.breakBefore || ''}`)
    || /always/.test(`${el.style?.pageBreakAfter || ''} ${el.style?.pageBreakBefore || ''}`)));
  Wurd.isTextBox = (el) => Boolean(el && el.tagName === 'DIV' && el.classList.contains('wurd-textbox'));

  // Sets a text box's size and placement from an import: { floating, leftPct, topPt, widthPct, minHeightPt }.
  // A free box's left and width are shares of the page width, an in-line box's width of the text column.
  Wurd.setBoxGeometry = (box, info = {}) => {
    const r = (n, places = 2) => Math.round(n * 10 ** places) / 10 ** places;
    const positive = (n) => Number.isFinite(n) && n > 0;
    // A box with no outline (as a drawing's text frames usually are)
    if (info.plain) box.classList.add('wurd-textbox--plain');
    if (info.floating) {
      box.classList.add('wurd-textbox--float');
      box.style.left = `${r(Math.min(Math.max(0, info.leftPct || 0), 95))}%`;
      box.style.marginTop = `${r(Number.isFinite(info.topPt) ? info.topPt : 0, 1)}pt`;
      box.style.width = `${r(Math.min(positive(info.widthPct) ? info.widthPct : 40, 100))}%`;
    } else if (positive(info.widthPct) && info.widthPct < 99.5) {
      box.style.width = `${r(info.widthPct)}%`;
    }
    if (positive(info.minHeightPt)) box.style.minHeight = `${r(info.minHeightPt, 1)}pt`;
  };

  // An error whose message is written for the person using the app
  class UserError extends Error {}
  Wurd.UserError = UserError;

  // Thrown when someone cancels a password prompt; callers stop quietly
  class Cancelled extends Error {}
  Wurd.Cancelled = Cancelled;

  // Loads a vendored library on first use. A <script> element can load from file://, unlike fetch.
  const scripts = new Map();
  Wurd.loadScript = (src) => {
    if (!scripts.has(src)) {
      scripts.set(src, new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = () => {
          scripts.delete(src);
          reject(new UserError(`Couldn't load ${src}. Check the lib folder is next to index.html.`));
        };
        document.head.append(script);
      }));
    }
    return scripts.get(src);
  };

  // Strip scripts, event handlers and javascript: URLs before HTML goes into the editor.
  // <style> is dropped too, since it would restyle the whole app rather than just the document.
  Wurd.sanitizeHtml = (html) => DOMPurify.sanitize(html, { FORBID_TAGS: ['style'] });

  // A detached, sanitized copy of the document for exporters to walk
  Wurd.documentRoot = (html) => {
    const root = document.createElement('div');
    root.innerHTML = Wurd.sanitizeHtml(html);
    return root;
  };

  Wurd.bytesToBase64 = (bytes) => {
    let binary = '';
    for (let i = 0; i < bytes.length; i += 0x8000) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(binary);
  };

  Wurd.base64ToBytes = (base64) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));

  // Text files: honour a byte-order mark, otherwise UTF-8, falling back to Windows-1252
  Wurd.decodeText = (bytes) => {
    if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder('utf-8').decode(bytes.subarray(3));
    if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (e) {
      return new TextDecoder('windows-1252').decode(bytes);
    }
  };

  // U+FFFE and U+FFFF aren't allowed in XML (built from char codes to keep the source ASCII)
  const NON_CHARACTERS = new RegExp('[' + String.fromCharCode(0xfffe, 0xffff) + ']', 'g');

  // Escapes text for XML and drops characters XML 1.0 can't contain
  Wurd.escapeXml = (text) => String(text)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(NON_CHARACTERS, '')
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  Wurd.downloadBlob = (blob, fileName) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  };

  // Web Crypto only exists in secure contexts (file:// counts as one in current browsers)
  Wurd.cryptoAvailable = () => Boolean(window.isSecureContext && window.crypto && window.crypto.subtle);

  // Collapses whitespace the way HTML rendering does, keeping non-breaking spaces
  Wurd.collapseSpaces = (text) => text.replace(/[ \t\n\r\f]+/g, ' ');

  // CSS font-family list -> the first family name, unquoted
  Wurd.firstFontFamily = (css) => (css || '').split(',')[0].trim().replace(/^["']|["']$/g, '');

  // Which of the three Wurd fonts a font name is closest to
  Wurd.classifyFont = (name, { pitch, generic } = {}) => {
    if (pitch === 'fixed' || generic === 'modern' || /mono|courier|consol|menlo|code/i.test(name || '')) return Wurd.MONO_FONT;
    if (generic === 'roman' || /serif|georgia|times|cambria|garamond|book|palatino/i.test(name || '')) {
      if (!/sans/i.test(name || '')) return Wurd.SERIF_FONT;
    }
    return Wurd.DEFAULT_FONT;
  };
})(window.Wurd = window.Wurd || {});
