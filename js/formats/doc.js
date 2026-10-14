// Legacy Word .doc. Reading: binary Word 97-2003 files give their text (formatting isn't recovered);
// RTF, MHT and HTML files named .doc are recognised too. Writing: an HTML-based .doc that Word and LibreOffice open.
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});
  const root = typeof window !== 'undefined' ? window : globalThis;

  const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];
  const startsWith = (bytes, signature) => signature.every((b, i) => bytes[i] === b);

  // ---- Binary .doc text (MS-DOC: FIB -> Clx -> piece table) ------------------

  // 8-bit text is Latin-1 except for these code points (MS-DOC 2.4.1 FcCompressed)
  const FC_COMPRESSED = {
    0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030,
    0x8a: 0x0160, 0x8b: 0x2039, 0x8c: 0x0152, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022,
    0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153, 0x9f: 0x0178,
  };

  class DocError extends Error {
    constructor(kind) {
      super(kind);
      this.kind = kind;
    }
  }

  const toU8 = (content) => (content instanceof Uint8Array ? content : Uint8Array.from(content || []));

  // Returns the main document text of a binary .doc. Paragraph ends become blank lines, line breaks "\n".
  // Pure function of the bytes (needs the CFB library), so it can be tested outside the browser.
  const readDocText = (bytes) => {
    const CFB = root.CFB;
    let cfb;
    try {
      cfb = CFB.read(bytes, { type: 'array' });
    } catch (e) {
      throw new DocError('corrupt');
    }
    if (CFB.find(cfb, '/EncryptedPackage')) throw new DocError('encrypted');
    // Absolute path: embedded objects in ObjectPool have WordDocument streams of their own
    const wordEntry = CFB.find(cfb, '/WordDocument');
    if (!wordEntry) throw new DocError('not-word');
    const wd = toU8(wordEntry.content);
    if (wd.length < 0x1aa) throw new DocError('corrupt');
    const fib = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    if (fib.getUint16(0x00, true) !== 0xa5ec) throw new DocError('corrupt'); // FibBase.wIdent
    if (fib.getUint16(0x02, true) < 0x00c1) throw new DocError('word95'); // nFib before Word 97
    const flags = fib.getUint16(0x0a, true);
    if (flags & 0x0100) throw new DocError('encrypted'); // fEncrypted
    const tableEntry = CFB.find(cfb, flags & 0x0200 ? '/1Table' : '/0Table'); // fWhichTblStm
    if (!tableEntry) throw new DocError('corrupt');
    const table = toU8(tableEntry.content);
    const tbl = new DataView(table.buffer, table.byteOffset, table.byteLength);

    const csw = fib.getUint16(0x20, true);
    const rgLw = 0x22 + csw * 2 + 2; // FibRgLw97 (0x40 in practice)
    const cslw = fib.getUint16(rgLw - 2, true);
    const rgFcLcb = rgLw + cslw * 4 + 2; // FibRgFcLcb97 (0x9A in practice)
    if (rgFcLcb + 34 * 8 > wd.length) throw new DocError('corrupt');
    const ccpText = fib.getInt32(rgLw + 0x0c, true); // characters in the main document
    const fcClx = fib.getUint32(rgFcLcb + 33 * 8, true);
    const lcbClx = fib.getUint32(rgFcLcb + 33 * 8 + 4, true);
    if (ccpText < 0 || lcbClx < 5 || fcClx + lcbClx > table.length) throw new DocError('corrupt');

    // Clx: skip Prc entries (0x01), then the Pcdt (0x02) holds the piece table
    let p = fcClx;
    const clxEnd = fcClx + lcbClx;
    while (p < clxEnd && table[p] === 0x01) {
      if (p + 3 > clxEnd) throw new DocError('corrupt');
      p += 3 + tbl.getInt16(p + 1, true);
    }
    if (p + 5 > clxEnd || table[p] !== 0x02) throw new DocError('corrupt');
    const lcb = tbl.getUint32(p + 1, true);
    const plc = p + 5;
    const pieces = (lcb - 4) / 12;
    if (!Number.isInteger(pieces) || pieces < 1 || plc + lcb > table.length) throw new DocError('corrupt');

    let text = '';
    for (let i = 0; i < pieces; i++) {
      const cpStart = tbl.getUint32(plc + 4 * i, true);
      if (cpStart >= ccpText) break; // headers, footnotes and comments come after the main text
      const cpEnd = Math.min(tbl.getUint32(plc + 4 * (i + 1), true), ccpText);
      const count = cpEnd - cpStart;
      if (count <= 0) continue;
      const fcRaw = tbl.getUint32(plc + 4 * (pieces + 1) + 8 * i + 2, true);
      if (fcRaw & 0x40000000) {
        const start = (fcRaw & 0x3fffffff) >>> 1;
        if (start + count > wd.length) throw new DocError('corrupt');
        for (let k = 0; k < count; k++) {
          const b = wd[start + k];
          text += String.fromCharCode(FC_COMPRESSED[b] || b);
        }
      } else {
        const start = fcRaw & 0x3fffffff;
        if (start + count * 2 > wd.length) throw new DocError('corrupt');
        text += new TextDecoder('utf-16le').decode(wd.subarray(start, start + count * 2));
      }
    }

    // Fields are {0x13 instructions 0x14 result 0x15}: keep only results (a link's text, not HYPERLINK "...")
    let out = '';
    const fields = [];
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      if (c === 0x13) {
        fields.push(false);
        continue;
      }
      if (c === 0x14) {
        if (fields.length) fields[fields.length - 1] = true;
        continue;
      }
      if (c === 0x15) {
        fields.pop();
        continue;
      }
      if (fields.includes(false)) continue;
      if (c === 0x0d || c === 0x07 || c === 0x0c || c === 0x0e) out += '\n\n'; // paragraph, cell/row end, page/column break
      else if (c === 0x0b) out += '\n'; // line break
      else if (c === 0x09) out += '\t';
      else if (c === 0x1e) out += String.fromCharCode(0x2011); // non-breaking hyphen
      else if (c >= 0x20) out += text[i];
      // 0x1f optional hyphen, 0x01/0x08 objects and other controls are dropped
    }
    return out;
  };

  const DOC_MESSAGES = {
    encrypted: 'This Word document is password-protected. Remove the password in Word or LibreOffice, then open it again.',
    word95: "Word 6 and Word 95 documents aren't supported. Open it in Word or LibreOffice and save it as .docx.",
    corrupt: "This .doc file looks damaged, so Wurd can't read it.",
    'not-word': "This file isn't a Word document.",
  };

  // ---- RTF (sometimes saved with a .doc name): text only ---------------------

  const SKIP_DESTINATIONS = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'object', 'header', 'headerl',
    'headerr', 'headerf', 'footer', 'footerl', 'footerr', 'footerf', 'footnote', 'fldinst', 'listtable', 'listoverridetable',
    'themedata', 'colorschememapping', 'datastore', 'latentstyles', 'xmlnstbl', 'rsidtbl', 'generator', 'filetbl', 'revtbl',
    'listtext', 'pntext', 'pntxta', 'pntxtb', 'bkmkstart', 'bkmkend', 'shppict', 'nonshppict', 'annotation', 'atnid', 'atnauthor']);
  const RTF_SYMBOLS = { emdash: '—', endash: '–', lquote: '‘', rquote: '’', ldblquote: '“', rdblquote: '”', bullet: '•' };

  const rtfToText = (rtf) => {
    let out = '';
    let decoder = new TextDecoder('windows-1252');
    const stack = [];
    let skip = false;
    let uc = 1;
    let fallback = 0; // characters to skip after a \u character
    let i = 0;
    while (i < rtf.length) {
      const ch = rtf[i];
      if (ch === '{') {
        stack.push({ skip, uc });
        i++;
      } else if (ch === '}') {
        const state = stack.pop();
        if (state) ({ skip, uc } = state);
        i++;
      } else if (ch === '\\') {
        const next = rtf[i + 1];
        if (next === '\\' || next === '{' || next === '}') {
          if (!skip) out += next;
          i += 2;
        } else if (next === '~') {
          if (!skip) out += Wurd.NBSP;
          i += 2;
        } else if (next === '_') {
          if (!skip) out += String.fromCharCode(0x2011);
          i += 2;
        } else if (next === '*') {
          skip = true;
          i += 2;
        } else if (next === "'") {
          if (fallback > 0) fallback--;
          else if (!skip) out += decoder.decode(new Uint8Array([parseInt(rtf.substr(i + 2, 2), 16) || 0x3f]));
          i += 4;
        } else if (next === '\n' || next === '\r') {
          if (!skip) out += '\n\n';
          i += 2;
        } else {
          const match = /^([a-zA-Z]{1,32})(-?\d{1,10})? ?/.exec(rtf.slice(i + 1, i + 46));
          if (!match) {
            i += 2;
            continue;
          }
          i += 1 + match[0].length;
          const word = match[1];
          const param = match[2] === undefined ? null : Number(match[2]);
          if (SKIP_DESTINATIONS.has(word)) {
            skip = true;
          } else if (word === 'ansicpg' && param) {
            try {
              decoder = new TextDecoder(`windows-${param}`);
            } catch (e) {
              // Unknown code page: keep Windows-1252
            }
          } else if (!skip) {
            if (word === 'par' || word === 'sect' || word === 'page' || word === 'row') out += '\n\n';
            else if (word === 'line') out += '\n';
            else if (word === 'tab' || word === 'cell') out += '\t';
            else if (word === 'uc') uc = param ?? 1;
            else if (word === 'u') {
              out += String.fromCharCode(param < 0 ? param + 65536 : param);
              fallback = uc;
            } else if (RTF_SYMBOLS[word]) out += RTF_SYMBOLS[word];
          }
        }
      } else if (ch === '\r' || ch === '\n') {
        i++;
      } else {
        if (fallback > 0) fallback--;
        else if (!skip) out += ch;
        i++;
      }
    }
    return out;
  };

  // ---- MHT (Word's "Single File Web Page") ----------------------------------

  const latin1Bytes = (text) => Uint8Array.from(text, (c) => c.charCodeAt(0) & 0xff);

  const decodeTransfer = (body, encoding) => {
    if (/base64/i.test(encoding)) return Wurd.base64ToBytes(body.replace(/\s+/g, ''));
    if (/quoted-printable/i.test(encoding)) {
      return latin1Bytes(body.replace(/=\r?\n/g, '').replace(/=([0-9A-F]{2})/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16))));
    }
    return latin1Bytes(body);
  };

  const mhtToHtml = (text) => {
    const boundary = /boundary="?([^";\r\n]+)"?/i.exec(text)?.[1];
    if (!boundary) return null;
    let html = null;
    const resources = new Map();
    for (const part of text.split(`--${boundary}`).slice(1)) {
      const split = /\r?\n\r?\n/.exec(part);
      if (!split) continue;
      const headers = part.slice(0, split.index);
      const body = part.slice(split.index + split[0].length);
      const header = (name) => new RegExp(`^${name}:\\s*(.+)$`, 'im').exec(headers)?.[1]?.trim() || '';
      const type = header('Content-Type');
      const bytes = decodeTransfer(body, header('Content-Transfer-Encoding'));
      if (/^text\/html/i.test(type) && html === null) {
        const charset = /charset="?([\w-]+)/i.exec(type)?.[1] || 'utf-8';
        try {
          html = new TextDecoder(charset).decode(bytes);
        } catch (e) {
          html = new TextDecoder('utf-8').decode(bytes);
        }
      } else if (/^image\//i.test(type)) {
        const uri = `data:${type.split(';')[0]};base64,${Wurd.bytesToBase64(bytes)}`;
        const location = header('Content-Location');
        const id = header('Content-ID').replace(/^<|>$/g, '');
        if (location) {
          resources.set(location, uri);
          resources.set(location.split('/').pop(), uri);
        }
        if (id) resources.set(`cid:${id}`, uri);
      }
    }
    if (html === null) return null;
    const doc = new DOMParser().parseFromString(html, 'text/html');
    doc.querySelectorAll('img[src]').forEach((img) => {
      const src = img.getAttribute('src');
      const uri = resources.get(src) || resources.get(src.split('/').pop());
      if (uri) img.setAttribute('src', uri);
    });
    return doc.documentElement.outerHTML;
  };

  // ---- Open .doc / .docx ----------------------------------------------------

  const importWord = async (bytes, ext) => {
    if (startsWith(bytes, OLE_MAGIC)) {
      await Wurd.loadScript('lib/cfb.min.js');
      try {
        return { html: F.text.textToHtml(readDocText(bytes)), font: Wurd.DEFAULT_FONT };
      } catch (err) {
        if (err instanceof DocError) throw new Wurd.UserError(DOC_MESSAGES[err.kind]);
        throw err;
      }
    }
    // A .docx, possibly misnamed .doc
    if (startsWith(bytes, ZIP_MAGIC)) return F.docx.import(bytes);
    if (ext === 'docx') throw new Wurd.UserError("This file isn't a valid Word document.");

    const head = new TextDecoder('windows-1252').decode(bytes.subarray(0, 4096));
    if (/^\s*\{\\rtf/.test(head)) {
      const text = rtfToText(new TextDecoder('windows-1252').decode(bytes));
      return { html: F.text.textToHtml(text), font: Wurd.DEFAULT_FONT, notice: 'Opened an RTF document as text' };
    }
    if (/^MIME-Version:/im.test(head.slice(0, 2048))) {
      const html = mhtToHtml(new TextDecoder('windows-1252').decode(bytes));
      if (html) return { html, font: Wurd.DEFAULT_FONT };
    }
    // Case-sensitive: Word-HTML (including Wurd's own .doc files) has a <w:WordDocument> settings block
    if (/^\s*<\?xml[^>]*\?>\s*<\?mso-application\s+progid="Word\.Document"/i.test(head) || /<w:wordDocument[\s>]/.test(head)) {
      throw new Wurd.UserError("Word 2003 XML documents aren't supported. Open it in Word or LibreOffice and save it as .docx.");
    }
    // HTML-based .doc, including the ones Wurd saves (they keep their page font)
    if (/<(html|!doctype|head|body)\b/i.test(head)) return F.html.import(bytes);
    return { html: F.text.textToHtml(Wurd.decodeText(bytes)), font: Wurd.DEFAULT_FONT };
  };

  // ---- Save .doc (Word-HTML) --------------------------------------------------

  const exportDoc = (documentRoot, { title, pageFont }) => {
    const body = documentRoot.cloneNode(true);
    // Word honours table border attributes more reliably than CSS borders
    body.querySelectorAll('table').forEach((table) => {
      table.setAttribute('border', '1');
      table.setAttribute('cellspacing', '0');
      table.setAttribute('cellpadding', '5');
    });
    // Word uses the first font in a list; it has Georgia and Courier New but not Inter
    body.querySelectorAll('font[face]').forEach((font) => {
      if (/^\s*["']?Inter\b/i.test(font.getAttribute('face'))) font.setAttribute('face', 'Arial, sans-serif');
    });
    const esc = Wurd.escapeXml;
    const font = F.docx.wordFont(pageFont);
    // Nothing may come before <html: LibreOffice's format detection checks the start of the file
    const html = `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="ProgId" content="Word.Document">
<meta name="Generator" content="Wurd">
<meta name="wurd-page-font" content="${esc(pageFont)}">
<title>${esc(title)}</title>
<!--[if gte mso 9]><xml><o:DocumentProperties><o:Title>${esc(title)}</o:Title></o:DocumentProperties><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
@page WordSection1 { size: 595.3pt 841.9pt; margin: 72pt; mso-header-margin: 35.4pt; mso-footer-margin: 35.4pt; mso-paper-source: 0; }
div.WordSection1 { page: WordSection1; }
body, p, li, td, th { font-family: "${font}"; font-size: 12pt; }
p { margin: 0 0 8pt; }
h1 { font-size: 27pt; mso-outline-level: 1; }
h2 { font-size: 18pt; mso-outline-level: 2; }
h3 { font-size: 15pt; mso-outline-level: 3; }
h4 { font-size: 12pt; mso-outline-level: 4; }
blockquote { margin: 0 0 8pt 36pt; border-left: solid #D1D5DB 2.25pt; padding-left: 8pt; font-style: italic; }
pre, code { font-family: "Courier New"; font-size: 10pt; }
pre { background: #F3F4F6; }
table { border-collapse: collapse; }
td, th { border: solid #9CA3AF 1pt; padding: 3pt 5pt; vertical-align: top; }
</style>
</head>
<body lang="EN-GB" style="tab-interval: 36pt">
<div class="WordSection1">
${body.innerHTML}
</div>
</body>
</html>`;
    return new Blob([String.fromCharCode(0xfeff) + html], { type: 'application/msword' });
  };

  F.word = { import: importWord, exportDoc, readDocText, rtfToText, mhtToHtml, DocError };
})(typeof window !== 'undefined' ? (window.Wurd = window.Wurd || {}) : (globalThis.Wurd = globalThis.Wurd || {}));
