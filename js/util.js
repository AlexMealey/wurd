// Shared helpers. Every Wurd script is a classic script that attaches to window.Wurd
// (ES modules don't load from file:// pages, so the app avoids them).
(function (Wurd) {
  'use strict';

  Wurd.NBSP = String.fromCharCode(0xa0);

  // Must match the font <option> values in index.html
  Wurd.DEFAULT_FONT = 'Inter, sans-serif';
  Wurd.SERIF_FONT = 'Georgia, serif';
  Wurd.MONO_FONT = 'Courier New, monospace';

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
