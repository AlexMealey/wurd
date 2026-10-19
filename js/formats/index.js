// The format registry: which files Wurd opens and writes, and how
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});

  // Export choices, in the order the Export As dialog lists them
  F.EXPORTS = [
    { id: 'docx', label: 'Word document', hint: 'Opens in Microsoft Word, LibreOffice and Google Docs.' },
    { id: 'odt', label: 'OpenDocument text', hint: "Opens in LibreOffice and Word. Can be protected with a password, using LibreOffice's protection, which is weaker than .wurd's.", encryptable: true },
    { id: 'doc', label: 'Word 97–2003 document', hint: 'An HTML-based .doc that Word and LibreOffice open. Images may not show in Word, so use .docx for images.' },
    { id: 'html', label: 'Web page', hint: 'Opens in any web browser.' },
    { id: 'txt', label: 'Plain text', hint: 'Text only: formatting, images and links are left out.' },
    { id: 'wurd', label: 'Encrypted Wurd document', hint: 'Encrypted with AES-256, under a key that takes Argon2id and 64 MB of memory to derive from the password. Only Wurd can open it, and only with the password.', encrypted: true },
  ];

  const MIME_TYPES = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    odt: 'application/vnd.oasis.opendocument.text',
    doc: 'application/msword',
    html: 'text/html;charset=utf-8',
    txt: 'text/plain;charset=utf-8',
    wurd: 'application/json',
  };

  F.OPENABLE = ['docx', 'odt', 'doc', 'html', 'htm', 'txt', 'wurd'];

  F.extensionOf = (name) => {
    const match = /\.([^./\\]+)$/.exec(name || '');
    return match ? match[1].toLowerCase() : '';
  };

  // Reads a file into { html, title, format, font, lock?, notice? }.
  // html is not yet sanitized: the app sanitizes everything it loads.
  F.open = async (file, options = {}) => {
    const ext = F.extensionOf(file.name);
    const bytes = new Uint8Array(await file.arrayBuffer());
    let result;
    switch (ext) {
      case 'txt':
        result = F.text.import(bytes);
        break;
      case 'html':
      case 'htm':
        result = F.html.import(bytes);
        break;
      case 'docx':
      case 'doc':
        result = await F.word.import(bytes, ext);
        break;
      case 'odt':
        result = await F.odt.import(bytes, options);
        break;
      case 'wurd':
        result = await F.wurd.import(bytes, options);
        break;
      default:
        throw new Wurd.UserError(`Wurd can't open .${ext || 'unknown'} files. Open a .docx, .odt, .doc, .html, .txt or .wurd file.`);
    }
    return {
      title: file.name.replace(/\.[^/.]+$/, ''),
      format: ext === 'htm' ? 'html' : ext,
      font: Wurd.DEFAULT_FONT,
      ...result,
    };
  };

  // A lock for a new password, for the formats that encrypt: what the app keeps (instead of the password) so
  // Save can encrypt the document again
  F.lock = async (format, password) => {
    switch (format) {
      case 'wurd':
        return F.wurd.lock(password);
      case 'odt':
        return F.odfPackage.lock(password);
      default:
        return null;
    }
  };

  // Writes the document (editor HTML) in a format. ctx: { title, pageFont, lock? } (a lock from F.lock or F.open)
  F.export = async (format, html, ctx) => {
    const root = Wurd.documentRoot(html);
    let blob;
    switch (format) {
      case 'docx':
        blob = await F.docx.export(root, ctx);
        break;
      case 'odt':
        blob = await F.odt.export(root, ctx);
        break;
      case 'doc':
        blob = F.word.exportDoc(root, ctx);
        break;
      case 'html':
        blob = F.html.export(root, ctx);
        break;
      case 'txt':
        blob = F.text.export(root, ctx);
        break;
      case 'wurd':
        blob = await F.wurd.export(root, ctx);
        break;
      default:
        throw new Wurd.UserError(`Wurd can't save .${format} files.`);
    }
    // Blobs from libraries can come back untyped; downloads should carry the right type
    return blob.type ? blob : new Blob([blob], { type: MIME_TYPES[format] });
  };

  // Header cells are bold anyway; exporters make them bold explicitly for Word, so drop that on the way back
  F.unwrapHeaderBold = (root) => {
    root.querySelectorAll('th').forEach((th) => {
      // Look through a single wrapping paragraph (<th><p><strong>...)
      let only = th.children.length === 1 ? th.firstElementChild : null;
      if (only?.tagName === 'P' && only.children.length === 1) only = only.firstElementChild;
      if (only && /^(STRONG|B)$/.test(only.tagName) && th.textContent.trim() === only.textContent.trim()) only.replaceWith(...only.childNodes);
    });
  };

  F.fileName = (title, format) => `${(title || 'Untitled').replace(/[\\/:*?"<>|]+/g, '_').trim() || 'Untitled'}.${format}`;
})(window.Wurd = window.Wurd || {});
