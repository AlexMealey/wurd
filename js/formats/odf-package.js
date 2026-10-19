// OpenDocument packages: the zip container, its manifest, and ODF password encryption.
// Writing uses per-file AES-256-CBC with PBKDF2 (ODF 1.2/1.3), which every LibreOffice since 3.4 opens.
// Reading also handles AES-GCM and LibreOffice 24.8+'s default whole-package encryption with Argon2id.
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});

  const ODF_VERSION = '1.3';
  const MANIFEST_NS = 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0';
  const ODT_MIME = 'application/vnd.oasis.opendocument.text';
  const PBKDF2_ITERATIONS = 100000; // what LibreOffice itself writes for AES-CBC
  const MAX_ITERATIONS = 10000000; // refuse hostile files that would hang the tab
  const MAX_ARGON2_MEMORY_KIB = 1024 * 1024;

  const URI = {
    aes256cbc: 'http://www.w3.org/2001/04/xmlenc#aes256-cbc',
    sha256StartKey: 'http://www.w3.org/2000/09/xmldsig#sha256',
    sha256Checksum: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k',
  };

  // Thrown when a password doesn't decrypt the document
  class WrongPassword extends Error {}

  const utf8 = new TextEncoder();
  const subtle = () => {
    if (!Wurd.cryptoAvailable()) {
      throw new Wurd.UserError("Password-protected documents need the browser's Web Crypto, which isn't available here.");
    }
    return window.crypto.subtle;
  };

  const viaStream = async (bytes, transform) => new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(transform)).arrayBuffer());
  const deflateRaw = (bytes) => viaStream(bytes, new CompressionStream('deflate-raw'));
  const inflateRaw = (bytes) => viaStream(bytes, new DecompressionStream('deflate-raw'));

  const sameBytes = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);

  // ---- Writing --------------------------------------------------------------

  const encryptEntry = async (plain, startKey) => {
    const crypto = subtle();
    const compressed = await deflateRaw(plain);
    // ODF checksum: SHA-256 of the first 1024 bytes of the compressed, unencrypted data
    const checksum = new Uint8Array(await crypto.digest('SHA-256', compressed.subarray(0, 1024)));
    const salt = window.crypto.getRandomValues(new Uint8Array(16));
    const iv = window.crypto.getRandomValues(new Uint8Array(16));
    // PBKDF2 with HMAC-SHA1 (ODF's definition), keyed with the SHA-256 "start key" of the password
    const base = await crypto.importKey('raw', startKey, 'PBKDF2', false, ['deriveBits']);
    const bits = await crypto.deriveBits({ name: 'PBKDF2', hash: 'SHA-1', salt, iterations: PBKDF2_ITERATIONS }, base, 256);
    const key = await crypto.importKey('raw', bits, 'AES-CBC', false, ['encrypt']);
    // Web Crypto's PKCS#7 padding is a valid case of the W3C padding LibreOffice expects
    const stored = new Uint8Array(await crypto.encrypt({ name: 'AES-CBC', iv }, key, compressed));
    return { stored, info: { size: plain.length, checksum, salt, iv } };
  };

  const buildManifest = (entries, mimeType) => {
    const esc = Wurd.escapeXml;
    const b64 = Wurd.bytesToBase64;
    const lines = entries.map(({ path, mediaType, info }) => {
      if (!info) return ` <manifest:file-entry manifest:full-path="${esc(path)}" manifest:media-type="${mediaType}"/>`;
      return [
        ` <manifest:file-entry manifest:full-path="${esc(path)}" manifest:media-type="${mediaType}" manifest:size="${info.size}">`,
        `  <manifest:encryption-data manifest:checksum-type="${URI.sha256Checksum}" manifest:checksum="${b64(info.checksum)}">`,
        `   <manifest:algorithm manifest:algorithm-name="${URI.aes256cbc}" manifest:initialisation-vector="${b64(info.iv)}"/>`,
        `   <manifest:start-key-generation manifest:start-key-generation-name="${URI.sha256StartKey}" manifest:key-size="32"/>`,
        // key-size="32" is required: the default of 16 would mean AES-128
        `   <manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:key-size="32" manifest:iteration-count="${PBKDF2_ITERATIONS}" manifest:salt="${b64(info.salt)}"/>`,
        '  </manifest:encryption-data>',
        ' </manifest:file-entry>',
      ].join('\n');
    });
    return `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="${MANIFEST_NS}" manifest:version="${ODF_VERSION}">
 <manifest:file-entry manifest:full-path="/" manifest:version="${ODF_VERSION}" manifest:media-type="${mimeType}"/>
${lines.join('\n')}
</manifest:manifest>
`;
  };

  // ODF requires the first zip entry to be "mimetype", stored uncompressed, with no extra field
  const checkMimetypeEntry = (zipBytes, mimeType) => {
    const dv = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);
    const name = new TextDecoder().decode(zipBytes.subarray(30, 38));
    const data = new TextDecoder().decode(zipBytes.subarray(38, 38 + mimeType.length));
    const ok = dv.getUint32(0, true) === 0x04034b50 && dv.getUint16(6, true) === 0 && dv.getUint16(8, true) === 0
      && dv.getUint16(26, true) === 8 && dv.getUint16(28, true) === 0 && name === 'mimetype' && data === mimeType;
    if (!ok) throw new Error('The OpenDocument package was written incorrectly (mimetype entry).');
  };

  // ODF's "start key", SHA-256 of the password: all that writing needs. The app keeps this lock, not the password.
  const createLock = async (password) => {
    const bytes = utf8.encode(password);
    try {
      return { format: 'odt', startKey: new Uint8Array(await subtle().digest('SHA-256', bytes)) };
    } finally {
      bytes.fill(0);
    }
  };

  // parts: [{ path, data: string | Uint8Array, mediaType }]; startKey (from a lock) encrypts it
  const writePackage = async (parts, { startKey = null, mimeType = ODT_MIME } = {}) => {
    await Wurd.loadScript('lib/jszip.min.js');
    const zip = new window.JSZip();
    zip.file('mimetype', mimeType, { compression: 'STORE', createFolders: false });
    const entries = [];
    for (const part of parts) {
      const bytes = typeof part.data === 'string' ? utf8.encode(part.data) : part.data;
      if (startKey) {
        // Everything except mimetype and the manifest is encrypted, including meta.xml (it holds the title)
        const { stored, info } = await encryptEntry(bytes, startKey);
        zip.file(part.path, stored, { compression: 'STORE', createFolders: false, binary: true });
        entries.push({ ...part, info });
      } else {
        zip.file(part.path, bytes, { createFolders: false, binary: true });
        entries.push(part);
      }
    }
    zip.file('META-INF/manifest.xml', buildManifest(entries, mimeType), { createFolders: false });
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType });
    checkMimetypeEntry(bytes, mimeType);
    return new Blob([bytes], { type: mimeType });
  };

  // ---- Reading --------------------------------------------------------------

  // Attribute by local name, whatever namespace prefix the writer used (manifest:, loext:, ...)
  const attr = (el, localName) => {
    if (!el) return null;
    for (const a of el.attributes) if (a.localName === localName) return a.value;
    return null;
  };
  const child = (el, localName) => el && Array.from(el.children).find((c) => c.localName === localName);

  const parseManifest = (xml) => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const entries = new Map();
    for (const entry of doc.getElementsByTagNameNS(MANIFEST_NS, 'file-entry')) {
      const path = attr(entry, 'full-path');
      const data = child(entry, 'encryption-data');
      let encryption = null;
      if (data) {
        const algorithm = child(data, 'algorithm');
        const startKey = child(data, 'start-key-generation');
        const kdf = child(data, 'key-derivation');
        encryption = {
          checksumType: attr(data, 'checksum-type') || '',
          checksum: attr(data, 'checksum') ? Wurd.base64ToBytes(attr(data, 'checksum')) : null,
          algorithm: attr(algorithm, 'algorithm-name') || '',
          iv: attr(algorithm, 'initialisation-vector') ? Wurd.base64ToBytes(attr(algorithm, 'initialisation-vector')) : null,
          startKey: attr(startKey, 'start-key-generation-name') || 'SHA1',
          kdf: attr(kdf, 'key-derivation-name') || '',
          keySize: Number(attr(kdf, 'key-size')) || 16,
          iterations: Number(attr(kdf, 'iteration-count')) || 0,
          salt: attr(kdf, 'salt') ? Wurd.base64ToBytes(attr(kdf, 'salt')) : null,
          argon2: {
            iterations: Number(attr(kdf, 'argon2-iterations')) || 0,
            memory: Number(attr(kdf, 'argon2-memory')) || 0,
            lanes: Number(attr(kdf, 'argon2-lanes')) || 0,
          },
          pgp: Boolean(child(data, 'keyinfo')) || /pgp/i.test(attr(kdf, 'key-derivation-name') || ''),
        };
      }
      entries.set(path, { path, mediaType: attr(entry, 'media-type'), size: Number(attr(entry, 'size')) || null, encryption });
    }
    return entries;
  };

  // Formats Wurd can't decrypt get a clear message before any password is asked for
  const checkSupported = (enc) => {
    if (enc.pgp) throw new Wurd.UserError('This document is encrypted to an OpenPGP key, which Wurd can\'t open. Open it in LibreOffice.');
    if (/blowfish/i.test(enc.algorithm)) {
      throw new Wurd.UserError("This document uses old Blowfish encryption, which Wurd can't open. Open it in LibreOffice and save it again with a password.");
    }
    const cipher = /aes(128|192|256)-(cbc|gcm)$/i.exec(enc.algorithm);
    if (!cipher) throw new Wurd.UserError("This document's encryption method isn't supported.");
    if (cipher[1] === '192') throw new Wurd.UserError("This document uses AES-192 encryption, which browsers don't support.");
    if (!/pbkdf2|argon2id/i.test(enc.kdf)) throw new Wurd.UserError("This document's password protection method isn't supported.");
    if (/pbkdf2/i.test(enc.kdf) && (!enc.iterations || enc.iterations > MAX_ITERATIONS)) throw new Wurd.UserError("This document's encryption settings aren't valid.");
    if (/argon2id/i.test(enc.kdf) && (!enc.argon2.memory || enc.argon2.memory > MAX_ARGON2_MEMORY_KIB || enc.argon2.iterations > 100)) {
      throw new Wurd.UserError("This document's encryption settings aren't valid.");
    }
    return cipher[2].toLowerCase();
  };

  const deriveKey = async (enc, password, cache, mode) => {
    const crypto = subtle();
    const id = [enc.startKey, enc.kdf, enc.keySize, enc.iterations, enc.salt && Wurd.bytesToBase64(enc.salt),
      enc.argon2.iterations, enc.argon2.memory, enc.argon2.lanes, mode].join('|');
    if (cache.has(id)) return cache.get(id);
    const startHash = /sha256/i.test(enc.startKey) ? 'SHA-256' : 'SHA-1';
    const startKey = new Uint8Array(await crypto.digest(startHash, utf8.encode(password)));
    let raw;
    if (/argon2id/i.test(enc.kdf)) {
      await Wurd.loadScript('lib/argon2.umd.min.js');
      raw = await window.hashwasm.argon2id({
        password: startKey,
        salt: enc.salt,
        iterations: enc.argon2.iterations || 3,
        parallelism: enc.argon2.lanes || 4,
        memorySize: enc.argon2.memory,
        hashLength: enc.keySize,
        outputType: 'binary',
      });
    } else {
      const base = await crypto.importKey('raw', startKey, 'PBKDF2', false, ['deriveBits']);
      raw = new Uint8Array(await crypto.deriveBits({ name: 'PBKDF2', hash: 'SHA-1', salt: enc.salt, iterations: enc.iterations }, base, enc.keySize * 8));
    }
    const key = await crypto.importKey('raw', raw, mode === 'cbc' ? 'AES-CBC' : 'AES-GCM', false, mode === 'cbc' ? ['encrypt', 'decrypt'] : ['decrypt']);
    cache.set(id, key);
    return key;
  };

  // Web Crypto insists on PKCS#7 padding, but LibreOffice pads with random bytes (W3C/ISO 10126).
  // Appending one block that decrypts to a full PKCS#7 padding block lets Web Crypto succeed;
  // the document's own padding is then removed using its last byte.
  const decryptCbcUnpadded = async (key, iv, data) => {
    const crypto = subtle();
    const lastBlock = data.subarray(data.length - 16);
    const extra = new Uint8Array(await crypto.encrypt({ name: 'AES-CBC', iv: lastBlock }, key, new Uint8Array(16).fill(16))).subarray(0, 16);
    const joined = new Uint8Array(data.length + 16);
    joined.set(data);
    joined.set(extra, data.length);
    return new Uint8Array(await crypto.decrypt({ name: 'AES-CBC', iv }, key, joined));
  };

  const decryptEntry = async (stored, enc, password, cache) => {
    const crypto = subtle();
    const mode = checkSupported(enc);
    const key = await deriveKey(enc, password, cache, mode);
    let compressed;
    if (mode === 'cbc') {
      if (!stored.length || stored.length % 16 || !enc.iv) throw new Wurd.UserError('This document is damaged.');
      const padded = await decryptCbcUnpadded(key, enc.iv, stored);
      const pad = padded[padded.length - 1];
      if (pad < 1 || pad > 16) throw new WrongPassword();
      compressed = padded.subarray(0, padded.length - pad);
      if (enc.checksum) {
        const hash = /sha256/i.test(enc.checksumType) ? 'SHA-256' : 'SHA-1';
        const digest = new Uint8Array(await crypto.digest(hash, compressed.subarray(0, 1024)));
        if (!sameBytes(digest, enc.checksum)) throw new WrongPassword();
      }
    } else {
      // GCM entries are IV (12 bytes) + ciphertext + tag; a failed tag check means a wrong password
      try {
        compressed = new Uint8Array(await crypto.decrypt({ name: 'AES-GCM', iv: stored.subarray(0, 12), tagLength: 128 }, key, stored.subarray(12)));
      } catch (e) {
        throw new WrongPassword();
      }
    }
    try {
      return await inflateRaw(compressed);
    } catch (e) {
      // No checksum to tell a wrong password from damage; a wrong key almost never inflates
      throw new WrongPassword();
    }
  };

  // Keeps asking until the password decrypts the test entry (or the person cancels)
  const unlock = async (askPassword, attempt) => {
    let error;
    for (;;) {
      const password = await askPassword({ error });
      if (password === null || password === undefined) throw new Wurd.Cancelled();
      try {
        return { password, result: await attempt(password) };
      } catch (err) {
        if (!(err instanceof WrongPassword)) throw err;
        error = 'Wrong password, or the file is damaged. Try again.';
      }
    }
  };

  // Opens a package. Returns { lock, readBytes(path), readText(path) }; lock is null unless it was encrypted.
  const readPackage = async (bytes, { askPassword, mimePattern = /^application\/vnd\.oasis\.opendocument\.text(-template)?$/ } = {}) => {
    await Wurd.loadScript('lib/jszip.min.js');
    let zip;
    try {
      zip = await window.JSZip.loadAsync(bytes);
    } catch (e) {
      throw new Wurd.UserError("This file isn't a valid OpenDocument file.");
    }
    const mimetype = zip.file('mimetype') ? (await zip.file('mimetype').async('string')).trim() : '';
    if (!mimePattern.test(mimetype)) throw new Wurd.UserError("This file isn't an OpenDocument text document.");
    const manifestFile = zip.file('META-INF/manifest.xml');
    const manifest = manifestFile ? parseManifest(await manifestFile.async('string')) : new Map();
    const decoder = new TextDecoder('utf-8');
    // The lock is built for writing (always a SHA-256 start key), whichever start key this file was read with
    const api = async (readBytes, password) => ({
      lock: password === null ? null : await createLock(password),
      readBytes,
      readText: async (path) => {
        const data = await readBytes(path);
        return data ? decoder.decode(data) : null;
      },
    });

    // LibreOffice 24.8+ default: the whole document is one encrypted zip inside "encrypted-package"
    const wholesome = manifest.get('encrypted-package');
    if (wholesome?.encryption && zip.file('encrypted-package')) {
      checkSupported(wholesome.encryption);
      const stored = await zip.file('encrypted-package').async('uint8array');
      const { password, result: inner } = await unlock(askPassword, async (pw) => {
        const plain = await decryptEntry(stored, wholesome.encryption, pw, new Map());
        try {
          return await window.JSZip.loadAsync(plain);
        } catch (e) {
          throw new WrongPassword();
        }
      });
      const innerType = inner.file('mimetype') ? (await inner.file('mimetype').async('string')).trim() : '';
      if (innerType !== mimetype) throw new Wurd.UserError('This document is damaged.');
      return api(async (path) => (inner.file(path) ? inner.file(path).async('uint8array') : null), password);
    }

    const contentEntry = manifest.get('content.xml');
    if (!contentEntry?.encryption) {
      return api(async (path) => (zip.file(path) ? zip.file(path).async('uint8array') : null), null);
    }

    // Per-file encryption: content.xml is decrypted first and doubles as the password check
    checkSupported(contentEntry.encryption);
    const cache = new Map();
    const storedContent = await zip.file('content.xml').async('uint8array');
    const { password, result: content } = await unlock(askPassword, async (pw) => {
      cache.clear();
      return decryptEntry(storedContent, contentEntry.encryption, pw, cache);
    });
    return api(async (path) => {
      if (path === 'content.xml') return content;
      const file = zip.file(path);
      if (!file) return null;
      const raw = await file.async('uint8array');
      const info = manifest.get(path);
      if (!info?.encryption) return raw;
      try {
        return await decryptEntry(raw, info.encryption, password, cache);
      } catch (err) {
        if (err instanceof WrongPassword) throw new Wurd.UserError('This document is damaged.');
        throw err;
      }
    }, password);
  };

  F.odfPackage = { write: writePackage, read: readPackage, lock: createLock, parseManifest, ODF_VERSION, ODT_MIME };
})(window.Wurd = window.Wurd || {});
