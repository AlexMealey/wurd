// Encrypted Wurd documents (.wurd): the document is encrypted on this device with AES-256-GCM under a random
// document key. The file keeps that key wrapped (AES-256-GCM) under a key derived from the password with
// Argon2id, so every guess at the password costs an attacker 64 MiB of memory. Only Wurd opens these files.
// Version 1 files (PBKDF2-HMAC-SHA-256 straight to the document key) still open, and are saved as version 2.
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});

  // RFC 9106's second recommended setting: 64 MiB, 3 passes, 4 lanes
  const ARGON2 = { memory: 65536, iterations: 3, parallelism: 4 };
  // The most a file may ask for, so a hostile one can't make the tab hang or run out of memory
  const MAX_ARGON2 = { memory: 256 * 1024, iterations: 10, parallelism: 16 };
  const MAX_PBKDF2_ITERATIONS = 10000000;
  const KNOWN_FONTS = [Wurd.DEFAULT_FONT, Wurd.SERIF_FONT, Wurd.MONO_FONT];

  const encode = (text) => new TextEncoder().encode(text);
  const decode = (bytes) => new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  const randomBase64 = (length) => Wurd.bytesToBase64(crypto.getRandomValues(new Uint8Array(length)));

  // Authenticated with the wrapped key and with the document, so none of the header values can be swapped.
  // The document's own data covers the wrapped key, tying the two together.
  const keyAdditionalData = ({ kdf, key }) => encode(`wurd-encrypted-document|v2|key|Argon2id|${kdf.memory}|${kdf.iterations}|${kdf.parallelism}|${kdf.salt}|AES-GCM|${key.iv}`);
  const contentAdditionalData = ({ kdf, key }, iv) => encode(`wurd-encrypted-document|v2|content|${kdf.salt}|${key.iv}|${key.wrapped}|AES-GCM|${iv}`);
  const v1AdditionalData = ({ iterations, salt, iv }) => encode(`wurd-encrypted-document|v1|PBKDF2|SHA-256|${iterations}|${salt}|AES-GCM|${iv}`);

  const requireCrypto = () => {
    if (!Wurd.cryptoAvailable()) {
      throw new Wurd.UserError("Encrypted Wurd documents need the browser's Web Crypto, which isn't available here.");
    }
  };

  // The key that wraps the document key: Argon2id of the password. Its raw bytes, and the password's, are
  // wiped as soon as the browser holds the key.
  const passwordKey = async (password, kdf) => {
    await Wurd.loadScript('lib/argon2.umd.min.js');
    const passwordBytes = encode(password);
    let raw = null;
    try {
      raw = await window.hashwasm.argon2id({
        password: passwordBytes,
        salt: Wurd.base64ToBytes(kdf.salt),
        iterations: kdf.iterations,
        parallelism: kdf.parallelism,
        memorySize: kdf.memory,
        hashLength: 32,
        outputType: 'binary',
      });
      return await crypto.subtle.importKey('raw', raw, 'AES-GCM', false, ['wrapKey', 'unwrapKey']);
    } finally {
      passwordBytes.fill(0);
      raw?.fill(0);
    }
  };

  // The document key, usable but never readable: its bytes don't pass through JavaScript
  const unwrapDocumentKey = (wrappingKey, header) => crypto.subtle.unwrapKey(
    'raw', Wurd.base64ToBytes(header.key.wrapped), wrappingKey,
    { name: 'AES-GCM', iv: Wurd.base64ToBytes(header.key.iv), additionalData: keyAdditionalData(header) },
    'AES-GCM', false, ['encrypt', 'decrypt'],
  );

  // A new document key, wrapped under the password. The lock is what Save keeps instead of the password, so
  // saving again needs neither the password nor another Argon2id run.
  const createLock = async (password) => {
    requireCrypto();
    const kdf = { name: 'Argon2id', ...ARGON2, salt: randomBase64(16) };
    const key = { name: 'AES-GCM', iv: randomBase64(12) };
    const wrappingKey = await passwordKey(password, kdf);
    const fresh = await crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, true, ['encrypt', 'decrypt']);
    const wrapped = await crypto.subtle.wrapKey('raw', fresh, wrappingKey, { name: 'AES-GCM', iv: Wurd.base64ToBytes(key.iv), additionalData: keyAdditionalData({ kdf, key }) });
    key.wrapped = Wurd.bytesToBase64(new Uint8Array(wrapped));
    // Swap the new key, which could be read out, for one that can only be used
    return { format: 'wurd', kdf, key, dek: await unwrapDocumentKey(wrappingKey, { kdf, key }) };
  };

  const exportWurd = async (root, { title, pageFont, lock }) => {
    requireCrypto();
    if (lock?.format !== 'wurd') throw new Wurd.UserError('An encrypted Wurd document needs a password.');
    // A new IV for every save; the document key stays the same
    const iv = randomBase64(12);
    const plaintext = JSON.stringify({ app: 'Wurd', v: 1, title, font: pageFont, html: root.innerHTML, savedAt: new Date().toISOString() });
    const ciphertext = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: Wurd.base64ToBytes(iv), additionalData: contentAdditionalData(lock, iv) }, lock.dek, encode(plaintext)));
    const file = {
      wurd: 'encrypted-document',
      v: 2,
      kdf: lock.kdf,
      key: lock.key,
      cipher: { name: 'AES-GCM', iv },
      ciphertext: Wurd.bytesToBase64(ciphertext),
    };
    return new Blob([`${JSON.stringify(file, null, 2)}\n`], { type: 'application/json' });
  };

  const notWurd = () => new Wurd.UserError("This file isn't an encrypted Wurd document, or it's damaged.");

  const isInt = (n, min, max) => Number.isInteger(n) && n >= min && n <= max;

  // Decodes base64 header fields, refusing any that aren't valid or are shorter than needed
  const bytesOf = (fields) => {
    try {
      return fields.map(([text, minLength, maxLength = Infinity]) => {
        const bytes = Wurd.base64ToBytes(text);
        if (bytes.length < minLength || bytes.length > maxLength) throw notWurd();
        return bytes;
      });
    } catch (e) {
      throw notWurd();
    }
  };

  // Keeps asking until a password opens the file, or the person cancels. open() returns null for a wrong password.
  const unlock = async (askPassword, open) => {
    let error;
    for (;;) {
      const password = await askPassword({ error });
      if (password === null || password === undefined) throw new Wurd.Cancelled();
      const opened = await open(password);
      if (opened) return opened;
      // GCM's authentication fails for a wrong password and for any change to the key's part of the file
      error = 'Wrong password, or the file is damaged. Try again.';
    }
  };

  const openV2 = (file, askPassword) => {
    const { kdf, key, cipher } = file;
    const valid = kdf?.name === 'Argon2id' && isInt(kdf.parallelism, 1, MAX_ARGON2.parallelism)
      && isInt(kdf.memory, 8 * kdf.parallelism, MAX_ARGON2.memory) && isInt(kdf.iterations, 1, MAX_ARGON2.iterations)
      && typeof kdf.salt === 'string' && key?.name === 'AES-GCM' && typeof key.iv === 'string' && typeof key.wrapped === 'string'
      && cipher?.name === 'AES-GCM' && typeof cipher.iv === 'string' && typeof file.ciphertext === 'string';
    if (!valid) throw notWurd();
    const [, , , iv, ciphertext] = bytesOf([[kdf.salt, 16], [key.iv, 12, 12], [key.wrapped, 48, 48], [cipher.iv, 12, 12], [file.ciphertext, 16]]);
    // Only the known fields, so the lock never carries anything else from the file into the next save
    const header = {
      kdf: { name: 'Argon2id', memory: kdf.memory, iterations: kdf.iterations, parallelism: kdf.parallelism, salt: kdf.salt },
      key: { name: 'AES-GCM', iv: key.iv, wrapped: key.wrapped },
    };
    requireCrypto();
    return unlock(askPassword, async (password) => {
      const wrappingKey = await passwordKey(password, header.kdf);
      let dek;
      try {
        dek = await unwrapDocumentKey(wrappingKey, header);
      } catch (e) {
        return null;
      }
      // The password was right, so a document that fails its check was changed or damaged: asking again won't help
      try {
        const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv, additionalData: contentAdditionalData(header, cipher.iv) }, dek, ciphertext);
        return { plain, lock: { format: 'wurd', ...header, dek } };
      } catch (e) {
        throw new Wurd.UserError("This encrypted Wurd document is damaged and can't be opened.");
      }
    });
  };

  const openV1 = (file, askPassword) => {
    const { kdf, cipher } = file;
    const valid = kdf?.name === 'PBKDF2' && kdf.hash === 'SHA-256' && isInt(kdf.iterations, 1, MAX_PBKDF2_ITERATIONS)
      && typeof kdf.salt === 'string' && cipher?.name === 'AES-GCM' && typeof cipher.iv === 'string'
      && typeof file.ciphertext === 'string';
    if (!valid) throw notWurd();
    const [salt, iv, ciphertext] = bytesOf([[kdf.salt, 16], [cipher.iv, 12, 12], [file.ciphertext, 16]]);
    const header = { iterations: kdf.iterations, salt: kdf.salt, iv: cipher.iv };
    requireCrypto();
    return unlock(askPassword, async (password) => {
      const passwordBytes = encode(password);
      let plain;
      try {
        const base = await crypto.subtle.importKey('raw', passwordBytes, 'PBKDF2', false, ['deriveKey']);
        const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations: kdf.iterations }, base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
        plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv, additionalData: v1AdditionalData(header) }, key, ciphertext);
      } catch (e) {
        return null;
      } finally {
        passwordBytes.fill(0);
      }
      // A version 2 lock for the same password, so the next save upgrades the file
      return { plain, lock: await createLock(password) };
    });
  };

  const importWurd = async (bytes, { askPassword } = {}) => {
    let file;
    try {
      file = JSON.parse(decode(bytes));
    } catch (e) {
      throw notWurd();
    }
    if (file?.wurd !== 'encrypted-document' || !isInt(file.v, 1, Infinity)) throw notWurd();
    if (file.v > 2) throw new Wurd.UserError('This encrypted Wurd document was saved by a newer version of Wurd. Update Wurd to open it.');
    const { plain, lock } = await (file.v === 1 ? openV1 : openV2)(file, askPassword);
    let doc;
    try {
      doc = JSON.parse(decode(new Uint8Array(plain)));
    } catch (e) {
      throw notWurd();
    }
    if (typeof doc?.html !== 'string') throw notWurd();
    return { html: doc.html, font: KNOWN_FONTS.includes(doc.font) ? doc.font : Wurd.DEFAULT_FONT, lock };
  };

  F.wurd = { import: importWurd, export: exportWurd, lock: createLock };
})(window.Wurd = window.Wurd || {});
