// Encrypted Wurd documents (.wurd): the document is encrypted on this device with AES-256-GCM,
// using a key derived from the password with PBKDF2-HMAC-SHA-256. Only Wurd opens these files.
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});

  const ITERATIONS = 600000; // OWASP's 2023 recommendation for PBKDF2-HMAC-SHA-256
  const MAX_ITERATIONS = 10000000; // a hostile file can't make the tab hang for minutes
  const KNOWN_FONTS = [Wurd.DEFAULT_FONT, Wurd.SERIF_FONT, Wurd.MONO_FONT];

  const encode = (text) => new TextEncoder().encode(text);
  const decode = (bytes) => new TextDecoder('utf-8', { fatal: true }).decode(bytes);

  // Authenticated with the ciphertext, so none of the header values can be swapped
  const additionalData = ({ iterations, salt, iv }) => encode(`wurd-encrypted-document|v1|PBKDF2|SHA-256|${iterations}|${salt}|AES-GCM|${iv}`);

  const deriveKey = async (password, salt, iterations, usage) => {
    const base = await crypto.subtle.importKey('raw', encode(password), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations }, base, { name: 'AES-GCM', length: 256 }, false, [usage]);
  };

  const requireCrypto = () => {
    if (!Wurd.cryptoAvailable()) {
      throw new Wurd.UserError("Encrypted Wurd documents need the browser's Web Crypto, which isn't available here.");
    }
  };

  const exportWurd = async (root, { title, pageFont, password }) => {
    requireCrypto();
    if (!password) throw new Wurd.UserError('An encrypted Wurd document needs a password.');
    const salt = crypto.getRandomValues(new Uint8Array(16));
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const header = { iterations: ITERATIONS, salt: Wurd.bytesToBase64(salt), iv: Wurd.bytesToBase64(iv) };
    const plaintext = JSON.stringify({ app: 'Wurd', v: 1, title, font: pageFont, html: root.innerHTML, savedAt: new Date().toISOString() });
    const key = await deriveKey(password, salt, ITERATIONS, 'encrypt');
    const ciphertext = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv, additionalData: additionalData(header) }, key, encode(plaintext)));
    const file = {
      wurd: 'encrypted-document',
      v: 1,
      kdf: { name: 'PBKDF2', hash: 'SHA-256', iterations: ITERATIONS, salt: header.salt },
      cipher: { name: 'AES-GCM', iv: header.iv },
      ciphertext: Wurd.bytesToBase64(ciphertext),
    };
    return new Blob([`${JSON.stringify(file, null, 2)}\n`], { type: 'application/json' });
  };

  const notWurd = () => new Wurd.UserError("This file isn't an encrypted Wurd document, or it's damaged.");

  const importWurd = async (bytes, { askPassword } = {}) => {
    let file;
    try {
      file = JSON.parse(decode(bytes));
    } catch (e) {
      throw notWurd();
    }
    const { kdf, cipher } = file || {};
    const valid = file?.wurd === 'encrypted-document' && file.v === 1
      && kdf?.name === 'PBKDF2' && kdf.hash === 'SHA-256'
      && Number.isInteger(kdf.iterations) && kdf.iterations > 0 && kdf.iterations <= MAX_ITERATIONS
      && typeof kdf.salt === 'string' && cipher?.name === 'AES-GCM' && typeof cipher.iv === 'string'
      && typeof file.ciphertext === 'string';
    if (!valid) throw notWurd();
    requireCrypto();
    let salt;
    let iv;
    let ciphertext;
    try {
      salt = Wurd.base64ToBytes(kdf.salt);
      iv = Wurd.base64ToBytes(cipher.iv);
      ciphertext = Wurd.base64ToBytes(file.ciphertext);
    } catch (e) {
      throw notWurd();
    }
    if (salt.length < 16 || iv.length !== 12 || ciphertext.length < 16) throw notWurd();
    const header = { iterations: kdf.iterations, salt: kdf.salt, iv: cipher.iv };

    let error;
    for (;;) {
      const password = await askPassword({ error });
      if (password === null || password === undefined) throw new Wurd.Cancelled();
      let plain;
      try {
        const key = await deriveKey(password, salt, kdf.iterations, 'decrypt');
        plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv, additionalData: additionalData(header) }, key, ciphertext);
      } catch (e) {
        // GCM's authentication fails for a wrong password and for any change to the file
        error = 'Wrong password, or the file is damaged. Try again.';
        continue;
      }
      let doc;
      try {
        doc = JSON.parse(decode(new Uint8Array(plain)));
      } catch (e) {
        throw notWurd();
      }
      if (typeof doc?.html !== 'string') throw notWurd();
      return { html: doc.html, font: KNOWN_FONTS.includes(doc.font) ? doc.font : Wurd.DEFAULT_FONT, password };
    }
  };

  F.wurd = { import: importWurd, export: exportWurd };
})(window.Wurd = window.Wurd || {});
