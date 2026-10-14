// The optional Wurd server (wurd-server): keeps each user's documents, encrypted with a key from their
// token, and hands them back. Wurd works without it. Files go up as exactly the bytes Save would download,
// so the server never sees a document in any other form, and .wurd files stay encrypted.
(function (Wurd) {
  'use strict';

  // The address, if the user chose to remember it on this device. The token is never kept here.
  const ADDRESS_KEY = 'wurd-server';
  // The address and token, if the user chose to stay connected in this tab
  const SESSION_KEY = 'wurd-server-session';

  // A refusal from the server; status is the HTTP status, rev the file's current revision on a conflict
  class ServerError extends Wurd.UserError {
    constructor(message, status, rev = null) {
      super(message);
      this.status = status;
      this.rev = rev;
    }
  }

  // storage is 'localStorage' or 'sessionStorage'; even looking either up can throw when storage is blocked
  const read = (storage, key) => {
    try {
      return JSON.parse(window[storage].getItem(key));
    } catch (e) {
      return null;
    }
  };

  const write = (storage, key, value) => {
    try {
      if (value) window[storage].setItem(key, JSON.stringify(value));
      else window[storage].removeItem(key);
    } catch (e) {
      // Storage unavailable: nothing is remembered
    }
  };

  const loadSavedUrl = () => {
    const saved = read('localStorage', ADDRESS_KEY);
    if (typeof saved?.url !== 'string') return null;
    // Older versions kept the token here too; drop it
    if ('token' in saved) write('localStorage', ADDRESS_KEY, { url: saved.url });
    return saved.url;
  };

  const loadConnection = () => {
    const saved = read('sessionStorage', SESSION_KEY);
    return typeof saved?.url === 'string' && typeof saved?.token === 'string' ? saved : null;
  };

  let savedUrl = loadSavedUrl();
  let connection = loadConnection();

  const revOf = (response) => {
    const match = /^"(\d+)"$/.exec(response.headers.get('ETag') || '');
    return match ? Number(match[1]) : null;
  };

  const request = async (method, path, { body, headers = {} } = {}, conn = connection) => {
    if (!conn) throw new Wurd.UserError('Connect to a Wurd server first, in the Vault tab.');
    let response;
    try {
      response = await fetch(`${conn.url}/api${path}`, {
        method,
        body,
        headers: { ...(conn.token && { Authorization: `Bearer ${conn.token}` }), ...headers },
        cache: 'no-store',
        credentials: 'omit',
        referrerPolicy: 'no-referrer',
      });
    } catch (e) {
      throw new ServerError(`Couldn't reach the Wurd server at ${conn.url}.`, 0);
    }
    if (!response.ok) {
      let message = '';
      try {
        message = (await response.json()).error;
      } catch (e) {
        // Not one of the server's JSON errors (a proxy, or something else at that address)
      }
      throw new ServerError(message || `The Wurd server answered with error ${response.status}.`, response.status, revOf(response));
    }
    return response;
  };

  // "http://host:8026/" -> "http://host:8026"; anything that isn't an http(s) address is refused
  const normaliseUrl = (text) => {
    let url;
    try {
      url = new URL(text.trim());
    } catch (e) {
      throw new Wurd.UserError('Enter the server address, e.g. http://localhost:8026.');
    }
    if (!/^https?:$/.test(url.protocol)) throw new Wurd.UserError('The server address must start with http:// or https://.');
    return url.origin + url.pathname.replace(/\/+$/, '');
  };

  // Checks the address and token by listing files. If they work, remembers the address on this device
  // and the token in this tab, each only if asked to.
  const connect = async (urlText, token, { rememberUrl = false, keepInTab = false } = {}) => {
    const conn = { url: normaliseUrl(urlText), token: token.trim() };
    if (!conn.token) throw new Wurd.UserError('Enter your token, or request one.');
    await request('GET', '/files', {}, conn);
    connection = conn;
    savedUrl = rememberUrl ? conn.url : null;
    write('localStorage', ADDRESS_KEY, savedUrl && { url: savedUrl });
    write('sessionStorage', SESSION_KEY, keepInTab && conn);
  };

  // Forgets the token; a remembered address stays
  const disconnect = () => {
    connection = null;
    write('sessionStorage', SESSION_KEY, null);
  };

  // A new token for a new, empty vault on the server at that address. The server shows it only once.
  const requestToken = async (urlText) => {
    const response = await request('POST', '/tokens', {}, { url: normaliseUrl(urlText) });
    return (await response.json()).token;
  };

  const nameHeader = (name) => ({ 'X-File-Name': encodeURIComponent(name) });

  Wurd.server = {
    ServerError,
    connect,
    disconnect,
    requestToken,
    connected: () => Boolean(connection),
    url: () => connection?.url ?? null,
    savedUrl: () => savedUrl,
    keptInTab: () => Boolean(read('sessionStorage', SESSION_KEY)),

    // [{ id, name, format, size, rev, updatedAt }], newest first
    list: async () => (await request('GET', '/files')).json(),

    // The stored file as a File (to open like any other), and the revision it was read at
    download: async ({ id, name }) => {
      const response = await request('GET', `/files/${id}`);
      return { file: new File([await response.blob()], name), rev: revOf(response) };
    },

    create: async (name, blob) => (await request('POST', '/files', { body: blob, headers: nameHeader(name) })).json(),

    // Replaces the file if it's still at rev; a ServerError with status 412 if it changed elsewhere
    update: async (id, rev, name, blob) => (await request('PUT', `/files/${id}`, { body: blob, headers: { 'If-Match': `"${rev}"`, ...nameHeader(name) } })).json(),

    remove: async (id, rev) => {
      await request('DELETE', `/files/${id}`, { headers: { 'If-Match': `"${rev}"` } });
    },
  };
})(window.Wurd = window.Wurd || {});
