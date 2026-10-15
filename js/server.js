// The optional Wurd server (wurd-server): keeps documents and hands them back. Wurd works without it.
// Files go up as exactly the bytes Save would download, so the server never sees a document in any
// other form, and .wurd files stay encrypted.
(function (Wurd) {
  'use strict';

  const STORAGE_KEY = 'wurd-server';

  // A refusal from the server; status is the HTTP status, rev the file's current revision on a conflict
  class ServerError extends Wurd.UserError {
    constructor(message, status, rev = null) {
      super(message);
      this.status = status;
      this.rev = rev;
    }
  }

  const loadConnection = () => {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
      return typeof saved?.url === 'string' && typeof saved?.token === 'string' ? saved : null;
    } catch (e) {
      return null;
    }
  };

  let connection = loadConnection();

  const revOf = (response) => {
    const match = /^"(\d+)"$/.exec(response.headers.get('ETag') || '');
    return match ? Number(match[1]) : null;
  };

  // One API call to the server at url; a ServerError if it can't be reached or refuses
  const send = async (url, path, { method, body, headers = {} }) => {
    let response;
    try {
      response = await fetch(`${url}/api${path}`, {
        method,
        body,
        headers,
        cache: 'no-store',
        credentials: 'omit',
        referrerPolicy: 'no-referrer',
      });
    } catch (e) {
      throw new ServerError(`Couldn't reach the Wurd server at ${url}.`, 0);
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

  const request = async (method, path, { body, headers = {} } = {}, conn = connection) => {
    if (!conn) throw new Wurd.UserError('Connect to a Wurd server first, in the Vault tab.');
    return send(conn.url, path, { method, body, headers: { Authorization: `Bearer ${conn.token}`, ...headers } });
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

  // Checks the address and token by listing files, and remembers them on this device if they work
  const connect = async (urlText, token) => {
    const conn = { url: normaliseUrl(urlText), token: token.trim() };
    if (!conn.token) throw new Wurd.UserError('Enter the token.');
    await request('GET', '/files', {}, conn);
    connection = conn;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(conn));
    } catch (e) {
      // Storage unavailable: connected for this visit only
    }
  };

  // Asks the server for a new token, which opens a new, empty vault, and connects with it. Returns the
  // token: the server keeps no copy of it, so this is the only chance to show it.
  const requestToken = async (urlText) => {
    const url = normaliseUrl(urlText);
    const response = await send(url, '/tokens', { method: 'POST' });
    const { token } = await response.json().catch(() => ({}));
    if (typeof token !== 'string') throw new ServerError(`${url} didn't answer like a Wurd server.`, response.status);
    await connect(url, token);
    return token;
  };

  const disconnect = () => {
    connection = null;
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch (e) {
      // Nothing was stored
    }
  };

  const nameHeader = (name) => ({ 'X-File-Name': encodeURIComponent(name) });

  Wurd.server = {
    ServerError,
    connect,
    requestToken,
    disconnect,
    connected: () => Boolean(connection),
    url: () => connection?.url ?? null,

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
