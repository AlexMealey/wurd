// Drafts: the open document is kept in this tab's sessionStorage, so reloading the page doesn't lose it.
// Each tab keeps its own, the browser deletes it when the tab closes, and it expires after a few minutes.
// Password-protected documents are never kept, since the draft would hold them unencrypted.
(function (Wurd) {
  'use strict';

  const DRAFT_KEY = 'wurd-draft';
  const MINUTES_KEY = 'wurd-draft-minutes';
  // Must match the #settings-draft options in index.html; 0 turns drafts off
  const LIFETIMES = [0, 5, 15, 30, 60];
  const DEFAULT_MINUTES = 15;
  const FONTS = [Wurd.DEFAULT_FONT, Wurd.SERIF_FONT, Wurd.MONO_FONT];

  const loadMinutes = () => {
    try {
      const saved = Number(localStorage.getItem(MINUTES_KEY) ?? DEFAULT_MINUTES);
      return LIFETIMES.includes(saved) ? saved : DEFAULT_MINUTES;
    } catch (e) {
      return DEFAULT_MINUTES;
    }
  };

  let minutes = loadMinutes();
  // What the last save did, for the Settings dialog: 'kept', 'off', 'protected', 'failed' or 'none'
  let state = { status: minutes ? 'none' : 'off' };
  let expiry = 0;

  const lifetimeMs = () => minutes * 60000;

  const clear = (status = minutes ? 'none' : 'off') => {
    clearTimeout(expiry);
    state = { status };
    try {
      sessionStorage.removeItem(DRAFT_KEY);
    } catch (e) {
      // Storage unavailable: nothing was kept
    }
  };

  // Deletes the draft once its time is up, even while the tab stays open
  const keptUntil = (title, savedAt) => {
    state = { status: 'kept', title, savedAt };
    clearTimeout(expiry);
    expiry = setTimeout(() => clear(), savedAt + lifetimeMs() - Date.now());
  };

  // doc: { html, title, format, font, server: { url, id, rev } | null, protected }.
  // True when a reload will bring the document back.
  const save = (doc) => {
    if (!minutes || doc.protected) {
      clear(minutes ? 'protected' : 'off');
      return false;
    }
    const savedAt = Date.now();
    try {
      const { html, title, format, font, server } = doc;
      sessionStorage.setItem(DRAFT_KEY, JSON.stringify({ v: 1, html, title, format, font, server, savedAt }));
    } catch (e) {
      // Over the storage quota (large images) or storage is blocked. An older draft would bring back
      // stale content, so it goes too.
      clear('failed');
      return false;
    }
    keptUntil(doc.title, savedAt);
    return true;
  };

  // The draft left behind by a reload, if it's still fresh. Its HTML is not sanitized yet.
  const restore = () => {
    let d = null;
    try {
      d = JSON.parse(sessionStorage.getItem(DRAFT_KEY));
    } catch (e) {
      // Unreadable: treated as no draft
    }
    const age = Date.now() - d?.savedAt;
    const valid = d?.v === 1 && typeof d.html === 'string' && typeof d.title === 'string'
      && Wurd.formats.EXPORTS.some((f) => f.id === d.format) && FONTS.includes(d.font)
      && age >= 0 && age < lifetimeMs();
    if (!valid) {
      clear();
      return null;
    }
    const s = d.server;
    const server = typeof s?.url === 'string' && /^[0-9a-f]{32}$/.test(s.id) && Number.isInteger(s.rev) ? { url: s.url, id: s.id, rev: s.rev } : null;
    keptUntil(d.title, d.savedAt);
    return { html: d.html, title: d.title, format: d.format, font: d.font, server };
  };

  const setMinutes = (value) => {
    minutes = LIFETIMES.includes(value) ? value : DEFAULT_MINUTES;
    try {
      localStorage.setItem(MINUTES_KEY, String(minutes));
    } catch (e) {
      // Storage unavailable: the setting lasts for this visit
    }
    if (!minutes) clear();
  };

  const describe = () => {
    switch (state.status) {
      case 'kept': {
        const until = new Date(state.savedAt + lifetimeMs()).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        return `Kept: “${state.title || 'Untitled'}”, until ${until}.`;
      }
      case 'off': return 'Nothing is kept.';
      case 'protected': return 'Not kept: this document is password-protected.';
      case 'failed': return "Not kept: the document is too large for the browser's storage, or storage is blocked.";
      default: return 'Nothing kept right now.';
    }
  };

  Wurd.drafts = { save, restore, clear, setMinutes, minutes: () => minutes, describe };
})(window.Wurd = window.Wurd || {});
