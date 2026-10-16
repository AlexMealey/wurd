// Wurd's service worker: keeps a copy of the whole app so the installed app opens and works offline.
// Only registered when Wurd is served over http(s) (js/pwa.js); opened from file:// it isn't used.
//
// Bump VERSION whenever any file in PRECACHE changes. Browsers only notice a new version when this
// file's bytes change, and each version's files are kept together in their own cache, so a page
// never mixes files from two versions.
'use strict';

const VERSION = '1';
const CACHE = `wurd-${VERSION}`;

// Everything the app loads, including the libraries it only loads on first use (Wurd.loadScript),
// so opening and saving every format works offline
const PRECACHE = [
  'index.html',
  'manifest.webmanifest',
  'css/app.css',
  'icons/icon.svg',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
  'icons/apple-touch-icon.png',
  'js/theme.js',
  'js/util.js',
  'js/formats/text.js',
  'js/formats/html.js',
  'js/formats/docx.js',
  'js/formats/doc.js',
  'js/formats/odf-package.js',
  'js/formats/odt.js',
  'js/formats/wurd.js',
  'js/formats/index.js',
  'js/server.js',
  'js/drafts.js',
  'js/history.js',
  'js/dialogs.js',
  'js/shortcuts.js',
  'js/app.js',
  'js/pwa.js',
  'lib/lucide.min.js',
  'lib/purify.min.js',
  'lib/mammoth.browser.min.js',
  'lib/docx.iife.js',
  'lib/cfb.min.js',
  'lib/jszip.min.js',
  'lib/argon2.umd.min.js',
  'lib/fonts/inter.css',
  ...['cyrillic-ext', 'cyrillic', 'greek-ext', 'greek', 'latin-ext', 'latin', 'vietnamese'].flatMap((subset) => [
    `lib/fonts/files/inter-${subset}-wght-normal.woff2`,
    `lib/fonts/files/inter-${subset}-wght-italic.woff2`,
  ]),
];

const scopeUrl = (path) => new URL(path, self.registration.scope).href;

self.addEventListener('install', (event) => {
  // cache: 'reload' skips the HTTP cache, so a new version never stores an older file.
  // The new version then waits until the page chooses to switch (js/pwa.js), or every Wurd tab closes.
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(PRECACHE.map((path) => new Request(path, { cache: 'reload' })))));
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter((name) => name.startsWith('wurd-') && name !== CACHE).map((name) => caches.delete(name)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = new URL(request.url);
  // Wurd server API calls, other sites and anything that isn't a GET go straight to the network
  if (request.method !== 'GET' || url.origin !== self.location.origin || url.href.startsWith(scopeUrl('api/'))) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    // Opening the app (with or without a query string) gets the cached page
    const path = url.origin + url.pathname;
    const isApp = request.mode === 'navigate' && (path === scopeUrl('./') || path === scopeUrl('index.html'));
    const cached = await cache.match(isApp ? 'index.html' : request, { ignoreSearch: isApp });
    return cached || fetch(request);
  })());
});

// The page asks the waiting version to take over when the person chooses to update
self.addEventListener('message', (event) => {
  if (event.data === 'skip-waiting') self.skipWaiting();
});
