// Wurd as an installed app: the offline service worker (sw.js), the update banner, the Install
// button, opening files the operating system hands over, and a title bar that follows the theme.
// Opened from file://, only the title bar colour applies.
(function (Wurd) {
  'use strict';

  const $ = (selector) => document.querySelector(selector);

  const root = document.documentElement;
  const editor = $('#editor');
  const installButton = $('#install-app');
  const updateBanner = $('#update-banner');

  // ---- Title bar ----------------------------------------------------------

  // Both theme-color tags (one per OS colour scheme) take the app theme's header colour,
  // so the title bar matches even when the theme toggle disagrees with the OS
  const syncThemeColor = () => {
    const color = getComputedStyle(root).getPropertyValue('--header-bg').trim();
    document.querySelectorAll('meta[name="theme-color"]').forEach((meta) => {
      meta.content = color;
    });
  };

  syncThemeColor();
  new MutationObserver(syncThemeColor).observe(root, { attributes: true, attributeFilter: ['data-theme'] });

  // ---- Files from the operating system ------------------------------------

  // The installed app is listed under "Open with" for the formats in manifest.webmanifest.
  // Each launch gets a new window, so it never replaces a document that's already open.
  if ('launchQueue' in window) {
    window.launchQueue.setConsumer(async ({ files }) => {
      if (!files?.length) return;
      try {
        await Wurd.app.openFile(await files[0].getFile());
      } catch (err) {
        console.error(err);
        Wurd.app.notify("Couldn't open the file.");
      }
    });
  }

  // ---- Install ------------------------------------------------------------

  let installPrompt = null;

  const hideInstall = () => {
    installPrompt = null;
    if (installButton.contains(document.activeElement)) editor.focus();
    installButton.hidden = true;
  };

  // The browser offers installation; show our button instead of its mobile banner
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    if (window.matchMedia('(display-mode: standalone)').matches) return;
    installPrompt = e;
    installButton.hidden = false;
  });

  installButton.addEventListener('click', async () => {
    if (!installPrompt) return;
    const prompt = installPrompt;
    prompt.prompt();
    const { outcome } = await prompt.userChoice;
    // A prompt can only be shown once; the browser fires beforeinstallprompt again if it can retry
    if (outcome === 'accepted') hideInstall();
    else installPrompt = null;
  });

  window.addEventListener('appinstalled', () => {
    hideInstall();
    Wurd.app.notify('Wurd installed. Open it from your apps.');
  });

  // ---- Offline and updates ------------------------------------------------

  // Service workers need http(s) and a secure context (https, or localhost)
  if (!('serviceWorker' in navigator) || !/^https?:$/.test(location.protocol)) return;

  let waitingWorker = null;
  let reloading = false;

  const showUpdate = (worker) => {
    waitingWorker = worker;
    if (!updateBanner.hidden) return;
    updateBanner.hidden = false;
    Wurd.app.announce('A new version of Wurd is ready. Reload to use it.');
  };

  const hideUpdate = () => {
    if (updateBanner.contains(document.activeElement)) editor.focus();
    updateBanner.hidden = true;
  };

  // The new version takes over, then the page reloads into it. The open document comes back as a
  // draft (js/drafts.js), or the browser asks first when it can't be kept.
  $('#update-reload').addEventListener('click', () => {
    reloading = true;
    if (waitingWorker?.state === 'installed') waitingWorker.postMessage('skip-waiting');
    // Another tab may already have switched to the new version
    else location.reload();
  });

  $('#update-later').addEventListener('click', hideUpdate);

  // Fires when the first version claims the page too, which needs no reload
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (reloading) location.reload();
  });

  const watch = (registration) => {
    // Waiting from an earlier visit (other Wurd tabs kept the old version in use)
    if (registration.waiting && navigator.serviceWorker.controller) showUpdate(registration.waiting);
    registration.addEventListener('updatefound', () => {
      const worker = registration.installing;
      worker?.addEventListener('statechange', () => {
        // Without a controller this is the first install: the app now works offline, nothing to update
        if (worker.state === 'installed' && navigator.serviceWorker.controller) showUpdate(worker);
      });
    });
    // Check for a new version whenever Wurd comes back into view
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') registration.update().catch(() => {});
    });
  };

  navigator.serviceWorker
    // updateViaCache: 'none' checks sw.js itself past the HTTP cache, so new versions are noticed
    .register('sw.js', { scope: './', updateViaCache: 'none' })
    .then(watch)
    .catch((err) => console.warn("Wurd won't work offline: the service worker didn't register.", err));
})(window.Wurd = window.Wurd || {});
