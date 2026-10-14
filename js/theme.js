// Picks the theme before the page paints, so it never flashes the wrong one:
// the last choice made with the toggle, otherwise the operating system's preference.
(function () {
  var theme = null;
  try {
    theme = localStorage.getItem('wurd-theme');
  } catch (e) {
    // Storage can be unavailable (private windows, blocked site data); fall back to the OS preference
  }
  if (theme !== 'light' && theme !== 'dark') {
    theme = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  document.documentElement.dataset.theme = theme;
})();
