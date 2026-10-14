// HTML: import any web page (sanitized when loaded) and export a standalone page
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});
  const KNOWN_FONTS = [Wurd.DEFAULT_FONT, Wurd.SERIF_FONT, Wurd.MONO_FONT];

  // Pages Wurd saved remember their page font in a <meta> tag
  const pageFontOf = (doc) => {
    const font = doc.querySelector('meta[name="wurd-page-font"]')?.getAttribute('content');
    return KNOWN_FONTS.includes(font) ? font : Wurd.DEFAULT_FONT;
  };

  const importHtml = (bytes) => {
    const text = Wurd.decodeText(bytes);
    // DOMParser builds an inert document: nothing in it runs or loads
    const doc = new DOMParser().parseFromString(text, 'text/html');
    return { html: text, font: pageFontOf(doc) };
  };

  const PAGE_CSS = `
    body { max-width: 48rem; margin: 2rem auto; padding: 0 1rem; color: #18181b; line-height: 1.625; }
    h1, h2, h3, h4 { line-height: 1.25; }
    blockquote { margin-left: 0; padding-left: 1em; border-left: 4px solid #e5e7eb; font-style: italic; }
    pre { padding: 0.85em 1.15em; border-radius: 6px; background: #1f2937; color: #e5e7eb; overflow-x: auto; }
    table { border-collapse: collapse; }
    th, td { padding: 0.25em 0.5em; border: 1px solid #d1d5db; text-align: start; }
    img { max-width: 100%; height: auto; }`;

  const exportHtml = (root, { title, pageFont }) => {
    const esc = Wurd.escapeXml;
    const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Wurd">
<meta name="wurd-page-font" content="${esc(pageFont)}">
<title>${esc(title)}</title>
<style>
    body { font-family: ${pageFont.replace('Inter,', 'Inter, system-ui,')}; }${PAGE_CSS}
</style>
</head>
<body>
${root.innerHTML}
</body>
</html>
`;
    return new Blob([html], { type: 'text/html;charset=utf-8' });
  };

  F.html = { import: importHtml, export: exportHtml, pageFontOf };
})(window.Wurd = window.Wurd || {});
