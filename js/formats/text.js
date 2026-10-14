// Plain text: .txt import (paragraphs, line breaks and indentation kept) and export
(function (Wurd) {
  'use strict';

  const F = (Wurd.formats = Wurd.formats || {});
  const NBSP = Wurd.NBSP;
  const TAB_WIDTH = 4;

  // Expands tabs to the next tab stop, then turns spaces HTML would collapse into non-breaking spaces
  // so indentation and runs of spaces survive. Each run keeps one ordinary space so long lines still wrap.
  const preserveSpaces = (line) => {
    let expanded = '';
    for (const ch of line) {
      expanded += ch === '\t' ? ' '.repeat(TAB_WIDTH - (expanded.length % TAB_WIDTH)) : ch;
    }
    return expanded
      .replace(/^ +/, (indent) => NBSP.repeat(indent.length))
      .replace(/ {2,}/g, (run) => NBSP.repeat(run.length - 1) + ' ');
  };

  // Plain text to document HTML: blank lines separate paragraphs, single line breaks become <br>.
  // Built from text nodes, so characters like < and & stay literal text.
  const textToHtml = (text) => {
    const container = document.createElement('div');
    let p = null;
    for (const line of text.replace(/\r\n?/g, '\n').split('\n')) {
      if (!line.trim()) {
        // A blank (or whitespace-only) line ends the current paragraph
        p = null;
        continue;
      }
      if (p) {
        p.append(document.createElement('br'));
      } else {
        p = document.createElement('p');
        container.append(p);
      }
      p.append(preserveSpaces(line));
    }
    return container.hasChildNodes() ? container.innerHTML : '<p><br></p>';
  };

  const BLOCK_TAGS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'SECTION', 'ARTICLE',
    'HEADER', 'FOOTER', 'FIGURE', 'FIGCAPTION', 'ADDRESS', 'LI', 'DD', 'DT', 'DL']);

  const listText = (list, depth) => {
    const lines = [];
    let number = Number(list.getAttribute('start')) || 1;
    for (const child of list.children) {
      // execCommand can nest a list directly inside another list rather than inside an item
      if (child.tagName === 'UL' || child.tagName === 'OL') {
        lines.push(listText(child, depth + 1));
        continue;
      }
      if (child.tagName !== 'LI') continue;
      const item = child.cloneNode(true);
      item.querySelectorAll(':scope > ul, :scope > ol').forEach((sub) => sub.remove());
      const marker = list.tagName === 'OL' ? `${number++}. ` : '- ';
      const indent = '  '.repeat(depth);
      const body = htmlToText(item).trim().replace(/\n{2,}/g, '\n').split('\n');
      lines.push(indent + marker + body[0], ...body.slice(1).map((line) => indent + ' '.repeat(marker.length) + line));
      child.querySelectorAll(':scope > ul, :scope > ol').forEach((sub) => lines.push(listText(sub, depth + 1)));
    }
    return lines.join('\n');
  };

  const tableText = (table) => Array.from(table.rows, (row) =>
    Array.from(row.cells, (cell) => Wurd.collapseSpaces(cell.textContent).trim().split(NBSP).join(' ')).join('\t')
  ).join('\n');

  // Document HTML to plain text: blocks separated by blank lines, <br> as a line break,
  // lists with "-" or "1." markers, table cells separated by tabs
  const htmlToText = (root) => {
    const blocks = [];
    let line = '';
    const flush = () => {
      const text = line
        .split('\n')
        .map((part) => part.replace(/^ +| +$/g, ''))
        .join('\n')
        .split(NBSP)
        .join(' ');
      if (text.trim()) blocks.push(text.replace(/^\n+|\n+$/g, ''));
      line = '';
    };
    const walk = (node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        line += Wurd.collapseSpaces(node.data);
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const tag = node.tagName;
      if (tag === 'BR') {
        line += '\n';
      } else if (tag === 'IMG') {
        if (node.alt) line += `[${node.alt}]`;
      } else if (tag === 'PRE') {
        flush();
        blocks.push(node.textContent.replace(/\n$/, ''));
      } else if (tag === 'UL' || tag === 'OL') {
        flush();
        blocks.push(listText(node, 0));
      } else if (tag === 'TABLE') {
        flush();
        blocks.push(tableText(node));
      } else if (tag === 'HR') {
        flush();
        blocks.push('-'.repeat(10));
      } else if (BLOCK_TAGS.has(tag)) {
        flush();
        node.childNodes.forEach(walk);
        flush();
      } else {
        node.childNodes.forEach(walk);
      }
    };
    root.childNodes.forEach(walk);
    flush();
    return blocks.length ? `${blocks.join('\n\n')}\n` : '';
  };

  F.text = {
    preserveSpaces,
    textToHtml,
    htmlToText,
    import: (bytes) => ({ html: textToHtml(Wurd.decodeText(bytes)), font: Wurd.MONO_FONT }),
    export: (root) => new Blob([htmlToText(root)], { type: 'text/plain;charset=utf-8' }),
  };
})(window.Wurd = window.Wurd || {});
