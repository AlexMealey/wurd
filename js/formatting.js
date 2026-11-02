// Inline formatting the browser's execCommand can't do reliably: font, size, colour and highlight
// written as <span style>, with a real "remove" for each. Also paragraph and list indentation.
(function (Wurd) {
  'use strict';

  const BLOCK = /^(P|DIV|H[1-6]|UL|OL|LI|BLOCKQUOTE|PRE|HR|TABLE|THEAD|TBODY|TFOOT|TR|TD|TH|CAPTION|SECTION|ARTICLE|ASIDE|HEADER|FOOTER|NAV|MAIN|FIGURE|FIGCAPTION|ADDRESS|DL|DT|DD)$/;
  // Containers whose text children are only the whitespace between their rows or items
  const BLOCK_ONLY = /^(UL|OL|TABLE|THEAD|TBODY|TFOOT|TR|DL)$/;
  const LOOSE_SPACE = /^[ \t\n\r\f]*$/;
  // The <font> attribute that carries each property
  const FONT_ATTR = { 'font-family': 'face', color: 'color', 'font-size': 'size' };
  const PROPS = ['font-family', 'font-size', 'color', 'background-color'];

  let editor = null;

  const isElement = (node) => node && node.nodeType === Node.ELEMENT_NODE;
  const isBlock = (node) => isElement(node) && BLOCK.test(node.tagName);

  // The nearest block around a node (the editor itself when the text sits directly in it)
  const blockOf = (node) => {
    let el = isElement(node) ? node : node.parentNode;
    while (el && el !== editor && !isBlock(el)) el = el.parentNode;
    return el || editor;
  };

  const hasProp = (el, prop) => {
    if (!isElement(el)) return false;
    if (el.style && el.style.getPropertyValue(prop)) return true;
    if (el.tagName === 'FONT' && FONT_ATTR[prop] && el.hasAttribute(FONT_ATTR[prop])) return true;
    return prop === 'background-color' && el.tagName === 'MARK';
  };

  const unwrap = (el) => {
    el.replaceWith(...el.childNodes);
  };

  // Drops the property from an element; spans, fonts and marks left carrying nothing are unwrapped
  const removeProp = (el, prop) => {
    if (el.style) el.style.removeProperty(prop);
    if (el.tagName === 'FONT' && FONT_ATTR[prop]) el.removeAttribute(FONT_ATTR[prop]);
    if (el.getAttribute('style') === '') el.removeAttribute('style');
    if ((el.tagName === 'MARK' && prop === 'background-color') || (/^(SPAN|FONT)$/.test(el.tagName) && el.attributes.length === 0)) unwrap(el);
  };

  // Inline elements emptied by a split (no text, image or line break left) are removed
  const pruneEmpty = (root) => {
    if (!isElement(root)) return;
    Array.from(root.querySelectorAll('*')).reverse().forEach((el) => {
      if (!isBlock(el) && !/^(BR|IMG|HR)$/.test(el.tagName) && !el.firstChild) el.remove();
    });
  };

  // Splits `ancestor` at a point so the point becomes its edge. Returns [beforePart, afterPart];
  // a part that would be empty is null.
  const splitElement = (ancestor, container, offset) => {
    const range = document.createRange();
    range.setStart(container, offset);
    range.setEnd(ancestor, ancestor.childNodes.length);
    const tail = range.extractContents();
    const clone = ancestor.cloneNode(false);
    clone.append(tail);
    pruneEmpty(clone);
    pruneEmpty(ancestor);
    const hasContent = (el) => el.firstChild !== null;
    const after = hasContent(clone) ? clone : null;
    if (after) ancestor.after(after);
    const before = hasContent(ancestor) ? ancestor : null;
    if (!before) ancestor.remove();
    return [before, after];
  };

  // Splits every inline ancestor of `node` (below its block) that carries one of `props`, at the
  // node's start (side 'before') or end (side 'after'). The node itself is moved, never copied.
  const splitAncestors = (node, side, props, block) => {
    let el = node.parentNode;
    while (el && el !== block && el !== editor && !isBlock(el)) {
      if (props.some((p) => hasProp(el, p))) {
        // Split at the node's own edge; elements in between are split along with it
        const index = Array.prototype.indexOf.call(node.parentNode.childNodes, node);
        const [before, after] = splitElement(el, node.parentNode, side === 'before' ? index : index + 1);
        // Carry on from the part that still holds the node
        el = side === 'before' ? after || before : before || after;
      }
      el = el.parentNode;
    }
  };

  // Text nodes the range covers, after splitting the partly selected ones at its edges
  const selectedTextNodes = (range) => {
    const { endContainer, endOffset } = range;
    if (endContainer.nodeType === Node.TEXT_NODE && endOffset > 0 && endOffset < endContainer.length) endContainer.splitText(endOffset);
    // The live range moves its start onto the new node when it falls past the split
    const { startContainer, startOffset } = range;
    if (startContainer.nodeType === Node.TEXT_NODE && startOffset > 0 && startOffset < startContainer.length) startContainer.splitText(startOffset);
    const root = range.commonAncestorContainer;
    const nodes = [];
    const consider = (node) => {
      if (!node.length) return;
      if (range.comparePoint(node, 0) < 0 || range.comparePoint(node, node.length) > 0) return;
      if (!editor.contains(node)) return;
      if (LOOSE_SPACE.test(node.data)) {
        const parent = node.parentNode;
        if (parent === editor || BLOCK_ONLY.test(parent.nodeName) || isBlock(node.previousSibling) || isBlock(node.nextSibling)) return;
      }
      nodes.push(node);
    };
    if (root.nodeType === Node.TEXT_NODE) {
      consider(root);
    } else {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) consider(node);
    }
    return nodes;
  };

  const sameSpan = (a, b) => isElement(a) && isElement(b) && a.tagName === 'SPAN' && b.tagName === 'SPAN'
    && a.attributes.length === 1 && b.attributes.length === 1 && a.style.cssText === b.style.cssText;

  // Sets (or with value null, removes) one CSS property on the text a range covers.
  // Returns a range over the same text afterwards, or null when there was no text.
  const applyStyle = (range, prop, value) => {
    const nodes = selectedTextNodes(range);
    if (!nodes.length) return null;
    // Isolate each block's run of selected text from styled ancestors that reach past it
    const groups = new Map();
    nodes.forEach((node) => {
      const block = blockOf(node);
      if (!groups.has(block)) groups.set(block, []);
      groups.get(block).push(node);
    });
    groups.forEach((group, block) => {
      splitAncestors(group[group.length - 1], 'after', [prop], block);
      splitAncestors(group[0], 'before', [prop], block);
    });
    nodes.forEach((node) => {
      const block = blockOf(node);
      let el = node.parentNode;
      while (el && el !== block && el !== editor) {
        const parent = el.parentNode;
        if (hasProp(el, prop)) removeProp(el, prop);
        el = parent;
      }
    });
    if (value !== null && value !== undefined) {
      nodes.forEach((node) => {
        const parent = node.parentNode;
        if (parent.tagName === 'SPAN' && parent.childNodes.length === 1 && Array.from(parent.attributes).every((a) => a.name === 'style')) {
          parent.style.setProperty(prop, value);
          return;
        }
        const span = document.createElement('span');
        span.style.setProperty(prop, value);
        node.replaceWith(span);
        span.append(node);
      });
    }
    // Merge neighbouring spans that ended up with the same style
    nodes.forEach((node) => {
      const span = node.parentNode;
      if (span?.tagName !== 'SPAN') return;
      const prev = span.previousSibling;
      if (sameSpan(prev, span)) {
        prev.append(...span.childNodes);
        span.remove();
      }
    });
    const result = document.createRange();
    result.setStart(nodes[0], 0);
    const last = nodes[nodes.length - 1];
    result.setEnd(last, last.length);
    return result;
  };

  // ---- Style for the next typed text ----------------------------------------

  // Set at a collapsed caret: { prop: value|null }, applied to the next typed text
  let pending = null;

  const setPending = (prop, value) => {
    pending = pending || {};
    pending[prop] = value;
  };

  const clearPending = () => {
    pending = null;
  };

  const pendingValue = (prop) => (pending && prop in pending ? pending[prop] : undefined);

  // Runs before the history's own input listener, so the typed text is recorded already styled
  const onInput = (e) => {
    if (!pending || e.inputType !== 'insertText' || !e.data) return;
    const props = pending;
    pending = null;
    const selection = window.getSelection();
    if (!selection.rangeCount) return;
    const caret = selection.getRangeAt(0);
    const node = caret.endContainer;
    if (!caret.collapsed || node.nodeType !== Node.TEXT_NODE || caret.endOffset < e.data.length) return;
    let range = document.createRange();
    range.setStart(node, caret.endOffset - e.data.length);
    range.setEnd(node, caret.endOffset);
    for (const [prop, value] of Object.entries(props)) {
      range = applyStyle(range, prop, value) || range;
    }
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  };

  const NAVIGATION = /^(Arrow(Up|Down|Left|Right)|Home|End|PageUp|PageDown|Enter|Tab|Escape)$/;

  // ---- Reading the style at the caret ---------------------------------------

  const toPt = (px) => Math.round(parseFloat(px) * 0.75 * 2) / 2;

  // { font, size } at the selection's focus, with any pending choice taking precedence
  const currentStyle = () => {
    const selection = window.getSelection();
    if (!selection.rangeCount || !editor.contains(selection.focusNode)) return null;
    const el = isElement(selection.focusNode) ? selection.focusNode : selection.focusNode.parentElement;
    const computed = window.getComputedStyle(el);
    const pendingFont = pendingValue('font-family');
    const pendingSize = pendingValue('font-size');
    return {
      font: pendingFont || Wurd.classifyFont(Wurd.firstFontFamily(computed.fontFamily)),
      size: pendingSize ? parseFloat(pendingSize) : toPt(computed.fontSize),
    };
  };

  // ---- Selection that survives moving nodes ---------------------------------

  // Moving a node collapses live ranges inside it, so keep node/offset pairs and put them back
  const keepSelection = (fn) => {
    const selection = window.getSelection();
    const r = selection.rangeCount ? selection.getRangeAt(0) : null;
    const saved = r && [r.startContainer, r.startOffset, r.endContainer, r.endOffset];
    const result = fn();
    if (saved && editor.contains(saved[0]) && editor.contains(saved[2])) {
      try {
        const range = document.createRange();
        range.setStart(saved[0], Math.min(saved[1], maxOffset(saved[0])));
        range.setEnd(saved[2], Math.min(saved[3], maxOffset(saved[2])));
        selection.removeAllRanges();
        selection.addRange(range);
      } catch (e) {
        // The saved points no longer make a range; leave the selection where the browser put it
      }
    }
    return result;
  };

  const maxOffset = (node) => (node.nodeType === Node.TEXT_NODE ? node.length : node.childNodes.length);

  // ---- Indentation ----------------------------------------------------------

  const INDENTABLE = 'p, h1, h2, h3, h4, h5, h6, pre';

  // The paragraphs and list items a range touches, in document order. List items win over the
  // paragraphs inside them, and an item inside another selected item moves with it.
  const blocksIn = (range) => {
    const found = [];
    const add = (node) => {
      const el = isElement(node) ? node : node.parentElement;
      if (!el || !editor.contains(el)) return;
      const block = el.closest('li') || el.closest(INDENTABLE);
      if (block && block !== editor && editor.contains(block) && !found.includes(block)) found.push(block);
    };
    // A text node counts only if some of its text is selected (a triple-click ends at the next
    // paragraph's start, which shouldn't pull that paragraph in)
    const textSelected = (node) => {
      if (range.comparePoint(node, 0) > 0 || range.comparePoint(node, node.length) < 0) return false;
      if (node === range.endContainer && range.endOffset === 0) return false;
      return !(node === range.startContainer && range.startOffset === node.length);
    };
    if (!range.collapsed) {
      const root = range.commonAncestorContainer;
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      for (let node = walker.currentNode; node; node = walker.nextNode()) {
        if (node.nodeType === Node.TEXT_NODE ? node.data.trim() && textSelected(node) : /^(BR|IMG)$/.test(node.tagName) && range.intersectsNode(node)) add(node);
      }
    }
    // A caret, or a selection with no text in it (an empty paragraph)
    if (!found.length) add(range.startContainer);
    return found.filter((block) => !found.some((other) => other !== block && other.tagName === 'LI' && other.contains(block)));
  };

  const isList = (el) => isElement(el) && (el.tagName === 'UL' || el.tagName === 'OL');

  // Moves an item into a nested list at the end of the item before it
  const indentItem = (li) => {
    const prev = li.previousElementSibling;
    if (!prev || prev.tagName !== 'LI') return false;
    const tag = li.parentElement.tagName;
    let nested = prev.lastElementChild;
    if (!nested || nested.tagName !== tag) {
      nested = document.createElement(tag);
      prev.append(nested);
    }
    nested.append(li);
    return true;
  };

  // Moves an item out to its parent list, after the item it was nested in. The items after it
  // become its own nested list, so the order stays the same.
  const outdentItem = (li) => {
    const list = li.parentElement;
    const container = list.parentElement;
    let anchor;
    if (container?.tagName === 'LI') anchor = container;
    else if (isList(container)) anchor = list; // a list nested straight inside a list (old execCommand output)
    else return false;
    const following = [];
    for (let s = li.nextSibling; s; s = s.nextSibling) following.push(s);
    if (following.some((n) => isElement(n))) {
      let sub = li.lastElementChild;
      if (!isList(sub) || sub.tagName !== list.tagName) {
        sub = document.createElement(list.tagName);
        li.append(sub);
      }
      sub.append(...following);
    }
    anchor.after(li);
    if (!list.querySelector('li')) list.remove();
    return true;
  };

  const setIndentSteps = (el, steps) => {
    el.style.removeProperty('margin-inline-start');
    if (steps > 0) el.style.marginLeft = `${steps * Wurd.INDENT_PT}pt`;
    else el.style.removeProperty('margin-left');
    if (!el.getAttribute('style')) el.removeAttribute('style');
  };

  // Indents (direction 1) or outdents (-1) what the range touches. True if anything changed.
  const indent = (range, direction) => keepSelection(() => {
    let changed = false;
    for (const block of blocksIn(range)) {
      if (block.tagName === 'LI') {
        changed = (direction > 0 ? indentItem(block) : outdentItem(block)) || changed;
        continue;
      }
      const steps = Math.round(Wurd.indentOf(block) / Wurd.INDENT_PT);
      const next = Math.min(Math.max(steps + direction, 0), Wurd.MAX_INDENT_STEPS);
      if (next !== steps || (next === 0 && Wurd.indentOf(block))) {
        setIndentSteps(block, next);
        changed = true;
      }
    }
    return changed;
  });

  // ---- Setup ----------------------------------------------------------------

  // Call before the undo history is created, so typed text is styled before it's recorded
  const init = (editorEl) => {
    editor = editorEl;
    editor.addEventListener('input', onInput);
    editor.addEventListener('mousedown', clearPending);
    editor.addEventListener('keydown', (e) => {
      if (NAVIGATION.test(e.key)) clearPending();
    });
  };

  Wurd.formatting = {
    init,
    PROPS,
    applyStyle,
    setPending,
    clearPending,
    pendingValue,
    currentStyle,
    indent,
    keepSelection,
    blockOf: (node) => blockOf(node),
  };
})(window.Wurd = window.Wurd || {});
