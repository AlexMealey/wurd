// Undo/redo for the editor. The browser's own undo stack is unreliable once content is replaced
// with innerHTML, so Wurd keeps its own: whole-document snapshots plus where the caret was.
(function (Wurd) {
  'use strict';

  const TYPING = /^(insertText|insertReplacementText|insertCompositionText|deleteContentBackward|deleteContentForward|deleteWordBackward|deleteWordForward)$/;
  const COALESCE_MS = 1000;

  Wurd.createHistory = ({ editor, limit = 100, onChange = () => {} }) => {
    let undoStack = [];
    let redoStack = [];
    let suppressed = 0;
    let pendingSelection = null;
    let lastKind = null;
    let lastTime = 0;

    // Selections are stored as child-index paths from the editor, which stay valid after restoring the HTML
    const pathTo = (node) => {
      const path = [];
      while (node && node !== editor) {
        const parent = node.parentNode;
        if (!parent) return null;
        path.unshift(Array.prototype.indexOf.call(parent.childNodes, node));
        node = parent;
      }
      return node === editor ? path : null;
    };

    const nodeAt = (path) => {
      let node = editor;
      for (const index of path) {
        node = node.childNodes[index];
        if (!node) return null;
      }
      return node;
    };

    const readSelection = () => {
      const selection = window.getSelection();
      if (!selection || !selection.rangeCount) return null;
      const range = selection.getRangeAt(0);
      const start = pathTo(range.startContainer);
      const end = pathTo(range.endContainer);
      if (!start || !end) return null;
      return { start, startOffset: range.startOffset, end, endOffset: range.endOffset };
    };

    const maxOffset = (node) => (node.nodeType === Node.TEXT_NODE ? node.length : node.childNodes.length);

    const restoreSelection = (saved) => {
      const selection = window.getSelection();
      const startNode = saved && nodeAt(saved.start);
      const endNode = saved && nodeAt(saved.end);
      const range = document.createRange();
      if (startNode && endNode) {
        range.setStart(startNode, Math.min(saved.startOffset, maxOffset(startNode)));
        range.setEnd(endNode, Math.min(saved.endOffset, maxOffset(endNode)));
      } else {
        range.selectNodeContents(editor);
        range.collapse(true);
      }
      selection.removeAllRanges();
      selection.addRange(range);
      // Bring the restored caret into view
      const target = range.startContainer.nodeType === Node.ELEMENT_NODE ? range.startContainer : range.startContainer.parentElement;
      target?.scrollIntoView?.({ block: 'nearest' });
    };

    let current = { html: editor.innerHTML, selection: null };

    // Call after the document changed. Consecutive typing within a second becomes one step.
    const record = (kind = 'edit') => {
      if (suppressed) return;
      const html = editor.innerHTML;
      if (html === current.html) {
        pendingSelection = null;
        return;
      }
      const now = Date.now();
      const coalesce = kind === 'typing' && lastKind === 'typing' && now - lastTime < COALESCE_MS && undoStack.length > 0;
      if (!coalesce) {
        undoStack.push({ html: current.html, selection: pendingSelection || current.selection });
        if (undoStack.length > limit) undoStack.shift();
      }
      redoStack = [];
      current = { html, selection: readSelection() };
      lastKind = kind;
      lastTime = now;
      pendingSelection = null;
      onChange();
    };

    const apply = (state) => {
      editor.innerHTML = state.html;
      current = { html: editor.innerHTML, selection: state.selection };
      lastKind = null;
      pendingSelection = null;
      editor.focus({ preventScroll: true });
      restoreSelection(state.selection);
      onChange();
    };

    const undo = () => {
      // Catch any change that slipped past the input listener
      if (editor.innerHTML !== current.html) record('edit');
      const previous = undoStack.pop();
      if (!previous) return false;
      redoStack.push({ html: current.html, selection: readSelection() || current.selection });
      apply(previous);
      return true;
    };

    const redo = () => {
      const next = redoStack.pop();
      if (!next) return false;
      undoStack.push({ html: current.html, selection: readSelection() || current.selection });
      apply(next);
      return true;
    };

    const reset = () => {
      undoStack = [];
      redoStack = [];
      current = { html: editor.innerHTML, selection: null };
      lastKind = null;
      pendingSelection = null;
      onChange();
    };

    // Runs fn without recording the input events it causes (for execCommand), then records one step
    const batch = (kind, fn) => {
      // Undo should bring back the selection the command was applied to
      const before = readSelection();
      suppressed++;
      try {
        fn();
      } finally {
        suppressed--;
      }
      if (before) pendingSelection = before;
      record(kind);
    };

    editor.addEventListener('beforeinput', (e) => {
      // The Edit menu and context menu undo through the browser's stack; send them to ours instead
      if (e.inputType === 'historyUndo' || e.inputType === 'historyRedo') {
        e.preventDefault();
        if (e.inputType === 'historyUndo') undo();
        else redo();
        return;
      }
      if (!suppressed && !pendingSelection) pendingSelection = readSelection();
    });

    editor.addEventListener('input', (e) => {
      if (suppressed) return;
      record(TYPING.test(e.inputType || '') ? 'typing' : e.inputType || 'edit');
    });

    return {
      record,
      batch,
      undo,
      redo,
      reset,
      canUndo: () => undoStack.length > 0,
      canRedo: () => redoStack.length > 0,
      get size() {
        return undoStack.length;
      },
    };
  };
})(window.Wurd = window.Wurd || {});
