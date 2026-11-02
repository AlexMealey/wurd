// Moving and resizing text boxes. A box is either in line with the text (it takes up space in the
// flow and is dragged between paragraphs) or free (placed anywhere on the page, in front of text
// and images). A free box is anchored before the paragraph it sits over, so it moves with that
// paragraph; margin-top is its offset from there and left is a share of the page width.
//
//   in line: <div class="wurd-textbox" style="width: 60%; min-height: 54pt">
//   free:    <div class="wurd-textbox wurd-textbox--float" style="left: 30%; margin-top: 12pt; width: 40%">
(function (Wurd) {
  'use strict';

  const FLOAT = 'wurd-textbox--float';
  const MIN_WIDTH_PX = 48;
  const MIN_HEIGHT_PX = 24;
  const PT_PER_PX = 0.75;
  const round = (n, places = 2) => Math.round(n * 10 ** places) / 10 ** places;

  Wurd.isFloatingBox = (el) => Boolean(el && Wurd.isTextBox(el) && el.classList.contains(FLOAT));

  let editor = null;
  let history = null;
  let getScale = () => 1;
  let tools = null;
  let moveButton = null;
  let floatButton = null;
  let resizeButton = null;
  let dropLine = null;
  let active = null;
  let dragging = false;

  const isFloating = (box) => box.classList.contains(FLOAT);

  const contentWidth = () => {
    const style = window.getComputedStyle(editor);
    return editor.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
  };

  // A free box's width is a share of the page; an in-line box's, of the text column
  const basisWidth = (box) => (isFloating(box) ? editor.clientWidth : contentWidth());

  const tidyStyle = (box) => {
    if (!box.getAttribute('style')) box.removeAttribute('style');
  };

  // The blocks that make up the flow of the text (free boxes don't)
  const flowBlocks = (except) => Array.from(editor.children).filter((el) => el !== except && !Wurd.isFloatingBox(el));

  // Before a block, but ahead of any free boxes anchored to it, so they stay where they are
  const beforeBlock = (block, box) => {
    if (!block) {
      editor.append(box);
      return;
    }
    let ref = block;
    while (ref.previousElementSibling && ref.previousElementSibling !== box && Wurd.isFloatingBox(ref.previousElementSibling)) ref = ref.previousElementSibling;
    ref.before(box);
  };

  // Puts a free box's top-left corner at (left, top), in pixels from the page's top-left corner
  const placeFloat = (box, left, top) => {
    const x = Math.min(Math.max(0, left), Math.max(0, editor.clientWidth - box.offsetWidth));
    const y = Math.max(0, top);
    const anchor = flowBlocks(box).find((el) => el.offsetTop + el.offsetHeight > y) || null;
    if (anchor) anchor.before(box);
    else editor.append(box);
    // Measure where it sits with no offset, then offset it from there
    box.style.marginTop = '0px';
    box.style.marginTop = `${round((y - box.offsetTop) * PT_PER_PX, 1)}pt`;
    box.style.left = `${round((x / editor.clientWidth) * 100)}%`;
  };

  const setWidth = (box, px) => {
    const pct = round((Math.max(MIN_WIDTH_PX, px) / basisWidth(box)) * 100);
    if (!isFloating(box) && pct >= 99.5) box.style.removeProperty('width');
    else box.style.width = `${Math.min(pct, 100)}%`;
  };

  const setHeight = (box, px) => {
    box.style.minHeight = `${round(Math.max(MIN_HEIGHT_PX, px) * PT_PER_PX, 1)}pt`;
  };

  // Free or in line, keeping the box where it is on the page
  const setFloating = (box, on) => {
    if (on === isFloating(box)) return;
    const { offsetLeft: left, offsetTop: top, offsetWidth: width } = box;
    if (on) {
      box.classList.add(FLOAT);
      box.style.width = `${round((width / editor.clientWidth) * 100)}%`;
      placeFloat(box, left, top);
    } else {
      box.classList.remove(FLOAT);
      ['left', 'top', 'margin-top'].forEach((p) => box.style.removeProperty(p));
      setWidth(box, width);
      Wurd.insert.ensureParagraphAfter(box);
    }
    tidyStyle(box);
  };

  // Where an in-line box let go at height y goes: before the block returned, or at the end (null)
  const inlineTarget = (box, y) => flowBlocks(box).find((el) => el.offsetTop + el.offsetHeight / 2 > y) || null;

  const moveInline = (box, before) => {
    beforeBlock(before, box);
    Wurd.insert.ensureParagraphAfter(box);
  };

  // ---- The handles drawn over the active box --------------------------------

  const origin = () => ({ x: editor.offsetLeft + editor.clientLeft, y: editor.offsetTop + editor.clientTop });

  const position = () => {
    if (!tools) return;
    if (!active || !active.isConnected || editor.hidden || !editor.contains(active)) {
      active = null;
      tools.hidden = true;
      return;
    }
    const o = origin();
    tools.style.left = `${o.x + active.offsetLeft}px`;
    tools.style.top = `${o.y + active.offsetTop}px`;
    tools.style.width = `${active.offsetWidth}px`;
    tools.style.height = `${active.offsetHeight}px`;
    floatButton.setAttribute('aria-pressed', String(isFloating(active)));
    tools.hidden = false;
  };

  const showDrop = (box, target) => {
    const o = origin();
    const blocks = flowBlocks(box);
    const last = blocks[blocks.length - 1];
    const y = target ? target.offsetTop : last ? last.offsetTop + last.offsetHeight : 0;
    const style = window.getComputedStyle(editor);
    dropLine.style.left = `${o.x + parseFloat(style.paddingLeft)}px`;
    dropLine.style.top = `${o.y + y - 1}px`;
    dropLine.style.width = `${contentWidth()}px`;
    dropLine.hidden = false;
  };

  // Follows the selection: the box the caret is in gets the handles
  const update = () => {
    if (!tools) return;
    const selection = window.getSelection();
    const node = selection.rangeCount ? selection.anchorNode : null;
    const el = node && (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement);
    const box = el && editor.contains(el) ? el.closest('.wurd-textbox') : null;
    if (box) active = box;
    else if (!dragging && !tools.contains(document.activeElement)) active = null;
    position();
  };

  const commit = () => {
    history.record('text-box');
    position();
  };

  // Dragging the move handle or the resize corner
  const beginDrag = (e, kind) => {
    if (!active || e.button !== 0) return;
    e.preventDefault();
    const box = active;
    const handle = e.currentTarget;
    handle.setPointerCapture(e.pointerId);
    const scale = getScale();
    const rect = box.getBoundingClientRect();
    const start = { x: e.clientX, y: e.clientY, left: box.offsetLeft, top: box.offsetTop, width: box.offsetWidth, height: box.offsetHeight, grabY: (e.clientY - rect.top) / scale };
    const floating = isFloating(box);
    let target;
    let moved = false;
    dragging = true;
    box.classList.add('is-dragging');
    const delta = (ev) => [(ev.clientX - start.x) / scale, (ev.clientY - start.y) / scale];
    const onMove = (ev) => {
      const [dx, dy] = delta(ev);
      if (Math.abs(dx) + Math.abs(dy) > 3) moved = true;
      if (kind === 'resize') {
        setWidth(box, start.width + dx);
        setHeight(box, start.height + dy);
      } else {
        box.style.transform = `translate(${dx}px, ${dy}px)`;
        if (!floating) {
          target = inlineTarget(box, start.top + start.grabY + dy);
          showDrop(box, target);
        }
      }
      position();
    };
    const finish = (ev) => {
      handle.removeEventListener('pointermove', onMove);
      handle.removeEventListener('pointerup', finish);
      handle.removeEventListener('pointercancel', finish);
      box.classList.remove('is-dragging');
      box.style.removeProperty('transform');
      dropLine.hidden = true;
      dragging = false;
      if (kind === 'move' && moved && ev.type === 'pointerup') {
        const [dx, dy] = delta(ev);
        if (floating) placeFloat(box, start.left + dx, start.top + dy);
        else moveInline(box, target);
      }
      tidyStyle(box);
      commit();
    };
    handle.addEventListener('pointermove', onMove);
    handle.addEventListener('pointerup', finish);
    handle.addEventListener('pointercancel', finish);
  };

  // Arrow keys on the move handle: a free box moves (Shift for 1px steps); an in-line box moves
  // up or down past a paragraph
  const onMoveKey = (e) => {
    if (!active) return;
    const box = active;
    if (isFloating(box)) {
      const d = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
      if (!d) return;
      e.preventDefault();
      const step = e.shiftKey ? 1 : 8;
      placeFloat(box, box.offsetLeft + d[0] * step, box.offsetTop + d[1] * step);
    } else {
      if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
      e.preventDefault();
      const blocks = flowBlocks(box);
      const next = blocks.findIndex((el) => box.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING);
      const index = next < 0 ? blocks.length : next;
      if (e.key === 'ArrowUp' && index > 0) moveInline(box, blocks[index - 1]);
      else if (e.key === 'ArrowDown' && index < blocks.length) moveInline(box, blocks[index + 1] || null);
      else return;
      box.scrollIntoView({ block: 'nearest' });
    }
    commit();
  };

  // Arrow keys on the resize corner: Left/Right for width, Up/Down for height (Shift for small steps)
  const onResizeKey = (e) => {
    if (!active) return;
    const box = active;
    const step = e.shiftKey ? 0.01 : 0.05;
    if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') setWidth(box, box.offsetWidth + (e.key === 'ArrowRight' ? 1 : -1) * step * basisWidth(box));
    else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') setHeight(box, box.offsetHeight + (e.key === 'ArrowDown' ? 1 : -1) * (e.shiftKey ? 2 : 12));
    else return;
    e.preventDefault();
    commit();
  };

  const init = (options) => {
    ({ editor, history, getScale } = options);
    tools = document.getElementById('box-tools');
    moveButton = document.getElementById('box-move');
    floatButton = document.getElementById('box-float');
    resizeButton = document.getElementById('box-resize');
    dropLine = document.getElementById('box-drop');
    // The handles mustn't take the caret out of the box
    [moveButton, floatButton, resizeButton].forEach((button) => button.addEventListener('mousedown', (e) => e.preventDefault()));
    moveButton.addEventListener('pointerdown', (e) => beginDrag(e, 'move'));
    resizeButton.addEventListener('pointerdown', (e) => beginDrag(e, 'resize'));
    moveButton.addEventListener('keydown', onMoveKey);
    resizeButton.addEventListener('keydown', onResizeKey);
    floatButton.addEventListener('click', () => {
      if (!active) return;
      setFloating(active, !isFloating(active));
      commit();
    });
    tools.addEventListener('focusout', () => setTimeout(update, 0));
    window.addEventListener('resize', position);
  };

  // Where each top-level free box sits on its page, in document order: { topPt } from the top of its
  // page's text (the editor's content edge, or just after a page break)
  const layout = () => {
    if (!editor || editor.hidden) return null;
    let pageTop = parseFloat(window.getComputedStyle(editor).paddingTop);
    const out = [];
    for (const el of editor.children) {
      if (Wurd.isPageBreak(el)) pageTop = el.offsetTop + el.offsetHeight + parseFloat(window.getComputedStyle(el).marginBottom);
      else if (Wurd.isFloatingBox(el)) out.push({ topPt: round((el.offsetTop - pageTop) * PT_PER_PX, 1) });
    }
    return out;
  };

  Wurd.textboxes = { init, update, position, layout, setFloating, placeFloat, isFloating: (box) => Wurd.isFloatingBox(box) };
})(window.Wurd = window.Wurd || {});
