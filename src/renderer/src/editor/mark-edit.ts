import type { Editor } from "@tiptap/core";
import type { Mark, Node as PMNode, Schema } from "@tiptap/pm/model";

/**
 * Per-mark removal helpers.
 *
 * ProseMirror only offers `deleteMark`, which strips *every* mark from a
 * range — useless when a run carries `bold` *and* `comment` (or a tracked
 * mark). The correct operation is to rebuild the text node with the wanted
 * mark list, so every "remove one mark" flow (accept/reject revisions,
 * delete comment, ...) funnels through here. Ranges are collected first and
 * applied in descending order so each original position stays valid on the
 * shared transaction.
 */

function textNodesWithMark(
  doc: PMNode,
  predicate: (mark: Mark) => boolean,
): Array<{ from: number; to: number; node: PMNode; mark: Mark }> {
  const out: Array<{ from: number; to: number; node: PMNode; mark: Mark }> = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.marks.length) return;
    for (const mark of node.marks) {
      if (predicate(mark)) {
        out.push({ from: pos, to: pos + node.nodeSize, node, mark });
        return;
      }
    }
  });
  return out;
}

/**
 * Rebuild every text node carrying a matching mark without that mark
 * (`deleteText: true` removes the text itself instead — the reject-insertion
 * path). One transaction: one Ctrl+Z reverts the whole operation.
 */
export function rewriteMarkRanges(
  editor: Editor,
  predicate: (mark: Mark) => boolean,
  opts: { deleteText: boolean },
): boolean {
  const { state } = editor;
  const hits = textNodesWithMark(state.doc, predicate);
  if (hits.length === 0) return false;
  const tr = state.tr;
  for (let i = hits.length - 1; i >= 0; i--) {
    const hit = hits[i];
    if (opts.deleteText) {
      tr.delete(hit.from, hit.to);
      continue;
    }
    const kept = hit.node.marks.filter((m) => !predicate(m));
    // A text node with an *empty* mark list is still a plain text node — only
    // `deleteText` may remove the text itself.
    const fresh = state.schema.text(String(hit.node.text ?? ""), kept);
    tr.replaceWith(hit.from, hit.to, fresh);
  }
  editor.view.dispatch(tr);
  return true;
}

/** True when the node at the selection head carries a matching mark. */
export function markAroundSelection(
  editor: Editor,
  predicate: (mark: Mark) => boolean,
): { node: PMNode; from: number; mark: Mark } | null {
  const { state } = editor;
  const head = state.selection.head;
  const $head = state.doc.resolve(head);
  // The caret sits between two text nodes after typing: prefer the node just
  // before the caret, then the one after (Word's "change under cursor").
  const candidates: Array<{ parent: PMNode; parentStart: number; offset: number }> = [];
  if ($head.depth > 0) {
    const parentStart = head - $head.parentOffset;
    candidates.push({ parent: $head.parent, parentStart, offset: $head.parentOffset - 1 });
    candidates.push({ parent: $head.parent, parentStart, offset: $head.parentOffset });
  }
  for (const cand of candidates) {
    let from = 0;
    let index = 0;
    let found: PMNode | null = null;
    cand.parent.forEach((child, childOffset) => {
      if (found) return;
      const start = cand.parentStart + 1 + childOffset;
      if (cand.offset >= index && cand.offset < index + child.nodeSize) {
        found = child;
        from = start;
        return;
      }
      index += child.nodeSize;
    });
    if (found) {
      const node = found as PMNode;
      for (const mark of node.marks) {
        if (predicate(mark)) return { node, from, mark };
      }
    }
  }
  return null;
}

/** All document ranges carrying a given mark type, for panel jump-to. */
export function firstRangeWithMarkId(
  schema: Schema,
  doc: PMNode,
  markName: string,
  id: string,
): { from: number; to: number } | null {
  const type = schema.marks[markName];
  if (!type) return null;
  let found: { from: number; to: number } | null = null;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (!node.isText) return;
    if (node.marks.some((m) => m.type === type && m.attrs.id === id))
      found = { from: pos, to: pos + node.nodeSize };
  });
  return found;
}
