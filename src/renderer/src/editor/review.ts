import { Extension, Mark, mergeAttributes } from "@tiptap/core";
import { Fragment, Slice } from "@tiptap/pm/model";
import type { Mark as PMMark, Node as PMNode, Schema } from "@tiptap/pm/model";
import { NodeSelection, Plugin, PluginKey } from "@tiptap/pm/state";
import type { Editor } from "@tiptap/core";
import type { EditorView } from "@tiptap/pm/view";
import { getSettings } from "../features/appearance";
import { markAroundSelection, rewriteMarkRanges } from "./mark-edit";

/**
 * Track Changes (Word's review model on our own primitives):
 *
 * - `trackedInsert` / `trackedDelete` are ordinary inline marks carrying
 *   `{ author, date }`, so the revision state survives page autosave and
 *   lands in the document JSON.
 * - While the `trackChanges` doc attr is on, the capture plugin rewrites the
 *   obvious edit paths — typing (`handleTextInput`), pasting
 *   (`transformPasted`) and deleting (`beforeinput`) — so a "deletion" is
 *   really a mark, and rejected text can come back verbatim.
 * - Accept/reject are mark operations routed through `mark-edit.ts`, one
 *   transaction each (one Ctrl+Z reverts one operation).
 *
 * Deletions that hit non-text content (an image, a whole table, runs with
 * mixed block content) fall through to a real delete — matching the v1
 * scope note: revisions are character-level.
 */

export const TRACK_INSERT = "trackedInsert";
export const TRACK_DELETE = "trackedDelete";

function trackMark(name: string, cssClass: string) {
  return Mark.create({
    name,
    inclusive: false,
    excludes: "",

    addAttributes() {
      return {
        author: {
          default: "",
          parseHTML: (element: HTMLElement) => element.getAttribute("data-author") ?? "",
          renderHTML: (attrs: { author?: string }) => ({ "data-author": attrs.author ?? "" }),
        },
        date: {
          default: "",
          parseHTML: (element: HTMLElement) => element.getAttribute("data-date") ?? "",
          renderHTML: (attrs: { date?: string }) => ({ "data-date": attrs.date ?? "" }),
        },
      };
    },

    renderHTML({ HTMLAttributes }) {
      return ["span", mergeAttributes(HTMLAttributes, { class: cssClass, "data-track": name }), 0];
    },

    parseHTML() {
      return [
        {
          tag: `span[data-track="${name}"]`,
          getAttrs: (element: HTMLElement | string) => {
            if (typeof element === "string") return false;
            return {
              author: element.getAttribute("data-author") ?? "",
              date: element.getAttribute("data-date") ?? "",
            };
          },
        },
      ];
    },
  });
}

export const TrackedInsert = trackMark(TRACK_INSERT, "fw-ins");
export const TrackedDelete = trackMark(TRACK_DELETE, "fw-del");

export function isTrackingEnabled(editor: Editor): boolean {
  return editor.state.doc.attrs?.trackChanges === true;
}

/** Review-mode toggle (Review → Track Changes): lives on the doc attrs. */
export function setTracking(editor: Editor, on: boolean): void {
  if (editor.isDestroyed) return;
  editor.view.dispatch(editor.state.tr.setDocAttribute("trackChanges", on));
}

function revisionInfo(): { author: string; date: string } {
  return {
    author: getSettings().authorName || "You",
    date: new Date().toISOString(),
  };
}

/** Rebuild a pasted fragment with the insert mark stamped on every text node. */
function stampFragment(fragment: Fragment, schema: Schema, mark: PMMark): Fragment {
  const nodes: PMNode[] = [];
  fragment.forEach((node) => {
    if (node.isText) {
      nodes.push(schema.text(String(node.text ?? ""), [...node.marks, mark]));
    } else if (node.isInline) {
      nodes.push(node.type.create(node.attrs, stampFragment(node.content, schema, mark), node.marks));
    } else {
      nodes.push(node.copy(stampFragment(node.content, schema, mark)));
    }
  });
  return Fragment.fromArray(nodes);
}

export const ReviewTracking = Extension.create({
  name: "fwReviewTracking",

  addProseMirrorPlugins() {
    const editor = this.editor;
    const enabled = (): boolean => !editor.isDestroyed && isTrackingEnabled(editor);

    return [
      new Plugin({
        key: new PluginKey("fwReviewTracking"),
        props: {
          handleTextInput: (view: EditorView, from: number, to: number, text: string): boolean => {
            if (!enabled()) return false;
            const { state } = view;
            const type = state.schema.marks[TRACK_INSERT];
            if (!type) return false;
            const tr = state.tr;
            tr.insertText(text, from, to);
            tr.addMark(from, from + text.length, type.create(revisionInfo()));
            view.dispatch(tr);
            return true;
          },

          transformPasted: (slice: Slice, view: EditorView): Slice => {
            if (!enabled()) return slice;
            const type = view.state.schema.marks[TRACK_INSERT];
            if (!type) return slice;
            const mark = type.create(revisionInfo());
            return new Slice(
              stampFragment(slice.content, view.state.schema, mark),
              slice.openStart,
              slice.openEnd,
            );
          },

          handleDOMEvents: {
            beforeinput: (view: EditorView, event: Event): boolean => {
              const input = event as InputEvent;
              const type = input.inputType ?? "";
              if (!/^delete/.test(type) && type !== "cut") return false;
              if (!enabled() || view.state.selection instanceof NodeSelection) return false;
              const { state } = view;
              const sel = state.selection;
              let from = sel.from;
              let to = sel.to;
              if (sel.empty) {
                const $pos = sel.$from;
                if ($pos.depth === 0) return false;
                if (type === "deleteContentBackward") {
                  if ($pos.parentOffset === 0) return false; // paragraph merge stays hard
                  from = $pos.pos - 1;
                  to = $pos.pos;
                } else {
                  if ($pos.parentOffset >= $pos.parent.content.size) return false; // end of block
                  from = $pos.pos;
                  to = $pos.pos + 1;
                }
              }
              // Character-level only: any non-text *leaf* (image, hard break,
              // footnote ref, table) makes the whole delete a hard delete.
              // `nodesBetween` also visits the block ancestors — they descend.
              let blocked = false;
              state.doc.nodesBetween(from, to, (node) => {
                if (blocked) return false;
                if (node.isText) return false;
                if (node.isInline) {
                  blocked = true;
                  return false;
                }
                return true;
              });
              if (blocked) return false;
              const markType = state.schema.marks[TRACK_DELETE];
              if (!markType) return false;
              // A suppressed `cut` never reaches the browser's clipboard
              // step: without this, Ctrl+X in track-changes mode deletes
              // (softly) AND loses the text — nothing lands on the clipboard.
              if (type === "cut") {
                const text = state.doc.textBetween(from, to, "\n", " ");
                void navigator.clipboard?.writeText(text).catch(() => undefined);
              }
              view.dispatch(state.tr.addMark(from, to, markType.create(revisionInfo())));
              event.preventDefault();
              return true;
            },
          },
        },
      }),
    ];
  },
});

const isTrackMark = (m: PMMark): boolean => m.type.name === TRACK_INSERT || m.type.name === TRACK_DELETE;

/** Accept All Changes: revisions become permanent text (marks cleared). */
export function acceptAllChanges(editor: Editor): void {
  rewriteMarkRanges(editor, isTrackMark, { deleteText: false });
}

/** Reject All Changes: insertions vanish, deletions come back clean. */
export function rejectAllChanges(editor: Editor): void {
  rewriteMarkRanges(editor, (m) => m.type.name === TRACK_INSERT, { deleteText: true });
  rewriteMarkRanges(editor, (m) => m.type.name === TRACK_DELETE, { deleteText: false });
}

/** Accept the change under the caret (clears just that mark). */
export function acceptChangeAtCursor(editor: Editor): boolean {
  const hit = markAroundSelection(editor, isTrackMark);
  if (!hit) return false;
  return rewriteMarkRanges(editor, (m) => m === hit.mark, { deleteText: false });
}

/** Reject the change under the caret (drop an insertion / restore a deletion). */
export function rejectChangeAtCursor(editor: Editor): boolean {
  const hit = markAroundSelection(editor, isTrackMark);
  if (!hit) return false;
  const insert = hit.mark.type.name === TRACK_INSERT;
  return rewriteMarkRanges(editor, (m) => m === hit.mark, { deleteText: insert });
}
