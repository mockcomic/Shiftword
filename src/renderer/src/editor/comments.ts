import { Mark, mergeAttributes } from "@tiptap/core";
import type { Mark as PMMark } from "@tiptap/pm/model";
import { TextSelection } from "@tiptap/pm/state";
import type { Editor } from "@tiptap/core";
import type { CommentRecord, CommentStore } from "../../../shared/types";
import { promptModal } from "../features/modal";
import { getSettings } from "../features/appearance";
import { firstRangeWithMarkId, rewriteMarkRanges } from "./mark-edit";

/**
 * Comments, Word-style: a `comment` mark carries the thread id; the thread
 * itself (author / text / resolved) lives in the `comments` attr of the doc
 * node. Ids — not positions — connect them, so page re-partitioning can
 * never strand a comment, and both halves ride in the document JSON through
 * page autosave and the .docx save.
 */

export const COMMENT_MARK = "comment";

export const CommentMarkExtension = Mark.create({
  name: COMMENT_MARK,
  inclusive: false,
  excludes: "",

  addAttributes() {
    return {
      id: {
        default: "",
        parseHTML: (element: HTMLElement) => element.getAttribute("data-comment-id") ?? "",
        renderHTML: (attrs: { id?: string }) => ({ "data-comment-id": attrs.id ?? "" }),
      },
      resolved: {
        default: false,
        parseHTML: (element: HTMLElement) => element.getAttribute("data-comment-resolved") === "true",
        renderHTML: (attrs: { resolved?: boolean }) =>
          attrs.resolved ? { "data-comment-resolved": "true" } : {},
      },
    };
  },

  renderHTML({ HTMLAttributes, mark }) {
    return [
      "span",
      mergeAttributes(HTMLAttributes, {
        class: mark.attrs.resolved ? "fw-comment fw-comment-resolved" : "fw-comment",
      }),
      0,
    ];
  },

  parseHTML() {
    return [
      {
        tag: "span[data-comment-id]",
        getAttrs: (element: HTMLElement | string) => {
          if (typeof element === "string") return false;
          return {
            id: element.getAttribute("data-comment-id") ?? "",
            resolved: element.getAttribute("data-comment-resolved") === "true",
          };
        },
      },
    ];
  },
});

/** The thread store from the doc attrs (defensively typed — it crosses IPC). */
export function readCommentStore(editor: Editor): CommentStore {
  const raw = (editor.state.doc.attrs?.comments ?? {}) as Record<string, Partial<CommentRecord>>;
  const out: CommentStore = {};
  for (const [id, rec] of Object.entries(raw)) {
    if (!rec || typeof rec !== "object") continue;
    out[id] = {
      author: typeof rec.author === "string" ? rec.author : "Unknown",
      text: typeof rec.text === "string" ? rec.text : "",
      createdAt: typeof rec.createdAt === "number" ? rec.createdAt : Date.now(),
      resolved: rec.resolved === true,
    };
  }
  return out;
}

function newCommentId(): string {
  return `c-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/**
 * Add a comment over the current selection (bubble-toolbar 💬): the prompt
 * opens on top of the selection — the range is captured first, so the modal
 * taking focus can't invalidate it. An empty selection gets nothing (Word
 * behaves the same).
 */
export async function addCommentToSelection(editor: Editor): Promise<string | null> {
  if (editor.isDestroyed) return null;
  const { selection, schema, doc } = editor.state;
  if (selection.empty) return null;
  const from = selection.from;
  const to = selection.to;
  const text = await promptModal({
    title: "New comment",
    label: "Comment",
    okLabel: "Comment",
    multiline: true,
    validate: (v) => (v.trim() ? null : "Write something to comment."),
  });
  if (text === null || editor.isDestroyed) return null;
  const id = newCommentId();
  const type = schema.marks[COMMENT_MARK];
  if (!type) return null;
  const store = readCommentStore(editor);
  store[id] = {
    author: getSettings().authorName || "You",
    text,
    createdAt: Date.now(),
    resolved: false,
  };
  const tr = editor.state.tr;
  tr.addMark(from, to, type.create({ id, resolved: false }));
  tr.setDocAttribute("comments", store);
  // Guard: the doc may have been replaced while the prompt was open.
  if (doc === editor.state.doc) editor.view.dispatch(tr);
  return id;
}

/** Rewrite every mark range carrying `id` with fresh attrs (resolve toggle). */
function setMarkResolvedAttr(editor: Editor, id: string, resolved: boolean): void {
  const { state } = editor;
  const type = state.schema.marks[COMMENT_MARK];
  if (!type) return;
  const ranges: Array<{ from: number; to: number }> = [];
  state.doc.descendants((node, pos) => {
    if (!node.isText) return;
    if (node.marks.some((m) => m.type === type && m.attrs.id === id)) {
      ranges.push({ from: pos, to: pos + node.nodeSize });
    }
  });
  if (ranges.length === 0) return;
  const tr = state.tr;
  for (const r of ranges) tr.addMark(r.from, r.to, type.create({ id, resolved }));
  editor.view.dispatch(tr);
}

/** Edit / resolve / unresolve a thread from the panel. */
export async function updateComment(
  editor: Editor,
  id: string,
  patch: { text?: string; resolved?: boolean },
): Promise<void> {
  if (editor.isDestroyed) return;
  const store = readCommentStore(editor);
  const rec = store[id];
  if (!rec) return;
  if (patch.text !== undefined && patch.text !== rec.text) {
    const next = await promptModal({
      title: "Edit comment",
      label: "Comment",
      value: rec.text,
      okLabel: "Save",
      multiline: true,
    });
    if (next === null || editor.isDestroyed) return;
    rec.text = next;
  }
  if (patch.resolved !== undefined) rec.resolved = patch.resolved;
  store[id] = rec;
  editor.view.dispatch(editor.state.tr.setDocAttribute("comments", store));
  if (patch.resolved !== undefined) setMarkResolvedAttr(editor, id, patch.resolved);
}

/** Delete the thread and every mark range referencing it. */
export function deleteComment(editor: Editor, id: string): void {
  if (editor.isDestroyed) return;
  rewriteMarkRanges(editor, (mark: PMMark) => mark.type.name === COMMENT_MARK && mark.attrs.id === id, {
    deleteText: false,
  });
  const store = readCommentStore(editor);
  delete store[id];
  editor.view.dispatch(editor.state.tr.setDocAttribute("comments", store));
}

/** Move the selection to the comment's first range (panel jump). */
export function jumpToComment(editor: Editor, id: string): void {
  if (editor.isDestroyed) return;
  const range = firstRangeWithMarkId(editor.state.schema, editor.state.doc, COMMENT_MARK, id);
  if (!range) return;
  try {
    const tr = editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range.from, range.to));
    editor.view.dispatch(tr.scrollIntoView());
  } catch {
    /* stale range — the next click gets a fresh one */
  }
}
