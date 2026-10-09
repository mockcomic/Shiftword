import type { Editor } from "@tiptap/core";
import { deleteComment, jumpToComment, readCommentStore, updateComment } from "../editor/comments";

/**
 * The Comments rail (View → Comments Panel).
 *
 * A fixed element beside the canvas (not inside the zoomed `.fw-doc`) that
 * mirrors the comment store from the doc attrs: jump, edit, resolve, delete.
 * Everything flows back through transactions, so the store, the marks and
 * the dirty/autosave state never drift apart.
 */

let panelEl: HTMLElement | null = null;
/** Serialized store the DOM currently shows — the panel is rebuilt only when it changes. */
let renderedSignature = "";

export function isCommentsPanelOpen(): boolean {
  return panelEl !== null;
}

export function closeCommentsPanel(): void {
  if (!panelEl) return;
  panelEl.remove();
  panelEl = null;
  renderedSignature = "";
}

export function toggleCommentsPanel(editor: Editor): void {
  if (panelEl) closeCommentsPanel();
  else openCommentsPanel(editor);
}

export function openCommentsPanel(editor: Editor): void {
  if (editor.isDestroyed) return;
  if (!panelEl) {
    panelEl = document.createElement("aside");
    panelEl.id = "fw-comments";
    panelEl.innerHTML = `<h2>Comments</h2><div id="fw-comments-list"></div>`;
    document.body.appendChild(panelEl);
    // One delegated listener for the whole rail: every action derives the
    // card id and the store state at click time, so re-rendering never has
    // to re-attach handlers (or leak the old ones).
    panelEl.addEventListener("click", (e) => {
      const btn = (e.target as HTMLElement).closest?.("button[data-act]") as HTMLElement | null;
      if (!btn) return;
      const card = btn.closest<HTMLElement>(".fw-comment-card");
      const id = card?.dataset.id ?? "";
      if (!id || editor.isDestroyed) return;
      const act = btn.dataset.act;
      if (act === "jump") jumpToComment(editor, id);
      else if (act === "edit")
        void updateComment(editor, id, { text: "" }); // text:"" opens the prompt with the current value
      else if (act === "resolve") {
        const rec = readCommentStore(editor)[id];
        void updateComment(editor, id, { resolved: !(rec?.resolved ?? false) });
      } else if (act === "delete") deleteComment(editor, id);
    });
  }
  refreshCommentsPanel(editor);
}

/** Re-render the rail from the store — called on every editor update. */
export function refreshCommentsPanel(editor: Editor): void {
  if (!panelEl || editor.isDestroyed) return;
  const list = panelEl.querySelector<HTMLElement>("#fw-comments-list");
  if (!list) return;
  const store = readCommentStore(editor);
  const entries = Object.entries(store).sort((a, b) => b[1].createdAt - a[1].createdAt);
  // This runs from the editor `onUpdate` on EVERY keystroke. The comment
  // store only changes through comment operations, so rebuild the DOM (and
  // reset the panel's scroll position) only when it actually differs.
  const signature = JSON.stringify(entries);
  if (signature === renderedSignature) return;
  renderedSignature = signature;
  if (entries.length === 0) {
    list.innerHTML = `<p class="fw-comments-empty">No comments. Select text and hit 💬.</p>`;
    return;
  }
  list.innerHTML = entries
    .map(
      ([id, rec]) => `
      <article class="fw-comment-card${rec.resolved ? " resolved" : ""}" data-id="${id}">
        <header><strong>${escapeHtml(rec.author)}</strong><time>${new Date(rec.createdAt).toLocaleString()}</time></header>
        <p>${escapeHtml(rec.text)}</p>
        <footer>
          <button type="button" data-act="jump">Jump</button>
          <button type="button" data-act="edit">Edit</button>
          <button type="button" data-act="resolve">${rec.resolved ? "Unresolve" : "Resolve"}</button>
          <button type="button" data-act="delete">Delete</button>
        </footer>
      </article>`,
    )
    .join("");
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
