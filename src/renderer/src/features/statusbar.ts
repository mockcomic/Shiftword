import type { Editor } from "@tiptap/core";
import { getZoom } from "./zoom";

let debounceTimer: number | null = null;

/** Debounced status bar refresh (150 ms). */
export function scheduleStatus(editor: Editor): void {
  if (debounceTimer !== null) window.clearTimeout(debounceTimer);
  debounceTimer = window.setTimeout(() => {
    debounceTimer = null;
    if (!editor.isDestroyed) updateStatus(editor);
  }, 150);
}

/** Immediate status bar refresh: `Page X / Y`, word count, `DOCX · zoom%`. */
export function updateStatus(editor: Editor): void {
  const el = document.querySelector<HTMLElement>("#statusbar");
  if (!el || editor.isDestroyed) return;

  const doc = editor.state.doc;
  const text = doc.textBetween(0, doc.content.size, "\n", " ");
  const words = (text.match(/\S+/g) ?? []).length;

  // Page containing the selection head.
  let pageCount = 0;
  let currentPage = 1;
  const head = editor.state.selection.head;
  doc.forEach((page, pos) => {
    pageCount += 1;
    if (head >= pos) currentPage = pageCount;
  });

  // The page readout stays in the DOM even in Notion mode (where the CSS
  // hides it): the `--smoke-test` probe greps for `Page X / Y`, and it must
  // keep passing whichever mode the user last persisted.
  const notion = document.body.classList.contains("notion-mode");
  const tracking = doc.attrs?.trackChanges === true;
  el.innerHTML =
    `<span class="sb-left sb-page">Page ${currentPage} / ${pageCount}</span>` +
    `<span class="sb-center">${words} words</span>` +
    `<span class="sb-right">${tracking ? '<span class="sb-track" title="Track changes is on">● Tracking</span> ' : ""}${notion ? "Notion" : "DOCX"} · ${getZoom()}%</span>`;
}
