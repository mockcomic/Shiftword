import { Extension } from "@tiptap/core";
import type { Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

/**
 * Find & Replace (Ctrl+F / Ctrl+H), Word-style.
 *
 * Matches are computed over *text nodes* only — a match never spans a
 * hard break or a node boundary, which is exactly Word's behaviour when the
 * caret search runs out of a paragraph. Highlighting is a decoration plugin
 * (never an edit), while Replace dispatches real transactions so undo and
 * the dirty flag stay honest.
 *
 * The panel is a fixed-position DOM element outside the zoomed document
 * (same placement trick as the bubble toolbar), so the editor keeps its
 * selection and the match colours follow the theme instead of the zoom.
 */

export interface FindMatch {
  from: number;
  to: number;
}

interface FRState {
  query: string;
  caseSensitive: boolean;
  matches: FindMatch[];
  /** Index of the highlighted current match; -1 when none. */
  active: number;
}

const findKey = new PluginKey<FRState>("fwFindReplace");

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Walk every text node and collect case-(in)sensitive hits per node. */
function collectMatches(doc: PMNode, query: string, caseSensitive: boolean): FindMatch[] {
  if (!query) return [];
  const out: FindMatch[] = [];
  const re = new RegExp(escapeRegExp(query), caseSensitive ? "g" : "gi");
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(node.text)) !== null) {
      out.push({ from: pos + m.index, to: pos + m.index + m[0].length });
      if (m[0].length === 0) re.lastIndex++; // guard against zero-width loops
    }
  });
  return out;
}

function initialState(): FRState {
  return { query: "", caseSensitive: false, matches: [], active: -1 };
}

export const FindReplace = Extension.create({
  name: "fwFindReplace",

  addProseMirrorPlugins() {
    return [
      new Plugin<FRState>({
        key: findKey,
        state: {
          init: () => initialState(),
          apply(tr, prev, _old, next) {
            const meta = tr.getMeta(findKey) as
              { query?: string; caseSensitive?: boolean; clear?: boolean } | undefined;
            let { query, caseSensitive, active } = prev;
            if (meta?.clear) return initialState();
            if (meta) {
              query = meta.query ?? query;
              caseSensitive = meta.caseSensitive ?? caseSensitive;
            }
            const docChanged = tr.docChanged || meta !== undefined;
            const matches = docChanged ? collectMatches(next.doc, query, caseSensitive) : prev.matches;

            // The current match follows the caret: if the selection now sits
            // inside a match, that match becomes the active one (so typing or
            // clicking re-anchors F3 the way Word does).
            const sel = tr.selection;
            if (matches.length === 0) {
              active = -1;
            } else if (sel && !sel.empty) {
              const idx = matches.findIndex((mt) => mt.from === sel.from && mt.to === sel.to);
              active = idx >= 0 ? idx : Math.min(Math.max(active, 0), matches.length - 1);
            } else if (sel) {
              const idx = matches.findIndex((mt) => sel.from >= mt.from && sel.from < mt.to);
              if (idx >= 0) active = idx;
              else active = Math.min(Math.max(active, 0), matches.length - 1);
            }
            return { query, caseSensitive, matches, active };
          },
        },
        props: {
          decorations(state) {
            const fr = findKey.getState(state);
            if (!fr || fr.matches.length === 0) return DecorationSet.empty;
            return DecorationSet.create(
              state.doc,
              fr.matches.map((mt, i) =>
                Decoration.inline(mt.from, mt.to, {
                  class: i === fr.active ? "fw-match fw-match-active" : "fw-match",
                }),
              ),
            );
          },
        },
      }),
    ];
  },
});

// ---------------------------------------------------------------------------
// Commands (dispatch through the exported functions so the menu, the panel,
// and the tests all drive one identical path)
// ---------------------------------------------------------------------------

function getState(editor: Editor): FRState {
  return findKey.getState(editor.state) ?? initialState();
}

/** Set the query (or clear everything with an empty query + clear flag). */
function setQuery(editor: Editor, query: string, caseSensitive: boolean): void {
  editor.view.dispatch(editor.state.tr.setMeta(findKey, { query, caseSensitive }));
}

function clearFind(editor: Editor): void {
  editor.view.dispatch(editor.state.tr.setMeta(findKey, { clear: true }));
  refreshLabel(editor);
}

/** Move to the next match (wraps around) and select + scroll it into view. */
export function findNext(editor: Editor): void {
  const fr = getState(editor);
  if (fr.matches.length === 0) return;
  const target = (fr.active + 1) % fr.matches.length;
  selectMatch(editor, fr.matches[target]);
}

export function findPrev(editor: Editor): void {
  const fr = getState(editor);
  if (fr.matches.length === 0) return;
  const target = fr.active <= 0 ? fr.matches.length - 1 : fr.active - 1;
  selectMatch(editor, fr.matches[target]);
}

function selectMatch(editor: Editor, mt: FindMatch): void {
  const { state } = editor;
  try {
    const tr = state.tr.setSelection(TextSelection.create(state.doc, mt.from, mt.to));
    editor.view.dispatch(tr.scrollIntoView());
  } catch {
    /* a re-pagination between decoration paint and click can invalidate the
       range; the next keystroke refreshes the match list anyway */
  }
  refreshLabel(editor);
}

/** True when the current selection sits exactly on a match. */
function selectionOnMatch(editor: Editor): FindMatch | null {
  const fr = getState(editor);
  const sel = editor.state.selection;
  return fr.matches.find((mt) => mt.from === sel.from && mt.to === sel.to) ?? null;
}

/** Replace the current match (if selected) with the panel's text; advance. */
export function replaceCurrent(editor: Editor, replacement: string): void {
  const mt = selectionOnMatch(editor);
  if (!mt) {
    findNext(editor);
    return;
  }
  const tr = editor.state.tr.insertText(replacement, mt.from, mt.to);
  editor.view.dispatch(tr.scrollIntoView());
  refreshLabel(editor);
}

/** Replace every match in ONE transaction (single Ctrl+Z reverts all). */
export function replaceAll(editor: Editor, replacement: string): void {
  const fr = getState(editor);
  if (fr.matches.length === 0) return;
  const tr = editor.state.tr;
  for (let i = fr.matches.length - 1; i >= 0; i--) {
    const mt = fr.matches[i];
    tr.insertText(replacement, mt.from, mt.to);
  }
  editor.view.dispatch(tr);
  refreshLabel(editor);
}

// ---------------------------------------------------------------------------
// Panel
// ---------------------------------------------------------------------------

let panel: HTMLElement | null = null;
let editorRef: Editor | null = null;

export function closeFindPanel(): void {
  if (panel) {
    panel.remove();
    panel = null;
  }
  if (editorRef && !editorRef.isDestroyed) clearFind(editorRef);
}

/** Open (or re-show) the panel. `withReplace` reveals the replace row (Ctrl+H). */
export function openFindPanel(editor: Editor, withReplace: boolean): void {
  editorRef = editor;
  if (!panel) {
    panel = document.createElement("div");
    panel.id = "fw-find-panel";
    panel.className = "fw-find-panel";
    panel.innerHTML = `
      <div class="fr-row fr-search-row">
        <input id="fr-query" type="text" placeholder="Find" aria-label="Find" />
        <button id="fr-case" type="button" class="fr-toggle" title="Match case">Aa</button>
        <span id="fr-count" class="fr-count" aria-live="polite">0 / 0</span>
        <button id="fr-prev" type="button" title="Previous match (Shift+F3)">▲</button>
        <button id="fr-next" type="button" title="Next match (F3)">▼</button>
        <button id="fr-close" type="button" title="Close (Esc)">✕</button>
      </div>
      <div class="fr-row fr-replace-row">
        <input id="fr-replace" type="text" placeholder="Replace" aria-label="Replace" />
        <button id="fr-replace-one" type="button">Replace</button>
        <button id="fr-replace-all" type="button">All</button>
      </div>
    `;
    document.body.appendChild(panel);

    const queryInput = panel.querySelector<HTMLInputElement>("#fr-query")!;
    const replaceInput = panel.querySelector<HTMLInputElement>("#fr-replace")!;
    const caseBtn = panel.querySelector<HTMLButtonElement>("#fr-case")!;

    const runQuery = (): void => {
      const ed = editorRef;
      if (!ed || ed.isDestroyed) return;
      setQuery(ed, queryInput.value, caseBtn.classList.contains("active"));
      refreshLabel(ed);
      const fr = getState(ed);
      if (fr.matches.length > 0) {
        // Reveal the first match at or after the caret, like Word.
        const caret = ed.state.selection.from;
        const first = fr.matches.find((mt) => mt.to > caret) ?? fr.matches[0];
        selectMatch(ed, first);
      }
    };

    queryInput.addEventListener("input", runQuery);
    caseBtn.addEventListener("click", () => {
      caseBtn.classList.toggle("active");
      runQuery();
      queryInput.focus();
    });
    panel.querySelector("#fr-prev")!.addEventListener("click", () => {
      if (editorRef) findPrev(editorRef);
    });
    panel.querySelector("#fr-next")!.addEventListener("click", () => {
      if (editorRef) findNext(editorRef);
    });
    panel.querySelector("#fr-close")!.addEventListener("click", () => {
      const ed = editorRef;
      closeFindPanel();
      ed?.commands.focus(); // the ✕ must return the caret like Escape does
    });
    panel.querySelector("#fr-replace-one")!.addEventListener("click", () => {
      if (editorRef) replaceCurrent(editorRef, replaceInput.value);
    });
    panel.querySelector("#fr-replace-all")!.addEventListener("click", () => {
      if (editorRef) replaceAll(editorRef, replaceInput.value);
    });

    // Keys typed in the panel must never reach the editor's keymaps; Escape
    // closes it. Capture phase on the panel itself, plus a document-level
    // fallback so Escape works even when the editor holds focus.
    panel.addEventListener("keydown", (e) => {
      // While an IME composition is active, Enter/Escape belong to the
      // composition (commit/cancel candidates), not to the panel.
      if (e.isComposing || e.keyCode === 229) return;
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        closeFindPanel();
        editorRef?.commands.focus();
      } else if (e.key === "Enter") {
        e.preventDefault();
        const target = e.target as HTMLInputElement;
        if (target === replaceInput) {
          if (editorRef) replaceCurrent(editorRef as Editor, replaceInput.value);
        } else {
          if (editorRef) findNext(editorRef);
        }
      }
    });
  }

  panel.classList.toggle("with-replace", withReplace);
  const queryInput = panel.querySelector<HTMLInputElement>("#fr-query")!;
  queryInput.focus();
  queryInput.select();
  refreshLabel(editor);
}

/** Keep the `n / m` counter honest after any navigation or replacement. */
function refreshLabel(editor: Editor): void {
  if (!panel) return;
  const el = panel.querySelector<HTMLElement>("#fr-count");
  if (!el) return;
  const fr = getState(editor);
  el.textContent = fr.matches.length === 0 ? "No results" : `${fr.active + 1} / ${fr.matches.length}`;
  el.classList.toggle("empty", fr.matches.length === 0);
}

/** Menu-command entry point (app.ts dispatches `edit:find…` here). */
export function handleFindCommand(editor: Editor, id: string): void {
  switch (id) {
    case "edit:find":
      openFindPanel(editor, false);
      break;
    case "edit:replace":
      openFindPanel(editor, true);
      break;
    case "edit:findNext":
      findNext(editor);
      break;
    case "edit:findPrev":
      findPrev(editor);
      break;
  }
}

/** Escape closes the panel even while the editor (not the panel) has focus. */
if (typeof document !== "undefined") {
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && panel) {
      closeFindPanel();
    }
  });
}
