import { mergeAttributes, Node } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import type { Editor } from "@tiptap/core";
import { promptModal } from "../features/modal";

/**
 * Footnotes (Word's model, condensed): the note text lives in the atom's
 * `text` attr, so it survives pagination, page autosave and the .docx
 * round-trip without a side store. The superscript number is never stored:
 * it is the ref's rank among all refs in document order, recomputed by the
 * NodeView and by the page-end note area (a widget decoration at each
 * page's bottom). Word semantics for "footnotes at the bottom of the page"
 * without touching the pagination engine — the area is absolutely
 * positioned, so it never feeds a measured block height.
 *
 * The note text is plain text in v1 (no marks inside footnotes), matching
 * the "readable, if plainer" posture of the other custom nodes.
 */

export const FOOTNOTE_REF = "footnoteRef";

/** Every footnoteRef in document order (rank = index + 1). */
export function collectFootnotes(doc: PMNode): Array<{ pos: number; text: string }> {
  const out: Array<{ pos: number; text: string }> = [];
  doc.descendants((node, pos) => {
    if (node.type.name === FOOTNOTE_REF) {
      out.push({ pos, text: String(node.attrs.text ?? "") });
    }
  });
  return out;
}

/** Cheap identity for "the footnote set did / did not change". */
function footnoteSignature(doc: PMNode): string {
  return collectFootnotes(doc)
    .map((f) => `${f.pos}:${f.text}`)
    .join("|");
}

/**
 * Live marker painters, registered by each NodeView.
 *
 * A NodeView only gets `update()` when *its own* node changes — adding or
 * deleting ANOTHER footnote changes every other marker's rank without
 * touching them, so the old numbering would go stale ([2] staying [2]
 * after the real #2 vanished). The plugin's view.update re-paints the
 * registry whenever the collected set changes.
 */
const livePaints = new Set<() => void>();

/** The page-bottom `n. note text` list, rebuilt per decoration pass. */
function buildNoteArea(entries: Array<{ ordinal: number; text: string }>): HTMLElement {
  const area = document.createElement("div");
  area.className = "fw-page-footnotes";
  area.setAttribute("contenteditable", "false");
  const rule = document.createElement("hr");
  rule.className = "fw-fn-rule";
  area.appendChild(rule);
  for (const entry of entries) {
    const row = document.createElement("div");
    row.className = "fw-fn-entry";
    const sup = document.createElement("sup");
    sup.textContent = String(entry.ordinal);
    row.appendChild(sup);
    row.appendChild(document.createTextNode(` ${entry.text}`));
    area.appendChild(row);
  }
  return area;
}

export const FootnoteRef = Node.create({
  name: FOOTNOTE_REF,
  inline: true,
  group: "inline",
  atom: true,
  draggable: false,

  addAttributes() {
    return {
      text: {
        default: "",
        parseHTML: (element: HTMLElement) => element.getAttribute("data-fn") ?? "",
        renderHTML: (attrs: { text?: string }) => ({ "data-fn": attrs.text ?? "" }),
      },
    };
  },

  // Static rendering (clipboard copy, HTML export): the note text in
  // brackets. The live editor shows the auto-numbered marker via NodeView.
  renderHTML({ HTMLAttributes, node }) {
    return [
      "sup",
      mergeAttributes(HTMLAttributes, { class: "fw-fn-ref", "data-type": FOOTNOTE_REF }),
      `[${String(node.attrs.text ?? "")}]`,
    ];
  },

  parseHTML() {
    return [{ tag: `sup[data-type="${FOOTNOTE_REF}"]`, priority: 60 }];
  },

  addNodeView() {
    return ({ editor, getPos }) => {
      const dom = document.createElement("sup");
      dom.className = "fw-fn-ref";
      dom.setAttribute("data-type", FOOTNOTE_REF);
      dom.setAttribute("contenteditable", "false");

      const paint = (): void => {
        if (editor.isDestroyed) return;
        const all = collectFootnotes(editor.state.doc);
        const pos = getPos();
        const rank = typeof pos === "number" ? all.findIndex((f) => f.pos === pos) + 1 : 0;
        dom.textContent = `[${rank > 0 ? rank : "?"}]`;
      };

      dom.addEventListener("click", () => {
        if (editor.isDestroyed) return;
        const pos = getPos();
        if (typeof pos !== "number") return;
        const target = editor.state.doc.nodeAt(pos);
        if (!target) return;
        void (async () => {
          const value = await promptModal({
            title: "Footnote",
            label: "Note text",
            value: String(target.attrs.text ?? ""),
            okLabel: "Save",
            multiline: true,
          });
          if (value === null || editor.isDestroyed) return;
          const tr = editor.state.tr.setNodeMarkup(pos, undefined, { ...target.attrs, text: value });
          editor.view.dispatch(tr);
        })();
      });

      paint();
      livePaints.add(paint);
      return {
        dom,
        update: (node: PMNode) => {
          void node;
          paint();
          return true;
        },
        ignoreMutation: () => true,
        destroy: () => {
          livePaints.delete(paint);
        },
      };
    };
  },

  addProseMirrorPlugins() {
    let decoSignature = "";
    let decoCache: { doc: PMNode; set: DecorationSet } | null = null;
    // Tracked SEPARATELY from the decoration cache: `props.decorations` is
    // consulted before the plugin view's `update`, so a shared variable
    // would make the decoration pass consume the change and the re-rank
    // never run.
    let paintSignature = "";
    return [
      new Plugin({
        key: new PluginKey("fwFootnoteArea"),
        props: {
          decorations(state) {
            const { doc } = state;
            // A fresh DecorationSet (and fresh note-area DOM) on EVERY
            // transaction — every keystroke, every selection move — made
            // prosemirror-view recreate the page-end widgets each time,
            // because widget `toDOM` functions compare equal only by
            // reference. The set only depends on the footnote collection,
            // so it is cached until that changes.
            const signature = footnoteSignature(doc);
            if (decoCache && decoCache.doc === doc && signature === decoSignature) return decoCache.set;
            decoSignature = signature;
            const all = collectFootnotes(doc);
            if (all.length === 0) {
              decoCache = { doc, set: DecorationSet.empty };
              return decoCache.set;
            }
            const widgets: Decoration[] = [];
            doc.forEach((page, pagePos) => {
              if (page.type.name !== "page") return;
              const entries = all
                .map((f, i) => ({ ...f, ordinal: i + 1 }))
                .filter((f) => f.pos > pagePos && f.pos < pagePos + page.nodeSize);
              if (entries.length === 0) return;
              const area = buildNoteArea(entries);
              widgets.push(
                Decoration.widget(pagePos + page.nodeSize - 1, () => area, {
                  side: 1,
                  ignoreSelection: true,
                }),
              );
            });
            decoCache = { doc, set: DecorationSet.create(doc, widgets) };
            return decoCache.set;
          },
        },
        view: () => ({
          update: (view) => {
            // Re-rank every live marker when the footnote set changed —
            // their own NodeView.update never fires for another ref's
            // insertion or deletion.
            if (livePaints.size === 0) return;
            const signature = footnoteSignature(view.state.doc);
            if (signature === paintSignature) return;
            paintSignature = signature;
            for (const paint of livePaints) paint();
          },
        }),
      }),
    ];
  },
});

/**
 * Insert a footnote at the caret (Insert → Footnote / Ctrl+Alt+F): the ref
 * lands empty and the note editor opens immediately, the Word flow.
 */
export async function insertFootnoteAtCursor(editor: Editor): Promise<void> {
  if (editor.isDestroyed) return;
  const type = editor.state.schema.nodes[FOOTNOTE_REF];
  if (!type) return;
  editor
    .chain()
    .focus()
    .insertContent({ type: FOOTNOTE_REF, attrs: { text: "" } })
    .run();
  // An inline atom occupies exactly one position: the ref sits directly
  // before the caret's new position.
  const refAt = (): PMNode | null => {
    const pos = editor.state.selection.to - 1;
    const node = editor.state.doc.nodeAt(pos);
    return node && node.type.name === FOOTNOTE_REF ? node : null;
  };
  if (!refAt()) return;
  const pos = editor.state.selection.to - 1;
  const value = await promptModal({
    title: "Footnote",
    label: "Note text",
    value: "",
    okLabel: "Save",
    multiline: true,
  });
  if (editor.isDestroyed) return;
  const node = editor.state.doc.nodeAt(pos);
  if (!node || node.type.name !== FOOTNOTE_REF) return;
  if (value === null || value === "") {
    // An empty note is not worth keeping: drop the ref the user abandoned.
    editor.view.dispatch(editor.state.tr.delete(pos, pos + node.nodeSize));
    return;
  }
  editor.view.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, text: value }));
}
