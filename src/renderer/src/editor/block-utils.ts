import type { Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { TextSelection } from "@tiptap/pm/state";
import { BLOCK_NODE_NAMES, paragraphsFromText } from "./blocks";

/**
 * One catalogue of insertable / turn-into block types, shared by the slash
 * menu, the block-handle "Turn into" submenu and the toolbar's style select —
 * so all three offer exactly the same blocks with exactly the same behaviour.
 */

export interface BlockDef {
  /** Stable id used by menus. */
  id: string;
  label: string;
  /** Emoji shown in the slash menu (kept in the Unicode range CSP-safe). */
  icon: string;
  /** Extra words the filter matches against. */
  keywords: string;
  /** Hidden in Notion mode (page-level Word artefacts). */
  wordOnly?: boolean;
  /**
   * Insert `text` (possibly multi-line) as this block type at the top-level
   * block under the caret, replacing it when it is an empty paragraph and
   * otherwise appending after it. Returns false when the schema rejects it.
   */
  apply: (editor: Editor, text: string) => boolean;
}

/** The top-level block (direct child of a `page`) under the caret. */
export function topBlock(editor: Editor): { node: PMNode; before: number; after: number } | null {
  const { $from } = editor.state.selection;
  if ($from.depth < 2) return null;
  if ($from.node(1).type.name !== "page") return null;
  const node = $from.node(2);
  return { node, before: $from.before(2), after: $from.after(2) };
}

function blockText(node: PMNode): string {
  return node.textBetween(0, node.content.size, "\n", "");
}

/**
 * Replace the current top-level block (or follow it, when it carries text)
 * with `make`'s node(s). Shared by every BlockDef.apply.
 */
function replaceTopBlock(
  editor: Editor,
  text: string,
  make: (editor: Editor, text: string) => PMNode[] | null,
): boolean {
  const top = topBlock(editor);
  if (!top) return false;
  const nodes = make(editor, text);
  if (!nodes || nodes.length === 0) return false;

  const isEmptyParagraph = top.node.type.name === "paragraph" && top.node.content.size === 0;
  const tr = editor.state.tr;
  // An all-atom insert (divider, image, page break) leaves nothing to type
  // into: always trail it with a fresh paragraph so the caret has a home.
  const allAtoms = nodes.length > 0 && nodes.every((n) => n.isAtom);
  const withTrailer = allAtoms ? [...nodes, editor.state.schema.nodes["paragraph"].create()] : nodes;
  if (isEmptyParagraph || text) {
    // One atomic replace: a delete followed by an insert lets ProseMirror's
    // createAndFill repopulate the emptied page with a fresh paragraph that
    // the insert then lands *in front of* (the ghost block the two-step had).
    tr.replaceWith(top.before, top.after, withTrailer);
  } else {
    tr.insert(top.after, withTrailer);
  }
  try {
    const $start = tr.doc.resolve(isEmptyParagraph || text ? top.before + 1 : top.after + 1);
    tr.setSelection(TextSelection.near($start));
  } catch {
    /* keep the selection the replace left behind */
  }
  editor.view.dispatch(tr);
  return true;
}

function simpleBlock(typeName: string, attrs?: Record<string, unknown>) {
  return (editor: Editor, text: string): boolean =>
    replaceTopBlock(editor, text, (ed, t) => {
      const type = ed.state.schema.nodes[typeName];
      if (!type) return null;
      return [type.create(attrs ?? null, paragraphsFromText(ed.state.schema, t))];
    });
}

export const BLOCK_DEFS: BlockDef[] = [
  {
    id: "paragraph",
    label: "Text",
    icon: "¶",
    keywords: "paragraph body text plain",
    apply: simpleBlock("paragraph"),
  },
  {
    id: "heading1",
    label: "Heading 1",
    icon: "H1",
    keywords: "title heading big section",
    apply: (editor, text) =>
      replaceTopBlock(editor, text, (ed, t) => {
        const type = ed.state.schema.nodes["heading"];
        if (!type) return null;
        return [type.create({ level: 1 }, paragraphsFromText(ed.state.schema, t).slice(0, 1))];
      }),
  },
  {
    id: "heading2",
    label: "Heading 2",
    icon: "H2",
    keywords: "subtitle heading section",
    apply: (editor, text) =>
      replaceTopBlock(editor, text, (ed, t) => {
        const type = ed.state.schema.nodes["heading"];
        if (!type) return null;
        return [type.create({ level: 2 }, paragraphsFromText(ed.state.schema, t).slice(0, 1))];
      }),
  },
  {
    id: "heading3",
    label: "Heading 3",
    icon: "H3",
    keywords: "subtitle heading small section",
    apply: (editor, text) =>
      replaceTopBlock(editor, text, (ed, t) => {
        const type = ed.state.schema.nodes["heading"];
        if (!type) return null;
        return [type.create({ level: 3 }, paragraphsFromText(ed.state.schema, t).slice(0, 1))];
      }),
  },
  {
    id: "bulletList",
    label: "Bulleted list",
    icon: "•",
    keywords: "bullet unordered list ul",
    apply: (editor, text) =>
      replaceTopBlock(editor, text, (ed, t) => {
        const list = ed.state.schema.nodes["bulletList"];
        const item = ed.state.schema.nodes["listItem"];
        if (!list || !item) return null;
        const lines = t ? t.split("\n") : [""];
        return [
          list.create(
            null,
            lines.map((l) => item.create(null, paragraphsFromText(ed.schema, l))),
          ),
        ];
      }),
  },
  {
    id: "orderedList",
    label: "Numbered list",
    icon: "1.",
    keywords: "number ordered list ol",
    apply: (editor, text) =>
      replaceTopBlock(editor, text, (ed, t) => {
        const list = ed.state.schema.nodes["orderedList"];
        const item = ed.state.schema.nodes["listItem"];
        if (!list || !item) return null;
        const lines = t ? t.split("\n") : [""];
        return [
          list.create(
            null,
            lines.map((l) => item.create(null, paragraphsFromText(ed.schema, l))),
          ),
        ];
      }),
  },
  {
    id: "taskList",
    label: "To-do list",
    icon: "☑",
    keywords: "todo task checklist checkbox",
    apply: (editor, text) =>
      replaceTopBlock(editor, text, (ed, t) => {
        const list = ed.state.schema.nodes[BLOCK_NODE_NAMES.taskList];
        const item = ed.state.schema.nodes[BLOCK_NODE_NAMES.taskItem];
        if (!list || !item) return null;
        const lines = t ? t.split("\n") : [""];
        return [
          list.create(
            null,
            lines.map((l) => item.create({ checked: false }, paragraphsFromText(ed.schema, l))),
          ),
        ];
      }),
  },
  {
    id: "quote",
    label: "Quote",
    icon: "❝",
    keywords: "blockquote citation",
    apply: simpleBlock(BLOCK_NODE_NAMES.quote),
  },
  {
    id: "callout",
    label: "Callout",
    icon: "💡",
    keywords: "info note warning highlight box",
    apply: simpleBlock(BLOCK_NODE_NAMES.callout),
  },
  {
    id: "codeBlock",
    label: "Code",
    icon: "</>",
    keywords: "code block monospace pre snippet",
    apply: (editor, text) =>
      replaceTopBlock(editor, text, (ed, t) => {
        const type = ed.state.schema.nodes[BLOCK_NODE_NAMES.codeBlock];
        if (!type) return null;
        return [type.create(null, t ? ed.state.schema.text(t) : null)];
      }),
  },
  {
    id: "divider",
    label: "Divider",
    icon: "—",
    keywords: "separator hr line rule",
    apply: (editor) =>
      replaceTopBlock(editor, "", (ed) => {
        const type = ed.state.schema.nodes[BLOCK_NODE_NAMES.divider];
        return type ? [type.create()] : null;
      }),
  },
  {
    id: "image",
    label: "Image",
    icon: "🖼",
    keywords: "picture photo image media upload",
    // Opens the native picker instead of inserting a node; slash-menu.ts
    // special-cases this id, so a plain apply can never do anything.
    apply: () => false,
  },
  {
    id: "table",
    label: "Table",
    icon: "▦",
    keywords: "grid table 3x3",
    apply: (editor) => editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run(),
  },
  {
    id: "pageBreak",
    label: "Page break",
    icon: "⤓",
    keywords: "page break new page",
    wordOnly: true,
    apply: (editor) =>
      editor.chain().focus().insertContentAt(editor.state.selection, { type: "pageBreak" }).run(),
  },
];

/** Case-insensitive token filter over label + keywords. */
export function filterBlockDefs(defs: BlockDef[], query: string): BlockDef[] {
  const q = query.trim().toLowerCase();
  if (!q) return defs;
  const tokens = q.split(/\s+/);
  return defs.filter((def) => {
    const hay = `${def.label} ${def.keywords}`.toLowerCase();
    return tokens.every((t) => hay.includes(t));
  });
}

/** The block types offered for "Turn into" (excludes insert-only specials). */
export function turnIntoDefs(): BlockDef[] {
  return BLOCK_DEFS.filter(
    (d) => !d.wordOnly && d.id !== "table" && d.id !== "divider" && d.id !== "image",
  ).map((def) => ({
    ...def,
    // Turn-into always converts the existing text, never appends.
    apply: (editor: Editor, text: string) => {
      const top = topBlock(editor);
      if (!top) return false;
      const t = text || blockText(top.node);
      return def.apply(editor, t);
    },
  }));
}
