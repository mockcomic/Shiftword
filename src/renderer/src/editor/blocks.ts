import { mergeAttributes, Node } from "@tiptap/core";
import { InputRule } from "@tiptap/core";
import type { Node as PMNode, Schema } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";

/**
 * Notion-style block primitives, hand-rolled in the repo's existing pattern
 * (see code.ts / link.ts / hard-break.ts): no new npm dependencies, and every
 * node renders plain HTML that ProseMirror's DOMParser reads back — which is
 * also what makes the .docx round-trip recognisable: io/import-docx.ts
 * rebuilds these exact signatures from an exported file.
 *
 * All of them join the `block` group, so the custom `page` node (`block+`)
 * accepts them with no schema change.
 */

const TASK_LIST = "taskList";
const TASK_ITEM = "taskItem";
const CALLOUT = "callout";
const QUOTE = "fwQuote";
const DIVIDER = "divider";
const CODE_BLOCK = "fwCodeBlock";

/** Icons the user cycles through by clicking a callout's icon. */
export const CALLOUT_ICONS = ["💡", "ℹ️", "⚠️", "✅", "❗", "📌"] as const;

/** Node names for the catalog, slash menu, and turn-into helpers. */
export const BLOCK_NODE_NAMES = {
  taskList: TASK_LIST,
  taskItem: TASK_ITEM,
  callout: CALLOUT,
  quote: QUOTE,
  divider: DIVIDER,
  codeBlock: CODE_BLOCK,
} as const;

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

function paragraphWith(schema: Schema, text: string): PMNode {
  const para = schema.nodes["paragraph"];
  if (!text) return para.create();
  return para.create(null, schema.text(text));
}

/** One paragraph per line of `text` (empty text → one empty paragraph). */
export function paragraphsFromText(schema: Schema, text: string): PMNode[] {
  const lines = text ? text.split("\n") : [""];
  return lines.map((line) => paragraphWith(schema, line));
}

/** Absolute position of the taskItem ancestor of `pos`, or null. */
function taskItemPos(doc: PMNode, pos: number): number | null {
  const $pos = doc.resolve(pos);
  for (let d = $pos.depth; d > 0; d--) {
    if ($pos.node(d).type.name === TASK_ITEM) return $pos.before(d);
  }
  return null;
}

/** Position of the rendered node `dom` in the document (scans blocks). */
function posOfRenderedNode(view: EditorView, typeName: string, dom: HTMLElement): number | null {
  let found: number | null = null;
  view.state.doc.descendants((node, pos) => {
    if (found !== null) return false;
    if (node.type.name !== typeName) return true;
    try {
      if (view.nodeDOM(pos) === dom) found = pos;
    } catch {
      /* nodeDOM can throw for positions the view has not rendered */
    }
    return false;
  });
  return found;
}

/** Locate a block-type node by walking up from the selection. */
function ancestorPos(doc: PMNode, pos: number, typeName: string): number | null {
  const $pos = doc.resolve(pos);
  for (let d = $pos.depth; d > 0; d--) {
    if ($pos.node(d).type.name === typeName) return $pos.before(d);
  }
  return null;
}

// ---------------------------------------------------------------------------
// TaskList / TaskItem
// ---------------------------------------------------------------------------

export const TaskItem = Node.create({
  name: TASK_ITEM,
  // Deliberately NOT in the block group: only TaskList may contain items.
  content: "paragraph",
  defining: true,

  addAttributes() {
    return {
      checked: {
        default: false,
        parseHTML: (element: HTMLElement) => element.getAttribute("data-checked") === "true",
        renderHTML: (attrs: { checked?: boolean }) => ({
          "data-checked": attrs.checked ? "true" : "false",
        }),
      },
    };
  },

  renderHTML({ node, HTMLAttributes }) {
    const checked = Boolean(node.attrs.checked);
    return [
      "li",
      mergeAttributes(HTMLAttributes, { "data-type": TASK_ITEM, class: "fw-task-item" }),
      [
        "label",
        { class: "fw-task-label" },
        [
          "input",
          {
            type: "checkbox",
            class: "fw-task-checkbox",
            ...(checked ? { checked: "checked" } : {}),
          },
        ],
        ["span", { class: "fw-task-content" }, 0],
      ],
    ];
  },

  parseHTML() {
    return [{ tag: `li[data-type="${TASK_ITEM}"]`, priority: 60 }];
  },

  addKeyboardShortcuts() {
    const editor = this.editor;

    const splitOrManage = (): boolean => {
      const { state } = editor;
      const sel = state.selection;
      if (!sel.empty) return false;
      const $from = sel.$from;
      const itemPos = taskItemPos(state.doc, $from.pos);
      if (itemPos === null) return false;
      const item = state.doc.nodeAt(itemPos);
      if (!item) return false;
      // Only manage Enter when the caret sits in the item's own paragraph.
      if ($from.parent.type.name !== "paragraph" || $from.depth === 0) return false;

      const before = $from.parent.textBetween(0, $from.parentOffset);
      const after = $from.parent.textBetween($from.parentOffset, $from.parent.content.size);

      // Enter on an empty unchecked task exits the list (Notion semantics):
      // the item disappears and a plain paragraph takes it, with the caret
      // parked inside it. The paragraph must be created explicitly — a bare
      // delete plus a `TextSelection.near` drifts backward into the previous
      // item when there is nothing ahead to land in, which left the user
      // oscillating between item splits forever, unable to start a new line.
      if (before === "" && after === "") {
        const tr = state.tr;
        if (item.attrs.checked) {
          tr.setNodeMarkup(itemPos, undefined, { ...item.attrs, checked: false });
          editor.view.dispatch(tr);
          return true;
        }
        const listPos = ancestorPos(state.doc, itemPos, TASK_LIST);
        if (listPos === null) return false;
        const list = state.doc.nodeAt(listPos);
        if (!list) return false;

        const listEnd = listPos + list.nodeSize;
        if (list.childCount <= 1) {
          // The whole list collapses into one paragraph at the same spot.
          tr.replaceWith(listPos, listEnd, paragraphWith(state.schema, ""));
          tr.setSelection(TextSelection.create(tr.doc, listPos + 1));
        } else if (itemPos + item.nodeSize === listEnd - 1) {
          // Trailing item: exit below the list — delete the item, then insert
          // a paragraph right after the (now shorter) list and place the
          // caret inside it, wherever the page continues afterwards.
          tr.delete(itemPos, itemPos + item.nodeSize);
          const paraPos = tr.mapping.map(listEnd);
          tr.insert(paraPos, paragraphWith(state.schema, ""));
          tr.setSelection(TextSelection.create(tr.doc, paraPos + 1));
        } else {
          // Middle item: just drop it; the caret moves forward onto the next
          // item (bias 1 keeps it from drifting into the previous one).
          tr.delete(itemPos, itemPos + item.nodeSize);
          const at = tr.mapping.map(itemPos);
          tr.setSelection(TextSelection.near(tr.doc.resolve(at), 1));
        }
        editor.view.dispatch(tr);
        return true;
      }

      // Split the text across the current and a fresh (unchecked) item.
      const itemSchema = editor.state.schema;
      const prevItem = item.type.create({ ...item.attrs }, paragraphWith(itemSchema, before));
      const nextItem = item.type.create({ ...item.attrs, checked: false }, paragraphWith(itemSchema, after));
      const tr = state.tr.replaceWith(itemPos, itemPos + item.nodeSize, [prevItem, nextItem]);
      const nextStart = itemPos + prevItem.nodeSize + 1;
      tr.setSelection(TextSelection.near(tr.doc.resolve(nextStart)));
      editor.view.dispatch(tr);
      return true;
    };

    return {
      Enter: splitOrManage,
    };
  },
});

export const TaskList = Node.create({
  name: TASK_LIST,
  group: "block",
  content: `${TASK_ITEM}+`,

  renderHTML({ HTMLAttributes }) {
    return ["ul", mergeAttributes(HTMLAttributes, { "data-type": TASK_LIST, class: "fw-task-list" }), 0];
  },

  parseHTML() {
    // Priority above the bullet list's plain `ul` rule, which would otherwise
    // swallow every task list into a bulleted list on paste/import.
    return [{ tag: `ul[data-type="${TASK_LIST}"]`, priority: 60 }];
  },

  addInputRules() {
    return [
      new InputRule({
        find: /^\[( |x|X)\]\s$/,
        handler: ({ state, range, chain }) => {
          const schema = state.schema;
          const item = schema.nodes[TASK_ITEM].create({ checked: false }, paragraphWith(schema, ""));
          const list = schema.nodes[TASK_LIST].create(null, [item]);
          chain().deleteRange(range).insertContentAt(range.from, list).run();
        },
      }),
    ];
  },

  addProseMirrorPlugins() {
    // Checkbox clicks live here rather than in a NodeView: the checkbox is a
    // plain rendered leaf, and dispatching an attribute update keeps undo,
    // dirty-tracking and the .docx export honest.
    return [
      new Plugin({
        key: new PluginKey("fwTaskCheck"),
        props: {
          handleDOMEvents: {
            click: (view, event) => {
              const target = event.target as HTMLElement | null;
              const box = target?.closest?.("input.fw-task-checkbox") as HTMLInputElement | null;
              if (!box) return false;
              event.preventDefault();
              const li = box.closest("li") as HTMLElement | null;
              if (!li) return false;
              const pos = posOfRenderedNode(view, TASK_ITEM, li);
              if (pos === null) return false;
              const node = view.state.doc.nodeAt(pos);
              if (!node) return false;
              view.dispatch(
                view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, checked: !node.attrs.checked }),
              );
              return true;
            },
          },
        },
      }),
    ];
  },
});

// ---------------------------------------------------------------------------
// Callout
// ---------------------------------------------------------------------------

export const Callout = Node.create({
  name: CALLOUT,
  group: "block",
  content: "block+",

  addAttributes() {
    return {
      icon: {
        default: "💡",
        parseHTML: (element: HTMLElement) => element.getAttribute("data-icon") || "💡",
        renderHTML: (attrs: { icon?: string }) => ({ "data-icon": attrs.icon || "💡" }),
      },
    };
  },

  renderHTML({ HTMLAttributes }) {
    return [
      "div",
      mergeAttributes(HTMLAttributes, { class: "fw-callout", "data-type": CALLOUT }),
      ["span", { class: "fw-callout-icon", contenteditable: "false" }],
      ["div", { class: "fw-callout-body" }, 0],
    ];
  },

  parseHTML() {
    return [{ tag: `div[data-type="${CALLOUT}"]` }];
  },

  addInputRules() {
    return [
      new InputRule({
        find: /^>>\s$/,
        handler: ({ state, range, chain }) => {
          const node = state.schema.nodes[CALLOUT].create(null, paragraphWith(state.schema, ""));
          chain().deleteRange(range).insertContentAt(range.from, node).run();
        },
      }),
    ];
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("fwCalloutIcon"),
        props: {
          handleDOMEvents: {
            click: (view, event) => {
              const target = event.target as HTMLElement | null;
              const icon = target?.closest?.(".fw-callout-icon") as HTMLElement | null;
              if (!icon) return false;
              event.preventDefault();
              const box = icon.closest(`div[data-type="${CALLOUT}"]`) as HTMLElement | null;
              if (!box) return false;
              const pos = posOfRenderedNode(view, CALLOUT, box);
              if (pos === null) return false;
              const node = view.state.doc.nodeAt(pos);
              if (!node) return false;
              const current = String(node.attrs.icon || "💡");
              const idx = (CALLOUT_ICONS as readonly string[]).indexOf(current);
              const next = CALLOUT_ICONS[(idx + 1 + CALLOUT_ICONS.length) % CALLOUT_ICONS.length];
              view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, icon: next }));
              return true;
            },
          },
        },
      }),
    ];
  },
});

// ---------------------------------------------------------------------------
// Quote
// ---------------------------------------------------------------------------

export const Quote = Node.create({
  name: QUOTE,
  group: "block",
  content: "block+",
  defining: true,

  renderHTML({ HTMLAttributes }) {
    return ["blockquote", mergeAttributes(HTMLAttributes, { class: "fw-quote", "data-type": QUOTE }), 0];
  },

  parseHTML() {
    return [{ tag: "blockquote" }];
  },

  addInputRules() {
    return [
      new InputRule({
        find: /^\s?>\s$/,
        handler: ({ state, range, chain }) => {
          const node = state.schema.nodes[QUOTE].create(null, paragraphWith(state.schema, ""));
          chain().deleteRange(range).insertContentAt(range.from, node).run();
        },
      }),
    ];
  },
});

// ---------------------------------------------------------------------------
// Divider
// ---------------------------------------------------------------------------

export const Divider = Node.create({
  name: DIVIDER,
  group: "block",
  atom: true,
  draggable: false,

  renderHTML() {
    return ["hr", { class: "fw-divider", "data-type": DIVIDER }];
  },

  parseHTML() {
    return [{ tag: "hr" }];
  },

  addInputRules() {
    return [
      new InputRule({
        find: /^(?:---|___|\*\*\*)$/,
        handler: ({ state, range, chain }) => {
          const node = state.schema.nodes[DIVIDER].create();
          chain().deleteRange(range).insertContentAt(range.from, node).run();
        },
      }),
    ];
  },
});

// ---------------------------------------------------------------------------
// CodeBlock
// ---------------------------------------------------------------------------

export const CodeBlock = Node.create({
  name: CODE_BLOCK,
  group: "block",
  content: "text*",
  marks: "",
  code: true,
  defining: true,

  renderHTML({ HTMLAttributes }) {
    return ["pre", mergeAttributes(HTMLAttributes, { class: "fw-codeblock", "data-type": CODE_BLOCK }), 0];
  },

  parseHTML() {
    return [{ tag: "pre", preserveWhitespace: "full" }];
  },

  addInputRules() {
    return [
      new InputRule({
        find: /^```$/,
        handler: ({ state, range, chain }) => {
          const node = state.schema.nodes[CODE_BLOCK].create();
          chain().deleteRange(range).insertContentAt(range.from, node).run();
        },
      }),
    ];
  },

  addKeyboardShortcuts() {
    const editor = this.editor;
    return {
      // Literal newlines inside the block instead of trying to split a
      // `text*` block (which the content expression forbids anyway).
      Enter: () => {
        const { state } = editor;
        if (state.selection.$from.parent.type.name !== CODE_BLOCK) return false;
        return editor.commands.insertContent("\n");
      },
      Tab: () => {
        const { state } = editor;
        if (state.selection.$from.parent.type.name !== CODE_BLOCK) return false;
        return editor.commands.insertContent("  ");
      },
    };
  },
});
