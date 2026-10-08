import { Bold } from "@tiptap/extension-bold";
import { Color } from "@tiptap/extension-color";
import { Heading } from "@tiptap/extension-heading";
import { Highlight } from "@tiptap/extension-highlight";
import { History } from "@tiptap/extension-history";
import { BulletList, ListItem, OrderedList } from "@tiptap/extension-list";
import { Image } from "@tiptap/extension-image";
import { Paragraph } from "@tiptap/extension-paragraph";
import { Placeholder } from "@tiptap/extension-placeholder";
import { Strike } from "@tiptap/extension-strike";
import { Table } from "@tiptap/extension-table";
import { TableCell } from "@tiptap/extension-table-cell";
import { TableHeader } from "@tiptap/extension-table-header";
import { TableRow } from "@tiptap/extension-table-row";
import { TextAlign } from "@tiptap/extension-text-align";
import { FontFamily, FontSize, TextStyle } from "@tiptap/extension-text-style";
import { Text } from "@tiptap/extension-text";
import { Underline } from "@tiptap/extension-underline";
import { Italic } from "@tiptap/extension-italic";
import { FreeWordDoc, Page, PageBreak } from "./schema";
import { Code } from "./code";
import { Cursors } from "./cursors";
import { HardBreak } from "./hard-break";
import { Link } from "./link";
import { Callout, CodeBlock, Divider, Quote, TaskItem, TaskList } from "./blocks";
import { SlashMenu } from "../features/slash-menu";
import { BubbleToolbar } from "../features/bubble-toolbar";
import { BlockHandle } from "../features/block-handle";

/**
 * Highlight also recognises run shading as emitted by docx-preview: Word's
 * `w:highlight`/`w:shd` render as inline `background-color` on the run's
 * `<span>`, while the stock extension only parses `<mark>`. Whites and
 * transparent values are ignored so default shading does not become a
 * highlight (callout/code-block fills are paragraph-level and already
 * consumed by the block-signature pass before parsing).
 */
const WHITEISH = new Set([
  "transparent",
  "inherit",
  "initial",
  "unset",
  "rgb(255,255,255)",
  "#fff",
  "#ffffff",
]);

const ShadingAwareHighlight = Highlight.extend({
  parseHTML() {
    return [
      { tag: "mark" },
      {
        tag: "span",
        getAttrs: (element: HTMLElement | string) => {
          if (typeof element === "string") return false;
          const bg = (element.style.backgroundColor || "").replace(/\s+/g, "").toLowerCase();
          if (!bg || WHITEISH.has(bg)) return false;
          return {};
        },
      },
    ];
  },
});

/**
 * In TipTap v3, `textStyle` is the single mark that carries inline style
 * attributes (color, font-family, font-size) as global attributes; the
 * extensions below attach those attributes to it.
 */
export function createExtensions() {
  return [
    FreeWordDoc,
    Page,
    PageBreak,
    Paragraph.extend({
      // The block handle starts a NodeSelection drag; the ProseMirror spec
      // only lets a node be dragged when it declares itself draggable.
      draggable: true,
      addAttributes() {
        return {
          lineSpacing: {
            default: 1,
            parseHTML: (element: HTMLElement) => {
              const lh = element.style.getPropertyValue("line-height");
              if (!lh) return 1;
              const n = parseFloat(lh);
              return Number.isFinite(n) && n > 0 ? n : 1;
            },
            renderHTML: (attrs: { lineSpacing?: number }) => {
              const ls = Number(attrs.lineSpacing) || 1;
              return ls !== 1 ? { style: `line-height:${ls}` } : {};
            },
          },
        };
      },
    }),
    Text,
    Heading.extend({
      draggable: true,
      addOptions() {
        return { levels: [1, 2, 3] as (1 | 2 | 3)[], HTMLAttributes: {} };
      },
    }),
    Bold,
    Italic,
    Underline,
    Strike,
    // Inline primitives that used to be missing entirely: Shift+Enter produced
    // no soft break, pasted URLs stayed plain text, and there was no monospace
    // run. See the individual files for why they are hand-rolled.
    HardBreak,
    Link,
    Code,
    TextAlign.configure({ types: ["paragraph", "heading"] }),
    TextStyle,
    Color,
    FontFamily,
    FontSize,
    ShadingAwareHighlight.configure({ multicolor: true }),
    BulletList,
    OrderedList,
    ListItem,
    // Notion-style blocks (see blocks.ts): to-do lists, callouts, quote,
    // divider and fenced code — all in the `block` group, so the page node
    // accepts them without a schema change.
    TaskList,
    TaskItem,
    Callout,
    Quote,
    Divider,
    CodeBlock,
    // Tables (imported from .docx via table/td/th markup; inserted via toolbar).
    // Resizable: the drag handle comes from prosemirror-tables' columnResizing,
    // which @tiptap/extension-table installs through the table node view; widths
    // land in the cells' `colwidth` attr and are exported to the .docx.
    Table.configure({ resizable: true, handleWidth: 6, lastColumnResizable: true }),
    TableRow,
    TableHeader,
    TableCell,
    Cursors,
    // Images: allow data: URLs so images survive .docx round-trips (both
    // docx-preview output and user-inserted images are base64-embedded).
    Image.configure({ allowBase64: true, inline: false }),
    History,
    Placeholder.configure({
      placeholder: "Write, or press '/' for commands…",
    }),
    // Notion-style chrome: slash menu, floating selection toolbar, block
    // handle (all ProseMirror-plugin owned, so they die with the editor).
    SlashMenu,
    BubbleToolbar,
    BlockHandle,
  ];
}
