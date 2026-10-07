import { Node } from "@tiptap/core";

/**
 * Custom top-level node. ProseMirror requires a unique top-level node type
 * named "doc"; page sizes/margins live in its attrs and drive the CSS
 * variables applied to the editor element.
 */
export const FreeWordDoc = Node.create({
  name: "doc",
  topNode: true,
  content: "page+",
  addAttributes() {
    return {
      size: { default: "a4" },
      mt: { default: 25.4 },
      mr: { default: 25.4 },
      mb: { default: 25.4 },
      ml: { default: 25.4 },
    };
  },
  renderHTML() {
    return ["div", { class: "fw-doc" }, 0];
  },
});

/**
 * A fixed-size page. `block+` allows any block node (paragraph, heading,
 * lists, page breaks). The `group` must stay set so ProseMirror does not
 * treat it as a second top-level candidate next to `doc`.
 */
export const Page = Node.create({
  name: "page",
  group: "block",
  content: "block+",
  renderHTML() {
    return ["div", { class: "fw-page" }, 0];
  },
});

/**
 * Explicit page break marker. Created only via the Insert → Page Break
 * command (or imported from .docx); it is never parsed from plain HTML.
 */
export const PageBreak = Node.create({
  name: "pageBreak",
  group: "block",
  atom: true,
  selectable: false,
  draggable: false,
  parseHTML: () => [],
  renderHTML() {
    return ["div", { class: "fw-pagebreak" }];
  },
});
