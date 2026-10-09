import { Node } from "@tiptap/core";

/**
 * Height (CSS px) the pagination engine reserves inside the page's top /
 * bottom margin for the header / footer band while that band has text. The
 * bands themselves are absolutely positioned widgets — the reservation is
 * what stops page content from ever typing *under* them.
 */
export const FW_HEADER_RESERVE_PX = 28;
export const FW_FOOTER_RESERVE_PX = 28;

/**
 * Custom top-level node. ProseMirror requires a unique top-level node type
 * named "doc"; page sizes/margins live in its attrs and drive the CSS
 * variables applied to the editor element. The `header`/`footer` pair holds
 * the header/footer text (`{{page}}` / `{{pages}}` tokens supported),
 * `trackChanges` switches the review capture on, and `comments` is the
 * comment store the `comment` mark references — all of it travels inside
 * the document JSON, so page autosave and the .docx round-trip cover it.
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
      header: { default: null as string | null },
      headerAlign: { default: "left" as "left" | "center" | "right" },
      footer: { default: null as string | null },
      footerAlign: { default: "left" as "left" | "center" | "right" },
      trackChanges: { default: false },
      // A function default: a literal `{}` would hand the SAME mutable
      // object to every doc created without the attr.
      comments: { default: () => ({}) as Record<string, unknown> },
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
