import type { Editor } from "@tiptap/core";
import { getHTMLFromFragment } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import {
  FONT_STACK_BODY,
  FONT_STACK_CODE,
  HEADING_SIZE_PT,
  HYPERLINK_COLOR,
  LINE_HEIGHT,
  PARA_AFTER_PT,
  SIZE_BODY_PT,
} from "../../../shared/typography";

/** Deep copy of the current document as plain PM JSON (IPC-safe). */
export function docJson(editor: Editor): Record<string, unknown> {
  return JSON.parse(JSON.stringify(editor.getJSON()));
}

/**
 * Plain text with one line per paragraph block.
 *
 * Not `doc.textBetween` anymore: the "final" view must drop text the user
 * deleted *without accepting the revision* (Word exports the same view),
 * and the footnote refs stay silent — a txt file has no page foot to put
 * the notes in, and the bracketed number would only be noise.
 */
function inlineText(node: PMNode): string {
  let out = "";
  node.forEach((child) => {
    if (child.isText) {
      if (!child.marks.some((m) => m.type.name === "trackedDelete")) out += String(child.text ?? "");
      return;
    }
    if (child.type.name === "hardBreak") {
      out += "\n";
      return;
    }
    if (child.type.name === "footnoteRef") return; // the number is layout, not text
    out += " "; // other atoms keep the old leafText behaviour
  });
  return out;
}

function collectLines(node: PMNode, out: string[]): void {
  if (node.isTextblock) {
    out.push(inlineText(node));
    return;
  }
  if (node.type.name === "table") {
    node.forEach((row) => {
      const cells: string[] = [];
      row.forEach((cell) => {
        const inner: string[] = [];
        collectLines(cell, inner);
        cells.push(inner.join(" "));
      });
      out.push(cells.join("\t"));
    });
    return;
  }
  node.forEach((child) => collectLines(child, out));
}

export function toText(editor: Editor): string {
  const out: string[] = [];
  editor.state.doc.forEach((page) => collectLines(page, out));
  return out.join("\n");
}

/**
 * Standalone HTML export.
 *
 * Deliberately NOT `editor.getHTML()`: our top-level nodes render the *view*
 * chrome (`div.fw-doc`, one `div.fw-page` per page with the page CSS baked
 * in), and pagination is an accident of the current layout — it would leak
 * empty page boxes and arbitrary page breaks into a file someone opens in a
 * browser. Blocks never split across pages, so concatenating each page's
 * content fragment preserves both order and content.
 */
const HTML_STYLES = [
  `body{margin:2rem auto;max-width:52rem;font-family:${FONT_STACK_BODY};font-size:${SIZE_BODY_PT}pt;line-height:${LINE_HEIGHT};color:#000}`,
  `h1{font-size:${HEADING_SIZE_PT[1]}pt}h2{font-size:${HEADING_SIZE_PT[2]}pt}h3{font-size:${HEADING_SIZE_PT[3]}pt}`,
  `h1,h2,h3,p{margin:0 0 ${PARA_AFTER_PT}pt}`,
  "table{border-collapse:collapse;width:100%}",
  "td,th{border:1px solid #767171;padding:3px 6px;vertical-align:top}",
  "th{background:#f2f2f2;font-weight:bold}",
  "img{max-width:100%;height:auto}",
  `code{font-family:${FONT_STACK_CODE};background:#f2f2f2;padding:0 2px}`,
  `a{color:#${HYPERLINK_COLOR.toLowerCase()};text-decoration:underline}`,
  "blockquote{margin:0 0 6pt;padding-left:12pt;border-left:3px solid #d0d0d0;color:#333;font-style:italic}",
  "ul.fw-task-list{list-style:none;padding-left:6pt}",
  "ul.fw-task-list li{display:flex;gap:6pt;align-items:flex-start}",
  "ul.fw-task-list input[type=checkbox]{margin-top:2pt}",
  ".fw-callout{display:flex;gap:8pt;background:#f1f1f1;border-radius:6px;padding:8pt 10pt;margin:0 0 6pt}",
  ".fw-callout-body p{margin:0 0 4pt}",
  "hr.fw-divider{border:none;border-top:1px solid #e0e0e0;margin:10pt 0}",
  "pre.fw-codeblock{font-family:Consolas,'Courier New',monospace;font-size:0.95em;background:#f2f2f2;border-radius:4px;padding:6pt 8pt;white-space:pre-wrap;margin:0 0 6pt}",
  // Review marks keep their meaning in a plain HTML export: insertions stay
  // underlined, rejected-pending deletions stay struck through (shown, like
  // Word's "All Markup" view — the HTML export never hides content).
  ".fw-ins{color:#1a6b1a;text-decoration:underline}",
  ".fw-del{color:#b00020;text-decoration:line-through}",
  ".fw-comment{background:rgba(255,235,59,.35)}",
  ".fw-fn-ref{font-size:.75em;vertical-align:super;color:#444}",
].join("\n");

export function toHtml(editor: Editor): string {
  const { doc, schema } = editor.state;
  const body: string[] = [];
  doc.forEach((page) => {
    if (page.content.size === 0) return;
    body.push(getHTMLFromFragment(page.content, schema));
  });

  return [
    "<!doctype html>",
    '<html lang="en">',
    "<head>",
    '<meta charset="utf-8">',
    "<title>Shiftword Document</title>",
    "<style>",
    HTML_STYLES,
    "</style>",
    "</head>",
    "<body>",
    body.join("\n"),
    "</body>",
    "</html>",
    "",
  ].join("\n");
}
