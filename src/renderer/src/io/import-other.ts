import { DOMParser as PMDOMParser } from "@tiptap/pm/model";
import type { Schema } from "@tiptap/pm/model";
import { DEFAULT_DOC_ATTRS, flatBlockJson, normalizeImages, sanitizeAnchors } from "./import-docx";

function wrapPage(pageContent: unknown[]): Record<string, unknown> {
  return {
    type: "doc",
    attrs: { ...DEFAULT_DOC_ATTRS },
    content: [{ type: "page", content: pageContent }],
  };
}

/** Plain text → one paragraph per line (split on \r\n / \n / \r). */
export function textToDoc(text: string, _schema?: Schema): Record<string, unknown> {
  const lines = text.split(/\r\n|\n|\r/);
  const content = lines.map((line) =>
    line.length > 0 ? { type: "paragraph", content: [{ type: "text", text: line }] } : { type: "paragraph" },
  );
  return wrapPage(content);
}

/**
 * HTML → PM doc via the same schema parser as the .docx path.
 * Tables and data-URL images are kept; scripts/styles/external images are
 * dropped and h4–h6 are demoted to h3 (schema has levels 1–3).
 */
export function htmlToDoc(html: string, schema: Schema): Record<string, unknown> {
  const container = document.createElement("div");
  container.innerHTML = html;
  container.querySelectorAll("style, script, link").forEach((el) => el.remove());
  normalizeImages(container);
  sanitizeAnchors(container);
  container.querySelectorAll("h4, h5, h6").forEach((el) => {
    const h3 = document.createElement("h3");
    while (el.firstChild) h3.appendChild(el.firstChild);
    el.replaceWith(h3);
  });
  const parsed = PMDOMParser.fromSchema(schema).parse(container);
  // The DOM parser wraps top-level content in a `page` node (topNode
  // content is `page+`); flatten so we get a single page of blocks.
  return wrapPage(flatBlockJson(parsed));
}
