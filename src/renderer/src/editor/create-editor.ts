import { Editor } from "@tiptap/core";
import type { JSONContent } from "@tiptap/core";
import { PAGE_SIZES } from "../../../shared/types";
import type { PageSize } from "../../../shared/types";
import { createExtensions } from "./extensions";

export const INITIAL_DOC: JSONContent = {
  type: "doc",
  attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
  content: [{ type: "page", content: [{ type: "paragraph" }] }],
};

/**
 * Push the current doc attrs (page size / margins) onto the editor element
 * as CSS variables so the page chrome always reflects document settings.
 */
export function syncDocAttrStyles(editor: Editor): void {
  const attrs = editor.state.doc.attrs as { size: string; mt: number; mr: number; mb: number; ml: number };
  const el = editor.view.dom as HTMLElement;
  const ps = PAGE_SIZES[attrs.size as PageSize] ?? PAGE_SIZES.a4;
  el.style.setProperty("--page-w", `${ps.w}mm`);
  el.style.setProperty("--page-h", `${ps.h}mm`);
  el.style.setProperty("--m-t", `${attrs.mt}mm`);
  el.style.setProperty("--m-r", `${attrs.mr}mm`);
  el.style.setProperty("--m-b", `${attrs.mb}mm`);
  el.style.setProperty("--m-l", `${attrs.ml}mm`);
}

export function createEditor(onUpdate: () => void): Editor {
  const element = document.querySelector<HTMLElement>("#editor");
  if (!element) throw new Error("#editor element not found");

  const editor = new Editor({
    element,
    extensions: createExtensions(),
    content: INITIAL_DOC,
    onUpdate: () => onUpdate(),
  });

  element.classList.add("fw-doc");
  syncDocAttrStyles(editor);
  return editor;
}
