import type { Editor } from "@tiptap/core";
import { isSafeHref } from "../editor/link";
import { promptModal } from "./modal";

/**
 * Ask for a URL and apply (or clear) the link mark on the selection.
 * Lives here — not inside toolbar.ts — because both the ribbon toolbar and
 * the Notion-style bubble menu offer the link command and must validate and
 * prompt identically.
 */
export async function promptLink(editor: Editor): Promise<void> {
  const current = (editor.getAttributes("link").href as string | null) ?? "";
  const entered = await promptModal({
    title: "Insert link",
    label: "URL",
    value: current,
    placeholder: "https://example.com",
    okLabel: "Apply",
    // Rejected up front rather than dropped: a silent `setLink` failure would
    // look exactly like a button that does nothing.
    validate: (v) => (v && !isSafeHref(v) ? "Use an http(s), mailto: or tel: address." : null),
  });
  if (entered === null) return; // dismissed
  if (!entered) {
    editor.chain().focus().unsetLink().run();
    return;
  }
  editor.chain().focus().setLink({ href: entered }).run();
}
