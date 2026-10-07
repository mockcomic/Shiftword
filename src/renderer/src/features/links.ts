import type { Editor } from "@tiptap/core";

/**
 * Clicks on hyperlinks must not navigate the editor window or spawn a second
 * renderer. The closest `<a>` is intercepted and its URL handed to the main
 * process, which re-validates the scheme before calling `shell.openExternal`
 * — the renderer never gets to decide what opens.
 */
export function installLinkClicks(editor: Editor): void {
  const view = editor.view;
  view.dom.addEventListener("click", (event) => {
    const anchor = (event.target as HTMLElement | null)?.closest?.("a[href]") as HTMLAnchorElement | null;
    if (!anchor || !view.dom.contains(anchor)) return;

    // Word semantics: a plain click positions the caret inside the link so it
    // can be edited; Ctrl/Cmd+click is the explicit "open this" gesture, so an
    // accidental click never launches a browser.
    if (!(event.ctrlKey || event.metaKey)) return;

    const href = anchor.getAttribute("href");
    if (!href) return;
    event.preventDefault();
    event.stopPropagation();
    void window.fw?.openExternal(href);
  });
}
