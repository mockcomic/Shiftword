import { Mark, mergeAttributes } from "@tiptap/core";
import { safeExternalUrl } from "../../../shared/urls";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    fwLink: {
      /** Apply a hyperlink to the selection / stored marks. Rejects unsafe URLs. */
      setLink: (attrs: { href: string }) => ReturnType;
      /** Apply or remove the hyperlink mark. Rejects unsafe URLs. */
      toggleLink: (attrs: { href: string }) => ReturnType;
      /** Remove the hyperlink mark from the selection. */
      unsetLink: () => ReturnType;
    };
  }
}

/**
 * What may be stored on a link: an http(s)/mailto/tel URL by the shared policy,
 * or an in-document `#anchor`. Anything else — `javascript:`, `data:`,
 * `file:` — is refused so a hostile import cannot park an executable URL in the
 * document.
 */
export function isSafeHref(href: unknown): href is string {
  if (typeof href !== "string") return false;
  const trimmed = href.trim();
  if (!trimmed || trimmed === "#") return false;
  return trimmed.startsWith("#") || safeExternalUrl(trimmed) !== null;
}

/**
 * Normalise a user- or HTML-supplied URL: the trimmed URL when the scheme is
 * allowed, otherwise null so callers can drop the link instead of storing a
 * `javascript:` / `data:` payload that a later click would execute.
 */
export function normalizeHref(href: unknown): string | null {
  if (!isSafeHref(href)) return null;
  return href.trim();
}

export interface LinkAttributes {
  href: string | null;
}

/**
 * Hyperlink mark.
 *
 * Deliberately inert on click: the official `@tiptap/extension-link` opens
 * `window.open()` on click, which in Electron means a second renderer unless
 * something intercepts it. Clicks are instead routed through
 * `features/links.ts` → `fw.openExternal()` → a validated main-process
 * `shell.openExternal`, so the only navigation path is the checked one.
 */
export const Link = Mark.create({
  name: "link",
  // A mark has to be in the inline group or ProseMirror refuses to attach it
  // to text — the command then just returns false and the button looks dead.
  group: "inline",

  addAttributes() {
    return {
      href: {
        default: null,
        parseHTML: (element: HTMLElement) => normalizeHref(element.getAttribute("href")),
        renderHTML: (attrs: LinkAttributes) =>
          attrs.href ? { href: attrs.href, target: "_blank", rel: "noopener noreferrer" } : {},
      },
    };
  },

  parseHTML() {
    return [{ tag: "a[href]" }];
  },

  renderHTML({ HTMLAttributes }) {
    // An href that failed validation renders as plain text: no dead anchor in
    // the document, and nothing for a click handler to act on.
    if (!isSafeHref(HTMLAttributes.href)) return ["span", 0];
    return ["a", mergeAttributes(HTMLAttributes), 0];
  },

  addCommands() {
    return {
      setLink:
        (attrs: { href: string }) =>
        ({ chain }) => {
          const href = normalizeHref(attrs?.href);
          if (!href) return false;
          return chain().extendMarkRange("link").setMark("link", { href }).run();
        },
      toggleLink:
        (attrs: { href: string }) =>
        ({ chain }) => {
          const href = normalizeHref(attrs?.href);
          if (!href) return false;
          return chain().extendMarkRange("link").toggleMark("link", { href }).run();
        },
      unsetLink:
        () =>
        ({ chain }) =>
          chain().extendMarkRange("link").unsetMark("link").run(),
    };
  },
});
