import { Mark, mergeAttributes } from "@tiptap/core";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    fwCode: {
      /** Toggle the inline code (monospace) mark. */
      toggleCode: () => ReturnType;
    };
  }
}

/**
 * Inline code / monospace span — the Word "Consolas over the text" case rather
 * than a fenced block, because the document model here has no block-level
 * preformatted region and the page-break engine measures blocks, not widgets.
 *
 * `excludes` is deliberately NOT set, so code coexists with bold / italic /
 * colour the way Word's font dialog allows.
 */
export const Code = Mark.create({
  name: "code",
  // A mark outside the inline group cannot be attached to text at all: PM
  // rejects it, `toggleMark` silently returns false, and the button looks dead.
  group: "inline",
  code: true,
  inclusive: false,

  renderHTML({ HTMLAttributes }) {
    return ["code", mergeAttributes(HTMLAttributes), 0];
  },

  parseHTML() {
    return [
      { tag: "code" },
      { tag: "kbd" },
      { tag: "samp" },
      // Legacy HTML from exported web pages.
      { tag: "tt" },
    ];
  },

  addCommands() {
    return {
      toggleCode:
        () =>
        ({ chain }) =>
          chain().toggleMark("code", {}).run(),
    };
  },
});
