import { Node } from "@tiptap/core";

/**
 * Soft line break (Shift+Enter) — the Word "line break" that wraps inside a
 * paragraph instead of starting a new one.
 *
 * Hand-rolled rather than pulled from `@tiptap/extension-hard-break` so the
 * dependency list stays exactly what `npm ci` already installs; the official
 * package is a one-line swap here if that ever changes.
 */
export const HardBreak = Node.create({
  name: "hardBreak",
  inline: true,
  group: "inline",
  selectable: false,

  renderHTML() {
    return ["br"];
  },

  parseHTML() {
    return [{ tag: "br" }];
  },

  addKeyboardShortcuts() {
    return {
      "Shift-Enter": () => this.editor.commands.insertContent({ type: "hardBreak" }),
    };
  },
});
