import { Extension } from "@tiptap/core";
import { dropCursor } from "@tiptap/pm/dropcursor";
import { gapCursor } from "@tiptap/pm/gapcursor";

/**
 * Drop and gap cursors, wrapped from the ProseMirror plugins `@tiptap/pm`
 * already re-exports (`@tiptap/pm/dropcursor`, `@tiptap/pm/gapcursor`).
 *
 * Without them there is no line where a dragged image or table can land, and
 * no caret position at the boundary of two tables — both of which make those
 * blocks feel unreachable rather than merely hard to place.
 *
 * Wrapped here rather than imported from `@tiptap/extensions` because that
 * package is only present transitively today; depending on it directly would
 * silently ride somebody else's dependency tree.
 */
export const Cursors = Extension.create({
  name: "fwCursors",

  addProseMirrorPlugins() {
    return [
      // `gapCursor()` takes no options in the bundled prosemirror-gapcursor.
      dropCursor({ color: "#0067c0", width: 2 }),
      gapCursor(),
    ];
  },
});
