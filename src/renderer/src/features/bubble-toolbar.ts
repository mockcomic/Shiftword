import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import { CellSelection } from "@tiptap/pm/tables";
import { promptLink } from "./link-prompt";
import { addCommentToSelection } from "../editor/comments";

/**
 * Floating selection toolbar — the Notion affordance: select text and a small
 * rounded bar appears above it with the mark commands, so the ribbon toolbar
 * stops being the only way to format anything (and can hide entirely in
 * Notion mode).
 *
 * An Extension + plugin so it lives and dies with the editor view; the popup
 * is `position: fixed` (outside the zoomed document) and its buttons use
 * mousedown+preventDefault so the editor selection survives the click.
 */

const HIGHLIGHT_YELLOW = "#FFFF00";

export const BubbleToolbar = Extension.create({
  name: "fwBubbleToolbar",

  addProseMirrorPlugins() {
    const editor = this.editor;

    let popup: HTMLElement | null = null;
    /**
     * Signature of the last rendered (selection range + mark states). The
     * plugin `update` fires for EVERY transaction (pagination, decorations,
     * history meta…), and the old rebuild paid a full innerHTML rewrite,
     * 16 listener re-attachments and two forced `coordsAtPos` layouts each
     * time. Same signature → same DOM, zero layout reads.
     */
    let lastSignature = "";

    const close = (): void => {
      if (popup) {
        popup.remove();
        popup = null;
        lastSignature = "";
      }
    };

    const selectionIsBubbleWorthy = (view: EditorView): boolean => {
      const { selection } = view.state;
      if (selection.empty) return false;
      if (selection instanceof CellSelection) return false;
      // Inside a table the ribbon's table group is the working UI; a bubble
      // there would fight the column-resize handles for the same square pixels.
      const $from = selection.$from;
      for (let d = $from.depth; d > 0; d--) {
        if ($from.node(d).type.name === "table") return false;
      }
      // Code blocks: the mark buttons are meaningless there (marks: "").
      if ($from.parent.type.spec.code) return false;
      return true;
    };

    interface BubbleCommand {
      id: string;
      label: string;
      title: string;
      active: () => boolean;
      run: () => void;
    }

    const commands = (): BubbleCommand[] => [
      {
        id: "bold",
        label: "B",
        title: "Bold",
        active: () => editor.isActive("bold"),
        run: () => editor.chain().toggleMark("bold").run(),
      },
      {
        id: "italic",
        label: "I",
        title: "Italic",
        active: () => editor.isActive("italic"),
        run: () => editor.chain().toggleMark("italic").run(),
      },
      {
        id: "underline",
        label: "U",
        title: "Underline",
        active: () => editor.isActive("underline"),
        run: () => editor.chain().toggleMark("underline").run(),
      },
      {
        id: "strike",
        label: "S",
        title: "Strikethrough",
        active: () => editor.isActive("strike"),
        run: () => editor.chain().toggleMark("strike").run(),
      },
      {
        id: "code",
        label: "</>",
        title: "Monospace",
        active: () => editor.isActive("code"),
        run: () => editor.chain().toggleCode().run(),
      },
      {
        id: "link",
        label: "↗",
        title: "Link",
        active: () => editor.isActive("link"),
        run: () => {
          void promptLink(editor);
        },
      },
      {
        id: "highlight",
        label: "▟",
        title: "Yellow highlight",
        active: () => editor.isActive("highlight", { color: HIGHLIGHT_YELLOW }),
        run: () => editor.chain().toggleHighlight({ color: HIGHLIGHT_YELLOW }).run(),
      },
      {
        id: "comment",
        label: "💬",
        title: "Add comment",
        active: () => false,
        run: () => {
          // The comment prompt captures the selection itself before opening.
          void addCommentToSelection(editor);
        },
      },
    ];

    const render = (view: EditorView): void => {
      if (!selectionIsBubbleWorthy(view)) {
        close();
        return;
      }
      const defs = commands();
      const { selection } = view.state;
      const signature = `${selection.from}:${selection.to}:${defs.map((d) => (d.active() ? "1" : "0")).join("")}`;
      if (popup && signature === lastSignature) return; // nothing moved, nothing toggled
      lastSignature = signature;
      if (!popup) {
        popup = document.createElement("div");
        popup.className = "bubble-menu";
        document.body.appendChild(popup);
      }
      popup.innerHTML = defs
        .map(
          (def) =>
            `<button type="button" class="bubble-btn${def.active() ? " active" : ""}" data-id="${def.id}" title="${def.title}">${def.label}</button>`,
        )
        .join("");
      popup.querySelectorAll<HTMLButtonElement>(".bubble-btn").forEach((btn) => {
        btn.addEventListener("mousedown", (e) => {
          e.preventDefault(); // keep the editor selection
          e.stopPropagation();
        });
        btn.addEventListener("click", (e) => {
          e.preventDefault();
          const def = defs.find((d) => d.id === btn.dataset.id);
          if (!def) return;
          def.run();
          // Link and comment open prompts that take focus themselves; the
          // popup must not close or re-render under the modal.
          if (def.id !== "link" && def.id !== "comment") {
            view.focus();
            // Re-render so the active states reflect the command just run.
            close();
            render(view);
          }
        });
      });
      position(view);
    };

    const position = (view: EditorView): void => {
      if (!popup) return;
      const { selection } = view.state;
      try {
        const from = view.coordsAtPos(selection.from);
        const to = view.coordsAtPos(Math.max(selection.from, selection.to - 1));
        const rect = popup.getBoundingClientRect();
        const centerX = (from.left + to.right) / 2 - (rect.width || 200) / 2;
        const maxLeft = Math.max(8, (globalThis.window?.innerWidth ?? 1024) - (rect.width || 200) - 8);
        popup.style.left = `${Math.min(Math.max(8, centerX), maxLeft)}px`;
        const height = rect.height || 36;
        popup.style.top = `${Math.max(8, from.top - height - 8)}px`;
      } catch {
        close();
      }
    };

    return [
      new Plugin({
        key: new PluginKey("fwBubbleToolbar"),
        view: () => ({
          update: (view) => {
            if (selectionIsBubbleWorthy(view)) render(view);
            else close();
          },
          destroy: () => {
            close();
          },
        }),
      }),
    ];
  },
});
