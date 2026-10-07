import { Extension } from "@tiptap/core";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import { BLOCK_DEFS, filterBlockDefs } from "../editor/block-utils";
import type { BlockDef } from "../editor/block-utils";
import { insertImageAtCursor } from "./image-insert";

/**
 * Slash command menu: type `/` at the start of a line, or after a space, and
 * a filterable block picker opens — Notion's signature affordance.
 *
 * Registered as an Extension so the ProseMirror plugin is part of the
 * editor's state from construction (no unsupported runtime attach), and the
 * popup is destroyed together with the editor view — a previous boot can
 * never leak a live popup into the next one.
 */

const MAX_QUERY_LEN = 32;

export const SlashMenu = Extension.create({
  name: "fwSlashMenu",

  addProseMirrorPlugins() {
    const editor = this.editor;

    let popup: HTMLElement | null = null;
    let open = false;
    let slashPos = -1;
    let query = "";
    let active = 0;
    let items: BlockDef[] = [];

    const close = (): void => {
      open = false;
      query = "";
      active = 0;
      items = [];
      if (popup) {
        popup.remove();
        popup = null;
      }
    };

    const computeItems = (): void => {
      const notion = document.body.classList.contains("notion-mode");
      const visible = notion ? BLOCK_DEFS.filter((d) => !d.wordOnly) : BLOCK_DEFS;
      items = filterBlockDefs(visible, query);
      if (active >= items.length) active = Math.max(0, items.length - 1);
    };

    const position = (view: EditorView): void => {
      if (!popup) return;
      let coords: ReturnType<EditorView["coordsAtPos"]> | null = null;
      try {
        coords = view.coordsAtPos(slashPos);
      } catch {
        coords = null; // jsdom / detached view
      }
      if (!coords) {
        popup.style.left = "40%";
        popup.style.top = "30%";
        return;
      }
      const width = 280;
      const maxLeft = Math.max(8, (globalThis.window?.innerWidth ?? 1024) - width - 8);
      popup.style.left = `${Math.min(Math.max(8, coords.left), maxLeft)}px`;
      const height = popup.getBoundingClientRect().height || 240;
      const viewportH = globalThis.window?.innerHeight ?? 768;
      const belowFits = coords.bottom + 6 + height < viewportH - 8;
      popup.style.top = `${belowFits ? coords.bottom + 6 : Math.max(8, coords.top - height - 6)}px`;
    };

    const run = (view: EditorView, index: number): void => {
      const item = items[index];
      if (!item) return;
      const head = view.state.selection.head;
      const from = Math.min(slashPos, head);
      const to = Math.max(slashPos, head);
      close();
      // Remove the "/query" text, park the caret where the slash was, then
      // let the catalog insert the block (block-utils replaces the current
      // empty block or appends after a non-empty one).
      view.dispatch(view.state.tr.delete(from, to));
      try {
        const at = Math.min(from, view.state.doc.content.size - 1);
        view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, Math.max(1, at))));
      } catch {
        /* selection mapping may fail on an exotic doc — the delete still ran */
      }
      view.focus();
      if (item.id === "image") {
        void insertImageAtCursor(editor);
        return;
      }
      item.apply(editor, "");
    };

    const render = (view: EditorView): void => {
      if (!open) return;
      if (!popup) {
        popup = document.createElement("div");
        popup.className = "slash-menu";
        popup.setAttribute("role", "listbox");
        document.body.appendChild(popup);
      }
      if (items.length === 0) {
        popup.innerHTML = `<div class="slash-empty">No matching blocks</div>`;
      } else {
        popup.innerHTML = items
          .map(
            (item, i) =>
              `<div class="slash-item${i === active ? " active" : ""}" role="option" data-index="${i}">` +
              `<span class="slash-icon">${item.icon}</span><span class="slash-label">${item.label}</span></div>`,
          )
          .join("");
        popup.querySelectorAll<HTMLElement>(".slash-item").forEach((el) => {
          // mousedown (not click): preventDefault keeps the editor focused,
          // so the caret is still where the "/query" was when the block lands.
          el.addEventListener("mousedown", (e) => {
            e.preventDefault();
            e.stopPropagation();
            run(view, Number(el.dataset.index));
          });
          el.addEventListener("mousemove", () => {
            active = Number(el.dataset.index);
            popup
              ?.querySelectorAll(".slash-item")
              .forEach((n, i) => n.classList.toggle("active", i === active));
          });
        });
      }
      position(view);
    };

    const refreshFromDoc = (view: EditorView): void => {
      if (!open) return;
      const sel = view.state.selection;
      if (!sel.empty || sel.head < slashPos) {
        close();
        return;
      }
      const text = view.state.doc.textBetween(slashPos, sel.head, "\n", "\n");
      if (!text.startsWith("/") || text.length - 1 > MAX_QUERY_LEN) {
        close();
        return;
      }
      query = text.slice(1);
      active = 0;
      computeItems();
      render(view);
    };

    const tryOpen = (view: EditorView): void => {
      if (open) {
        refreshFromDoc(view);
        return;
      }
      const sel = view.state.selection;
      if (!sel.empty || !sel.$from.parent.isTextblock) return;
      const pos = sel.head;
      // Only the act of typing "/" opens the menu — the character must be the
      // one immediately before the caret. Every other update path re-fails
      // this check harmlessly.
      const before = view.state.doc.textBetween(Math.max(0, pos - 1), pos, "\n", " ");
      if (before !== "/") return;
      const lineStartOk = pos - 1 <= sel.$from.start();
      const prev = lineStartOk ? " " : view.state.doc.textBetween(Math.max(0, pos - 2), pos - 1, "\n", " ");
      if (!lineStartOk && /\S/.test(prev)) return;
      open = true;
      slashPos = pos - 1;
      query = "";
      active = 0;
      computeItems();
      render(view);
    };

    return [
      new Plugin({
        key: new PluginKey("fwSlashMenu"),
        props: {
          handleKeyDown: (view, event) => {
            if (!open) return false;
            switch (event.key) {
              case "Escape":
                close();
                return true;
              case "ArrowDown":
                active = Math.min(items.length - 1, active + 1);
                render(view);
                return true;
              case "ArrowUp":
                active = Math.max(0, active - 1);
                render(view);
                return true;
              case "Enter":
              case "Tab":
                if (items.length === 0) {
                  close();
                  return true;
                }
                run(view, active);
                return true;
              default:
                return false;
            }
          },
        },
        view: () => ({
          update: (view) => {
            // Handles both directions: opening on a freshly typed "/" and
            // refreshing the filter while open.
            tryOpen(view);
          },
          destroy: () => {
            close();
          },
        }),
      }),
    ];
  },
});
