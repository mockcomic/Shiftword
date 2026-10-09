import { Extension } from "@tiptap/core";
import { NodeSelection, Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import { topBlock } from "../editor/block-utils";
import { turnIntoDefs } from "../editor/block-utils";

/**
 * Block handle (Notion's ⋮⋮): appears in the gutter beside the hovered
 * top-level block. Click opens the block menu (turn into / duplicate /
 * delete / move), and dragging it selects the block so ProseMirror's native
 * drag-drop — with the existing gap-cursor as the drop target — moves it.
 *
 * Hover tracking is event-delegated on the editor element (one listener,
 * cheap `closest` lookup), never per-block observers.
 */

export const BlockHandle = Extension.create({
  name: "fwBlockHandle",

  addProseMirrorPlugins() {
    const editor = this.editor;
    let handle: HTMLElement | null = null;
    let menu: HTMLElement | null = null;
    let targetPos: number | null = null;

    const destroy = (): void => {
      hideHandle();
      closeMenu();
    };

    const hideHandle = (): void => {
      if (handle) {
        handle.remove();
        handle = null;
        targetPos = null;
      }
    };

    const closeMenu = (): void => {
      if (menu) {
        menu.remove();
        menu = null;
      }
    };

    /**
     * Find the absolute position of the top-level block whose rendered
     * element is `el` (or contains it): walk pages, then their children.
     *
     * The walk calls `view.nodeDOM()` for EVERY block of EVERY page, which
     * on raw mousemove is O(document) per pixel — it stalls the pointer on
     * long documents. Elements do not survive a document change, so a
     * WeakMap keyed by element (reset whenever the doc changes) turns the
     * whole mouse-hover path into one scan per document state.
     */
    const posCache = new WeakMap<HTMLElement, number>();
    let cachedDoc: unknown = null;
    const blockPosForDom = (view: EditorView, el: HTMLElement): number | null => {
      const doc = view.state.doc;
      if (doc !== cachedDoc) {
        cachedDoc = doc; // entries are keyed by detached elements; a fresh
        // doc invalidates them implicitly by rebuilding the walk below.
      }
      const cached = posCache.get(el);
      if (cached !== undefined) {
        // Fast path: trust the cache only while the node still resolves to
        // this very element (pagination reuses DOM, typing does not).
        try {
          if (view.nodeDOM(cached) === el) return cached;
        } catch {
          /* fall through to the walk */
        }
      }
      let found: number | null = null;
      doc.forEach((page, pagePos) => {
        if (found !== null) return;
        page.forEach((block, off) => {
          if (found !== null) return;
          const blockPos = pagePos + 1 + off;
          try {
            const nodeDom = view.nodeDOM(blockPos) as HTMLElement | null;
            if (nodeDom && (nodeDom === el || nodeDom.contains(el))) {
              found = blockPos;
              posCache.set(el, blockPos);
            }
          } catch {
            /* unrendered position */
          }
        });
      });
      return found;
    };

    const topBlockDomFor = (view: EditorView, pos: number): HTMLElement | null => {
      try {
        return view.nodeDOM(pos) as HTMLElement | null;
      } catch {
        return null;
      }
    };

    const showHandle = (view: EditorView, pos: number): void => {
      const el = topBlockDomFor(view, pos);
      if (!el) {
        hideHandle();
        return;
      }
      if (!handle) {
        handle = document.createElement("div");
        handle.className = "block-handle";
        handle.textContent = "⋮⋮";
        handle.title = "Drag to move, click for actions";
        handle.setAttribute("draggable", "true");
        document.body.appendChild(handle);

        handle.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          if (targetPos !== null) openMenu(view, targetPos);
        });

        // Drag: hand ProseMirror a NodeSelection of the block, then let the
        // browser's drag of that selection run (draggable specs + gap cursor
        // do the rest, including the drop line).
        handle.addEventListener("dragstart", (e) => {
          if (targetPos === null || !editor.state.schema.nodes.page) return;
          e.stopPropagation();
          try {
            const tr = editor.state.tr;
            tr.setSelection(NodeSelection.create(editor.state.doc, targetPos));
            editor.view.dispatch(tr);
          } catch {
            if (e.dataTransfer) e.dataTransfer.effectAllowed = "none";
          }
        });
      }
      const rect = el.getBoundingClientRect();
      handle.style.left = `${Math.max(2, rect.left - 30)}px`;
      handle.style.top = `${rect.top + Math.min(8, Math.max(0, (rect.height - 22) / 2))}px`;
      handle.style.display = "block";
      targetPos = pos;
    };

    const openMenu = (view: EditorView, pos: number): void => {
      closeMenu();
      const node = view.state.doc.nodeAt(pos);
      if (!node) return;
      menu = document.createElement("div");
      menu.className = "block-menu";
      const typeName = node.type.name;
      const items: Array<[string, string, () => void]> = [];

      // Turn into — the same catalog the slash menu uses. Selecting the node
      // first aligns topBlock()/turn-into with the block under the handle,
      // not wherever the caret happened to be.
      for (const def of turnIntoDefs()) {
        items.push([
          `turn:${def.id}`,
          `⇄ ${def.label}`,
          () => {
            selectNode(view, pos);
            def.apply(editor, node.textBetween(0, node.content.size, "\n", ""));
            closeMenu();
          },
        ]);
      }
      items.push([
        "duplicate",
        "⧉ Duplicate",
        () => {
          duplicate(view, pos);
          closeMenu();
        },
      ]);
      if (typeName !== "divider" && typeName !== "pageBreak") {
        items.push([
          "copy",
          "⧉ Copy text",
          () => {
            copyText(view, pos);
            closeMenu();
          },
        ]);
      }
      items.push([
        "up",
        "↑ Move up",
        () => {
          move(view, pos, -1);
          closeMenu();
        },
      ]);
      items.push([
        "down",
        "↓ Move down",
        () => {
          move(view, pos, 1);
          closeMenu();
        },
      ]);
      items.push([
        "delete",
        "🗑 Delete",
        () => {
          remove(view, pos);
          closeMenu();
        },
      ]);

      menu.innerHTML = items
        .map(
          ([id, label]) => `<button type="button" class="block-menu-item" data-id="${id}">${label}</button>`,
        )
        .join("");
      let built = 0;
      menu.querySelectorAll<HTMLButtonElement>(".block-menu-item").forEach((btn) => {
        const item = items[built++];
        btn.addEventListener("mousedown", (e) => e.preventDefault());
        btn.addEventListener("click", () => item?.[2]());
      });
      document.body.appendChild(menu);
      const el = topBlockDomFor(view, pos);
      const rect = el?.getBoundingClientRect();
      menu.style.left = `${rect ? Math.max(8, rect.left - 30) : 40}px`;
      menu.style.top = `${(rect?.bottom ?? 120) + 4}px`;
    };

    const selectNode = (view: EditorView, pos: number): void => {
      try {
        view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
      } catch {
        /* atom inside an empty page — nothing to select, the action may still work */
      }
    };

    const duplicate = (view: EditorView, pos: number): void => {
      const node = view.state.doc.nodeAt(pos);
      if (!node) return;
      const clone = node.type.create(node.attrs, node.content, node.marks);
      view.dispatch(view.state.tr.insert(pos + node.nodeSize, clone));
    };

    const copyText = (view: EditorView, pos: number): void => {
      const node = view.state.doc.nodeAt(pos);
      if (!node) return;
      void navigator.clipboard
        ?.writeText(node.textBetween(0, node.content.size, "\n", " "))
        .catch(() => undefined);
    };

    const remove = (view: EditorView, pos: number): void => {
      const node = view.state.doc.nodeAt(pos);
      if (!node) return;
      const tr = view.state.tr;
      // The page node requires `block+`: deleting its last block would make
      // an invalid document, so an emptied page gets a fresh paragraph.
      const $pos = view.state.doc.resolve(pos);
      const page = $pos.node(1);
      if (page && page.childCount <= 1) {
        tr.replaceWith(pos, pos + node.nodeSize, view.state.schema.nodes.paragraph.create());
      } else {
        tr.delete(pos, pos + node.nodeSize);
      }
      view.dispatch(tr);
    };

    const move = (view: EditorView, pos: number, dir: -1 | 1): void => {
      const node = view.state.doc.nodeAt(pos);
      if (!node) return;
      const $pos = view.state.doc.resolve(pos);
      const page = $pos.node(1);
      if (!page) return;
      const index = $pos.index(1);
      const swapIndex = dir === -1 ? index - 1 : index + 1;
      if (swapIndex < 0 || swapIndex >= page.childCount) return;
      const other = page.child(swapIndex);
      const firstPos = dir === -1 ? pos - other.nodeSize : pos;
      const to = dir === -1 ? pos + node.nodeSize : pos + node.nodeSize + other.nodeSize;
      const ordered = dir === -1 ? [other, node] : [node, other];
      const tr = view.state.tr.replaceWith(
        firstPos,
        to,
        ordered.map((n) => n.type.create(n.attrs, n.content, n.marks)),
      );
      try {
        tr.setSelection(TextSelection.near(tr.doc.resolve(firstPos + 1), 1));
      } catch {
        /* keep the old selection */
      }
      view.dispatch(tr);
    };

    return [
      new Plugin({
        key: new PluginKey("fwBlockHandle"),
        props: {
          handleDOMEvents: {
            mousemove: (view, event) => {
              if (menu) return false; // stay put while the menu is open
              const target = event.target as HTMLElement | null;
              if (handle && handle.contains(target)) return false;
              // Walk up to the direct child of the page — that is the block
              // the handle belongs to. (`closest(".fw-page > *")` is not
              // reliable across ancestor chains for child combinators.)
              let el: HTMLElement | null = target;
              while (el && el.parentElement && !el.parentElement.classList.contains("fw-page")) {
                el = el.parentElement;
              }
              if (!el || !el.parentElement?.classList.contains("fw-page")) {
                if (!target?.closest?.(".block-handle")) hideHandle();
                return false;
              }
              const pos = blockPosForDom(view, el);
              if (pos === null) return false;
              showHandle(view, pos);
              return false;
            },
            mouseleave: (view) => {
              void view;
              if (!menu) hideHandle();
              return false;
            },
          },
        },
        view: () => ({
          update: (view) => {
            // A transaction moved/removed the block the handle points at —
            // without this the handle would float over the wrong block.
            if (handle && targetPos !== null) {
              const node = view.state.doc.nodeAt(targetPos);
              if (!node) hideHandle();
              else {
                const el = topBlockDomFor(view, targetPos);
                if (el) {
                  const rect = el.getBoundingClientRect();
                  handle.style.left = `${Math.max(2, rect.left - 30)}px`;
                  handle.style.top = `${rect.top + Math.min(8, Math.max(0, (rect.height - 22) / 2))}px`;
                }
              }
            }
          },
          destroy,
        }),
      }),
    ];
  },
});

/** Public helper used by tests to reach the shared "turn into" catalog. */
export { turnIntoDefs, topBlock };
