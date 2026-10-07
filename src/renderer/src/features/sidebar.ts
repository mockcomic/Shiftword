import type { PageMeta } from "../../../shared/types";

/**
 * The workspace sidebar: a local page tree in the left gutter, Notion-style.
 *
 * Vanilla DOM in the repo's house style (see toolbar.ts's dropdown pattern):
 * one delegated listener set, full re-render on data change (the tree is
 * tens of rows — virtualisation would be the bug, not the optimisation).
 *
 * The element itself is created here and appended to <body>; visibility is
 * CSS (`body.sidebar-open #sidebar`), driven by the persisted setting.
 */

export interface SidebarHandlers {
  onOpenPage(id: string): void;
  onNewPage(parentId: string | null): void;
  onRename(id: string): void;
  onSetIcon(id: string): void;
  onDelete(id: string): void;
  onMove(id: string, parentId: string | null, order: number): void;
  onQuickSwitcher(): void;
}

export interface Sidebar {
  refresh(pages: PageMeta[]): void;
  setActive(id: string | null): void;
  element: HTMLElement;
}

export function createSidebar(handlers: SidebarHandlers): Sidebar {
  let pages: PageMeta[] = [];
  let activeId: string | null = null;
  const expanded = new Set<string>();

  const root = document.createElement("aside");
  root.id = "sidebar";
  root.innerHTML = `
    <div class="side-head">
      <span class="side-title">FreeWord</span>
      <span class="side-head-actions">
        <button type="button" id="side-search" title="Quick switcher (Ctrl+K)">🔍</button>
        <button type="button" id="side-new" title="New page (Ctrl+Alt+N)">＋</button>
      </span>
    </div>
    <div class="side-tree" id="side-tree"></div>
    <div class="side-foot">
      <button type="button" id="side-new-root">＋ New page</button>
    </div>
  `;
  document.body.appendChild(root);

  const tree = root.querySelector<HTMLElement>("#side-tree")!;

  root.querySelector("#side-new")!.addEventListener("click", () => handlers.onNewPage(null));
  root.querySelector("#side-new-root")!.addEventListener("click", () => handlers.onNewPage(null));
  root.querySelector("#side-search")!.addEventListener("click", () => handlers.onQuickSwitcher());

  // Drop onto the empty space below the tree = move to root, end of list.
  tree.addEventListener("dragover", (e) => {
    if ((e.target as HTMLElement).closest(".side-row")) return;
    e.preventDefault();
  });
  tree.addEventListener("drop", (e) => {
    const overRow = (e.target as HTMLElement).closest(".side-row") as HTMLElement | null;
    const id = e.dataTransfer?.getData("text/fw-page");
    if (!id || overRow) return;
    e.preventDefault();
    const roots = pages.filter((p) => p.parentId === null);
    handlers.onMove(id, null, roots.length === 0 ? 0 : Math.max(...roots.map((r) => r.order)) + 1);
  });

  function childrenOf(parentId: string | null): PageMeta[] {
    return pages
      .filter((p) => p.parentId === parentId)
      .sort((a, b) => a.order - b.order || a.title.localeCompare(b.title));
  }

  function hasChildren(id: string): boolean {
    return pages.some((p) => p.parentId === id);
  }

  /** Guard against cycles from the UI side: `descendant` check (store re-checks). */
  function isDescendant(candidateId: string, ancestorId: string): boolean {
    let cursor = pages.find((p) => p.id === candidateId);
    while (cursor) {
      if (cursor.id === ancestorId) return true;
      cursor = cursor.parentId ? pages.find((p) => p.id === cursor!.parentId) : undefined;
    }
    return false;
  }

  function renderRows(parentId: string | null, depth: number): HTMLElement[] {
    const out: HTMLElement[] = [];
    const sibs = childrenOf(parentId);
    for (const page of sibs) {
      const row = document.createElement("div");
      row.className = "side-row" + (page.id === activeId ? " active" : "");
      row.draggable = true;
      row.dataset.id = page.id;
      row.style.paddingLeft = `${8 + depth * 14}px`;
      const chev = document.createElement("button");
      chev.type = "button";
      chev.className = "side-chevron";
      chev.textContent = hasChildren(page.id) ? (expanded.has(page.id) ? "▾" : "▸") : "";
      chev.addEventListener("click", (e) => {
        e.stopPropagation();
        if (!hasChildren(page.id)) return;
        if (expanded.has(page.id)) expanded.delete(page.id);
        else expanded.add(page.id);
        rerender();
      });
      const icon = document.createElement("span");
      icon.className = "side-icon";
      icon.textContent = page.icon || "📄";
      const label = document.createElement("span");
      label.className = "side-label";
      label.textContent = page.title || "Untitled";
      label.title = page.title;
      const more = document.createElement("button");
      more.type = "button";
      more.className = "side-more";
      more.textContent = "⋯";
      more.title = "Page actions";
      more.addEventListener("click", (e) => {
        e.stopPropagation();
        openRowMenu(more, page);
      });
      row.append(chev, icon, label, more);
      row.addEventListener("click", () => handlers.onOpenPage(page.id));
      attachDnd(row, page);
      out.push(row);
      if (expanded.has(page.id)) out.push(...renderRows(page.id, depth + 1));
    }
    return out;
  }

  function attachDnd(row: HTMLElement, page: PageMeta): void {
    row.addEventListener("dragstart", (e) => {
      e.dataTransfer?.setData("text/fw-page", page.id);
      if (e.dataTransfer) e.dataTransfer.effectAllowed = "move";
    });
    row.addEventListener("dragover", (e) => {
      if (!e.dataTransfer) return;
      const dragged = draggedHint();
      if (!dragged || dragged === page.id || isDescendant(dragged, page.id)) return;
      e.preventDefault();
      const rect = row.getBoundingClientRect();
      const rel = (e.clientY - rect.top) / rect.height;
      row.classList.remove("drop-before", "drop-after", "drop-child");
      row.classList.add(rel < 0.25 ? "drop-before" : rel > 0.75 ? "drop-after" : "drop-child");
    });
    row.addEventListener("dragleave", () => {
      row.classList.remove("drop-before", "drop-after", "drop-child");
    });
    row.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      row.classList.remove("drop-before", "drop-after", "drop-child");
      const dragged = e.dataTransfer?.getData("text/fw-page") || draggedHint();
      if (!dragged || dragged === page.id || isDescendant(dragged, page.id)) return;
      const rect = row.getBoundingClientRect();
      const rel = ((e as DragEvent).clientY - rect.top) / rect.height;
      const siblings = childrenOf(page.parentId).filter((p) => p.id !== dragged);
      if (rel >= 0.25 && rel <= 0.75) {
        // Into this page.
        const kids = childrenOf(page.id).filter((p) => p.id !== dragged);
        handlers.onMove(dragged, page.id, kids.length === 0 ? 0 : Math.max(...kids.map((k) => k.order)) + 1);
        expanded.add(page.id);
      } else {
        const idx = childrenOf(page.parentId).findIndex((p) => p.id === page.id);
        const at = rel < 0.25 ? idx : idx + 1;
        // Re-number the sibling list with the moved row spliced in.
        const list = siblings.map((p) => p.id);
        list.splice(Math.max(0, Math.min(list.length, at)), 0, dragged);
        handlers.onMove(dragged, page.parentId, list.indexOf(dragged));
      }
    });
  }

  // The dataTransfer payload is unreadable during dragover; remember the id
  // from dragstart as the practical cross-browser answer.
  let lastDragged: string | null = null;
  const draggedHint = (): string | null => lastDragged;
  root.addEventListener("dragstart", (e) => {
    lastDragged = (e.target as HTMLElement).closest?.(".side-row")?.getAttribute("data-id") ?? null;
  });
  root.addEventListener("dragend", () => {
    lastDragged = null;
    rerender();
  });

  let rowMenu: HTMLElement | null = null;
  function closeRowMenu(): void {
    if (rowMenu) {
      rowMenu.remove();
      rowMenu = null;
    }
  }
  function openRowMenu(anchor: HTMLElement, page: PageMeta): void {
    closeRowMenu();
    rowMenu = document.createElement("div");
    rowMenu.className = "block-menu";
    const items: Array<[string, () => void]> = [
      [
        "✎ Rename",
        () => {
          closeRowMenu();
          handlers.onRename(page.id);
        },
      ],
      [
        "😀 Set icon",
        () => {
          closeRowMenu();
          handlers.onSetIcon(page.id);
        },
      ],
      [
        "＋ New subpage",
        () => {
          closeRowMenu();
          handlers.onNewPage(page.id);
        },
      ],
      [
        "🗑 Delete",
        () => {
          closeRowMenu();
          handlers.onDelete(page.id);
        },
      ],
    ];
    rowMenu.innerHTML = items
      .map(([label]) => `<button type="button" class="block-menu-item">${label}</button>`)
      .join("");
    Array.from(rowMenu.querySelectorAll<HTMLButtonElement>(".block-menu-item")).forEach((btn, i) => {
      btn.addEventListener("mousedown", (e) => e.preventDefault());
      btn.addEventListener("click", () => items[i][1]());
    });
    document.body.appendChild(rowMenu);
    const rect = anchor.getBoundingClientRect();
    rowMenu.style.left = `${Math.max(4, rect.left - 120)}px`;
    rowMenu.style.top = `${rect.bottom + 2}px`;
  }
  document.addEventListener("mousedown", (e) => {
    if (rowMenu && !rowMenu.contains(e.target as Node)) closeRowMenu();
  });

  function rerender(): void {
    tree.replaceChildren(...renderRows(null, 0));
    if (pages.length === 0) {
      const empty = document.createElement("div");
      empty.className = "side-empty";
      empty.textContent = "No pages yet — create one.";
      tree.appendChild(empty);
    }
  }

  return {
    element: root,
    refresh(newPages) {
      pages = newPages.slice();
      // Expand anything that has a parent chain so a fresh workspace reads as
      // a list, not a puzzle.
      for (const p of pages)
        if (hasChildren(p.id) && expanded.size === 0 && p.parentId === null) expanded.add(p.id);
      rerender();
    },
    setActive(id) {
      activeId = id;
      rerender();
    },
  };
}

/** Fuzzy filter for the quick switcher: every query token appears in the title. */
export function filterPages(pages: PageMeta[], query: string): PageMeta[] {
  const tokens = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return pages.slice(0, 20);
  const scored = pages.map((p) => {
    const hay = `${p.icon} ${p.title}`.toLowerCase();
    const hit = tokens.every((t) => hay.includes(t));
    const starts = tokens.every((t) => p.title.toLowerCase().startsWith(t));
    return { p, hit, starts };
  });
  return scored
    .filter((s) => s.hit)
    .sort((a, b) => Number(b.starts) - Number(a.starts) || b.p.updatedAt - a.p.updatedAt)
    .slice(0, 20)
    .map((s) => s.p);
}
