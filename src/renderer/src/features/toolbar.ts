import type { Editor } from "@tiptap/core";
import { CellSelection } from "@tiptap/pm/tables";
import { insertImageAtCursor } from "./image-insert";
import { promptLink } from "./link-prompt";
import { ZOOM_STEPS, getZoom, setZoom } from "./zoom";
import { BLOCK_DEFS, topBlock } from "../editor/block-utils";
import { BLOCK_NODE_NAMES } from "../editor/blocks";

const FONTS = [
  "Calibri",
  "Cambria",
  "Arial",
  "Times New Roman",
  "Georgia",
  "Verdana",
  "Courier New",
  "Consolas",
  "Segoe UI",
  "Tahoma",
];

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];

const COLOR_SWATCHES = [
  "#000000",
  "#7F7F7F",
  "#C00000",
  "#ED7D31",
  "#FFC000",
  "#70AD47",
  "#2E75B6",
  "#7030A0",
  "#E91E63",
  "#404040",
];

const HIGHLIGHTS: Array<{ label: string; color: string | null }> = [
  { label: "None", color: null },
  { label: "Yellow", color: "#FFFF00" },
  { label: "Green", color: "#92D050" },
];

interface TextStyleAttrs {
  fontFamily?: string | null;
  fontSize?: string | null;
  color?: string | null;
}

/** Read the textStyle mark's inline-style attrs (color / font-family / font-size) at the cursor. */
function textStyleAttrs(editor: Editor): TextStyleAttrs {
  return editor.getAttributes("textStyle") as TextStyleAttrs;
}

/** True when the selection spans table cells (drag-selected, or a row/column selection). */
function cellSelectionActive(editor: Editor): boolean {
  return editor.state.selection instanceof CellSelection;
}

/** Look up a catalog block type by its style-select id. */
function blockDef(id: string) {
  return BLOCK_DEFS.find((d) => d.id === id);
}

/**
 * Apply a style-select choice: plain paragraph/heading use `setNode` (the
 * Tiptap-native path), the Notion-style blocks go through the shared catalog
 * so the slash menu, the block menu and this select convert identically.
 */
function applyStyleChoice(editor: Editor, value: string): void {
  if (value === "paragraph") {
    editor.chain().focus().setNode("paragraph").run();
    return;
  }
  if (value === "h1" || value === "h2" || value === "h3") {
    editor
      .chain()
      .focus()
      .setNode("heading", { level: Number(value.slice(1)) })
      .run();
    return;
  }
  const def = blockDef(value);
  if (!def) return;
  const top = topBlock(editor);
  const text = top ? top.node.textBetween(0, top.node.content.size, "\n", "") : "";
  def.apply(editor, text);
}

let syncTimer: number | null = null;
let openMenu: HTMLElement | null = null;
let outsideDownHandler: ((e: MouseEvent) => void) | null = null;

function removeOutsideHandler(): void {
  if (outsideDownHandler) {
    document.removeEventListener("mousedown", outsideDownHandler, true);
    outsideDownHandler = null;
  }
}

function closeMenus(): void {
  removeOutsideHandler();
  if (openMenu) {
    openMenu.remove();
    openMenu = null;
  }
}

/**
 * Open (or close) a dropdown anchored to a toolbar button.
 *
 * The closed-state test used to compare `openMenu` against a
 * `[data-menu-for]` element looked up *inside the anchor's group* — but the
 * menu is appended to `document.body`, so that lookup was always `null`, and
 * with no menu open the check read `null === null` and returned early: the
 * color and highlight dropdowns could never open at all, and clicking the
 * anchor again could never close them. The menu records its own anchor id, and
 * that is what the toggle compares.
 */
function toggleMenu(anchor: HTMLElement, build: (menu: HTMLElement) => void): void {
  if (openMenu && openMenu.dataset.menuFor === anchor.id) {
    closeMenus();
    return;
  }
  closeMenus();
  const menu = document.createElement("div");
  menu.className = "tb-dropdown";
  menu.dataset.menuFor = anchor.id;
  build(menu);
  document.body.appendChild(menu);
  const rect = anchor.getBoundingClientRect();
  menu.style.top = `${rect.bottom + 2}px`;
  menu.style.left = `${Math.max(4, rect.left)}px`;
  openMenu = menu;
  // One outside-click handler at a time, removed whenever the menu closes —
  // the previous version added a new listener per menu and only ever removed
  // the one whose `openMenu === menu` guard happened to pass.
  const onDown = (e: MouseEvent): void => {
    if (!menu.contains(e.target as Node) && e.target !== anchor) closeMenus();
  };
  outsideDownHandler = onDown;
  // Deferred so the click that opened the menu does not also close it.
  window.setTimeout(() => {
    if (openMenu === menu && outsideDownHandler === onDown) {
      document.addEventListener("mousedown", onDown, true);
    }
  }, 0);
}

/** Debounced toolbar refresh; see `attachToolbarSync` for what triggers it. */
function scheduleSync(editor: Editor): void {
  if (syncTimer !== null) window.clearTimeout(syncTimer);
  syncTimer = window.setTimeout(() => {
    syncTimer = null;
    if (!editor.isDestroyed) syncToolbar(editor);
  }, 50);
}

/**
 * Keep the toolbar mirrors in step with the editor.
 *
 * `selectionUpdate` alone is not enough: toggling a mark with an empty
 * selection changes `storedMarks`, not the selection, so the Bold button used
 * to stay visually off until the caret moved. `transactions` catches that
 * case (and node/mark changes generally) while the 50 ms debounce collapses
 * the churn of typing.
 */
export function attachToolbarSync(editor: Editor): void {
  editor.on("selectionUpdate", () => scheduleSync(editor));
  editor.on("transaction", ({ transaction }) => {
    if (transaction.storedMarksSet || transaction.docChanged || transaction.selectionSet) {
      scheduleSync(editor);
    }
  });
}

export function createToolbar(editor: Editor): void {
  const root = document.querySelector<HTMLElement>("#toolbar");
  if (!root) return;

  root.innerHTML = `
    <div class="tb-group">
      <select id="tb-style" title="Paragraph style">
        <option value="paragraph">Paragraph</option>
        <option value="h1">Heading 1</option>
        <option value="h2">Heading 2</option>
        <option value="h3">Heading 3</option>
        <option value="taskList">To-do list</option>
        <option value="quote">Quote</option>
        <option value="callout">Callout</option>
        <option value="codeBlock">Code block</option>
      </select>
    </div>
    <div class="tb-group">
      <select id="tb-font" title="Font family"></select>
      <select id="tb-size" title="Font size"></select>
    </div>
    <div class="tb-group">
      <button id="tb-bold" title="Bold (Ctrl+B)"><b>B</b></button>
      <button id="tb-italic" title="Italic (Ctrl+I)"><i>I</i></button>
      <button id="tb-underline" title="Underline (Ctrl+U)">U&#8242;</button>
      <button id="tb-strike" title="Strikethrough (Ctrl+Shift+X)"><s>S</s></button>
      <button id="tb-code" title="Monospace">&lt;/&gt;</button>
    </div>
    <div class="tb-group">
      <button id="tb-link" title="Insert link">↗</button>
      <button id="tb-unlink" title="Remove link">⨯&#8203;↗</button>
    </div>
    <div class="tb-group">
      <button id="tb-color" title="Text color">A<span class="tb-swatch-bar" id="tb-color-bar"></span><span class="tb-caret">▾</span></button>
      <button id="tb-highlight" title="Highlight">▟<span class="tb-caret">▾</span></button>
    </div>
    <div class="tb-group">
      <button id="tb-align-left" title="Align left (Ctrl+L)">⯇</button>
      <button id="tb-align-center" title="Center (Ctrl+E)">≡</button>
      <button id="tb-align-right" title="Align right (Ctrl+R)">⯈</button>
      <button id="tb-align-justify" title="Justify">☰</button>
    </div>
    <div class="tb-group">
      <select id="tb-spacing" title="Line spacing">
        <option value="1">1.0</option>
        <option value="1.15">1.15</option>
        <option value="1.5">1.5</option>
        <option value="2">2.0</option>
      </select>
    </div>
    <div class="tb-group">
      <button id="tb-bullet" title="Bulleted list">•≡</button>
      <button id="tb-ordered" title="Numbered list">1≡</button>
      <button id="tb-indent" title="Indent list item (Tab)">→</button>
      <button id="tb-outdent" title="Outdent list item (Shift+Tab)">←</button>
    </div>
    <div class="tb-group">
      <button id="tb-table" title="Insert table (3 × 3)">⊞</button>
      <button id="tb-table-row-add" title="Add row below">Row +</button>
      <button id="tb-table-row-del" title="Delete row">Row −</button>
      <button id="tb-table-col-add" title="Add column">Col +</button>
      <button id="tb-table-col-del" title="Delete column">Col −</button>
      <button id="tb-table-header" title="Toggle header row">Hr</button>
      <button id="tb-table-merge" title="Merge selected cells">▣</button>
      <button id="tb-table-split" title="Split merged cells">▤</button>
      <button id="tb-table-del" title="Delete table">⊘</button>
    </div>
    <div class="tb-group">
      <button id="tb-image" title="Insert image">Image</button>
    </div>
    <div class="tb-group">
      <button id="tb-break" class="tb-word-only" title="Insert page break (Ctrl+Enter)">⤓</button>
    </div>
    <div class="tb-group tb-grow"></div>
    <div class="tb-group">
      <select id="tb-zoom" title="Zoom">
        ${ZOOM_STEPS.map((z) => `<option value="${z}">${z}%</option>`).join("")}
      </select>
    </div>
  `;

  const $ = <T extends HTMLElement>(id: string): T => root.querySelector<HTMLElement>(`#${id}`) as T;
  const fontSel = $("tb-font");
  const sizeSel = $("tb-size");
  fontSel.innerHTML = `<option value="">Auto</option>${FONTS.map((f) => `<option value="${f}">${f}</option>`).join("")}`;
  sizeSel.innerHTML = `<option value="">Auto</option>${SIZES.map((s) => `<option value="${s}pt">${s}</option>`).join("")}`;

  // --- paragraph style -------------------------------------------------
  $("tb-style").addEventListener("change", (e) => {
    applyStyleChoice(editor, (e.target as HTMLSelectElement).value);
  });

  // --- font family / size ----------------------------------------------
  fontSel.addEventListener("change", (e) => {
    const family = (e.target as HTMLSelectElement).value;
    if (family) editor.chain().focus().setFontFamily(family).run();
    else editor.chain().focus().unsetFontFamily().run();
  });
  sizeSel.addEventListener("change", (e) => {
    const size = (e.target as HTMLSelectElement).value;
    if (size) editor.chain().focus().setFontSize(size).run();
    else editor.chain().focus().unsetFontSize().run();
  });

  // --- marks -------------------------------------------------------------
  $("tb-bold").addEventListener("click", () => editor.chain().focus().toggleMark("bold").run());
  $("tb-italic").addEventListener("click", () => editor.chain().focus().toggleMark("italic").run());
  $("tb-underline").addEventListener("click", () => editor.chain().focus().toggleMark("underline").run());
  $("tb-strike").addEventListener("click", () => editor.chain().focus().toggleMark("strike").run());
  $("tb-code").addEventListener("click", () => editor.chain().focus().toggleCode().run());

  // --- links --------------------------------------------------------------
  $("tb-link").addEventListener("click", () => void promptLink(editor));
  $("tb-unlink").addEventListener("click", () => editor.chain().focus().unsetLink().run());

  // --- text color ---------------------------------------------------------
  $("tb-color").addEventListener("click", () => {
    toggleMenu($("tb-color"), (menu) => {
      const grid = document.createElement("div");
      grid.className = "swatch-grid";
      for (const color of COLOR_SWATCHES) {
        const swatch = document.createElement("button");
        swatch.className = "swatch";
        swatch.style.background = color;
        swatch.title = color;
        swatch.addEventListener("click", () => {
          const current = textStyleAttrs(editor).color;
          if (current === color) editor.chain().focus().unsetColor().run();
          else editor.chain().focus().setColor(color).run();
          closeMenus();
          syncToolbar(editor);
        });
        grid.appendChild(swatch);
      }
      const row = document.createElement("div");
      row.className = "swatch-row";
      const auto = document.createElement("button");
      auto.textContent = "Auto (no color)";
      auto.addEventListener("click", () => {
        editor.chain().focus().unsetColor().run();
        closeMenus();
        syncToolbar(editor);
      });
      row.appendChild(auto);
      menu.appendChild(grid);
      menu.appendChild(row);
    });
  });

  // --- highlight -----------------------------------------------------------
  $("tb-highlight").addEventListener("click", () => {
    toggleMenu($("tb-highlight"), (menu) => {
      const row = document.createElement("div");
      row.className = "swatch-row";
      for (const opt of HIGHLIGHTS) {
        const btn = document.createElement("button");
        btn.textContent = opt.label;
        if (opt.color) btn.style.background = opt.color;
        btn.addEventListener("click", () => {
          if (opt.color) editor.chain().focus().toggleHighlight({ color: opt.color }).run();
          else editor.chain().focus().toggleHighlight().run();
          closeMenus();
          syncToolbar(editor);
        });
        row.appendChild(btn);
      }
      menu.appendChild(row);
    });
  });

  // --- alignment ------------------------------------------------------------
  for (const [id, align] of [
    ["tb-align-left", "left"],
    ["tb-align-center", "center"],
    ["tb-align-right", "right"],
    ["tb-align-justify", "justify"],
  ] as const) {
    $(id).addEventListener("click", () => editor.chain().focus().setTextAlign(align).run());
  }

  // --- line spacing -----------------------------------------------------------
  $("tb-spacing").addEventListener("change", (e) => {
    const value = parseFloat((e.target as HTMLSelectElement).value);
    editor.chain().focus().updateAttributes("paragraph", { lineSpacing: value }).run();
  });

  // --- lists ---------------------------------------------------------------------
  $("tb-bullet").addEventListener("click", () => editor.chain().focus().toggleBulletList().run());
  $("tb-ordered").addEventListener("click", () => editor.chain().focus().toggleOrderedList().run());
  // The Tab / Shift+Tab keymaps live in the list extension; the buttons give the
  // same commands to anyone who cannot discover the key.
  $("tb-indent").addEventListener("click", () => editor.chain().focus().sinkListItem("listItem").run());
  $("tb-outdent").addEventListener("click", () => editor.chain().focus().liftListItem("listItem").run());

  // --- tables ------------------------------------------------------------------------
  $("tb-table").addEventListener("click", () => {
    editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
  });
  $("tb-table-row-add").addEventListener("click", () => editor.chain().focus().addRowAfter().run());
  $("tb-table-row-del").addEventListener("click", () => editor.chain().focus().deleteRow().run());
  $("tb-table-col-add").addEventListener("click", () => editor.chain().focus().addColumnAfter().run());
  $("tb-table-col-del").addEventListener("click", () => editor.chain().focus().deleteColumn().run());
  $("tb-table-header").addEventListener("click", () => editor.chain().focus().toggleHeaderRow().run());
  $("tb-table-merge").addEventListener("click", () => editor.chain().focus().mergeOrSplit().run());
  $("tb-table-split").addEventListener("click", () => editor.chain().focus().splitCell().run());
  $("tb-table-del").addEventListener("click", () => editor.chain().focus().deleteTable().run());

  // --- image ---------------------------------------------------------------------------
  $("tb-image").addEventListener("click", () => {
    void insertImageAtCursor(editor);
  });

  // --- page break ------------------------------------------------------------------
  $("tb-break").addEventListener("click", () => {
    editor.chain().focus().insertContentAt(editor.state.selection, { type: "pageBreak" }).run();
  });

  // --- zoom --------------------------------------------------------------------------
  $("tb-zoom").addEventListener("change", (e) => {
    setZoom(parseInt((e.target as HTMLSelectElement).value, 10), editor);
    syncToolbar(editor);
  });

  attachToolbarSync(editor);

  syncToolbar(editor);
}

/** Mirror the active node/mark state into toolbar controls. */
export function syncToolbar(editor: Editor): void {
  const root = document.querySelector<HTMLElement>("#toolbar");
  if (!root || editor.isDestroyed) return;
  const $ = (id: string): HTMLElement | null => root.querySelector(`#${id}`);

  const styleSel = $("tb-style") as HTMLSelectElement | null;
  if (styleSel) {
    styleSel.value = editor.isActive("heading", { level: 1 })
      ? "h1"
      : editor.isActive("heading", { level: 2 })
        ? "h2"
        : editor.isActive("heading", { level: 3 })
          ? "h3"
          : editor.isActive(BLOCK_NODE_NAMES.taskList)
            ? "taskList"
            : editor.isActive(BLOCK_NODE_NAMES.quote)
              ? "quote"
              : editor.isActive(BLOCK_NODE_NAMES.callout)
                ? "callout"
                : editor.isActive(BLOCK_NODE_NAMES.codeBlock)
                  ? "codeBlock"
                  : "paragraph";
  }

  const ts = textStyleAttrs(editor);
  const fontSel = $("tb-font") as HTMLSelectElement | null;
  const family = ts.fontFamily ? ts.fontFamily.replace(/^['"]+|['"]+$/g, "") : null;
  if (fontSel)
    fontSel.value = family && fontSel.querySelector(`option[value="${CSS.escape(family)}"]`) ? family : "";

  const sizeSel = $("tb-size") as HTMLSelectElement | null;
  if (sizeSel) sizeSel.value = ts.fontSize ?? "";

  const setActive = (id: string, active: boolean): void => {
    const el = $(id);
    el?.classList.toggle("active", active);
  };
  setActive("tb-bold", editor.isActive("bold"));
  setActive("tb-italic", editor.isActive("italic"));
  setActive("tb-underline", editor.isActive("underline"));
  setActive("tb-strike", editor.isActive("strike"));
  setActive("tb-code", editor.isActive("code"));

  const inLink = editor.isActive("link");
  setActive("tb-link", inLink);
  // Insert/edit is always available: disabling it when there is no link yet
  // would make the first link impossible to create (a disabled button never
  // fires its click handler). Only removing a link that is not there is a
  // no-op worth greying out.
  const unlinkEl = $("tb-unlink") as HTMLButtonElement | null;
  if (unlinkEl) unlinkEl.disabled = !inLink;

  for (const [id, align] of [
    ["tb-align-left", "left"],
    ["tb-align-center", "center"],
    ["tb-align-right", "right"],
    ["tb-align-justify", "justify"],
  ] as const) {
    setActive(id, editor.isActive({ textAlign: align }));
  }

  const spacingSel = $("tb-spacing") as HTMLSelectElement | null;
  const ls = Number(editor.getAttributes("paragraph").lineSpacing) || 1;
  if (spacingSel) spacingSel.value = String(ls);

  const inList = editor.isActive("bulletList") || editor.isActive("orderedList");
  setActive("tb-bullet", editor.isActive("bulletList"));
  setActive("tb-ordered", editor.isActive("orderedList"));
  for (const [id, canGo] of [
    ["tb-indent", editor.can().sinkListItem("listItem")],
    ["tb-outdent", editor.can().liftListItem("listItem")],
  ] as const) {
    const el = $(id) as HTMLButtonElement | null;
    if (el) el.disabled = !inList || !canGo;
  }

  const inTable = editor.isActive("table");
  for (const id of [
    "tb-table-row-add",
    "tb-table-row-del",
    "tb-table-col-add",
    "tb-table-col-del",
    "tb-table-header",
    "tb-table-merge",
    "tb-table-split",
    "tb-table-del",
  ]) {
    const el = $(id) as HTMLButtonElement | null;
    if (el) el.disabled = !inTable;
  }
  setActive("tb-table-header", editor.isActive("tableHeader"));
  // Merging needs a span of cells; splitting only makes sense on a merged cell.
  const mergeEl = $("tb-table-merge") as HTMLButtonElement | null;
  if (mergeEl) mergeEl.disabled = !inTable || !cellSelectionActive(editor);
  const splitEl = $("tb-table-split") as HTMLButtonElement | null;
  if (splitEl) splitEl.disabled = !inTable || !editor.can().splitCell();

  const zoomSel = $("tb-zoom") as HTMLSelectElement | null;
  if (zoomSel) {
    const zoom = getZoom();
    // Ctrl+= can leave the preset list (setZoom clamps 25–400); show the real
    // value instead of blanking the select when there is no matching option.
    if (!zoomSel.querySelector(`option[value="${zoom}"]`)) {
      const opt = document.createElement("option");
      opt.value = String(zoom);
      opt.textContent = `${zoom}%`;
      // Insert in place so the list stays ordered as the user steps past the
      // presets with Ctrl+= / Ctrl+-.
      const after = Array.from(zoomSel.querySelectorAll("option")).find((o) => Number(o.value) > zoom);
      zoomSel.insertBefore(opt, after ?? null);
    }
    zoomSel.value = String(zoom);
  }

  const colorBar = $("tb-color-bar") as HTMLElement | null;
  if (colorBar) colorBar.style.background = textStyleAttrs(editor).color ?? "transparent";

  const highlightBtn = $("tb-highlight") as HTMLElement | null;
  const hl = editor.getAttributes("highlight").color as string | null | undefined;
  if (highlightBtn) highlightBtn.style.background = hl ?? "transparent";
}
