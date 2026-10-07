import type { Editor, JSONContent } from "@tiptap/core";
import type { FwApi, PageMeta, PageSettings, WriteFileResult } from "../../shared/types";
import { PAGE_SIZES } from "../../shared/types";
import type { PageSize } from "../../shared/types";
import { createEditor, INITIAL_DOC, syncDocAttrStyles } from "./editor/create-editor";
import { runPagination, schedulePaginate } from "./editor/pagination";
import { createToolbar, syncToolbar } from "./features/toolbar";
import { insertImageAtCursor } from "./features/image-insert";
import { updateStatus, scheduleStatus } from "./features/statusbar";
import { openPageSetup } from "./features/page-setup";
import { setZoom, zoomStep } from "./features/zoom";
import { docxBufferToDoc } from "./io/import-docx";
import { htmlToDoc, textToDoc } from "./io/import-other";
import { docJson, toHtml, toText } from "./io/serialize";
import { alertDialog, confirmModal, promptModal } from "./features/modal";
import { installLinkClicks } from "./features/links";
import {
  beginPrintAppearance,
  endPrintAppearance,
  initAppearance,
  setTheme,
  toggleNotionMode,
  toggleSidebar,
  toggleToolbarPin,
} from "./features/appearance";
import { createSidebar } from "./features/sidebar";
import type { Sidebar } from "./features/sidebar";
import { openQuickSwitcher } from "./features/quick-switcher";

declare global {
  interface Window {
    fw?: FwApi;
  }
}

let editor: Editor | null = null;
let currentPath: string | null = null;
let docName = "Untitled";
let dirty = false;
let printMode = false;

// --- workspace page state ---------------------------------------------------
// A document has exactly one source: a picked file (today's flow, untouched)
// or a workspace page (autosaved — no Ctrl+S for it, the save *is* the edit).
let currentPageId: string | null = null;
let pagesCache: PageMeta[] = [];
let sidebar: Sidebar | null = null;
let pageSaveTimer: number | null = null;
/** Changes made since the page last hit disk (drives the quit flush). */
let pageUnsaved = false;

const fw = window.fw;

async function reloadPages(): Promise<void> {
  if (!fw || !sidebar) return;
  try {
    pagesCache = await fw.listPages();
  } catch {
    pagesCache = [];
  }
  sidebar.refresh(pagesCache);
}

/** Debounced autosave: a page reaches disk 600 ms after the last keystroke. */
function schedulePageSave(): void {
  if (currentPageId === null) return;
  if (pageSaveTimer !== null) window.clearTimeout(pageSaveTimer);
  pageSaveTimer = window.setTimeout(() => {
    pageSaveTimer = null;
    void flushPageSave();
  }, 600);
}

async function flushPageSave(): Promise<boolean> {
  if (pageSaveTimer !== null) {
    window.clearTimeout(pageSaveTimer);
    pageSaveTimer = null;
  }
  if (currentPageId === null || !editor || !fw || !pageUnsaved) return true;
  const ok = await fw
    .writePage(currentPageId, docJson(editor))
    .then((res) => !res || !res.error)
    .catch(() => false);
  if (ok) {
    pageUnsaved = false;
    markClean();
  } else {
    // Keep the document dirty so the quit prompt still protects the user.
    console.error("[FreeWord] page autosave failed");
  }
  return ok;
}

function isPageSource(): boolean {
  return currentPageId !== null;
}

async function openPage(id: string): Promise<void> {
  if (!editor || !fw) return;
  if (currentPageId === id) return;
  if (dirty) {
    const kept = isPageSource() ? await flushPageSave() : true;
    if (!kept || (!isPageSource() && !window.confirm("Discard unsaved changes and open another page?"))) {
      if (!isPageSource()) return;
    }
  }
  const page = await fw.readPage(id).catch(() => null);
  if (!page) {
    await alertDialog(`Could not open the page "${id}" — it may have been deleted.`);
    void reloadPages();
    return;
  }
  editor.commands.setContent(page.json as unknown as JSONContent, { emitUpdate: false });
  currentPageId = id;
  currentPath = null;
  docName = page.meta.title || "Untitled";
  pageUnsaved = false;
  markClean();
  syncDocAttrStyles(editor);
  runPagination(editor);
  editor.commands.focus("start");
  updateStatus(editor);
  syncToolbar(editor);
  // Refresh the sidebar first so the active highlight survives the re-render.
  await reloadPages();
  sidebar?.setActive(id);
}

async function createAndOpenPage(parentId: string | null): Promise<void> {
  if (!fw) return;
  const meta = await fw.createPage(parentId ? { parentId } : undefined).catch(() => null);
  if (!meta) {
    await alertDialog("Could not create the page.");
    return;
  }
  await reloadPages();
  await openPage(meta.id);
}

async function renameCurrentOrId(id: string): Promise<void> {
  if (!fw) return;
  const meta = pagesCache.find((p) => p.id === id);
  const title = await promptModal({
    title: "Rename page",
    label: "Title",
    value: meta?.title ?? "",
    placeholder: "Page title",
    okLabel: "Rename",
  });
  if (title === null || !title) return;
  await fw.renamePage(id, { title }).catch(() => null);
  if (currentPageId === id) {
    docName = title;
    updateTitle();
  }
  void reloadPages();
}

async function setIconFor(id: string): Promise<void> {
  if (!fw) return;
  const meta = pagesCache.find((p) => p.id === id);
  const icon = await promptModal({
    title: "Set page icon",
    label: "Emoji",
    value: meta?.icon ?? "",
    placeholder: "💡",
    okLabel: "Set",
    validate: (v) => ([...v].length <= 4 ? null : "One or two characters, please."),
  });
  if (icon === null) return;
  await fw.renamePage(id, { icon: icon || "📄" }).catch(() => null);
  void reloadPages();
}

/** The Ctrl+K modal over the page list; the sidebar header button opens it too. */
function openQuickSwitch(): void {
  if (!fw) return;
  void (async () => {
    await reloadPages();
    openQuickSwitcher(pagesCache, (id) => void openPage(id));
  })();
}

async function deletePageUI(id: string): Promise<void> {
  if (!fw) return;
  const meta = pagesCache.find((p) => p.id === id);
  const yes = await confirmModal({
    title: "Delete page",
    message: `Delete "${meta?.title ?? "this page"}" and its subpages? They are kept in the workspace trash.`,
    okLabel: "Delete",
    danger: true,
  });
  if (!yes) return;
  const res = await fw.deletePage(id).catch((err: unknown) => ({ error: String(err) }));
  if (res?.error) {
    await alertDialog(`Could not delete the page: ${res.error}`);
    return;
  }
  if (currentPageId === id) {
    // The deleted page must not haunt the editor: start a blank doc.
    currentPageId = null;
    pageUnsaved = false;
    sidebar?.setActive(null);
    editor?.commands.setContent(INITIAL_DOC, { emitUpdate: false });
    docName = "Untitled";
    markClean();
  }
  void reloadPages();
}

function updateTitle(): void {
  document.title = `${dirty ? "● " : ""}${docName} — FreeWord`;
}

function markDirty(): void {
  if (dirty) return;
  dirty = true;
  updateTitle();
}

function markClean(): void {
  // Always refresh the title, even when nothing was dirty: `openFile` adopts a
  // new document name while the document is clean, and an early return here
  // left the old name (and the old `● ` state) on the window title.
  dirty = false;
  updateTitle();
}

function onEditorUpdate(): void {
  markDirty();
  if (editor) {
    schedulePaginate(editor);
    scheduleStatus(editor);
    if (currentPageId !== null) {
      pageUnsaved = true;
      schedulePageSave();
    }
  }
}

function nextFrame(): Promise<void> {
  const { promise, resolve } = Promise.withResolvers<void>();
  requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
  return promise;
}

/**
 * Toggle print mode: hides chrome, zeroes canvas padding, forces zoom to 1,
 * and writes an @page rule reflecting the current document size BEFORE the
 * PDF is captured or the system print dialog is shown.
 */
function setPrintMode(on: boolean): void {
  if (printMode === on) return;
  printMode = on;
  document.body.classList.toggle("print-mode", on);
  // The print engine always gets the Word rendering: real fixed pages and a
  // settled pagination, even mid-session in Notion mode.
  if (on) beginPrintAppearance();
  let styleEl = document.getElementById("fw-print") as HTMLStyleElement | null;
  if (!on) {
    styleEl?.remove();
    endPrintAppearance();
    if (editor) runPagination(editor);
    return;
  }
  if (!styleEl) {
    styleEl = document.createElement("style");
    styleEl.id = "fw-print";
    document.head.appendChild(styleEl);
  }
  const attrs = (editor?.state.doc.attrs ?? {}) as { size?: string };
  const ps = PAGE_SIZES[(attrs.size ?? "a4") as PageSize] ?? PAGE_SIZES.a4;
  styleEl.textContent = [
    `@page { size: ${ps.w}mm ${ps.h}mm; margin: 0; }`,
    "#toolbar, #statusbar, #sidebar { display: none !important; }",
    "#canvas { background: white !important; padding: 0 !important; }",
    ".fw-doc { zoom: 1 !important; }",
    ".fw-page { margin: 0 !important; box-shadow: none !important; break-after: page; page-break-after: always; }",
    ".fw-page:last-child { break-after: auto; page-break-after: auto; }",
    "* { print-color-adjust: exact; -webkit-print-color-adjust: exact; }",
  ].join("\n");
}

function insertPageBreak(): void {
  if (!editor) return;
  editor.chain().focus().insertContentAt(editor.state.selection, { type: "pageBreak" }).run();
}

function insertTable(): void {
  if (!editor) return;
  editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
}

function applyPageSettings(s: PageSettings): void {
  if (!editor) return;
  const tr = editor.state.tr;
  // `setNodeAttribute(0, …)` addresses the node *at document offset 0* — the
  // first page — so the attrs used to land on a node type that does not define
  // them, and were silently dropped: Page Setup looked inert and both .docx
  // and PDF exports kept emitting the A4/25.4 mm defaults. Page size and
  // margins are attrs of the custom `doc` node, which `setDocAttribute` sets.
  tr.setDocAttribute("size", s.size);
  tr.setDocAttribute("mt", s.mt);
  tr.setDocAttribute("mr", s.mr);
  tr.setDocAttribute("mb", s.mb);
  tr.setDocAttribute("ml", s.ml);
  editor.view.dispatch(tr);
  syncDocAttrStyles(editor);
  runPagination(editor);
  markDirty();
  if (currentPageId !== null) {
    pageUnsaved = true;
    schedulePageSave();
  }
}

function openPageSetupModal(): void {
  if (!editor) return;
  const attrs = editor.state.doc.attrs as PageSettings;
  openPageSetup(
    {
      size: attrs.size ?? "a4",
      mt: attrs.mt ?? 25.4,
      mr: attrs.mr ?? 25.4,
      mb: attrs.mb ?? 25.4,
      ml: attrs.ml ?? 25.4,
    },
    applyPageSettings,
  );
}

async function newDoc(): Promise<void> {
  if (!editor) return;
  if (dirty && isPageSource()) {
    await flushPageSave();
  } else if (dirty && !window.confirm("Discard unsaved changes and start a new document?")) {
    return;
  }
  editor.commands.setContent(INITIAL_DOC, { emitUpdate: false });
  currentPath = null;
  currentPageId = null;
  pageUnsaved = false;
  sidebar?.setActive(null);
  docName = "Untitled";
  markClean();
  syncDocAttrStyles(editor);
  runPagination(editor);
  // Open at the top, the way Word does: setContent leaves the caret wherever
  // the mapping puts it, and a re-pagination below maps it by page — which can
  // otherwise land it on the last page of a freshly created document.
  editor.commands.focus("start");
  updateStatus(editor);
  syncToolbar(editor);
}

async function openFile(path?: string): Promise<void> {
  if (!editor || !fw) return;
  if (dirty && isPageSource()) {
    // The page autosaves — flush it, then the file takes over the editor.
    await flushPageSave();
  } else if (dirty && !window.confirm("Discard unsaved changes and open another document?")) {
    return;
  }
  const res = await fw.openFile(path);
  if (!res) return;
  try {
    const json =
      res.ext === "docx"
        ? await docxBufferToDoc(res.buffer, editor.state.schema)
        : res.ext === "txt"
          ? textToDoc(new TextDecoder().decode(res.buffer))
          : htmlToDoc(new TextDecoder().decode(res.buffer), editor.state.schema);
    editor.commands.setContent(json as unknown as JSONContent, { emitUpdate: false });
  } catch (err) {
    // An unreadable file must be visible without blocking the renderer the way
    // `window.alert` does (that call freezes the whole page, quit handshake included).
    void alertDialog(`Could not open "${res.name}": ${err instanceof Error ? err.message : String(err)}`);
    return;
  }
  currentPath = res.path;
  currentPageId = null;
  pageUnsaved = false;
  sidebar?.setActive(null);
  docName = res.name.replace(/\.(docx|txt|html|htm)$/i, "");
  markClean();
  syncDocAttrStyles(editor);
  runPagination(editor);
  // See newDoc: the caret belongs at the top of a document the user just opened.
  editor.commands.focus("start");
  updateStatus(editor);
  syncToolbar(editor);
}

type SaveOutcome = "saved" | "cancelled" | "failed";

/** Adopt a path the main process just wrote the .docx to: track it, rename, clear dirty. */
function commitSavedPath(savedPath: string): void {
  currentPath = savedPath;
  docName =
    savedPath
      .split(/[\\/]/)
      .pop()
      ?.replace(/\.(docx|txt|html|htm)$/i, "") ?? docName;
  markClean();
}

/** Show the native error the main process already reported, and log for the console. */
function reportFailure(what: string, err: unknown): void {
  const message = err instanceof Error ? err.message : String(err);
  console.error(`[FreeWord] ${what}: ${message}`);
  void alertDialog(`${what}: ${message}`);
}

/**
 * Serialize + hand the document to the main process.
 *
 * Every failure mode is reported to the user and returned as a status: the
 * call used to reject (or return an error) with nobody watching, which left
 * Ctrl+S doing nothing and the quit handshake waiting for a response that
 * never came.
 */
async function doSave(path: string | undefined): Promise<SaveOutcome> {
  if (!editor || !fw) return "cancelled";
  let res: WriteFileResult | null;
  try {
    res = await fw.saveDoc({ json: docJson(editor), defaultName: `${docName}.docx`, path });
  } catch (err) {
    reportFailure(`Could not save "${docName}"`, err);
    return "failed";
  }
  // Dialog dismissed by the user — not an error, keep the document dirty.
  if (!res) return "cancelled";
  if (res.error) {
    // The main process already surfaced a native error dialog for this.
    console.error(`[FreeWord] save failed: ${res.error}`);
    return "failed";
  }
  if (!res.path) return "cancelled";
  commitSavedPath(res.path);
  updateTitle();
  return "saved";
}

async function saveDoc(): Promise<SaveOutcome> {
  // A workspace page has no file identity: Ctrl+S is an immediate autosave
  // flush, not a Save-As. (Export to .docx stays available from the File menu.)
  if (isPageSource()) {
    const ok = await flushPageSave();
    return ok ? "saved" : "failed";
  }
  return doSave(currentPath ?? undefined);
}

/**
 * Export a copy in a plain-text format. This is an export, not a save: the
 * document keeps its .docx identity (and its dirty flag), because the file
 * just written does not contain the document — only its text.
 */
async function exportFormat(ext: "txt" | "html"): Promise<SaveOutcome> {
  if (!editor || !fw) return "cancelled";
  const content = ext === "txt" ? toText(editor) : toHtml(editor);
  let res: WriteFileResult | null;
  try {
    res = await fw.saveRaw({ content, defaultName: `${docName}.${ext}`, ext });
  } catch (err) {
    reportFailure(`Could not export as .${ext}`, err);
    return "failed";
  }
  if (!res) return "cancelled";
  if (res.error) {
    console.error(`[FreeWord] export .${ext} failed: ${res.error}`);
    return "failed";
  }
  return res.path ? "saved" : "cancelled";
}

async function exportPdf(): Promise<void> {
  if (!editor || !fw) return;
  setPrintMode(true);
  await nextFrame();
  try {
    const res = await fw.exportPdf(`${docName}.pdf`);
    if (res?.error) console.error(`[FreeWord] PDF export failed: ${res.error}`);
  } catch (err) {
    // A rejection here used to vanish into the void: `void exportPdf()` meant
    // choosing a path and getting nothing at all.
    reportFailure(`Could not export "${docName}.pdf"`, err);
  } finally {
    setPrintMode(false);
  }
}

async function printDoc(): Promise<void> {
  if (!editor || !fw) return;
  setPrintMode(true);
  await nextFrame();
  try {
    const res = await fw.print();
    if (res?.error) console.error(`[FreeWord] print failed: ${res.error}`);
  } catch (err) {
    reportFailure("Could not print", err);
  } finally {
    setPrintMode(false);
  }
}

export function initApp(): void {
  editor = createEditor(onEditorUpdate);
  createToolbar(editor);
  // Ctrl/Cmd+click on a link opens it in the desktop browser through the
  // main process; a plain click only places the caret so the link can be edited.
  installLinkClicks(editor);
  void initAppearance(editor, fw ?? null).then(() => {
    // Settings arrive asynchronously; the chrome that depends on the mode
    // (status-bar label, toolbar mirrors) was rendered for the default.
    if (editor && !editor.isDestroyed) {
      updateStatus(editor);
      syncToolbar(editor);
    }
  });
  updateStatus(editor);
  updateTitle();

  // --- workspace sidebar -------------------------------------------------
  sidebar = createSidebar({
    onOpenPage: (id) => void openPage(id),
    onNewPage: (parentId) => void createAndOpenPage(parentId),
    onRename: (id) => void renameCurrentOrId(id),
    onSetIcon: (id) => void setIconFor(id),
    onDelete: (id) => void deletePageUI(id),
    onMove: (id, parentId, order) => {
      void (async () => {
        const res = await fw?.movePage(id, parentId, order);
        if (res?.error) await alertDialog(`Could not move the page: ${res.error}`);
        await reloadPages();
      })();
    },
    onQuickSwitcher: () => openQuickSwitch(),
  });
  if (fw) {
    void reloadPages();
    // Losing focus is a good moment to put the debounce window's keystrokes on
    // disk — an alt-tab into a crash should not cost more than that.
    window.addEventListener("blur", () => {
      void flushPageSave();
    });
    // Ctrl/Cmd+K — Notion's quick switcher. The link command lost the tooltip
    // claim on this chord; nothing else binds it.
    document.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === "k") {
        e.preventDefault();
        openQuickSwitch();
      }
    });
  }

  document.addEventListener("fw:page-setup-request", openPageSetupModal);
  document.fonts.ready.then(() => {
    if (editor) runPagination(editor);
  });

  // Keep the toolbar/status bar fresh when zoom changes.
  window.addEventListener("fw:zoom-changed", () => {
    if (editor) {
      syncToolbar(editor);
      updateStatus(editor);
    }
  });

  // A re-pagination is layout-only (it sets `preventUpdate`), so the chrome
  // that reads the page structure — `Page X / Y` above all — has to be
  // refreshed here rather than from an `update` event.
  window.addEventListener("fw:layout-changed", () => {
    if (editor) {
      syncToolbar(editor);
      updateStatus(editor);
    }
  });

  if (!fw) return; // Standalone browser mode (dev smoke) — no IPC available.

  fw.onCommand(({ id, path }) => {
    const ed = editor;
    switch (id) {
      case "file:new":
        void newDoc();
        break;
      case "file:open":
        void openFile();
        break;
      case "recent:open":
        if (path) void openFile(path);
        break;
      case "file:save":
        void saveDoc();
        break;
      case "file:saveAs":
        void doSave(undefined);
        break;
      case "file:exportTxt":
        void exportFormat("txt");
        break;
      case "file:exportHtml":
        void exportFormat("html");
        break;
      case "file:exportPdf":
        void exportPdf();
        break;
      case "file:print":
        void printDoc();
        break;
      case "insert:pageBreak":
        insertPageBreak();
        break;
      case "insert:table":
        insertTable();
        break;
      case "insert:image":
        if (ed) void insertImageAtCursor(ed);
        break;
      case "view:zoomIn":
        if (ed) zoomStep(1, ed);
        break;
      case "view:zoomOut":
        if (ed) zoomStep(-1, ed);
        break;
      case "view:zoomReset":
        if (ed) setZoom(100, ed);
        break;
      case "view:pageSetup":
        openPageSetupModal();
        break;
      case "view:notionMode":
        void toggleNotionMode();
        break;
      case "view:toggleToolbarPin":
        void toggleToolbarPin();
        break;
      case "view:toggleSidebar":
        void toggleSidebar();
        break;
      case "view:quickSwitcher":
        openQuickSwitch();
        break;
      case "page:new":
        void createAndOpenPage(null);
        break;
      case "page:open":
        // Reuses the MenuCommand path field to carry the page id (the only
        // free-form string the menu already sends).
        if (path) void openPage(path);
        break;
      default:
        if (typeof id === "string" && id.startsWith("view:theme:")) {
          const t = id.slice("view:theme:".length);
          if (t === "light" || t === "dark" || t === "system") void setTheme(t);
        }
    }
    window.dispatchEvent(new Event("fw:zoom-changed"));
  });

  fw.onWillQuit(() => {
    // A page document's autosave lives in a debounce window; flush it first
    // so the answer main receives reflects what actually reached disk.
    void (async () => {
      try {
        await flushPageSave();
      } finally {
        fw.sendQuitResponse({ dirty, name: docName, path: currentPath });
      }
    })();
  });

  fw.onSaveThenQuit(() => {
    void (async () => {
      try {
        await saveDoc(); // routes to the page flush for page sources
      } finally {
        // Must run even if the save threw or was cancelled: main waits on this
        // reply, and without it the app looks frozen on quit.
        fw.sendQuitResponse({ dirty, name: docName, path: currentPath });
      }
    })();
  });
}
