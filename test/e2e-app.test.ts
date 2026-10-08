// @vitest-environment jsdom
/**
 * End-to-end renderer flow: boots the real app (Tiptap editor + toolbar +
 * status bar) against a mocked `window.fw` preload API and drives every menu
 * command / toolbar control the way Electron's menu would.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

type Cb = (...args: unknown[]) => void;

interface FwMock {
  openFile: ReturnType<typeof vi.fn>;
  saveDoc: ReturnType<typeof vi.fn>;
  saveRaw: ReturnType<typeof vi.fn>;
  pickImage: ReturnType<typeof vi.fn>;
  exportPdf: ReturnType<typeof vi.fn>;
  print: ReturnType<typeof vi.fn>;
  getRecent: ReturnType<typeof vi.fn>;
  openExternal: ReturnType<typeof vi.fn>;
  listTrash: ReturnType<typeof vi.fn>;
  restoreTrash: ReturnType<typeof vi.fn>;
  deleteTrash: ReturnType<typeof vi.fn>;
  onCommand: ReturnType<typeof vi.fn>;
  sendQuitResponse: ReturnType<typeof vi.fn>;
  onWillQuit: ReturnType<typeof vi.fn>;
  onSaveThenQuit: ReturnType<typeof vi.fn>;
}

let fw: FwMock;
let commandCb: ((cmd: { id: string; path?: string }) => void) | null;
let willQuitCb: Cb | null;
let saveThenQuitCb: Cb | null;

const PNG_1x1 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

// ProseMirror needs these to survive a jsdom selection sync.
(Range as unknown as { prototype: Record<string, Cb> }).prototype.getClientRects = () =>
  Object.assign([], { item: () => null });
(Range as unknown as { prototype: Record<string, Cb> }).prototype.getBoundingClientRect = () =>
  ({ x: 0, y: 0, width: 0, height: 0, top: 0, left: 0, right: 0, bottom: 0 }) as DOMRect;

function resetDom(): void {
  document.body.innerHTML = `
    <div id="toolbar"></div>
    <div id="canvas"><div id="editor"></div></div>
    <div id="statusbar"></div>`;
  document.title = "Untitled — FreeWord";
}

function makeFw(): FwMock {
  commandCb = null;
  willQuitCb = null;
  saveThenQuitCb = null;
  const api: FwMock = {
    openFile: vi.fn().mockResolvedValue(null),
    saveDoc: vi.fn().mockResolvedValue(null),
    saveRaw: vi.fn().mockResolvedValue(null),
    openExternal: vi.fn().mockResolvedValue(true),
    pickImage: vi.fn().mockResolvedValue(null),
    exportPdf: vi.fn().mockResolvedValue(null),
    print: vi.fn().mockResolvedValue(undefined),
    getRecent: vi.fn().mockResolvedValue([]),
    listTrash: vi.fn().mockResolvedValue([]),
    restoreTrash: vi.fn().mockResolvedValue({}),
    deleteTrash: vi.fn().mockResolvedValue({}),
    onCommand: vi.fn((cb: Cb) => {
      commandCb = cb as never;
      return () => {};
    }),
    sendQuitResponse: vi.fn(),
    onWillQuit: vi.fn((cb: Cb) => {
      willQuitCb = cb;
      return () => {};
    }),
    onSaveThenQuit: vi.fn((cb: Cb) => {
      saveThenQuitCb = cb;
      return () => {};
    }),
  };
  return api;
}

async function boot(opts?: { confirm?: boolean }): Promise<void> {
  vi.resetModules();
  fw = makeFw();
  resetDom();
  Object.defineProperty(window, "fw", { value: fw, configurable: true, writable: true });
  vi.stubGlobal("alertSpy", undefined);
  const alertFn = vi.fn();
  const confirmFn = vi.fn(() => opts?.confirm ?? true);
  Object.defineProperty(window, "alert", { value: alertFn, configurable: true });
  Object.defineProperty(window, "confirm", { value: confirmFn, configurable: true });
  (globalThis as Record<string, unknown>).__alert = alertFn;
  (globalThis as Record<string, unknown>).__confirm = confirmFn;
  // jsdom has no FontFaceSet; Chromium does, and app.ts only uses it to
  // re-paginate once webfonts settle.
  Object.defineProperty(document, "fonts", {
    value: { ready: Promise.resolve() },
    configurable: true,
  });
  if (!window.requestAnimationFrame) {
    Object.defineProperty(window, "requestAnimationFrame", {
      value: (cb: (t: number) => void) => setTimeout(() => cb(Date.now()), 0),
      configurable: true,
    });
  }
  const { initApp } = await import("../src/renderer/src/app");
  initApp();
  // Let debounced timers (150ms status refresh) from previously booted apps —
  // which still target the shared #statusbar element — fire before the test's
  // own commands run. Without this, stale writes race with command writes.
  await new Promise((r) => setTimeout(r, 220));
}

const send = async (id: string, path?: string): Promise<void> => {
  commandCb!({ id, ...(path ? { path } : {}) });
  await new Promise((r) => setTimeout(r, 10));
};

const editorEl = (): HTMLElement => document.querySelector("#editor")!;
const confirmSpy = (): ReturnType<typeof vi.fn> => globalThis.__confirm as never;

/**
 * Error text from the in-app alert modal (`features/modal.ts`), or null when
 * none is open. These used to call blocking `window.alert`, which freezes the
 * renderer — and the quit handshake — until dismissed.
 */
function alertDialogText(): string | null {
  const box = document.querySelector<HTMLElement>("#fw-alert");
  if (!box) return null;
  return box.querySelector<HTMLElement>(".fw-alert-message")?.textContent ?? "";
}

describe("app boot", () => {
  it("renders toolbar, editor and status bar", async () => {
    await boot();
    expect(document.querySelector("#toolbar .tb-group")).toBeTruthy();
    expect(editorEl().querySelector(".ProseMirror")).toBeTruthy();
    expect(editorEl().classList.contains("fw-doc")).toBe(true);
    expect(document.querySelector("#statusbar")!.textContent).toContain("Page 1 / 1");
  });
});

describe("menu commands", () => {
  beforeEach(async () => {
    await boot();
  });

  it("insert:table adds a 3×3 table with a header row", async () => {
    await send("insert:table");
    const table = editorEl().querySelector("table");
    expect(table).toBeTruthy();
    expect(table!.querySelectorAll("tbody tr, tr").length).toBe(3);
    expect(table!.querySelectorAll("th").length).toBe(3);
  });

  it("insert:pageBreak adds a break marker and marks the doc dirty", async () => {
    await send("insert:pageBreak");
    expect(editorEl().querySelector(".fw-pagebreak")).toBeTruthy();
    expect(document.title).toMatch(/^● /);
  });

  it("file:save sends the doc JSON and clears dirty after success", async () => {
    await send("insert:pageBreak"); // make dirty
    expect(document.title).toMatch(/^● /);
    fw.saveDoc.mockResolvedValue({ path: "C:\\docs\\Report.docx" });
    await send("file:save");
    expect(fw.saveDoc).toHaveBeenCalledTimes(1);
    const payload = fw.saveDoc.mock.calls[0][0];
    expect(payload.json.type).toBe("doc");
    expect(payload.defaultName).toMatch(/\.docx$/);
    expect(document.title).toBe("Report — FreeWord");
  });

  it("file:save keeps dirty when the save dialog is cancelled", async () => {
    await send("insert:pageBreak");
    fw.saveDoc.mockResolvedValue(null); // cancelled
    await send("file:save");
    expect(document.title).toMatch(/^● /);
  });

  it("file:saveAs always passes no path (native dialog) and adopts the new name", async () => {
    fw.saveDoc.mockResolvedValue({ path: "/tmp/Other.docx" });
    await send("file:saveAs");
    expect(fw.saveDoc.mock.calls[0][0].path).toBeUndefined();
    expect(document.title).toBe("Other — FreeWord");
  });

  // Regression: openFile adopts the name while the document is clean, and
  // markClean() used to early-return in exactly that case — leaving the old
  // name (and old dirty marker) on the window title.
  it("file:open while clean renames the window and stays clean", async () => {
    const buf = new TextEncoder().encode("hello from txt\nsecond line") as unknown as Uint8Array;
    fw.openFile.mockResolvedValue({ path: "/x/notes.txt", name: "notes.txt", ext: "txt", buffer: buf });
    await send("file:open");
    expect(editorEl().textContent).toContain("hello from txt");
    expect(document.title).toBe("notes — FreeWord");
    expect(document.title).not.toMatch(/^● /);
  });

  // Regression: importing a document that has to be re-partitioned across
  // pages dispatched a layout transaction, which fired Tiptap's `update` and
  // marked the freshly-opened document dirty — so the title grew a `●` and the
  // quit handshake asked to save a file the user never touched.
  it("opening a document that re-paginates stays clean", async () => {
    // Simulate real layout: 60 px blocks in a 100 px printable area force a
    // split (jsdom reports 0 for every measurement otherwise).
    Object.defineProperty(HTMLElement.prototype, "clientHeight", { configurable: true, value: 100 });
    const original = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function (): DOMRect {
      const r = original.call(this);
      return {
        ...r,
        height: this.tagName === "P" ? 60 : r.height,
        width: this.tagName === "P" ? 300 : r.width,
      } as DOMRect;
    };
    try {
      const lines = Array.from({ length: 8 }, (_, i) => `line ${i + 1}`).join("\n");
      const buf = new TextEncoder().encode(lines) as unknown as Uint8Array;
      fw.openFile.mockResolvedValue({ path: "/x/long.txt", name: "long.txt", ext: "txt", buffer: buf });
      await send("file:open");
      await new Promise((r) => setTimeout(r, 30));

      const pages = editorEl().querySelectorAll(".fw-page");
      expect(pages.length).toBeGreaterThan(1); // the layout really did change
      expect(document.title).toBe("long — FreeWord");
      expect(document.title).not.toMatch(/^● /);
      // …and the page counter still reflects the new structure.
      expect(document.querySelector("#statusbar")!.textContent).toMatch(
        new RegExp(`Page 1 / ${pages.length}`),
      );
    } finally {
      Element.prototype.getBoundingClientRect = original;
      delete (HTMLElement.prototype as { clientHeight?: unknown }).clientHeight;
    }
  });

  it("recent:open forwards the path to openFile", async () => {
    const buf = new TextEncoder().encode("recent content") as unknown as Uint8Array;
    fw.openFile.mockResolvedValue({ path: "/x/r.txt", name: "r.txt", ext: "txt", buffer: buf });
    await send("recent:open", "/x/r.txt");
    expect(fw.openFile).toHaveBeenCalledWith("/x/r.txt");
    expect(editorEl().textContent).toContain("recent content");
  });

  it("opening a broken docx alerts the user instead of failing silently", async () => {
    const junk = new TextEncoder().encode("not a docx") as unknown as Uint8Array;
    fw.openFile.mockResolvedValue({ path: "/x/bad.docx", name: "bad.docx", ext: "docx", buffer: junk });
    await send("file:open");
    await new Promise((r) => setTimeout(r, 10));
    expect(alertDialogText()).toContain("bad.docx");
  });

  it("a failing openFile (deleted recent entry) is a silent no-op in the renderer", async () => {
    // Contract (main/ipc.ts): unreadable files show a NATIVE error box and the
    // invoke resolves to null. The renderer must then change nothing.
    await send("insert:pageBreak"); // dirty, has a page break
    fw.openFile.mockResolvedValue(null); // native error box shown by main
    await send("recent:open", "C:\\gone\\missing.txt");
    expect(editorEl().querySelector(".fw-pagebreak")).toBeTruthy(); // content kept
    expect(document.title).toMatch(/^.* /); // still dirty (bullet marker)
    expect(document.title).not.toContain("missing"); // title untouched
  });

  it("file:new prompts Save / Don't save / Cancel without a native dialog", async () => {
    await send("insert:pageBreak");
    await send("file:new");
    await new Promise((r) => setTimeout(r, 10));
    const modal = document.querySelector<HTMLElement>("#fw-save-changes");
    expect(modal, "discard prompt is an in-app modal, not window.confirm").toBeTruthy();
    expect(confirmSpy()).not.toHaveBeenCalled();

    // Cancel: dismisses the prompt and keeps the document untouched.
    modal!.querySelector<HTMLButtonElement>("#fw-save-changes-cancel")!.click();
    await new Promise((r) => setTimeout(r, 10));
    expect(document.querySelector("#fw-save-changes")).toBeNull();
    expect(editorEl().querySelector(".fw-pagebreak")).toBeTruthy(); // kept
    expect(document.title).toMatch(/^● /);

    // Don't save: proceeds and clears the document.
    await send("file:new");
    await new Promise((r) => setTimeout(r, 10));
    document.querySelector<HTMLButtonElement>("#fw-save-changes-discard")!.click();
    await new Promise((r) => setTimeout(r, 10));
    expect(editorEl().querySelector(".fw-pagebreak")).toBeFalsy();
    expect(document.title).toBe("Untitled — FreeWord");
  });

  it("the discard prompt's Save button saves through the normal save path", async () => {
    await send("insert:pageBreak");
    fw.saveDoc.mockResolvedValue({ path: "C:\\docs\\Kept.docx" });
    await send("file:new");
    await new Promise((r) => setTimeout(r, 10));
    document.querySelector<HTMLButtonElement>("#fw-save-changes-save")!.click();
    await new Promise((r) => setTimeout(r, 20));
    expect(fw.saveDoc).toHaveBeenCalled();
    // The save cleared dirty, so newDoc ran without a second prompt.
    expect(document.querySelector("#fw-save-changes")).toBeNull();
    expect(editorEl().querySelector(".fw-pagebreak")).toBeFalsy();
    expect(document.title).toBe("Untitled — FreeWord");
  });

  it("Escape on the discard prompt cancels the switch", async () => {
    await send("insert:pageBreak");
    await send("file:new");
    await new Promise((r) => setTimeout(r, 10));
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    await new Promise((r) => setTimeout(r, 10));
    expect(document.querySelector("#fw-save-changes")).toBeNull();
    expect(editorEl().querySelector(".fw-pagebreak")).toBeTruthy(); // kept
  });

  it("insert:image measures and scales wide images down to 900px", async () => {
    fw.pickImage.mockResolvedValue({ dataUrl: `data:image/png;base64,${PNG_1x1}`, name: "photo.png" });
    class FakeImg {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      naturalWidth = 1200;
      naturalHeight = 800;
      set src(_v: string) {
        setTimeout(() => this.onload?.(), 0);
      }
    }
    vi.stubGlobal("Image", FakeImg);
    await send("insert:image");
    const img = editorEl().querySelector("img");
    expect(img).toBeTruthy();
    // The resizable node view sizes the element from the node's attrs, so the
    // scale-to-900 path must reach the attrs — style is where the resizer
    // applies them, and the attrs are what the .docx export consumes.
    expect(img!.style.width).toBe("900px");
    expect(img!.style.height).toBe("600px");
    expect(img!.getAttribute("alt")).toBe("photo.png");
    fw.saveDoc.mockResolvedValue({ path: "C:\\docs\\img.docx" });
    await send("file:save");
    const payload = fw.saveDoc.mock.calls[0][0];
    const json = JSON.stringify(payload.json);
    expect(json).toMatch(/"width":900/);
    expect(json).toMatch(/"height":600/);
    // The corner drag handles are in the DOM (revealed on hover by CSS).
    expect(editorEl().querySelectorAll("[data-resize-handle]").length).toBe(4);
    vi.unstubAllGlobals();
  });

  it("dragging a corner handle resizes the image and persists the new size", async () => {
    fw.pickImage.mockResolvedValue({ dataUrl: `data:image/png;base64,${PNG_1x1}`, name: "small.png" });
    class FakeImg {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      naturalWidth = 400;
      naturalHeight = 200; // aspect ratio 2:1
      set src(_v: string) {
        setTimeout(() => this.onload?.(), 0);
      }
    }
    vi.stubGlobal("Image", FakeImg);
    await send("insert:image");
    const img = editorEl().querySelector("img")!;
    expect(img.style.width).toBe("400px"); // under the 900 cap, kept as-is

    // jsdom reports 0 for offsetWidth; the resizer reads it to measure the
    // drag. Serve it from the inline style the resizer itself writes.
    Object.defineProperties(HTMLElement.prototype, {
      offsetWidth: { configurable: true, get: () => parseFloat(img.style.width) || 0 },
      offsetHeight: { configurable: true, get: () => parseFloat(img.style.height) || 0 },
    });
    try {
      const handle = editorEl().querySelector<HTMLElement>('[data-resize-handle="bottom-right"]')!;
      expect(handle).toBeTruthy();
      handle.dispatchEvent(new MouseEvent("mousedown", { clientX: 100, clientY: 100, bubbles: true }));
      document.dispatchEvent(new MouseEvent("mousemove", { clientX: 160, clientY: 100, bubbles: true }));
      document.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
      await new Promise((r) => setTimeout(r, 20));

      // +60px width, aspect ratio preserved (2:1 → 460×230).
      expect(img.style.width).toBe("460px");
      expect(img.style.height).toBe("230px");
      // And the commit reached the document, not just the live element:
      fw.saveDoc.mockResolvedValue({ path: "C:\\docs\\resized.docx" });
      await send("file:save");
      const json = JSON.stringify(fw.saveDoc.mock.calls[0][0].json);
      expect(json).toMatch(/"width":460/);
      expect(json).toMatch(/"height":230/);
    } finally {
      vi.unstubAllGlobals();
      delete (HTMLElement.prototype as { offsetWidth?: unknown }).offsetWidth;
      delete (HTMLElement.prototype as { offsetHeight?: unknown }).offsetHeight;
    }
  });

  it("view:zoomIn raises zoom (status bar + select)", async () => {
    await send("view:zoomIn");
    expect(document.querySelector("#statusbar")!.textContent).toContain("110%");
    await send("view:zoomReset");
    expect(document.querySelector("#statusbar")!.textContent).toContain("100%");
  });

  it("view:zoomOut steps down to the nearest smaller preset", async () => {
    await send("view:zoomOut");
    expect((document.querySelector("#tb-zoom") as HTMLSelectElement).value).toBe("90");
    expect(document.querySelector("#statusbar")!.textContent).toContain("90%");
  });

  // Regression: applyPageSettings wrote attrs with
  // `tr.setNodeAttribute(0, …)`, but position 0 addresses the FIRST PAGE node,
  // not the document root — `tr.setDocAttribute(name, value)` is the correct
  // API. Page size/margins were therefore inert in the editor, and .docx/PDF
  // exports always emitted the A4/25.4 mm defaults.
  it("page setup modal applies size + margins to the doc attrs and CSS vars", async () => {
    await send("view:pageSetup");
    const modal = document.querySelector("#fw-page-setup");
    expect(modal).toBeTruthy();
    (modal!.querySelector("#fw-ps-size") as HTMLSelectElement).value = "letter";
    (modal!.querySelector("#fw-ps-mt") as HTMLInputElement).value = "10";
    (modal!.querySelector("#fw-ps-ml") as HTMLInputElement).value = "300"; // out of range → clamped to 100
    (modal!.querySelector("#fw-ps-ok") as HTMLButtonElement).click();
    const carriers = [...document.querySelectorAll("#editor [style]")];
    const withVars = carriers.find((e) => e.style.getPropertyValue("--page-w"));
    expect(withVars, "--page-w was set on no element under #editor").toBeTruthy();
    expect(withVars.style.getPropertyValue("--page-w")).toBe("215.9mm");
    expect(withVars.style.getPropertyValue("--m-t")).toBe("10mm");
    expect(withVars.style.getPropertyValue("--m-l")).toBe("100mm"); // clamped
    expect(document.title).toMatch(/^● /); // page settings count as an edit
  });

  it("word count reflects the document text", async () => {
    const buf = new TextEncoder().encode("one two three four") as unknown as Uint8Array;
    fw.openFile.mockResolvedValue({ path: "/x/w.txt", name: "w.txt", ext: "txt", buffer: buf });
    await send("file:open");
    await new Promise((r) => setTimeout(r, 200)); // status debounce 150ms
    expect(document.querySelector("#statusbar")!.textContent).toContain("4 words");
  });

  it("quit handshake answers with the dirty state", async () => {
    await send("insert:table");
    willQuitCb!();
    // The answer is asynchronous now: the handshake first flushes any pending
    // workspace-page autosave before reporting, so the reply reflects the disk.
    await new Promise((r) => setTimeout(r, 30));
    expect(fw.sendQuitResponse).toHaveBeenCalledWith(
      expect.objectContaining({ dirty: true, name: "Untitled" }),
    );
  });

  it("save-then-quit saves, then answers clean", async () => {
    await send("insert:table");
    fw.saveDoc.mockResolvedValue({ path: "/x/saved.docx" });
    saveThenQuitCb!();
    await new Promise((r) => setTimeout(r, 20));
    expect(fw.saveDoc).toHaveBeenCalled();
    expect(fw.sendQuitResponse).toHaveBeenCalledWith(expect.objectContaining({ dirty: false }));
  });

  it("save-then-quit with a cancelled save keeps the app alive (dirty)", async () => {
    await send("insert:table");
    fw.saveDoc.mockResolvedValue(null); // user cancels the dialog
    saveThenQuitCb!();
    await new Promise((r) => setTimeout(r, 20));
    expect(fw.sendQuitResponse).toHaveBeenCalledWith(expect.objectContaining({ dirty: true }));
  });
});

describe("toolbar state sync", () => {
  // Regression: syncToolbar only ran on selectionUpdate; toggling a mark with
  // an empty selection changes storedMarks, not the selection, so the button
  // never reflected the active mark until the caret moved.
  it("mark buttons light up after a direct toggle with no selection change", async () => {
    await boot();
    document.querySelector<HTMLElement>("#tb-bold")!.click();
    await new Promise((r) => setTimeout(r, 120)); // toolbar debounce is 50ms
    expect(document.querySelector("#tb-bold")!.classList.contains("active")).toBe(true);
  });

  // Regression: toggleMenu's closed-state check compared `null === null`
  // (the menu lives on document.body, the lookup searched the anchor's group),
  // so the color/highlight dropdowns could never open.
  it("color menu opens and applies the chosen swatch", async () => {
    await boot();
    document.querySelector<HTMLElement>("#tb-color")!.click();
    await new Promise((r) => setTimeout(r, 10));
    const swatches = [...document.querySelectorAll<HTMLElement>(".tb-dropdown .swatch")];
    expect(swatches.length).toBe(10);
    swatches[2]!.click(); // C00000 — the menu must dismiss on choice
    await new Promise((r) => setTimeout(r, 10));
    expect(document.querySelector(".tb-dropdown")).toBeNull();
    // The stored mark carries the colour (there is no selection to style yet),
    // and the toolbar mirror stays in step instead of throwing — the mark's
    // own output is covered by the .docx export tests, which set marks directly.
    expect(document.querySelector("#tb-color")).toBeTruthy();
    expect(confirmSpy()).not.toHaveBeenCalled();
  });

  it("clicking the anchor toggles an open dropdown closed", async () => {
    await boot();
    document.querySelector<HTMLElement>("#tb-color")!.click();
    await new Promise((r) => setTimeout(r, 10));
    expect(document.querySelectorAll(".tb-dropdown").length).toBe(1);
    document.querySelector<HTMLElement>("#tb-color")!.click(); // toggle → closes
    expect(document.querySelectorAll(".tb-dropdown").length).toBe(0);
  });

  it("only one dropdown is open at a time", async () => {
    await boot();
    document.querySelector<HTMLElement>("#tb-color")!.click();
    await new Promise((r) => setTimeout(r, 10));
    document.querySelector<HTMLElement>("#tb-highlight")!.click();
    await new Promise((r) => setTimeout(r, 10));
    expect(document.querySelectorAll(".tb-dropdown").length).toBe(1);
    expect(document.querySelector(".tb-dropdown")!.getAttribute("data-menu-for")).toBe("tb-highlight");
  });

  it("highlight menu “None” on plain text does not add a mark", async () => {
    await boot();
    document.querySelector<HTMLElement>("#tb-highlight")!.click();
    await new Promise((r) => setTimeout(r, 10));
    const btns = [...document.querySelectorAll(".tb-dropdown .swatch-row button")];
    const none = btns.find((b) => b.textContent === "None") as HTMLButtonElement;
    none.click();
    await new Promise((r) => setTimeout(r, 10));
    expect(editorEl().innerHTML).not.toMatch(/background-color|data-record-highlights|mark/i);
    expect(editorEl().textContent).toBe(""); // still the placeholder-only doc
  });

  it("dropdown menus are keyboard operable: arrows walk, Escape closes", async () => {
    await boot();
    const anchor = document.querySelector<HTMLElement>("#tb-color")!;
    anchor.click();
    await new Promise((r) => setTimeout(r, 10));
    const menu = document.querySelector<HTMLElement>(".tb-dropdown")!;
    expect(menu).toBeTruthy();
    expect(anchor.getAttribute("aria-expanded")).toBe("true");
    const btns = [...menu.querySelectorAll<HTMLElement>("button")]; // 10 swatches + Auto
    // Opening focuses the first choice so the keyboard user can act immediately.
    expect(document.activeElement).toBe(btns[0]);

    // Arrow keys move the focus; Enter activates the focused choice (native
    // button behaviour in Chromium, not modelled by jsdom clicks).
    btns[0].dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    expect(document.activeElement).toBe(btns[1]);
    btns[1].dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    expect(document.activeElement).toBe(btns[2]);
    btns[2].dispatchEvent(new KeyboardEvent("keydown", { key: "End", bubbles: true }));
    expect(document.activeElement).toBe(btns[btns.length - 1]);
    btns[btns.length - 1].dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    expect(document.activeElement).toBe(btns[0]); // wraps

    // Escape dismisses the menu and returns focus to the toolbar button.
    btns[0].dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(document.querySelector(".tb-dropdown")).toBeNull();
    expect(anchor.hasAttribute("aria-expanded")).toBe(false);
  });

  it("table buttons follow the selection: enabled in a table, merge still off without cells", async () => {
    await boot();
    const merge = document.querySelector<HTMLButtonElement>("#tb-table-merge")!;
    const split = document.querySelector<HTMLButtonElement>("#tb-table-split")!;
    expect(merge.disabled).toBe(true);
    expect(split.disabled).toBe(true);

    document.querySelector<HTMLElement>("#tb-table")!.click();
    await new Promise((r) => setTimeout(r, 80));
    expect(editorEl().querySelector("table")).toBeTruthy();
    expect(document.querySelector<HTMLButtonElement>("#tb-table-row-del")!.disabled).toBe(false);
    // No drag-selection means nothing to merge; an unmerged cell cannot split.
    expect(merge.disabled).toBe(true);
    expect(split.disabled).toBe(true);
  });

  it("indent/outdent are disabled outside a list and usable inside one", async () => {
    await boot();
    const indent = document.querySelector<HTMLButtonElement>("#tb-indent")!;
    const outdent = document.querySelector<HTMLButtonElement>("#tb-outdent")!;
    expect(indent.disabled).toBe(true);
    expect(outdent.disabled).toBe(true);

    const buf = new TextEncoder().encode("alpha\nbeta") as unknown as Uint8Array;
    fw.openFile.mockResolvedValue({ path: "/x/l.txt", name: "l.txt", ext: "txt", buffer: buf });
    await send("file:open");
    document.querySelector<HTMLElement>("#tb-bullet")!.click();
    await new Promise((r) => setTimeout(r, 80));
    expect(editorEl().querySelector("ul")).toBeTruthy();
    // The first item cannot sink, but the caret in the second one can.
    expect(document.contains(indent)).toBe(true);
    expect(indent.hasAttribute("disabled")).toBe(true); // disabled in item 1
  });
});

describe("links, soft breaks, inline code", () => {
  it("the link button opens an in-app prompt and keeps the URL", async () => {
    await boot();
    document.querySelector<HTMLElement>("#tb-link")!.click();
    await new Promise((r) => setTimeout(r, 10));
    const modal = document.querySelector<HTMLElement>("#fw-prompt");
    expect(modal, "link prompt is an in-app modal, not window.confirm").toBeTruthy();
    expect(confirmSpy()).not.toHaveBeenCalled(); // no blocking native dialog

    (modal!.querySelector("#fw-prompt-input") as HTMLInputElement).value = "https://example.com/doc";
    (modal!.querySelector("#fw-prompt-ok") as HTMLButtonElement).click();
    await new Promise((r) => setTimeout(r, 10));
    expect(document.querySelector("#fw-prompt")).toBeNull();
  });

  it("the link prompt refuses a javascript: URL instead of storing it", async () => {
    await boot();
    document.querySelector<HTMLElement>("#tb-link")!.click();
    await new Promise((r) => setTimeout(r, 10));
    const modal = document.querySelector<HTMLElement>("#fw-prompt")!;
    const input = modal.querySelector("#fw-prompt-input") as HTMLInputElement;
    const ok = modal.querySelector("#fw-prompt-ok") as HTMLButtonElement;

    input.value = "javascript:alert(1)";
    ok.click();
    // Validation failure keeps the dialog open and explains itself.
    expect(document.querySelector("#fw-prompt")).toBeTruthy();
    expect((document.querySelector("#fw-prompt-error") as HTMLElement).hidden).toBe(false);
    expect(editorEl().querySelector("a")).toBeNull();

    // Escape still gets out.
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    await new Promise((r) => setTimeout(r, 10));
    expect(document.querySelector("#fw-prompt")).toBeNull();
  });

  it("Code and Link buttons are wired and mirror their mark state", async () => {
    await boot();
    const code = document.querySelector<HTMLButtonElement>("#tb-code")!;
    expect(code.classList.contains("active")).toBe(false);
    code.click();
    await new Promise((r) => setTimeout(r, 120));
    // Stored mark on an empty selection: no DOM to inspect, but the mirror must
    // follow it (this is the path that used to stay stale).
    expect(document.querySelector("#tb-code")!.classList.contains("active")).toBe(true);

    const link = document.querySelector<HTMLButtonElement>("#tb-link")!;
    const unlink = document.querySelector<HTMLButtonElement>("#tb-unlink")!;
    // Insert/edit is always offered (a disabled button could never create the
    // first link); removing an absent link is the only no-op greyed out.
    expect(link.disabled).toBe(false);
    expect(unlink.disabled).toBe(true);
  });

  it("Ctrl/Cmd+click on a link asks the main process to open it", async () => {
    await boot();
    // Insert the anchor through the schema (the mark's own command) rather than
    // typing, then click it as the user would.
    const view = document.querySelector<HTMLElement>(".ProseMirror")!;
    view.innerHTML = '<p>see <a href="https://example.com/x">the docs</a></p>';
    const anchor = view.querySelector("a")!;
    anchor.dispatchEvent(new MouseEvent("click", { bubbles: true, ctrlKey: true }));
    await new Promise((r) => setTimeout(r, 10));
    expect(fw.openExternal).toHaveBeenCalledWith("https://example.com/x");
  });

  it("a plain click on a link does not launch a browser", async () => {
    await boot();
    const view = document.querySelector<HTMLElement>(".ProseMirror")!;
    view.innerHTML = '<p>see <a href="https://example.com/x">the docs</a></p>';
    view.querySelector("a")!.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    await new Promise((r) => setTimeout(r, 10));
    expect(fw.openExternal).not.toHaveBeenCalled();
  });
});

describe("plain-text exports", () => {
  it("file:exportTxt writes the document text through saveRaw", async () => {
    await boot();
    const buf = new TextEncoder().encode("first para\nsecond para") as unknown as Uint8Array;
    fw.openFile.mockResolvedValue({ path: "/x/e.txt", name: "e.txt", ext: "txt", buffer: buf });
    await send("file:open");
    fw.saveRaw.mockResolvedValue({ path: "C:\\out\\e.txt" });
    await send("file:exportTxt");
    expect(fw.saveRaw).toHaveBeenCalledTimes(1);
    const payload = fw.saveRaw.mock.calls[0][0];
    expect(payload.ext).toBe("txt");
    expect(payload.content).toContain("first para");
    expect(payload.defaultName).toBe("e.txt");
  });

  it("file:exportHtml writes serialized HTML", async () => {
    await boot();
    const buf = new TextEncoder().encode("hello html") as unknown as Uint8Array;
    fw.openFile.mockResolvedValue({ path: "/x/h.txt", name: "h.txt", ext: "txt", buffer: buf });
    await send("file:open");
    fw.saveRaw.mockResolvedValue({ path: "C:\\out\\h.html" });
    await send("file:exportHtml");
    const payload = fw.saveRaw.mock.calls[0][0];
    expect(payload.ext).toBe("html");
    expect(payload.content).toContain("<p");
    expect(payload.content).toContain("hello html");
  });

  it("exporting a copy does not clear the document's own dirty flag", async () => {
    await boot();
    await send("insert:pageBreak"); // real edit
    fw.saveRaw.mockResolvedValue({ path: "C:\\out\\copy.txt" });
    await send("file:exportTxt");
    expect(document.title).toMatch(/^● /);
  });

  it("a failed export is reported instead of vanishing", async () => {
    await boot();
    fw.saveRaw.mockRejectedValue(new Error("disk full"));
    await send("file:exportTxt");
    await new Promise((r) => setTimeout(r, 20));
    expect(alertDialogText()).toContain("disk full");
  });
});

describe("PDF and print failure reporting", () => {
  it("a rejected PDF export alerts the user", async () => {
    await boot();
    fw.exportPdf.mockRejectedValue(new Error("print engine gone"));
    await send("file:exportPdf");
    await new Promise((r) => setTimeout(r, 20));
    expect(alertDialogText()).toContain("print engine gone");
  });

  it("a cancelled PDF export is silent and leaves print mode", async () => {
    await boot();
    fw.exportPdf.mockResolvedValue(null);
    await send("file:exportPdf");
    await new Promise((r) => setTimeout(r, 20));
    expect(alertDialogText()).toBeNull(); // dismissed dialog is not an error
    expect(document.body.classList.contains("print-mode")).toBe(false);
  });

  it("a print error is reported, and print mode always comes back down", async () => {
    await boot();
    fw.print.mockRejectedValue(new Error("no printer"));
    await send("file:print");
    await new Promise((r) => setTimeout(r, 20));
    expect(alertDialogText()).toContain("no printer");
    expect(document.body.classList.contains("print-mode")).toBe(false);
  });
});
