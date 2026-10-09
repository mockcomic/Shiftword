// @vitest-environment jsdom
/**
 * Workspace pages from the renderer's side: opening a page, the autosave
 * debounce, Ctrl+S as an immediate flush, the quit handshake flushing before
 * answering, page:new / page:open menu commands, and the quick switcher —
 * all against a mocked preload bridge with an in-memory store.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

interface Stored {
  meta: {
    id: string;
    title: string;
    icon: string;
    parentId: string | null;
    order: number;
    updatedAt: number;
  };
  json: Record<string, unknown>;
}

let store = new Map<string, Stored>();
let trashStore = new Map<string, Stored & { file: string; deletedAt: number }>();
let nextId = 1;

(Range as unknown as { prototype: Record<string, unknown> }).prototype.getClientRects = () =>
  Object.assign([], { item: () => null });
(Range as unknown as { prototype: Record<string, unknown> }).prototype.getBoundingClientRect = () =>
  ({ x: 0, y: 0, width: 0, height: 0, top: 0, left: 0, right: 0, bottom: 0 }) as DOMRect;

function makeStored(title: string, json?: Record<string, unknown>): Stored {
  const id = `p-test-${nextId++}`;
  return {
    meta: { id, title, icon: "📄", parentId: null, order: 0, updatedAt: Date.now() },
    json: json ?? {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [{ type: "page", content: [{ type: "paragraph" }] }],
    },
  };
}

let fw: Record<string, ReturnType<typeof vi.fn>>;
let commandCb: ((cmd: { id: string; path?: string }) => void) | null;
let saveThenQuitCb: (() => void) | null;

async function boot(): Promise<void> {
  vi.resetModules();
  store = new Map();
  trashStore = new Map();
  nextId = 1;
  document.body.innerHTML = `
    <div id="toolbar"></div>
    <div id="canvas"><div id="editor"></div></div>
    <div id="statusbar"></div>`;
  document.body.className = "";
  commandCb = null;
  saveThenQuitCb = null;
  const written: string[] = [];
  fw = {
    openFile: vi.fn().mockResolvedValue(null),
    saveDoc: vi.fn().mockResolvedValue(null),
    saveRaw: vi.fn().mockResolvedValue(null),
    pickImage: vi.fn().mockResolvedValue(null),
    exportPdf: vi.fn().mockResolvedValue(null),
    print: vi.fn().mockResolvedValue(undefined),
    getRecent: vi.fn().mockResolvedValue([]),
    openExternal: vi.fn().mockResolvedValue(true),
    getSettings: vi
      .fn()
      .mockResolvedValue({ notionMode: false, theme: "light", sidebarOpen: true, toolbarPinned: false }),
    setSettings: vi.fn().mockImplementation(async (p: Record<string, unknown>) => p),
    listPages: vi.fn(async () => [...store.values()].map((s) => s.meta)),
    readPage: vi.fn(async (id: string) => store.get(id) ?? null),
    writePage: vi.fn(async (id: string, json: Record<string, unknown>) => {
      const s = store.get(id);
      if (!s) return { error: "Unknown page." };
      s.json = json;
      written.push(id);
      return {};
    }),
    createPage: vi.fn(async (opts?: { title?: string; parentId?: string | null }) => {
      const s = makeStored(opts?.title || "Untitled");
      s.meta.parentId = opts?.parentId ?? null;
      store.set(s.meta.id, s);
      return s.meta;
    }),
    renamePage: vi.fn(async (id: string, patch: { title?: string; icon?: string }) => {
      const s = store.get(id);
      if (!s) return null;
      if (patch.title) s.meta.title = patch.title;
      if (patch.icon) s.meta.icon = patch.icon;
      return s.meta;
    }),
    duplicatePage: vi.fn(async (id: string) => {
      const s = store.get(id);
      if (!s) return { error: "Unknown page." };
      const copy = makeStored(`${s.meta.title} (copy)`, JSON.parse(JSON.stringify(s.json)));
      copy.meta.parentId = s.meta.parentId;
      copy.meta.icon = s.meta.icon;
      store.set(copy.meta.id, copy);
      return { meta: copy.meta };
    }),
    movePage: vi.fn(async (id: string, parentId: string | null) => {
      const s = store.get(id);
      if (!s) return { error: "Unknown page." };
      s.meta.parentId = parentId;
      return {};
    }),
    deletePage: vi.fn(async (id: string) => {
      const s = store.get(id);
      store.delete(id);
      if (s) {
        const file = `${id}-t${nextId++}.json`;
        trashStore.set(file, { ...s, file, deletedAt: Date.now() });
      }
      return {};
    }),
    listTrash: vi.fn(async () =>
      [...trashStore.values()]
        .map((t) => ({
          file: t.file,
          id: t.meta.id,
          title: t.meta.title,
          icon: t.meta.icon,
          parentId: t.meta.parentId,
          updatedAt: t.meta.updatedAt,
          deletedAt: t.deletedAt,
        }))
        .sort((a, b) => b.deletedAt - a.deletedAt),
    ),
    restoreTrash: vi.fn(async (file: string) => {
      const t = trashStore.get(file);
      if (!t) return { error: "Unknown trash entry." };
      trashStore.delete(file);
      store.set(t.meta.id, { meta: t.meta, json: t.json });
      return { meta: t.meta };
    }),
    deleteTrash: vi.fn(async (file: string) => {
      trashStore.delete(file);
      return {};
    }),
    onCommand: vi.fn((cb: (cmd: { id: string; path?: string }) => void) => {
      commandCb = cb;
      return () => {};
    }),
    sendQuitResponse: vi.fn(),
    onWillQuit: vi.fn((cb: () => void) => {
      (globalThis as Record<string, unknown>).__willQuit = cb;
      return () => {};
    }),
    onSaveThenQuit: vi.fn((cb: () => void) => {
      saveThenQuitCb = cb;
      return () => {};
    }),
  };
  Object.defineProperty(window, "fw", { value: fw, configurable: true, writable: true });
  Object.defineProperty(window, "confirm", { value: vi.fn(() => true), configurable: true });
  Object.defineProperty(document, "fonts", { value: { ready: Promise.resolve() }, configurable: true });
  if (!window.requestAnimationFrame) {
    Object.defineProperty(window, "requestAnimationFrame", {
      value: (cb: (t: number) => void) => setTimeout(() => cb(Date.now()), 0),
      configurable: true,
    });
  }
  const { initApp } = await import("../src/renderer/src/app");
  initApp();
  await new Promise((r) => setTimeout(r, 240));
  (fw as unknown as { __written: string[] }).__written = written;
}

const tick = (ms = 800): Promise<void> => new Promise((r) => setTimeout(r, ms));

describe("workspace pages", () => {
  beforeEach(async () => {
    await boot();
  });

  it("the sidebar renders pages from the store", async () => {
    store.set("p-a", makeStored("Alpha"));
    store.set("p-b", makeStored("Bef"));
    commandCb!({ id: "view:quickSwitcher" }); // triggers a reloadPages()
    await tick(50);
    document.getElementById("fw-quick-switcher")?.remove();
    const rows = Array.from(document.querySelectorAll("#sidebar .side-label")).map((n) => n.textContent);
    expect(rows).toContain("Alpha");
    expect(rows).toContain("Bef");
  });

  it("page:new creates a page, opens it, and the title follows", async () => {
    commandCb!({ id: "page:new" });
    await tick(50);
    const id = [...store.keys()][0];
    expect(id).toBeTruthy();
    expect(document.title).toContain("Untitled");
    expect(store.get(id)!.meta.title).toBe("Untitled");
  });

  it("typing autosaves the page after the debounce window", async () => {
    const s = makeStored("Scratch");
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    // Any editing command that changes the doc triggers the autosave path:
    // insert:table goes through the editor update handler.
    commandCb!({ id: "insert:table" });
    await tick(1000); // debounce (600ms) + margin
    expect(fw.writePage).toHaveBeenCalled();
    expect(JSON.stringify(store.get(s.meta.id)!.json)).toContain("table");
  });

  it("file:save on a page flushes instead of opening a dialog", async () => {
    const s = makeStored("Flush me");
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    commandCb!({ id: "insert:table" });
    commandCb!({ id: "file:save" });
    await tick(50);
    expect(fw.saveDoc, "Ctrl+S must not go to the file dialog for a page").not.toHaveBeenCalled();
    expect(fw.writePage).toHaveBeenCalled();
    expect(JSON.stringify(store.get(s.meta.id)!.json)).toContain("table");
  });

  it("the quit handshake flushes the page before answering", async () => {
    const s = makeStored("Quit me");
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    commandCb!({ id: "insert:table" });
    const willQuit = (globalThis as Record<string, unknown>).__willQuit as () => void;
    willQuit();
    await tick(50);
    expect(fw.sendQuitResponse).toHaveBeenCalled();
    const answer = fw.sendQuitResponse.mock.calls.at(-1)![0] as { dirty: boolean };
    expect(answer.dirty).toBe(false); // the flush completed first
    expect(JSON.stringify(store.get(s.meta.id)!.json)).toContain("table");
  });

  it("save-then-quit routes through the same flush", async () => {
    const s = makeStored("Quit2");
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    commandCb!({ id: "insert:table" });
    saveThenQuitCb!();
    await tick(50);
    expect(fw.writePage).toHaveBeenCalled();
    expect(fw.sendQuitResponse.mock.calls.at(-1)![0]).toMatchObject({ dirty: false });
  });

  it("the quick switcher filters titles and opens the picked page", async () => {
    const s1 = makeStored("Groceries list");
    const s2 = makeStored("Meeting notes");
    store.set(s1.meta.id, s1);
    store.set(s2.meta.id, s2);
    commandCb!({ id: "view:quickSwitcher" });
    await tick(50);
    const input = document.querySelector<HTMLInputElement>("#quick-input")!;
    expect(input).toBeTruthy();
    input.value = "groc";
    input.dispatchEvent(new Event("input"));
    await tick(20);
    const items = Array.from(document.querySelectorAll(".quick-item"));
    expect(items).toHaveLength(1);
    (items[0] as HTMLButtonElement).click();
    await tick(50);
    expect(document.title).toContain("Groceries list");
  });

  it("opening a page switches the document content and marks it clean", async () => {
    const s = makeStored("Doc", {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        {
          type: "page",
          content: [{ type: "paragraph", content: [{ type: "text", text: "page body text" }] }],
        },
      ],
    });
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    expect(document.querySelector(".ProseMirror")!.textContent).toContain("page body text");
    expect(document.title).not.toContain("● "); // clean after opening (setContent is silent)
    expect(document.title).toContain("Doc");
  });

  it("deleting the open page drops back to a blank document", async () => {
    const s = makeStored("Doomed");
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    // Delete through the sidebar flow — confirm modal must be accepted.
    const moreBtns = document.querySelectorAll<HTMLButtonElement>("#sidebar .side-more");
    expect(moreBtns.length).toBe(1);
    moreBtns[0].click();
    await tick(20);
    const delBtn = Array.from(document.querySelectorAll<HTMLButtonElement>(".block-menu-item")).find((b) =>
      b.textContent?.includes("Delete"),
    )!;
    delBtn.click();
    await tick(20);
    // accept the confirm dialog
    const ok = document.querySelector<HTMLButtonElement>("#fw-confirm-ok")!;
    expect(ok).toBeTruthy();
    ok.click();
    await tick(50);
    expect(document.title).toContain("Untitled");
  });

  it("the row menu duplicates a page beside the original", async () => {
    const s = makeStored("Original");
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    document.querySelector<HTMLButtonElement>("#sidebar .side-more")!.click();
    await tick(20);
    Array.from(document.querySelectorAll<HTMLButtonElement>(".block-menu-item"))
      .find((b) => b.textContent?.includes("Duplicate"))!
      .click();
    await tick(50);
    expect(fw.duplicatePage).toHaveBeenCalledWith(s.meta.id);
    expect([...store.values()].map((x) => x.meta.title)).toContain("Original (copy)");
    expect(document.querySelector("#side-tree")!.textContent).toContain("Original (copy)");
  });

  it("the sidebar trash lists deleted pages and restores them", async () => {
    const s = makeStored("Rescuable");
    store.set(s.meta.id, s);
    commandCb!({ id: "page:open", path: s.meta.id });
    await tick(50);
    // Delete via the row menu + confirm.
    document.querySelector<HTMLButtonElement>("#sidebar .side-more")!.click();
    await tick(20);
    Array.from(document.querySelectorAll<HTMLButtonElement>(".block-menu-item"))
      .find((b) => b.textContent?.includes("Delete"))!
      .click();
    await tick(20);
    document.querySelector<HTMLButtonElement>("#fw-confirm-ok")!.click();
    await tick(50);

    // Open the trash section and check the entry reads by title.
    const toggle = document.getElementById("side-trash-toggle") as HTMLButtonElement;
    toggle.click();
    await tick(20);
    const trashRows = [...document.querySelectorAll<HTMLElement>(".side-trash-row")];
    expect(trashRows).toHaveLength(1);
    expect(trashRows[0].textContent).toContain("Rescuable");

    // Restore puts it back in the tree and out of the trash.
    trashRows[0].querySelector<HTMLButtonElement>(".side-trash-restore")!.click();
    await tick(50);
    expect(document.querySelectorAll(".side-trash-row")).toHaveLength(0);
    const treeText = [...document.querySelectorAll("#side-tree .side-row")].map((r) => r.textContent);
    expect(treeText.some((t) => t?.includes("Rescuable"))).toBe(true);
    expect(store.has(s.meta.id)).toBe(true);
  });

  it("deleting forever and emptying trash erase entries after a confirm", async () => {
    const a = makeStored("Gone-a");
    const b = makeStored("Gone-b");
    store.set(a.meta.id, a);
    store.set(b.meta.id, b);
    commandCb!({ id: "page:open", path: a.meta.id });
    await tick(50);
    // Delete both through the bridge the way the row menu does.
    await fw.deletePage(a.meta.id);
    await fw.deletePage(b.meta.id);
    commandCb!({ id: "view:quickSwitcher" }); // triggers a reloadPages()
    document.getElementById("fw-quick-switcher")?.remove();
    await tick(30);
    (document.getElementById("side-trash-toggle") as HTMLButtonElement).click();
    await tick(20);
    expect(document.querySelectorAll(".side-trash-row")).toHaveLength(2);

    // Single permanent delete asks first, then erases only that entry.
    document.querySelectorAll<HTMLButtonElement>(".side-trash-delete")[0].click();
    await tick(20);
    const confirm = document.getElementById("fw-confirm");
    expect(confirm, "permanent delete asks for confirmation").toBeTruthy();
    document.querySelector<HTMLButtonElement>("#fw-confirm-ok")!.click();
    await tick(50);
    expect(document.querySelectorAll(".side-trash-row")).toHaveLength(1);

    // Empty trash clears the rest.
    document.querySelector<HTMLButtonElement>(".side-trash-empty-all")!.click();
    await tick(20);
    document.querySelector<HTMLButtonElement>("#fw-confirm-ok")!.click();
    await tick(50);
    expect(document.querySelectorAll(".side-trash-row")).toHaveLength(0);
    expect(trashStore.size).toBe(0);
  });
});
