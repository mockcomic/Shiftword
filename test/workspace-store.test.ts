/**
 * Workspace page store: CRUD over the userData/workspace directory, the
 * path-traversal guard on ids, the JSON size cap, subtree delete to .trash,
 * and the corrupt-store tolerance — plus the same surface through the real
 * IPC handlers (main-frame-only).
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const handlers = new Map<string, (...args: unknown[]) => unknown>();
let userDataDir = "";
const MAIN_FRAME = { id: "main-frame" };

vi.mock("electron", () => ({
  app: {
    getPath: (name: string) => (name === "userData" ? userDataDir : tmpdir()),
    getVersion: () => "1.0.0",
    quit: vi.fn(),
    on: vi.fn(),
    whenReady: () => ({ then: () => {} }),
    requestSingleInstanceLock: () => true,
    commandLine: { appendSwitch: vi.fn() },
  },
  BrowserWindow: class {},
  dialog: {
    showOpenDialog: vi.fn(),
    showSaveDialog: vi.fn(),
    showErrorBox: vi.fn(),
    showMessageBox: vi.fn(),
  },
  shell: { openExternal: vi.fn() },
  ipcMain: {
    handle: vi.fn((channel: string, fn: (...a: unknown[]) => unknown) => handlers.set(channel, fn)),
    on: vi.fn(),
  },
  Menu: { setApplicationMenu: vi.fn(), buildFromTemplate: vi.fn((t: unknown) => t) },
}));

const ws = await import("../src/main/workspace");
const { registerIpc } = await import("../src/main/ipc");

const call = (channel: string, ...args: unknown[]) =>
  handlers.get(channel)!({ senderFrame: MAIN_FRAME }, ...args);

beforeEach(async () => {
  handlers.clear();
  userDataDir = await mkdtemp(join(tmpdir(), "fw-workspace-"));
  registerIpc(
    () =>
      ({
        isDestroyed: () => false,
        webContents: { isDestroyed: () => false, mainFrame: MAIN_FRAME },
      }) as never,
    () => {},
  );
});

describe("workspace store", () => {
  it("lists nothing on a cold workspace", async () => {
    expect(await ws.listPages()).toEqual([]);
  });

  it("creates, lists, writes and reads back a page", async () => {
    const meta = await ws.createPage({ title: "Meeting notes" });
    expect(meta).toBeTruthy();
    expect((await ws.listPages()).map((p) => p.title)).toEqual(["Meeting notes"]);

    const doc = {
      type: "doc",
      attrs: {},
      content: [
        { type: "page", content: [{ type: "paragraph", content: [{ type: "text", text: "agenda" }] }] },
      ],
    };
    expect(await ws.writePage(meta!.id, doc)).toEqual({});
    const read = await ws.readPage(meta!.id);
    expect(read?.meta.id).toBe(meta!.id);
    expect(JSON.stringify(read?.json)).toContain("agenda");
  });

  it("refuses invalid ids instead of letting them steer a path", async () => {
    for (const evil of ["../../etc/passwd", "p/../x", "a", "p.x", "pages", "P-!!"]) {
      expect(await ws.writePage(evil, { type: "doc" })).toHaveProperty("error");
      expect(await ws.readPage(evil)).toBeNull();
      expect(await ws.deletePage(evil)).toHaveProperty("error");
    }
  });

  it("enforces the document size cap", async () => {
    const meta = await ws.createPage({ title: "Huge" });
    const huge = { type: "doc", filler: "x".repeat(11 * 1024 * 1024) };
    const res = await ws.writePage(meta!.id, huge);
    expect(res.error).toMatch(/larger than/);
  });

  it("renames, re-parents and re-orders", async () => {
    const parent = (await ws.createPage({ title: "Parent" }))!;
    const child = (await ws.createPage({ title: "Child", parentId: parent.id }))!;
    expect(child.parentId).toBe(parent.id);

    await ws.renamePage(child.id, { title: "Renamed", icon: "🚀" });
    const pages = await ws.listPages();
    const moved = pages.find((p) => p.id === child.id)!;
    expect(moved.title).toBe("Renamed");
    expect(moved.icon).toBe("🚀");

    await ws.movePage(child.id, null, 7);
    expect((await ws.listPages()).find((p) => p.id === child.id)!.parentId).toBeNull();
  });

  it("refuses a move that would create a cycle", async () => {
    const a = (await ws.createPage({ title: "A" }))!;
    const b = (await ws.createPage({ title: "B", parentId: a.id }))!;
    const res = await ws.movePage(a.id, b.id, 0);
    expect(res.error).toMatch(/cycle|itself/);
  });

  it("deletes a subtree into .trash", async () => {
    const parent = (await ws.createPage({ title: "Branch" }))!;
    const child = (await ws.createPage({ title: "Leaf", parentId: parent.id }))!;
    const res = await ws.deletePage(parent.id);
    expect(res).toEqual({});
    const pages = await ws.listPages();
    expect(pages.some((p) => p.id === parent.id)).toBe(false);
    expect(pages.some((p) => p.id === child.id)).toBe(false); // subtree followed
    const trashDir = join(userDataDir, "workspace", ".trash");
    expect(existsSync(trashDir)).toBe(true);
    const trashed = await readdir(trashDir);
    expect(trashed.length).toBe(2); // both doc files moved, not deleted
  });

  it("repairs a corrupt pages.json to an empty workspace", async () => {
    const meta = await ws.createPage({ title: "Fine" });
    await writeFile(join(userDataDir, "workspace", "pages.json"), "{{{");
    expect(await ws.listPages()).toEqual([]);
    expect(await ws.readPage(meta!.id)).toBeNull();
  });

  it("opens a metadata-only page as a clean document", async () => {
    const meta = (await ws.createPage({ title: "Docless" }))!;
    const file = join(userDataDir, "workspace", `${meta.id}.json`);
    const { rm } = await import("node:fs/promises");
    await rm(file);
    const read = await ws.readPage(meta.id);
    expect(read?.json.type).toBe("doc");
  });
});

describe("trash restore", () => {
  it("lists deleted pages by title and restores them intact", async () => {
    const meta = (await ws.createPage({ title: "Doomed draft" }))!;
    const doc = {
      type: "doc",
      attrs: {},
      content: [
        { type: "page", content: [{ type: "paragraph", content: [{ type: "text", text: "keep me" }] }] },
      ],
    };
    await ws.writePage(meta.id, doc);
    expect(await ws.deletePage(meta.id)).toEqual({});

    const trashed = await ws.listTrash();
    expect(trashed).toHaveLength(1);
    expect(trashed[0].title).toBe("Doomed draft");
    expect(trashed[0].id).toBe(meta.id);
    expect(trashed[0].deletedAt).toBeGreaterThan(0);

    const res = await ws.restoreTrash(trashed[0].file);
    expect(res.error).toBeUndefined();
    expect(res.meta?.title).toBe("Doomed draft");
    expect(await ws.listTrash()).toEqual([]);
    const read = await ws.readPage(meta.id);
    expect(JSON.stringify(read?.json)).toContain("keep me");
  });

  it("re-attaches a restored page to the root when its parent is gone too", async () => {
    const parent = (await ws.createPage({ title: "P" }))!;
    const child = (await ws.createPage({ title: "C", parentId: parent.id }))!;
    await ws.deletePage(child.id); // child trashed while parent still exists
    await ws.deletePage(parent.id); // then the parent goes too
    const childEntry = (await ws.listTrash()).find((t) => t.id === child.id)!;
    expect(childEntry.parentId).toBe(parent.id); // remembered at delete time
    const res = await ws.restoreTrash(childEntry.file);
    expect(res.meta?.parentId).toBeNull();
  });

  it("surfaces legacy raw-doc trash files and restores them as clean pages", async () => {
    const { mkdir } = await import("node:fs/promises");
    const trash = join(userDataDir, "workspace", ".trash");
    await mkdir(trash, { recursive: true });
    await writeFile(join(trash, "p-legacy1-abc.json"), JSON.stringify({ type: "doc", attrs: {} }));
    const entry = (await ws.listTrash()).find((t) => t.id === "p-legacy1")!;
    expect(entry.title).toBe("Untitled");
    const res = await ws.restoreTrash(entry.file);
    expect(res.error).toBeUndefined();
    expect(await ws.readPage("p-legacy1")).toBeTruthy();
  });

  it("deleteTrash erases one entry for good", async () => {
    const meta = (await ws.createPage({ title: "Erased" }))!;
    await ws.deletePage(meta.id);
    const entry = (await ws.listTrash())[0]!;
    expect(await ws.deleteTrash(entry.file)).toEqual({});
    expect(await ws.listTrash()).toEqual([]);
    expect(existsSync(join(userDataDir, "workspace", ".trash", entry.file))).toBe(false);
  });

  it("refuses trash names that could steer a path", async () => {
    for (const evil of ["../evil.json", "no-json", "a.json", "p a-1.json"]) {
      const r = await ws.restoreTrash(evil);
      expect(r.error).toMatch(/Invalid/);
      expect((await ws.deleteTrash(evil)).error).toMatch(/Invalid/);
    }
  });
});

describe("page IPC handlers", () => {
  it("create → list → write → read → delete through the bridge", async () => {
    const meta = (await call("page:create", { title: "Via IPC" })) as Record<string, unknown>;
    expect(typeof meta.id).toBe("string");
    const id = meta.id as string;
    const listed = (await call("page:list")) as Array<Record<string, unknown>>;
    expect(listed.map((p) => p.title)).toContain("Via IPC");
    expect(await call("page:write", id, { type: "doc", attrs: {}, content: [] })).toEqual({});
    const read = (await call("page:read", id)) as { meta: { title: string }; json: { type: string } };
    expect(read.meta.title).toBe("Via IPC");
    expect(read.json.type).toBe("doc");
    expect(await call("page:delete", id)).toEqual({});
    expect(((await call("page:list")) as unknown[]).length).toBe(0);
  });

  it("ignores untrusted senders", async () => {
    const rogue = { senderFrame: { id: "rogue" } };
    expect(await handlers.get("page:list")!(rogue)).toEqual([]);
    expect(await handlers.get("page:read")!(rogue, "whatever123")).toBeNull();
    expect(await handlers.get("page:create")!(rogue)).toBeNull();
    const w = (await handlers.get("page:write")!(rogue, "whatever123", {})) as { error?: string };
    expect(w.error).toBeTruthy();
    expect(await handlers.get("page:trashList")!(rogue)).toEqual([]);
    const rt = (await handlers.get("page:trashRestore")!(rogue, "p-x-y.json")) as { error?: string };
    expect(rt.error).toBeTruthy();
    const dt = (await handlers.get("page:trashDelete")!(rogue, "p-x-y.json")) as { error?: string };
    expect(dt.error).toBeTruthy();
  });

  it("delete → trashList → restore through the bridge", async () => {
    const meta = (await call("page:create", { title: "Via IPC trash" })) as { id: string };
    expect(await call("page:delete", meta.id)).toEqual({});
    const trashed = (await call("page:trashList")) as Array<{ file: string; title: string }>;
    expect(trashed.map((t) => t.title)).toContain("Via IPC trash");
    const res = (await call("page:trashRestore", trashed[0].file)) as { error?: string };
    expect(res.error).toBeUndefined();
    const listed = (await call("page:list")) as Array<{ id: string }>;
    expect(listed.map((p) => p.id)).toContain(meta.id);
  });

  it("keeps the metadata file valid JSON after many operations", async () => {
    for (let i = 0; i < 5; i++) await call("page:create", { title: `P${i}` });
    const raw = await readFile(join(userDataDir, "workspace", "pages.json"), "utf8");
    expect(() => JSON.parse(raw)).not.toThrow();
  });
});
