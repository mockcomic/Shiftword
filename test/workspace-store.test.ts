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
  });

  it("keeps the metadata file valid JSON after many operations", async () => {
    for (let i = 0; i < 5; i++) await call("page:create", { title: `P${i}` });
    const raw = await readFile(join(userDataDir, "workspace", "pages.json"), "utf8");
    expect(() => JSON.parse(raw)).not.toThrow();
  });
});
