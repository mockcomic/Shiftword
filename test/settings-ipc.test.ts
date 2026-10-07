/**
 * The settings IPC: persistence to userData/settings.json, key-by-key patch
 * filtering, corrupt-file tolerance, and the trusted-frame posture — an
 * untrusted caller must get defaults, never the user's settings.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const handlers = new Map<string, (...args: unknown[]) => unknown>();
let userDataDir = "";
const MAIN_FRAME = { id: "main-frame" };
const ROUGE_FRAME = { id: "other-frame" };

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

const { registerIpc } = await import("../src/main/ipc");

const call = (channel: string, ...args: unknown[]) =>
  handlers.get(channel)!({ senderFrame: MAIN_FRAME }, ...args);
const callFrom = (frame: unknown, channel: string, ...args: unknown[]) =>
  handlers.get(channel)!({ senderFrame: frame }, ...args);

beforeEach(async () => {
  handlers.clear();
  userDataDir = await mkdtemp(join(tmpdir(), "fw-settings-"));
  registerIpc(
    () =>
      ({
        isDestroyed: () => false,
        webContents: { isDestroyed: () => false, mainFrame: MAIN_FRAME },
      }) as never,
    () => {},
  );
});

describe("settings:get", () => {
  it("answers defaults when no file exists", async () => {
    await expect(call("settings:get")).resolves.toEqual({
      notionMode: false,
      theme: "light",
      sidebarOpen: false,
      toolbarPinned: false,
    });
  });

  it("repairs a corrupt settings.json instead of failing", async () => {
    await writeFile(join(userDataDir, "settings.json"), "{ not json");
    await expect(call("settings:get")).resolves.toMatchObject({ notionMode: false, theme: "light" });
  });

  it("rejects invalid enum values", async () => {
    await writeFile(
      join(userDataDir, "settings.json"),
      JSON.stringify({ notionMode: "yes", theme: "hotdog", sidebarOpen: 1 }),
    );
    const s = (await call("settings:get")) as Record<string, unknown>;
    expect(s.notionMode).toBe(false);
    expect(s.theme).toBe("light");
    expect(s.sidebarOpen).toBe(false);
  });

  it("refuses to leak settings to an untrusted frame", async () => {
    await call("settings:set", { notionMode: true });
    const leaked = (await callFrom(ROUGE_FRAME, "settings:get")) as Record<string, unknown>;
    expect(leaked.notionMode).toBe(false);
  });
});

describe("settings:set", () => {
  it("merges the patch, persists it, and answers the full merged settings", async () => {
    const next = await call("settings:set", { notionMode: true, theme: "dark" });
    expect(next).toEqual({ notionMode: true, theme: "dark", sidebarOpen: false, toolbarPinned: false });
    const raw = JSON.parse(await readFile(join(userDataDir, "settings.json"), "utf8"));
    expect(raw.notionMode).toBe(true);
    expect(raw.theme).toBe("dark");
  });

  it("ignores unknown or mistyped keys", async () => {
    const next = await call("settings:set", { theme: "neon", evil: 42, notionMode: "true" });
    expect(next).toMatchObject({ theme: "light", notionMode: false });
    expect(Object.keys(next)).not.toContain("evil");
  });

  it("rebuilds the native menu so its checkboxes mirror the change", async () => {
    let rebuilt = 0;
    handlers.clear();
    registerIpc(
      () =>
        ({
          isDestroyed: () => false,
          webContents: { isDestroyed: () => false, mainFrame: MAIN_FRAME },
        }) as never,
      () => rebuilt++,
    );
    await call("settings:set", { notionMode: true });
    expect(rebuilt).toBe(1);
  });

  it("survives an unwritable userData directory", async () => {
    // Block the store at the source: userData itself is a plain file, so both
    // mkdir -p and the write fail. The answer must still come back correct.
    const parent = userDataDir + "-parent";
    await writeFile(parent, "not a directory");
    userDataDir = parent;
    const next = await call("settings:set", { notionMode: true });
    expect(next.notionMode).toBe(true); // in-memory answer stays authoritative
  });
});
