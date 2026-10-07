// @vitest-environment jsdom
/**
 * Notion mode end-to-end in the renderer: settings load at boot, the toggle
 * flips the chrome and the pagination engine, the print wrap restores the
 * Word rendering regardless of the visible mode, and everything persists
 * through the (mocked) preload bridge.
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
  getSettings: ReturnType<typeof vi.fn>;
  setSettings: ReturnType<typeof vi.fn>;
  onCommand: ReturnType<typeof vi.fn>;
  sendQuitResponse: ReturnType<typeof vi.fn>;
  onWillQuit: ReturnType<typeof vi.fn>;
  onSaveThenQuit: ReturnType<typeof vi.fn>;
}

let fw: FwMock;
let commandCb: ((cmd: { id: string; path?: string }) => void) | null;
let persisted: Record<string, unknown> = {};

(Range as unknown as { prototype: Record<string, Cb> }).prototype.getClientRects = () =>
  Object.assign([], { item: () => null });
(Range as unknown as { prototype: Record<string, Cb> }).prototype.getBoundingClientRect = () =>
  ({ x: 0, y: 0, width: 0, height: 0, top: 0, left: 0, right: 0, bottom: 0 }) as DOMRect;

async function boot(
  initial: Partial<{ notionMode: boolean; theme: string; sidebarOpen: boolean; toolbarPinned: boolean }> = {},
): Promise<void> {
  vi.resetModules();
  document.body.className = "";
  delete document.body.dataset.theme;
  delete document.body.dataset.uiMode;
  document.body.innerHTML = `
    <div id="toolbar"></div>
    <div id="canvas"><div id="editor"></div></div>
    <div id="statusbar"></div>`;
  persisted = { ...initial };
  commandCb = null;
  fw = {
    openFile: vi.fn().mockResolvedValue(null),
    saveDoc: vi.fn().mockResolvedValue(null),
    saveRaw: vi.fn().mockResolvedValue(null),
    pickImage: vi.fn().mockResolvedValue(null),
    exportPdf: vi.fn().mockResolvedValue(null),
    print: vi.fn().mockResolvedValue(undefined),
    getRecent: vi.fn().mockResolvedValue([]),
    openExternal: vi.fn().mockResolvedValue(true),
    getSettings: vi.fn().mockImplementation(async () => ({
      notionMode: false,
      theme: "light",
      sidebarOpen: false,
      toolbarPinned: false,
      ...persisted,
    })),
    setSettings: vi.fn().mockImplementation(async (patch: Record<string, unknown>) => {
      persisted = { ...persisted, ...patch };
      return { ...persisted };
    }),
    onCommand: vi.fn((cb: Cb) => {
      commandCb = cb as never;
      return () => {};
    }),
    sendQuitResponse: vi.fn(),
    onWillQuit: vi.fn(() => () => {}),
    onSaveThenQuit: vi.fn(() => () => {}),
  };
  Object.defineProperty(window, "fw", { value: fw, configurable: true, writable: true });
  Object.defineProperty(document, "fonts", { value: { ready: Promise.resolve() }, configurable: true });
  if (!window.requestAnimationFrame) {
    Object.defineProperty(window, "requestAnimationFrame", {
      value: (cb: (t: number) => void) => setTimeout(() => cb(Date.now()), 0),
      configurable: true,
    });
  }
  const { initApp } = await import("../src/renderer/src/app");
  initApp();
  await new Promise((r) => setTimeout(r, 240)); // settings promise + debounces
}

const send = async (id: string): Promise<void> => {
  commandCb!({ id });
  await new Promise((r) => setTimeout(r, 20));
};

describe("boot with settings", () => {
  it("defaults keep the Word chrome and running pagination", async () => {
    await boot();
    expect(document.body.classList.contains("notion-mode")).toBe(false);
    const { isPaginationPaused } = await import("../src/renderer/src/editor/pagination");
    expect(isPaginationPaused()).toBe(false);
    expect(document.querySelector("#statusbar")!.textContent).toContain("Page 1 / 1");
    expect(document.querySelector("#statusbar")!.textContent).toContain("DOCX");
  });

  it("notionMode=true from disk applies the class, pauses pagination, flips the status label", async () => {
    await boot({ notionMode: true });
    expect(document.body.classList.contains("notion-mode")).toBe(true);
    const { isPaginationPaused } = await import("../src/renderer/src/editor/pagination");
    expect(isPaginationPaused()).toBe(true);
    // The page readout stays in the DOM (the smoke probe needs it) but the CSS
    // hides it via .sb-page in notion mode.
    expect(document.querySelector("#statusbar .sb-page")).toBeTruthy();
    expect(document.querySelector("#statusbar .sb-right")!.textContent).toContain("Notion");
  });

  it("applies theme=dark and follows the resolved value", async () => {
    await boot({ theme: "dark" });
    expect(document.body.dataset.theme).toBe("dark");
  });
});

describe("toggling", () => {
  beforeEach(async () => {
    await boot();
  });

  it("view:notionMode turns Notion mode on, persists it, and pauses pagination", async () => {
    await send("view:notionMode");
    expect(document.body.classList.contains("notion-mode")).toBe(true);
    const { isPaginationPaused } = await import("../src/renderer/src/editor/pagination");
    expect(isPaginationPaused()).toBe(true);
    // Turning Notion mode on also opens the sidebar (the user can close it
    // again and that sticks) — one persisted patch covering both.
    expect(fw.setSettings).toHaveBeenCalledWith({ notionMode: true, sidebarOpen: true });
    expect(persisted.notionMode).toBe(true);
  });

  it("toggles back off and re-partitions (runs pagination once)", async () => {
    await send("view:notionMode");
    await send("view:notionMode");
    expect(document.body.classList.contains("notion-mode")).toBe(false);
    const { isPaginationPaused } = await import("../src/renderer/src/editor/pagination");
    expect(isPaginationPaused()).toBe(false);
  });

  it("view:theme:* persists and applies the theme", async () => {
    await send("view:theme:dark");
    expect(document.body.dataset.theme).toBe("dark");
    expect(fw.setSettings).toHaveBeenCalledWith({ theme: "dark" });
  });

  it("view:toggleToolbarPin persists the pin", async () => {
    await send("view:notionMode");
    await send("view:toggleToolbarPin");
    expect(document.body.classList.contains("pinned-toolbar")).toBe(true);
    expect(persisted.toolbarPinned).toBe(true);
  });

  it("print wrap: Notion class comes off for the capture and back on after", async () => {
    await send("view:notionMode");
    const { beginPrintAppearance, endPrintAppearance } =
      await import("../src/renderer/src/features/appearance");
    beginPrintAppearance();
    expect(document.body.classList.contains("notion-mode")).toBe(false);
    const { isPaginationPaused } = await import("../src/renderer/src/editor/pagination");
    expect(isPaginationPaused()).toBe(false);
    endPrintAppearance();
    expect(document.body.classList.contains("notion-mode")).toBe(true);
    expect(isPaginationPaused()).toBe(true);
  });
});

describe("standalone resilience", () => {
  it("an fw mock without settings support still boots the Word chrome", async () => {
    vi.resetModules();
    document.body.innerHTML = `
      <div id="toolbar"></div>
      <div id="canvas"><div id="editor"></div></div>
      <div id="statusbar"></div>`;
    const bare = {
      openFile: vi.fn().mockResolvedValue(null),
      saveDoc: vi.fn().mockResolvedValue(null),
      saveRaw: vi.fn().mockResolvedValue(null),
      pickImage: vi.fn().mockResolvedValue(null),
      exportPdf: vi.fn().mockResolvedValue(null),
      print: vi.fn().mockResolvedValue(undefined),
      getRecent: vi.fn().mockResolvedValue([]),
      openExternal: vi.fn().mockResolvedValue(true),
      onCommand: vi.fn(() => () => {}),
      sendQuitResponse: vi.fn(),
      onWillQuit: vi.fn(() => () => {}),
      onSaveThenQuit: vi.fn(() => () => {}),
    };
    Object.defineProperty(window, "fw", { value: bare, configurable: true, writable: true });
    Object.defineProperty(document, "fonts", { value: { ready: Promise.resolve() }, configurable: true });
    const { initApp } = await import("../src/renderer/src/app");
    initApp();
    await new Promise((r) => setTimeout(r, 240));
    expect(document.body.classList.contains("notion-mode")).toBe(false);
    expect(document.querySelector("#statusbar")!.textContent).toContain("Page 1 / 1");
  });
});
