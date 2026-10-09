/**
 * Main-process e2e: drives the real IPC handlers with a mocked Electron
 * (dialog + window) and the real filesystem in a temp dir.
 */
import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

// ---------------------------------------------------------------------------
// Electron mock
// ---------------------------------------------------------------------------
const handlers = new Map<string, (...args: unknown[]) => unknown>();
const dialogMock = {
  showOpenDialog: vi.fn(),
  showSaveDialog: vi.fn(),
  showErrorBox: vi.fn(),
  showMessageBox: vi.fn(),
};
let userDataDir = "";
// The main process only answers handlers called by *its own* main frame, so the
// fake event has to carry the same object the fake window reports.
const MAIN_FRAME = { id: "main-frame" };
const ipcEvent = { senderFrame: MAIN_FRAME };
/** A frame from anywhere else: an iframe, another window, an attacker. */
const ROUGE_FRAME = { id: "other-frame" };

const shellMock = {
  openExternal: vi.fn().mockResolvedValue(undefined),
};

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
  dialog: dialogMock,
  shell: shellMock,
  ipcMain: {
    handle: vi.fn((channel: string, fn: (...a: unknown[]) => unknown) => handlers.set(channel, fn)),
    on: vi.fn(),
  },
  Menu: { setApplicationMenu: vi.fn(), buildFromTemplate: vi.fn((t: unknown) => t) },
}));

const { registerIpc } = await import("../src/main/ipc");
const { readRecent, upsertRecent } = await import("../src/main/recent");

let win: Record<string, unknown>;
const makeWin = () => ({
  isDestroyed: () => false,
  webContents: {
    isDestroyed: () => false,
    mainFrame: MAIN_FRAME,
    printToPDF: vi.fn().mockResolvedValue(Buffer.from("%PDF-1.4 fake")),
    print: vi.fn((_opts: unknown, cb: () => void) => cb()),
  },
});

const call = (channel: string, ...args: unknown[]) => handlers.get(channel)!(ipcEvent, ...args);
const callFrom = (frame: unknown, channel: string, ...args: unknown[]) =>
  handlers.get(channel)!({ senderFrame: frame }, ...args);

beforeAll(async () => {
  userDataDir = await mkdtemp(join(tmpdir(), "fw-userdata-"));
});

beforeEach(async () => {
  handlers.clear();
  vi.clearAllMocks();
  win = makeWin();
  registerIpc(
    () => win as never,
    () => {},
  );
  await writeFile(join(userDataDir, "recent.json"), "[]");
});

describe("file:open", () => {
  it("returns null when the dialog is cancelled", async () => {
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: true, filePaths: [] });
    await expect(call("file:open")).resolves.toBeNull();
  });

  it("rejects legacy .doc with a dedicated error", async () => {
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: ["C:\\a\\old.doc"] });
    await expect(call("file:open")).resolves.toBeNull();
    expect(dialogMock.showErrorBox).toHaveBeenCalledWith(
      "Unsupported format",
      expect.stringContaining(".doc"),
    );
  });

  it("rejects unsupported extensions", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-open-"));
    const p = join(dir, "thing.pdf");
    await writeFile(p, "x");
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [p] });
    await expect(call("file:open")).resolves.toBeNull();
    expect(dialogMock.showErrorBox).toHaveBeenCalled();
  });

  it("opens a txt file, normalizes buffer + records it in recents", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-open-"));
    const p = join(dir, "note.txt");
    await writeFile(p, "hi there");
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [p] });
    const res = (await call("file:open")) as Record<string, unknown>;
    expect(res.ext).toBe("txt");
    expect(res.name).toBe("note.txt");
    expect(res.buffer).toBeInstanceOf(Uint8Array);
    expect(new TextDecoder().decode(res.buffer as Uint8Array)).toBe("hi there");
    const recents = await readRecent();
    expect(recents.map((r) => r.path)).toContain(p);
  });

  it("maps .htm to the html importer", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-open-"));
    const p = join(dir, "page.htm");
    await writeFile(p, "<p>x</p>");
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [p] });
    const res = (await call("file:open")) as { ext: string };
    expect(res.ext).toBe("html");
  });

  it("opening a deleted recent file shows a native error box and returns null", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-open-"));
    const p = join(dir, "was-here.txt");
    await writeFile(p, "content");
    // Only paths the user already opened/saved are honoured (recent store), so
    // the entry has to exist before the file disappears.
    await upsertRecent({ path: p, name: "was-here.txt", mtimeMs: 1 });
    const { rm } = await import("node:fs/promises");
    await rm(p);
    const res = await call("file:open", p);
    expect(res).toBeNull();
    expect(dialogMock.showErrorBox).toHaveBeenCalledWith(
      "Could not open file",
      expect.stringContaining("was-here.txt"),
    );
  });

  it("ignores an explicit path that is not in the recent store (no arbitrary read)", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-open-"));
    const secret = join(dir, "secret.txt");
    await writeFile(secret, "do not read me");
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: true, filePaths: [] });
    // A renderer naming an absolute path must not have it read back: the
    // handler drops the path and asks the user instead.
    const res = await call("file:open", secret);
    expect(res).toBeNull();
    expect(dialogMock.showOpenDialog).toHaveBeenCalled();
    expect(dialogMock.showErrorBox).not.toHaveBeenCalled();
  });

  it("rejects a call from a frame that is not the app window's main frame", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-open-"));
    const p = join(dir, "note2.txt");
    await writeFile(p, "hi");
    await upsertRecent({ path: p, name: "note2.txt", mtimeMs: 1 });
    await expect(callFrom(ROUGE_FRAME, "file:open", p)).resolves.toBeNull();
    await expect(callFrom(null, "file:open", p)).resolves.toBeNull();
    // The legitimate frame still works.
    const ok = (await call("file:open", p)) as { path: string };
    expect(ok.path).toBe(p);
  });

  it("refuses documents far too large to open", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-open-"));
    const p = join(dir, "huge.txt");
    // Sparse write: only the last byte forces the size, no 100 MB buffer.
    const { open } = await import("node:fs/promises");
    const handle = await open(p, "w");
    await handle.truncate(100 * 1024 * 1024 + 1024);
    await handle.close();
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [p] });
    await expect(call("file:open")).resolves.toBeNull();
    expect(dialogMock.showErrorBox).toHaveBeenCalledWith("File too large", expect.any(String));
  });

  it("surfaces docx write failures on save instead of hanging", async () => {
    const res = (await call("docx:save", {
      json: { type: "doc", content: [{ type: "page", content: [{ type: "paragraph" }] }] },
      defaultName: "broken.docx",
      path: "C:\\nonexistent-dir-zz\\broken.docx",
    })) as { error?: string; path?: string };
    expect(res?.error).toBeTruthy(); // { error } — renderer stays in control
    expect(dialogMock.showErrorBox).toHaveBeenCalled();
  });
});

describe("docx:save", () => {
  it("cancelling the dialog saves nothing", async () => {
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: true, filePath: "" });
    await expect(
      call("docx:save", { json: { type: "doc", content: [] }, defaultName: "a.docx" }),
    ).resolves.toBeNull();
  });

  it("writes a real zip (.docx) and records recents", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-save-"));
    const p = join(dir, "out.docx");
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: false, filePath: p });
    const json = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        { type: "page", content: [{ type: "paragraph", content: [{ type: "text", text: "saved" }] }] },
      ],
    };
    const res = (await call("docx:save", { json, defaultName: "out.docx" })) as { path: string };
    expect(res.path).toBe(p);
    const bytes = await readFile(p);
    expect(bytes.subarray(0, 2).toString()).toBe("PK"); // zip magic
    const recents = await readRecent();
    expect(recents[0].path).toBe(p);
  });

  it("Ctrl+S overwrites the known path without a dialog", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-save-"));
    const p = join(dir, "quiet.docx");
    const json = { type: "doc", content: [{ type: "page", content: [{ type: "paragraph" }] }] };
    // First save goes through the dialog — that is how the path becomes one
    // main "knows" and the renderer is later allowed to name directly.
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: false, filePath: p });
    const first = (await call("docx:save", { json, defaultName: "quiet.docx" })) as { path: string };
    expect(first.path).toBe(p);
    dialogMock.showSaveDialog.mockClear();
    const res = (await call("docx:save", { json, defaultName: "quiet.docx", path: p })) as { path: string };
    expect(res.path).toBe(p);
    expect(dialogMock.showSaveDialog).not.toHaveBeenCalled();
  });

  it("refuses a renderer-invented path that never came from a dialog", async () => {
    const json = { type: "doc", content: [{ type: "page", content: [{ type: "paragraph" }] }] };
    const res = (await call("docx:save", {
      json,
      defaultName: "evil.docx",
      path: join(tmpdir(), "fw-invented-42", "startup.docx"),
    })) as { error?: string };
    expect(res.error).toMatch(/Unrecognized save path/);
  });
});

describe("file:pickImage", () => {
  it("returns a data URL with the right mime", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-img-"));
    const p = join(dir, "pic.png");
    const png = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
      "base64",
    );
    await writeFile(p, png);
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [p] });
    const res = (await call("file:pickImage")) as { dataUrl: string; name: string };
    expect(res.dataUrl.startsWith("data:image/png;base64,")).toBe(true);
    expect(res.name).toBe("pic.png");
  });

  it("rejects images over 10 MB", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-img-"));
    const p = join(dir, "big.bmp");
    await writeFile(p, Buffer.alloc(10 * 1024 * 1024 + 1, 1));
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [p] });
    await expect(call("file:pickImage")).resolves.toBeNull();
    expect(dialogMock.showErrorBox).toHaveBeenCalledWith("Image too large", expect.any(String));
  });

  it("rejects unknown image extensions", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-img-"));
    const p = join(dir, "weird.webp");
    await writeFile(p, Buffer.from("RIFFxxxxWEBP"));
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [p] });
    await expect(call("file:pickImage")).resolves.toBeNull();
    expect(dialogMock.showErrorBox).toHaveBeenCalled();
  });
});

describe("pdf:export", () => {
  it("captures with CSS page size and zero margins", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-pdf-"));
    const p = join(dir, "doc.pdf");
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: false, filePath: p });
    const res = (await call("pdf:export", "doc.pdf")) as { path: string };
    expect(res.path).toBe(p);
    const opts = (win.webContents as unknown as { printToPDF: ReturnType<typeof vi.fn> }).printToPDF.mock
      .calls[0][0];
    expect(opts.preferCSSPageSize).toBe(true);
    expect(opts.printBackground).toBe(true);
    expect(opts.margins).toEqual({ top: 0, bottom: 0, left: 0, right: 0 });
    expect((await readFile(p)).subarray(0, 5).toString()).toBe("%PDF-");
  });

  it("returns null when cancelled — and never renders the PDF", async () => {
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: true, filePath: "" });
    await expect(call("pdf:export", "x.pdf")).resolves.toBeNull();
    // The old order paid the full printToPDF cost before the dialog, so every
    // cancelled export rendered a document nobody asked for.
    const printToPDF = (win.webContents as unknown as { printToPDF: ReturnType<typeof vi.fn> }).printToPDF;
    expect(printToPDF).not.toHaveBeenCalled();
  });

  it("reports a render failure as { error } instead of rejecting", async () => {
    const printToPDF = (win.webContents as unknown as { printToPDF: ReturnType<typeof vi.fn> }).printToPDF;
    printToPDF.mockRejectedValue(new Error("print engine unavailable"));
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: false, filePath: join(tmpdir(), "x.pdf") });
    const res = (await call("pdf:export", "x.pdf")) as { error?: string };
    expect(res.error).toMatch(/print engine unavailable/);
    expect(dialogMock.showErrorBox).toHaveBeenCalledWith("Could not export PDF", expect.any(String));
  });

  it("reports a write failure (locked/full target) as { error }", async () => {
    dialogMock.showSaveDialog.mockResolvedValue({
      canceled: false,
      filePath: "C:\\nonexistent-dir-zz\\out.pdf",
    });
    const res = (await call("pdf:export", "out.pdf")) as { error?: string };
    expect(res.error).toBeTruthy();
    expect(dialogMock.showErrorBox).toHaveBeenCalledWith("Could not export PDF", expect.any(String));
  });
});

describe("file:saveRaw", () => {
  it("writes plain text and records it in recents", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-raw-"));
    const p = join(dir, "note.txt");
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: false, filePath: p });
    const res = (await call("file:saveRaw", {
      content: "line one\nline two",
      defaultName: "note.txt",
      ext: "txt",
    })) as { path: string };
    expect(res.path).toBe(p);
    expect(await readFile(p, "utf8")).toBe("line one\nline two");
    expect((await readRecent()).map((r) => r.path)).toContain(p);
    // showSaveDialog is called as (window, options).
    expect(dialogMock.showSaveDialog.mock.calls[0][1].filters[0].extensions).toEqual(["txt"]);
  });

  it("offers an html filter for the html format", async () => {
    const dir = await mkdtemp(join(tmpdir(), "fw-raw-"));
    const p = join(dir, "page.html");
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: false, filePath: p });
    const res = (await call("file:saveRaw", {
      content: "<p>hi</p>",
      defaultName: "page.html",
      ext: "html",
    })) as { path: string };
    expect(res.path).toBe(p);
    expect(dialogMock.showSaveDialog.mock.calls[0][1].filters[0].extensions).toEqual(["html", "htm"]);
  });

  it("refuses an extension outside the allowlist", async () => {
    const res = (await call("file:saveRaw", {
      content: "x",
      defaultName: "evil.exe",
      ext: "exe",
    })) as { error?: string };
    expect(res.error).toMatch(/Unsupported export format/);
    expect(dialogMock.showErrorBox).toHaveBeenCalled();
  });

  it("returns null when the dialog is cancelled", async () => {
    dialogMock.showSaveDialog.mockResolvedValue({ canceled: true, filePath: "" });
    await expect(
      call("file:saveRaw", { content: "x", defaultName: "a.txt", ext: "txt" }),
    ).resolves.toBeNull();
  });
});

describe("shell:openExternal", () => {
  it("opens http/https/mailto/tel links", async () => {
    for (const url of ["https://example.com", "http://example.com", "mailto:a@b.c", "tel:+15551234"]) {
      shellMock.openExternal.mockClear();
      await expect(call("shell:openExternal", url)).resolves.toBe(true);
      expect(shellMock.openExternal).toHaveBeenCalledWith(url);
    }
  });

  it("refuses javascript:, data:, file: and garbage", async () => {
    for (const url of [
      "javascript:alert(1)",
      "JavaScript:alert(1)",
      "data:text/html;base64,PHNjcmlwdD4=",
      "file:///C:/Windows/win.ini",
      "",
      "#top",
      "not a url",
      42,
      null,
    ]) {
      shellMock.openExternal.mockClear();
      await expect(call("shell:openExternal", url)).resolves.toBe(false);
      expect(shellMock.openExternal).not.toHaveBeenCalled();
    }
  });

  it("ignores calls from a foreign frame", async () => {
    shellMock.openExternal.mockClear();
    await expect(callFrom(ROUGE_FRAME, "shell:openExternal", "https://example.com")).resolves.toBe(false);
    expect(shellMock.openExternal).not.toHaveBeenCalled();
  });
});

describe("print", () => {
  it("invokes the system print dialog (non-silent)", async () => {
    await call("print");
    const print = (win.webContents as unknown as { print: ReturnType<typeof vi.fn> }).print;
    expect(print).toHaveBeenCalledWith(
      expect.objectContaining({ silent: false, printBackground: true }),
      expect.any(Function),
    );
  });

  // The handler must not resolve before Chromium's print callback: the
  // renderer tears print mode down (toolbar back, re-pagination, zoom
  // restored) as soon as this resolves, and doing that mid-capture puts the
  // chrome into the printed page.
  it("handler resolves only after the print job finishes (no print-mode teardown race)", async () => {
    const print = (win.webContents as unknown as { print: ReturnType<typeof vi.fn> }).print;
    let jobDone = false;
    print.mockImplementation((_o: unknown, cb: () => void) => {
      setTimeout(() => {
        jobDone = true;
        cb();
      }, 250);
    });
    await call("print");
    expect(jobDone).toBe(true);
  });

  it("reports a print failure as { error } instead of rejecting", async () => {
    const print = (win.webContents as unknown as { print: ReturnType<typeof vi.fn> }).print;
    print.mockImplementation(() => {
      throw new Error("no printer available");
    });
    const res = (await call("print")) as { error?: string };
    expect(res.error).toMatch(/no printer available/);
    expect(dialogMock.showErrorBox).toHaveBeenCalledWith("Could not print", expect.any(String));
  });
});

describe("recent store", () => {
  it("dedupes by path, orders by activity, caps at 8, survives corruption", async () => {
    for (let i = 0; i < 10; i++) {
      await upsertRecent({ path: `C:\\f\\f${i}.docx`, name: `f${i}.docx`, mtimeMs: 0 });
    }
    let list = await readRecent();
    expect(list.length).toBe(8);
    expect(list[0].path).toBe("C:\\f\\f9.docx"); // most recent first

    await upsertRecent({ path: "C:\\f\\f0.docx", name: "f0.docx", mtimeMs: 0 }); // re-open oldest
    list = await readRecent();
    expect(list.filter((e) => e.path === "C:\\f\\f0.docx")).toHaveLength(1);
    expect(list[0].path).toBe("C:\\f\\f0.docx");

    // corrupt file → empty, never throws
    await writeFile(join(userDataDir, "recent.json"), "{not json");
    await expect(readRecent()).resolves.toEqual([]);
    await writeFile(join(userDataDir, "recent.json"), '[{"path":123}]'); // invalid entries
    await expect(readRecent()).resolves.toEqual([]);
  });
});

describe("application menu", () => {
  it("Edit menu uses native undo/redo roles (shadowing TipTap history — see report)", async () => {
    const menu = await import("../src/main/menu");
    menu.configureMenu(
      () => win as never,
      () => {},
    );
    await menu.rebuildMenu();
    const electronMenu = await import("electron");
    const build = (electronMenu as unknown as { Menu: { buildFromTemplate: ReturnType<typeof vi.fn> } }).Menu
      .buildFromTemplate;
    expect(build).toHaveBeenCalled();
    const template = build.mock.calls[build.mock.calls.length - 1][0] as Array<{
      label: string;
      submenu?: Array<{ label?: string; role?: string; accelerator?: string }>;
    }>;
    const edit = template.find((t) => t.label === "Edit");
    const roles = (edit?.submenu ?? []).map((i) => i.role).filter(Boolean);
    // Documents the current design: native roles capture Ctrl+Z / Ctrl+Y
    // before the page, bypassing ProseMirror's history (see bug report §M4).
    expect(roles).toEqual(expect.arrayContaining(["undo", "redo"]));
    const file = template.find((t) => t.label === "File");
    const labels = (file?.submenu ?? []).map((i) => i.label);
    expect(labels).toEqual(
      expect.arrayContaining([
        "New",
        "Open…",
        "Save",
        "Save As…",
        "Export PDF…",
        // Plain-text formats reach users who never see a "Save As" dialog.
        "Save As Plain Text…",
        "Save As HTML…",
        "Print…",
        "Recent",
        "Quit",
      ]),
    );
  });
});

describe("context menu", () => {
  interface Item {
    role?: string;
    label?: string;
    enabled?: boolean;
    type?: string;
  }

  /** `Menu.buildFromTemplate` is mocked as the identity, so read the template. */
  async function contextMenuFor(params: Record<string, unknown>): Promise<Item[]> {
    const menu = await import("../src/main/menu");
    menu.buildContextMenu(win as never, params as never);
    const electron = await import("electron");
    const build = (electron as unknown as { Menu: { buildFromTemplate: ReturnType<typeof vi.fn> } }).Menu
      .buildFromTemplate;
    return build.mock.calls[build.mock.calls.length - 1][0] as Item[];
  }

  const editable = {
    isEditable: true,
    selectionText: "",
    editFlags: { isEditable: true, canCut: true, canCopy: true, canPaste: true, canSelectAll: true },
  };

  it("offers edit roles and the spell-checker's suggestions for a misspelled word", async () => {
    const items = await contextMenuFor({
      ...editable,
      misspelledWord: "teh",
      dictionarySuggestions: ["the", "tea", "ten"],
    });
    const roles = items.map((i) => i.role);
    expect(roles).toEqual(expect.arrayContaining(["cut", "copy", "paste", "selectAll"]));
    const labels = items.map((i) => i.label);
    expect(labels).toEqual(expect.arrayContaining(["the", "tea", "ten"]));
    // Suggestions are useless without the word they replace.
    expect(labels.some((l) => typeof l === "string" && l.includes("teh"))).toBe(true);
  });

  it("caps the suggestion list rather than dumping the dictionary", async () => {
    const many = Array.from({ length: 30 }, (_, i) => `word${i}`);
    const items = await contextMenuFor({ ...editable, misspelledWord: "xyzzy", dictionarySuggestions: many });
    const suggested = items.filter((i) => typeof i.label === "string" && i.label?.startsWith("word"));
    expect(suggested.length).toBeGreaterThan(0);
    expect(suggested.length).toBeLessThanOrEqual(6);
  });

  it("says so when there are no suggestions, instead of showing an empty group", async () => {
    const items = await contextMenuFor({
      ...editable,
      misspelledWord: "qwertyuiop",
      dictionarySuggestions: [],
    });
    expect(items.some((i) => i.label === "No suggestions" && i.enabled === false)).toBe(true);
  });

  it("offers copy only, with nothing editable", async () => {
    const items = await contextMenuFor({ isEditable: false, selectionText: "selected words", editFlags: {} });
    expect(items.filter((i) => !i.type).map((i) => i.role)).toEqual(["copy"]);
  });

  it("opens a link only when its scheme is one the app trusts", async () => {
    const safe = await contextMenuFor({
      isEditable: false,
      selectionText: "",
      editFlags: {},
      linkURL: "https://example.com/",
    });
    const unsafe = await contextMenuFor({
      isEditable: false,
      selectionText: "",
      editFlags: {},
      linkURL: "javascript:alert(1)",
    });
    expect(safe.map((i) => i.label)).toContain("Open Link in Browser");
    expect(unsafe.map((i) => i.label)).not.toContain("Open Link in Browser");
  });
});
