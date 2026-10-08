import { dialog, ipcMain, shell } from "electron";
import type { BrowserWindow } from "electron";
import { readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import type { AppSettings, PrintResult, SaveRawPayload, WriteFileResult } from "../shared/types";
import { DEFAULT_SETTINGS } from "../shared/types";
import { safeExternalUrl } from "../shared/urls";
import { toDocxBuffer } from "./docx-export";
import { readRecent, upsertRecent } from "./recent";
import { readSettings, setSettings as writeSettings } from "./settings";
import {
  createPage,
  deletePage,
  deleteTrash,
  listPages,
  listTrash,
  movePage,
  readPage,
  renamePage,
  restoreTrash,
  writePage,
} from "./workspace";

type GetWindow = () => BrowserWindow | null;

const SUPPORTED_EXT = ["docx", "txt", "html", "htm"];

/** Formats that `file:saveRaw` may write — the renderer cannot name its own extension. */
const RAW_FORMATS: Record<SaveRawPayload["ext"], { label: string; extensions: string[] }> = {
  txt: { label: "Plain Text", extensions: ["txt"] },
  html: { label: "HTML", extensions: ["html", "htm"] },
};

/** Anything larger is a mistake or an attack, not a document: reading it would hang the renderer. */
const MAX_OPEN_BYTES = 100 * 1024 * 1024;

function sleep(ms: number): Promise<void> {
  const { promise, resolve } = Promise.withResolvers<void>();
  setTimeout(resolve, ms);
  return promise;
}

/** A short, user-presentable description of a caught error. */
function messageOf(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (typeof err === "string") return err;
  return String(err);
}

/**
 * Only the app window's own main frame may drive these handlers.
 *
 * Without this, any frame that ever appears in the renderer (now or after a
 * future change) can ask for arbitrary documents, PDFs and image picks. The
 * check fails closed: no window, or a frame that is not that window's main
 * frame, gets nothing — the same answers the user gets for "dismissed".
 */
function trustedSender(event: { senderFrame: unknown }, getWindow: GetWindow): boolean {
  const win = getWindow();
  if (!win || win.isDestroyed()) return false;
  const mainFrame = (win.webContents as { mainFrame?: unknown }).mainFrame;
  if (!mainFrame) return false;
  return event.senderFrame === mainFrame;
}

/** Path equality as the OS sees it — Recents are keyed by the path we recorded. */
function samePath(a: string, b: string): boolean {
  const na = path.resolve(a);
  const nb = path.resolve(b);
  return process.platform === "linux" ? na === nb : na.toLowerCase() === nb.toLowerCase();
}

export function registerIpc(getWindow: GetWindow, onFilesChanged: () => void): void {
  ipcMain.handle("file:open", async (event, requestedPath?: string) => {
    if (!trustedSender(event, getWindow)) return null;
    const win = getWindow();
    if (!win) return null;

    let filePath: string | undefined;
    if (typeof requestedPath === "string" && requestedPath) {
      // An explicit path only counts when the user already opened or saved it
      // through this dialog — i.e. it is in the recent store. That is what
      // makes the menu's Recent entries safe: a renderer (or a bug, or a
      // compromised page) cannot name an arbitrary absolute path and have the
      // main process read it back. Anything else falls back to the dialog, so
      // a stale/removed entry behaves exactly like a fresh File → Open.
      const recents = await readRecent();
      if (recents.some((r) => samePath(r.path, requestedPath))) filePath = requestedPath;
    }
    if (!filePath) {
      const result = await dialog.showOpenDialog(win, {
        properties: ["openFile"],
        filters: [
          { name: "Word Document", extensions: ["docx"] },
          { name: "Text", extensions: ["txt"] },
          { name: "HTML", extensions: ["html", "htm"] },
          { name: "All Files", extensions: ["*"] },
        ],
      });
      if (result.canceled || result.filePaths.length === 0) return null;
      filePath = result.filePaths[0];
    }

    const ext = path.extname(filePath).toLowerCase().replace(/^\./, "");
    if (ext === "doc") {
      dialog.showErrorBox(
        "Unsupported format",
        "Legacy .doc files are not supported. Please save the file as .docx.",
      );
      return null;
    }
    if (!SUPPORTED_EXT.includes(ext)) {
      dialog.showErrorBox("Unsupported format", `Files with the extension .${ext} are not supported.`);
      return null;
    }
    const buffer = await readFile(filePath).catch((err: unknown) => {
      dialog.showErrorBox(
        "Could not open file",
        `"${path.basename(filePath!)}" could not be read:\n${messageOf(err)}`,
      );
      return null;
    });
    if (!buffer) return null;
    if (buffer.byteLength > MAX_OPEN_BYTES) {
      dialog.showErrorBox(
        "File too large",
        `"${path.basename(filePath)}" is ${Math.round(buffer.byteLength / (1024 * 1024))} MB. ` +
          `Documents larger than ${Math.round(MAX_OPEN_BYTES / (1024 * 1024))} MB are not supported.`,
      );
      return null;
    }
    const st = await stat(filePath).catch(() => null);
    const name = path.basename(filePath);
    if (st) await upsertRecent({ path: filePath, name, mtimeMs: st.mtimeMs });
    onFilesChanged();
    return {
      path: filePath,
      name,
      ext: (ext === "htm" ? "html" : ext) as "docx" | "txt" | "html",
      buffer: new Uint8Array(buffer),
    };
  });

  ipcMain.handle(
    "docx:save",
    async (event, payload: { json: Record<string, unknown>; defaultName: string; path?: string }) => {
      if (!trustedSender(event, getWindow)) return null;
      const win = getWindow();
      if (!win) return null;
      let filePath = payload.path;
      if (!filePath) {
        const result = await dialog.showSaveDialog(win, {
          defaultPath: payload.defaultName,
          filters: [{ name: "Word Document", extensions: ["docx"] }],
        });
        if (result.canceled || !result.filePath) return null;
        filePath = result.filePath;
      }
      const buf = await toDocxBuffer(payload.json).catch((err: unknown) => {
        dialog.showErrorBox(
          "Could not save document",
          `"${path.basename(filePath!)}" could not be produced:\n${messageOf(err)}`,
        );
        return null;
      });
      if (!buf) return { error: "The document could not be converted to .docx." };
      try {
        await writeFile(filePath, buf);
      } catch (err) {
        // The usual causes on Windows: the target is open in Word (EBUSY),
        // read-only / protected location (EACCES), or a full disk (ENOSPC).
        dialog.showErrorBox(
          "Could not save document",
          `"${filePath}" could not be written:\n${messageOf(err)}`,
        );
        return { error: messageOf(err) };
      }
      const st = await stat(filePath).catch(() => null);
      await upsertRecent({
        path: filePath,
        name: path.basename(filePath),
        mtimeMs: st?.mtimeMs ?? Date.now(),
      });
      onFilesChanged();
      return { path: filePath };
    },
  );

  // Plain-text formats (.txt / .html): the renderer serializes the content, the
  // main process only ever writes bytes to a path the user picked.
  ipcMain.handle("file:saveRaw", async (event, payload: SaveRawPayload) => {
    if (!trustedSender(event, getWindow)) return null;
    const win = getWindow();
    if (!win) return null;
    const format = RAW_FORMATS[payload?.ext as SaveRawPayload["ext"]];
    if (!format) {
      dialog.showErrorBox("Unsupported format", `Cannot export to .${String(payload?.ext)}.`);
      return { error: `Unsupported export format: ${String(payload?.ext)}` };
    }
    if (typeof payload.content !== "string") return { error: "Nothing to write." };

    let filePath = payload.path;
    if (!filePath) {
      const result = await dialog.showSaveDialog(win, {
        defaultPath: payload.defaultName,
        filters: [{ name: format.label, extensions: format.extensions }],
      });
      if (result.canceled || !result.filePath) return null;
      filePath = result.filePath;
    }
    try {
      await writeFile(filePath, payload.content, "utf8");
    } catch (err) {
      dialog.showErrorBox("Could not export", `"${filePath}" could not be written:\n${messageOf(err)}`);
      return { error: messageOf(err) };
    }
    const st = await stat(filePath).catch(() => null);
    await upsertRecent({ path: filePath, name: path.basename(filePath), mtimeMs: st?.mtimeMs ?? Date.now() });
    onFilesChanged();
    return { path: filePath };
  });

  ipcMain.handle("pdf:export", async (event, defaultName: string): Promise<WriteFileResult | null> => {
    if (!trustedSender(event, getWindow)) return null;
    const win = getWindow();
    if (!win) return null;
    // Ask first: rendering the PDF is the expensive part, and paying it before
    // the user commits to a path meant every cancelled export did the work.
    const result = await dialog.showSaveDialog(win, {
      defaultPath: defaultName,
      filters: [{ name: "PDF Document", extensions: ["pdf"] }],
    });
    if (result.canceled || !result.filePath) return null;
    // Give the renderer a moment to settle print mode.
    await sleep(100);
    let data: Uint8Array;
    try {
      data = await win.webContents.printToPDF({
        printBackground: true,
        preferCSSPageSize: true,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
      });
    } catch (err) {
      dialog.showErrorBox("Could not export PDF", `The document could not be rendered:\n${messageOf(err)}`);
      return { error: messageOf(err) };
    }
    try {
      await writeFile(result.filePath, data);
    } catch (err) {
      dialog.showErrorBox(
        "Could not export PDF",
        `"${result.filePath}" could not be written:\n${messageOf(err)}`,
      );
      return { error: messageOf(err) };
    }
    return { path: result.filePath };
  });

  ipcMain.handle("print", async (event): Promise<PrintResult | null> => {
    if (!trustedSender(event, getWindow)) return null;
    const win = getWindow();
    if (!win) return null;
    await sleep(100);
    try {
      // The system print dialog handles the user's choice (print or cancel).
      await new Promise<void>((resolve) => {
        win.webContents.print({ silent: false, printBackground: true }, () => resolve());
      });
    } catch (err) {
      dialog.showErrorBox("Could not print", `Printing failed:\n${messageOf(err)}`);
      return { error: messageOf(err) };
    }
    return {};
  });

  ipcMain.handle("file:pickImage", async (event) => {
    if (!trustedSender(event, getWindow)) return null;
    const win = getWindow();
    if (!win) return null;
    const result = await dialog.showOpenDialog(win, {
      properties: ["openFile"],
      filters: [
        {
          name: "Images",
          extensions: ["png", "jpg", "jpeg", "gif", "bmp"],
        },
      ],
    });
    if (result.canceled || result.filePaths.length === 0) return null;
    const filePath = result.filePaths[0];
    const buffer = await readFile(filePath).catch((err: unknown) => {
      dialog.showErrorBox("Could not read image", `"${path.basename(filePath)}":\n${messageOf(err)}`);
      return null;
    });
    if (!buffer) return null;
    const MAX_IMAGE_BYTES = 10 * 1024 * 1024;
    if (buffer.byteLength > MAX_IMAGE_BYTES) {
      dialog.showErrorBox("Image too large", "Images larger than 10 MB are not supported.");
      return null;
    }
    const ext = path.extname(filePath).toLowerCase().replace(/^\./, "");
    const MIME_BY_EXT: Record<string, string> = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp",
    };
    const mime = MIME_BY_EXT[ext];
    if (!mime) {
      dialog.showErrorBox("Unsupported format", "Images must be PNG, JPEG, GIF, or BMP.");
      return null;
    }
    return {
      dataUrl: `data:${mime};base64,${buffer.toString("base64")}`,
      name: path.basename(filePath),
    };
  });

  ipcMain.handle("shell:openExternal", async (event, url: unknown) => {
    if (!trustedSender(event, getWindow)) return false;
    const safe = safeExternalUrl(url);
    if (!safe) {
      console.warn("[FreeWord] refused to open an external URL with a disallowed scheme");
      return false;
    }
    try {
      // Electron's `openExternal` resolves with void; the boolean is our own
      // "the URL passed policy and the shell accepted it" answer.
      await shell.openExternal(safe);
      return true;
    } catch (err) {
      dialog.showErrorBox("Could not open link", `${safe}\n${messageOf(err)}`);
      return false;
    }
  });

  ipcMain.handle("recent:get", () => readRecent());

  // Settings (Notion mode / theme / sidebar). The patch is filtered
  // key-by-key in settings.ts, so the renderer cannot smuggle in new keys.
  ipcMain.handle("settings:get", async (event): Promise<AppSettings> => {
    // Untrusted frames get the defaults, never the user's actual settings.
    if (!trustedSender(event, getWindow)) return { ...DEFAULT_SETTINGS };
    return readSettings();
  });
  ipcMain.handle("settings:set", async (event, patch: Partial<AppSettings>) => {
    if (!trustedSender(event, getWindow)) return { ...DEFAULT_SETTINGS };
    const safe: Partial<AppSettings> = {};
    if (patch && typeof patch === "object") {
      if (typeof patch.notionMode === "boolean") safe.notionMode = patch.notionMode;
      if (typeof patch.sidebarOpen === "boolean") safe.sidebarOpen = patch.sidebarOpen;
      if (typeof patch.toolbarPinned === "boolean") safe.toolbarPinned = patch.toolbarPinned;
      if (patch.theme === "light" || patch.theme === "dark" || patch.theme === "system") {
        safe.theme = patch.theme;
      }
    }
    const next = await writeSettings(safe);
    // The View menu mirrors notionMode/theme as checkbox/radio items.
    onFilesChanged();
    return next;
  });

  // Workspace pages. Every id crossing the boundary is re-validated in
  // workspace.ts (path-traversal guard) and every handler is main-frame-only.
  ipcMain.handle("page:list", async (event) => {
    if (!trustedSender(event, getWindow)) return [];
    return listPages();
  });
  ipcMain.handle("page:read", async (event, id: unknown) => {
    if (!trustedSender(event, getWindow) || typeof id !== "string") return null;
    return readPage(id);
  });
  ipcMain.handle("page:write", async (event, id: unknown, json: unknown) => {
    if (!trustedSender(event, getWindow)) return { error: "Untrusted sender." };
    if (typeof id !== "string") return { error: "Invalid page id." };
    return writePage(id, json);
  });
  ipcMain.handle("page:create", async (event, opts?: { title?: string; parentId?: string | null }) => {
    if (!trustedSender(event, getWindow)) return null;
    const title = typeof opts?.title === "string" ? opts.title.slice(0, 200) : undefined;
    const parentId = typeof opts?.parentId === "string" ? opts.parentId : null;
    return createPage({ title, parentId });
  });
  ipcMain.handle("page:rename", async (event, id: unknown, patch?: { title?: string; icon?: string }) => {
    if (!trustedSender(event, getWindow) || typeof id !== "string") return null;
    return renamePage(id, patch ?? {});
  });
  ipcMain.handle("page:move", async (event, id: unknown, parentId: unknown, order: unknown) => {
    if (!trustedSender(event, getWindow) || typeof id !== "string") return { error: "Untrusted sender." };
    return movePage(id, typeof parentId === "string" ? parentId : null, Number(order) || 0);
  });
  ipcMain.handle("page:delete", async (event, id: unknown) => {
    if (!trustedSender(event, getWindow) || typeof id !== "string") return { error: "Untrusted sender." };
    return deletePage(id);
  });
  ipcMain.handle("page:trashList", async (event) => {
    if (!trustedSender(event, getWindow)) return [];
    return listTrash();
  });
  ipcMain.handle("page:trashRestore", async (event, file: unknown) => {
    if (!trustedSender(event, getWindow) || typeof file !== "string") return { error: "Untrusted sender." };
    return restoreTrash(file);
  });
  ipcMain.handle("page:trashDelete", async (event, file: unknown) => {
    if (!trustedSender(event, getWindow) || typeof file !== "string") return { error: "Untrusted sender." };
    return deleteTrash(file);
  });
}
