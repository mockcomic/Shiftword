/** Shared types used by main, preload, and renderer. */

/** A recently-opened/saved file entry. */
export interface RecentEntry {
  path: string;
  name: string;
  mtimeMs: number;
}

/** Supported page sizes in millimetres. */
export const PAGE_SIZES = {
  a4: { w: 210, h: 297 },
  letter: { w: 215.9, h: 279.4 },
} as const;

export type PageSize = keyof typeof PAGE_SIZES;

/** Page settings stored on the custom `doc` node attrs. */
export interface PageSettings {
  size: PageSize;
  mt: number; // mm, top margin
  mr: number; // mm, right margin
  mb: number; // mm, bottom margin
  ml: number; // mm, left margin
}

export const DEFAULT_PAGE_SETTINGS: PageSettings = {
  size: "a4",
  mt: 25.4,
  mr: 25.4,
  mb: 25.4,
  ml: 25.4,
};

/** IPC: file:open result */
export interface OpenFileResult {
  path: string;
  name: string;
  ext: "docx" | "txt" | "html";
  buffer: Uint8Array;
}

/** IPC: file:pickImage result */
export interface PickImageResult {
  /** Data URL (png/jpeg/gif/bmp) safe to embed in the document JSON. */
  dataUrl: string;
  /** Original file name (used as alt text). */
  name: string;
}

/** IPC: docx:save payload */
export interface SaveDocPayload {
  json: Record<string, unknown>;
  defaultName: string;
  path?: string;
}

/**
 * IPC result of "write these bytes to a path the user picked".
 *
 * `error` is returned (not thrown) so the renderer always regains control:
 * a rejected `invoke` used to leave the save command, and the quit handshake
 * with it, hanging with no feedback at all.
 */
export type WriteFileResult = { path: string; error?: undefined } | { error: string; path?: undefined };

/** @deprecated Kept as an alias — every file write uses the same shape. */
export type SaveDocResult = WriteFileResult;

/** IPC: file:saveRaw payload (plain-text formats, no .docx conversion). */
export interface SaveRawPayload {
  /** File contents, already serialized by the renderer. */
  content: string;
  defaultName: string;
  /** Extension the content is for; the main process re-checks this allowlist. */
  ext: "txt" | "html";
  path?: string;
}

/** IPC: print result — an error string when the print job could not start. */
export type PrintResult = { error?: string };

/** Menu command sent from main to the renderer. */
export interface MenuCommand {
  id: string;
  /** Only set for `recent:open` — the file to open. */
  path?: string;
}

/** Colour scheme. `system` follows the OS light/dark setting. */
export type Theme = "light" | "dark" | "system";

/**
 * User preferences persisted to `userData/settings.json` (main process) and
 * applied by the renderer. `notionMode` switches the chrome to the
 * continuous-canvas Notion look; the document model is identical either way.
 */
export interface AppSettings {
  notionMode: boolean;
  theme: Theme;
  /** Left page-tree sidebar (workspace pages). */
  sidebarOpen: boolean;
  /** Ribbon toolbar visibility inside Notion mode (hidden unless pinned). */
  toolbarPinned: boolean;
  /**
   * Name stamped on comments and tracked changes. The main process defaults
   * it to the OS user name; the Edit menu offers an override.
   */
  authorName: string;
}

export const DEFAULT_SETTINGS: AppSettings = {
  notionMode: false,
  theme: "light",
  sidebarOpen: false,
  toolbarPinned: false,
  authorName: "You",
};

/**
 * One comment thread (v1 is single-message per thread). Stored under the
 * `comments` attr of the `doc` node keyed by id; the `comment` mark on text
 * references that id. Keeping the store in the document JSON means
 * workspace-page autosave, `.docx` save and the quit handshake cover
 * comments with no extra plumbing.
 */
export interface CommentRecord {
  author: string;
  text: string;
  createdAt: number;
  resolved: boolean;
}

export type CommentStore = Record<string, CommentRecord>;

/**
 * Workspace pages (the Notion-style page tree, stored locally under
 * `userData/workspace/`). Metadata lives in `pages.json`; the document JSON
 * per page in `<id>.json` — the exact shape the .docx exporter consumes, so
 * a page and a file round-trip through the same pipeline.
 */
export interface PageMeta {
  id: string;
  title: string;
  /** Emoji shown in the sidebar (Unicode only — the CSP stays closed). */
  icon: string;
  parentId: string | null;
  order: number;
  updatedAt: number;
}

/** A page plus its document JSON (`readPage` answer / write payload). */
export interface WorkspacePage {
  meta: PageMeta;
  json: Record<string, unknown>;
}

/** A page document as trashed (restorable) — keyed by its `.trash` file name. */
export interface TrashEntry {
  /** File name inside `workspace/.trash/` — the unique key for restore/delete. */
  file: string;
  id: string;
  title: string;
  icon: string;
  parentId: string | null;
  updatedAt: number;
  deletedAt: number;
}

/** Anything larger is refused: an image-laden page over IPC must stay sane. */
export const MAX_PAGE_JSON_BYTES = 10 * 1024 * 1024;

/** The API exposed on window.fw by the preload script. */
export interface FwApi {
  /** Open a file. With a path (recent files) it is read directly; otherwise a dialog is shown. */
  openFile(path?: string): Promise<OpenFileResult | null>;
  saveDoc(payload: SaveDocPayload): Promise<WriteFileResult | null>;
  /**
   * Write a non-.docx format the renderer has already serialized. `ext` is
   * re-checked against an allowlist in the main process.
   */
  saveRaw(payload: SaveRawPayload): Promise<WriteFileResult | null>;
  /** Open a native image picker; resolves to the image as a data URL. */
  pickImage(): Promise<PickImageResult | null>;
  /** Print-to-PDF. Null when the user dismissed the Save dialog. */
  exportPdf(defaultName: string): Promise<WriteFileResult | null>;
  print(): Promise<PrintResult | null>;
  /**
   * Recent files, for callers outside the native menu. The menu builds its own
   * list in the main process (src/main/menu.ts), so the renderer currently has
   * no UI for this — kept as the entry point for an in-app list.
   */
  getRecent(): Promise<RecentEntry[]>;
  /** Read the persisted user settings (Notion mode / theme / sidebar). */
  getSettings(): Promise<AppSettings>;
  /** Merge a patch into the settings and return the full merged result. */
  setSettings(patch: Partial<AppSettings>): Promise<AppSettings>;
  /** All page metadata, ordered (tree order is parentId + order). */
  listPages(): Promise<PageMeta[]>;
  /** Read a page with its document JSON; null when the id is unknown. */
  readPage(id: string): Promise<WorkspacePage | null>;
  /** Persist a page's document JSON (debounced autosave uses this). */
  writePage(id: string, json: Record<string, unknown>): Promise<{ error?: string }>;
  /** Create a page (optionally under `parentId`); null when the write failed. */
  createPage(opts?: { title?: string; parentId?: string | null }): Promise<PageMeta | null>;
  /** Rename a page / change its icon. */
  renamePage(id: string, patch: { title?: string; icon?: string }): Promise<PageMeta | null>;
  /** Re-parent and/or re-order a page (sidebar drag). */
  movePage(id: string, parentId: string | null, order: number): Promise<{ error?: string }>;
  /** Delete a page — moved to the workspace trash, not erased. */
  deletePage(id: string): Promise<{ error?: string }>;
  /** Deleted pages, newest first — the trash the sidebar offers to restore. */
  listTrash(): Promise<TrashEntry[]>;
  /** Put a trashed page back into the tree (keyed by the trash file name). */
  restoreTrash(file: string): Promise<{ meta?: PageMeta; error?: string }>;
  /** Erase one trash entry for good. */
  deleteTrash(file: string): Promise<{ error?: string }>;
  /**
   * Open an external URL in the user's desktop browser. The main process
   * rejects anything but http/https/mailto/tel; returns whether it opened.
   */
  openExternal(url: string): Promise<boolean>;
  /** Subscribe to menu commands; returns an unsubscribe function. */
  onCommand(cb: (cmd: MenuCommand) => void): () => void;
  /** Respond to the main-process quit handshake. */
  sendQuitResponse(payload: { dirty: boolean; name: string; path: string | null }): void;
  /** Subscribe to `app:will-quit`; returns an unsubscribe function. */
  onWillQuit(cb: () => void): () => void;
  /** Subscribe to `app:save-then-quit`; returns an unsubscribe function. */
  onSaveThenQuit(cb: () => void): () => void;
}
