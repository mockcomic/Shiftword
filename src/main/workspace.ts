import { app } from "electron";
import { mkdir, readFile, readdir, rename, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import type { PageMeta, TrashEntry, WorkspacePage } from "../shared/types";
import { MAX_PAGE_JSON_BYTES } from "../shared/types";

/**
 * The workspace: a local page tree under `userData/workspace/`.
 *
 * Layout — `pages.json` holds the metadata array; every page's document JSON
 * is its own `<id>.json` file (the same shape `docx-export` consumes, so the
 * .docx pipeline needs no knowledge of pages at all). Deleting a page moves
 * its file into `.trash/` wrapped with its metadata so the UI can offer
 * restore; `listTrash` / `restoreTrash` / `deleteTrash` surface that.
 *
 * Same best-effort posture as recent.ts: a corrupt store must degrade to an
 * empty workspace, never break open/save or the quit handshake.
 */

const SAFE_ID = /^[A-Za-z0-9_-]{4,64}$/;
/** Trash entries are keyed by their file name (`<id>-<ts36>.json`). */
const SAFE_TRASH_FILE = /^[A-Za-z0-9_-]{4,140}\.json$/;

/** How many trash entries the sidebar ever needs — older ones can be emptied on disk. */
const MAX_TRASH_ENTRIES = 200;

function workspaceDir(): string {
  return path.join(app.getPath("userData"), "workspace");
}
function metaFile(): string {
  return path.join(workspaceDir(), "pages.json");
}
function docFile(id: string): string {
  return path.join(workspaceDir(), `${id}.json`);
}
function trashDir(): string {
  return path.join(workspaceDir(), ".trash");
}

/** A short, user-presentable description of a caught error. */
function messageOf(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

/** Ids reach us over IPC from the renderer — they must never steer a path. */
function isValidId(id: unknown): id is string {
  return typeof id === "string" && SAFE_ID.test(id) && id !== "pages";
}

function initialDoc(): Record<string, unknown> {
  return {
    type: "doc",
    attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
    content: [{ type: "page", content: [{ type: "paragraph" }] }],
  };
}

interface MetaStore {
  pages: PageMeta[];
}

function validMeta(m: unknown): m is PageMeta {
  if (!m || typeof m !== "object") return false;
  const c = m as Record<string, unknown>;
  return (
    typeof c.id === "string" &&
    SAFE_ID.test(c.id) &&
    typeof c.title === "string" &&
    typeof c.updatedAt === "number" &&
    (c.parentId === null || typeof c.parentId === "string") &&
    typeof c.order === "number"
  );
}

async function readStore(): Promise<MetaStore> {
  try {
    const raw = await readFile(metaFile(), "utf8");
    const data: unknown = JSON.parse(raw);
    const list = Array.isArray(data) ? data : (data as MetaStore)?.pages;
    const pages = Array.isArray(list) ? list.filter(validMeta) : [];
    return { pages };
  } catch {
    return { pages: [] };
  }
}

async function writeStore(store: MetaStore): Promise<void> {
  await mkdir(workspaceDir(), { recursive: true });
  await writeFile(metaFile(), JSON.stringify({ pages: store.pages }, null, 2), "utf8");
}

function newId(): string {
  const rand = Math.random().toString(36).slice(2, 10);
  return `p-${Date.now().toString(36)}-${rand}`;
}

export async function listPages(): Promise<PageMeta[]> {
  const { pages } = await readStore();
  return pages;
}

export async function readPage(id: string): Promise<WorkspacePage | null> {
  if (!isValidId(id)) return null;
  const { pages } = await readStore();
  const meta = pages.find((p) => p.id === id);
  if (!meta) return null;
  try {
    const raw = await readFile(docFile(id), "utf8");
    const json: unknown = JSON.parse(raw);
    if (!json || typeof json !== "object" || (json as Record<string, unknown>).type !== "doc") {
      return { meta, json: initialDoc() };
    }
    return { meta, json: json as Record<string, unknown> };
  } catch {
    // Metadata outlived the doc file (or it never existed): open a clean page.
    return { meta, json: initialDoc() };
  }
}

export async function writePage(id: string, json: unknown): Promise<{ error?: string }> {
  if (!isValidId(id)) return { error: "Invalid page id." };
  let serialized: string;
  try {
    serialized = JSON.stringify(json);
  } catch {
    return { error: "The document could not be serialized." };
  }
  if (serialized === undefined) return { error: "Nothing to save." };
  if (serialized.length > MAX_PAGE_JSON_BYTES) {
    return { error: `The page is larger than ${Math.round(MAX_PAGE_JSON_BYTES / (1024 * 1024))} MB.` };
  }
  const { pages } = await readStore();
  if (!pages.some((p) => p.id === id)) return { error: "Unknown page." };
  try {
    await mkdir(workspaceDir(), { recursive: true });
    await writeFile(docFile(id), serialized, "utf8");
    const meta = pages.find((p) => p.id === id)!;
    meta.updatedAt = Date.now();
    await writeStore({ pages });
    return {};
  } catch (err) {
    return { error: err instanceof Error ? err.message : String(err) };
  }
}

export async function createPage(opts?: {
  title?: string;
  parentId?: string | null;
}): Promise<PageMeta | null> {
  const store = await readStore();
  const parentId = opts?.parentId && store.pages.some((p) => p.id === opts.parentId) ? opts.parentId : null;
  const siblings = store.pages.filter((p) => p.parentId === parentId);
  const meta: PageMeta = {
    id: newId(),
    title: (opts?.title ?? "").trim() || "Untitled",
    icon: "📄",
    parentId,
    order: siblings.length === 0 ? 0 : Math.max(...siblings.map((s) => s.order)) + 1,
    updatedAt: Date.now(),
  };
  store.pages.push(meta);
  try {
    await mkdir(workspaceDir(), { recursive: true });
    await writeStore(store);
    await writeFile(docFile(meta.id), JSON.stringify(initialDoc()), "utf8");
    return meta;
  } catch {
    return null;
  }
}

export async function renamePage(
  id: string,
  patch: { title?: string; icon?: string },
): Promise<PageMeta | null> {
  if (!isValidId(id)) return null;
  const store = await readStore();
  const meta = store.pages.find((p) => p.id === id);
  if (!meta) return null;
  if (typeof patch.title === "string" && patch.title.trim()) meta.title = patch.title.trim().slice(0, 200);
  if (typeof patch.icon === "string" && [...patch.icon].length <= 4) meta.icon = patch.icon;
  meta.updatedAt = Date.now();
  await writeStore(store);
  return meta;
}

export async function movePage(
  id: string,
  parentId: string | null,
  order: number,
): Promise<{ error?: string }> {
  if (!isValidId(id)) return { error: "Invalid page id." };
  const store = await readStore();
  const meta = store.pages.find((p) => p.id === id);
  if (!meta) return { error: "Unknown page." };
  if (parentId !== null) {
    if (!isValidId(parentId)) return { error: "Invalid parent id." };
    if (parentId === id) return { error: "A page cannot contain itself." };
    if (!store.pages.some((p) => p.id === parentId)) return { error: "Unknown parent." };
    // Refuse cycles: the new parent must not be a descendant of the page.
    let cursor: PageMeta | undefined = store.pages.find((p) => p.id === parentId);
    while (cursor) {
      if (cursor.id === id) return { error: "That would create a cycle." };
      cursor = cursor.parentId ? store.pages.find((p) => p.id === cursor!.parentId) : undefined;
    }
  }
  meta.parentId = parentId;
  meta.order = Number.isFinite(order) ? Math.max(0, Math.floor(order)) : 0;
  meta.updatedAt = Date.now();
  await writeStore(store);
  return {};
}

export async function deletePage(id: string): Promise<{ error?: string }> {
  if (!isValidId(id)) return { error: "Invalid page id." };
  const store = await readStore();
  if (!store.pages.some((p) => p.id === id)) return { error: "Unknown page." };

  // Collect the whole subtree (children go to the trash with the parent —
  // a half-orphaned tree is worse than a deleted branch).
  const doomed = new Set<string>([id]);
  let grew = true;
  while (grew) {
    grew = false;
    for (const p of store.pages) {
      if (p.parentId && doomed.has(p.parentId) && !doomed.has(p.id)) {
        doomed.add(p.id);
        grew = true;
      }
    }
  }

  const trash = trashDir();
  try {
    await mkdir(trash, { recursive: true });
    for (const d of doomed) {
      const file = docFile(d);
      if (!existsSync(file)) continue;
      const trashName = `${d}-${Date.now().toString(36)}.json`;
      try {
        // Wrap the doc with its metadata so the trash UI can show a title
        // and restore re-adopts the page unchanged.
        const raw = await readFile(file, "utf8");
        let json: unknown;
        try {
          json = JSON.parse(raw);
        } catch {
          json = initialDoc();
        }
        const payload = {
          deletedAt: Date.now(),
          meta: store.pages.find((p) => p.id === d)!,
          json,
        };
        await writeFile(path.join(trash, trashName), JSON.stringify(payload), "utf8");
        await rm(file);
      } catch {
        // Fallback to the plain move: the bytes stay recoverable by hand.
        await rename(file, path.join(trash, trashName)).catch(() => undefined);
      }
    }
    store.pages = store.pages.filter((p) => !doomed.has(p.id));
    await writeStore(store);
    return {};
  } catch (err) {
    return { error: err instanceof Error ? err.message : String(err) };
  }
}

/* --- trash (restore UI for deleted pages) ----------------------------------
 *
 * A trash file is either `{ deletedAt, meta, json }` (current format) or the
 * bare doc JSON (entries made before the restore UI existed). The file name
 * `<id>-<ts36>.json` doubles as the unique key across restore/delete calls,
 * because the same id can end up in the trash more than once.
 */

function idFromTrashName(name: string): string {
  const base = name.replace(/\.json$/, "");
  const cut = base.lastIndexOf("-");
  return cut > 0 ? base.slice(0, cut) : base;
}

function deletedAtFromTrashName(name: string): number {
  const base = name.replace(/\.json$/, "");
  const cut = base.lastIndexOf("-");
  const ts = cut > 0 ? Number.parseInt(base.slice(cut + 1), 36) : NaN;
  return Number.isFinite(ts) ? ts : 0;
}

interface TrashPayload {
  deletedAt?: number;
  meta?: PageMeta;
  json?: unknown;
}

function toEntry(name: string, data: TrashPayload | null): TrashEntry {
  const meta = data?.meta && validMeta(data.meta) ? data.meta : null;
  return {
    file: name,
    id: meta?.id ?? idFromTrashName(name),
    title: meta?.title || "Untitled",
    icon: meta?.icon || "📄",
    parentId: meta?.parentId ?? null,
    updatedAt: meta?.updatedAt ?? 0,
    deletedAt: data?.deletedAt ?? deletedAtFromTrashName(name),
  };
}

async function readTrashPayload(fileName: string): Promise<TrashPayload | null> {
  try {
    const raw = await readFile(path.join(trashDir(), fileName), "utf8");
    const data: unknown = JSON.parse(raw);
    return data && typeof data === "object" ? (data as TrashPayload) : null;
  } catch {
    return null;
  }
}

/** Deleted pages, newest first — unreadable files are skipped, not fatal. */
export async function listTrash(): Promise<TrashEntry[]> {
  let names: string[];
  try {
    names = (await readdir(trashDir())).filter((n) => SAFE_TRASH_FILE.test(n));
  } catch {
    return []; // no trash yet
  }
  const entries: TrashEntry[] = [];
  for (const name of names.slice(0, MAX_TRASH_ENTRIES)) {
    entries.push(toEntry(name, await readTrashPayload(name)));
  }
  return entries.sort((a, b) => b.deletedAt - a.deletedAt);
}

/** Put a trashed page back. Unknown parents re-attach to the root, not vanish. */
export async function restoreTrash(fileName: string): Promise<{ meta?: PageMeta; error?: string }> {
  if (!SAFE_TRASH_FILE.test(fileName)) return { error: "Invalid trash entry." };
  const full = path.join(trashDir(), fileName);
  const data = await readTrashPayload(fileName);
  const storedMeta = data?.meta && validMeta(data.meta) ? data.meta : null;
  const id = storedMeta?.id ?? idFromTrashName(fileName);
  if (!isValidId(id)) return { error: "Invalid page id." };
  const store = await readStore();
  if (store.pages.some((p) => p.id === id)) return { error: "A page with this id already exists." };
  const parentId =
    storedMeta?.parentId && store.pages.some((p) => p.id === storedMeta.parentId)
      ? storedMeta.parentId
      : null;
  const siblings = store.pages.filter((p) => p.parentId === parentId);
  const meta: PageMeta = {
    id,
    title: (storedMeta?.title ?? "").trim() || "Untitled",
    icon: storedMeta?.icon || "📄",
    parentId,
    order: siblings.length === 0 ? 0 : Math.max(...siblings.map((s) => s.order)) + 1,
    updatedAt: Date.now(),
  };
  const json =
    data?.json && typeof data.json === "object" && (data.json as Record<string, unknown>).type === "doc"
      ? data.json
      : initialDoc();
  try {
    await mkdir(workspaceDir(), { recursive: true });
    await writeFile(docFile(id), JSON.stringify(json), "utf8");
    store.pages.push(meta);
    await writeStore(store);
    await rm(full).catch(() => undefined);
    return { meta };
  } catch (err) {
    return { error: messageOf(err) };
  }
}

/** Erase one trash entry for good. */
export async function deleteTrash(fileName: string): Promise<{ error?: string }> {
  if (!SAFE_TRASH_FILE.test(fileName)) return { error: "Invalid trash entry." };
  try {
    await rm(path.join(trashDir(), fileName));
    return {};
  } catch (err) {
    return { error: messageOf(err) };
  }
}
