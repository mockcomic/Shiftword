import { app } from "electron";
import { mkdir, readFile, readdir, rename, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import type { PageMeta, TrashEntry, WorkspacePage } from "../shared/types";
import { MAX_PAGE_JSON_BYTES, isStorableCover } from "../shared/types";

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

/**
 * Serialize every store touch.
 *
 * Each operation reads `pages.json`, mutates it in memory, and writes the
 * whole file back — so two operations running at once (an autosave flush
 * racing a rename or a create) let the later writer clobber the earlier
 * one's changes with its stale snapshot. A module-level promise chain makes
 * each read-modify-write atomic with respect to the others in this process,
 * which is the only writer (the app is single-instance).
 */
let storeLock: Promise<unknown> = Promise.resolve();
function withStoreLock<T>(fn: () => Promise<T>): Promise<T> {
  const run = storeLock.then(fn, fn);
  // Keep the chain alive even if an operation rejects, and never let a
  // rejection escape into an unhandled `storeLock` itself.
  storeLock = run.catch(() => undefined);
  return run;
}

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

/**
 * Ids reach us over IPC from the renderer — they must never steer a path.
 * `pages` (any casing — Windows/macOS volumes are case-insensitive) is
 * rejected because doc files live next to `pages.json`, and `PAGES.json`
 * would be that same file on a case-insensitive volume.
 */
function isValidId(id: unknown): id is string {
  return typeof id === "string" && SAFE_ID.test(path.basename(id)) && id.toLowerCase() !== "pages";
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

/** Drop a malformed cover (wrong type, oversized, not a data URL) on read. */
function normalizeMeta(m: PageMeta): PageMeta {
  if (m.cover !== undefined && m.cover !== null && !isStorableCover(m.cover)) m.cover = undefined;
  return m;
}

async function readStore(): Promise<MetaStore> {
  try {
    const raw = await readFile(metaFile(), "utf8");
    const data: unknown = JSON.parse(raw);
    const list = Array.isArray(data) ? data : (data as MetaStore)?.pages;
    const pages = Array.isArray(list) ? list.filter(validMeta).map(normalizeMeta) : [];
    return { pages };
  } catch {
    return { pages: [] };
  }
}

async function ensureWorkspaceDir(): Promise<void> {
  // No caching of "already created": the directory can disappear underneath
  // us (manual cleanup, roaming-profile moves) and a stale flag would make
  // every write fail with ENOENT forever. `mkdir -p` is idempotent.
  await mkdir(workspaceDir(), { recursive: true });
}

async function writeStore(store: MetaStore): Promise<void> {
  await ensureWorkspaceDir();
  await writeFileAtomic(metaFile(), JSON.stringify({ pages: store.pages }, null, 2));
}

/**
 * tmp + rename: an interrupted whole-file write would otherwise leave a
 * half-written `pages.json`, which every reader treats as an empty store —
 * silently erasing the user's page tree.
 */
async function writeFileAtomic(file: string, content: string): Promise<void> {
  const tmp = `${file}.tmp`;
  await writeFile(tmp, content, "utf8");
  await rename(tmp, file);
}

function newId(): string {
  const rand = Math.random().toString(36).slice(2, 10);
  return `p-${Date.now().toString(36)}-${rand}`;
}

export async function listPages(): Promise<PageMeta[]> {
  return withStoreLock(async () => (await readStore()).pages);
}

export async function readPage(id: string): Promise<WorkspacePage | null> {
  if (!isValidId(id)) return null;
  return withStoreLock(async () => {
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
  });
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
  // The cap is a BYTE budget; `serialized.length` counts UTF-16 units, which
  // under-counts CJK/emoji content by up to 3x.
  if (Buffer.byteLength(serialized, "utf8") > MAX_PAGE_JSON_BYTES) {
    return { error: `The page is larger than ${Math.round(MAX_PAGE_JSON_BYTES / (1024 * 1024))} MB.` };
  }
  return withStoreLock(async () => {
    const { pages } = await readStore();
    if (!pages.some((p) => p.id === id)) return { error: "Unknown page." };
    try {
      await ensureWorkspaceDir();
      await writeFileAtomic(docFile(id), serialized);
      const meta = pages.find((p) => p.id === id)!;
      meta.updatedAt = Date.now();
      await writeStore({ pages });
      return {};
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });
}

export async function createPage(opts?: {
  title?: string;
  parentId?: string | null;
}): Promise<PageMeta | null> {
  return withStoreLock(async () => {
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
      await ensureWorkspaceDir();
      await writeFileAtomic(docFile(meta.id), JSON.stringify(initialDoc()));
      await writeStore(store);
      return meta;
    } catch {
      return null;
    }
  });
}

export async function renamePage(
  id: string,
  patch: { title?: string; icon?: string; cover?: string | null },
): Promise<PageMeta | null> {
  if (!isValidId(id)) return null;
  return withStoreLock(async () => {
    const store = await readStore();
    const meta = store.pages.find((p) => p.id === id);
    if (!meta) return null;
    if (typeof patch.title === "string" && patch.title.trim()) meta.title = patch.title.trim().slice(0, 200);
    if (typeof patch.icon === "string" && [...patch.icon].length <= 4) meta.icon = patch.icon;
    if (patch.cover === null || patch.cover === "") {
      meta.cover = null;
    } else if (isStorableCover(patch.cover)) {
      // A cover is an image data URL, never an arbitrary string: the size and
      // `data:image/` checks keep `pages.json` bounded and the CSP happy.
      meta.cover = patch.cover;
    }
    meta.updatedAt = Date.now();
    await writeStore(store);
    return meta;
  });
}

/**
 * Deep-copy a page and its whole subtree (Notion's Duplicate). The copy lands
 * beside the original with the same icon/cover and a "(copy)" title; every
 * node in the copied subtree gets a fresh id, and the document files are
 * cloned alongside the metadata so the copy opens with the same content.
 */
export async function duplicatePage(id: string): Promise<{ meta?: PageMeta; error?: string }> {
  if (!isValidId(id)) return { error: "Invalid page id." };
  return withStoreLock(async () => {
    const store = await readStore();
    const source = store.pages.find((p) => p.id === id);
    if (!source) return { error: "Unknown page." };

    // The source subtree in parent-before-child order, so each clone can be
    // created while its (already cloned) parent is known.
    const subtree: PageMeta[] = [source];
    for (let i = 0; i < subtree.length; i++) {
      for (const p of store.pages) {
        if (p.parentId === subtree[i].id && !subtree.some((s) => s.id === p.id)) subtree.push(p);
      }
    }

    const idMap = new Map<string, string>();
    for (const p of subtree) idMap.set(p.id, newId());

    try {
      await ensureWorkspaceDir();
      const clones: PageMeta[] = [];
      for (const p of subtree) {
        const cloneId = idMap.get(p.id)!;
        // Read the doc file best-effort; an unreadable page duplicates as a
        // blank one instead of failing the whole copy.
        let json: unknown;
        try {
          json = JSON.parse(await readFile(docFile(p.id), "utf8"));
        } catch {
          json = initialDoc();
        }
        await writeFileAtomic(docFile(cloneId), JSON.stringify(json));
        clones.push({
          id: cloneId,
          title: p.id === source.id ? `${p.title} (copy)`.slice(0, 200) : p.title,
          icon: p.icon,
          cover: p.cover ?? null,
          // The original parent maps to its clone; the root copy keeps the
          // source's parent (or root).
          parentId: p.parentId ? (idMap.get(p.parentId) ?? null) : p.parentId,
          order: p.order,
          updatedAt: Date.now(),
        });
      }
      store.pages.push(...clones);
      await writeStore(store);
      return { meta: clones[0] };
    } catch (err) {
      return { error: messageOf(err) };
    }
  });
}

export async function movePage(
  id: string,
  parentId: string | null,
  order: number,
): Promise<{ error?: string }> {
  if (!isValidId(id)) return { error: "Invalid page id." };
  return withStoreLock(async () => {
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
  });
}

function isNotFound(err: unknown): boolean {
  return typeof err === "object" && err !== null && (err as { code?: string }).code === "ENOENT";
}

let trashSeq = 0;

export async function deletePage(id: string): Promise<{ error?: string }> {
  if (!isValidId(id)) return { error: "Invalid page id." };
  return withStoreLock(async () => {
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
        // Read best-effort: a page whose doc file never landed on disk (or is
        // unreadable) still gets a trash entry, so the page itself stays
        // restorable instead of vanishing with the store row.
        let json: unknown = initialDoc();
        let hadFile = false;
        try {
          const raw = await readFile(file, "utf8");
          hadFile = true;
          json = JSON.parse(raw);
        } catch (err) {
          if (!isNotFound(err)) console.warn(`[FreeWord] trashing unreadable page doc ${d}`, err);
        }
        const stamp = `${Date.now().toString(36)}-${(trashSeq++).toString(36)}${Math.random()
          .toString(36)
          .slice(2, 5)}`;
        const trashName = `${d}-${stamp}.json`;
        const payload = {
          deletedAt: Date.now(),
          meta: store.pages.find((p) => p.id === d)!,
          json,
        };
        await writeFile(path.join(trash, trashName), JSON.stringify(payload), "utf8");
        if (hadFile) await rm(file).catch(() => undefined);
      }
      store.pages = store.pages.filter((p) => !doomed.has(p.id));
      await writeStore(store);
      return {};
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });
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

/**
 * Deleted pages, newest first — unreadable files are skipped, not fatal.
 *
 * The name-derived timestamp orders the list BEFORE the cap is taken (the
 * old slice-in-readdir-order let the newest deletions fall off the end), and
 * the surviving entries are read in parallel instead of one serial round
 * after another on every sidebar refresh.
 */
export async function listTrash(): Promise<TrashEntry[]> {
  let names: string[];
  try {
    names = (await readdir(trashDir())).filter((n) => SAFE_TRASH_FILE.test(n));
  } catch {
    return []; // no trash yet
  }
  names.sort((a, b) => deletedAtFromTrashName(b) - deletedAtFromTrashName(a));
  const payloads = await Promise.all(names.slice(0, MAX_TRASH_ENTRIES).map(readTrashPayload));
  const entries = payloads.map((data, i) => toEntry(names[i], data));
  return entries.sort((a, b) => b.deletedAt - a.deletedAt);
}

/** Put a trashed page back. Unknown parents re-attach to the root, not vanish. */
export async function restoreTrash(fileName: string): Promise<{ meta?: PageMeta; error?: string }> {
  // The regex runs on the RAW name — a `basename()` pass first would let
  // `../evil.json` sneak past the character class (which already forbids
  // every path separator).
  if (!SAFE_TRASH_FILE.test(fileName)) return { error: "Invalid trash entry." };
  return withStoreLock(async () => {
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
      cover: storedMeta?.cover ?? null,
    };
    const json =
      data?.json && typeof data.json === "object" && (data.json as Record<string, unknown>).type === "doc"
        ? data.json
        : initialDoc();
    try {
      await ensureWorkspaceDir();
      await writeFileAtomic(docFile(id), JSON.stringify(json));
      store.pages.push(meta);
      await writeStore(store);
      // Only remove the trash file once the restore has actually landed —
      // a failed rm would leave a ghost row whose second restore always
      // fails with "A page with this id already exists".
      await rm(full);
      return { meta };
    } catch (err) {
      return { error: messageOf(err) };
    }
  });
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
