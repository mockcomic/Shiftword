import { app } from "electron";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import type { PageMeta, WorkspacePage } from "../shared/types";
import { MAX_PAGE_JSON_BYTES } from "../shared/types";

/**
 * The workspace: a local page tree under `userData/workspace/`.
 *
 * Layout — `pages.json` holds the metadata array; every page's document JSON
 * is its own `<id>.json` file (the same shape `docx-export` consumes, so the
 * .docx pipeline needs no knowledge of pages at all). Deleting a page renames
 * its file into `.trash/` — recovery stays possible on disk even though the
 * v1 UI does not surface it.
 *
 * Same best-effort posture as recent.ts: a corrupt store must degrade to an
 * empty workspace, never break open/save or the quit handshake.
 */

const SAFE_ID = /^[A-Za-z0-9_-]{4,64}$/;

function workspaceDir(): string {
  return path.join(app.getPath("userData"), "workspace");
}
function metaFile(): string {
  return path.join(workspaceDir(), "pages.json");
}
function docFile(id: string): string {
  return path.join(workspaceDir(), `${id}.json`);
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

  const trashDir = path.join(workspaceDir(), ".trash");
  try {
    await mkdir(trashDir, { recursive: true });
    for (const d of doomed) {
      const file = docFile(d);
      if (existsSync(file)) {
        await rename(file, path.join(trashDir, `${d}-${Date.now().toString(36)}.json`)).catch(
          () => undefined,
        );
      }
    }
    store.pages = store.pages.filter((p) => !doomed.has(p.id));
    await writeStore(store);
    return {};
  } catch (err) {
    return { error: err instanceof Error ? err.message : String(err) };
  }
}
