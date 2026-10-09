import { app } from "electron";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import path from "node:path";
import type { RecentEntry } from "../shared/types";

const MAX_RECENT = 8;

/** Same path identity as the ipc.ts trust check — casing must not fork entries. */
function samePath(a: string, b: string): boolean {
  const na = path.resolve(a);
  const nb = path.resolve(b);
  return process.platform === "linux" ? na === nb : na.toLowerCase() === nb.toLowerCase();
}

/** Serialize the read-modify-write so back-to-back upserts cannot lose entries. */
let chain: Promise<unknown> = Promise.resolve();

export async function readRecent(): Promise<RecentEntry[]> {
  try {
    const raw = await readFile(path.join(app.getPath("userData"), "recent.json"), "utf8");
    const data: unknown = JSON.parse(raw);
    if (!Array.isArray(data)) return [];
    const valid = data.filter(
      (e): e is RecentEntry =>
        !!e &&
        typeof (e as RecentEntry).path === "string" &&
        typeof (e as RecentEntry).name === "string" &&
        typeof (e as RecentEntry).mtimeMs === "number",
    );
    valid.sort((a, b) => b.mtimeMs - a.mtimeMs);
    return valid.slice(0, MAX_RECENT);
  } catch {
    // Missing or corrupt file → treat as empty.
    return [];
  }
}

export function upsertRecent(entry: RecentEntry): Promise<void> {
  const run = chain.then(async () => {
    const list = await readRecent();
    const filtered = list.filter((e) => !samePath(e.path, entry.path));
    // Order by most recent activity, not the file mtime.
    filtered.unshift({ ...entry, mtimeMs: Date.now() });
    const file = path.join(app.getPath("userData"), "recent.json");
    try {
      await mkdir(path.dirname(file), { recursive: true });
      // tmp + rename: an interrupted write must not erase the whole list.
      await writeFile(`${file}.tmp`, JSON.stringify(filtered.slice(0, MAX_RECENT), null, 2), "utf8");
      await rename(`${file}.tmp`, file);
    } catch {
      // Best effort — never break an open/save over the recent store.
    }
  });
  chain = run.catch(() => undefined);
  return run;
}
