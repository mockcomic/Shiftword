import { app } from "electron";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import type { RecentEntry } from "../shared/types";

const MAX_RECENT = 8;

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

export async function upsertRecent(entry: RecentEntry): Promise<void> {
  const list = await readRecent();
  const filtered = list.filter((e) => e.path !== entry.path);
  // Order by most recent activity, not the file mtime.
  filtered.unshift({ ...entry, mtimeMs: Date.now() });
  try {
    await mkdir(path.dirname(path.join(app.getPath("userData"), "recent.json")), { recursive: true });
    await writeFile(
      path.join(app.getPath("userData"), "recent.json"),
      JSON.stringify(filtered.slice(0, MAX_RECENT), null, 2),
      "utf8",
    );
  } catch {
    // Best effort — never break an open/save over the recent store.
  }
}
