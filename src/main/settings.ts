import { app } from "electron";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { DEFAULT_SETTINGS } from "../shared/types";
import type { AppSettings, Theme } from "../shared/types";

/**
 * Persisted user preferences (`userData/settings.json`). Same posture as
 * recent.ts: best-effort, never throws, corrupt file → defaults — a broken
 * settings file must not stop the app from starting.
 */

const THEMES: Theme[] = ["light", "dark", "system"];

/** Memoized: `os.userInfo()` is synchronous and slow on domain-joined Windows. */
let cachedAuthor: string | null = null;
function defaultAuthor(): string {
  if (cachedAuthor !== null) return cachedAuthor;
  try {
    cachedAuthor = os.userInfo().username || DEFAULT_SETTINGS.authorName;
  } catch {
    cachedAuthor = DEFAULT_SETTINGS.authorName;
  }
  return cachedAuthor;
}

function authorNameOf(value: unknown): string {
  return typeof value === "string" && value.trim() ? value.trim() : defaultAuthor();
}

function settingsFile(): string {
  return path.join(app.getPath("userData"), "settings.json");
}

/** Serialize the read-modify-write so concurrent setSettings calls cannot lose a key. */
let chain: Promise<unknown> = Promise.resolve();

export async function readSettings(): Promise<AppSettings> {
  try {
    const raw = await readFile(settingsFile(), "utf8");
    const data: unknown = JSON.parse(raw);
    if (!data || typeof data !== "object") return { ...DEFAULT_SETTINGS };
    const s = data as Record<string, unknown>;
    return {
      notionMode: s.notionMode === true,
      theme: THEMES.includes(s.theme as Theme) ? (s.theme as Theme) : DEFAULT_SETTINGS.theme,
      sidebarOpen: s.sidebarOpen === true,
      toolbarPinned: s.toolbarPinned === true,
      authorName: authorNameOf(s.authorName),
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function setSettings(patch: Partial<AppSettings>): Promise<AppSettings> {
  const run = chain.then(async () => {
    const current = await readSettings();
    const next: AppSettings = {
      notionMode: typeof patch.notionMode === "boolean" ? patch.notionMode : current.notionMode,
      theme: THEMES.includes(patch.theme as Theme) ? (patch.theme as Theme) : current.theme,
      sidebarOpen: typeof patch.sidebarOpen === "boolean" ? patch.sidebarOpen : current.sidebarOpen,
      toolbarPinned: typeof patch.toolbarPinned === "boolean" ? patch.toolbarPinned : current.toolbarPinned,
      authorName: authorNameOf(patch.authorName ?? current.authorName),
    };
    const file = settingsFile();
    try {
      await mkdir(path.dirname(file), { recursive: true });
      // tmp + rename: half a settings.json must not erase the user's theme.
      await writeFile(`${file}.tmp`, JSON.stringify(next, null, 2), "utf8");
      await rename(`${file}.tmp`, file);
    } catch {
      // Best effort — the in-memory answer stays authoritative for the session.
    }
    return next;
  });
  chain = run.catch(() => undefined);
  return run;
}
