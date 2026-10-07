import { app } from "electron";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { DEFAULT_SETTINGS } from "../shared/types";
import type { AppSettings, Theme } from "../shared/types";

/**
 * Persisted user preferences (`userData/settings.json`). Same posture as
 * recent.ts: best-effort, never throws, corrupt file → defaults — a broken
 * settings file must not stop the app from starting.
 */

const THEMES: Theme[] = ["light", "dark", "system"];

function settingsFile(): string {
  return path.join(app.getPath("userData"), "settings.json");
}

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
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export async function setSettings(patch: Partial<AppSettings>): Promise<AppSettings> {
  const current = await readSettings();
  const next: AppSettings = {
    notionMode: typeof patch.notionMode === "boolean" ? patch.notionMode : current.notionMode,
    theme: THEMES.includes(patch.theme as Theme) ? (patch.theme as Theme) : current.theme,
    sidebarOpen: typeof patch.sidebarOpen === "boolean" ? patch.sidebarOpen : current.sidebarOpen,
    toolbarPinned: typeof patch.toolbarPinned === "boolean" ? patch.toolbarPinned : current.toolbarPinned,
  };
  try {
    await mkdir(path.dirname(settingsFile()), { recursive: true });
    await writeFile(settingsFile(), JSON.stringify(next, null, 2), "utf8");
  } catch {
    // Best effort — the in-memory answer stays authoritative for the session.
  }
  return next;
}
