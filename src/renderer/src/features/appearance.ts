import type { Editor } from "@tiptap/core";
import type { AppSettings, FwApi, Theme } from "../../../shared/types";
import { DEFAULT_SETTINGS } from "../../../shared/types";
import {
  invalidateMeasureCache,
  runPagination,
  schedulePaginate,
  setPaginationPaused,
} from "../editor/pagination";
import { syncToolbar } from "./toolbar";
import { updateStatus } from "./statusbar";

/**
 * Notion mode + theme, applied to the DOM and persisted through the main
 * process (userData/settings.json). The toggle is deliberately cheap: the
 * document model never changes — Notion mode pauses the pagination engine
 * and lets CSS collapse the pages into a continuous canvas, so the .docx
 * export path is untouched and switching back re-partitions from scratch.
 */

let settings: AppSettings = { ...DEFAULT_SETTINGS };
let editorRef: Editor | null = null;
let fwRef: FwApi | null = null;

export function getSettings(): AppSettings {
  return settings;
}

function resolvedTheme(): "light" | "dark" {
  if (settings.theme === "system") {
    return globalThis.matchMedia?.("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }
  return settings.theme;
}

/** Push the current settings onto <body> and the pagination engine. */
export function applyAppearance(): void {
  const body = document.body;
  const notion = settings.notionMode;
  body.classList.toggle("notion-mode", notion);
  body.classList.toggle("pinned-toolbar", settings.toolbarPinned);
  body.classList.toggle("sidebar-open", settings.sidebarOpen);
  body.dataset.theme = resolvedTheme();
  body.dataset.uiMode = notion ? "notion" : "word";
  // A theme or mode swap changes fonts and page geometry — cached block
  // heights from the old appearance must not steer the new pagination.
  invalidateMeasureCache();
  setPaginationPaused(notion);
  if (!notion && editorRef && !editorRef.isDestroyed) {
    // Back in Word mode: the pages were never re-partitioned while hidden,
    // so settle the layout now (one pass re-splits what flowed together).
    runPagination(editorRef);
  }
}

export async function initAppearance(editor: Editor, fw: FwApi | null): Promise<void> {
  editorRef = editor;
  fwRef = fw;
  if (fw) {
    try {
      settings = await fw.getSettings();
    } catch {
      settings = { ...DEFAULT_SETTINGS };
    }
  }
  applyAppearance();
  // `system` follows the OS live.
  if (globalThis.matchMedia) {
    const mq = globalThis.matchMedia("(prefers-color-scheme: dark)");
    mq.addEventListener?.("change", () => {
      if (settings.theme === "system") applyAppearance();
    });
  }
}

/** Flip Notion mode; persisted when the bridge is available. */
export async function toggleNotionMode(): Promise<void> {
  const next = !settings.notionMode;
  // Turning Notion mode on also brings the sidebar up (Notion ships with it
  // open); the user can close it again and that sticks.
  const patch = next ? { notionMode: true, sidebarOpen: true } : { notionMode: false };
  settings = fwRef ? await fwRef.setSettings(patch) : { ...settings, ...patch };
  applyAppearance();
  if (editorRef && !editorRef.isDestroyed) {
    schedulePaginate(editorRef);
    syncToolbar(editorRef);
    updateStatus(editorRef);
  }
}

export async function toggleSidebar(): Promise<void> {
  const next = !settings.sidebarOpen;
  settings = fwRef ? await fwRef.setSettings({ sidebarOpen: next }) : { ...settings, sidebarOpen: next };
  applyAppearance();
}

export async function setTheme(theme: Theme): Promise<void> {
  settings = fwRef ? await fwRef.setSettings({ theme }) : { ...settings, theme };
  applyAppearance();
}

export async function toggleToolbarPin(): Promise<void> {
  const next = !settings.toolbarPinned;
  settings = fwRef ? await fwRef.setSettings({ toolbarPinned: next }) : { ...settings, toolbarPinned: next };
  applyAppearance();
}

/**
 * Print/PDF wrap: the print engine needs the *Word* rendering — fixed pages
 * and a settled pagination — no matter which mode is on screen. Call this
 * before capture and `endPrintAppearance` after.
 */
export function beginPrintAppearance(): void {
  document.body.classList.remove("notion-mode");
  setPaginationPaused(false);
  if (editorRef && !editorRef.isDestroyed) runPagination(editorRef);
}

export function endPrintAppearance(): void {
  applyAppearance();
}
