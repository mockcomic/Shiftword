import type { Editor } from "@tiptap/core";
import { invalidateMeasureCache, runPagination } from "../editor/pagination";

/** Zoom levels offered in the toolbar (percent). */
export const ZOOM_STEPS = [50, 75, 90, 100, 110, 125, 150, 200] as const;

let currentZoom = 100;

export function getZoom(): number {
  return currentZoom;
}

/**
 * Apply CSS `zoom` to the document element. Chromium's zoom is layout-aware,
 * so all px-based measurements (pagination) stay consistent after zooming.
 * Print mode overrides zoom to 1 via `body.print-mode` CSS.
 */
export function setZoom(percent: number, editor?: Editor): void {
  const clamped = Math.min(400, Math.max(25, Math.round(percent)));
  if (clamped === currentZoom) return;
  currentZoom = clamped;
  const el = document.querySelector<HTMLElement>("#editor");
  if (el) el.style.zoom = String(clamped / 100);
  // Measured block heights are zoom-dependent — forget the cached ones.
  invalidateMeasureCache();
  if (editor) runPagination(editor);
}

/** Step to the next/previous zoom level relative to the current zoom. */
export function zoomStep(direction: 1 | -1, editor?: Editor): void {
  const next =
    direction === 1
      ? (ZOOM_STEPS.find((z) => z > currentZoom) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1])
      : ([...ZOOM_STEPS].reverse().find((z) => z < currentZoom) ?? ZOOM_STEPS[0]);
  setZoom(next, editor);
}
