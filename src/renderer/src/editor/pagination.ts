import type { Editor } from "@tiptap/core";
import { Fragment, Slice } from "@tiptap/pm/model";
import type { Node as PMNode } from "@tiptap/pm/model";
import { AllSelection, NodeSelection, TextSelection } from "@tiptap/pm/state";
import type { Selection } from "@tiptap/pm/state";
import { FW_FOOTER_RESERVE_PX, FW_HEADER_RESERVE_PX } from "./schema";

/**
 * Pagination engine: content must fit exactly one page per `page` node.
 *
 * The pure core (layoutPages / computePageLayout) is DOM-free and
 * unit-tested; the DOM driver (runPagination) measures blocks, applies the
 * computed layout as a single transaction, and re-measures until settled.
 */

export interface BlockInfo {
  /** Absolute ProseMirror position of the block's start. */
  pos: number;
  /** Measured rendered height in CSS px. */
  h: number;
  isPageBreak: boolean;
}

export interface Move {
  fromPageIdx: number;
  /** Absolute position of the block being moved. */
  pos: number;
  /** Final page index in the re-paginated sequence. */
  toPageIdx: number;
}

export const MAX_PAGINATION_PASSES = 30;

/**
 * Compute the re-paginated layout.
 *
 * Rules (pages processed left → right):
 * 1. Explicit break: the first `pageBreak` in a page acts as a barrier —
 *    every block after it moves to a new page after the current one.
 *    Consecutive barriers naturally yield forced blank pages (two breaks
 *    = one empty page, Word semantics).
 * 2. Overflow: the first block that would push the page over
 *    `pageContentHeight` (± tolerance) and every block after it moves to
 *    the next page. A single block taller than a full page is left in
 *    place (documented MVP limitation); the blocks after it still move.
 * 3. A page whose only child is a trailing `pageBreak` is never merged away.
 *
 * Block order is always preserved: the result is a re-partition of the
 * same flat block sequence.
 */
export function layoutPages(pages: BlockInfo[][], pageContentHeight: number, tolerance = 2): BlockInfo[][] {
  const cur: BlockInfo[][] = pages.map((p) => p.slice());
  let i = 0;
  while (i < cur.length) {
    let page = cur[i];

    // 1. Explicit barrier: everything after the first pageBreak → new page.
    const brIdx = page.findIndex((b) => b.isPageBreak);
    if (brIdx >= 0 && brIdx < page.length - 1) {
      cur.splice(i + 1, 0, page.slice(brIdx + 1));
      cur[i] = page.slice(0, brIdx + 1);
      page = cur[i];
    }

    // 2. Overflow.
    let acc = 0;
    let cut = -1;
    for (let j = 0; j < page.length; j++) {
      if (acc + page[j].h > pageContentHeight + tolerance) {
        cut = j;
        break;
      }
      acc += page[j].h;
    }
    if (cut >= 0) {
      const first = page[cut];
      if (first.h > pageContentHeight + tolerance) {
        const after = page.slice(cut + 1);
        if (after.length > 0) {
          cur.splice(i + 1, 0, after);
          cur[i] = page.slice(0, cut + 1);
        }
      } else {
        cur.splice(i + 1, 0, page.slice(cut));
        cur[i] = page.slice(0, cut);
      }
    }
    i++;
  }
  return cur;
}

/** Decision function over heights: returns the moves needed to settle. */
export function computePageLayout(pages: BlockInfo[][], pageContentHeight: number, tolerance = 2): Move[] {
  const finalLayout = layoutPages(pages, pageContentHeight, tolerance);
  const finalIndex = new Map<number, number>();
  finalLayout.forEach((pg, fi) => pg.forEach((b) => finalIndex.set(b.pos, fi)));
  const moves: Move[] = [];
  pages.forEach((pg, pi) => {
    for (const b of pg) {
      const fi = finalIndex.get(b.pos);
      if (fi !== undefined && fi !== pi) moves.push({ fromPageIdx: pi, pos: b.pos, toPageIdx: fi });
    }
  });
  return moves;
}

// ---------------------------------------------------------------------------
// DOM driver
// ---------------------------------------------------------------------------

interface PageMeasure {
  /** Absolute position of the page node. */
  start: number;
  /** nodeSize (content size + 2). */
  size: number;
  blocks: BlockInfo[];
}

interface Range {
  start: number;
  size: number;
}

let rafHandle: number | null = null;

/**
 * Notion mode pauses the engine entirely: on the continuous canvas the pages
 * have no fixed height, so measuring blocks against `pageContentHeight` would
 * produce nonsense. The `page` nodes stay in the schema (exports are
 * unaffected) — they just stop being re-partitioned until Word mode returns.
 */
let paginationPaused = false;

export function setPaginationPaused(paused: boolean): void {
  paginationPaused = paused;
}

export function isPaginationPaused(): boolean {
  return paginationPaused;
}

/** Debounced pagination: at most one run per animation frame. */
export function schedulePaginate(editor: Editor): void {
  if (paginationPaused || rafHandle !== null) return;
  rafHandle = requestAnimationFrame(() => {
    rafHandle = null;
    if (!editor.isDestroyed && !paginationPaused) runPagination(editor);
  });
}

/**
 * Absolute document positions of a page's direct block children.
 *
 * `Node.forEach` reports offsets relative to the page's own content, while
 * `view.nodeDOM` and `doc.nodeAt` take absolute positions — the page's start
 * plus the token that opens it must be added. Getting this wrong makes every
 * page after the first resolve to the wrong node (a page, another page's
 * block, or nothing at all).
 */
export function blockPositionsOfPage(
  pageNode: PMNode,
  pagePos: number,
): Array<{ pos: number; isPageBreak: boolean }> {
  const out: Array<{ pos: number; isPageBreak: boolean }> = [];
  pageNode.forEach((child, childOffset) => {
    out.push({ pos: pagePos + 1 + childOffset, isPageBreak: child.type.name === "pageBreak" });
  });
  return out;
}

/**
 * Measured-height cache keyed by the *node object*.
 *
 * ProseMirror nodes are immutable and reused by identity in every state
 * that did not touch them, so one full-document measurement pass (the old
 * per-keystroke cost: a forced-layout `getBoundingClientRect` per block,
 * every frame, every settle-loop pass) becomes "measure the blocks that
 * are new, reuse the heights of everything else". Block width is the same
 * on every page, so a node's height only depends on the node — until
 * something that reflows the whole document changes: zoom, theme/fonts or
 * the window. Those call `invalidateMeasureCache`.
 */
let heightCache = new WeakMap<PMNode, number>();

export function invalidateMeasureCache(): void {
  heightCache = new WeakMap();
}

if (typeof window !== "undefined") {
  window.addEventListener("resize", () => invalidateMeasureCache());
}

function measureBlockHeight(el: HTMLElement): number {
  // `getBoundingClientRect().height` excludes CSS margins — but the blocks
  // carry a real `margin: 0 0 6pt` (typography.ts), and ignoring it made
  // every page overfill by ~8px per block against the 2px tolerance.
  const rect = el.getBoundingClientRect();
  let h = rect.height;
  const cs = getComputedStyle(el);
  h += parseFloat(cs.marginTop) || 0;
  h += parseFloat(cs.marginBottom) || 0;
  return h;
}

function measurePages(editor: Editor): { pages: PageMeasure[]; contentHeight: number } | null {
  const { view, state } = editor;
  const doc = state.doc;
  const pages: PageMeasure[] = [];
  let contentHeight: number | null = null;

  doc.forEach((pageNode, pagePos) => {
    const pageEl = view.nodeDOM(pagePos) as HTMLElement | null;
    if (!pageEl) return;
    if (contentHeight === null) {
      const cs = getComputedStyle(pageEl);
      const padTop = parseFloat(cs.paddingTop) || 0;
      const padBottom = parseFloat(cs.paddingBottom) || 0;
      contentHeight = pageEl.clientHeight - padTop - padBottom;
      // The header/footer bands render absolutely inside the page's margin
      // padding; while they carry text, shrink the usable content box so
      // typed blocks never flow underneath them (see schema.ts reserves).
      const attrs = (state.doc.attrs ?? {}) as { header?: string | null; footer?: string | null };
      if (attrs.header) contentHeight -= FW_HEADER_RESERVE_PX;
      if (attrs.footer) contentHeight -= FW_FOOTER_RESERVE_PX;
    }
    const blocks: BlockInfo[] = [];
    pageNode.forEach((child, childOffset) => {
      const pos = pagePos + 1 + childOffset;
      let h = heightCache.get(child);
      if (h === undefined) {
        const el = view.nodeDOM(pos) as HTMLElement | null;
        h = el ? measureBlockHeight(el) : 0;
        heightCache.set(child, h);
      }
      blocks.push({ pos, h, isPageBreak: child.type.name === "pageBreak" });
    });
    pages.push({ start: pagePos, size: pageNode.nodeSize, blocks });
  });

  if (contentHeight === null || contentHeight <= 0 || pages.length === 0) return null;
  return { pages, contentHeight };
}

/** True when the last page is a single empty paragraph not following a pageBreak. */
function hasTrailingEmptyPage(editor: Editor): boolean {
  const doc = editor.state.doc;
  if (doc.childCount < 2) return false;
  const last = doc.child(doc.childCount - 1);
  const prev = doc.child(doc.childCount - 2);
  const lastBlock = prev.lastChild;
  if (lastBlock && lastBlock.type.name === "pageBreak") return false;
  return last.childCount === 1 && last.child(0).type.name === "paragraph" && !last.child(0).firstChild;
}

/** Map an old doc position onto the re-paginated doc (block offsets preserved). */
function makeMapper(oldPages: Range[], newPages: Range[]): (p: number) => number {
  const pairs: Array<[number, number]> = [];
  const n = Math.min(oldPages.length, newPages.length);
  for (let i = 0; i < n; i++) {
    pairs.push([oldPages[i].start, newPages[i].start]);
    pairs.push([oldPages[i].start + oldPages[i].size, newPages[i].start + newPages[i].size]);
  }
  if (pairs.length === 0) return (p) => p;
  const lastOldEnd = pairs[pairs.length - 1][0];
  return (p: number): number => {
    const clamped = Math.min(p, lastOldEnd);
    // Binary search for the last pair with start <= clamped — the pairs are
    // sorted, and a linear scan here is O(pages) per mapped position on a
    // path that runs once per selection component per applied pass.
    let lo = 0;
    let hi = pairs.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (pairs[mid][0] <= clamped) lo = mid;
      else hi = mid - 1;
    }
    const best = pairs[lo];
    return best[1] + (clamped - best[0]);
  };
}

function mapSelection(newDoc: PMNode, oldSel: Selection, map: (p: number) => number): Selection {
  try {
    if (oldSel instanceof NodeSelection) return NodeSelection.create(newDoc, map(oldSel.anchor));
    if (oldSel instanceof AllSelection) return new AllSelection(newDoc);
    return TextSelection.create(newDoc, map(oldSel.anchor), map(oldSel.head));
  } catch {
    return TextSelection.atStart(newDoc);
  }
}

/**
 * Measure → compute → apply. Repeats until a pass makes no changes.
 * Returns true if any layout change was applied.
 */
export function runPagination(editor: Editor): boolean {
  if (editor.isDestroyed || paginationPaused) return false;
  let appliedCount = 0;

  for (let pass = 0; pass < MAX_PAGINATION_PASSES; pass++) {
    const measured = measurePages(editor);
    if (!measured) break;
    const { pages, contentHeight } = measured;
    const trailing = hasTrailingEmptyPage(editor);
    const effective = trailing ? pages.slice(0, -1) : pages;

    const finalLayout = layoutPages(
      effective.map((pg) => pg.blocks),
      contentHeight,
    );
    const unchanged =
      !trailing &&
      effective.length === finalLayout.length &&
      effective.every(
        (pg, i) =>
          pg.blocks.length === finalLayout[i].length &&
          pg.blocks.every((b, j) => b.pos === finalLayout[i][j].pos),
      );
    if (unchanged) break;

    const doc = editor.state.doc;
    const firstPage = doc.childCount > 0 ? doc.child(0) : null;
    if (!firstPage) break;

    let newPageNodes: PMNode[];
    try {
      newPageNodes = finalLayout.map((pg) => {
        const blocks = pg.map((b) => {
          const node = doc.nodeAt(b.pos);
          // Never drop a block on the floor: this transaction replaces every
          // page, so a missing node here means data loss. Abort instead.
          if (!node) throw new RangeError(`block at ${b.pos} is no longer in the document`);
          return node;
        });
        // Content is the SECOND argument — passing it inside the attrs object
        // makes createChecked throw RangeError: Invalid content for node page.
        return firstPage.type.createChecked(null, Fragment.fromArray(blocks));
      });
    } catch (err) {
      // A layout problem must never break the document: leave the current
      // pages as they are instead of throwing out of the update/rAF callback.
      console.error("[FreeWord] pagination could not be applied", err);
      break;
    }
    // Page starts must use the same convention on both sides of the selection
    // mapper: `measurePages` records the position *before* the page node (0 for
    // the first page), and so does the re-paginated document.
    let newStart = 0;
    const newRanges: Range[] = newPageNodes.map((np) => {
      const r = { start: newStart, size: np.nodeSize };
      newStart += np.nodeSize;
      return r;
    });

    const tr = editor.state.tr.replace(0, doc.content.size, new Slice(Fragment.from(newPageNodes), 0, 0));
    tr.setMeta("addToHistory", false); // layout-only change: not part of the undo history
    // Also not an edit. Tiptap emits `update` for any transaction whose doc
    // actually changed, so without this a re-partition — which happens on the
    // very first pass after importing a multi-page .docx — would run
    // `onEditorUpdate`, mark a freshly-opened document dirty and make the quit
    // handshake ask to save a file the user never touched.
    tr.setMeta("preventUpdate", true);
    tr.setSelection(mapSelection(tr.doc, editor.state.selection, makeMapper(pages, newRanges)));
    editor.view.dispatch(tr);
    appliedCount++;
  }

  if (appliedCount >= MAX_PAGINATION_PASSES) {
    console.warn(
      `FreeWord pagination did not settle within ${MAX_PAGINATION_PASSES} passes; stopping to avoid a jitter loop.`,
    );
  }
  if (appliedCount > 0 && typeof window !== "undefined") {
    // The status bar's `Page X / Y` and the toolbar's table states are derived
    // from the page structure this pass just rewrote — and `update` no longer
    // fires for a layout-only transaction, so say it out loud instead.
    window.dispatchEvent(new Event("fw:layout-changed"));
  }
  return appliedCount > 0;
}
