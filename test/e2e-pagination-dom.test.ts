// @vitest-environment jsdom
/**
 * DOM-driver tests for the pagination engine. These drive `runPagination`
 * with a real ProseMirror state but a spy on `view.nodeDOM`, so we can verify
 * (a) the positions the engine asks the view to measure, and (b) that the
 * applied layout preserves block content and selection.
 */
import { describe, expect, it } from "vitest";
import { getSchema } from "@tiptap/core";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import type { Editor } from "@tiptap/core";
import { createExtensions } from "../src/renderer/src/editor/extensions";
import { runPagination } from "../src/renderer/src/editor/pagination";

const schema = getSchema(createExtensions());

type PM = Record<string, unknown>;
const para = (text: string): PM => ({ type: "paragraph", content: [{ type: "text", text }] });
const emptyPara = (): PM => ({ type: "paragraph" });
const docJson = (pages: PM[][]): PM => ({
  type: "doc",
  attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
  content: pages.map((blocks) => ({ type: "page", content: blocks })),
});

/**
 * Build a fake editor whose `view.nodeDOM(pos)` returns an element ONLY for
 * positions that genuinely address that node (mirroring real ProseMirror
 * addressing), recording every position that was requested.
 *
 * `pageHeight` → clientHeight of page elements; blockHeights → rect heights
 * keyed by block text (defaults 10).
 */
function makeFake(doc: PM, opts: { pageHeight?: number; blockHeights?: Record<string, number> }) {
  const node = schema.nodeFromJSON(doc);
  node.check();
  let state = EditorState.create({ doc: node, selection: TextSelection.atStart(node) });

  const elements = new Map<object, HTMLElement>();
  node.forEach((page, _pagePos) => {
    const pe = document.createElement("div");
    Object.defineProperty(pe, "clientHeight", { get: () => opts.pageHeight ?? 1000 });
    elements.set(page, pe);
    page.forEach((child) => {
      const ce = document.createElement("div");
      const text = child.textContent.trim();
      const h = opts.blockHeights?.[text] ?? 10;
      ce.getBoundingClientRect = () => ({ height: h }) as DOMRect;
      elements.set(child, ce);
    });
  });

  const requested: number[] = [];
  const dispatched: Array<{ getMeta(k: string): unknown; doc: unknown }> = [];

  const fake = {
    isDestroyed: false,
    get state() {
      return state;
    },
    view: {
      nodeDOM(pos: number): HTMLElement | null {
        requested.push(pos);
        const n = state.doc.nodeAt(pos);
        return (n && elements.get(n)) ?? null;
      },
      dispatch(tr: { getMeta(k: string): unknown; doc: unknown }) {
        dispatched.push(tr);
        state = state.apply(tr as never);
      },
    },
  };
  return {
    editor: fake as unknown as Editor,
    requested,
    dispatched,
    state: () => state,
    select: (pos: number) => {
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    },
  };
}

/** Every block's absolute position: pageStart + 1 + offset-in-content. */
function blockPositions(doc: PM): number[] {
  const node = schema.nodeFromJSON(doc);
  const out: number[] = [];
  node.forEach((page, pagePos) => {
    let offset = 1; // inside the page node
    page.forEach((child) => {
      out.push(pagePos + offset);
      offset += child.nodeSize;
    });
  });
  return out;
}

describe("pagination DOM driver — measurement addressing", () => {
  it("asks the view for every block at its ABSOLUTE position", () => {
    const json = docJson([
      [para("a1"), para("a2")],
      [para("b1"), para("b2")],
    ]);
    const f = makeFake(json, { pageHeight: 1000, blockHeights: { a1: 10, a2: 10, b1: 10, b2: 10 } });
    runPagination(f.editor);

    const expected = blockPositions(json);
    const missing = expected.filter((p) => !f.requested.includes(p));
    // All block positions must be measured directly; the page nodes are also
    // legitimately requested. Anything NOT in the expected+page set would be
    // mis-addressed, but the clearest signal is `missing`.
    expect(missing).toEqual([]);
  });

  it("leaves a fitting multi-page document untouched", () => {
    const json = docJson([[para("a1")], [para("b1")]]);
    const f = makeFake(json, { pageHeight: 1000, blockHeights: { a1: 10, b1: 10 } });
    const changed = runPagination(f.editor);
    expect(changed).toBe(false);
    expect(f.state().doc.childCount).toBe(2);
  });

  it("marks layout transactions out of the undo history and out of `update`", () => {
    const json = docJson([[para("a1"), para("a2"), para("a3"), para("a4")]]);
    // contentHeight 1000; each block 400 → overflow after the second block.
    const f = makeFake(json, {
      pageHeight: 1000,
      blockHeights: { a1: 400, a2: 400, a3: 400, a4: 400 },
    });
    runPagination(f.editor);
    expect(f.dispatched.length).toBeGreaterThan(0);
    for (const tr of f.dispatched) {
      expect(tr.getMeta("addToHistory")).toBe(false);
      // Tiptap suppresses its `update` event on this meta — which is what keeps
      // a re-pagination from marking a freshly opened document dirty.
      expect(tr.getMeta("preventUpdate")).toBe(true);
    }
  });

  it("moves overflowing blocks to a new page and keeps their content", () => {
    const json = docJson([[para("one"), para("two"), para("three")]]);
    const f = makeFake(json, { pageHeight: 1000, blockHeights: { one: 400, two: 400, three: 400 } });
    runPagination(f.editor);
    const doc = f.state().doc;
    expect(doc.childCount).toBe(2);
    expect(doc.child(0).child(0).textContent).toBe("one");
    expect(doc.child(0).child(1).textContent).toBe("two");
    expect(doc.child(1).child(0).textContent).toBe("three");
  });

  it("keeps the caret in the moved block after re-pagination", () => {
    const json = docJson([[para("one"), para("two"), para("three")]]);
    const f = makeFake(json, { pageHeight: 1000, blockHeights: { one: 400, two: 400, three: 400 } });
    // Put the caret inside "three" (absolute position of its text start).
    const doc = schema.nodeFromJSON(json);
    const pos = doc.nodeSize - 3; // near the very end
    f.select(pos);
    runPagination(f.editor);
    const head = f.state().selection.head;
    const newDoc = f.state().doc;
    const page2Start = newDoc.child(0).nodeSize; // position of page 2 node
    expect(head).toBeGreaterThanOrEqual(page2Start + 1);
  });

  it("absorbs a trailing empty page and keeps the caret in the document", () => {
    const json = docJson([[para("last content")], [emptyPara()]]);
    const f = makeFake(json, { pageHeight: 1000, blockHeights: { "last content": 10 } });
    expect(() => runPagination(f.editor)).not.toThrow();
    expect(f.state().doc.childCount).toBe(1);
    // The caret must land somewhere valid (not silently flung to the top of
    // an unrelated position); at minimum it must remain in the doc.
    const sel = f.state().selection;
    expect(sel.from).toBeGreaterThanOrEqual(0);
    expect(sel.to).toBeLessThanOrEqual(f.state().doc.nodeSize);
  });

  it("never splits an explicit pageBreak barrier page", () => {
    const json = docJson([[para("a"), { type: "pageBreak" }, para("b")]]);
    const f = makeFake(json, { pageHeight: 1000, blockHeights: { a: 10, b: 10 } });
    runPagination(f.editor);
    const doc = f.state().doc;
    expect(doc.childCount).toBe(2);
    expect(doc.child(0).lastChild!.type.name).toBe("pageBreak");
    expect(doc.child(1).child(0).textContent).toBe("b");
  });
});

describe("pagination DOM driver — tables the user merged and resized", () => {
  /** A 2-row table: a header cell spanning 2 columns, then a merged pair. */
  const cell = (text: string, attrs: PM = {}, header = false): PM => ({
    type: header ? "tableHeader" : "tableCell",
    attrs,
    content: [para(text)],
  });
  const table = (): PM => ({
    type: "table",
    content: [
      {
        type: "tableRow",
        content: [cell("tbl", { colspan: 2, rowspan: 1, colwidth: [140, 140] }, true)],
      },
      {
        type: "tableRow",
        content: [
          cell("below", { colspan: 1, rowspan: 2, colwidth: [140] }),
          cell("side", { colspan: 1, rowspan: 1, colwidth: [140] }),
        ],
      },
    ],
  });

  it("keeps merged cells and dragged column widths intact when the page reflows", () => {
    const json = docJson([[table(), para("filler")]]);
    // `textContent` of a table concatenates every cell's text, which is the key
    // the fake measures by.
    const f = makeFake(json, { pageHeight: 100, blockHeights: { tblbelowside: 80, filler: 80 } });
    const changed = runPagination(f.editor);
    expect(changed).toBe(true); // the two blocks no longer fit on one page

    const doc = f.state().doc;
    const all = JSON.stringify(doc.toJSON());
    // Nothing was dropped or flattened on the way across the page boundary.
    expect(all).toContain('"colspan":2');
    expect(all).toContain('"rowspan":2');
    expect(all).toContain('"colwidth":[140,140]');
    expect(doc.childCount).toBeGreaterThanOrEqual(2);

    // The table object itself is the same shape it went in with.
    const tables: PM[] = [];
    doc.forEach((page) =>
      page.forEach((block) => {
        if (block.type.name === "table") tables.push(block.toJSON() as unknown as PM);
      }),
    );
    expect(tables).toHaveLength(1);
    expect(JSON.stringify(tables[0])).toBe(JSON.stringify(schema.nodeFromJSON(table()).toJSON()));
  });
});
