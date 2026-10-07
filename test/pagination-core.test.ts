import { describe, expect, it } from "vitest";
import { Schema } from "@tiptap/pm/model";
import type { Node as PMNode } from "@tiptap/pm/model";
import { blockPositionsOfPage, computePageLayout, layoutPages } from "../src/renderer/src/editor/pagination";
import type { BlockInfo } from "../src/renderer/src/editor/pagination";

const B = (pos: number, h: number, isPageBreak = false): BlockInfo => ({ pos, h, isPageBreak });

describe("computePageLayout", () => {
  it("no moves when everything fits", () => {
    const moves = computePageLayout([[B(1, 20), B(12, 30)], [B(44, 40)]], 100);
    expect(moves).toEqual([]);
  });

  it("respects the tolerance window", () => {
    // 101px block in a 100px page is within the 2px tolerance: no move.
    const moves = computePageLayout([[B(1, 101)]], 100);
    expect(moves).toEqual([]);
  });

  it("moves overflowing blocks to a new page", () => {
    const moves = computePageLayout([[B(1, 60), B(12, 60)]], 100);
    expect(moves).toEqual([{ fromPageIdx: 0, pos: 12, toPageIdx: 1 }]);
  });

  it("keeps an oversized single block in place and moves the rest", () => {
    const moves = computePageLayout([[B(1, 50), B(12, 200), B(33, 30)]], 100);
    expect(moves).toEqual([{ fromPageIdx: 0, pos: 33, toPageIdx: 1 }]);
  });

  it("barrier splits the remainder into a new page", () => {
    const moves = computePageLayout([[B(1, 30), B(13, 20, true), B(34, 30), B(65, 30)]], 100);
    expect(moves).toEqual([
      { fromPageIdx: 0, pos: 34, toPageIdx: 1 },
      { fromPageIdx: 0, pos: 65, toPageIdx: 1 },
    ]);
  });

  it("two consecutive barriers preserve a forced blank page", () => {
    const moves = computePageLayout([[B(1, 30), B(13, 20, true), B(34, 20, true), B(55, 30)]], 100);
    expect(moves).toEqual([
      { fromPageIdx: 0, pos: 34, toPageIdx: 1 },
      { fromPageIdx: 0, pos: 55, toPageIdx: 2 },
    ]);
    const layout = layoutPages([[B(1, 30), B(13, 20, true), B(34, 20, true), B(55, 30)]], 100);
    expect(layout).toEqual([[B(1, 30), B(13, 20, true)], [B(34, 20, true)], [B(55, 30)]]);
  });

  it("moves a barrier block to the next page when it no longer fits", () => {
    const layout = layoutPages([[B(1, 90), B(12, 20, true)]], 100);
    expect(layout).toEqual([[B(1, 90)], [B(12, 20, true)]]);
  });
});

describe("blockPositionsOfPage", () => {
  // The driver feeds these positions to `view.nodeDOM` and `doc.nodeAt`, both
  // of which take ABSOLUTE positions while `Node.forEach` reports offsets
  // relative to the page — a page after the first used to resolve to the wrong
  // node (another page, a foreign block, or nothing), which corrupted the
  // re-paginated document on save.
  const schema = new Schema({
    nodes: {
      doc: { content: "page+" },
      page: { content: "block+", group: "block" },
      paragraph: { content: "text*", group: "block" },
      pageBreak: { group: "block" },
      text: { group: "inline" },
    },
  });
  const para = (s: string) => schema.nodes.paragraph.create(null, [schema.text(s)]);

  const positions = (node: PMNode, pagePos: number): number[] =>
    blockPositionsOfPage(node, pagePos).map((b) => b.pos);

  it("returns absolute positions that resolve back to the same blocks", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.page.create(null, [para("aaa"), para("bb")]),
      schema.nodes.page.create(null, [para("cccc"), para("dddddd")]),
    ]);
    const pagePos: number[] = [];
    doc.forEach((_, pos) => pagePos.push(pos));

    let index = 0;
    doc.forEach((pageNode, pos) => {
      const blocks: PMNode[] = [];
      pageNode.forEach((c) => blocks.push(c));
      positions(pageNode, pos).forEach((abs, i) => {
        const resolved = doc.nodeAt(abs);
        expect(resolved, `page ${index} block ${i}`).not.toBeNull();
        expect(resolved!.type.name, `page ${index} block ${i}`).not.toBe("page");
        expect(resolved!.textContent, `page ${index} block ${i}`).toBe(blocks[i].textContent);
      });
      index++;
    });
    expect(index).toBe(2);
  });

  it("matches the absolute positions the layout functions expect", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.page.create(null, [para("aaa"), para("bb")]),
      schema.nodes.page.create(null, [para("cccc")]),
    ]);
    const first = doc.child(0);
    // Ground truth from the model itself.
    expect(positions(first, 0)).toEqual([1, 6]);
    const second = doc.child(1);
    expect(positions(second, 11)).toEqual([12]);
  });

  it("flags page-break blocks", () => {
    const page = schema.nodes.page.create(null, [para("a"), schema.nodes.pageBreak.create(), para("b")]);
    expect(blockPositionsOfPage(page, 0)).toEqual([
      { pos: 1, isPageBreak: false },
      { pos: 4, isPageBreak: true },
      { pos: 5, isPageBreak: false },
    ]);
  });
});
