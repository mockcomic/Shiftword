import { describe, expect, it } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import { toDocxBuffer } from "../src/main/docx-export";

type PM = Record<string, unknown>;

const T = (text: string, marks?: PM[]): PM => ({ type: "text", text, ...(marks ? { marks } : {}) });
const P = (...children: PM[]): PM => ({
  type: "paragraph",
  ...(children.length ? { content: children } : {}),
});
const DOC = (...pages: PM[]): PM => ({
  type: "doc",
  attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
  content: pages.map((p) => ({ type: "page", content: Array.isArray(p) ? p : [p] })),
});
const page = (...blocks: PM[]): PM[] => blocks;

async function unzip(pm: PM): Promise<{ document: string; numbering: string; styles: string }> {
  const buf = await toDocxBuffer(pm);
  const files = unzipSync(new Uint8Array(buf));
  const get = (n: string): string => (files[n] ? strFromU8(files[n]) : "");
  return {
    document: get("word/document.xml"),
    numbering: get("word/numbering.xml"),
    styles: get("word/styles.xml"),
  };
}

// Minimal valid raster fixtures for header parsing -------------------------
const PNG_1x1 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
function bmp(w: number, h: number): string {
  const b = Buffer.alloc(54);
  b.write("BM", 0);
  b.writeInt32LE(w, 18);
  b.writeInt32LE(h, 22);
  return b.toString("base64");
}
function gif(w: number, h: number): string {
  const b = Buffer.alloc(13);
  b.write("GIF89a", 0);
  b.writeUInt16LE(w, 6);
  b.writeUInt16LE(h, 8);
  return b.toString("base64");
}
function jpeg(w: number, h: number): string {
  const b = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x00, 0x00, 0x00, 0x03, 0x01]);
  b.writeUInt16BE(h, 7);
  b.writeUInt16BE(w, 9);
  return b.toString("base64");
}
const img = (attrs: PM): PM => ({ type: "image", attrs });

describe("docx export — page setup", () => {
  it("letter size and margins map to canonical twips", async () => {
    const { document } = await unzip({
      type: "doc",
      attrs: { size: "letter", mt: 20, mr: 15, mb: 10, ml: 25 },
      content: [{ type: "page", content: [P(T("x"))] }],
    });
    expect(document).toMatch(/w:pgSz[^/]*w="12240"/);
    expect(document).toMatch(/w:pgSz[^/]*h="15840"/);
    expect(document).toMatch(/w:top="1134"/);
    expect(document).toMatch(/w:right="850"/);
    expect(document).toMatch(/w:bottom="567"/);
    expect(document).toMatch(/w:left="1417"/); // 25mm → 1417.32 → rounded
  });

  it("suppresses the between-page break when a page already ends with pageBreak", async () => {
    const { document } = await unzip(DOC(page(P(T("one")), { type: "pageBreak" }), page(P(T("two")))));
    const breaks = document.match(/w:type="page"/g) ?? [];
    expect(breaks).toHaveLength(1); // only the explicit one, no extra boundary break
  });

  it("adds exactly one boundary break between pages without explicit break", async () => {
    const { document } = await unzip(DOC(page(P(T("one"))), page(P(T("two")))));
    const breaks = document.match(/w:type="page"/g) ?? [];
    expect(breaks).toHaveLength(1);
  });
});

describe("docx export — runs and marks", () => {
  it("converts rgb() colors to hex", async () => {
    const { document } = await unzip(
      DOC(page(P(T("red", [{ type: "textStyle", attrs: { color: "rgb(255, 0, 0)" } }])))),
    );
    expect(document).toMatch(/w:color w:val="FF0000"/);
  });

  it("supports fractional point sizes", async () => {
    const { document } = await unzip(
      DOC(page(P(T("s", [{ type: "textStyle", attrs: { fontSize: "10.5pt" } }])))),
    );
    expect(document).toMatch(/w:sz w:val="21"/);
  });

  it("maps known highlight colors and drops unknown ones", async () => {
    const { document } = await unzip(
      DOC(
        page(
          P(
            T("green", [{ type: "highlight", attrs: { color: "#92D050" } }]),
            T(" bogus", [{ type: "highlight", attrs: { color: "#FF00FF" } }]),
          ),
        ),
      ),
    );
    expect(document).toMatch(/w:highlight w:val="green"/);
    expect(document).not.toMatch(/w:highlight w:val="blue|w:highlight w:val="magenta"/);
  });

  it("clamps heading levels above 3 to Heading3", async () => {
    const { document } = await unzip(
      DOC(page({ type: "heading", attrs: { level: 5 }, content: [T("deep")] })),
    );
    expect(document).toMatch(/w:pStyle w:val="Heading3"/);
  });

  it("writes alignment and line spacing", async () => {
    const { document } = await unzip(
      DOC(
        page({
          type: "paragraph",
          attrs: { textAlign: "center", lineSpacing: 1.5 },
          content: [T("c")],
        }),
      ),
    );
    expect(document).toMatch(/w:jc w:val="center"/);
    expect(document).toMatch(/w:line="360"/);
    expect(document).toMatch(/w:lineRule="auto"/);
  });
});

describe("docx export — lists", () => {
  it("restarts numbering for each separate ordered list (Word semantics)", async () => {
    const { document } = await unzip(
      DOC(
        page(
          { type: "orderedList", content: [{ type: "listItem", content: [P(T("a"))] }] },
          P(T("between")),
          { type: "orderedList", content: [{ type: "listItem", content: [P(T("b"))] }] },
        ),
      ),
    );
    // Two list items ⇒ two numbering paragraphs ⇒ two different numIds. One
    // shared numId is exactly the defect: Word renders the second list as a
    // continuation (3, 4) of the first (1, 2).
    const numIds = [...document.matchAll(/<w:numId w:val="(\d+)"/g)].map((m) => m[1]);
    expect(numIds).toHaveLength(2);
    expect(new Set(numIds).size).toBe(2);
  });

  it("numbers a repeated list item from the shared sequence, not two references", async () => {
    const { document } = await unzip(
      DOC(
        page({
          type: "orderedList",
          content: [
            { type: "listItem", content: [P(T("one"))] },
            { type: "listItem", content: [P(T("two"))] },
          ],
        }),
      ),
    );
    // Items of the SAME list must keep counting, i.e. share one numId.
    const numIds = [...document.matchAll(/<w:numId w:val="(\d+)"/g)].map((m) => m[1]);
    expect(numIds).toHaveLength(2);
    expect(new Set(numIds).size).toBe(1);
  });

  it("defines numbering levels deep enough for nested ordered lists (4 levels)", async () => {
    const leaf: PM = { type: "listItem", content: [P(T("leaf"))] };
    const build = (level: number): PM =>
      level === 0
        ? leaf
        : {
            type: "listItem",
            content: [P(T(`n${level}`)), { type: "orderedList", content: [build(level - 1)] }],
          };
    const { document, numbering } = await unzip(
      DOC(page({ type: "orderedList", content: [build(3)] } as PM)),
    );
    const used = [
      ...new Set([...document.matchAll(/<w:ilvl w:val="(\d+)"/g)].map((m) => Number(m[1]))),
    ].sort();
    expect(used).toContain(3);
    // A level is only "defined" if Word gets a decimal format and a `%N.`
    // numbering text for it — prosemirror's defaults would silently render
    // deep items as bullets.
    for (const level of used) {
      const lvl = new RegExp(
        `<w:lvl w:ilvl="${level}"[\\s\\S]*?<w:numFmt w:val="decimal"\\/>[\\s\\S]*?<w:lvlText w:val="%${level + 1}\\."`,
      );
      expect(numbering, `level ${level} has no decimal definition`).toMatch(lvl);
    }
  });
});

describe("docx export — images", () => {
  it("reads natural size from PNG headers", async () => {
    const { document } = await unzip(DOC(page(img({ src: `data:image/png;base64,${PNG_1x1}` }))));
    expect(document).toMatch(/<wp:extent cx="9525" cy="9525"/);
  });

  it("reads natural size from BMP / GIF / JPEG headers", async () => {
    const { document } = await unzip(
      DOC(
        page(
          img({ src: `data:image/bmp;base64,${bmp(2, 3)}`, width: 0, height: 0 }),
          img({ src: `data:image/gif;base64,${gif(4, 5)}` }),
          img({ src: `data:image/jpeg;base64,${jpeg(6, 7)}` }),
        ),
      ),
    );
    // 2x3, 4x5, 6x7 px in EMU (× 9525)
    expect(document).toMatch(/cx="19050" cy="28575"/);
    expect(document).toMatch(/cx="38100" cy="47625"/);
    expect(document).toMatch(/cx="57150" cy="66675"/);
  });

  it("clamps oversized images to the page content width", async () => {
    const { document } = await unzip(
      DOC(page(img({ src: `data:image/png;base64,${PNG_1x1}`, width: 5000, height: 5000 }))),
    );
    const m = document.match(/<wp:extent cx="(\d+)" cy="(\d+)"/);
    expect(m).toBeTruthy();
    // A4 default margins → content width ≈ 601.7px; 5000px image scaled to fit.
    const cx = Number(m![1]);
    expect(cx).toBeLessThanOrEqual(603 * 9525);
    expect(cx).toBeGreaterThan(599 * 9525);
  });

  it("unsupported image sources degrade to an empty paragraph", async () => {
    const { document } = await unzip(DOC(page(img({ src: "blob:http://x/y" }))));
    expect(document).not.toMatch(/<wp:extent/);
  });
});

describe("docx export — tables", () => {
  it("writes header flags and colspans", async () => {
    const cell = (hdr: boolean, colspan: number): PM => ({
      type: hdr ? "tableHeader" : "tableCell",
      attrs: { colspan, rowspan: 1 },
      content: [P(T("c"))],
    });
    const { document } = await unzip(
      DOC(
        page({
          type: "table",
          content: [
            { type: "tableRow", content: [cell(true, 2)] },
            { type: "tableRow", content: [cell(false, 1), cell(false, 1)] },
          ],
        }),
      ),
    );
    expect(document).toMatch(/<w:tblHeader\/>/);
    expect(document).toMatch(/<w:gridSpan w:val="2"\/>/);
  });

  it("keeps every table cell with at least one paragraph", async () => {
    const { document } = await unzip(
      DOC(
        page({
          type: "table",
          content: [{ type: "tableRow", content: [{ type: "tableCell" }] }],
        }),
      ),
    );
    const tc = document.match(/<w:tc>.*?<\/w:tc>/s);
    expect(tc).toBeTruthy();
    expect(tc![0]).toMatch(/<w:p\/>|<w:p>.*<\/w:p>/s);
  });

  it("writes the width the user dragged the column to, in twips", async () => {
    const cell = (colwidth: number[] | null): PM => ({
      type: "tableCell",
      attrs: { colspan: 1, rowspan: 1, ...(colwidth ? { colwidth } : {}) },
      content: [P(T("c"))],
    });
    const { document } = await unzip(
      DOC(
        page({
          type: "table",
          content: [{ type: "tableRow", content: [cell([100]), cell([200]), cell(null)] }],
        }),
      ),
    );
    // 100 px × 15 twips/px = 1500; 200 px = 3000. A column nobody resized gets
    // no explicit width, so Word can still lay it out automatically.
    expect(document).toMatch(/<w:tcW w:type="dxa" w:w="1500"\/>/);
    expect(document).toMatch(/<w:tcW w:type="dxa" w:w="3000"\/>/);
    expect(document.match(/<w:tcW/g)).toHaveLength(2);
  });

  it("writes vertical merges (rowspan) so merged cells round-trip", async () => {
    const cell = (rowspan: number): PM => ({
      type: "tableCell",
      attrs: { colspan: 1, rowspan },
      content: [P(T(`r${rowspan}`))],
    });
    const { document } = await unzip(
      DOC(
        page({
          type: "table",
          content: [
            { type: "tableRow", content: [cell(2), cell(1)] },
            { type: "tableRow", content: [cell(1)] },
          ],
        }),
      ),
    );
    expect(document).toMatch(/<w:vMerge w:val="restart"\/>/);
    expect(document).toMatch(/<w:vMerge\/>|<w:vMerge w:val="continue"\/>/);
  });
});

describe("docx export — degenerate documents", () => {
  it("does not crash on a doc with no pages", async () => {
    await expect(toDocxBuffer({ type: "doc", attrs: {}, content: [] })).resolves.toBeInstanceOf(Buffer);
  });

  it("does not crash on a page with no blocks", async () => {
    await expect(
      toDocxBuffer({
        type: "doc",
        attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
        content: [{ type: "page" }],
      }),
    ).resolves.toBeInstanceOf(Buffer);
  });
});
