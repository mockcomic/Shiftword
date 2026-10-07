import { describe, expect, it } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import { toDocxBuffer } from "../src/main/docx-export";

const fixture = {
  type: "doc",
  attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
  content: [
    {
      type: "page",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "Hello " },
            { type: "text", text: "bold", marks: [{ type: "bold" }] },
            {
              type: "text",
              text: " styled",
              marks: [
                {
                  type: "textStyle",
                  attrs: { color: "#FF0000", fontFamily: "Times New Roman", fontSize: "14pt" },
                },
                { type: "highlight", attrs: { color: "#FFFF00" } },
              ],
            },
          ],
        },
        { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Title" }] },
        {
          type: "bulletList",
          content: [
            {
              type: "listItem",
              content: [{ type: "paragraph", content: [{ type: "text", text: "item one" }] }],
            },
            {
              type: "listItem",
              content: [
                { type: "paragraph", content: [{ type: "text", text: "item two" }] },
                {
                  type: "bulletList",
                  content: [
                    {
                      type: "listItem",
                      content: [{ type: "paragraph", content: [{ type: "text", text: "nested" }] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
        {
          type: "orderedList",
          content: [
            {
              type: "listItem",
              content: [{ type: "paragraph", content: [{ type: "text", text: "first" }] }],
            },
          ],
        },
        { type: "pageBreak" },
      ],
    },
    {
      type: "page",
      content: [{ type: "paragraph", content: [{ type: "text", text: "Page two content" }] }],
    },
    {
      type: "page",
      content: [{ type: "paragraph", content: [{ type: "text", text: "Page three" }] }],
    },
  ],
};

describe("toDocxBuffer", () => {
  it("exports a docx with formatting, lists, page breaks and A4 page size", async () => {
    const buf = await toDocxBuffer(fixture);
    const zip = unzipSync(new Uint8Array(buf));
    const docXml = strFromU8(zip["word/document.xml"]);
    const stylesXml = strFromU8(zip["word/styles.xml"]);

    // Page size (A4) and margins from doc attrs.
    expect(docXml).toContain('w:w="11906"');
    expect(docXml).toContain('w:h="16838"');
    expect(docXml).toContain('w:top="1440"');

    // Marks.
    expect(docXml).toContain("<w:b/>");
    expect(docXml).toContain('w:color w:val="FF0000"');
    expect(docXml).toContain('w:sz w:val="28"');
    expect(docXml).toContain("Times New Roman");
    expect(docXml).toContain('w:highlight w:val="yellow"');
    // Default run styling is style-based now (Calibri 11pt = 22 half-points
    // in the document defaults), NOT stamped on every run — a run-level
    // w:sz would override the Heading 1–3 styles in Word/Google Docs.
    expect(stylesXml).toContain('w:ascii="Calibri"');
    expect(stylesXml).toContain('w:sz w:val="22"');
    expect(docXml).not.toContain('w:sz w:val="22"');
    // Headings mirror the editor CSS (24/18/14 pt, black, bold) via styles.
    expect(stylesXml).toContain('w:sz w:val="48"');
    expect(stylesXml).toContain('w:sz w:val="36"');
    expect(stylesXml).toContain('w:color w:val="000000"');
    // Paragraph gap mirrors `margin: 0 0 6pt` (120 twips) on the styles.
    expect(stylesXml).toContain('w:line="276"');
    expect(docXml).toContain('w:after="120"');
    // The named Fw* block styles exist for Word's gallery.
    expect(stylesXml).toContain('w:styleId="FwCallout"');
    expect(stylesXml).toContain('w:styleId="FwCodeBlock"');

    // Structure.
    expect(docXml).toContain('w:val="Heading1"');
    expect(docXml).toContain("w:numId");
    expect(docXml).toContain("item one");
    expect(docXml).toContain("nested");

    // Exactly two page breaks: the explicit pageBreak at the end of page 1
    // plus the boundary break between pages 2 and 3. No double break after
    // page 1 (no-double-break rule).
    const pageBreaks = docXml.match(/w:type="page"/g) ?? [];
    expect(pageBreaks).toHaveLength(2);
  });

  it("exports tables with header rows, borders and empty cells", async () => {
    const doc = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        {
          type: "page",
          content: [
            {
              type: "table",
              content: [
                {
                  type: "tableRow",
                  content: [
                    {
                      type: "tableHeader",
                      content: [{ type: "paragraph", content: [{ type: "text", text: "Head A" }] }],
                    },
                    {
                      type: "tableHeader",
                      content: [{ type: "paragraph", content: [{ type: "text", text: "Head B" }] }],
                    },
                  ],
                },
                {
                  type: "tableRow",
                  content: [
                    {
                      type: "tableCell",
                      content: [{ type: "paragraph", content: [{ type: "text", text: "Body 1" }] }],
                    },
                    // Empty cell: OOXML requires a paragraph; the exporter must add one.
                    { type: "tableCell", content: [] },
                  ],
                },
              ],
            },
            { type: "paragraph", content: [{ type: "text", text: "after table" }] },
          ],
        },
      ],
    };
    const buf = await toDocxBuffer(doc);
    const zip = unzipSync(new Uint8Array(buf));
    const docXml = strFromU8(zip["word/document.xml"]);

    expect(docXml).toContain("<w:tbl>");
    expect(docXml).toContain("Head A");
    expect(docXml).toContain("Head B");
    expect(docXml).toContain("Body 1");
    // Header row is marked as a repeating table header.
    expect(docXml).toContain("<w:tblHeader/>");
    // Grid borders on the table.
    expect(docXml).toContain("<w:tblBorders>");
    expect(docXml).toContain('w:val="single"');
    // Table is full page width (percentage).
    expect(docXml).toContain('w:type="pct"');
    // The empty cell still produced a paragraph (count of tc > 3, each with a w:p).
    expect((docXml.match(/<w:tc>/g) ?? []).length).toBe(4);
    expect(docXml.indexOf("after table") > docXml.indexOf("</w:tbl>")).toBe(true);
  });

  it("exports images as embedded media with the requested size", async () => {
    // 1×1 transparent PNG.
    const png =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    const doc = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        {
          type: "page",
          content: [
            { type: "paragraph", content: [{ type: "text", text: "before" }] },
            { type: "image", attrs: { src: `data:image/png;base64,${png}`, width: 120, height: 60 } },
            { type: "paragraph", content: [{ type: "text", text: "after" }] },
          ],
        },
      ],
    };
    const buf = await toDocxBuffer(doc);
    const zip = unzipSync(new Uint8Array(buf));
    const docXml = strFromU8(zip["word/document.xml"]);

    // An inline drawing is emitted in the document body.
    expect(docXml).toContain("<w:drawing>");
    expect(docXml).toContain("<wp:inline");
    // Requested 120×60 px → 120*9525 × 60*9525 EMU.
    expect(docXml).toContain('cx="1143000"');
    expect(docXml).toContain('cy="571500"');
    // The image bytes are embedded under word/media/<hash>.png.
    const mediaKeys = Object.keys(zip).filter((k) => k.startsWith("word/media/") && k.endsWith(".png"));
    expect(mediaKeys).toHaveLength(1);
    expect(zip[mediaKeys[0]].length).toBe(Buffer.from(png, "base64").length);
  });

  it("clamps oversized images to the page content width", async () => {
    const png =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    const doc = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        {
          type: "page",
          content: [
            { type: "image", attrs: { src: `data:image/png;base64,${png}`, width: 5000, height: 5000 } },
          ],
        },
      ],
    };
    const buf = await toDocxBuffer(doc);
    const zip = unzipSync(new Uint8Array(buf));
    const docXml = strFromU8(zip["word/document.xml"]);
    // A4 content width = 210 − 2×25.4 = 159.2 mm → 601.5 px at 96 DPI → 5,729,793 EMU (rounded).
    const m = docXml.match(/cx="(\d+)"/);
    expect(m).not.toBeNull();
    expect(Number(m?.[1])).toBeLessThanOrEqual(6020 * 9525);
  });

  // ---------------------------------------------------------------------------
  // Regression: saving must never abort on degenerate input. Each of these
  // used to throw out of the `docx:save` handler, which nobody caught — the
  // save silently did nothing (and a save-during-quit hung the app).
  // ---------------------------------------------------------------------------

  it("survives degenerate tables instead of aborting the save", async () => {
    // `docx` throws RangeError: Invalid array length for a table with no rows.
    const degenerate: Array<[string, unknown[]]> = [
      ["no rows", []],
      ["rows without cells", [{ type: "tableRow", content: [] }]],
      ["only non-row children", [{ type: "paragraph", content: [{ type: "text", text: "stray" }] }]],
      ["row with only unsupported cells", [{ type: "tableRow", content: [{ type: "paragraph" }] }]],
    ];

    for (const [label, content] of degenerate) {
      const doc = {
        type: "doc",
        attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
        content: [{ type: "page", content: [{ type: "table", content }] }],
      };
      const buf = await toDocxBuffer(doc as unknown as Record<string, unknown>);
      const docXml = strFromU8(unzipSync(new Uint8Array(buf))["word/document.xml"]);
      // Each degenerate table is exported as exactly one row with one cell.
      expect(docXml, label).toContain("<w:tbl>");
      expect((docXml.match(/<w:tr>/g) ?? []).length, label).toBe(1);
      expect((docXml.match(/<w:tc>/g) ?? []).length, label).toBe(1);
    }
  });

  it("keeps surrounding content when a table is degenerate", async () => {
    const doc = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        {
          type: "page",
          content: [
            { type: "paragraph", content: [{ type: "text", text: "before table" }] },
            { type: "table", content: [] },
            { type: "paragraph", content: [{ type: "text", text: "after table" }] },
          ],
        },
      ],
    };
    const buf = await toDocxBuffer(doc);
    const docXml = strFromU8(unzipSync(new Uint8Array(buf))["word/document.xml"]);
    expect(docXml).toContain("before table");
    expect(docXml).toContain("after table");
    expect(docXml.indexOf("before table") < docXml.indexOf("after table")).toBe(true);
  });

  it("clamps nonsense colspan/rowspan values to whole cells", async () => {
    const cellWith = (attrs: Record<string, unknown>) => ({
      type: "tableCell",
      attrs,
      content: [{ type: "paragraph", content: [{ type: "text", text: "c" }] }],
    });
    const doc = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        {
          type: "page",
          content: [
            {
              type: "table",
              content: [
                {
                  type: "tableRow",
                  content: [cellWith({ colspan: 0, rowspan: -3 }), cellWith({ colspan: "x", rowspan: 1.9 })],
                },
              ],
            },
          ],
        },
      ],
    };
    const buf = await toDocxBuffer(doc);
    const docXml = strFromU8(unzipSync(new Uint8Array(buf))["word/document.xml"]);
    // Every span collapses to a whole cell (docx always writes gridSpan, so
    // the assertion is about the values, not their presence).
    const spans = [...docXml.matchAll(/<w:(?:gridSpan|vMerge) w:val="([^"]*)"/g)].map((m) => m[1]);
    expect(spans.length).toBeGreaterThan(0);
    expect(
      spans.every((v) => /^\d+$/.test(v) && Number(v) >= 1),
      spans.join(","),
    ).toBe(true);
    expect((docXml.match(/<w:tc>/g) ?? []).length).toBe(2);
  });

  it("keeps a legitimate colspan as a gridSpan", async () => {
    const doc = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [
        {
          type: "page",
          content: [
            {
              type: "table",
              content: [
                {
                  type: "tableRow",
                  content: [
                    {
                      type: "tableCell",
                      attrs: { colspan: 2, rowspan: 1 },
                      content: [{ type: "paragraph", content: [{ type: "text", text: "wide" }] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };
    const buf = await toDocxBuffer(doc);
    const docXml = strFromU8(unzipSync(new Uint8Array(buf))["word/document.xml"]);
    expect(docXml).toContain('<w:gridSpan w:val="2"/>');
  });

  it("tolerates a malformed document payload crossing IPC", async () => {
    for (const payload of [null, undefined, "not a doc", 42, {}, { type: "doc", attrs: "bad" }]) {
      const buf = await toDocxBuffer(payload as unknown as Record<string, unknown>);
      const docXml = strFromU8(unzipSync(new Uint8Array(buf))["word/document.xml"]);
      expect(docXml).toContain("<w:document");
    }
  });
});
