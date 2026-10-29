// @vitest-environment jsdom
// Full import integration: a .docx produced by the `docx` library (table +
// embedded image) is rendered with docx-preview and converted back into a
// Shiftword doc JSON via docxBufferToDoc.
import { describe, expect, it } from "vitest";
import { Document, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx";
import { unzipSync, strFromU8 } from "fflate";
import { getSchema } from "@tiptap/core";
import { createExtensions } from "../src/renderer/src/editor/extensions";
import { docxBufferToDoc } from "../src/renderer/src/io/import-docx";
import { toDocxBuffer } from "../src/main/docx-export";

const schema = getSchema(createExtensions());

// 1×1 transparent PNG.
const PNG_1X1 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

interface PMJson {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PMJson[];
  text?: string;
}

/** Build a .docx buffer with a 2×2 table and an inline 100×100 px PNG. */
async function makeDocxBuffer(): Promise<Uint8Array> {
  const doc = new Document({
    sections: [
      {
        properties: { page: { size: { width: 11906, height: 16838 } } },
        children: [
          new Paragraph({ children: [new TextRun("Before table")] }),
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ children: [new TextRun("A1")] })] }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun("A2")] })] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ children: [new TextRun("B1")] })] }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun("B2")] })] }),
                ],
              }),
            ],
          }),
          new Paragraph({
            children: [
              new ImageRun({
                type: "png",
                data: PNG_1X1,
                transformation: { width: 100, height: 100 },
              }),
            ],
          }),
          new Paragraph({ children: [new TextRun("After image")] }),
        ],
      },
    ],
  });
  return Packer.toBuffer(doc);
}

describe("docx round-trip import (docx-preview)", () => {
  it("imports tables and embedded images from a real .docx buffer", async () => {
    const buf = await makeDocxBuffer();
    const json = (await docxBufferToDoc(buf, schema)) as unknown as {
      type: string;
      content: PMJson[];
    };

    expect(json.type).toBe("doc");
    expect(json.content).toHaveLength(1);
    const page = json.content[0];
    expect(page.type).toBe("page");

    const types = page.content!.map((c) => c.type);
    expect(types).toContain("paragraph");
    expect(types).toContain("table");
    expect(types).toContain("image");

    // Table keeps its cells.
    const table = page.content!.find((c) => c.type === "table")!;
    expect(table.content!.length).toBe(2); // two rows
    expect(table.content![0].content!.length).toBe(2); // two cells

    // The image comes back as a data: URL (not a blob: URL).
    const img = page.content!.find((c) => c.type === "image")!;
    expect(String(img.attrs?.src)).toMatch(/^data:image\/png;base64,.+$/);
    // 100 px was stored in the docx ext; docx-preview renders it as
    // css pt which normalizeImages converts back to ~100 px.
    expect(Number(img.attrs?.width)).toBeGreaterThan(0);
    expect(Number(img.attrs?.height)).toBeGreaterThan(0);
  }, 30000);
});

// ---------------------------------------------------------------------------
// Formatting fidelity: what the exporter writes, the importer must read back
// — the Shiftword ⇄ Word ⇄ Google Docs consistency contract.
// ---------------------------------------------------------------------------

function docOf(attrs: Record<string, unknown>, blocks: unknown[]) {
  return { type: "doc", attrs, content: [{ type: "page", content: blocks }] };
}
const PAGE_ATTRS = { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 };

interface PMJson {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PMJson[];
  text?: string;
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>;
}

describe("docx round-trip formatting (export → docx-preview → import)", () => {
  it("keeps user-set run font, size and colour as textStyle marks", async () => {
    const json = docOf(PAGE_ATTRS, [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "styled",
            marks: [
              {
                type: "textStyle",
                attrs: { color: "#C00000", fontFamily: "Georgia", fontSize: "14pt" },
              },
            ],
          },
        ],
      },
    ]);
    const back = (await docxBufferToDoc(await toDocxBuffer(json), schema)) as unknown as {
      content: PMJson[];
    };
    const text = back.content[0].content![0].content!.find((c) => c.text === "styled")!;
    const ts = text.marks?.find((m) => m.type === "textStyle")?.attrs ?? {};
    expect(String(ts.fontFamily)).toMatch(/Georgia/i);
    expect(String(ts.fontSize)).toMatch(/^14pt$/i);
    // docx-preview renders colours as rgb(); the exporter accepts both forms.
    expect(String(ts.color)).toMatch(/rgb\(192,\s*0,\s*0\)|C00000/i);
  }, 30000);

  it("imports run shading (w:highlight) as a highlight mark that re-exports", async () => {
    const json = docOf(PAGE_ATTRS, [
      {
        type: "paragraph",
        content: [
          { type: "text", text: "plain " },
          { type: "text", text: "marked", marks: [{ type: "highlight", attrs: { color: "#FFFF00" } }] },
        ],
      },
    ]);
    const back = (await docxBufferToDoc(await toDocxBuffer(json), schema)) as unknown as {
      content: PMJson[];
    };
    const marked = back.content[0].content![0].content!.find((c) => (c.text ?? "").includes("marked"))!;
    expect(marked.marks?.some((m) => m.type === "highlight")).toBe(true);

    // And it must survive a second export (the colour arrives as rgb() form).
    const docXml = strFromU8(
      unzipSync(new Uint8Array(await toDocxBuffer(back as unknown as Record<string, unknown>)))[
        "word/document.xml"
      ],
    );
    expect(docXml).toContain('w:highlight w:val="yellow"');
  }, 30000);

  it("imports the real page size and margins instead of hardcoding A4", async () => {
    const json = docOf({ size: "letter", mt: 20, mr: 30, mb: 20, ml: 30 }, [
      { type: "paragraph", content: [{ type: "text", text: "US letter body" }] },
    ]);
    const back = (await docxBufferToDoc(await toDocxBuffer(json), schema)) as unknown as {
      type: string;
      attrs: Record<string, unknown>;
      content: PMJson[];
    };
    expect(back.attrs.size).toBe("letter");
    expect(Number(back.attrs.mt)).toBeCloseTo(20, 0);
    expect(Number(back.attrs.mr)).toBeCloseTo(30, 0);
    expect(Number(back.attrs.mb)).toBeCloseTo(20, 0);
    expect(Number(back.attrs.ml)).toBeCloseTo(30, 0);
  }, 30000);

  it("re-imports Notion-style blocks exported for Word", async () => {
    const json = docOf(PAGE_ATTRS, [
      {
        type: "taskList",
        content: [
          {
            type: "taskItem",
            attrs: { checked: true },
            content: [{ type: "paragraph", content: [{ type: "text", text: "done thing" }] }],
          },
          {
            type: "taskItem",
            attrs: { checked: false },
            content: [{ type: "paragraph", content: [{ type: "text", text: "open thing" }] }],
          },
        ],
      },
      {
        type: "callout",
        attrs: { icon: "💡" },
        content: [{ type: "paragraph", content: [{ type: "text", text: "remember this" }] }],
      },
      {
        type: "fwQuote",
        content: [{ type: "paragraph", content: [{ type: "text", text: "quoted" }] }],
      },
      { type: "divider" },
      {
        type: "fwCodeBlock",
        attrs: { language: "" },
        content: [{ type: "text", text: "const a = 1;\nconst b = 2;" }],
      },
    ]);
    const back = (await docxBufferToDoc(await toDocxBuffer(json), schema)) as unknown as {
      content: PMJson[];
    };
    const types = back.content[0].content!.map((c) => c.type);
    expect(types).toContain("taskList");
    expect(types).toContain("callout");
    expect(types).toContain("fwQuote");
    expect(types).toContain("divider");
    expect(types).toContain("fwCodeBlock");
  }, 30000);
});
