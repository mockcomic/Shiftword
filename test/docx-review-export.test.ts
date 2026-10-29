// @vitest-environment jsdom
/**
 * Feature-complete exports: headers/footers (with PAGE/NUMPAGES fields),
 * footnotes (`footnotes.xml` + references), comments (`comments.xml`), and
 * native Word tracked changes (`w:ins`/`w:del`) — plus a round-trip of
 * everything Shiftword itself writes.
 */
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { getSchema } from "@tiptap/core";
import { toDocxBuffer } from "../src/main/docx-export";
import { docxBufferToDoc } from "../src/renderer/src/io/import-docx";
import { createExtensions } from "../src/renderer/src/editor/extensions";

const schema = getSchema(createExtensions());

type PM = Record<string, unknown>;

const richDoc: PM = {
  type: "doc",
  attrs: {
    size: "a4",
    mt: 25.4,
    mr: 25.4,
    mb: 25.4,
    ml: 25.4,
    header: "CONFIDENTIAL — {{page}} of {{pages}}",
    footer: "keepit.io",
    comments: { "c-1": { author: "Ada", text: "Why?", createdAt: 1700000000000, resolved: false } },
  },
  content: [
    {
      type: "page",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "plain " },
            {
              type: "text",
              text: "added",
              marks: [{ type: "trackedInsert", attrs: { author: "Ada", date: "2026-01-01T00:00:00.000Z" } }],
            },
            {
              type: "text",
              text: " and removed",
              marks: [{ type: "trackedDelete", attrs: { author: "Ada", date: "2026-01-01T00:00:00.000Z" } }],
            },
            { type: "footnoteRef", attrs: { text: "the note" } },
            { type: "text", text: ", commented", marks: [{ type: "comment", attrs: { id: "c-1" } }] },
          ],
        },
      ],
    },
  ],
};

describe("docx export — review features", () => {
  it("emits section headers/footers with PAGE/NUMPAGES fields", async () => {
    const buffer = await toDocxBuffer(richDoc as never);
    const zip = unzipSync(buffer);
    const names = Object.keys(zip);
    const headerFiles = names.filter((n) => /header\d+\.xml/.test(n));
    const footerFiles = names.filter((n) => /footer\d+\.xml/.test(n));
    expect(headerFiles.length).toBeGreaterThanOrEqual(1);
    expect(footerFiles.length).toBeGreaterThanOrEqual(1);
    const header = headerFiles.map((n) => strFromU8(zip[n]!)).join("\n");
    expect(header).toContain("<w:hdr");
    expect(header).toContain("CONFIDENTIAL");
    expect(header).toContain("w:instrText");
    expect(header).toContain("PAGE");
    expect(header).toContain("NUMPAGES");
  });

  it("emits footnotes.xml and in-text references", async () => {
    const buffer = await toDocxBuffer(richDoc as never);
    const zip = unzipSync(buffer);
    expect(zip["word/footnotes.xml"]).toBeTruthy();
    const footnotes = strFromU8(zip["word/footnotes.xml"]!);
    expect(footnotes).toContain("<w:footnote");
    expect(footnotes).toContain("the note");
    const body = strFromU8(zip["word/document.xml"]!);
    expect(body).toContain("w:footnoteReference");
  });

  it("emits comments.xml with the thread and range markers", async () => {
    const buffer = await toDocxBuffer(richDoc as never);
    const zip = unzipSync(buffer);
    expect(zip["word/comments.xml"]).toBeTruthy();
    const comments = strFromU8(zip["word/comments.xml"]!);
    expect(comments).toContain("Ada");
    expect(comments).toContain("Why?");
    const body = strFromU8(zip["word/document.xml"]!);
    expect(body).toContain("w:commentRangeStart");
    expect(body).toContain("w:commentRangeEnd");
    expect(body).toContain("w:commentReference");
  });

  it("emits native Word tracked changes (w:ins / w:del)", async () => {
    const buffer = await toDocxBuffer(richDoc as never);
    const zip = unzipSync(buffer);
    const body = strFromU8(zip["word/document.xml"]!);
    expect(body).toContain("<w:ins");
    expect(body).toContain("<w:del");
    expect(body).toContain('w:author="Ada"');
    // The revision content itself: the deleted text lives in w:delText.
    expect(body).toContain("<w:delText");
    expect(body).toContain("and removed");
  });

  it("leaves the review features out entirely when the doc has none", async () => {
    const buffer = await toDocxBuffer({
      type: "doc",
      content: [{ type: "page", content: [{ type: "paragraph" }] }],
    } as never);
    const zip = unzipSync(buffer);
    expect(Object.keys(zip).some((n) => /header\d+\.xml/.test(n))).toBe(false);
    expect(Object.keys(zip).some((n) => /footer\d+\.xml/.test(n))).toBe(false);
    // The docx library always carries the separator/continuation footnote
    // stubs — what must be absent is any *reference* to them or to comments.
    const body = strFromU8(zip["word/document.xml"]!);
    expect(body).not.toContain("w:footnoteReference");
    expect(body).not.toContain("w:commentRangeStart");
    expect(body).not.toContain("<w:ins");
    expect(body).not.toContain("<w:del");
  });
});

describe("docx import — round-trips what Shiftword wrote", () => {
  it("imports footnotes back as footnoteRef nodes carrying the note text", async () => {
    const doc: PM = {
      type: "doc",
      attrs: { size: "a4" },
      content: [
        {
          type: "page",
          content: [
            {
              type: "paragraph",
              content: [
                { type: "text", text: "body " },
                { type: "footnoteRef", attrs: { text: "deep thought" } },
              ],
            },
          ],
        },
      ],
    };
    const buffer = await toDocxBuffer(doc as never);
    const back = await docxBufferToDoc(new Uint8Array(buffer), schema);
    const json = JSON.stringify(back);
    expect(json).toContain("footnoteRef");
    expect(json).toContain("deep thought");
    // The rendered note list must not smuggle the number into the body text.
    expect(json).not.toMatch(/"text":"body 1/);
  });

  it("imports the exported header/footer text back into the doc attrs", async () => {
    const doc: PM = {
      type: "doc",
      attrs: { size: "a4", header: "CONFIDENTIAL", footer: "keepit.io" },
      content: [{ type: "page", content: [{ type: "paragraph" }] }],
    };
    const buffer = await toDocxBuffer(doc as never);
    const back = (await docxBufferToDoc(new Uint8Array(buffer), schema)) as {
      attrs: Record<string, unknown>;
    };
    expect(back.attrs.header).toContain("CONFIDENTIAL");
    expect(back.attrs.footer).toContain("keepit.io");
  });
});
