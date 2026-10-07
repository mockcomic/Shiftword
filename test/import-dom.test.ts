// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { getSchema } from "@tiptap/core";
import { createExtensions } from "../src/renderer/src/editor/extensions";
import { domToDoc } from "../src/renderer/src/io/import-docx";
import { htmlToDoc } from "../src/renderer/src/io/import-other";

const schema = getSchema(createExtensions());

/** Minimal structural type for ProseMirror JSON (the shape domToDoc returns). */
interface PMJson {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PMJson[];
  text?: string;
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>;
}

function asDoc(json: Record<string, unknown>): PMJson {
  // domToDoc is specified to return exactly this JSON shape.
  return json as unknown as PMJson;
}

function childOf(parent: PMJson, type: string): PMJson {
  const found = (parent.content ?? []).find((c) => c.type === type);
  if (!found) throw new Error(`expected a ${type} node under a ${parent.type} node`);
  return found;
}

/**
 * Fixture mirroring docx-preview's rendered output: a `div.docx-wrapper`
 * containing one `section.docx` per Word page, spans with inline styles
 * (docx-preview writes font-family / font-size / color inline).
 */
const FIXTURE = `
<div class="docx-wrapper">
  <section class="docx">
    <h1>Title here</h1>
    <p style="text-align: center;"><span style="font-family: 'Times New Roman'; font-size: 14pt; color: #FF0000; font-weight: bold;">Hello</span><span style="font-family: Calibri; font-size: 14pt;"> world</span></p>
    <ul><li><p>Item one</p></li><li><p>Item two</p><ul><li><p>Nested</p></li></ul></li></ul>
    <ol><li><p>First</p></li></ol>
    <p>Before break</p>
  </section>
  <section class="docx">
    <p>After break</p>
  </section>
</div>`;

function parseContainer(): HTMLElement {
  const container = document.createElement("div");
  container.innerHTML = FIXTURE;
  return container;
}

describe("domToDoc", () => {
  it("produces a single page doc with the expected structure", () => {
    const doc = asDoc(domToDoc(parseContainer(), null, schema));
    expect(doc.type).toBe("doc");
    expect(doc.content).toHaveLength(1);
    const page = doc.content![0];
    expect(page.type).toBe("page");
    const types = page.content!.map((c) => c.type);
    expect(types).toContain("heading");
    expect(types).toContain("bulletList");
    expect(types).toContain("orderedList");
    expect(types).toContain("pageBreak");
    expect(types.filter((t) => t === "pageBreak")).toHaveLength(1);
  });

  it("parses heading level", () => {
    const doc = asDoc(domToDoc(parseContainer(), null, schema));
    const page = doc.content![0];
    const h1 = childOf(page, "heading");
    expect(h1.attrs?.level).toBe(1);
    expect(childOf(h1, "text").text).toBe("Title here");
  });

  it("parses inline styles into textStyle attrs (font, size, color) + bold mark", () => {
    const doc = asDoc(domToDoc(parseContainer(), null, schema));
    const page = doc.content![0];
    const para = page.content!.find(
      (c) => c.type === "paragraph" && (c.content ?? []).some((t) => t.text === "Hello"),
    );
    expect(para).toBeDefined();
    expect(para!.attrs?.textAlign).toBe("center");
    const marked = para!.content!.find((t) => t.text === "Hello")!;
    const markTypes = new Set((marked.marks ?? []).map((m) => m.type));
    expect(markTypes.has("textStyle")).toBe(true);
    expect(markTypes.has("bold")).toBe(true);
    const style = (marked.marks ?? []).find((m) => m.type === "textStyle")!.attrs ?? {};
    // Browsers/jsdom may normalize #FF0000 to rgb(255, 0, 0); the exporter accepts both.
    expect(style.fontSize).toBe("14pt");
    const family = String(style.fontFamily ?? "").replace(/^['"]+|['"]+$/g, "");
    expect(family).toBe("Times New Roman");
    const color = String(style.color ?? "").toLowerCase();
    expect(["#ff0000", "rgb(255, 0, 0)", "rgb(255,0,0)"]).toContain(color);
  });

  it("round-trips through toDocxBuffer with correct run styling", async () => {
    const { toDocxBuffer } = await import("../src/main/docx-export");
    const { unzipSync, strFromU8 } = await import("fflate");
    const json = domToDoc(parseContainer(), null, schema);
    const buf = await toDocxBuffer(json);
    const docXml = strFromU8(unzipSync(new Uint8Array(buf))["word/document.xml"]);
    expect(docXml).toContain('w:color w:val="FF0000"');
    expect(docXml).toContain('w:sz w:val="28"');
    expect(docXml).toContain("Times New Roman");
    expect(docXml).toContain("<w:b/>");
  });

  it("parses nested bullet lists", () => {
    const doc = asDoc(domToDoc(parseContainer(), null, schema));
    const page = doc.content![0];
    const ul = childOf(page, "bulletList");
    expect(ul.content).toHaveLength(2);
    const second = ul.content![1];
    expect(second.type).toBe("listItem");
    const nested = childOf(second, "bulletList");
    expect(nested.content![0].type).toBe("listItem");
    const nestedPara = childOf(nested.content![0], "paragraph");
    expect(childOf(nestedPara, "text").text).toBe("Nested");
  });

  it("places a pageBreak between rendered pages", () => {
    const doc = asDoc(domToDoc(parseContainer(), null, schema));
    const children = doc.content![0].content!;
    const breakIdx = children.findIndex((c) => c.type === "pageBreak");
    expect(breakIdx).toBeGreaterThan(0);
    const before = children[breakIdx - 1];
    const after = children[breakIdx + 1];
    expect(childOf(before, "text").text).toBe("Before break");
    expect(childOf(after, "text").text).toBe("After break");
  });

  it("converts docx-preview heading paragraphs to heading nodes", () => {
    const container = document.createElement("div");
    container.innerHTML = `<div class="docx-wrapper"><section class="docx">
      <p class="docx_heading1"><span>Heading One</span></p>
      <p class="docx_heading5"><span>Deep heading</span></p>
      <p><span>plain</span></p>
    </section></div>`;
    const doc = asDoc(domToDoc(container, null, schema));
    const children = doc.content![0].content!;
    expect(children[0].type).toBe("heading");
    expect(children[0].attrs?.level).toBe(1);
    expect(children[1].type).toBe("heading");
    expect(children[1].attrs?.level).toBe(3);
    expect(children[2].type).toBe("paragraph");
  });

  it("converts docx-preview list paragraphs into ordered/bullet lists", () => {
    const container = document.createElement("div");
    container.innerHTML = `<div class="docx-wrapper"><section class="docx">
      <p class="docx_listparagraph docx-num-2-0"><span>one</span></p>
      <p class="docx_listparagraph docx-num-2-0"><span>two</span></p>
      <p class="docx_listparagraph docx-num-1-0"><span>bullet</span></p>
    </section></div>`;
    const styles = document.createElement("div");
    styles.innerHTML = `<style>
      p.docx-num-2-0:before { content: counter(docx-num-2-0, decimal); counter-increment: docx-num-2-0; }
      p.docx-num-1-0:before { content: "●"; counter-increment: docx-num-1-0; }
    </style>`;
    const doc = asDoc(domToDoc(container, styles, schema));
    const children = doc.content![0].content!;
    const ol = children.find((c) => c.type === "orderedList");
    expect(ol).toBeDefined();
    expect(ol!.content).toHaveLength(2);
    const ul = children.find((c) => c.type === "bulletList");
    expect(ul).toBeDefined();
    expect(ul!.content).toHaveLength(1);
    const itemPara = childOf(ul!.content![0], "paragraph");
    expect(childOf(itemPara, "text").text).toBe("bullet");
  });

  it("demotes h4-h6 to heading level 3", () => {
    const container = document.createElement("div");
    container.innerHTML = `<div class="docx-wrapper"><section class="docx"><h4>Deep heading</h4></section></div>`;
    const doc = asDoc(domToDoc(container, null, schema));
    const h = childOf(doc.content![0], "heading");
    expect(h.attrs?.level).toBe(3);
    expect(childOf(h, "text").text).toBe("Deep heading");
  });

  it("imports tables with header and body cells", () => {
    const container = document.createElement("div");
    container.innerHTML = `<div class="docx-wrapper"><section class="docx">
      <p>Before</p>
      <table>
        <tr><th><p><span>H1</span></p></th><th><p><span>H2</span></p></th></tr>
        <tr><td><p><span>Cell 1</span></p></td><td><p><span>Cell 2</span></p></td></tr>
      </table>
      <p>After</p>
    </section></div>`;
    const doc = asDoc(domToDoc(container, null, schema));
    const page = doc.content![0];
    expect(page.content!.map((c) => c.type)).toEqual(["paragraph", "table", "paragraph"]);

    const table = childOf(page, "table");
    expect(table.content).toHaveLength(2);

    const headerRow = table.content![0];
    expect(headerRow.type).toBe("tableRow");
    expect(headerRow.content!.map((c) => c.type)).toEqual(["tableHeader", "tableHeader"]);
    expect(childOf(childOf(headerRow.content![0], "paragraph"), "text").text).toBe("H1");
    expect(childOf(childOf(headerRow.content![1], "paragraph"), "text").text).toBe("H2");

    const bodyRow = table.content![1];
    expect(bodyRow.content!.map((c) => c.type)).toEqual(["tableCell", "tableCell"]);
    expect(childOf(childOf(bodyRow.content![0], "paragraph"), "text").text).toBe("Cell 1");
    expect(childOf(childOf(bodyRow.content![1], "paragraph"), "text").text).toBe("Cell 2");
  });

  it("imports data-URL images with their rendered size and drops unsupported image sources", () => {
    const png1x1 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    const container = document.createElement("div");
    container.innerHTML = `<div class="docx-wrapper"><section class="docx">
      <p>Before</p>
      <img src="data:image/png;base64,${png1x1}" style="width: 200pt; height: 100pt;">
      <img src="http://example.com/unsupported.png">
      <p>After</p>
    </section></div>`;
    const doc = asDoc(domToDoc(container, null, schema));
    const page = doc.content![0];
    expect(page.content!.map((c) => c.type)).toEqual(["paragraph", "image", "paragraph"]);

    const img = childOf(page, "image");
    expect(String(img.attrs?.src)).toBe(`data:image/png;base64,${png1x1}`);
    // 200pt → 266.67px (rounded 267), 100pt → 133.33px (rounded 133).
    expect(Number(img.attrs?.width)).toBe(267);
    expect(Number(img.attrs?.height)).toBe(133);
  });
});

describe("htmlToDoc", () => {
  it("keeps tables and data-URL images, drops scripts and external images", () => {
    const png1x1 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    const html = `
      <h2>Section</h2>
      <p>Text</p>
      <table><tr><th>H</th></tr><tr><td><p>Val</p></td></tr></table>
      <img src="data:image/png;base64,${png1x1}" style="width: 100px; height: 50px;">
      <img src="http://example.com/x.png">
      <script>bad()</script>
    `;
    const doc = asDoc(htmlToDoc(html, schema));
    const page = doc.content![0];
    const types = page.content!.map((c) => c.type);
    expect(types).toContain("heading");
    expect(types).toContain("paragraph");
    expect(types).toContain("table");
    expect(types).toContain("image");

    const table = childOf(page, "table");
    expect(childOf(childOf(table.content![0], "tableHeader"), "paragraph")).toBeDefined();
    const img = childOf(page, "image");
    expect(String(img.attrs?.src)).toBe(`data:image/png;base64,${png1x1}`);
    expect(Number(img.attrs?.width)).toBe(100);
    expect(Number(img.attrs?.height)).toBe(50);
  });

  it("demotes h4-h6 to heading level 3", () => {
    const doc = asDoc(htmlToDoc("<h5>Deep</h5>", schema));
    const h = childOf(doc.content![0], "heading");
    expect(h.attrs?.level).toBe(3);
    expect(childOf(h, "text").text).toBe("Deep");
  });
});
