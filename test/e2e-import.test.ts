// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { getSchema } from "@tiptap/core";
import { createExtensions } from "../src/renderer/src/editor/extensions";
import { domToDoc } from "../src/renderer/src/io/import-docx";
import { htmlToDoc, textToDoc } from "../src/renderer/src/io/import-other";

const schema = getSchema(createExtensions());

type PM = Record<string, unknown>;
const docOf = (json: Record<string, unknown>) => json as unknown as { content: Array<{ content?: PM[] }> };
const blocks = (json: Record<string, unknown>): PM[] => docOf(json).content[0].content ?? [];

/** Returns an error string when the JSON does not produce a valid schema node. */
function validityError(json: Record<string, unknown>): string | null {
  try {
    schema.nodeFromJSON(json).check();
    return null;
  } catch (e) {
    return e instanceof Error ? e.message : String(e);
  }
}

function container(html: string): HTMLElement {
  const el = document.createElement("div");
  el.innerHTML = html;
  return el;
}

describe("empty-document imports", () => {
  it("empty .txt produces a valid single-paragraph doc", () => {
    const json = textToDoc("");
    expect(validityError(json)).toBeNull();
    expect(blocks(json)).toHaveLength(1);
  });

  it("empty .html produces a valid doc (editor must be able to load it)", () => {
    const json = htmlToDoc("", schema);
    expect(validityError(json)).toBeNull();
  });

  it("whitespace-only .html produces a valid doc", () => {
    const json = htmlToDoc("   \n  ", schema);
    expect(validityError(json)).toBeNull();
  });

  it("docx-preview output with a bodyless section produces a valid doc", () => {
    const json = domToDoc(
      container(`<div class="docx-wrapper"><section class="docx"></section></div>`),
      null,
      schema,
    );
    expect(validityError(json)).toBeNull();
  });
});

describe("textToDoc", () => {
  it("splits on all newline flavors", () => {
    const json = textToDoc("a\r\nb\rc\nd");
    expect(blocks(json)).toHaveLength(4);
    expect(JSON.stringify(json)).toContain('"a"');
  });

  it("preserves unicode and emoji", () => {
    const json = textToDoc("héllo 🌍 日本語");
    expect(JSON.stringify(json)).toContain("日本語");
  });
});

describe("htmlToDoc", () => {
  it("keeps data-url images and records rendered size", () => {
    const png =
      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
    const json = htmlToDoc(`<p><img src="${png}" style="width: 50pt; height: 50pt"></p>`, schema);
    type PM = Record<string, unknown>;
    const blocks = (json as unknown as { content: Array<{ content?: PM[] }> }).content[0].content ?? [];
    const image = blocks.find((b) => b.type === "image");
    expect(image).toBeTruthy();
    const attrs = image!.attrs as Record<string, unknown>;
    expect(String(attrs.src)).toContain("data:image/png");
    // 50pt → 66.67px → rounded to 67 px so export preserves the size.
    expect(Number(attrs.width)).toBe(67);
    expect(Number(attrs.height)).toBe(67);
  });

  it("drops external and webp images but keeps the text", () => {
    const json = htmlToDoc(
      `<p>a<img src="https://x/y.png">b<img src="data:image/webp;base64,AAAA">c</p>`,
      schema,
    );
    const s = JSON.stringify(json);
    expect(s).not.toContain("https://x");
    expect(s).not.toContain("webp");
    expect(s).toContain("abc");
  });

  it("demotes h4–h6 to h3 and drops script/style/link", () => {
    const json = htmlToDoc(
      `<h4>four</h4><h6>six</h6><script>alert(1)</script><style>p{}</style><link href="x"><p>ok</p>`,
      schema,
    );
    const s = JSON.stringify(json);
    expect(s).toContain('"heading"');
    expect(s.match(/"level":3/g)).toHaveLength(2);
    expect(s).not.toContain("alert");
  });

  it("imports nested bullet + ordered lists", () => {
    const json = htmlToDoc(
      `<ul><li>one<ul><li>nested</li></ul></li></ul><ol><li>a</li><li>b</li></ol>`,
      schema,
    );
    const s = JSON.stringify(json);
    expect(s).toContain("bulletList");
    expect(s).toContain("orderedList");
    expect(s).toContain("listItem");
  });

  // Regression: a paragraph containing ONLY an image is hoisted to block
  // level, and the emptied <p> used to survive — every imported image gained
  // a spurious blank paragraph above it (visible whitespace, and it travelled
  // into the exported .docx as well).
  it("image-only paragraph does not leave an empty companion paragraph", () => {
    const png =
      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
    const blocksOf = (html: string): string[] => {
      const json = htmlToDoc(html, schema);
      type PM = Record<string, unknown>;
      const blocks = (json as unknown as { content: Array<{ content?: PM[] }> }).content[0].content ?? [];
      return blocks.map((b) => String(b.type));
    };

    expect(blocksOf(`<p><img src="${png}"></p>`)).toEqual(["image"]);
    // A paragraph that also holds text keeps its text; the block image lands
    // beside it. The bug was the *empty* companion, never this case.
    expect(blocksOf(`<p>caption <img src="${png}"></p>`)).toEqual(["paragraph", "image"]);
  });
});

describe("domToDoc — docx-preview normalization", () => {
  it("detects ordered lists via counter() CSS tokens", () => {
    const containerEl = container(`
      <div class="docx-wrapper"><section class="docx">
        <p class="docx_listparagraph docx-num-1-0">alpha</p>
        <p class="docx_listparagraph docx-num-1-0">beta</p>
      </section></div>`);
    const styleContainer = document.createElement("div");
    styleContainer.innerHTML = `<style>p.docx-num-1-0:before { content: counter(list1) "."; }</style>`;
    const s = JSON.stringify(domToDoc(containerEl, styleContainer, schema));
    expect(s).toContain("orderedList");
    expect(s).not.toContain("bulletList");
  });

  it("treats glyph-prefixed tokens as bullets", () => {
    const containerEl = container(`
      <div class="docx-wrapper"><section class="docx">
        <p class="docx_listparagraph docx-num-2-0">bullet one</p>
      </section></div>`);
    const styleContainer = document.createElement("div");
    styleContainer.innerHTML = `<style>p.docx-num-2-0:before { content: "•"; }</style>`;
    const s = JSON.stringify(domToDoc(containerEl, styleContainer, schema));
    expect(s).toContain("bulletList");
  });

  it("splits adjacent runs at a kind change (bullet → ordered of same ref)", () => {
    const containerEl = container(`
      <div class="docx-wrapper"><section class="docx">
        <p class="docx_listparagraph docx-num-1-0">b1</p>
        <p class="docx_listparagraph docx-num-1-1">o1</p>
        <p class="docx_listparagraph docx-num-1-1">o2</p>
      </section></div>`);
    const styleContainer = document.createElement("div");
    styleContainer.innerHTML = `
      <style>p.docx-num-1-0:before { content: "•"; }</style>
      <style>p.docx-num-1-1:before { content: counter(x); }</style>`;
    const s = JSON.stringify(domToDoc(containerEl, styleContainer, schema));
    expect(s).toContain("bulletList");
    expect(s).toContain("orderedList");
  });

  it("rebuilds nesting from level classes", () => {
    const containerEl = container(`
      <div class="docx-wrapper"><section class="docx">
        <p class="docx_listparagraph docx-num-3-0">top</p>
        <p class="docx_listparagraph docx-num-3-1">kid</p>
        <p class="docx_listparagraph docx-num-3-0">top2</p>
      </section></div>`);
    const styleContainer = document.createElement("div");
    styleContainer.innerHTML = `<style>p.docx-num-3-0:before { content: "•"; } p.docx-num-3-1:before { content: "◦"; }</style>`;
    const json = domToDoc(containerEl, styleContainer, schema) as unknown as {
      content: Array<{ content: Array<{ type: string; content?: unknown[] }> }>;
    };
    const pageBlocks = json.content[0].content!;
    expect(pageBlocks[0].type).toBe("bulletList");
    const deep = JSON.stringify(pageBlocks);
    expect(deep).toContain("kid");
    // nested list must be inside the first top-level item
    expect(JSON.stringify(pageBlocks[0])).toMatch(/bulletList/);
  });

  it("maps heading classes and title to h1–h3, demoting deeper levels", () => {
    const containerEl = container(`
      <div class="docx-wrapper"><section class="docx">
        <p class="docx_heading1">a</p><p class="docx_heading5">b</p><p class="docx_title">c</p>
      </section></div>`);
    const s = JSON.stringify(domToDoc(containerEl, null, schema));
    expect(s.match(/"level":1/g)).toHaveLength(2); // heading1 + title
    expect(s).toContain('"level":3');
  });

  it("drops header/footer/footnote chrome but keeps tables", () => {
    const containerEl = container(`
      <div class="docx-wrapper"><section class="docx">
        <div class="docx-header"><p>hdr</p></div>
        <div class="docx-footer"><p>ftr</p></div>
        <table><tbody><tr><td><p>cell</p></td><td><p>two</p></td></tr></tbody></table>
      </section></div>`);
    const s = JSON.stringify(domToDoc(containerEl, null, schema));
    expect(s).not.toContain("hdr");
    expect(s).not.toContain("ftr");
    expect(s).toContain("table");
    expect(s).toContain("cell");
  });
});
