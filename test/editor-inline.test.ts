// @vitest-environment jsdom
/**
 * The inline primitives the schema used to lack — soft line break, hyperlink,
 * monospace — checked where they are actually defined: the ProseMirror schema,
 * the HTML parse path, and the OOXML the exporter emits for them.
 */
import { describe, expect, it } from "vitest";
import { Editor, getSchema } from "@tiptap/core";
import { DOMParser as PMDOMParser } from "@tiptap/pm/model";
import { unzipSync } from "fflate";
import { createExtensions } from "../src/renderer/src/editor/extensions";
import { isSafeHref } from "../src/renderer/src/editor/link";
import { htmlToDoc } from "../src/renderer/src/io/import-other";
import { toHtml } from "../src/renderer/src/io/serialize";
import { toDocxBuffer } from "../src/main/docx-export";

const schema = getSchema(createExtensions());

type PM = Record<string, unknown>;

function parseHtml(html: string): PM {
  const el = document.createElement("div");
  el.innerHTML = html;
  return PMDOMParser.fromSchema(schema).parse(el).toJSON() as PM;
}

function flatten(json: unknown): string[] {
  const out: string[] = [];
  const walk = (n: unknown): void => {
    if (!n || typeof n !== "object") return;
    const node = n as PM;
    if (node.type) out.push(String(node.type));
    (node.content as unknown[] | undefined)?.forEach(walk);
    (node.marks as unknown[] | undefined)?.forEach(walk);
  };
  walk(json);
  return out;
}

describe("schema registration", () => {
  it("registers the three inline primitives in the inline groups", () => {
    expect(schema.nodes.hardBreak, "hardBreak node missing").toBeTruthy();
    expect(schema.marks.link, "link mark missing").toBeTruthy();
    expect(schema.marks.code, "code mark missing").toBeTruthy();
    // A mark outside the inline group cannot be attached to text at all, which
    // is why the toolbar buttons looked dead before this was set.
    expect(schema.marks.link.spec.group).toBe("inline");
    expect(schema.marks.code.spec.group).toBe("inline");
  });

  it("lets text actually carry the marks", () => {
    const text = schema.text("x", [schema.marks.code.create({})]);
    expect(text.marks.map((m) => m.type.name)).toContain("code");

    const linked = schema.text("y", [schema.marks.link.create({ href: "https://example.com" })]);
    expect(linked.marks[0].attrs?.href).toBe("https://example.com");
  });
});

describe("HTML import of inline content", () => {
  it("keeps a soft line break as a hardBreak node", () => {
    expect(flatten(parseHtml("<p>one<br>two</p>"))).toContain("hardBreak");
  });

  it("keeps a trusted link with its href", () => {
    const json = parseHtml('<p>see <a href="https://example.com/doc">docs</a></p>');
    const types = flatten(json);
    expect(types).toContain("link");
    const marks = JSON.stringify(json);
    expect(marks).toContain("https://example.com/doc");
  });

  it("keeps monospace text as the code mark", () => {
    expect(flatten(parseHtml("<p>use <code>npm test</code> now</p>"))).toContain("code");
  });

  it("drops an unsafe anchor through the html importer", () => {
    const json = htmlToDoc('<p>click <a href="javascript:alert(1)">me</a></p>', schema);
    expect(flatten(json)).not.toContain("link");
    expect(JSON.stringify(json)).not.toContain("javascript:");
    // The linked text itself survives — only the payload is gone.
    expect(JSON.stringify(json)).toContain("me");
  });

  it("isSafeHref accepts the document schemes and refuses the rest", () => {
    for (const ok of ["https://a.b", "http://a.b", "mailto:a@b.c", "tel:+15551234", "#section"]) {
      expect(isSafeHref(ok), ok).toBe(true);
    }
    for (const bad of [
      "javascript:alert(1)",
      "JaVaScRiPt:alert(1)",
      "data:text/html,x",
      "file:///c:/x",
      "",
      "   ",
      "#",
    ]) {
      expect(isSafeHref(bad), bad).toBe(false);
    }
  });
});

describe("DOCX export of inline content", () => {
  const doc = (blocks: PM[]): PM => ({
    type: "doc",
    attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
    content: [{ type: "page", content: blocks }],
  });
  const text = (t: string, marks: PM[] = []): PM => ({
    type: "text",
    text: t,
    ...(marks.length ? { marks } : {}),
  });
  const para = (...content: PM[]): PM => ({ type: "paragraph", content });

  async function xmlOf(pm: PM): Promise<{ document: string; rels: string }> {
    const buf = await toDocxBuffer(pm);
    const files = unzipSync(new Uint8Array(buf));
    const dec = new TextDecoder();
    return {
      document: dec.decode(files["word/document.xml"]),
      rels: files["word/_rels/document.xml.rels"] ? dec.decode(files["word/_rels/document.xml.rels"]) : "",
    };
  }

  it("emits a real line break for a soft break", async () => {
    const { document } = await xmlOf(doc([para(text("up"), { type: "hardBreak" }, text("down"))]));
    expect(document).toMatch(/<w:br\s*\/>/);
  });

  it("emits a hyperlink with an external relationship for a link mark", async () => {
    const { document, rels } = await xmlOf(
      doc([para(text("docs", [{ type: "link", attrs: { href: "https://example.com/x" } }]))]),
    );
    expect(document).toMatch(/<w:hyperlink/);
    expect(document).toMatch(/r:id="rId/);
    expect(rels).toContain('Target="https://example.com/x"');
    expect(rels).toContain('TargetMode="External"');
  });

  it("refuses to export a link whose href never passed validation", async () => {
    const { document, rels } = await xmlOf(
      doc([para(text("x", [{ type: "link", attrs: { href: "javascript:alert(1)" } }]))]),
    );
    expect(document).not.toMatch(/<w:hyperlink/);
    expect(rels).not.toContain("javascript:");
    // The text is still there, unlinked.
    expect(document).toContain("x");
  });

  it("renders the code mark in a monospace face", async () => {
    const { document } = await xmlOf(doc([para(text("npm ci", [{ type: "code" }]))]));
    expect(document).toMatch(/w:ascii="Consolas"|w:hAnsi="Consolas"/);
  });
});

describe("soft-break keyboard shortcut", () => {
  function bootEditor(): Editor {
    document.body.innerHTML = `<div id="editor"></div>`;
    (Range as unknown as { prototype: Record<string, never> }).prototype.getClientRects = () =>
      Object.assign([], { item: () => null }) as never;
    (Range as unknown as { prototype: Record<string, never> }).prototype.getBoundingClientRect = () =>
      ({ x: 0, y: 0, width: 0, height: 0, top: 0, left: 0, right: 0, bottom: 0 }) as never;
    return new Editor({
      element: document.querySelector<HTMLElement>("#editor")!,
      extensions: createExtensions(),
      content: {
        type: "doc",
        attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
        content: [
          { type: "page", content: [{ type: "paragraph", content: [{ type: "text", text: "first" }] }] },
        ],
      },
    });
  }

  it("Shift+Enter inserts a hard line break without ending the paragraph", () => {
    const editor = bootEditor();
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);

    const handled = editor.view.someProp("handleKeyDown", (f) =>
      f(
        editor.view,
        new KeyboardEvent("keydown", { key: "Enter", shiftKey: true, bubbles: true, cancelable: true }),
      ),
    );
    expect(handled, "Shift+Enter has no binding").toBe(true);

    const blocks = (editor.state.doc.child(0).toJSON() as PM).content as PM[];
    expect(JSON.stringify(blocks)).toContain("hardBreak");
    // Still ONE paragraph: a soft break must not start a new block.
    expect(blocks.filter((b) => b.type === "paragraph")).toHaveLength(1);
    editor.destroy();
  });

  it("plain Enter still splits the paragraph (the shortcut must not steal it)", () => {
    const editor = bootEditor();
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);

    editor.view.someProp("handleKeyDown", (f) =>
      f(editor.view, new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true })),
    );

    const blocks = (editor.state.doc.child(0).toJSON() as PM).content as PM[];
    expect(JSON.stringify(blocks)).not.toContain("hardBreak");
    editor.destroy();
  });
});

describe("HTML export", () => {
  function editorWithPages(pages: PM[][]): Editor {
    document.body.innerHTML = `<div id="editor"></div>`;
    return new Editor({
      element: document.querySelector<HTMLElement>("#editor")!,
      extensions: createExtensions(),
      content: {
        type: "doc",
        attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
        content: pages.map((blocks) => ({ type: "page", content: blocks })),
      },
    });
  }

  it("is a standalone document, not page chrome", () => {
    const editor = editorWithPages([
      [{ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Title" }] }],
      [{ type: "paragraph", content: [{ type: "text", text: "Body text" }] }],
    ]);
    const html = toHtml(editor);
    editor.destroy();

    expect(html).toMatch(/^<!doctype html>/);
    expect(html).toMatch(/charset="utf-8"/i); // so non-ASCII text survives a double-click
    expect(html).toContain("<h1");
    expect(html).toContain("Title");
    expect(html).toContain("Body text"); // content from BOTH pages, in order
    // Pagination is a view concern: no page boxes, no doc wrapper, no CSS vars.
    expect(html).not.toContain("fw-page");
    expect(html).not.toContain("fw-doc");
    expect(html).not.toContain("--page-w");
  });
});
