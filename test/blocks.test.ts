// @vitest-environment jsdom
/**
 * Phase 1 (Notion feel): the new block types, their HTML parse path, the
 * shared block catalog (slash menu / turn-into), and the full .docx
 * export → import round-trip that must keep them intact.
 */
import { beforeEach, describe, expect, it } from "vitest";
import { Editor, getSchema } from "@tiptap/core";
import type { JSONContent } from "@tiptap/core";
import { DOMParser as PMDOMParser } from "@tiptap/pm/model";
import type { TextSelection } from "@tiptap/pm/state";
import { createExtensions } from "../src/renderer/src/editor/extensions";
import { BLOCK_DEFS, filterBlockDefs, topBlock } from "../src/renderer/src/editor/block-utils";
import { toDocxBuffer } from "../src/main/docx-export";
import { docxBufferToDoc } from "../src/renderer/src/io/import-docx";

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
  };
  walk(json);
  return out;
}

function findNode(json: PM, type: string): PM | undefined {
  let found: PM | undefined;
  const walk = (n: PM): void => {
    if (found) return;
    if (n.type === type) found = n;
    (Array.isArray(n.content) ? (n.content as PM[]) : []).forEach(walk);
  };
  walk(json);
  return found;
}

describe("block schema registration", () => {
  it("registers the new blocks in the block group", () => {
    for (const name of ["taskList", "taskItem", "callout", "fwQuote", "divider", "fwCodeBlock"]) {
      expect(schema.nodes[name], `${name} missing`).toBeTruthy();
    }
    expect(schema.nodes["taskList"].spec.group).toBe("block");
    expect(schema.nodes["callout"].spec.group).toBe("block");
    expect(schema.nodes["fwQuote"].spec.group).toBe("block");
    expect(schema.nodes["divider"].spec.group).toBe("block");
    expect(schema.nodes["fwCodeBlock"].spec.group).toBe("block");
    // taskItem lives ONLY inside taskList — no block group, or the page node
    // would let bare items escape the list.
    expect(schema.nodes["taskItem"].spec.group ?? "").not.toContain("block");
  });
});

describe("HTML parse path", () => {
  it("parses a task list with checked state", () => {
    const json = parseHtml(
      '<ul data-type="taskList"><li data-type="taskItem" data-checked="true"><p>done</p></li>' +
        '<li data-type="taskItem" data-checked="false"><p>todo</p></li></ul>',
    );
    const list = findNode(json, "taskList");
    expect(list).toBeTruthy();
    const items = (list!.content as PM[]) ?? [];
    expect(items).toHaveLength(2);
    expect(items[0].attrs?.checked).toBe(true);
    expect(items[1].attrs?.checked).toBe(false);
  });

  it("does not mistake a task list for a bulleted list", () => {
    const json = parseHtml('<ul data-type="taskList"><li data-type="taskItem"><p>x</p></li></ul>');
    expect(flatten(json)).toContain("taskList");
    expect(flatten(json)).not.toContain("bulletList");
  });

  it("parses callout, quote, divider and code block", () => {
    const json = parseHtml(
      '<div data-type="callout" data-icon="⚠️"><div class="fw-callout-body"><p>careful</p></div></div>' +
        "<blockquote><p>cited</p></blockquote>" +
        "<hr>" +
        "<pre>const a = 1\nconst b = 2</pre>",
    );
    const types = flatten(json);
    expect(types).toContain("callout");
    expect(types).toContain("fwQuote");
    expect(types).toContain("divider");
    expect(types).toContain("fwCodeBlock");
    expect(JSON.stringify(json)).toContain("careful");
    expect(JSON.stringify(json)).toContain("const a = 1");
    const callout = findNode(json, "callout");
    expect(callout?.attrs?.icon).toBe("⚠️");
  });
});

describe("block catalog", () => {
  it("filters by label and keywords", () => {
    expect(filterBlockDefs(BLOCK_DEFS, "todo").map((d) => d.id)).toContain("taskList");
    expect(filterBlockDefs(BLOCK_DEFS, "head").map((d) => d.id)).toEqual(
      expect.arrayContaining(["heading1", "heading2", "heading3"]),
    );
    expect(filterBlockDefs(BLOCK_DEFS, "")).toBe(BLOCK_DEFS);
    expect(filterBlockDefs(BLOCK_DEFS, "zzz")).toHaveLength(0);
  });

  describe("apply() against a live editor", () => {
    let host: HTMLElement;
    let editor: Editor;

    const doc: JSONContent = {
      type: "doc",
      attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
      content: [{ type: "page", content: [{ type: "paragraph" }] }],
    };

    beforeEach(() => {
      document.body.innerHTML = '<div id="host"></div>';
      host = document.getElementById("host")!;
      editor = new Editor({ element: host, extensions: createExtensions(), content: doc });
    });

    it("replaces an empty paragraph with the chosen block", () => {
      editor.commands.setTextSelection(2);
      const def = BLOCK_DEFS.find((d) => d.id === "callout")!;
      expect(def.apply(editor, "")).toBe(true);
      const json = editor.getJSON();
      expect(flatten(json)).toContain("callout");
      // The placeholder paragraph was consumed, not left above the callout.
      const page = (json.content as PM[])[0];
      expect((page.content as PM[]).map((c) => c.type)).toEqual(["callout"]);
    });

    it("turn-into carries the block's text", () => {
      editor.commands.setContent({
        type: "doc",
        attrs: doc.attrs,
        content: [
          { type: "page", content: [{ type: "paragraph", content: [{ type: "text", text: "hello" }] }] },
        ],
      } as unknown as JSONContent);
      editor.commands.setTextSelection(3);
      const def = BLOCK_DEFS.find((d) => d.id === "taskList")!;
      const top = topBlock(editor);
      expect(top).toBeTruthy();
      const text = top!.node.textBetween(0, top!.node.content.size, "\n", "");
      expect(def.apply(editor, text)).toBe(true);
      expect(JSON.stringify(editor.getJSON())).toContain("hello");
      expect(flatten(editor.getJSON())).toContain("taskList");
    });

    it("divider inserts as its own block without destroying the caret block", () => {
      editor.commands.setTextSelection(2);
      const def = BLOCK_DEFS.find((d) => d.id === "divider")!;
      expect(def.apply(editor, "")).toBe(true);
      const page = (editor.getJSON().content as PM[])[0];
      const types = (page.content as PM[]).map((c) => c.type);
      expect(types).toContain("divider");
      expect(types).toContain("paragraph");
    });
  });
});

describe("task item Enter behavior", () => {
  let host: HTMLElement;
  let editor: Editor;

  const baseDoc: JSONContent = {
    type: "doc",
    attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
    content: [{ type: "page", content: [{ type: "paragraph" }] }],
  };

  beforeEach(() => {
    document.body.innerHTML = '<div id="host"></div>';
    host = document.getElementById("host")!;
    editor = new Editor({ element: host, extensions: createExtensions(), content: baseDoc });
  });

  /** Dispatch Enter exactly as ProseMirror's keymap pipeline would. */
  function pressEnter(): boolean {
    const event = new KeyboardEvent("keydown", {
      key: "Enter",
      keyCode: 13,
      which: 13,
      bubbles: true,
      cancelable: true,
    });
    return Boolean(editor.view.someProp("handleKeyDown", (f) => f(editor.view, event)));
  }

  /** Replace the page's single empty paragraph with a task list of `items`. */
  function mkList(items: (string | { text: string; checked?: boolean })[]) {
    const list = editor.state.schema.nodes["taskList"];
    const item = editor.state.schema.nodes["taskItem"];
    const para = editor.state.schema.nodes["paragraph"];
    const node = list.create(
      null,
      items.map((raw) => {
        const spec = typeof raw === "string" ? { text: raw } : raw;
        return item.create(
          { checked: Boolean(spec.checked) },
          spec.text ? para.create(null, editor.state.schema.text(spec.text)) : para.create(),
        );
      }),
    );
    const $page = editor.state.doc.resolve(1);
    editor.view.dispatch(editor.state.tr.replaceWith($page.before(1) + 1, $page.after(1) - 1, node));
  }

  function pageChildren(): PM[] {
    return (((editor.getJSON().content as PM[]) ?? [])[0] as PM).content as PM[];
  }

  /** Put the caret inside the nth (0-based) task item's paragraph. */
  function setCaretInItem(index: number): void {
    let seen = -1;
    let target = -1;
    editor.state.doc.descendants((n, p) => {
      if (n.type.name === "taskItem") {
        seen += 1;
        if (seen === index) target = p + 2;
      }
      return true;
    });
    expect(target, `no task item #${index}`).toBeGreaterThan(-1);
    editor.commands.setTextSelection(target);
  }

  it("Enter on the trailing empty item exits into a fresh paragraph below the list", () => {
    mkList(["first", "second", ""]);
    setCaretInItem(2);
    expect(pressEnter()).toBe(true);

    const children = pageChildren();
    expect(children.map((c) => c.type)).toEqual(["taskList", "paragraph"]);
    const list = children[0];
    expect((list.content as PM[]).map((i) => (i.content as PM[])[0].content)).toEqual([
      [{ type: "text", text: "first" }],
      [{ type: "text", text: "second" }],
    ]);
    // The caret must sit in the new paragraph, not back inside the list.
    const sel = editor.state.selection as TextSelection;
    expect(sel.$from.depth).toBe(2);
    expect(sel.$from.parent.type.name).toBe("paragraph");

    // Regression for the oscillation: typing lands on the new line, and the
    // previous item is left untouched.
    editor.commands.insertContent("X");
    const after = pageChildren();
    expect((after[1].content as PM[] | undefined)?.[0]?.text ?? "").toBe("X");
    expect(JSON.stringify(after[0])).toContain('"second"');
    expect(JSON.stringify(after[1])).toContain('"X"');
    expect(JSON.stringify(after[0])).not.toContain("secondX");
  });

  it("Enter on the trailing empty item exits deterministically when a block follows the list", () => {
    mkList(["first", ""]);
    // Append a trailing paragraph so the page continues after the list.
    const para = editor.state.schema.nodes["paragraph"];
    const tailPos = editor.state.doc.content.size - 1;
    editor.view.dispatch(
      editor.state.tr.insert(tailPos, para.create(null, editor.state.schema.text("tail"))),
    );
    setCaretInItem(1);
    expect(pressEnter()).toBe(true);

    const types = pageChildren().map((c) => c.type);
    expect(types).toEqual(["taskList", "paragraph", "paragraph"]);
    const sel = editor.state.selection as TextSelection;
    expect(sel.$from.depth).toBe(2);
    // The exit paragraph is the one right after the list, and it is empty.
    expect(sel.from).toBeLessThan(editor.state.doc.content.size - 6);
    expect(sel.$from.parent.textContent).toBe("");
  });

  it("Enter on a middle empty item drops it and moves to the next item", () => {
    mkList(["first", "", "second"]);
    setCaretInItem(1);
    expect(pressEnter()).toBe(true);

    const children = pageChildren();
    expect(children.map((c) => c.type)).toEqual(["taskList"]);
    const items = (children[0].content as PM[]).map((i) => (i.content as PM[])[0]);
    expect(items.map((p) => p.content)).toEqual([
      [{ type: "text", text: "first" }],
      [{ type: "text", text: "second" }],
    ]);
    // Caret is at the start of the following item's paragraph.
    const sel = editor.state.selection as TextSelection;
    expect(sel.$from.depth).toBe(4);
    expect(sel.$from.parent.textContent).toBe("second");
    expect(sel.$from.parentOffset).toBe(0);
  });

  it("Enter on the only (empty) item collapses the list into a paragraph with a usable caret", () => {
    mkList([""]);
    setCaretInItem(0);
    expect(pressEnter()).toBe(true);

    expect(pageChildren().map((c) => c.type)).toEqual(["paragraph"]);
    const sel = editor.state.selection as TextSelection;
    expect(sel.$from.depth).toBe(2);
    editor.commands.insertContent("hello");
    expect(JSON.stringify(pageChildren()[0])).toContain('"hello"');
  });

  it("Enter on a filled item splits it and lands in the new unchecked item", () => {
    mkList(["buy milk"]);
    // Caret at the end of the item's text.
    let target = -1;
    editor.state.doc.descendants((n, p) => {
      if (n.type.name === "paragraph" && n.textContent === "buy milk") target = p + n.nodeSize - 1;
      return true;
    });
    expect(target).toBeGreaterThan(-1);
    editor.commands.setTextSelection(target);
    expect(pressEnter()).toBe(true);

    const items = (pageChildren()[0].content as PM[]).map((i) => i);
    expect(items).toHaveLength(2);
    expect((items[0].content as PM[])[0].content).toEqual([{ type: "text", text: "buy milk" }]);
    expect(items[1].attrs).toEqual({ checked: false });
    const sel = editor.state.selection as TextSelection;
    expect(sel.$from.depth).toBe(4);
    expect(sel.$from.parent.textContent).toBe("");
  });

  it("Enter on an empty CHECKED item only unchecks it", () => {
    mkList([{ text: "", checked: true }]);
    setCaretInItem(0);
    expect(pressEnter()).toBe(true);

    expect(pageChildren().map((c) => c.type)).toEqual(["taskList"]);
    const item = (pageChildren()[0].content as PM[])[0];
    expect(item.attrs).toEqual({ checked: false });
  });

  it("creating via the '[ ] ' input rule leaves a caret where Enter can extend the list", () => {
    // Faithful typing: textinput handlers run BEFORE the text is applied.
    const type = (text: string) => {
      for (const ch of text) {
        const { to } = editor.state.selection;
        const claimed = editor.view.someProp("handleTextInput", (f) => f(editor.view, to, to, ch));
        if (!claimed) editor.commands.insertContent(ch);
      }
    };
    editor.commands.setTextSelection(2);
    type("[ ] ");
    expect(pageChildren().map((c) => c.type)).toEqual(["taskList"]);
    type("buy milk");
    expect(pressEnter()).toBe(true);
    const items = pageChildren()[0].content as PM[];
    expect(items).toHaveLength(2);
    expect(JSON.stringify(items[0])).toContain("buy milk");
    const sel = editor.state.selection as TextSelection;
    expect(sel.$from.depth).toBe(4);
    expect(sel.$from.parent.textContent).toBe("");
  });
});

describe("Notion blocks survive the .docx round-trip", () => {
  const richDoc: PM = {
    type: "doc",
    attrs: { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 },
    content: [
      {
        type: "page",
        content: [
          { type: "paragraph", content: [{ type: "text", text: "intro" }] },
          {
            type: "taskList",
            content: [
              {
                type: "taskItem",
                attrs: { checked: false },
                content: [{ type: "paragraph", content: [{ type: "text", text: "buy milk" }] }],
              },
              {
                type: "taskItem",
                attrs: { checked: true },
                content: [{ type: "paragraph", content: [{ type: "text", text: "ship it" }] }],
              },
            ],
          },
          {
            type: "callout",
            attrs: { icon: "⚠️" },
            content: [{ type: "paragraph", content: [{ type: "text", text: "careful here" }] }],
          },
          {
            type: "fwQuote",
            content: [{ type: "paragraph", content: [{ type: "text", text: "to be or not" }] }],
          },
          { type: "divider" },
          { type: "fwCodeBlock", content: [{ type: "text", text: "let x = 1\necho x" }] },
        ],
      },
    ],
  };

  it("exports and re-imports every new block type", async () => {
    const buf = await toDocxBuffer(richDoc as never);
    const json = (await docxBufferToDoc(buf, schema)) as unknown as PM;
    const types = flatten(json);
    expect(types).toContain("taskList");
    expect(types).toContain("callout");
    expect(types).toContain("fwQuote");
    expect(types).toContain("divider");
    expect(types).toContain("fwCodeBlock");

    // To-do check states survive.
    const list = findNode(json, "taskList")!;
    const items = (list?.content as PM[]) ?? [];
    expect(items.length).toBe(2);
    expect(items.some((i) => i.attrs?.checked === true)).toBe(true);
    expect(items.some((i) => i.attrs?.checked === false)).toBe(true);
    expect(JSON.stringify(json)).toContain("buy milk");

    // Callout keeps its icon and body.
    const callout = findNode(json, "callout")!;
    expect(callout?.attrs?.icon).toBe("⚠️");
    expect(JSON.stringify(json)).toContain("careful here");

    // Quote and code content survive.
    expect(JSON.stringify(json)).toContain("to be or not");
    expect(JSON.stringify(json)).toContain("let x = 1");

    // No stray ☐/☑ glyphs leak into plain paragraphs.
    const paras = JSON.stringify(json);
    expect(paras).not.toContain("☐ buy");
  }, 30000);

  it("degrades to readable paragraphs when the style ids are dropped", async () => {
    // Simulate a Word re-save: the importer must fall back to the visual
    // signatures (glyphs / shading) when the custom style ids are gone, so
    // content is never lost.
    const el = document.createElement("div");
    el.innerHTML =
      '<section class="docx"><p style="background-color: rgb(241, 241, 241);">💡 tip via shading</p>' +
      "<p>☑ done via glyph</p></section>";
    const { convertFwBlockSignatures } = await import("../src/renderer/src/io/import-docx");
    convertFwBlockSignatures(el);
    const html = el.innerHTML;
    expect(html).toContain('data-type="callout"');
    expect(html).toContain('data-type="taskItem"');
    expect(html).toContain('data-checked="true"');
    expect(html).toContain("tip via shading");
    expect(html).toContain("done via glyph");
    expect(html).not.toContain("☑ done");
  });
});
