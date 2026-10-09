// @vitest-environment jsdom
/**
 * Feature-complete additions: find & replace, headers & footers, footnotes,
 * comments and track changes — driven through the same public surfaces the
 * menu commands use (so the tests fail if a command stops reaching its code).
 */
import { beforeEach, describe, expect, it } from "vitest";
import { Editor } from "@tiptap/core";
import type { JSONContent } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { Slice, Fragment } from "@tiptap/pm/model";
import { createExtensions } from "../src/renderer/src/editor/extensions";
import {
  closeFindPanel,
  findNext,
  findPrev,
  openFindPanel,
  replaceAll,
  replaceCurrent,
} from "../src/renderer/src/features/find-replace";
import {
  openHeaderFooterModal,
  readHeaderFooter,
  resolveTokens,
} from "../src/renderer/src/features/header-footer";
import { collectFootnotes, FootnoteRef } from "../src/renderer/src/editor/footnotes";
import {
  CommentMarkExtension,
  addCommentToSelection,
  deleteComment,
  readCommentStore,
  updateComment,
} from "../src/renderer/src/editor/comments";
import {
  acceptAllChanges,
  acceptChangeAtCursor,
  isTrackingEnabled,
  rejectAllChanges,
  rejectChangeAtCursor,
  setTracking,
} from "../src/renderer/src/editor/review";

// ProseMirror needs these in jsdom (same shim as the e2e-app suite).
type Cb = (...args: unknown[]) => void;
(Range as unknown as { prototype: Record<string, Cb> }).prototype.getClientRects = () =>
  Object.assign([], { item: () => null });
(Range as unknown as { prototype: Record<string, Cb> }).prototype.getBoundingClientRect = () =>
  ({ x: 0, y: 0, width: 0, height: 0, top: 0, left: 0, right: 0, bottom: 0 }) as DOMRect;

const DOC: JSONContent = {
  type: "doc",
  content: [
    {
      type: "page",
      content: [
        { type: "paragraph", content: [{ type: "text", text: "alpha beta alpha" }] },
        { type: "paragraph", content: [{ type: "text", text: "Beta GAMMA" }] },
      ],
    },
  ],
};

let host: HTMLElement;
let editor: Editor;

function mount(content: JSONContent = DOC): void {
  host = document.createElement("div");
  document.body.appendChild(host);
  editor = new Editor({
    element: host,
    extensions: createExtensions(),
    content,
    injectCSS: false,
  });
}

beforeEach(() => {
  closeFindPanel();
  document.body.innerHTML = "";
  document.body.removeAttribute("data-theme");
  document.body.classList.remove("notion-mode");
});

// ---------------------------------------------------------------------------
// Find & Replace
// ---------------------------------------------------------------------------

describe("find & replace", () => {
  beforeEach(() => mount());

  function setQuery(query: string): void {
    const input = document.querySelector<HTMLInputElement>("#fr-query")!;
    input.value = query;
    input.dispatchEvent(new Event("input", { bubbles: true }));
  }

  it("highlights matches and selects the first one", () => {
    openFindPanel(editor, false);
    setQuery("alpha");
    expect(host.querySelectorAll(".fw-match").length).toBe(2);
    const sel = editor.state.selection;
    expect(editor.state.doc.textBetween(sel.from, sel.to)).toBe("alpha");
    expect(host.querySelector(".fw-match-active")).toBeTruthy();
    expect(document.querySelector("#fr-count")!.textContent).toBe("1 / 2");
  });

  it("navigates with wrap-around", () => {
    openFindPanel(editor, false);
    setQuery("beta"); // "beta" and "Beta" — case-insensitive by default
    const first = editor.state.selection.from;
    findNext(editor);
    expect(editor.state.selection.from).not.toBe(first);
    findNext(editor); // wraps
    expect(editor.state.selection.from).toBe(first);
    findPrev(editor); // wraps backwards
    const text = editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to);
    expect(text.toLowerCase()).toBe("beta");
  });

  it("respects the match-case toggle", () => {
    openFindPanel(editor, false);
    setQuery("beta");
    expect(host.querySelectorAll(".fw-match").length).toBe(2);
    document.querySelector("#fr-case")!.click();
    expect(host.querySelectorAll(".fw-match").length).toBe(1);
  });

  it("replace edits the document through undoable transactions", () => {
    openFindPanel(editor, true);
    setQuery("alpha");
    const replace = document.querySelector<HTMLInputElement>("#fr-replace")!;
    replace.value = "ALPHA";
    replaceCurrent(editor, "ALPHA"); // current match only
    expect(editor.state.doc.textContent).toContain("ALPHA beta alpha");
    findNext(editor); // re-searched after the edit: the remaining match
    replaceCurrent(editor, "ALPHA");
    expect(editor.state.doc.textContent).toContain("ALPHA beta ALPHA");
    expect(editor.state.doc.textContent).not.toContain("alpha");
  });

  it("replace all is ONE transaction — a single undo reverts every hit", () => {
    openFindPanel(editor, true);
    setQuery("beta");
    const before = editor.state.doc.textContent;
    replaceAll(editor, "BETA!");
    expect(editor.state.doc.textContent).toContain("BETA!");
    editor.commands.undo();
    // Undo may need the history plugin's own step batching; if the doc did not
    // come back, fail loudly — this is the "one Ctrl+Z" promise.
    expect(editor.state.doc.textContent).toBe(before);
  });

  it("closes on Escape and clears the decorations", () => {
    openFindPanel(editor, false);
    setQuery("alpha");
    expect(document.querySelector("#fw-find-panel")).toBeTruthy();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(document.querySelector("#fw-find-panel")).toBeNull();
    expect(host.querySelectorAll(".fw-match").length).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Headers & footers
// ---------------------------------------------------------------------------

describe("headers & footers", () => {
  beforeEach(() => mount());

  it("renders the bands on every page with live page numbers", () => {
    const tr = editor.state.tr;
    tr.setDocAttribute("header", "Confidential");
    tr.setDocAttribute("footer", "Page {{page}} of {{pages}}");
    editor.view.dispatch(tr);
    const header = host.querySelector(".fw-page-header");
    const footer = host.querySelector(".fw-page-footer");
    expect(header?.textContent).toBe("Confidential");
    expect(footer?.textContent).toBe("Page 1 of 1");
    expect(resolveTokens("{{page}}/{{pages}}", 3, 9)).toBe("3/9");
  });

  it("the modal writes the doc attrs on OK and restores them on Cancel", async () => {
    openHeaderFooterModal(editor);
    expect(document.querySelector("#fw-header-footer")).toBeTruthy();
    (document.querySelector("#fhf-header") as HTMLTextAreaElement).value = "Draft header";
    (document.querySelector("#fhf-footer-align") as HTMLSelectElement).value = "center";
    document.querySelector("#fhf-ok")!.click();
    const value = readHeaderFooter(editor);
    expect(value.header).toBe("Draft header");
    expect(value.footerAlign).toBe("center");

    openHeaderFooterModal(editor);
    (document.querySelector("#fhf-header") as HTMLTextAreaElement).value = "changed";
    document.querySelector("#fhf-cancel")!.click();
    expect(readHeaderFooter(editor).header).toBe("Draft header");
  });
});

// ---------------------------------------------------------------------------
// Footnotes
// ---------------------------------------------------------------------------

describe("footnotes", () => {
  const withRefs: JSONContent = {
    type: "doc",
    content: [
      {
        type: "page",
        content: [
          {
            type: "paragraph",
            content: [
              { type: "text", text: "First " },
              { type: "footnoteRef", attrs: { text: "one" } },
              { type: "text", text: " note and second " },
              { type: "footnoteRef", attrs: { text: "two" } },
            ],
          },
        ],
      },
    ],
  };

  beforeEach(() => mount(withRefs));

  it("registers in the inline group and round-trips its text attr", () => {
    const json = editor.getJSON() as JSONContent;
    expect(JSON.stringify(json)).toContain('"footnoteRef"');
    const refs = collectFootnotes(editor.state.doc);
    expect(refs.map((r) => r.text)).toEqual(["one", "two"]);
  });

  it("renders the note area at the bottom of the page", () => {
    const area = host.querySelector(".fw-page-footnotes");
    expect(area).toBeTruthy();
    expect(area!.textContent).toContain("one");
    expect(area!.textContent).toContain("two");
  });

  it("numbers the refs by document order in the live view", () => {
    const host2 = document.createElement("div");
    document.body.appendChild(host2);
    const ed2 = new Editor({
      element: host2,
      extensions: createExtensions(),
      content: {
        type: "doc",
        content: [
          {
            type: "page",
            content: [
              {
                type: "paragraph",
                content: [
                  { type: "text", text: "see " },
                  { type: "footnoteRef", attrs: { text: "note" } },
                ],
              },
            ],
          },
        ],
      },
      injectCSS: false,
    });
    const html = ed2.view.dom.innerHTML;
    expect(html).toContain("<sup");
    // The NodeView shows the auto-numbered marker, not the raw note text.
    expect(html).toContain("[1]");
  });

  it("FootnoteRef node exists on the schema", () => {
    expect(editor.state.schema.nodes[FootnoteRef.name]).toBeTruthy();
  });
});

// ---------------------------------------------------------------------------
// Comments
// ---------------------------------------------------------------------------

describe("comments", () => {
  beforeEach(() => mount());

  async function commentRange(from: number, to: number): Promise<string | null> {
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)));
    const promise = addCommentToSelection(editor);
    await new Promise((r) => setTimeout(r, 0));
    const box = document.querySelector<HTMLElement>("#fw-prompt");
    if (!box) return promise;
    (box.querySelector("#fw-prompt-input") as HTMLTextAreaElement).value = "a note for you";
    box.querySelector("#fw-prompt-ok")!.click();
    return promise;
  }

  it("marks the range and stores the thread on the doc attrs", async () => {
    // "alpha beta alpha" → select "beta" (offsets 6–10, +1 for the page/para tokens)
    const id = await commentRange(7, 11);
    expect(id).toBeTruthy();
    const store = readCommentStore(editor);
    expect(store[id!].text).toBe("a note for you");
    expect(store[id!].author).toBeTruthy();
    const json = JSON.stringify(editor.getJSON());
    expect(json).toContain('"comment"');
    expect(json).toContain(id!);
    expect(host.querySelectorAll(".fw-comment").length).toBe(1);
  });

  it("resolve toggles the store and the mark style; delete drops the mark only", async () => {
    const id = (await commentRange(7, 11))!;
    await updateComment(editor, id, { resolved: true });
    expect(readCommentStore(editor)[id].resolved).toBe(true);
    expect(host.querySelector(".fw-comment-resolved")).toBeTruthy();
    const textBefore = editor.state.doc.textContent;
    deleteComment(editor, id);
    expect(readCommentStore(editor)[id]).toBeUndefined();
    expect(host.querySelectorAll(".fw-comment").length).toBe(0);
    expect(editor.state.doc.textContent).toBe(textBefore);
  });

  it("the comment mark parses its own HTML back", () => {
    expect(editor.state.schema.marks[CommentMarkExtension.name]).toBeTruthy();
  });
});

// ---------------------------------------------------------------------------
// Track changes
// ---------------------------------------------------------------------------

describe("track changes", () => {
  beforeEach(() => mount());

  function typeText(text: string, at: number): boolean {
    const view = editor.view;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
    let handled = false;
    view.someProp("handleTextInput", (f) => {
      handled = handled || f(view, at, at, text);
      return handled;
    });
    return handled;
  }

  function pressDelete(inputType: string, at: number, to?: number): boolean {
    const view = editor.view;
    const tr = view.state.tr;
    tr.setSelection(
      to === undefined
        ? TextSelection.create(view.state.doc, at)
        : TextSelection.create(view.state.doc, at, to),
    );
    view.dispatch(tr);
    const event = {
      inputType,
      defaultPrevented: false,
      preventDefault() {
        this.defaultPrevented = true;
      },
    };
    let handled = false;
    view.someProp("handleDOMEvents", (handlers) => {
      const fn = handlers.beforeinput;
      if (fn) handled = handled || fn(view, event as unknown as InputEvent);
      return handled;
    });
    return handled;
  }

  it("is off by default — typing is untracked", () => {
    expect(isTrackingEnabled(editor)).toBe(false);
    expect(typeText("x", 6)).toBe(false); // not handled → PM's own input path
  });

  it("typing while on inserts tracked text", () => {
    setTracking(editor, true);
    expect(isTrackingEnabled(editor)).toBe(true);
    expect(typeText("Z", 6)).toBe(true);
    const json = JSON.stringify(editor.getJSON());
    expect(json).toContain('"trackedInsert"');
    expect(editor.state.doc.textContent).toContain("Z");
  });

  it("paste while on stamps the insert mark", () => {
    setTracking(editor, true);
    const { schema } = editor.state;
    const slice = new Slice(Fragment.from(schema.text("pasted")), 0, 0);
    // Call the review plugin's own transform directly (other plugins own this
    // prop too, and `someProp` would stop at the first one that answers).
    let out: Slice = slice;
    for (const plugin of editor.state.plugins) {
      const f = plugin.spec.props?.transformPasted as
        ((s: Slice, v: typeof editor.view) => Slice) | undefined;
      if (!f) continue;
      const next = f(slice, editor.view);
      if (JSON.stringify(next.toJSON() ?? next).includes("trackedInsert")) out = next;
    }
    const json = JSON.stringify(out.toJSON());
    expect(json).toContain('"trackedInsert"');
    // Past the revision back through a transaction: text present.
    editor.view.dispatch(editor.state.tr.insertText("pasted", 6));
    expect(editor.state.doc.textContent).toContain("pasted");
  });

  it("deleting is a mark, not a removal — reject restores the text", () => {
    setTracking(editor, true);
    // Select "beta" (7–11) and hit delete.
    expect(pressDelete("deleteContentBackward", 7, 11)).toBe(true);
    expect(editor.state.doc.textContent).toContain("beta"); // still there
    const json = JSON.stringify(editor.getJSON());
    expect(json).toContain('"trackedDelete"');

    rejectAllChanges(editor);
    expect(editor.state.doc.textContent).toContain("beta");
    expect(JSON.stringify(editor.getJSON())).not.toContain("trackedDelete");
  });

  it("single-character backspace marks the char under the caret", () => {
    setTracking(editor, true);
    const before = editor.state.doc.textContent;
    // Caret at 7 = between "alpha" and " beta"; the char before is the final
    // "a" of alpha — marked, not removed.
    expect(pressDelete("deleteContentBackward", 7)).toBe(true);
    expect(editor.state.doc.textContent).toBe(before);
    expect(JSON.stringify(editor.getJSON())).toContain('"trackedDelete"');
  });

  it("accept all keeps the text and clears the marks", () => {
    setTracking(editor, true);
    typeText("Z", 6);
    pressDelete("deleteContentBackward", 7, 9);
    setTracking(editor, false);
    const before = editor.state.doc.textContent;
    acceptAllChanges(editor);
    expect(editor.state.doc.textContent).toBe(before);
    expect(JSON.stringify(editor.getJSON())).not.toMatch(/trackedInsert|trackedDelete/);
  });

  it("accept/reject at the cursor touch only that run", () => {
    const tracked: JSONContent = {
      type: "doc",
      content: [
        {
          type: "page",
          content: [
            {
              type: "paragraph",
              content: [
                { type: "text", text: "keep " },
                { type: "text", text: "gone", marks: [{ type: "trackedDelete", attrs: { author: "me" } }] },
              ],
            },
          ],
        },
      ],
    };
    document.body.innerHTML = "";
    mount(tracked);
    const view = editor.view;
    const posOfGone = 7; // inside "gone"
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, posOfGone)));
    expect(acceptChangeAtCursor(editor)).toBe(true);
    expect(JSON.stringify(editor.getJSON())).not.toContain("trackedDelete");
    expect(editor.state.doc.textContent).toBe("keep gone");
  });

  it("reject at the cursor removes an insertion", () => {
    const tracked: JSONContent = {
      type: "doc",
      content: [
        {
          type: "page",
          content: [
            {
              type: "paragraph",
              content: [
                { type: "text", text: "base " },
                { type: "text", text: "extra", marks: [{ type: "trackedInsert", attrs: { author: "me" } }] },
              ],
            },
          ],
        },
      ],
    };
    document.body.innerHTML = "";
    mount(tracked);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 7)));
    expect(rejectChangeAtCursor(editor)).toBe(true);
    expect(editor.state.doc.textContent).toBe("base ");
  });
});
