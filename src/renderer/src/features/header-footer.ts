import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import type { Editor } from "@tiptap/core";
import { runPagination } from "../editor/pagination";

/**
 * Page headers and footers.
 *
 * The text lives in the `header` / `footer` attrs of the doc node (see
 * schema.ts) with `{{page}}` / `{{pages}}` tokens, which keeps pagination,
 * page autosave and the .docx exporter on their existing rails. In the
 * editor each page gets two absolutely-positioned widgets inside its margin
 * band — never an editable node, so the pagination engine never measures
 * them (the engine *reserves* their height separately, see pagination.ts)
 * — and Notion mode hides the bands via CSS (`body.notion-mode`), exactly
 * like the other Word-only chrome. Print/PDF captures them for free since
 * they are in the page DOM.
 */

export const PAGE_TOKEN = "{{page}}";
export const PAGES_TOKEN = "{{pages}}";

export type BandAlign = "left" | "center" | "right";

export interface HeaderFooterValue {
  header: string | null;
  headerAlign: BandAlign;
  footer: string | null;
  footerAlign: BandAlign;
}

export function readHeaderFooter(editor: Editor): HeaderFooterValue {
  const attrs = (editor.state.doc.attrs ?? {}) as Partial<HeaderFooterValue>;
  return {
    header: typeof attrs.header === "string" && attrs.header ? attrs.header : null,
    headerAlign: attrs.headerAlign === "center" || attrs.headerAlign === "right" ? attrs.headerAlign : "left",
    footer: typeof attrs.footer === "string" && attrs.footer ? attrs.footer : null,
    footerAlign: attrs.footerAlign === "center" || attrs.footerAlign === "right" ? attrs.footerAlign : "left",
  };
}

/** Replace the page-number tokens for one concrete page. */
export function resolveTokens(text: string, pageIndex: number, pageCount: number): string {
  return text.replaceAll(PAGE_TOKEN, String(pageIndex)).replaceAll(PAGES_TOKEN, String(pageCount));
}

function makeBand(className: string, text: string, align: BandAlign): HTMLElement {
  const band = document.createElement("div");
  band.className = `${className} fw-band-${align}`;
  band.setAttribute("contenteditable", "false");
  band.textContent = text;
  return band;
}

export const HeaderFooter = Extension.create({
  name: "fwHeaderFooter",

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("fwHeaderFooter"),
        props: {
          decorations(state) {
            const attrs = (state.doc.attrs ?? {}) as Partial<HeaderFooterValue>;
            const header = typeof attrs.header === "string" && attrs.header ? attrs.header : null;
            const footer = typeof attrs.footer === "string" && attrs.footer ? attrs.footer : null;
            if (!header && !footer) return DecorationSet.empty;
            const pageCount = state.doc.childCount;
            const widgets: Decoration[] = [];
            let index = 0;
            state.doc.forEach((page, pagePos) => {
              index += 1;
              if (page.type.name !== "page") return;
              if (header) {
                const band = makeBand(
                  "fw-page-header",
                  resolveTokens(header, index, pageCount),
                  attrs.headerAlign ?? "left",
                );
                widgets.push(Decoration.widget(pagePos + 1, () => band, { side: 0, ignoreSelection: true }));
              }
              if (footer) {
                const band = makeBand(
                  "fw-page-footer",
                  resolveTokens(footer, index, pageCount),
                  attrs.footerAlign ?? "left",
                );
                widgets.push(
                  Decoration.widget(pagePos + page.nodeSize - 1, () => band, {
                    side: 1,
                    ignoreSelection: true,
                  }),
                );
              }
            });
            return DecorationSet.create(state.doc, widgets);
          },
        },
      }),
    ];
  },
});

/**
 * The Header & Footer modal (Insert → Header & Footer…). Cancel restores
 * the previous values; applying dispatches one doc-attrs transaction and
 * re-paginates (the content box shrinks/grows with the bands).
 */
export function openHeaderFooterModal(editor: Editor): void {
  if (editor.isDestroyed) return;
  document.getElementById("fw-header-footer")?.remove();
  const current = readHeaderFooter(editor);

  const overlay = document.createElement("div");
  overlay.id = "fw-header-footer";
  overlay.className = "fw-modal-overlay";
  overlay.innerHTML = `
    <div class="fw-modal" role="dialog" aria-modal="true" aria-label="Header and Footer">
      <h2>Header &amp; Footer</h2>
      <div class="row">
        <label for="fhf-header">Header</label>
        <textarea id="fhf-header" rows="2" placeholder="Leave empty for no header"></textarea>
        <select id="fhf-header-align" aria-label="Header alignment">
          <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
        </select>
        <button type="button" id="fhf-insert-page-h" title="Insert page number">1/n</button>
      </div>
      <div class="row">
        <label for="fhf-footer">Footer</label>
        <textarea id="fhf-footer" rows="2" placeholder="e.g. Page {{page}} of {{pages}}"></textarea>
        <select id="fhf-footer-align" aria-label="Footer alignment">
          <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
        </select>
        <button type="button" id="fhf-insert-page-f" title="Insert page number">1/n</button>
      </div>
      <p class="fhf-hint"><code>{{page}}</code> and <code>{{pages}}</code> become the live page numbers.</p>
      <div class="actions">
        <button id="fhf-cancel" type="button">Cancel</button>
        <button id="fhf-clear" type="button">Clear both</button>
        <button id="fhf-ok" type="button" class="primary">OK</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const headerEl = overlay.querySelector<HTMLTextAreaElement>("#fhf-header")!;
  const footerEl = overlay.querySelector<HTMLTextAreaElement>("#fhf-footer")!;
  const headerAlign = overlay.querySelector<HTMLSelectElement>("#fhf-header-align")!;
  const footerAlign = overlay.querySelector<HTMLSelectElement>("#fhf-footer-align")!;
  headerEl.value = current.header ?? "";
  footerEl.value = current.footer ?? "";
  headerAlign.value = current.headerAlign;
  footerAlign.value = current.footerAlign;

  const insertToken = (field: HTMLTextAreaElement): void => {
    const start = field.selectionStart ?? field.value.length;
    field.value = `${field.value.slice(0, start)}${PAGE_TOKEN} / ${PAGES_TOKEN}${field.value.slice(start)}`;
    field.focus();
  };
  overlay.querySelector("#fhf-insert-page-h")!.addEventListener("click", () => insertToken(headerEl));
  overlay.querySelector("#fhf-insert-page-f")!.addEventListener("click", () => insertToken(footerEl));

  const close = (): void => overlay.remove();
  const apply = (value: HeaderFooterValue): void => {
    close();
    const tr = editor.state.tr;
    tr.setDocAttribute("header", value.header);
    tr.setDocAttribute("headerAlign", value.headerAlign);
    tr.setDocAttribute("footer", value.footer);
    tr.setDocAttribute("footerAlign", value.footerAlign);
    editor.view.dispatch(tr);
    runPagination(editor);
    editor.commands.focus();
  };

  const collect = (): HeaderFooterValue => ({
    header: headerEl.value.trim() ? headerEl.value : null,
    headerAlign: headerAlign.value as BandAlign,
    footer: footerEl.value.trim() ? footerEl.value : null,
    footerAlign: footerAlign.value as BandAlign,
  });

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) close();
  });
  overlay.querySelector("#fhf-cancel")!.addEventListener("click", close);
  overlay
    .querySelector("#fhf-clear")!
    .addEventListener("click", () =>
      apply({ header: null, headerAlign: "left", footer: null, footerAlign: "left" }),
    );
  overlay.querySelector("#fhf-ok")!.addEventListener("click", () => apply(collect()));
  headerEl.focus();
}
