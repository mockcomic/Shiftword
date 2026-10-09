import { DOMParser as PMDOMParser } from "@tiptap/pm/model";
import type { Fragment as PMFragment, Node as PMNode, Schema } from "@tiptap/pm/model";
import { renderAsync } from "docx-preview";
import { isSafeHref } from "../editor/link";
import { PAGE_SIZES } from "../../../shared/types";
import { CALLOUT_FILL, CODE_FILL, FONT_CODE, hexToRgbCss } from "../../../shared/typography";

/** Doc attrs applied to imported documents (page size / margins in mm). */
export const DEFAULT_DOC_ATTRS = { size: "a4", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4 } as const;

/**
 * docx-preview renders headings as `<p class="docx_headingN">`, numbered and
 * bulleted items as flat `<p class="docx_listparagraph docx-num-<ref>-<lvl>">`
 * paragraphs (numbering lives in injected CSS), and each page as
 * `<section class="docx">`. These helpers convert that DOM into the plain
 * HTML the schema parses.
 */

/**
 * Collect the list tokens (e.g. "2-0") that the injected CSS marks as
 * ordered: ordered lists use `counter(...)` in the `::before` content,
 * bullets use a literal glyph.
 */
function collectOrderedListTokens(styleContainer: HTMLElement | null): Set<string> {
  const ordered = new Set<string>();
  if (!styleContainer) return ordered;
  const re = /p\.docx-num-([^-{][^{]*?):before\s*\{[^}]*counter\(/g;
  for (const styleEl of Array.from(styleContainer.querySelectorAll("style"))) {
    const css = styleEl.textContent ?? "";
    let m: RegExpExecArray | null;
    while ((m = re.exec(css)) !== null) ordered.add(m[1]);
  }
  return ordered;
}

function listToken(el: HTMLElement): string {
  return el.className.match(/\bdocx-num-([^\s]+)/)?.[1] ?? "";
}

/**
 * True when `next` directly follows `prev` among *element* siblings.
 * Whitespace-only text nodes (common in hand-written or pretty-printed
 * HTML) are ignored so list runs are not split by incidental whitespace.
 */
function isAdjacentListSibling(prev: HTMLElement, next: HTMLElement): boolean {
  if (next.parentElement !== prev.parentElement) return false;
  let n: Node | null = prev.nextSibling;
  while (n) {
    if (n === next) return true;
    if (n.nodeType !== Node.TEXT_NODE) return false;
    n = n.nextSibling;
  }
  return false;
}

function tokenLevel(el: HTMLElement): number {
  return Number(listToken(el).split("-").pop()) || 0;
}

function tokenRef(el: HTMLElement): string {
  return listToken(el).replace(/-\d+$/, "");
}

interface ListFrame {
  list: HTMLElement;
  level: number;
}

/**
 * Replace a run of consecutive `p.docx_listparagraph` siblings (same
 * numbering reference, same bullet/ordered kind) with a `<ul>` or `<ol>`;
 * nesting is rebuilt from the level encoded in the class.
 */
function replaceRunWithList(run: HTMLElement[], isOrdered: (token: string) => boolean): void {
  const first = run[0];
  const parent = first.parentElement;
  if (!parent) return;
  const rootLevel = tokenLevel(first);
  const rootList = document.createElement(isOrdered(listToken(first)) ? "ol" : "ul");
  // Capture the insertion anchor BEFORE the loop removes the run — a removed
  // node's `nextSibling` is null, and inserting against that stale anchor
  // appended the rebuilt list to the END of the section, silently moving it
  // past every later paragraph (compare convertFwBlockSignatures, which does
  // capture the sibling first).
  const anchor = run[run.length - 1].nextSibling;
  const stack: ListFrame[] = [];
  let lastLi: HTMLLIElement | null = null;
  for (const p of run) {
    const level = tokenLevel(p);
    if (level > rootLevel) {
      if (stack.length === 0 || level > stack[stack.length - 1].level) {
        const nested = document.createElement(isOrdered(listToken(p)) ? "ol" : "ul");
        const attachTo = stack.length === 0 ? rootList : lastLi;
        if (attachTo) {
          attachTo.appendChild(nested);
          stack.push({ list: nested, level });
        }
      } else {
        while (stack.length > 0 && stack[stack.length - 1].level > level) stack.pop();
      }
    }
    const li = document.createElement("li");
    while (p.firstChild) li.appendChild(p.firstChild);
    p.remove();
    (stack.length > 0 ? stack[stack.length - 1].list : rootList).appendChild(li);
    lastLi = li;
  }
  parent.insertBefore(rootList, anchor);
}

function convertListParagraphs(root: HTMLElement, isOrdered: (token: string) => boolean): void {
  const items = Array.from(root.querySelectorAll<HTMLElement>("p.docx_listparagraph"));
  if (items.length === 0) return;
  let i = 0;
  while (i < items.length) {
    const ref0 = tokenRef(items[i]);
    const ordered0 = isOrdered(listToken(items[i]));
    let j = i + 1;
    while (j < items.length) {
      const prev = items[j - 1];
      const next = items[j];
      if (!isAdjacentListSibling(prev, next)) break;
      if (tokenRef(next) !== ref0 || isOrdered(listToken(next)) !== ordered0) break;
      j++;
    }
    replaceRunWithList(items.slice(i, j), isOrdered);
    i = j;
  }
}

function replaceElement(el: HTMLElement, tag: string): void {
  const replacement = document.createElement(tag);
  while (el.firstChild) replacement.appendChild(el.firstChild);
  el.replaceWith(replacement);
}

// ---------------------------------------------------------------------------
// Notion-style block signatures (written by docx-export.ts)
// ---------------------------------------------------------------------------

/** True when the element carries `name` among its classes, case-insensitively. */
function hasClass(el: HTMLElement, name: string): boolean {
  return Array.from(el.classList).some((c) => c.toLowerCase() === name);
}

/** Background color of an inline style, normalised to "r,g,b" — "" when unset. */
function bg(el: HTMLElement): string {
  return (el.style.backgroundColor || "").replace(/\s+/g, "");
}

function isTaskPara(el: HTMLElement): boolean {
  return hasClass(el, "docx_fwtaskitem") || /^[☐☑]\s?/.test(el.textContent ?? "");
}

function isCalloutPara(el: HTMLElement): boolean {
  return hasClass(el, "docx_fwcallout") || bg(el) === hexToRgbCss(CALLOUT_FILL);
}

function isQuotePara(el: HTMLElement): boolean {
  return hasClass(el, "docx_fwquote");
}

function isToggleHeadPara(el: HTMLElement): boolean {
  return hasClass(el, "docx_fwtoggle") || /^(▸|▾)\s/.test((el.textContent ?? "").trimStart());
}

function isToggleBodyPara(el: HTMLElement): boolean {
  return hasClass(el, "docx_fwtogglebody");
}

function isTocPara(el: HTMLElement): boolean {
  return hasClass(el, "docx_fwtoc");
}

function isCodePara(el: HTMLElement): boolean {
  if (hasClass(el, "docx_fwcodeblock")) return true;
  if (bg(el) !== hexToRgbCss(CODE_FILL)) return false;
  const text = el.textContent ?? "";
  if (!text) return false;
  const spans = Array.from(el.querySelectorAll("span"));
  return spans.length > 0 && spans.every((s) => new RegExp(FONT_CODE, "i").test(s.style.fontFamily));
}

function isDividerPara(el: HTMLElement): boolean {
  if (hasClass(el, "docx_fwdivider")) return true;
  const borderBottom = el.style.borderBottom || "";
  return borderBottom !== "" && !(el.textContent ?? "").trim() && el.childElementCount === 0;
}

/** Consecutive-siblings run detection over a predicate (whitespace-tolerant). */
function findRuns(
  items: HTMLElement[],
  sameKind: (a: HTMLElement, b: HTMLElement) => boolean,
): HTMLElement[][] {
  const runs: HTMLElement[][] = [];
  let current: HTMLElement[] = [];
  for (const item of items) {
    if (
      current.length > 0 &&
      isAdjacentListSibling(current[current.length - 1], item) &&
      sameKind(current[0], item)
    ) {
      current.push(item);
    } else {
      if (current.length > 0) runs.push(current);
      current = [item];
    }
  }
  if (current.length > 0) runs.push(current);
  return runs;
}

/**
 * Move `p`'s content into a fresh clean paragraph. `stripLeading` removes a
 * leading string (the ☐/☑ or callout icon) — searching *through* element
 * children, because docx-preview gives every run its own `<span>` and the
 * glyph usually lives alone in the first one.
 */
function cleanParagraphFrom(p: HTMLElement, stripLeading?: string): HTMLElement {
  const para = document.createElement("p");
  let remaining: string | null = stripLeading ?? null;
  const walk = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      let text = node.textContent ?? "";
      if (remaining !== null && text.startsWith(remaining)) {
        text = text.slice(remaining.length);
        remaining = null;
      } else if (text.trim() !== "") {
        remaining = null; // the first real text did not match; stop hunting
      }
      if (text) para.appendChild(document.createTextNode(text));
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = node as HTMLElement;
    if (el.tagName === "BR") {
      para.appendChild(document.createElement("br"));
      return;
    }
    const clone = document.createElement(el.tagName.toLowerCase());
    for (const attr of Array.from(el.attributes)) clone.setAttribute(attr.name, attr.value);
    for (const child of Array.from(el.childNodes)) walk(child);
    if ((clone.textContent ?? "").length > 0 || el.tagName === "IMG") para.appendChild(clone);
  };
  for (const child of Array.from(p.childNodes)) walk(child);
  return para;
}

/**
 * Convert docx-preview's rendering of our exported Notion blocks back into
 * the HTML the schema parses: `ul[data-type=taskList] > li[data-type=taskItem]`,
 * `div[data-type=callout]`, `blockquote[data-type=fwQuote]`, `hr`, and
 * `pre.fw-codeblock`. Recognition is by the custom style ids we export; the
 * visual signatures (glyph prefix, shading, Consolas) are the fallback for
 * files that were re-saved by an editor that dropped the style ids.
 */
export function convertFwBlockSignatures(root: HTMLElement): void {
  // Dividers: an empty bordered paragraph (or tagged one) is a bare <hr>.
  root.querySelectorAll<HTMLElement>("p").forEach((p) => {
    if (isDividerPara(p)) {
      const hr = document.createElement("hr");
      hr.setAttribute("data-type", "divider");
      p.replaceWith(hr);
    }
  });

  const paragraphs = () => Array.from(root.querySelectorAll<HTMLElement>("p"));

  // To-do lists.
  for (const run of findRuns(paragraphs().filter(isTaskPara), (a, b) => isTaskPara(a) && isTaskPara(b))) {
    const parent = run[0].parentElement;
    // Anchor after the LAST item of the run: every p in the run gets removed,
    // so a mid-run sibling anchor would no longer be in the parent.
    const anchor = run[run.length - 1].nextSibling;
    const list = document.createElement("ul");
    list.setAttribute("data-type", "taskList");
    list.className = "fw-task-list";
    for (const p of run) {
      const text = (p.textContent ?? "").trimStart();
      const glyph = text.startsWith("☑") || text.startsWith("☐") ? text.slice(0, 1) : undefined;
      const checked = text.startsWith("☑");
      const li = document.createElement("li");
      li.setAttribute("data-type", "taskItem");
      li.setAttribute("data-checked", checked ? "true" : "false");
      li.appendChild(cleanParagraphFrom(p, glyph ? `${glyph} ` : glyph));
      p.remove();
      list.appendChild(li);
    }
    parent?.insertBefore(list, anchor);
  }

  // Callouts (icon stripped from the first paragraph into data-icon).
  for (const run of findRuns(
    paragraphs().filter(isCalloutPara),
    (a, b) => isCalloutPara(a) && isCalloutPara(b),
  )) {
    const parent = run[0].parentElement;
    const anchor = run[run.length - 1].nextSibling;
    const first = run[0].textContent ?? "";
    const knownIcon = ["💡", "ℹ️", "⚠️", "✅", "❗", "📌"].find((i) => first.trimStart().startsWith(i));
    const box = document.createElement("div");
    box.setAttribute("data-type", "callout");
    box.className = "fw-callout";
    box.setAttribute("data-icon", knownIcon ?? "💡");
    const body = document.createElement("div");
    box.appendChild(body);
    run.forEach((p, i) => {
      body.appendChild(cleanParagraphFrom(p, i === 0 && knownIcon ? `${knownIcon} ` : undefined));
      p.remove();
    });
    parent?.insertBefore(box, anchor);
  }

  // Quotes.
  for (const run of findRuns(paragraphs().filter(isQuotePara), (a, b) => isQuotePara(a) && isQuotePara(b))) {
    const parent = run[0].parentElement;
    const anchor = run[run.length - 1].nextSibling;
    const quote = document.createElement("blockquote");
    quote.setAttribute("data-type", "fwQuote");
    for (const p of run) {
      quote.appendChild(cleanParagraphFrom(p));
      p.remove();
    }
    parent?.insertBefore(quote, anchor);
  }

  // Code blocks: consecutive monospace shaded paragraphs collapse into one
  // <pre>, their text joined with newlines.
  for (const run of findRuns(paragraphs().filter(isCodePara), (a, b) => isCodePara(a) && isCodePara(b))) {
    const parent = run[0].parentElement;
    const anchor = run[run.length - 1].nextSibling;
    const pre = document.createElement("pre");
    pre.setAttribute("data-type", "fwCodeBlock");
    pre.className = "fw-codeblock";
    pre.textContent = run.map((p) => (p.textContent ?? "").replace(/\u00a0/g, " ")).join("\n");
    run.forEach((p) => p.remove());
    parent?.insertBefore(pre, anchor);
  }

  // Toggles: a FwToggle summary paragraph (▸/▾ glyph carries the collapsed
  // flag when style ids survive a third-party re-save as well), absorbing
  // the immediately following FwToggleBody paragraphs into the body.
  for (const head of paragraphs().filter(isToggleHeadPara)) {
    if (!head.parentElement) continue;
    const body: HTMLElement[] = [];
    let cursor: Node | null = head.nextSibling;
    while (cursor) {
      if (cursor.nodeType === Node.TEXT_NODE && !(cursor.textContent ?? "").trim()) {
        cursor = cursor.nextSibling;
        continue;
      }
      if (cursor.nodeType === Node.ELEMENT_NODE && isToggleBodyPara(cursor as HTMLElement)) {
        body.push(cursor as HTMLElement);
        cursor = cursor.nextSibling;
        continue;
      }
      break;
    }
    const text = (head.textContent ?? "").trimStart();
    const glyph = text.startsWith("▸") || text.startsWith("▾") ? text.slice(0, 1) : undefined;
    const collapsed = glyph === "▸" || head.getAttribute("data-collapsed") === "true";
    const box = document.createElement("div");
    box.setAttribute("data-type", "fwToggle");
    box.className = "fw-toggle";
    box.setAttribute("data-collapsed", collapsed ? "true" : "false");
    const caret = document.createElement("span");
    caret.className = "fw-toggle-caret";
    caret.setAttribute("contenteditable", "false");
    const content = document.createElement("div");
    content.className = "fw-toggle-content";
    box.append(caret, content);
    content.appendChild(cleanParagraphFrom(head, glyph ? `${glyph} ` : glyph));
    for (const p of body) {
      content.appendChild(cleanParagraphFrom(p));
      p.remove();
    }
    head.parentElement.insertBefore(box, head);
    head.remove();
  }

  // Table of contents: a run of FwToc paragraphs collapses into one marker
  // node — the editor re-renders the live heading list from the document.
  for (const run of findRuns(paragraphs().filter(isTocPara), (a, b) => isTocPara(a) && isTocPara(b))) {
    const parent = run[0].parentElement;
    const anchor = run[run.length - 1].nextSibling;
    const box = document.createElement("div");
    box.setAttribute("data-type", "fwToc");
    box.className = "fw-toc";
    run.forEach((p) => p.remove());
    parent?.insertBefore(box, anchor);
  }
}

/**
 * Headers/footers and footnote text recovered from the section rendering.
 * Page-number fields come through as their rendered first-page number (no
 * token recovery), which is the documented best-effort import path.
 */
export interface ImportedBands {
  header: string | null;
  headerAlign: "left" | "center" | "right";
  footer: string | null;
  footerAlign: "left" | "center" | "right";
}

function bandAlign(el: HTMLElement): "left" | "center" | "right" {
  const firstPara = el.querySelector("p") ?? el;
  const a = (firstPara instanceof HTMLElement ? firstPara.style.textAlign : "").toLowerCase();
  if (a === "center" || a === "right") return a;
  return "left";
}

function bandText(el: HTMLElement | null): string | null {
  if (!el) return null;
  const text = (el.textContent ?? "").replace(/\s+/g, " ").trim();
  return text ? text : null;
}

/**
 * Read the first `header`/`footer` element of the rendered document (the
 * elements themselves are dropped from the body by `normalizeDom`).
 */
export function readHeaderFooterTexts(root: HTMLElement): ImportedBands {
  const header = root.querySelector<HTMLElement>("header, .docx-header");
  const footer = root.querySelector<HTMLElement>("footer, .docx-footer");
  return {
    header: bandText(header),
    headerAlign: header ? bandAlign(header) : "left",
    footer: bandText(footer),
    footerAlign: footer ? bandAlign(footer) : "left",
  };
}

/**
 * Convert docx-preview's footnote rendering back into `footnoteRef` nodes:
 * each plain numeric `<sup>` (the reference, numbered per section) is
 * replaced by our schema signature carrying the matching note text from the
 * section's trailing note list, which is then removed from the body.
 */
export function convertFootnotes(root: HTMLElement): void {
  const sections = Array.from(root.querySelectorAll<HTMLElement>("section.docx"));
  const targets: HTMLElement[] = sections.length > 0 ? sections : [root];
  for (const section of targets) {
    const refs = Array.from(section.querySelectorAll<HTMLElement>("sup")).filter(
      (el) => /^\d+$/.test((el.textContent ?? "").trim()) && !el.hasAttribute("data-type"),
    );
    // No numeric references → no footnote chrome to harvest, and any plain
    // `<ol>` in the section is genuine content, not a note list.
    if (refs.length === 0) continue;
    const kids = Array.from(section.children) as HTMLElement[];
    // The note list is the last `<ol>` written into the section (the footer,
    // when rendered, trails it) whose entries are bare paragraphs/items.
    const noteList = [...kids]
      .reverse()
      .find(
        (el) =>
          el.tagName === "OL" &&
          Array.from(el.children).every((c) => c.tagName === "LI" || c.tagName === "P"),
      );
    const notes = noteList
      ? Array.from(noteList.children).map((li) =>
          (li.textContent ?? "")
            .replace(/^\s*\d+[.\s]*/, "")
            .replace(/\s+/g, " ")
            .trim(),
        )
      : [];
    refs.forEach((sup, i) => {
      const replacement = document.createElement("sup");
      replacement.setAttribute("data-type", "footnoteRef");
      replacement.setAttribute("data-fn", notes[i] ?? sup.textContent?.trim() ?? "");
      sup.replaceWith(replacement);
    });
    noteList?.remove();
  }
}

/** Convert a CSS length (px/pt/in/cm/mm) to CSS pixels; null if unparseable. */
function cssLenToPx(v: string | null): number | null {
  if (!v) return null;
  const m = v.trim().match(/^([\d.]+)\s*(pt|px|in|cm|mm)?$/);
  if (!m) return null;
  const n = parseFloat(m[1]);
  if (!Number.isFinite(n) || n <= 0) return null;
  switch (m[2]) {
    case "pt":
      return n * (96 / 72);
    case "in":
      return n * 96;
    case "cm":
      return n * (96 / 2.54);
    case "mm":
      return n * (96 / 25.4);
    default:
      return n; // px or unitless
  }
}

/** CSS pixels → millimetres at 96 DPI. */
function pxToMm(px: number): number {
  return (px * 25.4) / 96;
}

/**
 * Recover the section geometry (page size + margins) from docx-preview's
 * rendering: with `ignoreWidth/ignoreHeight` off it writes the section's
 * `pgSz` as an inline `width` and `pgMar` as `padding-*` on `section.docx`.
 * Returns only the values it could read — callers merge over the defaults.
 */
export function readPageAttrs(
  root: HTMLElement,
): Partial<{ size: "a4" | "letter"; mt: number; mr: number; mb: number; ml: number }> {
  const out: Partial<{ size: "a4" | "letter"; mt: number; mr: number; mb: number; ml: number }> = {};
  const section = root.querySelector<HTMLElement>("section.docx");
  if (!section) return out;
  const mm = (prop: string): number | null => {
    const px = cssLenToPx(section.style.getPropertyValue(prop));
    if (px == null || !Number.isFinite(px) || px <= 0) return null;
    const v = pxToMm(px);
    return v > 0 && v < 2000 ? Math.round(v * 10) / 10 : null;
  };
  const widthMm = (() => {
    const px = cssLenToPx(section.style.getPropertyValue("width"));
    return px == null ? null : pxToMm(px);
  })();
  if (widthMm != null && widthMm > 50) {
    // Snap to the closest supported size; Word/GDocs files are overwhelmingly
    // A4 or Letter, and small rendering deltas must not flip the choice.
    const [a4, letter] = [PAGE_SIZES.a4.w, PAGE_SIZES.letter.w];
    out.size = Math.abs(widthMm - letter) < Math.abs(widthMm - a4) ? "letter" : "a4";
  }
  const sides = { mt: "padding-top", mr: "padding-right", mb: "padding-bottom", ml: "padding-left" } as const;
  for (const [attr, prop] of Object.entries(sides) as Array<[keyof typeof sides, string]>) {
    const v = mm(prop);
    if (v != null) out[attr] = Math.min(v, 150); // absurd margins fall back below
  }
  return out;
}

/**
 * Drop links the app refuses to store or open (`javascript:`, `data:`, `file:`,
 * …) by unwrapping them: the linked text survives, the payload does not. The
 * link mark validates again when it renders, so this is defence in depth on the
 * import path rather than the only check.
 */
export function sanitizeAnchors(root: HTMLElement): void {
  root.querySelectorAll<HTMLAnchorElement>("a[href]").forEach((el) => {
    if (isSafeHref(el.getAttribute("href"))) return;
    el.removeAttribute("href");
    el.removeAttribute("target");
    replaceElement(el, "span");
  });
}

/**
 * Keep only images FreeWord can persist (data: URLs of formats the DOCX
 * writer supports) and record their rendered size as width/height attributes
 * so the editor and export keep the original dimensions.
 */
export function normalizeImages(root: HTMLElement): void {
  const supported = /^data:image\/(png|jpe?g|gif|bmp);base64,.+$/;
  root.querySelectorAll<HTMLImageElement>("img").forEach((el) => {
    const src = el.getAttribute("src") ?? "";
    if (!supported.test(src)) {
      el.remove();
      return;
    }
    const attrPx = (v: string | null): number | null => {
      const n = v == null ? NaN : Number(v);
      return Number.isFinite(n) && n > 0 ? n : null;
    };
    const w = cssLenToPx(el.style.getPropertyValue("width")) ?? attrPx(el.getAttribute("width"));
    const h = cssLenToPx(el.style.getPropertyValue("height")) ?? attrPx(el.getAttribute("height"));
    if (w) el.setAttribute("width", String(Math.round(w)));
    if (h) el.setAttribute("height", String(Math.round(h)));

    // The image node is block-level, so ProseMirror hoists it out of its
    // paragraph — and the emptied `<p>` used to survive the parse, giving every
    // imported image a spurious blank paragraph above it (visible whitespace,
    // and it travelled into the exported .docx too). Drop the wrapper here.
    const parent = el.parentElement;
    if (parent && parent.tagName === "P") {
      const holdsOnlyThisImage = Array.from(parent.childNodes).every(
        (n) =>
          n === el ||
          (n.nodeType === Node.TEXT_NODE && !(n.textContent ?? "").trim()) ||
          (n.nodeType === Node.ELEMENT_NODE && (n as Element).tagName === "BR"),
      );
      if (holdsOnlyThisImage) parent.replaceWith(el);
    }
  });
}

/**
 * Strip unsupported elements and normalize docx-preview's structure:
 * - scripts / styles / link tags / headers & footers are dropped
 * - tables and data-URL images are kept (normalized via normalizeImages)
 * - `p.docx_heading1..3` → h1–h3, `p.docx_title` → h1, h4–h6 /
 *   `p.docx_heading4..6` → h3 (schema only has heading levels 1–3)
 * - `p.docx_listparagraph` runs → ul/ol + li (nested by level)
 * - empty text spans (page-break artifacts) are removed
 */
function normalizeDom(root: HTMLElement, orderedTokens: Set<string>): void {
  // Headers, footers and the trailing per-section footnote list are rendered
  // as sibling elements of the body; the band text is scraped by
  // `readHeaderFooterTexts`/`convertFootnotes` (which removes the note list)
  // before this runs, so everything left here is dropped from the body text.
  root
    .querySelectorAll("style, script, link, header, footer, .docx-header, .docx-footer")
    .forEach((el) => el.remove());

  normalizeImages(root);
  sanitizeAnchors(root);
  convertFwBlockSignatures(root);

  root.querySelectorAll<HTMLElement>("p[class]").forEach((el) => {
    const cls = Array.from(el.classList).find((c) => /^docx_heading\d$/.test(c));
    if (!cls) return;
    const level = Number(cls.slice("docx_heading".length));
    replaceElement(el, level <= 3 ? `h${level}` : "h3");
  });
  root.querySelectorAll<HTMLElement>("p.docx_title").forEach((el) => replaceElement(el, "h1"));
  root.querySelectorAll<HTMLElement>("h4, h5, h6").forEach((el) => replaceElement(el, "h3"));

  convertListParagraphs(root, (token) => orderedTokens.has(token));

  root.querySelectorAll<HTMLElement>("span").forEach((el) => {
    if (el.textContent === "" && el.childElementCount === 0) el.remove();
  });
}

/**
 * Convert a rendered docx-preview container (or any Word-like HTML block
 * container) into a FreeWord doc JSON:
 *
 *   <div class="docx-wrapper">
 *     <section class="docx">…page 1…</section>
 *     <section class="docx">…page 2…</section>
 *   </div>
 *
 * Each `section.docx` becomes the content of the single imported page;
 * consecutive sections are separated by an explicit `pageBreak` node.
 */
/** Plain ProseMirror JSON (the shape returned by `Node.toJSON()`). */
export interface PMJson {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PMJson[];
  text?: string;
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>;
}

/** Convert parsed ProseMirror content (a Fragment) into plain JSON children. */
export function fragmentToJSON(frag: PMFragment): PMJson[] {
  const out: PMJson[] = [];
  frag.forEach((n) => out.push(n.toJSON() as PMJson));
  return out;
}

/**
 * The DOM parser wraps section content in a `page` node (the top node's
 * content expression is `page+`), so flatten any `page` wrappers to expose
 * the inner blocks directly.
 */
export function flatBlockJson(parsed: PMNode): PMJson[] {
  const out: PMJson[] = [];
  for (const child of fragmentToJSON(parsed.content)) {
    if (child.type === "page") out.push(...(child.content ?? []));
    else out.push(child);
  }
  return out;
}

export function domToDoc(
  container: HTMLElement,
  styleContainer: HTMLElement | null,
  schema: Schema,
): Record<string, unknown> {
  const root = container.querySelector<HTMLElement>(":scope > .docx-wrapper") ?? container;
  // Scrape the bands and re-hydrate footnote refs *before* the body
  // normalization strips the header/footer/footer-list chrome away.
  const bands = readHeaderFooterTexts(root);
  convertFootnotes(root);
  normalizeDom(root, collectOrderedListTokens(styleContainer));

  const sections = Array.from(root.querySelectorAll<HTMLElement>("section.docx"));
  const parser = PMDOMParser.fromSchema(schema);

  let pageContent: unknown[];
  if (sections.length === 0) {
    const parsed = parser.parse(root);
    pageContent = flatBlockJson(parsed);
  } else {
    pageContent = [];
    sections.forEach((section, i) => {
      if (i > 0) pageContent.push({ type: "pageBreak" });
      const parsed = parser.parse(section);
      pageContent.push(...flatBlockJson(parsed));
    });
  }

  return {
    type: "doc",
    attrs: {
      ...DEFAULT_DOC_ATTRS,
      ...readPageAttrs(root),
      header: bands.header,
      headerAlign: bands.headerAlign,
      footer: bands.footer,
      footerAlign: bands.footerAlign,
    },
    content: [{ type: "page", content: pageContent }],
  };
}

/**
 * Render a .docx buffer with docx-preview into a detached container and
 * convert it to a FreeWord doc JSON.
 */
export async function docxBufferToDoc(buffer: Uint8Array, schema: Schema): Promise<Record<string, unknown>> {
  const container = document.createElement("div");
  const styleContainer = document.createElement("div");
  await renderAsync(buffer, container, styleContainer, {
    inWrapper: true,
    breakPages: true,
    // The bands and the footnote list are rendered into the section so the
    // importers can scrape them (readHeaderFooterTexts / convertFootnotes);
    // normalizeDom then removes the chrome from the body.
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: false,
    // Comments import as plain text (docx-preview injects popover chrome at
    // the reference, which would pollute the body text).
    renderComments: false,
    // Keep the section geometry (width/minHeight/padding inline styles) —
    // readPageAttrs recovers the real page size and margins from them.
    ignoreWidth: false,
    ignoreHeight: false,
    // Inline document images as data: URLs so they persist in the doc JSON
    // and round-trip through DOCX export (default is blob: URLs, which die
    // with the session).
    useBase64URL: true,
  });
  return domToDoc(container, styleContainer, schema);
}
