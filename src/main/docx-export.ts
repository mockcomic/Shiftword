import {
  AlignmentType,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  ExternalHyperlink,
  Footer,
  FootnoteReferenceRun,
  Header,
  HeadingLevel,
  ImageRun,
  InsertedTextRun,
  LevelFormat,
  LineRuleType,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlignTable,
  WidthType,
} from "docx";
// Only used through `typeof` to name the highlight colour union.
import type { HighlightColor } from "docx";
import type { PageSize } from "../shared/types";
import { PAGE_SIZES } from "../shared/types";
import {
  BODY_SIZE_HALF_POINTS,
  CALLOUT_FILL,
  CODE_FILL,
  CODE_SIZE_HALF_POINTS,
  DIVIDER_COLOR,
  FONT_BODY,
  FONT_CODE,
  FONT_EMOJI,
  FONT_SYMBOL,
  HEADING_COLOR,
  HEADING_SIZE_HALF_POINTS,
  HIGHLIGHT_GREEN,
  HIGHLIGHT_YELLOW,
  HYPERLINK_COLOR,
  LINE_SPACING_TWIPS,
  PARA_AFTER_TWIPS,
  ptToTwips,
} from "../shared/typography";

type AlignmentValue = (typeof AlignmentType)[keyof typeof AlignmentType];
type HeadingValue = (typeof HeadingLevel)[keyof typeof HeadingLevel];
type HighlightValue = (typeof HighlightColor)[keyof typeof HighlightColor];

// ---------------------------------------------------------------------------
// Lightweight ProseMirror JSON types (plain objects crossing the IPC boundary)
// ---------------------------------------------------------------------------
interface Mark {
  type: string;
  attrs?: Record<string, unknown>;
}
interface PMNode {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PMNode[];
  marks?: Mark[];
  text?: string;
}
interface PMDoc {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PMNode[];
}

function childrenOf(node: PMNode | null | undefined): PMNode[] {
  return Array.isArray(node?.content) ? (node.content as PMNode[]) : [];
}

function num(v: unknown, fallback: number): number {
  const n = typeof v === "number" ? v : parseFloat(String(v ?? ""));
  return Number.isFinite(n) ? n : fallback;
}

/** As `num`, but always a whole number ≥ 1 — what `columnSpan`/`rowSpan` require. */
function positiveInt(v: unknown, fallback: number): number {
  const n = Math.floor(num(v, NaN));
  return Number.isFinite(n) && n >= 1 ? n : fallback;
}

/** Millimetres → twips (1 mm = 1440/25.4 ≈ 56.6929 twips).
 *  Rounding yields the canonical Word values (A4 = 11906 × 16838, Letter = 12240 × 15840). */
function mmToTwip(mm: number): number {
  return Math.round((mm * 1440) / 25.4);
}

/** CSS pixels → twips (1 px = 15 twips at 96 DPI) — Tiptap stores `colwidth` in px. */
function pxToTwip(px: number): number {
  return Math.round(px * 15);
}

/** Millimetres → CSS pixels at 96 DPI. */
function mmToPx(mm: number): number {
  return (mm * 96) / 25.4;
}

type ParagraphChild =
  | TextRun
  | ExternalHyperlink
  | FootnoteReferenceRun
  | InsertedTextRun
  | DeletedTextRun
  | CommentRangeStart
  | CommentRangeEnd
  | CommentReference;
type DocChild = Paragraph | Table;

/**
 * One-pass context computed from the whole document before any run is
 * emitted, so ordinals are stable wherever the node appears:
 *
 * - `footnotes` maps each `footnoteRef` JSON node (object identity — the
 *   same graph the renderer serialized) to its Word ordinal, which doubles
 *   as the `word/footnotes.xml` id.
 * - `comments` maps a comment id (the string on the `comment` mark) to a
 *   numeric OOXML comment id; `commentDefs` carries the thread text.
 * - `revision` numbers the `w:ins`/`w:del` ids of tracked changes.
 */
interface ExportContext {
  footnotes: Map<PMNode, number>;
  footnoteTexts: Map<number, string>;
  comments: Map<string, number>;
  commentDefs: Array<{
    id: number;
    author: string;
    initials: string;
    date: Date;
    children: Paragraph[];
  }>;
  revision: { n: number };
  /**
   * Every heading in document order (level + plain text) — the table-of-
   * contents node carries no content in the JSON, so its export renders from
   * this list, exactly like the editor's live NodeView.
   */
  headings: Array<{ level: number; text: string }>;
}

/**
 * The payload arrives over IPC — structured clone PRESERVES reference
 * cycles, so a self-referential document JSON (buggy producer or hostile
 * renderer) would otherwise recurse until the stack blows and the whole
 * save fails with an opaque RangeError.
 */
const MAX_WALK_DEPTH = 256;

function walkPm(node: PMNode, visit: (n: PMNode) => void, depth = 0): void {
  if (depth > MAX_WALK_DEPTH) throw new RangeError("document JSON nests too deeply (reference cycle?)");
  visit(node);
  childrenOf(node).forEach((c) => walkPm(c, visit, depth + 1));
}

/** Concatenated plain text of a node subtree (bounded like `walkPm`). */
function flattenText(node: PMNode, depth = 0): string {
  if (depth > MAX_WALK_DEPTH) return "";
  let out = typeof node.text === "string" ? node.text : "";
  for (const child of childrenOf(node)) out += flattenText(child, depth + 1);
  return out;
}

interface RawCommentRecord {
  author?: string;
  text?: string;
  createdAt?: number;
}

function buildExportContext(doc: PMDoc): ExportContext {
  const ctx: ExportContext = {
    footnotes: new Map(),
    footnoteTexts: new Map(),
    comments: new Map(),
    commentDefs: [],
    revision: { n: 0 },
    headings: [],
  };
  const store = (doc.attrs?.comments ?? {}) as Record<string, RawCommentRecord>;
  const top: PMNode = { type: "doc", content: childrenOf(doc as unknown as PMNode) };
  walkPm(top, (node) => {
    if (node.type === "footnoteRef") {
      const ordinal = ctx.footnotes.size + 1;
      ctx.footnotes.set(node, ordinal);
      ctx.footnoteTexts.set(ordinal, String(node.attrs?.text ?? ""));
    }
    if (node.type === "heading") {
      // Plain text of the whole heading subtree, in document order — the
      // TOC block's export entries (see the `fwToc` case below).
      const level = Math.min(3, Math.max(1, Math.round(num(node.attrs?.level, 1))));
      const text = flattenText(node).trim();
      ctx.headings.push({ level, text });
    }
    for (const mark of node.marks ?? []) {
      if (mark.type !== "comment") continue;
      const id = String(mark.attrs?.id ?? "");
      if (!id || ctx.comments.has(id)) continue;
      const numId = ctx.comments.size + 1;
      ctx.comments.set(id, numId);
      const rec = store[id] ?? {};
      const author = typeof rec.author === "string" && rec.author ? rec.author : "Unknown";
      ctx.commentDefs.push({
        id: numId,
        author,
        initials: author.slice(0, 2).toUpperCase(),
        date: new Date(typeof rec.createdAt === "number" ? rec.createdAt : Date.now()),
        children: [new Paragraph({ children: [new TextRun(String(rec.text ?? ""))] })],
      });
    }
  });
  return ctx;
}

const CELL_BORDER = { style: "single" as const, size: 4, color: "000000" };

/** Link schemes that become a real OOXML hyperlink; `#anchor` stays plain text. */
const EXTERNAL_HREF = /^(https?:|mailto:|tel:)/i;

/**
 * Ordered-list numbering references, one per list.
 *
 * OOXML numbers by `numId`: every paragraph that shares a reference shares a
 * counter, so one global reference made a second ordered list continue the
 * first (1, 2 … then 3, 4). Each top-level ordered list therefore gets its own
 * reference, and `numbering.config` is built from what the document actually
 * used.
 */
class OrderedNumbering {
  private readonly references = new Set<string>();
  private counter = 0;

  /** Claim a fresh reference for a top-level ordered list. */
  next(): string {
    this.counter += 1;
    const reference = `fw-ordered-${this.counter}`;
    this.references.add(reference);
    return reference;
  }

  toConfig() {
    return [...this.references].map((reference) => ({ reference, levels: orderedLevels() }));
  }
}

/** Decimal levels 0–8 (Word's full depth) for one ordered-list reference. */
function orderedLevels() {
  return [0, 1, 2, 3, 4, 5, 6, 7, 8].map((level) => ({
    level,
    format: LevelFormat.DECIMAL,
    text: `%${level + 1}.`,
    alignment: AlignmentType.START,
    style: { paragraph: { indent: { left: 720 + level * 360, hanging: 360 } } },
  }));
}

/**
 * Natural pixel size of a base64-encoded raster image (PNG/GIF/JPEG/BMP).
 * Returns null when the format is unknown or the header is malformed.
 *
 * Only the header is decoded (64 KB covers any realistic JPEG SOF even with
 * large EXIF/APP segments): decoding the whole base64 payload per image node
 * just to read 8–26 bytes was a full wasted decode of every photo, on top
 * of the docx library's own decode.
 */
const IMAGE_HEADER_BASE64_BYTES = 64 * 1024;
function imageSizeFromBase64(b64: string): { w: number; h: number } | null {
  const header = b64.length > IMAGE_HEADER_BASE64_BYTES ? b64.slice(0, IMAGE_HEADER_BASE64_BYTES) : b64;
  const buf = Buffer.from(header, "base64");
  // PNG: IHDR width/height at offsets 16/20 (big-endian u32).
  if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) {
    return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  }
  // GIF: logical screen width/height at offsets 6/8 (little-endian u16).
  if (buf.length > 10 && buf.toString("ascii", 0, 4) === "GIF8") {
    return { w: buf.readUInt16LE(6), h: buf.readUInt16LE(8) };
  }
  // JPEG: scan for a SOF marker (C0–CF, excluding C4/C8/CC).
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
    let off = 2;
    while (off + 9 <= buf.length) {
      if (buf[off] !== 0xff) {
        off++;
        continue;
      }
      const marker = buf[off + 1];
      if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7)) {
        off += 2;
        continue;
      }
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { h: buf.readUInt16BE(off + 5), w: buf.readUInt16BE(off + 7) };
      }
      off += 2 + buf.readUInt16BE(off + 2);
    }
  }
  // BMP: header width/height (int32 LE) at offsets 18/22.
  if (buf.length > 26 && buf[0] === 0x42 && buf[1] === 0x4d) {
    return { w: Math.abs(buf.readInt32LE(18)), h: Math.abs(buf.readInt32LE(22)) };
  }
  return null;
}

/** Decode a `data:image/...;base64,` URL into docx `ImageRun` fields. */
function dataUrlToImage(src: string): { type: "png" | "jpg" | "gif" | "bmp"; data: string } | null {
  // indexOf + slice instead of a `(.+)` capture: the regex engine copies the
  // entire (megabyte-scale) base64 payload into the capture group.
  const header = /^data:image\/(png|jpe?g|gif|bmp);base64,/i.exec(src);
  if (!header) return null;
  const type = header[1].toLowerCase() === "jpeg" ? "jpg" : header[1].toLowerCase();
  return { type: type as "png" | "jpg" | "gif" | "bmp", data: src.slice(header[0].length) };
}

/** One paragraph containing the image, clamped to the page content width. */
function imageToParagraph(node: PMNode, contentWidthPx: number): Paragraph {
  const attrs = node.attrs ?? {};
  const src = typeof attrs.src === "string" ? attrs.src : "";
  const img = dataUrlToImage(src);
  // Unsupported sources (blob:, webp, …) render as an empty paragraph.
  if (!img) return new Paragraph({});

  let w = num(attrs.width, 0);
  let h = num(attrs.height, 0);
  if (!(w > 0 && h > 0)) {
    const natural = imageSizeFromBase64(img.data);
    if (natural) {
      w = natural.w;
      h = natural.h;
    }
  }
  if (!(w > 0 && h > 0)) {
    w = 480;
    h = 360;
  }
  if (contentWidthPx > 0 && w > contentWidthPx) {
    const scale = contentWidthPx / w;
    w *= scale;
    h *= scale;
  }
  return new Paragraph({
    children: [
      new ImageRun({
        type: img.type,
        data: img.data,
        transformation: { width: Math.round(w), height: Math.round(h) },
      }),
    ],
  });
}

function ptToHalfPoints(v: unknown): number | undefined {
  const s = String(v ?? "");
  const m = s.match(/^([\d.]+)\s*pt$/);
  if (m) return Math.round(parseFloat(m[1]) * 2);
  const n = parseFloat(s);
  return Number.isFinite(n) ? Math.round(n * 2) : undefined;
}

/** Accepts "#RRGGBB", "RRGGBB", "rgb(r, g, b)" or "rgba(r, g, b, a)". */
function toHex(v: unknown): string | undefined {
  if (typeof v !== "string") return undefined;
  const s = v.trim();
  const hex = s.replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return hex.toUpperCase();
  const m = s.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (m) {
    return [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0").toUpperCase()).join("");
  }
  return undefined;
}

interface RunOptions {
  bold?: boolean;
  italics?: boolean;
  strike?: boolean;
  underline?: { type: "single" };
  color?: string;
  size?: number;
  font?: { name: string };
  highlight?: HighlightValue;
}

const MONO_FONT = { name: FONT_CODE };

/**
 * Inline children of a paragraph → docx runs.
 *
 * A `text` child becomes one run carrying every mark it holds; a `hardBreak`
 * becomes a `<w:br/>`; a `footnoteRef` becomes a real `FootnoteReferenceRun`;
 * a `comment` mark wraps the run in `CommentRangeStart/End` + `CommentReference`;
 * `trackedInsert`/`trackedDelete` become native `w:ins`/`w:del` runs so Word
 * shows them in its own Review pane; and a `link` mark wraps the run in an
 * `ExternalHyperlink` so Word shows a real, clickable link (the mark itself was
 * already validated to an http/https/mailto/tel URL when it was stored).
 */
function buildRuns(children: PMNode[], ctx: ExportContext): ParagraphChild[] {
  const runs: ParagraphChild[] = [];
  for (const child of children) {
    if (child.type === "hardBreak") {
      // A soft line break: Word wants one `<w:br/>` run, not a new paragraph.
      runs.push(new TextRun({ break: 1 }));
      continue;
    }
    if (child.type === "footnoteRef") {
      const ordinal = ctx.footnotes.get(child);
      if (ordinal !== undefined) runs.push(new FootnoteReferenceRun(ordinal));
      continue;
    }
    if (child.type !== "text") continue;
    const text = String(child.text ?? "");
    if (text.length === 0) continue;
    const opts: RunOptions = {};
    let href: string | null = null;
    let commentId: number | undefined;
    let inserted = false;
    let deleted = false;
    let trackAuthor = "Unknown";
    let trackDate = new Date().toISOString();
    for (const mark of child.marks ?? []) {
      const a = mark.attrs ?? {};
      switch (mark.type) {
        case "bold":
          opts.bold = true;
          break;
        case "italic":
          opts.italics = true;
          break;
        case "strike":
          opts.strike = true;
          break;
        case "underline":
          opts.underline = { type: "single" };
          break;
        case "code":
          // Word has no inline "code" character style here; monospace + the
          // grey shading is the closest stable rendering.
          opts.font = MONO_FONT;
          break;
        case "comment":
          commentId = ctx.comments.get(String(a.id ?? ""));
          break;
        case "trackedInsert":
          inserted = true;
          trackAuthor = String(a.author ?? "") || trackAuthor;
          if (typeof a.date === "string" && a.date) trackDate = a.date;
          break;
        case "trackedDelete":
          deleted = true;
          trackAuthor = String(a.author ?? "") || trackAuthor;
          if (typeof a.date === "string" && a.date) trackDate = a.date;
          break;
        case "link": {
          const candidate = typeof a.href === "string" ? a.href.trim() : "";
          href = EXTERNAL_HREF.test(candidate) ? candidate : null;
          break;
        }
        case "textStyle": {
          const f = String(a.fontFamily ?? "")
            .trim()
            .replace(/^['"]+|['"]+$/g, "");
          if (f) opts.font = { name: f };
          const s = ptToHalfPoints(a.fontSize);
          if (s) opts.size = s;
          const c = toHex(a.color);
          if (c) opts.color = c;
          break;
        }
        case "highlight": {
          // docx-preview hands us Word's ST_HighlightColor enum names verbatim
          // (`w:highlight val="yellow"` → span background "yellow"), so accept
          // those too and map them back to the palette the UI stores.
          const NAMED_HIGHLIGHTS: Record<string, string> = {
            yellow: HIGHLIGHT_YELLOW,
            green: HIGHLIGHT_GREEN,
          };
          const raw = typeof a.color === "string" ? a.color.trim().toLowerCase() : "";
          const c = NAMED_HIGHLIGHTS[raw] ?? toHex(a.color);
          if (c === HIGHLIGHT_YELLOW) opts.highlight = "yellow";
          else if (c === HIGHLIGHT_GREEN) opts.highlight = "green";
          // Other highlight colors are dropped in v1.
          break;
        }
        default:
          break;
      }
    }
    // Font/size intentionally NOT stamped here: body Calibri 11 pt lives in
    // the document defaults (see `documentStyles`), and heading sizes in the
    // Heading 1–3 styles, so Word/Google Docs render the paragraph styles
    // instead of having every run's direct formatting override them. Runs
    // only carry direct font/size when the user set it via a textStyle mark.

    ctx.revision.n += 1;
    let core: ParagraphChild[];
    if (href) {
      // The link's own colour/underline win over whatever the run carried;
      // Word users expect the blue underlined treatment.
      core = [
        new ExternalHyperlink({
          link: href,
          children: [new TextRun({ text, ...opts, color: HYPERLINK_COLOR, underline: { type: "single" } })],
        }),
      ];
    } else if (deleted) {
      // Native Word revision: the text lives inside `w:del` and Word's Review
      // pane can accept or reject it without our marks.
      core = [
        new DeletedTextRun({ text, id: ctx.revision.n, author: trackAuthor, date: trackDate, ...opts }),
      ];
    } else if (inserted) {
      core = [
        new InsertedTextRun({ text, id: ctx.revision.n, author: trackAuthor, date: trackDate, ...opts }),
      ];
    } else {
      core = [new TextRun({ text, ...opts })];
    }

    if (commentId !== undefined) {
      runs.push(new CommentRangeStart(commentId));
      runs.push(...core);
      runs.push(new CommentRangeEnd(commentId), new CommentReference(commentId));
    } else {
      runs.push(...core);
    }
  }
  return runs;
}

const ALIGN_MAP: Record<string, AlignmentValue> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

interface ListContext {
  numbering?: { reference: string; level: number };
  bullet?: { level: number };
}

/** Extra paragraph shaping for the Notion-style blocks (see blocks.ts). */
interface ParagraphExtras {
  /** OOXML style id we also recognise on import (docx-preview → `docx_<id>` class). */
  styleName?: string;
  shadingFill?: string;
  indentLeftTwips?: number;
  borderBottom?: boolean;
  /**
   * `w:spacing w:after` in twips. Mirrors the CSS `margin: 0 0 6pt` on
   * `.fw-page p` (collapsed to 0 for the last child, like the CSS
   * `p:last-child` rule); callers pass it explicitly at every container
   * boundary so Word/Docs show the same gaps as the editor.
   */
  spacingAfterTwips?: number;
}

function makeParagraph(
  node: PMNode,
  listCtx: ListContext | null,
  extra: ParagraphExtras | undefined,
  ctx: ExportContext,
): Paragraph {
  const attrs = node.attrs ?? {};
  const alignment = typeof attrs.textAlign === "string" ? ALIGN_MAP[attrs.textAlign] : undefined;
  const lineSpacing = num(attrs.lineSpacing, 1);
  const spacing: {
    line?: number;
    lineRule?: (typeof LineRuleType)[keyof typeof LineRuleType];
    after?: number;
  } = {};
  if (lineSpacing !== 1) {
    spacing.line = Math.round(240 * lineSpacing);
    spacing.lineRule = LineRuleType.AUTO;
  }
  if (extra?.spacingAfterTwips !== undefined) spacing.after = extra.spacingAfterTwips;

  let heading: HeadingValue | undefined;
  if (node.type === "heading") {
    const lvl = Math.min(3, Math.max(1, Math.round(num(attrs.level, 1))));
    heading = (HeadingLevel as Record<string, HeadingValue>)[`HEADING_${lvl}`];
  }

  const paragraph: Paragraph = new Paragraph({
    children: buildRuns(childrenOf(node), ctx),
    ...(alignment !== undefined ? { alignment } : {}),
    ...(Object.keys(spacing).length > 0 ? { spacing } : {}),
    ...(heading !== undefined ? { heading } : {}),
    ...(listCtx?.numbering !== undefined ? { numbering: listCtx.numbering } : {}),
    ...(listCtx?.bullet !== undefined ? { bullet: listCtx.bullet } : {}),
    ...(extra?.styleName !== undefined ? { style: extra.styleName } : {}),
    ...(extra?.shadingFill !== undefined ? { shading: { fill: extra.shadingFill } } : {}),
    ...(extra?.indentLeftTwips !== undefined ? { indent: { left: extra.indentLeftTwips } } : {}),
    ...(extra?.borderBottom
      ? { border: { bottom: { style: "single" as const, size: 6, color: DIVIDER_COLOR } } }
      : {}),
  });
  return paragraph;
}

/**
 * `w:spacing w:after` for the *last* paragraph of a container, mirroring the
 * CSS `p:last-child { margin-bottom: 0 }` collapse: the gap only survives when
 * the container itself has a bottom margin (it is not the last block of its
 * parent) — otherwise the next block would inherit a phantom 6 pt gap.
 */
function lastChildAfterTwips(containerIsLastInParent: boolean): number {
  return containerIsLastInParent ? 0 : PARA_AFTER_TWIPS;
}

/**
 * `w:spacing w:after` for a non-last paragraph — always the full 6 pt gap.
 */
const MIDDLE_AFTER_TWIPS = PARA_AFTER_TWIPS;

interface ListSiblings {
  /** True when this item is the last `<li>` of its (nested) list. */
  isLastItem: boolean;
  /** After-spacing the whole list contributes *below* its last paragraph (0 for nested `li > ul` / last page block). */
  listBottomTwips: number;
}

function listItemToParagraphs(
  item: PMNode,
  kind: "bullet" | "ordered",
  level: number,
  orderedRef: string | null,
  siblings: ListSiblings,
  ctx: ExportContext,
): Paragraph[] {
  const ctxList: ListContext =
    kind === "bullet"
      ? { bullet: { level } }
      : orderedRef
        ? { numbering: { reference: orderedRef, level } }
        : {};
  const out: Paragraph[] = [];
  const kids = childrenOf(item);
  kids.forEach((child, k) => {
    const isLastChild = k === kids.length - 1;
    if (child.type === "paragraph" || child.type === "heading") {
      const after = isLastChild ? (siblings.isLastItem ? siblings.listBottomTwips : 0) : MIDDLE_AFTER_TWIPS;
      out.push(makeParagraph(child, ctxList, { spacingAfterTwips: after }, ctx));
    } else if (child.type === "bulletList") {
      const subs = childrenOf(child).filter((c) => c.type === "listItem");
      subs.forEach((sub, si) => {
        out.push(
          ...listItemToParagraphs(
            sub,
            "bullet",
            level + 1,
            null,
            {
              isLastItem: si === subs.length - 1 && isLastChild && siblings.isLastItem,
              // `li > ul { margin: 0 }` — a nested list only shows the outer
              // list's bottom margin when it is itself the very last element.
              listBottomTwips: isLastChild ? siblings.listBottomTwips : MIDDLE_AFTER_TWIPS,
            },
            ctx,
          ),
        );
      });
    } else if (child.type === "orderedList") {
      const subs = childrenOf(child).filter((c) => c.type === "listItem");
      subs.forEach((sub, si) => {
        // A nested ordered list continues its own parent's sequence (same
        // reference, deeper level) — that is how Word nests a numbered list.
        out.push(
          ...listItemToParagraphs(
            sub,
            "ordered",
            level + 1,
            orderedRef,
            {
              isLastItem: si === subs.length - 1 && isLastChild && siblings.isLastItem,
              listBottomTwips: isLastChild ? siblings.listBottomTwips : MIDDLE_AFTER_TWIPS,
            },
            ctx,
          ),
        );
      });
    }
  });
  return out;
}

/** All items of one (nested) list with correct last-item flags. */
function listItemsToParagraphs(
  list: PMNode,
  kind: "bullet" | "ordered",
  level: number,
  orderedRef: string | null,
  listIsLastInParent: boolean,
  ctx: ExportContext,
): Paragraph[] {
  const items = childrenOf(list).filter((c) => c.type === (kind === "bullet" ? "listItem" : "listItem"));
  const out: Paragraph[] = [];
  items.forEach((item, i) => {
    out.push(
      ...listItemToParagraphs(
        item,
        kind,
        level,
        orderedRef,
        {
          isLastItem: i === items.length - 1,
          listBottomTwips: listIsLastInParent ? 0 : MIDDLE_AFTER_TWIPS,
        },
        ctx,
      ),
    );
  });
  return out;
}

/**
 * A task item exports as a paragraph prefixed with ☐/☑ (Word has no native
 * checkbox list) tagged with the `FwTaskItem` style so the importer can
 * re-hydrate it into a real to-do. Task-list items are tight in the editor
 * (`li.fw-task-item p { margin: 0 }`), so only the list's 6 pt bottom gap —
 * passed in as `afterTwips` — is mirrored, on the final item.
 */
function taskItemToParagraphs(item: PMNode, afterTwips: number, ctx: ExportContext): Paragraph[] {
  const inner = childrenOf(item).find((c) => c.type === "paragraph" || c.type === "heading");
  const glyph = item.attrs?.checked === true ? "☑ " : "☐ ";
  const runs: ParagraphChild[] = [
    new TextRun({ text: glyph, font: { name: FONT_SYMBOL }, size: BODY_SIZE_HALF_POINTS }),
  ];
  if (inner) runs.push(...buildRuns(childrenOf(inner), ctx));
  return [new Paragraph({ style: "FwTaskItem", children: runs, spacing: { after: afterTwips } })];
}

/** Task list: tight items; the list gap rides on the last item's paragraph. */
function taskListToParagraphs(node: PMNode, isLastInParent: boolean, ctx: ExportContext): Paragraph[] {
  const items = childrenOf(node).filter((c) => c.type === "taskItem");
  return items.flatMap((c, i) =>
    taskItemToParagraphs(c, i === items.length - 1 && !isLastInParent ? MIDDLE_AFTER_TWIPS : 0, ctx),
  );
}

/** A callout exports as its body paragraphs, shaded, with the icon on the first line. */
function calloutToParagraphs(node: PMNode, isLastInParent: boolean, ctx: ExportContext): Paragraph[] {
  const blocks = childrenOf(node);
  const icon = typeof node.attrs?.icon === "string" ? node.attrs.icon : "💡";
  const out: Paragraph[] = [];
  blocks.forEach((block, i) => {
    const isLastBlock = i === blocks.length - 1;
    const after = isLastBlock ? lastChildAfterTwips(isLastInParent) : MIDDLE_AFTER_TWIPS;
    if (block.type === "paragraph" || block.type === "heading") {
      const runs = buildRuns(childrenOf(block), ctx);
      if (i === 0) {
        runs.unshift(
          new TextRun({ text: `${icon} `, font: { name: FONT_EMOJI }, size: BODY_SIZE_HALF_POINTS }),
        );
      }
      out.push(
        new Paragraph({
          style: "FwCallout",
          shading: { fill: CALLOUT_FILL },
          spacing: { after },
          children: runs,
        }),
      );
      return;
    }
    // Anything exotic inside a callout still exports as best-effort paragraphs.
    out.push(
      ...blockToDocChildren(block, 0, new OrderedNumbering(), isLastBlock && isLastInParent, ctx).filter(
        (c): c is Paragraph => c instanceof Paragraph,
      ),
    );
  });
  if (out.length === 0) {
    out.push(
      new Paragraph({
        style: "FwCallout",
        spacing: { after: lastChildAfterTwips(isLastInParent) },
        children: [],
      }),
    );
  }
  return out;
}

/** A quote exports as its paragraphs, indented and tagged `FwQuote`. */
function quoteToParagraphs(
  node: PMNode,
  contentWidthPx: number,
  ordered: OrderedNumbering,
  isLastInParent: boolean,
  ctx: ExportContext,
): DocChild[] {
  const out: DocChild[] = [];
  const blocks = childrenOf(node);
  blocks.forEach((block, i) => {
    const isLastBlock = i === blocks.length - 1;
    const after = isLastBlock ? lastChildAfterTwips(isLastInParent) : MIDDLE_AFTER_TWIPS;
    if (block.type === "paragraph" || block.type === "heading") {
      out.push(
        makeParagraph(
          block,
          null,
          { styleName: "FwQuote", indentLeftTwips: 720, spacingAfterTwips: after },
          ctx,
        ),
      );
    } else {
      out.push(...blockToDocChildren(block, contentWidthPx, ordered, isLastBlock && isLastInParent, ctx));
    }
  });
  return out.length > 0 ? out : [new Paragraph({ style: "FwQuote", children: [] })];
}

/**
 * A code block exports as one monospace, grey-shaded paragraph per source
 * line, tagged `FwCodeBlock`; the importer collapses a run of them back into
 * a single `fwCodeBlock` node. Lines are tight (the editor renders one `pre`);
 * only the block's 6 pt bottom gap is mirrored, on the final line.
 */
function codeBlockToParagraphs(node: PMNode, isLastInParent: boolean): Paragraph[] {
  const text = node.content
    ? childrenOf(node)
        .map((c) => String(c.text ?? ""))
        .join("")
    : "";
  const lines = text.length > 0 ? text.split("\n") : [""];
  return lines.map(
    (line, i) =>
      new Paragraph({
        style: "FwCodeBlock",
        shading: { fill: CODE_FILL },
        spacing: { after: i === lines.length - 1 ? lastChildAfterTwips(isLastInParent) : 0 },
        children: [new TextRun({ text: line, font: MONO_FONT, size: CODE_SIZE_HALF_POINTS })],
      }),
  );
}

/**
 * A toggle exports like a callout's flat shape: the summary paragraph carries
 * the `FwToggle` style plus a ▸/▾ glyph (which doubles as the collapsed flag
 * on import when style ids survive), and every body paragraph is tagged
 * `FwToggleBody` so the importer can re-attach it to the toggle.
 */
function toggleToParagraphs(node: PMNode, isLastInParent: boolean, ctx: ExportContext): Paragraph[] {
  const blocks = childrenOf(node);
  const collapsed = node.attrs?.collapsed === true;
  const glyph = collapsed ? "▸" : "▾";
  const out: Paragraph[] = [];
  blocks.forEach((block, i) => {
    const isLastBlock = i === blocks.length - 1;
    const after = isLastBlock ? lastChildAfterTwips(isLastInParent) : MIDDLE_AFTER_TWIPS;
    const style = i === 0 ? "FwToggle" : "FwToggleBody";
    if (block.type === "paragraph" || block.type === "heading") {
      const runs = buildRuns(childrenOf(block), ctx);
      if (i === 0) {
        runs.unshift(
          new TextRun({ text: `${glyph} `, font: { name: FONT_SYMBOL }, size: BODY_SIZE_HALF_POINTS }),
        );
      }
      out.push(
        new Paragraph({
          style,
          spacing: { after },
          ...(i === 0 ? {} : { indent: { left: 360 } }),
          children: runs,
        }),
      );
      return;
    }
    out.push(
      ...blockToDocChildren(block, 0, new OrderedNumbering(), isLastBlock && isLastInParent, ctx).filter(
        (c): c is Paragraph => c instanceof Paragraph,
      ),
    );
  });
  if (out.length === 0) {
    out.push(
      new Paragraph({
        style: "FwToggle",
        spacing: { after: lastChildAfterTwips(isLastInParent) },
        children: [],
      }),
    );
  }
  return out;
}

/**
 * The TOC node stores no content — the export renders the heading snapshot
 * collected in `buildExportContext` as `FwToc`-tagged paragraphs (one per
 * entry plus the title), which the importer collapses back into one node.
 */
function tocToParagraphs(ctx: ExportContext): Paragraph[] {
  const out: Paragraph[] = [new Paragraph({ style: "FwToc", children: [new TextRun("Table of contents")] })];
  for (const h of ctx.headings) {
    out.push(
      new Paragraph({
        style: "FwToc",
        indent: { left: (h.level - 1) * 360 },
        children: [new TextRun(h.text || "Untitled heading")],
      }),
    );
  }
  if (ctx.headings.length === 0) {
    out.push(new Paragraph({ style: "FwToc", children: [new TextRun("(no headings)")] }));
  }
  return out;
}

function blockToDocChildren(
  node: PMNode,
  contentWidthPx: number,
  ordered: OrderedNumbering,
  isLast = false,
  ctx: ExportContext,
): DocChild[] {
  switch (node.type) {
    case "pageBreak":
      return [new Paragraph({ children: [new PageBreak()] })];
    case "paragraph":
      return [makeParagraph(node, null, { spacingAfterTwips: isLast ? 0 : MIDDLE_AFTER_TWIPS }, ctx)];
    case "heading":
      // The CSS collapse targets `p:last-child`, so headings keep their 6 pt
      // gap even as the last block — matching the editor.
      return [makeParagraph(node, null, { spacingAfterTwips: MIDDLE_AFTER_TWIPS }, ctx)];
    // --- Notion-style blocks (see renderer/src/editor/blocks.ts) ----------
    // Word has no native task list / callout / quote-block / code-block, so
    // each export carries a custom style id (recognised on import) plus a
    // visual fallback that stays readable in Word itself.
    case "taskList":
      return taskListToParagraphs(node, isLast, ctx);
    case "callout":
      return calloutToParagraphs(node, isLast, ctx);
    case "fwQuote":
      return quoteToParagraphs(node, contentWidthPx, ordered, isLast, ctx);
    case "divider":
      return [
        new Paragraph({
          style: "FwDivider",
          children: [],
          spacing: { before: ptToTwips(10), after: isLast ? 0 : ptToTwips(10) },
          border: { bottom: { style: "single" as const, size: 6, color: DIVIDER_COLOR } },
        }),
      ];
    case "fwCodeBlock":
      return codeBlockToParagraphs(node, isLast);
    case "fwToggle":
      return toggleToParagraphs(node, isLast, ctx);
    case "fwToc":
      return tocToParagraphs(ctx);
    case "bulletList":
      return listItemsToParagraphs(node, "bullet", 0, null, isLast, ctx);
    case "orderedList": {
      const reference = ordered.next();
      return listItemsToParagraphs(node, "ordered", 0, reference, isLast, ctx);
    }
    case "table":
      return [pmTableToDocx(node, contentWidthPx, ordered, ctx)];
    case "image":
      return [imageToParagraph(node, contentWidthPx)];
    default:
      // Unknown block nodes are skipped; recurse so nested content is not lost.
      return childrenOf(node).flatMap((c, ci) =>
        blockToDocChildren(c, contentWidthPx, ordered, ci === childrenOf(node).length - 1 && isLast, ctx),
      );
  }
}

/** Cell content: any block nodes (paragraphs, lists, nested tables, images). */
function cellBlocks(
  cell: PMNode,
  contentWidthPx: number,
  ordered: OrderedNumbering,
  ctx: ExportContext,
): DocChild[] {
  const kids = childrenOf(cell);
  const out = kids.flatMap((c, ci) =>
    blockToDocChildren(c, contentWidthPx, ordered, ci === kids.length - 1, ctx),
  );
  // OOXML requires at least one paragraph in every cell.
  if (out.length === 0) out.push(new Paragraph({}));
  return out;
}

/**
 * Width the user dragged the column to, in twips.
 *
 * Tiptap keeps `colwidth` as an array of CSS pixel values (the resize handle
 * writes it); a table without resized columns has null/absent widths and is
 * left for Word to lay out automatically.
 */
function cellWidthTwips(attrs: Record<string, unknown>): number | undefined {
  const raw = attrs.colwidth;
  const first = Array.isArray(raw) ? raw[0] : raw;
  const px = num(first, 0);
  return px > 0 ? pxToTwip(px) : undefined;
}

/** Convert a Tiptap `table` node into a docx `Table` with simple grid borders. */
function pmTableToDocx(
  node: PMNode,
  contentWidthPx: number,
  ordered: OrderedNumbering,
  ctx: ExportContext,
): Table {
  const rows: TableRow[] = [];
  for (const rowNode of childrenOf(node)) {
    if (rowNode.type !== "tableRow") continue;
    const cells: TableCell[] = [];
    let isHeaderRow = false;
    for (const cellNode of childrenOf(rowNode)) {
      if (cellNode.type !== "tableCell" && cellNode.type !== "tableHeader") continue;
      if (cellNode.type === "tableHeader") isHeaderRow = true;
      const ca = cellNode.attrs ?? {};
      const widthTwips = cellWidthTwips(ca);
      cells.push(
        new TableCell({
          children: cellBlocks(cellNode, contentWidthPx, ordered, ctx),
          columnSpan: positiveInt(ca.colspan, 1),
          rowSpan: positiveInt(ca.rowspan, 1),
          verticalAlign: VerticalAlignTable.TOP,
          ...(widthTwips !== undefined ? { width: { size: widthTwips, type: WidthType.DXA } } : {}),
        }),
      );
    }
    if (cells.length === 0) continue;
    rows.push(new TableRow({ children: cells, ...(isHeaderRow ? { tableHeader: true } : {}) }));
  }
  // A table whose rows were all dropped (no `tableRow` children, or rows that
  // lost every cell — degenerate JSON from an import or an external editor)
  // still has to produce valid OOXML: `docx` throws RangeError on an empty
  // row list, which would abort the whole save. Export one empty cell instead.
  if (rows.length === 0) {
    rows.push(new TableRow({ children: [new TableCell({ children: [new Paragraph({})] })] }));
  }
  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: CELL_BORDER,
      bottom: CELL_BORDER,
      left: CELL_BORDER,
      right: CELL_BORDER,
      insideHorizontal: CELL_BORDER,
      insideVertical: CELL_BORDER,
    },
  });
}

/**
 * Document-level styles mirroring the editor's `.fw-page` CSS so Word and
 * Google Docs render the same look: body Calibri 11 pt with 1.15 lines,
 * black bold headings at 24/18/14 pt (the library defaults are blue and
 * 16/13/12 pt), and named definitions for the custom `Fw*` block styles so
 * they appear in Word's gallery with sane spacing. Runs carry font/size
 * directly only when a textStyle/code mark asked for it — direct formatting
 * always beats the style, and that is precisely why the old unconditional
 * 11 pt stamp used to flatten headings.
 */
function documentStyles() {
  return {
    default: {
      document: {
        run: { font: FONT_BODY, size: BODY_SIZE_HALF_POINTS },
        paragraph: { spacing: { line: LINE_SPACING_TWIPS, lineRule: LineRuleType.AUTO } },
      },
      heading1: { run: { size: HEADING_SIZE_HALF_POINTS[1], bold: true, color: HEADING_COLOR } },
      heading2: { run: { size: HEADING_SIZE_HALF_POINTS[2], bold: true, color: HEADING_COLOR } },
      heading3: { run: { size: HEADING_SIZE_HALF_POINTS[3], bold: true, color: HEADING_COLOR } },
    },
    // Definitions only — the visual signatures (shading, ☐/☑ glyphs, Consolas)
    // stay as direct formatting because Google Docs can drop style ids on
    // re-export and the importer's fallback detection keys on the visuals.
    paragraphStyles: [
      { id: "FwTaskItem", name: "FreeWord Task Item", basedOn: "Normal" as const },
      { id: "FwCallout", name: "FreeWord Callout", basedOn: "Normal" as const },
      { id: "FwQuote", name: "FreeWord Quote", basedOn: "Normal" as const, next: "Normal" as const },
      {
        id: "FwCodeBlock",
        name: "FreeWord Code Block",
        basedOn: "Normal" as const,
        run: { font: FONT_CODE, size: CODE_SIZE_HALF_POINTS },
      },
      { id: "FwDivider", name: "FreeWord Divider", basedOn: "Normal" as const },
      { id: "FwToggle", name: "FreeWord Toggle", basedOn: "Normal" as const },
      { id: "FwToggleBody", name: "FreeWord Toggle Body", basedOn: "Normal" as const },
      { id: "FwToc", name: "FreeWord Table of Contents", basedOn: "Normal" as const },
    ],
  };
}

/**
 * Header/footer text → docx runs: the `{{page}}` / `{{pages}}` tokens become
 * real `PAGE` / `NUMPAGES` fields so Word prints live page numbers.
 */
function bandRuns(text: string): ParagraphChild[] {
  const out: ParagraphChild[] = [];
  const parts = text.split(/(\{\{pages?\}\})/);
  for (const part of parts) {
    if (part === "{{page}}") out.push(new TextRun({ children: [PageNumber.CURRENT] }));
    else if (part === "{{pages}}") out.push(new TextRun({ children: [PageNumber.TOTAL_PAGES] }));
    else if (part) out.push(new TextRun({ text: part }));
  }
  return out;
}

function bandParagraph(text: unknown, align: unknown): Paragraph {
  const a = typeof align === "string" ? ALIGN_MAP[align] : undefined;
  return new Paragraph({
    children: bandRuns(String(text ?? "")),
    ...(a !== undefined ? { alignment: a } : {}),
  });
}

export async function toDocxBuffer(pm: Record<string, unknown>): Promise<Buffer> {
  // The payload arrives over IPC, so it may be anything at all. Treat a
  // missing/non-object document as an empty one rather than throwing — a
  // throw here aborts the save with no diagnostic.
  const doc = (pm && typeof pm === "object" ? pm : {}) as unknown as PMDoc;
  const attrs = typeof doc.attrs === "object" && doc.attrs !== null ? doc.attrs : {};
  const size: PageSize = attrs.size === "letter" ? "letter" : "a4";
  const ps = PAGE_SIZES[size];
  const mt = num(attrs.mt, 25.4);
  const mr = num(attrs.mr, 25.4);
  const mb = num(attrs.mb, 25.4);
  const ml = num(attrs.ml, 25.4);

  const ctx = buildExportContext(doc);

  const pages = childrenOf(doc).filter((n) => n.type === "page");
  // Page content width in CSS px (96 DPI), used to clamp oversized images.
  const contentWidthPx = mmToPx(Math.max(10, ps.w - ml - mr));
  const ordered = new OrderedNumbering();
  const children: DocChild[] = [];
  pages.forEach((page, i) => {
    const pageBlocks = childrenOf(page);
    const lastIsBreak = pageBlocks.length > 0 && pageBlocks[pageBlocks.length - 1].type === "pageBreak";
    pageBlocks.forEach((block, bi) => {
      children.push(...blockToDocChildren(block, contentWidthPx, ordered, bi === pageBlocks.length - 1, ctx));
    });
    // Page-boundary break between explicit pages, suppressed when the page
    // already ends with an explicit pageBreak node (no double breaks).
    if (i < pages.length - 1 && !lastIsBreak) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
  });

  const headerText = typeof attrs.header === "string" && attrs.header ? attrs.header : null;
  const footerText = typeof attrs.footer === "string" && attrs.footer ? attrs.footer : null;

  const docx = new Document({
    creator: "FreeWord",
    title: "FreeWord Document",
    styles: documentStyles(),
    numbering: { config: ordered.toConfig() },
    ...(ctx.commentDefs.length > 0 ? { comments: { children: ctx.commentDefs } } : {}),
    ...(ctx.footnoteTexts.size > 0
      ? {
          footnotes: Object.fromEntries(
            [...ctx.footnoteTexts.entries()].map(([ordinal, text]) => [
              String(ordinal),
              { children: [new Paragraph({ children: [new TextRun(`${ordinal}.  ${text}`)] })] },
            ]),
          ),
        }
      : {}),
    sections: [
      {
        properties: {
          page: {
            size: {
              width: mmToTwip(ps.w),
              height: mmToTwip(ps.h),
            },
            margin: {
              top: mmToTwip(mt),
              right: mmToTwip(mr),
              bottom: mmToTwip(mb),
              left: mmToTwip(ml),
            },
          },
        },
        ...(headerText !== null
          ? { headers: { default: new Header({ children: [bandParagraph(headerText, attrs.headerAlign)] }) } }
          : {}),
        ...(footerText !== null
          ? { footers: { default: new Footer({ children: [bandParagraph(footerText, attrs.footerAlign)] }) } }
          : {}),
        children,
      },
    ],
  });

  return Packer.toBuffer(docx);
}
