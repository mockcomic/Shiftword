/**
 * Typographic tokens — the single source of truth for the document's look.
 *
 * The editor CSS (`.fw-page` rules in renderer/src/styles.css), the standalone
 * HTML export (io/serialize.ts) and the DOCX export (main/docx-export.ts) all
 * derive their values from here so a document renders identically in
 * Shiftword, Microsoft Word, and Google Docs. The test-suite parses
 * styles.css and fails if it drifts from these numbers.
 */

/** Body font written into docx styles/defaults and the HTML export. */
export const FONT_BODY = "Calibri";

/** CSS font stack for the body (HTML export + editor chrome parity). */
export const FONT_STACK_BODY = `${FONT_BODY},'Segoe UI',system-ui,sans-serif`;

/** Monospace font for inline code and code blocks (Word has no code charstyle). */
export const FONT_CODE = "Consolas";

/** CSS font stack for monospace runs. */
export const FONT_STACK_CODE = `${FONT_CODE},'Courier New',monospace`;

/** Symbol fonts the exporter stamps on task-list glyphs and callout icons. */
export const FONT_SYMBOL = "Segoe UI Symbol";
export const FONT_EMOJI = "Segoe UI Emoji";

/** Body text size in points (11 pt, like Word's default Calibri documents). */
export const SIZE_BODY_PT = 11;

/** Code-block text size in points (0.95 em of 11 pt ≈ 10 pt). */
export const SIZE_CODE_PT = 10;

/** Heading sizes in points, mirroring `.fw-page h1..h3` in styles.css. */
export const HEADING_SIZE_PT: Record<1 | 2 | 3, number> = { 1: 24, 2: 18, 3: 14 };

/** Headings are black (Word's stock headings are blue — the editor's are not). */
export const HEADING_COLOR = "000000";

/** Paragraph line height as a multiplier (CSS unitless line-height). */
export const LINE_HEIGHT = 1.15;

/** Space below paragraphs/headings in points (CSS `margin: 0 0 6pt`). */
export const PARA_AFTER_PT = 6;

/** Word's hyperlink blue. */
export const HYPERLINK_COLOR = "0563C1";

// Visual signatures shared between the DOCX exporter and the importer's
// fallback block detection (io/import-docx.ts recognises blocks by these).
export const CALLOUT_FILL = "F1F1F1";
export const CODE_FILL = "F2F2F2";
export const DIVIDER_COLOR = "BFBFBF";

/** The two highlight colours the DOCX export can map to a real w:highlight. */
export const HIGHLIGHT_YELLOW = "FFFF00";
export const HIGHLIGHT_GREEN = "92D050";

/** Points → OOXML half-points. */
export function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2);
}

export const BODY_SIZE_HALF_POINTS = ptToHalfPoints(SIZE_BODY_PT); // 22
export const CODE_SIZE_HALF_POINTS = ptToHalfPoints(SIZE_CODE_PT); // 20
export const HEADING_SIZE_HALF_POINTS: Record<1 | 2 | 3, number> = {
  1: ptToHalfPoints(HEADING_SIZE_PT[1]), // 48
  2: ptToHalfPoints(HEADING_SIZE_PT[2]), // 36
  3: ptToHalfPoints(HEADING_SIZE_PT[3]), // 28
};

/** Line-height multiplier → OOXML `w:spacing w:line` (240 = single). */
export const LINE_SPACING_TWIPS = Math.round(240 * LINE_HEIGHT); // 276

/** Points → OOXML twips (1 pt = 20 twips). */
export function ptToTwips(pt: number): number {
  return Math.round(pt * 20);
}

export const PARA_AFTER_TWIPS = ptToTwips(PARA_AFTER_PT); // 120

/** "RRGGBB" → the normalised `el.style.backgroundColor` form used in imports. */
export function hexToRgbCss(hex: string): string {
  const n = parseInt(hex.replace(/^#/, ""), 16);
  return `rgb(${(n >> 16) & 0xff},${(n >> 8) & 0xff},${n & 0xff})`;
}
