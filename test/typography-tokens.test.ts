// Drift guard: the shared typography tokens must match the editor CSS.
// The DOCX/HTML exporters render from the tokens; the editor renders from
// styles.css — if either side changes without the other, this fails.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  FONT_BODY,
  HEADING_SIZE_PT,
  LINE_HEIGHT,
  PARA_AFTER_PT,
  SIZE_BODY_PT,
} from "../src/shared/typography";

const css = readFileSync(fileURLToPath(new URL("../src/renderer/src/styles.css", import.meta.url)), "utf8");

describe("typography tokens vs editor CSS", () => {
  it("body font family and size match `.fw-page`", () => {
    expect(css).toMatch(new RegExp(`\\.fw-page\\s*\\{[^}]*font-family:\\s*${FONT_BODY}\\b`));
    expect(css).toMatch(new RegExp(`\\.fw-page\\s*\\{[^}]*font-size:\\s*${SIZE_BODY_PT}pt\\b`));
  });

  it("heading sizes match `.fw-page h1..h3`", () => {
    for (const level of [1, 2, 3] as const) {
      expect(css, `h${level}`).toMatch(
        new RegExp(`\\.fw-page h${level}\\s*\\{[^}]*font-size:\\s*${HEADING_SIZE_PT[level]}pt\\b`),
      );
    }
  });

  it("line height and paragraph gap match the block rules", () => {
    expect(css).toMatch(new RegExp(`line-height:\\s*${LINE_HEIGHT}\\b`));
    expect(css).toMatch(new RegExp(`margin:\\s*0\\s+0\\s+${PARA_AFTER_PT}pt`));
  });
});
