import { describe, expect, it } from "vitest";

import {
  escapeInline,
  renderStyledTokens,
  type StyledToken,
  wrapEmphasis,
} from "../src/inline.js";

describe("escapeInline", () => {
  it("escapes every character that would otherwise be Markdown syntax", () => {
    expect(escapeInline("a\\b`c*d_e[f]g<h>i|j")).toBe(
      "a\\\\b\\`c\\*d\\_e\\[f\\]g\\<h\\>i\\|j",
    );
  });

  it("leaves ordinary prose untouched", () => {
    expect(escapeInline("Hello, world. 100% done!")).toBe(
      "Hello, world. 100% done!",
    );
  });
});

describe("wrapEmphasis", () => {
  const plain: StyledToken = { text: "" };

  it("wraps bold, italic and both", () => {
    expect(wrapEmphasis("x", { ...plain, bold: true })).toBe("**x**");
    expect(wrapEmphasis("x", { ...plain, italic: true })).toBe("*x*");
    expect(wrapEmphasis("x", { ...plain, bold: true, italic: true })).toBe(
      "***x***",
    );
  });

  it("returns escaped text when unstyled", () => {
    expect(wrapEmphasis("a*b", plain)).toBe("a\\*b");
  });

  it("uses code spans for monospace and does not escape inside them", () => {
    expect(wrapEmphasis("a*b_c", { ...plain, monospace: true })).toBe("`a*b_c`");
  });

  it("swaps backticks for quotes so a code span cannot break out", () => {
    expect(wrapEmphasis("a`b", { ...plain, monospace: true })).toBe("`a'b`");
  });

  it("lets monospace win over bold and italic", () => {
    expect(
      wrapEmphasis("x", { ...plain, monospace: true, bold: true, italic: true }),
    ).toBe("`x`");
  });
});

describe("renderStyledTokens", () => {
  it("merges adjacent same-styled tokens into one emphasis span", () => {
    // The bug this guards: Word splits a bold word across two runs, and
    // wrapping each separately emits `**bo****ld**`, which CommonMark renders
    // literally rather than as emphasis.
    expect(
      renderStyledTokens([
        { text: "bo", bold: true },
        { text: "ld", bold: true },
      ]),
    ).toBe("**bold**");
  });

  it("does not merge tokens whose styles differ", () => {
    expect(
      renderStyledTokens([
        { text: "bold", bold: true },
        { text: "italic", italic: true },
      ]),
    ).toBe("**bold***italic*");
  });

  it("keeps whitespace outside the delimiters", () => {
    // `** bold **` is not emphasis in CommonMark; the markers must hug the word.
    expect(
      renderStyledTokens([{ text: " bold ", bold: true }, { text: "tail" }]),
    ).toBe("**bold** tail");
  });

  it("emits literal tokens verbatim, unescaped and unmerged", () => {
    expect(
      renderStyledTokens([
        { text: "see " },
        { text: "[docs](https://example.com/a_b)", literal: true },
        { text: " now" },
      ]),
    ).toBe("see [docs](https://example.com/a_b) now");
  });

  it("never merges two literals together", () => {
    expect(
      renderStyledTokens([
        { text: "[a](x)", literal: true },
        { text: "[b](y)", literal: true },
      ]),
    ).toBe("[a](x)[b](y)");
  });

  it("collapses runs of spaces and tabs but preserves newlines", () => {
    // Hard line breaks arrive as literals and must survive the collapse.
    expect(
      renderStyledTokens([
        { text: "a   \t b" },
        { text: "\n", literal: true },
        { text: "c" },
      ]),
    ).toBe("a b\nc");
  });

  it("drops empty tokens and trims the result", () => {
    expect(
      renderStyledTokens([{ text: "" }, { text: "  hi  " }, { text: "" }]),
    ).toBe("hi");
  });

  it("passes through a whitespace-only token without emphasis markers", () => {
    expect(
      renderStyledTokens([
        { text: "a", bold: true },
        { text: " ", bold: true },
        { text: "b", bold: true },
      ]),
    ).toBe("**a b**");
  });

  it("returns an empty string for no tokens", () => {
    expect(renderStyledTokens([])).toBe("");
  });

  it("escapes syntax inside styled text", () => {
    expect(renderStyledTokens([{ text: "a_b", bold: true }])).toBe("**a\\_b**");
  });
});
