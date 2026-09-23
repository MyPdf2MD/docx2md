import { describe, expect, it } from "vitest";

import { renderMarkdown } from "../src/md-renderer.js";
import type { ASTNode } from "../src/types.js";

const para = (content: string): ASTNode => ({ type: "paragraph", content });

describe("renderMarkdown", () => {
  it("returns an empty string for no nodes", () => {
    expect(renderMarkdown([])).toBe("");
  });

  it("terminates non-empty output with exactly one newline", () => {
    expect(renderMarkdown([para("hello")])).toBe("hello\n");
  });

  it("separates blocks with one blank line", () => {
    expect(renderMarkdown([para("one"), para("two")])).toBe("one\n\ntwo\n");
  });

  it("drops nodes that render to nothing, without leaving a gap", () => {
    expect(renderMarkdown([para("one"), para("   "), para("two")])).toBe(
      "one\n\ntwo\n",
    );
  });
});

describe("headings", () => {
  it("emits one hash per level", () => {
    expect(
      renderMarkdown([{ type: "heading", level: 3, content: "Title" }]),
    ).toBe("### Title\n");
  });

  it("defaults to level 1 when no level is given", () => {
    expect(renderMarkdown([{ type: "heading", content: "Title" }])).toBe(
      "# Title\n",
    );
  });

  it("clamps out-of-range levels into 1–6", () => {
    expect(renderMarkdown([{ type: "heading", level: 9, content: "x" }])).toBe(
      "###### x\n",
    );
    expect(renderMarkdown([{ type: "heading", level: 0, content: "x" }])).toBe(
      "# x\n",
    );
    expect(renderMarkdown([{ type: "heading", level: -3, content: "x" }])).toBe(
      "# x\n",
    );
  });

  it("skips an empty heading rather than emitting bare hashes", () => {
    expect(renderMarkdown([{ type: "heading", level: 2, content: "  " }])).toBe(
      "",
    );
  });
});

describe("list items", () => {
  it("keeps items of one list on consecutive lines", () => {
    expect(
      renderMarkdown([
        { type: "list_item", content: "a" },
        { type: "list_item", content: "b" },
      ]),
    ).toBe("- a\n- b\n");
  });

  it("indents by two spaces per nesting level", () => {
    expect(
      renderMarkdown([
        { type: "list_item", content: "a" },
        { type: "list_item", level: 1, content: "b" },
        { type: "list_item", level: 2, content: "c" },
      ]),
    ).toBe("- a\n  - b\n    - c\n");
  });

  it("numbers ordered items sequentially when the source gave no ordinal", () => {
    expect(
      renderMarkdown([
        { type: "list_item", isOrdered: true, content: "a" },
        { type: "list_item", isOrdered: true, content: "b" },
        { type: "list_item", isOrdered: true, content: "c" },
      ]),
    ).toBe("1. a\n2. b\n3. c\n");
  });

  it("trusts the ordinal printed in the source when there is one", () => {
    // A PDF that starts its list at 5 should stay at 5 — renumbering it would
    // silently contradict the document.
    expect(
      renderMarkdown([
        { type: "list_item", isOrdered: true, ordinal: 5, content: "a" },
        { type: "list_item", isOrdered: true, content: "b" },
      ]),
    ).toBe("5. a\n6. b\n");
  });

  it("restarts numbering after a non-list block interrupts", () => {
    expect(
      renderMarkdown([
        { type: "list_item", isOrdered: true, content: "a" },
        para("interruption"),
        { type: "list_item", isOrdered: true, content: "b" },
      ]),
    ).toBe("1. a\n\ninterruption\n\n1. b\n");
  });

  it("tracks a separate counter per nesting depth", () => {
    expect(
      renderMarkdown([
        { type: "list_item", isOrdered: true, content: "a" },
        { type: "list_item", isOrdered: true, level: 1, content: "a1" },
        { type: "list_item", isOrdered: true, level: 1, content: "a2" },
        { type: "list_item", isOrdered: true, content: "b" },
      ]),
    ).toBe("1. a\n  1. a1\n  2. a2\n2. b\n");
  });

  it("skips an empty item", () => {
    expect(renderMarkdown([{ type: "list_item", content: "  " }])).toBe("");
  });
});

describe("tables", () => {
  it("emits a GFM pipe table with a delimiter row", () => {
    expect(
      renderMarkdown([
        {
          type: "table",
          content: "",
          rows: [
            ["a", "b"],
            ["1", "2"],
          ],
        },
      ]),
    ).toBe("| a | b |\n| --- | --- |\n| 1 | 2 |\n");
  });

  it("pads short rows to the widest row", () => {
    expect(
      renderMarkdown([
        { type: "table", content: "", rows: [["a", "b", "c"], ["1"]] },
      ]),
    ).toBe("| a | b | c |\n| --- | --- | --- |\n| 1 |  |  |\n");
  });

  it("replaces newlines inside a cell with <br> so the row survives", () => {
    expect(
      renderMarkdown([
        { type: "table", content: "", rows: [["a\nb", "c"]] },
      ]),
    ).toBe("| a<br>b | c |\n| --- | --- |\n");
  });

  it("drops rows that are entirely empty", () => {
    expect(
      renderMarkdown([
        {
          type: "table",
          content: "",
          rows: [
            ["a", "b"],
            ["", "  "],
            ["1", "2"],
          ],
        },
      ]),
    ).toBe("| a | b |\n| --- | --- |\n| 1 | 2 |\n");
  });

  it("refuses a single-column table — that is a paragraph, not a table", () => {
    expect(
      renderMarkdown([{ type: "table", content: "", rows: [["a"], ["b"]] }]),
    ).toBe("");
  });

  it("renders nothing when there are no rows at all", () => {
    expect(renderMarkdown([{ type: "table", content: "", rows: [] }])).toBe("");
    expect(renderMarkdown([{ type: "table", content: "" }])).toBe("");
  });
});

describe("code blocks", () => {
  it("fences with three backticks", () => {
    expect(
      renderMarkdown([{ type: "code_block", content: "const x = 1;" }]),
    ).toBe("```\nconst x = 1;\n```\n");
  });

  it("widens the fence past the longest backtick run in the code", () => {
    // A ``` inside the body would otherwise close the block early.
    expect(
      renderMarkdown([{ type: "code_block", content: "a ``` b" }]),
    ).toBe("````\na ``` b\n````\n");
  });

  it("skips a blank code block", () => {
    expect(renderMarkdown([{ type: "code_block", content: "  \n " }])).toBe("");
  });

  it("keeps interior indentation but strips trailing whitespace", () => {
    expect(
      renderMarkdown([{ type: "code_block", content: "if (x) {\n  y();\n}  \n" }]),
    ).toBe("```\nif (x) {\n  y();\n}\n```\n");
  });
});

describe("blockquotes", () => {
  it("prefixes every line", () => {
    expect(
      renderMarkdown([{ type: "blockquote", content: "one\ntwo" }]),
    ).toBe("> one\n> two\n");
  });

  it("emits a bare > for a blank interior line, with no trailing space", () => {
    expect(
      renderMarkdown([{ type: "blockquote", content: "one\n\ntwo" }]),
    ).toBe("> one\n>\n> two\n");
  });

  it("skips an empty quote", () => {
    expect(renderMarkdown([{ type: "blockquote", content: " " }])).toBe("");
  });
});

describe("paragraph block-start escaping", () => {
  it.each([
    ["# not a heading", "\\# not a heading"],
    ["> not a quote", "\\> not a quote"],
    ["- not a bullet", "\\- not a bullet"],
    ["* not a bullet", "\\* not a bullet"],
    ["+ not a bullet", "\\+ not a bullet"],
    ["1. not a list", "\\1. not a list"],
    ["12) not a list", "\\12) not a list"],
  ])("escapes %j so it is not re-parsed as a block", (input, expected) => {
    expect(renderMarkdown([para(input)])).toBe(`${expected}\n`);
  });

  it("leaves text that only resembles syntax mid-line alone", () => {
    expect(renderMarkdown([para("a # b")])).toBe("a # b\n");
    expect(renderMarkdown([para("2024 was a year")])).toBe("2024 was a year\n");
  });
});
