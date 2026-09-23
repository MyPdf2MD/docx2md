/**
 * The DOCX parser's specification.
 *
 * `node`, with no DOM: `src/xml.ts` reads the OOXML, so the parser under test
 * here is byte for byte the one that ships to a browser. It used to need jsdom
 * for `DOMParser`, and a Node-only stand-in would have exercised a different
 * namespace implementation than the shipped one — which, for a parser whose
 * every element match is namespace-aware, is exactly the wrong thing to stub.
 */

import { describe, expect, it } from "vitest";

import {
  cell,
  makeDocx,
  numberingDefinition,
  numPr,
  paragraph,
  pStyle,
  row,
  run,
  styleDefinition,
  table,
} from "./docx-fixtures.js";
import { parseDocx } from "../src/index.js";
import { renderMarkdown } from "../src/md-renderer.js";
import { DocxErrorCode, isConversionError } from "../src/types.js";

/** Parse straight to Markdown — what the package actually produces. */
async function toMarkdown(parts: Parameters<typeof makeDocx>[0]) {
  const { nodes } = await parseDocx(makeDocx(parts));
  return renderMarkdown(nodes);
}

const HEADING_STYLES = [
  styleDefinition("Heading1", { name: "heading 1" }),
  styleDefinition("Heading2", { name: "heading 2" }),
  styleDefinition("Title", { name: "Title" }),
  styleDefinition("Subtitle", { name: "Subtitle" }),
].join("");
/* ------------------------------------------------------------------ */

describe("headings", () => {
  it("maps Word's heading styles to the matching Markdown level", async () => {
    const markdown = await toMarkdown({
      styles: HEADING_STYLES,
      body: [
        paragraph(run("Chapter"), pStyle("Heading1")),
        paragraph(run("Section"), pStyle("Heading2")),
        paragraph(run("Body")),
      ].join(""),
    });

    expect(markdown).toBe("# Chapter\n\n## Section\n\nBody\n");
  });

  it("treats Title and Subtitle as the document's top two levels", async () => {
    const markdown = await toMarkdown({
      styles: HEADING_STYLES,
      body: [
        paragraph(run("The Report"), pStyle("Title")),
        paragraph(run("Second quarter"), pStyle("Subtitle")),
      ].join(""),
    });

    expect(markdown).toBe("# The Report\n\n## Second quarter\n");
  });

  it("follows a basedOn chain to find the heading level", async () => {
    // Custom template styles inherit from the built-ins rather than repeating
    // their names, so a converter that only reads the style's own name sees
    // "Report Section Head" and returns a paragraph.
    const markdown = await toMarkdown({
      styles:
        HEADING_STYLES +
        styleDefinition("ReportHead", {
          name: "Report Section Head",
          basedOn: "Heading2",
        }),
      body: paragraph(run("Findings"), pStyle("ReportHead")),
    });

    expect(markdown).toBe("## Findings\n");
  });

  it("reads the style id when the display name is localized", async () => {
    // A German Word writes the name as "Überschrift 1" and keeps the id.
    const markdown = await toMarkdown({
      styles: styleDefinition("Heading1", { name: "Überschrift 1" }),
      body: paragraph(run("Einleitung"), pStyle("Heading1")),
    });

    expect(markdown).toBe("# Einleitung\n");
  });

  it("falls back to outlineLvl only when no name says anything", async () => {
    const markdown = await toMarkdown({
      styles: styleDefinition("Custom", { name: "Custom", outlineLevel: 2 }),
      body: paragraph(run("Third level"), pStyle("Custom")),
    });

    expect(markdown).toBe("### Third level\n");
  });

  it("does not emphasise inside a heading", async () => {
    // "# **Title**" is noise: the heading is already the emphasis.
    const markdown = await toMarkdown({
      styles: HEADING_STYLES,
      body: paragraph(run("Bold title", "<w:b/>"), pStyle("Heading1")),
    });

    expect(markdown).toBe("# Bold title\n");
  });

  it("survives a basedOn cycle rather than hanging", async () => {
    const markdown = await toMarkdown({
      styles:
        styleDefinition("A", { name: "A", basedOn: "B" }) +
        styleDefinition("B", { name: "B", basedOn: "A" }),
      body: paragraph(run("Text"), pStyle("A")),
    });

    expect(markdown).toBe("Text\n");
  });
});

describe("lists", () => {
  it("distinguishes bulleted from numbered levels", async () => {
    const markdown = await toMarkdown({
      numbering:
        numberingDefinition("1", "10", ["bullet"]) +
        numberingDefinition("2", "20", ["decimal"]),
      body: [
        paragraph(run("Loose item"), numPr("1")),
        paragraph(run("First step"), numPr("2")),
        paragraph(run("Second step"), numPr("2")),
      ].join(""),
    });

    expect(markdown).toBe("- Loose item\n\n1. First step\n2. Second step\n");
  });

  it("nests by the declared ilvl", async () => {
    const markdown = await toMarkdown({
      numbering: numberingDefinition("1", "10", ["decimal", "lowerLetter", "lowerRoman"]),
      body: [
        paragraph(run("Top"), numPr("1", 0)),
        paragraph(run("Middle"), numPr("1", 1)),
        paragraph(run("Deep"), numPr("1", 2)),
      ].join(""),
    });

    expect(markdown).toBe("1. Top\n  1. Middle\n    1. Deep\n");
  });

  it("treats a numFmt of none as a bullet rather than inventing a number", async () => {
    const markdown = await toMarkdown({
      numbering: numberingDefinition("1", "10", ["none"]),
      body: paragraph(run("Unmarked"), numPr("1")),
    });

    expect(markdown).toBe("- Unmarked\n");
  });

  it("lets a heading style outrank numbering", async () => {
    // Word's numbered-heading styles carry both. Emitting "1. Introduction" as
    // a list item loses the document's entire outline.
    const markdown = await toMarkdown({
      styles: HEADING_STYLES,
      numbering: numberingDefinition("1", "10", ["decimal"]),
      body: paragraph(run("Introduction"), pStyle("Heading1") + numPr("1")),
    });

    expect(markdown).toBe("# Introduction\n");
  });

  it("ignores numId 0, which is Word removing a style's numbering", async () => {
    const markdown = await toMarkdown({
      numbering: numberingDefinition("1", "10", ["decimal"]),
      body: paragraph(run("Not a list"), numPr("0")),
    });

    expect(markdown).toBe("Not a list\n");
  });
});

describe("tables", () => {
  it("emits a GFM pipe table with the first row as the header", async () => {
    const markdown = await toMarkdown({
      body: table(
        row(cell(paragraph(run("Name"))) + cell(paragraph(run("Size")))) +
          row(cell(paragraph(run("Report"))) + cell(paragraph(run("2 MB")))),
      ),
    });

    expect(markdown).toBe(
      "| Name | Size |\n| --- | --- |\n| Report | 2 MB |\n",
    );
  });

  it("pads a horizontally merged cell so later columns stay aligned", async () => {
    // Without the padding, "Total" lands under "Q1" instead of under "Q2".
    const markdown = await toMarkdown({
      body: table(
        row(
          cell(paragraph(run("Q1"))) +
            cell(paragraph(run("Q2"))) +
            cell(paragraph(run("Q3"))),
        ) +
          row(
            cell(paragraph(run("Merged")), '<w:gridSpan w:val="2"/>') +
              cell(paragraph(run("Total"))),
          ),
      ),
    });

    expect(markdown).toBe(
      "| Q1 | Q2 | Q3 |\n| --- | --- | --- |\n| Merged |  | Total |\n",
    );
  });

  it("leaves a vertical-merge continuation empty rather than repeating it", async () => {
    const markdown = await toMarkdown({
      body: table(
        row(cell(paragraph(run("Region"))) + cell(paragraph(run("Value")))) +
          row(cell(paragraph(run("North"))) + cell(paragraph(run("10")))) +
          row(
            cell(paragraph(run("North")), '<w:vMerge w:val="continue"/>') +
              cell(paragraph(run("20"))),
          ),
      ),
    });

    expect(markdown).toBe(
      "| Region | Value |\n| --- | --- |\n| North | 10 |\n|  | 20 |\n",
    );
  });

  it("emits a single-column table as paragraphs instead of losing it", async () => {
    // `md-renderer` drops tables narrower than two columns, and Word uses
    // exactly that shape for callouts and sidebars.
    const markdown = await toMarkdown({
      body: table(
        row(cell(paragraph(run("A callout")))) +
          row(cell(paragraph(run("A second line")))),
      ),
    });

    expect(markdown).toBe("A callout\n\nA second line\n");
  });

  it("turns a line break inside a cell into a newline, not a hard break", async () => {
    // `md-renderer` converts newlines in cells to <br>; a CommonMark hard
    // break (backslash + newline) would split the pipe row instead.
    const markdown = await toMarkdown({
      body: table(
        row(
          cell(
            paragraph(`${run("One")}<w:r><w:br/></w:r>${run("Two")}`),
          ) + cell(paragraph(run("Other"))),
        ) + row(cell(paragraph(run("a"))) + cell(paragraph(run("b")))),
      ),
    });

    expect(markdown).toContain("| One<br>Two | Other |");
  });

  it("flattens a nested table to lines rather than dropping it", async () => {
    const markdown = await toMarkdown({
      body: table(
        row(
          cell(
            paragraph(run("Outer")) +
              table(row(cell(paragraph(run("in1"))) + cell(paragraph(run("in2"))))),
          ) + cell(paragraph(run("Right"))),
        ) + row(cell(paragraph(run("x"))) + cell(paragraph(run("y")))),
      ),
    });

    expect(markdown).toContain("Outer<br>in1 in2");
  });
});

describe("inline formatting", () => {
  it("carries bold, italic, strikethrough and monospace", async () => {
    const markdown = await toMarkdown({
      body: paragraph(
        run("bold", "<w:b/>") +
          run(" ") +
          run("italic", "<w:i/>") +
          run(" ") +
          run("struck", "<w:strike/>") +
          run(" ") +
          run("code", '<w:rFonts w:ascii="Consolas"/>'),
      ),
    });

    expect(markdown).toBe("**bold** *italic* ~~struck~~ `code`\n");
  });

  it("honours a toggle switched off rather than treating presence as true", async () => {
    // `<w:b w:val="0"/>` turns bold OFF. Reading the element's presence as
    // "true" bolds every run that explicitly disables it.
    const markdown = await toMarkdown({
      body: paragraph(run("plain", '<w:b w:val="0"/>')),
    });

    expect(markdown).toBe("plain\n");
  });

  it("reads the complex-script twins that carry RTL and Indic styling", async () => {
    const markdown = await toMarkdown({
      body: paragraph(run("عريض", "<w:bCs/>")),
    });

    expect(markdown).toBe("**عريض**\n");
  });

  it("merges runs Word split mid-word instead of emitting **bo****ld**", async () => {
    const markdown = await toMarkdown({
      body: paragraph(run("bo", "<w:b/>") + run("ld", "<w:b/>")),
    });

    expect(markdown).toBe("**bold**\n");
  });

  it("resolves a hyperlink through the relationships part", async () => {
    const markdown = await toMarkdown({
      rels:
        '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a" TargetMode="External"/>',
      body: paragraph(
        `<w:hyperlink r:id="rId7">${run("the docs")}</w:hyperlink>`,
      ),
    });

    expect(markdown).toBe("[the docs](https://example.com/a)\n");
  });

  it("brackets a destination containing spaces or parens", async () => {
    const markdown = await toMarkdown({
      rels:
        '<Relationship Id="rId7" Type="hyperlink" Target="https://example.com/a b(c)"/>',
      body: paragraph(`<w:hyperlink r:id="rId7">${run("link")}</w:hyperlink>`),
    });

    expect(markdown).toBe("[link](<https://example.com/a b(c)>)\n");
  });

  it("degrades an internal anchor to its label", async () => {
    // A bookmark jump points into a document that stops existing once this is
    // Markdown. The label is the content; the jump is not.
    const markdown = await toMarkdown({
      body: paragraph(
        `<w:hyperlink w:anchor="_Toc123">${run("See chapter 2")}</w:hyperlink>`,
      ),
    });

    expect(markdown).toBe("See chapter 2\n");
  });

  it("keeps the spaces that live in their own runs", async () => {
    const markdown = await toMarkdown({
      body: paragraph(run("one") + run(" ") + run("two")),
    });

    expect(markdown).toBe("one two\n");
  });

  it("escapes Markdown punctuation that came from the document's text", async () => {
    const markdown = await toMarkdown({
      body: paragraph(run("a * b _ c [d]")),
    });

    expect(markdown).toBe("a \\* b \\_ c \\[d\\]\n");
  });
});

describe("tracked changes and fields", () => {
  it("keeps insertions and drops deletions", async () => {
    const markdown = await toMarkdown({
      body: paragraph(
        run("The ") +
          `<w:ins>${run("new ")}</w:ins>` +
          `<w:del><w:r><w:delText>old </w:delText></w:r></w:del>` +
          run("wording"),
      ),
    });

    expect(markdown).toBe("The new wording\n");
  });

  it("skips field instructions, which are machine syntax and not content", async () => {
    const markdown = await toMarkdown({
      body: paragraph(
        `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>` +
          run("Contents"),
      ),
    });

    expect(markdown).toBe("Contents\n");
  });

  it("recurses into content controls, which wrap real text", async () => {
    const markdown = await toMarkdown({
      body: `<w:sdt><w:sdtContent>${paragraph(run("Inside a control"))}</w:sdtContent></w:sdt>`,
    });

    expect(markdown).toBe("Inside a control\n");
  });
});

describe("code blocks and quotes", () => {
  it("fuses consecutive code paragraphs into one fence", async () => {
    // Word has no multi-line code block: five lines of code are five
    // paragraphs. Five fences instead of one is the most visible failure a
    // DOCX conversion can have.
    const markdown = await toMarkdown({
      styles: styleDefinition("SourceCode", { name: "Source Code" }),
      body: [
        paragraph(run("const a = 1;"), pStyle("SourceCode")),
        paragraph(run("const b = 2;"), pStyle("SourceCode")),
        paragraph(run("After")),
      ].join(""),
    });

    expect(markdown).toBe("```\nconst a = 1;\nconst b = 2;\n```\n\nAfter\n");
  });

  it("detects a code paragraph from every run being monospaced", async () => {
    const markdown = await toMarkdown({
      body: paragraph(run("npm install", '<w:rFonts w:ascii="Courier New"/>')),
    });

    expect(markdown).toBe("```\nnpm install\n```\n");
  });

  it("fuses consecutive quote paragraphs into one blockquote", async () => {
    const markdown = await toMarkdown({
      styles: styleDefinition("Quote", { name: "Quote" }),
      body: [
        paragraph(run("First line"), pStyle("Quote")),
        paragraph(run("Second line"), pStyle("Quote")),
      ].join(""),
    });

    expect(markdown).toBe("> First line\n> Second line\n");
  });
});

describe("images", () => {
  it("counts a drawing rather than dropping it silently", async () => {
    const { droppedImageCount } = await parseDocx(
      makeDocx({
        body: paragraph(`<w:r><w:drawing/></w:r>${run("Figure 1")}`),
      }),
    );

    expect(droppedImageCount).toBe(1);
  });

  it("counts an AlternateContent figure once, not once per representation", async () => {
    // mc:Choice and mc:Fallback are two encodings of the SAME object. Walking
    // both reports twice as many lost figures as the document contains.
    const { droppedImageCount } = await parseDocx(
      makeDocx({
        body: paragraph(
          "<w:r><mc:AlternateContent>" +
            "<mc:Choice Requires=\"wps\"><w:drawing/></mc:Choice>" +
            "<mc:Fallback><w:pict/></mc:Fallback>" +
            "</mc:AlternateContent></w:r>",
        ),
      }),
    );

    expect(droppedImageCount).toBe(1);
  });

  it("reads a text box as text and does not count it as a lost image", async () => {
    // Word stores floating text in w:drawing too. Counting one of those as a
    // dropped figure would warn about losing a picture while actually losing
    // a paragraph.
    const { nodes, droppedImageCount } = await parseDocx(
      makeDocx({
        body: paragraph(
          `<w:r><w:drawing><w:txbxContent>${paragraph(run("Pull quote"))}</w:txbxContent></w:drawing></w:r>`,
        ),
      }),
    );

    expect(droppedImageCount).toBe(0);
    expect(renderMarkdown(nodes)).toBe("Pull quote\n");
  });
});

describe("failure modes", () => {
  async function codeFor(bytes: Uint8Array): Promise<string> {
    try {
      await parseDocx(bytes);
      return "NO_ERROR";
    } catch (error) {
      return isConversionError(error) ? error.code : "NOT_A_CONVERSION_ERROR";
    }
  }

  it("rejects bytes that are not a ZIP at all", async () => {
    expect(await codeFor(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toBe(
      DocxErrorCode.INVALID_DOCX,
    );
  });

  it("rejects a ZIP with no word/document.xml — an .xlsx renamed, usually", async () => {
    expect(
      await codeFor(makeDocx({ body: "", omitDocument: true, extra: { "xl/workbook.xml": "<x/>" } })),
    ).toBe(DocxErrorCode.INVALID_DOCX);
  });

  it("rejects malformed XML rather than converting it to silence", async () => {
    const broken = makeDocx({ body: "" });
    // Rebuild with deliberately unbalanced markup.
    const bytes = makeDocx({ body: "<w:p><w:r><w:t>unclosed" });
    expect(broken.length).toBeGreaterThan(0);
    expect(await codeFor(bytes)).toBe(DocxErrorCode.INVALID_DOCX);
  });

  it("reports an already-aborted signal as a cancellation", async () => {
    const controller = new AbortController();
    controller.abort();

    let code = "NO_ERROR";
    try {
      await parseDocx(makeDocx({ body: paragraph(run("Text")) }), {
        signal: controller.signal,
      });
    } catch (error) {
      code = isConversionError(error) ? error.code : "NOT_A_CONVERSION_ERROR";
    }

    expect(code).toBe(DocxErrorCode.CANCELLED);
  });

  it("reports a passed deadline as a timeout", async () => {
    let code = "NO_ERROR";
    try {
      await parseDocx(makeDocx({ body: paragraph(run("Text")) }), {
        deadline: performance.now() - 1,
      });
    } catch (error) {
      code = isConversionError(error) ? error.code : "NOT_A_CONVERSION_ERROR";
    }

    expect(code).toBe(DocxErrorCode.TIMEOUT);
  });

  it("converts a document with no styles or numbering parts at all", async () => {
    // Minimal generators (Google Docs export, some scripts) omit both.
    const markdown = await toMarkdown({ body: paragraph(run("Just text")) });
    expect(markdown).toBe("Just text\n");
  });

  it("drops empty paragraphs rather than emitting blank blocks", async () => {
    const markdown = await toMarkdown({
      body: paragraph(run("One")) + "<w:p/>" + paragraph(run("Two")),
    });

    expect(markdown).toBe("One\n\nTwo\n");
  });
});

describe("a whole document", () => {
  it("converts headings, lists, a table and a link together", async () => {
    const markdown = await toMarkdown({
      styles: HEADING_STYLES,
      numbering: numberingDefinition("1", "10", ["decimal"]),
      rels:
        '<Relationship Id="rId1" Type="hyperlink" Target="https://example.com"/>',
      body: [
        paragraph(run("Quarterly Report"), pStyle("Title")),
        paragraph(run("Summary"), pStyle("Heading1")),
        paragraph(
          run("Revenue rose. See ") +
            `<w:hyperlink r:id="rId1">${run("the appendix")}</w:hyperlink>` +
            run("."),
        ),
        paragraph(run("Grew headcount"), numPr("1")),
        paragraph(run("Opened two offices"), numPr("1")),
        table(
          row(cell(paragraph(run("Metric"))) + cell(paragraph(run("Value")))) +
            row(cell(paragraph(run("Revenue"))) + cell(paragraph(run("1.2M")))),
        ),
      ].join(""),
    });

    expect(markdown).toBe(
      [
        "# Quarterly Report",
        "",
        "# Summary",
        "",
        "Revenue rose. See [the appendix](https://example.com).",
        "",
        "1. Grew headcount",
        "2. Opened two offices",
        "",
        "| Metric | Value |",
        "| --- | --- |",
        "| Revenue | 1.2M |",
        "",
      ].join("\n"),
    );
  });
});
