# @mypdf2md/docx2md

Convert `.docx` (OOXML) to GitHub Flavored Markdown.

No DOM, no network, no filesystem, no child processes, no LibreOffice. One
dependency ([fflate](https://github.com/101arrowz/fflate), for the ZIP). The
same code runs in Node, in a browser and in a Web Worker.

A `.docx` states its structure outright — `w:pStyle` names the heading,
`w:numPr` marks the list item, `w:tbl` marks the table — so this library
translates rather than guesses. That is the whole design: nothing here infers
structure from font sizes or indentation.

## Install

Not yet published to npm. Install from the repository:

```sh
npm install github:MyPdf2MD/docx2md
```

npm builds the package as it installs it, so nothing further is needed.

To work on it, clone instead:

```sh
git clone https://github.com/MyPdf2MD/docx2md.git
cd docx2md
npm install
npm run build
```

`npm install` already runs the build; the explicit `npm run build` is there
for rebuilding after a change. `npm test` runs the suite.

## Use

```ts
import { readFile } from "node:fs/promises";
import { convert } from "@mypdf2md/docx2md";

const markdown = await convert(await readFile("report.docx"));
```

In a browser, hand it the bytes from a file input:

```ts
const markdown = await convert(await file.arrayBuffer());
```

`convert` accepts a `Uint8Array`, a Node `Buffer` or an `ArrayBuffer`, and
returns a string.

## CLI

From a clone, after `npm install`:

```sh
node dist/cli.js report.docx              # to stdout
node dist/cli.js report.docx -o report.md # to a file
cat report.docx | node dist/cli.js        # from stdin
```

Exits `0` on success, `1` on a document it cannot read, `2` on bad usage.

## Before and after

The source `word/document.xml`, as Word writes it — eight runs, an `rPr` on two
of them, and a hyperlink that points at a relationship ID rather than a URL:

```xml
<w:p><w:r><w:t xml:space="preserve">The exporter now writes </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>UTF-8</w:t></w:r><w:r><w:t xml:space="preserve"> by default. See </w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>the migration guide</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> before upgrading. The old </w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t>--latin1</w:t></w:r><w:r><w:t xml:space="preserve"> flag is gone.</w:t></w:r></w:p>
```

and the whole document out the other side, verbatim:

````md
# Release Notes

## What changed

The exporter now writes **UTF-8** by default. See [the migration guide](https://example.com/migrating) before upgrading. The old ~~--latin1~~ flag is gone.

- Encoding is detected, not guessed
- Tables keep their headers
  - Including merged cells

## Compatibility

| Version | Status |
| --- | --- |
| 2.x | Supported |
| 1.x | Security fixes only |

> Upgrading is not reversible. Take a backup.

```
npm install exporter@3
exporter --check
```

Thanks to everyone who tested the beta.
````

Three things in that output are worth pointing at, because they are where naive
converters go wrong:

- The eight runs became one sentence. Word splits a sentence into runs wherever
  formatting, spell-check state or a stray edit says so; emitting each one
  separately produces `**bo****ld**`.
- `r:id="rId1"` was resolved through `word/_rels/document.xml.rels` to a real
  URL.
- The document contained a `PAGEREF _Toc4711 \h` field instruction and an
  image. The instruction is gone (Word never shows it; it is the single largest
  source of garbage in a naive conversion) and the image was counted rather than
  emitted as a broken `![]()`.

## What it handles

| | |
|---|---|
| **Headings** | `w:pStyle` by name, then by style ID, then `w:outlineLvl`, following `w:basedOn` chains. Localized Word included — `Überschrift 1` with the ID `Heading1` resolves either way. `Title` and `Subtitle` map to H1 and H2. |
| **Lists** | Ordered vs bulleted resolved properly, through `numId` → `abstractNumId` → the level's `numFmt`. Nesting via `w:ilvl`. `numId="0"` correctly *removes* numbering a style applied. A numbered heading stays a heading. |
| **Tables** | GFM pipe tables. `gridSpan` pads the columns it spans; a `vMerge` continuation is emptied rather than repeating the cell above; a nested table flattens to text (GFM has no nested tables); a one-column table becomes paragraphs (GFM has no one-column table). |
| **Emphasis** | `w:b`, `w:i`, `w:strike`, `w:dstrike`, plus the complex-script twins `w:bCs`/`w:iCs` that carry styling for Arabic, Hebrew and Devanagari. Toggle semantics are respected: `<w:b w:val="0"/>` turns bold *off*. |
| **Links** | `w:hyperlink` resolved through the relationship part. Destinations containing spaces or parens are wrapped in `<>`. An internal anchor degrades to its label rather than producing a dead link. |
| **Code** | Fenced blocks from code styles, and inline code inferred from a monospace `w:rFonts`. Consecutive code paragraphs merge into one fence, because Word has no multi-line code block. |
| **Blockquotes** | From Word's `Quote`, `Intense Quote` and `Block Text` styles. Consecutive ones merge. |
| **Tracked changes** | `w:ins` content is kept; `w:del`, `w:delText` and `w:moveFrom` are dropped. The result is the document as it now stands. |
| **Wrappers** | Content controls (`w:sdt`), smart tags, `mc:AlternateContent` (the Choice *or* the Fallback, never both — otherwise every shape counts twice), and text boxes, whose contents are recovered rather than lost. |
| **Cancellation** | An `AbortSignal` and a `timeoutMs`, checked every 64 nodes. |

It also inflates only the four XML parts it needs, so a 40 MB `.docx` that is
40 MB of JPEGs costs a few hundred KB of work.

## What it does not do

Listed because a converter that fails silently is worse than one that says so.

- **Images.** Markdown cannot carry bytes and this library will not upload them.
  Pictures are counted — `parseDocx` returns `droppedImageCount` — and never
  emitted. There is no extraction, no base64, no output directory.
- **Footnotes and endnotes.** `word/footnotes.xml` is not read. The reference
  marker and its text are both lost.
- **Headers and footers.** Not read.
- **Comments.** Not read.
- **Equations.** OMML (`m:oMath`) is not understood. There is no LaTeX output;
  loose glyph text from an equation may appear unstructured.
- **Encrypted documents.** A password-protected `.docx` is an OLE container
  rather than a ZIP. Rejected with `OLE_CONTAINER`.
- **Legacy `.doc`.** Word 97–2003 is a different format entirely, also OLE. Same
  rejection. Only the filename distinguishes the two cases, and this library is
  handed bytes, so it reports what it can see.
- **Page numbers, page counts, page breaks.** Word repaginates at render time
  and stores no pagination, so any number here would be invented.
- **Colour, highlighting, underline, alignment, font sizes, columns.** No
  Markdown equivalent.
- **Bookmarks.**
- **`.docx` output.** This converts one way.

## API

```ts
convert(input: Uint8Array | ArrayBuffer, options?: ConvertOptions): Promise<string>
```

`ConvertOptions`: `signal?: AbortSignal`, `timeoutMs?: number`,
`onProgress?: (p: ConversionProgress) => void`.

For anything other than Markdown, take the two stages separately. `parseDocx`
returns a small block AST — six node types, no format-specific nodes — which you
can truncate, count, or render yourself:

```ts
import { parseDocx, renderMarkdown } from "@mypdf2md/docx2md";

const { nodes, droppedImageCount } = await parseDocx(bytes);
const markdown = renderMarkdown(nodes.slice(0, 50));
```

### Errors

Every rejection is a `ConversionError` carrying a `code` from `DocxErrorCode`:
`INVALID_DOCX`, `OLE_CONTAINER`, `TIMEOUT` or `CANCELLED`. Nothing throws a bare
`Error`, so a caller can always map the failure onto a message of its own.

```ts
import { convert, isConversionError } from "@mypdf2md/docx2md";

try {
  await convert(bytes);
} catch (error) {
  if (isConversionError(error)) console.error(error.code);
}
```

`ConversionError` is generic over its code — `ConversionError<MyCodes>` — so an
application can widen the union with codes of its own while still catching a
single class.

## Requirements

Node 18 or newer, or any browser with `TextDecoder` and `Promise`. TypeScript
types are included; there is no separate `@types` package.

## Why it exists

It is the DOCX half of the converter at
[mypdf2md.com](https://mypdf2md.com), which runs entirely in the browser — the
same code, published rather than forked, so the two cannot drift.

## License

MIT
