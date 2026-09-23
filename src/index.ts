/**
 * `.docx` → GitHub Flavored Markdown.
 *
 * A `.docx` states its structure outright — `w:pStyle` names the heading,
 * `w:numPr` marks the list item, `w:tbl` marks the table — so this package
 * translates rather than guesses. Two stages, both exported: `parseDocx` reads
 * OOXML into a small block AST, and `renderMarkdown` writes that AST out.
 * `convert` is the two of them in one call, which is what almost everyone
 * wants.
 *
 * No network, no filesystem, no DOM. The same code runs in Node, in a browser
 * and in a worker.
 */

import { parseDocx } from "./docx-parser.js";
import { renderMarkdown } from "./md-renderer.js";
import { ConversionError, DocxErrorCode, type ConvertOptions } from "./types.js";

export { parseDocx } from "./docx-parser.js";
export { renderMarkdown } from "./md-renderer.js";
export {
  escapeInline,
  renderStyledTokens,
  wrapEmphasis,
  type StyledToken,
} from "./inline.js";
export {
  ConversionError,
  DocxErrorCode,
  isConversionError,
  type ASTNode,
  type ASTNodeType,
  type ConversionProgress,
  type ConversionStage,
  type ConvertOptions,
  type ParseResult,
} from "./types.js";

/**
 * Convert a `.docx` to Markdown.
 *
 * Accepts anything that is already bytes — a Node `Buffer`, a `Uint8Array`, or
 * an `ArrayBuffer` from `File.arrayBuffer()` in a browser. Rejects with a
 * `ConversionError` whose `code` is one of `DocxErrorCode`.
 *
 * ```ts
 * import { readFile } from "node:fs/promises";
 * import { convert } from "@mypdf2md/docx2md";
 *
 * const markdown = await convert(await readFile("report.docx"));
 * ```
 */
export async function convert(
  input: Uint8Array | ArrayBuffer,
  options: ConvertOptions = {},
): Promise<string> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);

  if (bytes.byteLength === 0) {
    throw new ConversionError(DocxErrorCode.INVALID_DOCX, "the file is empty");
  }

  const { nodes } = await parseDocx(bytes, {
    signal: options.signal,
    // The parser thinks in absolute deadlines because it checks the clock
    // inside its walk; a caller thinks in durations. Translating here keeps
    // `performance.now()` out of the public API.
    deadline:
      options.timeoutMs === undefined
        ? undefined
        : performance.now() + options.timeoutMs,
    onProgress: options.onProgress,
  });

  return renderMarkdown(nodes);
}
