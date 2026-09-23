/**
 * Synthetic `.docx` builders for the parser specs.
 *
 * Test-only, but a module rather than a helper inside the spec file: the
 * fixtures ARE the specification of what OOXML this package claims to read, and
 * writing them out as real markup keeps that claim inspectable. A binary
 * fixture checked into the repo would be neither readable in a diff nor
 * editable when a case needs a variation.
 */

import { zipSync, strToU8 } from "fflate";

export const W_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const MC_NS =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const PKG_REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

export interface DocxParts {
  /** Inner XML of `w:body`. */
  body: string;
  /** Inner XML of `w:styles`. */
  styles?: string;
  /** Inner XML of `w:numbering`. */
  numbering?: string;
  /** Inner XML of the relationships root. */
  rels?: string;
  /** Extra entries, for the "a ZIP but not a document" cases. */
  extra?: Record<string, string>;
  /** Drop `word/document.xml` entirely. */
  omitDocument?: boolean;
}

export function makeDocx({
  body,
  styles,
  numbering,
  rels,
  extra,
  omitDocument = false,
}: DocxParts): Uint8Array {
  const files: Record<string, Uint8Array> = {};

  if (!omitDocument) {
    files["word/document.xml"] = strToU8(
      `${DECLARATION}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:mc="${MC_NS}"><w:body>${body}</w:body></w:document>`,
    );
  }

  if (styles !== undefined) {
    files["word/styles.xml"] = strToU8(
      `${DECLARATION}<w:styles xmlns:w="${W_NS}">${styles}</w:styles>`,
    );
  }

  if (numbering !== undefined) {
    files["word/numbering.xml"] = strToU8(
      `${DECLARATION}<w:numbering xmlns:w="${W_NS}">${numbering}</w:numbering>`,
    );
  }

  if (rels !== undefined) {
    files["word/_rels/document.xml.rels"] = strToU8(
      `${DECLARATION}<Relationships xmlns="${PKG_REL_NS}">${rels}</Relationships>`,
    );
  }

  for (const [name, content] of Object.entries(extra ?? {})) {
    files[name] = strToU8(content);
  }

  return zipSync(files);
}

/** A run of plain text, optionally styled. */
export function run(
  text: string,
  properties = "",
  { preserveSpace = true } = {},
): string {
  const rPr = properties === "" ? "" : `<w:rPr>${properties}</w:rPr>`;
  const space = preserveSpace ? ' xml:space="preserve"' : "";
  return `<w:r>${rPr}<w:t${space}>${escapeXml(text)}</w:t></w:r>`;
}

/** A paragraph, optionally carrying `w:pPr` children. */
export function paragraph(children: string, properties = ""): string {
  const pPr = properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`;
  return `<w:p>${pPr}${children}</w:p>`;
}

export function pStyle(id: string): string {
  return `<w:pStyle w:val="${id}"/>`;
}

export function numPr(numId: string, ilvl = 0): string {
  return `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
}

export function styleDefinition(
  id: string,
  { name, basedOn, outlineLevel }: {
    name?: string;
    basedOn?: string;
    outlineLevel?: number;
  } = {},
): string {
  const parts = [
    name === undefined ? "" : `<w:name w:val="${escapeXml(name)}"/>`,
    basedOn === undefined ? "" : `<w:basedOn w:val="${basedOn}"/>`,
    outlineLevel === undefined
      ? ""
      : `<w:pPr><w:outlineLvl w:val="${outlineLevel}"/></w:pPr>`,
  ];
  return `<w:style w:type="paragraph" w:styleId="${id}">${parts.join("")}</w:style>`;
}

/**
 * A numbering definition, mapping one `numId` to per-level formats.
 *
 * `formats[0]` is `ilvl` 0. Anything other than "bullet" or "none" makes the
 * level ordered, which mirrors what `numbering.xml` actually means.
 */
export function numberingDefinition(
  numId: string,
  abstractId: string,
  formats: string[],
): string {
  const levels = formats
    .map(
      (format, index) =>
        `<w:lvl w:ilvl="${index}"><w:numFmt w:val="${format}"/></w:lvl>`,
    )
    .join("");
  return (
    `<w:abstractNum w:abstractNumId="${abstractId}">${levels}</w:abstractNum>` +
    `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`
  );
}

export function cell(children: string, properties = ""): string {
  const tcPr = properties === "" ? "" : `<w:tcPr>${properties}</w:tcPr>`;
  return `<w:tc>${tcPr}${children}</w:tc>`;
}

export function row(cells: string): string {
  return `<w:tr>${cells}</w:tr>`;
}

export function table(rows: string): string {
  return `<w:tbl>${rows}</w:tbl>`;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
