/**
 * DOCX (OOXML) → the shared `ASTNode[]`.
 *
 * The counterpart to `ast-builder.ts`, and a much simpler job. A PDF stores
 * positioned glyphs and nothing else, so structure has to be *inferred* from
 * geometry. A `.docx` states its structure outright — `w:pStyle` names the
 * heading, `w:numPr` marks the list item, `w:tbl` marks the table — so this
 * file translates rather than guesses, and its output is correspondingly more
 * reliable than the PDF path's.
 *
 * What it deliberately does not do:
 *
 * - **No new AST node types.** Everything lands in the six `ASTNodeType`s the
 *   Markdown renderer already handles. A format-specific node would mean a
 *   format-specific renderer, and then the same document saved two ways would
 *   convert two ways.
 * - **No image bytes.** Markdown cannot carry them and this package will not
 *   upload them, so figures are counted into `droppedImageCount` and reported
 *   rather than dropped in silence.
 * - **No full unzip.** A `.docx` is a ZIP, and a 40MB one is 40MB of JPEGs with
 *   a few hundred KB of XML in it. Only the four parts below are inflated.
 */

import { unzip, unzipSync, type Unzipped } from "fflate";

import { renderStyledTokens, type StyledToken } from "./inline.js";
import {
  type ASTNode,
  ConversionError,
  DocxErrorCode,
  type ConversionProgress,
  type ParseResult,
} from "./types.js";
import {
  isElement,
  parseXml as readXml,
  type XmlDocument,
  type XmlElement,
  type XmlNode,
} from "./xml.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const PART_DOCUMENT = "word/document.xml";
const PART_STYLES = "word/styles.xml";
const PART_NUMBERING = "word/numbering.xml";
const PART_RELS = "word/_rels/document.xml.rels";

const WANTED_PARTS = new Set([
  PART_DOCUMENT,
  PART_STYLES,
  PART_NUMBERING,
  PART_RELS,
]);

/** Checked against `w:rFonts/@w:ascii`. Word has no "this run is code" flag —
 *  a monospace face is the only signal the document carries. */
const MONOSPACE_FONTS =
  /^(courier|consolas|monaco|menlo|inconsolata|source code|fira ?(code|mono)|jetbrains ?mono|roboto mono|dejavu sans mono|liberation mono|lucida console|andale mono|sf mono|ibm plex mono|cascadia)/i;

const QUOTE_STYLES = /^(intense ?)?(quote|block ?text)$/i;
const CODE_STYLES = /^(html ?preformatted|source ?code|code|plain ?text|preformatted ?text)$/i;

export interface DocxParseOptions {
  signal?: AbortSignal;
  /** `performance.now()` value after which the parse must abort. */
  deadline?: number;
  onProgress?: (progress: ConversionProgress) => void;
}

/* ------------------------------------------------------------------ *
 * ZIP
 * ------------------------------------------------------------------ */

async function readParts(data: Uint8Array): Promise<Map<string, string>> {
  const filter = (file: { name: string }) => WANTED_PARTS.has(file.name);

  const decode = (files: Unzipped): Map<string, string> => {
    const decoder = new TextDecoder("utf-8");
    const parts = new Map<string, string>();
    for (const [name, bytes] of Object.entries(files)) {
      if (bytes.length > 0) parts.set(name, decoder.decode(bytes));
    }
    return parts;
  };

  try {
    return await new Promise<Map<string, string>>((resolve, reject) => {
      unzip(data, { filter }, (error, files) => {
        if (error) reject(error);
        else resolve(decode(files));
      });
    });
  } catch {
    // fflate's async path builds its worker from a Blob URL, which the Tauri
    // shell's CSP refuses and some hardened browser configurations block. The
    // synchronous path needs no worker; it costs a frame on a large file, which
    // is strictly better than failing.
    return decode(unzipSync(data, { filter }));
  }
}

/* ------------------------------------------------------------------ *
 * XML helpers
 * ------------------------------------------------------------------ */

function parseXml(text: string): XmlDocument {
  try {
    return readXml(text);
  } catch (error) {
    // The reader throws on malformed input rather than handing back a
    // document with a `<parsererror>` in it the way `DOMParser` does, so this
    // is the only thing standing between a corrupt file and a silently empty
    // conversion.
    throw new ConversionError(DocxErrorCode.INVALID_DOCX, undefined, error);
  }
}

/** Namespace-aware match, with a raw-prefix fallback for a document that
 *  declares no namespaces at all. */
function isNamed(node: XmlNode, ns: string, prefix: string, name: string): boolean {
  if (!isElement(node)) return false;
  if (node.namespaceURI === ns) return node.localName === name;
  return node.nodeName === `${prefix}:${name}`;
}

const isW = (node: XmlNode, name: string) => isNamed(node, W_NS, "w", name);
const isMc = (node: XmlNode, name: string) => isNamed(node, MC_NS, "mc", name);

function attrNS(el: XmlElement, ns: string, prefix: string, name: string): string | null {
  return el.getAttributeNS(ns, name) ?? el.getAttribute(`${prefix}:${name}`);
}

const wAttr = (el: XmlElement, name = "val") => attrNS(el, W_NS, "w", name);

function childW(el: XmlElement, name: string): XmlElement | null {
  for (const child of el.childNodes) {
    if (isW(child, name)) return child as XmlElement;
  }
  return null;
}

function childrenW(el: XmlElement, name: string): XmlElement[] {
  return el.childNodes.filter((child): child is XmlElement => isW(child, name));
}

function descendantsW(el: XmlElement, name: string): XmlElement[] {
  const byNs = el.getElementsByTagNameNS(W_NS, name);
  return byNs.length > 0 ? byNs : el.getElementsByTagName(`w:${name}`);
}

/**
 * OOXML boolean attributes are *toggles*: a bare `<w:b/>` turns bold on, and
 * `<w:b w:val="0"/>` turns it off again. Treating the element's presence as
 * "true" bolds every run that explicitly switches bold off.
 */
function toggleOn(el: XmlElement | null): boolean {
  if (!el) return false;
  const value = wAttr(el);
  return value === null || !/^(0|false|off)$/i.test(value);
}

/* ------------------------------------------------------------------ *
 * styles.xml / numbering.xml / rels
 * ------------------------------------------------------------------ */

interface StyleInfo {
  name: string;
  basedOn: string | null;
  outlineLevel: number | null;
}

function readStyles(doc: XmlDocument | null): Map<string, StyleInfo> {
  const styles = new Map<string, StyleInfo>();
  if (!doc?.documentElement) return styles;

  for (const style of descendantsW(doc.documentElement, "style")) {
    const id = wAttr(style, "styleId");
    if (!id) continue;

    const nameEl = childW(style, "name");
    const basedOnEl = childW(style, "basedOn");
    const pPr = childW(style, "pPr");
    const outlineEl = pPr ? childW(pPr, "outlineLvl") : null;
    const outlineRaw = outlineEl ? wAttr(outlineEl) : null;

    styles.set(id, {
      name: nameEl ? (wAttr(nameEl) ?? "") : "",
      basedOn: basedOnEl ? wAttr(basedOnEl) : null,
      outlineLevel: outlineRaw === null ? null : Number.parseInt(outlineRaw, 10),
    });
  }

  return styles;
}

function headingLevelFrom(label: string): number | null {
  const name = label.trim().toLowerCase();
  if (name === "") return null;
  // Word's own Title/Subtitle pair is a document's H1/H2 in every practical
  // sense, and mapping them to paragraphs loses the only structure the author
  // gave the top of the page.
  if (name === "title") return 1;
  if (name === "subtitle") return 2;

  const match =
    /^heading\s*([1-9])$/.exec(name) ??
    /^heading([1-9])$/.exec(name) ??
    /^h([1-6])$/.exec(name);
  return match?.[1] ? Math.min(6, Number.parseInt(match[1], 10)) : null;
}

/**
 * A style's heading rank, following `w:basedOn` until something says.
 *
 * Both the display name and the style id are tried, in that order, because
 * neither alone is reliable: a localized Word writes the name as
 * "Überschrift 1" while keeping the id "Heading1", and a document built from a
 * custom template does the reverse.
 */
function resolveHeadingLevel(
  styleId: string | null,
  styles: Map<string, StyleInfo>,
): number | null {
  let current = styleId;
  const seen = new Set<string>();

  while (current && !seen.has(current)) {
    seen.add(current);
    const info = styles.get(current);

    const level = headingLevelFrom(info?.name ?? "") ?? headingLevelFrom(current);
    if (level !== null) return level;

    // `w:outlineLvl` is 0-based and only consulted once the names have said
    // nothing — plenty of body styles carry an outline level without being
    // headings, so it is the weakest of the three signals, not the first.
    if (info && info.outlineLevel !== null && info.outlineLevel >= 0 && info.outlineLevel <= 5) {
      return info.outlineLevel + 1;
    }

    current = info?.basedOn ?? null;
  }

  return null;
}

function resolveStyleName(
  styleId: string | null,
  styles: Map<string, StyleInfo>,
): string {
  let current = styleId;
  const seen = new Set<string>();
  while (current && !seen.has(current)) {
    seen.add(current);
    const info = styles.get(current);
    const name = info?.name?.trim();
    if (name) return name;
    current = info?.basedOn ?? null;
  }
  return styleId ?? "";
}

/** `numId` → `ilvl` → is this level numbered rather than bulleted. */
type NumberingMap = Map<string, Map<number, boolean>>;

function readNumbering(doc: XmlDocument | null): NumberingMap {
  const numbering: NumberingMap = new Map();
  if (!doc?.documentElement) return numbering;

  const root = doc.documentElement;

  const abstractFormats = new Map<string, Map<number, boolean>>();
  for (const abstract of descendantsW(root, "abstractNum")) {
    const id = wAttr(abstract, "abstractNumId");
    if (!id) continue;

    const levels = new Map<number, boolean>();
    for (const lvl of childrenW(abstract, "lvl")) {
      const ilvlRaw = wAttr(lvl, "ilvl");
      const fmtEl = childW(lvl, "numFmt");
      const fmt = fmtEl ? (wAttr(fmtEl) ?? "") : "";
      const ilvl = ilvlRaw === null ? 0 : Number.parseInt(ilvlRaw, 10);
      // "none" means the level prints no marker at all. Rendering it as `1.`
      // invents a number the document does not have, so it stays a bullet.
      levels.set(ilvl, !/^(bullet|none)$/i.test(fmt));
    }
    abstractFormats.set(id, levels);
  }

  for (const num of descendantsW(root, "num")) {
    const numId = wAttr(num, "numId");
    if (!numId) continue;
    const abstractEl = childW(num, "abstractNumId");
    const abstractId = abstractEl ? wAttr(abstractEl) : null;
    const levels = abstractId ? abstractFormats.get(abstractId) : undefined;
    if (levels) numbering.set(numId, levels);
  }

  return numbering;
}

function readRelationships(doc: XmlDocument | null): Map<string, string> {
  const rels = new Map<string, string>();
  if (!doc?.documentElement) return rels;

  const root = doc.documentElement;
  const byNs = root.getElementsByTagNameNS(PKG_REL_NS, "Relationship");
  const entries = byNs.length > 0 ? byNs : root.getElementsByTagName("Relationship");

  for (const entry of entries) {
    const id = entry.getAttribute("Id");
    const target = entry.getAttribute("Target");
    if (id && target) rels.set(id, target);
  }

  return rels;
}

/* ------------------------------------------------------------------ *
 * Inline content
 * ------------------------------------------------------------------ */

interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  monospace?: boolean;
}

interface InlineContext {
  rels: Map<string, string>;
  /** Inside a table cell, a `w:br` must become a bare newline: `md-renderer`
   *  turns those into `<br>`, and a CommonMark hard break would instead break
   *  the pipe row it sits in. */
  inTableCell: boolean;
  /** Headings suppress emphasis — `# **Title**` is noise, not structure. */
  plainText: boolean;
  images: { count: number };
}

function readRunStyle(rPr: XmlElement | null): RunStyle {
  if (!rPr) return {};

  const fonts = childW(rPr, "rFonts");
  const ascii = fonts ? (wAttr(fonts, "ascii") ?? wAttr(fonts, "hAnsi") ?? "") : "";

  return {
    // `bCs` / `iCs` are the complex-script twins, and are what carry the
    // styling for Arabic, Hebrew and Devanagari runs.
    bold: toggleOn(childW(rPr, "b")) || toggleOn(childW(rPr, "bCs")),
    italic: toggleOn(childW(rPr, "i")) || toggleOn(childW(rPr, "iCs")),
    strike: toggleOn(childW(rPr, "strike")) || toggleOn(childW(rPr, "dstrike")),
    monospace: MONOSPACE_FONTS.test(ascii),
  };
}

function pushText(out: StyledToken[], text: string, style: RunStyle, ctx: InlineContext) {
  if (text === "") return;
  out.push(ctx.plainText ? { text } : { ...style, text });
}

/**
 * Walk a paragraph (or any container of runs) into styled tokens.
 *
 * The element names skipped and recursed into here are the whole difference
 * between a clean conversion and one littered with machine syntax — see the
 * comments at each branch.
 */
function walkInline(
  el: XmlElement,
  ctx: InlineContext,
  style: RunStyle,
  out: StyledToken[],
): void {
  for (const child of el.childNodes) {
    if (!isElement(child)) continue;

    // --- runs -------------------------------------------------------
    if (isW(child, "r")) {
      const runStyle = { ...style, ...readRunStyle(childW(child, "rPr")) };
      walkRun(child, ctx, runStyle, out);
      continue;
    }

    // --- hyperlinks -------------------------------------------------
    if (isW(child, "hyperlink")) {
      emitHyperlink(child, ctx, style, out);
      continue;
    }

    // --- tracked changes --------------------------------------------
    // An insertion is content the document now has; a deletion is content it
    // no longer has. Emitting both produces a document that contradicts
    // itself, and emitting the deletion is the worse of the two errors.
    if (isW(child, "ins")) {
      walkInline(child, ctx, style, out);
      continue;
    }
    if (isW(child, "del") || isW(child, "moveFrom")) continue;

    // --- wrappers that hold real content ----------------------------
    if (isW(child, "sdt")) {
      const content = childW(child, "sdtContent");
      if (content) walkInline(content, ctx, style, out);
      continue;
    }
    if (isW(child, "smartTag") || isW(child, "fldSimple") || isW(child, "moveTo")) {
      walkInline(child, ctx, style, out);
      continue;
    }

    if (handleShared(child, ctx, style, out)) continue;

    if (isW(child, "bookmarkStart") || isW(child, "bookmarkEnd") || isW(child, "proofErr")) {
      continue;
    }

    // Anything unrecognised that might still contain runs.
    if (child.childNodes.length > 0) walkInline(child, ctx, style, out);
  }
}

function walkRun(
  run: XmlElement,
  ctx: InlineContext,
  style: RunStyle,
  out: StyledToken[],
): void {
  for (const child of run.childNodes) {
    if (!isElement(child)) continue;

    if (isW(child, "t")) {
      // No trimming: `xml:space="preserve"` runs carry the spaces *between*
      // words, and a document is full of runs that are a single space.
      pushText(out, child.textContent, style, ctx);
      continue;
    }

    if (isW(child, "tab")) {
      pushText(out, "\t", style, ctx);
      continue;
    }

    if (isW(child, "br") || isW(child, "cr")) {
      out.push({ text: ctx.inTableCell ? "\n" : "\\\n", literal: true });
      continue;
    }

    if (isW(child, "noBreakHyphen")) {
      pushText(out, "-", style, ctx);
      continue;
    }

    // `w:instrText` is a field *instruction* — `PAGEREF _Toc123 \h`, `TOC \o
    // "1-3"`. It is machine syntax that Word never shows, and it is the single
    // largest source of garbage in a naive DOCX conversion. `w:delText` is the
    // text of a tracked deletion.
    if (isW(child, "instrText") || isW(child, "delText")) continue;
    if (isW(child, "softHyphen") || isW(child, "sym") || isW(child, "rPr")) continue;

    if (handleShared(child, ctx, style, out)) continue;

    if (child.childNodes.length > 0) walkInline(child, ctx, style, out);
  }
}

/**
 * Element handling both walkers need, in one place so they cannot drift.
 *
 * They did drift. Drawings were counted only when they sat directly inside a
 * `w:r`, and `mc:AlternateContent` was unwrapped only outside one — so a
 * figure written the way Word actually writes a shape (an AlternateContent
 * inside a run) hit neither branch and was reported as zero lost images.
 *
 * Returns whether the node was dealt with.
 */
function handleShared(
  child: XmlElement,
  ctx: InlineContext,
  style: RunStyle,
  out: StyledToken[],
): boolean {
  if (isW(child, "drawing") || isW(child, "pict") || isW(child, "object")) {
    emitDrawing(child, ctx, style, out);
    return true;
  }

  // `mc:Choice` is the modern representation and `mc:Fallback` the legacy one
  // *of the same object*. Walking both counts every such figure twice.
  if (isMc(child, "AlternateContent")) {
    const children = child.childNodes;
    const chosen =
      children.find((node) => isMc(node, "Choice")) ??
      children.find((node) => isMc(node, "Fallback"));
    if (chosen) walkInline(chosen as XmlElement, ctx, style, out);
    return true;
  }

  return false;
}

/**
 * A drawing is either a picture we cannot carry or a text box we must not lose.
 *
 * Word stores floating text in `w:drawing` too, wrapped in `w:txbxContent`.
 * Counting one of those as a dropped image would warn about losing a figure
 * while actually losing a paragraph, so the two cases are told apart here.
 */
function emitDrawing(
  el: XmlElement,
  ctx: InlineContext,
  style: RunStyle,
  out: StyledToken[],
): void {
  const textBoxes = descendantsW(el, "txbxContent");
  if (textBoxes.length === 0) {
    ctx.images.count += 1;
    return;
  }
  for (const box of textBoxes) {
    for (const paragraph of childrenW(box, "p")) {
      walkInline(paragraph, ctx, style, out);
    }
  }
}

function emitHyperlink(
  el: XmlElement,
  ctx: InlineContext,
  style: RunStyle,
  out: StyledToken[],
): void {
  const labelTokens: StyledToken[] = [];
  walkInline(el, ctx, style, labelTokens);
  const label = renderStyledTokens(labelTokens);
  if (label === "") return;

  const relId = attrNS(el, R_NS, "r", "id");
  const target = relId ? ctx.rels.get(relId) : null;

  // An internal anchor points at a bookmark in a document that no longer
  // exists once this is Markdown. The label is the content; the jump is not.
  if (!target) {
    out.push(...labelTokens);
    return;
  }

  // Angle brackets are the only way a Markdown destination can contain a space
  // or an unbalanced paren without the link falling apart.
  const destination = /[\s()]/.test(target) ? `<${target}>` : target;
  out.push({ text: `[${label}](${destination})`, literal: true });
}

/* ------------------------------------------------------------------ *
 * Block content
 * ------------------------------------------------------------------ */

interface ParserContext {
  styles: Map<string, StyleInfo>;
  numbering: NumberingMap;
  rels: Map<string, string>;
  images: { count: number };
  signal?: AbortSignal;
  deadline?: number;
}

function paragraphStyleId(paragraph: XmlElement): string | null {
  const pPr = childW(paragraph, "pPr");
  if (!pPr) return null;
  const style = childW(pPr, "pStyle");
  return style ? wAttr(style) : null;
}

function allRunsMonospaced(paragraph: XmlElement): boolean {
  const runs = descendantsW(paragraph, "r");
  const withText = runs.filter((run) => descendantsW(run, "t").length > 0);
  if (withText.length === 0) return false;
  return withText.every((run) => readRunStyle(childW(run, "rPr")).monospace);
}

function convertParagraph(paragraph: XmlElement, ctx: ParserContext): ASTNode | null {
  const styleId = paragraphStyleId(paragraph);
  const headingLevel = resolveHeadingLevel(styleId, ctx.styles);
  const styleName = resolveStyleName(styleId, ctx.styles);

  const pPr = childW(paragraph, "pPr");
  const numPr = pPr ? childW(pPr, "numPr") : null;

  const inline: InlineContext = {
    rels: ctx.rels,
    inTableCell: false,
    // A heading is already emphasised by being a heading.
    plainText: headingLevel !== null,
    images: ctx.images,
  };

  const tokens: StyledToken[] = [];
  walkInline(paragraph, inline, {}, tokens);
  const content = renderStyledTokens(tokens);
  if (content === "") return null;

  if (headingLevel !== null) {
    // Deliberately ahead of the numbering check: Word's numbered-heading
    // styles carry both a heading style and a `w:numPr`, and "1. Introduction"
    // as a list item loses the document's entire outline.
    return { type: "heading", level: headingLevel, content };
  }

  if (numPr) {
    const ilvlEl = childW(numPr, "ilvl");
    const numIdEl = childW(numPr, "numId");
    const ilvlRaw = ilvlEl ? wAttr(ilvlEl) : null;
    const numId = numIdEl ? wAttr(numIdEl) : null;

    // `numId="0"` is Word's way of *removing* numbering that a style applied.
    if (numId !== null && numId !== "0") {
      const depth = ilvlRaw === null ? 0 : Math.max(0, Number.parseInt(ilvlRaw, 10) || 0);
      const isOrdered = ctx.numbering.get(numId)?.get(depth) ?? false;
      return { type: "list_item", level: depth, content, isOrdered };
    }
  }

  if (QUOTE_STYLES.test(styleName)) {
    return { type: "blockquote", content };
  }

  if (CODE_STYLES.test(styleName) || allRunsMonospaced(paragraph)) {
    // The fence is added by the renderer, so the content must be the raw text —
    // `renderStyledTokens` has already wrapped monospace runs in backticks,
    // which inside a fence would be literal characters.
    return { type: "code_block", content: content.replace(/`/g, "") };
  }

  return { type: "paragraph", content };
}

/** Flatten a table cell to the text of its blocks, one per line. */
function cellText(cell: XmlElement, ctx: ParserContext): string {
  const lines: string[] = [];

  for (const node of cell.childNodes) {
    if (!isElement(node)) continue;

    if (isW(node, "p")) {
      const inline: InlineContext = {
        rels: ctx.rels,
        inTableCell: true,
        plainText: false,
        images: ctx.images,
      };
      const tokens: StyledToken[] = [];
      walkInline(node, inline, {}, tokens);
      const text = renderStyledTokens(tokens);
      if (text !== "") lines.push(text);
      continue;
    }

    // A nested table cannot be a table in Markdown — GFM has no such syntax —
    // so its cells become one line of text rather than vanishing.
    if (isW(node, "tbl")) {
      for (const row of childrenW(node, "tr")) {
        const inner = childrenW(row, "tc")
          .map((innerCell) => cellText(innerCell, ctx))
          .filter((text) => text !== "");
        if (inner.length > 0) lines.push(inner.join(" "));
      }
    }
  }

  return lines.join("\n");
}

function convertTable(table: XmlElement, ctx: ParserContext): ASTNode[] {
  const rows: string[][] = [];

  for (const row of childrenW(table, "tr")) {
    const cells: string[] = [];

    for (const cell of childrenW(row, "tc")) {
      const tcPr = childW(cell, "tcPr");
      const vMerge = tcPr ? childW(tcPr, "vMerge") : null;
      // A vertical-merge continuation is the *same* cell as the one above.
      // Repeating its text would duplicate content down the column.
      const isContinuation =
        vMerge !== null && (wAttr(vMerge) ?? "continue").toLowerCase() === "continue";

      const text = isContinuation ? "" : cellText(cell, ctx);
      cells.push(text);

      // A horizontally merged cell occupies several grid columns. Padding the
      // extra ones keeps every later cell in the row under the header it
      // actually belongs to.
      const gridSpanEl = tcPr ? childW(tcPr, "gridSpan") : null;
      const span = gridSpanEl ? Number.parseInt(wAttr(gridSpanEl) ?? "1", 10) : 1;
      for (let i = 1; i < (Number.isFinite(span) ? span : 1); i += 1) cells.push("");
    }

    if (cells.length > 0) rows.push(cells);
  }

  const populated = rows.filter((row) => row.some((cell) => cell.trim() !== ""));
  if (populated.length === 0) return [];

  const columns = Math.max(...populated.map((row) => row.length));

  // `md-renderer` drops any table narrower than two columns, because a
  // one-column pipe table is not valid GFM. Word uses exactly that shape for
  // callouts and sidebars, so the content is emitted as paragraphs instead of
  // disappearing.
  if (columns < 2) {
    return populated
      .flatMap((row) => row)
      .filter((cell) => cell.trim() !== "")
      .map((cell) => ({ type: "paragraph" as const, content: cell.replace(/\n/g, " ") }));
  }

  return [{ type: "table", content: "", rows: populated }];
}

function checkInterrupted(ctx: {
  signal?: AbortSignal;
  deadline?: number;
}): void {
  if (ctx.signal?.aborted) {
    throw new ConversionError(DocxErrorCode.CANCELLED);
  }
  if (ctx.deadline !== undefined && performance.now() > ctx.deadline) {
    throw new ConversionError(DocxErrorCode.TIMEOUT);
  }
}

function convertBody(body: XmlElement, ctx: ParserContext): ASTNode[] {
  const nodes: ASTNode[] = [];
  let sinceCheck = 0;

  const visit = (container: XmlElement): void => {
    for (const node of container.childNodes) {
      if (!isElement(node)) continue;

      if ((sinceCheck += 1) % 64 === 0) checkInterrupted(ctx);

      if (isW(node, "p")) {
        const converted = convertParagraph(node, ctx);
        if (converted) nodes.push(converted);
        continue;
      }

      if (isW(node, "tbl")) {
        nodes.push(...convertTable(node, ctx));
        continue;
      }

      // Content controls wrap real block content at the body level too.
      if (isW(node, "sdt")) {
        const content = childW(node, "sdtContent");
        if (content) visit(content);
        continue;
      }
    }
  };

  visit(body);
  return mergeAdjacent(nodes);
}

/**
 * Fuse runs of same-kind blocks that Markdown expresses as one.
 *
 * Word has no multi-line code block: five lines of code are five paragraphs in
 * a code style. Emitting five fences instead of one is the single most visible
 * difference between a good DOCX conversion and a bad one.
 */
function mergeAdjacent(nodes: ASTNode[]): ASTNode[] {
  const merged: ASTNode[] = [];

  for (const node of nodes) {
    const previous = merged[merged.length - 1];
    const mergeable = node.type === "code_block" || node.type === "blockquote";

    if (previous && mergeable && previous.type === node.type) {
      previous.content = `${previous.content}\n${node.content}`;
      continue;
    }

    merged.push({ ...node });
  }

  return merged;
}

/* ------------------------------------------------------------------ *
 * Entry point
 * ------------------------------------------------------------------ */

/**
 * "PK\x03\x04" — every OOXML file is a ZIP.
 *
 * D0 CF 11 E0 A1 B1 1A E1 — an OLE2 compound document. Both a Word 97–2003
 * `.doc` and a `.docx` encrypted at rest look like this, and neither is
 * readable here. Saying so beats letting the unzip fail with a generic error.
 */
function sniff(bytes: Uint8Array): "zip" | "ole" | "unknown" {
  if (
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  ) {
    return "zip";
  }

  if (
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1
  ) {
    return "ole";
  }

  return "unknown";
}

/**
 * Read a `.docx` into the shared AST.
 *
 * `convert` in `index.ts` is the one-call version of this; reach for
 * `parseDocx` when you want the blocks themselves — to truncate a document to
 * a budget, to count something, or to render it as anything but Markdown.
 */
export async function parseDocx(
  data: Uint8Array,
  options: DocxParseOptions = {},
): Promise<ParseResult> {
  const { signal, deadline, onProgress } = options;

  try {
    onProgress?.({ stage: "loading", ratio: 0.1 });

    if (sniff(data) === "ole") {
      throw new ConversionError(DocxErrorCode.OLE_CONTAINER);
    }

    let parts: Map<string, string>;
    try {
      parts = await readParts(data);
    } catch {
      // Not a readable ZIP: truncated, corrupt, or never a `.docx` at all.
      throw new ConversionError(DocxErrorCode.INVALID_DOCX);
    }

    const documentXml = parts.get(PART_DOCUMENT);
    if (!documentXml) {
      // A ZIP without `word/document.xml` is some other OOXML file — an .xlsx
      // or .pptx renamed, most often — or a .docx missing its main part.
      throw new ConversionError(DocxErrorCode.INVALID_DOCX);
    }

    checkInterrupted({ signal, deadline });
    onProgress?.({ stage: "extracting", ratio: 0.35 });

    const documentDoc = parseXml(documentXml);
    const stylesXml = parts.get(PART_STYLES);
    const numberingXml = parts.get(PART_NUMBERING);
    const relsXml = parts.get(PART_RELS);

    const ctx: ParserContext = {
      styles: readStyles(stylesXml ? parseXml(stylesXml) : null),
      numbering: readNumbering(numberingXml ? parseXml(numberingXml) : null),
      rels: readRelationships(relsXml ? parseXml(relsXml) : null),
      images: { count: 0 },
      signal,
      deadline,
    };

    const root = documentDoc.documentElement;
    const body = root ? childW(root, "body") : null;
    if (!body) throw new ConversionError(DocxErrorCode.INVALID_DOCX);

    onProgress?.({ stage: "building", ratio: 0.7 });
    const nodes = convertBody(body, ctx);

    onProgress?.({ stage: "done", ratio: 1 });

    return { nodes, droppedImageCount: ctx.images.count };
  } catch (error) {
    // Anything unrecognised is a file this package could not read. The cause is
    // attached rather than logged: a library has no business writing to a
    // caller's console.
    if (error instanceof ConversionError) throw error;
    throw new ConversionError(DocxErrorCode.INVALID_DOCX, undefined, error);
  }
}
