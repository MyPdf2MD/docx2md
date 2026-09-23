/**
 * AST → GitHub Flavored Markdown.
 *
 * Inline emphasis and escaping already happened in the AST builder, so this
 * stage only assembles block syntax: heading hashes, list markers with their
 * numbering, fenced code, blockquote prefixes and pipe tables.
 */

import type { ASTNode } from "./types.js";

const MAX_HEADING_LEVEL = 6;

export function renderMarkdown(nodes: ASTNode[]): string {
  const blocks: string[] = [];
  /** Ordered-list counters, one per nesting depth. */
  let counters: number[] = [];
  let previous: ASTNode | null = null;

  for (const node of nodes) {
    if (node.type !== "list_item") counters = [];

    const rendered = renderNode(node, counters);
    if (rendered === "") continue;

    // Items of the same list stay on consecutive lines; everything else gets a
    // blank line so Markdown treats it as a new block.
    if (blocks.length > 0) {
      blocks.push(continuesList(previous, node) ? "\n" : "\n\n");
    }
    blocks.push(rendered);
    previous = node;
  }

  const markdown = blocks.join("").replace(/\n{3,}/g, "\n\n").trim();
  return markdown === "" ? "" : `${markdown}\n`;
}

/**
 * Whether two adjacent items belong to the same list.
 *
 * A change of marker kind starts a new list, so it needs the blank line that
 * every other block boundary gets. Without it a bulleted list followed by a
 * numbered one renders as `- item` and `1. item` on consecutive lines — which
 * CommonMark does read as two lists, but which is unreadable as source and
 * which some stricter renderers fold into one. DOCX makes this common: Word
 * documents switch between bullets and numbers constantly, and unlike the PDF
 * path the parser knows for certain which is which.
 */
function continuesList(previous: ASTNode | null, node: ASTNode): boolean {
  return (
    previous !== null &&
    previous.type === "list_item" &&
    node.type === "list_item" &&
    !!previous.isOrdered === !!node.isOrdered
  );
}

function renderNode(node: ASTNode, counters: number[]): string {
  switch (node.type) {
    case "heading":
      return renderHeading(node);
    case "list_item":
      return renderListItem(node, counters);
    case "table":
      return renderTable(node);
    case "code_block":
      return renderCodeBlock(node);
    case "blockquote":
      return renderBlockquote(node);
    case "paragraph":
    default:
      return escapeBlockStart(node.content.trim());
  }
}

function renderHeading(node: ASTNode): string {
  const content = node.content.trim();
  if (content === "") return "";
  const level = Math.min(MAX_HEADING_LEVEL, Math.max(1, node.level ?? 1));
  return `${"#".repeat(level)} ${content}`;
}

function renderListItem(node: ASTNode, counters: number[]): string {
  const content = node.content.trim();
  if (content === "") return "";

  const depth = Math.max(0, node.level ?? 0);
  counters.length = depth + 1;
  for (let i = 0; i <= depth; i += 1) counters[i] ??= 0;

  const indent = "  ".repeat(depth);
  if (!node.isOrdered) return `${indent}- ${content}`;

  // Trust the number printed in the PDF when there is one; otherwise continue
  // our own sequence for this depth.
  const ordinal = node.ordinal ?? (counters[depth] ?? 0) + 1;
  counters[depth] = ordinal;
  return `${indent}${ordinal}. ${content}`;
}

function renderCodeBlock(node: ASTNode): string {
  const content = node.content.replace(/\s+$/, "");
  if (content.trim() === "") return "";

  // Widen the fence if the code itself contains backtick runs.
  const longestRun = Math.max(
    0,
    ...[...content.matchAll(/`+/g)].map((match) => match[0].length),
  );
  const fence = "`".repeat(Math.max(3, longestRun + 1));
  return `${fence}\n${content}\n${fence}`;
}

function renderBlockquote(node: ASTNode): string {
  const content = node.content.trim();
  if (content === "") return "";
  return content
    .split("\n")
    .map((line) => (line === "" ? ">" : `> ${line}`))
    .join("\n");
}

function renderTable(node: ASTNode): string {
  const rows = (node.rows ?? []).filter((row) =>
    row.some((cell) => cell.trim() !== ""),
  );
  if (rows.length === 0) return "";

  const columns = Math.max(...rows.map((row) => row.length));
  if (columns < 2) return "";

  const normalise = (row: string[]) =>
    Array.from({ length: columns }, (_, index) =>
      (row[index] ?? "").replace(/\s*\n\s*/g, "<br>").trim(),
    );

  // `rows` is non-empty — the filter above would have returned early.
  const [header = [], ...body] = rows;
  const lines = [
    `| ${normalise(header).join(" | ")} |`,
    `| ${Array.from({ length: columns }, () => "---").join(" | ")} |`,
    ...body.map((row) => `| ${normalise(row).join(" | ")} |`),
  ];

  return lines.join("\n");
}

/**
 * Stop body text that happens to start with `#`, `>`, `- ` or `1.` from being
 * re-parsed as a heading, quote or list.
 */
function escapeBlockStart(content: string): string {
  return content.replace(/^(#{1,6}\s|>\s?|[-+*]\s|\d{1,3}[.)]\s)/, "\\$1");
}
