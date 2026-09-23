/**
 * Styled text runs → inline Markdown.
 *
 * Deliberately not tied to `.docx`. How a front end *discovers* styling varies
 * — this package reads it straight out of `w:rPr`, while a PDF reader has to
 * infer it from font names and glyph metrics — but once something has decided
 * that a stretch of text is bold, the rules for turning that into `**bold**`
 * are identical. Keeping them in one place is what stops the same document,
 * saved two ways, converting two ways.
 *
 * Exported because mypdf2md.com relies on exactly that: its PDF front end
 * shares this module with the DOCX one.
 */

export interface StyledToken {
  text: string;
  bold?: boolean;
  italic?: boolean;
  monospace?: boolean;
  /**
   * GFM `~~strikethrough~~`.
   *
   * Only the DOCX front end sets this today, because only it can know: Word
   * says so outright in `w:strike`, whereas a PDF stores a struck-through word
   * as the word plus an unrelated line-drawing operation somewhere on the page,
   * with nothing tying the two together.
   */
  strike?: boolean;
  /**
   * Emitted verbatim: already-valid Markdown that must not be escaped or
   * merged into a neighbour. Links and hard line breaks arrive this way,
   * because escaping `[`, `]` or `\` inside them would defeat the point.
   */
  literal?: boolean;
}

const ESCAPE_RE = /([\\`*_[\]<>|])/g;

export function escapeInline(text: string): string {
  return text.replace(ESCAPE_RE, "\\$1");
}

export function wrapEmphasis(core: string, token: StyledToken): string {
  const emphasised = applyEmphasis(core, token);
  // Outermost, so `~~**x**~~` rather than `**~~x~~**`. Both parse, but this
  // order survives a round trip through editors that normalise emphasis.
  return token.strike ? `~~${emphasised}~~` : emphasised;
}

function applyEmphasis(core: string, token: StyledToken): string {
  if (token.monospace) return `\`${core.replace(/`/g, "'")}\``;

  const body = escapeInline(core);
  if (token.bold && token.italic) return `***${body}***`;
  if (token.bold) return `**${body}**`;
  if (token.italic) return `*${body}*`;
  return body;
}

/**
 * Merge adjacent same-styled tokens, then wrap each in its emphasis markers.
 *
 * Merging first is what stops `**bo****ld**` — two bold runs that the source
 * document happened to split mid-word. Whitespace is kept outside the
 * delimiters because CommonMark will not parse `** bold **` as emphasis at all.
 */
export function renderStyledTokens(tokens: StyledToken[]): string {
  const merged: StyledToken[] = [];
  for (const token of tokens) {
    if (token.text === "") continue;
    const last = merged[merged.length - 1];
    if (
      last &&
      !last.literal &&
      !token.literal &&
      !!last.bold === !!token.bold &&
      !!last.italic === !!token.italic &&
      !!last.monospace === !!token.monospace &&
      !!last.strike === !!token.strike
    ) {
      last.text += token.text;
    } else {
      merged.push({ ...token });
    }
  }

  let out = "";
  for (const token of merged) {
    if (token.literal) {
      out += token.text;
      continue;
    }
    const match = /^(\s*)([\s\S]*?)(\s*)$/.exec(token.text);
    const [, lead = "", core = "", trail = ""] = match ?? [];
    if (core === "") {
      out += token.text;
      continue;
    }
    out += lead + wrapEmphasis(core, token) + trail;
  }

  // `[ \t]` rather than `\s`, so hard line breaks survive the collapse.
  return out.replace(/[ \t]+/g, " ").trim();
}
