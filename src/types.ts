/**
 * The package's public contract.
 *
 * Deliberately narrow. The AST has six node types and the Markdown renderer
 * handles all six — a format-specific node would mean a format-specific
 * renderer, and then the same document saved two ways would convert two ways.
 */

/** The six blocks every `.docx` is reduced to. */
export type ASTNodeType =
  | "heading"
  | "paragraph"
  | "list_item"
  | "table"
  | "code_block"
  | "blockquote";

export interface ASTNode {
  type: ASTNodeType;
  /** Heading rank (1–6) for `heading`; nesting depth (0-based) for `list_item`. */
  level?: number;
  /** Inline Markdown for every node type except `table`, which uses `rows`. */
  content: string;
  /** Row-major cell matrix for `table`; the first row is the header. */
  rows?: string[][];
  /** `list_item` only — ordered (`1.`) vs unordered (`-`). */
  isOrdered?: boolean;
  /** `list_item` only — an explicit number to print instead of counting. */
  ordinal?: number;
}

export type ConversionStage = "loading" | "extracting" | "building" | "done";

export interface ConversionProgress {
  stage: ConversionStage;
  /** Overall completion, 0–1, for driving a progress bar. */
  ratio: number;
}

/**
 * Every code this package can throw.
 *
 * `ConversionError` is intentionally generic over its code so that a host
 * application can widen the union with codes of its own — its own file-format
 * or network failures — while still catching a single class. See the README.
 */
export const DocxErrorCode = {
  /** Not a readable ZIP, not an OOXML wordprocessing document, or the XML
   *  inside it is malformed. */
  INVALID_DOCX: "INVALID_DOCX",
  /**
   * An OLE2 compound document rather than a ZIP.
   *
   * Either a Word 97–2003 `.doc` or a `.docx` encrypted at rest. Only a
   * filename can tell the two apart and this package is handed bytes, so it
   * reports what it can actually see.
   */
  OLE_CONTAINER: "OLE_CONTAINER",
  TIMEOUT: "TIMEOUT",
  CANCELLED: "CANCELLED",
} as const;

export type DocxErrorCode = (typeof DocxErrorCode)[keyof typeof DocxErrorCode];

/**
 * Every rejection from `convert` and `parseDocx` is one of these, so a caller
 * can map `code` straight onto a message of its own.
 */
export class ConversionError<Code extends string = string> extends Error {
  readonly code: Code;

  constructor(code: Code, message?: string, cause?: unknown) {
    super(message ?? code, { cause });
    this.name = "ConversionError";
    this.code = code;
  }
}

export function isConversionError(value: unknown): value is ConversionError {
  return value instanceof ConversionError;
}

export interface ConvertOptions {
  /** Caller-controlled cancellation. Rejects with `CANCELLED`. */
  signal?: AbortSignal;
  /** Abort the parse after this many milliseconds. Rejects with `TIMEOUT`.
   *  Omit for no limit. Both this and `signal` are checked every 64 nodes, so
   *  a parse stops promptly without paying for a clock read per element. */
  timeoutMs?: number;
  onProgress?: (progress: ConversionProgress) => void;
}

export interface ParseResult {
  nodes: ASTNode[];
  /**
   * Figures the conversion could not carry.
   *
   * Markdown cannot hold image bytes and this package will not upload them, so
   * pictures are counted rather than dropped in silence. A UI that converts a
   * document with figures in it should say so.
   */
  droppedImageCount: number;
}
