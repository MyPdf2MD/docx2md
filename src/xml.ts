/**
 * A small, namespace-aware XML reader.
 *
 * The parser used to call the browser's `DOMParser`, which meant it ran in a
 * browser and nowhere else — Node has no such global, and the test suite had to
 * boot jsdom to make up for it. Swapping in a DOM implementation for Node would
 * have shipped one namespace implementation and tested against another, which
 * for a parser whose every element match is namespace-aware is precisely the
 * wrong trade.
 *
 * So this reads the subset of XML that OOXML actually uses, and exposes the
 * handful of DOM-shaped accessors `docx-parser.ts` needs. One code path in
 * every runtime, no dependencies.
 *
 * Deliberately not implemented, because no `.docx` part contains them:
 * DTDs and entity declarations, processing instructions beyond the XML
 * declaration, mixed-content mutation, or anything writable. The five
 * predefined entities and numeric character references are supported;
 * an unknown named entity is left as written rather than throwing, which is
 * what every mainstream parser does with `&nbsp;` in an XML document.
 */

export class XmlError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "XmlError";
  }
}

/** Mirrors the DOM's numbering, so the shape stays familiar. `isolatedModules`
 *  rules out a `const enum`, and a frozen object reads the same at the call
 *  site anyway. */
export const XmlNodeType = { Element: 1, Text: 3 } as const;

export interface XmlText {
  readonly nodeType: typeof XmlNodeType.Text;
  readonly text: string;
}

export type XmlNode = XmlElement | XmlText;

export function isElement(node: XmlNode): node is XmlElement {
  return node.nodeType === XmlNodeType.Element;
}

interface RawAttr {
  /** As written, including any prefix: `w:val`, `xmlns:w`, `Id`. */
  qualified: string;
  /** Resolved namespace URI, or null for an unprefixed attribute. Per the
   *  spec an unprefixed attribute is in *no* namespace, never the element's. */
  ns: string | null;
  local: string;
  value: string;
}

export class XmlElement {
  readonly nodeType = XmlNodeType.Element;

  constructor(
    /** As written, including the prefix: `w:p`. */
    readonly nodeName: string,
    readonly localName: string,
    readonly namespaceURI: string | null,
    private readonly attrs: RawAttr[],
    readonly childNodes: XmlNode[],
  ) {}

  getAttribute(qualified: string): string | null {
    for (const attr of this.attrs) {
      if (attr.qualified === qualified) return attr.value;
    }
    return null;
  }

  getAttributeNS(ns: string, local: string): string | null {
    for (const attr of this.attrs) {
      if (attr.ns === ns && attr.local === local) return attr.value;
    }
    return null;
  }

  /** Element children only, in document order. */
  get children(): XmlElement[] {
    return this.childNodes.filter(isElement);
  }

  /** Descendants matching a namespace + local name, in document order.
   *  Self is excluded, matching the DOM. */
  getElementsByTagNameNS(ns: string, local: string): XmlElement[] {
    const found: XmlElement[] = [];
    const walk = (parent: XmlElement): void => {
      for (const child of parent.childNodes) {
        if (!isElement(child)) continue;
        if (child.namespaceURI === ns && child.localName === local) found.push(child);
        walk(child);
      }
    };
    walk(this);
    return found;
  }

  /** The prefix-matching counterpart, for a document that declares no
   *  namespaces. See `isNamed` in `docx-parser.ts` for why both exist. */
  getElementsByTagName(qualified: string): XmlElement[] {
    const found: XmlElement[] = [];
    const walk = (parent: XmlElement): void => {
      for (const child of parent.childNodes) {
        if (!isElement(child)) continue;
        if (child.nodeName === qualified) found.push(child);
        walk(child);
      }
    };
    walk(this);
    return found;
  }

  get textContent(): string {
    let out = "";
    const walk = (node: XmlNode): void => {
      if (!isElement(node)) {
        out += node.text;
        return;
      }
      for (const child of node.childNodes) walk(child);
    };
    walk(this);
    return out;
  }
}

export interface XmlDocument {
  readonly documentElement: XmlElement | null;
}

/* ------------------------------------------------------------------ *
 * Entities
 * ------------------------------------------------------------------ */

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

const ENTITY_RE = /&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z][A-Za-z0-9._-]*);/g;

function decodeEntities(text: string): string {
  // The fast path matters: `w:t` bodies are the bulk of a document and most
  // carry no entity at all.
  if (!text.includes("&")) return text;

  return text.replace(ENTITY_RE, (whole, body: string) => {
    if (body.charCodeAt(0) === 35 /* # */) {
      const code =
        body[1] === "x" || body[1] === "X"
          ? Number.parseInt(body.slice(2), 16)
          : Number.parseInt(body.slice(1), 10);
      // Lone surrogates and out-of-range code points would make
      // `fromCodePoint` throw on a file we can still mostly read.
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return whole;
      if (code >= 0xd800 && code <= 0xdfff) return whole;
      return String.fromCodePoint(code);
    }
    return NAMED_ENTITIES[body] ?? whole;
  });
}

/* ------------------------------------------------------------------ *
 * Parser
 * ------------------------------------------------------------------ */

const NAME_START = /[A-Za-z_:]/;
const NAME_CHAR = /[-A-Za-z0-9_:.]/;

/** Namespace prefix → URI, innermost declaration winning. */
type NamespaceScope = ReadonlyMap<string, string>;

const XML_NS = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

const ROOT_SCOPE: NamespaceScope = new Map([["xml", XML_NS]]);

class Reader {
  private pos = 0;

  constructor(private readonly src: string) {}

  get done(): boolean {
    return this.pos >= this.src.length;
  }

  peek(offset = 0): string {
    return this.src[this.pos + offset] ?? "";
  }

  startsWith(text: string): boolean {
    return this.src.startsWith(text, this.pos);
  }

  advance(count = 1): void {
    this.pos += count;
  }

  /** Consume up to and including `terminator`; returns the text before it. */
  takeUntil(terminator: string): string {
    const at = this.src.indexOf(terminator, this.pos);
    if (at === -1) throw new XmlError(`unterminated ${terminator} at ${this.pos}`);
    const text = this.src.slice(this.pos, at);
    this.pos = at + terminator.length;
    return text;
  }

  /** Text up to the next `<`, or to the end. */
  takeText(): string {
    const at = this.src.indexOf("<", this.pos);
    const end = at === -1 ? this.src.length : at;
    const text = this.src.slice(this.pos, end);
    this.pos = end;
    return text;
  }

  skipSpace(): void {
    while (this.pos < this.src.length) {
      const ch = this.src.charCodeAt(this.pos);
      // space, tab, LF, CR
      if (ch === 32 || ch === 9 || ch === 10 || ch === 13) this.pos += 1;
      else break;
    }
  }

  takeName(): string {
    const start = this.pos;
    if (!NAME_START.test(this.peek())) {
      throw new XmlError(`expected a name at ${this.pos}`);
    }
    this.pos += 1;
    while (this.pos < this.src.length && NAME_CHAR.test(this.src[this.pos]!)) {
      this.pos += 1;
    }
    return this.src.slice(start, this.pos);
  }

  expect(ch: string): void {
    if (this.peek() !== ch) {
      throw new XmlError(`expected ${ch} at ${this.pos}, found ${this.peek() || "EOF"}`);
    }
    this.pos += 1;
  }
}

interface PendingAttr {
  qualified: string;
  prefix: string | null;
  local: string;
  value: string;
}

function readAttributes(reader: Reader): PendingAttr[] {
  const attrs: PendingAttr[] = [];

  for (;;) {
    reader.skipSpace();
    const ch = reader.peek();
    if (ch === "" || ch === ">" || ch === "/") return attrs;

    const qualified = reader.takeName();
    reader.skipSpace();
    reader.expect("=");
    reader.skipSpace();

    const quote = reader.peek();
    if (quote !== '"' && quote !== "'") {
      throw new XmlError(`unquoted attribute value for ${qualified}`);
    }
    reader.advance();
    const raw = reader.takeUntil(quote);

    const colon = qualified.indexOf(":");
    attrs.push({
      qualified,
      prefix: colon === -1 ? null : qualified.slice(0, colon),
      local: colon === -1 ? qualified : qualified.slice(colon + 1),
      value: decodeEntities(raw),
    });
  }
}

/** Layer any `xmlns` declarations on this element over the inherited scope. */
function extendScope(parent: NamespaceScope, attrs: PendingAttr[]): NamespaceScope {
  let scope: Map<string, string> | null = null;

  for (const attr of attrs) {
    const isDefault = attr.qualified === "xmlns";
    if (!isDefault && attr.prefix !== "xmlns") continue;
    scope ??= new Map(parent);
    // `xmlns:w=""` undeclares the prefix. Keeping the parent's binding would
    // resolve an element to a namespace the document just said it is not in.
    const key = isDefault ? "" : attr.local;
    if (attr.value === "") scope.delete(key);
    else scope.set(key, attr.value);
  }

  return scope ?? parent;
}

function resolveAttrs(attrs: PendingAttr[], scope: NamespaceScope): RawAttr[] {
  return attrs.map((attr) => ({
    qualified: attr.qualified,
    // An unprefixed attribute is in no namespace — it does not inherit the
    // element's default. `Relationship/@Id` depends on exactly this.
    ns:
      attr.prefix === null
        ? null
        : attr.prefix === "xmlns"
          ? XMLNS_NS
          : (scope.get(attr.prefix) ?? null),
    local: attr.qualified === "xmlns" ? "xmlns" : attr.local,
    value: attr.value,
  }));
}

/**
 * Parse an XML document.
 *
 * Throws `XmlError` on malformed input rather than returning a document with a
 * `<parsererror>` in it the way `DOMParser` does — the caller has one thing to
 * handle instead of two, and a corrupt file can no longer convert to silence.
 */
export function parseXml(source: string): XmlDocument {
  // A UTF-8 BOM survives `TextDecoder` and is not a name-start character.
  const reader = new Reader(source.charCodeAt(0) === 0xfeff ? source.slice(1) : source);

  let documentElement: XmlElement | null = null;
  /** Open elements, innermost last. Each carries the scope in force inside it. */
  const stack: { element: XmlElement; scope: NamespaceScope }[] = [];

  const push = (node: XmlNode): void => {
    const parent = stack[stack.length - 1];
    if (parent) parent.element.childNodes.push(node);
  };

  while (!reader.done) {
    if (!reader.startsWith("<")) {
      const text = reader.takeText();
      // Whitespace-only text is kept, not dropped. It usually *is* just indent
      // between elements, and the parser's `isElement` filters skip it — but a
      // `<w:t xml:space="preserve"> </w:t>` is a single space that a document
      // needs, and nothing at this level can tell the two apart.
      if (stack.length > 0 && text !== "") {
        push({ nodeType: XmlNodeType.Text, text: decodeEntities(text) });
      }
      continue;
    }

    if (reader.startsWith("<!--")) {
      reader.advance(4);
      reader.takeUntil("-->");
      continue;
    }

    if (reader.startsWith("<![CDATA[")) {
      reader.advance(9);
      const text = reader.takeUntil("]]>");
      // CDATA is literal: no entity expansion.
      if (stack.length > 0 && text !== "") {
        push({ nodeType: XmlNodeType.Text, text });
      }
      continue;
    }

    // `<?xml ... ?>` and any other processing instruction.
    if (reader.startsWith("<?")) {
      reader.advance(2);
      reader.takeUntil("?>");
      continue;
    }

    // `<!DOCTYPE ...>`. OOXML has none, but a hand-edited file might.
    if (reader.startsWith("<!")) {
      reader.advance(2);
      reader.takeUntil(">");
      continue;
    }

    if (reader.startsWith("</")) {
      reader.advance(2);
      const name = reader.takeName();
      reader.skipSpace();
      reader.expect(">");

      const open = stack.pop();
      if (!open) throw new XmlError(`unexpected closing tag </${name}>`);
      if (open.element.nodeName !== name) {
        throw new XmlError(`</${name}> closes <${open.element.nodeName}>`);
      }
      continue;
    }

    // An opening tag.
    reader.advance();
    const qualified = reader.takeName();
    const pending = readAttributes(reader);

    const parentScope = stack[stack.length - 1]?.scope ?? ROOT_SCOPE;
    const scope = extendScope(parentScope, pending);

    const colon = qualified.indexOf(":");
    const prefix = colon === -1 ? "" : qualified.slice(0, colon);
    const local = colon === -1 ? qualified : qualified.slice(colon + 1);
    // An unprefixed element takes the default namespace; `Relationships` in a
    // `.rels` part is found this way.
    const namespaceURI = scope.get(prefix) ?? null;

    const element = new XmlElement(
      qualified,
      local,
      namespaceURI,
      resolveAttrs(pending, scope),
      [],
    );

    if (stack.length === 0) {
      if (documentElement) throw new XmlError("a second root element");
      documentElement = element;
    } else {
      push(element);
    }

    reader.skipSpace();
    if (reader.peek() === "/") {
      reader.advance();
      reader.expect(">");
      continue;
    }
    reader.expect(">");
    stack.push({ element, scope });
  }

  if (stack.length > 0) {
    throw new XmlError(`unclosed <${stack[stack.length - 1]!.element.nodeName}>`);
  }
  if (!documentElement) throw new XmlError("no root element");

  return { documentElement };
}
