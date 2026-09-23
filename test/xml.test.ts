/**
 * The bundled XML reader.
 *
 * `docx-parser.test.ts` already exercises it end to end through real OOXML, so
 * what is specified here is the part that file cannot reach on purpose: the
 * namespace edge cases a `.docx` in the wild will eventually contain, and the
 * malformed input that must throw rather than convert to silence.
 */

import { describe, expect, it } from "vitest";

import { isElement, parseXml, XmlError, type XmlElement } from "../src/xml.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function root(source: string): XmlElement {
  const element = parseXml(source).documentElement;
  if (!element) throw new Error("no root element");
  return element;
}

describe("namespaces", () => {
  it("resolves a prefix to the URI its declaration binds", () => {
    const el = root(`<w:p xmlns:w="${W}"/>`);
    expect(el.namespaceURI).toBe(W);
    expect(el.localName).toBe("p");
    expect(el.nodeName).toBe("w:p");
  });

  it("gives an unprefixed element the default namespace", () => {
    const el = root(`<Relationships xmlns="urn:rels"><Relationship Id="rId1"/></Relationships>`);
    expect(el.namespaceURI).toBe("urn:rels");
    expect(el.getElementsByTagNameNS("urn:rels", "Relationship")).toHaveLength(1);
  });

  it("leaves an unprefixed attribute in no namespace", () => {
    // `Relationship/@Id` depends on exactly this: it must not inherit the
    // element's default namespace the way the element name does.
    const el = root(`<Relationship xmlns="urn:rels" Id="rId1"/>`);
    expect(el.getAttribute("Id")).toBe("rId1");
    expect(el.getAttributeNS("urn:rels", "Id")).toBeNull();
  });

  it("lets an inner declaration shadow an outer one", () => {
    const el = root(`<a xmlns:p="urn:one"><b xmlns:p="urn:two"><p:c/></b><p:d/></a>`);
    const [inner] = el.getElementsByTagNameNS("urn:two", "c");
    const [outer] = el.getElementsByTagNameNS("urn:one", "d");
    expect(inner?.nodeName).toBe("p:c");
    expect(outer?.nodeName).toBe("p:d");
  });

  it("honours an undeclaration, rather than falling back to the parent", () => {
    // `xmlns=""` says this subtree is in no namespace. Keeping the parent's
    // binding would resolve an element into a namespace the document has just
    // said it is not in.
    const el = root(`<a xmlns="urn:one"><b xmlns=""><c/></b></a>`);
    expect(el.getElementsByTagNameNS("urn:one", "c")).toHaveLength(0);
    const [b] = el.getElementsByTagName("b");
    expect(b?.namespaceURI).toBeNull();
  });

  it("reads an undeclared prefix as no namespace rather than throwing", () => {
    // A namespace-blind document still has to convert; `isNamed` in the parser
    // falls back to matching the raw prefix for precisely this case.
    const el = root(`<w:document><w:body/></w:document>`);
    expect(el.namespaceURI).toBeNull();
    expect(el.getElementsByTagName("w:body")).toHaveLength(1);
  });

  it("binds the reserved xml prefix without a declaration", () => {
    const el = root(`<w:t xmlns:w="${W}" xml:space="preserve"> a </w:t>`);
    expect(
      el.getAttributeNS("http://www.w3.org/XML/1998/namespace", "space"),
    ).toBe("preserve");
  });
});

describe("text", () => {
  it("expands the five predefined entities", () => {
    const el = root(`<t>&lt;a&gt; &amp; &quot;b&quot; &apos;c&apos;</t>`);
    expect(el.textContent).toBe(`<a> & "b" 'c'`);
  });

  it("expands decimal and hex character references", () => {
    const el = root(`<t>&#65;&#x42;&#128512;</t>`);
    expect(el.textContent).toBe("AB\u{1F600}");
  });

  it("leaves an unknown named entity as written", () => {
    // Throwing on `&nbsp;` would refuse a document that every other reader
    // accepts, and dropping it would silently lose a character.
    expect(root(`<t>a&nbsp;b</t>`).textContent).toBe("a&nbsp;b");
  });

  it("treats CDATA as literal, with no entity expansion", () => {
    expect(root(`<t><![CDATA[a &amp; <b>]]></t>`).textContent).toBe("a &amp; <b>");
  });

  it("concatenates text across nested elements, in document order", () => {
    expect(root(`<p>one <r>two</r> three</p>`).textContent).toBe("one two three");
  });

  it("keeps significant whitespace inside an element", () => {
    // A `.docx` is full of runs that are a single space; trimming them welds
    // words together.
    expect(root(`<t> </t>`).textContent).toBe(" ");
  });

  it("keeps whitespace-only text, leaving it to the caller to skip", () => {
    // It is usually just indentation, but a `<w:t xml:space="preserve"> </w:t>`
    // is a space the document needs and nothing here can tell the two apart.
    // The parser filters with `isElement`; `children` is that filter.
    const el = root(`<body>\n  <p/>\n  <p/>\n</body>`);

    expect(el.children).toHaveLength(2);
    expect(el.childNodes.filter((node) => !isElement(node))).not.toHaveLength(0);
  });
});

describe("structure", () => {
  it("reads self-closing and empty elements alike", () => {
    expect(root(`<a><b/><c></c></a>`).children).toHaveLength(2);
  });

  it("tolerates whitespace around attributes and the closing slash", () => {
    const el = root(`<a  x = "1"   y='2'   />`);
    expect(el.getAttribute("x")).toBe("1");
    expect(el.getAttribute("y")).toBe("2");
  });

  it("skips the declaration, comments and a doctype", () => {
    const el = root(
      `<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE a><!-- note --><a><!-- x --><b/></a>`,
    );
    expect(el.nodeName).toBe("a");
    expect(el.children).toHaveLength(1);
  });

  it("skips a leading byte order mark", () => {
    expect(root(`﻿<a/>`).nodeName).toBe("a");
  });

  it("returns descendants in document order, excluding self", () => {
    const el = root(`<w:p xmlns:w="${W}"><w:r><w:t>a</w:t></w:r><w:t>b</w:t></w:p>`);
    expect(el.getElementsByTagNameNS(W, "t").map((n) => n.textContent)).toEqual([
      "a",
      "b",
    ]);
  });
});

describe("malformed input", () => {
  const cases: [string, string][] = [
    ["a mismatched closing tag", "<a><b></a>"],
    ["an unclosed element", "<a><b>text"],
    ["an unterminated attribute", '<a x="1/>'],
    ["an unquoted attribute value", "<a x=1/>"],
    ["a stray closing tag", "<a/></b>"],
    ["a second root element", "<a/><b/>"],
    ["an unterminated comment", "<a><!-- forever"],
    ["nothing at all", ""],
  ];

  for (const [label, source] of cases) {
    it(`throws on ${label}`, () => {
      expect(() => parseXml(source)).toThrow(XmlError);
    });
  }
});
