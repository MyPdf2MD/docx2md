/**
 * The public API: `convert`, and the promises the README makes about it.
 */

import { describe, expect, it } from "vitest";

import { convert, DocxErrorCode, isConversionError } from "../src/index.js";
import { makeDocx, paragraph, pStyle, run, styleDefinition } from "./docx-fixtures.js";

const STYLES = styleDefinition("Heading1", { name: "heading 1" });

const SAMPLE = () =>
  makeDocx({
    styles: STYLES,
    body: paragraph(run("Report"), pStyle("Heading1")) + paragraph(run("Body.")),
  });

async function codeFor(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
    return "NO_ERROR";
  } catch (error) {
    return isConversionError(error) ? error.code : "NOT_A_CONVERSION_ERROR";
  }
}

describe("convert", () => {
  it("takes a Uint8Array and returns Markdown", async () => {
    expect(await convert(SAMPLE())).toBe("# Report\n\nBody.\n");
  });

  it("takes an ArrayBuffer, which is what File.arrayBuffer() gives a browser", async () => {
    const bytes = SAMPLE();
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;

    expect(await convert(buffer)).toBe("# Report\n\nBody.\n");
  });

  it("takes a Node Buffer, which is what readFile gives", async () => {
    expect(await convert(Buffer.from(SAMPLE()))).toBe("# Report\n\nBody.\n");
  });

  it("rejects an empty input rather than returning an empty string", async () => {
    expect(await codeFor(convert(new Uint8Array(0)))).toBe(DocxErrorCode.INVALID_DOCX);
  });

  it("rejects an OLE container by name, so the advice can be specific", async () => {
    // A Word 97–2003 .doc, or a .docx encrypted at rest.
    const ole = new Uint8Array(512);
    ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(await codeFor(convert(ole))).toBe(DocxErrorCode.OLE_CONTAINER);
  });

  it("reports an already-aborted signal as a cancellation", async () => {
    const controller = new AbortController();
    controller.abort();
    expect(await codeFor(convert(SAMPLE(), { signal: controller.signal }))).toBe(
      DocxErrorCode.CANCELLED,
    );
  });

  it("turns timeoutMs into a deadline the parser can honour", async () => {
    expect(await codeFor(convert(SAMPLE(), { timeoutMs: -1 }))).toBe(
      DocxErrorCode.TIMEOUT,
    );
  });

  it("runs to completion with no timeout set", async () => {
    expect(await convert(SAMPLE(), {})).toContain("# Report");
  });

  it("reports progress through to done", async () => {
    const stages: string[] = [];
    await convert(SAMPLE(), { onProgress: (p) => stages.push(p.stage) });

    expect(stages[0]).toBe("loading");
    expect(stages.at(-1)).toBe("done");
  });

  it("never throws a bare Error — every rejection carries a code", async () => {
    const codes = await Promise.all([
      codeFor(convert(new Uint8Array([1, 2, 3, 4]))),
      codeFor(convert(new Uint8Array(0))),
      codeFor(convert(makeDocx({ body: "<w:p><w:r><w:t>unclosed" }))),
    ]);

    expect(codes).toEqual([
      DocxErrorCode.INVALID_DOCX,
      DocxErrorCode.INVALID_DOCX,
      DocxErrorCode.INVALID_DOCX,
    ]);
  });
});
