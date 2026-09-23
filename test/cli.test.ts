/**
 * The CLI, against the built `dist/cli.js` rather than the TypeScript source.
 *
 * Spawning the real thing is the only way to specify what the CLI actually
 * promises — an exit code, bytes on stdout, bytes on stderr — and it is also
 * the only test here that would catch the emitted ESM being unrunnable in
 * Node, which is a thing a type-checked source file can be.
 */

import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { makeDocx, paragraph, pStyle, run, styleDefinition } from "./docx-fixtures.js";

const CLI = fileURLToPath(new URL("../dist/cli.js", import.meta.url));

interface Run {
  code: number;
  stdout: string;
  stderr: string;
}

function runCli(args: string[], stdin?: Uint8Array): Promise<Run> {
  return new Promise((resolve) => {
    const child = execFile(
      process.execPath,
      [CLI, ...args],
      { encoding: "buffer" },
      (error, stdout, stderr) => {
        resolve({
          // `error.code` is the exit status for a non-zero exit.
          code: error && typeof error.code === "number" ? error.code : 0,
          stdout: stdout.toString("utf8"),
          stderr: stderr.toString("utf8"),
        });
      },
    );

    if (stdin) child.stdin?.end(Buffer.from(stdin));
    else child.stdin?.end();
  });
}

const DOCX = makeDocx({
  styles: styleDefinition("Heading1", { name: "heading 1" }),
  body: paragraph(run("Report"), pStyle("Heading1")) + paragraph(run("Body.")),
});

let dir: string;
let input: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "docx2md-cli-"));
  input = join(dir, "report.docx");
  await writeFile(input, DOCX);
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

describe("docx2md", () => {
  it("converts a path to stdout and exits 0", async () => {
    const result = await runCli([input]);

    expect(result.code).toBe(0);
    expect(result.stdout).toBe("# Report\n\nBody.\n");
    expect(result.stderr).toBe("");
  });

  it("writes to a file with -o, leaving stdout empty", async () => {
    const out = join(dir, "out.md");
    const result = await runCli([input, "-o", out]);

    expect(result.code).toBe(0);
    expect(result.stdout).toBe("");
    expect(await readFile(out, "utf8")).toBe("# Report\n\nBody.\n");
  });

  it("accepts --output=path as well", async () => {
    const out = join(dir, "out2.md");
    expect((await runCli([input, `--output=${out}`])).code).toBe(0);
    expect(await readFile(out, "utf8")).toContain("# Report");
  });

  it("reads stdin when given no path", async () => {
    const result = await runCli([], DOCX);

    expect(result.code).toBe(0);
    expect(result.stdout).toBe("# Report\n\nBody.\n");
  });

  it("reads stdin for an explicit -", async () => {
    expect((await runCli(["-"], DOCX)).stdout).toBe("# Report\n\nBody.\n");
  });

  it("exits 1 and says so on a file it cannot read", async () => {
    const bad = join(dir, "bad.docx");
    await writeFile(bad, Buffer.from([1, 2, 3, 4, 5, 6, 7, 8]));
    const result = await runCli([bad]);

    expect(result.code).toBe(1);
    expect(result.stdout).toBe("");
    expect(result.stderr).toContain("not a readable .docx");
  });

  it("exits 1 and names the path when the file is missing", async () => {
    const result = await runCli([join(dir, "nope.docx")]);

    expect(result.code).toBe(1);
    expect(result.stderr).toContain("nope.docx");
  });

  it("tells an OLE container apart, with advice that fits", async () => {
    const ole = join(dir, "legacy.docx");
    const bytes = Buffer.alloc(512);
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    await writeFile(ole, bytes);

    const result = await runCli([ole]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("re-save it as a plain .docx");
  });

  it("exits 2 on bad usage, which is not the same as a bad file", async () => {
    const result = await runCli(["--nonsense"], DOCX);

    expect(result.code).toBe(2);
    expect(result.stderr).toContain("unknown option --nonsense");
    expect(result.stderr).toContain("Usage:");
  });

  it("exits 2 when -o is given nothing to write to", async () => {
    expect((await runCli([input, "-o"])).code).toBe(2);
  });

  it("exits 2 on a second input file rather than silently ignoring one", async () => {
    expect((await runCli([input, input])).code).toBe(2);
  });

  it("prints usage and exits 0 for --help", async () => {
    const result = await runCli(["--help"]);

    expect(result.code).toBe(0);
    expect(result.stdout).toContain("Usage:");
  });

  it("prints the package version for --version", async () => {
    const result = await runCli(["--version"]);

    expect(result.code).toBe(0);
    expect(result.stdout.trim()).toMatch(/^\d+\.\d+\.\d+/);
  });
});
