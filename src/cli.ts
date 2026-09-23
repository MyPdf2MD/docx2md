#!/usr/bin/env node
/**
 * `docx2md file.docx [-o out.md]`
 *
 * Reads a path or stdin, writes Markdown to a file or stdout, and exits
 * non-zero on any failure. Nothing else — a converter that also reformatted,
 * watched directories or guessed output names would be a worse converter.
 */

import { readFile, writeFile } from "node:fs/promises";
import process from "node:process";

import { convert } from "./index.js";
import { isConversionError } from "./types.js";

const USAGE = `docx2md — convert a .docx to Markdown

Usage:
  docx2md <file.docx> [-o <out.md>]
  docx2md < file.docx > out.md

Options:
  -o, --output <path>   Write to <path> instead of stdout
  -h, --help            Show this message
  -v, --version         Show the version

Exits 0 on success, 1 on a file this package cannot read, 2 on bad usage.`;

/** Bad usage is worth its own code: a script can tell "you called me wrong"
 *  from "that document is unreadable" without parsing stderr. */
const EXIT_FAILURE = 1;
const EXIT_USAGE = 2;

interface Args {
  input: string | null;
  output: string | null;
}

function parseArgs(argv: string[]): Args {
  const args: Args = { input: null, output: null };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]!;

    if (arg === "-o" || arg === "--output") {
      const value = argv[i + 1];
      if (value === undefined || value.startsWith("-")) {
        throw new UsageError(`${arg} needs a path`);
      }
      args.output = value;
      i += 1;
      continue;
    }

    if (arg.startsWith("--output=")) {
      args.output = arg.slice("--output=".length);
      if (args.output === "") throw new UsageError("--output needs a path");
      continue;
    }

    // `--` ends the options, so a file named `-o.docx` is still reachable.
    if (arg === "--") {
      const rest = argv.slice(i + 1);
      if (rest.length > 1) throw new UsageError("expected one input file");
      if (rest[0] !== undefined) args.input = rest[0];
      break;
    }

    if (arg.startsWith("-") && arg !== "-") {
      throw new UsageError(`unknown option ${arg}`);
    }

    if (args.input !== null) throw new UsageError("expected one input file");
    args.input = arg;
  }

  return args;
}

class UsageError extends Error {}

async function readStdin(): Promise<Uint8Array> {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) chunks.push(chunk as Buffer);
  return new Uint8Array(Buffer.concat(chunks));
}

async function readVersion(): Promise<string> {
  const url = new URL("../package.json", import.meta.url);
  const manifest: unknown = JSON.parse(await readFile(url, "utf8"));
  return typeof manifest === "object" &&
    manifest !== null &&
    "version" in manifest &&
    typeof manifest.version === "string"
    ? manifest.version
    : "unknown";
}

async function main(argv: string[]): Promise<number> {
  if (argv.includes("-h") || argv.includes("--help")) {
    process.stdout.write(`${USAGE}\n`);
    return 0;
  }

  if (argv.includes("-v") || argv.includes("--version")) {
    process.stdout.write(`${await readVersion()}\n`);
    return 0;
  }

  const args = parseArgs(argv);

  // No path and a terminal on stdin means the user typed `docx2md` and
  // expected to be told what it wants, not to hang waiting for bytes.
  if (args.input === null && process.stdin.isTTY) {
    process.stderr.write(`${USAGE}\n`);
    return EXIT_USAGE;
  }

  const bytes =
    args.input === null || args.input === "-"
      ? await readStdin()
      : new Uint8Array(await readFile(args.input));

  const markdown = await convert(bytes);

  if (args.output === null) process.stdout.write(markdown);
  else await writeFile(args.output, markdown, "utf8");

  return 0;
}

const describe = (error: unknown): string => {
  if (isConversionError(error)) {
    switch (error.code) {
      case "INVALID_DOCX":
        return "not a readable .docx";
      case "OLE_CONTAINER":
        return "this is a Word 97–2003 .doc or an encrypted .docx; re-save it as a plain .docx";
      case "TIMEOUT":
        return "timed out";
      case "CANCELLED":
        return "cancelled";
      default:
        return error.code;
    }
  }
  // ENOENT, EACCES and friends already say what went wrong, in the user's
  // language and naming the path. Rewriting them would lose all of that.
  return error instanceof Error ? error.message : String(error);
};

main(process.argv.slice(2))
  .then((code) => {
    process.exitCode = code;
  })
  .catch((error: unknown) => {
    const prefix = error instanceof UsageError ? "docx2md" : "docx2md: error";
    process.stderr.write(`${prefix}: ${describe(error)}\n`);
    if (error instanceof UsageError) process.stderr.write(`\n${USAGE}\n`);
    process.exitCode = error instanceof UsageError ? EXIT_USAGE : EXIT_FAILURE;
  });
