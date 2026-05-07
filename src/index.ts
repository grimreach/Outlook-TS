#!/usr/bin/env node
import { readFile, writeFile } from "node:fs/promises";
import { evaluateDiagnostics } from "./rules.js";
import { renderConsoleReport, renderMarkdownReport } from "./report.js";
import type { OutlookDiagnostics } from "./types.js";

interface CliArgs {
  command: string;
  input?: string;
  markdown?: string;
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));

  if (args.command !== "diagnose") {
    printHelp();
    process.exitCode = args.command === "help" ? 0 : 1;
    return;
  }

  if (!args.input) {
    throw new Error("Missing --input <diagnostics.json>");
  }

  const diagnostics = JSON.parse(stripByteOrderMark(await readFile(args.input, "utf8"))) as OutlookDiagnostics;
  const findings = evaluateDiagnostics(diagnostics);
  const consoleReport = renderConsoleReport(diagnostics, findings);

  console.log(consoleReport);

  if (args.markdown) {
    await writeFile(args.markdown, renderMarkdownReport(diagnostics, findings), "utf8");
  }
}

function parseArgs(argv: string[]): CliArgs {
  const [command = "help", ...rest] = argv;
  const parsed: CliArgs = { command };

  for (let index = 0; index < rest.length; index += 1) {
    const arg = rest[index];
    if (arg === "--input") {
      parsed.input = rest[++index];
    } else if (arg === "--markdown") {
      parsed.markdown = rest[++index];
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }

  return parsed;
}

function stripByteOrderMark(value: string): string {
  return value.charCodeAt(0) === 0xfeff ? value.slice(1) : value;
}

function printHelp(): void {
  console.log(`Outlook-TS

Usage:
  outlook-ts diagnose --input <diagnostics.json> [--markdown report.md]

Examples:
  node dist/index.js diagnose --input samples/classic-modern-toggle.json
  node dist/index.js diagnose --input outlook-diagnostics.json --markdown report.md
`);
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
