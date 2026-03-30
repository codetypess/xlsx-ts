#!/usr/bin/env node

import { realpathSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Command, CommanderError, InvalidArgumentError } from "commander";

import { formatError, writeJson } from "./cli-json.js";
import type { Writer } from "./cli-json.js";
import { registerApplyCommands } from "./cli-apply-commands.js";
import { registerRecordCommands } from "./cli-record-commands.js";
import { registerTableCommands } from "./cli-table-commands.js";
import { registerWorkbookCommands } from "./cli-workbook-commands.js";
import { validateRoundtripFile } from "./roundtrip.js";

interface CliIo {
  cwd?: string;
  stderr?: Writer;
  stdout?: Writer;
}

class CliExitError extends Error {
  readonly exitCode: number;

  constructor(exitCode: number) {
    super(`CLI exited with code ${exitCode}`);
    this.exitCode = exitCode;
  }
}

export async function runCli(argv: string[], io: CliIo = {}): Promise<number> {
  const stdout = io.stdout ?? ((chunk: string) => process.stdout.write(chunk));
  const stderr = io.stderr ?? ((chunk: string) => process.stderr.write(chunk));
  const cwd = io.cwd ?? process.cwd();
  const program = createProgram({ cwd, stderr, stdout });

  try {
    await program.parseAsync(["node", "xlsx-ts", ...argv], { from: "node" });
    return 0;
  } catch (error) {
    if (error instanceof CliExitError) {
      return error.exitCode;
    }

    if (error instanceof CommanderError) {
      return error.exitCode;
    }

    stderr(`${formatError(error)}\n`);
    return 1;
  }
}

function createProgram(io: Required<CliIo>): Command {
  const program = new Command()
    .name("xlsx-ts")
    .description("Lossless-first XLSX inspection and editing CLI")
    .showHelpAfterError()
    .configureOutput({
      writeErr: io.stderr,
      writeOut: io.stdout,
    })
    .exitOverride();

  registerWorkbookCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
    {
      parsePositiveInteger,
      resolveOutputPath,
    },
  );

  registerRecordCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
    {
      parsePositiveInteger,
      resolveOutputPath,
    },
  );

  registerTableCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
    {
      parsePositiveInteger,
      resolveOutputPath,
    },
  );

  registerApplyCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
    {
      resolveOutputPath,
    },
  );

  program
    .command("validate")
    .argument("<file>", "input xlsx file")
    .option("--output <file>", "persist the roundtrip workbook to the given path")
    .action(async (file: string, options: { output?: string }) => {
      const result = await validateRoundtripFile(
        resolveFrom(io.cwd, file),
        options.output ? resolveFrom(io.cwd, options.output) : undefined,
      );
      writeJson(io.stdout, result);

      if (!result.ok) {
        throw new CliExitError(1);
      }
    });

  return program;
}

function resolveOutputPath(
  inputPath: string,
  options: {
    inPlace: boolean;
    output?: string;
  },
): string {
  if (options.inPlace && options.output) {
    throw new Error("Use either --output or --in-place, not both");
  }

  if (options.inPlace) {
    return inputPath;
  }

  if (options.output) {
    return options.output;
  }

  throw new Error("An output path is required; pass --output or use --in-place");
}

function parsePositiveInteger(value: string): number {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new InvalidArgumentError(`Expected a positive integer, got: ${value}`);
  }

  return parsed;
}

function resolveFrom(cwd: string, targetPath: string): string {
  return resolve(cwd, targetPath);
}

async function main(): Promise<void> {
  process.exitCode = await runCli(process.argv.slice(2));
}

if (
  process.argv[1] &&
  realpathSync.native(resolve(process.argv[1])) === realpathSync.native(fileURLToPath(import.meta.url))
) {
  await main();
}
