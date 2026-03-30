#!/usr/bin/env node

import { realpathSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Command, CommanderError } from "commander";

import { formatError } from "./cli-json.js";
import type { Writer } from "./cli-json.js";
import { registerApplyCommands } from "./cli-apply-commands.js";
import { registerRecordCommands } from "./cli-record-commands.js";
import { CliExitError } from "./cli-shared.js";
import { registerTableCommands } from "./cli-table-commands.js";
import { registerValidateCommands } from "./cli-validate-commands.js";
import { registerWorkbookCommands } from "./cli-workbook-commands.js";

interface CliIo {
  cwd?: string;
  stderr?: Writer;
  stdout?: Writer;
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
  );

  registerRecordCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
  );

  registerTableCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
  );

  registerApplyCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
  );

  registerValidateCommands(program, {
    cwd: io.cwd,
    stdout: io.stdout,
  });

  return program;
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
