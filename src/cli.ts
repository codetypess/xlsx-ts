#!/usr/bin/env node

import { existsSync, readFileSync, realpathSync } from "node:fs";
import { dirname, join, parse, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Command, CommanderError } from "commander";

import { formatError } from "./cli/cli-json.js";
import type { Writer } from "./cli/cli-json.js";
import { registerApplyCommands } from "./cli/cli-apply-commands.js";
import { registerRecordCommands } from "./cli/cli-record-commands.js";
import { CliExitError } from "./cli/cli-shared.js";
import { registerTableCommands } from "./cli/cli-table-commands.js";
import { registerValidateCommands } from "./cli/cli-validate-commands.js";
import { registerWorkbookCommands } from "./cli/cli-workbook-commands.js";

interface CliIo {
  cwd?: string;
  stderr?: Writer;
  stdout?: Writer;
}

const CLI_VERSION = resolveCliVersion();

export async function runCli(argv: string[], io: CliIo = {}): Promise<number> {
  const stdout = io.stdout ?? ((chunk: string) => process.stdout.write(chunk));
  const stderr = io.stderr ?? ((chunk: string) => process.stderr.write(chunk));
  const cwd = io.cwd ?? process.cwd();
  const program = createProgram({ cwd, stderr, stdout });

  try {
    await program.parseAsync(["node", "fastxlsx", ...argv], { from: "node" });
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
    .name("fastxlsx")
    .version(CLI_VERSION)
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

function resolveCliVersion(): string {
  let currentDirectory = dirname(fileURLToPath(import.meta.url));
  const filesystemRoot = parse(currentDirectory).root;

  while (true) {
    const packageJsonPath = join(currentDirectory, "package.json");
    if (existsSync(packageJsonPath)) {
      try {
        const packageJson = JSON.parse(readFileSync(packageJsonPath, "utf8")) as { version?: unknown };
        if (typeof packageJson.version === "string" && packageJson.version.length > 0) {
          return packageJson.version;
        }
      } catch {
        // Keep walking upward until a readable package.json with a version is found.
      }
    }

    if (currentDirectory === filesystemRoot) {
      break;
    }

    currentDirectory = dirname(currentDirectory);
  }

  return "0.0.0";
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
