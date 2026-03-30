import { Command } from "commander";

import { writeJson } from "./cli-json.js";
import { CliExitError, resolveFrom } from "./cli-shared.js";
import type { CliCommandIo } from "./cli-shared.js";
import { validateRoundtripFile } from "./roundtrip.js";

export function registerValidateCommands(program: Command, io: CliCommandIo): void {
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
}
