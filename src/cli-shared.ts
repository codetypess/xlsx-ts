import { resolve } from "node:path";

import { InvalidArgumentError } from "commander";

import type { Writer } from "./cli-json.js";

export interface CliCommandIo {
  cwd: string;
  stdout: Writer;
}

export class CliExitError extends Error {
  readonly exitCode: number;

  constructor(exitCode: number) {
    super(`CLI exited with code ${exitCode}`);
    this.exitCode = exitCode;
  }
}

export function resolveOutputPath(
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

export function parsePositiveInteger(value: string): number {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new InvalidArgumentError(`Expected a positive integer, got: ${value}`);
  }

  return parsed;
}

export function resolveFrom(cwd: string, targetPath: string): string {
  return resolve(cwd, targetPath);
}
