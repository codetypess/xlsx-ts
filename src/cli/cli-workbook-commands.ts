import { Command } from "commander";

import { parseJsonCellValue, writeJson } from "./cli-json.js";
import type { CellError, CellValue, DefinedName, SheetVisibility } from "../types.js";
import { parseNonNegativeInteger, parsePositiveInteger, resolveFrom, resolveOutputPath } from "./cli-shared.js";
import type { CliCommandIo } from "./cli-shared.js";
import { Workbook } from "../workbook.js";

interface InspectResult {
  activeSheet: string;
  definedNames: DefinedName[];
  file: string;
  sheets: Array<{
    columnCount: number;
    headers: string[];
    name: string;
    physicalRangeRef: string | null;
    rangeRef: string | null;
    rowCount: number;
    visibility: SheetVisibility;
  }>;
}

interface GetCellResult {
  backgroundColor: string | null;
  cell: string;
  exists: boolean;
  error: CellError | null;
  file: string;
  formula: string | null;
  numberFormat: string | null;
  rawType: string | null;
  sheet: string;
  styleId: number | null;
  type: string;
  value: CellValue;
}

export function registerWorkbookCommands(
  program: Command,
  io: CliCommandIo,
): void {
  const workbookCommand = program
    .command("workbook")
    .description("Workflow-oriented workbook metadata commands");

  const activeCommand = workbookCommand
    .command("active")
    .description("Workbook active-sheet commands");

  const visibilityCommand = workbookCommand
    .command("visibility")
    .description("Workbook sheet visibility commands");

  const definedNameCommand = workbookCommand
    .command("defined-name")
    .description("Workbook defined-name commands");

  program
    .command("create")
    .argument("<file>", "output xlsx file")
    .option("--sheet <name>", "initial sheet name", "Sheet1")
    .action(async (file: string, options: { sheet: string }) => {
      const outputPath = resolveFrom(io.cwd, file);
      const workbook = Workbook.create(options.sheet);
      await workbook.save(outputPath);
      writeJson(io.stdout, {
        action: "createWorkbook",
        output: outputPath,
        sheets: workbook.getSheets().map((sheet) => sheet.name),
      });
    });

  program
    .command("inspect")
    .argument("<file>", "input xlsx file")
    .option("--header-row <row>", "header row used for the sheet preview", parsePositiveInteger, 1)
    .action(async (file: string, options: { headerRow: number }) => {
      const result = await inspectWorkbook(resolveFrom(io.cwd, file), options.headerRow);
      writeJson(io.stdout, result);
    });

  program
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .action(async (file: string, options: { cell: string; sheet: string }) => {
      const result = await getCell(resolveFrom(io.cwd, file), options.sheet, options.cell);
      writeJson(io.stdout, result);
    });

  activeCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .action(async (file: string) => {
      const inputPath = resolveFrom(io.cwd, file);
      const workbook = await Workbook.open(inputPath);
      writeJson(io.stdout, {
        activeSheet: workbook.getActiveSheet().name,
        file: inputPath,
      });
    });

  activeCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.setActiveSheet(options.sheet);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "workbook.active.set",
          activeSheet: workbook.getActiveSheet().name,
          input: inputPath,
          output: outputPath,
        });
      },
    );

  visibilityCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const inputPath = resolveFrom(io.cwd, file);
      const workbook = await Workbook.open(inputPath);
      writeJson(io.stdout, {
        file: inputPath,
        sheet: options.sheet,
        visibility: workbook.getSheetVisibility(options.sheet),
      });
    });

  visibilityCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--visibility <value>", "visible, hidden, or veryHidden")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          sheet: string;
          visibility: SheetVisibility;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.setSheetVisibility(options.sheet, options.visibility);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "workbook.visibility.set",
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
          visibility: workbook.getSheetVisibility(options.sheet),
        });
      },
    );

  definedNameCommand
    .command("list")
    .argument("<file>", "input xlsx file")
    .option("--scope <sheet>", "optional sheet scope filter")
    .action(async (file: string, options: { scope?: string }) => {
      const inputPath = resolveFrom(io.cwd, file);
      const workbook = await Workbook.open(inputPath);
      const normalizedScope = options.scope ? (workbook.tryGetSheet(options.scope)?.name ?? options.scope) : undefined;
      const definedNames = options.scope
        ? workbook.getDefinedNames().filter((definedName) => definedName.scope === (normalizedScope ?? null))
        : workbook.getDefinedNames();
      writeJson(io.stdout, {
        definedNames,
        file: inputPath,
      });
    });

  definedNameCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--name <value>", "defined name")
    .option("--scope <sheet>", "optional sheet scope")
    .action(async (file: string, options: { name: string; scope?: string }) => {
      const inputPath = resolveFrom(io.cwd, file);
      const workbook = await Workbook.open(inputPath);
      const normalizedScope = options.scope ? (workbook.tryGetSheet(options.scope)?.name ?? options.scope) : null;
      const definedName = workbook.getDefinedNames().find(
        (candidate) => candidate.name === options.name && candidate.scope === normalizedScope,
      ) ?? null;
      writeJson(io.stdout, {
        definedName,
        file: inputPath,
      });
    });

  definedNameCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--name <value>", "defined name")
    .requiredOption("--value <ref>", "defined name formula or reference")
    .option("--scope <sheet>", "optional sheet scope")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          name: string;
          output?: string;
          scope?: string;
          value: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.setDefinedName(options.name, options.value, {
          scope: options.scope,
        });
        await workbook.save(outputPath);
        const normalizedScope = options.scope ? (workbook.tryGetSheet(options.scope)?.name ?? options.scope) : null;
        const definedName = workbook.getDefinedNames().find(
          (candidate) => candidate.name === options.name && candidate.scope === normalizedScope,
        ) ?? null;
        writeJson(io.stdout, {
          action: "workbook.definedName.set",
          definedName,
          input: inputPath,
          output: outputPath,
        });
      },
    );

  definedNameCommand
    .command("delete")
    .argument("<file>", "input xlsx file")
    .requiredOption("--name <value>", "defined name")
    .option("--scope <sheet>", "optional sheet scope")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          name: string;
          output?: string;
          scope?: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const deleted = workbook.getDefinedName(options.name, options.scope) !== null;
        workbook.deleteDefinedName(options.name, options.scope);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "workbook.definedName.delete",
          deleted,
          input: inputPath,
          output: outputPath,
        });
      },
    );

  program
    .command("add-sheet")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "new sheet name")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.addSheet(options.sheet);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "addSheet",
          input: inputPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("rename-sheet")
    .argument("<file>", "input xlsx file")
    .requiredOption("--from <name>", "current sheet name")
    .requiredOption("--to <name>", "new sheet name")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          from: string;
          inPlace?: boolean;
          output?: string;
          to: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.renameSheet(options.from, options.to);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "renameSheet",
          input: inputPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("move-sheet")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name to move")
    .requiredOption("--index <number>", "target sheet index (0-based)", parseNonNegativeInteger)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          index: number;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.moveSheet(options.sheet, options.index);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "moveSheet",
          input: inputPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("delete-sheet")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name to delete")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.deleteSheet(options.sheet);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "deleteSheet",
          input: inputPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("copy-style")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--from <cell>", "source cell address")
    .requiredOption("--to <cell>", "target cell address")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          from: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          to: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        sheet.copyStyle(options.from, options.to);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "copyStyle",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.to),
        });
      },
    );

  program
    .command("set-number-format")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .requiredOption("--format <code>", "number format code")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          cell: string;
          format: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.getSheet(options.sheet).setNumberFormat(options.cell, options.format);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setNumberFormat",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.cell),
        });
      },
    );

  program
    .command("set-background-color")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .option("--color <rgb>", "ARGB color, such as FFFF0000")
    .option("--clear", "remove the solid background fill")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          cell: string;
          clear?: boolean;
          color?: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const actionCount = Number(Boolean(options.clear)) + Number(options.color !== undefined);
        if (actionCount !== 1) {
          throw new Error("Exactly one of --color or --clear is required");
        }

        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.getSheet(options.sheet).setBackgroundColor(options.cell, options.clear ? null : options.color!);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setBackgroundColor",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.cell),
        });
      },
    );

  program
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .option("--value <json>", "JSON literal for the next cell value")
    .option("--text <value>", "plain string value without JSON quoting")
    .option("--formula <formula>", "formula text without the leading equals sign")
    .option("--cached-value <json>", "JSON literal for the formula cached value")
    .option("--cached-text <value>", "plain string cached value for a formula")
    .option("--clear", "delete the cell instead of writing a value")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          cachedText?: string;
          cachedValue?: string;
          cell: string;
          clear?: boolean;
          formula?: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          text?: string;
          value?: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const actionCount =
          Number(Boolean(options.clear)) +
          Number(options.formula !== undefined) +
          Number(options.text !== undefined) +
          Number(options.value !== undefined);

        if (actionCount !== 1) {
          throw new Error("Exactly one of --value, --text, --formula, or --clear is required");
        }

        if (options.formula !== undefined) {
          const cachedValue =
            options.cachedText !== undefined
              ? options.cachedText
              : options.cachedValue !== undefined
                ? parseJsonCellValue(options.cachedValue, "--cached-value")
                : undefined;
          sheet.setFormula(options.cell, options.formula, cachedValue === undefined ? {} : { cachedValue });
        } else if (options.clear) {
          sheet.deleteCell(options.cell);
        } else {
          const value =
            options.text !== undefined
              ? options.text
              : parseJsonCellValue(options.value!, "--value");
          sheet.setCell(options.cell, value);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: options.clear ? "clearCell" : options.formula !== undefined ? "setFormula" : "setCell",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.cell),
        });
      },
    );
}

async function inspectWorkbook(filePath: string, headerRow: number): Promise<InspectResult> {
  const workbook = await Workbook.open(filePath);
  const sheets = workbook.getSheets().map((sheet) => ({
    columnCount: sheet.columnCount,
    headers: trimTrailingEmptyStrings(sheet.getHeaders(headerRow)),
    name: sheet.name,
    physicalRangeRef: sheet.getPhysicalRangeRef(),
    rangeRef: sheet.getRangeRef(),
    rowCount: sheet.rowCount,
    visibility: workbook.getSheetVisibility(sheet.name),
  }));

  return {
    activeSheet: workbook.getActiveSheet().name,
    definedNames: workbook.getDefinedNames(),
    file: filePath,
    sheets,
  };
}

async function getCell(filePath: string, sheetName: string, cellAddress: string): Promise<GetCellResult> {
  const workbook = await Workbook.open(filePath);
  const sheet = workbook.getSheet(sheetName);
  const cell = sheet.cell(cellAddress);

  return {
    backgroundColor: sheet.getBackgroundColor(cellAddress),
    cell: cellAddress,
    exists: cell.exists,
    error: cell.error,
    file: filePath,
    formula: cell.formula,
    numberFormat: cell.numberFormat?.code ?? null,
    rawType: cell.rawType,
    sheet: sheetName,
    styleId: cell.styleId,
    type: cell.type,
    value: cell.value,
  };
}

function trimTrailingEmptyStrings(values: string[]): string[] {
  let end = values.length;

  while (end > 0 && values[end - 1] === "") {
    end -= 1;
  }

  return values.slice(0, end);
}
