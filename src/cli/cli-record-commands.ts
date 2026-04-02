import { readFile, writeFile } from "node:fs/promises";

import { Command } from "commander";

import {
  assertRecord,
  parseJsonCellRecord,
  parseJsonCellRecordArray,
  parseJsonDocument,
  parseJsonStringArray,
  writeJson,
} from "./cli-json.js";
import type { CellRecord } from "./cli-json.js";
import { parsePositiveInteger, resolveFrom, resolveOutputPath } from "./cli-shared.js";
import type { CliCommandIo } from "./cli-shared.js";
import { Workbook } from "../workbook.js";

export function registerRecordCommands(
  program: Command,
  io: CliCommandIo,
): void {
  const sheetCommand = program
    .command("sheet")
    .description("Workflow-oriented sheet import/export and comment commands");

  sheetCommand
    .command("export")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--format <format>", "export format: json or csv")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output path")
    .action(
      async (
        file: string,
        options: { format: string; headerRow: number; output?: string; sheet: string },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const format = parseSheetTransferFormat(options.format);
        const output = sheet.exportRecords({ format, headerRow: options.headerRow });

        if (options.output) {
          const outputPath = resolveFrom(io.cwd, options.output);
          const content =
            format === "json"
              ? `${JSON.stringify(output, null, 2)}\n`
              : typeof output === "string" && output.length > 0
                ? `${output}\n`
                : "";
          await writeFile(outputPath, content, "utf8");
          writeJson(io.stdout, {
            action: "sheet.export",
            format,
            input: inputPath,
            output: outputPath,
            sheet: options.sheet,
          });
          return;
        }

        io.stdout(
          format === "json"
            ? `${JSON.stringify(output, null, 2)}\n`
            : typeof output === "string" && output.length > 0
              ? `${output}\n`
              : "",
        );
      },
    );

  sheetCommand
    .command("import")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--format <format>", "import format: json or csv")
    .requiredOption("--from <file>", "source json/csv file")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--key-field <name>", "key field used for upsert mode")
    .option("--mode <mode>", "import mode: replace, append, or upsert")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          format: string;
          from: string;
          headerRow: number;
          inPlace?: boolean;
          keyField?: string;
          mode?: string;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const sourcePath = resolveFrom(io.cwd, options.from);
        const format = parseSheetTransferFormat(options.format);
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const mode = parseSheetImportMode(options.mode);

        const result =
          format === "json"
            ? sheet.importRecords(
                parseJsonCellRecordArray(await readFile(sourcePath, "utf8"), "--from"),
                {
                  headerRow: options.headerRow,
                  keyField: options.keyField,
                  mode,
                },
              )
            : sheet.importRecords(parseCsvAsRecords(await readFile(sourcePath, "utf8")), {
                headerRow: options.headerRow,
                keyField: options.keyField,
                mode,
              });

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.import",
          format,
          input: inputPath,
          mode: result.mode,
          output: outputPath,
          result,
          sheet: options.sheet,
          source: sourcePath,
        });
      },
    );

  const commentCommand = sheetCommand
    .command("comment")
    .description("Worksheet comment commands");

  const sheetRecordsCommand = sheetCommand
    .command("records")
    .description("Workflow-oriented sheet record commands");

  const layoutCommand = sheetCommand
    .command("layout")
    .description("Workflow-oriented sheet layout commands");

  layoutCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--column-widths <json>", "JSON object mapping column labels to widths")
    .option("--row-heights <json>", "JSON object mapping row numbers to heights")
    .option("--freeze-columns <count>", "number of frozen columns", parsePositiveInteger)
    .option("--freeze-rows <count>", "number of frozen rows", parsePositiveInteger)
    .option("--clear-freeze", "remove frozen panes")
    .option("--print-area <range>", "print area range")
    .option("--clear-print-area", "remove print area")
    .option("--print-title-rows <range>", "print title rows, such as 1:1")
    .option("--print-title-columns <range>", "print title columns, such as A:B")
    .option("--clear-print-titles", "remove print titles")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          clearFreeze?: boolean;
          clearPrintArea?: boolean;
          clearPrintTitles?: boolean;
          columnWidths?: string;
          freezeColumns?: number;
          freezeRows?: number;
          inPlace?: boolean;
          output?: string;
          printArea?: string;
          printTitleColumns?: string;
          printTitleRows?: string;
          rowHeights?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);

        if (options.columnWidths) {
          for (const [column, width] of Object.entries(parseWorksheetNumberMap(options.columnWidths, "--column-widths"))) {
            sheet.setColumnWidth(column, width);
          }
        }
        if (options.rowHeights) {
          for (const [rowNumberText, height] of Object.entries(parseWorksheetNumberMap(options.rowHeights, "--row-heights"))) {
            sheet.setRowHeight(Number(rowNumberText), height);
          }
        }

        if (options.clearFreeze) {
          sheet.unfreezePane();
        } else if (options.freezeColumns !== undefined || options.freezeRows !== undefined) {
          sheet.freezePane(options.freezeColumns ?? 0, options.freezeRows ?? 0);
        }

        if (options.clearPrintArea) {
          sheet.setPrintArea(null);
        } else if (options.printArea !== undefined) {
          sheet.setPrintArea(options.printArea);
        }

        if (options.clearPrintTitles) {
          sheet.setPrintTitles({ columns: null, rows: null });
        } else if (options.printTitleColumns !== undefined || options.printTitleRows !== undefined) {
          sheet.setPrintTitles({
            columns: options.printTitleColumns,
            rows: options.printTitleRows,
          });
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.layout.set",
          freezePane: sheet.getFreezePane(),
          input: inputPath,
          output: outputPath,
          printArea: sheet.getPrintArea(),
          printTitles: sheet.getPrintTitles(),
          sheet: options.sheet,
        });
      },
    );

  sheetRecordsCommand
    .command("upsert")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--key-field <name>", "key field used to match the record")
    .requiredOption("--record <json>", "JSON object keyed by header names")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          inPlace?: boolean;
          keyField: string;
          output?: string;
          record: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const record = parseJsonCellRecord(options.record, "--record");
        const result = sheet.upsertRecord(options.keyField, record, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.records.upsert",
          input: inputPath,
          output: outputPath,
          result,
          row: result.row,
          sheet: options.sheet,
        });
      },
    );

  commentCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address")
    .action(async (file: string, options: { cell: string; sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        comment: workbook.getSheet(options.sheet).getComment(options.cell),
        sheet: options.sheet,
      });
    });

  commentCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address")
    .requiredOption("--text <text>", "comment text")
    .option("--author <name>", "comment author")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          author?: string;
          cell: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          text: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const comment = workbook.getSheet(options.sheet).setComment(options.cell, options.text, {
          author: options.author,
        });
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.comment.set",
          comment,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  commentCommand
    .command("delete")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          cell: string;
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
        const sheet = workbook.getSheet(options.sheet);
        const existed = sheet.getComment(options.cell) !== null;
        sheet.removeComment(options.cell);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.comment.delete",
          deleted: existed,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("records")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .action(async (file: string, options: { headerRow: number; sheet: string }) => {
      const result = await getRecords(resolveFrom(io.cwd, file), options.sheet, options.headerRow);
      writeJson(io.stdout, result);
    });

  program
    .command("export-json")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output json path")
    .action(async (file: string, options: { headerRow: number; output?: string; sheet: string }) => {
      const result = await getRecords(resolveFrom(io.cwd, file), options.sheet, options.headerRow);

      if (options.output) {
        const outputPath = resolveFrom(io.cwd, options.output);
        await writeFile(outputPath, `${JSON.stringify(result.records, null, 2)}\n`, "utf8");
        writeJson(io.stdout, {
          action: "exportJson",
          input: result.file,
          output: outputPath,
          records: result.records.length,
          sheet: options.sheet,
        });
        return;
      }

      io.stdout(`${JSON.stringify(result.records, null, 2)}\n`);
    });

  program
    .command("export-csv")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output csv path")
    .action(async (file: string, options: { headerRow: number; output?: string; sheet: string }) => {
      const inputPath = resolveFrom(io.cwd, file);
      const workbook = await Workbook.open(inputPath);
      const csv = workbook.getSheet(options.sheet).toCsv(options.headerRow);

      if (options.output) {
        const outputPath = resolveFrom(io.cwd, options.output);
        await writeFile(outputPath, csv.length > 0 ? `${csv}\n` : "", "utf8");
        writeJson(io.stdout, {
          action: "exportCsv",
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
        return;
      }

      io.stdout(csv.length > 0 ? `${csv}\n` : "");
    });

  program
    .command("set-headers")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--headers <json>", "JSON array of header strings")
    .option("--header-row <row>", "target header row", parsePositiveInteger, 1)
    .option("--start-column <column>", "1-based start column", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          headers: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          startColumn: number;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const headers = parseJsonStringArray(options.headers, "--headers");
        sheet.setHeaders(headers, options.headerRow, options.startColumn);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setHeaders",
          headers: sheet.getHeaders(options.headerRow),
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("import-json")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--from-json <file>", "json file containing an array of records")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          fromJson: string;
          headerRow: number;
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
        const records = parseJsonCellRecordArray(
          await readFile(resolveFrom(io.cwd, options.fromJson), "utf8"),
          "--from-json",
        );
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        sheet.fromJson(records, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "importJson",
          input: inputPath,
          output: outputPath,
          records: sheet.toJson(options.headerRow),
          sheet: options.sheet,
        });
      },
    );

  program
    .command("import-csv")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--from-csv <file>", "csv file containing header row and records")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          fromCsv: string;
          headerRow: number;
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
        const csv = await readFile(resolveFrom(io.cwd, options.fromCsv), "utf8");
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        sheet.fromCsv(csv, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "importCsv",
          input: inputPath,
          output: outputPath,
          records: sheet.toJson(options.headerRow),
          sheet: options.sheet,
        });
      },
    );

  program
    .command("add-record")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--record <json>", "JSON object keyed by header names")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          inPlace?: boolean;
          output?: string;
          record: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const record = parseJsonCellRecord(options.record, "--record");
        sheet.addRecord(record, options.headerRow);
        await workbook.save(outputPath);
        const result = await getRecords(outputPath, options.sheet, options.headerRow);
        writeJson(io.stdout, {
          action: "addRecord",
          input: inputPath,
          output: outputPath,
          record,
          records: result.records,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("set-record")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--row <row>", "1-based row number", parsePositiveInteger)
    .requiredOption("--record <json>", "JSON object keyed by header names")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          inPlace?: boolean;
          output?: string;
          record: string;
          row: number;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const record = parseJsonCellRecord(options.record, "--record");
        sheet.setRecord(options.row, record, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setRecord",
          input: inputPath,
          output: outputPath,
          record: await getRecord(outputPath, options.sheet, options.row, options.headerRow),
          row: options.row,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("set-records")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--records <json>", "JSON array of record objects")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          inPlace?: boolean;
          output?: string;
          records: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const records = parseJsonCellRecordArray(options.records, "--records");
        sheet.setRecords(records, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setRecords",
          input: inputPath,
          output: outputPath,
          records: (await getRecords(outputPath, options.sheet, options.headerRow)).records,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("delete-record")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--row <row>", "1-based row number", parsePositiveInteger)
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          inPlace?: boolean;
          output?: string;
          row: number;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        sheet.deleteRecord(options.row, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "deleteRecord",
          input: inputPath,
          output: outputPath,
          records: (await getRecords(outputPath, options.sheet, options.headerRow)).records,
          row: options.row,
          sheet: options.sheet,
        });
      },
    );
}

function parseSheetTransferFormat(value: string): "json" | "csv" {
  if (value === "json" || value === "csv") {
    return value;
  }

  throw new Error(`Expected --format to be json or csv, got: ${value}`);
}

function parseSheetImportMode(value?: string): "append" | "replace" | "upsert" | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (value === "append" || value === "replace" || value === "upsert") {
    return value;
  }

  throw new Error(`Expected --mode to be append, replace, or upsert, got: ${value}`);
}

function parseCsvAsRecords(source: string): CellRecord[] {
  const rows = source.replace(/\r/g, "").split("\n");
  if (rows.at(-1) === "") {
    rows.pop();
  }
  if (rows.length === 0) {
    return [];
  }

  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");
  sheet.fromCsv(source, 1);
  return sheet.toJson() as CellRecord[];
}

function parseWorksheetNumberMap(source: string, label: string): Record<string, number | null> {
  const record = assertRecord(parseJsonDocument(source, label), label);
  const next: Record<string, number | null> = {};

  for (const [key, value] of Object.entries(record)) {
    if (value === null) {
      next[key] = null;
      continue;
    }

    if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
      throw new Error(`Expected ${label}.${key} to be a positive number or null`);
    }

    next[key] = value;
  }

  return next;
}

async function getRecords(
  filePath: string,
  sheetName: string,
  headerRow: number,
): Promise<{
  file: string;
  headerRow: number;
  records: CellRecord[];
  sheet: string;
}> {
  const workbook = await Workbook.open(filePath);
  const sheet = workbook.getSheet(sheetName);
  return {
    file: filePath,
    headerRow,
    records: sheet.getRecords(headerRow),
    sheet: sheetName,
  };
}

async function getRecord(
  filePath: string,
  sheetName: string,
  row: number,
  headerRow: number,
): Promise<CellRecord | null> {
  const workbook = await Workbook.open(filePath);
  return workbook.getSheet(sheetName).getRecord(row, headerRow);
}
