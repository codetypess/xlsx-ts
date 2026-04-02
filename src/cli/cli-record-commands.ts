import { readFile, writeFile } from "node:fs/promises";

import { Command } from "commander";

import {
  assertRecord,
  parseJsonCellRecord,
  parseJsonCellRecordArray,
  parseJsonDocument,
  parseJsonStringArray,
  resolveMatchValue,
  writeJson,
} from "./cli-json.js";
import type { CellRecord } from "./cli-json.js";
import { parseBooleanValue, parsePositiveInteger, resolveFrom, resolveOutputPath } from "./cli-shared.js";
import type { CliCommandIo } from "./cli-shared.js";
import { Workbook } from "../workbook.js";

export function registerRecordCommands(
  program: Command,
  io: CliCommandIo,
): void {
  const sheetCommand = program
    .command("sheet")
    .description("Workflow-oriented sheet metadata and record commands");

  sheetCommand
    .command("export")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--format <format>", "export format: json or csv")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--no-headers", "omit the CSV header row")
    .option("--crlf", "use CRLF line endings for CSV output")
    .option("--output <file>", "output path")
    .action(
      async (
        file: string,
        options: { crlf?: boolean; format: string; headerRow: number; headers?: boolean; output?: string; sheet: string },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const format = parseSheetTransferFormat(options.format);
        const output = sheet.exportRecords({
          format,
          headerRow: options.headerRow,
          includeHeaders: options.headers !== false,
          lineEnding: options.crlf ? "\r\n" : "\n",
        });

        if (options.output) {
          const outputPath = resolveFrom(io.cwd, options.output);
          const content =
            format === "json"
              ? `${JSON.stringify(output, null, 2)}\n`
              : typeof output === "string" && output.length > 0
                ? `${output}${options.crlf ? "\r\n" : "\n"}`
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
    .option("--trim-values", "trim CSV cell values before import")
    .option("--no-trim-headers", "preserve CSV header whitespace")
    .option("--no-infer-types", "keep imported CSV values as strings")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          format: string;
          from: string;
          headerRow: number;
          inferTypes?: boolean;
          inPlace?: boolean;
          keyField?: string;
          mode?: string;
          output?: string;
          sheet: string;
          trimHeaders?: boolean;
          trimValues?: boolean;
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
                  inferTypes: options.inferTypes !== false,
                  keyField: options.keyField,
                  mode,
                  trimHeaders: options.trimHeaders !== false,
                  trimValues: options.trimValues === true,
                },
              )
            : sheet.importRecords(parseCsvAsRecords(await readFile(sourcePath, "utf8"), {
                inferTypes: options.inferTypes !== false,
                trimHeaders: options.trimHeaders !== false,
                trimValues: options.trimValues === true,
              }), {
                headerRow: options.headerRow,
                inferTypes: options.inferTypes !== false,
                keyField: options.keyField,
                mode,
                trimHeaders: options.trimHeaders !== false,
                trimValues: options.trimValues === true,
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

  const hyperlinkCommand = sheetCommand
    .command("hyperlink")
    .description("Worksheet hyperlink commands");

  const filterCommand = sheetCommand
    .command("filter")
    .description("Worksheet auto-filter commands");

  const selectionCommand = sheetCommand
    .command("selection")
    .description("Worksheet selection commands");

  const validationCommand = sheetCommand
    .command("validation")
    .description("Worksheet data validation commands");

  const mergeCommand = sheetCommand
    .command("merge")
    .description("Worksheet merged range commands");

  const protectionCommand = sheetCommand
    .command("protection")
    .description("Worksheet protection commands");

  const sheetRecordsCommand = sheetCommand
    .command("records")
    .description("Workflow-oriented sheet record commands");

  const layoutCommand = sheetCommand
    .command("layout")
    .description("Workflow-oriented sheet layout commands");

  layoutCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--columns <json>", "JSON array of column labels to inspect")
    .option("--rows <json>", "JSON array of row numbers to inspect")
    .action(
      async (
        file: string,
        options: {
          columns?: string;
          rows?: string;
          sheet: string;
        },
      ) => {
        const workbook = await Workbook.open(resolveFrom(io.cwd, file));
        const sheet = workbook.getSheet(options.sheet);
        const columns = options.columns ? parseJsonStringArray(options.columns, "--columns") : [];
        const rows = options.rows ? parseJsonNumberArray(options.rows, "--rows") : [];

        writeJson(io.stdout, {
          action: "sheet.layout.get",
          columnWidths: Object.fromEntries(columns.map((column) => [column, sheet.getColumnWidth(column)])),
          freezePane: sheet.getFreezePane(),
          printArea: sheet.getPrintArea(),
          printTitles: sheet.getPrintTitles(),
          rowHeights: Object.fromEntries(rows.map((rowNumber) => [String(rowNumber), sheet.getRowHeight(rowNumber)])),
          sheet: options.sheet,
        });
      },
    );

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
    .command("list")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          sheet: string;
        },
      ) => {
        const workbook = await Workbook.open(resolveFrom(io.cwd, file));
        const sheet = workbook.getSheet(options.sheet);
        writeJson(io.stdout, {
          action: "sheet.records.list",
          headers: sheet.getHeaders(options.headerRow),
          records: sheet.getRecords(options.headerRow),
          sheet: options.sheet,
        });
      },
    );

  sheetRecordsCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--key-field <name>", "key field used to match the record")
    .option("--value <json>", "JSON scalar key value")
    .option("--text <value>", "plain string key value")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          keyField: string;
          sheet: string;
          text?: string;
          value?: string;
        },
      ) => {
        const workbook = await Workbook.open(resolveFrom(io.cwd, file));
        const sheet = workbook.getSheet(options.sheet);
        writeJson(io.stdout, {
          action: "sheet.records.get",
          record: sheet.findRecordBy(options.keyField, resolveMatchValue(options.value, options.text), options.headerRow),
          sheet: options.sheet,
        });
      },
    );

  sheetRecordsCommand
    .command("append")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--record <json>", "JSON object keyed by header names")
    .option("--records <json>", "JSON array of record objects keyed by header names")
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
          record?: string;
          records?: string;
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
        const records = parseSheetRecordInputs(options.record, options.records);
        sheet.appendRecords(records, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.records.append",
          appended: records.length,
          headers: sheet.getHeaders(options.headerRow),
          input: inputPath,
          output: outputPath,
          records: sheet.getRecords(options.headerRow),
          sheet: options.sheet,
        });
      },
    );

  sheetRecordsCommand
    .command("replace")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--record <json>", "JSON object keyed by header names")
    .option("--records <json>", "JSON array of record objects keyed by header names")
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
          record?: string;
          records?: string;
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
        const records = parseSheetRecordInputs(options.record, options.records);
        sheet.replaceRecords(records, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.records.replace",
          headers: sheet.getHeaders(options.headerRow),
          input: inputPath,
          output: outputPath,
          records: sheet.getRecords(options.headerRow),
          replaced: records.length,
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

  sheetRecordsCommand
    .command("delete")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--key-field <name>", "key field used to match the record")
    .option("--value <json>", "JSON scalar key value")
    .option("--text <value>", "plain string key value")
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
        const deleted = sheet.removeRecordBy(
          options.keyField,
          resolveMatchValue(options.value, options.text),
          options.headerRow,
        );
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.records.delete",
          deleted,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  commentCommand
    .command("list")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        comments: workbook.getSheet(options.sheet).getComments(),
        sheet: options.sheet,
      });
    });

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

  commentCommand
    .command("clear")
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
        const sheet = workbook.getSheet(options.sheet);
        const deleted = sheet.getComments().length;
        sheet.clearComments();
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.comment.clear",
          cleared: deleted,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  hyperlinkCommand
    .command("list")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        hyperlinks: workbook.getSheet(options.sheet).getHyperlinks(),
        sheet: options.sheet,
      });
    });

  hyperlinkCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address")
    .action(async (file: string, options: { cell: string; sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        hyperlink: workbook.getSheet(options.sheet).getHyperlink(options.cell),
        sheet: options.sheet,
      });
    });

  hyperlinkCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address")
    .requiredOption("--target <value>", "hyperlink target")
    .option("--text <value>", "replace the cell text")
    .option("--tooltip <value>", "hyperlink tooltip")
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
          target: string;
          text?: string;
          tooltip?: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        sheet.setHyperlink(options.cell, options.target, {
          text: options.text,
          tooltip: options.tooltip,
        });
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.hyperlink.set",
          hyperlink: sheet.getHyperlink(options.cell),
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  hyperlinkCommand
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
        const existed = sheet.getHyperlink(options.cell) !== null;
        sheet.removeHyperlink(options.cell);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.hyperlink.delete",
          deleted: existed,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  hyperlinkCommand
    .command("clear")
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
        const sheet = workbook.getSheet(options.sheet);
        const cleared = sheet.getHyperlinks().length;
        sheet.clearHyperlinks();
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.hyperlink.clear",
          cleared,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  filterCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        autoFilter: workbook.getSheet(options.sheet).getAutoFilter(),
        sheet: options.sheet,
      });
    });

  filterCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--range <ref>", "filter range")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          range: string;
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
        sheet.setAutoFilter(options.range);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.filter.set",
          autoFilter: sheet.getAutoFilter(),
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  filterCommand
    .command("clear")
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
        const sheet = workbook.getSheet(options.sheet);
        sheet.clearAutoFilter();
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.filter.clear",
          autoFilter: null,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  selectionCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        selection: workbook.getSheet(options.sheet).getSelection(),
        sheet: options.sheet,
      });
    });

  selectionCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--active-cell <address>", "active cell address")
    .option("--range <sqref>", "selection range")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          activeCell: string;
          inPlace?: boolean;
          output?: string;
          range?: string;
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
        sheet.setSelection(options.activeCell, options.range ?? options.activeCell);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.selection.set",
          input: inputPath,
          output: outputPath,
          selection: sheet.getSelection(),
          sheet: options.sheet,
        });
      },
    );

  selectionCommand
    .command("clear")
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
        const sheet = workbook.getSheet(options.sheet);
        sheet.clearSelection();
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.selection.clear",
          input: inputPath,
          output: outputPath,
          selection: null,
          sheet: options.sheet,
        });
      },
    );

  validationCommand
    .command("list")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        sheet: options.sheet,
        validations: workbook.getSheet(options.sheet).getDataValidations(),
      });
    });

  validationCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--range <sqref>", "validation range")
    .action(async (file: string, options: { range: string; sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        sheet: options.sheet,
        validation: workbook.getSheet(options.sheet).getDataValidation(options.range),
      });
    });

  validationCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--range <sqref>", "validation range")
    .option("--type <value>", "validation type")
    .option("--operator <value>", "validation operator")
    .option("--allow-blank <value>", "allow blank values (true/false)", parseBooleanValue)
    .option("--show-input-message <value>", "show the input prompt (true/false)", parseBooleanValue)
    .option("--show-error-message <value>", "show the error message (true/false)", parseBooleanValue)
    .option("--show-drop-down <value>", "show the dropdown arrow (true/false)", parseBooleanValue)
    .option("--error-style <value>", "error style")
    .option("--error-title <value>", "error title")
    .option("--error <value>", "error message")
    .option("--prompt-title <value>", "prompt title")
    .option("--prompt <value>", "prompt message")
    .option("--ime-mode <value>", "IME mode")
    .option("--formula1 <value>", "first formula")
    .option("--formula2 <value>", "second formula")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          allowBlank?: boolean;
          error?: string;
          errorStyle?: string;
          errorTitle?: string;
          formula1?: string;
          formula2?: string;
          imeMode?: string;
          inPlace?: boolean;
          operator?: string;
          output?: string;
          prompt?: string;
          promptTitle?: string;
          range: string;
          sheet: string;
          showDropDown?: boolean;
          showErrorMessage?: boolean;
          showInputMessage?: boolean;
          type?: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        sheet.setDataValidation(options.range, buildValidationOptions(options));
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.validation.set",
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
          validation: sheet.getDataValidation(options.range),
        });
      },
    );

  validationCommand
    .command("delete")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--range <sqref>", "validation range")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          range: string;
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
        const deleted = sheet.getDataValidation(options.range) !== null;
        sheet.removeDataValidation(options.range);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.validation.delete",
          deleted,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  validationCommand
    .command("clear")
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
        const sheet = workbook.getSheet(options.sheet);
        const cleared = sheet.getDataValidations().length;
        sheet.clearDataValidations();
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.validation.clear",
          cleared,
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  mergeCommand
    .command("list")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        mergedRanges: workbook.getSheet(options.sheet).getMergedRanges(),
        sheet: options.sheet,
      });
    });

  mergeCommand
    .command("add")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--range <ref>", "merged range")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          range: string;
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
        sheet.addMergedRange(options.range);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.merge.add",
          input: inputPath,
          mergedRanges: sheet.getMergedRanges(),
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  mergeCommand
    .command("remove")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--range <ref>", "merged range")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          range: string;
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
        const beforeCount = sheet.getMergedRanges().length;
        sheet.removeMergedRange(options.range);
        const deleted = sheet.getMergedRanges().length !== beforeCount;
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.merge.remove",
          deleted,
          input: inputPath,
          mergedRanges: sheet.getMergedRanges(),
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  mergeCommand
    .command("clear")
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
        const sheet = workbook.getSheet(options.sheet);
        const cleared = sheet.getMergedRanges().length;
        sheet.clearMergedRanges();
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.merge.clear",
          cleared,
          input: inputPath,
          mergedRanges: [],
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  protectionCommand
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .action(async (file: string, options: { sheet: string }) => {
      const workbook = await Workbook.open(resolveFrom(io.cwd, file));
      writeJson(io.stdout, {
        protection: workbook.getSheet(options.sheet).getProtection(),
        sheet: options.sheet,
      });
    });

  protectionCommand
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--password-hash <hash>", "precomputed legacy password hash")
    .option("--auto-filter", "allow auto filter while protected")
    .option("--delete-columns", "allow deleting columns while protected")
    .option("--delete-rows", "allow deleting rows while protected")
    .option("--format-cells", "allow formatting cells while protected")
    .option("--format-columns", "allow formatting columns while protected")
    .option("--format-rows", "allow formatting rows while protected")
    .option("--insert-columns", "allow inserting columns while protected")
    .option("--insert-hyperlinks", "allow inserting hyperlinks while protected")
    .option("--insert-rows", "allow inserting rows while protected")
    .option("--objects", "allow editing objects while protected")
    .option("--pivot-tables", "allow pivot table usage while protected")
    .option("--scenarios", "allow editing scenarios while protected")
    .option("--sort", "allow sorting while protected")
    .option("--select-locked-cells", "allow selecting locked cells")
    .option("--select-unlocked-cells", "allow selecting unlocked cells")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          autoFilter?: boolean;
          deleteColumns?: boolean;
          deleteRows?: boolean;
          formatCells?: boolean;
          formatColumns?: boolean;
          formatRows?: boolean;
          inPlace?: boolean;
          insertColumns?: boolean;
          insertHyperlinks?: boolean;
          insertRows?: boolean;
          objects?: boolean;
          output?: string;
          passwordHash?: string;
          pivotTables?: boolean;
          scenarios?: boolean;
          selectLockedCells?: boolean;
          selectUnlockedCells?: boolean;
          sheet: string;
          sort?: boolean;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const protection = workbook.getSheet(options.sheet).protect({
          autoFilter: options.autoFilter,
          deleteColumns: options.deleteColumns,
          deleteRows: options.deleteRows,
          formatCells: options.formatCells,
          formatColumns: options.formatColumns,
          formatRows: options.formatRows,
          insertColumns: options.insertColumns,
          insertHyperlinks: options.insertHyperlinks,
          insertRows: options.insertRows,
          objects: options.objects,
          passwordHash: options.passwordHash,
          pivotTables: options.pivotTables,
          scenarios: options.scenarios,
          selectLockedCells: options.selectLockedCells,
          selectUnlockedCells: options.selectUnlockedCells,
          sort: options.sort,
        });
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.protection.set",
          input: inputPath,
          output: outputPath,
          protection,
          sheet: options.sheet,
        });
      },
    );

  protectionCommand
    .command("clear")
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
        workbook.getSheet(options.sheet).unprotect();
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "sheet.protection.clear",
          input: inputPath,
          output: outputPath,
          protection: null,
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

function parseCsvAsRecords(
  source: string,
  options: {
    inferTypes: boolean;
    trimHeaders: boolean;
    trimValues: boolean;
  },
): CellRecord[] {
  const rows = source.replace(/\r/g, "").split("\n");
  if (rows.at(-1) === "") {
    rows.pop();
  }
  if (rows.length === 0) {
    return [];
  }

  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");
  sheet.fromCsv(source, options);
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

function parseJsonNumberArray(source: string, label: string): number[] {
  const values = parseJsonDocument(source, label);
  if (!Array.isArray(values)) {
    throw new Error(`Expected ${label} to be an array`);
  }

  return values.map((value, index) => {
    if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
      throw new Error(`Expected ${label}[${index}] to be a positive integer`);
    }

    return value;
  });
}

function buildValidationOptions(options: {
  allowBlank?: boolean;
  error?: string;
  errorStyle?: string;
  errorTitle?: string;
  formula1?: string;
  formula2?: string;
  imeMode?: string;
  operator?: string;
  prompt?: string;
  promptTitle?: string;
  showDropDown?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  type?: string;
}): {
  allowBlank?: boolean;
  error?: string;
  errorStyle?: string;
  errorTitle?: string;
  formula1?: string;
  formula2?: string;
  imeMode?: string;
  operator?: string;
  prompt?: string;
  promptTitle?: string;
  showDropDown?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  type?: string;
} {
  return {
    allowBlank: options.allowBlank,
    error: options.error,
    errorStyle: options.errorStyle,
    errorTitle: options.errorTitle,
    formula1: options.formula1,
    formula2: options.formula2,
    imeMode: options.imeMode,
    operator: options.operator,
    prompt: options.prompt,
    promptTitle: options.promptTitle,
    showDropDown: options.showDropDown,
    showErrorMessage: options.showErrorMessage,
    showInputMessage: options.showInputMessage,
    type: options.type,
  };
}

function parseSheetRecordInputs(record?: string, records?: string): CellRecord[] {
  if (record && records) {
    throw new Error("Use either --record or --records, not both");
  }

  if (record) {
    return [parseJsonCellRecord(record, "--record")];
  }

  if (records) {
    return parseJsonCellRecordArray(records, "--records");
  }

  throw new Error("Pass --record or --records");
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
