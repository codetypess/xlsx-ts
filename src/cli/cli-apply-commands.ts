import { readFile } from "node:fs/promises";

import { Command } from "commander";

import {
  assertArray,
  assertCellRecord,
  assertCellRecordArray,
  assertCellValue,
  assertNullableString,
  assertPositiveInteger,
  assertPositiveIntegerArray,
  assertRecord,
  assertSheetVisibility,
  assertString,
  assertStringArray,
  optionalPositiveInteger,
  optionalString,
  parseJsonDocument,
  writeJson,
} from "./cli-json.js";
import type { CellRecord } from "./cli-json.js";
import { resolveFrom, resolveOutputPath } from "./cli-shared.js";
import type { CliCommandIo } from "./cli-shared.js";
import type { CellValue, SheetVisibility } from "../types.js";
import { Workbook } from "../workbook.js";

type WorkbookOperation =
  | {
      headerRow?: number;
      record: CellRecord;
      sheet: string;
      type: "addRecord";
    }
  | {
      headerRow?: number;
      records: CellRecord[];
      sheet: string;
      type: "addRecords";
    }
  | {
      cell: string;
      color: string | null;
      sheet: string;
      type: "setBackgroundColor";
    }
  | {
      from: string;
      sheet: string;
      to: string;
      type: "copyStyle";
    }
  | {
      cell: string;
      sheet: string;
      type: "clearCell";
    }
  | {
      headerRow?: number;
      row: number;
      sheet: string;
      type: "deleteRecord";
    }
  | {
      headerRow?: number;
      rows: number[];
      sheet: string;
      type: "deleteRecords";
    }
  | {
      name: string;
      scope?: string;
      type: "deleteDefinedName";
    }
  | {
      sheet: string;
      type: "addSheet";
    }
  | {
      sheet: string;
      type: "deleteSheet";
    }
  | {
      from: string;
      to: string;
      type: "renameSheet";
    }
  | {
      headerRow?: number;
      record: CellRecord;
      row: number;
      sheet: string;
      type: "setRecord";
    }
  | {
      headerRow?: number;
      records: CellRecord[];
      sheet: string;
      type: "setRecords";
    }
  | {
      cachedValue?: CellValue;
      cell: string;
      formula: string;
      sheet: string;
      type: "setFormula";
    }
  | {
      cell: string;
      formatCode: string;
      sheet: string;
      type: "setNumberFormat";
    }
  | {
      cell: string;
      sheet: string;
      type: "setCell";
      value: CellValue;
    }
  | {
      sheet: string;
      type: "setActiveSheet";
    }
  | {
      headerRow?: number;
      headers: string[];
      sheet: string;
      startColumn?: number;
      type: "setHeaders";
    }
  | {
      name: string;
      scope?: string;
      type: "setDefinedName";
      value: string;
    }
  | {
      sheet: string;
      type: "setSheetVisibility";
      visibility: SheetVisibility;
    };

interface OpsDocument {
  actions: WorkbookOperation[];
  output?: string;
}

export function registerApplyCommands(
  program: Command,
  io: CliCommandIo,
): void {
  program
    .command("apply")
    .argument("<file>", "input xlsx file")
    .requiredOption("--ops <file>", "JSON document with workbook actions")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          ops: string;
          output?: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const opsPath = resolveFrom(io.cwd, options.ops);
        const document = await readOpsDocument(opsPath);
        // Keep CLI precedence explicit: `--output` wins over any embedded document output,
        // and both override the default "must choose output or in-place" guard.
        const configuredOutput = document.output ? resolveFrom(io.cwd, document.output) : undefined;
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : configuredOutput,
        });
        const workbook = await Workbook.open(inputPath);

        // Apply actions in source order so ops documents can express dependent steps,
        // such as creating or renaming a sheet before later cell edits target it.
        for (const action of document.actions) {
          applyWorkbookOperation(workbook, action);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          actions: document.actions.length,
          input: inputPath,
          ops: opsPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );
}

async function readOpsDocument(filePath: string): Promise<OpsDocument> {
  const parsed = parseJsonDocument(await readFile(filePath, "utf8"), filePath);

  // Support a compact array form for simple batches and an object form when callers
  // also need to embed output metadata alongside the action list.
  if (Array.isArray(parsed)) {
    return {
      actions: parsed.map((item, index) => parseWorkbookOperation(item, `${filePath}[${index}]`)),
    };
  }

  const record = assertRecord(parsed, filePath);
  const actions = assertArray(record.actions, `${filePath}.actions`);
  return {
    actions: actions.map((item, index) => parseWorkbookOperation(item, `${filePath}.actions[${index}]`)),
    output: record.output === undefined ? undefined : assertString(record.output, `${filePath}.output`),
  };
}

function parseWorkbookOperation(value: unknown, label: string): WorkbookOperation {
  const record = assertRecord(value, label);
  const type = assertString(record.type, `${label}.type`);

  switch (type) {
    case "addRecord":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        record: assertCellRecord(record.record, `${label}.record`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "addRecords":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        records: assertCellRecordArray(record.records, `${label}.records`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "addSheet":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "copyStyle":
      return {
        from: assertString(record.from, `${label}.from`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        to: assertString(record.to, `${label}.to`),
        type,
      };
    case "clearCell":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "deleteRecord":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        row: assertPositiveInteger(record.row, `${label}.row`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "deleteRecords":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        rows: assertPositiveIntegerArray(record.rows, `${label}.rows`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "deleteDefinedName":
      return {
        name: assertString(record.name, `${label}.name`),
        scope: optionalString(record.scope, `${label}.scope`),
        type,
      };
    case "deleteSheet":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "renameSheet":
      return {
        from: assertString(record.from, `${label}.from`),
        to: assertString(record.to, `${label}.to`),
        type,
      };
    case "setBackgroundColor":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        color: assertNullableString(record.color, `${label}.color`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setHeaders":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        headers: assertStringArray(record.headers, `${label}.headers`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        startColumn: optionalPositiveInteger(record.startColumn, `${label}.startColumn`),
        type,
      };
    case "setNumberFormat":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        formatCode: assertString(record.formatCode, `${label}.formatCode`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setRecord":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        record: assertCellRecord(record.record, `${label}.record`),
        row: assertPositiveInteger(record.row, `${label}.row`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setRecords":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        records: assertCellRecordArray(record.records, `${label}.records`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setActiveSheet":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setCell":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
        value: assertCellValue(record.value, `${label}.value`),
      };
    case "setDefinedName":
      return {
        name: assertString(record.name, `${label}.name`),
        scope: optionalString(record.scope, `${label}.scope`),
        type,
        value: assertString(record.value, `${label}.value`),
      };
    case "setFormula":
      return {
        cachedValue:
          record.cachedValue === undefined
            ? undefined
            : assertCellValue(record.cachedValue, `${label}.cachedValue`),
        cell: assertString(record.cell, `${label}.cell`),
        formula: assertString(record.formula, `${label}.formula`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setSheetVisibility":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
        visibility: assertSheetVisibility(record.visibility, `${label}.visibility`),
      };
    default:
      throw new Error(`Unsupported operation type at ${label}.type: ${type}`);
  }
}

function applyWorkbookOperation(workbook: Workbook, action: WorkbookOperation): void {
  switch (action.type) {
    case "addRecord":
      workbook.getSheet(action.sheet).addRecord(action.record, action.headerRow ?? 1);
      return;
    case "addRecords":
      workbook.getSheet(action.sheet).addRecords(action.records, action.headerRow ?? 1);
      return;
    case "addSheet":
      workbook.addSheet(action.sheet);
      return;
    case "copyStyle":
      workbook.getSheet(action.sheet).copyStyle(action.from, action.to);
      return;
    case "clearCell":
      workbook.getSheet(action.sheet).deleteCell(action.cell);
      return;
    case "deleteRecord":
      workbook.getSheet(action.sheet).deleteRecord(action.row, action.headerRow ?? 1);
      return;
    case "deleteRecords":
      workbook.getSheet(action.sheet).deleteRecords(action.rows, action.headerRow ?? 1);
      return;
    case "deleteDefinedName":
      workbook.deleteDefinedName(action.name, action.scope);
      return;
    case "deleteSheet":
      workbook.deleteSheet(action.sheet);
      return;
    case "renameSheet":
      workbook.renameSheet(action.from, action.to);
      return;
    case "setBackgroundColor":
      workbook.getSheet(action.sheet).setBackgroundColor(action.cell, action.color);
      return;
    case "setHeaders":
      workbook.getSheet(action.sheet).setHeaders(
        action.headers,
        action.headerRow ?? 1,
        action.startColumn ?? 1,
      );
      return;
    case "setRecord":
      workbook.getSheet(action.sheet).setRecord(action.row, action.record, action.headerRow ?? 1);
      return;
    case "setRecords":
      // `setRecords` is intentionally destructive for the addressed record region:
      // it rewrites the table snapshot rather than merging field-by-field.
      workbook.getSheet(action.sheet).setRecords(action.records, action.headerRow ?? 1);
      return;
    case "setActiveSheet":
      workbook.setActiveSheet(action.sheet);
      return;
    case "setCell":
      workbook.getSheet(action.sheet).setCell(action.cell, action.value);
      return;
    case "setDefinedName":
      workbook.setDefinedName(action.name, action.value, action.scope ? { scope: action.scope } : {});
      return;
    case "setFormula":
      workbook
        .getSheet(action.sheet)
        .setFormula(
          action.cell,
          action.formula,
          action.cachedValue === undefined ? {} : { cachedValue: action.cachedValue },
        );
      return;
    case "setNumberFormat":
      workbook.getSheet(action.sheet).setNumberFormat(action.cell, action.formatCode);
      return;
    case "setSheetVisibility":
      workbook.setSheetVisibility(action.sheet, action.visibility);
      return;
  }
}
