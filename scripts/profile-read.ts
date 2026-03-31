import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook } from "../src/index.js";

interface SheetProfile {
  name: string;
  rowCount: number;
  columnCount: number;
  usedRange: string | null;
  physicalCellCount: number;
  logicalUsedCellCount: number;
  blankPlaceholderCellCount: number;
  maxPhysicalColumn: number;
  denseReadCount: number;
  denseAmplification: number;
  indexMs: number;
  denseReadMs: number;
  entryReadMs: number;
}

async function main(): Promise<void> {
  const filePath = resolve(process.cwd(), process.argv[2] ?? "res/event.xlsx");
  const openedAt = performance.now();
  const workbook = await Workbook.open(filePath);
  const openMs = Number((performance.now() - openedAt).toFixed(1));

  const sheets: SheetProfile[] = [];

  for (const sheet of workbook.getSheets()) {
    const indexedAt = performance.now();
    const entries = sheet.getPhysicalCellEntries();
    const indexMs = Number((performance.now() - indexedAt).toFixed(1));

    const logicalEntries = sheet.getCellEntries();
    const logicalUsedCellCount = logicalEntries.length;
    const blankPlaceholderCellCount = entries.length - logicalUsedCellCount;
    const maxPhysicalColumn = entries.reduce(
      (currentMax, entry) => Math.max(currentMax, entry.columnNumber),
      0,
    );
    const denseReadCount = sheet.rowCount * sheet.columnCount;

    const denseReadStartedAt = performance.now();
    let denseNonNullCount = 0;
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber += 1) {
        if (sheet.getCell(rowNumber, columnNumber) !== null) {
          denseNonNullCount += 1;
        }
      }
    }
    const denseReadMs = Number((performance.now() - denseReadStartedAt).toFixed(1));

    const entryReadStartedAt = performance.now();
    let entryNonNullCount = 0;
    for (const entry of entries) {
      if (entry.value !== null || entry.formula !== null) {
        entryNonNullCount += 1;
      }
    }
    const entryReadMs = Number((performance.now() - entryReadStartedAt).toFixed(1));

    if (denseNonNullCount !== logicalUsedCellCount || entryNonNullCount !== logicalUsedCellCount) {
      throw new Error(
        `Profile mismatch in sheet ${sheet.name}: dense=${denseNonNullCount}, entries=${entryNonNullCount}, logical=${logicalUsedCellCount}`,
      );
    }

    sheets.push({
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      usedRange: sheet.getRangeRef(),
      physicalCellCount: entries.length,
      logicalUsedCellCount,
      blankPlaceholderCellCount,
      maxPhysicalColumn,
      denseReadCount,
      denseAmplification:
        logicalUsedCellCount === 0 ? 0 : Number((denseReadCount / logicalUsedCellCount).toFixed(2)),
      indexMs,
      denseReadMs,
      entryReadMs,
    });
  }

  console.log(JSON.stringify({ file: filePath, openMs, sheets }, null, 2));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}
