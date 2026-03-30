import { XlsxError } from "../errors.js";
import type { CellEntry } from "../types.js";
import type { LocatedCell } from "./sheet-index.js";
import { makeCellAddress, normalizeCellAddress, normalizeColumnNumber } from "./sheet-address.js";
import { assertRowNumber } from "./sheet-validation.js";

export function resolveCellAddress(addressOrRowNumber: string | number, column?: number | string): string {
  if (typeof addressOrRowNumber === "string") {
    if (column !== undefined) {
      throw new XlsxError("Column argument is not allowed when address is a string");
    }

    return normalizeCellAddress(addressOrRowNumber);
  }

  assertRowNumber(addressOrRowNumber);
  if (column === undefined) {
    throw new XlsxError(`Missing column index for row: ${addressOrRowNumber}`);
  }

  return makeCellAddress(addressOrRowNumber, normalizeColumnNumber(column));
}

export function createCellEntry(cell: LocatedCell): CellEntry {
  return {
    address: cell.address,
    rowNumber: cell.rowNumber,
    columnNumber: cell.columnNumber,
    ...cell.snapshot,
  };
}

export function normalizeEmptyRowXml(rowXml: string): string {
  return rowXml.replace(/>\s*<\/row>$/, "></row>");
}

export function resolveCopyStyleArguments(
  sourceAddressOrRowNumber: string | number,
  sourceColumnOrTargetAddress: number | string,
  targetRowNumber?: number,
  targetColumn?: number | string,
): { sourceAddress: string; targetAddress: string } {
  if (typeof sourceAddressOrRowNumber === "string") {
    if (typeof sourceColumnOrTargetAddress !== "string") {
      throw new XlsxError("Missing target address for copyStyle");
    }

    return {
      sourceAddress: resolveCellAddress(sourceAddressOrRowNumber),
      targetAddress: resolveCellAddress(sourceColumnOrTargetAddress),
    };
  }

  if (targetRowNumber === undefined || targetColumn === undefined) {
    throw new XlsxError("Missing target row or column for copyStyle");
  }

  return {
    sourceAddress: resolveCellAddress(sourceAddressOrRowNumber, sourceColumnOrTargetAddress),
    targetAddress: resolveCellAddress(targetRowNumber, targetColumn),
  };
}
