import { XlsxError } from "../errors.js";

export function assertFreezeSplit(columnCount: number, rowCount: number): void {
  if (!Number.isInteger(columnCount) || columnCount < 0) {
    throw new XlsxError(`Invalid freeze column count: ${columnCount}`);
  }

  if (!Number.isInteger(rowCount) || rowCount < 0) {
    throw new XlsxError(`Invalid freeze row count: ${rowCount}`);
  }

  if (columnCount === 0 && rowCount === 0) {
    throw new XlsxError("Freeze pane requires at least one frozen row or column");
  }
}

export function assertRowNumber(rowNumber: number): void {
  if (!Number.isInteger(rowNumber) || rowNumber < 1) {
    throw new XlsxError(`Invalid row number: ${rowNumber}`);
  }
}

export function assertColumnNumber(columnNumber: number): void {
  if (!Number.isInteger(columnNumber) || columnNumber < 1) {
    throw new XlsxError(`Invalid column number: ${columnNumber}`);
  }
}

export function assertInsertCount(count: number): void {
  if (!Number.isInteger(count) || count < 1) {
    throw new XlsxError(`Invalid insert count: ${count}`);
  }
}
