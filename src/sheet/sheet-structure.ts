import { XlsxError } from "../errors.js";
import type { LocatedCell, LocatedRow } from "./sheet-index.js";
import {
  columnLabelToNumber,
  formatRangeRef,
  makeCellAddress,
  numberToColumnLabel,
  parseRangeRef,
  splitCellAddress,
} from "./sheet-address.js";
import { buildXmlElement, rewriteXmlTagsByName } from "./sheet-xml.js";
import { decodeXmlText, escapeRegex, escapeXmlText, parseAttributes, serializeAttributes } from "../utils/xml.js";

const WORKSHEET_REF_TAGS = ["autoFilter", "sortState", "hyperlink"];
const WORKSHEET_SQREF_TAGS = [
  "conditionalFormatting",
  "dataValidation",
  "selection",
  "protectedRange",
  "ignoredError",
];
const WORKSHEET_CELL_REF_ATTRIBUTES: Array<[string, string]> = [
  ["selection", "activeCell"],
  ["pane", "topLeftCell"],
];

export function transformRowXml(
  sheetXml: string,
  row: LocatedRow,
  sheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const rowAttributes = parseAttributes(row.attributesSource);
  const nextRowAttributes = rowAttributes.map(([name, value]) => {
    if (name === "r") {
      return [name, String(shiftRowNumber(Number(value), targetRowNumber, rowCount))] as [string, string];
    }

    if (name === "spans") {
      return [name, shiftRowSpans(value, targetColumnNumber, columnCount)] as [string, string];
    }

    return [name, value] as [string, string];
  });

  const rowOpenTag = `<row ${serializeAttributes(nextRowAttributes)}>`;
  let nextInnerXml = "";
  let cursor = row.innerStart;

  for (const cell of row.cells) {
    nextInnerXml += sheetXml.slice(cursor, cell.start);
    nextInnerXml += transformCellXml(
      sheetXml.slice(cell.start, cell.end),
      cell,
      sheetName,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );
    cursor = cell.end;
  }

  nextInnerXml += sheetXml.slice(cursor, row.innerEnd);
  return `${rowOpenTag}${nextInnerXml}</row>`;
}

export function deleteRowTransform(
  sheetXml: string,
  row: LocatedRow,
  sheetName: string,
  targetRowNumber: number,
  count: number,
): string {
  const nextRowNumber = deleteShiftRowNumber(row.rowNumber, targetRowNumber, count);
  const rowAttributes = parseAttributes(row.attributesSource)
    .filter(([name]) => name !== "r")
    .map(([name, value]) => [name, value] as [string, string]);
  const nextRowAttributes: Array<[string, string]> = [["r", String(nextRowNumber)], ...rowAttributes];

  if (row.selfClosing || row.cells.length === 0) {
    return `<row ${serializeAttributes(nextRowAttributes)}/>`;
  }

  const nextCells = row.cells.map((cell) =>
    deleteRowCellTransform(sheetXml.slice(cell.start, cell.end), cell, sheetName, targetRowNumber, count),
  );
  return `<row ${serializeAttributes(nextRowAttributes)}>${nextCells.join("")}</row>`;
}

export function deleteColumnTransform(
  sheetXml: string,
  row: LocatedRow,
  sheetName: string,
  targetColumnNumber: number,
  count: number,
): string {
  const keptCells = row.cells
    .filter((cell) => !isColumnDeleted(cell.columnNumber, targetColumnNumber, count))
    .map((cell) => ({
      columnNumber: deleteShiftColumnNumber(cell.columnNumber, targetColumnNumber, count),
      xml: deleteColumnCellTransform(sheetXml.slice(cell.start, cell.end), cell, sheetName, targetColumnNumber, count),
    }));

  const baseAttributes = parseAttributes(row.attributesSource)
    .filter(([name]) => name !== "spans")
    .map(([name, value]) => [name, value] as [string, string]);

  if (keptCells.length === 0) {
    return `<row ${serializeAttributes(baseAttributes)}/>`;
  }

  const nextAttributes = [...baseAttributes];
  const spansIndex = nextAttributes.findIndex(([name]) => name === "spans");
  const spansValue = `${keptCells[0].columnNumber}:${keptCells[keptCells.length - 1].columnNumber}`;

  if (spansIndex === -1) {
    nextAttributes.push(["spans", spansValue]);
  } else {
    nextAttributes[spansIndex] = ["spans", spansValue];
  }

  return `<row ${serializeAttributes(nextAttributes)}>${keptCells.map((cell) => cell.xml).join("")}</row>`;
}

export function shiftRangeRefColumns(range: string, targetColumnNumber: number, count: number): string {
  return shiftRangeRef(range, targetColumnNumber, count, 0, 0);
}

export function shiftRangeRefRows(range: string, targetRowNumber: number, count: number): string {
  return shiftRangeRef(range, 0, 0, targetRowNumber, count);
}

export function deleteRangeRefColumns(range: string, targetColumnNumber: number, count: number): string | null {
  return deleteRangeRef(range, targetColumnNumber, count, 0, 0);
}

export function deleteRangeRefRows(range: string, targetRowNumber: number, count: number): string | null {
  return deleteRangeRef(range, 0, 0, targetRowNumber, count);
}

export function shiftRangeRef(
  range: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);

  return formatRangeRef(
    shiftRowNumber(startRow, targetRowNumber, rowCount),
    shiftColumnNumber(startColumn, targetColumnNumber, columnCount),
    shiftRowNumber(endRow, targetRowNumber, rowCount),
    shiftColumnNumber(endColumn, targetColumnNumber, columnCount),
  );
}

export function deleteRangeRef(
  range: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string | null {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  const nextColumns = deleteRangeAxis(startColumn, endColumn, targetColumnNumber, columnCount);
  const nextRows = deleteRangeAxis(startRow, endRow, targetRowNumber, rowCount);

  if (!nextColumns || !nextRows) {
    return null;
  }

  return formatRangeRef(nextRows.start, nextColumns.start, nextRows.end, nextColumns.end);
}

export function transformWorksheetStructureReferences(
  sheetXml: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
  mode: "shift" | "delete",
): string {
  const transformRange = (range: string) =>
    mode === "shift"
      ? shiftRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount)
      : deleteRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount);

  let nextSheetXml = sheetXml;

  for (const tagName of WORKSHEET_REF_TAGS) {
    nextSheetXml = rewriteWorksheetReferenceTag(nextSheetXml, tagName, "ref", false, transformRange);
  }

  for (const tagName of WORKSHEET_SQREF_TAGS) {
    nextSheetXml = rewriteWorksheetReferenceTag(nextSheetXml, tagName, "sqref", true, transformRange);
  }

  for (const [tagName, attributeName] of WORKSHEET_CELL_REF_ATTRIBUTES) {
    nextSheetXml = rewriteWorksheetCellReferenceAttribute(
      nextSheetXml,
      tagName,
      attributeName,
      mode,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );
  }

  nextSheetXml = rewriteCountedContainer(nextSheetXml, "dataValidations", "dataValidation", "count");
  nextSheetXml = rewriteEmptyContainer(nextSheetXml, "hyperlinks", "hyperlink");
  return nextSheetXml;
}

export function shiftFormulaReferences(
  formula: string,
  currentSheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
  includeUnqualifiedReferences = true,
): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );
    const previous = formula[cursor - 1];

    if (rangeMatch) {
      const [
        fullMatch,
        startSheetRef,
        startColumnDollar,
        startColumnLabel,
        startRowDollar,
        startRowText,
        endSheetRef,
        endColumnDollar,
        endColumnLabel,
        endRowDollar,
        endRowText,
      ] = rangeMatch;

      if (
        !matchesFormulaReference(startSheetRef, currentSheetName, includeUnqualifiedReferences, previous) ||
        (endSheetRef !== undefined && !matchesSheetReference(endSheetRef, currentSheetName))
      ) {
        nextFormula += fullMatch;
        cursor += fullMatch.length;
        continue;
      }

      const nextStartColumn = shiftColumnNumber(
        columnLabelToNumber(startColumnLabel),
        targetColumnNumber,
        columnCount,
      );
      const nextEndColumn = shiftColumnNumber(
        columnLabelToNumber(endColumnLabel),
        targetColumnNumber,
        columnCount,
      );
      const nextStartRow = shiftRowNumber(Number(startRowText), targetRowNumber, rowCount);
      const nextEndRow = shiftRowNumber(Number(endRowText), targetRowNumber, rowCount);
      const leftRef = `${startSheetRef ?? ""}${startColumnDollar}${numberToColumnLabel(nextStartColumn)}${startRowDollar}${nextStartRow}`;
      const rightRef = `${endSheetRef ?? ""}${endColumnDollar}${numberToColumnLabel(nextEndColumn)}${endRowDollar}${nextEndRow}`;
      nextFormula += `${leftRef}:${rightRef}`;
      cursor += fullMatch.length;
      continue;
    }

    const match = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);

    if (!match) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef, columnDollar, columnLabel, rowDollar, rowNumber] = match;

    if (!matchesFormulaReference(sheetRef, currentSheetName, includeUnqualifiedReferences, previous)) {
      nextFormula += fullMatch;
      cursor += fullMatch.length;
      continue;
    }

    const columnNumber = columnLabelToNumber(columnLabel);
    const nextColumnNumber = shiftColumnNumber(columnNumber, targetColumnNumber, columnCount);
    const nextRowNumber = shiftRowNumber(Number(rowNumber), targetRowNumber, rowCount);
    nextFormula += `${sheetRef ?? ""}${columnDollar}${numberToColumnLabel(nextColumnNumber)}${rowDollar}${String(nextRowNumber)}`;
    cursor += fullMatch.length;
  }

  return nextFormula;
}

export function deleteFormulaReferences(
  formula: string,
  currentSheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
  includeUnqualifiedReferences = true,
): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );

    if (rangeMatch) {
      const [
        fullMatch,
        startSheetRef,
        startColumnDollar,
        startColumnLabel,
        startRowDollar,
        startRowText,
        endSheetRef,
        endColumnDollar,
        endColumnLabel,
        endRowDollar,
        endRowText,
      ] = rangeMatch;
      const previous = formula[cursor - 1];

      if (
        !matchesFormulaReference(
          startSheetRef,
          currentSheetName,
          includeUnqualifiedReferences,
          previous,
        ) ||
        (endSheetRef !== undefined && !matchesSheetReference(endSheetRef, currentSheetName))
      ) {
        nextFormula += fullMatch;
        cursor += fullMatch.length;
        continue;
      }

      const nextRange = deleteRangeAxis(
        columnLabelToNumber(startColumnLabel),
        columnLabelToNumber(endColumnLabel),
        targetColumnNumber,
        columnCount,
      );
      const nextRows = deleteRangeAxis(
        Number(startRowText),
        Number(endRowText),
        targetRowNumber,
        rowCount,
      );

      if (!nextRange || !nextRows) {
        nextFormula += "#REF!";
      } else {
        const leftRef = `${startSheetRef ?? ""}${startColumnDollar}${numberToColumnLabel(nextRange.start)}${startRowDollar}${nextRows.start}`;
        const rightPrefix = endSheetRef ?? "";
        const rightRef = `${rightPrefix}${endColumnDollar}${numberToColumnLabel(nextRange.end)}${endRowDollar}${nextRows.end}`;
        nextFormula += `${leftRef}:${rightRef}`;
      }

      cursor += fullMatch.length;
      continue;
    }

    const refMatch = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);
    if (!refMatch) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef, columnDollar, columnLabel, rowDollar, rowText] = refMatch;
    const previous = formula[cursor - 1];

    if (
      !matchesFormulaReference(
        sheetRef,
        currentSheetName,
        includeUnqualifiedReferences,
        previous,
      )
    ) {
      nextFormula += fullMatch;
      cursor += fullMatch.length;
      continue;
    }

    const columnNumber = columnLabelToNumber(columnLabel);
    const rowNumber = Number(rowText);

    if (isColumnDeleted(columnNumber, targetColumnNumber, columnCount) || isRowDeleted(rowNumber, targetRowNumber, rowCount)) {
      nextFormula += "#REF!";
    } else {
      nextFormula += `${sheetRef ?? ""}${columnDollar}${numberToColumnLabel(deleteShiftColumnNumber(columnNumber, targetColumnNumber, columnCount))}${rowDollar}${deleteShiftRowNumber(rowNumber, targetRowNumber, rowCount)}`;
    }

    cursor += fullMatch.length;
  }

  return nextFormula;
}

export function deleteSheetFormulaReferences(formula: string, deletedSheetName: string): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );

    if (rangeMatch) {
      const [fullMatch, startSheetRef, , , , , endSheetRef] = rangeMatch;

      if (
        matchesSheetReference(startSheetRef, deletedSheetName) ||
        matchesSheetReference(endSheetRef, deletedSheetName)
      ) {
        nextFormula += "#REF!";
      } else {
        nextFormula += fullMatch;
      }

      cursor += fullMatch.length;
      continue;
    }

    const refMatch = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);
    if (!refMatch) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef] = refMatch;

    nextFormula += matchesSheetReference(sheetRef, deletedSheetName) ? "#REF!" : fullMatch;
    cursor += fullMatch.length;
  }

  return nextFormula;
}

export function renameSheetFormulaReferences(
  formula: string,
  previousSheetName: string,
  nextSheetName: string,
): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );

    if (rangeMatch) {
      const [
        fullMatch,
        startSheetRef,
        startColumnDollar,
        startColumnLabel,
        startRowDollar,
        startRowText,
        endSheetRef,
        endColumnDollar,
        endColumnLabel,
        endRowDollar,
        endRowText,
      ] = rangeMatch;
      const nextStartSheetRef = renameSheetReferencePrefix(startSheetRef, previousSheetName, nextSheetName);
      const nextEndSheetRef = renameSheetReferencePrefix(endSheetRef, previousSheetName, nextSheetName);

      nextFormula +=
        nextStartSheetRef === startSheetRef && nextEndSheetRef === endSheetRef
          ? fullMatch
          : `${nextStartSheetRef ?? ""}${startColumnDollar}${startColumnLabel}${startRowDollar}${startRowText}:${nextEndSheetRef ?? ""}${endColumnDollar}${endColumnLabel}${endRowDollar}${endRowText}`;
      cursor += fullMatch.length;
      continue;
    }

    const refMatch = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);
    if (!refMatch) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef, columnDollar, columnLabel, rowDollar, rowText] = refMatch;
    const nextSheetRef = renameSheetReferencePrefix(sheetRef, previousSheetName, nextSheetName);

    nextFormula +=
      nextSheetRef === sheetRef
        ? fullMatch
        : `${nextSheetRef ?? ""}${columnDollar}${columnLabel}${rowDollar}${rowText}`;
    cursor += fullMatch.length;
  }

  return nextFormula;
}

function transformCellXml(
  cellXml: string,
  cell: LocatedCell,
  sheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const attributes = parseAttributes(cell.attributesSource);
  const nextAttributes = attributes.map(([name, value]) => {
    if (name === "r") {
      return [
        name,
        shiftCellAddress(
          value,
          targetColumnNumber,
          columnCount,
          targetRowNumber,
          rowCount,
        ),
      ] as [string, string];
    }

    return [name, value] as [string, string];
  });

  const nextCellOpenTag = `<c ${serializeAttributes(nextAttributes)}`;
  if (!cellXml.includes("</c>")) {
    return `${nextCellOpenTag}/>`;
  }

  const innerStart = cellXml.indexOf(">") + 1;
  const innerEnd = cellXml.lastIndexOf("</c>");
  let nextInnerXml = cellXml.slice(innerStart, innerEnd);

  nextInnerXml = rewriteXmlTagsByName(nextInnerXml, "f", (formulaTag) => {
    const formulaAttributes = parseAttributes(formulaTag.attributesSource);
    const nextFormulaAttributes = formulaAttributes.map(([name, value]) => {
      if (name === "ref") {
        return [
          name,
          shiftRangeRef(
            value,
            targetColumnNumber,
            columnCount,
            targetRowNumber,
            rowCount,
          ),
        ] as [string, string];
      }

      return [name, value] as [string, string];
    });
    const serializedAttributes = serializeAttributes(nextFormulaAttributes);
    const shiftedFormula = shiftFormulaReferences(
      decodeXmlText(formulaTag.innerXml ?? ""),
      sheetName,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );

    return `<f${serializedAttributes ? ` ${serializedAttributes}` : ""}>${escapeXmlText(shiftedFormula)}</f>`;
  });

  return `${nextCellOpenTag}>${nextInnerXml}</c>`;
}

function deleteRowCellTransform(
  cellXml: string,
  cell: LocatedCell,
  sheetName: string,
  targetRowNumber: number,
  count: number,
): string {
  const attributes = parseAttributes(cell.attributesSource).map(([name, value]) => {
    if (name === "r") {
      const { columnNumber, rowNumber } = splitCellAddress(value);
      return [name, makeCellAddress(deleteShiftRowNumber(rowNumber, targetRowNumber, count), columnNumber)] as [string, string];
    }

    return [name, value] as [string, string];
  });

  return deleteTransformCellInnerXml(cellXml, attributes, sheetName, 0, 0, targetRowNumber, count);
}

function deleteColumnCellTransform(
  cellXml: string,
  cell: LocatedCell,
  sheetName: string,
  targetColumnNumber: number,
  count: number,
): string {
  const attributes = parseAttributes(cell.attributesSource).map(([name, value]) => {
    if (name === "r") {
      const { rowNumber, columnNumber } = splitCellAddress(value);
      return [name, makeCellAddress(rowNumber, deleteShiftColumnNumber(columnNumber, targetColumnNumber, count))] as [string, string];
    }

    return [name, value] as [string, string];
  });

  return deleteTransformCellInnerXml(cellXml, attributes, sheetName, targetColumnNumber, count, 0, 0);
}

function deleteTransformCellInnerXml(
  cellXml: string,
  attributes: Array<[string, string]>,
  sheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const cellOpenTag = `<c ${serializeAttributes(attributes)}`;
  if (!cellXml.includes("</c>")) {
    return `${cellOpenTag}/>`;
  }

  const innerStart = cellXml.indexOf(">") + 1;
  const innerEnd = cellXml.lastIndexOf("</c>");
  let nextInnerXml = cellXml.slice(innerStart, innerEnd);

  nextInnerXml = rewriteXmlTagsByName(nextInnerXml, "f", (formulaTag) => {
    const formulaAttributes = parseAttributes(formulaTag.attributesSource);
    const nextFormulaAttributes = formulaAttributes.map(([name, value]) => {
      if (name === "ref") {
        const nextRange = deleteRangeRef(
          value,
          targetColumnNumber,
          columnCount,
          targetRowNumber,
          rowCount,
        );

        return nextRange === null ? [name, "#REF!"] as [string, string] : [name, nextRange] as [string, string];
      }

      return [name, value] as [string, string];
    });

    const nextFormula = deleteFormulaReferences(
      decodeXmlText(formulaTag.innerXml ?? ""),
      sheetName,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );
    const serializedAttributes = serializeAttributes(nextFormulaAttributes);
    return `<f${serializedAttributes ? ` ${serializedAttributes}` : ""}>${escapeXmlText(nextFormula)}</f>`;
  });

  return `${cellOpenTag}>${nextInnerXml}</c>`;
}

function shiftCellAddress(
  address: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const { rowNumber, columnNumber } = splitCellAddress(address);
  return makeCellAddress(
    shiftRowNumber(rowNumber, targetRowNumber, rowCount),
    shiftColumnNumber(columnNumber, targetColumnNumber, columnCount),
  );
}

function shiftColumnNumber(columnNumber: number, targetColumnNumber: number, count: number): number {
  if (targetColumnNumber <= 0 || count <= 0) {
    return columnNumber;
  }

  return columnNumber >= targetColumnNumber ? columnNumber + count : columnNumber;
}

function shiftRowNumber(rowNumber: number, targetRowNumber: number, count: number): number {
  if (targetRowNumber <= 0 || count <= 0) {
    return rowNumber;
  }

  return rowNumber >= targetRowNumber ? rowNumber + count : rowNumber;
}

function deleteShiftColumnNumber(columnNumber: number, targetColumnNumber: number, count: number): number {
  if (targetColumnNumber <= 0 || count <= 0) {
    return columnNumber;
  }

  return columnNumber > targetColumnNumber + count - 1 ? columnNumber - count : columnNumber;
}

function deleteShiftRowNumber(rowNumber: number, targetRowNumber: number, count: number): number {
  if (targetRowNumber <= 0 || count <= 0) {
    return rowNumber;
  }

  return rowNumber > targetRowNumber + count - 1 ? rowNumber - count : rowNumber;
}

function isColumnDeleted(columnNumber: number, targetColumnNumber: number, count: number): boolean {
  return targetColumnNumber > 0 && columnNumber >= targetColumnNumber && columnNumber <= targetColumnNumber + count - 1;
}

function isRowDeleted(rowNumber: number, targetRowNumber: number, count: number): boolean {
  return targetRowNumber > 0 && rowNumber >= targetRowNumber && rowNumber <= targetRowNumber + count - 1;
}

function deleteRangeAxis(
  start: number,
  end: number,
  target: number,
  count: number,
): { start: number; end: number } | null {
  if (target <= 0 || count <= 0) {
    return { start, end };
  }

  const deleteEnd = target + count - 1;

  if (end < target) {
    return { start, end };
  }

  if (start > deleteEnd) {
    return { start: start - count, end: end - count };
  }

  const hasLeft = start < target;
  const hasRight = end > deleteEnd;

  if (!hasLeft && !hasRight) {
    return null;
  }

  const nextStart = hasLeft ? start : target;
  const nextEnd = hasRight ? end - count : deleteEnd >= start ? target - 1 : end;

  if (nextStart > nextEnd) {
    return null;
  }

  return { start: nextStart, end: nextEnd };
}

function rewriteWorksheetReferenceTag(
  sheetXml: string,
  tagName: string,
  attributeName: string,
  multipleRanges: boolean,
  transformRange: (range: string) => string | null,
): string {
  const regex = new RegExp(
    `<${escapeRegex(tagName)}\\b([^>]*?)(\\/>|>[\\s\\S]*?<\\/${escapeRegex(tagName)}>)`,
    "g",
  );

  return sheetXml.replace(regex, (match, attributesSource, bodySource) => {
    const attributes = parseAttributes(attributesSource);
    const attributeIndex = attributes.findIndex(([name]) => name === attributeName);

    if (attributeIndex === -1) {
      return match;
    }

    const currentValue = attributes[attributeIndex]?.[1] ?? "";
    const nextValue = transformWorksheetReferenceValue(currentValue, multipleRanges, transformRange);

    if (nextValue === currentValue) {
      return match;
    }

    if (nextValue === null) {
      return "";
    }

    const nextAttributes = [...attributes];
    nextAttributes[attributeIndex] = [attributeName, nextValue];
    const serializedAttributes = serializeAttributes(nextAttributes);
    const tagOpen = serializedAttributes.length > 0 ? `<${tagName} ${serializedAttributes}` : `<${tagName}`;

    if (bodySource === "/>") {
      return `${tagOpen}/>`;
    }

    const closingTag = `</${tagName}>`;
    const innerXml = bodySource.slice(1, -closingTag.length);
    return `${tagOpen}>${innerXml}${closingTag}`;
  });
}

function transformWorksheetReferenceValue(
  value: string,
  multipleRanges: boolean,
  transformRange: (range: string) => string | null,
): string | null {
  if (multipleRanges) {
    const nextRanges = value
      .trim()
      .split(/\s+/)
      .filter((range) => range.length > 0)
      .map((range) => transformRange(range))
      .filter((range): range is string => range !== null);

    return nextRanges.length > 0 ? nextRanges.join(" ") : null;
  }

  return transformRange(value);
}

function rewriteCountedContainer(
  sheetXml: string,
  containerTagName: string,
  childTagName: string,
  countAttributeName: string,
): string {
  const regex = new RegExp(
    `<${escapeRegex(containerTagName)}\\b([^>]*)>([\\s\\S]*?)<\\/${escapeRegex(containerTagName)}>`,
    "g",
  );

  return sheetXml.replace(regex, (_match, attributesSource, innerXml) => {
    const childMatches = innerXml.match(new RegExp(`<${escapeRegex(childTagName)}\\b`, "g")) ?? [];
    if (childMatches.length === 0) {
      return "";
    }

    const attributes = parseAttributes(attributesSource);
    const countIndex = attributes.findIndex(([name]) => name === countAttributeName);
    const nextAttributes = [...attributes];

    if (countIndex === -1) {
      nextAttributes.push([countAttributeName, String(childMatches.length)]);
    } else {
      nextAttributes[countIndex] = [countAttributeName, String(childMatches.length)];
    }

    const serializedAttributes = serializeAttributes(nextAttributes);
    return `<${containerTagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${innerXml}</${containerTagName}>`;
  });
}

function rewriteEmptyContainer(
  sheetXml: string,
  containerTagName: string,
  childTagName: string,
): string {
  const regex = new RegExp(
    `<${escapeRegex(containerTagName)}\\b([^>]*)>([\\s\\S]*?)<\\/${escapeRegex(containerTagName)}>`,
    "g",
  );

  return sheetXml.replace(regex, (match, _attributesSource, innerXml) => {
    return new RegExp(`<${escapeRegex(childTagName)}\\b`).test(innerXml) ? match : "";
  });
}

function rewriteWorksheetCellReferenceAttribute(
  sheetXml: string,
  tagName: string,
  attributeName: string,
  mode: "shift" | "delete",
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const regex = new RegExp(
    `<${escapeRegex(tagName)}\\b([^>]*?)(\\/?>|>[\\s\\S]*?<\\/${escapeRegex(tagName)}>)`,
    "g",
  );

  return sheetXml.replace(regex, (match, attributesSource, bodySource) => {
    const attributes = parseAttributes(attributesSource);
    const attributeIndex = attributes.findIndex(([name]) => name === attributeName);

    if (attributeIndex === -1) {
      return match;
    }

    const currentValue = attributes[attributeIndex]?.[1] ?? "";
    const nextValue =
      mode === "shift"
        ? shiftCellAddress(
            currentValue,
            targetColumnNumber,
            columnCount,
            targetRowNumber,
            rowCount,
          )
        : deleteCellReferenceAddress(
            currentValue,
            targetColumnNumber,
            columnCount,
            targetRowNumber,
            rowCount,
          );

    if (nextValue === currentValue) {
      return match;
    }

    const nextAttributes = [...attributes];
    if (nextValue === null) {
      nextAttributes.splice(attributeIndex, 1);
    } else {
      nextAttributes[attributeIndex] = [attributeName, nextValue];
    }

    const serializedAttributes = serializeAttributes(nextAttributes);
    const tagOpen = serializedAttributes.length > 0 ? `<${tagName} ${serializedAttributes}` : `<${tagName}`;

    if (bodySource === "/>") {
      return `${tagOpen}/>`;
    }

    const closingTag = `</${tagName}>`;
    const innerXml = bodySource.slice(1, -closingTag.length);
    return `${tagOpen}>${innerXml}${closingTag}`;
  });
}

function deleteCellReferenceAddress(
  address: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string | null {
  const { rowNumber, columnNumber } = splitCellAddress(address);

  if (isColumnDeleted(columnNumber, targetColumnNumber, columnCount) || isRowDeleted(rowNumber, targetRowNumber, rowCount)) {
    return null;
  }

  return makeCellAddress(
    deleteShiftRowNumber(rowNumber, targetRowNumber, rowCount),
    deleteShiftColumnNumber(columnNumber, targetColumnNumber, columnCount),
  );
}

function shiftRowSpans(spans: string, targetColumnNumber: number, count: number): string {
  const match = spans.match(/^(\d+):(\d+)$/);
  if (!match) {
    return spans;
  }

  const startColumn = Number(match[1]);
  const endColumn = Number(match[2]);
  return `${startColumn >= targetColumnNumber ? startColumn + count : startColumn}:${endColumn >= targetColumnNumber ? endColumn + count : endColumn}`;
}

function matchesFormulaReference(
  sheetRef: string | undefined,
  targetSheetName: string,
  includeUnqualifiedReferences: boolean,
  previousCharacter: string | undefined,
): boolean {
  if (!sheetRef) {
    return includeUnqualifiedReferences && !(previousCharacter && /[A-Za-z0-9_.]/.test(previousCharacter));
  }

  return matchesSheetReference(sheetRef, targetSheetName);
}

function matchesSheetReference(sheetRef: string | undefined, targetSheetName: string): boolean {
  if (!sheetRef) {
    return false;
  }

  const rawSheetName = sheetRef.slice(0, -1);
  const normalizedSheetName =
    rawSheetName.startsWith("'") && rawSheetName.endsWith("'")
      ? rawSheetName.slice(1, -1).replaceAll("''", "'")
      : rawSheetName;

  return normalizeSheetNameKey(normalizedSheetName) === normalizeSheetNameKey(targetSheetName);
}

function renameSheetReferencePrefix(
  sheetRef: string | undefined,
  previousSheetName: string,
  nextSheetName: string,
): string | undefined {
  if (!matchesSheetReference(sheetRef, previousSheetName)) {
    return sheetRef;
  }

  return `${formatSheetReference(nextSheetName)}!`;
}

function formatSheetReference(sheetName: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) {
    return sheetName;
  }

  return `'${sheetName.replaceAll("'", "''")}'`;
}

function normalizeSheetNameKey(sheetName: string): string {
  return sheetName.toUpperCase();
}
