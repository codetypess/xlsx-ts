import type { CellValue } from "./types.js";
import { XlsxError } from "./errors.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "./utils/xml-read.js";
import { escapeRegex, escapeXmlText, parseAttributes, serializeAttributes } from "./utils/xml.js";

export interface SheetTableReference {
  relationshipId: string;
  path: string;
}

export interface SheetTableMetadata {
  displayName: string;
  name: string;
  path: string;
  range: string;
}

export const TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

export const TABLE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

export function parseSheetTables(
  tableReferences: SheetTableReference[],
  readEntryText: (path: string) => string,
): SheetTableMetadata[] {
  const tables: SheetTableMetadata[] = [];

  for (const table of tableReferences) {
    const tableXml = readEntryText(table.path);
    const tableTag = findFirstXmlTag(tableXml, "table");
    if (!tableTag) {
      continue;
    }

    const name = getTagAttr(tableTag, "name");
    const displayName = getTagAttr(tableTag, "displayName");
    const range = getTagAttr(tableTag, "ref");

    if (!name || !displayName || !range) {
      continue;
    }

    tables.push({ name, displayName, range: normalizeRangeRef(range), path: table.path });
  }

  return tables;
}

export function findSheetTableReferenceByName(
  tableReferences: SheetTableReference[],
  readEntryText: (path: string) => string,
  name: string,
): SheetTableReference | null {
  return (
    tableReferences.find((table) => {
      const tableXml = readEntryText(table.path);
      const tableTag = findFirstXmlTag(tableXml, "table");
      if (!tableTag) {
        return false;
      }

      return getTagAttr(tableTag, "name") === name || getTagAttr(tableTag, "displayName") === name;
    }) ?? null
  );
}

export function getNextTablePath(entryPaths: string[]): string {
  let nextIndex = 1;

  for (const path of entryPaths) {
    const match = path.match(/^xl\/tables\/table(\d+)\.xml$/);
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `xl/tables/table${nextIndex}.xml`;
}

export function getNextTableId(
  entryPaths: string[],
  readEntryText: (path: string) => string,
): number {
  let nextId = 1;

  for (const path of entryPaths) {
    if (!/^xl\/tables\/table\d+\.xml$/.test(path)) {
      continue;
    }

    const tableXml = readEntryText(path);
    const tableTag = findFirstXmlTag(tableXml, "table");
    const idText = tableTag ? getTagAttr(tableTag, "id") : undefined;
    if (idText) {
      nextId = Math.max(nextId, Number(idText) + 1);
    }
  }

  return nextId;
}

export function getNextTableName(entryPaths: string[]): string {
  let nextIndex = 1;

  for (const path of entryPaths) {
    const match = path.match(/^xl\/tables\/table(\d+)\.xml$/);
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `Table${nextIndex}`;
}

export function assertTableName(name: string): void {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    throw new XlsxError(`Invalid table name: ${name}`);
  }
}

export function buildTableXml(
  range: string,
  id: number,
  name: string,
  headerValues: CellValue[],
): string {
  const parsedRange = parseRangeRef(range);
  const width = parsedRange.endColumn - parsedRange.startColumn + 1;
  const columnNames = buildTableColumnNames(headerValues, width);

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${escapeXmlText(name)}" displayName="${escapeXmlText(name)}" ref="${range}" totalsRowShown="0">` +
    `<autoFilter ref="${range}"/>` +
    `<tableColumns count="${columnNames.length}">` +
    columnNames
      .map((columnName, index) => `<tableColumn id="${index + 1}" name="${escapeXmlText(columnName)}"/>`)
      .join("") +
    `</tableColumns>` +
    `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>` +
    `</table>`
  );
}

export function getNextRelationshipIdFromXml(relationshipsXml: string): string {
  let nextId = 1;

  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    const relationshipId = getTagAttr(relationshipTag, "Id");
    if (!relationshipId?.startsWith("rId")) {
      continue;
    }

    nextId = Math.max(nextId, Number(relationshipId.slice(3)) + 1);
  }

  return `rId${nextId}`;
}

export function appendRelationship(
  relationshipsXml: string,
  relationshipId: string,
  type: string,
  target: string,
  targetMode?: string,
): string {
  const closingTag = "</Relationships>";
  const insertionIndex = relationshipsXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet relationships file is missing </Relationships>");
  }

  const attributes: Array<[string, string]> = [
    ["Id", relationshipId],
    ["Type", type],
    ["Target", target],
  ];
  if (targetMode) {
    attributes.push(["TargetMode", targetMode]);
  }

  const relationshipXml = `<Relationship ${serializeAttributes(attributes)}/>`;
  return relationshipsXml.slice(0, insertionIndex) + relationshipXml + relationshipsXml.slice(insertionIndex);
}

export function upsertRelationship(
  relationshipsXml: string,
  relationshipId: string,
  type: string,
  target: string,
  targetMode?: string,
): string {
  const nextRelationshipXml = buildRelationshipXml(relationshipId, type, target, targetMode);
  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    if (getTagAttr(relationshipTag, "Id") === relationshipId) {
      return replaceXmlTagSource(relationshipsXml, relationshipTag.source, nextRelationshipXml);
    }
  }

  return appendRelationship(relationshipsXml, relationshipId, type, target, targetMode);
}

export function removeRelationshipById(relationshipsXml: string, relationshipId: string): string {
  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    if (getTagAttr(relationshipTag, "Id") === relationshipId) {
      return replaceXmlTagSource(relationshipsXml, relationshipTag.source, "");
    }
  }

  return relationshipsXml;
}

export function makeRelativeSheetRelationshipTarget(sheetPath: string, targetPath: string): string {
  const fromParts = sheetPath.split("/").slice(0, -1).filter((part) => part.length > 0);
  const toParts = targetPath.split("/").filter((part) => part.length > 0);
  let commonLength = 0;

  while (
    commonLength < fromParts.length &&
    commonLength < toParts.length &&
    fromParts[commonLength] === toParts[commonLength]
  ) {
    commonLength += 1;
  }

  const upward = fromParts.slice(commonLength).map(() => "..");
  const downward = toParts.slice(commonLength);
  return [...upward, ...downward].join("/");
}

export function appendTablePart(sheetXml: string, relationshipId: string): string {
  const tablePartsTag = findFirstXmlTag(sheetXml, "tableParts");
  if (tablePartsTag && tablePartsTag.innerXml !== null) {
    const tableParts = findXmlTags(tablePartsTag.innerXml, "tablePart")
      .filter((tag) => tag.selfClosing)
      .map((tag) => tag.source);
    tableParts.push(`<tablePart r:id="${relationshipId}"/>`);
    const nextTablePartsXml = buildCountedXmlContainer("tableParts", tablePartsTag.attributesSource, "count", tableParts);
    return replaceXmlTagSource(sheetXml, tablePartsTag.source, nextTablePartsXml);
  }

  const closingTag = "</worksheet>";
  const insertionIndex = sheetXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return (
    sheetXml.slice(0, insertionIndex) +
    `<tableParts count="1"><tablePart r:id="${relationshipId}"/></tableParts>` +
    sheetXml.slice(insertionIndex)
  );
}

export function removeTablePartsFromSheetXml(sheetXml: string, relationshipIds: string[]): string {
  const tablePartsTag = findFirstXmlTag(sheetXml, "tableParts");
  if (!tablePartsTag || tablePartsTag.innerXml === null) {
    return sheetXml;
  }

  const keptTableParts = findXmlTags(tablePartsTag.innerXml, "tablePart")
    .map((tablePartTag) => ({
      relationshipId: getTagAttr(tablePartTag, "r:id"),
      xml: tablePartTag.source,
    }))
    .filter((tablePart) => tablePart.relationshipId && !relationshipIds.includes(tablePart.relationshipId));

  const nextTablePartsXml =
    keptTableParts.length === 0
      ? ""
      : buildCountedXmlContainer(
          "tableParts",
          tablePartsTag.attributesSource,
          "count",
          keptTableParts.map((tablePart) => tablePart.xml),
        );

  return replaceXmlTagSource(sheetXml, tablePartsTag.source, nextTablePartsXml);
}

export function addContentTypeOverride(contentTypesXml: string, partPath: string, contentType: string): string {
  if (new RegExp(`PartName\\s*=\\s*["']/${escapeRegex(partPath)}["']`).test(contentTypesXml)) {
    return contentTypesXml;
  }

  const closingTag = "</Types>";
  const insertionIndex = contentTypesXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Content types file is missing </Types>");
  }

  return (
    contentTypesXml.slice(0, insertionIndex) +
    `<Override PartName="/${escapeXmlText(partPath)}" ContentType="${escapeXmlText(contentType)}"/>` +
    contentTypesXml.slice(insertionIndex)
  );
}

export function removeContentTypeOverride(contentTypesXml: string, partPath: string): string {
  for (const overrideTag of findXmlTags(contentTypesXml, "Override")) {
    if (getTagAttr(overrideTag, "PartName") === `/${partPath}`) {
      return replaceXmlTagSource(contentTypesXml, overrideTag.source, "");
    }
  }

  return contentTypesXml;
}

function buildTableColumnNames(headerValues: CellValue[], width: number): string[] {
  const names: string[] = [];
  const seen = new Map<string, number>();

  for (let index = 0; index < width; index += 1) {
    const rawValue = headerValues[index];
    const baseName =
      typeof rawValue === "string" && rawValue.trim().length > 0 ? rawValue.trim() : `Column${index + 1}`;
    const nextCount = (seen.get(baseName) ?? 0) + 1;
    seen.set(baseName, nextCount);
    names.push(nextCount === 1 ? baseName : `${baseName}_${nextCount}`);
  }

  return names;
}

function buildRelationshipXml(
  relationshipId: string,
  type: string,
  target: string,
  targetMode?: string,
): string {
  const attributes: Array<[string, string]> = [
    ["Id", relationshipId],
    ["Type", type],
    ["Target", target],
  ];
  if (targetMode) {
    attributes.push(["TargetMode", targetMode]);
  }

  return `<Relationship ${serializeAttributes(attributes)}/>`;
}

function buildCountedXmlContainer(
  tagName: string,
  attributesSource: string,
  countAttributeName: string,
  childXml: string[],
): string {
  const attributes = parseAttributes(attributesSource);
  const nextAttributes = [...attributes];
  const countIndex = nextAttributes.findIndex(([name]) => name === countAttributeName);

  if (countIndex === -1) {
    nextAttributes.push([countAttributeName, String(childXml.length)]);
  } else {
    nextAttributes[countIndex] = [countAttributeName, String(childXml.length)];
  }

  const serializedAttributes = serializeAttributes(nextAttributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${childXml.join("")}</${tagName}>`;
}

export function findWorksheetChildInsertionIndex(sheetXml: string, followingTagNames: string[]): number {
  let insertionIndex = -1;

  for (const tagName of followingTagNames) {
    const match = sheetXml.match(new RegExp(`<${escapeRegex(tagName)}\\b`));
    if (!match || match.index === undefined) {
      continue;
    }

    if (insertionIndex === -1 || match.index < insertionIndex) {
      insertionIndex = match.index;
    }
  }

  if (insertionIndex !== -1) {
    return insertionIndex;
  }

  const closingTag = "</worksheet>";
  const closingTagIndex = sheetXml.indexOf(closingTag);
  if (closingTagIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return closingTagIndex;
}

function replaceXmlTagSource(xml: string, source: string, nextSource: string): string {
  const index = xml.indexOf(source);
  if (index === -1) {
    return xml;
  }

  return xml.slice(0, index) + nextSource + xml.slice(index + source.length);
}

function normalizeRangeRef(range: string): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  return formatRangeRef(startRow, startColumn, endRow, endColumn);
}

function parseRangeRef(range: string): {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
} {
  const normalizedRange = range.toUpperCase();
  const [startAddress, endAddress = normalizedRange] = normalizedRange.split(":");

  if (!startAddress || !endAddress) {
    throw new XlsxError(`Invalid range reference: ${range}`);
  }

  const start = splitCellAddress(startAddress);
  const end = splitCellAddress(endAddress);

  return {
    startRow: Math.min(start.rowNumber, end.rowNumber),
    endRow: Math.max(start.rowNumber, end.rowNumber),
    startColumn: Math.min(start.columnNumber, end.columnNumber),
    endColumn: Math.max(start.columnNumber, end.columnNumber),
  };
}

function splitCellAddress(address: string): { rowNumber: number; columnNumber: number } {
  assertCellAddress(address);
  const match = address.toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }

  return {
    columnNumber: columnLabelToNumber(match[1]),
    rowNumber: Number(match[2]),
  };
}

function assertCellAddress(address: string): void {
  if (!/^[A-Z]+[1-9]\d*$/i.test(address)) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }
}

function formatRangeRef(
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): string {
  const startAddress = makeCellAddress(startRow, startColumn);
  const endAddress = makeCellAddress(endRow, endColumn);
  return startAddress === endAddress ? startAddress : `${startAddress}:${endAddress}`;
}

function makeCellAddress(rowNumber: number, columnNumber: number): string {
  return `${numberToColumnLabel(columnNumber)}${rowNumber}`;
}

function compareCellAddresses(left: string, right: string): number {
  const leftCell = splitCellAddress(left);
  const rightCell = splitCellAddress(right);
  return leftCell.rowNumber - rightCell.rowNumber || leftCell.columnNumber - rightCell.columnNumber;
}

function columnLabelToNumber(label: string): number {
  let value = 0;

  for (const character of label) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }

  return value;
}

function numberToColumnLabel(columnNumber: number): string {
  if (!Number.isInteger(columnNumber) || columnNumber < 1) {
    throw new XlsxError(`Invalid column number: ${columnNumber}`);
  }

  let remaining = columnNumber;
  let label = "";

  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return label;
}
