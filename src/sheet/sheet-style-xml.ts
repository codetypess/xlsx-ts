import type { LocatedRow } from "./sheet-index.js";
import { buildXmlContainer, findWorksheetChildInsertionIndex, replaceXmlTagSource } from "./sheet-xml.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "../utils/xml-read.js";
import { getXmlAttr, parseAttributes, serializeAttributes } from "../utils/xml.js";

interface ColumnDefinition {
  min: number;
  max: number;
  attributes: Array<[string, string]>;
}

interface ColumnDefinitionPatch {
  hidden?: boolean | null;
  styleId?: number | null;
  width?: number | null;
}

interface RowDefinitionPatch {
  height?: number | null;
  hidden?: boolean | null;
  styleId?: number | null;
}

const COLS_FOLLOWING_TAGS = [
  "sheetData",
  "autoFilter",
  "sortState",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];

export function parseRowStyleId(attributesSource: string | undefined): number | null {
  if (!attributesSource) {
    return null;
  }

  const styleId = getXmlAttr(attributesSource, "s");
  return styleId === undefined ? null : Number(styleId);
}

export function parseColumnStyleId(sheetXml: string, columnNumber: number): number | null {
  let styleId: number | null = null;

  for (const definition of parseColumnDefinitions(sheetXml)) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      continue;
    }

    const styleText = getXmlAttr(serializeAttributes(definition.attributes), "style");
    styleId = styleText === undefined ? null : Number(styleText);
  }

  return styleId;
}

export function parseColumnHidden(sheetXml: string, columnNumber: number): boolean {
  let hidden = false;

  for (const definition of parseColumnDefinitions(sheetXml)) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      continue;
    }

    hidden = parseBooleanAttribute(definition.attributes, "hidden") ?? false;
  }

  return hidden;
}

export function parseColumnWidth(sheetXml: string, columnNumber: number): number | null {
  let width: number | null = null;

  for (const definition of parseColumnDefinitions(sheetXml)) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      continue;
    }

    width = parseNumericAttribute(definition.attributes, "width");
  }

  return width;
}

export function parseRowHidden(attributesSource: string | undefined): boolean {
  if (!attributesSource) {
    return false;
  }

  return parseBooleanAttribute(parseAttributes(attributesSource), "hidden") ?? false;
}

export function parseRowHeight(attributesSource: string | undefined): number | null {
  if (!attributesSource) {
    return null;
  }

  return parseNumericAttribute(parseAttributes(attributesSource), "ht");
}

export function buildStyledRowXml(sheetXml: string, row: LocatedRow, styleId: number | null): string {
  const serializedAttributes = serializeAttributes(buildRowAttributes(row.rowNumber, row.attributesSource, { styleId }));

  if (row.selfClosing) {
    return `<row ${serializedAttributes}/>`;
  }

  return `<row ${serializedAttributes}>${sheetXml.slice(row.innerStart, row.innerEnd)}</row>`;
}

export function buildEmptyStyledRowXml(rowNumber: number, styleId: number): string {
  return `<row ${serializeAttributes(buildRowAttributes(rowNumber, "", { styleId }))}/>`;
}

export function buildUpdatedRowXml(sheetXml: string, row: LocatedRow, patch: RowDefinitionPatch): string {
  const serializedAttributes = serializeAttributes(buildRowAttributes(row.rowNumber, row.attributesSource, patch));

  if (row.selfClosing) {
    return `<row ${serializedAttributes}/>`;
  }

  return `<row ${serializedAttributes}>${sheetXml.slice(row.innerStart, row.innerEnd)}</row>`;
}

export function buildEmptyRowXml(rowNumber: number, patch: RowDefinitionPatch): string | null {
  const attributes = buildRowAttributes(rowNumber, "", patch);
  if (attributes.length === 1) {
    return null;
  }

  return `<row ${serializeAttributes(attributes)}/>`;
}

export function updateColumnStyleIdInSheetXml(
  sheetXml: string,
  columnNumber: number,
  styleId: number | null,
): string {
  return updateColumnDefinitionInSheetXml(sheetXml, columnNumber, { styleId });
}

export function updateColumnHiddenInSheetXml(
  sheetXml: string,
  columnNumber: number,
  hidden: boolean,
): string {
  return updateColumnDefinitionInSheetXml(sheetXml, columnNumber, { hidden });
}

export function updateColumnWidthInSheetXml(
  sheetXml: string,
  columnNumber: number,
  width: number | null,
): string {
  return updateColumnDefinitionInSheetXml(sheetXml, columnNumber, { width });
}

export function transformColumnStyleDefinitions(
  sheetXml: string,
  targetColumnNumber: number,
  count: number,
  mode: "shift" | "delete",
): string {
  const existingDefinitions = parseColumnDefinitions(sheetXml);
  if (existingDefinitions.length === 0) {
    return sheetXml;
  }

  const nextDefinitions: ColumnDefinition[] = [];

  for (const definition of existingDefinitions) {
    if (mode === "shift") {
      if (definition.max < targetColumnNumber) {
        nextDefinitions.push(definition);
        continue;
      }

      if (definition.min >= targetColumnNumber) {
        nextDefinitions.push(buildColumnDefinition(definition.min + count, definition.max + count, definition.attributes));
        continue;
      }

      nextDefinitions.push(buildColumnDefinition(definition.min, targetColumnNumber - 1, definition.attributes));
      nextDefinitions.push(buildColumnDefinition(targetColumnNumber + count, definition.max + count, definition.attributes));
      continue;
    }

    const deleteEnd = targetColumnNumber + count - 1;
    if (definition.max < targetColumnNumber) {
      nextDefinitions.push(definition);
      continue;
    }

    if (definition.min > deleteEnd) {
      nextDefinitions.push(buildColumnDefinition(definition.min - count, definition.max - count, definition.attributes));
      continue;
    }

    if (definition.min < targetColumnNumber) {
      nextDefinitions.push(buildColumnDefinition(definition.min, targetColumnNumber - 1, definition.attributes));
    }

    if (definition.max > deleteEnd) {
      nextDefinitions.push(buildColumnDefinition(targetColumnNumber, definition.max - count, definition.attributes));
    }
  }

  return replaceColumnDefinitions(sheetXml, normalizeColumnDefinitions(nextDefinitions));
}

function buildRowAttributes(
  rowNumber: number,
  existingAttributesSource = "",
  patch: RowDefinitionPatch = {},
): Array<[string, string]> {
  const attributes = parseAttributes(existingAttributesSource);
  const preserved = attributes.filter(
    ([name]) =>
      name !== "r" &&
      name !== "s" &&
      name !== "customFormat" &&
      name !== "hidden" &&
      name !== "ht" &&
      name !== "customHeight",
  );
  const nextAttributes: Array<[string, string]> = [["r", String(rowNumber)]];

  if (patch.styleId !== undefined ? patch.styleId !== null : parseNumericAttribute(attributes, "s") !== null) {
    const styleId = patch.styleId ?? parseNumericAttribute(attributes, "s");
    if (styleId !== null) {
      nextAttributes.push(["s", String(styleId)], ["customFormat", "1"]);
    }
  }

  const hidden = patch.hidden ?? parseBooleanAttribute(attributes, "hidden") ?? false;
  if (hidden) {
    nextAttributes.push(["hidden", "1"]);
  }

  const height = patch.height !== undefined ? patch.height : parseNumericAttribute(attributes, "ht");
  if (height !== null) {
    nextAttributes.push(["ht", String(height)], ["customHeight", "1"]);
  }

  nextAttributes.push(...preserved);
  return nextAttributes;
}

function parseColumnDefinitions(sheetXml: string): ColumnDefinition[] {
  const colsTag = findFirstXmlTag(sheetXml, "cols");
  if (!colsTag || colsTag.innerXml === null) {
    return [];
  }

  return findXmlTags(colsTag.innerXml, "col")
    .map((colTag) => {
      const attributes = parseAttributes(colTag.attributesSource);
      const min = Number(getTagAttr(colTag, "min") ?? "0");
      const max = Number(getTagAttr(colTag, "max") ?? "0");

      return {
        min,
        max,
        attributes,
      };
    })
    .filter(
      (definition) =>
        Number.isInteger(definition.min) &&
        Number.isInteger(definition.max) &&
        definition.min > 0 &&
        definition.max >= definition.min,
    );
}

function replaceColumnDefinitions(sheetXml: string, definitions: ColumnDefinition[]): string {
  const colsTag = findFirstXmlTag(sheetXml, "cols");
  const colsXml =
    definitions.length === 0
      ? ""
      : buildXmlContainer(
          "cols",
          colsTag?.attributesSource ?? "",
          definitions.map((definition) => serializeColumnDefinition(definition)).join(""),
        );

  if (colsTag) {
    return replaceXmlTagSource(sheetXml, colsTag, colsXml);
  }

  if (definitions.length === 0) {
    return sheetXml;
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, COLS_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + colsXml + sheetXml.slice(insertionIndex);
}

function normalizeColumnDefinitions(definitions: ColumnDefinition[]): ColumnDefinition[] {
  const filtered = definitions
    .filter((definition) => definition.min <= definition.max)
    .sort((left, right) => left.min - right.min || left.max - right.max);
  const merged: ColumnDefinition[] = [];

  for (const definition of filtered) {
    const previous = merged.at(-1);
    if (
      previous &&
      previous.max + 1 === definition.min &&
      haveEquivalentColumnDefinitionAttributes(previous.attributes, definition.attributes)
    ) {
      previous.max = definition.max;
      continue;
    }

    merged.push({
      min: definition.min,
      max: definition.max,
      attributes: [...definition.attributes],
    });
  }

  return merged;
}

function buildColumnDefinition(
  min: number,
  max: number,
  existingAttributes: Array<[string, string]>,
): ColumnDefinition {
  const preserved = existingAttributes.filter(([name]) => name !== "min" && name !== "max");
  return {
    min,
    max,
    attributes: [["min", String(min)], ["max", String(max)], ...preserved],
  };
}

function buildColumnDefinitionWithPatch(
  min: number,
  max: number,
  existingAttributes: Array<[string, string]>,
  patch: ColumnDefinitionPatch,
): ColumnDefinition | null {
  const preserved = existingAttributes.filter(
    ([name]) =>
      name !== "min" &&
      name !== "max" &&
      name !== "style" &&
      name !== "hidden" &&
      name !== "width" &&
      name !== "customWidth",
  );
  const attributes: Array<[string, string]> = [["min", String(min)], ["max", String(max)]];

  const styleId = patch.styleId !== undefined ? patch.styleId : parseNumericAttribute(existingAttributes, "style");
  if (styleId !== null) {
    attributes.push(["style", String(styleId)]);
  }

  const hidden = patch.hidden !== undefined ? patch.hidden : parseBooleanAttribute(existingAttributes, "hidden");
  if (hidden) {
    attributes.push(["hidden", "1"]);
  }

  const width = patch.width !== undefined ? patch.width : parseNumericAttribute(existingAttributes, "width");
  if (width !== null) {
    attributes.push(["width", String(width)], ["customWidth", "1"]);
  }

  attributes.push(...preserved);

  if (attributes.length === 2) {
    return null;
  }

  return { min, max, attributes };
}

function updateColumnDefinitionInSheetXml(
  sheetXml: string,
  columnNumber: number,
  patch: ColumnDefinitionPatch,
): string {
  const existingDefinitions = parseColumnDefinitions(sheetXml);
  const nextDefinitions: ColumnDefinition[] = [];
  let handled = false;

  for (const definition of existingDefinitions) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      nextDefinitions.push(definition);
      continue;
    }

    handled = true;
    if (definition.min < columnNumber) {
      nextDefinitions.push(buildColumnDefinition(definition.min, columnNumber - 1, definition.attributes));
    }

    const nextDefinition = buildColumnDefinitionWithPatch(columnNumber, columnNumber, definition.attributes, patch);
    if (nextDefinition) {
      nextDefinitions.push(nextDefinition);
    }

    if (columnNumber < definition.max) {
      nextDefinitions.push(buildColumnDefinition(columnNumber + 1, definition.max, definition.attributes));
    }
  }

  if (!handled) {
    const nextDefinition = buildColumnDefinitionWithPatch(columnNumber, columnNumber, [], patch);
    if (nextDefinition) {
      nextDefinitions.push(nextDefinition);
    }
  }

  return replaceColumnDefinitions(sheetXml, normalizeColumnDefinitions(nextDefinitions));
}

function serializeColumnDefinition(definition: ColumnDefinition): string {
  const attributes = definition.attributes.map(([name, value]) => {
    if (name === "min") {
      return [name, String(definition.min)] as [string, string];
    }
    if (name === "max") {
      return [name, String(definition.max)] as [string, string];
    }
    return [name, value] as [string, string];
  });

  return `<col ${serializeAttributes(attributes)}/>`;
}

function haveEquivalentColumnDefinitionAttributes(
  left: Array<[string, string]>,
  right: Array<[string, string]>,
): boolean {
  const normalize = (attributes: Array<[string, string]>) =>
    serializeAttributes(attributes.filter(([name]) => name !== "min" && name !== "max"));

  return normalize(left) === normalize(right);
}

function parseBooleanAttribute(attributes: Array<[string, string]>, name: string): boolean | null {
  const value = attributes.find(([attributeName]) => attributeName === name)?.[1];
  if (value === undefined) {
    return null;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function parseNumericAttribute(attributes: Array<[string, string]>, name: string): number | null {
  const value = attributes.find(([attributeName]) => attributeName === name)?.[1];
  if (value === undefined) {
    return null;
  }

  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}
