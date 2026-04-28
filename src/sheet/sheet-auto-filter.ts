import { XlsxError } from "../errors.js";
import type {
  AutoFilterColumn,
  AutoFilterCondition,
  AutoFilterDefinition,
  DateGroupItem,
  SortStateDefinition,
} from "../types.js";
import { findFirstXmlTag, getTagAttr, type XmlTag } from "../utils/xml-read.js";
import { parseAttributes } from "../utils/xml.js";
import { formatRangeRef, normalizeRangeRef, parseRangeRef } from "./sheet-address.js";
import {
  buildXmlElement,
  buildSelfClosingXmlElement,
  findWorksheetChildInsertionIndex,
  getXmlTagInnerStart,
  replaceXmlTagSource,
} from "./sheet-xml.js";

const WORKSHEET_AUTO_FILTER_FOLLOWING_TAGS = [
  "sortState",
  "dataConsolidate",
  "customSheetViews",
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
  "drawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
] as const;

interface ParsedAutoFilterState {
  range: string;
  autoFilterAttributes: Array<[string, string]>;
  rootItems: AutoFilterRootItem[];
  sortState: SortStateDefinition | null;
  sortStateLocation: "autoFilter" | "external" | null;
}

type AutoFilterRootItem =
  | {
      kind: "column";
      columnNumber: number;
      definition: AutoFilterColumn | null;
      source: string;
    }
  | {
      kind: "sortState";
      source: string;
    }
  | {
      kind: "raw";
      source: string;
    };

export function parseWorksheetAutoFilterDefinition(sheetXml: string): AutoFilterDefinition | null {
  return parseAutoFilterDefinitionFromTags(
    findFirstXmlTag(sheetXml, "autoFilter"),
    findFirstXmlTag(sheetXml, "sortState"),
  );
}

export function parseTableAutoFilterDefinition(tableXml: string): AutoFilterDefinition | null {
  return parseAutoFilterDefinitionFromTags(findFirstXmlTag(tableXml, "autoFilter"), null);
}

export function setWorksheetAutoFilterDefinitionInSheetXml(
  sheetXml: string,
  definition: AutoFilterDefinition,
): string {
  const normalizedDefinition = normalizeAutoFilterDefinition(definition);
  const currentState = parseAutoFilterStateFromTags(
    findFirstXmlTag(sheetXml, "autoFilter"),
    findFirstXmlTag(sheetXml, "sortState"),
  );
  const nextRootItems = buildRootItemsForFullDefinition(currentState, normalizedDefinition);
  const nextAutoFilterXml = buildAutoFilterXml(
    normalizedDefinition.range,
    nextRootItems,
    currentState?.autoFilterAttributes ?? [],
  );
  let nextSheetXml = upsertWorksheetAutoFilterXml(sheetXml, nextAutoFilterXml);
  nextSheetXml = upsertWorksheetSortStateXml(nextSheetXml, normalizedDefinition.sortState ?? null);
  return nextSheetXml;
}

export function setWorksheetAutoFilterRangeInSheetXml(sheetXml: string, range: string): string {
  const normalizedRange = normalizeRangeRef(range);
  const autoFilterTag = findFirstXmlTag(sheetXml, "autoFilter");
  const externalSortStateTag = findFirstXmlTag(sheetXml, "sortState");
  const currentState = parseAutoFilterStateFromTags(autoFilterTag, externalSortStateTag);

  if (!currentState) {
    return upsertWorksheetAutoFilterXml(sheetXml, buildSelfClosingXmlElement("autoFilter", [["ref", normalizedRange]]));
  }

  const nextRootItems = currentState.rootItems
    .map((item) => rewriteRootItemForRange(item, normalizedRange))
    .filter((item): item is AutoFilterRootItem => item !== null);
  const nextAutoFilterXml = buildAutoFilterXml(normalizedRange, nextRootItems, currentState.autoFilterAttributes);
  let nextSheetXml = upsertWorksheetAutoFilterXml(sheetXml, nextAutoFilterXml);

  if (currentState.sortStateLocation === "external") {
    const currentExternalSortStateTag = findFirstXmlTag(nextSheetXml, "sortState");
    const nextSortStateXml = currentExternalSortStateTag
      ? rewriteSortStateTagForRange(currentExternalSortStateTag, normalizedRange)
      : null;
    if (currentExternalSortStateTag) {
      nextSheetXml = replaceXmlTagSource(nextSheetXml, currentExternalSortStateTag, nextSortStateXml ?? "");
    }
  }

  return nextSheetXml;
}

export function setWorksheetSortStateInSheetXml(sheetXml: string, sortState: SortStateDefinition | null): string {
  return upsertWorksheetSortStateXml(sheetXml, sortState);
}

export function setWorksheetAutoFilterColumnInSheetXml(
  sheetXml: string,
  column: AutoFilterColumn,
): string {
  const currentState = parseAutoFilterStateFromTags(
    findFirstXmlTag(sheetXml, "autoFilter"),
    findFirstXmlTag(sheetXml, "sortState"),
  );
  if (!currentState) {
    throw new XlsxError("Worksheet autoFilter range is not set");
  }

  assertColumnInRange(column.columnNumber, currentState.range);
  const nextRootItems = buildRootItemsForColumnUpdate(currentState, column);
  const nextAutoFilterXml = buildAutoFilterXml(
    currentState.range,
    nextRootItems,
    currentState.autoFilterAttributes,
  );
  return upsertWorksheetAutoFilterXml(sheetXml, nextAutoFilterXml);
}

export function clearWorksheetAutoFilterColumnsInSheetXml(
  sheetXml: string,
  columnNumbers?: number[],
): string {
  const currentState = parseAutoFilterStateFromTags(
    findFirstXmlTag(sheetXml, "autoFilter"),
    findFirstXmlTag(sheetXml, "sortState"),
  );
  if (!currentState) {
    return sheetXml;
  }

  const nextRootItems = buildRootItemsForClear(currentState, columnNumbers);
  const nextAutoFilterXml = buildAutoFilterXml(
    currentState.range,
    nextRootItems,
    currentState.autoFilterAttributes,
  );
  return upsertWorksheetAutoFilterXml(sheetXml, nextAutoFilterXml);
}

export function setTableAutoFilterDefinitionInTableXml(
  tableXml: string,
  definition: AutoFilterDefinition,
): string {
  const normalizedDefinition = normalizeAutoFilterDefinition(definition);
  const currentState = parseAutoFilterStateFromTags(findFirstXmlTag(tableXml, "autoFilter"), null);
  const nextRootItems = buildRootItemsForFullDefinition(currentState, normalizedDefinition);
  if (normalizedDefinition.sortState && currentState?.sortStateLocation !== "autoFilter") {
    insertSortStateRootItem(nextRootItems, {
      kind: "sortState",
      source: buildSortStateXml(normalizedDefinition.sortState),
    });
  }
  const nextAutoFilterXml = buildAutoFilterXml(
    normalizedDefinition.range,
    nextRootItems,
    currentState?.autoFilterAttributes ?? [],
  );
  return upsertTableAutoFilterXml(tableXml, nextAutoFilterXml);
}

export function setTableAutoFilterColumnInTableXml(
  tableXml: string,
  column: AutoFilterColumn,
): string {
  const currentState = parseAutoFilterStateFromTags(findFirstXmlTag(tableXml, "autoFilter"), null);
  const tableRange = parseTableRange(tableXml);
  const currentRange = currentState?.range ?? tableRange;

  if (!currentRange) {
    throw new XlsxError("Table range is missing");
  }

  assertColumnInRange(column.columnNumber, currentRange);
  const nextRootItems = buildRootItemsForColumnUpdate(currentState ?? createEmptyState(currentRange), column);
  const nextAutoFilterXml = buildAutoFilterXml(
    currentRange,
    nextRootItems,
    currentState?.autoFilterAttributes ?? [],
  );
  return upsertTableAutoFilterXml(tableXml, nextAutoFilterXml);
}

export function clearTableAutoFilterColumnsInTableXml(
  tableXml: string,
  columnNumbers?: number[],
): string {
  const currentState = parseAutoFilterStateFromTags(findFirstXmlTag(tableXml, "autoFilter"), null);
  if (!currentState) {
    return tableXml;
  }

  const nextRootItems = buildRootItemsForClear(currentState, columnNumbers);
  const nextAutoFilterXml = buildAutoFilterXml(
    currentState.range,
    nextRootItems,
    currentState.autoFilterAttributes,
  );
  return upsertTableAutoFilterXml(tableXml, nextAutoFilterXml);
}

export function rewriteAutoFilterTagWithTransformedRefs(
  autoFilterTag: XmlTag,
  transformRange: (range: string) => string | null,
  targetColumnNumber: number,
  columnCount: number,
  mode: "shift" | "delete",
): string | null {
  const ref = getTagAttr(autoFilterTag, "ref");
  if (!ref) {
    return autoFilterTag.source;
  }

  const currentRange = normalizeRangeRef(ref);
  const nextRange = transformRange(currentRange);
  if (nextRange === null) {
    return null;
  }

  if (autoFilterTag.innerXml === null) {
    return rebuildTagWithRef(autoFilterTag, "autoFilter", nextRange, null);
  }

  const { startColumn: currentStartColumn } = parseRangeRef(currentRange);
  const { startColumn: nextStartColumn } = parseRangeRef(nextRange);
  const nextChildXml = findTopLevelXmlTags(autoFilterTag.innerXml)
    .map((childTag) => {
      if (childTag.tagName !== "filterColumn") {
        return childTag.source;
      }

      return rewriteFilterColumnTagWithTransformedColId(
        childTag,
        currentStartColumn,
        nextStartColumn,
        targetColumnNumber,
        columnCount,
        mode,
      );
    })
    .filter((childXml): childXml is string => childXml !== null)
    .join("");

  return rebuildTagWithRef(autoFilterTag, "autoFilter", nextRange, nextChildXml);
}

export function rewriteSortStateTagWithTransformedRefs(
  sortStateTag: XmlTag,
  transformRange: (range: string) => string | null,
): string | null {
  const ref = getTagAttr(sortStateTag, "ref");
  if (!ref) {
    return sortStateTag.source;
  }

  const nextRange = transformRange(normalizeRangeRef(ref));
  if (nextRange === null) {
    return null;
  }

  const childTags = findTopLevelXmlTags(sortStateTag.innerXml ?? "");
  let validSortConditionCount = 0;
  const nextChildXml = childTags
    .map((childTag) => {
      if (childTag.tagName !== "sortCondition") {
        return childTag.source;
      }

      const nextSortConditionXml = rewriteSortConditionTagWithTransformedRef(childTag, transformRange);
      if (nextSortConditionXml !== null) {
        validSortConditionCount += 1;
      }
      return nextSortConditionXml;
    })
    .filter((childXml): childXml is string => childXml !== null)
    .join("");

  if (validSortConditionCount === 0) {
    return null;
  }

  return rebuildTagWithRef(sortStateTag, "sortState", nextRange, nextChildXml);
}

function parseAutoFilterDefinitionFromTags(
  autoFilterTag: XmlTag | null,
  externalSortStateTag: XmlTag | null,
): AutoFilterDefinition | null {
  const state = parseAutoFilterStateFromTags(autoFilterTag, externalSortStateTag);
  if (!state) {
    return null;
  }

  return {
    range: state.range,
    columns: state.rootItems
      .filter((item): item is Extract<AutoFilterRootItem, { kind: "column" }> => item.kind === "column")
      .map((item) => item.definition)
      .filter((column): column is AutoFilterColumn => column !== null),
    sortState: state.sortState,
  };
}

function parseAutoFilterStateFromTags(
  autoFilterTag: XmlTag | null,
  externalSortStateTag: XmlTag | null,
): ParsedAutoFilterState | null {
  if (!autoFilterTag) {
    return null;
  }

  const ref = getTagAttr(autoFilterTag, "ref");
  if (!ref) {
    return null;
  }

  const range = normalizeRangeRef(ref);
  const rangeBounds = parseRangeRef(range);
  const rootItems: AutoFilterRootItem[] = [];
  let sortState = externalSortStateTag ? parseSortStateTag(externalSortStateTag) : null;
  let sortStateLocation: "autoFilter" | "external" | null = externalSortStateTag ? "external" : null;

  for (const childTag of findTopLevelXmlTags(autoFilterTag.innerXml ?? "")) {
    if (childTag.tagName === "filterColumn") {
      const parsedColumn = parseFilterColumnTag(childTag, rangeBounds.startColumn);
      rootItems.push({
        kind: "column",
        columnNumber: parsedColumn.columnNumber,
        definition: parsedColumn.definition,
        source: childTag.source,
      });
      continue;
    }

    if (childTag.tagName === "sortState" && sortStateLocation === null) {
      sortState = parseSortStateTag(childTag);
      sortStateLocation = "autoFilter";
      rootItems.push({ kind: "sortState", source: childTag.source });
      continue;
    }

    rootItems.push({ kind: "raw", source: childTag.source });
  }

  return {
    range,
    autoFilterAttributes: parseAttributes(autoFilterTag.attributesSource),
    rootItems,
    sortState,
    sortStateLocation,
  };
}

function parseFilterColumnTag(
  filterColumnTag: XmlTag,
  startColumnNumber: number,
): {
  columnNumber: number;
  definition: AutoFilterColumn | null;
} {
  const colIdText = getTagAttr(filterColumnTag, "colId");
  const colId = colIdText ? Number(colIdText) : 0;
  const columnNumber = Number.isInteger(colId) && colId >= 0 ? startColumnNumber + colId : startColumnNumber;
  const childTags = findTopLevelXmlTags(filterColumnTag.innerXml ?? "");
  const filtersTag = childTags.find((tag) => tag.tagName === "filters") ?? null;
  const customFiltersTag = childTags.find((tag) => tag.tagName === "customFilters") ?? null;
  const colorFilterTag = childTags.find((tag) => tag.tagName === "colorFilter") ?? null;
  const dynamicFilterTag = childTags.find((tag) => tag.tagName === "dynamicFilter") ?? null;
  const top10Tag = childTags.find((tag) => tag.tagName === "top10") ?? null;
  const iconFilterTag = childTags.find((tag) => tag.tagName === "iconFilter") ?? null;
  const supportedChildCount = [
    filtersTag,
    customFiltersTag,
    colorFilterTag,
    dynamicFilterTag,
    top10Tag,
    iconFilterTag,
  ].filter((tag) => tag !== null).length;

  if (supportedChildCount === 1 && filtersTag) {
    return {
      columnNumber,
      definition: parseFiltersColumnTag(filtersTag, columnNumber),
    };
  }

  if (supportedChildCount === 1 && customFiltersTag) {
    return {
      columnNumber,
      definition: parseCustomFilterColumnTag(customFiltersTag, columnNumber),
    };
  }

  if (supportedChildCount === 1 && colorFilterTag) {
    return {
      columnNumber,
      definition: parseColorFilterColumnTag(colorFilterTag, columnNumber),
    };
  }

  if (supportedChildCount === 1 && dynamicFilterTag) {
    return {
      columnNumber,
      definition: parseDynamicFilterColumnTag(dynamicFilterTag, columnNumber),
    };
  }

  if (supportedChildCount === 1 && top10Tag) {
    return {
      columnNumber,
      definition: parseTop10FilterColumnTag(top10Tag, columnNumber),
    };
  }

  if (supportedChildCount === 1 && iconFilterTag) {
    return {
      columnNumber,
      definition: parseIconFilterColumnTag(iconFilterTag, columnNumber),
    };
  }

  return {
    columnNumber,
    definition: null,
  };
}

function parseFiltersColumnTag(filtersTag: XmlTag, columnNumber: number): AutoFilterColumn | null {
  const childTags = findTopLevelXmlTags(filtersTag.innerXml ?? "");
  const filterTags = childTags.filter((tag) => tag.tagName === "filter");
  const dateGroupTags = childTags.filter((tag) => tag.tagName === "dateGroupItem");
  const unknownTags = childTags.filter((tag) => tag.tagName !== "filter" && tag.tagName !== "dateGroupItem");
  const blank = parseOptionalXmlBoolean(getTagAttr(filtersTag, "blank"));

  if (unknownTags.length > 0) {
    return null;
  }

  if (dateGroupTags.length > 0 && filterTags.length === 0) {
    const items = dateGroupTags
      .map((tag) => parseDateGroupItemTag(tag))
      .filter((item): item is DateGroupItem => item !== null);
    if (items.length === 0) {
      return null;
    }

    return {
      columnNumber,
      kind: "dateGroup",
      items,
    };
  }

  if (filterTags.length === 0) {
    if (blank === null || dateGroupTags.length > 0) {
      return null;
    }

    return {
      columnNumber,
      kind: "blank",
      mode: blank ? "blank" : "nonBlank",
    };
  }

  const values = filterTags
    .map((tag) => getTagAttr(tag, "val"))
    .filter((value): value is string => value !== undefined);

  return {
    columnNumber,
    kind: "values",
    values,
    ...(blank === true ? { includeBlank: true } : {}),
  };
}

function parseCustomFilterColumnTag(customFiltersTag: XmlTag, columnNumber: number): AutoFilterColumn | null {
  const childTags = findTopLevelXmlTags(customFiltersTag.innerXml ?? "");
  const customFilterTags = childTags.filter((tag) => tag.tagName === "customFilter");
  const unknownTags = childTags.filter((tag) => tag.tagName !== "customFilter");

  if (customFilterTags.length === 0 || unknownTags.length > 0) {
    return null;
  }

  const conditions = customFilterTags
    .map((tag) => parseCustomFilterConditionTag(tag))
    .filter((condition): condition is AutoFilterCondition => condition !== null);
  if (conditions.length !== customFilterTags.length) {
    return null;
  }

  return {
    columnNumber,
    kind: "custom",
    join: parseOptionalXmlBoolean(getTagAttr(customFiltersTag, "and")) ? "and" : "or",
    conditions,
  };
}

function parseColorFilterColumnTag(colorFilterTag: XmlTag, columnNumber: number): AutoFilterColumn | null {
  const dxfId = parseIntegerAttr(colorFilterTag, "dxfId");
  if (dxfId === null || dxfId < 0) {
    return null;
  }

  return {
    columnNumber,
    kind: "color",
    dxfId,
    cellColor: parseOptionalXmlBoolean(getTagAttr(colorFilterTag, "cellColor")) ?? false,
  };
}

function parseDynamicFilterColumnTag(dynamicFilterTag: XmlTag, columnNumber: number): AutoFilterColumn | null {
  const type = getTagAttr(dynamicFilterTag, "type");
  if (!type) {
    return null;
  }

  const val = parseNumberAttr(dynamicFilterTag, "val");
  const maxVal = parseNumberAttr(dynamicFilterTag, "maxVal");
  const valIso = getTagAttr(dynamicFilterTag, "valIso");
  const maxValIso = getTagAttr(dynamicFilterTag, "maxValIso");

  return {
    columnNumber,
    kind: "dynamic",
    type,
    ...(val !== null ? { val } : {}),
    ...(maxVal !== null ? { maxVal } : {}),
    ...(valIso !== undefined ? { valIso } : {}),
    ...(maxValIso !== undefined ? { maxValIso } : {}),
  };
}

function parseTop10FilterColumnTag(top10Tag: XmlTag, columnNumber: number): AutoFilterColumn | null {
  const value = parseNumberAttr(top10Tag, "val");
  if (value === null) {
    return null;
  }

  const filterValue = parseNumberAttr(top10Tag, "filterVal");
  return {
    columnNumber,
    kind: "top10",
    top: parseOptionalXmlBoolean(getTagAttr(top10Tag, "top")) ?? true,
    percent: parseOptionalXmlBoolean(getTagAttr(top10Tag, "percent")) ?? false,
    value,
    ...(filterValue !== null ? { filterValue } : {}),
  };
}

function parseIconFilterColumnTag(iconFilterTag: XmlTag, columnNumber: number): AutoFilterColumn | null {
  const iconSet = getTagAttr(iconFilterTag, "iconSet");
  const iconId = parseIntegerAttr(iconFilterTag, "iconId");
  if (iconSet === undefined && iconId === null) {
    return null;
  }

  return {
    columnNumber,
    kind: "icon",
    ...(iconSet !== undefined ? { iconSet } : {}),
    ...(iconId !== null ? { iconId } : {}),
  };
}

function parseDateGroupItemTag(tag: XmlTag): DateGroupItem | null {
  const year = parseIntegerAttr(tag, "year");
  const dateTimeGrouping = getTagAttr(tag, "dateTimeGrouping");
  if (year === null || !isDateTimeGrouping(dateTimeGrouping)) {
    return null;
  }

  const month = parseIntegerAttr(tag, "month");
  const day = parseIntegerAttr(tag, "day");
  const hour = parseIntegerAttr(tag, "hour");
  const minute = parseIntegerAttr(tag, "minute");
  const second = parseIntegerAttr(tag, "second");

  return {
    year,
    ...(month !== null ? { month } : {}),
    ...(day !== null ? { day } : {}),
    ...(hour !== null ? { hour } : {}),
    ...(minute !== null ? { minute } : {}),
    ...(second !== null ? { second } : {}),
    dateTimeGrouping,
  };
}

function parseCustomFilterConditionTag(tag: XmlTag): AutoFilterCondition | null {
  const rawOperator = getTagAttr(tag, "operator") ?? "equal";
  const value = getTagAttr(tag, "val");
  if (value === undefined) {
    return null;
  }

  const textCondition = parseWildcardTextCondition(rawOperator, value);
  if (textCondition) {
    return textCondition;
  }

  switch (rawOperator) {
    case "equal":
      return { operator: "equals", value };
    case "notEqual":
      return { operator: "notEquals", value };
    case "greaterThan":
      return { operator: "greaterThan", value };
    case "greaterThanOrEqual":
      return { operator: "greaterThanOrEqual", value };
    case "lessThan":
      return { operator: "lessThan", value };
    case "lessThanOrEqual":
      return { operator: "lessThanOrEqual", value };
    default:
      return null;
  }
}

function parseWildcardTextCondition(operator: string, value: string): AutoFilterCondition | null {
  if (operator === "equal" && value.startsWith("*") && value.endsWith("*") && value.length >= 2) {
    return { operator: "contains", value: unescapeAutoFilterPattern(value.slice(1, -1)) };
  }

  if (operator === "notEqual" && value.startsWith("*") && value.endsWith("*") && value.length >= 2) {
    return { operator: "notContains", value: unescapeAutoFilterPattern(value.slice(1, -1)) };
  }

  if (operator === "equal" && value.endsWith("*") && !value.startsWith("*")) {
    return { operator: "beginsWith", value: unescapeAutoFilterPattern(value.slice(0, -1)) };
  }

  if (operator === "equal" && value.startsWith("*") && !value.endsWith("*")) {
    return { operator: "endsWith", value: unescapeAutoFilterPattern(value.slice(1)) };
  }

  return null;
}

function parseSortStateTag(tag: XmlTag): SortStateDefinition | null {
  const ref = getTagAttr(tag, "ref");
  if (!ref) {
    return null;
  }

  const range = normalizeRangeRef(ref);
  const conditions = findTopLevelXmlTags(tag.innerXml ?? "")
    .filter((childTag) => childTag.tagName === "sortCondition")
    .map((sortConditionTag) => {
      const conditionRef = getTagAttr(sortConditionTag, "ref");
      if (!conditionRef) {
        return null;
      }

      const normalizedConditionRef = normalizeRangeRef(conditionRef);
      const { startColumn } = parseRangeRef(normalizedConditionRef);
      const descending = parseOptionalXmlBoolean(getTagAttr(sortConditionTag, "descending"));
      return {
        columnNumber: startColumn,
        ...(descending ? { descending: true } : {}),
      };
    })
    .filter((condition): condition is { columnNumber: number; descending?: boolean } => condition !== null);

  return {
    range,
    conditions,
  };
}

function buildRootItemsForFullDefinition(
  currentState: ParsedAutoFilterState | null,
  definition: AutoFilterDefinition,
): AutoFilterRootItem[] {
  const nextSupportedColumns = new Map(
    definition.columns.map((column) => [column.columnNumber, buildFilterColumnRootItem(column, definition.range)]),
  );
  const nextItems: AutoFilterRootItem[] = [];
  let hasSortStateItem = false;
  const nextSortStateItem =
    definition.sortState && currentState?.sortStateLocation === "autoFilter"
      ? ({
          kind: "sortState",
          source: buildSortStateXml(definition.sortState),
        } satisfies AutoFilterRootItem)
      : null;

  for (const item of currentState?.rootItems ?? []) {
    if (item.kind === "column") {
      if (!isColumnInRange(item.columnNumber, definition.range)) {
        continue;
      }

      const replacement = nextSupportedColumns.get(item.columnNumber);
      if (replacement) {
        nextItems.push(replacement);
        nextSupportedColumns.delete(item.columnNumber);
        continue;
      }

      if (item.definition === null) {
        nextItems.push(item);
      }
      continue;
    }

    if (item.kind === "sortState") {
      if (nextSortStateItem) {
        nextItems.push(nextSortStateItem);
        hasSortStateItem = true;
      }
      continue;
    }

    nextItems.push(item);
  }

  for (const column of [...nextSupportedColumns.values()].sort((left, right) => left.columnNumber - right.columnNumber)) {
    insertColumnRootItem(nextItems, column);
  }

  if (nextSortStateItem && !hasSortStateItem) {
    insertSortStateRootItem(nextItems, nextSortStateItem);
  }

  return nextItems;
}

function buildRootItemsForColumnUpdate(
  currentState: ParsedAutoFilterState,
  column: AutoFilterColumn,
): AutoFilterRootItem[] {
  const replacement = buildFilterColumnRootItem(column, currentState.range);
  const nextItems: AutoFilterRootItem[] = [];
  let replaced = false;

  for (const item of currentState.rootItems) {
    if (item.kind === "column" && item.columnNumber === column.columnNumber) {
      nextItems.push(replacement);
      replaced = true;
      continue;
    }

    nextItems.push(item);
  }

  if (!replaced) {
    insertColumnRootItem(nextItems, replacement);
  }

  return nextItems;
}

function buildRootItemsForClear(
  currentState: ParsedAutoFilterState,
  columnNumbers?: number[],
): AutoFilterRootItem[] {
  if (columnNumbers === undefined) {
    return currentState.rootItems.filter((item) => item.kind !== "column");
  }

  const targetColumns = new Set(columnNumbers);
  return currentState.rootItems.filter(
    (item) => item.kind !== "column" || !targetColumns.has(item.columnNumber),
  );
}

function rewriteRootItemForRange(item: AutoFilterRootItem, range: string): AutoFilterRootItem | null {
  if (item.kind === "column") {
    if (!isColumnInRange(item.columnNumber, range)) {
      return null;
    }

    return {
      ...item,
      source: rewriteFilterColumnSourceForRange(item.source, item.columnNumber, range),
    };
  }

  if (item.kind === "sortState") {
    const sortStateTag = findFirstXmlTag(item.source, "sortState");
    if (!sortStateTag) {
      return item;
    }

    const nextSortStateXml = rewriteSortStateTagForRange(sortStateTag, range);
    return nextSortStateXml ? { kind: "sortState", source: nextSortStateXml } : null;
  }

  return item;
}

function buildFilterColumnRootItem(column: AutoFilterColumn, range: string): Extract<AutoFilterRootItem, { kind: "column" }> {
  return {
    kind: "column",
    columnNumber: column.columnNumber,
    definition: column,
    source: buildFilterColumnXml(column, range),
  };
}

function buildFilterColumnXml(column: AutoFilterColumn, range: string): string {
  assertColumnInRange(column.columnNumber, range);
  const { startColumn } = parseRangeRef(range);
  const attributes: Array<[string, string]> = [["colId", String(column.columnNumber - startColumn)]];

  switch (column.kind) {
    case "values":
      return buildXmlElement("filterColumn", attributes, buildValuesFilterXml(column));
    case "blank":
      return buildXmlElement("filterColumn", attributes, buildBlankFilterXml(column));
    case "custom":
      return buildXmlElement("filterColumn", attributes, buildCustomFilterXml(column));
    case "dateGroup":
      return buildXmlElement("filterColumn", attributes, buildDateGroupFilterXml(column));
    case "color":
      return buildXmlElement("filterColumn", attributes, buildColorFilterXml(column));
    case "dynamic":
      return buildXmlElement("filterColumn", attributes, buildDynamicFilterXml(column));
    case "top10":
      return buildXmlElement("filterColumn", attributes, buildTop10FilterXml(column));
    case "icon":
      return buildXmlElement("filterColumn", attributes, buildIconFilterXml(column));
    default:
      return buildSelfClosingXmlElement("filterColumn", attributes);
  }
}

function buildValuesFilterXml(column: Extract<AutoFilterColumn, { kind: "values" }>): string {
  const attributes: Array<[string, string]> = [];
  if (column.includeBlank !== undefined) {
    attributes.push(["blank", column.includeBlank ? "1" : "0"]);
  }

  const childXml = column.values.map((value) => buildSelfClosingXmlElement("filter", [["val", value]]));
  return childXml.length === 0
    ? buildSelfClosingXmlElement("filters", attributes)
    : buildXmlElement("filters", attributes, childXml.join(""));
}

function buildBlankFilterXml(column: Extract<AutoFilterColumn, { kind: "blank" }>): string {
  return buildSelfClosingXmlElement("filters", [["blank", column.mode === "blank" ? "1" : "0"]]);
}

function buildCustomFilterXml(column: Extract<AutoFilterColumn, { kind: "custom" }>): string {
  const attributes: Array<[string, string]> = [];
  if (column.join === "and") {
    attributes.push(["and", "1"]);
  }

  const childXml = column.conditions.map((condition) =>
    buildSelfClosingXmlElement("customFilter", buildCustomFilterConditionAttributes(condition)),
  );
  return buildXmlElement("customFilters", attributes, childXml.join(""));
}

function buildDateGroupFilterXml(column: Extract<AutoFilterColumn, { kind: "dateGroup" }>): string {
  const childXml = column.items.map((item) => {
    const attributes: Array<[string, string]> = [
      ["year", String(item.year)],
      ["dateTimeGrouping", item.dateTimeGrouping],
    ];
    appendOptionalNumberAttribute(attributes, "month", item.month);
    appendOptionalNumberAttribute(attributes, "day", item.day);
    appendOptionalNumberAttribute(attributes, "hour", item.hour);
    appendOptionalNumberAttribute(attributes, "minute", item.minute);
    appendOptionalNumberAttribute(attributes, "second", item.second);
    return buildSelfClosingXmlElement("dateGroupItem", attributes);
  });
  return buildXmlElement("filters", [], childXml.join(""));
}

function buildColorFilterXml(column: Extract<AutoFilterColumn, { kind: "color" }>): string {
  return buildSelfClosingXmlElement("colorFilter", [
    ["dxfId", String(column.dxfId)],
    ["cellColor", column.cellColor ? "1" : "0"],
  ]);
}

function buildDynamicFilterXml(column: Extract<AutoFilterColumn, { kind: "dynamic" }>): string {
  const attributes: Array<[string, string]> = [["type", column.type]];
  appendOptionalNumberAttribute(attributes, "val", column.val);
  appendOptionalNumberAttribute(attributes, "maxVal", column.maxVal);
  appendOptionalStringAttribute(attributes, "valIso", column.valIso);
  appendOptionalStringAttribute(attributes, "maxValIso", column.maxValIso);
  return buildSelfClosingXmlElement("dynamicFilter", attributes);
}

function buildTop10FilterXml(column: Extract<AutoFilterColumn, { kind: "top10" }>): string {
  const attributes: Array<[string, string]> = [
    ["top", column.top ? "1" : "0"],
    ["percent", column.percent ? "1" : "0"],
    ["val", String(column.value)],
  ];
  appendOptionalNumberAttribute(attributes, "filterVal", column.filterValue);
  return buildSelfClosingXmlElement("top10", attributes);
}

function buildIconFilterXml(column: Extract<AutoFilterColumn, { kind: "icon" }>): string {
  const attributes: Array<[string, string]> = [];
  appendOptionalStringAttribute(attributes, "iconSet", column.iconSet);
  if (column.iconId !== undefined) {
    attributes.push(["iconId", String(column.iconId)]);
  }
  return buildSelfClosingXmlElement("iconFilter", attributes);
}

function buildCustomFilterConditionAttributes(condition: AutoFilterCondition): Array<[string, string]> {
  switch (condition.operator) {
    case "equals":
      return [["operator", "equal"], ["val", String(condition.value)]];
    case "notEquals":
      return [["operator", "notEqual"], ["val", String(condition.value)]];
    case "greaterThan":
      return [["operator", "greaterThan"], ["val", String(condition.value)]];
    case "greaterThanOrEqual":
      return [["operator", "greaterThanOrEqual"], ["val", String(condition.value)]];
    case "lessThan":
      return [["operator", "lessThan"], ["val", String(condition.value)]];
    case "lessThanOrEqual":
      return [["operator", "lessThanOrEqual"], ["val", String(condition.value)]];
    case "contains":
      return [["operator", "equal"], ["val", `*${escapeAutoFilterPattern(condition.value)}*`]];
    case "notContains":
      return [["operator", "notEqual"], ["val", `*${escapeAutoFilterPattern(condition.value)}*`]];
    case "beginsWith":
      return [["operator", "equal"], ["val", `${escapeAutoFilterPattern(condition.value)}*`]];
    case "endsWith":
      return [["operator", "equal"], ["val", `*${escapeAutoFilterPattern(condition.value)}`]];
  }
}

function buildSortStateXml(sortState: SortStateDefinition): string {
  const normalizedRange = normalizeRangeRef(sortState.range);
  const { startRow, endRow } = parseRangeRef(normalizedRange);
  const childXml = sortState.conditions.map((condition) =>
    buildSelfClosingXmlElement("sortCondition", [
      ["ref", formatRangeRef(startRow, condition.columnNumber, endRow, condition.columnNumber)],
      ...(condition.descending ? ([["descending", "1"]] as Array<[string, string]>) : []),
    ]),
  );
  return buildXmlElement("sortState", [["ref", normalizedRange]], childXml.join(""));
}

function buildAutoFilterXml(
  range: string,
  rootItems: AutoFilterRootItem[],
  existingAttributes: Array<[string, string]>,
): string {
  const normalizedRange = normalizeRangeRef(range);
  const attributes = [...existingAttributes];
  const refIndex = attributes.findIndex(([name]) => name === "ref");
  if (refIndex === -1) {
    attributes.push(["ref", normalizedRange]);
  } else {
    attributes[refIndex] = ["ref", normalizedRange];
  }

  const innerXml = rootItems.map((item) => item.source).join("");
  return innerXml.length === 0
    ? buildSelfClosingXmlElement("autoFilter", attributes)
    : buildXmlElement("autoFilter", attributes, innerXml);
}

function upsertWorksheetAutoFilterXml(sheetXml: string, autoFilterXml: string): string {
  const autoFilterTag = findFirstXmlTag(sheetXml, "autoFilter");
  if (autoFilterTag) {
    return replaceXmlTagSource(sheetXml, autoFilterTag, autoFilterXml);
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, [...WORKSHEET_AUTO_FILTER_FOLLOWING_TAGS]);
  return sheetXml.slice(0, insertionIndex) + autoFilterXml + sheetXml.slice(insertionIndex);
}

function upsertWorksheetSortStateXml(sheetXml: string, sortState: SortStateDefinition | null): string {
  const sortStateTag = findFirstXmlTag(sheetXml, "sortState");
  if (!sortState) {
    return sortStateTag ? replaceXmlTagSource(sheetXml, sortStateTag, "") : sheetXml;
  }

  const sortStateXml = buildSortStateXml(sortState);
  if (sortStateTag) {
    return replaceXmlTagSource(sheetXml, sortStateTag, sortStateXml);
  }

  const autoFilterTag = findFirstXmlTag(sheetXml, "autoFilter");
  if (!autoFilterTag) {
    const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, [...WORKSHEET_AUTO_FILTER_FOLLOWING_TAGS]);
    return sheetXml.slice(0, insertionIndex) + sortStateXml + sheetXml.slice(insertionIndex);
  }

  return sheetXml.slice(0, autoFilterTag.end) + sortStateXml + sheetXml.slice(autoFilterTag.end);
}

function upsertTableAutoFilterXml(tableXml: string, autoFilterXml: string): string {
  const autoFilterTag = findFirstXmlTag(tableXml, "autoFilter");
  if (autoFilterTag) {
    return replaceXmlTagSource(tableXml, autoFilterTag, autoFilterXml);
  }

  const tableTag = findFirstXmlTag(tableXml, "table");
  if (!tableTag) {
    throw new XlsxError("Table XML is missing <table>");
  }

  const childTags = findTopLevelXmlTags(tableTag.innerXml ?? "");
  const insertionIndex =
    childTags.length > 0 ? getXmlTagInnerStart(tableTag) + childTags[0].start : getXmlTagInnerStart(tableTag);
  return tableXml.slice(0, insertionIndex) + autoFilterXml + tableXml.slice(insertionIndex);
}

function parseTableRange(tableXml: string): string | null {
  const tableTag = findFirstXmlTag(tableXml, "table");
  if (!tableTag) {
    return null;
  }

  const ref = getTagAttr(tableTag, "ref");
  return ref ? normalizeRangeRef(ref) : null;
}

function createEmptyState(range: string): ParsedAutoFilterState {
  return {
    range: normalizeRangeRef(range),
    autoFilterAttributes: [],
    rootItems: [],
    sortState: null,
    sortStateLocation: null,
  };
}

function normalizeAutoFilterDefinition(definition: AutoFilterDefinition): AutoFilterDefinition {
  const range = normalizeRangeRef(definition.range);
  const seenColumns = new Set<number>();
  const columns = [...definition.columns]
    .map((column) => normalizeAutoFilterColumn(column, range))
    .sort((left, right) => left.columnNumber - right.columnNumber);

  for (const column of columns) {
    if (seenColumns.has(column.columnNumber)) {
      throw new XlsxError(`Duplicate autoFilter column: ${column.columnNumber}`);
    }

    seenColumns.add(column.columnNumber);
  }

  const sortState = definition.sortState ? normalizeSortStateDefinition(definition.sortState, range) : null;
  return {
    range,
    columns,
    sortState,
  };
}

function normalizeAutoFilterColumn(column: AutoFilterColumn, range: string): AutoFilterColumn {
  assertColumnInRange(column.columnNumber, range);

  switch (column.kind) {
    case "values":
      return {
        columnNumber: column.columnNumber,
        kind: "values",
        values: [...column.values],
        ...(column.includeBlank !== undefined ? { includeBlank: column.includeBlank } : {}),
      };
    case "blank":
      return {
        columnNumber: column.columnNumber,
        kind: "blank",
        mode: column.mode,
      };
    case "custom":
      if (column.conditions.length === 0 || column.conditions.length > 2) {
        throw new XlsxError(`Invalid custom autoFilter condition count for column ${column.columnNumber}`);
      }
      return {
        columnNumber: column.columnNumber,
        kind: "custom",
        join: column.join,
        conditions: column.conditions.map((condition) => ({ ...condition })),
      };
    case "dateGroup":
      if (column.items.length === 0) {
        throw new XlsxError(`Date-group autoFilter column must include at least one item: ${column.columnNumber}`);
      }
      return {
        columnNumber: column.columnNumber,
        kind: "dateGroup",
        items: column.items.map((item) => ({ ...item })),
      };
    case "color":
      if (!Number.isInteger(column.dxfId) || column.dxfId < 0) {
        throw new XlsxError(`Invalid color autoFilter dxfId for column ${column.columnNumber}`);
      }
      return {
        columnNumber: column.columnNumber,
        kind: "color",
        dxfId: column.dxfId,
        cellColor: column.cellColor,
      };
    case "dynamic":
      return {
        columnNumber: column.columnNumber,
        kind: "dynamic",
        type: column.type,
        ...(column.val !== undefined ? { val: column.val } : {}),
        ...(column.maxVal !== undefined ? { maxVal: column.maxVal } : {}),
        ...(column.valIso !== undefined ? { valIso: column.valIso } : {}),
        ...(column.maxValIso !== undefined ? { maxValIso: column.maxValIso } : {}),
      };
    case "top10":
      return {
        columnNumber: column.columnNumber,
        kind: "top10",
        top: column.top,
        percent: column.percent,
        value: column.value,
        ...(column.filterValue !== undefined ? { filterValue: column.filterValue } : {}),
      };
    case "icon":
      if (column.iconId !== undefined && (!Number.isInteger(column.iconId) || column.iconId < 0)) {
        throw new XlsxError(`Invalid icon autoFilter iconId for column ${column.columnNumber}`);
      }
      return {
        columnNumber: column.columnNumber,
        kind: "icon",
        ...(column.iconSet !== undefined ? { iconSet: column.iconSet } : {}),
        ...(column.iconId !== undefined ? { iconId: column.iconId } : {}),
      };
    default:
      return column;
  }
}

function normalizeSortStateDefinition(sortState: SortStateDefinition, filterRange: string): SortStateDefinition {
  const range = normalizeRangeRef(sortState.range);
  for (const condition of sortState.conditions) {
    assertColumnInRange(condition.columnNumber, filterRange);
  }

  return {
    range,
    conditions: sortState.conditions.map((condition) => ({ ...condition })),
  };
}

function assertColumnInRange(columnNumber: number, range: string): void {
  if (!isColumnInRange(columnNumber, range)) {
    throw new XlsxError(`AutoFilter column ${columnNumber} is outside range ${normalizeRangeRef(range)}`);
  }
}

function isColumnInRange(columnNumber: number, range: string): boolean {
  const { startColumn, endColumn } = parseRangeRef(range);
  return columnNumber >= startColumn && columnNumber <= endColumn;
}

function insertColumnRootItem(
  rootItems: AutoFilterRootItem[],
  item: Extract<AutoFilterRootItem, { kind: "column" }>,
): void {
  let firstNonColumnIndex: number | null = null;

  for (let index = 0; index < rootItems.length; index += 1) {
    const currentItem = rootItems[index];
    if (currentItem.kind === "column") {
      if (currentItem.columnNumber > item.columnNumber) {
        rootItems.splice(index, 0, item);
        return;
      }
      continue;
    }

    if (firstNonColumnIndex === null) {
      firstNonColumnIndex = index;
    }
  }

  rootItems.splice(firstNonColumnIndex ?? rootItems.length, 0, item);
}

function insertSortStateRootItem(
  rootItems: AutoFilterRootItem[],
  item: Extract<AutoFilterRootItem, { kind: "sortState" }>,
): void {
  const firstRawIndex = rootItems.findIndex((rootItem) => rootItem.kind === "raw");
  rootItems.splice(firstRawIndex === -1 ? rootItems.length : firstRawIndex, 0, item);
}

function rewriteFilterColumnTagWithTransformedColId(
  filterColumnTag: XmlTag,
  currentStartColumn: number,
  nextStartColumn: number,
  targetColumnNumber: number,
  columnCount: number,
  mode: "shift" | "delete",
): string | null {
  const attributes = parseAttributes(filterColumnTag.attributesSource);
  const colIdIndex = attributes.findIndex(([name]) => name === "colId");
  if (colIdIndex === -1) {
    return filterColumnTag.source;
  }

  const currentColId = Number(attributes[colIdIndex]?.[1] ?? "");
  if (!Number.isInteger(currentColId) || currentColId < 0) {
    return filterColumnTag.source;
  }

  const currentColumnNumber = currentStartColumn + currentColId;
  const nextColumnNumber =
    mode === "shift"
      ? shiftColumnNumber(currentColumnNumber, targetColumnNumber, columnCount)
      : deleteShiftColumnNumber(currentColumnNumber, targetColumnNumber, columnCount);

  if (nextColumnNumber === null || nextColumnNumber < nextStartColumn) {
    return null;
  }

  const nextAttributes = [...attributes];
  nextAttributes[colIdIndex] = ["colId", String(nextColumnNumber - nextStartColumn)];
  return filterColumnTag.innerXml === null
    ? buildSelfClosingXmlElement("filterColumn", nextAttributes)
    : buildXmlElement("filterColumn", nextAttributes, filterColumnTag.innerXml);
}

function rewriteSortConditionTagWithTransformedRef(
  sortConditionTag: XmlTag,
  transformRange: (range: string) => string | null,
): string | null {
  const attributes = parseAttributes(sortConditionTag.attributesSource);
  const refIndex = attributes.findIndex(([name]) => name === "ref");
  if (refIndex === -1) {
    return null;
  }

  const currentRef = attributes[refIndex]?.[1] ?? "";
  const nextRef = transformRange(normalizeRangeRef(currentRef));
  if (nextRef === null) {
    return null;
  }

  const nextAttributes = [...attributes];
  nextAttributes[refIndex] = ["ref", nextRef];
  return sortConditionTag.innerXml === null
    ? buildSelfClosingXmlElement("sortCondition", nextAttributes)
    : buildXmlElement("sortCondition", nextAttributes, sortConditionTag.innerXml);
}

function rewriteFilterColumnSourceForRange(source: string, columnNumber: number, range: string): string {
  const filterColumnTag = findFirstXmlTag(source, "filterColumn");
  if (!filterColumnTag) {
    return source;
  }

  const { startColumn } = parseRangeRef(range);
  const attributes = parseAttributes(filterColumnTag.attributesSource);
  const nextAttributes = [...attributes];
  const colIdIndex = nextAttributes.findIndex(([name]) => name === "colId");
  const nextColId = String(columnNumber - startColumn);

  if (colIdIndex === -1) {
    nextAttributes.push(["colId", nextColId]);
  } else {
    nextAttributes[colIdIndex] = ["colId", nextColId];
  }

  return filterColumnTag.innerXml === null
    ? buildSelfClosingXmlElement("filterColumn", nextAttributes)
    : buildXmlElement("filterColumn", nextAttributes, filterColumnTag.innerXml);
}

function rewriteSortStateTagForRange(sortStateTag: XmlTag, range: string): string | null {
  const normalizedRange = normalizeRangeRef(range);
  const { startRow, endRow } = parseRangeRef(normalizedRange);
  const childTags = findTopLevelXmlTags(sortStateTag.innerXml ?? "");
  let validSortConditionCount = 0;
  const nextChildXml = childTags
    .map((childTag) => {
      if (childTag.tagName !== "sortCondition") {
        return childTag.source;
      }

      const conditionRef = getTagAttr(childTag, "ref");
      if (!conditionRef) {
        return null;
      }

      const { startColumn } = parseRangeRef(normalizeRangeRef(conditionRef));
      if (!isColumnInRange(startColumn, normalizedRange)) {
        return null;
      }

      validSortConditionCount += 1;
      const attributes = parseAttributes(childTag.attributesSource);
      const nextAttributes = [...attributes];
      const refIndex = nextAttributes.findIndex(([name]) => name === "ref");
      const nextRef = formatRangeRef(startRow, startColumn, endRow, startColumn);

      if (refIndex === -1) {
        nextAttributes.push(["ref", nextRef]);
      } else {
        nextAttributes[refIndex] = ["ref", nextRef];
      }

      return childTag.innerXml === null
        ? buildSelfClosingXmlElement("sortCondition", nextAttributes)
        : buildXmlElement("sortCondition", nextAttributes, childTag.innerXml);
    })
    .filter((childXml): childXml is string => childXml !== null)
    .join("");

  if (validSortConditionCount === 0) {
    return null;
  }

  return rebuildTagWithRef(sortStateTag, "sortState", normalizedRange, nextChildXml);
}

function rebuildTagWithRef(
  tag: XmlTag,
  tagName: "autoFilter" | "sortState",
  nextRef: string,
  innerXml: string | null,
): string {
  const attributes = parseAttributes(tag.attributesSource);
  const refIndex = attributes.findIndex(([name]) => name === "ref");
  const nextAttributes = [...attributes];

  if (refIndex === -1) {
    nextAttributes.push(["ref", nextRef]);
  } else {
    nextAttributes[refIndex] = ["ref", nextRef];
  }

  return innerXml === null || innerXml.length === 0
    ? buildSelfClosingXmlElement(tagName, nextAttributes)
    : buildXmlElement(tagName, nextAttributes, innerXml);
}

function shiftColumnNumber(columnNumber: number, targetColumnNumber: number, count: number): number {
  if (targetColumnNumber <= 0 || count <= 0) {
    return columnNumber;
  }

  return columnNumber >= targetColumnNumber ? columnNumber + count : columnNumber;
}

function deleteShiftColumnNumber(columnNumber: number, targetColumnNumber: number, count: number): number | null {
  if (targetColumnNumber <= 0 || count <= 0) {
    return columnNumber;
  }

  if (columnNumber >= targetColumnNumber && columnNumber <= targetColumnNumber + count - 1) {
    return null;
  }

  return columnNumber > targetColumnNumber + count - 1 ? columnNumber - count : columnNumber;
}

function appendOptionalNumberAttribute(
  attributes: Array<[string, string]>,
  name: string,
  value: number | undefined,
): void {
  if (value !== undefined) {
    attributes.push([name, String(value)]);
  }
}

function appendOptionalStringAttribute(
  attributes: Array<[string, string]>,
  name: string,
  value: string | undefined,
): void {
  if (value !== undefined) {
    attributes.push([name, value]);
  }
}

function parseIntegerAttr(tag: XmlTag, name: string): number | null {
  const value = getTagAttr(tag, name);
  if (value === undefined) {
    return null;
  }

  const parsed = Number(value);
  return Number.isInteger(parsed) ? parsed : null;
}

function parseNumberAttr(tag: XmlTag, name: string): number | null {
  const value = getTagAttr(tag, name);
  if (value === undefined) {
    return null;
  }

  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseOptionalXmlBoolean(value: string | undefined): boolean | null {
  if (value === undefined) {
    return null;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function isDateTimeGrouping(value: string | undefined): value is DateGroupItem["dateTimeGrouping"] {
  return value === "year" || value === "month" || value === "day" || value === "hour" || value === "minute" || value === "second";
}

function escapeAutoFilterPattern(value: string): string {
  return value.replaceAll("~", "~~").replaceAll("*", "~*").replaceAll("?", "~?");
}

function unescapeAutoFilterPattern(value: string): string {
  let nextValue = "";

  for (let index = 0; index < value.length; index += 1) {
    const current = value[index];
    if (current === "~" && index + 1 < value.length) {
      nextValue += value[index + 1];
      index += 1;
      continue;
    }

    nextValue += current;
  }

  return nextValue;
}

function findTopLevelXmlTags(xml: string): XmlTag[] {
  const tags: XmlTag[] = [];
  let searchStart = 0;

  while (searchStart < xml.length) {
    const tagStart = xml.indexOf("<", searchStart);
    if (tagStart === -1) {
      break;
    }

    if (xml.startsWith("<!--", tagStart)) {
      const commentEnd = xml.indexOf("-->", tagStart + 4);
      searchStart = commentEnd === -1 ? xml.length : commentEnd + 3;
      continue;
    }

    if (xml.startsWith("<?", tagStart)) {
      const declarationEnd = xml.indexOf("?>", tagStart + 2);
      searchStart = declarationEnd === -1 ? xml.length : declarationEnd + 2;
      continue;
    }

    if (xml[tagStart + 1] === "/") {
      searchStart = tagStart + 2;
      continue;
    }

    const tagNameMatch = xml.slice(tagStart + 1).match(/^([A-Za-z_][\w:.-]*)/);
    if (!tagNameMatch) {
      searchStart = tagStart + 1;
      continue;
    }

    const tagName = tagNameMatch[1];
    if (!isTagBoundary(xml, tagStart + 1 + tagName.length)) {
      searchStart = tagStart + 1 + tagName.length;
      continue;
    }

    const tagOpenEnd = findTagOpenEnd(xml, tagStart + 1);
    if (tagOpenEnd === -1) {
      break;
    }

    const tagOpenSource = xml.slice(tagStart + 1 + tagName.length, tagOpenEnd);
    const selfClosing = isSelfClosingTagSource(tagOpenSource);

    if (selfClosing) {
      tags.push({
        attributesSource: trimTagAttributesSource(tagOpenSource),
        end: tagOpenEnd + 1,
        innerXml: null,
        selfClosing: true,
        source: xml.slice(tagStart, tagOpenEnd + 1),
        start: tagStart,
        tagName,
      });
      searchStart = tagOpenEnd + 1;
      continue;
    }

    const closingPattern = `</${tagName}>`;
    const closeStart = xml.indexOf(closingPattern, tagOpenEnd + 1);
    if (closeStart === -1) {
      break;
    }

    tags.push({
      attributesSource: trimTagAttributesSource(tagOpenSource),
      end: closeStart + closingPattern.length,
      innerXml: xml.slice(tagOpenEnd + 1, closeStart),
      selfClosing: false,
      source: xml.slice(tagStart, closeStart + closingPattern.length),
      start: tagStart,
      tagName,
    });
    searchStart = closeStart + closingPattern.length;
  }

  return tags;
}

function findTagOpenEnd(xml: string, start: number): number {
  let quote: number | null = null;

  for (let index = start; index < xml.length; index += 1) {
    const code = xml.charCodeAt(index);

    if (quote !== null) {
      if (code === quote) {
        quote = null;
      }
      continue;
    }

    if (code === 34 || code === 39) {
      quote = code;
      continue;
    }

    if (code === 62) {
      return index;
    }
  }

  return -1;
}

function trimTagAttributesSource(source: string): string {
  let end = source.length;
  while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
    end -= 1;
  }

  if (end > 0 && source.charCodeAt(end - 1) === 47) {
    end -= 1;
    while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
      end -= 1;
    }
  }

  let start = 0;
  while (start < end && isXmlWhitespaceCode(source.charCodeAt(start))) {
    start += 1;
  }

  return source.slice(start, end);
}

function isTagBoundary(xml: string, index: number): boolean {
  if (index >= xml.length) {
    return true;
  }

  const code = xml.charCodeAt(index);
  return code === 47 || code === 62 || isXmlWhitespaceCode(code);
}

function isSelfClosingTagSource(source: string): boolean {
  let index = source.length - 1;

  while (index >= 0 && isXmlWhitespaceCode(source.charCodeAt(index))) {
    index -= 1;
  }

  return index >= 0 && source.charCodeAt(index) === 47;
}

function isXmlWhitespaceCode(code: number): boolean {
  return code === 9 || code === 10 || code === 13 || code === 32;
}
