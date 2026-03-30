import type { Sheet } from "../sheet.js";
import { deleteSheetFormulaReferences, renameSheetFormulaReferences } from "../sheet/sheet-structure.js";
import { XlsxError } from "../errors.js";
import type { SheetVisibility } from "../types.js";
import {
  buildDefinedNameTagSource,
  buildDefinedNameTagXml,
  rewriteDefinedNamesInWorkbookXml,
} from "./workbook-defined-names.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "../utils/xml-read.js";
import { decodeXmlText, escapeRegex, parseAttributes, serializeAttributes } from "../utils/xml.js";

export function getNextSheetId(workbookXml: string): number {
  let nextSheetId = 1;

  for (const sheetTag of findXmlTags(workbookXml, "sheet")) {
    const sheetId = getTagAttr(sheetTag, "sheetId");
    if (sheetId === undefined) {
      continue;
    }

    nextSheetId = Math.max(nextSheetId, Number(sheetId) + 1);
  }

  return nextSheetId;
}

export function getNextRelationshipId(relationshipsXml: string): string {
  let nextId = 1;

  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    const id = getTagAttr(relationshipTag, "Id");
    if (!id?.startsWith("rId")) {
      continue;
    }

    nextId = Math.max(nextId, Number(id.slice(3)) + 1);
  }

  return `rId${nextId}`;
}

export function getNextWorksheetPath(workbookDir: string, entryOrder: string[]): string {
  let nextIndex = 1;
  const prefix = workbookDir ? `${workbookDir}/worksheets/` : "worksheets/";

  for (const path of entryOrder) {
    const match = path.match(new RegExp(`^${escapeRegex(prefix)}sheet(\\d+)\\.xml$`));
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `${prefix}sheet${nextIndex}.xml`;
}

export function toRelationshipTarget(workbookDir: string, path: string): string {
  return workbookDir && path.startsWith(`${workbookDir}/`) ? path.slice(workbookDir.length + 1) : path;
}

export function insertBeforeClosingTag(xml: string, tagName: string, snippet: string): string {
  const closingTag = `</${tagName}>`;
  const insertionIndex = xml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError(`Missing closing tag: ${closingTag}`);
  }

  return xml.slice(0, insertionIndex) + snippet + xml.slice(insertionIndex);
}

export function renameSheetInWorkbookXml(
  workbookXml: string,
  relationshipId: string,
  currentSheetName: string,
  nextSheetName: string,
): string {
  const renamedWorkbookXml = rewriteSheetsInWorkbookXml(workbookXml, (sheetTag) => {
    const attributes = parseAttributes(sheetTag.attributesSource);
    const relationshipIndex = attributes.findIndex(([name]) => name === "r:id");

    if (relationshipIndex === -1 || attributes[relationshipIndex]?.[1] !== relationshipId) {
      return sheetTag.source;
    }

    const nextAttributes = attributes.map(([name, value]) => {
      if (name === "name") {
        return [name, nextSheetName] as [string, string];
      }

      return [name, value] as [string, string];
    });
    return buildSheetTagXml(nextAttributes);
  }).workbookXml;

  return rewriteDefinedNamesInWorkbookXml(renamedWorkbookXml, (tag) => {
    const nameText = decodeXmlText(tag.innerXml ?? "");
    const nextNameText = renameSheetFormulaReferences(nameText, currentSheetName, nextSheetName);
    return nextNameText === nameText ? tag.source : buildDefinedNameTagSource(tag.attributesSource, nextNameText);
  }).workbookXml;
}

export function reorderWorkbookXmlSheets(
  workbookXml: string,
  currentSheets: Sheet[],
  nextSheets: Sheet[],
): string {
  const sheetsTag = findFirstXmlTag(workbookXml, "sheets");
  if (!sheetsTag || sheetsTag.innerXml === null) {
    throw new XlsxError("Workbook is missing <sheets>");
  }

  const sheetNodes = new Map<string, string>();
  for (const sheetTag of findXmlTags(sheetsTag.innerXml, "sheet")) {
    const relationshipId = getTagAttr(sheetTag, "r:id");
    if (relationshipId) {
      sheetNodes.set(relationshipId, sheetTag.source);
    }
  }

  const reorderedSheetsXml = nextSheets
    .map((sheet) => {
      const sheetXml = sheetNodes.get(sheet.relationshipId);
      if (!sheetXml) {
        throw new XlsxError(`Sheet relationship not found: ${sheet.relationshipId}`);
      }

      return sheetXml;
    })
    .join("");
  const localSheetIdMap = buildLocalSheetIdMap(currentSheets, nextSheets);
  const nextWorkbookXml = replaceXmlTagSource(workbookXml, sheetsTag, `<sheets>${reorderedSheetsXml}</sheets>`);
  const nextActiveTab = localSheetIdMap.get(parseActiveSheetIndex(workbookXml, currentSheets.length));
  const rewrittenDefinedNamesWorkbookXml = rewriteDefinedNamesInWorkbookXml(nextWorkbookXml, (tag) => {
    const attributes = parseAttributes(tag.attributesSource);
    const localSheetIdIndex = attributes.findIndex(([name]) => name === "localSheetId");
    if (localSheetIdIndex === -1) {
      return tag.source;
    }

    const localSheetIdText = attributes[localSheetIdIndex]?.[1];
    if (localSheetIdText === undefined) {
      return tag.source;
    }

    const nextLocalSheetId = localSheetIdMap.get(Number(localSheetIdText));
    if (nextLocalSheetId === undefined) {
      return tag.source;
    }

    attributes[localSheetIdIndex] = ["localSheetId", String(nextLocalSheetId)];
    return buildDefinedNameTagXml(attributes, decodeXmlText(tag.innerXml ?? ""));
  }).workbookXml;

  if (nextActiveTab !== undefined) {
    return updateWorkbookViewActiveTab(rewrittenDefinedNamesWorkbookXml, nextActiveTab);
  }

  return rewrittenDefinedNamesWorkbookXml;
}

export function parseSheetVisibility(workbookXml: string, relationshipId: string): SheetVisibility {
  for (const sheetTag of findXmlTags(workbookXml, "sheet")) {
    if (getTagAttr(sheetTag, "r:id") !== relationshipId) {
      continue;
    }

    const state = getTagAttr(sheetTag, "state");
    if (state === "hidden" || state === "veryHidden") {
      return state;
    }

    return "visible";
  }

  throw new XlsxError(`Sheet relationship not found: ${relationshipId}`);
}

export function updateSheetVisibilityInWorkbookXml(
  workbookXml: string,
  relationshipId: string,
  visibility: SheetVisibility,
): string {
  const replacement = rewriteSheetsInWorkbookXml(workbookXml, (sheetTag) => {
    const attributes = parseAttributes(sheetTag.attributesSource);
    const relationshipIndex = attributes.findIndex(([name]) => name === "r:id");

    if (relationshipIndex === -1 || attributes[relationshipIndex]?.[1] !== relationshipId) {
      return sheetTag.source;
    }

    const withoutState = attributes.filter(([name]) => name !== "state");
    const nextAttributes =
      visibility === "visible"
        ? withoutState
        : [...withoutState, ["state", visibility] as [string, string]];
    return buildSheetTagXml(nextAttributes);
  });

  if (!replacement.changed) {
    throw new XlsxError(`Sheet relationship not found: ${relationshipId}`);
  }

  return replacement.workbookXml;
}

export function parseActiveSheetIndex(workbookXml: string, sheetCount: number): number {
  const workbookViewTag = findFirstXmlTag(workbookXml, "workbookView");
  const activeTabText = workbookViewTag ? getTagAttr(workbookViewTag, "activeTab") : undefined;
  const activeTab = activeTabText === undefined ? 0 : Number(activeTabText);

  if (!Number.isInteger(activeTab) || activeTab < 0 || activeTab >= sheetCount) {
    return 0;
  }

  return activeTab;
}

export function updateActiveSheetInWorkbookXml(workbookXml: string, activeSheetIndex: number): string {
  return updateWorkbookViewActiveTab(workbookXml, activeSheetIndex);
}

export function removeSheetFromWorkbookXml(
  workbookXml: string,
  relationshipId: string,
  deletedSheetName: string,
  deletedSheetIndex: number,
): string {
  const withoutSheet = rewriteSheetsInWorkbookXml(workbookXml, (sheetTag) =>
    getTagAttr(sheetTag, "r:id") === relationshipId ? "" : sheetTag.source,
  ).workbookXml;

  return rewriteDefinedNamesInWorkbookXml(withoutSheet, (tag) => {
    const attributes = parseAttributes(tag.attributesSource);
    const localSheetIdIndex = attributes.findIndex(([name]) => name === "localSheetId");
    const localSheetIdText = localSheetIdIndex === -1 ? undefined : attributes[localSheetIdIndex]?.[1];

    if (localSheetIdText !== undefined) {
      const localSheetId = Number(localSheetIdText);
      if (localSheetId === deletedSheetIndex) {
        return "";
      }

      if (localSheetId > deletedSheetIndex) {
        attributes[localSheetIdIndex] = ["localSheetId", String(localSheetId - 1)];
      }
    }

    const nameText = decodeXmlText(tag.innerXml ?? "");
    const nextNameText = deleteSheetFormulaReferences(nameText, deletedSheetName);
    if (nextNameText === nameText && localSheetIdText === undefined) {
      return tag.source;
    }

    return buildDefinedNameTagXml(attributes, nextNameText);
  }).workbookXml;
}

function buildLocalSheetIdMap(currentSheets: Array<{ relationshipId: string }>, nextSheets: Array<{ relationshipId: string }>): Map<number, number> {
  const nextIndexesByRelationshipId = new Map<string, number>();
  nextSheets.forEach((sheet, index) => {
    nextIndexesByRelationshipId.set(sheet.relationshipId, index);
  });

  const localSheetIdMap = new Map<number, number>();
  currentSheets.forEach((sheet, index) => {
    const nextIndex = nextIndexesByRelationshipId.get(sheet.relationshipId);
    if (nextIndex !== undefined) {
      localSheetIdMap.set(index, nextIndex);
    }
  });

  return localSheetIdMap;
}

function getXmlTagInnerStart(tag: XmlTag): number {
  if (tag.innerXml === null) {
    return tag.end;
  }

  return tag.end - tag.innerXml.length - `</${tag.tagName}>`.length;
}

function replaceXmlTagSource(xml: string, tag: XmlTag, nextSource: string): string {
  return xml.slice(0, tag.start) + nextSource + xml.slice(tag.end);
}

function replaceNestedXmlTagSource(xml: string, parentTag: XmlTag, childTag: XmlTag, nextSource: string): string {
  const parentInnerStart = getXmlTagInnerStart(parentTag);
  return (
    xml.slice(0, parentInnerStart + childTag.start) +
    nextSource +
    xml.slice(parentInnerStart + childTag.end)
  );
}

function buildSheetTagXml(attributes: Array<[string, string]>): string {
  return `<sheet ${serializeAttributes(attributes)}/>`;
}

function rewriteSheetsInWorkbookXml(
  workbookXml: string,
  transform: (tag: XmlTag) => string,
): { changed: boolean; workbookXml: string } {
  const sheetsTag = findFirstXmlTag(workbookXml, "sheets");
  if (!sheetsTag || sheetsTag.innerXml === null) {
    return { changed: false, workbookXml };
  }

  const nextSheetSources: string[] = [];
  let changed = false;

  for (const sheetTag of findXmlTags(sheetsTag.innerXml, "sheet")) {
    const nextSource = transform(sheetTag);
    if (nextSource !== sheetTag.source) {
      changed = true;
    }

    if (nextSource.length > 0) {
      nextSheetSources.push(nextSource);
    }
  }

  if (!changed) {
    return { changed: false, workbookXml };
  }

  return {
    changed: true,
    workbookXml: replaceXmlTagSource(workbookXml, sheetsTag, `<sheets>${nextSheetSources.join("")}</sheets>`),
  };
}

function updateWorkbookViewActiveTab(workbookXml: string, activeSheetIndex: number): string {
  const workbookViewXml = `<workbookView activeTab="${activeSheetIndex}"/>`;
  const bookViewsTag = findFirstXmlTag(workbookXml, "bookViews");

  if (!bookViewsTag) {
    const sheetsTag = findFirstXmlTag(workbookXml, "sheets");
    if (!sheetsTag) {
      return workbookXml;
    }

    return workbookXml.slice(0, sheetsTag.start) + `<bookViews>${workbookViewXml}</bookViews>` + workbookXml.slice(sheetsTag.start);
  }

  const workbookViewTag =
    bookViewsTag.innerXml === null ? null : findFirstXmlTag(bookViewsTag.innerXml, "workbookView");
  if (!workbookViewTag) {
    return replaceXmlTagSource(workbookXml, bookViewsTag, `<bookViews>${workbookViewXml}</bookViews>`);
  }

  const attributes = parseAttributes(workbookViewTag.attributesSource);
  const activeTabIndex = attributes.findIndex(([name]) => name === "activeTab");
  if (activeTabIndex === -1) {
    attributes.push(["activeTab", String(activeSheetIndex)]);
  } else {
    attributes[activeTabIndex] = ["activeTab", String(activeSheetIndex)];
  }

  const serializedAttributes = serializeAttributes(attributes);
  return replaceNestedXmlTagSource(
    workbookXml,
    bookViewsTag,
    workbookViewTag,
    `<workbookView${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`,
  );
}
