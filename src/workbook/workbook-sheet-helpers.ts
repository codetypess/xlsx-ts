import type { Sheet } from "../sheet.js";
import { XlsxError } from "../errors.js";
import { renameHyperlinkLocation } from "./workbook-sheet-package.js";
import { rewriteXmlTagsByName } from "./workbook-xml.js";
import { decodeXmlText, escapeXmlText, parseAttributes, serializeAttributes } from "../utils/xml.js";

export function normalizeSheetNameKey(sheetName: string): string {
  return sheetName.toUpperCase();
}

export function sheetNamesEqual(left: string, right: string): boolean {
  return normalizeSheetNameKey(left) === normalizeSheetNameKey(right);
}

export function findSheetIndexByName(sheets: Sheet[], sheetName: string): number {
  const normalizedSheetName = normalizeSheetNameKey(sheetName);
  return sheets.findIndex((candidate) => normalizeSheetNameKey(candidate.name) === normalizedSheetName);
}

export function requireSheetByName(sheets: Sheet[], sheetName: string): Sheet {
  const sheet = findSheetByName(sheets, sheetName);
  if (!sheet) {
    throw new XlsxError(`Sheet not found: ${sheetName}`);
  }

  return sheet;
}

export function findSheetByName(sheets: Sheet[], sheetName: string): Sheet | null {
  const sheetIndex = findSheetIndexByName(sheets, sheetName);
  return sheetIndex === -1 ? null : sheets[sheetIndex] ?? null;
}

export function resolveLocalSheetId(sheets: Sheet[], scope: string | null): number | null {
  if (scope === null) {
    return null;
  }

  const localSheetId = findSheetIndexByName(sheets, scope);
  if (localSheetId === -1) {
    throw new XlsxError(`Sheet not found: ${scope}`);
  }

  return localSheetId;
}

export function countVisibleSheets(
  workbookXml: string,
  sheets: Sheet[],
  parseSheetVisibility: (workbookXml: string, relationshipId: string) => "visible" | "hidden" | "veryHidden",
): number {
  return sheets.filter(
    (candidate) => parseSheetVisibility(workbookXml, candidate.relationshipId) === "visible",
  ).length;
}

export function rewriteFormulaXml(
  sheetXml: string,
  transformFormula: (formula: string) => string,
): { changed: boolean; sheetXml: string } {
  let changed = false;
  const nextSheetXml = rewriteXmlTagsByName(sheetXml, "f", (formulaTag) => {
    const formula = decodeXmlText(formulaTag.innerXml ?? "");
    const nextFormula = transformFormula(formula);

    if (nextFormula === formula) {
      return formulaTag.source;
    }

    changed = true;
    const serializedAttributes = serializeAttributes(parseAttributes(formulaTag.attributesSource));
    return `<f${serializedAttributes ? ` ${serializedAttributes}` : ""}>${escapeXmlText(nextFormula)}</f>`;
  });

  return { changed, sheetXml: nextSheetXml };
}

export function rewriteHyperlinkLocationXml(
  sheetXml: string,
  currentSheetName: string,
  nextSheetName: string,
): { changed: boolean; sheetXml: string } {
  let changed = false;
  const nextSheetXml = rewriteXmlTagsByName(sheetXml, "hyperlink", (hyperlinkTag) => {
    const attributes = parseAttributes(hyperlinkTag.attributesSource);
    const locationIndex = attributes.findIndex(([name]) => name === "location");

    if (locationIndex === -1) {
      return hyperlinkTag.source;
    }

    const location = attributes[locationIndex]?.[1] ?? "";
    const nextLocation = renameHyperlinkLocation(location, currentSheetName, nextSheetName);
    if (nextLocation === location) {
      return hyperlinkTag.source;
    }

    changed = true;
    const nextAttributes = [...attributes];
    nextAttributes[locationIndex] = ["location", nextLocation];
    const serializedAttributes = serializeAttributes(nextAttributes);
    return `<hyperlink${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
  });

  return { changed, sheetXml: nextSheetXml };
}
