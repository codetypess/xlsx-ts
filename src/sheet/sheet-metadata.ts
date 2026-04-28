import type {
  DataValidation,
  Hyperlink,
  SetDataValidationOptions,
  SheetProtection,
  SheetProtectionOptions,
} from "../types.js";
import { XlsxError } from "../errors.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "../utils/xml-read.js";
import { decodeXmlText, escapeXmlText, parseAttributes, serializeAttributes } from "../utils/xml.js";
import {
  compareCellAddresses,
  normalizeCellAddress,
  normalizeRangeRef,
  normalizeSqref,
} from "./sheet-address.js";
import {
  buildXmlElement,
  buildSelfClosingXmlElement,
  buildCountedXmlContainer,
  findWorksheetChildInsertionIndex,
  replaceXmlTagSource,
} from "./sheet-xml.js";

export const HYPERLINK_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const AUTO_FILTER_FOLLOWING_TAGS = [
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

const DATA_VALIDATIONS_FOLLOWING_TAGS = [
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

const SHEET_PROTECTION_FOLLOWING_TAGS = [
  "protectedRanges",
  "scenarios",
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

const SHEET_PROTECTION_BOOLEAN_ATTRIBUTES = [
  "sheet",
  "objects",
  "scenarios",
  "formatCells",
  "formatColumns",
  "formatRows",
  "insertColumns",
  "insertRows",
  "insertHyperlinks",
  "deleteColumns",
  "deleteRows",
  "selectLockedCells",
  "sort",
  "autoFilter",
  "pivotTables",
  "selectUnlockedCells",
] as const;

export function parseSheetAutoFilter(sheetXml: string): string | null {
  const autoFilterTag = findFirstXmlTag(sheetXml, "autoFilter");
  if (!autoFilterTag) {
    return null;
  }

  const ref = getTagAttr(autoFilterTag, "ref");
  return ref ? normalizeRangeRef(ref) : null;
}

export function upsertAutoFilterInSheetXml(sheetXml: string, range: string): string {
  const normalizedRange = normalizeRangeRef(range);
  const autoFilterTag = findFirstXmlTag(sheetXml, "autoFilter");

  if (autoFilterTag) {
    const attributes = parseAttributes(autoFilterTag.attributesSource);
    const refIndex = attributes.findIndex(([name]) => name === "ref");
    const nextAttributes = [...attributes];

    if (refIndex === -1) {
      nextAttributes.push(["ref", normalizedRange]);
    } else {
      nextAttributes[refIndex] = ["ref", normalizedRange];
    }

    const autoFilterXml =
      autoFilterTag.innerXml === null
        ? buildSelfClosingXmlElement("autoFilter", nextAttributes)
        : buildXmlElement("autoFilter", nextAttributes, autoFilterTag.innerXml);
    return replaceXmlTagSource(sheetXml, autoFilterTag, autoFilterXml);
  }

  const autoFilterXml = `<autoFilter ref="${normalizedRange}"/>`;
  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, AUTO_FILTER_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + autoFilterXml + sheetXml.slice(insertionIndex);
}

export function removeAutoFilterFromSheetXml(sheetXml: string): string {
  let nextSheetXml = sheetXml;

  const autoFilterTag = findFirstXmlTag(nextSheetXml, "autoFilter");
  if (autoFilterTag) {
    nextSheetXml = replaceXmlTagSource(nextSheetXml, autoFilterTag, "");
  }

  const sortStateTag = findFirstXmlTag(nextSheetXml, "sortState");
  if (sortStateTag) {
    nextSheetXml = replaceXmlTagSource(nextSheetXml, sortStateTag, "");
  }

  return nextSheetXml;
}

export function parseSheetDataValidations(sheetXml: string): DataValidation[] {
  const dataValidationsTag = findFirstXmlTag(sheetXml, "dataValidations");
  if (!dataValidationsTag || dataValidationsTag.innerXml === null) {
    return [];
  }

  return parseDataValidationEntries(dataValidationsTag.innerXml)
    .map((validationTag) => {
      const sqref = getTagAttr(validationTag, "sqref");
      if (!sqref) {
        return null;
      }

      const errorTitle = getTagAttr(validationTag, "errorTitle");
      const error = getTagAttr(validationTag, "error");
      const promptTitle = getTagAttr(validationTag, "promptTitle");
      const prompt = getTagAttr(validationTag, "prompt");
      const formula1 = findFirstXmlTag(validationTag.innerXml ?? "", "formula1")?.innerXml;
      const formula2 = findFirstXmlTag(validationTag.innerXml ?? "", "formula2")?.innerXml;

      return {
        range: normalizeSqref(sqref),
        type: getTagAttr(validationTag, "type") ?? null,
        operator: getTagAttr(validationTag, "operator") ?? null,
        allowBlank: parseOptionalXmlBoolean(getTagAttr(validationTag, "allowBlank")),
        showInputMessage: parseOptionalXmlBoolean(getTagAttr(validationTag, "showInputMessage")),
        showErrorMessage: parseOptionalXmlBoolean(getTagAttr(validationTag, "showErrorMessage")),
        showDropDown: parseOptionalXmlBoolean(getTagAttr(validationTag, "showDropDown")),
        errorStyle: getTagAttr(validationTag, "errorStyle") ?? null,
        errorTitle: errorTitle ? decodeXmlText(errorTitle) : null,
        error: error ? decodeXmlText(error) : null,
        promptTitle: promptTitle ? decodeXmlText(promptTitle) : null,
        prompt: prompt ? decodeXmlText(prompt) : null,
        imeMode: getTagAttr(validationTag, "imeMode") ?? null,
        formula1: formula1 ? decodeXmlText(formula1) : null,
        formula2: formula2 ? decodeXmlText(formula2) : null,
      };
    })
    .filter((validation): validation is DataValidation => validation !== null);
}

export function buildDataValidationXml(range: string, options: SetDataValidationOptions): string {
  const attributes: Array<[string, string]> = [["sqref", normalizeSqref(range)]];
  appendOptionalAttribute(attributes, "type", options.type);
  appendOptionalAttribute(attributes, "operator", options.operator);
  appendOptionalBooleanAttribute(attributes, "allowBlank", options.allowBlank);
  appendOptionalBooleanAttribute(attributes, "showInputMessage", options.showInputMessage);
  appendOptionalBooleanAttribute(attributes, "showErrorMessage", options.showErrorMessage);
  appendOptionalBooleanAttribute(attributes, "showDropDown", options.showDropDown);
  appendOptionalAttribute(attributes, "errorStyle", options.errorStyle);
  appendOptionalAttribute(attributes, "errorTitle", options.errorTitle);
  appendOptionalAttribute(attributes, "error", options.error);
  appendOptionalAttribute(attributes, "promptTitle", options.promptTitle);
  appendOptionalAttribute(attributes, "prompt", options.prompt);
  appendOptionalAttribute(attributes, "imeMode", options.imeMode);

  const formulas: string[] = [];
  if (options.formula1 !== undefined) {
    formulas.push(`<formula1>${escapeXmlText(options.formula1)}</formula1>`);
  }
  if (options.formula2 !== undefined) {
    formulas.push(`<formula2>${escapeXmlText(options.formula2)}</formula2>`);
  }

  return formulas.length === 0
    ? `<dataValidation ${serializeAttributes(attributes)}/>`
    : `<dataValidation ${serializeAttributes(attributes)}>${formulas.join("")}</dataValidation>`;
}

export function upsertDataValidationInSheetXml(sheetXml: string, dataValidationXml: string, range: string): string {
  const normalizedRange = normalizeSqref(range);
  const dataValidationsTag = findFirstXmlTag(sheetXml, "dataValidations");
  const dataValidations = ((dataValidationsTag?.innerXml ?? "")
    ? parseDataValidationEntries(dataValidationsTag?.innerXml ?? "").map((validationTag) => ({
        range: normalizeSqref(getTagAttr(validationTag, "sqref") ?? ""),
        xml: validationTag.source,
      }))
    : []
  ).filter((validation) => validation.range !== normalizedRange);

  dataValidations.push({ range: normalizedRange, xml: dataValidationXml });
  const nextDataValidationsXml = buildCountedXmlContainer(
    "dataValidations",
    dataValidationsTag?.attributesSource ?? "",
    "count",
    dataValidations.map((validation) => validation.xml),
  );

  if (dataValidationsTag) {
    return replaceXmlTagSource(sheetXml, dataValidationsTag, nextDataValidationsXml);
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, DATA_VALIDATIONS_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + nextDataValidationsXml + sheetXml.slice(insertionIndex);
}

export function removeDataValidationFromSheetXml(sheetXml: string, range: string): string {
  const normalizedRange = normalizeSqref(range);
  const dataValidationsTag = findFirstXmlTag(sheetXml, "dataValidations");
  if (!dataValidationsTag || dataValidationsTag.innerXml === null) {
    return sheetXml;
  }

  const keptDataValidations = parseDataValidationEntries(dataValidationsTag.innerXml).filter(
    (validationTag) => normalizeSqref(getTagAttr(validationTag, "sqref") ?? "") !== normalizedRange,
  );

  const nextDataValidationsXml =
    keptDataValidations.length === 0
      ? ""
      : buildCountedXmlContainer(
          "dataValidations",
          dataValidationsTag.attributesSource,
          "count",
          keptDataValidations.map((validationTag) => validationTag.source),
        );

  return replaceXmlTagSource(sheetXml, dataValidationsTag, nextDataValidationsXml);
}

export function parseSheetProtection(sheetXml: string): SheetProtection | null {
  const protectionTag = findFirstXmlTag(sheetXml, "sheetProtection");
  if (!protectionTag) {
    return null;
  }

  return {
    autoFilter: parseOptionalXmlBoolean(getTagAttr(protectionTag, "autoFilter")),
    deleteColumns: parseOptionalXmlBoolean(getTagAttr(protectionTag, "deleteColumns")),
    deleteRows: parseOptionalXmlBoolean(getTagAttr(protectionTag, "deleteRows")),
    formatCells: parseOptionalXmlBoolean(getTagAttr(protectionTag, "formatCells")),
    formatColumns: parseOptionalXmlBoolean(getTagAttr(protectionTag, "formatColumns")),
    formatRows: parseOptionalXmlBoolean(getTagAttr(protectionTag, "formatRows")),
    insertColumns: parseOptionalXmlBoolean(getTagAttr(protectionTag, "insertColumns")),
    insertHyperlinks: parseOptionalXmlBoolean(getTagAttr(protectionTag, "insertHyperlinks")),
    insertRows: parseOptionalXmlBoolean(getTagAttr(protectionTag, "insertRows")),
    objects: parseOptionalXmlBoolean(getTagAttr(protectionTag, "objects")),
    passwordHash: getTagAttr(protectionTag, "password") ?? null,
    pivotTables: parseOptionalXmlBoolean(getTagAttr(protectionTag, "pivotTables")),
    scenarios: parseOptionalXmlBoolean(getTagAttr(protectionTag, "scenarios")),
    selectLockedCells: parseOptionalXmlBoolean(getTagAttr(protectionTag, "selectLockedCells")),
    selectUnlockedCells: parseOptionalXmlBoolean(getTagAttr(protectionTag, "selectUnlockedCells")),
    sheet: parseOptionalXmlBoolean(getTagAttr(protectionTag, "sheet")) ?? true,
    sort: parseOptionalXmlBoolean(getTagAttr(protectionTag, "sort")),
  };
}

export function upsertSheetProtectionInSheetXml(
  sheetXml: string,
  options: SheetProtectionOptions = {},
): string {
  const attributes: Array<[string, string]> = [["sheet", "1"]];

  for (const name of SHEET_PROTECTION_BOOLEAN_ATTRIBUTES) {
    if (name === "sheet") {
      continue;
    }

    const value = options[name];
    if (value !== undefined) {
      attributes.push([name, value ? "1" : "0"]);
    }
  }

  if (options.passwordHash !== undefined) {
    attributes.push(["password", options.passwordHash]);
  }

  const protectionXml = `<sheetProtection ${serializeAttributes(attributes)}/>`;
  const protectionTag = findFirstXmlTag(sheetXml, "sheetProtection");

  if (protectionTag) {
    return replaceXmlTagSource(sheetXml, protectionTag, protectionXml);
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, SHEET_PROTECTION_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + protectionXml + sheetXml.slice(insertionIndex);
}

export function removeSheetProtectionFromSheetXml(sheetXml: string): string {
  const protectionTag = findFirstXmlTag(sheetXml, "sheetProtection");
  if (!protectionTag) {
    return sheetXml;
  }

  return replaceXmlTagSource(sheetXml, protectionTag, "");
}

export function parseSheetHyperlinks(
  sheetXml: string,
  relationshipTargets: Map<string, string>,
): Hyperlink[] {
  return findXmlTags(sheetXml, "hyperlink").map((tag) => {
    const address = getTagAttr(tag, "ref");
    const relationshipId = getTagAttr(tag, "r:id");
    const location = getTagAttr(tag, "location");
    const tooltip = getTagAttr(tag, "tooltip") ?? null;

    if (!address) {
      return null;
    }

    if (relationshipId) {
      const target = relationshipTargets.get(relationshipId);
      if (!target) {
        return null;
      }

      return {
        address: normalizeCellAddress(address),
        target,
        tooltip,
        type: "external" as const,
      };
    }

    if (!location) {
      return null;
    }

    return {
      address: normalizeCellAddress(address),
      target: location,
      tooltip,
      type: "internal" as const,
    };
  }).filter((hyperlink): hyperlink is Hyperlink => hyperlink !== null);
}

export function parseHyperlinkRelationshipTargets(relationshipsXml: string): Map<string, string> {
  const targets = new Map<string, string>();

  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    if (!relationshipTag.selfClosing) {
      continue;
    }

    const relationshipId = getTagAttr(relationshipTag, "Id");
    const type = getTagAttr(relationshipTag, "Type");
    const target = getTagAttr(relationshipTag, "Target");

    if (!relationshipId || !type || !target || type !== HYPERLINK_RELATIONSHIP_TYPE) {
      continue;
    }

    targets.set(relationshipId, decodeXmlText(target));
  }

  return targets;
}

export function getHyperlinkRelationshipId(sheetXml: string, address: string): string | null {
  const normalizedAddress = normalizeCellAddress(address);

  for (const hyperlinkTag of findXmlTags(sheetXml, "hyperlink")) {
    const ref = getTagAttr(hyperlinkTag, "ref");
    if (!ref || normalizeCellAddress(ref) !== normalizedAddress) {
      continue;
    }

    return getTagAttr(hyperlinkTag, "r:id") ?? null;
  }

  return null;
}

export function buildInternalHyperlinkXml(address: string, location: string, tooltip?: string): string {
  const attributes: Array<[string, string]> = [["ref", address], ["location", location]];
  if (tooltip) {
    attributes.push(["tooltip", tooltip]);
  }

  return `<hyperlink ${serializeAttributes(attributes)}/>`;
}

export function buildExternalHyperlinkXml(address: string, relationshipId: string, tooltip?: string): string {
  const attributes: Array<[string, string]> = [["ref", address], ["r:id", relationshipId]];
  if (tooltip) {
    attributes.push(["tooltip", tooltip]);
  }

  return `<hyperlink ${serializeAttributes(attributes)}/>`;
}

export function upsertHyperlinkInSheetXml(sheetXml: string, hyperlinkXml: string, address: string): string {
  const normalizedAddress = normalizeCellAddress(address);
  const hyperlinksTag = findFirstXmlTag(sheetXml, "hyperlinks");
  const hyperlinksInnerXml = hyperlinksTag?.innerXml ?? "";

  const hyperlinks = (hyperlinksInnerXml
    ? findXmlTags(hyperlinksInnerXml, "hyperlink").map((tag) => {
        const ref = getTagAttr(tag, "ref");
        return {
          address: ref ? normalizeCellAddress(ref) : "",
          xml: tag.source,
        };
      })
    : []
  ).filter((hyperlink) => hyperlink.address !== normalizedAddress);
  hyperlinks.push({ address: normalizedAddress, xml: hyperlinkXml });
  hyperlinks.sort((left, right) => compareCellAddresses(left.address, right.address));

  const nextHyperlinksXml = `<hyperlinks>${hyperlinks.map((hyperlink) => hyperlink.xml).join("")}</hyperlinks>`;

  if (hyperlinksTag) {
    return replaceXmlTagSource(sheetXml, hyperlinksTag, nextHyperlinksXml);
  }

  const closingTag = "</worksheet>";
  const insertionIndex = sheetXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return sheetXml.slice(0, insertionIndex) + nextHyperlinksXml + sheetXml.slice(insertionIndex);
}

export function removeHyperlinkFromSheetXml(sheetXml: string, address: string): string {
  const normalizedAddress = normalizeCellAddress(address);
  const hyperlinksTag = findFirstXmlTag(sheetXml, "hyperlinks");
  if (!hyperlinksTag) {
    return sheetXml;
  }

  const keptHyperlinks = findXmlTags(hyperlinksTag.innerXml ?? "", "hyperlink")
    .map((tag) => {
      const ref = getTagAttr(tag, "ref");
      return {
        address: ref ? normalizeCellAddress(ref) : "",
        xml: tag.source,
      };
    })
    .filter((hyperlink) => hyperlink.address !== normalizedAddress);

  const nextHyperlinksXml =
    keptHyperlinks.length === 0
      ? ""
      : `<hyperlinks>${keptHyperlinks.map((hyperlink) => hyperlink.xml).join("")}</hyperlinks>`;

  return replaceXmlTagSource(sheetXml, hyperlinksTag, nextHyperlinksXml);
}

function parseDataValidationEntries(innerXml: string): XmlTag[] {
  return findXmlTags(innerXml, "dataValidation");
}

function appendOptionalAttribute(attributes: Array<[string, string]>, name: string, value: string | undefined): void {
  if (value !== undefined) {
    attributes.push([name, value]);
  }
}

function appendOptionalBooleanAttribute(attributes: Array<[string, string]>, name: string, value: boolean | undefined): void {
  if (value !== undefined) {
    attributes.push([name, value ? "1" : "0"]);
  }
}

function parseOptionalXmlBoolean(value: string | undefined): boolean | null {
  if (value === undefined) {
    return null;
  }

  return value === "1" || value.toLowerCase() === "true";
}
