import { XlsxError } from "./errors.js";
import type { DefinedName } from "./types.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "./utils/xml-read.js";
import { decodeXmlText, escapeXmlText, serializeAttributes } from "./utils/xml.js";

export function parseDefinedNames(
  workbookXml: string,
  sheets: Array<{ name: string }>,
): DefinedName[] {
  return findXmlTags(workbookXml, "definedName")
    .filter((tag) => tag.innerXml !== null)
    .map((tag) => {
      const localSheetIdText = getTagAttr(tag, "localSheetId");
      const localSheetId = localSheetIdText === undefined ? null : Number(localSheetIdText);
      return {
        hidden: getTagAttr(tag, "hidden") === "1",
        name: getTagAttr(tag, "name") ?? "",
        scope: localSheetId === null ? null : (sheets[localSheetId]?.name ?? null),
        value: decodeXmlText(tag.innerXml ?? ""),
      };
    })
    .filter((definedName) => definedName.name.length > 0);
}

export function buildDefinedNameTagSource(attributesSource: string, value: string): string {
  return `<definedName${attributesSource ? ` ${attributesSource}` : ""}>${escapeXmlText(value)}</definedName>`;
}

export function buildDefinedNameTagXml(attributes: Array<[string, string]>, value: string): string {
  const serializedAttributes = serializeAttributes(attributes);
  return `<definedName${serializedAttributes ? ` ${serializedAttributes}` : ""}>${escapeXmlText(value)}</definedName>`;
}

export function rewriteDefinedNamesInWorkbookXml(
  workbookXml: string,
  transform: (tag: XmlTag) => string,
): { changed: boolean; workbookXml: string } {
  const definedNamesTag = findFirstXmlTag(workbookXml, "definedNames");
  if (!definedNamesTag || definedNamesTag.innerXml === null) {
    return { changed: false, workbookXml };
  }

  const nextDefinedNameSources: string[] = [];
  let changed = false;

  for (const definedNameTag of findXmlTags(definedNamesTag.innerXml, "definedName")) {
    const nextSource = transform(definedNameTag);
    if (nextSource !== definedNameTag.source) {
      changed = true;
    }

    if (nextSource.length > 0) {
      nextDefinedNameSources.push(nextSource);
    }
  }

  if (!changed) {
    return { changed: false, workbookXml };
  }

  if (nextDefinedNameSources.length === 0) {
    return {
      changed: true,
      workbookXml: workbookXml.slice(0, definedNamesTag.start) + workbookXml.slice(definedNamesTag.end),
    };
  }

  return {
    changed: true,
    workbookXml: replaceXmlTagSource(
      workbookXml,
      definedNamesTag,
      `<definedNames>${nextDefinedNameSources.join("")}</definedNames>`,
    ),
  };
}

export function buildDefinedNameXml(name: string, value: string, localSheetId: number | null): string {
  const attributes: Array<[string, string]> = [["name", name]];
  if (localSheetId !== null) {
    attributes.push(["localSheetId", String(localSheetId)]);
  }

  return `<definedName ${serializeAttributes(attributes)}>${escapeXmlText(value)}</definedName>`;
}

export function insertDefinedNameIntoWorkbookXml(workbookXml: string, definedNameXml: string): string {
  const definedNamesTag = findFirstXmlTag(workbookXml, "definedNames");
  if (definedNamesTag) {
    const insertionIndex = definedNamesTag.end - "</definedNames>".length;
    return workbookXml.slice(0, insertionIndex) + definedNameXml + workbookXml.slice(insertionIndex);
  }

  return insertBeforeClosingTag(workbookXml, "workbook", `<definedNames>${definedNameXml}</definedNames>`);
}

export function removeDefinedNameFromWorkbookXml(
  workbookXml: string,
  name: string,
  localSheetId: number | null,
): string {
  return rewriteDefinedNamesInWorkbookXml(workbookXml, (tag) => {
    const candidateName = getTagAttr(tag, "name");
    const candidateLocalSheetIdText = getTagAttr(tag, "localSheetId");
    const candidateLocalSheetId = candidateLocalSheetIdText === undefined ? null : Number(candidateLocalSheetIdText);
    return candidateName === name && candidateLocalSheetId === localSheetId ? "" : tag.source;
  }).workbookXml;
}

function replaceXmlTagSource(xml: string, tag: XmlTag, nextSource: string): string {
  return xml.slice(0, tag.start) + nextSource + xml.slice(tag.end);
}

function insertBeforeClosingTag(xml: string, tagName: string, snippet: string): string {
  const closingTag = `</${tagName}>`;
  const insertionIndex = xml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError(`Missing closing tag: ${closingTag}`);
  }

  return xml.slice(0, insertionIndex) + snippet + xml.slice(insertionIndex);
}
