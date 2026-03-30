import { XlsxError } from "./errors.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "./utils/xml-read.js";
import { escapeXmlText } from "./utils/xml.js";

export function buildEmptyWorksheetXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>`
  );
}

export function renameHyperlinkLocation(
  location: string,
  currentSheetName: string,
  nextSheetName: string,
): string {
  const hashPrefix = location.startsWith("#") ? "#" : "";
  const target = hashPrefix ? location.slice(1) : location;
  const bangIndex = target.indexOf("!");

  if (bangIndex === -1) {
    return location;
  }

  const sheetToken = target.slice(0, bangIndex);
  const normalizedSheetName =
    sheetToken.startsWith("'") && sheetToken.endsWith("'")
      ? sheetToken.slice(1, -1).replaceAll("''", "'")
      : sheetToken;

  if (normalizedSheetName !== currentSheetName) {
    return location;
  }

  return `${hashPrefix}${formatSheetNameForReference(nextSheetName)}${target.slice(bangIndex)}`;
}

export function removeRelationshipById(relationshipsXml: string, relationshipId: string): string {
  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    if (getTagAttr(relationshipTag, "Id") === relationshipId) {
      return replaceXmlTagSource(relationshipsXml, relationshipTag.source, "");
    }
  }

  return relationshipsXml;
}

export function removeContentTypeOverride(contentTypesXml: string, partPath: string): string {
  for (const overrideTag of findXmlTags(contentTypesXml, "Override")) {
    if (getTagAttr(overrideTag, "PartName") === `/${partPath}`) {
      return replaceXmlTagSource(contentTypesXml, overrideTag.source, "");
    }
  }

  return contentTypesXml;
}

export function updateAppSheetNames(appXml: string, sheetNames: string[]): string {
  const hasHeadingPairs = findFirstXmlTag(appXml, "HeadingPairs") !== null;
  const hasTitlesOfParts = findFirstXmlTag(appXml, "TitlesOfParts") !== null;

  if (!hasHeadingPairs && !hasTitlesOfParts) {
    return appXml;
  }

  let nextAppXml = appXml;

  if (hasHeadingPairs) {
    const nextHeadingPairs =
      `<HeadingPairs><vt:vector size="2" baseType="variant">` +
      `<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>` +
      `<vt:variant><vt:i4>${sheetNames.length}</vt:i4></vt:variant>` +
      `</vt:vector></HeadingPairs>`;
    const headingPairsTag = findFirstXmlTag(nextAppXml, "HeadingPairs");
    if (headingPairsTag) {
      nextAppXml = replaceXmlTagSource(nextAppXml, headingPairsTag.source, nextHeadingPairs);
    }
  }

  if (hasTitlesOfParts) {
    const nextTitlesOfParts =
      `<TitlesOfParts><vt:vector size="${sheetNames.length}" baseType="lpstr">` +
      sheetNames.map((sheetName) => `<vt:lpstr>${escapeXmlText(sheetName)}</vt:lpstr>`).join("") +
      `</vt:vector></TitlesOfParts>`;
    const titlesOfPartsTag = findFirstXmlTag(nextAppXml, "TitlesOfParts");
    if (titlesOfPartsTag) {
      nextAppXml = replaceXmlTagSource(nextAppXml, titlesOfPartsTag.source, nextTitlesOfParts);
    }
  }

  return nextAppXml;
}

export function formatSheetNameForReference(sheetName: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) {
    return sheetName;
  }

  return `'${sheetName.replaceAll("'", "''")}'`;
}

function replaceXmlTagSource(xml: string, source: string, nextSource: string): string {
  const index = xml.indexOf(source);
  if (index === -1) {
    throw new XlsxError("XML tag source not found");
  }

  return xml.slice(0, index) + nextSource + xml.slice(index + source.length);
}
