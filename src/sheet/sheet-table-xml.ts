import { basenamePosix, dirnamePosix, resolveRelationshipTarget } from "../utils/path.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "../utils/xml-read.js";
import { parseAttributes } from "../utils/xml.js";
import { buildXmlElement, buildSelfClosingXmlElement, replaceXmlTagSource } from "./sheet-xml.js";

export interface TableReference {
  relationshipId: string;
  path: string;
}

export const EMPTY_RELATIONSHIPS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

export function listTableReferences(
  sheetXml: string,
  sheetPath: string,
  entryPaths: string[],
  readEntryText: (path: string) => string,
): TableReference[] {
  const sheetRelationshipIds = findXmlTags(sheetXml, "tablePart")
    .filter((tag) => tag.selfClosing)
    .map((tag) => getTagAttr(tag, "r:id"))
    .filter((relationshipId): relationshipId is string => relationshipId !== undefined);
  if (sheetRelationshipIds.length === 0) {
    return [];
  }

  const relationshipsPath = getSheetRelationshipsPath(sheetPath);
  if (!entryPaths.includes(relationshipsPath)) {
    return [];
  }

  const relationshipsXml = readEntryText(relationshipsPath);
  const baseDir = dirnamePosix(sheetPath);
  const tables: TableReference[] = [];

  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    if (!relationshipTag.selfClosing) {
      continue;
    }

    const relationshipId = getTagAttr(relationshipTag, "Id");
    const type = getTagAttr(relationshipTag, "Type");
    const target = getTagAttr(relationshipTag, "Target");

    if (
      !relationshipId ||
      !type ||
      !target ||
      !sheetRelationshipIds.includes(relationshipId) ||
      !/\/table$/.test(type)
    ) {
      continue;
    }

    tables.push({
      relationshipId,
      path: resolveRelationshipTarget(baseDir, target),
    });
  }

  return tables;
}

export function rewriteTableReferenceXml(
  tableXml: string,
  transformRange: (range: string) => string | null,
): string | null {
  const tableTag = findFirstXmlTag(tableXml, "table");
  if (!tableTag) {
    return tableXml;
  }

  const tableAttributes = parseAttributes(tableTag.attributesSource);
  const refIndex = tableAttributes.findIndex(([name]) => name === "ref");
  if (refIndex === -1) {
    return tableXml;
  }

  const currentRange = tableAttributes[refIndex]?.[1] ?? "";
  const nextRange = transformRange(currentRange);
  if (nextRange === null) {
    return null;
  }

  const nextTableAttributes = [...tableAttributes];
  nextTableAttributes[refIndex] = ["ref", nextRange];
  let nextTableXml = replaceXmlTagSource(
    tableXml,
    tableTag,
    buildXmlElement("table", nextTableAttributes, tableTag.innerXml ?? ""),
  );
  const autoFilterTag = findFirstXmlTag(nextTableXml, "autoFilter");

  if (autoFilterTag) {
    const attributes = parseAttributes(autoFilterTag.attributesSource);
    const autoFilterRefIndex = attributes.findIndex(([name]) => name === "ref");

    if (autoFilterRefIndex !== -1) {
      const autoFilterRange = attributes[autoFilterRefIndex]?.[1] ?? "";
      const nextAutoFilterRange = transformRange(autoFilterRange);

      if (nextAutoFilterRange === null) {
        nextTableXml = replaceXmlTagSource(nextTableXml, autoFilterTag, "");
      } else {
        const nextAttributes = [...attributes];
        nextAttributes[autoFilterRefIndex] = ["ref", nextAutoFilterRange];
        nextTableXml = replaceXmlTagSource(
          nextTableXml,
          autoFilterTag,
          buildSelfClosingXmlElement("autoFilter", nextAttributes),
        );
      }
    }
  }

  return nextTableXml;
}

export function getSheetRelationshipsPath(sheetPath: string): string {
  return `${dirnamePosix(sheetPath)}/_rels/${basenamePosix(sheetPath)}.rels`;
}
