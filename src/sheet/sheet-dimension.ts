import { XlsxError } from "../errors.js";
import type { SheetIndex } from "./sheet-index.js";
import { formatRangeRef } from "./sheet-address.js";
import { getXmlTagInnerStart, replaceXmlTagSource } from "./sheet-xml.js";
import { findFirstXmlTag } from "../utils/xml-read.js";

export function updateDimensionRef(sheetIndex: SheetIndex): string {
  const usedRange = formatUsedRangeBounds(sheetIndex.usedBounds);
  const dimensionTag = findFirstXmlTag(sheetIndex.xml, "dimension");

  if (!usedRange) {
    if (!dimensionTag) {
      return sheetIndex.xml;
    }

    return replaceXmlTagSource(sheetIndex.xml, dimensionTag, "");
  }

  const dimensionXml = `<dimension ref="${usedRange}"/>`;

  if (dimensionTag) {
    return replaceXmlTagSource(sheetIndex.xml, dimensionTag, dimensionXml);
  }

  const worksheetTag = findFirstXmlTag(sheetIndex.xml, "worksheet");
  if (!worksheetTag) {
    throw new XlsxError("Worksheet is missing opening tag");
  }

  const worksheetInnerStart = getXmlTagInnerStart(worksheetTag);
  return sheetIndex.xml.slice(0, worksheetInnerStart) + dimensionXml + sheetIndex.xml.slice(worksheetInnerStart);
}

export function formatUsedRangeBounds(bounds: SheetIndex["usedBounds"]): string | null {
  return bounds ? formatRangeRef(bounds.minRow, bounds.minColumn, bounds.maxRow, bounds.maxColumn) : null;
}
