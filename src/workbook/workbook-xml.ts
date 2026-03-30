import { findXmlTags, type XmlTag } from "../utils/xml-read.js";

export function replaceXmlTagSource(xml: string, tag: XmlTag, nextSource: string): string {
  return xml.slice(0, tag.start) + nextSource + xml.slice(tag.end);
}

export function rewriteXmlTagsByName(
  xml: string,
  tagName: string,
  rewriteTag: (tag: XmlTag) => string,
): string {
  const tags = findXmlTags(xml, tagName);
  if (tags.length === 0) {
    return xml;
  }

  let nextXml = "";
  let cursor = 0;

  for (const tag of tags) {
    nextXml += xml.slice(cursor, tag.start);
    nextXml += rewriteTag(tag);
    cursor = tag.end;
  }

  nextXml += xml.slice(cursor);
  return nextXml;
}
