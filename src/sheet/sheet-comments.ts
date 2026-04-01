import type { SheetComment } from "../types.js";
import { XlsxError } from "../errors.js";
import { basenamePosix, dirnamePosix, resolveRelationshipTarget } from "../utils/path.js";
import { decodeXmlText, escapeRegex, escapeXmlText, parseAttributes, serializeAttributes } from "../utils/xml.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "../utils/xml-read.js";
import { compareCellAddresses, normalizeCellAddress, splitCellAddress } from "./sheet-address.js";
import { buildSelfClosingXmlElement, findWorksheetChildInsertionIndex, replaceXmlTagSource } from "./sheet-xml.js";

export const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
export const COMMENTS_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
export const VML_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.vmlDrawing";
export const VML_DRAWING_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

const LEGACY_DRAWING_FOLLOWING_TAGS = [
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];

export interface ParsedComments {
  authors: string[];
  comments: SheetComment[];
}

export interface SheetCommentParts {
  commentsPath: string | null;
  commentsRelationshipId: string | null;
  legacyDrawingRelationshipId: string | null;
  vmlPath: string | null;
  vmlRelationshipId: string | null;
}

export function parseCommentsXml(commentsXml: string): ParsedComments {
  const authorsTag = findFirstXmlTag(commentsXml, "authors");
  const commentListTag = findFirstXmlTag(commentsXml, "commentList");
  const authors = authorsTag?.innerXml
    ? findXmlTags(authorsTag.innerXml, "author").map((tag) => decodeXmlText(tag.innerXml ?? ""))
    : [];
  const comments = commentListTag?.innerXml
    ? findXmlTags(commentListTag.innerXml, "comment")
        .map((tag) => {
          const address = getTagAttr(tag, "ref");
          if (!address) {
            return null;
          }

          const authorId = Number(getTagAttr(tag, "authorId") ?? "-1");
          const textTag = findFirstXmlTag(tag.innerXml ?? "", "text");
          const text = textTag ? parseCommentText(textTag.innerXml ?? "") : "";

          return {
            address: normalizeCellAddress(address),
            author: Number.isInteger(authorId) && authorId >= 0 ? (authors[authorId] ?? null) : null,
            text,
          };
        })
        .filter((comment): comment is SheetComment => comment !== null)
    : [];

  comments.sort((left, right) => compareCellAddresses(left.address, right.address));
  return { authors, comments };
}

export function buildCommentsXml(comments: SheetComment[]): string {
  const normalizedComments = [...comments].sort((left, right) => compareCellAddresses(left.address, right.address));
  const authors: string[] = [];
  const authorIds = new Map<string, number>();

  for (const comment of normalizedComments) {
    const author = comment.author ?? "fastxlsx";
    if (!authorIds.has(author)) {
      authorIds.set(author, authors.length);
      authors.push(author);
    }
  }

  const authorsXml = authors.map((author) => `<author>${escapeXmlText(author)}</author>`).join("");
  const commentsXml = normalizedComments
    .map((comment) => {
      const author = comment.author ?? "fastxlsx";
      const authorId = authorIds.get(author);
      if (authorId === undefined) {
        throw new XlsxError(`Comment author not found: ${author}`);
      }

      return (
        `<comment ref="${escapeXmlText(comment.address)}" authorId="${authorId}">` +
        `<text><t>${escapeXmlText(comment.text)}</t></text>` +
        `</comment>`
      );
    })
    .join("");

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<authors>${authorsXml}</authors>` +
    `<commentList>${commentsXml}</commentList>` +
    `</comments>`
  );
}

export function buildCommentsVmlDrawingXml(comments: SheetComment[]): string {
  const shapesXml = [...comments]
    .sort((left, right) => compareCellAddresses(left.address, right.address))
    .map((comment, index) => {
      const { rowNumber, columnNumber } = splitCellAddress(comment.address);
      return (
        `<v:shape id="_x0000_s${1025 + index}" type="#_x0000_t202" ` +
        `style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:${index + 1};visibility:hidden" ` +
        `fillcolor="#ffffe1" o:insetmode="auto">` +
        `<v:fill color2="#ffffe1"/>` +
        `<v:shadow on="t" color="black" obscured="t"/>` +
        `<v:path o:connecttype="none"/>` +
        `<v:textbox style="mso-direction-alt:auto"/>` +
        `<x:ClientData ObjectType="Note">` +
        `<x:MoveWithCells/>` +
        `<x:SizeWithCells/>` +
        `<x:Row>${rowNumber - 1}</x:Row>` +
        `<x:Column>${columnNumber - 1}</x:Column>` +
        `</x:ClientData>` +
        `</v:shape>`
      );
    })
    .join("");

  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>` +
    `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">` +
    `<v:stroke joinstyle="miter"/>` +
    `<v:path gradientshapeok="t" o:connecttype="rect"/>` +
    `</v:shapetype>` +
    shapesXml +
    `</xml>`
  );
}

export function findSheetCommentParts(sheetXml: string, sheetPath: string, relationshipsXml: string): SheetCommentParts {
  const relationshipsPathDir = dirnamePosix(sheetPath);
  const commentsRelationship = findRelationshipByType(relationshipsXml, COMMENTS_RELATIONSHIP_TYPE);
  const legacyDrawingTag = findFirstXmlTag(sheetXml, "legacyDrawing");
  const legacyDrawingRelationshipId = legacyDrawingTag ? (getTagAttr(legacyDrawingTag, "r:id") ?? null) : null;
  const vmlRelationship =
    legacyDrawingRelationshipId !== null
      ? findRelationshipById(relationshipsXml, legacyDrawingRelationshipId)
      : findRelationshipByType(relationshipsXml, VML_DRAWING_RELATIONSHIP_TYPE);

  return {
    commentsPath:
      commentsRelationship?.target !== undefined
        ? resolveRelationshipTarget(relationshipsPathDir, commentsRelationship.target)
        : null,
    commentsRelationshipId: commentsRelationship?.id ?? null,
    legacyDrawingRelationshipId,
    vmlPath:
      vmlRelationship?.target !== undefined
        ? resolveRelationshipTarget(relationshipsPathDir, vmlRelationship.target)
        : null,
    vmlRelationshipId: vmlRelationship?.id ?? null,
  };
}

export function ensureLegacyDrawingInSheetXml(sheetXml: string, relationshipId: string): string {
  const legacyDrawingTag = findFirstXmlTag(sheetXml, "legacyDrawing");
  const nextLegacyDrawingXml = buildSelfClosingXmlElement("legacyDrawing", [["r:id", relationshipId]]);

  if (legacyDrawingTag) {
    return replaceXmlTagSource(sheetXml, legacyDrawingTag, nextLegacyDrawingXml);
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, LEGACY_DRAWING_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + nextLegacyDrawingXml + sheetXml.slice(insertionIndex);
}

export function removeLegacyDrawingFromSheetXml(sheetXml: string): string {
  const legacyDrawingTag = findFirstXmlTag(sheetXml, "legacyDrawing");
  if (!legacyDrawingTag) {
    return sheetXml;
  }

  return replaceXmlTagSource(sheetXml, legacyDrawingTag, "");
}

export function ensureDefaultContentType(contentTypesXml: string, extension: string, contentType: string): string {
  if (new RegExp(`<Default\\b[^>]*Extension\\s*=\\s*["']${escapeRegex(extension)}["']`, "i").test(contentTypesXml)) {
    return contentTypesXml;
  }

  const closingTag = "</Types>";
  const insertionIndex = contentTypesXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Content types file is missing </Types>");
  }

  return (
    contentTypesXml.slice(0, insertionIndex) +
    `<Default Extension="${escapeXmlText(extension)}" ContentType="${escapeXmlText(contentType)}"/>` +
    contentTypesXml.slice(insertionIndex)
  );
}

export function getNextCommentsPath(entryPaths: string[]): string {
  let nextIndex = 1;

  for (const path of entryPaths) {
    const match = path.match(/^xl\/comments(\d+)\.xml$/);
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `xl/comments${nextIndex}.xml`;
}

export function getNextVmlDrawingPath(entryPaths: string[]): string {
  let nextIndex = 1;

  for (const path of entryPaths) {
    const match = path.match(/^xl\/drawings\/vmlDrawing(\d+)\.vml$/);
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `xl/drawings/vmlDrawing${nextIndex}.vml`;
}

function parseCommentText(textXml: string): string {
  const runs = findXmlTags(textXml, "t");
  if (runs.length === 0) {
    return "";
  }

  return runs.map((tag) => decodeXmlText(tag.innerXml ?? "")).join("");
}

function findRelationshipByType(
  relationshipsXml: string,
  relationshipType: string,
): { id: string; target: string } | null {
  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    const id = getTagAttr(relationshipTag, "Id");
    const type = getTagAttr(relationshipTag, "Type");
    const target = getTagAttr(relationshipTag, "Target");

    if (id && type === relationshipType && target) {
      return { id, target };
    }
  }

  return null;
}

function findRelationshipById(
  relationshipsXml: string,
  relationshipId: string,
): { id: string; target: string } | null {
  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    const id = getTagAttr(relationshipTag, "Id");
    const target = getTagAttr(relationshipTag, "Target");

    if (id === relationshipId && target) {
      return { id, target };
    }
  }

  return null;
}
