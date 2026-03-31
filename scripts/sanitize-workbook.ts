import { Zip } from "../src/zip.js";
import { escapeXmlText, getXmlAttr } from "../src/utils/xml.js";

const inputPath = process.argv[2];
const outputPath = process.argv[3] ?? inputPath;

if (!inputPath) {
  throw new Error("Usage: node --import tsx scripts/sanitize-workbook.ts <input.xlsx> [output.xlsx]");
}

const adapter = new Zip();
const entries = await adapter.readArchive(inputPath);
const sanitizedEntries = entries.map((entry) => ({
  path: entry.path,
  data: shouldSanitizeTextEntry(entry.path)
    ? new TextEncoder().encode(sanitizeEntryXml(entry.path, new TextDecoder().decode(entry.data)))
    : new Uint8Array(entry.data),
}));

await adapter.writeArchive(outputPath, sanitizedEntries);

function shouldSanitizeTextEntry(path: string): boolean {
  return (
    path === "xl/sharedStrings.xml" ||
    path === "xl/workbook.xml" ||
    path === "docProps/core.xml" ||
    path === "docProps/app.xml" ||
    path === "docProps/custom.xml" ||
    /^xl\/comments\d+\.xml$/.test(path) ||
    /^xl\/externalLinks\/externalLink\d+\.xml$/.test(path) ||
    /^xl\/worksheets\/sheet\d+\.xml$/.test(path) ||
    /\.rels$/.test(path)
  );
}

function sanitizeEntryXml(path: string, xml: string): string {
  if (path === "xl/sharedStrings.xml") {
    return sanitizeSharedStringsXml(xml);
  }

  if (path === "xl/workbook.xml") {
    return sanitizeWorkbookXml(xml);
  }

  if (/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) {
    return sanitizeWorksheetXml(xml);
  }

  if (/\.rels$/.test(path)) {
    return sanitizeRelationshipsXml(xml);
  }

  if (/^xl\/externalLinks\/externalLink\d+\.xml$/.test(path)) {
    return sanitizeExternalLinkXml(xml);
  }

  if (/^xl\/comments\d+\.xml$/.test(path)) {
    return sanitizeCommentsXml(xml);
  }

  if (path === "docProps/core.xml") {
    return sanitizeCorePropertiesXml(xml);
  }

  if (path === "docProps/app.xml") {
    return sanitizeAppPropertiesXml(xml);
  }

  if (path === "docProps/custom.xml") {
    return sanitizeCustomPropertiesXml(xml);
  }

  return xml;
}

function sanitizeSharedStringsXml(xml: string): string {
  let stringIndex = 0;
  return xml.replace(/<t\b([^>]*?)(\/>|>([\s\S]*?)<\/t>)/g, (_match, attributesSource, tail, textSource) => {
    const length = tail === "/>" ? 0 : decodeXmlEntities(textSource).length;
    const sanitized = buildMaskedText(length, stringIndex);
    stringIndex += 1;
    return `<t${attributesSource}>${escapeXmlText(sanitized)}</t>`;
  });
}

function sanitizeWorksheetXml(xml: string): string {
  let cellIndex = 0;

  return xml.replace(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, (match, attributesSource, tail, innerXml) => {
    const normalizedAttributes = attributesSource.trim();
    const rawType = getXmlAttr(normalizedAttributes, "t") ?? null;

    if (tail === "/>") {
      cellIndex += 1;
      return match;
    }

    let nextInnerXml = innerXml;
    const localSeed = cellIndex;
    const valueMatch = innerXml.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
    const currentValue = valueMatch?.[1];

    nextInnerXml = nextInnerXml.replace(/<f\b([^>]*?)(\/>|>([\s\S]*?)<\/f>)/g, (_formulaMatch, formulaAttributesSource, formulaTail) => {
      if (formulaTail === "/>") {
        return `<f${formulaAttributesSource}/>`;
      }

      return `<f${formulaAttributesSource}>${escapeXmlText(buildMaskedFormula(rawType, currentValue, localSeed))}</f>`;
    });

    nextInnerXml = nextInnerXml.replace(/<v\b([^>]*?)(\/>|>([\s\S]*?)<\/v>)/g, (_valueMatch, valueAttributesSource, valueTail, valueSource) => {
      if (valueTail === "/>") {
        return `<v${valueAttributesSource}/>`;
      }

      return `<v${valueAttributesSource}>${escapeXmlText(buildMaskedValue(rawType, valueSource, localSeed))}</v>`;
    });

    if (rawType === "inlineStr") {
      let inlineIndex = 0;
      nextInnerXml = nextInnerXml.replace(/<t\b([^>]*?)(\/>|>([\s\S]*?)<\/t>)/g, (_textMatch, textAttributesSource, textTail, textSource) => {
        const length = textTail === "/>" ? 0 : decodeXmlEntities(textSource).length;
        const sanitized = buildMaskedText(length, localSeed + inlineIndex);
        inlineIndex += 1;
        return `<t${textAttributesSource}>${escapeXmlText(sanitized)}</t>`;
      });
    }

    cellIndex += 1;
    return replaceCellInnerXml(match, nextInnerXml);
  });
}

function sanitizeRelationshipsXml(xml: string): string {
  let externalIndex = 0;

  return xml.replace(/<Relationship\b([^>]*?)\/>/g, (match, attributesSource) => {
    const targetMode = getXmlAttr(attributesSource, "TargetMode");
    if (targetMode !== "External") {
      return match;
    }

    const nextTarget = `https://example.invalid/resource/${externalIndex + 1}`;
    externalIndex += 1;
    return match.replace(
      /\bTarget="([^"]*)"/,
      `Target="${escapeXmlText(nextTarget)}"`,
    );
  });
}

function sanitizeWorkbookXml(xml: string): string {
  let nextXml = xml.replace(/\burl="([^"]*)"/g, () => `url="${escapeXmlText("https://example.invalid/workbook")}"`);
  nextXml = nextXml.replace(/\bdocumentId="([^"]*)"/g, 'documentId="sanitized-document-id"');
  nextXml = nextXml.replace(/\bxr10:uidLastSave="([^"]*)"/g, 'xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"');
  return nextXml;
}

function sanitizeCorePropertiesXml(xml: string): string {
  const replacements: Record<string, string> = {
    "dc:title": "Sanitized Workbook",
    "dc:subject": "Sanitized",
    "dc:creator": "xlsx-ts",
    "cp:keywords": "sanitized",
    "dc:description": "Sanitized benchmark workbook",
    "cp:lastModifiedBy": "xlsx-ts",
    "cp:category": "benchmark",
    "cp:contentStatus": "sanitized",
  };

  let nextXml = xml;
  for (const [tagName, value] of Object.entries(replacements)) {
    nextXml = replaceTagText(nextXml, tagName, value);
  }
  return nextXml;
}

function sanitizeAppPropertiesXml(xml: string): string {
  return replaceTagText(replaceTagText(xml, "Company", "Sanitized"), "Manager", "Sanitized");
}

function sanitizeCustomPropertiesXml(xml: string): string {
  let stringIndex = 0;
  let nextXml = xml.replace(/<vt:lpwstr>([\s\S]*?)<\/vt:lpwstr>/g, (_match, textSource) => {
    const sanitized = buildMaskedText(Math.max(8, decodeXmlEntities(textSource).length), stringIndex);
    stringIndex += 1;
    return `<vt:lpwstr>${escapeXmlText(sanitized)}</vt:lpwstr>`;
  });

  nextXml = nextXml.replace(/<vt:i4>([\s\S]*?)<\/vt:i4>/g, (_match, valueSource) => {
    return `<vt:i4>${escapeXmlText(buildMaskedNumber(valueSource, 0))}</vt:i4>`;
  });

  nextXml = nextXml.replace(/<vt:bool>([\s\S]*?)<\/vt:bool>/g, "<vt:bool>false</vt:bool>");
  return nextXml;
}

function sanitizeCommentsXml(xml: string): string {
  let authorIndex = 0;
  let textIndex = 0;
  let nextXml = xml.replace(/<author>([\s\S]*?)<\/author>/g, () => {
    authorIndex += 1;
    return `<author>Author ${authorIndex}</author>`;
  });

  nextXml = nextXml.replace(/<t\b([^>]*?)(\/>|>([\s\S]*?)<\/t>)/g, (_match, attributesSource, tail, textSource) => {
    const length = tail === "/>" ? 0 : decodeXmlEntities(textSource).length;
    const sanitized = buildMaskedText(length, textIndex);
    textIndex += 1;
    return `<t${attributesSource}>${escapeXmlText(sanitized)}</t>`;
  });

  return nextXml;
}

function sanitizeExternalLinkXml(xml: string): string {
  let textIndex = 0;
  let numberIndex = 0;
  let sheetIndex = 0;
  let nextXml = xml.replace(/<sheetName\b([^>]*?)\/>/g, (_match, attributesSource) => {
    const currentIndex = sheetIndex;
    sheetIndex += 1;
    return replaceAttrValue(`<sheetName${attributesSource}/>`, "val", `Sheet${currentIndex + 1}`);
  });

  nextXml = nextXml.replace(/<cell\b([^>]*?)(\/>|>([\s\S]*?)<\/cell>)/g, (match, attributesSource, tail, innerXml) => {
    if (tail === "/>") {
      return match;
    }

    const rawType = getXmlAttr(attributesSource.trim(), "t") ?? null;
    const nextInnerXml = innerXml.replace(/<v\b([^>]*?)(\/>|>([\s\S]*?)<\/v>)/g, (_valueMatch, valueAttributesSource, valueTail, valueSource) => {
      if (valueTail === "/>") {
        return `<v${valueAttributesSource}/>`;
      }

      if (rawType === "str") {
        const sanitized = buildMaskedText(decodeXmlEntities(valueSource).length, textIndex);
        textIndex += 1;
        return `<v${valueAttributesSource}>${escapeXmlText(sanitized)}</v>`;
      }

      const sanitized = buildMaskedNumber(valueSource, numberIndex);
      numberIndex += 1;
      return `<v${valueAttributesSource}>${escapeXmlText(sanitized)}</v>`;
    });

    return replaceCellInnerXml(match, nextInnerXml, "cell");
  });

  return nextXml;
}

function replaceTagText(xml: string, tagName: string, value: string): string {
  const tagPattern = new RegExp(`<${escapeTagName(tagName)}\\b[^>]*>[\\s\\S]*?<\\/${escapeTagName(tagName)}>`);
  const exactPattern = new RegExp(`(<${escapeTagName(tagName)}\\b[^>]*>)[\\s\\S]*?(<\\/${escapeTagName(tagName)}>)`);

  return tagPattern.test(xml)
    ? xml.replace(exactPattern, `$1${escapeXmlText(value)}$2`)
    : xml;
}

function escapeTagName(tagName: string): string {
  return tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function replaceCellInnerXml(cellXml: string, innerXml: string, tagName = "c"): string {
  const innerStart = cellXml.indexOf(">") + 1;
  const innerEnd = cellXml.lastIndexOf(`</${tagName}>`);
  return `${cellXml.slice(0, innerStart)}${innerXml}${cellXml.slice(innerEnd)}`;
}

function replaceAttrValue(tagSource: string, attributeName: string, value: string): string {
  const escapedValue = escapeXmlText(value);
  const quotedAttrPattern = new RegExp(`\\b${escapeTagName(attributeName)}="([^"]*)"`);

  if (quotedAttrPattern.test(tagSource)) {
    return tagSource.replace(quotedAttrPattern, `${attributeName}="${escapedValue}"`);
  }

  return tagSource;
}

function buildMaskedFormula(rawType: string | null, valueSource: string | undefined, seed: number): string {
  if (rawType === "str") {
    return `"${buildMaskedText(Math.max(4, decodeXmlEntities(valueSource ?? "").length || 4), seed)}"`;
  }

  if (rawType === "b") {
    return seed % 2 === 0 ? "TRUE" : "FALSE";
  }

  return `${1000 + (seed % 9000)}`;
}

function buildMaskedValue(rawType: string | null, valueSource: string, seed: number): string {
  if (rawType === "s") {
    return valueSource;
  }

  if (rawType === "inlineStr") {
    return valueSource;
  }

  if (rawType === "str") {
    const length = decodeXmlEntities(valueSource).length;
    return buildMaskedText(length, seed);
  }

  if (rawType === "b") {
    return seed % 2 === 0 ? "1" : "0";
  }

  return buildMaskedNumber(valueSource, seed);
}

function buildMaskedNumber(valueSource: string, seed: number): string {
  let digitIndex = 0;
  let next = "";

  for (const character of valueSource) {
    if (character < "0" || character > "9") {
      next += character;
      continue;
    }

    let digit = (seed + digitIndex * 7 + 3) % 10;
    if (digitIndex === 0 && digit === 0) {
      digit = 1;
    }

    next += String(digit);
    digitIndex += 1;
  }

  return digitIndex === 0 ? String(1000 + (seed % 9000)) : next;
}

function buildMaskedText(length: number, seed: number): string {
  if (length <= 0) {
    return "";
  }

  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let text = "";

  for (let index = 0; index < length; index += 1) {
    text += alphabet[(seed * 13 + index * 17) % alphabet.length];
  }

  return text;
}

function decodeXmlEntities(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", "\"")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}
