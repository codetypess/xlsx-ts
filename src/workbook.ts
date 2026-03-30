import type {
  ArchiveEntry,
  CellBorderColor,
  CellBorderColorPatch,
  CellBorderDefinition,
  CellBorderPatch,
  CellBorderSideDefinition,
  CellBorderSidePatch,
  CellFillColor,
  CellFillColorPatch,
  CellFillDefinition,
  CellFillPatch,
  CellFontColor,
  CellFontColorPatch,
  CellFontDefinition,
  CellFontPatch,
  CellNumberFormatDefinition,
  CellStyleAlignment,
  CellStyleAlignmentPatch,
  CellStyleDefinition,
  CellStylePatch,
  DefinedName,
  SetDefinedNameOptions,
  SheetVisibility,
} from "./types.js";
import { XlsxError } from "./errors.js";
import {
  Sheet,
} from "./sheet.js";
import {
  deleteFormulaReferences,
  deleteSheetFormulaReferences,
  renameSheetFormulaReferences,
  shiftFormulaReferences,
} from "./sheet/sheet-structure.js";
import { parseSharedStrings } from "./workbook/shared-strings.js";
import {
  buildDefinedNameTagSource,
  buildDefinedNameTagXml,
  buildDefinedNameXml,
  insertDefinedNameIntoWorkbookXml,
  parseDefinedNames,
  removeDefinedNameFromWorkbookXml,
  rewriteDefinedNamesInWorkbookXml,
} from "./workbook/workbook-defined-names.js";
import {
  getNextRelationshipId,
  getNextSheetId,
  getNextWorksheetPath,
  insertBeforeClosingTag,
  parseActiveSheetIndex,
  parseSheetVisibility,
  removeSheetFromWorkbookXml,
  renameSheetInWorkbookXml,
  reorderWorkbookXmlSheets,
  toRelationshipTarget,
  updateActiveSheetInWorkbookXml,
  updateSheetVisibilityInWorkbookXml,
} from "./workbook/workbook-sheet-metadata.js";
import {
  buildEmptyWorksheetXml,
  removeContentTypeOverride,
  removeRelationshipById,
  updateAppSheetNames,
} from "./workbook/workbook-sheet-package.js";
import {
  countVisibleSheets,
  requireSheetByName,
  resolveLocalSheetId,
  rewriteFormulaXml,
  rewriteHyperlinkLocationXml,
} from "./workbook/workbook-sheet-helpers.js";
import {
  assertCellBorderPatch,
  assertCellFillPatch,
  assertCellFontPatch,
  assertCellStylePatch,
  assertDefinedName,
  assertFormatCode,
  assertSheetIndex,
  assertSheetName,
  assertSheetVisibility,
  assertStyleId,
} from "./workbook/workbook-validation.js";
import {
  appendBorderToStylesXml,
  appendCellXfToStylesXml,
  appendFillToStylesXml,
  appendFontToStylesXml,
  replaceBorderInStylesXml,
  replaceCellXfInStylesXml,
  replaceFillInStylesXml,
  replaceFontInStylesXml,
  upsertNumberFormatInStylesXml,
} from "./workbook/workbook-styles-container.js";
import {
  buildPatchedBorderXml,
  buildPatchedCellXfXml,
  buildPatchedFillXml,
  buildPatchedFontXml,
  cloneCellBorderDefinition,
  cloneCellFillDefinition,
  cloneCellFontDefinition,
  cloneCellStyleDefinition,
  getNextCustomNumberFormatId,
} from "./workbook/workbook-styles-build.js";
import {
  parseStylesXml,
  type ParsedBorder,
  type ParsedCellStyle,
  type ParsedFill,
  type ParsedFont,
  type StylesCache,
} from "./workbook/workbook-styles-parse.js";
import { replaceXmlTagSource } from "./workbook/workbook-xml.js";
import { Zip } from "./zip.js";
import type { WorkbookContext } from "./workbook/workbook-context.js";
import { resolveWorkbookContext } from "./workbook/workbook-context.js";
import { basenamePosix, dirnamePosix } from "./utils/path.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "./utils/xml-read.js";
import {
  escapeXmlText,
  decodeXmlText,
  escapeRegex,
  getXmlAttr,
  parseAttributes,
  serializeAttributes,
} from "./utils/xml.js";

const XML_DECODER = new TextDecoder();
const XML_ENCODER = new TextEncoder();
const WORKSHEET_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const WORKSHEET_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

const BUILTIN_NUMBER_FORMATS = new Map<number, string>([
  [0, "General"],
  [1, "0"],
  [2, "0.00"],
  [3, "#,##0"],
  [4, "#,##0.00"],
  [9, "0%"],
  [10, "0.00%"],
  [11, "0.00E+00"],
  [12, "# ?/?"],
  [13, "# ??/??"],
  [14, "mm-dd-yy"],
  [15, "d-mmm-yy"],
  [16, "d-mmm"],
  [17, "mmm-yy"],
  [18, "h:mm AM/PM"],
  [19, "h:mm:ss AM/PM"],
  [20, "h:mm"],
  [21, "h:mm:ss"],
  [22, "m/d/yy h:mm"],
  [37, "#,##0 ;(#,##0)"],
  [38, "#,##0 ;[Red](#,##0)"],
  [39, "#,##0.00;(#,##0.00)"],
  [40, "#,##0.00;[Red](#,##0.00)"],
  [45, "mm:ss"],
  [46, "[h]:mm:ss"],
  [47, "mmss.0"],
  [48, "##0.0E+0"],
  [49, "@"],
]);

export class Workbook {
  private readonly adapter: Zip;
  private readonly entryOrder: string[];
  private readonly entries: Map<string, Uint8Array>;
  private workbookContext?: WorkbookContext;
  private sharedStringsCache?: string[];
  private stylesCache?: StylesCache | null;

  constructor(
    entries: Iterable<ArchiveEntry>,
    adapter = new Zip(),
    options: { cloneEntryData?: boolean } = {},
  ) {
    this.adapter = adapter;
    this.entries = new Map();
    this.entryOrder = [];
    const cloneEntryData = options.cloneEntryData ?? true;

    for (const entry of entries) {
      this.entryOrder.push(entry.path);
      this.entries.set(entry.path, cloneEntryData ? new Uint8Array(entry.data) : entry.data);
    }
  }

  static async open(filePath: string): Promise<Workbook> {
    const adapter = new Zip();
    const entries = await adapter.readArchive(filePath);
    return new Workbook(entries, adapter, { cloneEntryData: false });
  }

  static fromEntries(entries: Iterable<ArchiveEntry>): Workbook {
    return new Workbook(entries);
  }

  listEntries(): string[] {
    return [...this.entryOrder];
  }

  getSheets(): Sheet[] {
    return [...this.getWorkbookContext().sheets];
  }

  getSheet(sheetName: string): Sheet {
    return requireSheetByName(this.getWorkbookContext().sheets, sheetName);
  }

  getActiveSheet(): Sheet {
    const context = this.getWorkbookContext();
    const activeSheetIndex = parseActiveSheetIndex(this.readEntryText(context.workbookPath), context.sheets.length);
    return context.sheets[activeSheetIndex] ?? context.sheets[0]!;
  }

  getStyle(styleId: number): CellStyleDefinition | null {
    assertStyleId(styleId);
    return cloneCellStyleDefinition(this.getStylesCache()?.cellXfs[styleId]?.definition ?? null);
  }

  getNumberFormat(numFmtId: number): CellNumberFormatDefinition | null {
    assertStyleId(numFmtId);
    const styles = this.getStylesCache();
    const customCode = styles?.numberFormats.get(numFmtId);
    if (customCode !== undefined) {
      return {
        builtin: false,
        code: customCode,
        numFmtId,
      };
    }

    const builtinCode = BUILTIN_NUMBER_FORMATS.get(numFmtId);
    if (builtinCode !== undefined) {
      return {
        builtin: true,
        code: builtinCode,
        numFmtId,
      };
    }

    return null;
  }

  getFont(fontId: number): CellFontDefinition | null {
    assertStyleId(fontId);
    return cloneCellFontDefinition(this.getStylesCache()?.fonts[fontId]?.definition ?? null);
  }

  getFill(fillId: number): CellFillDefinition | null {
    assertStyleId(fillId);
    return cloneCellFillDefinition(this.getStylesCache()?.fills[fillId]?.definition ?? null);
  }

  getBorder(borderId: number): CellBorderDefinition | null {
    assertStyleId(borderId);
    return cloneCellBorderDefinition(this.getStylesCache()?.borders[borderId]?.definition ?? null);
  }

  updateNumberFormat(numFmtId: number, formatCode: string): void {
    assertStyleId(numFmtId);
    assertFormatCode(formatCode);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    if (!styles.numberFormats.has(numFmtId)) {
      if (BUILTIN_NUMBER_FORMATS.has(numFmtId)) {
        throw new XlsxError(`Cannot update builtin number format: ${numFmtId}`);
      }

      throw new XlsxError(`Number format not found: ${numFmtId}`);
    }

    this.writeEntryText(styles.path, upsertNumberFormatInStylesXml(styles.xml, numFmtId, formatCode));
  }

  cloneNumberFormat(numFmtId: number, formatCode?: string): number {
    assertStyleId(numFmtId);
    if (formatCode !== undefined) {
      assertFormatCode(formatCode);
    }

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceCode = formatCode ?? styles.numberFormats.get(numFmtId) ?? BUILTIN_NUMBER_FORMATS.get(numFmtId);
    if (sourceCode === undefined) {
      throw new XlsxError(`Number format not found: ${numFmtId}`);
    }

    const nextNumFmtId = getNextCustomNumberFormatId(styles.numberFormats);
    this.writeEntryText(styles.path, upsertNumberFormatInStylesXml(styles.xml, nextNumFmtId, sourceCode));
    return nextNumFmtId;
  }

  ensureNumberFormat(formatCode: string): number {
    assertFormatCode(formatCode);

    for (const [numFmtId, builtinCode] of BUILTIN_NUMBER_FORMATS) {
      if (builtinCode === formatCode) {
        return numFmtId;
      }
    }

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    for (const [numFmtId, existingCode] of styles.numberFormats) {
      if (existingCode === formatCode) {
        return numFmtId;
      }
    }

    const nextNumFmtId = getNextCustomNumberFormatId(styles.numberFormats);
    this.writeEntryText(styles.path, upsertNumberFormatInStylesXml(styles.xml, nextNumFmtId, formatCode));
    return nextNumFmtId;
  }

  updateFont(fontId: number, patch: CellFontPatch): void {
    assertStyleId(fontId);
    assertCellFontPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFont = styles.fonts[fontId];
    if (!sourceFont) {
      throw new XlsxError(`Font not found: ${fontId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceFontInStylesXml(styles.xml, fontId, buildPatchedFontXml(sourceFont, patch)),
    );
  }

  cloneFont(fontId: number, patch: CellFontPatch = {}): number {
    assertStyleId(fontId);
    assertCellFontPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFont = styles.fonts[fontId];
    if (!sourceFont) {
      throw new XlsxError(`Font not found: ${fontId}`);
    }

    const nextFontId = styles.fonts.length;
    this.writeEntryText(
      styles.path,
      appendFontToStylesXml(styles.xml, buildPatchedFontXml(sourceFont, patch)),
    );
    return nextFontId;
  }

  updateFill(fillId: number, patch: CellFillPatch): void {
    assertStyleId(fillId);
    assertCellFillPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFill = styles.fills[fillId];
    if (!sourceFill) {
      throw new XlsxError(`Fill not found: ${fillId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceFillInStylesXml(styles.xml, fillId, buildPatchedFillXml(sourceFill, patch)),
    );
  }

  cloneFill(fillId: number, patch: CellFillPatch = {}): number {
    assertStyleId(fillId);
    assertCellFillPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFill = styles.fills[fillId];
    if (!sourceFill) {
      throw new XlsxError(`Fill not found: ${fillId}`);
    }

    const nextFillId = styles.fills.length;
    this.writeEntryText(
      styles.path,
      appendFillToStylesXml(styles.xml, buildPatchedFillXml(sourceFill, patch)),
    );
    return nextFillId;
  }

  updateBorder(borderId: number, patch: CellBorderPatch): void {
    assertStyleId(borderId);
    assertCellBorderPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceBorder = styles.borders[borderId];
    if (!sourceBorder) {
      throw new XlsxError(`Border not found: ${borderId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceBorderInStylesXml(styles.xml, borderId, buildPatchedBorderXml(sourceBorder, patch)),
    );
  }

  cloneBorder(borderId: number, patch: CellBorderPatch = {}): number {
    assertStyleId(borderId);
    assertCellBorderPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceBorder = styles.borders[borderId];
    if (!sourceBorder) {
      throw new XlsxError(`Border not found: ${borderId}`);
    }

    const nextBorderId = styles.borders.length;
    this.writeEntryText(
      styles.path,
      appendBorderToStylesXml(styles.xml, buildPatchedBorderXml(sourceBorder, patch)),
    );
    return nextBorderId;
  }

  updateStyle(styleId: number, patch: CellStylePatch): void {
    assertStyleId(styleId);
    assertCellStylePatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceStyle = styles.cellXfs[styleId];
    if (!sourceStyle) {
      throw new XlsxError(`Style not found: ${styleId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceCellXfInStylesXml(styles.xml, styleId, buildPatchedCellXfXml(sourceStyle, patch)),
    );
  }

  cloneStyle(styleId: number, patch: CellStylePatch = {}): number {
    assertStyleId(styleId);
    assertCellStylePatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceStyle = styles.cellXfs[styleId];
    if (!sourceStyle) {
      throw new XlsxError(`Style not found: ${styleId}`);
    }

    const nextStyleId = styles.cellXfs.length;
    this.writeEntryText(
      styles.path,
      appendCellXfToStylesXml(styles.xml, buildPatchedCellXfXml(sourceStyle, patch)),
    );
    return nextStyleId;
  }

  setActiveSheet(sheetName: string): Sheet {
    const context = this.getWorkbookContext();
    const targetIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    if (targetIndex === -1) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    if (this.getSheetVisibility(sheetName) !== "visible") {
      throw new XlsxError(`Cannot activate hidden sheet: ${sheetName}`);
    }

    const workbookPath = context.workbookPath;
    this.writeEntryText(
      workbookPath,
      updateActiveSheetInWorkbookXml(this.readEntryText(workbookPath), targetIndex),
    );
    return this.getSheet(sheetName);
  }

  getSheetVisibility(sheetName: string): SheetVisibility {
    const context = this.getWorkbookContext();
    const sheet = requireSheetByName(context.sheets, sheetName);

    return parseSheetVisibility(this.readEntryText(context.workbookPath), sheet.relationshipId);
  }

  setSheetVisibility(sheetName: string, visibility: SheetVisibility): void {
    assertSheetVisibility(visibility);

    const context = this.getWorkbookContext();
    const sheet = requireSheetByName(context.sheets, sheetName);

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    const currentVisibility = parseSheetVisibility(workbookXml, sheet.relationshipId);

    if (currentVisibility === visibility) {
      return;
    }

    const visibleSheetCount = countVisibleSheets(workbookXml, context.sheets, parseSheetVisibility);
    if (currentVisibility === "visible" && visibility !== "visible" && visibleSheetCount === 1) {
      throw new XlsxError("Workbook must contain at least one visible sheet");
    }

    this.writeEntryText(
      workbookPath,
      updateSheetVisibilityInWorkbookXml(workbookXml, sheet.relationshipId, visibility),
    );
  }

  getDefinedNames(): DefinedName[] {
    const context = this.getWorkbookContext();
    return parseDefinedNames(this.readEntryText(context.workbookPath), context.sheets);
  }

  getDefinedName(name: string, scope?: string): string | null {
    const definedName = this.getDefinedNames().find(
      (candidate) => candidate.name === name && candidate.scope === (scope ?? null),
    );
    return definedName?.value ?? null;
  }

  setDefinedName(name: string, value: string, options: SetDefinedNameOptions = {}): void {
    assertDefinedName(name);

    const context = this.getWorkbookContext();
    const localSheetId = resolveLocalSheetId(context.sheets, options.scope ?? null);

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    const nextDefinedNameXml = buildDefinedNameXml(name, value, localSheetId);
    const replacement = rewriteDefinedNamesInWorkbookXml(workbookXml, (tag) => {
      const candidateName = getTagAttr(tag, "name");
      const candidateLocalSheetId = getTagAttr(tag, "localSheetId");
      const candidateScope = candidateLocalSheetId === undefined ? null : Number(candidateLocalSheetId);

      return candidateName === name && candidateScope === localSheetId ? nextDefinedNameXml : tag.source;
    });

    if (replacement.changed) {
      this.writeEntryText(workbookPath, replacement.workbookXml);
      return;
    }

    this.writeEntryText(workbookPath, insertDefinedNameIntoWorkbookXml(workbookXml, nextDefinedNameXml));
  }

  deleteDefinedName(name: string, scope?: string): void {
    const context = this.getWorkbookContext();
    const localSheetId = resolveLocalSheetId(context.sheets, scope ?? null);

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    const nextWorkbookXml = removeDefinedNameFromWorkbookXml(workbookXml, name, localSheetId);

    if (nextWorkbookXml !== workbookXml) {
      this.writeEntryText(workbookPath, nextWorkbookXml);
    }
  }

  renameSheet(currentSheetName: string, nextSheetName: string): Sheet {
    assertSheetName(nextSheetName);

    const context = this.getWorkbookContext();
    const renamedSheet = requireSheetByName(context.sheets, currentSheetName);

    if (currentSheetName === nextSheetName) {
      return renamedSheet;
    }

    if (context.sheets.some((sheet) => sheet.name === nextSheetName)) {
      throw new XlsxError(`Sheet already exists: ${nextSheetName}`);
    }

    for (const sheet of context.sheets) {
      this.rewriteSheetFormulaTexts(sheet.path, (formula) =>
        renameSheetFormulaReferences(formula, currentSheetName, nextSheetName),
      );
      this.rewriteSheetHyperlinkLocations(sheet.path, currentSheetName, nextSheetName);
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    this.writeEntryText(
      context.workbookPath,
      renameSheetInWorkbookXml(workbookXml, renamedSheet.relationshipId, currentSheetName, nextSheetName),
    );
    this.rewriteAppSheetNames(
      context.sheets.map((sheet) => (sheet.name === currentSheetName ? nextSheetName : sheet.name)),
    );
    renamedSheet.name = nextSheetName;
    return renamedSheet;
  }

  moveSheet(sheetName: string, targetIndex: number): Sheet {
    const context = this.getWorkbookContext();
    const sourceIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    requireSheetByName(context.sheets, sheetName);

    assertSheetIndex(targetIndex, context.sheets.length);
    if (sourceIndex === targetIndex) {
      return context.sheets[sourceIndex]!;
    }

    const nextSheets = [...context.sheets];
    const [movedSheet] = nextSheets.splice(sourceIndex, 1);
    nextSheets.splice(targetIndex, 0, movedSheet!);

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    this.writeEntryText(
      workbookPath,
      reorderWorkbookXmlSheets(workbookXml, context.sheets, nextSheets),
    );
    this.rewriteAppSheetNames(nextSheets.map((sheet) => sheet.name));
    return this.getSheet(sheetName);
  }

  addSheet(sheetName: string): Sheet {
    assertSheetName(sheetName);

    const context = this.getWorkbookContext();
    if (context.sheets.some((sheet) => sheet.name === sheetName)) {
      throw new XlsxError(`Sheet already exists: ${sheetName}`);
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    const workbookRelsXml = this.readEntryText(context.workbookRelsPath);
    const nextSheetId = getNextSheetId(workbookXml);
    const nextRelationshipId = getNextRelationshipId(workbookRelsXml);
    const nextSheetPath = getNextWorksheetPath(context.workbookDir, this.entryOrder);
    const relationshipTarget = toRelationshipTarget(context.workbookDir, nextSheetPath);
    const contentTypesXml = this.readEntryText("[Content_Types].xml");

    this.writeEntryText(nextSheetPath, buildEmptyWorksheetXml());
    this.writeEntryText(
      context.workbookPath,
      insertBeforeClosingTag(
        workbookXml,
        "sheets",
        `<sheet name="${escapeXmlText(sheetName)}" sheetId="${nextSheetId}" r:id="${nextRelationshipId}"/>`,
      ),
    );
    this.writeEntryText(
      context.workbookRelsPath,
      insertBeforeClosingTag(
        workbookRelsXml,
        "Relationships",
        `<Relationship Id="${nextRelationshipId}" Type="${WORKSHEET_RELATIONSHIP_TYPE}" Target="${escapeXmlText(relationshipTarget)}"/>`,
      ),
    );
    this.writeEntryText(
      "[Content_Types].xml",
      insertBeforeClosingTag(
        contentTypesXml,
        "Types",
        `<Override PartName="/${escapeXmlText(nextSheetPath)}" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`,
      ),
    );
    this.rewriteAppSheetNames([...context.sheets.map((sheet) => sheet.name), sheetName]);

    return this.getSheet(sheetName);
  }

  deleteSheet(sheetName: string): void {
    const context = this.getWorkbookContext();
    if (context.sheets.length === 1) {
      throw new XlsxError("Cannot delete the last sheet");
    }

    const deletedSheetIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    requireSheetByName(context.sheets, sheetName);

    const deletedSheet = context.sheets[deletedSheetIndex];
    if (!deletedSheet) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    const workbookRelsXml = this.readEntryText(context.workbookRelsPath);
    const contentTypesXml = this.readEntryText("[Content_Types].xml");

    for (const sheet of context.sheets) {
      if (sheet.path === deletedSheet.path) {
        continue;
      }

      this.rewriteSheetFormulaTexts(sheet.path, (formula) =>
        deleteSheetFormulaReferences(formula, sheetName),
      );
    }

    this.writeEntryText(
      context.workbookPath,
      removeSheetFromWorkbookXml(workbookXml, deletedSheet.relationshipId, sheetName, deletedSheetIndex),
    );
    this.writeEntryText(
      context.workbookRelsPath,
      removeRelationshipById(workbookRelsXml, deletedSheet.relationshipId),
    );
    this.writeEntryText(
      "[Content_Types].xml",
      removeContentTypeOverride(contentTypesXml, deletedSheet.path),
    );
    this.rewriteAppSheetNames(
      context.sheets.filter((sheet) => sheet.name !== sheetName).map((sheet) => sheet.name),
    );
    this.removeEntry(deletedSheet.path);

    const sheetRelsPath = `${dirnamePosix(deletedSheet.path)}/_rels/${basenamePosix(deletedSheet.path)}.rels`;
    if (this.entries.has(sheetRelsPath)) {
      this.removeEntry(sheetRelsPath);
    }
  }

  async save(filePath: string): Promise<void> {
    await this.adapter.writeArchive(filePath, this.getEntriesView());
  }

  toEntries(): ArchiveEntry[] {
    return this.entryOrder.map((path) => {
      const data = this.entries.get(path);
      if (!data) {
        throw new XlsxError(`Entry missing from map: ${path}`);
      }

      return { path, data: new Uint8Array(data) };
    });
  }

  private *getEntriesView(): Iterable<ArchiveEntry> {
    for (const path of this.entryOrder) {
      const data = this.entries.get(path);
      if (!data) {
        throw new XlsxError(`Entry missing from map: ${path}`);
      }

      yield { path, data };
    }
  }

  private getWorkbookContext(): WorkbookContext {
    if (this.workbookContext) {
      return this.workbookContext;
    }

    this.workbookContext = resolveWorkbookContext(this, (path) => this.readEntryText(path));

    return this.workbookContext;
  }

  readSharedStrings(): string[] {
    return [...this.getSharedStringsCache()];
  }

  getSharedString(index: number): string | null {
    return this.getSharedStringsCache()[index] ?? null;
  }

  private getSharedStringsCache(): string[] {
    if (this.sharedStringsCache) {
      return this.sharedStringsCache;
    }

    const sharedStringsPath = this.getWorkbookContext().sharedStringsPath;
    if (!sharedStringsPath || !this.entries.has(sharedStringsPath)) {
      this.sharedStringsCache = [];
      return this.sharedStringsCache;
    }

    this.sharedStringsCache = parseSharedStrings(this.readEntryText(sharedStringsPath));
    return this.sharedStringsCache;
  }

  private getStylesCache(): StylesCache | null {
    if (this.stylesCache !== undefined) {
      return this.stylesCache;
    }

    const stylesPath = this.getWorkbookContext().stylesPath;
    if (!stylesPath || !this.entries.has(stylesPath)) {
      this.stylesCache = null;
      return this.stylesCache;
    }

    this.stylesCache = parseStylesXml(stylesPath, this.readEntryText(stylesPath));
    return this.stylesCache;
  }

  rewriteDefinedNamesForSheetStructure(
    sheetName: string,
    targetColumnNumber: number,
    columnCount: number,
    targetRowNumber: number,
    rowCount: number,
    mode: "shift" | "delete",
  ): void {
    const context = this.getWorkbookContext();
    const localSheetIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    if (localSheetIndex === -1) {
      return;
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    const replacement = rewriteDefinedNamesInWorkbookXml(workbookXml, (tag) => {
      const localSheetIdText = getTagAttr(tag, "localSheetId");
      const includeUnqualifiedReferences =
        localSheetIdText !== undefined && Number(localSheetIdText) === localSheetIndex;
      const nameText = decodeXmlText(tag.innerXml ?? "");
      const nextNameText =
        mode === "shift"
          ? shiftFormulaReferences(
              nameText,
              sheetName,
              targetColumnNumber,
              columnCount,
              targetRowNumber,
              rowCount,
              includeUnqualifiedReferences,
            )
          : deleteFormulaReferences(
              nameText,
              sheetName,
              targetColumnNumber,
              columnCount,
              targetRowNumber,
              rowCount,
              includeUnqualifiedReferences,
            );

      return nextNameText === nameText ? tag.source : buildDefinedNameTagSource(tag.attributesSource, nextNameText);
    });

    if (replacement.changed) {
      this.writeEntryText(context.workbookPath, replacement.workbookXml);
    }
  }

  private rewriteSheetFormulaTexts(
    path: string,
    transformFormula: (formula: string) => string,
  ): void {
    const result = rewriteFormulaXml(this.readEntryText(path), transformFormula);

    if (result.changed) {
      this.writeEntryText(path, result.sheetXml);
    }
  }

  private rewriteSheetHyperlinkLocations(
    path: string,
    currentSheetName: string,
    nextSheetName: string,
  ): void {
    const result = rewriteHyperlinkLocationXml(this.readEntryText(path), currentSheetName, nextSheetName);

    if (result.changed) {
      this.writeEntryText(path, result.sheetXml);
    }
  }

  private rewriteAppSheetNames(sheetNames: string[]): void {
    const appPath = "docProps/app.xml";
    if (!this.entries.has(appPath)) {
      return;
    }

    const appXml = this.readEntryText(appPath);
    const nextAppXml = updateAppSheetNames(appXml, sheetNames);

    if (nextAppXml !== appXml) {
      this.writeEntryText(appPath, nextAppXml);
    }
  }

  readEntryText(path: string): string {
    const entry = this.entries.get(path);
    if (!entry) {
      throw new XlsxError(`Entry not found: ${path}`);
    }

    return XML_DECODER.decode(entry);
  }

  writeEntryText(path: string, text: string): void {
    if (!this.entries.has(path)) {
      this.entryOrder.push(path);
    }

    if (this.workbookContext?.sharedStringsPath === path) {
      this.sharedStringsCache = undefined;
    }

    if (this.workbookContext?.stylesPath === path) {
      this.stylesCache = undefined;
    }

    if (
      this.workbookContext &&
      (this.workbookContext.workbookPath === path || this.workbookContext.workbookRelsPath === path)
    ) {
      this.stylesCache = undefined;
      this.workbookContext = undefined;
    }

    this.entries.set(path, XML_ENCODER.encode(text));
  }

  removeEntry(path: string): void {
    if (!this.entries.delete(path)) {
      return;
    }

    const entryIndex = this.entryOrder.indexOf(path);
    if (entryIndex !== -1) {
      this.entryOrder.splice(entryIndex, 1);
    }

    if (
      this.workbookContext &&
      (this.workbookContext.sharedStringsPath === path ||
        this.workbookContext.stylesPath === path ||
        this.workbookContext.workbookPath === path ||
        this.workbookContext.workbookRelsPath === path)
    ) {
      this.sharedStringsCache = undefined;
      this.stylesCache = undefined;
      this.workbookContext = undefined;
    }
  }
}
