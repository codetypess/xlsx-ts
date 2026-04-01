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
  WorkbookCreateOptions,
  WorkbookCreateSheetOptions,
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
  findSheetByName,
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
  removeEntryOrderPath,
  resolveSharedStringsCache,
  resolveStylesCache,
  shouldResetSharedStringsCache,
  shouldResetStylesCache,
  shouldResetWorkbookContext,
} from "./workbook/workbook-cache.js";
import {
  parseStylesXml,
  type ParsedBorder,
  type ParsedCellStyle,
  type ParsedFill,
  type ParsedFont,
  type StylesCache,
} from "./workbook/workbook-styles-parse.js";
import { buildWorkbookTemplateEntries } from "./workbook/workbook-template.js";
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

/**
 * Workbook-level API for reading and mutating OOXML spreadsheet packages.
 *
 * Style, font, fill, border, and number format ids on this class are the raw
 * OOXML ids from `styles.xml`. They are not worksheet row/column indexes.
 */
export class Workbook {
  private readonly adapter: Zip;
  private readonly entryOrder: string[];
  private readonly entries: Map<string, Uint8Array>;
  private batchDepth = 0;
  private readonly batchedSheets = new Set<Sheet>();
  private workbookContext?: WorkbookContext;
  private sharedStringsCache?: string[];
  private stylesCache?: StylesCache | null;

  /**
   * Creates a workbook wrapper around archive entries.
   *
   * Use {@link Workbook.open} for files on disk or {@link Workbook.fromEntries}
   * when the package is already loaded in memory.
   */
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

  /**
   * Opens an `.xlsx` file from disk.
   */
  static async open(filePath: string): Promise<Workbook> {
    const adapter = new Zip();
    const entries = await adapter.readArchive(filePath);
    return new Workbook(entries, adapter, { cloneEntryData: false });
  }

  /**
   * Opens an `.xlsx` archive that is already loaded as bytes.
   */
  static fromUint8Array(data: Uint8Array): Workbook {
    const adapter = new Zip();
    return new Workbook(adapter.readArchiveData(data), adapter, { cloneEntryData: false });
  }

  /**
   * Opens an `.xlsx` archive from an ArrayBuffer.
   */
  static fromArrayBuffer(data: ArrayBuffer): Workbook {
    return Workbook.fromUint8Array(new Uint8Array(data));
  }

  /**
   * Creates a workbook from archive entries that are already in memory.
   */
  static fromEntries(entries: Iterable<ArchiveEntry>): Workbook {
    return new Workbook(entries);
  }

  /**
   * Creates a new workbook from a minimal built-in template.
   */
  static create(sheetName?: string): Workbook;
  static create(options: WorkbookCreateOptions): Workbook;
  static create(sheetNameOrOptions: string | WorkbookCreateOptions = "Sheet1"): Workbook {
    if (typeof sheetNameOrOptions === "string") {
      assertSheetName(sheetNameOrOptions);
      return new Workbook(buildWorkbookTemplateEntries({ sheetName: sheetNameOrOptions }));
    }

    const options = sheetNameOrOptions;
    const normalizedSheets = normalizeWorkbookCreateSheets(options.sheets);
    if (normalizedSheets.length === 0) {
      throw new XlsxError("Workbook.create requires at least one sheet");
    }

    if (normalizedSheets.every((sheet) => sheet.visibility !== "visible")) {
      throw new XlsxError("Workbook.create requires at least one visible sheet");
    }

    const creator = options.author ?? "fastxlsx";
    const workbook = new Workbook(
      buildWorkbookTemplateEntries({
        createdAt: options.createdAt,
        creator,
        lastModifiedBy: options.modifiedBy ?? creator,
        sheetName: normalizedSheets[0]!.name,
      }),
    );

    workbook.batch((currentWorkbook) => {
      for (let index = 1; index < normalizedSheets.length; index += 1) {
        currentWorkbook.addSheet(normalizedSheets[index]!.name);
      }

      for (const sheetConfig of normalizedSheets) {
        const sheet = currentWorkbook.getSheet(sheetConfig.name);
        if (sheetConfig.headers && sheetConfig.headers.length > 0) {
          sheet.setHeaders(sheetConfig.headers);
        }
        if (sheetConfig.records && sheetConfig.records.length > 0) {
          sheet.addRecords(sheetConfig.records);
        }
      }

      for (const sheetConfig of normalizedSheets) {
        if (sheetConfig.visibility !== "visible") {
          currentWorkbook.setSheetVisibility(sheetConfig.name, sheetConfig.visibility);
        }
      }

      const activeSheetName = options.activeSheet ?? normalizedSheets.find((sheet) => sheet.visibility === "visible")!.name;
      currentWorkbook.setActiveSheet(activeSheetName);
    });

    return workbook;
  }

  /**
   * Lists archive entry paths in workbook order.
   */
  listEntries(): string[] {
    return [...this.entryOrder];
  }

  /**
   * Returns worksheet objects in workbook order.
   */
  getSheets(): Sheet[] {
    return [...this.getWorkbookContext().sheets];
  }

  /**
   * Groups multiple workbook or sheet mutations into one logical batch.
   */
  batch<Result>(applyChanges: (workbook: Workbook) => Result): Result {
    this.batchDepth += 1;

    try {
      return applyChanges(this);
    } finally {
      this.batchDepth -= 1;

      if (this.batchDepth === 0) {
        for (const sheet of this.batchedSheets) {
          sheet.finalizeBatchWrite();
        }
        this.batchedSheets.clear();
      }
    }
  }

  /**
   * Returns worksheet names in workbook order.
   */
  getSheetNames(): string[] {
    return this.getWorkbookContext().sheets.map((sheet) => sheet.name);
  }

  /**
   * Returns a worksheet by name.
   */
  getSheet(sheetName: string): Sheet {
    return requireSheetByName(this.getWorkbookContext().sheets, sheetName);
  }

  /**
   * Returns whether a worksheet with the given name exists.
   */
  hasSheet(sheetName: string): boolean {
    return this.tryGetSheet(sheetName) !== null;
  }

  /**
   * Returns a worksheet by name, or null when it does not exist.
   */
  tryGetSheet(sheetName: string): Sheet | null {
    return findSheetByName(this.getWorkbookContext().sheets, sheetName);
  }

  /**
   * Returns the currently active visible sheet according to workbook metadata.
   */
  getActiveSheet(): Sheet {
    const context = this.getWorkbookContext();
    const activeSheetIndex = parseActiveSheetIndex(this.readEntryText(context.workbookPath), context.sheets.length);
    return context.sheets[activeSheetIndex] ?? context.sheets[0]!;
  }

  /**
   * Reads a cell style definition by raw OOXML `cellXf` id.
   */
  getStyle(styleId: number): CellStyleDefinition | null {
    assertStyleId(styleId);
    return cloneCellStyleDefinition(this.getStylesCache()?.cellXfs[styleId]?.definition ?? null);
  }

  /**
   * Reads a number format definition by raw OOXML `numFmtId`.
   */
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

  /**
   * Reads a font definition by raw OOXML font id.
   */
  getFont(fontId: number): CellFontDefinition | null {
    assertStyleId(fontId);
    return cloneCellFontDefinition(this.getStylesCache()?.fonts[fontId]?.definition ?? null);
  }

  /**
   * Reads a fill definition by raw OOXML fill id.
   */
  getFill(fillId: number): CellFillDefinition | null {
    assertStyleId(fillId);
    return cloneCellFillDefinition(this.getStylesCache()?.fills[fillId]?.definition ?? null);
  }

  /**
   * Reads a border definition by raw OOXML border id.
   */
  getBorder(borderId: number): CellBorderDefinition | null {
    assertStyleId(borderId);
    return cloneCellBorderDefinition(this.getStylesCache()?.borders[borderId]?.definition ?? null);
  }

  /**
   * Rewrites an existing custom number format in `styles.xml`.
   */
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

  /**
   * Clones a builtin or custom number format and returns the new `numFmtId`.
   */
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

  /**
   * Returns the existing `numFmtId` for a format code or creates one.
   */
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

  /**
   * Rewrites an existing font definition by raw OOXML font id.
   */
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

  /**
   * Clones a font definition and returns the new raw OOXML font id.
   */
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

  /**
   * Rewrites an existing fill definition by raw OOXML fill id.
   */
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

  /**
   * Clones a fill definition and returns the new raw OOXML fill id.
   */
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

  /**
   * Rewrites an existing border definition by raw OOXML border id.
   */
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

  /**
   * Clones a border definition and returns the new raw OOXML border id.
   */
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

  /**
   * Rewrites an existing `cellXf` style by raw OOXML style id.
   */
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

  /**
   * Clones a `cellXf` style and returns the new raw OOXML style id.
   */
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

  /**
   * Sets the active sheet tab.
   *
   * Hidden sheets cannot be activated.
   */
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

  /**
   * Reads the workbook visibility state for a sheet.
   */
  getSheetVisibility(sheetName: string): SheetVisibility {
    const context = this.getWorkbookContext();
    const sheet = requireSheetByName(context.sheets, sheetName);

    return parseSheetVisibility(this.readEntryText(context.workbookPath), sheet.relationshipId);
  }

  /**
   * Updates the workbook visibility state for a sheet.
   *
   * The workbook must always keep at least one visible sheet.
   */
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

  /**
   * Lists workbook defined names, including sheet-scoped names.
   */
  getDefinedNames(): DefinedName[] {
    const context = this.getWorkbookContext();
    return parseDefinedNames(this.readEntryText(context.workbookPath), context.sheets);
  }

  /**
   * Reads one defined name value by name and optional sheet scope.
   */
  getDefinedName(name: string, scope?: string): string | null {
    const definedName = this.getDefinedNames().find(
      (candidate) => candidate.name === name && candidate.scope === (scope ?? null),
    );
    return definedName?.value ?? null;
  }

  /**
   * Creates or replaces a defined name.
   *
   * When `options.scope` is set, the name is sheet-scoped.
   */
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

  /**
   * Deletes a defined name by name and optional sheet scope.
   */
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

  /**
   * Renames a sheet and rewrites workbook-level references.
   */
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

  /**
   * Reorders a sheet to a zero-based workbook position.
   */
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

  /**
   * Appends a new worksheet to the workbook.
   */
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

  /**
   * Deletes a worksheet and rewrites workbook-level references.
   *
   * The last remaining sheet cannot be deleted.
   */
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

  /**
   * Writes the current workbook contents to disk as an `.xlsx` file.
   */
  async save(filePath: string): Promise<void> {
    await this.adapter.writeArchive(filePath, this.getEntriesView());
  }

  /**
   * Exports the current workbook contents as a zipped `.xlsx` byte array.
   */
  toUint8Array(): Uint8Array {
    return this.adapter.writeArchiveData(this.getEntriesView());
  }

  /**
   * Exports the workbook as detached archive entries.
   */
  toEntries(): ArchiveEntry[] {
    this.flushBatchedSheets();
    return this.entryOrder.map((path) => {
      const data = this.entries.get(path);
      if (!data) {
        throw new XlsxError(`Entry missing from map: ${path}`);
      }

      return { path, data: new Uint8Array(data) };
    });
  }

  private *getEntriesView(): Iterable<ArchiveEntry> {
    this.flushBatchedSheets();
    for (const path of this.entryOrder) {
      const data = this.entries.get(path);
      if (!data) {
        throw new XlsxError(`Entry missing from map: ${path}`);
      }

      yield { path, data };
    }
  }

  private flushBatchedSheets(): void {
    for (const sheet of this.batchedSheets) {
      sheet.finalizeBatchWrite();
    }
  }

  private getWorkbookContext(): WorkbookContext {
    if (this.workbookContext) {
      return this.workbookContext;
    }

    this.workbookContext = resolveWorkbookContext(this, (path) => this.readEntryText(path));

    return this.workbookContext;
  }

  /**
   * Returns the shared strings table as plain strings.
   */
  readSharedStrings(): string[] {
    return [...this.getSharedStringsCache()];
  }

  /**
   * Reads one shared string by zero-based shared string table index.
   */
  getSharedString(index: number): string | null {
    return this.getSharedStringsCache()[index] ?? null;
  }

  private getSharedStringsCache(): string[] {
    this.sharedStringsCache = resolveSharedStringsCache(
      this.sharedStringsCache,
      this.getWorkbookContext(),
      (path) => this.entries.has(path),
      (path) => this.readEntryText(path),
    );
    return this.sharedStringsCache;
  }

  private getStylesCache(): StylesCache | null {
    this.stylesCache = resolveStylesCache(
      this.stylesCache,
      this.getWorkbookContext(),
      (path) => this.entries.has(path),
      (path) => this.readEntryText(path),
    );
    return this.stylesCache;
  }

  /**
   * Internal helper used by sheet structure edits to keep defined names in sync.
   */
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

    if (shouldResetSharedStringsCache(this.workbookContext, path)) {
      this.sharedStringsCache = undefined;
    }

    if (shouldResetStylesCache(this.workbookContext, path)) {
      this.stylesCache = undefined;
    }

    if (shouldResetWorkbookContext(this.workbookContext, path)) {
      this.workbookContext = undefined;
    }

    this.entries.set(path, XML_ENCODER.encode(text));
  }

  removeEntry(path: string): void {
    if (!this.entries.delete(path)) {
      return;
    }

    removeEntryOrderPath(this.entryOrder, path);

    if (shouldResetSharedStringsCache(this.workbookContext, path)) {
      this.sharedStringsCache = undefined;
    }
    if (shouldResetStylesCache(this.workbookContext, path)) {
      this.stylesCache = undefined;
    }
    if (shouldResetWorkbookContext(this.workbookContext, path) || this.workbookContext?.sharedStringsPath === path) {
      this.workbookContext = undefined;
    }
  }

  isBatching(): boolean {
    return this.batchDepth > 0;
  }

  markSheetDirty(sheet: Sheet): void {
    this.batchedSheets.add(sheet);
  }
}

function normalizeWorkbookCreateSheets(
  sheets: Array<string | WorkbookCreateSheetOptions> | undefined,
): Array<WorkbookCreateSheetOptions & { visibility: SheetVisibility }> {
  const normalizedInput = sheets && sheets.length > 0 ? sheets : ["Sheet1"];

  return normalizedInput.map((sheet, index) => {
    const normalizedSheet =
      typeof sheet === "string"
        ? { name: sheet }
        : {
            headers: sheet.headers ? [...sheet.headers] : undefined,
            name: sheet.name,
            records: sheet.records ? sheet.records.map((record) => ({ ...record })) : undefined,
            visibility: sheet.visibility,
          };

    assertSheetName(normalizedSheet.name);
    const visibility = normalizedSheet.visibility ?? "visible";
    assertSheetVisibility(visibility);

    if (index > 0 && normalizedInput.some((candidate, candidateIndex) => candidateIndex < index && resolveCreateSheetName(candidate) === normalizedSheet.name)) {
      throw new XlsxError(`Sheet already exists: ${normalizedSheet.name}`);
    }

    return {
      ...normalizedSheet,
      visibility,
    };
  });
}

function resolveCreateSheetName(sheet: string | WorkbookCreateSheetOptions): string {
  return typeof sheet === "string" ? sheet : sheet.name;
}
