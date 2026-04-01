import { Cell } from "./cell.js";
import type {
  CellStyleAlignment,
  CellStyleAlignmentPatch,
  CellBorderDefinition,
  CellBorderPatch,
  CellEntry,
  CellFillDefinition,
  CellFillPatch,
  CellFontDefinition,
  CellFontPatch,
  CellNumberFormatDefinition,
  CellSnapshot,
  CellStyleDefinition,
  CellStylePatch,
  CellValue,
  SheetExportRecordsOptions,
  SheetImportRecordsOptions,
  DataValidation,
  FreezePane,
  Hyperlink,
  SheetComment,
  SheetCommentWriteOptions,
  SetDataValidationOptions,
  SetFormulaOptions,
  SetHyperlinkOptions,
  SheetImportRecordsResult,
  SheetSelection,
  SheetPrintTitles,
  SheetUpsertRecordResult,
} from "./types.js";
import { XlsxError } from "./errors.js";
import {
  buildSheetIndex,
  getLocatedCell,
  parseCellSnapshot,
  type LocatedCell,
  type LocatedRow,
  type SheetIndex,
} from "./sheet/sheet-index.js";
import {
  compareCellAddresses,
  makeCellAddress,
  normalizeCellAddress,
  normalizeColumnNumber,
  normalizeRangeRef,
  normalizeSqref,
  parseRangeRef,
  splitCellAddress,
} from "./sheet/sheet-address.js";
import {
  createCellEntry,
  normalizeEmptyRowXml,
  resolveCellAddress,
  resolveCopyStyleArguments,
} from "./sheet/sheet-common.js";
import { formatUsedRangeBounds, updateDimensionRef } from "./sheet/sheet-dimension.js";
import { removeSheetHyperlink, setSheetHyperlink } from "./sheet/sheet-hyperlink-ops.js";
import { parseMergedRanges, updateMergedRanges } from "./sheet/sheet-merge.js";
import { buildHeaderMap, writeRecordValues } from "./sheet/sheet-records.js";
import {
  assertStyleId,
  resolveCloneStylePatch,
  resolveSetAlignmentPatch,
  resolveSetBorderPatch,
  resolveSetFillPatch,
  resolveSetFontPatch,
  resolveSetStyleId,
} from "./sheet/sheet-style-input.js";
import {
  buildEmptyRowXml,
  buildEmptyStyledRowXml,
  buildStyledRowXml,
  buildUpdatedRowXml,
  parseColumnHidden,
  parseColumnStyleId,
  parseColumnWidth,
  parseRowHeight,
  parseRowHidden,
  parseRowStyleId,
  transformColumnStyleDefinitions,
  updateColumnHiddenInSheetXml,
  updateColumnStyleIdInSheetXml,
  updateColumnWidthInSheetXml,
} from "./sheet/sheet-style-xml.js";
import {
  EMPTY_RELATIONSHIPS_XML,
  getSheetRelationshipsPath,
  listTableReferences,
  rewriteTableReferenceXml,
  type TableReference,
} from "./sheet/sheet-table-xml.js";
import {
  assertColumnNumber,
  assertFreezeSplit,
  assertInsertCount,
  assertRowNumber,
} from "./sheet/sheet-validation.js";
import {
  buildCommentsVmlDrawingXml,
  buildCommentsXml,
  COMMENTS_CONTENT_TYPE,
  COMMENTS_RELATIONSHIP_TYPE,
  ensureDefaultContentType,
  ensureLegacyDrawingInSheetXml,
  findSheetCommentParts,
  getNextCommentsPath,
  getNextVmlDrawingPath,
  parseCommentsXml,
  removeLegacyDrawingFromSheetXml,
  VML_CONTENT_TYPE,
  VML_DRAWING_RELATIONSHIP_TYPE,
} from "./sheet/sheet-comments.js";
import {
  buildDataValidationXml,
  parseHyperlinkRelationshipTargets,
  parseSheetAutoFilter,
  parseSheetDataValidations,
  parseSheetHyperlinks,
  removeAutoFilterFromSheetXml,
  removeDataValidationFromSheetXml,
  upsertAutoFilterInSheetXml,
  upsertDataValidationInSheetXml,
} from "./sheet/sheet-metadata.js";
import {
  addContentTypeOverride,
  appendRelationship,
  appendTablePart,
  assertTableName,
  buildTableXml,
  findSheetTableReferenceByName,
  getNextRelationshipIdFromXml,
  getNextTableId,
  getNextTableName,
  getNextTablePath,
  makeRelativeSheetRelationshipTarget,
  parseSheetTables,
  removeContentTypeOverride,
  removeRelationshipById,
  upsertRelationship,
  removeTablePartsFromSheetXml,
  TABLE_CONTENT_TYPE,
  TABLE_RELATIONSHIP_TYPE,
} from "./sheet/sheet-package.js";
import {
  deleteColumnTransform,
  deleteFormulaReferences,
  deleteRangeRef,
  deleteRangeRefColumns,
  deleteRangeRefRows,
  deleteRowTransform,
  deleteSheetFormulaReferences,
  renameSheetFormulaReferences,
  shiftRangeRef,
  shiftFormulaReferences,
  shiftRangeRefColumns,
  shiftRangeRefRows,
  transformRowXml,
  transformWorksheetStructureReferences,
} from "./sheet/sheet-structure.js";
import { formatSheetNameForReference } from "./workbook/workbook-sheet-package.js";
import {
  buildFormulaCellXml,
  buildStyledCellXml,
  buildValueCellXml,
  findRowInsertionIndex,
  insertCell,
  resolveSetCellValue,
  resolveSetFormulaArguments,
} from "./sheet/sheet-write.js";
import {
  buildXmlElement,
  replaceXmlTagSource,
  rewriteXmlTagsByName,
} from "./sheet/sheet-xml.js";
import {
  parseSheetFreezePane,
  parseSheetSelection,
  removeFreezePaneFromSheetXml,
  upsertFreezePaneInSheetXml,
  upsertSheetSelectionInSheetXml,
} from "./sheet/sheet-view-metadata.js";
import type { Workbook } from "./workbook.js";
import { resolvePosix } from "./utils/path.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "./utils/xml-read.js";
import {
  decodeXmlText,
  escapeRegex,
  escapeXmlText,
  parseAttributes,
} from "./utils/xml.js";

/**
 * Worksheet read/write API.
 *
 * Numeric row and column arguments are 1-based across this class:
 * `1` means the first row or first column, not index `0`.
 */
export class Sheet {
  /**
   * Current worksheet name as stored in workbook metadata.
   */
  name: string;

  /**
   * Worksheet part path inside the OOXML package.
   */
  readonly path: string;

  /**
   * Workbook relationship id that points to this worksheet part.
   */
  readonly relationshipId: string;

  private readonly cellHandles = new Map<string, Cell>();
  private hasPendingCellMutations = false;
  private readonly pendingCellMutations = new Map<string, PendingCellMutation>();
  private hasPendingBatchWrite = false;
  private revision = 0;
  private readonly workbook: Workbook;
  private sheetIndex?: SheetIndex;

  /**
   * Creates a worksheet handle bound to a parent workbook.
   *
   * Most callers obtain instances from {@link Workbook.getSheet},
   * {@link Workbook.getSheets}, or related workbook APIs.
   */
  constructor(
    workbook: Workbook,
    options: {
      name: string;
      path: string;
      relationshipId: string;
    },
  ) {
    this.workbook = workbook;
    this.name = options.name;
    this.path = options.path;
    this.relationshipId = options.relationshipId;
  }

  /**
   * Returns a cached cell handle for an address.
   *
   * When using numeric coordinates, both `rowNumber` and numeric `column`
   * are 1-based.
   */
  cell(address: string): Cell;
  cell(rowNumber: number, column: number | string): Cell;
  cell(addressOrRowNumber: string | number, column?: number | string): Cell {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, column);
    let cell = this.cellHandles.get(normalizedAddress);

    if (!cell) {
      cell = new Cell(this, normalizedAddress);
      this.cellHandles.set(normalizedAddress, cell);
    }

    return cell;
  }

  /**
   * Groups multiple worksheet mutations and flushes sheet metadata once at the end.
   */
  batch<Result>(applyChanges: (sheet: Sheet) => Result): Result {
    return this.workbook.batch(() => applyChanges(this));
  }

  /**
   * Reads the current cell value.
   *
   * Numeric row and column arguments are 1-based.
   */
  getCell(address: string): CellValue;
  getCell(rowNumber: number, column: number | string): CellValue;
  getCell(addressOrRowNumber: string | number, column?: number | string): CellValue {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).value;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).value;
  }

  /**
   * Reads a best-effort display string for the current cell value.
   *
   * This is intended for user-facing inspection, not full Excel-format emulation.
   */
  getDisplayValue(address: string): string | null;
  getDisplayValue(rowNumber: number, column: number | string): string | null;
  getDisplayValue(addressOrRowNumber: string | number, column?: number | string): string | null {
    const snapshot =
      typeof addressOrRowNumber === "number"
        ? this.readCellSnapshotByIndexes(addressOrRowNumber, column)
        : this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column));
    return formatCellDisplayValue(snapshot);
  }

  /**
   * Reads the raw style id assigned to a cell.
   *
   * Numeric row and column arguments are 1-based.
   */
  getStyleId(address: string): number | null;
  getStyleId(rowNumber: number, column: number | string): number | null;
  getStyleId(addressOrRowNumber: string | number, column?: number | string): number | null {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).styleId;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).styleId;
  }

  /**
   * Resolves the effective cell style definition.
   *
   * Numeric row and column arguments are 1-based.
   */
  getStyle(address: string): CellStyleDefinition | null;
  getStyle(rowNumber: number, column: number | string): CellStyleDefinition | null;
  getStyle(addressOrRowNumber: string | number, column?: number | string): CellStyleDefinition | null {
    const styleId =
      typeof addressOrRowNumber === "number"
        ? this.readCellSnapshotByIndexes(addressOrRowNumber, column).styleId
        : this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).styleId;
    return this.workbook.getStyle(styleId ?? 0);
  }

  /**
   * Resolves the alignment portion of a cell style.
   *
   * Numeric row and column arguments are 1-based.
   */
  getAlignment(address: string): CellStyleAlignment | null;
  getAlignment(rowNumber: number, column: number | string): CellStyleAlignment | null;
  getAlignment(addressOrRowNumber: string | number, column?: number | string): CellStyleAlignment | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style?.alignment ?? null;
  }

  /**
   * Resolves the font portion of a cell style.
   *
   * Numeric row and column arguments are 1-based.
   */
  getFont(address: string): CellFontDefinition | null;
  getFont(rowNumber: number, column: number | string): CellFontDefinition | null;
  getFont(addressOrRowNumber: string | number, column?: number | string): CellFontDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getFont(style.fontId) : null;
  }

  /**
   * Resolves the fill portion of a cell style.
   *
   * Numeric row and column arguments are 1-based.
   */
  getFill(address: string): CellFillDefinition | null;
  getFill(rowNumber: number, column: number | string): CellFillDefinition | null;
  getFill(addressOrRowNumber: string | number, column?: number | string): CellFillDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getFill(style.fillId) : null;
  }

  /**
   * Reads the solid background color for a cell, if one is set.
   *
   * Numeric row and column arguments are 1-based.
   */
  getBackgroundColor(address: string): string | null;
  getBackgroundColor(rowNumber: number, column: number | string): string | null;
  getBackgroundColor(addressOrRowNumber: string | number, column?: number | string): string | null {
    const fill =
      typeof addressOrRowNumber === "number"
        ? this.getFill(addressOrRowNumber, column!)
        : this.getFill(addressOrRowNumber);
    if (!fill || fill.patternType !== "solid") {
      return null;
    }

    return fill.fgColor?.rgb ?? null;
  }

  /**
   * Resolves the border portion of a cell style.
   *
   * Numeric row and column arguments are 1-based.
   */
  getBorder(address: string): CellBorderDefinition | null;
  getBorder(rowNumber: number, column: number | string): CellBorderDefinition | null;
  getBorder(addressOrRowNumber: string | number, column?: number | string): CellBorderDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getBorder(style.borderId) : null;
  }

  /**
   * Resolves the number format portion of a cell style.
   *
   * Numeric row and column arguments are 1-based.
   */
  getNumberFormat(address: string): CellNumberFormatDefinition | null;
  getNumberFormat(rowNumber: number, column: number | string): CellNumberFormatDefinition | null;
  getNumberFormat(addressOrRowNumber: string | number, column?: number | string): CellNumberFormatDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getNumberFormat(style.numFmtId) : null;
  }

  /**
   * Reads the style id assigned to a column.
   *
   * Numeric column indexes are 1-based.
   */
  getColumnStyleId(column: number | string): number | null {
    const columnNumber = normalizeColumnNumber(column);
    return parseColumnStyleId(this.getSheetIndex().xml, columnNumber);
  }

  /**
   * Reads the effective style definition for a whole column.
   *
   * Numeric column indexes are 1-based.
   */
  getColumnStyle(column: number | string): CellStyleDefinition | null {
    const styleId = this.getColumnStyleId(column);
    return styleId === null ? null : this.workbook.getStyle(styleId);
  }

  /**
   * Reads whether a column is hidden.
   *
   * Numeric column indexes are 1-based.
   */
  getColumnHidden(column: number | string): boolean {
    return parseColumnHidden(this.getSheetIndex().xml, normalizeColumnNumber(column));
  }

  /**
   * Reads the explicit column width, if one is set.
   *
   * Numeric column indexes are 1-based.
   */
  getColumnWidth(column: number | string): number | null {
    return parseColumnWidth(this.getSheetIndex().xml, normalizeColumnNumber(column));
  }

  /**
   * Copies the style id from one cell to another without changing the value.
   *
   * Numeric row and column arguments are 1-based.
   */
  copyStyle(sourceAddress: string, targetAddress: string): void;
  copyStyle(
    sourceRowNumber: number,
    sourceColumn: number | string,
    targetRowNumber: number,
    targetColumn: number | string,
  ): void;
  copyStyle(
    sourceAddressOrRowNumber: string | number,
    sourceColumnOrTargetAddress: number | string,
    targetRowNumber?: number,
    targetColumn?: number | string,
  ): void {
    const { sourceAddress, targetAddress } = resolveCopyStyleArguments(
      sourceAddressOrRowNumber,
      sourceColumnOrTargetAddress,
      targetRowNumber,
      targetColumn,
    );
    this.setStyleId(targetAddress, this.getStyleId(sourceAddress));
  }

  /**
   * Clones the current cell style with a patch and applies it.
   *
   * Numeric row and column arguments are 1-based.
   */
  setStyle(address: string, patch: CellStylePatch): number;
  setStyle(rowNumber: number, column: number | string, patch: CellStylePatch): number;
  setStyle(
    addressOrRowNumber: string | number,
    columnOrPatch: number | string | CellStylePatch,
    patch?: CellStylePatch,
  ): number {
    if (typeof addressOrRowNumber === "number") {
      return this.cloneStyle(addressOrRowNumber, columnOrPatch as number | string, patch);
    }

    return this.cloneStyle(addressOrRowNumber, (columnOrPatch as CellStylePatch | undefined) ?? {});
  }

  /**
   * Clones the current style and replaces only the alignment part.
   *
   * Numeric row and column arguments are 1-based.
   */
  setAlignment(address: string, patch: CellStyleAlignmentPatch | null): number;
  setAlignment(rowNumber: number, column: number | string, patch: CellStyleAlignmentPatch | null): number;
  setAlignment(
    addressOrRowNumber: string | number,
    columnOrPatch: number | string | CellStyleAlignmentPatch | null,
    patch?: CellStyleAlignmentPatch | null,
  ): number {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrPatch as number | string) : undefined,
    );
    const nextPatch = resolveSetAlignmentPatch(addressOrRowNumber, columnOrPatch, patch);
    const nextStyleId = this.workbook.cloneStyle(this.getStyleId(normalizedAddress) ?? 0, {
      alignment: nextPatch,
      applyAlignment: nextPatch === null ? null : true,
    });
    this.setStyleId(normalizedAddress, nextStyleId);
    return nextStyleId;
  }

  /**
   * Clones the current style and replaces only the font part.
   *
   * Numeric row and column arguments are 1-based.
   */
  setFont(address: string, patch: CellFontPatch): number;
  setFont(rowNumber: number, column: number | string, patch: CellFontPatch): number;
  setFont(
    addressOrRowNumber: string | number,
    columnOrPatch: number | string | CellFontPatch,
    patch?: CellFontPatch,
  ): number {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrPatch as number | string) : undefined,
    );
    const nextPatch = resolveSetFontPatch(addressOrRowNumber, columnOrPatch, patch);
    const currentStyleId = this.getStyleId(normalizedAddress) ?? 0;
    const currentStyle = this.workbook.getStyle(currentStyleId);
    if (!currentStyle) {
      throw new XlsxError("Cell style not found");
    }

    const nextFontId = this.workbook.cloneFont(currentStyle.fontId, nextPatch);
    const nextStyleId = this.workbook.cloneStyle(currentStyleId, {
      fontId: nextFontId,
      applyFont: true,
    });
    this.setStyleId(normalizedAddress, nextStyleId);
    return nextFontId;
  }

  /**
   * Clones the current style and replaces only the fill part.
   *
   * Numeric row and column arguments are 1-based.
   */
  setFill(address: string, patch: CellFillPatch): number;
  setFill(rowNumber: number, column: number | string, patch: CellFillPatch): number;
  setFill(
    addressOrRowNumber: string | number,
    columnOrPatch: number | string | CellFillPatch,
    patch?: CellFillPatch,
  ): number {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrPatch as number | string) : undefined,
    );
    const nextPatch = resolveSetFillPatch(addressOrRowNumber, columnOrPatch, patch);
    const currentStyleId = this.getStyleId(normalizedAddress) ?? 0;
    const currentStyle = this.workbook.getStyle(currentStyleId);
    if (!currentStyle) {
      throw new XlsxError("Cell style not found");
    }

    const nextFillId = this.workbook.cloneFill(currentStyle.fillId, nextPatch);
    const nextStyleId = this.workbook.cloneStyle(currentStyleId, {
      fillId: nextFillId,
      applyFill: true,
    });
    this.setStyleId(normalizedAddress, nextStyleId);
    return nextFillId;
  }

  /**
   * Convenience helper for setting a solid background color.
   *
   * Numeric row and column arguments are 1-based.
   */
  setBackgroundColor(address: string, color: string | null): number;
  setBackgroundColor(rowNumber: number, column: number | string, color: string | null): number;
  setBackgroundColor(
    addressOrRowNumber: string | number,
    columnOrColor: number | string | null,
    color?: string | null,
  ): number {
    const nextColor = typeof addressOrRowNumber === "number" ? (color ?? null) : (columnOrColor as string | null);
    const fillPatch: CellFillPatch =
      nextColor === null
        ? {
            patternType: "none",
            fgColor: null,
            bgColor: null,
          }
        : {
            patternType: "solid",
            fgColor: { rgb: nextColor },
            bgColor: null,
          };

    if (typeof addressOrRowNumber === "number") {
      return this.setFill(addressOrRowNumber, columnOrColor as number | string, fillPatch);
    }

    return this.setFill(addressOrRowNumber, fillPatch);
  }

  /**
   * Clones the current style and replaces only the border part.
   *
   * Numeric row and column arguments are 1-based.
   */
  setBorder(address: string, patch: CellBorderPatch): number;
  setBorder(rowNumber: number, column: number | string, patch: CellBorderPatch): number;
  setBorder(
    addressOrRowNumber: string | number,
    columnOrPatch: number | string | CellBorderPatch,
    patch?: CellBorderPatch,
  ): number {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrPatch as number | string) : undefined,
    );
    const nextPatch = resolveSetBorderPatch(addressOrRowNumber, columnOrPatch, patch);
    const currentStyleId = this.getStyleId(normalizedAddress) ?? 0;
    const currentStyle = this.workbook.getStyle(currentStyleId);
    if (!currentStyle) {
      throw new XlsxError("Cell style not found");
    }

    const nextBorderId = this.workbook.cloneBorder(currentStyle.borderId, nextPatch);
    const nextStyleId = this.workbook.cloneStyle(currentStyleId, {
      borderId: nextBorderId,
      applyBorder: true,
    });
    this.setStyleId(normalizedAddress, nextStyleId);
    return nextBorderId;
  }

  /**
   * Clones the current style and replaces only the number format part.
   *
   * Numeric row and column arguments are 1-based.
   */
  setNumberFormat(address: string, formatCode: string): number;
  setNumberFormat(rowNumber: number, column: number | string, formatCode: string): number;
  setNumberFormat(
    addressOrRowNumber: string | number,
    columnOrFormatCode: number | string,
    formatCode?: string,
  ): number {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrFormatCode as number | string) : undefined,
    );
    const nextFormatCode =
      typeof addressOrRowNumber === "number" ? (formatCode ?? "") : (columnOrFormatCode as string);
    const currentStyleId = this.getStyleId(normalizedAddress) ?? 0;
    const nextNumFmtId = this.workbook.ensureNumberFormat(nextFormatCode);
    const nextStyleId = this.workbook.cloneStyle(currentStyleId, {
      numFmtId: nextNumFmtId,
      applyNumberFormat: true,
    });
    this.setStyleId(normalizedAddress, nextStyleId);
    return nextNumFmtId;
  }

  /**
   * Clones the current cell style and returns the new style id.
   *
   * Numeric row and column arguments are 1-based.
   */
  cloneStyle(address: string, patch?: CellStylePatch): number;
  cloneStyle(rowNumber: number, column: number | string, patch?: CellStylePatch): number;
  cloneStyle(
    addressOrRowNumber: string | number,
    columnOrPatch?: number | string | CellStylePatch,
    patch?: CellStylePatch,
  ): number {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrPatch as number | string) : undefined,
    );
    const nextPatch = resolveCloneStylePatch(addressOrRowNumber, columnOrPatch, patch);
    const nextStyleId = this.workbook.cloneStyle(this.getStyleId(normalizedAddress) ?? 0, nextPatch);
    this.setStyleId(normalizedAddress, nextStyleId);
    return nextStyleId;
  }

  /**
   * Renames this worksheet through the parent workbook.
   */
  rename(name: string): void {
    this.workbook.renameSheet(this.name, name);
  }

  /**
   * Reads the cell formula, if present.
   *
   * Numeric row and column arguments are 1-based.
   */
  getFormula(address: string): string | null;
  getFormula(rowNumber: number, column: number | string): string | null;
  getFormula(addressOrRowNumber: string | number, column?: number | string): string | null {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).formula;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).formula;
  }

  /**
   * Last used worksheet row number.
   */
  get rowCount(): number {
    return this.getSheetIndex().usedBounds?.maxRow ?? 0;
  }

  /**
   * Last used worksheet column number.
   */
  get columnCount(): number {
    return this.getSheetIndex().usedBounds?.maxColumn ?? 0;
  }

  /**
   * Reads a header row as strings.
   *
   * `headerRowNumber` is 1-based and defaults to the first row.
   */
  getHeaders(headerRowNumber = 1): string[] {
    assertRowNumber(headerRowNumber);
    return this.getRow(headerRowNumber).map((value) => (typeof value === "string" ? value : ""));
  }

  /**
   * Reads the row-level style id.
   *
   * `rowNumber` is 1-based.
   */
  getRowStyleId(rowNumber: number): number | null {
    assertRowNumber(rowNumber);
    return parseRowStyleId(this.getSheetIndex().rows.get(rowNumber)?.attributesSource);
  }

  /**
   * Reads the effective style definition for a row.
   *
   * `rowNumber` is 1-based.
   */
  getRowStyle(rowNumber: number): CellStyleDefinition | null {
    const styleId = this.getRowStyleId(rowNumber);
    return styleId === null ? null : this.workbook.getStyle(styleId);
  }

  /**
   * Reads whether a row is hidden.
   *
   * `rowNumber` is 1-based.
   */
  getRowHidden(rowNumber: number): boolean {
    assertRowNumber(rowNumber);
    return parseRowHidden(this.getSheetIndex().rows.get(rowNumber)?.attributesSource);
  }

  /**
   * Reads the explicit row height, if one is set.
   *
   * `rowNumber` is 1-based.
   */
  getRowHeight(rowNumber: number): number | null {
    assertRowNumber(rowNumber);
    return parseRowHeight(this.getSheetIndex().rows.get(rowNumber)?.attributesSource);
  }

  /**
   * Reads a worksheet row as a dense array up to the last used column.
   *
   * `rowNumber` is 1-based.
   */
  getRow(rowNumber: number): CellValue[] {
    assertRowNumber(rowNumber);

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row || row.cells.length === 0) {
      return [];
    }

    const values: CellValue[] = [];
    const maxColumn = row.maxColumnNumber;

    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      values.push(this.getCell(rowNumber, columnNumber));
    }

    return values;
  }

  /**
   * Reads logical cell entries present in a row.
   *
   * `rowNumber` is 1-based.
   */
  getRowEntries(rowNumber: number): CellEntry[] {
    return this.getPhysicalRowEntries(rowNumber).filter((cell) => isLogicalCellEntry(cell));
  }

  /**
   * Reads physical worksheet `<c>` nodes present in a row.
   *
   * `rowNumber` is 1-based.
   */
  getPhysicalRowEntries(rowNumber: number): CellEntry[] {
    assertRowNumber(rowNumber);

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row) {
      return [];
    }

    return row.cells.map((cell) => createCellEntry(cell));
  }

  /**
   * Reads a worksheet column as a dense array up to the last used row.
   *
   * Numeric column indexes are 1-based.
   */
  getColumn(column: number | string): CellValue[] {
    const columnNumber = normalizeColumnNumber(column);
    const index = this.getSheetIndex();
    const cells = index.rowNumbers
      .map((rowNumber) => index.rows.get(rowNumber)?.cellsByColumn[columnNumber])
      .filter((cell): cell is NonNullable<typeof cell> => cell !== undefined);

    if (cells.length === 0) {
      return [];
    }

    const values: CellValue[] = [];
    const maxRow = cells[cells.length - 1].rowNumber;

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
      values.push(this.getCell(rowNumber, columnNumber));
    }

    return values;
  }

  /**
   * Reads logical cell entries present in a column.
   *
   * Numeric column indexes are 1-based.
   */
  getColumnEntries(column: number | string): CellEntry[] {
    return this.getPhysicalColumnEntries(column).filter((cell) => isLogicalCellEntry(cell));
  }

  /**
   * Reads physical worksheet `<c>` nodes present in a column.
   *
   * Numeric column indexes are 1-based.
   */
  getPhysicalColumnEntries(column: number | string): CellEntry[] {
    const columnNumber = normalizeColumnNumber(column);
    const entries: CellEntry[] = [];
    const index = this.getSheetIndex();

    for (const rowNumber of index.rowNumbers) {
      const cell = index.rows.get(rowNumber)?.cellsByColumn[columnNumber];
      if (cell) {
        entries.push(createCellEntry(cell));
      }
    }

    return entries;
  }

  /**
   * Reads rows below the header as key-value records.
   *
   * Header names are taken from `headerRowNumber`, which is 1-based.
   * Blank or non-string headers are skipped.
   */
  getRecords(headerRowNumber = 1): Array<Record<string, CellValue>> {
    const headers = this.getRow(headerRowNumber);
    let lastHeaderColumn = 0;

    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      const value = headers[columnIndex];
      if (value !== null) {
        lastHeaderColumn = columnIndex + 1;
      }
    }

    if (lastHeaderColumn === 0) {
      return [];
    }

    const records: Array<Record<string, CellValue>> = [];
    const maxRow = this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber;

    for (let rowNumber = headerRowNumber + 1; rowNumber <= maxRow; rowNumber += 1) {
      const row = this.getRow(rowNumber);
      const hasAnyValue = row.some((value) => value !== null);
      if (!hasAnyValue) {
        continue;
      }

      const record: Record<string, CellValue> = {};

      for (let columnIndex = 0; columnIndex < lastHeaderColumn; columnIndex += 1) {
        const header = headers[columnIndex];
        if (typeof header !== "string" || header.length === 0) {
          continue;
        }

        record[header] = row[columnIndex] ?? null;
      }

      records.push(record);
    }

    return records;
  }

  /**
   * Reads one header-mapped record row.
   *
   * `rowNumber` and `headerRowNumber` are 1-based.
   */
  getRecord(rowNumber: number, headerRowNumber = 1): Record<string, CellValue> | null {
    assertRowNumber(rowNumber);

    const row = this.getRow(rowNumber);
    if (row.length === 0 || row.every((value) => value === null)) {
      return null;
    }

    const headers = this.getRow(headerRowNumber);
    const record: Record<string, CellValue> = {};

    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      const header = headers[columnIndex];
      if (typeof header !== "string" || header.length === 0) {
        continue;
      }

      record[header] = row[columnIndex] ?? null;
    }

    return record;
  }

  /**
   * Reads one record by matching a header field value.
   */
  getRecordBy(field: string, value: CellValue, headerRowNumber = 1): Record<string, CellValue> | null {
    const rowNumber = this.findRecordRow(field, value, headerRowNumber);
    return rowNumber === null ? null : this.getRecord(rowNumber, headerRowNumber);
  }

  /**
   * Alias for {@link getRecordBy}.
   */
  findRecordBy(field: string, value: CellValue, headerRowNumber = 1): Record<string, CellValue> | null {
    return this.getRecordBy(field, value, headerRowNumber);
  }

  /**
   * Exports header-mapped records as JSON-ready objects.
   */
  toJson(headerRowNumber = 1): Array<Record<string, CellValue>> {
    return this.getRecords(headerRowNumber).map((record) => ({ ...record }));
  }

  /**
   * Replaces the current record set from JSON-ready objects.
   */
  fromJson(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    const headers = collectRecordHeaders(records);
    if (headers.length > 0) {
      this.setHeaders(headers, headerRowNumber);
    }

    this.setRecords(records, headerRowNumber);
  }

  /**
   * Exports header-mapped records as CSV text.
   */
  toCsv(headerRowNumber = 1): string {
    const headers = trimTrailingEmptyHeaderNames(this.getHeaders(headerRowNumber));
    if (headers.length === 0) {
      return "";
    }

    const rows = [
      headers,
      ...this.getRecords(headerRowNumber).map((record) => headers.map((header) => formatCsvCellValue(record[header] ?? null))),
    ];

    return rows.map((row) => row.map((value) => escapeCsvField(value)).join(",")).join("\n");
  }

  /**
   * Replaces the current record set from CSV text using the first row as headers.
   */
  fromCsv(csv: string, headerRowNumber = 1): void {
    const rows = parseCsvRows(csv);
    if (rows.length === 0) {
      return;
    }

    const headers = rows[0]!.map((value) => value.trim());
    const records = rows.slice(1).map((row) => {
      const record: Record<string, CellValue> = {};

      for (let index = 0; index < headers.length; index += 1) {
        const header = headers[index];
        if (!header) {
          continue;
        }

        record[header] = parseCsvCellValue(row[index] ?? "");
      }

      return record;
    });

    this.setHeaders(headers, headerRowNumber);
    this.setRecords(records, headerRowNumber);
  }

  /**
   * Exports records in the requested format.
   */
  exportRecords(options: SheetExportRecordsOptions = {}): Array<Record<string, CellValue>> | string {
    const headerRow = options.headerRow ?? 1;
    const format = options.format ?? "json";
    return format === "csv" ? this.toCsv(headerRow) : this.toJson(headerRow);
  }

  /**
   * Imports records with a higher-level workflow mode.
   */
  importRecords(records: Array<Record<string, CellValue>>, options: SheetImportRecordsOptions = {}): SheetImportRecordsResult {
    const headerRow = options.headerRow ?? 1;
    const mode = options.mode ?? "replace";
    const headers = collectRecordHeaders(records);

    if (mode === "replace") {
      this.fromJson(records, headerRow);
      return {
        headers: headers.length > 0 ? headers : trimTrailingEmptyHeaderNames(this.getHeaders(headerRow)),
        imported: records.length,
        inserted: records.length,
        mode,
        rowCount: this.getRecords(headerRow).length,
        updated: 0,
      };
    }

    if (mode === "append") {
      this.addRecords(records, headerRow);
      return {
        headers: trimTrailingEmptyHeaderNames(this.getHeaders(headerRow)),
        imported: records.length,
        inserted: records.length,
        mode,
        rowCount: this.getRecords(headerRow).length,
        updated: 0,
      };
    }

    const keyField = options.keyField;
    if (!keyField) {
      throw new XlsxError("importRecords with mode=upsert requires keyField");
    }

    let inserted = 0;
    let updated = 0;
    this.batch((currentSheet) => {
      for (const record of records) {
        const result = currentSheet.upsertRecord(keyField, record, headerRow);
        if (result.inserted) {
          inserted += 1;
        } else {
          updated += 1;
        }
      }
    });

    return {
      headers: trimTrailingEmptyHeaderNames(this.getHeaders(headerRow)),
      imported: records.length,
      inserted,
      mode,
      rowCount: this.getRecords(headerRow).length,
      updated,
    };
  }

  /**
   * Synchronizes records with replace or upsert semantics.
   */
  syncRecords(records: Array<Record<string, CellValue>>, options: SheetImportRecordsOptions = {}): SheetImportRecordsResult {
    return this.importRecords(records, {
      ...options,
      mode: options.mode ?? (options.keyField ? "upsert" : "replace"),
    });
  }

  /**
   * Reads the local print area defined name for this sheet.
   */
  getPrintArea(): string | null {
    return this.workbook.getDefinedName("_xlnm.Print_Area", this.name);
  }

  /**
   * Creates or removes the local print area defined name for this sheet.
   */
  setPrintArea(range: string | null): string | null {
    if (range === null) {
      this.workbook.deleteDefinedName("_xlnm.Print_Area", this.name);
      return null;
    }

    const normalizedRange = normalizeRangeRef(range);
    this.workbook.setDefinedName("_xlnm.Print_Area", normalizedRange, { scope: this.name });
    return normalizedRange;
  }

  /**
   * Reads the row and column print-title references for this sheet.
   */
  getPrintTitles(): SheetPrintTitles {
    const value = this.workbook.getDefinedName("_xlnm.Print_Titles", this.name);
    if (!value) {
      return { columns: null, rows: null };
    }

    const titles: SheetPrintTitles = { columns: null, rows: null };
    for (const part of splitPrintTitleParts(value)) {
      const bangIndex = part.indexOf("!");
      const reference = bangIndex === -1 ? part : part.slice(bangIndex + 1);

      if (isPrintTitleRowRef(reference)) {
        titles.rows = normalizePrintTitleRowRef(reference);
      } else if (isPrintTitleColumnRef(reference)) {
        titles.columns = normalizePrintTitleColumnRef(reference);
      }
    }

    return titles;
  }

  /**
   * Creates, replaces, or removes the local print titles defined name for this sheet.
   */
  setPrintTitles(options: { columns?: string | null; rows?: string | null }): SheetPrintTitles {
    const rows = options.rows === undefined ? this.getPrintTitles().rows : options.rows;
    const columns = options.columns === undefined ? this.getPrintTitles().columns : options.columns;
    const references: string[] = [];
    const result: SheetPrintTitles = {
      columns: columns === null || columns === undefined ? null : normalizePrintTitleColumnRef(columns),
      rows: rows === null || rows === undefined ? null : normalizePrintTitleRowRef(rows),
    };

    if (rows !== null && rows !== undefined) {
      references.push(`${formatSheetNameForReference(this.name)}!${result.rows}`);
    }
    if (columns !== null && columns !== undefined) {
      references.push(`${formatSheetNameForReference(this.name)}!${result.columns}`);
    }

    if (references.length === 0) {
      this.workbook.deleteDefinedName("_xlnm.Print_Titles", this.name);
      return { columns: null, rows: null };
    }

    this.workbook.setDefinedName("_xlnm.Print_Titles", references.join(","), { scope: this.name });
    return result;
  }

  /**
   * Reads a rectangular cell range using A1 notation.
   */
  getRange(range: string): CellValue[][] {
    const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
    const values: CellValue[][] = [];

    for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
      const rowValues: CellValue[] = [];

      for (let columnNumber = startColumn; columnNumber <= endColumn; columnNumber += 1) {
        rowValues.push(this.getCell(makeCellAddress(rowNumber, columnNumber)));
      }

      values.push(rowValues);
    }

    return values;
  }

  /**
   * Applies a style patch to every cell in a range.
   */
  setRangeStyle(range: string, patch: CellStylePatch): void {
    this.batch((currentSheet) => {
      forEachCellInRange(range, (address) => {
        currentSheet.setStyle(address, patch);
      });
    });
  }

  /**
   * Applies a number format to every cell in a range.
   */
  setRangeNumberFormat(range: string, formatCode: string): void {
    this.batch((currentSheet) => {
      forEachCellInRange(range, (address) => {
        currentSheet.setNumberFormat(address, formatCode);
      });
    });
  }

  /**
   * Applies a solid background color to every cell in a range.
   */
  setRangeBackgroundColor(range: string, color: string | null): void {
    this.batch((currentSheet) => {
      forEachCellInRange(range, (address) => {
        currentSheet.setBackgroundColor(address, color);
      });
    });
  }

  /**
   * Copies styles from one rectangular range to another range of the same size.
   */
  copyRangeStyle(sourceRange: string, targetRange: string): void {
    const source = parseRangeRef(sourceRange);
    const target = parseRangeRef(targetRange);
    const sourceHeight = source.endRow - source.startRow;
    const sourceWidth = source.endColumn - source.startColumn;
    const targetHeight = target.endRow - target.startRow;
    const targetWidth = target.endColumn - target.startColumn;

    if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
      throw new XlsxError("Source and target ranges must have the same shape");
    }

    this.batch((currentSheet) => {
      for (let rowOffset = 0; rowOffset <= sourceHeight; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset <= sourceWidth; columnOffset += 1) {
          currentSheet.copyStyle(
            source.startRow + rowOffset,
            source.startColumn + columnOffset,
            target.startRow + rowOffset,
            target.startColumn + columnOffset,
          );
        }
      }
    });
  }

  /**
   * Returns all logical cell entries in row-major order.
   */
  getCellEntries(): CellEntry[] {
    return Array.from(this.iterCellEntries());
  }

  /**
   * Returns all physical worksheet `<c>` nodes in row-major order.
   */
  getPhysicalCellEntries(): CellEntry[] {
    return Array.from(this.iterPhysicalCellEntries());
  }

  /**
   * Iterates logical cells in row-major order.
   */
  *iterCellEntries(): IterableIterator<CellEntry> {
    const index = this.getSheetIndex();

    for (const rowNumber of index.rowNumbers) {
      const row = index.rows.get(rowNumber);
      if (!row) {
        continue;
      }

      for (const cell of row.cells) {
        if (!isLogicalCellEntry(cell.snapshot)) {
          continue;
        }

        yield createCellEntry(cell);
      }
    }
  }

  /**
   * Iterates physical worksheet `<c>` nodes in row-major order.
   */
  *iterPhysicalCellEntries(): IterableIterator<CellEntry> {
    const index = this.getSheetIndex();

    for (const rowNumber of index.rowNumbers) {
      const row = index.rows.get(rowNumber);
      if (!row) {
        continue;
      }

      for (const cell of row.cells) {
        yield createCellEntry(cell);
      }
    }
  }

  /**
   * Returns the current logical range in A1 notation.
   */
  getRangeRef(): string | null {
    return formatUsedRangeBounds(this.getSheetIndex().usedBounds);
  }

  /**
   * Returns the current physical worksheet `<c>` bounds in A1 notation.
   */
  getPhysicalRangeRef(): string | null {
    return formatUsedRangeBounds(this.getPhysicalBounds());
  }

  /**
   * Returns the current used range in A1 notation.
   */
  getMergedRanges(): string[] {
    return parseMergedRanges(this.getSheetIndex().xml);
  }

  /**
   * Reads the worksheet auto-filter range, if present.
   */
  getAutoFilter(): string | null {
    return parseSheetAutoFilter(this.getSheetIndex().xml);
  }

  /**
   * Reads the current freeze pane state.
   */
  getFreezePane(): FreezePane | null {
    return parseSheetFreezePane(this.getSheetIndex().xml);
  }

  private getPhysicalBounds(): {
    minRow: number;
    maxRow: number;
    minColumn: number;
    maxColumn: number;
  } | null {
    const index = this.getSheetIndex();
    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = 0;
    let minColumn = Number.POSITIVE_INFINITY;
    let maxColumn = 0;
    let hasCells = false;

    for (const rowNumber of index.rowNumbers) {
      const row = index.rows.get(rowNumber);
      if (!row || row.cells.length === 0) {
        continue;
      }

      hasCells = true;
      minRow = Math.min(minRow, rowNumber);
      maxRow = Math.max(maxRow, rowNumber);
      minColumn = Math.min(minColumn, row.cells[0]?.columnNumber ?? Number.POSITIVE_INFINITY);
      maxColumn = Math.max(maxColumn, row.cells[row.cells.length - 1]?.columnNumber ?? 0);
    }

    return hasCells ? { minRow, maxRow, minColumn, maxColumn } : null;
  }

  /**
   * Reads the current worksheet selection.
   */
  getSelection(): SheetSelection | null {
    return parseSheetSelection(this.getSheetIndex().xml);
  }

  /**
   * Lists worksheet data validation rules.
   */
  getDataValidations(): DataValidation[] {
    return parseSheetDataValidations(this.getSheetIndex().xml);
  }

  /**
   * Lists worksheet tables with their ranges and backing part paths.
   */
  getTables(): Array<{ name: string; displayName: string; range: string; path: string }> {
    return parseSheetTables(this.getTableReferences(), (path) => this.workbook.readEntryText(path));
  }

  /**
   * Lists worksheet comments.
   */
  getComments(): SheetComment[] {
    const parts = findSheetCommentParts(this.getSheetIndex().xml, this.path, this.readSheetRelationshipsXml());
    if (!parts.commentsPath || !this.workbook.listEntries().includes(parts.commentsPath)) {
      return [];
    }

    return parseCommentsXml(this.workbook.readEntryText(parts.commentsPath)).comments;
  }

  /**
   * Reads one worksheet comment by cell address.
   */
  getComment(address: string): SheetComment | null {
    const normalizedAddress = normalizeCellAddress(address);
    return this.getComments().find((comment) => comment.address === normalizedAddress) ?? null;
  }

  /**
   * Lists worksheet hyperlinks.
   */
  getHyperlinks(): Hyperlink[] {
    return parseSheetHyperlinks(this.getSheetIndex().xml, parseHyperlinkRelationshipTargets(this.readSheetRelationshipsXml()));
  }

  /**
   * Creates or replaces a hyperlink at the given cell address.
   */
  setHyperlink(address: string, target: string, options: SetHyperlinkOptions = {}): void {
    const normalizedAddress = normalizeCellAddress(address);
    if (options.text !== undefined) {
      this.setCell(normalizedAddress, options.text);
    }

    const nextState = setSheetHyperlink(
      this.getSheetIndex().xml,
      this.readSheetRelationshipsXml(),
      normalizedAddress,
      target,
      options,
    );
    this.writeSheetXml(nextState.sheetXml);
    this.writeSheetRelationshipsXml(nextState.relationshipsXml);
  }

  /**
   * Removes a hyperlink from the given cell address.
   */
  removeHyperlink(address: string): void {
    const normalizedAddress = normalizeCellAddress(address);
    const currentSheetXml = this.getSheetIndex().xml;
    const currentRelationshipsXml = this.readSheetRelationshipsXml();
    const nextState = removeSheetHyperlink(currentSheetXml, currentRelationshipsXml, normalizedAddress);

    if (nextState.sheetXml !== currentSheetXml) {
      this.writeSheetXml(nextState.sheetXml);
    }

    if (nextState.relationshipsXml !== currentRelationshipsXml) {
      this.writeSheetRelationshipsXml(nextState.relationshipsXml);
    }
  }

  /**
   * Creates or replaces a worksheet comment.
   */
  setComment(address: string, text: string, options: SheetCommentWriteOptions = {}): SheetComment {
    const normalizedAddress = normalizeCellAddress(address);
    const currentSheetXml = this.getSheetIndex().xml;
    let nextSheetXml = currentSheetXml;
    let nextRelationshipsXml = this.readSheetRelationshipsXml();
    let nextContentTypesXml = this.readContentTypesXml();
    const entryPaths = this.workbook.listEntries();
    const parts = findSheetCommentParts(currentSheetXml, this.path, nextRelationshipsXml);
    const existingComments =
      parts.commentsPath && entryPaths.includes(parts.commentsPath)
        ? parseCommentsXml(this.workbook.readEntryText(parts.commentsPath)).comments
        : [];
    const previousComment = existingComments.find((comment) => comment.address === normalizedAddress) ?? null;
    const nextComments = [
      ...existingComments.filter((comment) => comment.address !== normalizedAddress),
      {
        address: normalizedAddress,
        author: options.author ?? previousComment?.author ?? existingComments[0]?.author ?? "fastxlsx",
        text,
      },
    ].sort((left, right) => compareCellAddresses(left.address, right.address));

    let commentsPath = parts.commentsPath;
    if (!commentsPath) {
      commentsPath = getNextCommentsPath(entryPaths);
      const relationshipId = getNextRelationshipIdFromXml(nextRelationshipsXml);
      nextRelationshipsXml = appendRelationship(
        nextRelationshipsXml,
        relationshipId,
        COMMENTS_RELATIONSHIP_TYPE,
        makeRelativeSheetRelationshipTarget(this.path, commentsPath),
      );
      nextContentTypesXml = addContentTypeOverride(nextContentTypesXml, commentsPath, COMMENTS_CONTENT_TYPE);
    }

    let vmlRelationshipId = parts.legacyDrawingRelationshipId;
    let vmlPath = parts.vmlPath;
    if (!vmlRelationshipId || !vmlPath) {
      vmlRelationshipId = parts.legacyDrawingRelationshipId ?? getNextRelationshipIdFromXml(nextRelationshipsXml);
      vmlPath = parts.vmlPath ?? getNextVmlDrawingPath(entryPaths);
      nextRelationshipsXml = upsertRelationship(
        nextRelationshipsXml,
        vmlRelationshipId,
        VML_DRAWING_RELATIONSHIP_TYPE,
        makeRelativeSheetRelationshipTarget(this.path, vmlPath),
      );
      nextSheetXml = ensureLegacyDrawingInSheetXml(nextSheetXml, vmlRelationshipId);
      nextContentTypesXml = ensureDefaultContentType(nextContentTypesXml, "vml", VML_CONTENT_TYPE);
    }

    this.workbook.writeEntryText(commentsPath, buildCommentsXml(nextComments));
    this.workbook.writeEntryText(vmlPath, buildCommentsVmlDrawingXml(nextComments));
    if (nextSheetXml !== currentSheetXml) {
      this.writeSheetXml(nextSheetXml);
    }
    if (nextRelationshipsXml !== this.readSheetRelationshipsXml()) {
      this.writeSheetRelationshipsXml(nextRelationshipsXml);
    }
    if (nextContentTypesXml !== this.readContentTypesXml()) {
      this.writeContentTypesXml(nextContentTypesXml);
    }

    return {
      address: normalizedAddress,
      author: options.author ?? previousComment?.author ?? existingComments[0]?.author ?? "fastxlsx",
      text,
    };
  }

  /**
   * Removes a worksheet comment by cell address.
   */
  removeComment(address: string): void {
    const normalizedAddress = normalizeCellAddress(address);
    const currentSheetXml = this.getSheetIndex().xml;
    const currentRelationshipsXml = this.readSheetRelationshipsXml();
    const parts = findSheetCommentParts(currentSheetXml, this.path, currentRelationshipsXml);
    if (!parts.commentsPath || !this.workbook.listEntries().includes(parts.commentsPath)) {
      return;
    }

    const existingComments = parseCommentsXml(this.workbook.readEntryText(parts.commentsPath)).comments;
    const nextComments = existingComments.filter((comment) => comment.address !== normalizedAddress);
    if (nextComments.length === existingComments.length) {
      return;
    }

    if (nextComments.length > 0) {
      this.workbook.writeEntryText(parts.commentsPath, buildCommentsXml(nextComments));
      if (parts.vmlPath) {
        this.workbook.writeEntryText(parts.vmlPath, buildCommentsVmlDrawingXml(nextComments));
      }
      return;
    }

    let nextRelationshipsXml = currentRelationshipsXml;
    let nextSheetXml = currentSheetXml;
    let nextContentTypesXml = this.readContentTypesXml();

    this.workbook.removeEntry(parts.commentsPath);
    nextContentTypesXml = removeContentTypeOverride(nextContentTypesXml, parts.commentsPath);
    if (parts.commentsRelationshipId) {
      nextRelationshipsXml = removeRelationshipById(nextRelationshipsXml, parts.commentsRelationshipId);
    }

    if (parts.vmlPath) {
      this.workbook.removeEntry(parts.vmlPath);
    }
    if (parts.vmlRelationshipId) {
      nextRelationshipsXml = removeRelationshipById(nextRelationshipsXml, parts.vmlRelationshipId);
    }
    nextSheetXml = removeLegacyDrawingFromSheetXml(nextSheetXml);

    if (nextSheetXml !== currentSheetXml) {
      this.writeSheetXml(nextSheetXml);
    }
    if (nextRelationshipsXml !== currentRelationshipsXml) {
      this.writeSheetRelationshipsXml(nextRelationshipsXml);
    }
    if (nextContentTypesXml !== this.readContentTypesXml()) {
      this.writeContentTypesXml(nextContentTypesXml);
    }
  }

  /**
   * Sets the worksheet auto-filter range.
   */
  setAutoFilter(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const nextSheetXml = upsertAutoFilterInSheetXml(this.getSheetIndex().xml, normalizedRange);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Removes the worksheet auto-filter.
   */
  removeAutoFilter(): void {
    const nextSheetXml = removeAutoFilterFromSheetXml(this.getSheetIndex().xml);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Freezes the top rows and left columns.
   *
   * Numeric arguments are counts, not zero-based indexes.
   */
  freezePane(columnCount: number, rowCount = 0): void {
    assertFreezeSplit(columnCount, rowCount);
    const nextSheetXml = upsertFreezePaneInSheetXml(this.getSheetIndex().xml, columnCount, rowCount);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Clears any existing freeze pane state.
   */
  unfreezePane(): void {
    const nextSheetXml = removeFreezePaneFromSheetXml(this.getSheetIndex().xml);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Sets the active cell and selected range.
   */
  setSelection(activeCell: string, range = activeCell): void {
    const normalizedActiveCell = normalizeCellAddress(activeCell);
    const normalizedRange = normalizeSqref(range);
    const nextSheetXml = upsertSheetSelectionInSheetXml(
      this.getSheetIndex().xml,
      normalizedActiveCell,
      normalizedRange,
    );

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Creates or replaces a data validation rule for a range.
   */
  setDataValidation(range: string, options: SetDataValidationOptions = {}): void {
    const normalizedRange = normalizeSqref(range);
    const nextSheetXml = upsertDataValidationInSheetXml(
      this.getSheetIndex().xml,
      buildDataValidationXml(normalizedRange, options),
      normalizedRange,
    );

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Removes data validation rules matching a range.
   */
  removeDataValidation(range: string): void {
    const normalizedRange = normalizeSqref(range);
    const nextSheetXml = removeDataValidationFromSheetXml(this.getSheetIndex().xml, normalizedRange);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Adds a table part for a worksheet range.
   */
  addTable(
    range: string,
  options: { name?: string } = {},
  ): { name: string; displayName: string; range: string; path: string } {
    const normalizedRange = normalizeRangeRef(range);
    const existingTables = this.getTables();
    const name = options.name ?? getNextTableName(this.workbook.listEntries());
    assertTableName(name);

    if (existingTables.some((table) => table.name === name || table.displayName === name)) {
      throw new XlsxError(`Table already exists: ${name}`);
    }

    const tablePath = getNextTablePath(this.workbook.listEntries());
    const tableId = getNextTableId(this.workbook.listEntries(), (path) => this.workbook.readEntryText(path));
    const relationshipId = getNextRelationshipIdFromXml(this.readSheetRelationshipsXml());
    const tableXml = buildTableXml(normalizedRange, tableId, name, this.getRange(normalizedRange)[0] ?? []);

    this.workbook.writeEntryText(tablePath, tableXml);
    this.writeSheetRelationshipsXml(
      appendRelationship(
        this.readSheetRelationshipsXml(),
        relationshipId,
        TABLE_RELATIONSHIP_TYPE,
        makeRelativeSheetRelationshipTarget(this.path, tablePath),
      ),
    );
    this.writeSheetXml(appendTablePart(this.getSheetIndex().xml, relationshipId));
    this.writeContentTypesXml(addContentTypeOverride(this.readContentTypesXml(), tablePath, TABLE_CONTENT_TYPE));

    return {
      name,
      displayName: name,
      range: normalizedRange,
      path: tablePath,
    };
  }

  /**
   * Removes a worksheet table by table or display name.
   */
  removeTable(name: string): void {
    const tableReference = findSheetTableReferenceByName(
      this.getTableReferences(),
      (path) => this.workbook.readEntryText(path),
      name,
    );

    if (!tableReference) {
      throw new XlsxError(`Table not found: ${name}`);
    }

    this.writeSheetXml(removeTablePartsFromSheetXml(this.getSheetIndex().xml, [tableReference.relationshipId]));
    this.writeSheetRelationshipsXml(removeRelationshipById(this.readSheetRelationshipsXml(), tableReference.relationshipId));
    this.writeContentTypesXml(removeContentTypeOverride(this.readContentTypesXml(), tableReference.path));
    this.workbook.removeEntry(tableReference.path);
  }

  /**
   * Inserts worksheet rows before `rowNumber`.
   *
   * `rowNumber` is 1-based.
   */
  insertRow(rowNumber: number, count = 1): void {
    assertRowNumber(rowNumber);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges().map((range) =>
      shiftRangeRefRows(range, rowNumber, count),
    );

    for (const sourceRowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(sourceRowNumber);
      if (!row) {
        continue;
      }

      const nextRowXml = transformRowXml(
        index.xml,
        row,
        this.name,
        0,
        0,
        rowNumber,
        count,
      );
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformWorksheetStructureReferences(nextSheetXml, 0, 0, rowNumber, count, "shift");
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      shiftFormulaReferences(formula, this.name, 0, 0, rowNumber, count, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, 0, 0, rowNumber, count, "shift");
    this.updateTableReferences(0, 0, rowNumber, count, "shift");
  }

  /**
   * Inserts worksheet columns before `column`.
   *
   * Numeric column indexes are 1-based.
   */
  insertColumn(column: number | string, count = 1): void {
    const columnNumber = normalizeColumnNumber(column);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges().map((range) =>
      shiftRangeRefColumns(range, columnNumber, count),
    );

    for (const rowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(rowNumber);
      if (!row || row.selfClosing || row.cells.length === 0) {
        continue;
      }

      const nextRowXml = transformRowXml(
        index.xml,
        row,
        this.name,
        columnNumber,
        count,
        0,
        0,
      );
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformColumnStyleDefinitions(nextSheetXml, columnNumber, count, "shift");
    nextSheetXml = transformWorksheetStructureReferences(
      nextSheetXml,
      columnNumber,
      count,
      0,
      0,
      "shift",
    );
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      shiftFormulaReferences(formula, this.name, columnNumber, count, 0, 0, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, columnNumber, count, 0, 0, "shift");
    this.updateTableReferences(columnNumber, count, 0, 0, "shift");
  }

  /**
   * Deletes worksheet rows starting at `rowNumber`.
   *
   * `rowNumber` is 1-based.
   */
  deleteRow(rowNumber: number, count = 1): void {
    assertRowNumber(rowNumber);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    const deleteEndRow = rowNumber + count - 1;
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges()
      .map((range) => deleteRangeRefRows(range, rowNumber, count))
      .filter((range): range is string => range !== null);

    for (const sourceRowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(sourceRowNumber);
      if (!row) {
        continue;
      }

      if (sourceRowNumber >= rowNumber && sourceRowNumber <= deleteEndRow) {
        nextSheetXml = nextSheetXml.slice(0, row.start) + nextSheetXml.slice(row.end);
        continue;
      }

      const nextRowXml = deleteRowTransform(index.xml, row, this.name, rowNumber, count);
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformWorksheetStructureReferences(nextSheetXml, 0, 0, rowNumber, count, "delete");
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      deleteFormulaReferences(formula, this.name, 0, 0, rowNumber, count, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, 0, 0, rowNumber, count, "delete");
    this.updateTableReferences(0, 0, rowNumber, count, "delete");
  }

  /**
   * Deletes worksheet columns starting at `column`.
   *
   * Numeric column indexes are 1-based.
   */
  deleteColumn(column: number | string, count = 1): void {
    const columnNumber = normalizeColumnNumber(column);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges()
      .map((range) => deleteRangeRefColumns(range, columnNumber, count))
      .filter((range): range is string => range !== null);

    for (const rowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(rowNumber);
      if (!row) {
        continue;
      }

      const nextRowXml = deleteColumnTransform(index.xml, row, this.name, columnNumber, count);
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformColumnStyleDefinitions(nextSheetXml, columnNumber, count, "delete");
    nextSheetXml = transformWorksheetStructureReferences(
      nextSheetXml,
      columnNumber,
      count,
      0,
      0,
      "delete",
    );
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      deleteFormulaReferences(formula, this.name, columnNumber, count, 0, 0, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, columnNumber, count, 0, 0, "delete");
    this.updateTableReferences(columnNumber, count, 0, 0, "delete");
  }

  /**
   * Writes a cell value.
   *
   * Numeric row and column arguments are 1-based.
   */
  setCell(address: string, value: CellValue): void;
  setCell(rowNumber: number, column: number | string, value: CellValue): void;
  setCell(addressOrRowNumber: string | number, columnOrValue: number | string | CellValue, value?: CellValue): void {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? columnOrValue as number | string : undefined,
    );
    const currentCell = this.getCurrentCellWriteState(normalizedAddress);
    const nextValue = resolveSetCellValue(addressOrRowNumber, columnOrValue, value);
    const nextCellXml = buildValueCellXml(normalizedAddress, nextValue, currentCell.attributesSource);

    if (this.workbook.isBatching()) {
      this.stagePendingCellMutation(normalizedAddress, {
        attributesSource: extractCellAttributesSource(nextCellXml),
        kind: "set",
        snapshot: buildValueCellSnapshot(nextValue, currentCell.snapshot.styleId),
        xml: nextCellXml,
      });
      return;
    }

    this.writeCellXml(normalizedAddress, nextCellXml);
  }

  /**
   * Assigns a raw style id to a cell.
   *
   * Numeric row and column arguments are 1-based.
   */
  setStyleId(address: string, styleId: number | null): void;
  setStyleId(rowNumber: number, column: number | string, styleId: number | null): void;
  setStyleId(
    addressOrRowNumber: string | number,
    columnOrStyleId: number | string | null,
    styleId?: number | null,
  ): void {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrStyleId as number | string) : undefined,
    );
    const nextStyleId = resolveSetStyleId(addressOrRowNumber, columnOrStyleId, styleId);
    const currentCell = this.getCurrentCellWriteState(normalizedAddress);
    const nextCellXml = buildStyledCellXml(
      normalizedAddress,
      nextStyleId,
      currentCell.attributesSource,
      currentCell.cellXml,
    );

    if (this.workbook.isBatching()) {
      this.stagePendingCellMutation(normalizedAddress, {
        attributesSource: extractCellAttributesSource(nextCellXml),
        kind: "set",
        snapshot: buildStyledCellSnapshot(currentCell.snapshot, nextStyleId),
        xml: nextCellXml,
      });
      return;
    }

    this.writeCellXml(normalizedAddress, nextCellXml);
  }

  /**
   * Assigns a raw style id to a whole column.
   *
   * Numeric column indexes are 1-based.
   */
  setColumnStyleId(column: number | string, styleId: number | null): void {
    const columnNumber = normalizeColumnNumber(column);
    assertStyleId(styleId);

    const nextSheetXml = updateColumnStyleIdInSheetXml(
      this.getSheetIndex().xml,
      columnNumber,
      styleId,
    );

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Sets whether a column is hidden.
   *
   * Numeric column indexes are 1-based.
   */
  setColumnHidden(column: number | string, hidden: boolean): void {
    const columnNumber = normalizeColumnNumber(column);
    const currentSheetXml = this.getSheetIndex().xml;
    const nextSheetXml = updateColumnHiddenInSheetXml(currentSheetXml, columnNumber, hidden);

    if (nextSheetXml !== currentSheetXml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Sets or clears an explicit column width.
   *
   * Numeric column indexes are 1-based.
   */
  setColumnWidth(column: number | string, width: number | null): void {
    const columnNumber = normalizeColumnNumber(column);
    assertOptionalWorksheetSize(width, "column width");

    const currentSheetXml = this.getSheetIndex().xml;
    const nextSheetXml = updateColumnWidthInSheetXml(currentSheetXml, columnNumber, width);

    if (nextSheetXml !== currentSheetXml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  /**
   * Clones and applies a column style.
   */
  setColumnStyle(column: number | string, patch: CellStylePatch): number {
    return this.cloneColumnStyle(column, patch);
  }

  /**
   * Clones the effective column style and returns the new style id.
   */
  cloneColumnStyle(column: number | string, patch: CellStylePatch = {}): number {
    const nextStyleId = this.workbook.cloneStyle(this.getColumnStyleId(column) ?? 0, patch);
    this.setColumnStyleId(column, nextStyleId);
    return nextStyleId;
  }

  /**
   * Removes a cell node from the worksheet.
   *
   * Numeric row and column arguments are 1-based.
   */
  deleteCell(address: string): void;
  deleteCell(rowNumber: number, column: number | string): void;
  deleteCell(addressOrRowNumber: string | number, column?: number | string): void {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, column);
    const currentCell = this.getCurrentCellWriteState(normalizedAddress);

    if (!currentCell.snapshot.exists) {
      return;
    }

    if (this.workbook.isBatching()) {
      this.stagePendingCellMutation(normalizedAddress, {
        kind: "delete",
        snapshot: createMissingCellSnapshot(),
      });
      return;
    }

    const index = this.getSheetIndex();
    const existingCell = getLocatedCell(index, normalizedAddress);

    if (!existingCell) {
      return;
    }

    const row = index.rows.get(existingCell.rowNumber);
    if (row && row.cells.length === 1) {
      const nextRowXml = normalizeEmptyRowXml(
        index.xml.slice(row.start, existingCell.start) + index.xml.slice(existingCell.end, row.end),
      );
      this.writeSheetXml(index.xml.slice(0, row.start) + nextRowXml + index.xml.slice(row.end));
      return;
    }

    this.writeSheetXml(index.xml.slice(0, existingCell.start) + index.xml.slice(existingCell.end));
  }

  /**
   * Writes a formula cell and optional cached value.
   *
   * Numeric row and column arguments are 1-based.
   */
  setFormula(address: string, formula: string, options?: SetFormulaOptions): void;
  setFormula(rowNumber: number, column: number | string, formula: string, options?: SetFormulaOptions): void;
  setFormula(
    addressOrRowNumber: string | number,
    columnOrFormula: number | string,
    formulaOrOptions?: string | SetFormulaOptions,
    options: SetFormulaOptions = {},
  ): void {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? columnOrFormula as number | string : undefined,
    );
    const currentCell = this.getCurrentCellWriteState(normalizedAddress);
    const { formula, formulaOptions } = resolveSetFormulaArguments(
      addressOrRowNumber,
      columnOrFormula,
      formulaOrOptions,
      options,
    );
    const cachedValue = formulaOptions.cachedValue ?? null;
    const nextCellXml = buildFormulaCellXml(
      normalizedAddress,
      formula,
      cachedValue,
      currentCell.attributesSource,
    );

    if (this.workbook.isBatching()) {
      this.stagePendingCellMutation(normalizedAddress, {
        attributesSource: extractCellAttributesSource(nextCellXml),
        kind: "set",
        snapshot: buildFormulaCellSnapshot(formula, cachedValue, currentCell.snapshot.styleId),
        xml: nextCellXml,
      });
      return;
    }

    this.writeCellXml(normalizedAddress, nextCellXml);
  }

  /**
   * Returns the current internal revision counter for cache invalidation.
   */
  getRevision(): number {
    return this.revision;
  }

  /**
   * Writes a header row.
   *
   * `headerRowNumber` and `startColumn` are 1-based.
   */
  setHeaders(headers: string[], headerRowNumber = 1, startColumn = 1): void {
    assertRowNumber(headerRowNumber);
    assertColumnNumber(startColumn);
    this.setRow(headerRowNumber, headers, startColumn);
  }

  /**
   * Assigns a raw style id to a whole row.
   *
   * `rowNumber` is 1-based.
   */
  setRowStyleId(rowNumber: number, styleId: number | null): void {
    assertRowNumber(rowNumber);
    assertStyleId(styleId);

    const index = this.getSheetIndex();
    const row = index.rows.get(rowNumber);

    if (!row) {
      if (styleId === null) {
        return;
      }

      const insertionIndex = findRowInsertionIndex(index, rowNumber);
      this.writeSheetXml(
        index.xml.slice(0, insertionIndex) +
          buildEmptyStyledRowXml(rowNumber, styleId) +
          index.xml.slice(insertionIndex),
      );
      return;
    }

    this.writeSheetXml(
      index.xml.slice(0, row.start) +
        buildStyledRowXml(index.xml, row, styleId) +
        index.xml.slice(row.end),
    );
  }

  /**
   * Sets whether a row is hidden.
   *
   * `rowNumber` is 1-based.
   */
  setRowHidden(rowNumber: number, hidden: boolean): void {
    assertRowNumber(rowNumber);

    const index = this.getSheetIndex();
    const row = index.rows.get(rowNumber);
    if (!row) {
      if (!hidden) {
        return;
      }

      const nextRowXml = buildEmptyRowXml(rowNumber, { hidden: true });
      if (!nextRowXml) {
        return;
      }

      const insertionIndex = findRowInsertionIndex(index, rowNumber);
      this.writeSheetXml(index.xml.slice(0, insertionIndex) + nextRowXml + index.xml.slice(insertionIndex));
      return;
    }

    this.writeSheetXml(index.xml.slice(0, row.start) + buildUpdatedRowXml(index.xml, row, { hidden }) + index.xml.slice(row.end));
  }

  /**
   * Sets or clears an explicit row height.
   *
   * `rowNumber` is 1-based.
   */
  setRowHeight(rowNumber: number, height: number | null): void {
    assertRowNumber(rowNumber);
    assertOptionalWorksheetSize(height, "row height");

    const index = this.getSheetIndex();
    const row = index.rows.get(rowNumber);
    if (!row) {
      if (height === null) {
        return;
      }

      const nextRowXml = buildEmptyRowXml(rowNumber, { height });
      if (!nextRowXml) {
        return;
      }

      const insertionIndex = findRowInsertionIndex(index, rowNumber);
      this.writeSheetXml(index.xml.slice(0, insertionIndex) + nextRowXml + index.xml.slice(insertionIndex));
      return;
    }

    this.writeSheetXml(index.xml.slice(0, row.start) + buildUpdatedRowXml(index.xml, row, { height }) + index.xml.slice(row.end));
  }

  /**
   * Clones and applies a row style.
   */
  setRowStyle(rowNumber: number, patch: CellStylePatch): number {
    return this.cloneRowStyle(rowNumber, patch);
  }

  /**
   * Clones the effective row style and returns the new style id.
   */
  cloneRowStyle(rowNumber: number, patch: CellStylePatch = {}): number {
    assertRowNumber(rowNumber);
    const nextStyleId = this.workbook.cloneStyle(this.getRowStyleId(rowNumber) ?? 0, patch);
    this.setRowStyleId(rowNumber, nextStyleId);
    return nextStyleId;
  }

  /**
   * Adds a merged range if it is not already present.
   */
  addMergedRange(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const ranges = this.getMergedRanges();
    if (ranges.includes(normalizedRange)) {
      return;
    }

    this.writeSheetXml(updateMergedRanges(this.getSheetIndex().xml, [...ranges, normalizedRange]));
  }

  /**
   * Removes a merged range by normalized A1 reference.
   */
  removeMergedRange(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const ranges = this.getMergedRanges().filter((candidate) => candidate !== normalizedRange);
    this.writeSheetXml(updateMergedRanges(this.getSheetIndex().xml, ranges));
  }

  /**
   * Writes consecutive values into one row.
   *
   * `rowNumber` and `startColumn` are 1-based.
   */
  setRow(rowNumber: number, values: CellValue[], startColumn = 1): void {
    assertRowNumber(rowNumber);
    assertColumnNumber(startColumn);

    for (let columnOffset = 0; columnOffset < values.length; columnOffset += 1) {
      this.setCell(makeCellAddress(rowNumber, startColumn + columnOffset), values[columnOffset]);
    }
  }

  /**
   * Appends one row after the current last used row.
   *
   * `startColumn` is 1-based. Returns the appended 1-based row number.
   */
  appendRow(values: CellValue[], startColumn = 1): number {
    assertColumnNumber(startColumn);
    const rowNumber = (this.getSheetIndex().rowNumbers.at(-1) ?? 0) + 1;
    this.setRow(rowNumber, values, startColumn);
    return rowNumber;
  }

  /**
   * Appends multiple rows and returns their 1-based row numbers.
   *
   * `startColumn` is 1-based.
   */
  appendRows(rows: CellValue[][], startColumn = 1): number[] {
    assertColumnNumber(startColumn);

    const rowNumbers: number[] = [];
    let nextRowNumber = (this.getSheetIndex().rowNumbers.at(-1) ?? 0) + 1;

    for (const row of rows) {
      this.setRow(nextRowNumber, row, startColumn);
      rowNumbers.push(nextRowNumber);
      nextRowNumber += 1;
    }

    return rowNumbers;
  }

  /**
   * Writes consecutive values into one column.
   *
   * Numeric column indexes and `startRow` are 1-based.
   */
  setColumn(column: number | string, values: CellValue[], startRow = 1): void {
    const columnNumber = normalizeColumnNumber(column);
    assertRowNumber(startRow);

    for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
      this.setCell(makeCellAddress(startRow + rowOffset, columnNumber), values[rowOffset]);
    }
  }

  /**
   * Appends one header-mapped record after the current used range.
   */
  addRecord(record: Record<string, CellValue>, headerRowNumber = 1): void {
    if (Object.keys(record).length === 0) {
      return;
    }

    const headerMap = this.resolveRecordHeaderMap(headerRowNumber, [record], true);
    const nextRowNumber = Math.max(headerRowNumber + 1, (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1);
    this.writeRecordRow(nextRowNumber, record, headerMap, false);
  }

  /**
   * Alias for {@link addRecord}.
   */
  appendRecord(record: Record<string, CellValue>, headerRowNumber = 1): void {
    this.addRecord(record, headerRowNumber);
  }

  /**
   * Appends multiple header-mapped records after the current used range.
   */
  addRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    if (records.length === 0) {
      return;
    }

    const headerMap = this.resolveRecordHeaderMap(headerRowNumber, records, true);
    let nextRowNumber = Math.max(headerRowNumber + 1, (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1);

    for (const record of records) {
      if (Object.keys(record).length === 0) {
        nextRowNumber += 1;
        continue;
      }

      this.writeRecordRow(nextRowNumber, record, headerMap, false);
      nextRowNumber += 1;
    }
  }

  /**
   * Alias for {@link addRecords}.
   */
  appendRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    this.addRecords(records, headerRowNumber);
  }

  /**
   * Writes one header-mapped record into an existing row.
   *
   * `rowNumber` and `headerRowNumber` are 1-based.
   */
  setRecord(rowNumber: number, record: Record<string, CellValue>, headerRowNumber = 1): void {
    assertRowNumber(rowNumber);

    if (Object.keys(record).length === 0) {
      return;
    }

    const headerMap = this.resolveRecordHeaderMap(headerRowNumber, [record], rowNumber > headerRowNumber);
    this.writeRecordRow(rowNumber, record, headerMap, false);
  }

  /**
   * Replaces the record set below a header row.
   *
   * `headerRowNumber` is 1-based.
   */
  setRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    const headerMap = this.resolveRecordHeaderMap(headerRowNumber, records, true);
    const existingRecordRows = this.getSheetIndex().rowNumbers.filter(
      (rowNumber) => rowNumber > headerRowNumber && this.getRecord(rowNumber, headerRowNumber) !== null,
    );
    const targetRows: number[] = [];

    for (let index = 0; index < records.length; index += 1) {
      const rowNumber = headerRowNumber + 1 + index;
      this.writeRecordRow(rowNumber, records[index], headerMap, true);
      targetRows.push(rowNumber);
    }

    const rowsToDelete = existingRecordRows.filter((rowNumber) => !targetRows.includes(rowNumber));
    this.deleteRecords(rowsToDelete, headerRowNumber);
  }

  /**
   * Alias for {@link setRecords}.
   */
  replaceRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    this.setRecords(records, headerRowNumber);
  }

  /**
   * Creates or updates a record matched by one header field.
   *
   * Returns the 1-based row number that was written.
   */
  upsertRecord(field: string, record: Record<string, CellValue>, headerRowNumber = 1): SheetUpsertRecordResult {
    if (!Object.hasOwn(record, field)) {
      throw new XlsxError(`Record is missing match field: ${field}`);
    }

    const rowNumber = this.findRecordRow(field, record[field] ?? null, headerRowNumber);
    if (rowNumber === null) {
      const nextRowNumber = Math.max(
        headerRowNumber + 1,
        (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1,
      );
      this.addRecord(record, headerRowNumber);
      return {
        inserted: true,
        record: { ...record },
        row: nextRowNumber,
      };
    }

    this.setRecord(rowNumber, record, headerRowNumber);
    return {
      inserted: false,
      record: { ...record },
      row: rowNumber,
    };
  }

  /**
   * Deletes a single record row.
   *
   * `rowNumber` and `headerRowNumber` are 1-based.
   */
  deleteRecord(rowNumber: number, headerRowNumber = 1): void {
    assertRowNumber(rowNumber);
    assertRowNumber(headerRowNumber);

    if (rowNumber <= headerRowNumber) {
      throw new XlsxError(`Cannot delete header row: ${rowNumber}`);
    }

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row) {
      return;
    }

    const nextSheetXml = this.getSheetIndex().xml.slice(0, row.start) + this.getSheetIndex().xml.slice(row.end);
    this.writeSheetXml(nextSheetXml);
  }

  /**
   * Deletes multiple record rows.
   *
   * `headerRowNumber` is 1-based.
   */
  deleteRecords(rowNumbers: number[], headerRowNumber = 1): void {
    assertRowNumber(headerRowNumber);

    const uniqueRows = [...new Set(rowNumbers)];
    uniqueRows.sort((left, right) => right - left);

    for (const rowNumber of uniqueRows) {
      this.deleteRecord(rowNumber, headerRowNumber);
    }
  }

  /**
   * Deletes the first record matched by one header field.
   */
  deleteRecordBy(field: string, value: CellValue, headerRowNumber = 1): boolean {
    const rowNumber = this.findRecordRow(field, value, headerRowNumber);
    if (rowNumber === null) {
      return false;
    }

    this.deleteRecord(rowNumber, headerRowNumber);
    return true;
  }

  /**
   * Alias for {@link deleteRecordBy}.
   */
  removeRecordBy(field: string, value: CellValue, headerRowNumber = 1): boolean {
    return this.deleteRecordBy(field, value, headerRowNumber);
  }

  /**
   * Reads the full parsed cell snapshot for an address.
   */
  readCellSnapshot(address: string): CellSnapshot {
    const normalizedAddress = normalizeCellAddress(address);
    if (this.hasPendingCellMutations) {
      const pendingCell = this.pendingCellMutations.get(normalizedAddress);
      if (pendingCell) {
        return pendingCell.snapshot;
      }
    }

    const locatedCell = getLocatedCell(this.getSheetIndex(false), normalizedAddress);
    return parseCellSnapshot(locatedCell);
  }

  private readCellSnapshotByIndexes(
    rowNumber: number,
    column: number | string | undefined,
  ): CellSnapshot {
    assertRowNumber(rowNumber);
    if (column === undefined) {
      throw new XlsxError(`Missing column index for row: ${rowNumber}`);
    }

    const columnNumber = normalizeColumnNumber(column);
    if (this.hasPendingCellMutations) {
      const pendingCell = this.pendingCellMutations.get(makeCellAddress(rowNumber, columnNumber));
      if (pendingCell) {
        return pendingCell.snapshot;
      }
    }

    const row = this.getSheetIndex(false).rows.get(rowNumber);
    return parseCellSnapshot(row?.cellsByColumn[columnNumber]);
  }

  private getHeaderMap(headerRowNumber: number): Map<string, number> {
    assertRowNumber(headerRowNumber);
    return buildHeaderMap(this.getRow(headerRowNumber));
  }

  private resolveRecordHeaderMap(
    headerRowNumber: number,
    records: Array<Record<string, CellValue>>,
    allowHeaderInitialization: boolean,
  ): Map<string, number> {
    const headerMap = this.getHeaderMap(headerRowNumber);
    if (headerMap.size > 0 || !allowHeaderInitialization || !isEmptyHeaderRow(this.getRow(headerRowNumber))) {
      return headerMap;
    }

    const inferredHeaders = collectRecordHeaders(records);
    if (inferredHeaders.length === 0) {
      return headerMap;
    }

    this.setHeaders(inferredHeaders, headerRowNumber);
    return this.getHeaderMap(headerRowNumber);
  }

  private findRecordRow(field: string, value: CellValue, headerRowNumber: number): number | null {
    assertRowNumber(headerRowNumber);

    const maxRow = this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber;
    for (let rowNumber = headerRowNumber + 1; rowNumber <= maxRow; rowNumber += 1) {
      const record = this.getRecord(rowNumber, headerRowNumber);
      if (record && Object.hasOwn(record, field) && record[field] === value) {
        return rowNumber;
      }
    }

    return null;
  }

  private writeRecordRow(
    rowNumber: number,
    record: Record<string, CellValue>,
    headerMap: Map<string, number>,
    replaceMissingKeys: boolean,
  ): void {
    writeRecordValues(rowNumber, record, headerMap, replaceMissingKeys, (address, value) =>
      this.setCell(address, value),
    );
  }

  /**
   * Writes a rectangular value matrix starting at an address.
   */
  setRange(startAddress: string, values: CellValue[][]): void {
    const normalizedStartAddress = normalizeCellAddress(startAddress);
    if (values.length === 0) {
      return;
    }

    const expectedWidth = values[0]?.length ?? 0;
    if (expectedWidth === 0) {
      throw new XlsxError("Range values must contain at least one column");
    }

    for (const row of values) {
      if (row.length !== expectedWidth) {
        throw new XlsxError("Range values must be rectangular");
      }
    }

    const { rowNumber: startRow, columnNumber: startColumn } = splitCellAddress(normalizedStartAddress);

    for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
      const row = values[rowOffset];

      for (let columnOffset = 0; columnOffset < row.length; columnOffset += 1) {
        this.setCell(makeCellAddress(startRow + rowOffset, startColumn + columnOffset), row[columnOffset]);
      }
    }
  }

  private getSheetIndex(flushPendingMutations = true): SheetIndex {
    if (flushPendingMutations && this.hasPendingCellMutations) {
      this.flushPendingCellMutations();
    }

    if (this.sheetIndex) {
      return this.sheetIndex;
    }

    this.sheetIndex = buildSheetIndex(this.workbook, this.workbook.readEntryText(this.path));
    return this.sheetIndex;
  }

  private getCurrentCellWriteState(address: string): CurrentCellWriteState {
    if (this.hasPendingCellMutations) {
      const pendingCell = this.pendingCellMutations.get(address);
      if (pendingCell) {
        return pendingCell.kind === "delete"
          ? { snapshot: pendingCell.snapshot }
          : {
              attributesSource: pendingCell.attributesSource,
              cellXml: pendingCell.xml,
              snapshot: pendingCell.snapshot,
            };
      }
    }

    const index = this.getSheetIndex(false);
    const locatedCell = getLocatedCell(index, address);
    if (!locatedCell) {
      return { snapshot: createMissingCellSnapshot() };
    }

    return {
      attributesSource: locatedCell.attributesSource,
      cellXml: index.xml.slice(locatedCell.start, locatedCell.end),
      snapshot: locatedCell.snapshot,
    };
  }

  private stagePendingCellMutation(
    address: string,
    mutation: Omit<PendingCellMutation, "address" | "columnNumber" | "rowNumber">,
  ): void {
    const { rowNumber, columnNumber } = splitCellAddress(address);
    this.pendingCellMutations.set(address, {
      address,
      columnNumber,
      rowNumber,
      ...mutation,
    });
    this.hasPendingCellMutations = true;
    this.hasPendingBatchWrite = true;
    this.workbook.markSheetDirty(this);
    this.revision += 1;
  }

  private flushPendingCellMutations(): void {
    if (!this.hasPendingCellMutations) {
      return;
    }

    const baseIndex = this.getSheetIndex(false);
    const nextSheetXml = applyPendingCellMutationsToSheetXml(baseIndex, this.pendingCellMutations.values());
    this.workbook.writeEntryText(this.path, nextSheetXml);
    this.sheetIndex = buildSheetIndex(this.workbook, nextSheetXml);
    this.pendingCellMutations.clear();
    this.hasPendingCellMutations = false;
  }

  private writeCellXml(address: string, cellXml: string): void {
    const index = this.getSheetIndex();
    const existingCell = getLocatedCell(index, address);
    const nextSheetXml = existingCell
      ? index.xml.slice(0, existingCell.start) + cellXml + index.xml.slice(existingCell.end)
      : insertCell(index, address, cellXml);

    this.writeSheetXml(nextSheetXml);
  }

  private syncReferencedFormulasInOtherSheets(
    transformFormula: (formula: string) => string,
  ): void {
    for (const sheet of this.workbook.getSheets()) {
      if (sheet.path === this.path) {
        continue;
      }

      sheet.rewriteFormulaTexts(transformFormula);
    }
  }

  private rewriteFormulaTexts(transformFormula: (formula: string) => string): void {
    const sheetXml = this.getSheetIndex().xml;
    let changed = false;
    const nextSheetXml = rewriteXmlTagsByName(sheetXml, "f", (formulaTag) => {
      const formula = decodeXmlText(formulaTag.innerXml ?? "");
      const nextFormula = transformFormula(formula);

      if (nextFormula === formula) {
        return formulaTag.source;
      }

      changed = true;
      return buildXmlElement("f", parseAttributes(formulaTag.attributesSource), escapeXmlText(nextFormula));
    });

    if (changed) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  private getTableReferences(): TableReference[] {
    return listTableReferences(this.getSheetIndex().xml, this.path, this.workbook.listEntries(), (path) =>
      this.workbook.readEntryText(path),
    );
  }

  private readSheetRelationshipsXml(): string {
    const relationshipsPath = getSheetRelationshipsPath(this.path);
    return this.workbook.listEntries().includes(relationshipsPath)
      ? this.workbook.readEntryText(relationshipsPath)
      : EMPTY_RELATIONSHIPS_XML;
  }

  private writeSheetRelationshipsXml(relationshipsXml: string): void {
    this.workbook.writeEntryText(getSheetRelationshipsPath(this.path), relationshipsXml);
  }

  private readContentTypesXml(): string {
    return this.workbook.readEntryText("[Content_Types].xml");
  }

  private writeContentTypesXml(contentTypesXml: string): void {
    this.workbook.writeEntryText("[Content_Types].xml", contentTypesXml);
  }

  private updateTableReferences(
    targetColumnNumber: number,
    columnCount: number,
    targetRowNumber: number,
    rowCount: number,
    mode: "shift" | "delete",
  ): void {
    const transformRange = (range: string) =>
      mode === "shift"
        ? shiftRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount)
        : deleteRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount);
    const removedTables: TableReference[] = [];

    for (const table of this.getTableReferences()) {
      const tableXml = this.workbook.readEntryText(table.path);
      const nextTableXml = rewriteTableReferenceXml(tableXml, transformRange);

      if (nextTableXml === null) {
        removedTables.push(table);
        continue;
      }

      if (nextTableXml !== tableXml) {
        this.workbook.writeEntryText(table.path, nextTableXml);
      }
    }

    if (removedTables.length === 0) {
      return;
    }

    const removedRelationshipIds = removedTables.map((table) => table.relationshipId);
    const nextSheetXml = removeTablePartsFromSheetXml(this.getSheetIndex().xml, removedRelationshipIds);
    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }

    this.writeSheetRelationshipsXml(
      removedRelationshipIds.reduce(
        (relationshipsXml, relationshipId) => removeRelationshipById(relationshipsXml, relationshipId),
        this.readSheetRelationshipsXml(),
      ),
    );

    let nextContentTypesXml = this.readContentTypesXml();
    for (const table of removedTables) {
      nextContentTypesXml = removeContentTypeOverride(nextContentTypesXml, table.path);
      this.workbook.removeEntry(table.path);
    }
    this.writeContentTypesXml(nextContentTypesXml);
  }

  private writeSheetXml(nextSheetXml: string): void {
    const indexedSheet = buildSheetIndex(this.workbook, nextSheetXml);
    if (this.workbook.isBatching()) {
      this.workbook.writeEntryText(this.path, nextSheetXml);
      this.sheetIndex = indexedSheet;
      this.hasPendingBatchWrite = true;
      this.workbook.markSheetDirty(this);
      this.revision += 1;
      return;
    }

    const normalizedSheetXml = updateDimensionRef(indexedSheet);

    this.workbook.writeEntryText(this.path, normalizedSheetXml);
    this.sheetIndex =
      normalizedSheetXml === nextSheetXml ? indexedSheet : buildSheetIndex(this.workbook, normalizedSheetXml);
    this.revision += 1;
  }

  finalizeBatchWrite(): void {
    if (this.hasPendingCellMutations) {
      this.flushPendingCellMutations();
    }

    if (!this.hasPendingBatchWrite || !this.sheetIndex) {
      return;
    }

    const normalizedSheetXml = updateDimensionRef(this.sheetIndex);
    this.workbook.writeEntryText(this.path, normalizedSheetXml);
    this.sheetIndex =
      normalizedSheetXml === this.sheetIndex.xml
        ? this.sheetIndex
        : buildSheetIndex(this.workbook, normalizedSheetXml);
    this.hasPendingBatchWrite = false;
    this.revision += 1;
  }
}

interface PendingCellMutation {
  address: string;
  attributesSource?: string;
  columnNumber: number;
  kind: "delete" | "set";
  rowNumber: number;
  snapshot: CellSnapshot;
  xml?: string;
}

interface CurrentCellWriteState {
  attributesSource?: string;
  cellXml?: string;
  snapshot: CellSnapshot;
}

function isLogicalCellEntry(cell: Pick<CellEntry, "formula" | "value">): boolean {
  return cell.formula !== null || cell.value !== null;
}

function createMissingCellSnapshot(): CellSnapshot {
  return {
    exists: false,
    error: null,
    formula: null,
    rawType: null,
    styleId: null,
    type: "missing",
    value: null,
  };
}

function buildValueCellSnapshot(value: CellValue, styleId: number | null): CellSnapshot {
  return {
    exists: true,
    error: null,
    formula: null,
    rawType:
      typeof value === "string"
        ? "inlineStr"
        : typeof value === "boolean"
          ? "b"
          : null,
    styleId,
    type:
      value === null
        ? "blank"
        : typeof value === "string"
          ? "string"
          : typeof value === "number"
            ? "number"
            : "boolean",
    value,
  };
}

function buildStyledCellSnapshot(cell: CellSnapshot, styleId: number | null): CellSnapshot {
  if (!cell.exists) {
    return {
      exists: true,
      error: null,
      formula: null,
      rawType: null,
      styleId,
      type: "blank",
      value: null,
    };
  }

  return {
    ...cell,
    exists: true,
    styleId,
    type: cell.type === "missing" ? "blank" : cell.type,
  };
}

function collectRecordHeaders(records: Array<Record<string, CellValue>>): string[] {
  const headers: string[] = [];
  const seen = new Set<string>();

  for (const record of records) {
    for (const key of Object.keys(record)) {
      if (key.length === 0 || seen.has(key)) {
        continue;
      }

      seen.add(key);
      headers.push(key);
    }
  }

  return headers;
}

function isEmptyHeaderRow(values: CellValue[]): boolean {
  return values.length === 0 || values.every((value) => value === null || value === "");
}

function forEachCellInRange(range: string, visit: (address: string) => void): void {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    for (let columnNumber = startColumn; columnNumber <= endColumn; columnNumber += 1) {
      visit(makeCellAddress(rowNumber, columnNumber));
    }
  }
}

function assertOptionalWorksheetSize(value: number | null, label: string): void {
  if (value === null) {
    return;
  }

  if (!Number.isFinite(value) || value <= 0) {
    throw new XlsxError(`Invalid ${label}: ${value}`);
  }
}

function trimTrailingEmptyHeaderNames(headers: string[]): string[] {
  let end = headers.length;

  while (end > 0 && headers[end - 1] === "") {
    end -= 1;
  }

  return headers.slice(0, end);
}

function formatCsvCellValue(value: CellValue): string {
  if (value === null) {
    return "";
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  return String(value);
}

function parseCsvCellValue(value: string): CellValue {
  if (value.length === 0) {
    return null;
  }

  if (value === "TRUE") {
    return true;
  }
  if (value === "FALSE") {
    return false;
  }

  const numericValue = Number(value);
  if (value.trim() !== "" && Number.isFinite(numericValue)) {
    return numericValue;
  }

  return value;
}

function escapeCsvField(value: string): string {
  if (!/[",\n\r]/.test(value)) {
    return value;
  }

  return `"${value.replaceAll("\"", "\"\"")}"`;
}

function parseCsvRows(csv: string): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentValue = "";
  let index = 0;
  let inQuotes = false;

  while (index < csv.length) {
    const character = csv[index]!;

    if (inQuotes) {
      if (character === "\"") {
        if (csv[index + 1] === "\"") {
          currentValue += "\"";
          index += 2;
          continue;
        }

        inQuotes = false;
        index += 1;
        continue;
      }

      currentValue += character;
      index += 1;
      continue;
    }

    if (character === "\"") {
      inQuotes = true;
      index += 1;
      continue;
    }
    if (character === ",") {
      currentRow.push(currentValue);
      currentValue = "";
      index += 1;
      continue;
    }
    if (character === "\n") {
      currentRow.push(currentValue);
      rows.push(currentRow);
      currentRow = [];
      currentValue = "";
      index += 1;
      continue;
    }
    if (character === "\r") {
      index += 1;
      continue;
    }

    currentValue += character;
    index += 1;
  }

  if (inQuotes) {
    throw new XlsxError("Invalid CSV: unterminated quoted field");
  }

  if (currentValue.length > 0 || currentRow.length > 0) {
    currentRow.push(currentValue);
    rows.push(currentRow);
  }

  return rows;
}

function splitPrintTitleParts(value: string): string[] {
  const parts: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]!;
    if (character === "'") {
      if (inQuotes && value[index + 1] === "'") {
        current += "''";
        index += 1;
        continue;
      }

      inQuotes = !inQuotes;
      current += character;
      continue;
    }

    if (character === "," && !inQuotes) {
      parts.push(current);
      current = "";
      continue;
    }

    current += character;
  }

  if (current.length > 0) {
    parts.push(current);
  }

  return parts.map((part) => part.trim()).filter((part) => part.length > 0);
}

function isPrintTitleRowRef(value: string): boolean {
  return /^\$?\d+:\$?\d+$/.test(value);
}

function isPrintTitleColumnRef(value: string): boolean {
  return /^\$?[A-Z]+:\$?[A-Z]+$/i.test(value);
}

function normalizePrintTitleRowRef(value: string): string {
  const match = value.match(/^\$?(\d+):\$?(\d+)$/);
  if (!match) {
    throw new XlsxError(`Invalid print title row reference: ${value}`);
  }

  return `$${match[1]}:$${match[2]}`;
}

function normalizePrintTitleColumnRef(value: string): string {
  const match = value.toUpperCase().match(/^\$?([A-Z]+):\$?([A-Z]+)$/);
  if (!match) {
    throw new XlsxError(`Invalid print title column reference: ${value}`);
  }

  return `$${match[1]}:$${match[2]}`;
}

function buildFormulaCellSnapshot(
  formula: string,
  cachedValue: CellValue,
  styleId: number | null,
): CellSnapshot {
  return {
    exists: true,
    error: null,
    formula,
    rawType:
      typeof cachedValue === "string"
        ? "str"
        : typeof cachedValue === "boolean"
          ? "b"
          : null,
    styleId,
    type: "formula",
    value: cachedValue,
  };
}

function extractCellAttributesSource(cellXml: string): string {
  const openTagEnd = cellXml.indexOf(">");
  if (openTagEnd === -1) {
    throw new XlsxError("Cell XML is missing opening tag");
  }

  let source = cellXml.slice(2, openTagEnd);
  let end = source.length;
  while (end > 0 && /\s/.test(source[end - 1]!)) {
    end -= 1;
  }

  if (end > 0 && source[end - 1] === "/") {
    end -= 1;
    while (end > 0 && /\s/.test(source[end - 1]!)) {
      end -= 1;
    }
  }

  let start = 0;
  while (start < end && /\s/.test(source[start]!)) {
    start += 1;
  }

  source = source.slice(start, end);
  return source;
}

function applyPendingCellMutationsToSheetXml(
  baseIndex: SheetIndex,
  pendingMutations: Iterable<PendingCellMutation>,
): string {
  const mutationsByRow = new Map<number, PendingCellMutation[]>();

  for (const mutation of pendingMutations) {
    const rowMutations = mutationsByRow.get(mutation.rowNumber);
    if (rowMutations) {
      rowMutations.push(mutation);
    } else {
      mutationsByRow.set(mutation.rowNumber, [mutation]);
    }
  }

  let nextSheetXml = baseIndex.xml;
  const rowNumbers = [...mutationsByRow.keys()].sort((left, right) => right - left);

  for (const rowNumber of rowNumbers) {
    const rowMutations = mutationsByRow.get(rowNumber);
    if (!rowMutations) {
      continue;
    }

    const row = baseIndex.rows.get(rowNumber);
    if (!row) {
      const nextRowXml = buildRowXmlFromPendingMutations(rowNumber, `r="${rowNumber}"`, rowMutations);
      if (!nextRowXml) {
        continue;
      }

      const insertionIndex = findRowInsertionIndex(baseIndex, rowNumber);
      nextSheetXml = nextSheetXml.slice(0, insertionIndex) + nextRowXml + nextSheetXml.slice(insertionIndex);
      continue;
    }

    const nextRowXml = applyPendingMutationsToRow(baseIndex.xml, row, rowMutations);
    nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
  }

  return nextSheetXml;
}

function applyPendingMutationsToRow(
  sheetXml: string,
  row: LocatedRow,
  rowMutations: PendingCellMutation[],
): string {
  if (row.selfClosing) {
    return buildRowXmlFromPendingMutations(row.rowNumber, row.attributesSource, rowMutations)
      ?? sheetXml.slice(row.start, row.end);
  }

  let nextRowXml = sheetXml.slice(row.start, row.end);
  const sortedMutations = [...rowMutations].sort((left, right) => right.columnNumber - left.columnNumber);

  for (const mutation of sortedMutations) {
    const existingCell = row.cellsByColumn[mutation.columnNumber];
    if (existingCell) {
      const relativeStart = existingCell.start - row.start;
      const relativeEnd = existingCell.end - row.start;
      nextRowXml =
        mutation.kind === "delete"
          ? nextRowXml.slice(0, relativeStart) + nextRowXml.slice(relativeEnd)
          : nextRowXml.slice(0, relativeStart) + mutation.xml! + nextRowXml.slice(relativeEnd);
      continue;
    }

    if (mutation.kind === "delete") {
      continue;
    }

    const insertionIndex = findPendingCellInsertionIndex(row, mutation.columnNumber) - row.start;
    nextRowXml = nextRowXml.slice(0, insertionIndex) + mutation.xml! + nextRowXml.slice(insertionIndex);
  }

  return normalizeEmptyRowXml(nextRowXml);
}

function buildRowXmlFromPendingMutations(
  rowNumber: number,
  rowAttributesSource: string,
  rowMutations: PendingCellMutation[],
): string | null {
  const nextCells = [...rowMutations]
    .filter((mutation) => mutation.kind === "set")
    .sort((left, right) => left.columnNumber - right.columnNumber)
    .map((mutation) => mutation.xml!);

  if (nextCells.length === 0) {
    return null;
  }

  const attributesSource = rowAttributesSource.length > 0 ? rowAttributesSource : `r="${rowNumber}"`;
  return `<row ${attributesSource}>${nextCells.join("")}</row>`;
}

function findPendingCellInsertionIndex(row: Pick<LocatedRow, "cells" | "innerEnd">, columnNumber: number): number {
  for (const cell of row.cells) {
    if (cell.columnNumber > columnNumber) {
      return cell.start;
    }
  }

  return row.innerEnd;
}

function formatCellDisplayValue(cell: Pick<CellSnapshot, "error" | "value">): string | null {
  if (cell.error) {
    return cell.error.text;
  }

  if (cell.value === null) {
    return null;
  }

  if (typeof cell.value === "boolean") {
    return cell.value ? "TRUE" : "FALSE";
  }

  return String(cell.value);
}
