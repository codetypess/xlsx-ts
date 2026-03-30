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
  DataValidation,
  FreezePane,
  Hyperlink,
  SetDataValidationOptions,
  SetFormulaOptions,
  SetHyperlinkOptions,
  SheetSelection,
} from "./types.js";
import { XlsxError } from "./errors.js";
import {
  buildSheetIndex,
  parseCellSnapshot,
  type SheetIndex,
} from "./sheet/sheet-index.js";
import {
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
import { parseMergedRanges, updateMergedRanges } from "./sheet/sheet-merge.js";
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
  buildEmptyStyledRowXml,
  buildStyledRowXml,
  parseColumnStyleId,
  parseRowStyleId,
  transformColumnStyleDefinitions,
  updateColumnStyleIdInSheetXml,
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
  buildDataValidationXml,
  buildExternalHyperlinkXml,
  buildInternalHyperlinkXml,
  getHyperlinkRelationshipId,
  HYPERLINK_RELATIONSHIP_TYPE,
  parseHyperlinkRelationshipTargets,
  parseSheetAutoFilter,
  parseSheetDataValidations,
  parseSheetHyperlinks,
  removeAutoFilterFromSheetXml,
  removeDataValidationFromSheetXml,
  removeHyperlinkFromSheetXml,
  upsertAutoFilterInSheetXml,
  upsertDataValidationInSheetXml,
  upsertHyperlinkInSheetXml,
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
  removeTablePartsFromSheetXml,
  TABLE_CONTENT_TYPE,
  TABLE_RELATIONSHIP_TYPE,
  upsertRelationship,
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

export class Sheet {
  name: string;
  readonly path: string;
  readonly relationshipId: string;

  private readonly cellHandles = new Map<string, Cell>();
  private revision = 0;
  private readonly workbook: Workbook;
  private sheetIndex?: SheetIndex;

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

  getCell(address: string): CellValue;
  getCell(rowNumber: number, column: number | string): CellValue;
  getCell(addressOrRowNumber: string | number, column?: number | string): CellValue {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).value;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).value;
  }

  getStyleId(address: string): number | null;
  getStyleId(rowNumber: number, column: number | string): number | null;
  getStyleId(addressOrRowNumber: string | number, column?: number | string): number | null {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).styleId;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).styleId;
  }

  getStyle(address: string): CellStyleDefinition | null;
  getStyle(rowNumber: number, column: number | string): CellStyleDefinition | null;
  getStyle(addressOrRowNumber: string | number, column?: number | string): CellStyleDefinition | null {
    const styleId =
      typeof addressOrRowNumber === "number"
        ? this.readCellSnapshotByIndexes(addressOrRowNumber, column).styleId
        : this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).styleId;
    return this.workbook.getStyle(styleId ?? 0);
  }

  getAlignment(address: string): CellStyleAlignment | null;
  getAlignment(rowNumber: number, column: number | string): CellStyleAlignment | null;
  getAlignment(addressOrRowNumber: string | number, column?: number | string): CellStyleAlignment | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style?.alignment ?? null;
  }

  getFont(address: string): CellFontDefinition | null;
  getFont(rowNumber: number, column: number | string): CellFontDefinition | null;
  getFont(addressOrRowNumber: string | number, column?: number | string): CellFontDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getFont(style.fontId) : null;
  }

  getFill(address: string): CellFillDefinition | null;
  getFill(rowNumber: number, column: number | string): CellFillDefinition | null;
  getFill(addressOrRowNumber: string | number, column?: number | string): CellFillDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getFill(style.fillId) : null;
  }

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

  getBorder(address: string): CellBorderDefinition | null;
  getBorder(rowNumber: number, column: number | string): CellBorderDefinition | null;
  getBorder(addressOrRowNumber: string | number, column?: number | string): CellBorderDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getBorder(style.borderId) : null;
  }

  getNumberFormat(address: string): CellNumberFormatDefinition | null;
  getNumberFormat(rowNumber: number, column: number | string): CellNumberFormatDefinition | null;
  getNumberFormat(addressOrRowNumber: string | number, column?: number | string): CellNumberFormatDefinition | null {
    const style =
      typeof addressOrRowNumber === "number"
        ? this.getStyle(addressOrRowNumber, column!)
        : this.getStyle(addressOrRowNumber);
    return style ? this.workbook.getNumberFormat(style.numFmtId) : null;
  }

  getColumnStyleId(column: number | string): number | null {
    const columnNumber = normalizeColumnNumber(column);
    return parseColumnStyleId(this.getSheetIndex().xml, columnNumber);
  }

  getColumnStyle(column: number | string): CellStyleDefinition | null {
    const styleId = this.getColumnStyleId(column);
    return styleId === null ? null : this.workbook.getStyle(styleId);
  }

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

  rename(name: string): void {
    this.workbook.renameSheet(this.name, name);
  }

  getFormula(address: string): string | null;
  getFormula(rowNumber: number, column: number | string): string | null;
  getFormula(addressOrRowNumber: string | number, column?: number | string): string | null {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).formula;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).formula;
  }

  get rowCount(): number {
    return this.getSheetIndex().usedBounds?.maxRow ?? 0;
  }

  get columnCount(): number {
    return this.getSheetIndex().usedBounds?.maxColumn ?? 0;
  }

  getHeaders(headerRowNumber = 1): string[] {
    assertRowNumber(headerRowNumber);
    return this.getRow(headerRowNumber).map((value) => (typeof value === "string" ? value : ""));
  }

  getRowStyleId(rowNumber: number): number | null {
    assertRowNumber(rowNumber);
    return parseRowStyleId(this.getSheetIndex().rows.get(rowNumber)?.attributesSource);
  }

  getRowStyle(rowNumber: number): CellStyleDefinition | null {
    const styleId = this.getRowStyleId(rowNumber);
    return styleId === null ? null : this.workbook.getStyle(styleId);
  }

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

  getRowEntries(rowNumber: number): CellEntry[] {
    assertRowNumber(rowNumber);

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row) {
      return [];
    }

    return row.cells.map((cell) => createCellEntry(cell));
  }

  getColumn(column: number | string): CellValue[] {
    const columnNumber = normalizeColumnNumber(column);
    const cells = [...this.getSheetIndex().cells.values()]
      .filter((cell) => cell.columnNumber === columnNumber)
      .sort((left, right) => left.rowNumber - right.rowNumber);

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

  getColumnEntries(column: number | string): CellEntry[] {
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

  getCellEntries(): CellEntry[] {
    return Array.from(this.iterCellEntries());
  }

  *iterCellEntries(): IterableIterator<CellEntry> {
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

  getUsedRange(): string | null {
    return formatUsedRangeBounds(this.getSheetIndex().usedBounds);
  }

  getMergedRanges(): string[] {
    return parseMergedRanges(this.getSheetIndex().xml);
  }

  getAutoFilter(): string | null {
    return parseSheetAutoFilter(this.getSheetIndex().xml);
  }

  getFreezePane(): FreezePane | null {
    return parseSheetFreezePane(this.getSheetIndex().xml);
  }

  getSelection(): SheetSelection | null {
    return parseSheetSelection(this.getSheetIndex().xml);
  }

  getDataValidations(): DataValidation[] {
    return parseSheetDataValidations(this.getSheetIndex().xml);
  }

  getTables(): Array<{ name: string; displayName: string; range: string; path: string }> {
    return parseSheetTables(this.getTableReferences(), (path) => this.workbook.readEntryText(path));
  }

  getHyperlinks(): Hyperlink[] {
    return parseSheetHyperlinks(this.getSheetIndex().xml, parseHyperlinkRelationshipTargets(this.readSheetRelationshipsXml()));
  }

  setHyperlink(address: string, target: string, options: SetHyperlinkOptions = {}): void {
    const normalizedAddress = normalizeCellAddress(address);
    if (options.text !== undefined) {
      this.setCell(normalizedAddress, options.text);
    }

    const currentRelationshipId = getHyperlinkRelationshipId(this.getSheetIndex().xml, normalizedAddress);
    let relationshipsXml = this.readSheetRelationshipsXml();
    let relationshipId: string | null = currentRelationshipId;

    if (target.startsWith("#")) {
      if (currentRelationshipId) {
        relationshipsXml = removeRelationshipById(relationshipsXml, currentRelationshipId);
      }

      this.writeSheetXml(
        upsertHyperlinkInSheetXml(
          this.getSheetIndex().xml,
          buildInternalHyperlinkXml(normalizedAddress, target, options.tooltip),
          normalizedAddress,
        ),
      );
      this.writeSheetRelationshipsXml(relationshipsXml);
      return;
    }

    relationshipId ??= getNextRelationshipIdFromXml(relationshipsXml);
    relationshipsXml = upsertRelationship(
      relationshipsXml,
      relationshipId,
      HYPERLINK_RELATIONSHIP_TYPE,
      target,
      "External",
    );
    this.writeSheetXml(
      upsertHyperlinkInSheetXml(
        this.getSheetIndex().xml,
        buildExternalHyperlinkXml(normalizedAddress, relationshipId, options.tooltip),
        normalizedAddress,
      ),
    );
    this.writeSheetRelationshipsXml(relationshipsXml);
  }

  removeHyperlink(address: string): void {
    const normalizedAddress = normalizeCellAddress(address);
    const currentRelationshipId = getHyperlinkRelationshipId(this.getSheetIndex().xml, normalizedAddress);
    const nextSheetXml = removeHyperlinkFromSheetXml(this.getSheetIndex().xml, normalizedAddress);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }

    if (currentRelationshipId) {
      this.writeSheetRelationshipsXml(removeRelationshipById(this.readSheetRelationshipsXml(), currentRelationshipId));
    }
  }

  setAutoFilter(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const nextSheetXml = upsertAutoFilterInSheetXml(this.getSheetIndex().xml, normalizedRange);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  removeAutoFilter(): void {
    const nextSheetXml = removeAutoFilterFromSheetXml(this.getSheetIndex().xml);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  freezePane(columnCount: number, rowCount = 0): void {
    assertFreezeSplit(columnCount, rowCount);
    const nextSheetXml = upsertFreezePaneInSheetXml(this.getSheetIndex().xml, columnCount, rowCount);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  unfreezePane(): void {
    const nextSheetXml = removeFreezePaneFromSheetXml(this.getSheetIndex().xml);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

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

  removeDataValidation(range: string): void {
    const normalizedRange = normalizeSqref(range);
    const nextSheetXml = removeDataValidationFromSheetXml(this.getSheetIndex().xml, normalizedRange);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

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

  setCell(address: string, value: CellValue): void;
  setCell(rowNumber: number, column: number | string, value: CellValue): void;
  setCell(addressOrRowNumber: string | number, columnOrValue: number | string | CellValue, value?: CellValue): void {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, typeof addressOrRowNumber === "number" ? columnOrValue as number | string : undefined);
    const existingCell = this.getSheetIndex().cells.get(normalizedAddress);
    const nextValue = resolveSetCellValue(addressOrRowNumber, columnOrValue, value);
    this.writeCellXml(
      normalizedAddress,
      buildValueCellXml(normalizedAddress, nextValue, existingCell?.attributesSource),
    );
  }

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
    const index = this.getSheetIndex();
    const existingCell = index.cells.get(normalizedAddress);

    this.writeCellXml(
      normalizedAddress,
      buildStyledCellXml(
        normalizedAddress,
        nextStyleId,
        existingCell?.attributesSource,
        existingCell ? index.xml.slice(existingCell.start, existingCell.end) : undefined,
      ),
    );
  }

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

  setColumnStyle(column: number | string, patch: CellStylePatch): number {
    return this.cloneColumnStyle(column, patch);
  }

  cloneColumnStyle(column: number | string, patch: CellStylePatch = {}): number {
    const nextStyleId = this.workbook.cloneStyle(this.getColumnStyleId(column) ?? 0, patch);
    this.setColumnStyleId(column, nextStyleId);
    return nextStyleId;
  }

  deleteCell(address: string): void;
  deleteCell(rowNumber: number, column: number | string): void;
  deleteCell(addressOrRowNumber: string | number, column?: number | string): void {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, column);
    const index = this.getSheetIndex();
    const existingCell = index.cells.get(normalizedAddress);

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

  setFormula(address: string, formula: string, options?: SetFormulaOptions): void;
  setFormula(rowNumber: number, column: number | string, formula: string, options?: SetFormulaOptions): void;
  setFormula(
    addressOrRowNumber: string | number,
    columnOrFormula: number | string,
    formulaOrOptions?: string | SetFormulaOptions,
    options: SetFormulaOptions = {},
  ): void {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, typeof addressOrRowNumber === "number" ? columnOrFormula as number | string : undefined);
    const existingCell = this.getSheetIndex().cells.get(normalizedAddress);
    const { formula, formulaOptions } = resolveSetFormulaArguments(
      addressOrRowNumber,
      columnOrFormula,
      formulaOrOptions,
      options,
    );
    this.writeCellXml(
      normalizedAddress,
      buildFormulaCellXml(
        normalizedAddress,
        formula,
        formulaOptions.cachedValue ?? null,
        existingCell?.attributesSource,
      ),
    );
  }

  getRevision(): number {
    return this.revision;
  }

  setHeaders(headers: string[], headerRowNumber = 1, startColumn = 1): void {
    assertRowNumber(headerRowNumber);
    assertColumnNumber(startColumn);
    this.setRow(headerRowNumber, headers, startColumn);
  }

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

  setRowStyle(rowNumber: number, patch: CellStylePatch): number {
    return this.cloneRowStyle(rowNumber, patch);
  }

  cloneRowStyle(rowNumber: number, patch: CellStylePatch = {}): number {
    assertRowNumber(rowNumber);
    const nextStyleId = this.workbook.cloneStyle(this.getRowStyleId(rowNumber) ?? 0, patch);
    this.setRowStyleId(rowNumber, nextStyleId);
    return nextStyleId;
  }

  addMergedRange(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const ranges = this.getMergedRanges();
    if (ranges.includes(normalizedRange)) {
      return;
    }

    this.writeSheetXml(updateMergedRanges(this.getSheetIndex().xml, [...ranges, normalizedRange]));
  }

  removeMergedRange(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const ranges = this.getMergedRanges().filter((candidate) => candidate !== normalizedRange);
    this.writeSheetXml(updateMergedRanges(this.getSheetIndex().xml, ranges));
  }

  setRow(rowNumber: number, values: CellValue[], startColumn = 1): void {
    assertRowNumber(rowNumber);
    assertColumnNumber(startColumn);

    for (let columnOffset = 0; columnOffset < values.length; columnOffset += 1) {
      this.setCell(makeCellAddress(rowNumber, startColumn + columnOffset), values[columnOffset]);
    }
  }

  appendRow(values: CellValue[], startColumn = 1): number {
    assertColumnNumber(startColumn);
    const rowNumber = (this.getSheetIndex().rowNumbers.at(-1) ?? 0) + 1;
    this.setRow(rowNumber, values, startColumn);
    return rowNumber;
  }

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

  setColumn(column: number | string, values: CellValue[], startRow = 1): void {
    const columnNumber = normalizeColumnNumber(column);
    assertRowNumber(startRow);

    for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
      this.setCell(makeCellAddress(startRow + rowOffset, columnNumber), values[rowOffset]);
    }
  }

  addRecord(record: Record<string, CellValue>, headerRowNumber = 1): void {
    const headerMap = this.getHeaderMap(headerRowNumber);
    if (Object.keys(record).length === 0) {
      return;
    }

    const nextRowNumber = Math.max(headerRowNumber + 1, (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1);
    this.writeRecordRow(nextRowNumber, record, headerMap, false);
  }

  addRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    if (records.length === 0) {
      return;
    }

    const headerMap = this.getHeaderMap(headerRowNumber);
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

  setRecord(rowNumber: number, record: Record<string, CellValue>, headerRowNumber = 1): void {
    assertRowNumber(rowNumber);

    const headerMap = this.getHeaderMap(headerRowNumber);
    if (Object.keys(record).length === 0) {
      return;
    }

    this.writeRecordRow(rowNumber, record, headerMap, false);
  }

  setRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    const headerMap = this.getHeaderMap(headerRowNumber);
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

  deleteRecords(rowNumbers: number[], headerRowNumber = 1): void {
    assertRowNumber(headerRowNumber);

    const uniqueRows = [...new Set(rowNumbers)];
    uniqueRows.sort((left, right) => right - left);

    for (const rowNumber of uniqueRows) {
      this.deleteRecord(rowNumber, headerRowNumber);
    }
  }

  readCellSnapshot(address: string): CellSnapshot {
    const locatedCell = this.getSheetIndex().cells.get(normalizeCellAddress(address));
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
    const row = this.getSheetIndex().rows.get(rowNumber);
    return parseCellSnapshot(row?.cellsByColumn[columnNumber]);
  }

  private getHeaderMap(headerRowNumber: number): Map<string, number> {
    assertRowNumber(headerRowNumber);

    const headers = this.getRow(headerRowNumber);
    const headerMap = new Map<string, number>();

    headers.forEach((value, index) => {
      if (typeof value === "string" && value.length > 0 && !headerMap.has(value)) {
        headerMap.set(value, index + 1);
      }
    });

    return headerMap;
  }

  private writeRecordRow(
    rowNumber: number,
    record: Record<string, CellValue>,
    headerMap: Map<string, number>,
    replaceMissingKeys: boolean,
  ): void {
    const keys = Object.keys(record);

    for (const key of keys) {
      if (!headerMap.has(key)) {
        throw new XlsxError(`Header not found: ${key}`);
      }
    }

    if (replaceMissingKeys) {
      for (const [header, columnNumber] of headerMap) {
        const nextValue = Object.hasOwn(record, header) ? record[header] ?? null : null;
        this.setCell(makeCellAddress(rowNumber, columnNumber), nextValue);
      }
      return;
    }

    for (const key of keys) {
      const columnNumber = headerMap.get(key);
      if (!columnNumber) {
        continue;
      }

      this.setCell(makeCellAddress(rowNumber, columnNumber), record[key] ?? null);
    }
  }

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

  private getSheetIndex(): SheetIndex {
    if (this.sheetIndex) {
      return this.sheetIndex;
    }

    this.sheetIndex = buildSheetIndex(this.workbook, this.workbook.readEntryText(this.path));
    return this.sheetIndex;
  }

  private writeCellXml(address: string, cellXml: string): void {
    const index = this.getSheetIndex();
    const existingCell = index.cells.get(address);
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
    const normalizedSheetXml = updateDimensionRef(indexedSheet);

    this.workbook.writeEntryText(this.path, normalizedSheetXml);
    this.sheetIndex =
      normalizedSheetXml === nextSheetXml ? indexedSheet : buildSheetIndex(this.workbook, normalizedSheetXml);
    this.revision += 1;
  }
}
