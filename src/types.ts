export interface ArchiveEntry {
  path: string;
  data: Uint8Array;
}

export type CellValue = string | number | boolean | null;
export type CellType = "missing" | "blank" | "string" | "number" | "boolean" | "error" | "formula";
export type SheetVisibility = "visible" | "hidden" | "veryHidden";
export type OpenXmlStringEnum = string & {};
export type CellStyleHorizontalAlignment =
  | "general"
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed"
  | OpenXmlStringEnum;
export type CellStyleVerticalAlignment =
  | "top"
  | "center"
  | "bottom"
  | "justify"
  | "distributed"
  | OpenXmlStringEnum;
export type CellFontUnderline =
  | "single"
  | "double"
  | "singleAccounting"
  | "doubleAccounting"
  | OpenXmlStringEnum;
export type CellFontScheme = "major" | "minor" | OpenXmlStringEnum;
export type CellFontVerticalAlign = "baseline" | "superscript" | "subscript" | OpenXmlStringEnum;
export type CellFillPatternType =
  | "none"
  | "solid"
  | "mediumGray"
  | "darkGray"
  | "lightGray"
  | "darkHorizontal"
  | "darkVertical"
  | "darkDown"
  | "darkUp"
  | "darkGrid"
  | "darkTrellis"
  | "lightHorizontal"
  | "lightVertical"
  | "lightDown"
  | "lightUp"
  | "lightGrid"
  | "lightTrellis"
  | "gray125"
  | "gray0625"
  | OpenXmlStringEnum;
export type CellBorderStyle =
  | "thin"
  | "medium"
  | "dashed"
  | "dotted"
  | "thick"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot"
  | OpenXmlStringEnum;

export interface CellError {
  code: number | null;
  text: string;
}

export interface CellSnapshot {
  exists: boolean;
  error: CellError | null;
  formula: string | null;
  rawType: string | null;
  styleId: number | null;
  type: CellType;
  value: CellValue;
}

export interface CellEntry extends CellSnapshot {
  address: string;
  rowNumber: number;
  columnNumber: number;
}

export interface CellStyleAlignment {
  horizontal?: CellStyleHorizontalAlignment;
  vertical?: CellStyleVerticalAlignment;
  textRotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  relativeIndent?: number;
  justifyLastLine?: boolean;
  readingOrder?: number;
}

export interface CellStyleAlignmentPatch {
  horizontal?: CellStyleHorizontalAlignment | null;
  vertical?: CellStyleVerticalAlignment | null;
  textRotation?: number | null;
  wrapText?: boolean | null;
  shrinkToFit?: boolean | null;
  indent?: number | null;
  relativeIndent?: number | null;
  justifyLastLine?: boolean | null;
  readingOrder?: number | null;
}

export interface CellStyleDefinition {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  xfId: number | null;
  quotePrefix: boolean | null;
  pivotButton: boolean | null;
  applyNumberFormat: boolean | null;
  applyFont: boolean | null;
  applyFill: boolean | null;
  applyBorder: boolean | null;
  applyAlignment: boolean | null;
  applyProtection: boolean | null;
  alignment: CellStyleAlignment | null;
}

export interface CellStylePatch {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number | null;
  quotePrefix?: boolean | null;
  pivotButton?: boolean | null;
  applyNumberFormat?: boolean | null;
  applyFont?: boolean | null;
  applyFill?: boolean | null;
  applyBorder?: boolean | null;
  applyAlignment?: boolean | null;
  applyProtection?: boolean | null;
  alignment?: CellStyleAlignmentPatch | null;
}

export interface CellNumberFormatDefinition {
  builtin: boolean;
  code: string | null;
  numFmtId: number;
}

export interface CellFontColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
}

export interface CellFontColorPatch {
  rgb?: string | null;
  theme?: number | null;
  indexed?: number | null;
  auto?: boolean | null;
  tint?: number | null;
}

export interface CellFontDefinition {
  bold: boolean | null;
  italic: boolean | null;
  underline: CellFontUnderline | null;
  strike: boolean | null;
  outline: boolean | null;
  shadow: boolean | null;
  condense: boolean | null;
  extend: boolean | null;
  size: number | null;
  name: string | null;
  family: number | null;
  charset: number | null;
  scheme: CellFontScheme | null;
  vertAlign: CellFontVerticalAlign | null;
  color: CellFontColor | null;
}

export interface CellFontPatch {
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: CellFontUnderline | null;
  strike?: boolean | null;
  outline?: boolean | null;
  shadow?: boolean | null;
  condense?: boolean | null;
  extend?: boolean | null;
  size?: number | null;
  name?: string | null;
  family?: number | null;
  charset?: number | null;
  scheme?: CellFontScheme | null;
  vertAlign?: CellFontVerticalAlign | null;
  color?: CellFontColorPatch | null;
}

export interface CellFillColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
}

export interface CellFillColorPatch {
  rgb?: string | null;
  theme?: number | null;
  indexed?: number | null;
  auto?: boolean | null;
  tint?: number | null;
}

export interface CellFillDefinition {
  patternType: CellFillPatternType | null;
  fgColor: CellFillColor | null;
  bgColor: CellFillColor | null;
}

export interface CellFillPatch {
  patternType?: CellFillPatternType | null;
  fgColor?: CellFillColorPatch | null;
  bgColor?: CellFillColorPatch | null;
}

export interface CellBorderColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
}

export interface CellBorderColorPatch {
  rgb?: string | null;
  theme?: number | null;
  indexed?: number | null;
  auto?: boolean | null;
  tint?: number | null;
}

export interface CellBorderSideDefinition {
  style: CellBorderStyle | null;
  color: CellBorderColor | null;
}

export interface CellBorderSidePatch {
  style?: CellBorderStyle | null;
  color?: CellBorderColorPatch | null;
}

export interface CellBorderDefinition {
  left: CellBorderSideDefinition | null;
  right: CellBorderSideDefinition | null;
  top: CellBorderSideDefinition | null;
  bottom: CellBorderSideDefinition | null;
  diagonal: CellBorderSideDefinition | null;
  vertical: CellBorderSideDefinition | null;
  horizontal: CellBorderSideDefinition | null;
  diagonalUp: boolean | null;
  diagonalDown: boolean | null;
  outline: boolean | null;
}

export interface CellBorderPatch {
  left?: CellBorderSidePatch | null;
  right?: CellBorderSidePatch | null;
  top?: CellBorderSidePatch | null;
  bottom?: CellBorderSidePatch | null;
  diagonal?: CellBorderSidePatch | null;
  vertical?: CellBorderSidePatch | null;
  horizontal?: CellBorderSidePatch | null;
  diagonalUp?: boolean | null;
  diagonalDown?: boolean | null;
  outline?: boolean | null;
}

export interface DefinedName {
  hidden: boolean;
  name: string;
  scope: string | null;
  value: string;
}

export interface Hyperlink {
  address: string;
  target: string;
  tooltip: string | null;
  type: "external" | "internal";
}

export interface FreezePane {
  columnCount: number;
  rowCount: number;
  topLeftCell: string;
  activePane: "bottomLeft" | "topRight" | "bottomRight" | null;
}

export interface SheetSelection {
  activeCell: string | null;
  range: string | null;
  pane: "bottomLeft" | "topRight" | "bottomRight" | null;
}

export interface SheetPrintTitles {
  columns: string | null;
  rows: string | null;
}

export interface SheetProtection {
  autoFilter: boolean | null;
  deleteColumns: boolean | null;
  deleteRows: boolean | null;
  formatCells: boolean | null;
  formatColumns: boolean | null;
  formatRows: boolean | null;
  insertColumns: boolean | null;
  insertHyperlinks: boolean | null;
  insertRows: boolean | null;
  objects: boolean | null;
  passwordHash: string | null;
  pivotTables: boolean | null;
  scenarios: boolean | null;
  selectLockedCells: boolean | null;
  selectUnlockedCells: boolean | null;
  sheet: boolean;
  sort: boolean | null;
}

export interface SheetProtectionOptions {
  autoFilter?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertHyperlinks?: boolean;
  insertRows?: boolean;
  objects?: boolean;
  passwordHash?: string;
  pivotTables?: boolean;
  scenarios?: boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  sort?: boolean;
}

export interface SheetComment {
  address: string;
  author: string | null;
  text: string;
}

export interface SheetCommentWriteOptions {
  author?: string;
}

export interface SheetImportRecordsResult {
  headers: string[];
  imported: number;
  inserted: number;
  mode: "append" | "replace" | "update" | "upsert";
  rowCount: number;
  updated: number;
}

export interface SheetUpdateRecordResult {
  record: Record<string, CellValue>;
  row: number | null;
  updated: boolean;
}

export interface SheetUpsertRecordResult {
  inserted: boolean;
  record: Record<string, CellValue>;
  row: number;
}

export interface AutoFilterDefinition {
  range: string;
  columns: AutoFilterColumn[];
  sortState?: SortStateDefinition | null;
}

export type AutoFilterColumn =
  | ValuesFilterColumn
  | CustomFilterColumn
  | DateGroupFilterColumn
  | BlankFilterColumn
  | ColorFilterColumn
  | DynamicFilterColumn
  | Top10FilterColumn
  | IconFilterColumn;

export interface ValuesFilterColumn {
  columnNumber: number;
  kind: "values";
  values: string[];
  includeBlank?: boolean;
}

export interface BlankFilterColumn {
  columnNumber: number;
  kind: "blank";
  mode: "blank" | "nonBlank";
}

export interface CustomFilterColumn {
  columnNumber: number;
  kind: "custom";
  join: "and" | "or";
  conditions: AutoFilterCondition[];
}

export type AutoFilterCondition =
  | {
      operator:
        | "equals"
        | "notEquals"
        | "greaterThan"
        | "greaterThanOrEqual"
        | "lessThan"
        | "lessThanOrEqual";
      value: string | number;
    }
  | {
      operator: "contains" | "notContains" | "beginsWith" | "endsWith";
      value: string;
    };

export interface DateGroupFilterColumn {
  columnNumber: number;
  kind: "dateGroup";
  items: DateGroupItem[];
}

export interface ColorFilterColumn {
  columnNumber: number;
  kind: "color";
  dxfId: number;
  cellColor: boolean;
}

export interface DynamicFilterColumn {
  columnNumber: number;
  kind: "dynamic";
  type: string;
  val?: number;
  maxVal?: number;
  valIso?: string;
  maxValIso?: string;
}

export interface Top10FilterColumn {
  columnNumber: number;
  kind: "top10";
  top: boolean;
  percent: boolean;
  value: number;
  filterValue?: number;
}

export interface IconFilterColumn {
  columnNumber: number;
  kind: "icon";
  iconSet?: string;
  iconId?: number;
}

export interface DateGroupItem {
  year: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
  dateTimeGrouping: "year" | "month" | "day" | "hour" | "minute" | "second";
}

export interface SortStateDefinition {
  range: string;
  conditions: SortConditionDefinition[];
}

export interface SortConditionDefinition {
  columnNumber: number;
  descending?: boolean;
}

export interface SortRangeOptions {
  conditions: SortConditionDefinition[];
  hasHeaderRow?: boolean;
}

export interface SheetTable {
  readonly name: string;
  readonly displayName: string;
  readonly range: string;
  readonly path: string;

  getAutoFilterDefinition(): AutoFilterDefinition | null;
  setAutoFilterDefinition(definition: AutoFilterDefinition): void;
  setAutoFilterColumn(column: AutoFilterColumn): void;
  clearAutoFilterColumns(columnNumbers?: number[]): void;
}

export interface SheetTableSummary {
  name: string;
  displayName: string;
  range: string;
  path: string;
}

export interface SheetTableWithAutoFilterSummary extends SheetTableSummary {
  autoFilter: AutoFilterDefinition | null;
}

export interface DataValidation {
  range: string;
  type: string | null;
  operator: string | null;
  allowBlank: boolean | null;
  showInputMessage: boolean | null;
  showErrorMessage: boolean | null;
  showDropDown: boolean | null;
  errorStyle: string | null;
  errorTitle: string | null;
  error: string | null;
  promptTitle: string | null;
  prompt: string | null;
  imeMode: string | null;
  formula1: string | null;
  formula2: string | null;
}

export interface SetFormulaOptions {
  cachedValue?: CellValue;
}

export interface RecalculateSummary {
  cells: number;
  sheets: number;
  updated: number;
}

export interface SetDefinedNameOptions {
  scope?: string;
}

export interface WorkbookCreateSheetOptions {
  columnWidths?: Record<string, number | null>;
  comments?: SheetComment[];
  frozenPane?: {
    columnCount: number;
    rowCount?: number;
  };
  headers?: string[];
  headerStyle?: CellStylePatch;
  name: string;
  printArea?: string | null;
  printTitles?: {
    columns?: string | null;
    rows?: string | null;
  };
  rangeStyles?: Array<{
    backgroundColor?: string | null;
    numberFormat?: string;
    patch?: CellStylePatch;
    range: string;
  }>;
  records?: Array<Record<string, CellValue>>;
  rowHeights?: Record<string, number | null>;
  visibility?: SheetVisibility;
}

export interface WorkbookCreateOptions {
  activeSheet?: string;
  author?: string;
  createdAt?: Date;
  modifiedBy?: string;
  sheets?: Array<string | WorkbookCreateSheetOptions>;
}

export interface CreateTableSheetOptions {
  headerRow?: number;
  headers?: string[];
  records?: Array<Record<string, CellValue>>;
}

export interface SheetExportRecordsOptions {
  format?: "csv" | "json";
  headerRow?: number;
  includeHeaders?: boolean;
  lineEnding?: "\n" | "\r\n";
}

export interface SheetImportRecordsOptions {
  headerRow?: number;
  headerOrder?: string[];
  inferTypes?: boolean;
  keyField?: string;
  mode?: "append" | "replace" | "update" | "upsert";
  trimHeaders?: boolean;
  trimValues?: boolean;
}

export interface SetHyperlinkOptions {
  text?: string;
  tooltip?: string;
}

export interface SetDataValidationOptions {
  type?: string;
  operator?: string;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  showDropDown?: boolean;
  errorStyle?: string;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
  imeMode?: string;
  formula1?: string;
  formula2?: string;
}
