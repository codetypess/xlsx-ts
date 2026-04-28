# AutoFilter and Sort Metadata

## Scope

This document describes the shipped worksheet/table auto-filter surface in `fastxlsx`, plus the physical row-sorting API that now sits next to it.

The original gap was that `fastxlsx` only exposed worksheet `autoFilter` as a range string and would flatten nested filter XML when that range changed. That gap is now closed.

## Delivered Behavior

- Worksheet auto-filter metadata is available as a structured object, not only as `ref`.
- Table auto-filter metadata is available through a table handle and optional table summaries.
- `setAutoFilter(range)` preserves nested children instead of rewriting `<autoFilter/>` destructively.
- Targeted filter edits preserve unrelated `filterColumn` XML, unknown attributes, and `extLst`.
- Row/column structure transforms update filter ranges, column indexes, and sort-condition references.
- `sheet.sortRange(range, options)` performs physical row movement and updates linked worksheet metadata.

## Compatibility

The following existing APIs remain supported:

- `sheet.getAutoFilter(): string | null`
- `sheet.setAutoFilter(range: string): void`
- `sheet.removeAutoFilter(): void`
- `sheet.clearAutoFilter(): void`
- `sheet.getTables(): Array<{ name: string; displayName: string; range: string; path: string }>`

Their semantics are upgraded:

- `getAutoFilter()` still returns only the normalized range.
- `setAutoFilter(range)` updates the logical filter range while preserving existing nested filter content when possible.
- `removeAutoFilter()` remains destructive and removes worksheet `autoFilter` plus worksheet `sortState`.
- `getTables()` remains source-compatible, while `getTables({ includeAutoFilter: true })` adds structured filter metadata.

## Public Surface

### Worksheet APIs

```ts
interface Sheet {
  getAutoFilter(): string | null;
  getAutoFilterDefinition(): AutoFilterDefinition | null;
  setAutoFilter(range: string): void;
  setAutoFilterDefinition(definition: AutoFilterDefinition): void;
  setAutoFilterColumn(column: AutoFilterColumn): void;
  clearAutoFilterColumns(columnNumbers?: number[]): void;

  getTables(): SheetTableSummary[];
  getTables(options: { includeAutoFilter: true }): SheetTableWithAutoFilterSummary[];
  getTable(name: string): SheetTable;
  tryGetTable(name: string): SheetTable | null;

  sortRange(range: string, options: SortRangeOptions): void;
}
```

### Table Handle

```ts
interface SheetTable {
  readonly name: string;
  readonly displayName: string;
  readonly range: string;
  readonly path: string;

  getAutoFilterDefinition(): AutoFilterDefinition | null;
  setAutoFilterDefinition(definition: AutoFilterDefinition): void;
  setAutoFilterColumn(column: AutoFilterColumn): void;
  clearAutoFilterColumns(columnNumbers?: number[]): void;
}
```

`SheetTable` is a live handle over the backing table part. It does not copy table XML into a detached cache.

### Summary Types

```ts
interface SheetTableSummary {
  name: string;
  displayName: string;
  range: string;
  path: string;
}

interface SheetTableWithAutoFilterSummary extends SheetTableSummary {
  autoFilter: AutoFilterDefinition | null;
}
```

## Types

```ts
interface AutoFilterDefinition {
  range: string;
  columns: AutoFilterColumn[];
  sortState?: SortStateDefinition | null;
}

type AutoFilterColumn =
  | ValuesFilterColumn
  | BlankFilterColumn
  | CustomFilterColumn
  | DateGroupFilterColumn
  | ColorFilterColumn
  | DynamicFilterColumn
  | Top10FilterColumn
  | IconFilterColumn;

interface ValuesFilterColumn {
  columnNumber: number;
  kind: "values";
  values: string[];
  includeBlank?: boolean;
}

interface BlankFilterColumn {
  columnNumber: number;
  kind: "blank";
  mode: "blank" | "nonBlank";
}

interface CustomFilterColumn {
  columnNumber: number;
  kind: "custom";
  join: "and" | "or";
  conditions: AutoFilterCondition[];
}

type AutoFilterCondition =
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

interface DateGroupFilterColumn {
  columnNumber: number;
  kind: "dateGroup";
  items: DateGroupItem[];
}

interface DateGroupItem {
  year: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
  dateTimeGrouping: "year" | "month" | "day" | "hour" | "minute" | "second";
}

interface ColorFilterColumn {
  columnNumber: number;
  kind: "color";
  dxfId: number;
  cellColor: boolean;
}

interface DynamicFilterColumn {
  columnNumber: number;
  kind: "dynamic";
  type: string;
  val?: number;
  maxVal?: number;
  valIso?: string;
  maxValIso?: string;
}

interface Top10FilterColumn {
  columnNumber: number;
  kind: "top10";
  top: boolean;
  percent: boolean;
  value: number;
  filterValue?: number;
}

interface IconFilterColumn {
  columnNumber: number;
  kind: "icon";
  iconSet?: string;
  iconId?: number;
}

interface SortStateDefinition {
  range: string;
  conditions: SortConditionDefinition[];
}

interface SortConditionDefinition {
  columnNumber: number;
  descending?: boolean;
}

interface SortRangeOptions {
  conditions: SortConditionDefinition[];
  hasHeaderRow?: boolean;
}
```

### Public Type Rules

- Public filter APIs use absolute worksheet `columnNumber`.
- Public filter APIs do not expose OOXML-relative `colId`.
- Blank and non-blank filters are first-class through `kind: "blank"`.
- Color filters currently expose OOXML-oriented metadata through `dxfId` and `cellColor`, not resolved RGB values.
- Worksheet and table filters use the same logical type model.

## Mutation Semantics

### `setAutoFilter(range)`

- Creates a new empty `autoFilter` if none exists.
- Updates only the range when a filter already exists.
- Preserves nested `filterColumn` children, `sortState`, and supported/unsupported descendants when still valid.
- Re-bases `filterColumn@colId` when the range anchor column changes.
- Drops only columns or sort conditions that become invalid under the new range.

### `setAutoFilterDefinition(definition)`

- Replaces supported column definitions for the `columnNumber`s present in `definition.columns`.
- Replaces the worksheet or table `sortState` with `definition.sortState`.
- Preserves unrelated unsupported root children and unsupported in-range columns.
- Treats omission of an in-range previously supported column as an explicit clear of that supported column.

### `setAutoFilterColumn(column)`

- Rewrites one logical column filter.
- Preserves unrelated columns and ordering.
- Preserves unsupported children for the same column unless they conflict with the supported rewritten shape.
- Requires an existing filter range.

### `clearAutoFilterColumns(columnNumbers?)`

- With explicit column numbers, removes only those filter columns.
- With no argument, removes all filter columns but keeps the filter range.
- Does not remove `sortState`.
- Full destructive removal remains `removeAutoFilter()`.

### Table Rules

- Table filter setters follow the same merge/preservation rules as worksheet filters.
- `table.setAutoFilterDefinition(definition)` validates that `definition.range` matches the current table range.
- Table writes create nested `sortState` when needed and preserve unrelated table metadata.

## OOXML Mapping

### Worksheet

Supported structured data is read from worksheet-level `autoFilter` and worksheet-level `sortState`.

Handled nodes and attributes include:

- `autoFilter@ref`
- `autoFilter/filterColumn@colId`
- `filters/filter@val`
- `filters@blank`
- `blank`
- `customFilters@and`
- `customFilters/customFilter@operator`
- `customFilters/customFilter@val`
- `dateGroupItem`
- `colorFilter@dxfId`
- `colorFilter@cellColor`
- `dynamicFilter`
- `top10`
- `iconFilter`
- `sortState@ref`
- `sortState/sortCondition`

### Table

Supported logical data is read from the table part:

- table `ref`
- nested table `autoFilter`
- nested table `sortState`

Worksheet and table APIs intentionally expose the same logical shape even though the XML storage locations differ.

## Non-destructive Roundtrip

The implementation is designed to be source-stable for unrelated filter metadata:

- `setAutoFilter(range)` preserves nested worksheet children instead of collapsing them into `<autoFilter ref="..."/>`.
- Table range rewrites preserve nested table `autoFilter` children.
- Unsupported `filterColumn` nodes survive supported-column rewrites.
- `extLst` survives targeted edits.
- Structured reads only surface supported logical shapes, but unsupported XML is still retained internally and roundtripped.

## Structure Transform Semantics

When rows or columns are inserted or deleted, filter metadata stays attached to the same logical sheet columns and ranges.

Maintained metadata includes:

- filter `range`
- sort `range`
- `filterColumn@colId`
- `sortCondition@ref`

### Column Insert/Delete

- Stored filters are interpreted in absolute `columnNumber` coordinates.
- Column transforms apply in absolute coordinates first.
- Deleted columns are dropped.
- OOXML-relative `colId` is recomputed from the transformed range start column.
- The same logic is applied for table filters relative to the current table range.

### Row Insert/Delete

- Filter and sort ranges are updated.
- Column-bound filter definitions are preserved.
- Sort state is removed only if the transformed sort range becomes invalid.

## Physical Sorting

`sheet.sortRange(range, options)` performs a real row reorder inside one rectangular range.

Current behavior:

- Supports single-column and multi-column sorts.
- `hasHeaderRow` is supported and defaults to `true`.
- Moves cell content, styles, formulas, merged ranges, hyperlinks, and data validations inside the sorted rectangle.
- Updates worksheet `sortState`.
- Updates matching table `sortState` when the sorted range matches the table range or table auto-filter range.
- Leaves sorting as an explicit operation separate from metadata-only `sortState` writes.

### Current Limits

- Worksheet comments inside the sortable data area are not supported; the method throws.
- Metadata ranges that overlap the sortable rectangle without being fully contained are rejected.
- Metadata ranges that would become non-contiguous after row permutation are rejected.

These guards are intentional. They prevent silent corruption in cases where safe rewrite rules are not yet implemented.

## CLI

Structured worksheet filter commands are available:

- `sheet filter get --definition`
- `sheet filter set`
- `sheet filter set-definition`
- `sheet filter set-column`
- `sheet filter clear-columns`
- `sheet filter clear`

The CLI accepts and emits JSON using the same `AutoFilterDefinition` and `AutoFilterColumn` shapes as the library API.

## Validation

Coverage exists across:

- `test/lossless.test.ts`
- `test/cli.test.ts`
- `test/interop-matrix.test.ts`
- `test/real-files.test.ts`

Acceptance status:

- Worksheet filters can be read as structured definitions.
- Worksheet filters can be edited without flattening nested XML.
- Table filters can be read and edited through a table handle.
- Unsupported filter XML survives partial edits and roundtrip.
- Insert/delete row and column operations keep filter metadata consistent.
- Advanced typed filter kinds are exposed publicly.
- Structured filter CLI behavior is available.
- `sortRange()` performs physical sorting with explicit safety checks.
