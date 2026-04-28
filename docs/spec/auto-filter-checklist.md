# AutoFilter Checklist

This file tracks the delivered worksheet/table auto-filter and sorting work.

## Current Status

- Structured worksheet auto-filter APIs are shipped:
  - `sheet.getAutoFilterDefinition()`
  - `sheet.setAutoFilterDefinition(definition)`
  - `sheet.setAutoFilterColumn(column)`
  - `sheet.clearAutoFilterColumns(columnNumbers?)`
- Table-level filter APIs are shipped:
  - `sheet.getTable(name)`
  - `sheet.tryGetTable(name)`
  - `sheet.getTables({ includeAutoFilter: true })`
  - `table.getAutoFilterDefinition()`
  - `table.setAutoFilterDefinition(definition)`
  - `table.setAutoFilterColumn(column)`
  - `table.clearAutoFilterColumns(columnNumbers?)`
- Typed filter support now covers:
  - values filters
  - blank and non-blank filters
  - custom filters
  - text conditions via wildcard mapping
  - date-group filters
  - color filters
  - dynamic filters
  - top-10 filters
  - icon filters
  - `sortState`
- Non-destructive behavior is in place:
  - `setAutoFilter(range)` preserves nested worksheet filter content
  - table/table-range rewrites preserve nested `autoFilter` content
  - targeted column edits preserve unrelated `filterColumn` XML, attributes, and `extLst`
  - row/column structure transforms keep `autoFilter@ref`, `sortState@ref`, `filterColumn@colId`, and `sortCondition@ref` aligned
- Physical sorting is shipped through `sheet.sortRange(range, options)`.

## MVP

- [x] Add structured worksheet-level auto-filter read support.
- [x] Add structured worksheet-level auto-filter write support.
- [x] Preserve nested `filterColumn` content when only `autoFilter@ref` changes.
- [x] Preserve nested table `autoFilter` content when table `ref` changes.
- [x] Support column-local worksheet filter updates.
- [x] Support column-local table filter updates.
- [x] Expose absolute `columnNumber` in the public API instead of OOXML `colId`.
- [x] Read and write worksheet `sortState`.
- [x] Read and write table `autoFilter`.
- [x] Preserve unsupported or unknown `filterColumn` XML during partial edits.
- [x] Preserve `extLst` during partial edits.
- [x] Keep existing primitive APIs working:
  - `getAutoFilter(): string | null`
  - `setAutoFilter(range)`
  - `removeAutoFilter()`
  - `clearAutoFilter()`
- [x] Keep row/column structure transforms in sync with filter metadata:
  - top-level `autoFilter@ref`
  - top-level `sortState@ref`
  - `filterColumn@colId`
  - `sortCondition@ref`

## Phase 2

- [x] Add typed public support for color filters.
- [x] Add typed public support for dynamic filters.
- [x] Add typed public support for top-10 filters.
- [x] Add typed public support for icon filters.
- [x] Preserve and roundtrip the above filter kinds during targeted rewrites.
- [x] Extend CLI surfaces for structured filter definitions.
- [x] Add producer-interop coverage for structured filter definitions and advanced filter roundtrip.

## Sorting

- [x] Add `sheet.sortRange(range, options)`.
- [x] Support single-column and multi-column sorts.
- [x] Support header-row-aware sorting.
- [x] Ensure sort operations move:
  - cell values
  - styles
  - formulas
  - merged ranges
  - table sort/filter metadata
  - data validations
  - hyperlinks
- [x] Keep metadata-only `sortState` writes separate from real physical sorting behavior.

## Validation

- [x] `npm run build`
- [x] `npm test`

## Current Limits

- `sortRange()` currently throws if worksheet comments exist inside the sortable data area.
- `sortRange()` throws when overlapping metadata ranges would need partial rewrites or would become non-contiguous after row permutation.
- Color filters use OOXML-oriented metadata: `kind: "color"`, `dxfId`, and `cellColor`.

## Non-goals

- Filter picker search UI.
- Distinct value counting for UI display.
- `Auto Apply` interaction logic.
- Frontend popup, header highlight, and editor presentation logic.

## Delivered Slices

- [x] Structured worksheet and table filter APIs.
- [x] Non-destructive `ref` rewrites and column-local filter updates.
- [x] Advanced typed filter kinds and structured CLI commands.
- [x] Physical `sortRange(range, options)` support with metadata-safe row movement.
