# fastxlsx

[中文 README](README.zh.md)

A prototype XLSX reader/writer built around a "lossless first" principle.

The goal is not to map the entire Excel object model into a huge JS object graph first.
The first baseline is simpler and stricter:

`read(xlsx) -> write(xlsx)`

After that roundtrip, the extracted package parts should stay byte-for-byte identical unless a part was intentionally edited.

Once that baseline holds, styles, themes, comments, relationship files, and unknown extension nodes are preserved naturally.
Then higher-level APIs can be added on top with much lower risk.

## Install

```bash
npm i fastxlsx
```

## CLI Usage

For published npm usage, run the installed binary or invoke it directly with `npx`:

```bash
npx fastxlsx create path/to/new.xlsx --sheet Sheet1
npx fastxlsx inspect path/to/file.xlsx
npx fastxlsx move-sheet path/to/file.xlsx --sheet Archive --index 0 --in-place
npx fastxlsx get path/to/file.xlsx --sheet Sheet1 --cell B2
```

If the package is already installed in a project, use the exposed binary:

```bash
npx fastxlsx inspect path/to/file.xlsx
```

For repository development only, keep using the local source runner:

```bash
npm run cli -- create path/to/new.xlsx --sheet Sheet1
npm run cli -- inspect path/to/file.xlsx
npm run cli -- move-sheet path/to/file.xlsx --sheet Archive --index 0 --in-place
```

## Library Usage

Install the package and import it directly in your project:

```bash
npm i fastxlsx
```

```ts
import { Workbook } from "fastxlsx";
```

## Quick Start

```ts
import { Workbook } from "fastxlsx";

const workbook = Workbook.create("Sheet1");
workbook.getSheet("Sheet1").setCell("A1", "Hello");
await workbook.save("new.xlsx");
```

Create multiple sheets with initial records:

```ts
const workbook = Workbook.create({
  activeSheet: "Data",
  author: "fastxlsx",
  sheets: [
    { name: "Config", headers: ["Key", "Value"] },
    { name: "Data", records: [{ id: 1001, name: "Alpha" }] },
  ],
});
```

Import and export records as JSON / CSV:

```ts
const sheet = workbook.getSheet("Data");

sheet.fromJson([{ id: 1001, name: "Alpha" }]);
const records = sheet.toJson();
const csv = sheet.toCsv();
sheet.fromCsv(csv);
```

Create from scratch, then edit an existing workbook:

```ts
import { Workbook } from "fastxlsx";

const workbook = await Workbook.open("input.xlsx");
const sheet = workbook.tryGetSheet("Sheet1");

if (!sheet) {
  throw new Error("Sheet1 not found");
}

console.log(sheet.getCell("B2"));
console.log(sheet.getDisplayValue("B2"));
console.log(sheet.cell("B2").text, sheet.cell("B2").error);

sheet.batch((currentSheet) => {
  currentSheet.setCell("A1", "Hello");
  currentSheet.setCell("B2", 42);
  currentSheet.setFormula("C2", "SUM(B2,8)", { cachedValue: 50 });
});

await workbook.save("output.xlsx");
```

## Common Tasks

Open from memory instead of disk:

```ts
const workbook = Workbook.fromUint8Array(xlsxBytes);
const sameWorkbook = Workbook.fromArrayBuffer(arrayBuffer);
const nextBytes = workbook.toUint8Array();
```

Safely discover sheets before reading:

```ts
console.log(workbook.getSheetNames());
console.log(workbook.hasSheet("Config"));

const configSheet = workbook.tryGetSheet("Config");
if (configSheet) {
  console.log(configSheet.getRecords());
}
```

Read user-facing values and Excel errors:

```ts
const cell = workbook.getSheet("Sheet1").cell("C5");

console.log(cell.value);
console.log(cell.text);
console.log(cell.error);
```

Batch multiple writes together:

```ts
workbook.batch((currentWorkbook) => {
  currentWorkbook.getSheet("Sheet1").setCell("A1", "left");
  currentWorkbook.getSheet("Sheet2").setCell("A1", "right");
});
```

On blank sheets, `addRecord()`, `addRecords()`, `setRecord()`, and `setRecords()` initialize the header row from record keys automatically.

Build an export-ready template with layout, comments, and print settings:

```ts
const workbook = Workbook.create({
  activeSheet: "Data",
  sheets: [
    {
      name: "Data",
      headers: ["id", "name", "score"],
      records: [{ id: 1001, name: "Alpha", score: 98 }],
      columnWidths: { A: 12, B: 24, C: 12 },
      rowHeights: { "1": 24 },
      frozenPane: { columnCount: 1, rowCount: 1 },
      printArea: "A1:C20",
      printTitles: { rows: "1:1" },
      comments: [{ address: "C2", author: "fastxlsx", text: "Final score" }],
      headerStyle: {
        applyAlignment: true,
        alignment: { horizontal: "center" },
      },
    },
  ],
});
```

Sync records by key instead of replacing the whole sheet:

```ts
const sheet = workbook.getSheet("Data");

sheet.syncRecords(
  [
    { id: 1002, name: "Beta", score: 91 },
    { id: 1003, name: "Gamma", score: 87 },
  ],
  { keyField: "id" },
);
```

Use the higher-level workbook helpers for config and table sheets:

```ts
workbook.createConfigSheet("Config", {
  records: [
    { Key: "timeout", Value: "30" },
    { Key: "region", Value: "cn" },
  ],
});

workbook.createTableSheet("Rewards", {
  records: [
    { id: 1, item: "Gold", amount: 100 },
    { id: 2, item: "Gem", amount: 5 },
  ],
});
```

Use workflow CLI commands for import/export and comments:

```bash
npm run cli -- workbook active set out.xlsx --sheet Summary --in-place
npm run cli -- workbook visibility set out.xlsx --sheet Archive --visibility hidden --in-place
npm run cli -- workbook defined-name set out.xlsx --name Scores --value 'Summary!$A$1:$B$10' --in-place
npm run cli -- sheet import input.xlsx --sheet Data --format json --from rows.json --output out.xlsx
npm run cli -- sheet export out.xlsx --sheet Data --format csv --output rows.csv
npm run cli -- sheet records append out.xlsx --sheet Data --records '[{"id":1003,"name":"Gamma"}]' --in-place
npm run cli -- sheet records upsert out.xlsx --sheet Data --key-field id --record '{"id":1002,"name":"Beta"}' --in-place
npm run cli -- sheet hyperlink set out.xlsx --sheet Data --cell A2 --target https://example.com --text "Open" --in-place
npm run cli -- sheet filter set out.xlsx --sheet Data --range A1:C20 --in-place
npm run cli -- sheet selection set out.xlsx --sheet Data --active-cell C3 --range C3:D4 --in-place
npm run cli -- sheet validation set out.xlsx --sheet Data --range B2:B20 --type whole --operator between --formula1 1 --formula2 10 --in-place
npm run cli -- sheet merge add out.xlsx --sheet Data --range A1:B2 --in-place
npm run cli -- sheet protection set out.xlsx --sheet Data --sort --auto-filter --in-place
npm run cli -- sheet comment set out.xlsx --sheet Data --cell C2 --text "Final score" --in-place
```

## Design

The library is split into two layers:

1. `Lossless package layer`
   - Treat an `.xlsx` file as a zip package.
   - Keep every entry as raw bytes first.
   - Write untouched entries back exactly as they were, without re-serialization.

2. `Editable workbook layer`
   - Apply targeted XML patches only to the parts that actually need to change.
   - The current implementation already covers workbook metadata, cell values, formulas, styles, row and column edits, structured record helpers, tables, hyperlinks, filters, frozen panes, selections, data validation, and defined names.
   - Style-related `s="..."` attributes are preserved, so styles are not lost when values change.

## Why This Works For Style Preservation

Most style loss is not caused by failing to parse `styles.xml`.
It usually happens because the workbook is regenerated wholesale on write, which tends to break things like:

- unknown nodes
- attribute ordering
- namespaces and extension markers
- relationship file ordering
- coupling between shared strings, worksheets, and styles

The lossless-first direction flips that approach:

- preserve every package part first
- edit only the parts that must change
- write untouched parts back exactly as-is

That makes it much easier to satisfy a strict "roundtrip without diffs" requirement.

## Current API

- `Workbook.open(path)`
- `Workbook.create(sheetName?)`
- `Workbook.create(options)`
- `Workbook.fromEntries(entries)`
- `Workbook.fromUint8Array(data)`
- `Workbook.fromArrayBuffer(data)`
- `workbook.listEntries()`
- `workbook.toUint8Array()`
- `workbook.getSheets()`
- `workbook.getSheetNames()`
- `workbook.getSheet(name)`
- `workbook.hasSheet(name)`
- `workbook.tryGetSheet(name)`
- `workbook.getActiveSheet()`
- `workbook.batch(applyChanges)`
- `workbook.getNumberFormat(numFmtId)`
- `workbook.updateNumberFormat(numFmtId, formatCode)`
- `workbook.cloneNumberFormat(numFmtId, formatCode?)`
- `workbook.getBorder(borderId)`
- `workbook.updateBorder(borderId, patch)`
- `workbook.cloneBorder(borderId, patch?)`
- `workbook.getFill(fillId)`
- `workbook.updateFill(fillId, patch)`
- `workbook.cloneFill(fillId, patch?)`
- `workbook.getFont(fontId)`
- `workbook.updateFont(fontId, patch)`
- `workbook.cloneFont(fontId, patch?)`
- `workbook.getStyle(styleId)`
- `workbook.updateStyle(styleId, patch)`
- `workbook.cloneStyle(styleId, patch?)`
- `workbook.getSheetVisibility(name)`
- `workbook.getDefinedNames()`
- `workbook.getDefinedName(name, scope?)`
- `workbook.setDefinedName(name, value, options?)`
- `workbook.deleteDefinedName(name, scope?)`
- `workbook.renameSheet(currentName, nextName)`
- `workbook.moveSheet(name, targetIndex)`
- `workbook.addSheet(name)`
- `workbook.deleteSheet(name)`
- `workbook.setSheetVisibility(name, visibility)`
- `workbook.setActiveSheet(name)`
- `sheet.cell(address)`
- `sheet.cell(rowNumber, column)`
- `sheet.rename(name)`
- `sheet.getCell(address)`
- `sheet.getCell(rowNumber, column)`
- `sheet.getDisplayValue(address)`
- `sheet.getDisplayValue(rowNumber, column)`
- `sheet.getAlignment(address)`
- `sheet.getAlignment(rowNumber, column)`
- `sheet.getNumberFormat(address)`
- `sheet.getNumberFormat(rowNumber, column)`
- `sheet.getBorder(address)`
- `sheet.getBorder(rowNumber, column)`
- `sheet.getBackgroundColor(address)`
- `sheet.getBackgroundColor(rowNumber, column)`
- `sheet.getFill(address)`
- `sheet.getFill(rowNumber, column)`
- `sheet.getFont(address)`
- `sheet.getFont(rowNumber, column)`
- `sheet.getStyleId(address)`
- `sheet.getStyleId(rowNumber, column)`
- `sheet.getStyle(address)`
- `sheet.getStyle(rowNumber, column)`
- `sheet.getRowStyleId(rowNumber)`
- `sheet.getRowStyle(rowNumber)`
- `sheet.getRowHidden(rowNumber)`
- `sheet.getRowHeight(rowNumber)`
- `sheet.getColumnStyleId(column)`
- `sheet.getColumnStyle(column)`
- `sheet.getColumnHidden(column)`
- `sheet.getColumnWidth(column)`
- `sheet.copyStyle(sourceAddress, targetAddress)`
- `sheet.copyStyle(sourceRowNumber, sourceColumn, targetRowNumber, targetColumn)`
- `sheet.setAlignment(address, patch)`
- `sheet.setAlignment(rowNumber, column, patch)`
- `sheet.setNumberFormat(address, formatCode)`
- `sheet.setNumberFormat(rowNumber, column, formatCode)`
- `sheet.setBorder(address, patch)`
- `sheet.setBorder(rowNumber, column, patch)`
- `sheet.setBackgroundColor(address, color)`
- `sheet.setBackgroundColor(rowNumber, column, color)`
- `sheet.setFill(address, patch)`
- `sheet.setFill(rowNumber, column, patch)`
- `sheet.setFont(address, patch)`
- `sheet.setFont(rowNumber, column, patch)`
- `sheet.setStyle(address, patch)`
- `sheet.setStyle(rowNumber, column, patch)`
- `sheet.setRowStyle(rowNumber, patch)`
- `sheet.setColumnStyle(column, patch)`
- `sheet.cloneStyle(address, patch?)`
- `sheet.cloneStyle(rowNumber, column, patch?)`
- `sheet.cloneRowStyle(rowNumber, patch?)`
- `sheet.cloneColumnStyle(column, patch?)`
- `sheet.getCellEntries()`
- `sheet.getPhysicalCellEntries()`
- `sheet.iterCellEntries()`
- `sheet.iterPhysicalCellEntries()`
- `sheet.rowCount`
- `sheet.columnCount`
- `sheet.getHeaders(headerRowNumber?)`
- `sheet.getRecord(rowNumber, headerRowNumber?)`
- `sheet.getRecordBy(field, value, headerRowNumber?)`
- `sheet.getRecords(headerRowNumber?)`
- `sheet.toJson(headerRowNumber?)`
- `sheet.toCsv(headerRowNumber?)`
- `sheet.getColumn(column)`
- `sheet.getColumnEntries(column)`
- `sheet.getPhysicalColumnEntries(column)`
- `sheet.getRow(rowNumber)`
- `sheet.getRowEntries(rowNumber)`
- `sheet.getPhysicalRowEntries(rowNumber)`
- `sheet.getRange(range)`
- `sheet.getRangeRef()`
- `sheet.getPhysicalRangeRef()`
- `sheet.getMergedRanges()`
- `sheet.getAutoFilter()`
- `sheet.getFreezePane()`
- `sheet.getSelection()`
- `sheet.getDataValidations()`
- `sheet.getDataValidation(range)`
- `sheet.getTables()`
- `sheet.getComments()`
- `sheet.getComment(address)`
- `sheet.getHyperlink(address)`
- `sheet.hyperlink(address)`
- `sheet.getHyperlinks()`
- `sheet.getPrintArea()`
- `sheet.getPrintTitles()`
- `sheet.getProtection()`
- `sheet.addTable(range, options?)`
- `sheet.removeTable(name)`
- `sheet.setComment(address, text, options?)`
- `sheet.clearComments()`
- `sheet.removeComment(address)`
- `sheet.protect(options?)`
- `sheet.unprotect()`
- `sheet.setHyperlink(address, target, options?)`
- `sheet.clearHyperlinks()`
- `sheet.removeHyperlink(address)`
- `sheet.setAutoFilter(range)`
- `sheet.clearAutoFilter()`
- `sheet.setPrintArea(range)`
- `sheet.setPrintTitles(options)`
- `sheet.freezePane(columnCount, rowCount?)`
- `sheet.unfreezePane()`
- `sheet.setSelection(activeCell, range?)`
- `sheet.clearSelection()`
- `sheet.removeAutoFilter()`
- `sheet.setDataValidation(range, options?)`
- `sheet.clearDataValidations()`
- `sheet.removeDataValidation(range)`
- `sheet.setCell(address, value)`
- `sheet.setCell(rowNumber, column, value)`
- `sheet.setStyleId(address, styleId)`
- `sheet.setStyleId(rowNumber, column, styleId)`
- `sheet.setRowStyleId(rowNumber, styleId)`
- `sheet.setColumnStyleId(column, styleId)`
- `sheet.setRowHidden(rowNumber, hidden)`
- `sheet.setRowHeight(rowNumber, height)`
- `sheet.setColumnHidden(column, hidden)`
- `sheet.setColumnWidth(column, width)`
- `sheet.copyStyle(sourceAddress, targetAddress)`
- `sheet.copyStyle(sourceRowNumber, sourceColumn, targetRowNumber, targetColumn)`
- `sheet.deleteCell(address)`
- `sheet.deleteCell(rowNumber, column)`
- `sheet.deleteRow(row, count?)`
- `sheet.deleteColumn(column, count?)`
- `sheet.insertRow(row, count?)`
- `sheet.insertColumn(column, count?)`
- `sheet.setHeaders(headers, headerRowNumber?, startColumn?)`
- `sheet.setRecord(rowNumber, record, headerRowNumber?)`
- `sheet.setRecords(records, headerRowNumber?)`
- `sheet.fromJson(records, headerRowNumber?)`
- `sheet.fromCsv(csv, headerRowNumber?)`
- `sheet.upsertRecord(field, record, headerRowNumber?)`
- `sheet.deleteRecord(rowNumber, headerRowNumber?)`
- `sheet.deleteRecords(rowNumbers, headerRowNumber?)`
- `sheet.deleteRecordBy(field, value, headerRowNumber?)`
- `sheet.addRecord(record, headerRowNumber?)`
- `sheet.addRecords(records, headerRowNumber?)`
- `sheet.appendRow(values, startColumn?)`
- `sheet.appendRows(rows, startColumn?)`
- `sheet.setColumn(column, values, startRow?)`
- `sheet.setRow(rowNumber, values, startColumn?)`
- `sheet.setRange(startAddress, values)`
- `sheet.setRangeStyle(range, patch)`
- `sheet.setRangeNumberFormat(range, formatCode)`
- `sheet.setRangeBackgroundColor(range, color)`
- `sheet.copyRangeStyle(sourceRange, targetRange)`
- `sheet.addMergedRange(range)`
- `sheet.clearMergedRanges()`
- `sheet.removeMergedRange(range)`
- `sheet.getFormula(address)`
- `sheet.getFormula(rowNumber, column)`
- `sheet.setFormula(address, formula, options?)`
- `sheet.setFormula(rowNumber, column, formula, options?)`
- `sheet.batch(applyChanges)`
- `workbook.save(path)`

Example:

```ts
import { Workbook } from "fastxlsx";

const workbook = await Workbook.open("input.xlsx");
const sheet = workbook.getSheet("Sheet1");
const scoreCell = sheet.cell("B2");
const scoreValue = sheet.getCell(2, 2);
const scoreText = sheet.getDisplayValue(2, 2);
const scoreStyleId = sheet.getStyleId(2, 2);
const headerRowStyleId = sheet.getRowStyleId(1);
const scoreColumnStyleId = sheet.getColumnStyleId(2);
const detailSheet = workbook.addSheet("Detail");
const activeSheet = workbook.getActiveSheet();

console.log(sheet.getTables());
console.log(sheet.getHyperlinks());
console.log(sheet.rowCount, sheet.columnCount);
console.log(sheet.getFreezePane(), sheet.getSelection(), activeSheet.name);
console.log(scoreCell.text, scoreCell.error, scoreText);

workbook.batch((currentWorkbook) => {
  currentWorkbook.setDefinedName("Scores", "Summary!$A$1:$B$10");
  currentWorkbook.setDefinedName("LocalScore", "$B$2", { scope: "Summary" });
  currentWorkbook.renameSheet("Sheet1", "Summary");
  currentWorkbook.moveSheet("Summary", 0);
  currentWorkbook.setActiveSheet("Summary");
  currentWorkbook.setSheetVisibility("Summary", "hidden");
  detailSheet.rename("Detail 2026");
  sheet.addTable("A1:B10", { name: "Scores" });
  sheet.setHyperlink("A1", "https://example.com", { text: "Hello", tooltip: "Open link" });
  sheet.setHyperlink("B2", "#Summary!A1");
  sheet.setAutoFilter("A1:F20");
  sheet.freezePane(1, 1);
  sheet.setSelection("B2", "B2:C4");
  sheet.setDataValidation("B2:B100", { type: "whole", operator: "between", formula1: "0", formula2: "100" });
  sheet.setCell(3, 2, 98);
  sheet.setStyleId(3, 2, scoreStyleId);
  sheet.setRowStyleId(1, headerRowStyleId);
  sheet.setColumnStyleId(2, scoreColumnStyleId);
  sheet.copyStyle("B2", "C2");
  sheet.setCell("A1", "Hello");
  sheet.deleteRow(8);
  sheet.deleteColumn("G");
  sheet.insertRow(2);
  sheet.setHeaders(["Name", "Score"]);
  sheet.insertColumn("B");
  sheet.setRecord(2, { Name: "Alice", Score: 98 });
  sheet.setRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);
  sheet.deleteRecord(4);
  sheet.deleteRecords([6, 7]);
  sheet.addRecord({ Name: "Alice", Score: 98 });
  sheet.addRecords([
    { Name: "Bob", Score: 87 },
    { Name: "Cara", Score: 91 },
  ]);
  sheet.appendRow(["tail", 1]);
  sheet.appendRows([
    ["tail-2", 2],
    ["tail-3", 3],
  ]);
  sheet.setColumn("F", ["Q1", "Q2"], 2);
  sheet.setRow(5, ["Name", "Score"], 2);
  sheet.setRange("B2", [
    [1, 2],
    [3, 4],
  ]);
  sheet.addMergedRange("D1:E1");
  sheet.setFormula("B1", "SUM(1,2)", { cachedValue: 3 });
  sheet.setFormula(4, 3, "SUM(A4:B4)", { cachedValue: 12 });
  sheet.removeHyperlink("B2");
  sheet.unfreezePane();
  sheet.removeAutoFilter();
  sheet.removeDataValidation("B2:B100");
  sheet.removeTable("Scores");
  detailSheet.setCell("A1", "created");
  currentWorkbook.setSheetVisibility("Summary", "visible");
  currentWorkbook.deleteDefinedName("LocalScore", "Summary");
  currentWorkbook.deleteSheet("Temp");
});

console.log(workbook.getSheetNames(), workbook.hasSheet("Summary"), workbook.tryGetSheet("Temp"));
console.log(workbook.getDefinedNames(), workbook.getDefinedName("LocalScore", "Summary"));
console.log(scoreValue, scoreCell.value, scoreCell.styleId, scoreCell.formula);

await workbook.save("output.xlsx");
```

Notes:

- On first read/write access, a sheet scans `sheetData` once and builds indexes for rows and cells.
- `Workbook.fromUint8Array()`, `Workbook.fromArrayBuffer()`, and `workbook.toUint8Array()` cover in-memory workflows for browser, RPC, and upload pipelines.
- `workbook.getSheetNames()`, `hasSheet()`, and `tryGetSheet()` are the safe lookup APIs; `getSheet()` remains the strict variant that throws when a sheet is missing.
- `workbook.batch()` and `sheet.batch()` group related edits and flush pending sheet-dimension normalization once at the end of the outer batch.
- `sheet.cell(address)` returns a reusable `Cell` handle whose parsed value/formula/style-index state is cached by sheet revision. It now also exposes `cell.text`, `cell.error`, `cell.style`, `cell.alignment`, `cell.font`, `cell.fill`, `cell.backgroundColor`, `cell.border`, `cell.numberFormat`, `cell.setStyle(patch)`, `cell.setAlignment(patch)`, `cell.setFont(patch)`, `cell.setFill(patch)`, `cell.setBackgroundColor(color)`, `cell.setBorder(patch)`, `cell.setNumberFormat(formatCode)`, and `cell.cloneStyle(patch?)`.
- `sheet.cell()`, `getCell()`, `setCell()`, `getFormula()`, and `setFormula()` now support both `A1` addresses and `(rowNumber, column)` calls. Row and column indexes are 1-based.
- `sheet.getDisplayValue()` and `cell.text` are best-effort user-facing readers: they preserve Excel error text, stringify booleans as `TRUE` / `FALSE`, and otherwise stringify the cached value directly.
- Cells with cached Excel errors now expose `cell.error` / `snapshot.error` metadata. Pure error cells use `type: "error"`, while formula cells with error caches keep `type: "formula"` and still expose the structured error payload.
- Later `getCell()` and `getFormula()` calls use those indexes directly instead of running a full string match on every read.
- `sheet.rowCount` and `sheet.columnCount` mean the logical used bounds based on cells that currently have a value or formula. Pure blank placeholder `<c>` nodes and blank-only physical rows do not extend the used range. Empty sheets return `0`.
- `sheet.getCellEntries()`, `iterCellEntries()`, `getRowEntries()`, `getColumnEntries()`, and `getRangeRef()` are the default logical read APIs. They skip blank placeholder `<c>` nodes that have neither a value nor a formula, and they follow the logical used bounds.
- `sheet.getPhysicalCellEntries()`, `iterPhysicalCellEntries()`, `getPhysicalRowEntries()`, `getPhysicalColumnEntries()`, and `getPhysicalRangeRef()` expose exact physical worksheet `<c>` node boundaries when you need to inspect low-level package structure.
- `sheet.deleteCell()` removes the worksheet `<c>` node entirely; if you want to keep a styled placeholder but clear the value, continue using `setCell(..., null)`.
- `workbook.getStyle()` reads `cellXfs` definitions from `styles.xml`, `workbook.updateStyle()` patches an existing `<xf>` in place, and `workbook.cloneStyle()` appends a new `<xf>` derived from an existing one and returns the new style id.
- `workbook.getFont()`, `updateFont()`, and `cloneFont()` work directly on the `<fonts>` section in `styles.xml`, which is useful when you want to manage reusable `fontId` values explicitly.
- `workbook.getFill()`, `updateFill()`, and `cloneFill()` work directly on the `<fills>` section in `styles.xml`, which is useful when you want to manage reusable `fillId` values explicitly.
- `workbook.getBorder()`, `updateBorder()`, and `cloneBorder()` work directly on the `<borders>` section in `styles.xml`, which is useful when you want to manage reusable `borderId` values explicitly.
- `workbook.getNumberFormat()`, `updateNumberFormat()`, and `cloneNumberFormat()` work directly on `<numFmt>` entries in `styles.xml`; new custom formats are allocated from `numFmtId` `164` upward.
- `sheet.getFont()` resolves the font definition currently used by the cell.
- `sheet.setFont()` and `cell.setFont()` clone the current `fontId`, then clone the current `styleId` with that new font attached, so only the targeted cell changes and other cells sharing the old font/style stay untouched.
- `sheet.getFill()` resolves the fill definition currently used by the cell.
- `sheet.setFill()` and `cell.setFill()` clone the current `fillId`, then clone the current `styleId` with that new fill attached, so only the targeted cell changes and other cells sharing the old fill/style stay untouched.
- `sheet.getBackgroundColor()` and `cell.backgroundColor` are fill-based convenience readers. They return an ARGB value only when the cell uses a `solid` fill with `fgColor.rgb`.
- `sheet.setBackgroundColor()` and `cell.setBackgroundColor()` are fill-based convenience writers: passing a color writes `solid + fgColor.rgb`, and passing `null` resets the fill to `none`.
- `sheet.getBorder()` resolves the border definition currently used by the cell.
- `sheet.setBorder()` and `cell.setBorder()` clone the current `borderId`, then clone the current `styleId` with that new border attached, so only the targeted cell changes and other cells sharing the old border/style stay untouched.
- `sheet.getNumberFormat()` resolves the number-format definition currently used by the cell, including common builtin format codes.
- `sheet.setNumberFormat()` and `cell.setNumberFormat()` reuse or create the target `numFmtId`, then clone the current `styleId` with that new number format attached, so only the targeted cell changes.
- `sheet.getAlignment()` resolves the alignment definition currently used by the cell.
- `sheet.setAlignment()` and `cell.setAlignment()` clone the current `styleId`, patch only the alignment / `applyAlignment` layer, and apply the new style back to the target cell; passing `null` removes the `<alignment>` node.
- `sheet.getStyle()` resolves the cell's current style definition; when the cell has no explicit `s="..."`, it falls back to the default style `0`.
- `sheet.setStyle()` clones the current cell style into a new `styleId`, writes that new definition into `styles.xml`, and applies it back to the same cell so other cells sharing the old style id stay untouched.
- `sheet.cloneStyle()` clones the current cell style, writes the new definition into `styles.xml`, applies it back to the same cell, and supports both `A1` and `(rowNumber, column)` calls.
- `sheet.getStyleId()` and `setStyleId()` still only read and write the cell-level `s="..."` style index itself.
- `sheet.getRowStyleId()` and `setRowStyleId()` currently read and write the row-level `<row s="..." customFormat="1">` style index; that layer still does not modify `styles.xml`.
- `sheet.getRowStyle()` and `sheet.getColumnStyle()` resolve the currently assigned row/column style definition; they return `null` when no explicit row/column style is present.
- `sheet.setRowStyle()` and `sheet.setColumnStyle()` are convenience entry points for row/column style edits. They reuse the same clone semantics: derive a new `styleId` from the current row/column style and apply it immediately.
- `sheet.cloneRowStyle()` and `sheet.cloneColumnStyle()` clone the current row/column style into a new `styleId` and apply it immediately; when no explicit row/column style exists yet, they clone from the default style `0`.
- `sheet.getColumnStyleId()` and `setColumnStyleId()` currently read and write the column-level `<cols><col ... style="..."/>` style index, and those ranges are shifted during column insert/delete operations.
- `sheet.copyStyle()` currently copies the source cell's `styleId` onto the target cell without changing the target cell's value or formula; both address and `(rowNumber, column)` calls are supported.
- `sheet.getFreezePane()`, `freezePane()`, and `unfreezePane()` currently manage worksheet `sheetViews/sheetView/pane`; `topLeftCell` keeps tracking row and column insert/delete operations.
- `sheet.getSelection()` and `setSelection()` currently read and write worksheet `sheetViews/sheetView/selection`; when a frozen pane exists, they target the selection for the current active pane.
- Outside batches, each write rebuilds the sheet index immediately so later reads always see the latest content. Inside `sheet.batch()` / `workbook.batch()`, the final dimension sync is deferred until the batch flushes.
- Worksheet edits keep `<dimension ref="...">` in sync so used-range metadata does not go stale.
- `deleteRow()` and `deleteColumn()` currently update cell coordinates, formulas, merged ranges, worksheet `dimension`, common `ref` and `sqref` attributes, `definedNames`, and explicit formulas in other sheets that reference the edited sheet.
- `insertRow()` currently updates cell coordinates, formulas, merged ranges, worksheet `dimension`, common `ref` and `sqref` attributes, `definedNames`, and explicit formulas in other sheets that reference the edited sheet.
- `insertColumn()` currently updates cell coordinates, formulas, merged ranges, worksheet `dimension`, common `ref` and `sqref` attributes, `definedNames`, and explicit formulas in other sheets that reference the edited sheet.
- `sheet.getTables()` currently reads existing table names, display names, ranges, and part paths.
- `sheet.getHyperlinks()` currently reads internal and external hyperlinks from the sheet; external link targets are resolved through the sheet relationships part.
- `sheet.getAutoFilter()`, `sheet.setAutoFilter()`, and `sheet.removeAutoFilter()` currently manage the worksheet-level `autoFilter`; removing it also clears the top-level `sortState`.
- `sheet.getDataValidations()`, `sheet.setDataValidation()`, and `sheet.removeDataValidation()` currently manage worksheet-level `dataValidations`, including common attributes plus `formula1` and `formula2`, and keep `sqref` updated during row and column edits.
- `sheet.addTable()` currently creates the basic table part, sheet relationship, `[Content_Types].xml` override, and table XML. Column names default to the first row in the range, and blank names fall back to `ColumnN`.
- `sheet.removeTable()` currently removes the current sheet's `tableParts`, sheet relationship, table XML, and matching content type override.
- Existing linked tables keep their own `ref` and `autoFilter` updated during row and column insert/delete operations. If a table becomes empty, its `tableParts` entry is removed from the sheet.
- `sheet.setHyperlink()` and `sheet.removeHyperlink()` currently manage worksheet `<hyperlinks>` plus the matching sheet relationship for external links. Internal targets use a format like `#Sheet1!A1`.
- `workbook.getDefinedNames()`, `getDefinedName()`, `setDefinedName()`, and `deleteDefinedName()` currently support both global and local defined names.
- `workbook.getSheetVisibility()` and `setSheetVisibility()` currently support `visible`, `hidden`, and `veryHidden`, and prevent hiding the last visible sheet in the workbook.
- `workbook.getActiveSheet()` and `setActiveSheet()` currently read and write `workbookView.activeTab`; if the workbook does not yet contain `bookViews`, they are created automatically, and hidden sheets cannot be activated.
- `workbook.renameSheet()` and `sheet.rename()` currently update sheet names, explicit formula references in other sheets, `definedNames`, internal hyperlink locations, and document properties.
- `workbook.moveSheet()` currently uses a 0-based `targetIndex` and keeps workbook `<sheets>` order, worksheet order in `docProps/app.xml`, local defined-name `localSheetId` values, and `workbookView.activeTab` aligned.
- `workbook.addSheet()` and `workbook.deleteSheet()` currently maintain `workbook.xml`, workbook rels, and `[Content_Types].xml`, and adjust remaining formulas and `definedNames` when a sheet is deleted.

## Benchmarking

The repo now includes a sanitized large benchmark workbook at [`res/monster.xlsx`](res/monster.xlsx), intended for repeatable performance regression checks.

Common commands:

- `npm run bench:monster`
  - Run a 3-iteration benchmark on `res/monster.xlsx`
- `npm run bench:check`
  - Run a 5-iteration benchmark on `res/monster.xlsx` and validate the non-null count plus the configured read/write thresholds from `benchmarks/monster-baseline.json`
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5`
  - Run the benchmark with a custom file path and iteration count; the JSON output includes dense traversal (`result`), sparse traversal (`sparseResult`), a batch write scenario (`writeResult`), and per-sheet amplification stats
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5 --check benchmarks/monster-baseline.json`
  - Run the regression check against any benchmark file; the process exits non-zero when the workbook count or configured read/write timing thresholds are exceeded

The batch write benchmark intentionally targets the worksheet with the most physical cell nodes and overwrites up to 30 existing `A` column cells inside one `sheet.batch(...)` call. That keeps the regression check focused on the hot path for repeated in-memory cell edits on large sheets.

## Current Limits

- The zip backend now uses pure JS via `fflate`, so it no longer depends on system `python3` or `zip`.
- The full zip package and all entries are still loaded into memory today, so peak memory usage for very large files can still be improved.
- String writes use `inlineStr` to avoid rebuilding `sharedStrings.xml` for simple value updates.
- APIs for merged comments, rich text, images, and similar parts are still missing.
- XML writes are implemented as local patches, not as a full OOXML object model.

## Development

```bash
npm run build
npm test
npm run pack:check
npm run validate:task
```

Where:

- `npm test` runs the TypeScript tests through `tsx`
- `npm run pack:check` verifies that `npm pack --dry-run` contains exactly the expected build outputs and no stale files from older builds
- `npm run validate:task` runs the TypeScript validation script through `tsx`
- `npm run build` cleans `dist` first, then produces a fresh build

The automated checks currently cover:

1. Untouched roundtrip stability, including byte-for-byte part preservation.
2. Workbook editing flows across cells, formulas, styles, rows, columns, tables, hyperlinks, filters, frozen panes, selections, data validations, defined names, and CLI commands.
3. Package metadata plus dry-run tarball validation, so the published npm package stays aligned with the current source tree.

## Real File Validation

[`res/task.xlsx`](res/task.xlsx) in the repository is a useful regression sample.

```bash
npm run validate:task
```

To validate any other file:

```bash
npm run validate:roundtrip -- path/to/file.xlsx
```
