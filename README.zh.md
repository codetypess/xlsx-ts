# fastxlsx

[English README](README.md)

一个以“无损优先”为核心的 XLSX 读写库原型。

目标不是先把 Excel 全模型映射成庞大的 JS 对象，而是先保证这一条底线：

`read(xlsx) -> write(xlsx)` 之后，解压出来的各个部件文件内容保持一致。

这条底线一旦成立，样式、主题、批注、关系文件、未知扩展节点都能被天然保住。后续再往上叠加单元格、公式、批注、图片等 API，风险会低很多。

## 安装

```bash
npm i fastxlsx
```

## CLI 使用方式

如果是已经发布到 npm 的包，可以直接通过 `npx` 调用，或者使用已经安装到项目里的命令：

```bash
npx fastxlsx create path/to/new.xlsx --sheet Sheet1
npx fastxlsx inspect path/to/file.xlsx
npx fastxlsx move-sheet path/to/file.xlsx --sheet Archive --index 0 --in-place
npx fastxlsx get path/to/file.xlsx --sheet Sheet1 --cell B2
```

如果包已经安装到项目里，可以直接使用暴露出来的命令：

```bash
npx fastxlsx inspect path/to/file.xlsx
```

只有在这个仓库内开发时，才继续使用本地脚本入口：

```bash
npm run cli -- create path/to/new.xlsx --sheet Sheet1
npm run cli -- inspect path/to/file.xlsx
npm run cli -- move-sheet path/to/file.xlsx --sheet Archive --index 0 --in-place
```

## 作为库使用

先安装包，再在项目里直接导入：

```bash
npm i fastxlsx
```

```ts
import { Workbook } from "fastxlsx";
```

## 快速开始

```ts
import { Workbook } from "fastxlsx";

const workbook = Workbook.create("Sheet1");
workbook.getSheet("Sheet1").setCell("A1", "Hello");
await workbook.save("new.xlsx");
```

一次性创建多张 sheet，并写入初始记录：

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

以 JSON / CSV 导入导出记录：

```ts
const sheet = workbook.getSheet("Data");

sheet.fromJson([{ id: 1001, name: "Alpha" }]);
const records = sheet.toJson();
const csv = sheet.toCsv();
sheet.fromCsv(csv);
```

从零创建后，再编辑已有 workbook：

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

## 常见任务

从内存字节而不是磁盘打开：

```ts
const workbook = Workbook.fromUint8Array(xlsxBytes);
const sameWorkbook = Workbook.fromArrayBuffer(arrayBuffer);
const nextBytes = workbook.toUint8Array();
```

先安全探测 sheet，再决定是否读取：

```ts
console.log(workbook.getSheetNames());
console.log(workbook.hasSheet("Config"));

const configSheet = workbook.tryGetSheet("Config");
if (configSheet) {
  console.log(configSheet.getRecords());
}
```

读取更贴近用户看到的值，以及 Excel 错误信息：

```ts
const cell = workbook.getSheet("Sheet1").cell("C5");

console.log(cell.value);
console.log(cell.text);
console.log(cell.error);
```

把多次改动放进一个批次：

```ts
workbook.batch((currentWorkbook) => {
  currentWorkbook.getSheet("Sheet1").setCell("A1", "left");
  currentWorkbook.getSheet("Sheet2").setCell("A1", "right");
});
```

对于空白 sheet，`addRecord()`、`addRecords()`、`setRecord()`、`setRecords()` 会自动根据 record key 初始化表头行。

对于已经命中的现有行，`setRecord()` 和 `updateRecordBy()` 会保留未提供的字段。
`replaceRecord()`、`upsertRecord()` 以及默认 keyed 模式下的 `syncRecords()` 会整体替换命中的行，并把未提供的字段清空。

创建一个带布局、注释和打印设置的导出模板：

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

如果只是改部分字段，用 patch 式更新；只有在命中行 payload 完整时，才使用 keyed sync：

```ts
const sheet = workbook.getSheet("Data");

sheet.updateRecordBy("id", 1002, { score: 91 });

sheet.syncRecords(
  [
    { id: 1002, name: "Beta", score: 91 },
    { id: 1003, name: "Gamma", score: 87 },
  ],
  { keyField: "id" },
);
```

当提供 `keyField` 且未显式指定 `mode` 时，`syncRecords()` 默认走 `upsert`。
如果要做部分更新，请使用 `importRecords(..., { mode: "update" })` 或 `updateRecordBy()`；只有在允许整行替换时才使用 `upsert`。

使用更高层的 workbook helper 创建配置表和数据表：

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

通过工作流 CLI 做导入导出和注释：

```bash
npm run cli -- workbook active set out.xlsx --sheet Summary --in-place
npm run cli -- workbook visibility set out.xlsx --sheet Archive --visibility hidden --in-place
npm run cli -- workbook defined-name set out.xlsx --name Scores --value 'Summary!$A$1:$B$10' --in-place
npm run cli -- sheet import input.xlsx --sheet Data --format json --from rows.json --output out.xlsx
npm run cli -- sheet export out.xlsx --sheet Data --format csv --output rows.csv
npm run cli -- sheet records append out.xlsx --sheet Data --records '[{"id":1003,"name":"Gamma"}]' --in-place
npm run cli -- sheet records update out.xlsx --sheet Data --key-field id --value 1002 --record '{"name":"Beta"}' --in-place
npm run cli -- sheet records upsert out.xlsx --sheet Data --key-field id --record '{"id":1004,"name":"Delta"}' --in-place
npm run cli -- sheet hyperlink set out.xlsx --sheet Data --cell A2 --target https://example.com --text "Open" --in-place
npm run cli -- sheet filter set out.xlsx --sheet Data --range A1:C20 --in-place
npm run cli -- sheet selection set out.xlsx --sheet Data --active-cell C3 --range C3:D4 --in-place
npm run cli -- sheet validation set out.xlsx --sheet Data --range B2:B20 --type whole --operator between --formula1 1 --formula2 10 --in-place
npm run cli -- sheet merge add out.xlsx --sheet Data --range A1:B2 --in-place
npm run cli -- sheet protection set out.xlsx --sheet Data --sort --auto-filter --in-place
npm run cli -- sheet comment set out.xlsx --sheet Data --cell C2 --text "Final score" --in-place
```

只想改传入字段时请使用 `update`。
只有在“缺失行需要插入、命中行允许整行替换”时才使用 `upsert`。

## 设计思路

仓库后续功能开发统一遵循 [docs/spec-driven-development.md](docs/spec-driven-development.md) 里的 SDD 规范。

库分成两层：

1. `Lossless package layer`
   - 把 xlsx 当成 zip 包处理。
   - 所有 entry 先按原始字节保存。
   - 未修改的 entry 永远原样写回，不做重新序列化。

2. `Editable workbook layer`
   - 只针对确实需要改动的 XML 部件做局部 patch。
   - 当前实现已经覆盖 workbook 元数据、单元格值、公式、样式、行列编辑、记录式读写、表格、超链接、筛选、冻结窗格、选区、数据验证和 defined names。
   - 样式依赖的 `s="..."` 属性保留不动，因此样式不会因为写值被丢掉。

## 为什么这条路线适合“保样式”

大多数样式丢失，根源都不是 `styles.xml` 不会读，而是“写回时整个工作簿被重新生成”，导致：

- 未知节点丢失
- 属性顺序变化
- namespace/扩展标记被清洗
- 关系文件被重排
- shared strings / worksheet / styles 之间的耦合被误改

无损优先的路线反过来做：

- 先完整保住 zip 内所有 part
- 再只改需要改的 part
- 没改的 part 一律原样回写

这样就更容易通过“解压后内容一致”的验收标准。

## 当前能力

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
- `sheet.replaceRecord(rowNumber, record, headerRowNumber?)`
- `sheet.setRecords(records, headerRowNumber?)`
- `sheet.fromJson(records, headerRowNumber?)`
- `sheet.fromCsv(csv, headerRowNumber?)`
- `sheet.importRecords(records, options?)`
- `sheet.exportRecords(options?)`
- `sheet.updateRecordBy(field, record, headerRowNumber?)`
- `sheet.updateRecordBy(field, value, record, headerRowNumber?)`
- `sheet.upsertRecord(field, record, headerRowNumber?)`
- `sheet.syncRecords(records, options?)`
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

示例：

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

说明：

- 同一张工作表首次读写时会扫描一次 `sheetData`，建立单元格与行的位置索引
- `Workbook.fromUint8Array()`、`Workbook.fromArrayBuffer()` 和 `workbook.toUint8Array()` 覆盖了内存态工作流，适合浏览器、RPC 和上传处理场景
- `workbook.getSheetNames()`、`hasSheet()`、`tryGetSheet()` 是更安全的查找接口；`getSheet()` 仍然保留“找不到就抛错”的严格语义
- `workbook.batch()` 和 `sheet.batch()` 可以把一组相关改动包成一个批次，等外层批次结束时再统一做 worksheet `dimension` 同步
- `sheet.cell(address)` 返回可复用的 `Cell` 句柄，值/公式/样式索引会按工作表 revision 缓存；现在也可以通过 `cell.text` / `cell.error` / `cell.style` / `cell.alignment` / `cell.font` / `cell.fill` / `cell.backgroundColor` / `cell.border` / `cell.numberFormat` 读取当前状态，并用 `cell.setStyle(patch)` / `cell.setAlignment(patch)` / `cell.setFont(patch)` / `cell.setFill(patch)` / `cell.setBackgroundColor(color)` / `cell.setBorder(patch)` / `cell.setNumberFormat(formatCode)` / `cell.cloneStyle(patch?)` 直接派生并应用新样式
- `sheet.cell()` / `getCell()` / `setCell()` / `getFormula()` / `setFormula()` 现在同时支持 `A1` 地址和 `(rowNumber, column)` 两种调用方式；行列索引是从 `1` 开始
- `sheet.getDisplayValue()` 和 `cell.text` 是 best-effort 的显示值读取接口：Excel 错误值会保留原文，布尔值会输出 `TRUE` / `FALSE`，其余情况直接基于缓存值做字符串化
- 现在 Excel 错误单元格会额外暴露 `cell.error` / `snapshot.error`。纯错误单元格会标成 `type: "error"`；公式单元格如果缓存结果是错误，则仍然是 `type: "formula"`，但会附带结构化错误信息
- 后续 `getCell` / `getFormula` 会直接走索引查找，不再每次整张表做字符串匹配
- `sheet.rowCount` / `sheet.columnCount` 表示逻辑 used range 的最大行号 / 最大列号，只统计当前带值或公式的单元格；纯空白占位的 `<c>` 节点和只包含空白占位单元格的物理行都不会扩展 used range。空表返回 `0`
- `sheet.getCellEntries()` / `iterCellEntries()` / `getRowEntries()` / `getColumnEntries()` / `getRangeRef()` 是默认的逻辑读取 API，会跳过既没有值也没有公式的空白占位 `<c>` 节点，并按逻辑 used bounds 工作
- `sheet.getPhysicalCellEntries()` / `iterPhysicalCellEntries()` / `getPhysicalRowEntries()` / `getPhysicalColumnEntries()` / `getPhysicalRangeRef()` 用来读取精确的物理 `<c>` 节点边界，适合排查底层 package 结构
- `sheet.deleteCell()` 会真正移除 worksheet 里的 `<c>` 节点；如果你只是想保留样式占位但把值清空，继续用 `setCell(..., null)`
- `workbook.getStyle()` 会读取 `styles.xml` 里的 `cellXfs` 样式定义；`workbook.updateStyle()` 会原位修改已有 `xf`；`workbook.cloneStyle()` 会基于已有 `xf` 追加一个新样式，并返回新的 `styleId`
- `workbook.getFont()` / `updateFont()` / `cloneFont()` 会直接操作 `styles.xml` 里的 `<fonts>`；适合你想复用或维护 `fontId` 时使用
- `workbook.getFill()` / `updateFill()` / `cloneFill()` 会直接操作 `styles.xml` 里的 `<fills>`；适合你想复用或维护 `fillId` 时使用
- `workbook.getBorder()` / `updateBorder()` / `cloneBorder()` 会直接操作 `styles.xml` 里的 `<borders>`；适合你想复用或维护 `borderId` 时使用
- `workbook.getNumberFormat()` / `updateNumberFormat()` / `cloneNumberFormat()` 会直接操作 `styles.xml` 里的 `<numFmt>`；自定义格式会从 `164` 开始分配新的 `numFmtId`
- `sheet.getFont()` 会解析当前单元格最终引用到的字体定义
- `sheet.setFont()` / `cell.setFont()` 会先 clone 当前 `fontId`，再 clone 当前 `styleId` 并把新 `fontId` 套上去，所以只会影响当前单元格，不会污染其它共用旧字体或旧样式的单元格
- `sheet.getFill()` 会解析当前单元格最终引用到的填充定义
- `sheet.setFill()` / `cell.setFill()` 会先 clone 当前 `fillId`，再 clone 当前 `styleId` 并把新 `fillId` 套上去，所以只会影响当前单元格，不会污染其它共用旧填充或旧样式的单元格
- `sheet.getBackgroundColor()` / `cell.backgroundColor` 是基于填充层的简化读取，只在 `solid` 填充且 `fgColor.rgb` 存在时返回 ARGB 颜色
- `sheet.setBackgroundColor()` / `cell.setBackgroundColor()` 是基于填充层的简化写法：传颜色时会写成 `solid + fgColor.rgb`，传 `null` 会回退成 `none`
- `sheet.getBorder()` 会解析当前单元格最终引用到的边框定义
- `sheet.setBorder()` / `cell.setBorder()` 会先 clone 当前 `borderId`，再 clone 当前 `styleId` 并把新 `borderId` 套上去，所以只会影响当前单元格，不会污染其它共用旧边框或旧样式的单元格
- `sheet.getNumberFormat()` 会解析当前单元格最终引用到的数字格式定义，内建格式会直接映射成常见 format code
- `sheet.setNumberFormat()` / `cell.setNumberFormat()` 会先复用或创建目标 `numFmtId`，再 clone 当前 `styleId` 并套上新的数字格式，所以也只影响当前单元格
- `sheet.getAlignment()` 会解析当前单元格最终引用到的 alignment 定义
- `sheet.setAlignment()` / `cell.setAlignment()` 会基于当前 `styleId` clone 出一个新样式，并只更新 alignment / `applyAlignment`，所以同样只影响当前单元格；传 `null` 可以移除 alignment 节点
- `sheet.getStyle()` 会按单元格当前的 `styleId` 读取样式定义；如果单元格本身没有 `s="..."`，会回退到默认样式 `0`
- `sheet.setStyle()` 会基于当前单元格样式克隆出一个新的 `styleId`，写回 `styles.xml`，并把新样式应用到该单元格；这样不会连带修改其它共用旧 `styleId` 的单元格
- `sheet.cloneStyle()` 会基于当前单元格样式克隆出一个新的 `styleId`，写回 `styles.xml`，并把新样式直接应用到该单元格；同样支持 `A1` 和 `(rowNumber, column)`
- `sheet.getStyleId()` / `setStyleId()` 仍然只负责读写单元格上的 `s="..."` 样式索引
- `sheet.getRowStyleId()` / `setRowStyleId()` 当前读写 `<row s="..." customFormat="1">` 这一层的行级样式索引；这一层本身不会修改 `styles.xml`
- `sheet.getRowStyle()` / `sheet.getColumnStyle()` 会把行/列当前引用的样式索引解析成样式定义；如果该行/列没有显式样式，返回 `null`
- `sheet.setRowStyle()` / `sheet.setColumnStyle()` 是行列级的便捷入口，会复用现有 clone 语义：基于当前行/列样式克隆一个新 `styleId`，再立刻应用回去
- `sheet.cloneRowStyle()` / `sheet.cloneColumnStyle()` 会基于当前行/列样式克隆出一个新的 `styleId` 并立即应用；如果当前没有显式样式，会从默认样式 `0` 克隆
- `sheet.getColumnStyleId()` / `setColumnStyleId()` 当前读写 `<cols><col ... style="..."/>` 这一层的列级样式索引；插删列时这些范围也会一起跟着移动
- `sheet.copyStyle()` 当前会把源单元格的 `styleId` 复制到目标单元格，不会改动目标单元格的值或公式；同样支持地址和 `(rowNumber, column)` 两种调用
- `sheet.getFreezePane()` / `freezePane()` / `unfreezePane()` 当前维护 worksheet `sheetViews/sheetView/pane`；插删行列时 `topLeftCell` 也会继续跟随更新
- `sheet.getSelection()` / `setSelection()` 当前读写 worksheet `sheetViews/sheetView/selection`；冻结窗格存在时会优先落在当前 active pane 对应的 selection 上
- 非 batch 模式下，每次写入后都会立刻重建 sheet index，确保后续读取看到的是最新内容；在 `sheet.batch()` / `workbook.batch()` 内，最终的 `dimension` 同步会延后到批次 flush 时统一完成
- 修改工作表后会同步维护 `<dimension ref="...">`，避免使用范围信息过期
- `deleteRow()` / `deleteColumn()` 当前会同步更新本 sheet 的单元格坐标、公式引用、合并区域、`dimension`、常见 `ref/sqref` 属性、`definedNames`，以及其它 sheet 里显式引用它的公式
- `insertRow()` 当前会同步更新本 sheet 的单元格坐标、公式引用、合并区域、`dimension`、常见 `ref/sqref` 属性、`definedNames`，以及其它 sheet 里显式引用它的公式
- `insertColumn()` 当前会同步更新本 sheet 的单元格坐标、公式引用、合并区域、`dimension`、常见 `ref/sqref` 属性、`definedNames`，以及其它 sheet 里显式引用它的公式
- `sheet.getTables()` 当前可以读取已有 table 的名称、显示名、范围和部件路径
- `sheet.getHyperlinks()` 当前可以读取当前 sheet 上的内部和外部超链接；外部链接会解析 sheet rel 里的目标地址
- `sheet.getAutoFilter()` / `sheet.setAutoFilter()` / `sheet.removeAutoFilter()` 当前支持读写 worksheet 顶层 `autoFilter`，移除时会一并清掉顶层 `sortState`
- `sheet.getDataValidations()` / `sheet.setDataValidation()` / `sheet.removeDataValidation()` 当前支持读写 worksheet 顶层 `dataValidations`，包括常见属性与 `formula1/formula2`，并继续跟随插删行列维护 `sqref`
- `sheet.addTable()` 当前会创建最基础的 table part、sheet rel、`[Content_Types].xml` override 和 table XML；列名默认取范围首行，空列名会回退到 `ColumnN`
- `sheet.removeTable()` 当前会同步移除当前 sheet 的 `tableParts`、sheet rel、table XML 和对应的 content type override
- `sheet.setHyperlink()` / `sheet.removeHyperlink()` 当前支持维护 worksheet `<hyperlinks>` 与外部链接对应的 sheet rel，内部链接 target 用 `#Sheet1!A1` 这种格式
- 已有关联 table 在插删行列时会同步维护它们自己的 `ref` / `autoFilter`；如果整块 table 被删空，会从当前 sheet 的 `tableParts` 里移除
- `workbook.getDefinedNames()` / `getDefinedName()` / `setDefinedName()` / `deleteDefinedName()` 当前支持读写全局和本地 `definedNames`
- `workbook.getSheetVisibility()` / `setSheetVisibility()` 当前支持 `visible` / `hidden` / `veryHidden`；并会阻止把最后一张可见 sheet 隐藏掉
- `workbook.getActiveSheet()` / `setActiveSheet()` 当前读写 `workbookView.activeTab`；如果 workbook 里还没有 `bookViews`，会自动补上；隐藏 sheet 不允许设为 active
- `workbook.renameSheet()` / `sheet.rename()` 当前会同步维护 sheet 名、其它 sheet 的显式公式引用、`definedNames`、内部超链接位置和文档属性
- `workbook.moveSheet()` 当前使用 0-based `targetIndex`，会同步维护 workbook 里的 `<sheets>` 顺序、`docProps/app.xml` 里的工作表顺序、本地 `definedNames` 的 `localSheetId`，以及 `workbookView.activeTab`
- `workbook.addSheet()` / `workbook.deleteSheet()` 当前会同步维护 `workbook.xml`、rels、`[Content_Types].xml`，并在删除 sheet 时修正剩余公式与 `definedNames`

## 基准测试

仓库内现在包含一份已脱敏的大型基准文件 [`res/monster.xlsx`](res/monster.xlsx)，可直接用于性能回归对比。

常用命令：

- `npm run bench:monster`
  - 对 `res/monster.xlsx` 运行 3 轮基准
- `npm run bench:check`
  - 对 `res/monster.xlsx` 运行 5 轮基准，并校验 `benchmarks/monster-baseline.json` 里的非空单元格数量，以及配置好的读写耗时阈值
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5`
  - 自定义文件路径和迭代次数；输出 JSON 会同时包含致密遍历结果 `result`、稀疏遍历结果 `sparseResult`、批量写入结果 `writeResult`，以及每个 sheet 的放大量统计
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5 --check benchmarks/monster-baseline.json`
  - 对任意基准文件执行回归检查；数量或配置的读写耗时超出阈值时进程会以非零状态退出

批量写入基准会故意挑选“物理单元格节点最多”的 worksheet，并在一次 `sheet.batch(...)` 里覆盖最多 30 个现有的 `A` 列单元格。这样更容易盯住大表里重复内存写入的热点路径。

## 当前限制

- zip 读写后端现在使用纯 JS 的 `fflate`，不再依赖系统里的 `python3` 与 `zip`
- 当前仍会把整个 zip 包与各个 entry 一起放进内存，对超大文件的峰值内存还可以继续优化
- 字符串写入使用 `inlineStr`，避免为了简单写值而重建 `sharedStrings.xml`
- 合并单元格、批注、富文本、图片等 API 还没加
- 对 XML 的写入是“局部 patch”，不是完整 OOXML 模型

## 开发

```bash
npm run build
npm test
npm run pack:check
npm run validate:task
```

其中：

- `npm test` 直接通过 `tsx` 运行 TypeScript 测试
- `npm run pack:check` 会校验 `npm pack --dry-run` 的结果，确保 npm 包里只有当前源码对应的构建产物，不会夹带旧版本遗留文件
- `npm run validate:task` 直接通过 `tsx` 运行 TypeScript 验证脚本
- `npm run build` 会先清理 `dist`，再生成全新的构建产物

自动化校验目前覆盖：

1. 无修改 roundtrip 的稳定性，包括包内各个 part 的逐字节一致性
2. 单元格、公式、样式、行列、表格、超链接、筛选、冻结窗格、选区、数据验证、defined names 和 CLI 命令等编辑路径
3. 包元数据与 dry-run 打包校验，保证实际发布到 npm 的内容和当前源码树一致

## 真实文件验证

仓库里的 [`res/task.xlsx`](res/task.xlsx) 可以作为后续回归验证样本。

```bash
npm run validate:task
```

如果想验证任意文件：

```bash
npm run validate:roundtrip -- path/to/file.xlsx
```
