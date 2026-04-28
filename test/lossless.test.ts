import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { Workbook, type CellEntry } from "../src/index.ts";

test("roundtrip keeps extracted parts identical", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const outputPath = join(tempRoot, "output.xlsx");
    const expectedEntries = await loadFixtureEntries(fixtureDir);

    const sourceDocument = Workbook.fromEntries(expectedEntries);
    await sourceDocument.save(inputPath);

    const reopened = await Workbook.open(inputPath);
    await reopened.save(outputPath);

    const actualEntries = await Workbook.open(outputPath);
    assertEntryMapsEqual(toEntryMap(expectedEntries), toEntryMap(actualEntries.toEntries()));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("editing a styled cell keeps its style index and leaves styles.xml untouched", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const originalStyles = entryText(entries, "xl/styles.xml");
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getCell("A1"), "Hello");

  sheet.setCell("A1", "World");

  const nextEntries = workbook.toEntries();
  const sheetXml = entryText(nextEntries, "xl/worksheets/sheet1.xml");
  const stylesXml = entryText(nextEntries, "xl/styles.xml");

  assert.match(sheetXml, /<c r="A1" t="inlineStr" s="1">/);
  assert.match(sheetXml, /<t>World<\/t>/);
  assert.equal(stylesXml, originalStyles);
});

test("workbook supports in-memory byte array open and save flows", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const expectedEntries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(expectedEntries);
  const zipped = workbook.toUint8Array();
  const reopened = Workbook.fromUint8Array(zipped);

  assertEntryMapsEqual(toEntryMap(expectedEntries), toEntryMap(reopened.toEntries()));

  reopened.getSheet("Sheet1").setCell("A1", "Bytes");
  const rewritten = Workbook.fromUint8Array(reopened.toUint8Array());

  assert.equal(rewritten.getSheet("Sheet1").getCell("A1"), "Bytes");
});

test("workbook can create a new workbook from a built-in template", async () => {
  const workbook = Workbook.create("Config");
  const sheet = workbook.getSheet("Config");

  assert.deepEqual(workbook.getSheetNames(), ["Config"]);
  assert.equal(workbook.getActiveSheet().name, "Config");
  assert.equal(sheet.rowCount, 0);
  assert.equal(sheet.columnCount, 0);

  sheet.setCell("A1", "Hello");
  workbook.addSheet("Meta");

  assert.equal(sheet.getCell("A1"), "Hello");
  assert.deepEqual(workbook.getSheetNames(), ["Config", "Meta"]);
  assert.match(
    entryText(workbook.toEntries(), "docProps/app.xml"),
    /<vt:lpstr>Config<\/vt:lpstr><vt:lpstr>Meta<\/vt:lpstr>/,
  );
});

test("workbook can create a configured workbook from options", async () => {
  const workbook = Workbook.create({
    activeSheet: "Data",
    author: "Alice",
    modifiedBy: "Bob",
    sheets: [
      {
        name: "Config",
        headers: ["Key", "Value"],
        headerStyle: {
          applyAlignment: true,
          alignment: { horizontal: "center" },
        },
        columnWidths: { A: 20, B: 30 },
        rowHeights: { "1": 24 },
        comments: [{ address: "A1", author: "Alice", text: "Config header" }],
      },
      {
        name: "Data",
        records: [
          { id: 1001, name: "Alpha" },
          { id: 1002, name: "Beta" },
        ],
        frozenPane: { columnCount: 1, rowCount: 1 },
        printArea: "A1:B3",
        printTitles: { rows: "1:1" },
        rangeStyles: [
          { range: "A2:B3", backgroundColor: "FFFFF2CC" },
          { range: "A2:A3", numberFormat: "0" },
        ],
      },
      { name: "Hidden", visibility: "hidden" },
    ],
  });

  assert.deepEqual(workbook.getSheetNames(), ["Config", "Data", "Hidden"]);
  assert.equal(workbook.getActiveSheet().name, "Data");
  assert.equal(workbook.getSheetVisibility("Hidden"), "hidden");
  assert.deepEqual(workbook.getSheet("Config").getHeaders(), ["Key", "Value"]);
  assert.deepEqual(workbook.getSheet("Data").getRecords(), [
    { id: 1001, name: "Alpha" },
    { id: 1002, name: "Beta" },
  ]);
  assert.equal(workbook.getSheet("Config").getColumnWidth("A"), 20);
  assert.equal(workbook.getSheet("Config").getColumnWidth("B"), 30);
  assert.equal(workbook.getSheet("Config").getRowHeight(1), 24);
  assert.deepEqual(workbook.getSheet("Config").getComment("A1"), {
    address: "A1",
    author: "Alice",
    text: "Config header",
  });
  assert.deepEqual(workbook.getSheet("Data").getFreezePane(), {
    activePane: "bottomRight",
    columnCount: 1,
    rowCount: 1,
    topLeftCell: "B2",
  });
  assert.equal(workbook.getSheet("Data").getPrintArea(), "A1:B3");
  assert.deepEqual(workbook.getSheet("Data").getPrintTitles(), { columns: null, rows: "$1:$1" });
  assert.equal(workbook.getSheet("Data").getBackgroundColor("B3"), "FFFFF2CC");
  assert.equal(workbook.getSheet("Data").getNumberFormat("A2")?.code, "0");
  assert.deepEqual(workbook.getSheet("Config").getAlignment("A1"), { horizontal: "center" });

  const coreXml = entryText(workbook.toEntries(), "docProps/core.xml");
  assert.match(coreXml, /<dc:creator>Alice<\/dc:creator>/);
  assert.match(coreXml, /<cp:lastModifiedBy>Bob<\/cp:lastModifiedBy>/);
});

test("workbook supports ArrayBuffer open flows", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const zipped = Workbook.fromEntries(await loadFixtureEntries(fixtureDir)).toUint8Array();
  const reopened = Workbook.fromArrayBuffer(zipped.buffer.slice(
    zipped.byteOffset,
    zipped.byteOffset + zipped.byteLength,
  ));

  assert.equal(reopened.getActiveSheet().name, "Sheet1");
  assert.equal(reopened.getSheet("Sheet1").getCell("A1"), "Hello");
});

test("workbook sheet name APIs are case-insensitive like Excel", () => {
  const workbook = Workbook.create({
    activeSheet: "summary",
    sheets: [{ name: "Data" }, { name: "Summary" }],
  });

  assert.equal(workbook.getActiveSheet().name, "Summary");
  assert.equal(workbook.getSheet("data").name, "Data");
  assert.equal(workbook.tryGetSheet("SUMMARY")?.name, "Summary");
  assert.equal(workbook.hasSheet("summary"), true);

  workbook.setDefinedName("ScopedCell", "$A$1", { scope: "data" });
  assert.equal(workbook.getDefinedName("ScopedCell", "DATA"), "$A$1");

  assert.throws(() => workbook.addSheet("data"), /Sheet already exists: data/);
  assert.throws(() => workbook.renameSheet("data", "summary"), /Sheet already exists: summary/);
  assert.throws(() => Workbook.create({ sheets: ["Data", "data"] }), /Sheet already exists: data/);

  workbook.renameSheet("data", "DATA");
  assert.deepEqual(workbook.getSheetNames(), ["DATA", "Summary"]);

  workbook.moveSheet("summary", 0);
  assert.deepEqual(workbook.getSheetNames(), ["Summary", "DATA"]);

  workbook.setActiveSheet("data");
  assert.equal(workbook.getActiveSheet().name, "DATA");

  workbook.deleteSheet("summary");
  assert.deepEqual(workbook.getSheetNames(), ["DATA"]);
});

test("error cells expose structured error metadata", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="e"><v>#REF!</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const cell = Workbook.fromEntries(entries).getSheet("Sheet1").cell("A1");

  assert.equal(cell.type, "error");
  assert.equal(cell.value, "#REF!");
  assert.deepEqual(cell.error, { code: 0x17, text: "#REF!" });
});

test("formula cells preserve structured error metadata for cached formula errors", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="e"><f>VLOOKUP(B1,C1:D2,2,0)</f><v>#N/A</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const cell = Workbook.fromEntries(entries).getSheet("Sheet1").cell("A1");

  assert.equal(cell.type, "formula");
  assert.equal(cell.formula, "VLOOKUP(B1,C1:D2,2,0)");
  assert.equal(cell.value, "#N/A");
  assert.deepEqual(cell.error, { code: 0x2a, text: "#N/A" });
});

test("sheet.batch groups multiple edits and flushes the final used range", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const workbook = Workbook.fromEntries(await loadFixtureEntries(fixtureDir));
  const sheet = workbook.getSheet("Sheet1");
  const handle = sheet.cell("A1");

  sheet.batch((currentSheet) => {
    currentSheet.setCell("A1", "Batch");
    currentSheet.setCell("C3", 99);
    assert.equal(handle.value, "Batch");
    assert.equal(currentSheet.getCell("C3"), 99);
  });

  assert.equal(sheet.getCell("A1"), "Batch");
  assert.equal(sheet.getCell("C3"), 99);
  assert.equal(sheet.getRangeRef(), "A1:C3");
  assert.match(entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml"), /<dimension ref="A1:C3"\/>/);
});

test("sheet.batch stages cell XML lazily until a full-sheet read needs a flush", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const workbook = Workbook.fromEntries(await loadFixtureEntries(fixtureDir));
  const sheet = workbook.getSheet("Sheet1");
  const readSheetXml = () => (workbook as any).readEntryText("xl/worksheets/sheet1.xml") as string;

  sheet.batch((currentSheet) => {
    currentSheet.setCell("C3", 99);

    assert.equal(currentSheet.getCell("C3"), 99);
    assert.doesNotMatch(readSheetXml(), /r="C3"/);

    assert.equal(currentSheet.getRangeRef(), "A1:C3");
    assert.match(readSheetXml(), /<c r="C3"><v>99<\/v><\/c>/);

    currentSheet.setFormula("B2", "SUM(C3,1)", { cachedValue: 100 });
    assert.equal(currentSheet.getFormula("B2"), "SUM(C3,1)");
    assert.doesNotMatch(readSheetXml(), /SUM\(C3,1\)/);
  });

  const finalXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(finalXml, /<c r="B2"><f>SUM\(C3,1\)<\/f><v>100<\/v><\/c>/);
});

test("workbook.batch can group edits across multiple sheets", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const workbook = Workbook.fromEntries(
    withSecondSheet(
      await loadFixtureEntries(fixtureDir),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Second</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
  );

  workbook.batch((currentWorkbook) => {
    currentWorkbook.getSheet("Sheet1").setCell("B2", "Left");
    currentWorkbook.getSheet("Sheet2").setCell("B2", "Right");
  });

  assert.equal(workbook.getSheet("Sheet1").getCell("B2"), "Left");
  assert.equal(workbook.getSheet("Sheet2").getCell("B2"), "Right");
  assert.equal(workbook.getSheet("Sheet1").getRangeRef(), "A1:B2");
  assert.equal(workbook.getSheet("Sheet2").getRangeRef(), "A1:B2");
});

test("display value helpers expose best-effort user-facing cell text", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const workbook = Workbook.fromEntries(await loadFixtureEntries(fixtureDir));
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("A1", true);
  sheet.setCell("B1", 42);

  assert.equal(sheet.getDisplayValue("A1"), "TRUE");
  assert.equal(sheet.getDisplayValue("B1"), "42");

  const errorWorkbook = Workbook.fromEntries(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="e"><v>#N/A</v></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
  ).getSheet("Sheet1");

  assert.equal(errorWorkbook.getDisplayValue("A1"), "#N/A");
  assert.equal(errorWorkbook.cell("A1").text, "#N/A");
});

test("workbook parsing accepts single-quoted XML attributes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = convertEntriesToSingleQuotedAttributes(await loadFixtureEntries(fixtureDir), [
    "_rels/.rels",
    "xl/_rels/workbook.xml.rels",
    "xl/workbook.xml",
    "xl/styles.xml",
    "xl/worksheets/sheet1.xml",
  ]);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(workbook.getActiveSheet().name, "Sheet1");
  assert.equal(sheet.getCell("A1"), "Hello");
  assert.equal(sheet.getStyleId("A1"), 1);
  assert.equal(sheet.getStyle("A1")?.fontId, 1);

  sheet.setCell("A1", "World");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="A1" t="inlineStr" s="1">/);
  assert.match(sheetXml, /<t>World<\/t>/);
});

test("workbook parsing decodes XML entities inside attribute values", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sales &amp; Ops" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="LocalValue" localSheetId="0">$B$2</definedName>
  </definedNames>
</workbook>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sales & Ops"]);
  assert.equal(workbook.getSheet("Sales & Ops").getCell("A1"), "Hello");
  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "LocalValue", scope: "Sales & Ops", value: "$B$2" },
  ]);
  assert.equal(workbook.getDefinedName("LocalValue", "Sales & Ops"), "$B$2");
});

test("workbook readers tolerate single-quoted workbookView and definedName tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/workbook.xml",
    `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <bookViews><workbookView activeTab='0'/></bookViews>
  <sheets>
    <sheet name='Sheet1' sheetId='1' r:id='rId1' state='hidden'/>
  </sheets>
  <definedNames>
    <definedName name='LocalValue' localSheetId='0'>$B$2</definedName>
  </definedNames>
</workbook>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.equal(workbook.getActiveSheet().name, "Sheet1");
  assert.equal(workbook.getSheetVisibility("Sheet1"), "hidden");
  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "LocalValue", scope: "Sheet1", value: "$B$2" },
  ]);
  assert.equal(workbook.getDefinedName("LocalValue", "Sheet1"), "$B$2");
});

test("sheet addTable and removeTable tolerate single-quoted relationship and content-type attributes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  let entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Score</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Alice</t></is></c>
      <c r="B2"><v>98</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  entries = entries.filter((entry) => entry.path !== "xl/worksheets/_rels/sheet1.xml.rels");
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.addTable("A1:B2", { name: "Scores" });

  let relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  let contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");

  relsXml = relsXml.replace(/"/g, "'");
  contentTypesXml = contentTypesXml.replace(/"/g, "'");
  sheetXml = sheetXml.replace(/"/g, "'");

  const reparsed = Workbook.fromEntries(
    replaceEntryText(
      replaceEntryText(
        replaceEntryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels", relsXml),
        "[Content_Types].xml",
        contentTypesXml,
      ),
      "xl/worksheets/sheet1.xml",
      sheetXml,
    ),
  );

  assert.deepEqual(reparsed.getSheet("Sheet1").getTables(), [
    { name: "Scores", displayName: "Scores", range: "A1:B2", path: "xl/tables/table1.xml" },
  ]);

  reparsed.getSheet("Sheet1").removeTable("Scores");

  assert.doesNotMatch(entryText(reparsed.toEntries(), "xl/worksheets/sheet1.xml"), /<tableParts\b/);
  assert.doesNotMatch(entryText(reparsed.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels"), /table1\.xml/);
  assert.doesNotMatch(entryText(reparsed.toEntries(), "[Content_Types].xml"), /tables\/table1\.xml/);
});

test("workbook addSheet and deleteSheet tolerate single-quoted workbook metadata attributes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = convertEntriesToSingleQuotedAttributes(await loadFixtureEntries(fixtureDir), [
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "[Content_Types].xml",
  ]);
  const workbook = Workbook.fromEntries(entries);

  workbook.addSheet("Sheet2");
  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sheet1", "Sheet2"]);

  let workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  let relsXml = entryText(workbook.toEntries(), "xl/_rels/workbook.xml.rels");
  let contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");

  workbookXml = workbookXml.replace(/"/g, "'");
  relsXml = relsXml.replace(/"/g, "'");
  contentTypesXml = contentTypesXml.replace(/"/g, "'");

  const reparsed = Workbook.fromEntries(
    replaceEntryText(
      replaceEntryText(
        replaceEntryText(workbook.toEntries(), "xl/workbook.xml", workbookXml),
        "xl/_rels/workbook.xml.rels",
        relsXml,
      ),
      "[Content_Types].xml",
      contentTypesXml,
    ),
  );

  reparsed.deleteSheet("Sheet2");

  assert.deepEqual(reparsed.getSheets().map((sheet) => sheet.name), ["Sheet1"]);
  assert.doesNotMatch(entryText(reparsed.toEntries(), "xl/workbook.xml"), /Sheet2/);
  assert.doesNotMatch(entryText(reparsed.toEntries(), "xl/_rels/workbook.xml.rels"), /sheet2\.xml/);
  assert.doesNotMatch(entryText(reparsed.toEntries(), "[Content_Types].xml"), /worksheets\/sheet2\.xml/);
});

test("workbook active sheet and visibility writers tolerate single-quoted workbook metadata tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      withSecondSheet(
        await loadFixtureEntries(fixtureDir),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
      ),
      "xl/workbook.xml",
      `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <bookViews><workbookView activeTab='0'/></bookViews>
  <sheets>
    <sheet name='Sheet1' sheetId='1' r:id='rId1'/>
    <sheet name='Sheet2' sheetId='2' r:id='rId3' state='hidden'/>
  </sheets>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.setSheetVisibility("Sheet2", "visible");
  workbook.setActiveSheet("Sheet2");

  assert.equal(workbook.getSheetVisibility("Sheet2"), "visible");
  assert.equal(workbook.getActiveSheet().name, "Sheet2");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<workbookView activeTab="1"\/>/);
  assert.match(workbookXml, /<sheet name="Sheet2" sheetId="2" r:id="rId3"\/>/);
  assert.doesNotMatch(workbookXml, /Sheet2'[^>]*state=/);
});

test("workbook moveSheet tolerates single-quoted sheet and workbookView tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      replaceEntryText(
        withSecondSheet(
          await loadFixtureEntries(fixtureDir),
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
        ),
        "xl/workbook.xml",
        `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <bookViews><workbookView activeTab='1'/></bookViews>
  <sheets>
    <sheet name='Sheet1' sheetId='1' r:id='rId1'/>
    <sheet name='Sheet2' sheetId='2' r:id='rId3'/>
    <sheet name='Sheet3' sheetId='3' r:id='rId4'/>
  </sheets>
</workbook>`,
      ),
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`,
    ),
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>fastxlsx</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr><vt:lpstr>Sheet3</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
  );
  const withThirdSheet = [
    ...entries,
    {
      path: "xl/worksheets/sheet3.xml",
      data: new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>3</v></c></row>
  </sheetData>
</worksheet>`),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
  const workbook = Workbook.fromEntries(withThirdSheet);

  workbook.moveSheet("Sheet3", 0);

  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sheet3", "Sheet1", "Sheet2"]);

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(
    workbookXml,
    /<sheets><sheet name='Sheet3' sheetId='3' r:id='rId4'\/><sheet name='Sheet1' sheetId='1' r:id='rId1'\/><sheet name='Sheet2' sheetId='2' r:id='rId3'\/><\/sheets>/,
  );
  assert.match(workbookXml, /<workbookView activeTab="2"\/>/);
});

test("sheet metadata readers tolerate single-quoted selection, merge, and hyperlink tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const encoder = new TextEncoder();
  const entries = [
    ...replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>
  <sheetViews>
    <sheetView workbookViewId='0'>
      <selection pane='bottomRight' activeCell='B2' sqref='B2:C3'/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r='1'><c r='A1' t='inlineStr'><is><t>Hello</t></is></c></row>
  </sheetData>
  <mergeCells count='1'><mergeCell ref='D4:E5'/></mergeCells>
  <hyperlinks>
    <hyperlink ref='A1' r:id='rId1' tooltip='Open site'/>
    <hyperlink ref='B2' location='#Sheet1!A1'/>
  </hyperlinks>
</worksheet>`,
    ),
    {
      path: "xl/worksheets/_rels/sheet1.xml.rels",
      data: encoder.encode(`<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.com' TargetMode='External'/>
</Relationships>`),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getSelection(), {
    activeCell: "B2",
    pane: "bottomRight",
    range: "B2:C3",
  });
  assert.deepEqual(sheet.getMergedRanges(), ["D4:E5"]);
  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "A1", target: "https://example.com", tooltip: "Open site", type: "external" },
    { address: "B2", target: "#Sheet1!A1", tooltip: null, type: "internal" },
  ]);
});

test("sheet metadata writers tolerate single-quoted sheetViews and hyperlinks containers", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>
  <sheetViews>
    <sheetView workbookViewId='0'>
      <selection activeCell='A1' sqref='A1'/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r='1'><c r='A1' t='inlineStr'><is><t>Hello</t></is></c></row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref='A1' location='#Sheet1!A1'/>
  </hyperlinks>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.freezePane(1, 1);
  assert.deepEqual(sheet.getFreezePane(), {
    activePane: "bottomRight",
    columnCount: 1,
    rowCount: 1,
    topLeftCell: "B2",
  });

  sheet.setSelection("C3", "C3:D4");
  assert.deepEqual(sheet.getSelection(), {
    activeCell: "C3",
    pane: "bottomRight",
    range: "C3:D4",
  });

  sheet.unfreezePane();
  assert.equal(sheet.getFreezePane(), null);
  assert.deepEqual(sheet.getSelection(), {
    activeCell: "C3",
    pane: null,
    range: "C3:D4",
  });

  sheet.setHyperlink("B2", "#Sheet1!A1", { tooltip: "Jump" });
  sheet.removeHyperlink("A1");

  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "B2", target: "#Sheet1!A1", tooltip: "Jump", type: "internal" },
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<pane\b/);
  assert.match(sheetXml, /<selection activeCell="C3" sqref="C3:D4"\/>/);
  assert.doesNotMatch(sheetXml, /<hyperlink ref="A1"/);
  assert.match(sheetXml, /<hyperlink ref="B2" location="#Sheet1!A1" tooltip="Jump"\/>/);
});

test("defined name deletion tolerates single-quoted definedNames tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/workbook.xml",
    `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <sheets>
    <sheet name='Sheet1' sheetId='1' r:id='rId1'/>
  </sheets>
  <definedNames>
    <definedName name='GlobalValue'>$A$1</definedName>
    <definedName name='LocalValue' localSheetId='0'>$B$2</definedName>
  </definedNames>
</workbook>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.deleteDefinedName("LocalValue", "Sheet1");
  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "GlobalValue", scope: null, value: "$A$1" },
  ]);

  workbook.deleteDefinedName("GlobalValue");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.doesNotMatch(workbookXml, /LocalValue/);
  assert.doesNotMatch(workbookXml, /GlobalValue/);
  assert.doesNotMatch(workbookXml, /<definedNames\b/);
});

test("defined name writers tolerate single-quoted definedName tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/workbook.xml",
    `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <sheets>
    <sheet name='Sheet1' sheetId='1' r:id='rId1'/>
  </sheets>
  <definedNames>
    <definedName name='GlobalValue'>Sheet1!$A$1</definedName>
    <definedName name='LocalValue' localSheetId='0'>$B$2</definedName>
  </definedNames>
</workbook>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.setDefinedName("GlobalValue", "Sheet1!$C$3");
  workbook.setDefinedName("LocalValue", "$D$4", { scope: "Sheet1" });

  assert.equal(workbook.getDefinedName("GlobalValue"), "Sheet1!$C$3");
  assert.equal(workbook.getDefinedName("LocalValue", "Sheet1"), "$D$4");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<definedName name="GlobalValue">Sheet1!\$C\$3<\/definedName>/);
  assert.match(workbookXml, /<definedName name="LocalValue" localSheetId="0">\$D\$4<\/definedName>/);
});

test("column structure rewrites tolerate single-quoted definedName tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/workbook.xml",
      `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <sheets>
    <sheet name='Sheet1' sheetId='1' r:id='rId1'/>
  </sheets>
  <definedNames>
    <definedName name='_xlnm.Print_Area' localSheetId='0'>$A$1:$C$4</definedName>
    <definedName name='DataRange'>Sheet1!$B$2:$C$4</definedName>
  </definedNames>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>
  <sheetData>
    <row r='1'>
      <c r='A1'><v>1</v></c>
      <c r='B1'><v>2</v></c>
      <c r='C1'><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.getSheet("Sheet1").insertColumn(2);

  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "_xlnm.Print_Area", scope: "Sheet1", value: "$A$1:$D$4" },
    { hidden: false, name: "DataRange", scope: null, value: "Sheet1!$C$2:$D$4" },
  ]);
});

test("sheet metadata rewrites tolerate single-quoted definedName tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    withSecondSheet(
      await loadFixtureEntries(fixtureDir),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
    ),
    "xl/workbook.xml",
    `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <bookViews><workbookView activeTab='1'/></bookViews>
  <sheets>
    <sheet name='Sheet1' sheetId='1' r:id='rId1'/>
    <sheet name='Sheet2' sheetId='2' r:id='rId3'/>
  </sheets>
  <definedNames>
    <definedName name='ExternalRef'>Sheet1!$A$1</definedName>
    <definedName name='LocalToSheet1' localSheetId='0'>$A$1</definedName>
    <definedName name='LocalToSheet2' localSheetId='1'>$A$1</definedName>
  </definedNames>
</workbook>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.renameSheet("Sheet1", "Data");
  workbook.moveSheet("Sheet2", 0);

  assert.equal(workbook.getActiveSheet().name, "Sheet2");
  assert.equal(workbook.getDefinedName("ExternalRef"), "Data!$A$1");
  assert.equal(workbook.getDefinedName("LocalToSheet1", "Data"), "$A$1");
  assert.equal(workbook.getDefinedName("LocalToSheet2", "Sheet2"), "$A$1");

  workbook.deleteSheet("Sheet2");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<definedName name=['"]ExternalRef['"]>Data!\$A\$1<\/definedName>/);
  assert.match(workbookXml, /<definedName name="LocalToSheet1" localSheetId="0">\$A\$1<\/definedName>/);
  assert.doesNotMatch(workbookXml, /LocalToSheet2/);
});

test("sheet reads stay coherent after repeated writes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("B1", 1);
  assert.equal(sheet.getCell("B1"), 1);

  sheet.setCell("B1", 2);
  assert.equal(sheet.getCell("B1"), 2);

  sheet.setCell("A2", "Tail");
  assert.equal(sheet.getCell("A2"), "Tail");
});

test("formula cells can be read and updated without dropping styles", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><f>SUM(1,2)</f><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getFormula("A1"), "SUM(1,2)");
  assert.equal(sheet.getCell("A1"), 3);

  sheet.setFormula("A1", 'CONCAT("He","llo")', { cachedValue: "Hello" });

  assert.equal(sheet.getFormula("A1"), 'CONCAT("He","llo")');
  assert.equal(sheet.getCell("A1"), "Hello");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="A1" t="str" s="1">/);
  assert.match(sheetXml, /<f>CONCAT\(&quot;He&quot;,&quot;llo&quot;\)<\/f>/);
  assert.match(sheetXml, /<v>Hello<\/v>/);
});

test("shared formula follower cells resolve translated formulas", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><f t="shared" ref="A1:B2" si="0">B1+$C$1+D$1+$E2+SUM(F1:G2)</f><v>1</v></c>
    </row>
    <row r="2">
      <c r="B2" s="1"><f t="shared" si="0"/><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getFormula("A1"), "B1+$C$1+D$1+$E2+SUM(F1:G2)");
  assert.equal(sheet.getFormula("B2"), "C2+$C$1+E$1+$E3+SUM(G2:H3)");
  assert.equal(sheet.cell("B2").type, "formula");
  assert.equal(sheet.getCell("B2"), 2);
});

test("formula cells with self-closing string cached values read as empty strings", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1" t="str"><f>IF(1=0,"x","")</f><v/></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getFormula("A1"), 'IF(1=0,"x","")');
  assert.equal(sheet.getCell("A1"), "");
});

test("formula caches stay stale until cell.recalculate is called manually", () => {
  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("A1", 1);
  sheet.setCell("B1", 2);
  sheet.setFormula("C1", "SUM(A1:B1)", { cachedValue: 3 });

  sheet.setCell("A1", 10);

  assert.equal(sheet.getCell("C1"), 3);
  assert.equal(sheet.cell("C1").value, 3);

  const snapshot = sheet.cell("C1").recalculate();

  assert.equal(snapshot.value, 12);
  assert.equal(sheet.getCell("C1"), 12);
  assert.match(entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml"), /<c r="C1"><f>SUM\(A1:B1\)<\/f><v>12<\/v><\/c>/);
});

test("sheet and workbook recalculate APIs stay manual and resolve cross-sheet names", () => {
  const workbook = Workbook.create({
    sheets: [{ name: "Data" }, { name: "Summary" }],
  });
  let data = workbook.getSheet("Data");
  let summary = workbook.getSheet("Summary");

  data.setCell("A1", 4);
  data.setCell("A2", 6);
  workbook.setDefinedName("Total", "SUM(Data!A1:A2)");
  data = workbook.getSheet("Data");
  summary = workbook.getSheet("Summary");
  summary.setFormula("B1", "Total", { cachedValue: 0 });
  summary.setFormula("B2", "Data!A1+1", { cachedValue: 0 });

  assert.equal(summary.getCell("B1"), 0);
  assert.equal(summary.getCell("B2"), 0);

  const sheetSummary = summary.recalculate();

  assert.deepEqual(sheetSummary, { cells: 2, sheets: 1, updated: 2 });
  assert.equal(summary.getCell("B1"), 10);
  assert.equal(summary.getCell("B2"), 5);

  data.setCell("A1", 10);

  assert.equal(summary.getCell("B1"), 10);
  assert.equal(summary.getCell("B2"), 5);

  const workbookSummary = workbook.recalculate();

  assert.deepEqual(workbookSummary, { cells: 2, sheets: 2, updated: 2 });
  assert.equal(summary.getCell("B1"), 16);
  assert.equal(summary.getCell("B2"), 11);
});

test("cross-sheet formulas and sheet rewrites ignore sheet name case", () => {
  const workbook = Workbook.create({
    sheets: [{ name: "Data" }, { name: "Summary" }],
  });
  const data = workbook.getSheet("Data");
  const summary = workbook.getSheet("Summary");

  data.setCell("A1", 4);
  data.setCell("A2", 6);
  summary.setFormula("A1", "data!A1+1", { cachedValue: 0 });
  summary.setFormula("A2", "SUM('data'!A1:A2)", { cachedValue: 0 });

  const recalc = summary.recalculate();

  assert.deepEqual(recalc, { cells: 2, sheets: 1, updated: 2 });
  assert.equal(summary.getCell("A1"), 5);
  assert.equal(summary.getCell("A2"), 10);

  workbook.renameSheet("DATA", "Data Set");
  assert.equal(workbook.getSheet("Summary").getFormula("A1"), "'Data Set'!A1+1");
  assert.equal(workbook.getSheet("Summary").getFormula("A2"), "SUM('Data Set'!A1:A2)");

  workbook.deleteSheet("data set");
  assert.equal(workbook.getSheet("Summary").getFormula("A1"), "#REF!+1");
  assert.equal(workbook.getSheet("Summary").getFormula("A2"), "SUM(#REF!)");
});

test("manual recalc writes cached formula errors with structured metadata", () => {
  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("A1", 0);
  sheet.setFormula("B1", "1/A1", { cachedValue: 1 });

  const snapshot = sheet.recalculateCell("B1");

  assert.equal(snapshot.value, "#DIV/0!");
  assert.deepEqual(snapshot.error, { code: 0x07, text: "#DIV/0!" });
  assert.equal(sheet.getCell("B1"), "#DIV/0!");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="B1" t="e"><f>1\/A1<\/f><v>#DIV\/0!<\/v><\/c>/);
});

test("manual recalc rejects shared formula cells for now", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><f t="shared" si="0">SUM(B1:C1)</f><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.throws(
    () => workbook.getSheet("Sheet1").cell("A1").recalculate(),
    /Unsupported formula shape at Sheet1!A1/,
  );
});

test("manual recalc supports MATCH and VLOOKUP lookups", () => {
  const workbook = Workbook.create({
    sheets: [{ name: "Data" }, { name: "Summary" }],
  });
  let data = workbook.getSheet("Data");
  let summary = workbook.getSheet("Summary");

  data.setRange("A1", [
    [1, "Low"],
    [5, "Mid"],
    [10, "High"],
    [20, "Top"],
  ]);
  data.setColumn("D", [20, 10, 5, 1], 1);
  data.setRange("F1", [
    ["Alpha", 11],
    ["Bravo", 22],
    ["Charlie", 33],
  ]);

  workbook.setDefinedName("LookupTable", "Data!$F$1:$G$3");
  summary = workbook.getSheet("Summary");

  summary.setFormula("A1", "MATCH(10,Data!A1:A4,0)", { cachedValue: 0 });
  summary.setFormula("A2", "MATCH(9,Data!A1:A4,1)", { cachedValue: 0 });
  summary.setFormula("A3", "MATCH(9,Data!D1:D4,-1)", { cachedValue: 0 });
  summary.setFormula("A4", 'VLOOKUP("bravo",LookupTable,2,0)', { cachedValue: 0 });
  summary.setFormula("A5", "VLOOKUP(9,Data!A1:B4,2,TRUE)", { cachedValue: "" });

  const recalc = summary.recalculate();

  assert.deepEqual(recalc, { cells: 5, sheets: 1, updated: 5 });
  assert.equal(summary.getCell("A1"), 3);
  assert.equal(summary.getCell("A2"), 2);
  assert.equal(summary.getCell("A3"), 2);
  assert.equal(summary.getCell("A4"), 22);
  assert.equal(summary.getCell("A5"), "Mid");
});

test("inline string cells decode numeric entities and ignore phonetic runs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr">
        <is>
          <r><t>Hel</t></r>
          <rPh sb="0" eb="3"><t>X</t></rPh>
          <r><t>lo&#10;World</t></r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getCell("A1"), "Hello\nWorld");
});

test("shared string cells decode numeric entities and ignore phonetic runs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const encoder = new TextEncoder();
  const entries = [
    ...replaceEntryText(
      replaceEntryText(
        await loadFixtureEntries(fixtureDir),
        "xl/worksheets/sheet1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`,
      ),
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
    ),
    {
      path: "xl/sharedStrings.xml",
      data: encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>Hel</t></r>
    <rPh sb="0" eb="3"><t>X</t></rPh>
    <r><t>lo&#10;World</t></r>
  </si>
</sst>`),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getCell("A1"), "Hello\nWorld");
});

test("shared string parsing tolerates single-quoted rich-text and phonetic tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const encoder = new TextEncoder();
  const entries = [
    ...replaceEntryText(
      replaceEntryText(
        await loadFixtureEntries(fixtureDir),
        "xl/worksheets/sheet1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`,
      ),
      "xl/_rels/workbook.xml.rels",
      `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet1.xml'/>
  <Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>
  <Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings' Target='sharedStrings.xml'/>
</Relationships>`,
    ),
    {
      path: "xl/sharedStrings.xml",
      data: encoder.encode(`<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<sst xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' count='1' uniqueCount='1'>
  <si>
    <r><t xml:space='preserve'>Hel</t></r>
    <rPh sb='0' eb='3'><t>X</t></rPh>
    <r><t>lo&#10;World</t></r>
  </si>
</sst>`),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
  const workbook = Workbook.fromEntries(entries);

  assert.equal(workbook.getSheet("Sheet1").getCell("A1"), "Hello\nWorld");
});

test("cell handle objects cache parsed state and refresh after writes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const first = sheet.cell("A1");
  const second = sheet.cell("A1");

  assert.equal(first, second);
  assert.equal(first.exists, true);
  assert.equal(first.type, "string");
  assert.equal(first.styleId, 1);
  assert.equal(first.formula, null);
  assert.equal(first.value, "Hello");

  first.setValue("World");

  assert.equal(first.value, "World");
  assert.equal(first.type, "string");

  sheet.setFormula("A1", "SUM(1,2)", { cachedValue: 3 });

  assert.equal(first.formula, "SUM(1,2)");
  assert.equal(first.type, "formula");
  assert.equal(first.value, 3);

  const missing = sheet.cell("C9");
  assert.equal(missing.exists, false);
  assert.equal(missing.type, "missing");
  assert.equal(missing.value, null);
});

test("cell APIs accept 1-based row and column indexes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getCell(1, 1), "Hello");
  assert.equal(sheet.cell(1, 1).value, "Hello");

  sheet.setCell(2, 3, "Tail");
  assert.equal(sheet.getCell("C2"), "Tail");
  assert.equal(sheet.getCell(2, 3), "Tail");

  sheet.setFormula(3, 2, "SUM(1,2)", { cachedValue: 3 });
  assert.equal(sheet.getFormula("B3"), "SUM(1,2)");
  assert.equal(sheet.getFormula(3, 2), "SUM(1,2)");
  assert.equal(sheet.getCell(3, 2), 3);
});

test("style id APIs read and write style indexes by address and indexes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="2"><f>SUM(1,2)</f><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("A1");

  assert.equal(sheet.getStyleId("A1"), 1);
  assert.equal(sheet.getStyleId(1, 2), 2);
  assert.equal(cell.styleId, 1);

  sheet.setStyleId(1, 1, 5);
  sheet.setStyleId("B1", 6);
  sheet.setStyleId(2, 3, 7);

  assert.equal(sheet.getStyleId("A1"), 5);
  assert.equal(sheet.getStyleId(1, 2), 6);
  assert.equal(sheet.getStyleId("C2"), 7);
  assert.equal(cell.styleId, 5);
  assert.equal(sheet.getCell("A1"), "Hello");
  assert.equal(sheet.getFormula("B1"), "SUM(1,2)");
  assert.equal(sheet.getCell("B1"), 3);

  cell.setStyleId(null);
  assert.equal(sheet.getStyleId("A1"), null);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="A1" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="6"><f>SUM\(1,2\)<\/f><v>3<\/v><\/c>/);
  assert.match(sheetXml, /<c r="C2" s="7"\/>/);
});

test("cell style patch APIs clone and apply styles without mutating shared style ids", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="1" t="inlineStr"><is><t>World</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("B1");

  const a1StyleId = sheet.setStyle("A1", {
    numFmtId: 14,
    applyNumberFormat: true,
    applyAlignment: true,
    alignment: {
      horizontal: "center",
    },
  });
  const b1StyleId = cell.setStyle({
    applyAlignment: true,
    alignment: {
      horizontal: "right",
    },
  });

  assert.equal(a1StyleId, 2);
  assert.equal(b1StyleId, 3);
  assert.equal(sheet.getStyleId("A1"), 2);
  assert.equal(sheet.getStyleId("B1"), 3);
  assert.equal(workbook.getStyle(1)?.numFmtId, 0);
  assert.equal(workbook.getStyle(1)?.alignment, null);
  assert.equal(sheet.getStyle("A1")?.numFmtId, 14);
  assert.equal(sheet.getStyle("A1")?.alignment?.horizontal, "center");
  assert.equal(cell.style?.alignment?.horizontal, "right");

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");

  assert.match(stylesXml, /<cellXfs count="4">/);
  assert.match(sheetXml, /<c r="A1" s="2" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="3" t="inlineStr"><is><t>World<\/t><\/is><\/c>/);
});

test("font definition APIs read, clone, and update workbook fonts", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);

  assert.deepEqual(workbook.getFont(0), {
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: 11,
    name: "Calibri",
    family: 2,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: null,
  });
  assert.deepEqual(workbook.getFont(1), {
    bold: true,
    italic: null,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: 11,
    name: "Calibri",
    family: 2,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: null,
  });

  const nextFontId = workbook.cloneFont(1, {
    italic: true,
    color: {
      rgb: "FFFF0000",
    },
  });
  workbook.updateFont(0, {
    name: "Arial",
    size: 12,
  });

  assert.equal(nextFontId, 2);
  assert.deepEqual(workbook.getFont(0), {
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: 12,
    name: "Arial",
    family: 2,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: null,
  });
  assert.deepEqual(workbook.getFont(2), {
    bold: true,
    italic: true,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: 11,
    name: "Calibri",
    family: 2,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: {
      rgb: "FFFF0000",
    },
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  assert.match(stylesXml, /<fonts count="3">/);
  assert.match(stylesXml, /<font><sz val="12"\/><name val="Arial"\/><family val="2"\/><\/font>/);
  assert.match(stylesXml, /<font><b\/><i\/><color rgb="FFFF0000"\/><sz val="11"\/><name val="Calibri"\/><family val="2"\/><\/font>/);
});

test("cell font APIs clone and apply fonts without mutating shared font ids", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="1" t="inlineStr"><is><t>World</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("B1");

  assert.equal(sheet.getFont("A1")?.bold, true);
  assert.equal(cell.font?.bold, true);

  const a1FontId = sheet.setFont("A1", {
    italic: true,
    color: {
      rgb: "FFFF0000",
    },
  });
  const b1FontId = cell.setFont({
    bold: null,
    name: "Arial",
    size: 12,
  });

  assert.equal(a1FontId, 2);
  assert.equal(b1FontId, 3);
  assert.equal(workbook.getFont(1)?.bold, true);
  assert.equal(workbook.getFont(1)?.italic, null);
  assert.deepEqual(sheet.getFont("A1"), {
    bold: true,
    italic: true,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: 11,
    name: "Calibri",
    family: 2,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: {
      rgb: "FFFF0000",
    },
  });
  assert.deepEqual(cell.font, {
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: 12,
    name: "Arial",
    family: 2,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: null,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(stylesXml, /<fonts count="4">/);
  assert.match(sheetXml, /<c r="A1" s="2" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="3" t="inlineStr"><is><t>World<\/t><\/is><\/c>/);
});

test("fill definition APIs read, clone, and update workbook fills", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);

  assert.deepEqual(workbook.getFill(0), {
    patternType: "none",
    fgColor: null,
    bgColor: null,
  });
  assert.deepEqual(workbook.getFill(1), {
    patternType: "gray125",
    fgColor: null,
    bgColor: null,
  });

  const nextFillId = workbook.cloneFill(0, {
    patternType: "solid",
    fgColor: {
      rgb: "FFFF0000",
    },
  });
  workbook.updateFill(1, {
    patternType: "solid",
    fgColor: {
      rgb: "FF00FF00",
    },
  });

  assert.equal(nextFillId, 2);
  assert.deepEqual(workbook.getFill(1), {
    patternType: "solid",
    fgColor: {
      rgb: "FF00FF00",
    },
    bgColor: null,
  });
  assert.deepEqual(workbook.getFill(2), {
    patternType: "solid",
    fgColor: {
      rgb: "FFFF0000",
    },
    bgColor: null,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  assert.match(stylesXml, /<fills count="3">/);
  assert.match(stylesXml, /<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"\/><\/patternFill><\/fill>/);
  assert.match(stylesXml, /<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"\/><\/patternFill><\/fill>/);
});

test("cell fill APIs clone and apply fills without mutating shared fill ids", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="1" t="inlineStr"><is><t>World</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("B1");

  assert.equal(sheet.getFill("A1")?.patternType, "none");
  assert.equal(cell.fill?.patternType, "none");

  const a1FillId = sheet.setFill("A1", {
    patternType: "solid",
    fgColor: {
      rgb: "FFFF0000",
    },
  });
  const b1FillId = cell.setFill({
    patternType: "solid",
    fgColor: {
      rgb: "FF00FF00",
    },
  });

  assert.equal(a1FillId, 2);
  assert.equal(b1FillId, 3);
  assert.equal(workbook.getFill(0)?.patternType, "none");
  assert.equal(workbook.getFill(0)?.fgColor, null);
  assert.deepEqual(sheet.getFill("A1"), {
    patternType: "solid",
    fgColor: {
      rgb: "FFFF0000",
    },
    bgColor: null,
  });
  assert.deepEqual(cell.fill, {
    patternType: "solid",
    fgColor: {
      rgb: "FF00FF00",
    },
    bgColor: null,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(stylesXml, /<fills count="4">/);
  assert.match(sheetXml, /<c r="A1" s="2" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="3" t="inlineStr"><is><t>World<\/t><\/is><\/c>/);
});

test("background color helper APIs set solid fills and can clear them", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="1" t="inlineStr"><is><t>World</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("B1");

  assert.equal(sheet.getBackgroundColor("A1"), null);
  assert.equal(cell.backgroundColor, null);

  const a1FillId = sheet.setBackgroundColor("A1", "FFFF0000");
  const b1FillId = cell.setBackgroundColor("FF00FF00");

  assert.equal(a1FillId, 2);
  assert.equal(b1FillId, 3);
  assert.equal(sheet.getBackgroundColor("A1"), "FFFF0000");
  assert.equal(cell.backgroundColor, "FF00FF00");
  assert.equal(workbook.getFill(0)?.patternType, "none");

  const clearedFillId = cell.setBackgroundColor(null);

  assert.equal(clearedFillId, 4);
  assert.equal(cell.backgroundColor, null);
  assert.deepEqual(cell.fill, {
    patternType: "none",
    fgColor: null,
    bgColor: null,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(stylesXml, /<fills count="5">/);
  assert.match(stylesXml, /<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"\/><\/patternFill><\/fill>/);
  assert.match(stylesXml, /<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"\/><\/patternFill><\/fill>/);
  assert.match(sheetXml, /<c r="A1" s="2" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="4" t="inlineStr"><is><t>World<\/t><\/is><\/c>/);
});

test("border definition APIs read, clone, and update workbook borders", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);

  assert.deepEqual(workbook.getBorder(0), {
    left: { style: null, color: null },
    right: { style: null, color: null },
    top: { style: null, color: null },
    bottom: { style: null, color: null },
    diagonal: { style: null, color: null },
    vertical: null,
    horizontal: null,
    diagonalUp: null,
    diagonalDown: null,
    outline: null,
  });

  const nextBorderId = workbook.cloneBorder(0, {
    bottom: {
      style: "double",
      color: {
        rgb: "FF00FF00",
      },
    },
  });
  workbook.updateBorder(0, {
    top: {
      style: "thin",
      color: {
        rgb: "FFFF0000",
      },
    },
  });

  assert.equal(nextBorderId, 1);
  assert.deepEqual(workbook.getBorder(0), {
    left: { style: null, color: null },
    right: { style: null, color: null },
    top: { style: "thin", color: { rgb: "FFFF0000" } },
    bottom: { style: null, color: null },
    diagonal: { style: null, color: null },
    vertical: null,
    horizontal: null,
    diagonalUp: null,
    diagonalDown: null,
    outline: null,
  });
  assert.deepEqual(workbook.getBorder(1), {
    left: { style: null, color: null },
    right: { style: null, color: null },
    top: { style: null, color: null },
    bottom: { style: "double", color: { rgb: "FF00FF00" } },
    diagonal: { style: null, color: null },
    vertical: null,
    horizontal: null,
    diagonalUp: null,
    diagonalDown: null,
    outline: null,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  assert.match(stylesXml, /<borders count="2">/);
  assert.match(
    stylesXml,
    /<border><left\/><right\/><top style="thin"><color rgb="FFFF0000"\/><\/top><bottom\/><diagonal\/><\/border>/,
  );
  assert.match(
    stylesXml,
    /<border><left\/><right\/><top\/><bottom style="double"><color rgb="FF00FF00"\/><\/bottom><diagonal\/><\/border>/,
  );
});

test("cell border APIs clone and apply borders without mutating shared border ids", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="1" t="inlineStr"><is><t>World</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("B1");

  assert.equal(sheet.getBorder("A1")?.top?.style, null);
  assert.equal(cell.border?.bottom?.style, null);

  const a1BorderId = sheet.setBorder("A1", {
    top: {
      style: "thin",
      color: {
        rgb: "FFFF0000",
      },
    },
  });
  const b1BorderId = cell.setBorder({
    bottom: {
      style: "double",
      color: {
        rgb: "FF00FF00",
      },
    },
  });

  assert.equal(a1BorderId, 1);
  assert.equal(b1BorderId, 2);
  assert.equal(workbook.getBorder(0)?.top?.style, null);
  assert.equal(workbook.getBorder(0)?.bottom?.style, null);
  assert.deepEqual(sheet.getBorder("A1"), {
    left: { style: null, color: null },
    right: { style: null, color: null },
    top: { style: "thin", color: { rgb: "FFFF0000" } },
    bottom: { style: null, color: null },
    diagonal: { style: null, color: null },
    vertical: null,
    horizontal: null,
    diagonalUp: null,
    diagonalDown: null,
    outline: null,
  });
  assert.deepEqual(cell.border, {
    left: { style: null, color: null },
    right: { style: null, color: null },
    top: { style: null, color: null },
    bottom: { style: "double", color: { rgb: "FF00FF00" } },
    diagonal: { style: null, color: null },
    vertical: null,
    horizontal: null,
    diagonalUp: null,
    diagonalDown: null,
    outline: null,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(stylesXml, /<borders count="3">/);
  assert.match(sheetXml, /<c r="A1" s="2" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="3" t="inlineStr"><is><t>World<\/t><\/is><\/c>/);
});

test("number format APIs read, clone, and update workbook numFmts", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);

  assert.deepEqual(workbook.getNumberFormat(0), {
    builtin: true,
    code: "General",
    numFmtId: 0,
  });
  assert.equal(workbook.getNumberFormat(164), null);

  const nextNumFmtId = workbook.cloneNumberFormat(0, '#,##0.00_);[Red](#,##0.00)');
  workbook.updateNumberFormat(nextNumFmtId, "0.000");

  assert.equal(nextNumFmtId, 164);
  assert.deepEqual(workbook.getNumberFormat(164), {
    builtin: false,
    code: "0.000",
    numFmtId: 164,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  assert.match(stylesXml, /<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"\/><\/numFmts><fonts/);
});

test("cell number format APIs clone and apply numFmt ids without mutating shared styles", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="1" t="inlineStr"><is><t>World</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("B1");

  assert.deepEqual(sheet.getNumberFormat("A1"), {
    builtin: true,
    code: "General",
    numFmtId: 0,
  });

  const a1NumFmtId = sheet.setNumberFormat("A1", "0.00%");
  const b1NumFmtId = cell.setNumberFormat('#,##0.00_);[Red](#,##0.00)');

  assert.equal(a1NumFmtId, 10);
  assert.equal(b1NumFmtId, 164);
  assert.deepEqual(sheet.getNumberFormat("A1"), {
    builtin: true,
    code: "0.00%",
    numFmtId: 10,
  });
  assert.deepEqual(cell.numberFormat, {
    builtin: false,
    code: '#,##0.00_);[Red](#,##0.00)',
    numFmtId: 164,
  });
  assert.deepEqual(workbook.getNumberFormat(0), {
    builtin: true,
    code: "General",
    numFmtId: 0,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(stylesXml, /<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00_\);\[Red\]\(#,##0.00\)"\/><\/numFmts>/);
  assert.match(sheetXml, /<c r="A1" s="2" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="3" t="inlineStr"><is><t>World<\/t><\/is><\/c>/);
});

test("cell alignment APIs clone and apply alignments without mutating shared styles", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="1" t="inlineStr"><is><t>World</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const cell = sheet.cell("B1");

  assert.equal(sheet.getAlignment("A1"), null);
  assert.equal(cell.alignment, null);

  const a1StyleId = sheet.setAlignment("A1", {
    horizontal: "center",
    wrapText: true,
  });
  const b1StyleId = cell.setAlignment({
    horizontal: "right",
  });

  assert.equal(a1StyleId, 2);
  assert.equal(b1StyleId, 3);
  assert.equal(workbook.getStyle(1)?.alignment, null);
  assert.deepEqual(sheet.getAlignment("A1"), {
    horizontal: "center",
    wrapText: true,
  });
  assert.deepEqual(cell.alignment, {
    horizontal: "right",
  });

  const clearedStyleId = cell.setAlignment(null);

  assert.equal(clearedStyleId, 4);
  assert.equal(cell.alignment, null);

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(stylesXml, /<cellXfs count="5">/);
  assert.match(
    stylesXml,
    /<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"\/><\/xf>/,
  );
  assert.match(
    stylesXml,
    /<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"\/><\/xf>/,
  );
  assert.match(sheetXml, /<c r="A1" s="2" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B1" s="4" t="inlineStr"><is><t>World<\/t><\/is><\/c>/);
});

test("copyStyle APIs copy style indexes without changing target values", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" s="2"><f>SUM(1,2)</f><v>3</v></c>
      <c r="C1" t="inlineStr"><is><t>Tail</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.copyStyle("A1", "C1");
  sheet.copyStyle(1, 2, 2, 1);

  assert.equal(sheet.getStyleId("C1"), 1);
  assert.equal(sheet.getCell("C1"), "Tail");
  assert.equal(sheet.getStyleId(2, 1), 2);
  assert.equal(sheet.getCell("A2"), null);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="C1" s="1" t="inlineStr"><is><t>Tail<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="A2" s="2"\/>/);
});

test("style definition APIs read and clone workbook cellXfs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(workbook.getStyle(0), {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: null,
    applyFill: null,
    applyBorder: null,
    applyAlignment: null,
    applyProtection: null,
    alignment: null,
  });
  assert.deepEqual(sheet.getStyle("A1"), {
    numFmtId: 0,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: true,
    applyFill: null,
    applyBorder: null,
    applyAlignment: null,
    applyProtection: null,
    alignment: null,
  });
  assert.equal(sheet.cell(1, 1).style?.fontId, 1);

  const clonedBoldStyleId = workbook.cloneStyle(1, {
    numFmtId: 14,
    applyNumberFormat: true,
    applyAlignment: true,
    alignment: {
      horizontal: "center",
      wrapText: true,
    },
  });
  const clonedDefaultStyleId = sheet.cloneStyle(2, 2, {
    applyAlignment: true,
    alignment: {
      horizontal: "right",
    },
  });

  assert.equal(clonedBoldStyleId, 2);
  assert.equal(clonedDefaultStyleId, 3);
  assert.deepEqual(workbook.getStyle(clonedBoldStyleId), {
    numFmtId: 14,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: true,
    applyFont: true,
    applyFill: null,
    applyBorder: null,
    applyAlignment: true,
    applyProtection: null,
    alignment: {
      horizontal: "center",
      wrapText: true,
    },
  });
  assert.deepEqual(sheet.getStyle(2, 2), {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: null,
    applyFill: null,
    applyBorder: null,
    applyAlignment: true,
    applyProtection: null,
    alignment: {
      horizontal: "right",
    },
  });
  assert.equal(sheet.getStyleId("B2"), 3);

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");

  assert.match(stylesXml, /<cellXfs count="4">/);
  assert.match(
    stylesXml,
    /<xf numFmtId="14" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"\/><\/xf>/,
  );
  assert.match(
    stylesXml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"\/><\/xf>/,
  );
  assert.match(sheetXml, /<c r="B2" s="3"\/>/);
});

test("style writers tolerate single-quoted styles containers", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = convertEntriesToSingleQuotedAttributes(await loadFixtureEntries(fixtureDir), ["xl/styles.xml"]);
  const workbook = Workbook.fromEntries(entries);

  workbook.updateFont(0, { name: "Arial", size: 12 });
  const clonedFontId = workbook.cloneFont(1, {
    italic: true,
    color: {
      rgb: "FFFF0000",
    },
  });

  workbook.updateFill(1, {
    patternType: "solid",
    fgColor: {
      rgb: "FF00FF00",
    },
  });
  const clonedFillId = workbook.cloneFill(0, {
    patternType: "solid",
    fgColor: {
      rgb: "FFFF0000",
    },
  });

  workbook.updateBorder(0, {
    top: {
      style: "thin",
      color: {
        rgb: "FFFF0000",
      },
    },
  });
  const clonedBorderId = workbook.cloneBorder(0, {
    bottom: {
      style: "double",
      color: {
        rgb: "FF00FF00",
      },
    },
  });

  const numFmtId = workbook.cloneNumberFormat(0, "0.00");
  workbook.updateNumberFormat(numFmtId, "0.000");

  workbook.updateStyle(1, {
    applyAlignment: true,
    alignment: {
      horizontal: "center",
    },
  });
  const clonedStyleId = workbook.cloneStyle(1, {
    fontId: clonedFontId,
    fillId: clonedFillId,
    borderId: clonedBorderId,
    numFmtId,
    applyFont: true,
    applyFill: true,
    applyBorder: true,
    applyNumberFormat: true,
  });

  assert.equal(workbook.getFont(0)?.name, "Arial");
  assert.equal(workbook.getFont(0)?.size, 12);
  assert.equal(workbook.getFont(clonedFontId)?.italic, true);
  assert.deepEqual(workbook.getFont(clonedFontId)?.color, { rgb: "FFFF0000" });
  assert.equal(workbook.getFill(1)?.patternType, "solid");
  assert.deepEqual(workbook.getFill(1)?.fgColor, { rgb: "FF00FF00" });
  assert.deepEqual(workbook.getFill(clonedFillId)?.fgColor, { rgb: "FFFF0000" });
  assert.equal(workbook.getBorder(0)?.top?.style, "thin");
  assert.deepEqual(workbook.getBorder(0)?.top?.color, { rgb: "FFFF0000" });
  assert.equal(workbook.getBorder(clonedBorderId)?.bottom?.style, "double");
  assert.deepEqual(workbook.getBorder(clonedBorderId)?.bottom?.color, { rgb: "FF00FF00" });
  assert.deepEqual(workbook.getNumberFormat(numFmtId), {
    builtin: false,
    code: "0.000",
    numFmtId,
  });
  assert.equal(workbook.getStyle(1)?.alignment?.horizontal, "center");
  assert.equal(workbook.getStyle(clonedStyleId)?.fontId, clonedFontId);
  assert.equal(workbook.getStyle(clonedStyleId)?.fillId, clonedFillId);
  assert.equal(workbook.getStyle(clonedStyleId)?.borderId, clonedBorderId);
  assert.equal(workbook.getStyle(clonedStyleId)?.numFmtId, numFmtId);

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");
  assert.match(stylesXml, /<fonts count="3">/);
  assert.match(stylesXml, /<fills count="3">/);
  assert.match(stylesXml, /<borders count="2">/);
  assert.match(stylesXml, /<numFmts count="1">/);
  assert.match(stylesXml, /<cellXfs count="3">/);
});

test("workbook updateStyle patches existing cellXfs in place", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  workbook.updateStyle(1, {
    numFmtId: 14,
    applyNumberFormat: true,
    applyAlignment: true,
    alignment: {
      horizontal: "center",
      wrapText: true,
    },
  });

  assert.deepEqual(workbook.getStyle(1), {
    numFmtId: 14,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: true,
    applyFont: true,
    applyFill: null,
    applyBorder: null,
    applyAlignment: true,
    applyProtection: null,
    alignment: {
      horizontal: "center",
      wrapText: true,
    },
  });
  assert.equal(sheet.getStyle("A1")?.numFmtId, 14);
  assert.equal(sheet.getStyle("A1")?.alignment?.horizontal, "center");

  workbook.updateStyle(1, {
    applyAlignment: null,
    alignment: null,
  });

  assert.deepEqual(workbook.getStyle(1), {
    numFmtId: 14,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: true,
    applyFont: true,
    applyFill: null,
    applyBorder: null,
    applyAlignment: null,
    applyProtection: null,
    alignment: null,
  });

  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");

  assert.match(stylesXml, /<cellXfs count="2">/);
  assert.match(
    stylesXml,
    /<xf numFmtId="14" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"\/>/,
  );
});

test("row style id APIs read and write row-level style indexes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1" s="1" customFormat="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
    <row r="3">
      <c r="A3"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getRowStyleId(1), 1);
  assert.equal(sheet.getRowStyleId(2), null);
  assert.equal(sheet.getRowStyleId(3), null);

  sheet.setRowStyleId(1, 5);
  sheet.setRowStyleId(2, 6);
  sheet.setRowStyleId(3, 7);
  sheet.setRowStyleId(1, null);

  assert.equal(sheet.getRowStyleId(1), null);
  assert.equal(sheet.getRowStyleId(2), 6);
  assert.equal(sheet.getRowStyleId(3), 7);
  assert.equal(sheet.getCell("A3"), 3);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="1">\s*<c r="A1" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>\s*<\/row>/);
  assert.match(sheetXml, /<row r="2" s="6" customFormat="1"\/>/);
  assert.match(sheetXml, /<row r="3" s="7" customFormat="1">\s*<c r="A3"><v>3<\/v><\/c>\s*<\/row>/);
});

test("row layout APIs read and write hidden and height attributes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1" hidden="1" ht="25" customHeight="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
    <row r="3">
      <c r="A3"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getRowHidden(1), true);
  assert.equal(sheet.getRowHeight(1), 25);
  assert.equal(sheet.getRowHidden(2), false);
  assert.equal(sheet.getRowHeight(2), null);

  sheet.setRowHidden(1, false);
  sheet.setRowHeight(1, null);
  sheet.setRowHidden(2, true);
  sheet.setRowHeight(3, 30);

  assert.equal(sheet.getRowHidden(1), false);
  assert.equal(sheet.getRowHeight(1), null);
  assert.equal(sheet.getRowHidden(2), true);
  assert.equal(sheet.getRowHeight(3), 30);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="1">\s*<c r="A1" t="inlineStr"><is><t>Hello<\/t><\/is><\/c>\s*<\/row>/);
  assert.match(sheetXml, /<row r="2" hidden="1"\/>/);
  assert.match(sheetXml, /<row r="3" ht="30" customHeight="1">\s*<c r="A3"><v>3<\/v><\/c>\s*<\/row>/);
});

test("row and column style definition APIs read and clone style definitions", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getRowStyle(1), null);
  assert.equal(sheet.getColumnStyle("A"), null);

  sheet.setRowStyleId(2, 1);
  sheet.setColumnStyleId("B", 1);

  assert.deepEqual(sheet.getRowStyle(2), {
    numFmtId: 0,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: true,
    applyFill: null,
    applyBorder: null,
    applyAlignment: null,
    applyProtection: null,
    alignment: null,
  });
  assert.deepEqual(sheet.getColumnStyle("B"), {
    numFmtId: 0,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: true,
    applyFill: null,
    applyBorder: null,
    applyAlignment: null,
    applyProtection: null,
    alignment: null,
  });

  const rowStyleId = sheet.cloneRowStyle(3, {
    applyAlignment: true,
    alignment: {
      horizontal: "right",
    },
  });
  const columnStyleId = sheet.cloneColumnStyle("C", {
    applyAlignment: true,
    alignment: {
      horizontal: "center",
    },
  });
  const rowSetStyleId = sheet.setRowStyle(4, {
    applyAlignment: true,
    alignment: {
      horizontal: "left",
    },
  });
  const columnSetStyleId = sheet.setColumnStyle("D", {
    applyAlignment: true,
    alignment: {
      horizontal: "justify",
    },
  });

  assert.equal(rowStyleId, 2);
  assert.equal(columnStyleId, 3);
  assert.equal(rowSetStyleId, 4);
  assert.equal(columnSetStyleId, 5);
  assert.equal(sheet.getRowStyleId(3), 2);
  assert.equal(sheet.getRowStyleId(4), 4);
  assert.equal(sheet.getColumnStyleId("C"), 3);
  assert.equal(sheet.getColumnStyleId("D"), 5);
  assert.deepEqual(sheet.getRowStyle(3), {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: null,
    applyFill: null,
    applyBorder: null,
    applyAlignment: true,
    applyProtection: null,
    alignment: {
      horizontal: "right",
    },
  });
  assert.deepEqual(sheet.getColumnStyle("C"), {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: null,
    applyFill: null,
    applyBorder: null,
    applyAlignment: true,
    applyProtection: null,
    alignment: {
      horizontal: "center",
    },
  });
  assert.deepEqual(sheet.getRowStyle(4), {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: null,
    applyFill: null,
    applyBorder: null,
    applyAlignment: true,
    applyProtection: null,
    alignment: {
      horizontal: "left",
    },
  });
  assert.deepEqual(sheet.getColumnStyle("D"), {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    quotePrefix: null,
    pivotButton: null,
    applyNumberFormat: null,
    applyFont: null,
    applyFill: null,
    applyBorder: null,
    applyAlignment: true,
    applyProtection: null,
    alignment: {
      horizontal: "justify",
    },
  });

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const stylesXml = entryText(workbook.toEntries(), "xl/styles.xml");

  assert.match(sheetXml, /<row r="2" s="1" customFormat="1"\/>/);
  assert.match(sheetXml, /<row r="3" s="2" customFormat="1"\/>/);
  assert.match(sheetXml, /<row r="4" s="4" customFormat="1"\/>/);
  assert.match(sheetXml, /<cols><col min="2" max="2" style="1"\/><col min="3" max="3" style="3"\/><col min="4" max="4" style="5"\/><\/cols>/);
  assert.match(stylesXml, /<cellXfs count="6">/);
});

test("column style id APIs read, write, and shift with column edits", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="2" style="1"/>
    <col min="4" max="4" style="3" hidden="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="D1"><v>4</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getColumnStyleId("A"), 1);
  assert.equal(sheet.getColumnStyleId(2), 1);
  assert.equal(sheet.getColumnStyleId("C"), null);
  assert.equal(sheet.getColumnStyleId("D"), 3);

  sheet.setColumnStyleId("B", 5);
  sheet.setColumnStyleId(3, 7);
  sheet.setColumnStyleId("D", null);

  assert.equal(sheet.getColumnStyleId("A"), 1);
  assert.equal(sheet.getColumnStyleId("B"), 5);
  assert.equal(sheet.getColumnStyleId("C"), 7);
  assert.equal(sheet.getColumnStyleId("D"), null);

  sheet.insertColumn("C");

  assert.equal(sheet.getColumnStyleId("B"), 5);
  assert.equal(sheet.getColumnStyleId("C"), null);
  assert.equal(sheet.getColumnStyleId("D"), 7);
  assert.equal(sheet.getColumnStyleId("E"), null);

  sheet.deleteColumn("B");

  assert.equal(sheet.getColumnStyleId("A"), 1);
  assert.equal(sheet.getColumnStyleId("B"), null);
  assert.equal(sheet.getColumnStyleId("C"), 7);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<cols><col min="1" max="1" style="1"\/><col min="3" max="3" style="7"\/><col min="4" max="4" hidden="1"\/><\/cols>/);
});

test("column layout APIs read and write hidden and width attributes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="18.5" customWidth="1"/>
    <col min="3" max="3" hidden="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getColumnWidth("A"), 18.5);
  assert.equal(sheet.getColumnWidth("B"), null);
  assert.equal(sheet.getColumnHidden("C"), true);
  assert.equal(sheet.getColumnHidden("D"), false);

  sheet.setColumnWidth("A", null);
  sheet.setColumnWidth("B", 24);
  sheet.setColumnHidden("C", false);
  sheet.setColumnHidden("D", true);

  assert.equal(sheet.getColumnWidth("A"), null);
  assert.equal(sheet.getColumnWidth("B"), 24);
  assert.equal(sheet.getColumnHidden("C"), false);
  assert.equal(sheet.getColumnHidden("D"), true);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<cols><col min="2" max="2" width="24" customWidth="1"\/><col min="4" max="4" hidden="1"\/><\/cols>/);
});

test("range APIs read and write rectangular values", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getRangeRef(), "A1");
  assert.deepEqual(sheet.getRange("A1:B2"), [["Hello", null], [null, null]]);

  sheet.setRange("B2", [
    [1, 2],
    [3, 4],
  ]);

  assert.equal(sheet.getRangeRef(), "A1:C3");
  assert.deepEqual(sheet.getRange("A1:C3"), [
    ["Hello", null, null],
    [null, 1, 2],
    [null, 3, 4],
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="B2"><v>1<\/v><\/c><c r="C2"><v>2<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="B3"><v>3<\/v><\/c><c r="C3"><v>4<\/v><\/c><\/row>/);
});

test("range style APIs apply and copy styles across rectangles", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRangeStyle("A1:B2", {
    applyAlignment: true,
    alignment: { horizontal: "center" },
  });
  sheet.setRangeNumberFormat("A1:B2", "0.00%");
  sheet.setRangeBackgroundColor("A1:B2", "FFFF0000");
  sheet.copyRangeStyle("A1:B2", "C3:D4");

  assert.equal(sheet.getBackgroundColor("B2"), "FFFF0000");
  assert.equal(sheet.getBackgroundColor("D4"), "FFFF0000");
  assert.equal(sheet.getNumberFormat("B2")?.code, "0.00%");
  assert.equal(sheet.getNumberFormat("D4")?.code, "0.00%");
  assert.deepEqual(sheet.getAlignment("B2"), { horizontal: "center" });
  assert.deepEqual(sheet.getAlignment("D4"), { horizontal: "center" });
});

test("sheet rowCount and columnCount track the used bounds", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData></sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.rowCount, 0);
  assert.equal(sheet.columnCount, 0);

  sheet.setCell("C5", 1);

  assert.equal(sheet.rowCount, 5);
  assert.equal(sheet.columnCount, 3);
  assert.equal(sheet.getRangeRef(), "C5");

  sheet.setCell("A1", "Top");

  assert.equal(sheet.rowCount, 5);
  assert.equal(sheet.columnCount, 3);
  assert.equal(sheet.getRangeRef(), "A1:C5");

  sheet.deleteColumn("B");

  assert.equal(sheet.rowCount, 5);
  assert.equal(sheet.columnCount, 2);
  assert.equal(sheet.getRangeRef(), "A1:B5");
});

test("insertColumn shifts cell addresses, formulas, and merged ranges together", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C2"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><f>SUM(A1:B1)</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>Sheet1!B1</f><v>2</v></c>
      <c r="B2"><v>4</v></c>
      <c r="C2"><v>5</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.insertColumn("B");

  assert.equal(sheet.getCell("A1"), 1);
  assert.equal(sheet.getCell("B1"), null);
  assert.equal(sheet.getCell("C1"), 2);
  assert.equal(sheet.getCell("D1"), 3);
  assert.equal(sheet.getFormula("D1"), "SUM(A1:C1)");
  assert.equal(sheet.getFormula("A2"), "Sheet1!C1");
  assert.deepEqual(sheet.getMergedRanges(), ["C2:D2"]);
  assert.equal(sheet.getRangeRef(), "A1:D2");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="C1"><v>2<\/v><\/c>/);
  assert.match(sheetXml, /<c r="D1"><f>SUM\(A1:C1\)<\/f><v>3<\/v><\/c>/);
  assert.match(sheetXml, /<c r="A2"><f>Sheet1!C1<\/f><v>2<\/v><\/c>/);
  assert.match(sheetXml, /<mergeCell ref="C2:D2"\/>/);
  assert.match(sheetXml, /<dimension ref="A1:D2"\/>/);
});

test("insertRow shifts cell addresses, formulas, and merged ranges together", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>SUM(A2:B2)</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>2</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><f>Sheet1!A2</f><v>2</v></c>
      <c r="B3"><v>5</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A2:B3"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.insertRow(2);

  assert.equal(sheet.getCell("A1"), 1);
  assert.equal(sheet.getCell("A2"), null);
  assert.equal(sheet.getCell("A3"), 2);
  assert.equal(sheet.getCell("A4"), 2);
  assert.equal(sheet.getFormula("B1"), "SUM(A3:B3)");
  assert.equal(sheet.getFormula("A4"), "Sheet1!A3");
  assert.deepEqual(sheet.getMergedRanges(), ["A3:B4"]);
  assert.equal(sheet.getRangeRef(), "A1:B4");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="3">[\s\S]*<c r="A3"><v>2<\/v><\/c>[\s\S]*<c r="B3"><v>4<\/v><\/c>[\s\S]*<\/row>/);
  assert.match(sheetXml, /<row r="4">[\s\S]*<c r="A4"><f>Sheet1!A3<\/f><v>2<\/v><\/c>[\s\S]*<\/row>/);
  assert.match(sheetXml, /<c r="B1"><f>SUM\(A3:B3\)<\/f><v>3<\/v><\/c>/);
  assert.match(sheetXml, /<mergeCell ref="A3:B4"\/>/);
  assert.match(sheetXml, /<dimension ref="A1:B4"\/>/);
});

test("insertColumn updates worksheet ref attributes and defined names", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">$A$1:$C$4</definedName>
    <definedName name="DataRange">Sheet1!$B$2:$C$4</definedName>
  </definedNames>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2:C2"/></sheetView></sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>4</v></c>
      <c r="B2"><v>5</v></c>
      <c r="C2"><v>6</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:C4">
    <filterColumn colId="0"><filters><filter val="Alpha"/></filters></filterColumn>
    <filterColumn colId="1"><filters><filter val="Beta"/></filters></filterColumn>
  </autoFilter>
  <sortState ref="A2:C4">
    <sortCondition ref="B2:B4"/>
    <sortCondition ref="C2:C4" descending="1"/>
  </sortState>
  <conditionalFormatting sqref="B2:C4"><cfRule type="expression" priority="1"><formula>B2&gt;0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="whole" sqref="A2:B4"/></dataValidations>
  <hyperlinks><hyperlink ref="C2" location="#Sheet1!A1"/></hyperlinks>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.insertColumn("B");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");

  assert.match(sheetXml, /<selection activeCell="C2" sqref="C2:D2"\/>/);
  assert.match(sheetXml, /<autoFilter ref="A1:D4">/);
  assert.match(sheetXml, /<filterColumn colId="0"><filters><filter val="Alpha"\/><\/filters><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="2"><filters><filter val="Beta"\/><\/filters><\/filterColumn>/);
  assert.match(sheetXml, /<sortState ref="A2:D4">/);
  assert.match(sheetXml, /<sortCondition ref="C2:C4"\/>/);
  assert.match(sheetXml, /<sortCondition ref="D2:D4" descending="1"\/>/);
  assert.match(sheetXml, /<conditionalFormatting sqref="C2:D4">/);
  assert.match(sheetXml, /<dataValidations count="1"><dataValidation type="whole" sqref="A2:C4"\/><\/dataValidations>/);
  assert.match(sheetXml, /<hyperlinks><hyperlink ref="D2" location="#Sheet1!A1"\/><\/hyperlinks>/);
  assert.match(workbookXml, /<definedName name="_xlnm.Print_Area" localSheetId="0">\$A\$1:\$D\$4<\/definedName>/);
  assert.match(workbookXml, /<definedName name="DataRange">Sheet1!\$C\$2:\$D\$4<\/definedName>/);
});

test("insertColumn updates formulas in other sheets that reference the edited sheet", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSecondSheet(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>SUM(Sheet1!A1:B1)</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>Sheet1!B1</f><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet1 = workbook.getSheet("Sheet1");
  const sheet2 = workbook.getSheet("Sheet2");

  sheet1.insertColumn("B");

  assert.equal(sheet2.getFormula("A1"), "SUM(Sheet1!A1:C1)");
  assert.equal(sheet2.getFormula("A2"), "Sheet1!C1");

  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");
  assert.match(sheet2Xml, /<c r="A1"><f>SUM\(Sheet1!A1:C1\)<\/f><v>3<\/v><\/c>/);
  assert.match(sheet2Xml, /<c r="A2"><f>Sheet1!C1<\/f><v>2<\/v><\/c>/);
});

test("deleteColumn shifts cells, shrinks ranges, and emits #REF! for deleted refs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D2"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><f>SUM(A1:D1)</f><v>10</v></c>
      <c r="D1"><f>B1</f><v>2</v></c>
    </row>
    <row r="2">
      <c r="B2"><v>4</v></c>
      <c r="C2"><v>5</v></c>
      <c r="D2"><v>6</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="B2:D2"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteColumn("B");

  assert.equal(sheet.getCell("A1"), 1);
  assert.equal(sheet.getCell("B1"), 10);
  assert.equal(sheet.getFormula("B1"), "SUM(A1:C1)");
  assert.equal(sheet.getCell("C1"), 2);
  assert.equal(sheet.getFormula("C1"), "#REF!");
  assert.deepEqual(sheet.getMergedRanges(), ["B2:C2"]);
  assert.equal(sheet.getRangeRef(), "A1:C2");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<c r="B2"><v>4<\/v><\/c>/);
  assert.match(sheetXml, /<c r="B1"><f>SUM\(A1:C1\)<\/f><v>10<\/v><\/c>/);
  assert.match(sheetXml, /<c r="C1"><f>#REF!<\/f><v>2<\/v><\/c>/);
  assert.match(sheetXml, /<mergeCell ref="B2:C2"\/>/);
});

test("deleteColumn updates worksheet autoFilter column ids and sort conditions", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><v>3</v></c>
      <c r="D1"><v>4</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>5</v></c>
      <c r="B2"><v>6</v></c>
      <c r="C2"><v>7</v></c>
      <c r="D2"><v>8</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>9</v></c>
      <c r="B3"><v>10</v></c>
      <c r="C3"><v>11</v></c>
      <c r="D3"><v>12</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:D3">
    <filterColumn colId="0"><filters><filter val="A"/></filters></filterColumn>
    <filterColumn colId="1"><filters><filter val="B"/></filters></filterColumn>
    <filterColumn colId="3"><filters><filter val="D"/></filters></filterColumn>
  </autoFilter>
  <sortState ref="A2:D3">
    <sortCondition ref="B2:B3"/>
    <sortCondition ref="D2:D3" descending="1"/>
  </sortState>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteColumn("B");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<autoFilter ref="A1:C3">/);
  assert.match(sheetXml, /<filterColumn colId="0"><filters><filter val="A"\/><\/filters><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="2"><filters><filter val="D"\/><\/filters><\/filterColumn>/);
  assert.doesNotMatch(sheetXml, /<filterColumn colId="1"><filters><filter val="B"\/><\/filters><\/filterColumn>/);
  assert.match(sheetXml, /<sortState ref="A2:C3"><sortCondition ref="C2:C3" descending="1"\/><\/sortState>/);
});

test("deleteRow updates formulas in other sheets that reference the edited sheet", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSecondSheet(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>5</v></c>
      <c r="B3"><v>6</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>7</v></c>
      <c r="B4"><v>8</v></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>SUM(Sheet1!A1:B4)</f><v>36</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>Sheet1!A2</f><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet1 = workbook.getSheet("Sheet1");
  const sheet2 = workbook.getSheet("Sheet2");

  sheet1.deleteRow(2);

  assert.equal(sheet2.getFormula("A1"), "SUM(Sheet1!A1:B3)");
  assert.equal(sheet2.getFormula("A2"), "#REF!");

  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");
  assert.match(sheet2Xml, /<c r="A1"><f>SUM\(Sheet1!A1:B3\)<\/f><v>36<\/v><\/c>/);
  assert.match(sheet2Xml, /<c r="A2"><f>#REF!<\/f><v>3<\/v><\/c>/);
});

test("deleteRow shifts cells, shrinks ranges, and emits #REF! for deleted refs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B4"/>
  <sheetData>
    <row r="1">
      <c r="A1"><f>SUM(A1:B4)</f><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>2</v></c>
    </row>
    <row r="3">
      <c r="A3"><f>A2</f><v>2</v></c>
      <c r="B3"><v>3</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>4</v></c>
      <c r="B4"><v>5</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A2:B4"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteRow(2);

  assert.equal(sheet.getFormula("A1"), "SUM(A1:B3)");
  assert.equal(sheet.getFormula("A2"), "#REF!");
  assert.equal(sheet.getCell("A3"), 4);
  assert.deepEqual(sheet.getMergedRanges(), ["A2:B3"]);
  assert.equal(sheet.getRangeRef(), "A1:B3");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<row r="4">/);
  assert.match(sheetXml, /<row r="2">[\s\S]*<c r="A2"><f>#REF!<\/f><v>2<\/v><\/c>[\s\S]*<c r="B2"><v>3<\/v><\/c>[\s\S]*<\/row>/);
  assert.match(sheetXml, /<mergeCell ref="A2:B3"\/>/);
});

test("deleteRow updates worksheet ref attributes and defined names", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">$A$1:$C$4</definedName>
    <definedName name="DataRange">Sheet1!$B$2:$C$4</definedName>
  </definedNames>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="B3" sqref="B3:C3"/></sheetView></sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>4</v></c>
      <c r="B2"><v>5</v></c>
      <c r="C2"><v>6</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>7</v></c>
      <c r="B3"><v>8</v></c>
      <c r="C3"><v>9</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>10</v></c>
      <c r="B4"><v>11</v></c>
      <c r="C4"><v>12</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:C4"/>
  <sortState ref="A2:C4">
    <sortCondition ref="B2:B4"/>
  </sortState>
  <conditionalFormatting sqref="B2:C4"><cfRule type="expression" priority="1"><formula>B2&gt;0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="whole" sqref="A2:B4"/></dataValidations>
  <hyperlinks><hyperlink ref="C3" location="#Sheet1!A1"/></hyperlinks>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteRow(2);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");

  assert.match(sheetXml, /<selection activeCell="B2" sqref="B2:C2"\/>/);
  assert.match(sheetXml, /<autoFilter ref="A1:C3"\/>/);
  assert.match(sheetXml, /<sortState ref="A2:C3"><sortCondition ref="B2:B3"\/><\/sortState>/);
  assert.match(sheetXml, /<conditionalFormatting sqref="B2:C3">/);
  assert.match(sheetXml, /<dataValidations count="1"><dataValidation type="whole" sqref="A2:B3"\/><\/dataValidations>/);
  assert.match(sheetXml, /<hyperlinks><hyperlink ref="C2" location="#Sheet1!A1"\/><\/hyperlinks>/);
  assert.match(workbookXml, /<definedName name="_xlnm.Print_Area" localSheetId="0">\$A\$1:\$C\$3<\/definedName>/);
  assert.match(workbookXml, /<definedName name="DataRange">Sheet1!\$B\$2:\$C\$3<\/definedName>/);
});

test("sheet getTables reads existing tables and insertColumn updates table refs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSheetTable(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>5</v></c>
      <c r="B3"><v>6</v></c>
    </row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3" totalsRowShown="0">
  <autoFilter ref="A1:B3"/>
  <tableColumns count="2">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
  </tableColumns>
</table>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getTables(), [
    { name: "Sales", displayName: "Sales", range: "A1:B3", path: "xl/tables/table1.xml" },
  ]);

  sheet.insertColumn("B");

  assert.deepEqual(sheet.getTables(), [
    { name: "Sales", displayName: "Sales", range: "A1:C3", path: "xl/tables/table1.xml" },
  ]);

  const tableXml = entryText(workbook.toEntries(), "xl/tables/table1.xml");
  assert.match(tableXml, /<table [^>]*ref="A1:C3"[^>]*>/);
  assert.match(tableXml, /<autoFilter ref="A1:C3"\/>/);
});

test("deleteRow removes table parts when the full table range is deleted", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSheetTable(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>5</v></c>
      <c r="B3"><v>6</v></c>
    </row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3" totalsRowShown="0">
  <autoFilter ref="A1:B3"/>
  <tableColumns count="2">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
  </tableColumns>
</table>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteRow(1, 3);

  assert.deepEqual(sheet.getTables(), []);
  assert.equal(workbook.listEntries().includes("xl/tables/table1.xml"), false);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  const contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  assert.doesNotMatch(sheetXml, /<tableParts\b/);
  assert.doesNotMatch(relsXml, /relationships\/table/);
  assert.doesNotMatch(contentTypesXml, /spreadsheetml\.table\+xml/);
});

test("sheet addTable and removeTable manage package parts", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.setRow(2, ["Alice", 98]);
  sheet.setRow(3, ["Bob", 87]);

  const table = sheet.addTable("A1:B3", { name: "Scores" });

  assert.deepEqual(table, {
    name: "Scores",
    displayName: "Scores",
    range: "A1:B3",
    path: "xl/tables/table1.xml",
  });
  assert.deepEqual(sheet.getTables(), [table]);

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  let relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  let contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  let tableXml = entryText(workbook.toEntries(), "xl/tables/table1.xml");

  assert.match(sheetXml, /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts>/);
  assert.match(relsXml, /<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/table" Target="\.\.\/tables\/table1\.xml"\/>/);
  assert.match(contentTypesXml, /<Override PartName="\/xl\/tables\/table1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.table\+xml"\/>/);
  assert.match(tableXml, /<table [^>]*name="Scores" displayName="Scores" ref="A1:B3"[^>]*>/);
  assert.match(tableXml, /<tableColumn id="1" name="Name"\/>/);
  assert.match(tableXml, /<tableColumn id="2" name="Score"\/>/);

  sheet.removeTable("Scores");

  assert.deepEqual(sheet.getTables(), []);
  assert.equal(workbook.listEntries().includes("xl/tables/table1.xml"), false);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");

  assert.doesNotMatch(sheetXml, /<tableParts\b/);
  assert.doesNotMatch(relsXml, /relationships\/table/);
  assert.doesNotMatch(contentTypesXml, /spreadsheetml\.table\+xml/);
});

test("sheet table metadata tolerates single-quoted table and tableParts tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = convertEntriesToSingleQuotedAttributes(
    withSheetTable(
      replaceEntryText(
        await loadFixtureEntries(fixtureDir),
        "xl/worksheets/sheet1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`,
      ),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B2" totalsRowShown="0">
  <autoFilter ref="A1:B2"/>
  <tableColumns count="2">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
  </tableColumns>
</table>`,
    ),
    ["xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels", "xl/tables/table1.xml"],
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getTables(), [
    { name: "Sales", displayName: "Sales", range: "A1:B2", path: "xl/tables/table1.xml" },
  ]);

  sheet.insertColumn("B");

  assert.deepEqual(sheet.getTables(), [
    { name: "Sales", displayName: "Sales", range: "A1:C2", path: "xl/tables/table1.xml" },
  ]);

  let tableXml = entryText(workbook.toEntries(), "xl/tables/table1.xml");
  assert.match(tableXml, /<table [^>]*ref="A1:C2"[^>]*>/);
  assert.match(tableXml, /<autoFilter ref="A1:C2"\/>/);

  sheet.removeTable("Sales");

  tableXml = workbook.listEntries().includes("xl/tables/table1.xml") ? entryText(workbook.toEntries(), "xl/tables/table1.xml") : "";
  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.deepEqual(sheet.getTables(), []);
  assert.equal(tableXml, "");
  assert.doesNotMatch(sheetXml, /<tableParts\b/);
});

test("row APIs read sparse rows and write from a column offset", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getRow(1), ["Hello"]);
  assert.deepEqual(sheet.getRow(4), []);

  sheet.setRow(4, ["Name", null, "Score"], 2);

  assert.deepEqual(sheet.getRow(4), [null, "Name", null, "Score"]);
  assert.equal(sheet.getRangeRef(), "A1:D4");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="4"><c r="B4" t="inlineStr"><is><t>Name<\/t><\/is><\/c><c r="C4"\/><c r="D4" t="inlineStr"><is><t>Score<\/t><\/is><\/c><\/row>/);
  assert.match(sheetXml, /<dimension ref="A1:D4"\/>/);
});

test("column APIs read sparse columns and write from a row offset", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getColumn("A"), ["Hello"]);
  assert.deepEqual(sheet.getColumn(3), []);

  sheet.setColumn("C", ["Q1", null, "Q3"], 2);

  assert.deepEqual(sheet.getColumn("C"), [null, "Q1", null, "Q3"]);
  assert.equal(sheet.getRangeRef(), "A1:C4");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="C2" t="inlineStr"><is><t>Q1<\/t><\/is><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="C3"\/><\/row>/);
  assert.match(sheetXml, /<row r="4"><c r="C4" t="inlineStr"><is><t>Q3<\/t><\/is><\/c><\/row>/);
});

test("entry APIs iterate existing worksheet cells without dense null scans", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(2, ["Name", null, 98]);
  sheet.setCell("C4", "Tail");

  assert.deepEqual(
    summarizeCellEntries(sheet.getRowEntries(2)),
    [
      { address: "A2", rowNumber: 2, columnNumber: 1, type: "string", value: "Name" },
      { address: "C2", rowNumber: 2, columnNumber: 3, type: "number", value: 98 },
    ],
  );
  assert.deepEqual(
    summarizeCellEntries(sheet.getPhysicalRowEntries(2)),
    [
      { address: "A2", rowNumber: 2, columnNumber: 1, type: "string", value: "Name" },
      { address: "B2", rowNumber: 2, columnNumber: 2, type: "blank", value: null },
      { address: "C2", rowNumber: 2, columnNumber: 3, type: "number", value: 98 },
    ],
  );
  assert.deepEqual(
    summarizeCellEntries(sheet.getColumnEntries("C")),
    [
      { address: "C2", rowNumber: 2, columnNumber: 3, type: "number", value: 98 },
      { address: "C4", rowNumber: 4, columnNumber: 3, type: "string", value: "Tail" },
    ],
  );
  assert.deepEqual(
    summarizeCellEntries(sheet.getPhysicalColumnEntries("C")),
    [
      { address: "C2", rowNumber: 2, columnNumber: 3, type: "number", value: 98 },
      { address: "C4", rowNumber: 4, columnNumber: 3, type: "string", value: "Tail" },
    ],
  );

  const allEntries = summarizeCellEntries(sheet.getCellEntries());
  assert.deepEqual(allEntries, [
    { address: "A1", rowNumber: 1, columnNumber: 1, type: "string", value: "Hello" },
    { address: "A2", rowNumber: 2, columnNumber: 1, type: "string", value: "Name" },
    { address: "C2", rowNumber: 2, columnNumber: 3, type: "number", value: 98 },
    { address: "C4", rowNumber: 4, columnNumber: 3, type: "string", value: "Tail" },
  ]);
  assert.deepEqual(summarizeCellEntries(sheet.getPhysicalCellEntries()), [
    { address: "A1", rowNumber: 1, columnNumber: 1, type: "string", value: "Hello" },
    { address: "A2", rowNumber: 2, columnNumber: 1, type: "string", value: "Name" },
    { address: "B2", rowNumber: 2, columnNumber: 2, type: "blank", value: null },
    { address: "C2", rowNumber: 2, columnNumber: 3, type: "number", value: 98 },
    { address: "C4", rowNumber: 4, columnNumber: 3, type: "string", value: "Tail" },
  ]);
  assert.deepEqual(summarizeCellEntries([...sheet.iterCellEntries()]), allEntries);
  assert.deepEqual(summarizeCellEntries([...sheet.iterPhysicalCellEntries()]), [
    { address: "A1", rowNumber: 1, columnNumber: 1, type: "string", value: "Hello" },
    { address: "A2", rowNumber: 2, columnNumber: 1, type: "string", value: "Name" },
    { address: "B2", rowNumber: 2, columnNumber: 2, type: "blank", value: null },
    { address: "C2", rowNumber: 2, columnNumber: 3, type: "number", value: 98 },
    { address: "C4", rowNumber: 4, columnNumber: 3, type: "string", value: "Tail" },
  ]);
});

test("deleteCell removes worksheet cell nodes instead of leaving blank placeholders", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("C4", "Tail");
  sheet.deleteCell("A1");
  sheet.deleteCell(4, 3);

  assert.equal(sheet.getCell("A1"), null);
  assert.equal(sheet.getCell(4, 3), null);
  assert.equal(sheet.getRangeRef(), null);
  assert.deepEqual(sheet.getCellEntries(), []);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<c r="A1"/);
  assert.doesNotMatch(sheetXml, /<c r="C4"/);
  assert.match(sheetXml, /<sheetData>\s*<row r="1"><\/row>\s*<row r="4"><\/row>\s*<\/sheetData>/);
});

test("header APIs read and write header rows", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setHeaders(["Name", "Score"]);

  assert.deepEqual(sheet.getHeaders(), ["Name", "Score"]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(
    sheetXml,
    /<row r="1">[\s\S]*<c r="A1" t="inlineStr" s="1"><is><t>Name<\/t><\/is><\/c>[\s\S]*<c r="B1" t="inlineStr"><is><t>Score<\/t><\/is><\/c>[\s\S]*<\/row>/,
  );
});

test("append row APIs add rows at the sheet tail", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  const firstRow = sheet.appendRow(["Tail", 1], 2);
  const nextRows = sheet.appendRows([
    ["Tail-2", 2],
    ["Tail-3", 3],
  ], 2);

  assert.equal(firstRow, 2);
  assert.deepEqual(nextRows, [3, 4]);
  assert.deepEqual(sheet.getRow(4), [null, "Tail-3", 3]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="B2" t="inlineStr"><is><t>Tail<\/t><\/is><\/c><c r="C2"><v>1<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="4"><c r="B4" t="inlineStr"><is><t>Tail-3<\/t><\/is><\/c><c r="C4"><v>3<\/v><\/c><\/row>/);
});

test("record APIs map rows by header cells", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.setRow(2, ["Alice", 98]);
  sheet.setRow(4, ["Bob", 87]);

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  sheet.addRecord({ Name: "Cara", Score: 91 });

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
    { Name: "Cara", Score: 91 },
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="5"><c r="A5" t="inlineStr"><is><t>Cara<\/t><\/is><\/c><c r="B5"><v>91<\/v><\/c><\/row>/);
});

test("record APIs initialize headers on a blank created sheet", async () => {
  const workbook = Workbook.create("Config");
  const sheet = workbook.getSheet("Config");

  sheet.addRecord({ Name: "Alice", Score: 98 });

  assert.deepEqual(sheet.getHeaders(), ["Name", "Score"]);
  assert.deepEqual(sheet.getRecords(), [{ Name: "Alice", Score: 98 }]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="1"><c r="A1" t="inlineStr"><is><t>Name<\/t><\/is><\/c><c r="B1" t="inlineStr"><is><t>Score<\/t><\/is><\/c><\/row>/);
  assert.match(sheetXml, /<row r="2"><c r="A2" t="inlineStr"><is><t>Alice<\/t><\/is><\/c><c r="B2"><v>98<\/v><\/c><\/row>/);
});

test("record APIs can infer headers from multiple records on a blank sheet", async () => {
  const workbook = Workbook.create("Config");
  const sheet = workbook.getSheet("Config");

  sheet.setRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", City: "Shanghai" },
  ]);

  assert.deepEqual(sheet.getHeaders(), ["Name", "Score", "City"]);
  assert.deepEqual(sheet.getRecords(), [
    { Name: "Alice", Score: 98, City: null },
    { Name: "Bob", Score: null, City: "Shanghai" },
  ]);
});

test("record APIs can append multiple records in order", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="A2" t="inlineStr"><is><t>Alice<\/t><\/is><\/c><c r="B2"><v>98<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="A3" t="inlineStr"><is><t>Bob<\/t><\/is><\/c><c r="B3"><v>87<\/v><\/c><\/row>/);
});

test("record APIs can read and update a specific record row", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.setRow(2, ["Alice", 98]);

  assert.deepEqual(sheet.getRecord(2), { Name: "Alice", Score: 98 });

  sheet.setRecord(2, { Name: "Alicia", Score: 99 });

  assert.deepEqual(sheet.getRecord(2), { Name: "Alicia", Score: 99 });

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="A2" t="inlineStr"><is><t>Alicia<\/t><\/is><\/c><c r="B2"><v>99<\/v><\/c><\/row>/);
});

test("record APIs can replace the full record set", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
    { Name: "Cara", Score: 91 },
  ]);

  sheet.setRecords([
    { Name: "Zoe", Score: 100 },
    { Name: "Yan" },
  ]);

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Zoe", Score: 100 },
    { Name: "Yan", Score: null },
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="A2" t="inlineStr"><is><t>Zoe<\/t><\/is><\/c><c r="B2"><v>100<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="A3" t="inlineStr"><is><t>Yan<\/t><\/is><\/c><c r="B3"\/><\/row>/);
  assert.doesNotMatch(sheetXml, /<row r="4">/);
});

test("record key APIs can get, update, upsert, and delete by field", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["id", "name", "note"]);
  sheet.addRecords([
    { id: 1001, name: "Alpha", note: "first" },
    { id: 1002, name: "Beta", note: "second" },
  ]);

  assert.deepEqual(sheet.getRecordBy("id", 1002), { id: 1002, name: "Beta", note: "second" });

  const patchedRow = sheet.updateRecordBy("id", 1001, { note: "first-2" });
  const updatedRow = sheet.upsertRecord("id", { id: 1002, name: "Beta-2" });
  const insertedRow = sheet.upsertRecord("id", { id: 1003, name: "Gamma" });

  assert.deepEqual(patchedRow, {
    record: { note: "first-2" },
    row: 2,
    updated: true,
  });
  assert.deepEqual(updatedRow, {
    inserted: false,
    record: { id: 1002, name: "Beta-2" },
    row: 3,
  });
  assert.deepEqual(insertedRow, {
    inserted: true,
    record: { id: 1003, name: "Gamma" },
    row: 4,
  });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first-2" },
    { id: 1002, name: "Beta-2", note: null },
    { id: 1003, name: "Gamma", note: null },
  ]);

  assert.equal(sheet.deleteRecordBy("id", 1001), true);
  assert.equal(sheet.deleteRecordBy("id", 9999), false);
  assert.deepEqual(sheet.getRecords(), [
    { id: 1002, name: "Beta-2", note: null },
    { id: 1003, name: "Gamma", note: null },
  ]);

  assert.deepEqual(sheet.findRecordBy("id", 1002), { id: 1002, name: "Beta-2", note: null });
  assert.equal(sheet.removeRecordBy("id", 1003), true);
  assert.deepEqual(sheet.getRecords(), [{ id: 1002, name: "Beta-2", note: null }]);
});

test("sheet JSON helpers roundtrip header-mapped records", async () => {
  const workbook = Workbook.create("Data");
  const sheet = workbook.getSheet("Data");

  sheet.fromJson([
    { id: 1001, name: "Alpha", enabled: true },
    { id: 1002, name: "Beta", enabled: false },
  ]);

  assert.deepEqual(sheet.getHeaders(), ["id", "name", "enabled"]);
  assert.deepEqual(sheet.toJson(), [
    { id: 1001, name: "Alpha", enabled: true },
    { id: 1002, name: "Beta", enabled: false },
  ]);
});

test("sheet CSV helpers import and export header-mapped records", async () => {
  const workbook = Workbook.create("Data");
  const sheet = workbook.getSheet("Data");

  sheet.fromCsv('id,name,notes\n1001,Alpha,"A, B"\n1002,Beta,"line 1\nline 2"');

  assert.deepEqual(sheet.getHeaders(), ["id", "name", "notes"]);
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", notes: "A, B" },
    { id: 1002, name: "Beta", notes: "line 1\nline 2" },
  ]);

  assert.equal(
    sheet.toCsv(),
    'id,name,notes\n1001,Alpha,"A, B"\n1002,Beta,"line 1\nline 2"',
  );
});

test("sheet CSV helpers support export and import options", async () => {
  const workbook = Workbook.create("Data");
  const sheet = workbook.getSheet("Data");

  sheet.fromCsv(" id , name \n 1001 , Alpha \n", {
    trimHeaders: true,
    trimValues: true,
  });
  assert.deepEqual(sheet.getHeaders(), ["id", "name"]);
  assert.deepEqual(sheet.getRecords(), [{ id: 1001, name: "Alpha" }]);

  assert.equal(
    sheet.toCsv({ includeHeaders: false, lineEnding: "\r\n" }),
    "1001,Alpha",
  );
});

test("sheet JSON and CSV helpers support append, update, and upsert semantics", async () => {
  const workbook = Workbook.create("Data");
  const sheet = workbook.getSheet("Data");

  sheet.fromJson(
    [{ name: "Alpha", id: 1001, note: "first" }],
    { headerOrder: ["id", "name", "note"] },
  );
  assert.deepEqual(sheet.getHeaders(), ["id", "name", "note"]);
  assert.deepEqual(sheet.getRecords(), [{ id: 1001, name: "Alpha", note: "first" }]);

  sheet.fromJson([{ id: 1002, name: "Beta", note: "second" }], { mode: "append" });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first" },
    { id: 1002, name: "Beta", note: "second" },
  ]);

  sheet.fromJson([{ id: 1002, note: "second-patched" }], {
    keyField: "id",
    mode: "update",
  });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first" },
    { id: 1002, name: "Beta", note: "second-patched" },
  ]);

  sheet.fromCsv("id,note\n1001,first-csv\n", {
    keyField: "id",
    mode: "update",
  });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first-csv" },
    { id: 1002, name: "Beta", note: "second-patched" },
  ]);

  sheet.fromJson([{ id: 1002, name: "Beta-2" }], {
    keyField: "id",
    mode: "upsert",
  });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first-csv" },
    { id: 1002, name: "Beta-2", note: null },
  ]);
});

test("sheet record workflow APIs import, export, and sync records", async () => {
  const workbook = Workbook.create("Data");
  const sheet = workbook.getSheet("Data");

  const replaced = sheet.importRecords([
    { id: 1001, name: "Alpha", note: "first" },
    { id: 1002, name: "Beta", note: "second" },
  ]);
  assert.deepEqual(replaced, {
    headers: ["id", "name", "note"],
    imported: 2,
    inserted: 2,
    mode: "replace",
    rowCount: 2,
    updated: 0,
  });

  assert.deepEqual(sheet.exportRecords(), [
    { id: 1001, name: "Alpha", note: "first" },
    { id: 1002, name: "Beta", note: "second" },
  ]);
  assert.equal(sheet.exportRecords({ format: "csv" }), "id,name,note\n1001,Alpha,first\n1002,Beta,second");

  const appended = sheet.importRecords([{ id: 1003, name: "Gamma", note: "third" }], { mode: "append" });
  assert.deepEqual(appended, {
    headers: ["id", "name", "note"],
    imported: 1,
    inserted: 1,
    mode: "append",
    rowCount: 3,
    updated: 0,
  });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first" },
    { id: 1002, name: "Beta", note: "second" },
    { id: 1003, name: "Gamma", note: "third" },
  ]);

  const updated = sheet.importRecords([{ id: 1001, note: "first-updated" }, { id: 9999, note: "missing" }], {
    keyField: "id",
    mode: "update",
  });
  assert.deepEqual(updated, {
    headers: ["id", "name", "note"],
    imported: 2,
    inserted: 0,
    mode: "update",
    rowCount: 3,
    updated: 1,
  });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first-updated" },
    { id: 1002, name: "Beta", note: "second" },
    { id: 1003, name: "Gamma", note: "third" },
  ]);

  const synced = sheet.syncRecords(
    [
      { id: 1002, name: "Beta-2" },
      { id: 1004, name: "Delta", note: "fourth" },
    ],
    { keyField: "id" },
  );
  assert.deepEqual(synced, {
    headers: ["id", "name", "note"],
    imported: 2,
    inserted: 1,
    mode: "upsert",
    rowCount: 4,
    updated: 1,
  });
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha", note: "first-updated" },
    { id: 1002, name: "Beta-2", note: null },
    { id: 1003, name: "Gamma", note: "third" },
    { id: 1004, name: "Delta", note: "fourth" },
  ]);
});

test("workbook can create config and table sheets through workflow helpers", async () => {
  const workbook = Workbook.create("Root");

  workbook.createConfigSheet("Config", {
    records: [
      { Key: "timeout", Value: "30" },
      { Key: "region", Value: "cn" },
    ],
  });
  workbook.createTableSheet("Data", {
    records: [
      { id: 1001, name: "Alpha" },
      { id: 1002, name: "Beta" },
    ],
  });

  assert.deepEqual(workbook.getSheet("Config").getHeaders(), ["Key", "Value"]);
  assert.deepEqual(workbook.getSheet("Config").getRecords(), [
    { Key: "timeout", Value: "30" },
    { Key: "region", Value: "cn" },
  ]);
  assert.deepEqual(workbook.getSheet("Data").getHeaders(), ["id", "name"]);
  assert.deepEqual(workbook.getSheet("Data").getRecords(), [
    { id: 1001, name: "Alpha" },
    { id: 1002, name: "Beta" },
  ]);
});

test("record alias APIs match the primary record helpers", async () => {
  const workbook = Workbook.create("Data");
  const sheet = workbook.getSheet("Data");

  sheet.appendRecord({ id: 1001, name: "Alpha" });
  sheet.appendRecords([{ id: 1002, name: "Beta" }]);
  assert.deepEqual(sheet.getRecords(), [
    { id: 1001, name: "Alpha" },
    { id: 1002, name: "Beta" },
  ]);

  sheet.replaceRecords([{ id: 1003, name: "Gamma" }]);
  assert.deepEqual(sheet.getRecords(), [{ id: 1003, name: "Gamma" }]);
});

test("record APIs can delete a record row", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  sheet.deleteRecord(2);

  assert.equal(sheet.getRecord(2), null);
  assert.deepEqual(sheet.getRecords(), [{ Name: "Bob", Score: 87 }]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<row r="2">/);
  assert.match(sheetXml, /<dimension ref="A1:B3"\/>/);
});

test("record APIs can delete multiple record rows", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
    { Name: "Cara", Score: 91 },
  ]);

  sheet.deleteRecords([2, 4, 2]);

  assert.equal(sheet.getRecord(2), null);
  assert.equal(sheet.getRecord(4), null);
  assert.deepEqual(sheet.getRecords(), [{ Name: "Bob", Score: 87 }]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<row r="2">/);
  assert.match(sheetXml, /<row r="3"><c r="A3" t="inlineStr"><is><t>Bob<\/t><\/is><\/c><c r="B3"><v>87<\/v><\/c><\/row>/);
  assert.doesNotMatch(sheetXml, /<row r="4">/);
});

test("merged range APIs patch mergeCells without touching unrelated parts", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getMergedRanges(), []);

  sheet.addMergedRange("B2:A1");
  sheet.addMergedRange("C3:D4");
  sheet.addMergedRange("A1:B2");

  assert.deepEqual(sheet.getMergedRanges(), ["A1:B2", "C3:D4"]);

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(
    sheetXml,
    /<\/sheetData><mergeCells count="2"><mergeCell ref="A1:B2"\/><mergeCell ref="C3:D4"\/><\/mergeCells>\s*<\/worksheet>/,
  );

  sheet.removeMergedRange("A1:B2");
  assert.deepEqual(sheet.getMergedRanges(), ["C3:D4"]);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<mergeCells count="1"><mergeCell ref="C3:D4"\/><\/mergeCells>/);

  sheet.removeMergedRange("C3:D4");
  assert.deepEqual(sheet.getMergedRanges(), []);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<mergeCells\b/);

  sheet.addMergedRange("E5:D4");
  assert.deepEqual(sheet.getMergedRanges(), ["D4:E5"]);
  sheet.clearMergedRanges();
  assert.deepEqual(sheet.getMergedRanges(), []);
});

test("column and merged range writers tolerate single-quoted container tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = convertEntriesToSingleQuotedAttributes(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols outlineLevelCol='1'>
    <col min='1' max='2' style='1'/>
    <col min='4' max='4' style='3' hidden='1'/>
  </cols>
  <sheetData>
    <row r='1'>
      <c r='A1' t='inlineStr'><is><t>Hello</t></is></c>
      <c r='D1'><v>4</v></c>
    </row>
  </sheetData>
  <mergeCells count='1'><mergeCell ref='B2:C2'/></mergeCells>
</worksheet>`,
    ),
    ["xl/worksheets/sheet1.xml"],
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getMergedRanges(), ["B2:C2"]);
  assert.equal(sheet.getColumnStyleId("A"), 1);
  assert.equal(sheet.getColumnStyleId("D"), 3);

  sheet.setColumnStyleId("B", 5);
  sheet.insertColumn("C");
  sheet.addMergedRange("E5:D4");
  sheet.removeMergedRange("B2:D2");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.equal(sheet.getColumnStyleId("B"), 5);
  assert.equal(sheet.getColumnStyleId("E"), 3);
  assert.deepEqual(sheet.getMergedRanges(), ["D4:E5"]);
  assert.match(
    sheetXml,
    /<cols outlineLevelCol="1"><col min="1" max="1" style="1"\/><col min="2" max="2" style="5"\/><col min="5" max="5" style="3" hidden="1"\/><\/cols>/,
  );
  assert.match(sheetXml, /<mergeCells count="1"><mergeCell ref="D4:E5"\/><\/mergeCells>/);
});

test("writing cells keeps worksheet dimension ref in sync", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c></row></sheetData></worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("C4", 9);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<dimension ref="A1:C4"\/>/);
});

test("workbook can add a sheet and wire workbook metadata", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>fastxlsx</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
  );
  const workbook = Workbook.fromEntries(entries);

  const newSheet = workbook.addSheet("Sheet2");
  newSheet.setCell("A1", "New");

  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sheet1", "Sheet2"]);
  assert.equal(newSheet.getCell("A1"), "New");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  const relsXml = entryText(workbook.toEntries(), "xl/_rels/workbook.xml.rels");
  const contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");
  const appXml = entryText(workbook.toEntries(), "docProps/app.xml");

  assert.match(workbookXml, /<sheet name="Sheet2" sheetId="2" r:id="rId3"\/>/);
  assert.match(relsXml, /<Relationship Id="rId3" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/worksheet" Target="worksheets\/sheet2\.xml"\/>/);
  assert.match(contentTypesXml, /<Override PartName="\/xl\/worksheets\/sheet2\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.worksheet\+xml"\/>/);
  assert.match(sheet2Xml, /<row r="1"><c r="A1" t="inlineStr"><is><t>New<\/t><\/is><\/c><\/row>/);
  assert.match(appXml, /<vt:i4>2<\/vt:i4>/);
  assert.match(appXml, /<vt:lpstr>Sheet1<\/vt:lpstr><vt:lpstr>Sheet2<\/vt:lpstr>/);
});

test("workbook can delete a sheet and rewrite remaining references", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      replaceEntryText(
        withSecondSheet(
          await loadFixtureEntries(fixtureDir),
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>9</v></c>
    </row>
  </sheetData>
</worksheet>`,
        ),
        "xl/workbook.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="ExternalRef">Sheet2!$A$1</definedName>
    <definedName name="LocalToSheet2" localSheetId="1">$A$1</definedName>
  </definedNames>
</workbook>`,
      ),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>Sheet2!A1</f><v>9</v></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>fastxlsx</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.deleteSheet("Sheet2");

  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sheet1"]);
  assert.equal(workbook.getSheet("Sheet1").getFormula("A1"), "#REF!");
  assert.equal(workbook.listEntries().includes("xl/worksheets/sheet2.xml"), false);

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  const relsXml = entryText(workbook.toEntries(), "xl/_rels/workbook.xml.rels");
  const contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  const appXml = entryText(workbook.toEntries(), "docProps/app.xml");

  assert.doesNotMatch(workbookXml, /Sheet2/);
  assert.match(workbookXml, /<definedName name="ExternalRef">#REF!<\/definedName>/);
  assert.doesNotMatch(workbookXml, /LocalToSheet2/);
  assert.doesNotMatch(relsXml, /Target="worksheets\/sheet2\.xml"/);
  assert.doesNotMatch(contentTypesXml, /PartName="\/xl\/worksheets\/sheet2\.xml"/);
  assert.match(appXml, /<vt:i4>1<\/vt:i4>/);
  assert.match(appXml, /<vt:lpstr>Sheet1<\/vt:lpstr>/);
  assert.doesNotMatch(appXml, /Sheet2/);
});

test("workbook can move a sheet and remap local sheet scopes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      replaceEntryText(
        withSecondSheet(
          await loadFixtureEntries(fixtureDir),
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
        ),
        "xl/workbook.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="1"/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
    <sheet name="Sheet3" sheetId="3" r:id="rId4"/>
  </sheets>
  <definedNames>
    <definedName name="LocalToSheet1" localSheetId="0">$A$1</definedName>
    <definedName name="LocalToSheet2" localSheetId="1">$A$1</definedName>
    <definedName name="LocalToSheet3" localSheetId="2">$A$1</definedName>
  </definedNames>
</workbook>`,
      ),
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`,
    ),
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>fastxlsx</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr><vt:lpstr>Sheet3</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
  );
  const withThirdSheet = [
    ...entries,
    {
      path: "xl/worksheets/sheet3.xml",
      data: new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>3</v></c></row>
  </sheetData>
</worksheet>`),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
  const workbook = Workbook.fromEntries(withThirdSheet);

  workbook.moveSheet("Sheet3", 0);

  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sheet3", "Sheet1", "Sheet2"]);
  assert.equal(workbook.getDefinedName("LocalToSheet1", "Sheet1"), "$A$1");
  assert.equal(workbook.getDefinedName("LocalToSheet2", "Sheet2"), "$A$1");
  assert.equal(workbook.getDefinedName("LocalToSheet3", "Sheet3"), "$A$1");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  const appXml = entryText(workbook.toEntries(), "docProps/app.xml");

  assert.match(
    workbookXml,
    /<sheets><sheet name="Sheet3" sheetId="3" r:id="rId4"\/><sheet name="Sheet1" sheetId="1" r:id="rId1"\/><sheet name="Sheet2" sheetId="2" r:id="rId3"\/><\/sheets>/,
  );
  assert.match(workbookXml, /<workbookView activeTab="2"\/>/);
  assert.match(workbookXml, /<definedName name="LocalToSheet1" localSheetId="1">\$A\$1<\/definedName>/);
  assert.match(workbookXml, /<definedName name="LocalToSheet2" localSheetId="2">\$A\$1<\/definedName>/);
  assert.match(workbookXml, /<definedName name="LocalToSheet3" localSheetId="0">\$A\$1<\/definedName>/);
  assert.match(appXml, /<vt:lpstr>Sheet3<\/vt:lpstr><vt:lpstr>Sheet1<\/vt:lpstr><vt:lpstr>Sheet2<\/vt:lpstr>/);
});

test("workbook active sheet APIs read and write workbookView activeTab", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      withSecondSheet(
        await loadFixtureEntries(fixtureDir),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
      ),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="1"/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.equal(workbook.getActiveSheet().name, "Sheet2");
  workbook.setActiveSheet("Sheet1");
  assert.equal(workbook.getActiveSheet().name, "Sheet1");

  let workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<workbookView activeTab="0"\/>/);

  workbook.setSheetVisibility("Sheet2", "hidden");
  assert.throws(() => workbook.setActiveSheet("Sheet2"), /Cannot activate hidden sheet: Sheet2/);

  const withoutBookViews = Workbook.fromEntries(
    replaceEntryText(
      entries,
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
</workbook>`,
    ),
  );

  assert.equal(withoutBookViews.getActiveSheet().name, "Sheet1");
  withoutBookViews.setActiveSheet("Sheet2");
  workbookXml = entryText(withoutBookViews.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<bookViews><workbookView activeTab="1"\/><\/bookViews>\s*<sheets>/);
});

test("sheet freeze pane APIs read, write, shift, and remove pane state", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getFreezePane(), {
    columnCount: 1,
    rowCount: 1,
    topLeftCell: "B2",
    activePane: "bottomRight",
  });

  sheet.insertColumn("A");
  assert.deepEqual(sheet.getFreezePane(), {
    columnCount: 1,
    rowCount: 1,
    topLeftCell: "C2",
    activePane: "bottomRight",
  });

  sheet.freezePane(2, 1);
  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<pane state="frozen" xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight"\/>/);
  assert.match(sheetXml, /<selection pane="topRight"\/><selection pane="bottomLeft"\/><selection pane="bottomRight" activeCell="C2" sqref="C2"\/>/);

  sheet.unfreezePane();
  assert.equal(sheet.getFreezePane(), null);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<pane\b/);
  assert.match(sheetXml, /<selection activeCell="C2" sqref="C2"\/>/);

  sheet.freezePane(0, 2);
  assert.deepEqual(sheet.getFreezePane(), {
    columnCount: 0,
    rowCount: 2,
    topLeftCell: "A3",
    activePane: "bottomLeft",
  });
});

test("sheet selection APIs read, write, and follow frozen active panes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getSelection(), {
    activeCell: "B2",
    range: "B2",
    pane: "bottomRight",
  });

  sheet.setSelection("C3", "C3:D4");

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<selection pane="bottomRight" activeCell="C3" sqref="C3:D4"\/>/);
  assert.deepEqual(sheet.getSelection(), {
    activeCell: "C3",
    range: "C3:D4",
    pane: "bottomRight",
  });

  sheet.unfreezePane();
  assert.deepEqual(sheet.getSelection(), {
    activeCell: "C3",
    range: "C3:D4",
    pane: null,
  });

  sheet.setSelection("A1", "A1:B2");
  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<selection activeCell="A1" sqref="A1:B2"\/>/);

  sheet.clearSelection();
  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<selection\b/);
  assert.equal(sheet.getSelection(), null);
});

test("workbook sheet visibility APIs read and write hidden states", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      withSecondSheet(
        await loadFixtureEntries(fixtureDir),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
      ),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3" state="hidden"/>
  </sheets>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.equal(workbook.getSheetVisibility("Sheet1"), "visible");
  assert.equal(workbook.getSheetVisibility("Sheet2"), "hidden");

  workbook.setSheetVisibility("Sheet2", "veryHidden");
  assert.equal(workbook.getSheetVisibility("Sheet2"), "veryHidden");

  assert.throws(
    () => workbook.setSheetVisibility("Sheet1", "hidden"),
    /Workbook must contain at least one visible sheet/,
  );

  workbook.setSheetVisibility("Sheet2", "visible");
  workbook.setSheetVisibility("Sheet1", "hidden");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<sheet name="Sheet1" sheetId="1" r:id="rId1" state="hidden"\/>/);
  assert.match(workbookXml, /<sheet name="Sheet2" sheetId="2" r:id="rId3"\/>/);
  assert.equal(workbook.getSheetVisibility("Sheet1"), "hidden");
  assert.equal(workbook.getSheetVisibility("Sheet2"), "visible");
});

test("sheet rename updates workbook metadata, formulas, and hyperlink locations", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      replaceEntryText(
        withSecondSheet(
          await loadFixtureEntries(fixtureDir),
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>Sheet1!A1</f><v>1</v></c>
    </row>
  </sheetData>
  <hyperlinks><hyperlink ref="A1" location="#Sheet1!A1"/></hyperlinks>
</worksheet>`,
        ),
        "xl/workbook.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="ExternalRef">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>`,
      ),
      "docProps/app.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>fastxlsx</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>Sheet1!A1</f><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.rename("Data Set");

  assert.equal(sheet.name, "Data Set");
  assert.deepEqual(workbook.getSheets().map((candidate) => candidate.name), ["Data Set", "Sheet2"]);
  assert.equal(workbook.getSheet("Sheet2").getFormula("A1"), "'Data Set'!A1");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  const appXml = entryText(workbook.toEntries(), "docProps/app.xml");
  const sheet1Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");

  assert.match(workbookXml, /<sheet name="Data Set" sheetId="1" r:id="rId1"\/>/);
  assert.match(workbookXml, /<definedName name="ExternalRef">'Data Set'!\$A\$1<\/definedName>/);
  assert.match(sheet1Xml, /<c r="A1"><f>'Data Set'!A1<\/f><v>1<\/v><\/c>/);
  assert.match(sheet2Xml, /<c r="A1"><f>'Data Set'!A1<\/f><v>1<\/v><\/c>/);
  assert.match(sheet2Xml, /<hyperlinks><hyperlink ref="A1" location="#'Data Set'!A1"\/><\/hyperlinks>/);
  assert.match(appXml, /<vt:lpstr>Data Set<\/vt:lpstr><vt:lpstr>Sheet2<\/vt:lpstr>/);
});

test("sheet hyperlink APIs read, write, replace, and delete hyperlinks", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    withSecondSheet(
      await loadFixtureEntries(fixtureDir),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><is><t>Old</t></is></c></row>
    <row r="2"><c r="B2"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getHyperlink("A1"), null);

  sheet.setHyperlink("A1", "https://example.com", { text: "Open", tooltip: "Go" });
  sheet.setHyperlink("B2", "#Sheet2!A1");

  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "A1", target: "https://example.com", tooltip: "Go", type: "external" },
    { address: "B2", target: "#Sheet2!A1", tooltip: null, type: "internal" },
  ]);
  assert.deepEqual(sheet.getHyperlink("A1"), {
    address: "A1",
    target: "https://example.com",
    tooltip: "Go",
    type: "external",
  });
  assert.deepEqual(sheet.hyperlink("B2"), {
    address: "B2",
    target: "#Sheet2!A1",
    tooltip: null,
    type: "internal",
  });
  assert.equal(sheet.getCell("A1"), "Open");

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  let relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  assert.match(
    sheetXml,
    /<hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Go"\/><hyperlink ref="B2" location="#Sheet2!A1"\/><\/hyperlinks>/,
  );
  assert.match(
    relsXml,
    /<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink" Target="https:\/\/example\.com" TargetMode="External"\/>/,
  );

  sheet.setHyperlink("A1", "#Sheet2!B3");

  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "A1", target: "#Sheet2!B3", tooltip: null, type: "internal" },
    { address: "B2", target: "#Sheet2!A1", tooltip: null, type: "internal" },
  ]);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  assert.match(
    sheetXml,
    /<hyperlinks><hyperlink ref="A1" location="#Sheet2!B3"\/><hyperlink ref="B2" location="#Sheet2!A1"\/><\/hyperlinks>/,
  );
  assert.doesNotMatch(relsXml, /relationships\/hyperlink/);

  sheet.removeHyperlink("A1");

  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "B2", target: "#Sheet2!A1", tooltip: null, type: "internal" },
  ]);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<hyperlinks><hyperlink ref="B2" location="#Sheet2!A1"\/><\/hyperlinks>/);

  sheet.clearHyperlinks();

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<hyperlinks>/);
  assert.deepEqual(sheet.getHyperlinks(), []);
  assert.equal(sheet.getHyperlink("B2"), null);
});

test("sheet comment APIs create, update, read, and delete comments", async () => {
  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getComments(), []);
  assert.equal(sheet.getComment("B2"), null);

  const createdComment = sheet.setComment("B2", "Hello comment", { author: "Alice" });
  assert.deepEqual(createdComment, {
    address: "B2",
    author: "Alice",
    text: "Hello comment",
  });

  assert.deepEqual(sheet.getComment("B2"), {
    address: "B2",
    author: "Alice",
    text: "Hello comment",
  });
  assert.deepEqual(sheet.comment("B2"), {
    address: "B2",
    author: "Alice",
    text: "Hello comment",
  });

  let entries = workbook.toEntries();
  let sheetXml = entryText(entries, "xl/worksheets/sheet1.xml");
  let commentsXml = entryText(entries, "xl/comments1.xml");
  let relsXml = entryText(entries, "xl/worksheets/_rels/sheet1.xml.rels");
  let contentTypesXml = entryText(entries, "[Content_Types].xml");
  let vmlXml = entryText(entries, "xl/drawings/vmlDrawing1.vml");

  assert.match(sheetXml, /<legacyDrawing r:id="rId2"\/>/);
  assert.match(commentsXml, /<authors><author>Alice<\/author><\/authors>/);
  assert.match(commentsXml, /<comment ref="B2" authorId="0"><text><t>Hello comment<\/t><\/text><\/comment>/);
  assert.match(
    relsXml,
    /<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/comments" Target="\.\.\/comments1\.xml"\/>/,
  );
  assert.match(
    relsXml,
    /<Relationship Id="rId2" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/vmlDrawing" Target="\.\.\/drawings\/vmlDrawing1\.vml"\/>/,
  );
  assert.match(contentTypesXml, /<Override PartName="\/xl\/comments1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.comments\+xml"\/>/);
  assert.match(contentTypesXml, /<Default Extension="vml" ContentType="application\/vnd\.openxmlformats-officedocument\.vmlDrawing"\/>/);
  assert.match(vmlXml, /<x:Row>1<\/x:Row>/);
  assert.match(vmlXml, /<x:Column>1<\/x:Column>/);

  const updatedComment = sheet.setComment("B2", "Updated");
  const firstComment = sheet.setComment("A1", "First");
  assert.deepEqual(updatedComment, {
    address: "B2",
    author: "Alice",
    text: "Updated",
  });
  assert.deepEqual(firstComment, {
    address: "A1",
    author: "Alice",
    text: "First",
  });

  assert.deepEqual(sheet.getComments(), [
    { address: "A1", author: "Alice", text: "First" },
    { address: "B2", author: "Alice", text: "Updated" },
  ]);

  assert.deepEqual(
    sheet.setComments([
      { address: "C3", author: "Bob", text: "Third" },
      { address: "D4", author: null, text: "Fourth" },
    ]),
    [
      { address: "C3", author: "Bob", text: "Third" },
      { address: "D4", author: "Bob", text: "Fourth" },
    ],
  );
  assert.deepEqual(sheet.getComments(), [
    { address: "C3", author: "Bob", text: "Third" },
    { address: "D4", author: "Bob", text: "Fourth" },
  ]);

  sheet.removeComment("A1");
  assert.deepEqual(sheet.getComments(), [
    { address: "C3", author: "Bob", text: "Third" },
    { address: "D4", author: "Bob", text: "Fourth" },
  ]);

  sheet.clearComments();
  assert.deepEqual(sheet.getComments(), []);
  assert.equal(sheet.getComment("B2"), null);

  entries = workbook.toEntries();
  sheetXml = entryText(entries, "xl/worksheets/sheet1.xml");
  contentTypesXml = entryText(entries, "[Content_Types].xml");

  assert.doesNotMatch(sheetXml, /<legacyDrawing\b/);
  assert.doesNotMatch(contentTypesXml, /\/xl\/comments1\.xml/);
  assert.equal(workbook.listEntries().includes("xl/comments1.xml"), false);
  assert.equal(workbook.listEntries().includes("xl/drawings/vmlDrawing1.vml"), false);
});

test("sheet autoFilter definition reads supported columns and sort state", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="B1"><v>Header 1</v></c><c r="C1"><v>Header 2</v></c><c r="D1"><v>Header 3</v></c></row>
    <row r="2"><c r="B2"><v>Alpha</v></c><c r="C2"><v>11</v></c><c r="D2"><v>2024-04</v></c></row>
    <row r="3"><c r="B3"><v>Beta</v></c><c r="C3"><v>15</v></c><c r="D3"><v>2024-05</v></c></row>
    <row r="4"><c r="B4"><v>Gamma</v></c><c r="C4"><v>21</v></c><c r="D4"><v>2024-06</v></c></row>
    <row r="5"><c r="B5"><v></v></c><c r="C5"><v>18</v></c><c r="D5"><v>2024-07</v></c></row>
  </sheetData>
  <autoFilter ref="B2:D5">
    <filterColumn colId="0">
      <filters blank="1">
        <filter val="Alpha"/>
        <filter val="Beta"/>
      </filters>
    </filterColumn>
    <filterColumn colId="1">
      <customFilters and="1">
        <customFilter operator="greaterThan" val="10"/>
        <customFilter operator="lessThanOrEqual" val="20"/>
      </customFilters>
    </filterColumn>
    <filterColumn colId="2">
      <filters>
        <dateGroupItem year="2024" month="4" dateTimeGrouping="month"/>
      </filters>
    </filterColumn>
    <extLst><ext uri="{urn:test:autoFilter}"/></extLst>
  </autoFilter>
  <sortState ref="B2:D5">
    <sortCondition ref="C2:C5" descending="1"/>
    <sortCondition ref="D2:D5"/>
  </sortState>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getAutoFilterDefinition(), {
    range: "B2:D5",
    columns: [
      {
        columnNumber: 2,
        kind: "values",
        values: ["Alpha", "Beta"],
        includeBlank: true,
      },
      {
        columnNumber: 3,
        kind: "custom",
        join: "and",
        conditions: [
          { operator: "greaterThan", value: "10" },
          { operator: "lessThanOrEqual", value: "20" },
        ],
      },
      {
        columnNumber: 4,
        kind: "dateGroup",
        items: [{ year: 2024, month: 4, dateTimeGrouping: "month" }],
      },
    ],
    sortState: {
      range: "B2:D5",
      conditions: [{ columnNumber: 3, descending: true }, { columnNumber: 4 }],
    },
  });
});

test("sheet setAutoFilter preserves nested filter XML when only the range changes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c></row>
    <row r="2"><c r="A2"><v>5</v></c><c r="B2"><v>6</v></c><c r="C2"><v>7</v></c><c r="D2"><v>8</v></c></row>
  </sheetData>
  <autoFilter ref="A1:C2">
    <filterColumn colId="0">
      <filters blank="1"><filter val="Alpha"/></filters>
    </filterColumn>
    <extLst><ext uri="{urn:test:autoFilter}"/></extLst>
  </autoFilter>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setAutoFilter("A1:D2");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<autoFilter ref="A1:D2">/);
  assert.match(sheetXml, /<filterColumn colId="0">\s*<filters blank="1"><filter val="Alpha"\/><\/filters>\s*<\/filterColumn>/);
  assert.match(sheetXml, /<extLst><ext uri="\{urn:test:autoFilter\}"\/><\/extLst>/);
});

test("sheet setAutoFilter rebases filter columns and preserves sortState children when the range anchor changes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>0</v></c><c r="B1"><v>1</v></c><c r="C1"><v>2</v></c><c r="D1"><v>3</v></c></row>
    <row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>5</v></c><c r="C2"><v>6</v></c><c r="D2"><v>7</v></c></row>
  </sheetData>
  <autoFilter ref="B1:D2">
    <filterColumn colId="1"><filters><filter val="Alpha"/></filters></filterColumn>
    <extLst><ext uri="{urn:test:autoFilter}"/></extLst>
  </autoFilter>
  <sortState ref="B1:D2" caseSensitive="1">
    <sortCondition ref="C1:C2" descending="1"/>
    <extLst><ext uri="{urn:test:sortState}"/></extLst>
  </sortState>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setAutoFilter("A1:D2");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<autoFilter ref="A1:D2">/);
  assert.match(sheetXml, /<filterColumn colId="2"><filters><filter val="Alpha"\/><\/filters><\/filterColumn>/);
  assert.match(sheetXml, /<extLst><ext uri="\{urn:test:autoFilter\}"\/><\/extLst>/);
  assert.match(sheetXml, /<sortState ref="A1:D2" caseSensitive="1">/);
  assert.match(sheetXml, /<sortCondition ref="C1:C2" descending="1"\/>/);
  assert.match(sheetXml, /<extLst><ext uri="\{urn:test:sortState\}"\/><\/extLst>/);
});

test("sheet autoFilter definition updates supported columns and preserves unrelated unsupported XML", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>H1</v></c><c r="B1"><v>H2</v></c><c r="C1"><v>H3</v></c><c r="D1"><v>H4</v></c></row>
    <row r="2"><c r="A2"><v>A</v></c><c r="B2"><v>B</v></c><c r="C2"><v>C</v></c><c r="D2"><v>D</v></c></row>
    <row r="3"><c r="A3"><v>E</v></c><c r="B3"><v>F</v></c><c r="C3"><v>G</v></c><c r="D3"><v>H</v></c></row>
    <row r="4"><c r="A4"><v>I</v></c><c r="B4"><v>J</v></c><c r="C4"><v>K</v></c><c r="D4"><v>L</v></c></row>
  </sheetData>
  <autoFilter ref="A1:D4">
    <filterColumn colId="0"><filters><filter val="Old"/></filters></filterColumn>
    <filterColumn colId="1"><fooFilter answer="42"/></filterColumn>
    <extLst><ext uri="{urn:test:autoFilter}"/></extLst>
  </autoFilter>
  <sortState ref="A1:D4"><sortCondition ref="A1:A4"/></sortState>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setAutoFilterDefinition({
    range: "A1:D4",
    columns: [
      { columnNumber: 1, kind: "values", values: ["Alpha"], includeBlank: true },
      { columnNumber: 4, kind: "custom", join: "or", conditions: [{ operator: "contains", value: "foo" }] },
    ],
    sortState: {
      range: "A1:D4",
      conditions: [{ columnNumber: 4, descending: true }],
    },
  });

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<filterColumn colId="0"><filters blank="1"><filter val="Alpha"\/><\/filters><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="1"><fooFilter answer="42"\/><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="3"><customFilters><customFilter operator="equal" val="\*foo\*"\/><\/customFilters><\/filterColumn>/);
  assert.match(sheetXml, /<extLst><ext uri="\{urn:test:autoFilter\}"\/><\/extLst>/);
  assert.match(sheetXml, /<sortState ref="A1:D4"><sortCondition ref="D1:D4" descending="1"\/><\/sortState>/);

  assert.deepEqual(sheet.getAutoFilterDefinition(), {
    range: "A1:D4",
    columns: [
      {
        columnNumber: 1,
        kind: "values",
        values: ["Alpha"],
        includeBlank: true,
      },
      {
        columnNumber: 4,
        kind: "custom",
        join: "or",
        conditions: [{ operator: "contains", value: "foo" }],
      },
    ],
    sortState: {
      range: "A1:D4",
      conditions: [{ columnNumber: 4, descending: true }],
    },
  });

  sheet.clearAutoFilterColumns([2]);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<fooFilter\b/);
  assert.match(sheetXml, /<filterColumn colId="0"><filters blank="1"><filter val="Alpha"\/><\/filters><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="3"><customFilters><customFilter operator="equal" val="\*foo\*"\/><\/customFilters><\/filterColumn>/);
});

test("sheet autoFilter definition reads and writes blank and advanced supported filter kinds", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>H1</v></c><c r="B1"><v>H2</v></c><c r="C1"><v>H3</v></c><c r="D1"><v>H4</v></c><c r="E1"><v>H5</v></c></row>
    <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c><c r="D2"><v>4</v></c><c r="E2"><v>5</v></c></row>
  </sheetData>
  <autoFilter ref="A1:E2">
    <filterColumn colId="0"><filters blank="0"/></filterColumn>
    <filterColumn colId="1"><colorFilter dxfId="7" cellColor="1"/></filterColumn>
    <filterColumn colId="2"><dynamicFilter type="today" valIso="2026-04-28T00:00:00Z" maxValIso="2026-04-29T00:00:00Z"/></filterColumn>
    <filterColumn colId="3"><top10 top="0" percent="1" val="15" filterVal="42"/></filterColumn>
    <filterColumn colId="4"><iconFilter iconSet="3Arrows" iconId="2"/></filterColumn>
  </autoFilter>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getAutoFilterDefinition(), {
    range: "A1:E2",
    columns: [
      { columnNumber: 1, kind: "blank", mode: "nonBlank" },
      { columnNumber: 2, kind: "color", dxfId: 7, cellColor: true },
      {
        columnNumber: 3,
        kind: "dynamic",
        type: "today",
        valIso: "2026-04-28T00:00:00Z",
        maxValIso: "2026-04-29T00:00:00Z",
      },
      {
        columnNumber: 4,
        kind: "top10",
        top: false,
        percent: true,
        value: 15,
        filterValue: 42,
      },
      { columnNumber: 5, kind: "icon", iconSet: "3Arrows", iconId: 2 },
    ],
    sortState: null,
  });

  sheet.setAutoFilterDefinition({
    range: "A1:E2",
    columns: [
      { columnNumber: 1, kind: "blank", mode: "blank" },
      { columnNumber: 2, kind: "color", dxfId: 9, cellColor: false },
      { columnNumber: 3, kind: "dynamic", type: "belowAverage", valIso: "2026-04-28T00:00:00Z" },
      { columnNumber: 4, kind: "top10", top: true, percent: false, value: 5 },
      { columnNumber: 5, kind: "icon", iconSet: "5Arrows", iconId: 4 },
    ],
  });

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<filterColumn colId="0"><filters blank="1"\/><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="1"><colorFilter dxfId="9" cellColor="0"\/><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="2"><dynamicFilter type="belowAverage" valIso="2026-04-28T00:00:00Z"\/><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="3"><top10 top="1" percent="0" val="5"\/><\/filterColumn>/);
  assert.match(sheetXml, /<filterColumn colId="4"><iconFilter iconSet="5Arrows" iconId="4"\/><\/filterColumn>/);

  assert.deepEqual(sheet.getAutoFilterDefinition(), {
    range: "A1:E2",
    columns: [
      { columnNumber: 1, kind: "blank", mode: "blank" },
      { columnNumber: 2, kind: "color", dxfId: 9, cellColor: false },
      { columnNumber: 3, kind: "dynamic", type: "belowAverage", valIso: "2026-04-28T00:00:00Z" },
      { columnNumber: 4, kind: "top10", top: true, percent: false, value: 5 },
      { columnNumber: 5, kind: "icon", iconSet: "5Arrows", iconId: 4 },
    ],
    sortState: null,
  });
});

test("table autoFilter handle reads and updates filters without dropping nested XML on ref rewrites", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSheetTable(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>4</v></c>
      <c r="B2"><v>5</v></c>
      <c r="C2"><v>6</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>7</v></c>
      <c r="B3"><v>8</v></c>
      <c r="C3"><v>9</v></c>
    </row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3" totalsRowShown="0">
  <autoFilter ref="A1:B3">
    <filterColumn colId="0"><filters blank="1"><filter val="Alpha"/></filters></filterColumn>
    <sortState ref="A1:B3" caseSensitive="1">
      <sortCondition ref="B1:B3" descending="1"/>
      <extLst><ext uri="{urn:test:tableSort}"/></extLst>
    </sortState>
    <extLst><ext uri="{urn:test:autoFilter}"/></extLst>
  </autoFilter>
  <tableColumns count="2">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
  </tableColumns>
</table>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const table = sheet.getTable("Sales");

  assert.equal(sheet.tryGetTable("Missing"), null);
  assert.throws(() => sheet.getTable("Missing"), /Table not found: Missing/);
  assert.deepEqual(sheet.getTables({ includeAutoFilter: true }), [
    {
      name: "Sales",
      displayName: "Sales",
      range: "A1:B3",
      path: "xl/tables/table1.xml",
      autoFilter: {
        range: "A1:B3",
        columns: [{ columnNumber: 1, kind: "values", values: ["Alpha"], includeBlank: true }],
        sortState: {
          range: "A1:B3",
          conditions: [{ columnNumber: 2, descending: true }],
        },
      },
    },
  ]);
  assert.deepEqual(table.getAutoFilterDefinition(), {
    range: "A1:B3",
    columns: [{ columnNumber: 1, kind: "values", values: ["Alpha"], includeBlank: true }],
    sortState: {
      range: "A1:B3",
      conditions: [{ columnNumber: 2, descending: true }],
    },
  });

  table.setAutoFilterColumn({
    columnNumber: 2,
    kind: "custom",
    join: "or",
    conditions: [{ operator: "endsWith", value: "Corp" }],
  });

  let tableXml = entryText(workbook.toEntries(), "xl/tables/table1.xml");
  assert.match(tableXml, /<filterColumn colId="0"><filters blank="1"><filter val="Alpha"\/><\/filters><\/filterColumn>/);
  assert.match(tableXml, /<filterColumn colId="1"><customFilters><customFilter operator="equal" val="\*Corp"\/><\/customFilters><\/filterColumn>/);
  assert.match(tableXml, /<sortState ref="A1:B3" caseSensitive="1">/);
  assert.match(tableXml, /<sortCondition ref="B1:B3" descending="1"\/>/);
  assert.match(tableXml, /<extLst><ext uri="\{urn:test:tableSort\}"\/><\/extLst>/);
  assert.match(tableXml, /<extLst><ext uri="\{urn:test:autoFilter\}"\/><\/extLst>/);

  sheet.insertColumn("B");

  tableXml = entryText(workbook.toEntries(), "xl/tables/table1.xml");
  assert.equal(table.range, "A1:C3");
  assert.match(tableXml, /<table [^>]*ref="A1:C3"[^>]*>/);
  assert.match(tableXml, /<autoFilter ref="A1:C3">/);
  assert.match(tableXml, /<filterColumn colId="0"><filters blank="1"><filter val="Alpha"\/><\/filters><\/filterColumn>/);
  assert.match(tableXml, /<filterColumn colId="2"><customFilters><customFilter operator="equal" val="\*Corp"\/><\/customFilters><\/filterColumn>/);
  assert.match(tableXml, /<sortState ref="A1:C3" caseSensitive="1">/);
  assert.match(tableXml, /<sortCondition ref="C1:C3" descending="1"\/>/);
  assert.match(tableXml, /<extLst><ext uri="\{urn:test:tableSort\}"\/><\/extLst>/);
  assert.match(tableXml, /<extLst><ext uri="\{urn:test:autoFilter\}"\/><\/extLst>/);
});

test("sheet autoFilter APIs read, write, shift, and remove filters", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>
    <row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>5</v></c><c r="C2"><v>6</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="E1:F1"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getAutoFilter(), null);

  sheet.setAutoFilter("A1:C2");
  assert.equal(sheet.getAutoFilter(), "A1:C2");

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(
    sheetXml,
    /<\/sheetData>\s*<autoFilter ref="A1:C2"\/><mergeCells count="1"><mergeCell ref="E1:F1"\/><\/mergeCells>/,
  );

  sheet.insertColumn("B");
  assert.equal(sheet.getAutoFilter(), "A1:D2");

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<autoFilter ref="A1:D2"\/>/);

  workbook.writeEntryText(
    "xl/worksheets/sheet1.xml",
    sheetXml.replace(
      /<autoFilter ref="A1:D2"\/>/,
      `<autoFilter ref="A1:D2"/><sortState ref="A2:D2"/>`,
    ),
  );

  assert.equal(sheet.getAutoFilter(), "A1:D2");

  sheet.removeAutoFilter();

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.equal(sheet.getAutoFilter(), null);
  assert.doesNotMatch(sheetXml, /<autoFilter\b/);
  assert.doesNotMatch(sheetXml, /<sortState\b/);

  sheet.setAutoFilter("A1:C2");
  assert.equal(sheet.getAutoFilter(), "A1:C2");
  sheet.clearAutoFilter();
  assert.equal(sheet.getAutoFilter(), null);
});

test("sheet dataValidation APIs read, write, shift, and remove validations", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c></row>
    <row r="2"><c r="A2"><v>5</v></c><c r="B2"><v>6</v></c><c r="C2"><v>7</v></c><c r="D2"><v>8</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getDataValidations(), []);
  assert.equal(sheet.getDataValidation("A2:B4"), null);

  sheet.setDataValidation("A2:B4", {
    type: "whole",
    operator: "between",
    allowBlank: true,
    showErrorMessage: true,
    errorStyle: "stop",
    errorTitle: "Invalid",
    error: "Enter 1-10",
    formula1: "1",
    formula2: "10",
  });
  sheet.setDataValidation("D2", {
    type: "list",
    showDropDown: false,
    promptTitle: "Pick",
    prompt: "Choose one",
    formula1: "\"Yes,No\"",
  });

  assert.deepEqual(sheet.getDataValidations(), [
    {
      range: "A2:B4",
      type: "whole",
      operator: "between",
      allowBlank: true,
      showInputMessage: null,
      showErrorMessage: true,
      showDropDown: null,
      errorStyle: "stop",
      errorTitle: "Invalid",
      error: "Enter 1-10",
      promptTitle: null,
      prompt: null,
      imeMode: null,
      formula1: "1",
      formula2: "10",
    },
    {
      range: "D2",
      type: "list",
      operator: null,
      allowBlank: null,
      showInputMessage: null,
      showErrorMessage: null,
      showDropDown: false,
      errorStyle: null,
      errorTitle: null,
      error: null,
      promptTitle: "Pick",
      prompt: "Choose one",
      imeMode: null,
      formula1: "\"Yes,No\"",
      formula2: null,
    },
  ]);
  assert.deepEqual(sheet.getDataValidation("A2:B4"), {
    range: "A2:B4",
    type: "whole",
    operator: "between",
    allowBlank: true,
    showInputMessage: null,
    showErrorMessage: true,
    showDropDown: null,
    errorStyle: "stop",
    errorTitle: "Invalid",
    error: "Enter 1-10",
    promptTitle: null,
    prompt: null,
    imeMode: null,
    formula1: "1",
    formula2: "10",
  });

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<dataValidations count="2">/);
  assert.match(
    sheetXml,
    /<dataValidation sqref="A2:B4" type="whole" operator="between" allowBlank="1" showErrorMessage="1" errorStyle="stop" errorTitle="Invalid" error="Enter 1-10"><formula1>1<\/formula1><formula2>10<\/formula2><\/dataValidation>/,
  );
  assert.match(
    sheetXml,
    /<dataValidation sqref="D2" type="list" showDropDown="0" promptTitle="Pick" prompt="Choose one"><formula1>&quot;Yes,No&quot;<\/formula1><\/dataValidation>/,
  );

  sheet.insertColumn("B");

  assert.deepEqual(sheet.getDataValidations(), [
    {
      range: "A2:C4",
      type: "whole",
      operator: "between",
      allowBlank: true,
      showInputMessage: null,
      showErrorMessage: true,
      showDropDown: null,
      errorStyle: "stop",
      errorTitle: "Invalid",
      error: "Enter 1-10",
      promptTitle: null,
      prompt: null,
      imeMode: null,
      formula1: "1",
      formula2: "10",
    },
    {
      range: "E2",
      type: "list",
      operator: null,
      allowBlank: null,
      showInputMessage: null,
      showErrorMessage: null,
      showDropDown: false,
      errorStyle: null,
      errorTitle: null,
      error: null,
      promptTitle: "Pick",
      prompt: "Choose one",
      imeMode: null,
      formula1: "\"Yes,No\"",
      formula2: null,
    },
  ]);

  sheet.removeDataValidation("A2:C4");

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<dataValidations count="1"><dataValidation sqref="E2" type="list" showDropDown="0" promptTitle="Pick" prompt="Choose one"><formula1>&quot;Yes,No&quot;<\/formula1><\/dataValidation><\/dataValidations>/);

  sheet.removeDataValidation("E2");

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.deepEqual(sheet.getDataValidations(), []);
  assert.doesNotMatch(sheetXml, /<dataValidations\b/);

  sheet.setDataValidation("C2", {
    type: "list",
    formula1: "\"Yes,No\"",
  });
  assert.notEqual(sheet.getDataValidation("C2"), null);
  sheet.clearDataValidations();
  assert.deepEqual(sheet.getDataValidations(), []);
});

test("sheet sortRange reorders rows and moves linked metadata with the sorted records", () => {
  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRange("A1", [
    ["Name", "Score", "Group", "Formula"],
    ["Beta", 2, "Low", null],
    ["Alpha", 5, "High", null],
    ["Gamma", 3, "Mid", null],
  ]);
  sheet.setFormula("D2", "B2*10", { cachedValue: 20 });
  sheet.setFormula("D3", "B3*10", { cachedValue: 50 });
  sheet.setFormula("D4", "B4*10", { cachedValue: 30 });
  sheet.setBackgroundColor("A2", "FFFF0000");
  sheet.setBackgroundColor("A3", "FF00FF00");
  sheet.setBackgroundColor("A4", "FF0000FF");
  sheet.addMergedRange("C2:D2");
  sheet.setHyperlink("A2", "https://example.com/beta");
  sheet.setDataValidation("C2:D2", { type: "list", formula1: "\"Low,Mid,High\"" });
  sheet.setAutoFilterDefinition({ range: "A1:D4", columns: [], sortState: null });
  sheet.addTable("A1:D4", { name: "Scores" });

  sheet.sortRange("A1:D4", {
    conditions: [{ columnNumber: 2, descending: true }],
    hasHeaderRow: true,
  });

  assert.deepEqual(sheet.getRange("A1:D4"), [
    ["Name", "Score", "Group", "Formula"],
    ["Alpha", 5, "High", 50],
    ["Gamma", 3, "Mid", 30],
    ["Beta", 2, "Low", 20],
  ]);
  assert.equal(sheet.getFormula("D2"), "B2*10");
  assert.equal(sheet.getFormula("D3"), "B3*10");
  assert.equal(sheet.getFormula("D4"), "B4*10");
  assert.equal(sheet.getBackgroundColor("A2"), "FF00FF00");
  assert.equal(sheet.getBackgroundColor("A3"), "FF0000FF");
  assert.equal(sheet.getBackgroundColor("A4"), "FFFF0000");
  assert.deepEqual(sheet.getMergedRanges(), ["C4:D4"]);
  assert.equal(sheet.getHyperlink("A4")?.target, "https://example.com/beta");
  assert.equal(sheet.getHyperlink("A2"), null);
  assert.deepEqual(sheet.getDataValidation("C4:D4"), {
    range: "C4:D4",
    type: "list",
    operator: null,
    allowBlank: null,
    showInputMessage: null,
    showErrorMessage: null,
    showDropDown: null,
    errorStyle: null,
    errorTitle: null,
    error: null,
    promptTitle: null,
    prompt: null,
    imeMode: null,
    formula1: "\"Low,Mid,High\"",
    formula2: null,
  });
  assert.deepEqual(sheet.getAutoFilterDefinition(), {
    range: "A1:D4",
    columns: [],
    sortState: {
      range: "A1:D4",
      conditions: [{ columnNumber: 2, descending: true }],
    },
  });
  assert.deepEqual(sheet.getTable("Scores").getAutoFilterDefinition(), {
    range: "A1:D4",
    columns: [],
    sortState: {
      range: "A1:D4",
      conditions: [{ columnNumber: 2, descending: true }],
    },
  });
  assert.equal(sheet.getTable("Scores").range, "A1:D4");
});

test("sheet sortRange supports multi-column sorting", () => {
  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRange("A1", [
    ["Name", "Score", "Group"],
    ["Beta", 2, "B"],
    ["Alpha", 2, "C"],
    ["Gamma", 1, "A"],
    ["Delta", 2, "A"],
  ]);

  sheet.sortRange("A1:C5", {
    conditions: [
      { columnNumber: 2 },
      { columnNumber: 3, descending: true },
    ],
    hasHeaderRow: true,
  });

  assert.deepEqual(sheet.getRange("A1:C5"), [
    ["Name", "Score", "Group"],
    ["Gamma", 1, "A"],
    ["Alpha", 2, "C"],
    ["Beta", 2, "B"],
    ["Delta", 2, "A"],
  ]);
});

test("sheet autoFilter and dataValidation APIs tolerate single-quoted worksheet tags", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = convertEntriesToSingleQuotedAttributes(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>
    <row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>5</v></c><c r="C2"><v>6</v></c></row>
  </sheetData>
  <autoFilter ref="A1:C2"/>
  <dataValidations count="1">
    <dataValidation sqref="A2" type="whole" allowBlank="1">
      <formula1>1</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>`,
    ),
    ["xl/worksheets/sheet1.xml"],
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getAutoFilter(), "A1:C2");
  assert.deepEqual(sheet.getDataValidations(), [
    {
      range: "A2",
      type: "whole",
      operator: null,
      allowBlank: true,
      showInputMessage: null,
      showErrorMessage: null,
      showDropDown: null,
      errorStyle: null,
      errorTitle: null,
      error: null,
      promptTitle: null,
      prompt: null,
      imeMode: null,
      formula1: "1",
      formula2: null,
    },
  ]);

  sheet.setAutoFilter("A1:D2");
  sheet.setDataValidation("C2", {
    type: "list",
    showDropDown: false,
    formula1: "\"Yes,No\"",
  });

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<autoFilter ref="A1:D2"\/>/);
  assert.match(sheetXml, /<dataValidations count="2">/);
  assert.match(
    sheetXml,
    /<dataValidation sqref="C2" type="list" showDropDown="0"><formula1>&quot;Yes,No&quot;<\/formula1><\/dataValidation>/,
  );

  sheet.removeAutoFilter();
  sheet.removeDataValidation("A2");

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.equal(sheet.getAutoFilter(), null);
  assert.deepEqual(sheet.getDataValidations(), [
    {
      range: "C2",
      type: "list",
      operator: null,
      allowBlank: null,
      showInputMessage: null,
      showErrorMessage: null,
      showDropDown: false,
      errorStyle: null,
      errorTitle: null,
      error: null,
      promptTitle: null,
      prompt: null,
      imeMode: null,
      formula1: "\"Yes,No\"",
      formula2: null,
    },
  ]);
  assert.doesNotMatch(sheetXml, /<autoFilter\b/);
  assert.match(
    sheetXml,
    /<dataValidations count="1"><dataValidation sqref="C2" type="list" showDropDown="0"><formula1>&quot;Yes,No&quot;<\/formula1><\/dataValidation><\/dataValidations>/,
  );
});

test("workbook defined name APIs read, write, and delete global and local names", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="GlobalValue">Sheet1!$A$1</definedName>
    <definedName name="LocalValue" localSheetId="0">$B$2</definedName>
  </definedNames>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="B2"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "GlobalValue", scope: null, value: "Sheet1!$A$1" },
    { hidden: false, name: "LocalValue", scope: "Sheet1", value: "$B$2" },
  ]);
  assert.equal(workbook.getDefinedName("GlobalValue"), "Sheet1!$A$1");
  assert.equal(workbook.getDefinedName("LocalValue", "Sheet1"), "$B$2");

  workbook.setDefinedName("GlobalValue", "Sheet1!$C$3");
  workbook.setDefinedName("NewLocal", "$D$4", { scope: "Sheet1" });

  let workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<definedName name="GlobalValue">Sheet1!\$C\$3<\/definedName>/);
  assert.match(workbookXml, /<definedName name="NewLocal" localSheetId="0">\$D\$4<\/definedName>/);

  workbook.deleteDefinedName("LocalValue", "Sheet1");
  workbook.deleteDefinedName("GlobalValue");

  workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.doesNotMatch(workbookXml, /LocalValue/);
  assert.doesNotMatch(workbookXml, /GlobalValue/);
  assert.match(workbookXml, /<definedName name="NewLocal" localSheetId="0">\$D\$4<\/definedName>/);
  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "NewLocal", scope: "Sheet1", value: "$D$4" },
  ]);
});

test("sheet print area and print titles APIs manage local defined names", async () => {
  const workbook = Workbook.create("Sheet1");
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getPrintArea(), null);
  assert.deepEqual(sheet.getPrintTitles(), { columns: null, rows: null });

  assert.equal(sheet.setPrintArea("A1:C4"), "A1:C4");
  assert.deepEqual(sheet.setPrintTitles({ rows: "1:2", columns: "A:B" }), { columns: "$A:$B", rows: "$1:$2" });

  assert.equal(sheet.getPrintArea(), "A1:C4");
  assert.deepEqual(sheet.getPrintTitles(), { columns: "$A:$B", rows: "$1:$2" });
  assert.equal(workbook.getDefinedName("_xlnm.Print_Area", "Sheet1"), "A1:C4");
  assert.equal(
    workbook.getDefinedName("_xlnm.Print_Titles", "Sheet1"),
    "Sheet1!$1:$2,Sheet1!$A:$B",
  );

  assert.deepEqual(sheet.setPrintTitles({ rows: null }), { columns: "$A:$B", rows: null });
  assert.deepEqual(sheet.getPrintTitles(), { columns: "$A:$B", rows: null });
  assert.equal(workbook.getDefinedName("_xlnm.Print_Titles", "Sheet1"), "Sheet1!$A:$B");

  assert.equal(sheet.setPrintArea(null), null);
  assert.deepEqual(sheet.setPrintTitles({ columns: null }), { columns: null, rows: null });
  assert.equal(sheet.getPrintArea(), null);
  assert.deepEqual(sheet.getPrintTitles(), { columns: null, rows: null });
  assert.equal(workbook.getDefinedName("_xlnm.Print_Area", "Sheet1"), null);
  assert.equal(workbook.getDefinedName("_xlnm.Print_Titles", "Sheet1"), null);
});

test("deleting the last defined name removes the definedNames container", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="OnlyOne">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.deleteDefinedName("OnlyOne");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.doesNotMatch(workbookXml, /<definedNames>/);
  assert.deepEqual(workbook.getDefinedNames(), []);
});

async function loadFixtureEntries(rootDirectory: string): Promise<Array<{ path: string; data: Uint8Array }>> {
  const entries: Array<{ path: string; data: Uint8Array }> = [];
  const stack = [rootDirectory];

  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }

    const names = await readdir(current);

    for (const name of names) {
      const absolutePath = join(current, name);
      const info = await stat(absolutePath);

      if (info.isDirectory()) {
        stack.push(absolutePath);
        continue;
      }

      const relativePath = absolutePath.slice(rootDirectory.length + 1).replaceAll("\\", "/");
      entries.push({
        path: relativePath,
        data: await readFile(absolutePath),
      });
    }
  }

  entries.sort((left, right) => left.path.localeCompare(right.path));
  return entries;
}

function toEntryMap(
  entries: Array<{ path: string; data: Uint8Array }>,
): Map<string, string> {
  return new Map(entries.map((entry) => [entry.path, Buffer.from(entry.data).toString("utf8")]));
}

function assertEntryMapsEqual(expected: Map<string, string>, actual: Map<string, string>): void {
  assert.deepEqual([...actual.keys()].sort(), [...expected.keys()].sort());

  for (const [path, text] of expected) {
    assert.equal(actual.get(path), text, `content mismatch for ${path}`);
  }
}

function entryText(entries: Array<{ path: string; data: Uint8Array }>, path: string): string {
  const entry = entries.find((candidate) => candidate.path === path);
  if (!entry) {
    throw new Error(`Missing entry: ${path}`);
  }

  return Buffer.from(entry.data).toString("utf8");
}

function replaceEntryText(
  entries: Array<{ path: string; data: Uint8Array }>,
  path: string,
  text: string,
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();
  let replaced = false;
  const nextEntries = entries.map((entry) => {
    if (entry.path !== path) {
      return entry;
    }

    replaced = true;
    return {
      path,
      data: encoder.encode(text),
    };
  });

  if (!replaced) {
    throw new Error(`Missing entry for replacement: ${path}`);
  }

  return nextEntries;
}

function convertEntriesToSingleQuotedAttributes(
  entries: Array<{ path: string; data: Uint8Array }>,
  paths: string[],
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();
  const selected = new Set(paths);

  return entries.map((entry) => {
    if (!selected.has(entry.path)) {
      return entry;
    }

    const text = Buffer.from(entry.data).toString("utf8");
    return {
      path: entry.path,
      data: encoder.encode(text.replace(/([A-Za-z_][\w:.-]*)="([^"]*)"/g, "$1='$2'")),
    };
  });
}

function summarizeCellEntries(entries: CellEntry[]): Array<{
  address: string;
  rowNumber: number;
  columnNumber: number;
  type: CellEntry["type"];
  value: CellEntry["value"];
}> {
  return entries.map((entry) => ({
    address: entry.address,
    rowNumber: entry.rowNumber,
    columnNumber: entry.columnNumber,
    type: entry.type,
    value: entry.value,
  }));
}

function withSecondSheet(
  entries: Array<{ path: string; data: Uint8Array }>,
  sheetXml: string,
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();

  return [
    ...replaceEntryText(
      replaceEntryText(
        replaceEntryText(
          entries,
          "xl/workbook.xml",
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
</workbook>`,
        ),
        "xl/_rels/workbook.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
      ),
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    ),
    {
      path: "xl/worksheets/sheet2.xml",
      data: encoder.encode(sheetXml),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
}

function withSheetTable(
  entries: Array<{ path: string; data: Uint8Array }>,
  tableXml: string,
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();

  return [
    ...replaceEntryText(
      entries,
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`,
    ),
    {
      path: "xl/worksheets/_rels/sheet1.xml.rels",
      data: encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`),
    },
    {
      path: "xl/tables/table1.xml",
      data: encoder.encode(tableXml),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
}
