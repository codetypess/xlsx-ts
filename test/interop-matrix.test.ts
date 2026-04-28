import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { Workbook, type CellValue, validateRoundtripFile } from "../src/index.ts";

interface MutationMatrixCase {
  cell: string;
  filePath: string;
  name: string;
  sheet: string;
  value: CellValue;
}

test("real workbook mutation matrix preserves sheet metadata and untouched package parts", async () => {
  const cases: MutationMatrixCase[] = [
    {
      cell: "B2",
      filePath: resolve("res/task.xlsx"),
      name: "task-conf",
      sheet: "conf",
      value: "interop-task",
    },
    {
      cell: "B2",
      filePath: resolve("res/event.xlsx"),
      name: "event",
      sheet: "event",
      value: "interop-event",
    },
    {
      cell: "B3",
      filePath: resolve("res/producers/openpyxl-sample.xlsx"),
      name: "openpyxl",
      sheet: "Data",
      value: 123,
    },
    {
      cell: "B3",
      filePath: resolve("res/producers/xlsxwriter-sample.xlsx"),
      name: "xlsxwriter",
      sheet: "Data",
      value: 123,
    },
  ];

  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-interop-matrix-"));

  try {
    for (const testCase of cases) {
      const workbook = await Workbook.open(testCase.filePath);
      const sheet = workbook.getSheet(testCase.sheet);
      const targetSheetPath = sheet.path;
      const beforeEntries = workbook.toEntries();
      const beforeSnapshot = collectSheetInteropSnapshot(workbook, sheet, testCase.cell);
      const outputPath = join(tempRoot, `${testCase.name}.xlsx`);

      sheet.setCell(testCase.cell, testCase.value);
      await workbook.save(outputPath);

      const editedWorkbook = await Workbook.open(outputPath);
      const editedSheet = editedWorkbook.getSheet(testCase.sheet);
      const changedPaths = diffEntryPaths(beforeEntries, editedWorkbook.toEntries());

      assert.deepEqual(
        collectSheetInteropSnapshot(editedWorkbook, editedSheet, testCase.cell),
        beforeSnapshot,
        `metadata drift for ${testCase.name}`,
      );
      assert.equal(editedSheet.getCell(testCase.cell), testCase.value, `edited value mismatch for ${testCase.name}`);
      assert.deepEqual(changedPaths, [targetSheetPath], `unexpected package diffs for ${testCase.name}`);
    }
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("synthetic rich-part workbook roundtrips and keeps drawing/comment/extLst parts during cell edits", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-rich-parts-"));

  try {
    const entries = buildRichInteropEntries(await loadFixtureEntries(fixtureDir));
    const inputPath = join(tempRoot, "rich-parts.xlsx");
    const outputPath = join(tempRoot, "rich-parts-edited.xlsx");
    await Workbook.fromEntries(entries).save(inputPath);

    const roundtrip = await validateRoundtripFile(inputPath);
    assert.equal(roundtrip.ok, true);
    assert.deepEqual(roundtrip.diffs, []);

    const workbook = await Workbook.open(inputPath);
    const sheet = workbook.getSheet("Sheet1");
    const beforeEntries = workbook.toEntries();

    sheet.setCell("A1", "Rich");
    await workbook.save(outputPath);

    const editedWorkbook = await Workbook.open(outputPath);
    const editedEntries = editedWorkbook.toEntries();
    const changedPaths = diffEntryPaths(beforeEntries, editedEntries);
    const editedSheetXml = entryText(editedEntries, "xl/worksheets/sheet1.xml");

    assert.equal(editedWorkbook.getSheet("Sheet1").getCell("A1"), "Rich");
    assert.deepEqual(changedPaths, ["xl/worksheets/sheet1.xml"]);
    assert.match(editedSheetXml, /<legacyDrawing r:id="rIdVml1"\/>/);
    assert.match(editedSheetXml, /<drawing r:id="rIdDrawing1"\/>/);
    assert.match(editedSheetXml, /<extLst>/);

    for (const path of [
      "xl/workbook.xml",
      "xl/comments1.xml",
      "xl/drawings/drawing1.xml",
      "xl/drawings/vmlDrawing1.vml",
      "xl/worksheets/_rels/sheet1.xml.rels",
      "[Content_Types].xml",
    ]) {
      assertEntryBufferEqual(beforeEntries, editedEntries, path);
    }
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("synthetic advanced autoFilter workbook saves and reloads structured filter definitions", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-advanced-filter-"));

  try {
    const filePath = join(tempRoot, "advanced-filter.xlsx");
    const workbook = Workbook.create("Sheet1");
    const sheet = workbook.getSheet("Sheet1");

    sheet.setRange("A1", [
      ["A", "B", "C", "D", "E"],
      [1, 2, 3, 4, 5],
    ]);
    sheet.setAutoFilterDefinition({
      range: "A1:E2",
      columns: [
        { columnNumber: 1, kind: "blank", mode: "nonBlank" },
        { columnNumber: 2, kind: "color", dxfId: 3, cellColor: true },
        { columnNumber: 3, kind: "dynamic", type: "today" },
        { columnNumber: 4, kind: "top10", top: false, percent: true, value: 10 },
        { columnNumber: 5, kind: "icon", iconSet: "3Arrows", iconId: 1 },
      ],
      sortState: {
        range: "A1:E2",
        conditions: [{ columnNumber: 2, descending: true }],
      },
    });

    await workbook.save(filePath);

    const reopened = await Workbook.open(filePath);
    assert.deepEqual(reopened.getSheet("Sheet1").getAutoFilterDefinition(), {
      range: "A1:E2",
      columns: [
        { columnNumber: 1, kind: "blank", mode: "nonBlank" },
        { columnNumber: 2, kind: "color", dxfId: 3, cellColor: true },
        { columnNumber: 3, kind: "dynamic", type: "today" },
        { columnNumber: 4, kind: "top10", top: false, percent: true, value: 10 },
        { columnNumber: 5, kind: "icon", iconSet: "3Arrows", iconId: 1 },
      ],
      sortState: {
        range: "A1:E2",
        conditions: [{ columnNumber: 2, descending: true }],
      },
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

function collectSheetInteropSnapshot(workbook: Workbook, sheet: ReturnType<Workbook["getSheet"]>, cell: string) {
  return {
    activeSheet: workbook.getActiveSheet().name,
    autoFilter: sheet.getAutoFilter(),
    autoFilterDefinition: sheet.getAutoFilterDefinition(),
    backgroundColor: sheet.getBackgroundColor(cell),
    dataValidations: sheet.getDataValidations(),
    definedNames: workbook.getDefinedNames(),
    freezePane: sheet.getFreezePane(),
    hyperlinks: sheet.getHyperlinks(),
    mergedRanges: sheet.getMergedRanges(),
    physicalRangeRef: sheet.getPhysicalRangeRef(),
    rangeRef: sheet.getRangeRef(),
    styleId: sheet.getStyleId(cell),
  };
}

function buildRichInteropEntries(
  entries: Array<{ path: string; data: Uint8Array }>,
): Array<{ path: string; data: Uint8Array }> {
  let nextEntries = setEntryText(
    entries,
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="urn:fastxlsx:test">
  <workbookPr defaultThemeVersion="166925"/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <calcPr calcId="191029"/>
  <extLst>
    <ext uri="{11111111-2222-3333-4444-555555555555}">
      <mx:payload name="workbook-ext" value="keep"/>
    </ext>
  </extLst>
</workbook>`,
  );

  nextEntries = setEntryText(
    nextEntries,
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="urn:fastxlsx:test">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
  </sheetData>
  <legacyDrawing r:id="rIdVml1"/>
  <drawing r:id="rIdDrawing1"/>
  <extLst>
    <ext uri="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}">
      <mx:payload note="sheet-ext"/>
    </ext>
  </extLst>
</worksheet>`,
  );

  nextEntries = setEntryText(
    nextEntries,
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  );

  nextEntries = setEntryText(
    nextEntries,
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rIdVml1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`,
  );

  nextEntries = setEntryText(
    nextEntries,
    "xl/comments1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>fastxlsx</author></authors>
  <commentList>
    <comment ref="A1" authorId="0"><text><t>Hello comment</t></text></comment>
  </commentList>
</comments>`,
  );

  nextEntries = setEntryText(
    nextEntries,
    "xl/drawings/drawing1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 1"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr/>
      <xdr:txBody><a:bodyPr/><a:lstStyle/><a:p/></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
  );

  return setEntryText(
    nextEntries,
    "xl/drawings/vmlDrawing1.vml",
    `<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1"/>
  </o:shapelayout>
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/>
    <v:shadow on="t" color="black" obscured="t"/>
    <v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto"/>
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:Row>0</x:Row>
      <x:Column>0</x:Column>
    </x:ClientData>
  </v:shape>
</xml>`,
  );
}

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

function setEntryText(
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

  if (replaced) {
    return nextEntries;
  }

  return [...nextEntries, { path, data: encoder.encode(text) }]
    .sort((left, right) => left.path.localeCompare(right.path));
}

function entryText(entries: Array<{ path: string; data: Uint8Array }>, path: string): string {
  const entry = entries.find((candidate) => candidate.path === path);
  if (!entry) {
    throw new Error(`Missing entry: ${path}`);
  }

  return Buffer.from(entry.data).toString("utf8");
}

function assertEntryBufferEqual(
  leftEntries: Array<{ path: string; data: Uint8Array }>,
  rightEntries: Array<{ path: string; data: Uint8Array }>,
  path: string,
): void {
  const left = findEntryData(leftEntries, path);
  const right = findEntryData(rightEntries, path);
  assert.equal(Buffer.compare(left, right), 0, `content mismatch for ${path}`);
}

function diffEntryPaths(
  leftEntries: Array<{ path: string; data: Uint8Array }>,
  rightEntries: Array<{ path: string; data: Uint8Array }>,
): string[] {
  const leftMap = new Map(leftEntries.map((entry) => [entry.path, Buffer.from(entry.data)]));
  const rightMap = new Map(rightEntries.map((entry) => [entry.path, Buffer.from(entry.data)]));
  const allPaths = [...new Set([...leftMap.keys(), ...rightMap.keys()])].sort();
  const diffs: string[] = [];

  for (const path of allPaths) {
    const left = leftMap.get(path);
    const right = rightMap.get(path);
    if (!left || !right || Buffer.compare(left, right) !== 0) {
      diffs.push(path);
    }
  }

  return diffs;
}

function findEntryData(entries: Array<{ path: string; data: Uint8Array }>, path: string): Buffer {
  const entry = entries.find((candidate) => candidate.path === path);
  if (!entry) {
    throw new Error(`Missing entry: ${path}`);
  }

  return Buffer.from(entry.data);
}
