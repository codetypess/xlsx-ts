import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdir, mkdtemp, readFile, readdir, rm, stat, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { runCli } from "../src/cli.ts";
import { Workbook } from "../src/index.ts";

test("symlinked CLI entry prints help output", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const symlinkPath = join(tempRoot, "fastxlsx");
    await symlink(resolve("src/cli.ts"), symlinkPath);

    const result = spawnSync(process.execPath, ["--import", "tsx", symlinkPath, "--help"], {
      encoding: "utf8",
    });

    assert.equal(result.status, 0);
    assert.match(result.stdout, /Usage: fastxlsx \[options\] \[command\]/);
    assert.match(result.stdout, /display help for command/);
    assert.equal(result.stderr, "");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("cli prints the package version", async () => {
  const packageJson = JSON.parse(await readFile(resolve("package.json"), "utf8")) as { version: string };
  const result = await runCliCapture(["--version"]);

  assert.equal(result.exitCode, 0);
  assert.equal(result.stdout.trim(), packageJson.version);
  assert.equal(result.stderr, "");
});

test("inspect reports workbook structure as JSON", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const result = await runCliCapture(["inspect", inputPath]);

    assert.equal(result.exitCode, 0);

    const output = JSON.parse(result.stdout);
    assert.equal(output.file, inputPath);
    assert.equal(output.activeSheet, "Sheet1");
    assert.deepEqual(output.definedNames, []);
    assert.deepEqual(output.sheets, [
      {
        columnCount: 1,
        headers: ["Hello"],
        name: "Sheet1",
        physicalRangeRef: "A1",
        rangeRef: "A1",
        rowCount: 1,
        visibility: "visible",
      },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("get returns translated shared formulas for shared-formula followers", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
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
    const inputPath = join(tempRoot, "shared-formula.xlsx");
    await workbook.save(inputPath);

    const result = await runCliCapture(["get", inputPath, "--sheet", "Sheet1", "--cell", "B2"]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.formula, "C2+$C$1+E$1+$E3+SUM(G2:H3)");
    assert.equal(payload.type, "formula");
    assert.equal(payload.value, 2);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("create builds a new workbook through the direct CLI command", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const outputPath = join(tempRoot, "created.xlsx");
    const result = await runCliCapture([
      "create",
      outputPath,
      "--sheet",
      "Config",
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.output, outputPath);
    assert.deepEqual(payload.sheets, ["Config"]);

    const workbook = await Workbook.open(outputPath);
    const sheet = workbook.getSheet("Config");
    assert.deepEqual(workbook.getSheetNames(), ["Config"]);
    assert.equal(workbook.getActiveSheet().name, "Config");
    assert.equal(sheet.rowCount, 0);
    assert.equal(sheet.columnCount, 0);

    const validation = await runCliCapture(["validate", outputPath]);
    assert.equal(validation.exitCode, 0);
    assert.equal(JSON.parse(validation.stdout).ok, true);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("set writes a cell value to a new workbook and preserves the style id", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "set-output.xlsx");
    const result = await runCliCapture([
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--text",
      "World",
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.output, outputPath);
    assert.equal(payload.result.value, "World");
    assert.equal(payload.result.styleId, 1);

    const workbook = await Workbook.open(outputPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.equal(sheet.getCell("A1"), "World");
    assert.equal(sheet.getStyleId("A1"), 1);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("add-sheet creates a new worksheet through the direct CLI command", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "add-sheet-output.xlsx");
    const result = await runCliCapture([
      "add-sheet",
      inputPath,
      "--sheet",
      "Config",
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.sheets, ["Sheet1", "Config"]);

    const workbook = await Workbook.open(outputPath);
    assert.deepEqual(
      workbook.getSheets().map((sheet) => sheet.name),
      ["Sheet1", "Config"],
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("rename-sheet and delete-sheet manage worksheets through direct commands", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const withExtraSheetPath = join(tempRoot, "with-extra-sheet.xlsx");
    const renamedPath = join(tempRoot, "renamed.xlsx");
    const deletedPath = join(tempRoot, "deleted.xlsx");

    let result = await runCliCapture([
      "add-sheet",
      inputPath,
      "--sheet",
      "Scratch",
      "--output",
      withExtraSheetPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "rename-sheet",
      withExtraSheetPath,
      "--from",
      "Sheet1",
      "--to",
      "Config",
      "--output",
      renamedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "delete-sheet",
      renamedPath,
      "--sheet",
      "Scratch",
      "--output",
      deletedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(deletedPath);
    assert.deepEqual(
      workbook.getSheets().map((sheet) => sheet.name),
      ["Config"],
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("move-sheet reorders worksheets through the direct CLI command", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const withSecondSheetPath = join(tempRoot, "two-sheets.xlsx");
    const movedPath = join(tempRoot, "moved.xlsx");

    let result = await runCliCapture([
      "add-sheet",
      inputPath,
      "--sheet",
      "Sheet2",
      "--output",
      withSecondSheetPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "move-sheet",
      withSecondSheetPath,
      "--sheet",
      "Sheet2",
      "--index",
      "0",
      "--output",
      movedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).sheets, ["Sheet2", "Sheet1"]);

    const workbook = await Workbook.open(movedPath);
    assert.deepEqual(
      workbook.getSheets().map((sheet) => sheet.name),
      ["Sheet2", "Sheet1"],
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented workbook active and visibility commands manage workbook metadata", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const withSecondSheetPath = join(tempRoot, "two-sheets.xlsx");
    const activePath = join(tempRoot, "active.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "add-sheet",
      inputPath,
      "--sheet",
      "Sheet2",
      "--output",
      withSecondSheetPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "workbook",
      "active",
      "set",
      withSecondSheetPath,
      "--sheet",
      "Sheet2",
      "--output",
      activePath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).activeSheet, "Sheet2");

    result = await runCliCapture([
      "workbook",
      "active",
      "get",
      activePath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).activeSheet, "Sheet2");

    result = await runCliCapture([
      "workbook",
      "visibility",
      "set",
      activePath,
      "--sheet",
      "Sheet1",
      "--visibility",
      "hidden",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).visibility, "hidden");

    result = await runCliCapture([
      "workbook",
      "visibility",
      "get",
      activePath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).visibility, "hidden");

    const workbook = await Workbook.open(activePath);
    assert.equal(workbook.getActiveSheet().name, "Sheet2");
    assert.equal(workbook.getSheetVisibility("Sheet1"), "hidden");
    assert.equal(workbook.getSheetVisibility("Sheet2"), "visible");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented workbook defined-name commands manage global and local names", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const withSecondSheetPath = join(tempRoot, "two-sheets.xlsx");
    const namesPath = join(tempRoot, "names.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "add-sheet",
      inputPath,
      "--sheet",
      "Sheet2",
      "--output",
      withSecondSheetPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "workbook",
      "defined-name",
      "set",
      withSecondSheetPath,
      "--name",
      "GlobalValue",
      "--value",
      "Sheet1!$A$1",
      "--output",
      namesPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).definedName, {
      hidden: false,
      name: "GlobalValue",
      scope: null,
      value: "Sheet1!$A$1",
    });

    result = await runCliCapture([
      "workbook",
      "defined-name",
      "set",
      namesPath,
      "--name",
      "LocalValue",
      "--value",
      "$B$2",
      "--scope",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).definedName, {
      hidden: false,
      name: "LocalValue",
      scope: "Sheet1",
      value: "$B$2",
    });

    result = await runCliCapture([
      "workbook",
      "defined-name",
      "list",
      namesPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).definedNames.length, 2);

    result = await runCliCapture([
      "workbook",
      "defined-name",
      "get",
      namesPath,
      "--name",
      "LocalValue",
      "--scope",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).definedName, {
      hidden: false,
      name: "LocalValue",
      scope: "Sheet1",
      value: "$B$2",
    });

    result = await runCliCapture([
      "workbook",
      "defined-name",
      "delete",
      namesPath,
      "--name",
      "GlobalValue",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).deleted, true);

    const workbook = await Workbook.open(namesPath);
    assert.equal(workbook.getDefinedName("GlobalValue"), null);
    assert.equal(workbook.getDefinedName("LocalValue", "Sheet1"), "$B$2");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("record commands manage header-based sheet data through the CLI", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const headersPath = join(tempRoot, "headers.xlsx");
    const recordsPath = join(tempRoot, "records.xlsx");
    const replacedPath = join(tempRoot, "replaced.xlsx");
    const deletedPath = join(tempRoot, "deleted.xlsx");

    let result = await runCliCapture([
      "set-headers",
      inputPath,
      "--sheet",
      "Sheet1",
      "--headers",
      '["Key","Value"]',
      "--output",
      headersPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "add-record",
      headersPath,
      "--sheet",
      "Sheet1",
      "--record",
      '{"Key":"alpha","Value":"1"}',
      "--output",
      recordsPath,
    ]);
    assert.equal(result.exitCode, 0);
    const addRecordPayload = JSON.parse(result.stdout);
    assert.equal(addRecordPayload.recordCount, 1);
    assert.equal("records" in addRecordPayload, false);

    result = await runCliCapture([
      "set-records",
      recordsPath,
      "--sheet",
      "Sheet1",
      "--records",
      '[{"Key":"alpha","Value":"10"},{"Key":"beta","Value":"20"}]',
      "--output",
      replacedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const setRecordsPayload = JSON.parse(result.stdout);
    assert.equal(setRecordsPayload.recordCount, 2);
    assert.equal("records" in setRecordsPayload, false);

    result = await runCliCapture([
      "delete-record",
      replacedPath,
      "--sheet",
      "Sheet1",
      "--row",
      "2",
      "--output",
      deletedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const deleteRecordPayload = JSON.parse(result.stdout);
    assert.equal(deleteRecordPayload.recordCount, 1);
    assert.equal("records" in deleteRecordPayload, false);

    result = await runCliCapture(["records", deletedPath, "--sheet", "Sheet1"]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).records, [{ Key: "beta", Value: "20" }]);

    const workbook = await Workbook.open(deletedPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getHeaders(), ["Key", "Value"]);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecords(), [{ Key: "beta", Value: "20" }]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("json and csv record commands import and export sheet records", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const jsonPath = join(tempRoot, "records.json");
    const csvPath = join(tempRoot, "records.csv");
    const fromJsonPath = join(tempRoot, "from-json.xlsx");
    const fromCsvPath = join(tempRoot, "from-csv.xlsx");

    await writeFile(
      jsonPath,
      `${JSON.stringify([{ id: 1001, name: "Alpha" }, { id: 1002, name: "Beta" }], null, 2)}\n`,
    );
    await writeFile(csvPath, 'id,name,notes\n1003,Gamma,"A, B"\n1004,Delta,"line 1\nline 2"\n');

    let result = await runCliCapture([
      "import-json",
      inputPath,
      "--sheet",
      "Sheet1",
      "--from-json",
      jsonPath,
      "--output",
      fromJsonPath,
    ]);
    assert.equal(result.exitCode, 0);
    const importJsonPayload = JSON.parse(result.stdout);
    assert.equal(importJsonPayload.recordCount, 2);
    assert.equal("records" in importJsonPayload, false);

    result = await runCliCapture([
      "export-json",
      fromJsonPath,
      "--sheet",
      "Sheet1",
      "--output",
      jsonPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(
      JSON.parse(await readFile(jsonPath, "utf8")),
      [{ id: 1001, name: "Alpha" }, { id: 1002, name: "Beta" }],
    );

    result = await runCliCapture([
      "import-csv",
      inputPath,
      "--sheet",
      "Sheet1",
      "--from-csv",
      csvPath,
      "--output",
      fromCsvPath,
    ]);
    assert.equal(result.exitCode, 0);
    const importCsvPayload = JSON.parse(result.stdout);
    assert.equal(importCsvPayload.recordCount, 2);
    assert.equal("records" in importCsvPayload, false);

    result = await runCliCapture([
      "export-csv",
      fromCsvPath,
      "--sheet",
      "Sheet1",
      "--output",
      csvPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(
      await readFile(csvPath, "utf8"),
      'id,name,notes\n1003,Gamma,"A, B"\n1004,Delta,"line 1\nline 2"\n',
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet commands handle import/export and comments", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const jsonPath = join(tempRoot, "sheet.json");
    const csvPath = join(tempRoot, "sheet.csv");
    const importedPath = join(tempRoot, "sheet-imported.xlsx");
    const commentedPath = join(tempRoot, "sheet-commented.xlsx");

    await writeFile(jsonPath, `${JSON.stringify([{ id: 1001, name: "Alpha" }], null, 2)}\n`);
    await writeFile(csvPath, "id,name\n1002,Beta\n");

    let result = await runCliCapture([
      "sheet",
      "import",
      inputPath,
      "--sheet",
      "Sheet1",
      "--format",
      "json",
      "--from",
      jsonPath,
      "--output",
      importedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).result.imported, 1);

    result = await runCliCapture([
      "sheet",
      "export",
      importedPath,
      "--sheet",
      "Sheet1",
      "--format",
      "csv",
      "--output",
      csvPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(await readFile(csvPath, "utf8"), "id,name\n1001,Alpha\n");

    result = await runCliCapture([
      "sheet",
      "comment",
      "set",
      importedPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "B2",
      "--text",
      "Note",
      "--author",
      "Alice",
      "--output",
      commentedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).comment, {
      address: "B2",
      author: "Alice",
      text: "Note",
    });

    result = await runCliCapture([
      "sheet",
      "comment",
      "list",
      commentedPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).comments, [
      {
        address: "B2",
        author: "Alice",
        text: "Note",
      },
    ]);

    result = await runCliCapture([
      "sheet",
      "comment",
      "get",
      commentedPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "B2",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).comment, {
      address: "B2",
      author: "Alice",
      text: "Note",
    });

    result = await runCliCapture([
      "sheet",
      "comment",
      "clear",
      commentedPath,
      "--sheet",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).cleared, 1);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet import update mode patches matched records only", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const jsonPath = join(tempRoot, "patch.json");
    const outputPath = join(tempRoot, "patched.xlsx");
    const workbook = Workbook.create("Sheet1");
    workbook.getSheet("Sheet1").setRecords([{ id: 1001, name: "Alpha", note: "first" }]);
    await workbook.save(inputPath);
    await writeFile(
      jsonPath,
      `${JSON.stringify([{ id: 1001, note: "patched" }, { id: 9999, note: "missing" }], null, 2)}\n`,
    );

    const result = await runCliCapture([
      "sheet",
      "import",
      inputPath,
      "--sheet",
      "Sheet1",
      "--format",
      "json",
      "--from",
      jsonPath,
      "--mode",
      "update",
      "--key-field",
      "id",
      "--output",
      outputPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).result, {
      headers: ["id", "name", "note"],
      imported: 2,
      inserted: 0,
      mode: "update",
      rowCount: 1,
      updated: 1,
    });

    const patchedWorkbook = await Workbook.open(outputPath);
    assert.deepEqual(patchedWorkbook.getSheet("Sheet1").getRecords(), [
      { id: 1001, name: "Alpha", note: "patched" },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet import/export commands support CSV formatting options", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const csvPath = join(tempRoot, "sheet.csv");
    const outputPath = join(tempRoot, "output.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    await writeFile(csvPath, " id , name \n 1001 , Alpha \n");

    result = await runCliCapture([
      "sheet",
      "import",
      inputPath,
      "--sheet",
      "Sheet1",
      "--format",
      "csv",
      "--from",
      csvPath,
      "--trim-values",
      "--output",
      outputPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).result.headers[0], "id");

    result = await runCliCapture([
      "sheet",
      "export",
      outputPath,
      "--sheet",
      "Sheet1",
      "--format",
      "csv",
      "--no-headers",
      "--crlf",
      "--output",
      csvPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(await readFile(csvPath, "utf8"), "1001,Alpha\r\n");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet record upsert replaces matched rows and update patches them", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const outputPath = join(tempRoot, "records-upsert.xlsx");
    const updatedPath = join(tempRoot, "records-update.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "records",
      "upsert",
      inputPath,
      "--sheet",
      "Sheet1",
      "--key-field",
      "id",
      "--record",
      '{"id":1001,"name":"Alpha","note":"first"}',
      "--output",
      outputPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).result, {
      inserted: true,
      record: { id: 1001, name: "Alpha", note: "first" },
      row: 2,
    });

    result = await runCliCapture([
      "sheet",
      "records",
      "upsert",
      outputPath,
      "--sheet",
      "Sheet1",
      "--key-field",
      "id",
      "--record",
      '{"id":1001,"name":"Alpha-2"}',
      "--output",
      updatedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).result, {
      inserted: false,
      record: { id: 1001, name: "Alpha-2" },
      row: 2,
    });

    result = await runCliCapture([
      "sheet",
      "records",
      "get",
      updatedPath,
      "--sheet",
      "Sheet1",
      "--key-field",
      "id",
      "--value",
      "1001",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).record, {
      id: 1001,
      name: "Alpha-2",
      note: null,
    });

    result = await runCliCapture([
      "sheet",
      "records",
      "update",
      updatedPath,
      "--sheet",
      "Sheet1",
      "--key-field",
      "id",
      "--value",
      "1001",
      "--record",
      '{"note":"restored"}',
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).result, {
      record: { note: "restored" },
      row: 2,
      updated: true,
    });

    result = await runCliCapture([
      "sheet",
      "records",
      "get",
      updatedPath,
      "--sheet",
      "Sheet1",
      "--key-field",
      "id",
      "--value",
      "1001",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).record, {
      id: 1001,
      name: "Alpha-2",
      note: "restored",
    });

    result = await runCliCapture([
      "sheet",
      "records",
      "delete",
      updatedPath,
      "--sheet",
      "Sheet1",
      "--key-field",
      "id",
      "--value",
      "1001",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).deleted, true);

    result = await runCliCapture([
      "sheet",
      "records",
      "get",
      updatedPath,
      "--sheet",
      "Sheet1",
      "--key-field",
      "id",
      "--value",
      "1001",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).record, null);

    const workbook = await Workbook.open(updatedPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getHeaders(), ["id", "name", "note"]);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecords(), []);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet record list, append, and replace commands manage record sets", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const appendedPath = join(tempRoot, "appended.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "records",
      "append",
      inputPath,
      "--sheet",
      "Sheet1",
      "--records",
      '[{"id":1001,"name":"Alpha"},{"id":1002,"name":"Beta"}]',
      "--output",
      appendedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const appendPayload = JSON.parse(result.stdout);
    assert.equal(appendPayload.appended, 2);
    assert.equal(appendPayload.recordCount, 2);
    assert.equal("headers" in appendPayload, false);
    assert.equal("records" in appendPayload, false);

    result = await runCliCapture([
      "sheet",
      "records",
      "list",
      appendedPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).records, [
      { id: 1001, name: "Alpha" },
      { id: 1002, name: "Beta" },
    ]);

    result = await runCliCapture([
      "sheet",
      "records",
      "replace",
      appendedPath,
      "--sheet",
      "Sheet1",
      "--record",
      '{"id":2001,"name":"Gamma"}',
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    const replacePayload = JSON.parse(result.stdout);
    assert.equal(replacePayload.replaced, 1);
    assert.equal(replacePayload.recordCount, 1);
    assert.equal("headers" in replacePayload, false);
    assert.equal("records" in replacePayload, false);

    const workbook = await Workbook.open(appendedPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecords(), [
      { id: 2001, name: "Gamma" },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet layout command updates widths, heights, freeze, and print settings", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const outputPath = join(tempRoot, "layout.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "layout",
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--column-widths",
      '{"A":12,"B":24}',
      "--row-heights",
      '{"1":22}',
      "--freeze-columns",
      "1",
      "--freeze-rows",
      "1",
      "--print-area",
      "A1:B20",
      "--print-title-rows",
      "1:1",
      "--output",
      outputPath,
    ]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.printArea, "A1:B20");
    assert.deepEqual(payload.printTitles, { columns: null, rows: "$1:$1" });
    assert.deepEqual(payload.freezePane, {
      activePane: "bottomRight",
      columnCount: 1,
      rowCount: 1,
      topLeftCell: "B2",
    });

    const workbook = await Workbook.open(outputPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.equal(sheet.getColumnWidth("A"), 12);
    assert.equal(sheet.getColumnWidth("B"), 24);
    assert.equal(sheet.getRowHeight(1), 22);
    assert.equal(sheet.getPrintArea(), "A1:B20");
    assert.deepEqual(sheet.getPrintTitles(), { columns: null, rows: "$1:$1" });

    result = await runCliCapture([
      "sheet",
      "layout",
      "get",
      outputPath,
      "--sheet",
      "Sheet1",
      "--columns",
      '["A","B"]',
      "--rows",
      "[1]",
    ]);
    assert.equal(result.exitCode, 0);
    const inspected = JSON.parse(result.stdout);
    assert.deepEqual(inspected.columnWidths, { A: 12, B: 24 });
    assert.deepEqual(inspected.rowHeights, { "1": 22 });
    assert.equal(inspected.printArea, "A1:B20");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet protection commands read, write, and clear protection", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const outputPath = join(tempRoot, "protected.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "protection",
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--sort",
      "--auto-filter",
      "--format-cells",
      "--delete-rows",
      "--pivot-tables",
      "--password-hash",
      "ABCD",
      "--output",
      outputPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).protection.passwordHash, "ABCD");

    result = await runCliCapture([
      "sheet",
      "protection",
      "get",
      outputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    const protection = JSON.parse(result.stdout).protection;
    assert.equal(protection.sheet, true);
    assert.equal(protection.sort, true);
    assert.equal(protection.autoFilter, true);
    assert.equal(protection.formatCells, true);
    assert.equal(protection.deleteRows, true);
    assert.equal(protection.pivotTables, true);

    result = await runCliCapture([
      "sheet",
      "protection",
      "clear",
      outputPath,
      "--sheet",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).protection, null);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet hyperlink and filter commands manage metadata", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const linkedPath = join(tempRoot, "linked.xlsx");
    const filteredPath = join(tempRoot, "filtered.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "hyperlink",
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--target",
      "https://example.com",
      "--text",
      "Open",
      "--tooltip",
      "Go",
      "--output",
      linkedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).hyperlink, {
      address: "A1",
      target: "https://example.com",
      tooltip: "Go",
      type: "external",
    });

    result = await runCliCapture([
      "sheet",
      "hyperlink",
      "list",
      linkedPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).hyperlinks, [
      {
        address: "A1",
        target: "https://example.com",
        tooltip: "Go",
        type: "external",
      },
    ]);

    result = await runCliCapture([
      "sheet",
      "hyperlink",
      "get",
      linkedPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).hyperlink.target, "https://example.com");

    result = await runCliCapture([
      "sheet",
      "filter",
      "set",
      linkedPath,
      "--sheet",
      "Sheet1",
      "--range",
      "A1:B3",
      "--output",
      filteredPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).autoFilter, "A1:B3");

    result = await runCliCapture([
      "sheet",
      "filter",
      "get",
      filteredPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).autoFilter, "A1:B3");

    result = await runCliCapture([
      "sheet",
      "hyperlink",
      "clear",
      filteredPath,
      "--sheet",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).cleared, 1);

    result = await runCliCapture([
      "sheet",
      "filter",
      "clear",
      filteredPath,
      "--sheet",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).autoFilter, null);

    const workbook = await Workbook.open(filteredPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.equal(sheet.getHyperlinks().length, 0);
    assert.equal(sheet.getAutoFilter(), null);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("structured sheet filter CLI commands manage autoFilter definitions and columns", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const filteredPath = join(tempRoot, "filtered.xlsx");
    const definition = {
      range: "A1:C4",
      columns: [
        { columnNumber: 1, kind: "values", values: ["Alpha"], includeBlank: true },
        { columnNumber: 2, kind: "top10", top: true, percent: false, value: 3 },
        { columnNumber: 3, kind: "blank", mode: "nonBlank" },
      ],
      sortState: {
        range: "A1:C4",
        conditions: [{ columnNumber: 2, descending: true }],
      },
    };

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "filter",
      "set-definition",
      inputPath,
      "--sheet",
      "Sheet1",
      "--definition",
      JSON.stringify(definition),
      "--output",
      filteredPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).autoFilterDefinition, definition);

    result = await runCliCapture([
      "sheet",
      "filter",
      "get",
      filteredPath,
      "--sheet",
      "Sheet1",
      "--definition",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).autoFilterDefinition, definition);

    result = await runCliCapture([
      "sheet",
      "filter",
      "set-column",
      filteredPath,
      "--sheet",
      "Sheet1",
      "--column",
      JSON.stringify({ columnNumber: 2, kind: "dynamic", type: "today" }),
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).autoFilterDefinition, {
      range: "A1:C4",
      columns: [
        { columnNumber: 1, kind: "values", values: ["Alpha"], includeBlank: true },
        { columnNumber: 2, kind: "dynamic", type: "today" },
        { columnNumber: 3, kind: "blank", mode: "nonBlank" },
      ],
      sortState: {
        range: "A1:C4",
        conditions: [{ columnNumber: 2, descending: true }],
      },
    });

    result = await runCliCapture([
      "sheet",
      "filter",
      "clear-columns",
      filteredPath,
      "--sheet",
      "Sheet1",
      "--column-numbers",
      JSON.stringify([1, 3]),
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).autoFilterDefinition, {
      range: "A1:C4",
      columns: [{ columnNumber: 2, kind: "dynamic", type: "today" }],
      sortState: {
        range: "A1:C4",
        conditions: [{ columnNumber: 2, descending: true }],
      },
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet selection and validation commands manage view metadata", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const selectedPath = join(tempRoot, "selected.xlsx");
    const validatedPath = join(tempRoot, "validated.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "selection",
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--active-cell",
      "C3",
      "--range",
      "C3:D4",
      "--output",
      selectedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).selection, {
      activeCell: "C3",
      pane: null,
      range: "C3:D4",
    });

    result = await runCliCapture([
      "sheet",
      "selection",
      "get",
      selectedPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).selection.activeCell, "C3");

    result = await runCliCapture([
      "sheet",
      "validation",
      "set",
      selectedPath,
      "--sheet",
      "Sheet1",
      "--range",
      "A2:A10",
      "--type",
      "whole",
      "--operator",
      "between",
      "--allow-blank",
      "true",
      "--show-error-message",
      "true",
      "--error-title",
      "Invalid",
      "--error",
      "Use 1-10",
      "--formula1",
      "1",
      "--formula2",
      "10",
      "--output",
      validatedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).validation.range, "A2:A10");

    result = await runCliCapture([
      "sheet",
      "validation",
      "list",
      validatedPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).validations.length, 1);

    result = await runCliCapture([
      "sheet",
      "validation",
      "get",
      validatedPath,
      "--sheet",
      "Sheet1",
      "--range",
      "A2:A10",
    ]);
    assert.equal(result.exitCode, 0);
    const validation = JSON.parse(result.stdout).validation;
    assert.equal(validation.type, "whole");
    assert.equal(validation.formula1, "1");
    assert.equal(validation.formula2, "10");

    result = await runCliCapture([
      "sheet",
      "selection",
      "clear",
      validatedPath,
      "--sheet",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).selection, null);

    result = await runCliCapture([
      "sheet",
      "validation",
      "clear",
      validatedPath,
      "--sheet",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).cleared, 1);

    const workbook = await Workbook.open(validatedPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.equal(sheet.getSelection(), null);
    assert.deepEqual(sheet.getDataValidations(), []);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("workflow-oriented sheet merge commands manage merged ranges", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const mergedPath = join(tempRoot, "merged.xlsx");

    let result = await runCliCapture([
      "create",
      inputPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "merge",
      "add",
      inputPath,
      "--sheet",
      "Sheet1",
      "--range",
      "B2:A1",
      "--output",
      mergedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).mergedRanges, ["A1:B2"]);

    result = await runCliCapture([
      "sheet",
      "merge",
      "list",
      mergedPath,
      "--sheet",
      "Sheet1",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).mergedRanges, ["A1:B2"]);

    result = await runCliCapture([
      "sheet",
      "merge",
      "remove",
      mergedPath,
      "--sheet",
      "Sheet1",
      "--range",
      "A1:B2",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).deleted, true);

    result = await runCliCapture([
      "sheet",
      "merge",
      "add",
      mergedPath,
      "--sheet",
      "Sheet1",
      "--range",
      "C3:D4",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "sheet",
      "merge",
      "clear",
      mergedPath,
      "--sheet",
      "Sheet1",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).cleared, 1);

    const workbook = await Workbook.open(mergedPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getMergedRanges(), []);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("add-record initializes headers on a newly created workbook", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const createdPath = join(tempRoot, "created.xlsx");
    const recordsPath = join(tempRoot, "records.xlsx");

    let result = await runCliCapture([
      "create",
      createdPath,
      "--sheet",
      "Config",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "add-record",
      createdPath,
      "--sheet",
      "Config",
      "--record",
      '{"Key":"alpha","Value":"1"}',
      "--output",
      recordsPath,
    ]);
    assert.equal(result.exitCode, 0);
    const createdRecordPayload = JSON.parse(result.stdout);
    assert.equal(createdRecordPayload.recordCount, 1);
    assert.equal("records" in createdRecordPayload, false);

    const workbook = await Workbook.open(recordsPath);
    assert.deepEqual(workbook.getSheet("Config").getHeaders(), ["Key", "Value"]);
    assert.deepEqual(workbook.getSheet("Config").getRecords(), [{ Key: "alpha", Value: "1" }]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("config-table command group supports high-level config workflows", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const initializedPath = join(tempRoot, "config-init.xlsx");
    const upsertedPath = join(tempRoot, "config-upserted.xlsx");
    const updatedPath = join(tempRoot, "config-updated.xlsx");
    const deletedPath = join(tempRoot, "config-deleted.xlsx");
    const replacedPath = join(tempRoot, "config-replaced.xlsx");

    let result = await runCliCapture([
      "config-table",
      "init",
      inputPath,
      "--sheet",
      "Config",
      "--headers",
      '["Key","Value","Note"]',
      "--output",
      initializedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "config-table",
      "upsert",
      initializedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--record",
      '{"Key":"timeout","Value":"30","Note":"seconds"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const configInsertPayload = JSON.parse(result.stdout);
    assert.equal(configInsertPayload.inserted, true);
    assert.equal(configInsertPayload.recordCount, 1);
    assert.equal(configInsertPayload.row, 2);
    assert.equal("records" in configInsertPayload, false);
    assert.equal("rows" in configInsertPayload, false);

    result = await runCliCapture([
      "config-table",
      "upsert",
      upsertedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--record",
      '{"Key":"timeout","Value":"60"}',
      "--output",
      updatedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const configUpsertPayload = JSON.parse(result.stdout);
    assert.equal(configUpsertPayload.inserted, false);
    assert.equal(configUpsertPayload.recordCount, 1);
    assert.equal(configUpsertPayload.row, 2);
    assert.equal("records" in configUpsertPayload, false);
    assert.equal("rows" in configUpsertPayload, false);

    result = await runCliCapture([
      "config-table",
      "get",
      updatedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--text",
      "timeout",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).record, {
      record: { Key: "timeout", Value: "60", Note: null },
      row: 2,
    });

    result = await runCliCapture([
      "config-table",
      "update",
      updatedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--text",
      "timeout",
      "--record",
      '{"Note":"restored"}',
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    const configUpdatePayload = JSON.parse(result.stdout);
    assert.equal(configUpdatePayload.updated, true);
    assert.equal(configUpdatePayload.recordCount, 1);
    assert.equal("records" in configUpdatePayload, false);
    assert.equal("rows" in configUpdatePayload, false);

    result = await runCliCapture([
      "config-table",
      "get",
      updatedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--text",
      "timeout",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).record, {
      record: { Key: "timeout", Value: "60", Note: "restored" },
      row: 2,
    });

    result = await runCliCapture([
      "config-table",
      "delete",
      updatedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--text",
      "timeout",
      "--output",
      deletedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const configDeletePayload = JSON.parse(result.stdout);
    assert.equal(configDeletePayload.deleted, true);
    assert.equal(configDeletePayload.recordCount, 0);
    assert.equal("records" in configDeletePayload, false);
    assert.equal("rows" in configDeletePayload, false);

    result = await runCliCapture([
      "config-table",
      "replace",
      deletedPath,
      "--sheet",
      "Config",
      "--records",
      '[{"Key":"region","Value":"cn"},{"Key":"retries","Value":"3"}]',
      "--output",
      replacedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const configReplacePayload = JSON.parse(result.stdout);
    assert.equal(configReplacePayload.recordCount, 2);
    assert.equal("records" in configReplacePayload, false);
    assert.equal("rows" in configReplacePayload, false);

    result = await runCliCapture([
      "config-table",
      "list",
      replacedPath,
      "--sheet",
      "Config",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 2, record: { Key: "region", Value: "cn", Note: null } },
      { row: 3, record: { Key: "retries", Value: "3", Note: null } },
    ]);

    const workbook = await Workbook.open(replacedPath);
    assert.deepEqual(workbook.getSheet("Config").getHeaders(), ["Key", "Value", "Note"]);
    assert.deepEqual(workbook.getSheet("Config").getRecords(), [
      { Key: "region", Value: "cn", Note: null },
      { Key: "retries", Value: "3", Note: null },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("config-table sync imports JSON config objects in replace, update, and upsert modes", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const replaceJsonPath = join(tempRoot, "replace.json");
    const replaceOutputPath = join(tempRoot, "sync-replace.xlsx");
    const upsertJsonPath = join(tempRoot, "upsert.json");
    const updateJsonPath = join(tempRoot, "update.json");

    await writeFile(
      replaceJsonPath,
      JSON.stringify(
        {
          timeout: "30",
          region: "cn",
        },
        null,
        2,
      ),
    );

    let result = await runCliCapture([
      "config-table",
      "sync",
      inputPath,
      "--sheet",
      "Config",
      "--from-json",
      replaceJsonPath,
      "--output",
      replaceOutputPath,
    ]);
    assert.equal(result.exitCode, 0);
    const configSyncReplacePayload = JSON.parse(result.stdout);
    assert.equal(configSyncReplacePayload.mode, "replace");
    assert.equal(configSyncReplacePayload.recordCount, 2);
    assert.equal("rows" in configSyncReplacePayload, false);

    result = await runCliCapture([
      "config-table",
      "init",
      replaceOutputPath,
      "--sheet",
      "Config",
      "--headers",
      '["Key","Value","Note"]',
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);

    await writeFile(
      upsertJsonPath,
      JSON.stringify(
        {
          timeout: "60",
          retries: "3",
        },
        null,
        2,
      ),
    );

    result = await runCliCapture([
      "config-table",
      "sync",
      replaceOutputPath,
      "--sheet",
      "Config",
      "--from-json",
      upsertJsonPath,
      "--mode",
      "upsert",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    const configSyncUpsertPayload = JSON.parse(result.stdout);
    assert.equal(configSyncUpsertPayload.mode, "upsert");
    assert.equal(configSyncUpsertPayload.recordCount, 3);
    assert.equal("rows" in configSyncUpsertPayload, false);

    await writeFile(
      updateJsonPath,
      `${JSON.stringify([{ Key: "timeout", Note: "patched" }, { Key: "missing", Note: "ignored" }], null, 2)}\n`,
    );

    result = await runCliCapture([
      "config-table",
      "sync",
      replaceOutputPath,
      "--sheet",
      "Config",
      "--from-json",
      updateJsonPath,
      "--mode",
      "update",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    const configSyncUpdatePayload = JSON.parse(result.stdout);
    assert.equal(configSyncUpdatePayload.mode, "update");
    assert.equal(configSyncUpdatePayload.recordCount, 3);
    assert.equal("rows" in configSyncUpdatePayload, false);

    result = await runCliCapture([
      "config-table",
      "list",
      replaceOutputPath,
      "--sheet",
      "Config",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 2, record: { Key: "timeout", Value: "60", Note: "patched" } },
      { row: 3, record: { Key: "region", Value: "cn", Note: null } },
      { row: 4, record: { Key: "retries", Value: "3", Note: null } },
    ]);

    const workbook = await Workbook.open(replaceOutputPath);
    assert.deepEqual(workbook.getSheet("Config").getHeaders(), ["Key", "Value", "Note"]);
    assert.deepEqual(workbook.getSheet("Config").getRecords(), [
      { Key: "timeout", Value: "60", Note: "patched" },
      { Key: "region", Value: "cn", Note: null },
      { Key: "retries", Value: "3", Note: null },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table command group respects explicit data row boundaries", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeStructuredTableWorkbook(tempRoot);
    const upsertedPath = join(tempRoot, "table-upsert.xlsx");
    const syncJsonPath = join(tempRoot, "table-sync.json");
    const syncedPath = join(tempRoot, "table-sync.xlsx");

    let result = await runCliCapture([
      "table",
      "inspect",
      inputPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).dataRowCount, 2);

    result = await runCliCapture([
      "table",
      "get",
      inputPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--key",
      "1001",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).row, {
      row: 6,
      record: { id: 1001, name: "Alpha" },
    });

    result = await runCliCapture([
      "table",
      "upsert",
      inputPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--record",
      '{"id":1002,"name":"Beta-2"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const tableUpsertPayload = JSON.parse(result.stdout);
    assert.equal(tableUpsertPayload.inserted, false);
    assert.equal(tableUpsertPayload.recordCount, 2);
    assert.equal("rows" in tableUpsertPayload, false);

    await writeFile(
      syncJsonPath,
      JSON.stringify(
        [
          { id: 1001, name: "Alpha-2" },
          { id: 1003, name: "Gamma" },
        ],
        null,
        2,
      ),
    );

    result = await runCliCapture([
      "table",
      "sync",
      upsertedPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--from-json",
      syncJsonPath,
      "--output",
      syncedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const tableSyncPayload = JSON.parse(result.stdout);
    assert.equal(tableSyncPayload.recordCount, 2);
    assert.equal("rows" in tableSyncPayload, false);

    result = await runCliCapture([
      "table",
      "list",
      syncedPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 6, record: { id: 1001, name: "Alpha-2" } },
      { row: 7, record: { id: 1003, name: "Gamma" } },
    ]);

    const workbook = await Workbook.open(syncedPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.deepEqual(sheet.getRow(2), ["int", "string"]);
    assert.deepEqual(sheet.getRow(3), [">>", "client"]);
    assert.deepEqual(sheet.getRow(4), ["!!!", "x"]);
    assert.deepEqual(sheet.getRow(5), ["###", "display"]);
    assert.deepEqual(sheet.getRecord(6, 1), { id: 1001, name: "Alpha-2" });
    assert.deepEqual(sheet.getRecord(7, 1), { id: 1003, name: "Gamma" });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table update patches structured rows without clearing omitted fields", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writePatchStructuredTableWorkbook(tempRoot);
    const upsertedPath = join(tempRoot, "table-upsert.xlsx");
    const updatedPath = join(tempRoot, "table-update.xlsx");
    const syncJsonPath = join(tempRoot, "table-update-sync.json");
    const syncedPath = join(tempRoot, "table-update-sync.xlsx");

    let result = await runCliCapture([
      "table",
      "upsert",
      inputPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--record",
      '{"id":1002,"name":"Beta-2"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const patchedUpsertPayload = JSON.parse(result.stdout);
    assert.equal(patchedUpsertPayload.inserted, false);
    assert.equal(patchedUpsertPayload.recordCount, 2);
    assert.equal("rows" in patchedUpsertPayload, false);

    result = await runCliCapture([
      "table",
      "update",
      upsertedPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--key",
      "1001",
      "--record",
      '{"note":"first-patched"}',
      "--output",
      updatedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const tableUpdatePayload = JSON.parse(result.stdout);
    assert.equal(tableUpdatePayload.updated, true);
    assert.equal(tableUpdatePayload.recordCount, 2);
    assert.equal("rows" in tableUpdatePayload, false);

    await writeFile(
      syncJsonPath,
      `${JSON.stringify([{ id: 1001, name: "Alpha-2" }, { id: 9999, note: "missing" }], null, 2)}\n`,
    );

    result = await runCliCapture([
      "table",
      "sync",
      updatedPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--from-json",
      syncJsonPath,
      "--mode",
      "update",
      "--output",
      syncedPath,
    ]);
    assert.equal(result.exitCode, 0);
    const tableUpdateSyncPayload = JSON.parse(result.stdout);
    assert.equal(tableUpdateSyncPayload.mode, "update");
    assert.equal(tableUpdateSyncPayload.recordCount, 2);
    assert.equal("rows" in tableUpdateSyncPayload, false);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table command group supports profile presets for structured sheets", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeStructuredTableWorkbook(tempRoot);
    const profilesPath = join(tempRoot, "table-profiles.json");
    const upsertedPath = join(tempRoot, "profile-upsert.xlsx");

    await writeFile(
      profilesPath,
      JSON.stringify(
        {
          profiles: {
            demo: {
              sheet: "Sheet1",
              headerRow: 1,
              dataStartRow: 6,
              keyFields: ["id"],
            },
          },
        },
        null,
        2,
      ),
    );

    let result = await runCliCapture([
      "table",
      "list",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 6, record: { id: 1001, name: "Alpha" } },
      { row: 7, record: { id: 1002, name: "Beta" } },
    ]);

    result = await runCliCapture([
      "table",
      "get",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
      "--key",
      "1002",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).row, {
      row: 7,
      record: { id: 1002, name: "Beta" },
    });

    result = await runCliCapture([
      "table",
      "upsert",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
      "--record",
      '{"id":1002,"name":"Beta-2"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.keyFields, ["id"]);
    assert.equal(payload.inserted, false);
    assert.equal(payload.recordCount, 2);
    assert.equal("rows" in payload, false);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("explicit table options override profile values", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeStructuredTableWorkbook(tempRoot);
    const profilesPath = join(tempRoot, "table-profiles.json");

    await writeFile(
      profilesPath,
      JSON.stringify(
        {
          profiles: {
            demo: {
              sheet: "Sheet1",
              headerRow: 1,
              dataStartRow: 7,
              keyFields: ["id"],
            },
          },
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "table",
      "list",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
      "--data-start-row",
      "6",
    ]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.dataStartRow, 6);
    assert.deepEqual(payload.rows, [
      { row: 6, record: { id: 1001, name: "Alpha" } },
      { row: 7, record: { id: 1002, name: "Beta" } },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table command group supports composite key profiles", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeCompositeStructuredTableWorkbook(tempRoot);
    const profilesPath = join(tempRoot, "table-profiles.json");
    const upsertedPath = join(tempRoot, "composite-profile-upsert.xlsx");

    await writeFile(
      profilesPath,
      JSON.stringify(
        {
          profiles: {
            defineLike: {
              sheet: "Sheet1",
              headerRow: 2,
              dataStartRow: 7,
              keyFields: ["key1", "key2"],
            },
          },
        },
        null,
        2,
      ),
    );

    let result = await runCliCapture([
      "table",
      "get",
      inputPath,
      "--profile",
      "defineLike",
      "--profiles-file",
      profilesPath,
      "--key",
      '{"key1":"TASK_TYPE","key2":"MAIN"}',
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).row, {
      row: 7,
      record: {
        id: "-",
        comment: "任务类型",
        key1: "TASK_TYPE",
        key2: "MAIN",
        value_comment: "主线任务",
        value: 1,
        value_type: "int",
        enum: "TaskType",
        enum_option: "true",
      },
    });

    result = await runCliCapture([
      "table",
      "upsert",
      inputPath,
      "--profile",
      "defineLike",
      "--profiles-file",
      profilesPath,
      "--record",
      '{"id":"-","comment":"任务类型","key1":"TASK_TYPE","key2":"MAIN","value_comment":"主线任务-新","value":10,"value_type":"int","enum":"TaskType","enum_option":"true"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(upsertedPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecord(7, 2), {
      id: "-",
      comment: "任务类型",
      key1: "TASK_TYPE",
      key2: "MAIN",
      value_comment: "主线任务-新",
      value: 10,
      value_type: "int",
      enum: "TaskType",
      enum_option: "true",
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles scans full workbooks and supports multiple xlsx inputs", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeProfileGenerationWorkbook(tempRoot);
    const secondInputPath = await writeCompositeStructuredTableWorkbook(join(tempRoot, "second"), "define.xlsx");
    const outputPath = join(tempRoot, "generated-profiles.json");
    const result = await runCliCapture([
      "table",
      "generate-profiles",
      inputPath,
      secondInputPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    assert.equal(result.stdout, `Generated profile file: ${outputPath}\n`);

    const outputPayload = JSON.parse(await readFile(outputPath, "utf8"));
    assert.deepEqual(outputPayload, {
      profiles: {
        "input#Sheet1": {
          sheet: "Sheet1",
          headerRow: 1,
          dataStartRow: 6,
          keyFields: ["id"],
        },
        "input#Config Values": {
          sheet: "Config Values",
          headerRow: 2,
          dataStartRow: 7,
          keyFields: ["key"],
        },
        "define#Sheet1": {
          sheet: "Sheet1",
          headerRow: 2,
          dataStartRow: 7,
          keyFields: ["key1", "key2"],
        },
      },
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles prints full profiles when no output file is provided", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeProfileGenerationWorkbook(tempRoot);
    const result = await runCliCapture([
      "table",
      "generate-profiles",
      inputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.output, null);
    assert.deepEqual(payload.profiles, {
      "input#Sheet1": {
        sheet: "Sheet1",
        headerRow: 1,
        dataStartRow: 6,
        keyFields: ["id"],
      },
      "input#Config Values": {
        sheet: "Config Values",
        headerRow: 2,
        dataStartRow: 7,
        keyFields: ["key"],
      },
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles can read xlsx inputs from a file list", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeProfileGenerationWorkbook(tempRoot);
    const secondInputPath = await writeCompositeStructuredTableWorkbook(join(tempRoot, "second"), "define.xlsx");
    const filesPath = join(tempRoot, "xlsx-files.txt");
    await writeFile(filesPath, `${inputPath}\n${secondInputPath}\n`);

    const result = await runCliCapture([
      "table",
      "generate-profiles",
      "--files-from",
      filesPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.files, [inputPath, secondInputPath]);
    assert.deepEqual(payload.profileNames, [
      "input#Sheet1",
      "input#Config Values",
      "define#Sheet1",
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles skips sheets with uninferable table profiles", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeProfileGenerationWorkbook(tempRoot);
    const workbook = await Workbook.open(inputPath);
    workbook.addSheet("Scratch");
    await workbook.save(inputPath);

    const result = await runCliCapture([
      "table",
      "generate-profiles",
      inputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.profileNames, [
      "input#Sheet1",
      "input#Config Values",
    ]);
    assert.deepEqual(payload.skipped, [
      {
        file: inputPath,
        reason: "Unable to infer table header row for sheet: Scratch",
        sheet: "Scratch",
      },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles skips duplicate generated profile names", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeCompositeStructuredTableWorkbook(join(tempRoot, "first"));
    const duplicateInputPath = await writeCompositeStructuredTableWorkbook(join(tempRoot, "second"));
    const result = await runCliCapture([
      "table",
      "generate-profiles",
      inputPath,
      duplicateInputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.profileNames, ["input#Sheet1"]);
    assert.deepEqual(payload.profiles, {
      "input#Sheet1": {
        sheet: "Sheet1",
        headerRow: 2,
        dataStartRow: 7,
        keyFields: ["key1", "key2"],
      },
    });
    assert.deepEqual(payload.skipped, [
      {
        file: duplicateInputPath,
        profileName: "input#Sheet1",
        reason: "Duplicate generated profile name: input#Sheet1",
        sheet: "Sheet1",
      },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles records workbook open failures", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeCompositeStructuredTableWorkbook(tempRoot);
    const badInputPath = join(tempRoot, "bad.xlsx");
    const outputPath = join(tempRoot, "profiles.json");
    await writeFile(badInputPath, "not a real workbook");

    const result = await runCliCapture([
      "table",
      "generate-profiles",
      inputPath,
      badInputPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);
    assert.equal(result.stdout, `Generated profile file: ${outputPath}\n`);

    const outputPayload = JSON.parse(await readFile(outputPath, "utf8"));
    assert.deepEqual(outputPayload.profiles, {
      "input#Sheet1": {
        sheet: "Sheet1",
        headerRow: 2,
        dataStartRow: 7,
        keyFields: ["key1", "key2"],
      },
    });
    assert.equal(outputPayload.skipped.length, 1);
    assert.equal(outputPayload.skipped[0].file, badInputPath);
    assert.equal(outputPayload.skipped[0].sheet, undefined);
    assert.equal(typeof outputPayload.skipped[0].reason, "string");
    assert.notEqual(outputPayload.skipped[0].reason.length, 0);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles can scan directories recursively and ignore files", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeProfileGenerationWorkbook(tempRoot);
    const secondInputPath = await writeCompositeStructuredTableWorkbook(join(tempRoot, "second"), "define.xlsx");
    const ignoredInputPath = await writeCompositeStructuredTableWorkbook(join(tempRoot, "ignored"), "ignored.xlsx");

    const result = await runCliCapture([
      "table",
      "generate-profiles",
      "--from-dir",
      tempRoot,
      "--ignore",
      ignoredInputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.files, [inputPath, secondInputPath]);
    assert.deepEqual(payload.profileNames, [
      "input#Sheet1",
      "input#Config Values",
      "define#Sheet1",
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles skips temporary workbook files during directory scans", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeProfileGenerationWorkbook(tempRoot);
    await writeFile(join(tempRoot, "~$input.xlsx"), "not a real workbook");
    await writeFile(join(tempRoot, "~scratch.xlsx"), "not a real workbook");

    const result = await runCliCapture([
      "table",
      "generate-profiles",
      "--from-dir",
      tempRoot,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.files, [inputPath]);
    assert.deepEqual(payload.profileNames, [
      "input#Sheet1",
      "input#Config Values",
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles infers composite keys from key1/key2 headers", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeCompositeStructuredTableWorkbook(tempRoot);
    const result = await runCliCapture([
      "table",
      "generate-profiles",
      inputPath,
    ]);

    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).profiles, {
      "input#Sheet1": {
        sheet: "Sheet1",
        headerRow: 2,
        dataStartRow: 7,
        keyFields: ["key1", "key2"],
      },
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table generate-profiles filters sheets by regular expression", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeProfileGenerationWorkbook(tempRoot);
    const secondInputPath = await writeCompositeStructuredTableWorkbook(join(tempRoot, "second"), "define.xlsx");
    const result = await runCliCapture([
      "table",
      "generate-profiles",
      inputPath,
      secondInputPath,
      "--sheet-filter",
      "^(Sheet1|conf)$",
    ]);

    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).profiles, {
      "input#Sheet1": {
        sheet: "Sheet1",
        headerRow: 1,
        dataStartRow: 6,
        keyFields: ["id"],
      },
      "define#Sheet1": {
        sheet: "Sheet1",
        headerRow: 2,
        dataStartRow: 7,
        keyFields: ["key1", "key2"],
      },
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("style commands update formatting and can copy styles through the CLI", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const withValuePath = join(tempRoot, "with-value.xlsx");
    const formattedPath = join(tempRoot, "formatted.xlsx");

    let result = await runCliCapture([
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "B1",
      "--text",
      "Tail",
      "--output",
      withValuePath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "set-background-color",
      withValuePath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--color",
      "FFFF0000",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "set-number-format",
      withValuePath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--format",
      "0.00%",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "copy-style",
      withValuePath,
      "--sheet",
      "Sheet1",
      "--from",
      "A1",
      "--to",
      "B1",
      "--output",
      formattedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.result.backgroundColor, "FFFF0000");
    assert.equal(payload.result.numberFormat, "0.00%");
    assert.equal(payload.result.value, "Tail");

    const workbook = await Workbook.open(formattedPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.equal(sheet.getBackgroundColor("B1"), "FFFF0000");
    assert.equal(sheet.getNumberFormat("B1")?.code, "0.00%");
    assert.equal(sheet.getCell("B1"), "Tail");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply executes structured workbook operations", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "apply-output.xlsx");
    const opsPath = join(tempRoot, "ops.json");

    await writeFile(
      opsPath,
      JSON.stringify(
        {
          actions: [
            { type: "renameSheet", from: "Sheet1", to: "Config" },
            { type: "setCell", sheet: "Config", cell: "A1", value: "Updated" },
            { type: "setDefinedName", name: "Greeting", value: "Config!$A$1" },
            { type: "setActiveSheet", sheet: "Config" },
          ],
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.actions, 4);
    assert.deepEqual(payload.sheets, ["Config"]);

    const workbook = await Workbook.open(outputPath);
    const sheet = workbook.getSheet("Config");
    assert.equal(sheet.getCell("A1"), "Updated");
    assert.equal(workbook.getDefinedName("Greeting"), "Config!$A$1");
    assert.equal(workbook.getActiveSheet().name, "Config");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply supports worksheet and style operations", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "apply-style-output.xlsx");
    const opsPath = join(tempRoot, "style-ops.json");

    await writeFile(
      opsPath,
      JSON.stringify(
        {
          actions: [
            { type: "addSheet", sheet: "Scratch" },
            { type: "renameSheet", from: "Sheet1", to: "Config" },
            { type: "setCell", sheet: "Config", cell: "B1", value: "Tail" },
            { type: "setBackgroundColor", sheet: "Config", cell: "A1", color: "FF00FF00" },
            { type: "setNumberFormat", sheet: "Config", cell: "A1", formatCode: "0.00%" },
            { type: "copyStyle", sheet: "Config", from: "A1", to: "B1" },
            { type: "deleteSheet", sheet: "Scratch" },
          ],
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(outputPath);
    assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Config"]);
    assert.equal(workbook.getSheet("Config").getBackgroundColor("B1"), "FF00FF00");
    assert.equal(workbook.getSheet("Config").getNumberFormat("B1")?.code, "0.00%");
    assert.equal(workbook.getSheet("Config").getCell("B1"), "Tail");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply supports record and header operations", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "apply-records-output.xlsx");
    const opsPath = join(tempRoot, "records-ops.json");

    await writeFile(
      opsPath,
      JSON.stringify(
        {
          actions: [
            { type: "setHeaders", sheet: "Sheet1", headers: ["Key", "Value"] },
            { type: "addRecords", sheet: "Sheet1", records: [{ Key: "a", Value: "1" }, { Key: "b", Value: "2" }] },
            { type: "deleteRecord", sheet: "Sheet1", row: 2 },
          ],
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(outputPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getHeaders(), ["Key", "Value"]);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecords(), [{ Key: "b", Value: "2" }]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("validate returns a successful roundtrip result for the fixture workbook", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const result = await runCliCapture(["validate", inputPath]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.input, inputPath);
    assert.equal(payload.ok, true);
    assert.deepEqual(payload.diffs, []);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("set rejects conflicting output modes", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "conflict.xlsx");
    const result = await runCliCapture([
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--text",
      "World",
      "--output",
      outputPath,
      "--in-place",
    ]);

    assert.equal(result.exitCode, 1);
    assert.equal(result.stdout, "");
    assert.match(result.stderr, /Use either --output or --in-place, not both/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("set rejects invalid JSON values", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "invalid-json.xlsx");
    const result = await runCliCapture([
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--value",
      "{",
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 1);
    assert.equal(result.stdout, "");
    assert.match(result.stderr, /Invalid JSON in --value/);
    await assert.rejects(stat(outputPath));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply requires an output path unless --in-place is used", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const opsPath = join(tempRoot, "ops.json");
    await writeFile(opsPath, "[]");

    const result = await runCliCapture(["apply", inputPath, "--ops", opsPath]);

    assert.equal(result.exitCode, 1);
    assert.equal(result.stdout, "");
    assert.match(result.stderr, /An output path is required; pass --output or use --in-place/);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply rejects invalid JSON ops documents", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const opsPath = join(tempRoot, "ops.json");
    const outputPath = join(tempRoot, "output.xlsx");
    await writeFile(opsPath, "{");

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 1);
    assert.equal(result.stdout, "");
    assert.match(result.stderr, /Invalid JSON in .*ops\.json/);
    await assert.rejects(stat(outputPath));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply rejects unsupported operation types", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const opsPath = join(tempRoot, "ops.json");
    const outputPath = join(tempRoot, "output.xlsx");
    await writeFile(opsPath, JSON.stringify([{ type: "explodeSheet", sheet: "Sheet1" }]));

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 1);
    assert.equal(result.stdout, "");
    assert.match(result.stderr, /Unsupported operation type/);
    await assert.rejects(stat(outputPath));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

async function runCliCapture(argv: string[]): Promise<{
  exitCode: number;
  stderr: string;
  stdout: string;
}> {
  let stdout = "";
  let stderr = "";
  const exitCode = await runCli(argv, {
    stderr: (chunk) => {
      stderr += chunk;
    },
    stdout: (chunk) => {
      stdout += chunk;
    },
  });

  return { exitCode, stderr, stdout };
}

async function writeFixtureWorkbook(rootDirectory: string, fileName = "input.xlsx"): Promise<string> {
  await mkdir(rootDirectory, { recursive: true });
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const inputPath = join(rootDirectory, fileName);
  await workbook.save(inputPath);
  return inputPath;
}

async function writeStructuredTableWorkbook(rootDirectory: string): Promise<string> {
  const inputPath = await writeFixtureWorkbook(rootDirectory);
  const workbook = await Workbook.open(inputPath);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["id", "name"]);
  sheet.setRow(2, ["int", "string"]);
  sheet.setRow(3, [">>", "client"]);
  sheet.setRow(4, ["!!!", "x"]);
  sheet.setRow(5, ["###", "display"]);
  sheet.setRow(6, [1001, "Alpha"]);
  sheet.setRow(7, [1002, "Beta"]);

  await workbook.save(inputPath);
  return inputPath;
}

async function writePatchStructuredTableWorkbook(rootDirectory: string): Promise<string> {
  const inputPath = await writeFixtureWorkbook(rootDirectory);
  const workbook = await Workbook.open(inputPath);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["id", "name", "note"]);
  sheet.setRow(2, ["int", "string", "string"]);
  sheet.setRow(3, [">>", "client", "client"]);
  sheet.setRow(4, ["!!!", "x", "x"]);
  sheet.setRow(5, ["###", "display", "display"]);
  sheet.setRow(6, [1001, "Alpha", "first"]);
  sheet.setRow(7, [1002, "Beta", "second"]);

  await workbook.save(inputPath);
  return inputPath;
}

async function writeCompositeStructuredTableWorkbook(rootDirectory: string, fileName = "input.xlsx"): Promise<string> {
  const inputPath = await writeFixtureWorkbook(rootDirectory, fileName);
  const workbook = await Workbook.open(inputPath);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["@define"]);
  sheet.setRow(2, [
    "id",
    "comment",
    "key1",
    "key2",
    "value_comment",
    "value",
    "value_type",
    "enum",
    "enum_option",
  ]);
  sheet.setRow(3, ["auto", "string?", "string", "string?", "string?", "@value_type", "string", "string?", "bool?"]);
  sheet.setRow(4, [">>", "client|server", null, null, null, null, null, null, null]);
  sheet.setRow(5, ["!!!", "x", "x", "x", "x", "x", "x", "x", "x"]);
  sheet.setRow(6, ["###", "注释", null, null, "注释", null, null, null, null]);
  sheet.setRow(7, ["-", "任务类型", "TASK_TYPE", "MAIN", "主线任务", 1, "int", "TaskType", "true"]);
  sheet.setRow(8, ["-", null, "TASK_TYPE", "BRANCH", "支线任务", 2, "int", null, null]);

  await workbook.save(inputPath);
  return inputPath;
}

async function writeProfileGenerationWorkbook(rootDirectory: string): Promise<string> {
  const inputPath = await writeStructuredTableWorkbook(rootDirectory);
  const workbook = await Workbook.open(inputPath);
  const sheet = workbook.addSheet("Config Values");

  sheet.setRow(1, ["@config"]);
  sheet.setRow(2, ["id", "key", "value", "value_type", "value_comment"]);
  sheet.setRow(3, ["auto", "string", "string", "string", "string"]);
  sheet.setRow(4, [">>", null, null, null, null]);
  sheet.setRow(5, ["!!!", "x", "x", "x", "x"]);
  sheet.setRow(6, ["###", "键", "值", "值类型", "描述"]);
  sheet.setRow(7, ["-", "FOO", 1, "int", "示例"]);

  await workbook.save(inputPath);
  return inputPath;
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

function replaceEntryText(
  entries: Array<{ path: string; data: Uint8Array }>,
  path: string,
  text: string,
): Array<{ path: string; data: Uint8Array }> {
  const nextEntries = entries.map((entry) =>
    entry.path === path
      ? { path, data: new TextEncoder().encode(text) }
      : entry,
  );

  return nextEntries.some((entry) => entry.path === path)
    ? nextEntries
    : [...nextEntries, { path, data: new TextEncoder().encode(text) }];
}
