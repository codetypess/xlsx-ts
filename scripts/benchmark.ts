import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook } from "../src/index.js";

export interface BenchmarkResult {
  mode: "dense" | "sparse";
  runs: number[];
  averageMs: number;
  nonNull: number;
  visitedCells: number;
}

export interface BatchWriteBenchmarkResult {
  mode: "batch-write";
  runs: number[];
  averageMs: number;
  sheet: string;
  writes: number;
  targetRange: string;
  rowCount: number;
  columnCount: number;
}

export interface SheetBenchmarkSummary {
  name: string;
  rowCount: number;
  columnCount: number;
  usedRange: string | null;
  physicalCellCount: number;
  nonNull: number;
  maxPhysicalColumn: number;
  denseReadCount: number;
  denseAmplification: number;
}

export interface BenchmarkComparison {
  denseVisitedCells: number;
  sparseVisitedCells: number;
  denseAmplification: number;
}

export interface BenchmarkBaseline {
  expectedNonNull: number;
  maxAverageMs?: number;
  maxSparseAverageMs?: number;
  expectedBatchWriteSheet?: string;
  expectedBatchWriteWrites?: number;
  maxBatchWriteAverageMs?: number;
}

export async function runBenchmark(options: {
  filePath?: string;
  iterations?: number;
} = {}): Promise<{
  file: string;
  iterations: number;
  result: BenchmarkResult;
  sparseResult: BenchmarkResult;
  writeResult: BatchWriteBenchmarkResult;
  comparison: BenchmarkComparison;
  sheets: SheetBenchmarkSummary[];
}> {
  const filePath = options.filePath ?? resolve(process.cwd(), "res/monster.xlsx");
  const iterations = options.iterations ?? 3;
  const summary = await summarizeWorkbook(filePath);
  const result = await benchmark(iterations, "dense", () => benchmarkDenseWorkbook(filePath));
  const sparseResult = await benchmark(iterations, "sparse", () => benchmarkSparseWorkbook(filePath));
  const writeResult = await benchmarkBatchWrites(filePath, iterations);
  const comparison = {
    denseVisitedCells: result.visitedCells,
    sparseVisitedCells: sparseResult.visitedCells,
    denseAmplification:
      sparseResult.nonNull === 0 ? 0 : Number((result.visitedCells / sparseResult.nonNull).toFixed(2)),
  };

  return {
    file: filePath,
    iterations,
    result,
    sparseResult,
    writeResult,
    comparison,
    sheets: summary,
  };
}

async function benchmark(
  iterations: number,
  mode: BenchmarkResult["mode"],
  runOnce: () => Promise<{ nonNull: number; visitedCells: number }> | { nonNull: number; visitedCells: number },
): Promise<BenchmarkResult> {
  const runs: number[] = [];
  let nonNull = 0;
  let visitedCells = 0;

  for (let index = 0; index < iterations; index += 1) {
    const startedAt = performance.now();
    const run = await runOnce();
    nonNull = run.nonNull;
    visitedCells = run.visitedCells;
    runs.push(Number((performance.now() - startedAt).toFixed(1)));
  }

  return {
    mode,
    runs,
    averageMs: Number((runs.reduce((sum, value) => sum + value, 0) / runs.length).toFixed(1)),
    nonNull,
    visitedCells,
  };
}

async function benchmarkDenseWorkbook(filePath: string): Promise<{ nonNull: number; visitedCells: number }> {
  const workbook = await Workbook.open(filePath);
  let nonNull = 0;
  let visitedCells = 0;

  for (const sheet of workbook.getSheets()) {
    const rowCount = sheet.rowCount;
    const columnCount = sheet.columnCount;

    for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
      for (let columnNumber = 1; columnNumber <= columnCount; columnNumber += 1) {
        visitedCells += 1;
        const cell = sheet.getCell(rowNumber, columnNumber);
        if (cell !== null) {
          cell.toString();
          nonNull += 1;
        }
      }
    }
  }

  return { nonNull, visitedCells };
}

async function benchmarkSparseWorkbook(filePath: string): Promise<{ nonNull: number; visitedCells: number }> {
  const workbook = await Workbook.open(filePath);
  let nonNull = 0;
  let visitedCells = 0;

  for (const sheet of workbook.getSheets()) {
    for (const cell of sheet.iterCellEntries()) {
      visitedCells += 1;
      if (cell.value !== null) {
        cell.value.toString();
        nonNull += 1;
      }
    }
  }

  return { nonNull, visitedCells };
}

async function benchmarkBatchWrites(filePath: string, iterations: number): Promise<BatchWriteBenchmarkResult> {
  const runs: number[] = [];
  let summary:
    | Omit<BatchWriteBenchmarkResult, "averageMs" | "mode" | "runs">
    | undefined;

  for (let index = 0; index < iterations; index += 1) {
    const run = await benchmarkBatchWriteWorkbook(filePath);
    runs.push(run.elapsedMs);

    if (!summary) {
      summary = {
        columnCount: run.columnCount,
        rowCount: run.rowCount,
        sheet: run.sheet,
        targetRange: run.targetRange,
        writes: run.writes,
      };
    }
  }

  if (!summary) {
    throw new Error("Batch write benchmark did not produce any runs");
  }

  return {
    mode: "batch-write",
    runs,
    averageMs: Number((runs.reduce((sum, value) => sum + value, 0) / runs.length).toFixed(1)),
    ...summary,
  };
}

async function benchmarkBatchWriteWorkbook(filePath: string): Promise<{
  elapsedMs: number;
  sheet: string;
  writes: number;
  targetRange: string;
  rowCount: number;
  columnCount: number;
}> {
  const workbook = await Workbook.open(filePath);
  const sheet = selectBatchWriteTargetSheet(workbook);
  const scenario = resolveBatchWriteScenario(sheet);
  const startedAt = performance.now();

  sheet.batch((currentSheet) => {
    for (let offset = 0; offset < scenario.writes; offset += 1) {
      currentSheet.setCell(scenario.startRow + offset, 1, 900_000 + offset);
    }
  });

  return {
    elapsedMs: Number((performance.now() - startedAt).toFixed(1)),
    sheet: sheet.name,
    writes: scenario.writes,
    targetRange: scenario.targetRange,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
  };
}

function selectBatchWriteTargetSheet(workbook: Workbook) {
  const sheets = workbook.getSheets();
  const [firstSheet] = sheets;
  if (!firstSheet) {
    throw new Error("Workbook has no worksheets to benchmark");
  }

  let targetSheet = firstSheet;
  let maxPhysicalCellCount = firstSheet.getPhysicalCellEntries().length;

  for (let index = 1; index < sheets.length; index += 1) {
    const sheet = sheets[index]!;
    const physicalCellCount = sheet.getPhysicalCellEntries().length;
    if (physicalCellCount > maxPhysicalCellCount) {
      targetSheet = sheet;
      maxPhysicalCellCount = physicalCellCount;
    }
  }

  return targetSheet;
}

function resolveBatchWriteScenario(sheet: ReturnType<Workbook["getSheets"]>[number]): {
  startRow: number;
  writes: number;
  targetRange: string;
} {
  const startRow = sheet.rowCount >= 2 ? 2 : 1;
  const maxWritableExistingRows = Math.max(sheet.rowCount - startRow + 1, 1);
  const writes = Math.min(30, maxWritableExistingRows);
  const endRow = startRow + writes - 1;

  return {
    startRow,
    writes,
    targetRange: writes === 1 ? `A${startRow}` : `A${startRow}:A${endRow}`,
  };
}

async function summarizeWorkbook(filePath: string): Promise<SheetBenchmarkSummary[]> {
  const workbook = await Workbook.open(filePath);
  const summaries: SheetBenchmarkSummary[] = [];

  for (const sheet of workbook.getSheets()) {
    const physicalEntries = sheet.getPhysicalCellEntries();
    const logicalEntries = sheet.getCellEntries();
    const physicalCellCount = physicalEntries.length;
    const nonNull = logicalEntries.length;
    const maxPhysicalColumn = physicalEntries.reduce((currentMax, entry) => Math.max(currentMax, entry.columnNumber), 0);
    const denseReadCount = sheet.rowCount * sheet.columnCount;

    summaries.push({
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      usedRange: sheet.getRangeRef(),
      physicalCellCount,
      nonNull,
      maxPhysicalColumn,
      denseReadCount,
      denseAmplification: nonNull === 0 ? 0 : Number((denseReadCount / nonNull).toFixed(2)),
    });
  }

  return summaries;
}

async function main(): Promise<void> {
  const { filePathArg, iterationsArg, baselinePathArg } = parseCliArgs(process.argv.slice(2));
  const result = await runBenchmark({
    filePath: filePathArg ? resolve(process.cwd(), filePathArg) : undefined,
    iterations: iterationsArg ? Number(iterationsArg) : undefined,
  });

  if (baselinePathArg) {
    const baselinePath = resolve(process.cwd(), baselinePathArg);
    const baseline = JSON.parse(await readFile(baselinePath, "utf8")) as BenchmarkBaseline;
    const failures = validateAgainstBaseline(result, baseline);
    const output = {
      ...result,
      check: {
        ok: failures.length === 0,
        baseline: baselinePath,
        failures,
      },
    };

    if (failures.length > 0) {
      console.error(JSON.stringify(output, null, 2));
      process.exitCode = 1;
      return;
    }

    console.log(JSON.stringify(output, null, 2));
    return;
  }

  console.log(JSON.stringify(result, null, 2));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}

function parseCliArgs(args: string[]): {
  filePathArg?: string;
  iterationsArg?: string;
  baselinePathArg?: string;
} {
  const positional: string[] = [];
  let baselinePathArg: string | undefined;

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--check") {
      baselinePathArg = args[index + 1];
      if (!baselinePathArg) {
        throw new Error("Missing baseline path after --check");
      }
      index += 1;
      continue;
    }

    positional.push(argument);
  }

  return {
    filePathArg: positional[0],
    iterationsArg: positional[1],
    baselinePathArg,
  };
}

function validateAgainstBaseline(
  result: Awaited<ReturnType<typeof runBenchmark>>,
  baseline: BenchmarkBaseline,
): string[] {
  const failures: string[] = [];
  const local = result.result;

  if (local.nonNull !== baseline.expectedNonNull) {
    failures.push(`Expected nonNull=${baseline.expectedNonNull}, got ${local.nonNull}`);
  }

  if (baseline.maxAverageMs !== undefined && local.averageMs > baseline.maxAverageMs) {
    failures.push(`Average ${local.averageMs}ms exceeded ${baseline.maxAverageMs}ms`);
  }

  if (baseline.maxSparseAverageMs !== undefined && result.sparseResult.averageMs > baseline.maxSparseAverageMs) {
    failures.push(`Sparse average ${result.sparseResult.averageMs}ms exceeded ${baseline.maxSparseAverageMs}ms`);
  }

  if (baseline.expectedBatchWriteSheet !== undefined && result.writeResult.sheet !== baseline.expectedBatchWriteSheet) {
    failures.push(`Expected batch write sheet=${baseline.expectedBatchWriteSheet}, got ${result.writeResult.sheet}`);
  }

  if (baseline.expectedBatchWriteWrites !== undefined && result.writeResult.writes !== baseline.expectedBatchWriteWrites) {
    failures.push(`Expected batch write writes=${baseline.expectedBatchWriteWrites}, got ${result.writeResult.writes}`);
  }

  if (baseline.maxBatchWriteAverageMs !== undefined && result.writeResult.averageMs > baseline.maxBatchWriteAverageMs) {
    failures.push(`Batch write average ${result.writeResult.averageMs}ms exceeded ${baseline.maxBatchWriteAverageMs}ms`);
  }

  return failures;
}
