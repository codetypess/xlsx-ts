import type { Sheet } from "../sheet.js";
import type {
  CellError,
  CellSnapshot,
  CellValue,
  RecalculateSummary,
} from "../types.js";
import type { Workbook } from "../workbook.js";
import { XlsxError } from "../errors.js";
import {
  makeCellAddress,
  normalizeCellAddress,
  parseRangeRef,
  splitCellAddress,
} from "../sheet/sheet-address.js";
import { buildSheetIndex } from "../sheet/sheet-index.js";
import { parseFormulaTagInfo } from "../sheet/sheet-formula-xml.js";
import { normalizeSheetNameKey } from "../workbook/workbook-sheet-helpers.js";

const FORMULA_ERROR_CODES: Record<string, number> = {
  "#NULL!": 0x00,
  "#DIV/0!": 0x07,
  "#VALUE!": 0x0f,
  "#REF!": 0x17,
  "#NAME?": 0x1d,
  "#NUM!": 0x24,
  "#N/A": 0x2a,
  "#GETTING_DATA": 0x2b,
  "#SPILL!": 0x2c,
  "#CALC!": 0x2d,
  "#FIELD!": 0x2e,
  "#BLOCKED!": 0x2f,
  "#UNKNOWN!": 0x30,
};

const FORMULA_ERROR_TEXTS = Object.keys(FORMULA_ERROR_CODES).sort((left, right) => right.length - left.length);

type AstNode =
  | { kind: "binary"; left: AstNode; operator: string; right: AstNode }
  | { kind: "boolean"; value: boolean }
  | { kind: "error"; errorText: string }
  | { kind: "function"; arguments: AstNode[]; name: string }
  | { kind: "name"; name: string }
  | { kind: "number"; value: number }
  | { kind: "range"; endAddress: string; sheetName: string | null; startAddress: string }
  | { kind: "reference"; address: string; sheetName: string | null }
  | { kind: "string"; value: string }
  | { kind: "unary"; operand: AstNode; operator: string };

type EvaluatedValue =
  | { height: number; kind: "range"; values: ScalarValue[]; width: number }
  | { kind: "scalar"; value: ScalarValue };

type RangeValue = Extract<EvaluatedValue, { kind: "range" }>;

interface FormulaCellDefinition {
  address: string;
  cellXml: string;
  formula: string;
  formulaAttributesSource: string;
  sheetName: string;
}

interface ScalarValue {
  error: CellError | null;
  value: CellValue;
}

interface Token {
  type:
    | "bang"
    | "colon"
    | "comma"
    | "eof"
    | "error"
    | "identifier"
    | "lparen"
    | "number"
    | "operator"
    | "quoted_identifier"
    | "rparen"
    | "string";
  value: string;
}

class FormulaParser {
  private index = 0;
  private readonly tokens: Token[];

  constructor(source: string) {
    this.tokens = tokenizeFormula(source);
  }

  parse(): AstNode {
    const expression = this.parseComparison();
    this.expect("eof");
    return expression;
  }

  private parseComparison(): AstNode {
    let node = this.parseConcat();

    while (this.matchOperator("=", "<>", "<", ">", "<=", ">=")) {
      const operator = this.previous().value;
      const right = this.parseConcat();
      node = { kind: "binary", left: node, operator, right };
    }

    return node;
  }

  private parseConcat(): AstNode {
    let node = this.parseAdditive();

    while (this.matchOperator("&")) {
      const operator = this.previous().value;
      const right = this.parseAdditive();
      node = { kind: "binary", left: node, operator, right };
    }

    return node;
  }

  private parseAdditive(): AstNode {
    let node = this.parseMultiplicative();

    while (this.matchOperator("+", "-")) {
      const operator = this.previous().value;
      const right = this.parseMultiplicative();
      node = { kind: "binary", left: node, operator, right };
    }

    return node;
  }

  private parseMultiplicative(): AstNode {
    let node = this.parsePower();

    while (this.matchOperator("*", "/")) {
      const operator = this.previous().value;
      const right = this.parsePower();
      node = { kind: "binary", left: node, operator, right };
    }

    return node;
  }

  private parsePower(): AstNode {
    let node = this.parseUnary();

    while (this.matchOperator("^")) {
      const operator = this.previous().value;
      const right = this.parseUnary();
      node = { kind: "binary", left: node, operator, right };
    }

    return node;
  }

  private parseUnary(): AstNode {
    if (this.matchOperator("+", "-")) {
      const operator = this.previous().value;
      return {
        kind: "unary",
        operand: this.parseUnary(),
        operator,
      };
    }

    return this.parsePrimary();
  }

  private parsePrimary(): AstNode {
    const token = this.peek();

    if (this.match("number")) {
      return { kind: "number", value: Number(token.value) };
    }

    if (this.match("string")) {
      return { kind: "string", value: token.value };
    }

    if (this.match("error")) {
      return { kind: "error", errorText: token.value };
    }

    if (this.match("lparen")) {
      const expression = this.parseComparison();
      this.expect("rparen");
      return expression;
    }

    let sheetName: string | null = null;
    if ((token.type === "identifier" || token.type === "quoted_identifier") && this.peek(1).type === "bang") {
      sheetName = token.value;
      this.advance();
      this.advance();
    }

    const primary = this.peek();
    if (primary.type !== "identifier") {
      throw new XlsxError(`Unsupported formula token: ${primary.value || primary.type}`);
    }

    if (sheetName === null && this.peek(1).type === "lparen") {
      const name = primary.value;
      this.advance();
      this.expect("lparen");
      const argumentsList: AstNode[] = [];
      if (!this.check("rparen")) {
        do {
          argumentsList.push(this.parseComparison());
        } while (this.match("comma"));
      }
      this.expect("rparen");
      return { kind: "function", arguments: argumentsList, name };
    }

    const identifier = primary.value;
    this.advance();

    if (sheetName !== null || isCellReferenceToken(identifier)) {
      const startAddress = normalizeReferenceAddress(identifier);
      if (this.match("colon")) {
        const nextToken = this.expect("identifier");
        const endAddress = normalizeReferenceAddress(nextToken.value);
        return { kind: "range", endAddress, sheetName, startAddress };
      }

      return { kind: "reference", address: startAddress, sheetName };
    }

    if (identifier.toUpperCase() === "TRUE") {
      return { kind: "boolean", value: true };
    }

    if (identifier.toUpperCase() === "FALSE") {
      return { kind: "boolean", value: false };
    }

    return { kind: "name", name: identifier };
  }

  private check(type: Token["type"]): boolean {
    return this.peek().type === type;
  }

  private expect(type: Token["type"]): Token {
    const token = this.peek();
    if (token.type !== type) {
      throw new XlsxError(`Expected ${type} in formula, received ${token.value || token.type}`);
    }

    this.index += 1;
    return token;
  }

  private match(type: Token["type"]): boolean {
    if (!this.check(type)) {
      return false;
    }

    this.index += 1;
    return true;
  }

  private matchOperator(...operators: string[]): boolean {
    const token = this.peek();
    if (token.type !== "operator" || !operators.includes(token.value)) {
      return false;
    }

    this.index += 1;
    return true;
  }

  private peek(offset = 0): Token {
    return this.tokens[this.index + offset] ?? this.tokens[this.tokens.length - 1]!;
  }

  private previous(): Token {
    return this.tokens[this.index - 1]!;
  }

  private advance(): Token {
    const token = this.peek();
    this.index += 1;
    return token;
  }
}

class FormulaRuntime {
  private readonly astCache = new Map<string, AstNode>();
  private readonly formulaCellCaches = new Map<string, Map<string, FormulaCellDefinition>>();
  private readonly formulaResults = new Map<string, ScalarValue>();
  private readonly nameResults = new Map<string, EvaluatedValue>();
  private readonly activeKeys: string[] = [];
  private readonly sheets = new Map<string, Sheet>();

  constructor(private readonly workbook: Workbook) {
    for (const sheet of workbook.getSheets()) {
      this.sheets.set(normalizeSheetNameKey(sheet.name), sheet);
    }
  }

  evaluateFormulaCell(sheetName: string, address: string): ScalarValue {
    const resolvedSheetName = this.resolveSheetName(sheetName);
    const key = makeFormulaKey(resolvedSheetName, address);
    const cached = this.formulaResults.get(key);
    if (cached) {
      return cached;
    }

    const definition = this.getFormulaCellDefinition(resolvedSheetName, address);
    if (!definition) {
      return scalarFromSnapshot(this.requireSheet(resolvedSheetName).readCellSnapshot(address));
    }

    if (definition.formulaAttributesSource.length > 0) {
      throw new XlsxError(
        `Unsupported formula shape at ${definition.sheetName}!${definition.address}: <f ${definition.formulaAttributesSource}>`,
      );
    }

    this.enterStack(key, `Circular formula reference: ${definition.sheetName}!${definition.address}`);
    try {
      const evaluated = this.evaluateExpression(definition.formula, definition.sheetName);
      const scalar = ensureScalarValue(evaluated);
      this.formulaResults.set(key, scalar);
      return scalar;
    } finally {
      this.leaveStack(key);
    }
  }

  evaluateExpression(expression: string, currentSheetName: string): EvaluatedValue {
    const resolvedSheetName = this.resolveSheetName(currentSheetName);
    const normalized = expression.startsWith("=") ? expression.slice(1) : expression;
    const ast = this.getParsedAst(normalized);
    return this.evaluateNode(ast, resolvedSheetName);
  }

  getFormulaResults(): Map<string, ScalarValue> {
    return this.formulaResults;
  }

  listFormulaAddresses(sheetName: string): string[] {
    return [...this.getFormulaCellsForSheet(this.resolveSheetName(sheetName)).keys()];
  }

  private enterStack(key: string, message: string): void {
    if (this.activeKeys.includes(key)) {
      throw new XlsxError(message);
    }

    this.activeKeys.push(key);
  }

  private evaluateFunction(node: Extract<AstNode, { kind: "function" }>, currentSheetName: string): EvaluatedValue {
    const name = node.name.toUpperCase();

    switch (name) {
      case "IF": {
        if (node.arguments.length < 2 || node.arguments.length > 3) {
          throw new XlsxError("IF expects two or three arguments");
        }

        const condition = ensureScalarValue(this.evaluateNode(node.arguments[0]!, currentSheetName));
        if (condition.error !== null) {
          return scalarResult(condition);
        }

        const branchIndex = toBoolean(condition) ? 1 : 2;
        return node.arguments[branchIndex]
          ? this.evaluateNode(node.arguments[branchIndex]!, currentSheetName)
          : scalarResult({ error: null, value: null });
      }
      case "LEN": {
        if (node.arguments.length !== 1) {
          throw new XlsxError("LEN expects one argument");
        }

        const value = ensureScalarValue(this.evaluateNode(node.arguments[0]!, currentSheetName));
        if (value.error !== null) {
          return scalarResult(value);
        }

        return scalarResult({
          error: null,
          value: stringifyScalar(value).length,
        });
      }
      case "NOT": {
        if (node.arguments.length !== 1) {
          throw new XlsxError("NOT expects one argument");
        }

        const value = ensureScalarValue(this.evaluateNode(node.arguments[0]!, currentSheetName));
        if (value.error !== null) {
          return scalarResult(value);
        }

        return scalarResult({ error: null, value: !toBoolean(value) });
      }
      case "MATCH":
        return this.evaluateMatchFunction(node, currentSheetName);
      case "VLOOKUP":
        return this.evaluateVlookupFunction(node, currentSheetName);
    }

    const evaluatedArguments = node.arguments.map((argument) => this.evaluateNode(argument, currentSheetName));
    const flattened = evaluatedArguments.flatMap(flattenValues);
    const firstError = findFirstError(flattened);
    if (firstError) {
      return scalarResult(firstError);
    }

    switch (name) {
      case "AND":
        return scalarResult({
          error: null,
          value: evaluatedArguments.every((argument) => toBoolean(ensureScalarValue(argument))),
        });
      case "AVERAGE": {
        const numbers = collectNumericValues(evaluatedArguments, { ignoreText: true, includeBlankAsZero: false });
        if (!Array.isArray(numbers)) {
          return scalarResult(numbers);
        }
        if (numbers.length === 0) {
          return scalarResult(makeErrorValue("#DIV/0!"));
        }

        return scalarResult({ error: null, value: numbers.reduce((sum, value) => sum + value, 0) / numbers.length });
      }
      case "CONCAT":
      case "CONCATENATE":
        return scalarResult({
          error: null,
          value: flattened.map(stringifyScalar).join(""),
        });
      case "COUNT":
        return scalarResult({
          error: null,
          value: flattened.filter((value) => typeof value.value === "number").length,
        });
      case "COUNTA":
        return scalarResult({
          error: null,
          value: flattened.filter((value) => value.value !== null).length,
        });
      case "MAX": {
        const numbers = collectNumericValues(evaluatedArguments, { ignoreText: true, includeBlankAsZero: false });
        if (!Array.isArray(numbers)) {
          return scalarResult(numbers);
        }
        return scalarResult({ error: null, value: numbers.length === 0 ? 0 : Math.max(...numbers) });
      }
      case "MIN": {
        const numbers = collectNumericValues(evaluatedArguments, { ignoreText: true, includeBlankAsZero: false });
        if (!Array.isArray(numbers)) {
          return scalarResult(numbers);
        }
        return scalarResult({ error: null, value: numbers.length === 0 ? 0 : Math.min(...numbers) });
      }
      case "OR":
        return scalarResult({
          error: null,
          value: evaluatedArguments.some((argument) => toBoolean(ensureScalarValue(argument))),
        });
      case "SUM":
      {
        const numbers = collectNumericValues(evaluatedArguments, { ignoreText: true, includeBlankAsZero: true });
        if (!Array.isArray(numbers)) {
          return scalarResult(numbers);
        }

        return scalarResult({
          error: null,
          value: numbers.reduce((sum, value) => sum + value, 0),
        });
      }
      default:
        return scalarResult(makeErrorValue("#NAME?"));
    }
  }

  private evaluateMatchFunction(node: Extract<AstNode, { kind: "function" }>, currentSheetName: string): EvaluatedValue {
    if (node.arguments.length < 2 || node.arguments.length > 3) {
      throw new XlsxError("MATCH expects two or three arguments");
    }

    const lookupValue = ensureScalarValue(this.evaluateNode(node.arguments[0]!, currentSheetName));
    if (lookupValue.error !== null) {
      return scalarResult(lookupValue);
    }

    const lookupArray = coerceToRangeValue(this.evaluateNode(node.arguments[1]!, currentSheetName));
    const vector = extractLookupVector(lookupArray);
    if (!vector) {
      return scalarResult(makeErrorValue("#N/A"));
    }

    const matchType = resolveMatchType(
      node.arguments[2] ? ensureScalarValue(this.evaluateNode(node.arguments[2]!, currentSheetName)) : null,
    );
    if (typeof matchType !== "number") {
      return scalarResult(matchType);
    }

    const index =
      matchType === 0
        ? findExactLookupIndex(vector, lookupValue)
        : findApproximateLookupIndex(vector, lookupValue, matchType);

    return scalarResult(index === -1 ? makeErrorValue("#N/A") : { error: null, value: index + 1 });
  }

  private evaluateVlookupFunction(
    node: Extract<AstNode, { kind: "function" }>,
    currentSheetName: string,
  ): EvaluatedValue {
    if (node.arguments.length < 3 || node.arguments.length > 4) {
      throw new XlsxError("VLOOKUP expects three or four arguments");
    }

    const lookupValue = ensureScalarValue(this.evaluateNode(node.arguments[0]!, currentSheetName));
    if (lookupValue.error !== null) {
      return scalarResult(lookupValue);
    }

    const table = coerceToRangeValue(this.evaluateNode(node.arguments[1]!, currentSheetName));
    const columnIndex = resolveColumnIndex(
      ensureScalarValue(this.evaluateNode(node.arguments[2]!, currentSheetName)),
      table.width,
    );
    if (typeof columnIndex !== "number") {
      return scalarResult(columnIndex);
    }

    const rangeLookup = resolveRangeLookup(
      node.arguments[3] ? ensureScalarValue(this.evaluateNode(node.arguments[3]!, currentSheetName)) : null,
    );
    if (typeof rangeLookup !== "boolean") {
      return scalarResult(rangeLookup);
    }

    const firstColumn = extractTableColumn(table, 0);
    const rowIndex = rangeLookup
      ? findApproximateLookupIndex(firstColumn, lookupValue, 1)
      : findExactLookupIndex(firstColumn, lookupValue);

    return scalarResult(rowIndex === -1 ? makeErrorValue("#N/A") : getRangeValue(table, rowIndex, columnIndex - 1));
  }

  private evaluateName(node: Extract<AstNode, { kind: "name" }>, currentSheetName: string): EvaluatedValue {
    const key = `name:${currentSheetName}:${node.name.toUpperCase()}`;
    const cached = this.nameResults.get(key);
    if (cached) {
      return cached;
    }

    const definition = this.workbook.getDefinedName(node.name, currentSheetName) ?? this.workbook.getDefinedName(node.name);
    if (definition === null) {
      return scalarResult(makeErrorValue("#NAME?"));
    }

    this.enterStack(key, `Circular defined name reference: ${node.name}`);
    try {
      const evaluated = this.evaluateExpression(definition, currentSheetName);
      this.nameResults.set(key, evaluated);
      return evaluated;
    } finally {
      this.leaveStack(key);
    }
  }

  private evaluateNode(node: AstNode, currentSheetName: string): EvaluatedValue {
    switch (node.kind) {
      case "boolean":
        return scalarResult({ error: null, value: node.value });
      case "error":
        return scalarResult(makeErrorValue(node.errorText));
      case "function":
        return this.evaluateFunction(node, currentSheetName);
      case "name":
        return this.evaluateName(node, currentSheetName);
      case "number":
        return scalarResult({ error: null, value: node.value });
      case "range":
        return collectRangeValues(
          this,
          node.sheetName ?? currentSheetName,
          node.startAddress,
          node.endAddress,
        );
      case "reference":
        return scalarResult(this.readCellValue(node.sheetName ?? currentSheetName, node.address));
      case "string":
        return scalarResult({ error: null, value: node.value });
      case "unary":
        return scalarResult(evaluateUnary(node.operator, ensureScalarValue(this.evaluateNode(node.operand, currentSheetName))));
      case "binary":
        return scalarResult(
          evaluateBinary(
            node.operator,
            ensureScalarValue(this.evaluateNode(node.left, currentSheetName)),
            ensureScalarValue(this.evaluateNode(node.right, currentSheetName)),
          ),
        );
    }
  }

  private getFormulaCellDefinition(sheetName: string, address: string): FormulaCellDefinition | undefined {
    return this.getFormulaCellsForSheet(sheetName).get(normalizeCellAddress(address));
  }

  private getFormulaCellsForSheet(sheetName: string): Map<string, FormulaCellDefinition> {
    let definitions = this.formulaCellCaches.get(sheetName);
    if (definitions) {
      return definitions;
    }

    const sheet = this.requireSheet(sheetName);
    sheet.finalizeBatchWrite();

    const sheetXml = this.workbook.readEntryText(sheet.path);
    const sheetIndex = buildSheetIndex(this.workbook, sheetXml);
    definitions = new Map<string, FormulaCellDefinition>();

    for (const rowNumber of sheetIndex.rowNumbers) {
      const row = sheetIndex.rows.get(rowNumber);
      if (!row) {
        continue;
      }

      for (const cell of row.cells) {
        if (cell.snapshot.formula === null) {
          continue;
        }

        const cellXml = sheetXml.slice(cell.start, cell.end);
        const formulaInfo = parseFormulaTagInfo(cellXml);
        definitions.set(cell.address, {
          address: cell.address,
          cellXml,
          formula: formulaInfo.formula ?? cell.snapshot.formula,
          formulaAttributesSource: formulaInfo.attributesSource,
          sheetName,
        });
      }
    }

    this.formulaCellCaches.set(sheetName, definitions);
    return definitions;
  }

  private getParsedAst(expression: string): AstNode {
    let ast = this.astCache.get(expression);
    if (ast) {
      return ast;
    }

    ast = new FormulaParser(expression).parse();
    this.astCache.set(expression, ast);
    return ast;
  }

  private leaveStack(key: string): void {
    const activeKey = this.activeKeys.pop();
    if (activeKey !== key) {
      throw new XlsxError("Formula evaluation stack became inconsistent");
    }
  }

  private readCellValue(sheetName: string, address: string): ScalarValue {
    const resolvedSheetName = this.resolveSheetName(sheetName);
    const normalizedAddress = normalizeCellAddress(address);
    const sheet = this.requireSheet(resolvedSheetName);
    const snapshot = sheet.readCellSnapshot(normalizedAddress);
    if (snapshot.formula !== null) {
      return this.evaluateFormulaCell(resolvedSheetName, normalizedAddress);
    }

    return scalarFromSnapshot(snapshot);
  }

  private requireSheet(sheetName: string): Sheet {
    const sheet = this.sheets.get(normalizeSheetNameKey(sheetName));
    if (!sheet) {
      throw new XlsxError(`Sheet not found during formula evaluation: ${sheetName}`);
    }

    return sheet;
  }

  private resolveSheetName(sheetName: string): string {
    return this.requireSheet(sheetName).name;
  }
}

export function recalculateCellFormula(workbook: Workbook, sheet: Sheet, address: string): CellSnapshot {
  const normalizedAddress = normalizeCellAddress(address);
  const snapshot = sheet.readCellSnapshot(normalizedAddress);
  if (snapshot.formula === null) {
    return snapshot;
  }

  const runtime = createRuntime(workbook);
  runtime.evaluateFormulaCell(sheet.name, normalizedAddress);
  applyFormulaResults(workbook, runtime.getFormulaResults());
  return sheet.readCellSnapshot(normalizedAddress);
}

export function recalculateSheetFormulas(workbook: Workbook, sheet: Sheet): RecalculateSummary {
  const runtime = createRuntime(workbook);
  for (const address of runtime.listFormulaAddresses(sheet.name)) {
    runtime.evaluateFormulaCell(sheet.name, address);
  }

  return applyFormulaResults(workbook, runtime.getFormulaResults(), new Set([sheet.name]));
}

export function recalculateWorkbookFormulas(workbook: Workbook): RecalculateSummary {
  const runtime = createRuntime(workbook);
  const sheetNames = new Set(workbook.getSheetNames());
  for (const sheetName of sheetNames) {
    for (const address of runtime.listFormulaAddresses(sheetName)) {
      runtime.evaluateFormulaCell(sheetName, address);
    }
  }

  return applyFormulaResults(workbook, runtime.getFormulaResults(), sheetNames);
}

function applyFormulaResults(
  workbook: Workbook,
  results: Map<string, ScalarValue>,
  targetSheets?: Set<string>,
): RecalculateSummary {
  let updated = 0;

  workbook.batch((currentWorkbook) => {
    for (const [key, result] of results) {
      const { address, sheetName } = splitFormulaKey(key);
      const sheet = currentWorkbook.getSheet(sheetName);
      const snapshot = sheet.readCellSnapshot(address);

      if (!isFormulaSnapshotChanged(snapshot, result)) {
        continue;
      }

      updated += 1;
      sheet.applyRecalculatedFormulaValue(address, result.value, result.error);
    }
  });

  return {
    cells: results.size,
    sheets: targetSheets?.size ?? new Set([...results.keys()].map((key) => splitFormulaKey(key).sheetName)).size,
    updated,
  };
}

function collectNumericValues(
  values: EvaluatedValue[],
  options: { ignoreText: boolean; includeBlankAsZero: boolean },
): ScalarValue | number[] {
  const numbers: number[] = [];

  for (const value of values.flatMap(flattenValues)) {
    if (value.value === null) {
      if (options.includeBlankAsZero) {
        numbers.push(0);
      }
      continue;
    }

    if (typeof value.value === "number") {
      numbers.push(value.value);
      continue;
    }

    if (typeof value.value === "boolean") {
      numbers.push(value.value ? 1 : 0);
      continue;
    }

    const trimmed = value.value.trim();
    if (trimmed.length === 0) {
      if (options.includeBlankAsZero) {
        numbers.push(0);
      }
      continue;
    }

    const numericValue = Number(trimmed);
    if (Number.isFinite(numericValue)) {
      numbers.push(numericValue);
      continue;
    }

    if (!options.ignoreText) {
      return makeErrorValue("#VALUE!");
    }
  }

  return numbers;
}

function collectRangeValues(
  runtime: FormulaRuntime,
  sheetName: string,
  startAddress: string,
  endAddress: string,
): RangeValue {
  const start = splitCellAddress(startAddress);
  const end = splitCellAddress(endAddress);
  const range = parseRangeRef(`${makeCellAddress(start.rowNumber, start.columnNumber)}:${makeCellAddress(end.rowNumber, end.columnNumber)}`);
  const width = range.endColumn - range.startColumn + 1;
  const height = range.endRow - range.startRow + 1;
  const values: ScalarValue[] = [];

  for (let rowNumber = range.startRow; rowNumber <= range.endRow; rowNumber += 1) {
    for (let columnNumber = range.startColumn; columnNumber <= range.endColumn; columnNumber += 1) {
      values.push(runtime.evaluateFormulaCell(sheetName, makeCellAddress(rowNumber, columnNumber)));
    }
  }

  return rangeResult(values, width, height);
}

function createRuntime(workbook: Workbook): FormulaRuntime {
  for (const sheet of workbook.getSheets()) {
    sheet.finalizeBatchWrite();
  }

  return new FormulaRuntime(workbook);
}

function ensureScalarValue(value: EvaluatedValue): ScalarValue {
  if (value.kind === "scalar") {
    return value.value;
  }

  if (value.values.length === 1) {
    return value.values[0]!;
  }

  return makeErrorValue("#VALUE!");
}

function evaluateBinary(operator: string, left: ScalarValue, right: ScalarValue): ScalarValue {
  if (left.error !== null) {
    return left;
  }
  if (right.error !== null) {
    return right;
  }

  if (operator === "&") {
    return { error: null, value: stringifyScalar(left) + stringifyScalar(right) };
  }

  if (operator === "=" || operator === "<>" || operator === "<" || operator === ">" || operator === "<=" || operator === ">=") {
    const comparison = compareScalarValues(left, right);
    return {
      error: null,
      value:
        operator === "="
          ? comparison === 0
          : operator === "<>"
            ? comparison !== 0
            : operator === "<"
              ? comparison < 0
              : operator === ">"
                ? comparison > 0
                : operator === "<="
                  ? comparison <= 0
                  : comparison >= 0,
    };
  }

  const leftNumber = toNumber(left);
  if (leftNumber.error !== null) {
    return leftNumber;
  }

  const rightNumber = toNumber(right);
  if (rightNumber.error !== null) {
    return rightNumber;
  }

  switch (operator) {
    case "+":
      return { error: null, value: (leftNumber.value as number) + (rightNumber.value as number) };
    case "-":
      return { error: null, value: (leftNumber.value as number) - (rightNumber.value as number) };
    case "*":
      return { error: null, value: (leftNumber.value as number) * (rightNumber.value as number) };
    case "/":
      return rightNumber.value === 0 ? makeErrorValue("#DIV/0!") : { error: null, value: (leftNumber.value as number) / (rightNumber.value as number) };
    case "^":
      return { error: null, value: (leftNumber.value as number) ** (rightNumber.value as number) };
    default:
      throw new XlsxError(`Unsupported binary operator: ${operator}`);
  }
}

function evaluateUnary(operator: string, operand: ScalarValue): ScalarValue {
  if (operand.error !== null) {
    return operand;
  }

  const numericOperand = toNumber(operand);
  if (numericOperand.error !== null) {
    return numericOperand;
  }

  return {
    error: null,
    value: operator === "-" ? -(numericOperand.value as number) : numericOperand.value,
  };
}

function flattenValues(value: EvaluatedValue): ScalarValue[] {
  return value.kind === "range" ? value.values : [value.value];
}

function coerceToRangeValue(value: EvaluatedValue): RangeValue {
  return value.kind === "range" ? value : rangeResult([value.value], 1, 1);
}

function extractLookupVector(range: RangeValue): ScalarValue[] | null {
  return range.width === 1 || range.height === 1 ? range.values : null;
}

function extractTableColumn(range: RangeValue, columnIndex: number): ScalarValue[] {
  const values: ScalarValue[] = [];

  for (let rowIndex = 0; rowIndex < range.height; rowIndex += 1) {
    values.push(getRangeValue(range, rowIndex, columnIndex));
  }

  return values;
}

function findApproximateLookupIndex(values: ScalarValue[], lookupValue: ScalarValue, matchType: number): number {
  let bestIndex = -1;
  let bestValue: ScalarValue | undefined;

  for (const [index, candidate] of values.entries()) {
    if (candidate.error !== null) {
      continue;
    }

    const candidateComparison = compareLookupValues(candidate, lookupValue);
    if ((matchType > 0 && candidateComparison > 0) || (matchType < 0 && candidateComparison < 0)) {
      continue;
    }

    if (
      bestValue === undefined ||
      (matchType > 0 && compareLookupValues(candidate, bestValue) >= 0) ||
      (matchType < 0 && compareLookupValues(candidate, bestValue) <= 0)
    ) {
      bestIndex = index;
      bestValue = candidate;
    }
  }

  return bestIndex;
}

function findExactLookupIndex(values: ScalarValue[], lookupValue: ScalarValue): number {
  for (const [index, candidate] of values.entries()) {
    if (candidate.error !== null) {
      continue;
    }

    if (compareLookupValues(candidate, lookupValue) === 0) {
      return index;
    }
  }

  return -1;
}

function isCellReferenceToken(value: string): boolean {
  return /^\$?[A-Z]+\$?\d+$/i.test(value);
}

function isFormulaSnapshotChanged(snapshot: CellSnapshot, result: ScalarValue): boolean {
  const nextErrorText = result.error?.text ?? null;
  const currentErrorText = snapshot.error?.text ?? null;
  if (currentErrorText !== nextErrorText) {
    return true;
  }

  return snapshot.value !== result.value;
}

function makeErrorValue(errorText: string): ScalarValue {
  return {
    error: {
      code: FORMULA_ERROR_CODES[errorText] ?? null,
      text: errorText,
    },
    value: errorText,
  };
}

function makeFormulaKey(sheetName: string, address: string): string {
  return `${sheetName}::${normalizeCellAddress(address)}`;
}

function normalizeReferenceAddress(value: string): string {
  return normalizeCellAddress(value.replaceAll("$", ""));
}

function scalarFromSnapshot(snapshot: CellSnapshot): ScalarValue {
  if (snapshot.error !== null) {
    return {
      error: snapshot.error,
      value: snapshot.error.text,
    };
  }

  return {
    error: null,
    value: snapshot.value,
  };
}

function scalarResult(value: ScalarValue): EvaluatedValue {
  return { kind: "scalar", value };
}

function compareScalarValues(left: ScalarValue, right: ScalarValue): number {
  if (left.value === null && right.value === null) {
    return 0;
  }

  const leftNumber = toNumber(left, true);
  const rightNumber = toNumber(right, true);
  if (leftNumber.error === null && rightNumber.error === null) {
    return (leftNumber.value as number) - (rightNumber.value as number);
  }

  return compareTextValues(stringifyScalar(left), stringifyScalar(right));
}

function compareLookupValues(left: ScalarValue, right: ScalarValue): number {
  return compareScalarValues(left, right);
}

function splitFormulaKey(key: string): { address: string; sheetName: string } {
  const separatorIndex = key.indexOf("::");
  return {
    address: key.slice(separatorIndex + 2),
    sheetName: key.slice(0, separatorIndex),
  };
}

function stringifyScalar(value: ScalarValue): string {
  if (value.error !== null) {
    return value.error.text;
  }

  if (value.value === null) {
    return "";
  }

  if (typeof value.value === "boolean") {
    return value.value ? "TRUE" : "FALSE";
  }

  return String(value.value);
}

function toBoolean(value: ScalarValue): boolean {
  if (value.value === null) {
    return false;
  }

  if (typeof value.value === "boolean") {
    return value.value;
  }

  if (typeof value.value === "number") {
    return value.value !== 0;
  }

  const upper = value.value.trim().toUpperCase();
  if (upper === "TRUE") {
    return true;
  }
  if (upper === "FALSE" || upper.length === 0) {
    return false;
  }

  const numericValue = Number(upper);
  return Number.isFinite(numericValue) ? numericValue !== 0 : true;
}

function findFirstError(values: ScalarValue[]): ScalarValue | null {
  return values.find((value) => value.error !== null) ?? null;
}

function getRangeValue(range: RangeValue, rowIndex: number, columnIndex: number): ScalarValue {
  return range.values[rowIndex * range.width + columnIndex]!;
}

function rangeResult(values: ScalarValue[], width: number, height: number): RangeValue {
  return {
    height,
    kind: "range",
    values,
    width,
  };
}

function resolveColumnIndex(value: ScalarValue, width: number): ScalarValue | number {
  if (value.error !== null) {
    return value;
  }

  const numericValue = toNumber(value);
  if (numericValue.error !== null || typeof numericValue.value !== "number" || !Number.isFinite(numericValue.value)) {
    return makeErrorValue("#VALUE!");
  }

  const columnIndex = Math.trunc(numericValue.value);
  if (columnIndex < 1) {
    return makeErrorValue("#VALUE!");
  }
  if (columnIndex > width) {
    return makeErrorValue("#REF!");
  }

  return columnIndex;
}

function resolveMatchType(value: ScalarValue | null): ScalarValue | number {
  if (value === null) {
    return 1;
  }

  if (value.error !== null) {
    return value;
  }

  const numericValue = toNumber(value);
  if (numericValue.error !== null || typeof numericValue.value !== "number" || !Number.isFinite(numericValue.value)) {
    return makeErrorValue("#VALUE!");
  }

  if (numericValue.value === 0) {
    return 0;
  }

  return numericValue.value > 0 ? 1 : -1;
}

function resolveRangeLookup(value: ScalarValue | null): boolean | ScalarValue {
  if (value === null) {
    return true;
  }

  if (value.error !== null) {
    return value;
  }

  return toBoolean(value);
}

function compareTextValues(left: string, right: string): number {
  const leftText = left.toUpperCase();
  const rightText = right.toUpperCase();

  return leftText < rightText ? -1 : leftText > rightText ? 1 : 0;
}

function toNumber(value: ScalarValue, allowText = false): ScalarValue {
  if (value.error !== null) {
    return value;
  }

  if (value.value === null) {
    return { error: null, value: 0 };
  }

  if (typeof value.value === "number") {
    return value;
  }

  if (typeof value.value === "boolean") {
    return { error: null, value: value.value ? 1 : 0 };
  }

  const trimmed = value.value.trim();
  if (trimmed.length === 0) {
    return { error: null, value: 0 };
  }

  const numericValue = Number(trimmed);
  if (Number.isFinite(numericValue)) {
    return { error: null, value: numericValue };
  }

  return allowText ? { error: makeErrorValue("#VALUE!").error, value: null } : makeErrorValue("#VALUE!");
}

function tokenizeFormula(source: string): Token[] {
  const tokens: Token[] = [];
  let index = 0;

  while (index < source.length) {
    const character = source[index]!;

    if (/\s/.test(character)) {
      index += 1;
      continue;
    }

    if (character === "\"") {
      let value = "";
      index += 1;

      while (index < source.length) {
        const current = source[index]!;
        if (current === "\"") {
          if (source[index + 1] === "\"") {
            value += "\"";
            index += 2;
            continue;
          }

          index += 1;
          break;
        }

        value += current;
        index += 1;
      }

      tokens.push({ type: "string", value });
      continue;
    }

    if (character === "'") {
      let value = "";
      index += 1;

      while (index < source.length) {
        const current = source[index]!;
        if (current === "'") {
          if (source[index + 1] === "'") {
            value += "'";
            index += 2;
            continue;
          }

          index += 1;
          break;
        }

        value += current;
        index += 1;
      }

      tokens.push({ type: "quoted_identifier", value });
      continue;
    }

    const errorText = FORMULA_ERROR_TEXTS.find((candidate) =>
      source.slice(index, index + candidate.length).toUpperCase() === candidate,
    );
    if (errorText) {
      tokens.push({ type: "error", value: errorText });
      index += errorText.length;
      continue;
    }

    const twoCharacterOperator = source.slice(index, index + 2);
    if (twoCharacterOperator === "<=" || twoCharacterOperator === ">=" || twoCharacterOperator === "<>") {
      tokens.push({ type: "operator", value: twoCharacterOperator });
      index += 2;
      continue;
    }

    if ("=<>+-*/^&".includes(character)) {
      tokens.push({ type: "operator", value: character });
      index += 1;
      continue;
    }

    if (character === "(") {
      tokens.push({ type: "lparen", value: character });
      index += 1;
      continue;
    }
    if (character === ")") {
      tokens.push({ type: "rparen", value: character });
      index += 1;
      continue;
    }
    if (character === ",") {
      tokens.push({ type: "comma", value: character });
      index += 1;
      continue;
    }
    if (character === ":") {
      tokens.push({ type: "colon", value: character });
      index += 1;
      continue;
    }
    if (character === "!") {
      tokens.push({ type: "bang", value: character });
      index += 1;
      continue;
    }

    if (/\d/.test(character) || (character === "." && /\d/.test(source[index + 1] ?? ""))) {
      let end = index + 1;
      while (end < source.length && /[\d.]/.test(source[end]!)) {
        end += 1;
      }
      if (/[Ee]/.test(source[end] ?? "")) {
        let exponentEnd = end + 1;
        if (/[+-]/.test(source[exponentEnd] ?? "")) {
          exponentEnd += 1;
        }
        while (exponentEnd < source.length && /\d/.test(source[exponentEnd]!)) {
          exponentEnd += 1;
        }
        end = exponentEnd;
      }

      tokens.push({ type: "number", value: source.slice(index, end) });
      index = end;
      continue;
    }

    let end = index + 1;
    while (end < source.length && !/[\s()+\-*/^&,<>:=!]/.test(source[end]!)) {
      end += 1;
    }

    tokens.push({ type: "identifier", value: source.slice(index, end) });
    index = end;
  }

  tokens.push({ type: "eof", value: "" });
  return tokens;
}
