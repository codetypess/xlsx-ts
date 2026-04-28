# `fastxlsx` Specification-Driven Development Guide

## Purpose

This repository adopts `SDD` as `Specification-Driven Development`.

For `fastxlsx`, SDD is not a process document detached from the codebase. It is the rule that every non-trivial feature starts from an explicit behavior spec, then lands in code through a controlled sequence:

`spec -> internal model -> non-destructive writer -> structure sync -> public API/CLI -> tests -> docs`

The reason is simple: this project is built around `lossless-first`. If implementation starts directly from ad hoc XML rewriting, it is very easy to ship a feature that "works once" but silently drops styles, unknown nodes, relationship ordering, or metadata children during roundtrip.

## Repository Context

`fastxlsx` is organized around a few stable layers:

- `src/workbook.ts`, `src/sheet.ts`
  - Public workbook and worksheet APIs.
- `src/workbook/*`, `src/sheet/*`
  - Focused XML readers, writers, structure transforms, and package helpers.
- `src/cli/*`, `src/cli.ts`
  - Thin CLI routing over library behavior.
- `src/types.ts`, `src/index.ts`
  - Public type and export surface.
- `test/lossless.test.ts`
  - Primary lossless and mutation regression coverage.
- `test/real-files.test.ts`
  - Real workbook interoperability checks.
- `test/interop-matrix.test.ts`
  - Cross-surface snapshot coverage.
- `test/cli.test.ts`
  - CLI contract coverage.
- `test/xml-fuzz.test.ts`
  - XML parser and serializer robustness.
- `src/roundtrip.ts`, `scripts/validate-roundtrip.ts`
  - Roundtrip validation entrypoints.

## Core Principles

### 1. `Lossless-first` is the top-level constraint

Any feature that edits workbook XML must preserve untouched package parts byte-for-byte whenever possible.

If a feature needs to change one attribute in one tag, the implementation should target that attribute instead of reconstructing the whole part.

### 2. Public behavior is designed before XML patching starts

Every feature spec must define:

- the user-visible behavior
- the public API shape
- mutation semantics
- compatibility expectations
- failure behavior
- validation plan

The XML representation is an implementation detail, not the starting point.

### 3. Unknown or unsupported XML must survive roundtrip

When a node contains unsupported children, attributes, namespace markers, or `extLst`, changing one known field must not erase the rest.

If the implementation cannot preserve unknown content, the feature is not ready.

### 4. Structure-changing operations are part of the feature

For `fastxlsx`, a metadata feature is incomplete if it only supports read/write on a static workbook.

If the feature references rows, columns, ranges, or cells, the spec must also define behavior for:

- `insertRow`
- `deleteRow`
- `insertColumn`
- `deleteColumn`
- sheet rename or related reference rewrites, when applicable

### 5. CLI is a delivery surface, not the source of truth

Library semantics are defined first. CLI commands should expose those semantics with minimal translation logic.

If a feature exists only in CLI behavior and not as a coherent library capability, it is usually a design smell.

### 6. Tests are executable spec, not cleanup work

The spec must name the exact tests that prove the feature is safe:

- parsing tests
- write and overwrite tests
- roundtrip preservation tests
- structure transform tests
- real-file interoperability tests
- CLI tests, if the feature is exposed in CLI

## When SDD Is Required

Use an SDD work item for any change that touches one or more of the following:

- public API or TypeScript types
- CLI behavior or CLI JSON output
- workbook or worksheet metadata
- XML read/write helpers
- row or column structural transforms
- table behavior
- defined names, formulas, hyperlinks, styles, comments, validations, filters, merges, print settings, protection, or similar metadata systems
- any feature where lossless roundtrip could regress

Tiny typo fixes or comment-only changes do not need a separate feature SDD.

## Required Feature Spec Sections

Every medium or large feature should be written as a short SDD note in the issue, PR, or a dedicated file under `docs/spec/`.

At minimum, include the following sections.

### 1. Background

- What user scenario is blocked today?
- What current behavior is insufficient?
- Which existing APIs or code paths are involved?

### 2. Goals

- What must work in MVP?
- What is explicitly in scope for this change?

### 3. Non-goals

- What is intentionally deferred?
- What UI, editor, or advanced behavior is out of scope?

### 4. Public Surface

Define the exact external shape first:

- new methods
- changed methods
- new types
- return values
- error behavior
- CLI command additions or output changes

For public coordinate semantics, prefer stable domain terms such as absolute `columnNumber` over OOXML-relative fields like `colId`.

### 5. Internal Model

Describe the structured in-memory representation that bridges read and write paths.

If the feature currently exposes only a primitive, but the real workbook state is structured, introduce a structured internal model before expanding behavior.

### 6. OOXML Mapping

List the exact tags and attributes affected.

For each affected XML structure, define:

- what is parsed
- what is preserved verbatim
- what is rewritten
- what is left unsupported but must roundtrip unchanged

### 7. Mutation Semantics

Define how updates behave:

- full replace vs partial patch
- clear vs remove
- "set range but preserve existing children" vs "replace whole node"
- precedence when overlapping values exist

No setter should silently discard nested state unless its contract explicitly says it performs destructive replacement.

### 8. Structure Transform Semantics

Define what happens when rows or columns move relative to the feature.

This must include both:

- direct range refs such as `ref` or `sqref`
- deeper references such as column indexes, sort conditions, or child-node ranges

### 9. Compatibility and Migration

Clarify:

- whether the change is additive or breaking
- whether old APIs remain supported
- whether existing behavior must be preserved for callers that do not opt into the new feature

Prefer additive APIs over changing existing return types when compatibility matters.

### 10. Test Matrix

List the exact assertions to add, grouped by test suite.

### 11. Acceptance

State the minimal conditions for merge.

## Recommended Implementation Order

Unless there is a strong reason not to, implement features in this order:

1. Write the spec.
2. Identify all affected XML parts and structural transforms.
3. Add a structured parser for the current workbook state.
4. Add or refine internal types.
5. Implement non-destructive write helpers.
6. Implement row and column structure maintenance.
7. Expose public API.
8. Expose CLI, if needed.
9. Add regression coverage.
10. Update README or feature docs if the public surface changed.

This order is important. In this repository, writing setters before defining read model and transform behavior creates the highest regression risk.

## Project-specific Design Rules

### API Rules

- Prefer additive APIs for richer behavior.
  - Example: keep a legacy primitive getter if needed, but add a structured `get...Definition()` API for real feature work.
- Public APIs should express workbook semantics, not OOXML quirks.
- Use explicit method names for destructive behavior such as `clear...()` or `remove...()`.
- If user intent implies real data movement, metadata-only APIs are not enough.
  - Example: a future sort feature should not stop at writing `sortState` if callers expect rows to be reordered.

### XML Rewrite Rules

- Do not replace an existing container with a self-closing tag if that would drop supported or unsupported children.
- Preserve unknown attributes and child nodes unless the operation explicitly removes the whole feature.
- Prefer attribute-level updates or nested-tag replacement over rebuilding large XML fragments.
- Reuse existing helper patterns in `src/sheet/*`, `src/workbook/*`, and `src/utils/xml*.ts` before introducing new serializer logic.

### Structure Rules

- If Excel supports a concept at both worksheet level and table level, the spec must consider both.
- Any feature keyed by columns must define how insert/delete column affects its indexes.
- Any feature keyed by ranges must define how insert/delete row and column affects those ranges.
- Any feature that rewrites formulas, names, tables, validations, or metadata must be reviewed against existing structure rewrite helpers instead of assuming `ref` updates are sufficient.

### CLI Rules

- Keep CLI logic thin.
- CLI JSON output should remain deterministic.
- If CLI exposes a new structured feature, output should mirror library naming where practical.

## Test Strategy by Suite

Use the existing test layout deliberately.

### `test/lossless.test.ts`

Add coverage for:

- parse existing workbook state
- write new state
- rewrite existing state
- preserve unrelated XML
- insert/delete row and column behavior
- destructive clear/remove behavior

This is the main place for feature-specific XML regression tests.

### `test/real-files.test.ts`

Add or extend checks when the feature should read workbooks produced by real spreadsheet tools.

Use this when compatibility with producer-generated files matters.

### `test/interop-matrix.test.ts`

Add snapshot fields when the feature becomes part of the observable sheet/workbook surface.

Use this to prevent drift across multiple APIs reading the same workbook.

### `test/cli.test.ts`

Required when CLI commands are added or behavior changes.

### `test/xml-fuzz.test.ts`

Use when parser or serializer robustness is touched, especially for attribute handling, quoting, or nested tag rewriting.

## Definition of Done

A feature is ready to merge only when all applicable items below are true:

- The feature has a written spec with goals, non-goals, API, mutation semantics, structure semantics, and tests.
- The implementation preserves `lossless-first` expectations for untouched XML.
- Unknown or unsupported nested XML is not lost during partial edits.
- Row and column structure transforms are covered where relevant.
- Public types and exports are updated where relevant.
- CLI behavior is updated where relevant.
- README or user-facing docs are updated when the public surface changed.
- Tests covering the feature are present and passing.
- Build passes.

Baseline verification commands:

```bash
npm run build
npm test
```

Use these when relevant:

```bash
npm run validate:roundtrip -- path/to/file.xlsx
npm run pack:check
```

## Minimal SDD Template

Use the following template for new feature work:

```md
# Feature Name

## Background

## Goals

## Non-goals

## Public Surface

## Internal Model

## OOXML Mapping

## Mutation Semantics

## Structure Transform Semantics

## Compatibility

## Test Matrix

## Acceptance
```

## Current Default for `fastxlsx`

For future feature work in this repository, assume the following default stance unless the spec explicitly says otherwise:

- preserve first, rewrite second
- parse existing state before inventing new state
- partial patch before full replacement
- worksheet and table behavior must be reviewed together when the format supports both
- structure transforms are part of the feature, not follow-up cleanup
- tests are required before calling the feature done

This document is the default development contract for the repository.
