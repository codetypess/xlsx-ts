import type { WorkbookContext } from "./workbook-context.js";
import { parseSharedStrings } from "./shared-strings.js";
import { parseStylesXml, type StylesCache } from "./workbook-styles-parse.js";

export function resolveSharedStringsCache(
  currentCache: string[] | undefined,
  context: WorkbookContext,
  hasEntry: (path: string) => boolean,
  readEntryText: (path: string) => string,
): string[] {
  if (currentCache) {
    return currentCache;
  }

  const sharedStringsPath = context.sharedStringsPath;
  if (!sharedStringsPath || !hasEntry(sharedStringsPath)) {
    return [];
  }

  return parseSharedStrings(readEntryText(sharedStringsPath));
}

export function resolveStylesCache(
  currentCache: StylesCache | null | undefined,
  context: WorkbookContext,
  hasEntry: (path: string) => boolean,
  readEntryText: (path: string) => string,
): StylesCache | null {
  if (currentCache !== undefined) {
    return currentCache;
  }

  const stylesPath = context.stylesPath;
  if (!stylesPath || !hasEntry(stylesPath)) {
    return null;
  }

  return parseStylesXml(stylesPath, readEntryText(stylesPath));
}

export function removeEntryOrderPath(entryOrder: string[], path: string): void {
  const entryIndex = entryOrder.indexOf(path);
  if (entryIndex !== -1) {
    entryOrder.splice(entryIndex, 1);
  }
}

export function shouldResetWorkbookContext(context: WorkbookContext | undefined, path: string): boolean {
  return !!(
    context &&
    (context.workbookPath === path || context.workbookRelsPath === path)
  );
}

export function shouldResetStylesCache(context: WorkbookContext | undefined, path: string): boolean {
  return !!(
    context &&
    (context.stylesPath === path || context.workbookPath === path || context.workbookRelsPath === path)
  );
}

export function shouldResetSharedStringsCache(context: WorkbookContext | undefined, path: string): boolean {
  return !!(context && context.sharedStringsPath === path);
}
