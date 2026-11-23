/**
 * PS Office: a built-in query refresh updates the open workbook in place, as
 * Excel's does — the window keeps its sheet, scroll and selection, and only
 * the refreshed tables and the formula results that use them change. The
 * refresh saves the file first (the file stays the single source of truth),
 * then the screen adopts a session on the saved file and re-reads just the
 * regions the save changed (adoptRefreshedWorkbook in univer-sync.ts).
 *
 * In place is taken only when the screen and the saved file can be shown to
 * match by re-reading those regions: nothing pending in the edit journal, no
 * engine-owned cells (formula mode, closure, recalc overlay), no running
 * preload, and the same sheets, tables and styles in both. Anything else
 * reopens the workbook, which is always exact.
 */
import type { IRange } from '@univerjs/core'

import type { WorkbookFile } from '../shared/desktop-api'
import type { LazyWorkbookState } from './univer-state'

export interface RefreshRegions {
  readonly sheet: string
  readonly ranges: readonly IRange[]
}

/** Why the refreshed file can't be adopted in place (the caller reopens it), or null. */
export function refreshInPlaceRefusal(
  state: LazyWorkbookState,
  reopened: WorkbookFile,
  pendingEdits: number,
): string | null {
  const file = state.file
  if (!file.path || reopened.path !== file.path) return 'another file'
  if (pendingEdits > 0) return 'edits are pending'
  if (state.formulaMode) return 'formulas are live in the grid'
  if (state.flags.preloadRunning) return 'the workbook is still loading'
  if (state.closure.status === 'active' || state.closure.status === 'pending') return 'closure mode'
  if (state.closure.pinned.size > 0) return 'closure mode'
  if (state.recalc.running || state.recalc.overlay.size > 0) return 'recalculated values on screen'
  if (state.streamAliases.size > 0) return 'duplicated sheets'
  for (const ops of state.editJournal.structuralOps.values()) {
    if (ops.length > 0) return 'rows or columns were inserted or deleted'
  }
  if (reopened.styles.length !== file.styles.length) return 'the styles changed'
  if (reopened.definedNames.length !== file.definedNames.length) return 'the defined names changed'
  for (const [index, name] of file.definedNames.entries()) {
    const next = reopened.definedNames[index]!
    if (next.name !== name.name || next.sheetIndex !== name.sheetIndex) {
      return 'the defined names changed'
    }
  }
  if (reopened.sheets.length !== file.sheets.length) return 'the sheets changed'
  for (const [index, sheet] of file.sheets.entries()) {
    const next = reopened.sheets[index]!
    if (next.id !== sheet.id || next.name !== sheet.name) return 'the sheets changed'
    if (next.columnCount > sheet.columnCount) return 'a sheet grew wider'
    if (next.tables.length !== sheet.tables.length) return 'the tables changed'
    for (const [tableIndex, table] of sheet.tables.entries()) {
      const after = next.tables[tableIndex]!
      if (after.name !== table.name || after.headerRowCount !== table.headerRowCount) {
        return 'the tables changed'
      }
    }
    if (next.pivotTables.length !== sheet.pivotTables.length) return 'the pivot tables changed'
  }
  return null
}

/** A range clamped to a sheet's extent, or null when it lies outside. */
export function clampToExtent(range: IRange, rows: number, columns: number): IRange | null {
  const endRow = Math.min(range.endRow, rows - 1)
  const endColumn = Math.min(range.endColumn, columns - 1)
  if (range.startRow > endRow || range.startColumn > endColumn) return null
  return { ...range, endRow, endColumn }
}

/** Row chunks of a range that fit one sidecar read (`maxCells`), top to bottom. */
export function rangeChunks(range: IRange, maxCells: number): IRange[] {
  const width = range.endColumn - range.startColumn + 1
  const rows = Math.max(1, Math.floor(maxCells / Math.max(1, width)))
  const out: IRange[] = []
  for (let startRow = range.startRow; startRow <= range.endRow; startRow += rows) {
    out.push({ ...range, startRow, endRow: Math.min(range.endRow, startRow + rows - 1) })
  }
  return out
}

/** Drops the `row:column` keys inside any of the ranges from a per-cell cache. */
export function dropKeysInRanges(cache: Map<string, unknown>, ranges: readonly IRange[]): void {
  if (ranges.length === 0 || cache.size === 0) return
  for (const key of [...cache.keys()]) {
    const colon = key.indexOf(':')
    const row = Number(key.slice(0, colon))
    const column = Number(key.slice(colon + 1))
    for (const range of ranges) {
      if (
        row >= range.startRow &&
        row <= range.endRow &&
        column >= range.startColumn &&
        column <= range.endColumn
      ) {
        cache.delete(key)
        break
      }
    }
  }
}

/** Defined names whose formula the refresh changed (Excel's hidden ExternalData_N follow their tables). */
export function changedDefinedNames(
  before: WorkbookFile['definedNames'],
  after: WorkbookFile['definedNames'],
): WorkbookFile['definedNames'] {
  const key = (name: { name: string; sheetIndex?: number | undefined }): string =>
    `${name.sheetIndex ?? ''}|${name.name.toLowerCase()}`
  const old = new Map(before.map((name) => [key(name), name.formula]))
  return after.filter((name) => old.has(key(name)) && old.get(key(name)) !== name.formula)
}
