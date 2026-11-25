/**
 * PS Office: sorting a value-mode workbook moves its formulas with their
 * rows, as Excel's sort does. The grid of a large workbook holds only each
 * formula's cached value (the formulas are not live there), so the sort's
 * journal snapshot writes the moved formula itself: the formula the source
 * cell carried, with its references moved the way Excel moves them — same-
 * sheet references shift by the row distance like a copy, references written
 * with a sheet name and `$`-anchored rows keep their address (checked against
 * Excel through COM). Their results come from the built-in engine
 * (queueBuiltInRecalc).
 */
import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'

import type { LazyWorkbookState } from './univer-state'

/** The formula a cell carries now: a journaled edit's (none when it wrote a value), else the file's. */
export function currentFormulaAt(
  state: Pick<LazyWorkbookState, 'editJournal' | 'formulaText' | 'sharedFormulaGroups'>,
  sheetId: string,
  row: number,
  column: number,
): string | undefined {
  const key = `${row}:${column}`
  const journaled = state.editJournal.cells.get(sheetId)?.get(key)
  if (journaled?.hasValue) return journaled.formula
  return (
    state.formulaText.get(sheetId)?.get(key) ??
    state.sharedFormulaGroups.get(sheetId)?.textAt(row, column)
  )
}

/** A formula moved by a sort from `fromRow` to `toRow`, as Excel's sort rewrites it. */
export function sortMovedFormula(formula: string, fromRow: number, toRow: number): string {
  return offsetFormulaRefs(formula, toRow - fromRow, 0, { keepSheetQualified: true })
}

/**
 * Whether a value-mode sort may move the formulas of a sheet: every formula
 * must be known (the per-sheet index did not overflow), the coordinates must
 * be the file's (no inserted or deleted rows or columns this session), and
 * the built-in engine must be there to compute the moved formulas.
 */
export function canMoveFormulas(
  state: Pick<LazyWorkbookState, 'formulaTextTruncated' | 'editJournal'>,
  sheetId: string,
  hasBuiltInRecalc: boolean,
): boolean {
  if (!hasBuiltInRecalc) return false
  if (state.formulaTextTruncated.has(sheetId)) return false
  return (state.editJournal.structuralOps.get(sheetId)?.length ?? 0) === 0
}
