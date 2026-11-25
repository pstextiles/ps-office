/**
 * PS Office: a value-mode sort moves formulas with their rows, as Excel's
 * sort does, journaling them for the save while the grid shows the built-in
 * engine's results.
 */
import { describe, expect, it } from 'vitest'

import {
  VALUE_MODE_FORMULA_KEY,
  createEditJournal,
  recordSetRangeValues,
} from '../src/renderer/edit-journal'
import { reorderGate } from '../src/renderer/reorder-gate'
import { canMoveFormulas, currentFormulaAt, sortMovedFormula } from '../src/renderer/sort-formulas'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

const SHEET = 'sheet-4'

describe('journaling a moved formula', () => {
  it('keeps the formula for the save and marks it as never installed in the grid', () => {
    const journal = createEditJournal()
    const [entry] = recordSetRangeValues(journal, SHEET, {
      9: { 1: { v: 46305, f: '=IF($I10="","",$I10)', [VALUE_MODE_FORMULA_KEY]: true } },
    })
    expect(entry).toMatchObject({
      formula: '=IF($I10="","",$I10)',
      valueModeFormula: true,
      value: null,
    })
  })

  it('stays a moved formula when the engine writes its result, and stops being one when typed over', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, SHEET, {
      9: { 1: { f: '=A10', [VALUE_MODE_FORMULA_KEY]: true } },
    })
    recordSetRangeValues(journal, SHEET, { 9: { 1: { v: 7 } } })
    expect(journal.cells.get(SHEET)?.get('9:1')).toMatchObject({
      formula: '=A10',
      valueModeFormula: true,
    })
    recordSetRangeValues(journal, SHEET, { 9: { 1: { f: '=B10' } } })
    expect(journal.cells.get(SHEET)?.get('9:1')?.valueModeFormula).toBeUndefined()
  })

  it('drops a formula journaled before when a sort moves a plain value in', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, SHEET, {
      9: { 1: { f: '=A10', [VALUE_MODE_FORMULA_KEY]: true } },
    })
    recordSetRangeValues(journal, SHEET, { 9: { 1: { v: 'Castlerock', f: null } } })
    expect(journal.cells.get(SHEET)?.get('9:1')).toMatchObject({ value: 'Castlerock' })
    expect(journal.cells.get(SHEET)?.get('9:1')?.formula).toBeUndefined()
  })
})

describe('the formula a cell carries before the sort', () => {
  const state = (journal = createEditJournal()) =>
    ({
      editJournal: journal,
      formulaText: new Map([[SHEET, new Map([['5:1', '=A6*10']])]]),
      sharedFormulaGroups: new Map(),
      formulaTextTruncated: new Set<string>(),
    }) as unknown as LazyWorkbookState

  it("is the file's, unless a journaled edit replaced it", () => {
    expect(currentFormulaAt(state(), SHEET, 5, 1)).toBe('=A6*10')
    const journal = createEditJournal()
    recordSetRangeValues(journal, SHEET, { 5: { 1: { v: 3 } } })
    expect(currentFormulaAt(state(journal), SHEET, 5, 1)).toBeUndefined()
    const moved = createEditJournal()
    recordSetRangeValues(moved, SHEET, {
      5: { 1: { f: '=A9*10', [VALUE_MODE_FORMULA_KEY]: true } },
    })
    expect(currentFormulaAt(state(moved), SHEET, 5, 1)).toBe('=A9*10')
  })

  it('moves by the row distance, as Excel sorts (row 4 to row 2 in the probe)', () => {
    expect(sortMovedFormula('=$A$2+A4', 3, 1)).toBe('=$A$2+A2')
    expect(sortMovedFormula('=Other!A7+A4', 3, 1)).toBe('=Other!A7+A2')
  })

  it('is allowed only with a complete formula index, file coordinates and the engine', () => {
    const base = state()
    expect(canMoveFormulas(base, SHEET, true)).toBe(true)
    expect(canMoveFormulas(base, SHEET, false)).toBe(false)
    const truncated = { ...base, formulaTextTruncated: new Set([SHEET]) } as LazyWorkbookState
    expect(canMoveFormulas(truncated, SHEET, true)).toBe(false)
  })
})

describe('the sort gate', () => {
  const input = {
    formulaMode: false,
    preloadComplete: true,
    preloadRunning: false,
    isAddedSheet: false,
    cellCounts: { gridCells: 1, storedCells: 1, valueCells: 1 },
    rangeCells: 180_000,
    rangeHasStreamedFormulas: () => true,
  }

  it('lets a sort that moves the formulas through, and still refuses the rest', () => {
    expect(reorderGate({ ...input, canMoveFormulas: true })).toBe('allow')
    expect(reorderGate(input)).toBe('valueModeFormulas')
    expect(reorderGate({ ...input, preloadComplete: false, canMoveFormulas: true })).toBe(
      'offerFullLoad',
    )
  })
})
