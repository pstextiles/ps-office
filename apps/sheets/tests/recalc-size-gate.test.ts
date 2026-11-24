import { describe, expect, it } from 'vitest'

import {
  RECALC_MAX_FILE_BYTES,
  RECALC_MAX_GRID_CELLS,
  RECALC_MAX_VALUE_CELLS,
  recalcOverBudgetAtOpen,
} from '../src/renderer/univer-sync'

describe('recalcOverBudgetAtOpen', () => {
  it('disables the IronCalc fallback past the byte cap', () => {
    expect(recalcOverBudgetAtOpen(RECALC_MAX_FILE_BYTES + 1, 1000)).toBe(true)
    // the 133MB formula-heavy upload class
    expect(recalcOverBudgetAtOpen(133 * 1024 * 1024, 1000)).toBe(true)
  })

  it('disables it past the grid-cell cap even for small files', () => {
    // the 31MB / 88k×99 pure-data supplier class that burned ~4GB at open
    expect(recalcOverBudgetAtOpen(31 * 1024 * 1024, 88_000 * 99)).toBe(true)
    expect(recalcOverBudgetAtOpen(undefined, RECALC_MAX_GRID_CELLS + 1)).toBe(true)
  })

  it('disables it past the value-cell cap (PS Office: a formula-heavy plan never finished)', () => {
    expect(recalcOverBudgetAtOpen(4 * 1024 * 1024, 1_200_000, 419_356)).toBe(true)
    expect(recalcOverBudgetAtOpen(undefined, 300_000, RECALC_MAX_VALUE_CELLS + 1)).toBe(true)
    expect(recalcOverBudgetAtOpen(undefined, 300_000, RECALC_MAX_VALUE_CELLS)).toBe(false)
  })

  it('keeps the fallback for normal streamed files and sessions without a size', () => {
    expect(recalcOverBudgetAtOpen(RECALC_MAX_FILE_BYTES, RECALC_MAX_GRID_CELLS)).toBe(false)
    expect(recalcOverBudgetAtOpen(undefined, 200_000)).toBe(false)
  })
})
