/**
 * Recalculating after a save's edits (PS Office, built-in query refresh):
 * which cached values go back into the file, and that nothing does when a
 * formula is beyond the built-in calculation.
 */
import { describe, expect, it } from 'vitest'

import type {
  CellEdit,
  SheetFormulaValues,
  SheetTableEdit,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { dateSerial, recalculateAfterEdits } from '../src/main/formula-calc/recalc'
import { workbook, type SheetSpec } from './formula-calc-fixtures'

const RAW: SheetSpec = {
  name: 'Raw',
  table: { name: 'Tbl', ref: 'A1:B4', columns: ['PO', 'Qty'] },
  cells: { A1: 'PO', B1: 'Qty', A2: 'a', B2: 10, A3: 'b', B3: 20, A4: 'c', B4: 30 },
}
const PLAN: SheetSpec = {
  name: 'Plan',
  cells: {
    A1: { f: 'SUM(Tbl[Qty])', stored: 60 },
    A2: { f: 'XLOOKUP("b",Raw!A:A,Raw!B:B,"")', stored: 20 },
    A3: { f: 'COUNTA(Raw!A2:A100)', stored: 3 },
    // a dynamic array that listed three rows when the file was saved
    C1: { f: 'FILTER(Raw!A2:A100,Raw!B2:B100>15,"")', array: 'C1:C2', dynamic: true, stored: 'b' },
    C2: { spilled: 'c' },
    E1: { f: 'TEXTJOIN(",",TRUE,C1#)', stored: 'b,c' },
  },
}
const TODAY = 46346
const edit = (
  sheetName: string,
  row: number,
  column: number,
  value: string | number | null,
): CellEdit => ({
  sheetName,
  row,
  column,
  writeValue: true,
  cell: { value },
})
const cellsOf = (values: readonly SheetFormulaValues[], sheet: string): Record<string, unknown> =>
  Object.fromEntries(
    (values.find((v) => v.sheetName === sheet)?.cells ?? []).map((c) => [
      `${String.fromCharCode(65 + c.column)}${c.row + 1}`,
      c.spill
        ? { spill: c.value }
        : c.arrayRef
          ? { value: c.value, arrayRef: c.arrayRef }
          : c.value,
    ]),
  )

function recalc(
  sheets: readonly SheetSpec[],
  edits: readonly CellEdit[],
  tableEdits: readonly SheetTableEdit[] = [],
  culture = 'ko-KR',
) {
  return recalculateAfterEdits(workbook(sheets), edits, tableEdits, {
    today: TODAY,
    systemCulture: culture,
  })
}

describe('recalculating after edits', () => {
  it('returns nothing to write when nothing changed', () => {
    expect(recalc([RAW, PLAN], [])).toEqual({ ok: true, formulaValues: [], changed: 0 })
  })

  it('returns only the formula cells whose result changed', () => {
    // the refresh rewrites row 3 and adds a row; the table grows with it
    const outcome = recalc(
      [RAW, PLAN],
      [edit('Raw', 2, 1, 5), edit('Raw', 4, 0, 'd'), edit('Raw', 4, 1, 40)],
      [
        {
          sheetName: 'Raw',
          tableName: 'Tbl',
          area: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 1 },
        },
      ],
    )
    if (!outcome.ok) throw new Error(outcome.reason)
    expect(cellsOf(outcome.formulaValues, 'Plan')).toEqual({
      A1: 85,
      A2: 5,
      A3: 4,
      // the array now lists c and d: same size, so no new range; its first value and one spilled cell changed
      C1: 'c',
      C2: { spill: 'd' },
      E1: 'c,d',
    })
    expect(outcome.changed).toBe(6)
    expect(outcome.formulaValues.map((v) => v.sheetName)).toEqual(['Plan'])
  })

  it('gives a dynamic array its new range and empties the cells it left', () => {
    const shrunk = recalc([RAW, PLAN], [edit('Raw', 3, 1, 1)])
    if (!shrunk.ok) throw new Error(shrunk.reason)
    expect(cellsOf(shrunk.formulaValues, 'Plan')).toMatchObject({
      C1: { value: 'b', arrayRef: 'C1' },
      C2: { spill: null },
      E1: 'b',
    })
    const grown = recalc([RAW, PLAN], [edit('Raw', 4, 0, 'd'), edit('Raw', 4, 1, 40)])
    if (!grown.ok) throw new Error(grown.reason)
    expect(cellsOf(grown.formulaValues, 'Plan')).toMatchObject({
      C1: { value: 'b', arrayRef: 'C1:C3' },
      C3: { spill: 'd' },
      E1: 'b,c,d',
    })
  })

  it('computes the formulas the edits add, such as a calculated column in new rows', () => {
    const outcome = recalc(
      [RAW, { name: 'Plan', cells: { A1: { f: 'SUM(Raw!C:C)', stored: 0 } } }],
      [
        {
          sheetName: 'Raw',
          row: 1,
          column: 2,
          writeValue: true,
          cell: { value: null, formula: 'Tbl[[#This Row],[Qty]]*2' },
        },
      ],
      [
        {
          sheetName: 'Raw',
          tableName: 'Tbl',
          area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
        },
      ],
    )
    if (!outcome.ok) throw new Error(outcome.reason)
    expect(cellsOf(outcome.formulaValues, 'Raw')).toEqual({ C2: 20 })
    expect(cellsOf(outcome.formulaValues, 'Plan')).toEqual({ A1: 20 })
  })

  it('returns no value at all when one formula is beyond the calculation', () => {
    const outcome = recalc(
      [RAW, { name: 'Plan', cells: { ...PLAN.cells, G1: { f: 'OFFSET(A1,1,0)', stored: 20 } } }],
      [edit('Raw', 2, 1, 5)],
    )
    expect(outcome.ok).toBe(false)
    expect(!outcome.ok && outcome.reason).toContain('OFFSET')
  })

  it('does not guess how a text date reads where the locale was not checked', () => {
    const sheets = [RAW, { name: 'Plan', cells: { A1: { f: '"8/12"-7', stored: 46239 } } }]
    expect(recalc(sheets, [], [], 'ko-KR')).toEqual({ ok: true, formulaValues: [], changed: 0 })
    expect(recalc(sheets, [], [], 'es-GT').ok).toBe(false)
  })

  it('turns a local date into its serial', () => {
    expect(dateSerial(new Date(2026, 10, 20, 23, 59))).toBe(TODAY)
    expect(dateSerial(new Date(2026, 10, 20, 0, 0))).toBe(TODAY)
  })
})
