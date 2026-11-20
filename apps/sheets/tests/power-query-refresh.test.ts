/**
 * The built-in query refresh (PS Office): writing a query's result into its
 * query table as Excel's refresh does — the table keeps its columns and
 * takes the new row count, new rows take the styles of the row above and the
 * calculated columns' formulas, the table's sort is applied again — and
 * refusing (EngineLimit) whatever the engine can't match. Each rule here was
 * checked against Excel's own refresh of a real workbook.
 */
import { describe, expect, it } from 'vitest'

import type { QueryHost } from '../src/main/power-query'
import { partsToSerial } from '../src/main/power-query/convert'
import { findQueryTables } from '../src/main/power-query/query-tables'
import { planWorkbookRefresh, type RefreshPlan } from '../src/main/power-query/refresh'
import { readZip } from '../src/main/power-query/zip-read'
import { DEFAULT_ROWS, queryTableWorkbook, type QueryTableWorkbook } from './power-query-fixtures'

const host: QueryHost = {
  readFile: (path) => {
    throw new Error(`no file ${path}`)
  },
  listFolder: () => [],
}
const serial = (y: number, m: number, d: number): number => partsToSerial(y, m, d)!

function refresh(options: QueryTableWorkbook): ReturnType<typeof planWorkbookRefresh> {
  return planWorkbookRefresh(queryTableWorkbook(options), { host })
}
function planned(options: QueryTableWorkbook): RefreshPlan {
  const outcome = refresh(options)
  if (!outcome.ok) throw new Error(`refused: ${outcome.failure.message}`)
  return outcome.plan
}
function refusal(options: QueryTableWorkbook): {
  kind: string
  query: string | null
  message: string
} {
  const outcome = refresh(options)
  if (outcome.ok) throw new Error('expected a refusal')
  return outcome.failure
}
/** the edit at a 0-based cell of sheet Raw, or undefined */
const at = (
  plan: RefreshPlan,
  row: number,
  column: number,
): RefreshPlan['edits'][number] | undefined =>
  plan.edits.find((e) => e.sheetName === 'Raw' && e.row === row && e.column === column)

const table = (rows: string): string => `#table({"PO", "Qty", "Fecha"}, {${rows}})`
const FIVE = table(
  '{"a", 1, #date(2026, 1, 1)}, {"b", 2, #date(2026, 1, 2)}, {"c", 3, null}, {"d", 4, #date(2026, 1, 4)}, {"e", 5, #date(2026, 1, 5)}',
)

describe('finding the query tables', () => {
  it('reads the table, its bound and calculated columns, and the hidden range name', () => {
    const zip = readZip(queryTableWorkbook({ query: FIVE }))
    const [target, ...rest] = findQueryTables((path) => zip.text(path))
    expect(rest).toEqual([])
    expect(target).toMatchObject({
      query: 'Datos',
      sheetName: 'Raw',
      worksheetPath: 'xl/worksheets/sheet2.xml',
      tableName: 'Tbl',
      area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 3 },
      localName: { name: 'ExternalData_1', sheetIndex: 1 },
      sort: [],
      unsupported: [],
    })
    expect(target!.columns).toEqual([
      { name: 'PO', field: 'PO', formula: null },
      { name: 'Qty', field: 'Qty', formula: null },
      { name: 'Fecha', field: 'Fecha', formula: null },
      { name: 'Doble', field: null, formula: 'Tbl[[#This Row],[Qty]]*2' },
    ])
  })
})

describe('writing a result into its query table', () => {
  it("grows the table: new rows take the row above's styles and the calculated formula", () => {
    const plan = planned({ query: FIVE })
    expect(plan.tables).toEqual([
      { query: 'Datos', sheet: 'Raw', table: 'Tbl', rowsBefore: 3, rowsAfter: 5 },
    ])
    expect(plan.tableEdits).toEqual([
      {
        sheetName: 'Raw',
        tableName: 'Tbl',
        area: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 3 },
        localName: { name: 'ExternalData_1', sheetIndex: 1 },
      },
    ])
    // an existing row: values only, its cells keep their own styles and formula
    expect(at(plan, 1, 0)).toEqual({
      sheetName: 'Raw',
      row: 1,
      column: 0,
      writeValue: true,
      cell: { value: 'a' },
    })
    expect(at(plan, 1, 2)!.cell).toEqual({ value: serial(2026, 1, 1) })
    expect(at(plan, 1, 3)).toBeUndefined()
    expect(at(plan, 3, 2)!.cell).toEqual({ value: null })
    // new rows: the last row's style index per column (4, 5, 6, 7 in the fixture)
    expect(at(plan, 4, 0)).toMatchObject({ cell: { value: 'd' }, styleIndex: 4 })
    expect(at(plan, 4, 1)).toMatchObject({ cell: { value: 4 }, styleIndex: 5 })
    expect(at(plan, 5, 2)).toMatchObject({ cell: { value: serial(2026, 1, 5) }, styleIndex: 6 })
    expect(at(plan, 5, 3)).toMatchObject({
      cell: { value: null, formula: 'Tbl[[#This Row],[Qty]]*2' },
      styleIndex: 7,
    })
    // 3 bound columns × 5 rows + the formula in the 2 new rows
    expect(plan.edits).toHaveLength(17)
  })

  it('shrinks the table: rows that go away are emptied, formulas included', () => {
    const plan = planned({ query: table('{"z", 9, #date(2026, 3, 1)}') })
    expect(plan.tableEdits[0]!.area).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 1,
      endColumn: 3,
    })
    expect(at(plan, 1, 0)!.cell).toEqual({ value: 'z' })
    for (const row of [2, 3]) {
      for (const column of [0, 1, 2, 3]) {
        expect(at(plan, row, column)).toEqual({
          sheetName: 'Raw',
          row,
          column,
          writeValue: true,
          cell: { value: null },
        })
      }
    }
    expect(plan.edits).toHaveLength(3 + 8)
  })

  it('keeps one empty row for an empty result, and leaves the range alone when the count is the same', () => {
    const empty = planned({ query: table('') })
    expect(empty.tables[0]).toMatchObject({ rowsBefore: 3, rowsAfter: 1 })
    expect([0, 1, 2].map((c) => at(empty, 1, c)!.cell)).toEqual([
      { value: null },
      { value: null },
      { value: null },
    ])
    const same = planned({ query: table('{"x", 1, null}, {"y", 2, null}, {"z", 3, null}') })
    expect(same.tableEdits).toEqual([])
    expect(same.edits).toHaveLength(9)
  })

  it('writes an error value as an empty cell, as Excel does', () => {
    const plan = planned({
      query:
        'Table.TransformColumnTypes(#table({"PO", "Qty", "Fecha"}, {{"a", "x", null}, {"b", "2", null}, {"c", "3", null}}), {{"Qty", type number}})',
    })
    expect(at(plan, 1, 1)!.cell).toEqual({ value: null })
    expect(at(plan, 2, 1)!.cell).toEqual({ value: 2 })
  })

  it('leaves out the range name when the workbook has none', () => {
    const plan = planned({ query: FIVE, localName: false })
    expect(plan.tableEdits[0]).not.toHaveProperty('localName')
  })
})

describe("applying the table's sort again", () => {
  const sorted = (condition: string): string =>
    `<sortState ref="A2:D4"><sortCondition ${condition}/></sortState>`
  const MIXED = table(
    '{"late", 1, #date(2026, 1, 3)}, {"first", 2, #date(2026, 1, 1)}, {"gone", 3, "cancelado"}, {"blank", 4, null}, {"tie", 5, #date(2026, 1, 1)}, {"gone2", 6, "cancelado"}',
  )

  it('sorts numbers and dates ascending, one text after them, blanks last, ties in query order', () => {
    const plan = planned({ query: MIXED, tableHead: sorted('ref="C1:C4"') })
    expect([1, 2, 3, 4, 5, 6].map((row) => at(plan, row, 0)!.cell.value)).toEqual([
      'first',
      'tie',
      'late',
      'gone',
      'gone2',
      'blank',
    ])
    expect(plan.tableEdits[0]).toMatchObject({ sortFollowsRows: true })
  })

  it('sorts descending with blanks still last', () => {
    const plan = planned({
      query: table('{"a", 1, null}, {"b", 3, null}, {"c", null, null}, {"d", 2, null}'),
      tableHead: sorted('descending="1" ref="B1:B4"'),
    })
    expect([1, 2, 3, 4].map((row) => at(plan, row, 0)!.cell.value)).toEqual(['b', 'd', 'a', 'c'])
  })

  it('refuses what it cannot order as Excel does', () => {
    const texts = table('{"a", 1, "x"}, {"b", 2, "y"}, {"c", 3, null}')
    expect(refusal({ query: texts, tableHead: sorted('ref="C1:C4"') })).toMatchObject({
      kind: 'limit',
      query: 'Datos',
    })
    // text keys in a descending sort, a sort by colour, a sort on the calculated column
    const one = table('{"a", 1, "x"}, {"b", 2, null}, {"c", 3, null}')
    expect(refusal({ query: one, tableHead: sorted('descending="1" ref="C1:C4"') }).kind).toBe(
      'limit',
    )
    expect(
      refusal({ query: one, tableHead: sorted('sortBy="cellColor" ref="C1:C4" dxfId="0"') }).kind,
    ).toBe('limit')
    expect(refusal({ query: one, tableHead: sorted('ref="D1:D4"') }).kind).toBe('limit')
    // a sort that does not reach the last row can't follow a new row count
    expect(refusal({ query: FIVE, tableHead: sorted('ref="B1:B3"') }).message).toContain('sort')
  })
})

describe('refusing instead of guessing', () => {
  const SAME = table('{"x", 1, null}, {"y", 2, null}, {"z", 3, null}')

  it('refuses tables Excel would treat differently', () => {
    expect(refusal({ query: SAME, tableAttributes: 'totalsRowCount="1"' }).message).toContain(
      'totals row',
    )
    expect(
      refusal({
        query: SAME,
        tableHead:
          '<autoFilter ref="A1:D4"><filterColumn colId="0"><filters><filter val="a"/></filters></filterColumn></autoFilter>',
      }).message,
    ).toContain('filtered')
    // a column of the table's own without a formula holds values the query does not give
    expect(
      refusal({
        query: SAME,
        tableColumns:
          '<tableColumn id="1" name="PO" queryTableFieldId="1"/><tableColumn id="2" name="Qty" queryTableFieldId="2"/>' +
          '<tableColumn id="3" name="Fecha" queryTableFieldId="3"/><tableColumn id="4" name="Nota" queryTableFieldId="4"/>',
      }).message,
    ).toContain("'Nota'")
  })

  it('refuses a result whose columns changed', () => {
    const failure = refusal({ query: '#table({"PO", "Qty", "Entrega"}, {{"a", 1, null}})' })
    expect(failure).toMatchObject({ kind: 'limit', query: 'Datos' })
    expect(failure.message).toContain('columns changed')
  })

  it('refuses to grow over cells below the table', () => {
    const below = '<row r="6"><c r="B6"><v>99</v></c></row>'
    expect(refusal({ query: FIVE, rows: DEFAULT_ROWS + below }).message).toContain(
      'below the table',
    )
    // the same cell is no obstacle when the table does not reach it
    expect(
      refresh({
        query: table('{"x", 1, null}, {"y", 2, null}, {"z", 3, null}, {"w", 4, null}'),
        rows: DEFAULT_ROWS + below,
      }).ok,
    ).toBe(true)
  })

  it('tells an engine limit from a failing query, and names the query', () => {
    expect(refusal({ query: table('{"a", Number.From("$5"), null}') })).toMatchObject({
      kind: 'limit',
      query: 'Datos',
    })
    const failing = refusal({
      query: 'Excel.Workbook(File.Contents("missing.xlsx"), null, true){0}[Data]',
    })
    expect(failing).toMatchObject({ kind: 'error', query: 'Datos' })
    expect(failing.message).toContain('missing.xlsx')
  })
})
