/**
 * PS Office: a built-in query refresh updates the open workbook in place,
 * like Excel's — the regions the save changed, and when the screen may adopt
 * the saved file without reopening it.
 */
import { describe, expect, it } from 'vitest'

import { refreshRegions } from '../src/main/refresh-regions'
import {
  changedDefinedNames,
  clampToExtent,
  dropKeysInRanges,
  rangeChunks,
  refreshInPlaceRefusal,
} from '../src/renderer/refresh-in-place'
import type { WorkbookFile } from '../src/shared/desktop-api'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

describe('the regions a refresh changed', () => {
  it('folds a refreshed table into one rectangle and keeps separate bands apart', () => {
    const table = [] as { row: number; column: number }[]
    for (let row = 1; row <= 40; row++)
      for (let column = 0; column < 8; column++) table.push({ row, column })
    const regions = refreshRegions([
      { sheetName: 'Cut (raw)', cells: table },
      {
        sheetName: 'Plan',
        cells: [
          { row: 5, column: 3 },
          { row: 7, column: 9 },
          { row: 100, column: 2 },
        ],
      },
    ])
    expect(regions).toEqual([
      { sheet: 'Cut (raw)', ranges: [{ startRow: 1, endRow: 40, startColumn: 0, endColumn: 7 }] },
      {
        sheet: 'Plan',
        ranges: [
          { startRow: 5, endRow: 7, startColumn: 3, endColumn: 9 },
          { startRow: 100, endRow: 100, startColumn: 2, endColumn: 2 },
        ],
      },
    ])
  })

  it('merges the same sheet given in several pieces, and folds too many bands into one', () => {
    const cells = Array.from({ length: 300 }, (_, i) => ({ row: i * 10, column: i % 5 }))
    const [only] = refreshRegions([
      { sheetName: 'S', cells: cells.slice(0, 150) },
      { sheetName: 'S', cells: cells.slice(150) },
    ])
    expect(only!.ranges).toEqual([{ startRow: 0, endRow: 2990, startColumn: 0, endColumn: 4 }])
  })
})

const sheet = (id: string, name: string, extra: Partial<WorkbookFile['sheets'][number]> = {}) =>
  ({
    id,
    name,
    rowCount: 100,
    columnCount: 10,
    tables: [
      {
        name: 'Cut__raw',
        headerRowCount: 1,
        range: { startRow: 0, endRow: 50, startColumn: 0, endColumn: 7 },
      },
    ],
    pivotTables: [],
    ...extra,
  }) as unknown as WorkbookFile['sheets'][number]

const file = (patch: Partial<WorkbookFile> = {}): WorkbookFile =>
  ({
    path: 'C:/plan.xlsx',
    sessionId: 's1',
    sha256: 'a',
    styles: [{}, {}],
    definedNames: [
      { name: '_xlnm._FilterDatabase', formula: "'Cut (raw)'!$A$1:$H$51", sheetIndex: 0 },
    ],
    sheets: [sheet('sheet-1', 'Cut (raw)'), sheet('sheet-2', 'Plan')],
    visuals: [],
    ...patch,
  }) as unknown as WorkbookFile

const state = (patch: Record<string, unknown> = {}): LazyWorkbookState =>
  ({
    file: file(),
    formulaMode: false,
    flags: { preloadComplete: true, preloadRunning: false },
    closure: { status: 'unavailable', pinned: new Map() },
    recalc: { running: false, overlay: new Map() },
    streamAliases: new Map(),
    editJournal: { structuralOps: new Map() },
    ...patch,
  }) as unknown as LazyWorkbookState

describe('adopting the refreshed file in place', () => {
  it('is taken when only cells, table ranges and row counts changed', () => {
    const grown = file({
      sessionId: 's2',
      sheets: [
        sheet('sheet-1', 'Cut (raw)', {
          rowCount: 120,
          tables: [
            {
              name: 'Cut__raw',
              headerRowCount: 1,
              range: { startRow: 0, endRow: 70, startColumn: 0, endColumn: 7 },
            },
          ] as never,
        }),
        sheet('sheet-2', 'Plan'),
      ],
    })
    expect(refreshInPlaceRefusal(state(), grown, 0)).toBeNull()
  })

  it('reopens instead when the screen holds anything a re-read could not match', () => {
    const next = file({ sessionId: 's2' })
    expect(refreshInPlaceRefusal(state(), next, 3)).toBe('edits are pending')
    expect(refreshInPlaceRefusal(state({ formulaMode: true }), next, 0)).toMatch(/formulas/)
    expect(
      refreshInPlaceRefusal(
        state({ flags: { preloadComplete: false, preloadRunning: true } }),
        next,
        0,
      ),
    ).toMatch(/loading/)
    expect(
      refreshInPlaceRefusal(state({ closure: { status: 'active', pinned: new Map() } }), next, 0),
    ).toBe('closure mode')
    expect(
      refreshInPlaceRefusal(
        state({ recalc: { running: false, overlay: new Map([['x', new Map()]]) } }),
        next,
        0,
      ),
    ).toMatch(/recalculated/)
    expect(
      refreshInPlaceRefusal(
        state({ editJournal: { structuralOps: new Map([['sheet-1', [{}]]]) } }),
        next,
        0,
      ),
    ).toMatch(/inserted/)
    expect(refreshInPlaceRefusal(state(), file({ path: 'C:/other.xlsx' }), 0)).toBe('another file')
    expect(refreshInPlaceRefusal(state(), file({ styles: [{}, {}, {}] as never }), 0)).toMatch(
      /styles/,
    )
    expect(
      refreshInPlaceRefusal(state(), file({ sheets: [sheet('sheet-1', 'Cut (raw)')] }), 0),
    ).toMatch(/sheets/)
    expect(
      refreshInPlaceRefusal(
        state(),
        file({ sheets: [sheet('sheet-1', 'Cut (raw)', { tables: [] }), sheet('sheet-2', 'Plan')] }),
        0,
      ),
    ).toMatch(/tables/)
    expect(refreshInPlaceRefusal(state(), file({ definedNames: [] }), 0)).toMatch(/names/)
  })
})

describe('re-reading the changed regions', () => {
  it('clamps a region to the saved extent', () => {
    const range = { startRow: 40, endRow: 120, startColumn: 0, endColumn: 12 }
    expect(clampToExtent(range, 100, 10)).toEqual({
      startRow: 40,
      endRow: 99,
      startColumn: 0,
      endColumn: 9,
    })
    expect(clampToExtent({ ...range, startRow: 150 }, 100, 10)).toBeNull()
  })

  it('splits a region into reads of at most the given cells', () => {
    const chunks = rangeChunks({ startRow: 0, endRow: 24, startColumn: 0, endColumn: 9 }, 100)
    expect(chunks.map((c) => [c.startRow, c.endRow])).toEqual([
      [0, 9],
      [10, 19],
      [20, 24],
    ])
  })

  it('forgets cached cells inside the regions only', () => {
    const cache = new Map<string, unknown>([
      ['5:3', 1],
      ['6:9', 2],
      ['6:10', 3],
      ['100:2', 4],
    ])
    dropKeysInRanges(cache, [{ startRow: 5, endRow: 7, startColumn: 3, endColumn: 9 }])
    expect([...cache.keys()]).toEqual(['6:10', '100:2'])
  })

  it('finds the names the refresh moved (Excel keeps ExternalData_N on its table)', () => {
    const before = [
      { name: 'ExternalData_1', formula: "'Cut (raw)'!$A$1:$H$51", sheetIndex: 0 },
      { name: 'Total', formula: 'Plan!$B$2' },
    ]
    const after = [
      { name: 'ExternalData_1', formula: "'Cut (raw)'!$A$1:$H$71", sheetIndex: 0 },
      { name: 'Total', formula: 'Plan!$B$2' },
    ]
    expect(changedDefinedNames(before, after)).toEqual([after[0]])
  })
})
