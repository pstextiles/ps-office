/**
 * The built-in refresh run (PS Office): compute first, save only a complete
 * plan, one run at a time, report why nothing was written — and save the
 * formulas' new results with the tables only when all of them were computed.
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

import type { ArchiveClient } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import {
  builtInRefreshRunning,
  refreshWorkbookBuiltIn,
  type BuiltInRefreshDeps,
} from '../src/main/builtin-refresh'
import type { RefreshOutcome } from '../src/main/power-query/refresh'
import {
  runQueryRefreshJob,
  type QueryRefreshAnswer,
  type RefreshedFormulas,
} from '../src/main/query-refresh-worker'
import { queryTableWorkbook } from './power-query-fixtures'

const client = {} as ArchiveClient
const input = {
  client,
  sourcePath: 'C:\\snap\\book.xlsx',
  targetPath: 'C:\\docs\\book.xlsx',
  systemCulture: 'ko-KR',
}
const PLAN = {
  edits: [{ sheetName: 'Raw', row: 1, column: 0, writeValue: true, cell: { value: 'a' } }],
  tableEdits: [
    {
      sheetName: 'Raw',
      tableName: 'Tbl',
      area: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 3 },
    },
  ],
  tables: [{ query: 'Datos', sheet: 'Raw', table: 'Tbl', rowsBefore: 3, rowsAfter: 1 }],
}
const PLANNED: RefreshOutcome = { ok: true, plan: PLAN }
const VALUES = [{ sheetName: 'Plan', cells: [{ row: 0, column: 0, value: 85 }] }]
const RECALCULATED: RefreshedFormulas = { ok: true, values: VALUES, changed: 1 }

function deps(
  answer: QueryRefreshAnswer | Error,
  saveError?: unknown,
): BuiltInRefreshDeps & { save: ReturnType<typeof vi.fn> } {
  return {
    plan: vi.fn(async () => {
      if (answer instanceof Error) throw answer
      return answer
    }),
    save: vi.fn(async () => {
      if (saveError) throw saveError
      return { touchedEntries: [], removedEntries: [], addedEntries: [] }
    }),
  }
}

describe('refreshWorkbookBuiltIn', () => {
  it('saves the tables and the recalculated formulas in one save', async () => {
    const d = deps({ outcome: PLANNED, formulas: RECALCULATED })
    expect(await refreshWorkbookBuiltIn(input, d)).toEqual({
      ok: true,
      tables: PLAN.tables,
      // the screen re-reads these: the table's cell and the recalculated formula
      regions: [
        { sheet: 'Raw', ranges: [{ startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }] },
        { sheet: 'Plan', ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }] },
      ],
      formulas: 'recalculated',
    })
    expect(d.plan).toHaveBeenCalledWith({ path: input.sourcePath, systemCulture: 'ko-KR' })
    expect(d.save).toHaveBeenCalledTimes(1)
    expect(d.save.mock.calls[0]![0]).toEqual({
      client,
      sourcePath: input.sourcePath,
      targetPath: input.targetPath,
      edits: PLAN.edits,
      tableEdits: PLAN.tableEdits,
      formulaValues: VALUES,
    })
  })

  it('still refreshes the tables when a formula is beyond the calculation, and says the formulas are as before', async () => {
    const d = deps({
      outcome: PLANNED,
      formulas: { ok: false, reason: 'Formula limit: the function OFFSET' },
    })
    expect(await refreshWorkbookBuiltIn(input, d)).toEqual({
      ok: true,
      tables: PLAN.tables,
      // formulas kept their results: only the table's cells changed
      regions: [
        { sheet: 'Raw', ranges: [{ startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }] },
      ],
      formulas: 'stale',
      formulasDetail: 'Formula limit: the function OFFSET',
    })
    expect(d.save.mock.calls[0]![0]).not.toHaveProperty('formulaValues')
  })

  it('writes nothing when the engine cannot be exact or the query fails', async () => {
    const limit = deps({
      outcome: {
        ok: false,
        failure: { kind: 'limit', query: 'Datos', message: 'the table is filtered' },
      },
      formulas: null,
    })
    expect(await refreshWorkbookBuiltIn(input, limit)).toEqual({
      ok: false,
      code: 'limit',
      query: 'Datos',
      detail: 'the table is filtered',
    })
    expect(limit.save).not.toHaveBeenCalled()
    const failing = deps({
      outcome: { ok: false, failure: { kind: 'error', query: null, message: 'x'.repeat(900) } },
      formulas: null,
    })
    const failed = await refreshWorkbookBuiltIn(input, failing)
    expect(failed).toMatchObject({ ok: false, code: 'failed' })
    expect(failed.ok === false && failed.detail!.length).toBe(400)
    expect(failed).not.toHaveProperty('query')
    expect(failing.save).not.toHaveBeenCalled()
    expect(await refreshWorkbookBuiltIn(input, deps(new Error('worker died')))).toEqual({
      ok: false,
      code: 'failed',
      detail: 'worker died',
    })
  })

  it('tells a locked file from other save failures', async () => {
    const answer = { outcome: PLANNED, formulas: RECALCULATED }
    const locked = Object.assign(new Error('busy'), { code: 'EBUSY' })
    expect(await refreshWorkbookBuiltIn(input, deps(answer, locked))).toEqual({
      ok: false,
      code: 'locked',
    })
    expect(
      await refreshWorkbookBuiltIn(
        input,
        deps(answer, new Error('Table "Tbl" would overlap table A9:B9.')),
      ),
    ).toEqual({ ok: false, code: 'failed', detail: 'Table "Tbl" would overlap table A9:B9.' })
  })

  it('runs one refresh at a time', async () => {
    let release = (): void => {}
    const slow: BuiltInRefreshDeps = {
      plan: () =>
        new Promise<QueryRefreshAnswer>(
          (resolve) => (release = () => resolve({ outcome: PLANNED, formulas: RECALCULATED })),
        ),
      save: vi.fn(async () => ({ touchedEntries: [], removedEntries: [], addedEntries: [] })),
    }
    const first = refreshWorkbookBuiltIn(input, slow)
    expect(builtInRefreshRunning()).toBe(true)
    expect(
      await refreshWorkbookBuiltIn(input, deps({ outcome: PLANNED, formulas: RECALCULATED })),
    ).toEqual({
      ok: false,
      code: 'busy',
    })
    release()
    expect((await first).ok).toBe(true)
    expect(builtInRefreshRunning()).toBe(false)
  })
})

describe('the worker job', () => {
  it('plans the refresh from a workbook file and recalculates its formulas', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pso-refresh-'))
    try {
      const path = join(dir, 'book.xlsx')
      writeFileSync(
        path,
        queryTableWorkbook({
          query:
            '#table({"PO", "Qty", "Fecha"}, {{"a", 21, null}, {"b", 2, null}, {"c", 3, null}})',
        }),
      )
      const answer = runQueryRefreshJob({ path, systemCulture: 'ko-KR', today: 46346 })
      expect(answer.outcome.ok && answer.outcome.plan.tables).toEqual([
        { query: 'Datos', sheet: 'Raw', table: 'Tbl', rowsBefore: 3, rowsAfter: 3 },
      ])
      // the calculated column Doble (Qty × 2) follows the new Qty of the first row; the others are as stored
      expect(answer.formulas).toEqual({
        ok: true,
        values: [{ sheetName: 'Raw', cells: [{ row: 1, column: 3, value: 42 }] }],
        changed: 1,
      })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('answers a broken or missing file as a failure', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pso-refresh-'))
    try {
      const broken = join(dir, 'broken.xlsx')
      writeFileSync(broken, 'not a workbook')
      expect(runQueryRefreshJob({ path: broken })).toMatchObject({
        outcome: { ok: false },
        formulas: null,
      })
      expect(runQueryRefreshJob({ path: join(dir, 'missing.xlsx') })).toMatchObject({
        outcome: { ok: false },
        formulas: null,
      })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
