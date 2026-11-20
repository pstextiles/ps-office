/**
 * The built-in refresh run (PS Office): compute first, save only a complete
 * plan, one run at a time, and report why nothing was written.
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
import { runQueryRefreshJob } from '../src/main/query-refresh-worker'
import { queryTableWorkbook } from './power-query-fixtures'

const client = {} as ArchiveClient
const input = {
  client,
  sourcePath: 'C:\\snap\\book.xlsx',
  targetPath: 'C:\\docs\\book.xlsx',
  systemCulture: 'ko-KR',
}
const PLAN: RefreshOutcome = {
  ok: true,
  plan: {
    edits: [{ sheetName: 'Raw', row: 1, column: 0, writeValue: true, cell: { value: 'a' } }],
    tableEdits: [
      {
        sheetName: 'Raw',
        tableName: 'Tbl',
        area: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 3 },
      },
    ],
    tables: [{ query: 'Datos', sheet: 'Raw', table: 'Tbl', rowsBefore: 3, rowsAfter: 1 }],
  },
}
function deps(
  outcome: RefreshOutcome | Error,
  saveError?: unknown,
): BuiltInRefreshDeps & { save: ReturnType<typeof vi.fn> } {
  return {
    plan: vi.fn(async () => {
      if (outcome instanceof Error) throw outcome
      return outcome
    }),
    save: vi.fn(async () => {
      if (saveError) throw saveError
      return { touchedEntries: [], removedEntries: [], addedEntries: [] }
    }),
  }
}

describe('refreshWorkbookBuiltIn', () => {
  it('saves the plan from the snapshot to the file and reports the tables', async () => {
    const d = deps(PLAN)
    const result = await refreshWorkbookBuiltIn(input, d)
    expect(result).toEqual({ ok: true, tables: PLAN.ok ? PLAN.plan.tables : [] })
    expect(d.plan).toHaveBeenCalledWith({ path: input.sourcePath, systemCulture: 'ko-KR' })
    expect(d.save).toHaveBeenCalledTimes(1)
    expect(d.save.mock.calls[0]![0]).toMatchObject({
      client,
      sourcePath: input.sourcePath,
      targetPath: input.targetPath,
      edits: PLAN.ok ? PLAN.plan.edits : [],
      tableEdits: PLAN.ok ? PLAN.plan.tableEdits : [],
    })
  })

  it('writes nothing when the engine cannot be exact or the query fails', async () => {
    const limit = deps({
      ok: false,
      failure: { kind: 'limit', query: 'Datos', message: 'the table is filtered' },
    })
    expect(await refreshWorkbookBuiltIn(input, limit)).toEqual({
      ok: false,
      code: 'limit',
      query: 'Datos',
      detail: 'the table is filtered',
    })
    expect(limit.save).not.toHaveBeenCalled()
    const failing = deps({
      ok: false,
      failure: { kind: 'error', query: null, message: 'x'.repeat(900) },
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
    const locked = Object.assign(new Error('busy'), { code: 'EBUSY' })
    expect(await refreshWorkbookBuiltIn(input, deps(PLAN, locked))).toEqual({
      ok: false,
      code: 'locked',
    })
    expect(
      await refreshWorkbookBuiltIn(
        input,
        deps(PLAN, new Error('Table "Tbl" would overlap table A9:B9.')),
      ),
    ).toEqual({
      ok: false,
      code: 'failed',
      detail: 'Table "Tbl" would overlap table A9:B9.',
    })
  })

  it('runs one refresh at a time', async () => {
    let release = (): void => {}
    const slow: BuiltInRefreshDeps = {
      plan: () => new Promise<RefreshOutcome>((resolve) => (release = () => resolve(PLAN))),
      save: vi.fn(async () => ({ touchedEntries: [], removedEntries: [], addedEntries: [] })),
    }
    const first = refreshWorkbookBuiltIn(input, slow)
    expect(builtInRefreshRunning()).toBe(true)
    expect(await refreshWorkbookBuiltIn(input, deps(PLAN))).toEqual({ ok: false, code: 'busy' })
    release()
    expect((await first).ok).toBe(true)
    expect(builtInRefreshRunning()).toBe(false)
  })
})

describe('the worker job', () => {
  it('plans from a workbook file and answers a broken file as a failure', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pso-refresh-'))
    try {
      const path = join(dir, 'book.xlsx')
      writeFileSync(
        path,
        queryTableWorkbook({ query: '#table({"PO", "Qty", "Fecha"}, {{"a", 1, null}})' }),
      )
      const outcome = runQueryRefreshJob({ path })
      expect(outcome.ok && outcome.plan.tables).toEqual([
        { query: 'Datos', sheet: 'Raw', table: 'Tbl', rowsBefore: 3, rowsAfter: 1 },
      ])
      const broken = join(dir, 'broken.xlsx')
      writeFileSync(broken, 'not a workbook')
      expect(runQueryRefreshJob({ path: broken })).toMatchObject({ ok: false })
      expect(runQueryRefreshJob({ path: join(dir, 'missing.xlsx') })).toMatchObject({ ok: false })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
