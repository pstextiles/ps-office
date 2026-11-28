/**
 * One workbook refreshed in place (refresh-file.ts), shared by the app's
 * refresh-only run and the standalone PS Query Refresh program: the engine
 * reads a private copy, the file itself is the save's target, and the
 * workbook keeps its own recalculate-on-open setting when it can.
 */
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

import type { ArchiveClient } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import { refreshableFile, refreshFileInPlace } from '../src/main/refresh-file'

const client = {} as ArchiveClient

describe('refreshFileInPlace', () => {
  it('refuses what is not an existing .xlsx/.xlsm file', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pso-refresh-file-'))
    try {
      writeFileSync(join(dir, 'data.csv'), 'a,b\n')
      expect(refreshableFile(join(dir, 'data.csv'))).toBe(false)
      expect(refreshableFile(join(dir, 'missing.xlsx'))).toBe(false)
      expect(refreshableFile(dir)).toBe(false)
      expect(await refreshFileInPlace(join(dir, 'missing.xlsx'), { client })).toEqual({
        ok: false,
        code: 'not-file',
      })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('has the engine read a private copy and writes the file itself, trusting a full recalculation', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pso-refresh-file-'))
    try {
      const path = join(dir, 'Plan.XLSX')
      writeFileSync(path, 'workbook bytes')
      let planned = ''
      const plan = vi.fn(async (job: { path: string; trustRecalculatedValues?: boolean }) => {
        planned = job.path
        expect(existsSync(job.path)).toBe(true)
        return {
          outcome: { ok: false, failure: { kind: 'error', query: null, message: 'stop here' } },
          formulas: null,
        } as never
      })
      const result = await refreshFileInPlace(path, {
        client,
        systemCulture: 'es-GT',
        deps: { plan, save: vi.fn(), write: vi.fn() },
      })
      expect(result).toEqual({ ok: false, code: 'failed', detail: 'stop here' })
      expect(plan).toHaveBeenCalledWith({
        path: planned,
        systemCulture: 'es-GT',
        trustRecalculatedValues: true,
      })
      expect(planned).not.toBe(path)
      expect(planned.endsWith('.xlsx')).toBe(true)
      // the private copy is gone afterwards
      expect(existsSync(planned)).toBe(false)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
