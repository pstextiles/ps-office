/**
 * PS Office: one workbook file's queries refreshed and formulas recalculated
 * by the built-in engine, saved into the file itself — the refresh-only run
 * of the app (`--refresh-queries`) and the standalone PS Query Refresh
 * program (tools/query-refresh/) share it. No Electron here.
 *
 * The engine reads a private copy of the file, so the save checks the copy
 * it planned from (and a file being written meanwhile is not read half way).
 * When every formula was recalculated the file keeps its own
 * recalculate-on-open setting (calc-on-load.ts): an older Excel turns new
 * functions into #NAME? on a forced recalculation.
 */
import { existsSync, statSync } from 'node:fs'
import { copyFile, mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { extname, join } from 'node:path'
import type { ArchiveClient } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import {
  refreshWorkbookBuiltIn,
  type BuiltInRefreshDeps,
  type BuiltInRefreshResult,
} from './builtin-refresh'

export type FileRefreshResult =
  | BuiltInRefreshResult
  | { readonly ok: false; readonly code: 'not-file' }

/** only an existing .xlsx/.xlsm file can be refreshed in place */
export function refreshableFile(path: string): boolean {
  return /\.xls[xm]$/i.test(path) && existsSync(path) && statSync(path).isFile()
}

export async function refreshFileInPlace(
  path: string,
  options: {
    readonly client: ArchiveClient
    readonly systemCulture?: string
    readonly deps?: BuiltInRefreshDeps
  },
): Promise<FileRefreshResult> {
  if (!refreshableFile(path)) return { ok: false, code: 'not-file' }
  const dir = await mkdtemp(join(tmpdir(), 'psoffice-refresh-'))
  try {
    const snapshot = join(dir, `source${extname(path).toLowerCase()}`)
    await copyFile(path, snapshot)
    return await refreshWorkbookBuiltIn(
      {
        client: options.client,
        sourcePath: snapshot,
        targetPath: path,
        ...(options.systemCulture === undefined ? {} : { systemCulture: options.systemCulture }),
        trustRecalculatedValues: true,
      },
      ...(options.deps ? [options.deps] : []),
    )
  } finally {
    await rm(dir, { recursive: true, force: true }).catch(() => undefined)
  }
}
