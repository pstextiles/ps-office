/**
 * Data › Queries & Connections › Refresh (PS Office, built-in): the queries
 * are computed by the built-in engine (src/main/power-query/) in a worker
 * thread, and their results are saved into the workbook's query tables
 * through the ordinary save path. Nothing is written unless every query
 * table can be refreshed exactly as Excel would; otherwise the reason is
 * returned and the file is left as it was.
 *
 * The formulas of the workbook are then recalculated by the built-in
 * calculation (src/main/formula-calc/) and their new results saved with the
 * tables — but only when every one of them can be computed exactly as Excel
 * would. Otherwise they keep their stored results (`formulas: 'stale'`) until
 * Excel opens the file; every save asks Excel to recalculate on open.
 */
import { existsSync } from 'node:fs'
import { Worker } from 'node:worker_threads'
import {
  saveWorkbookViaSidecar,
  type ArchiveClient,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import type { RefreshedTable } from './power-query/refresh'
import queryRefreshWorkerPath from './query-refresh-worker?modulePath'
import {
  QUERY_REFRESH_WORKER_DATA,
  runQueryRefreshJob,
  type QueryRefreshAnswer,
  type QueryRefreshJob,
} from './query-refresh-worker'

/** a refresh that has not finished by then is stopped (a slow PC, a huge source) */
export const BUILTIN_REFRESH_TIMEOUT_MS = 15 * 60_000

export type BuiltInRefreshResult =
  | {
      readonly ok: true
      readonly tables: readonly RefreshedTable[]
      /** recalculated: the formulas' new results were saved too; stale: they keep their previous results */
      readonly formulas: 'recalculated' | 'stale'
      /** why the formulas were left as they were */
      readonly formulasDetail?: string
    }
  | {
      readonly ok: false
      /** limit: the engine can't do this exactly (Excel can); failed: the query or the save failed */
      readonly code: 'busy' | 'limit' | 'failed' | 'locked' | 'timeout'
      readonly query?: string
      readonly detail?: string
    }

export interface BuiltInRefreshDeps {
  readonly plan: (job: QueryRefreshJob) => Promise<QueryRefreshAnswer>
  readonly save: typeof saveWorkbookViaSidecar
}

/** The worker chunk when the build bundled one; otherwise (the test runner) the calling thread. */
export function planInWorker(job: QueryRefreshJob): Promise<QueryRefreshAnswer> {
  const path: unknown = queryRefreshWorkerPath
  if (typeof path !== 'string' || path === '' || !existsSync(path)) {
    return Promise.resolve(runQueryRefreshJob(job))
  }
  return new Promise((resolve, reject) => {
    const worker = new Worker(path, {
      workerData: QUERY_REFRESH_WORKER_DATA,
      // formulas that depend on one another row after row recurse deeply
      resourceLimits: { stackSizeMb: 64 },
    })
    const timer = setTimeout(() => {
      void worker.terminate()
      reject(new RefreshTimeout())
    }, BUILTIN_REFRESH_TIMEOUT_MS)
    const finish = (): void => {
      clearTimeout(timer)
      void worker.terminate()
    }
    worker.once('message', (answer: QueryRefreshAnswer) => {
      finish()
      resolve(answer)
    })
    worker.once('error', (error) => {
      finish()
      reject(error)
    })
    worker.once('exit', (code) => {
      clearTimeout(timer)
      reject(new Error(`The query worker stopped (exit code ${code}).`))
    })
    worker.postMessage(job)
  })
}

class RefreshTimeout extends Error {}

const defaultDeps: BuiltInRefreshDeps = { plan: planInWorker, save: saveWorkbookViaSidecar }

let running = false

export function builtInRefreshRunning(): boolean {
  return running
}

/** the file is open elsewhere (Excel, a sync client) or read-only */
function isLockedError(error: unknown): boolean {
  const code = (error as { code?: unknown } | null)?.code
  return code === 'EBUSY' || code === 'EPERM' || code === 'EACCES'
}

const short = (text: string): string => (text.length > 400 ? `${text.slice(0, 399)}…` : text)

export async function refreshWorkbookBuiltIn(
  input: {
    readonly client: ArchiveClient
    /** the session's bytes: what the person sees */
    readonly sourcePath: string
    /** the workbook file to write */
    readonly targetPath: string
    readonly systemCulture?: string
  },
  deps: BuiltInRefreshDeps = defaultDeps,
): Promise<BuiltInRefreshResult> {
  if (running) return { ok: false, code: 'busy' }
  running = true
  try {
    let answer: QueryRefreshAnswer
    try {
      answer = await deps.plan({
        path: input.sourcePath,
        ...(input.systemCulture === undefined ? {} : { systemCulture: input.systemCulture }),
      })
    } catch (error) {
      if (error instanceof RefreshTimeout) return { ok: false, code: 'timeout' }
      return {
        ok: false,
        code: 'failed',
        detail: short(error instanceof Error ? error.message : String(error)),
      }
    }
    const { outcome, formulas } = answer
    if (!outcome.ok) {
      const { kind, query, message } = outcome.failure
      return {
        ok: false,
        code: kind === 'limit' ? 'limit' : 'failed',
        ...(query === null ? {} : { query }),
        detail: short(message),
      }
    }
    try {
      await deps.save({
        client: input.client,
        sourcePath: input.sourcePath,
        targetPath: input.targetPath,
        edits: outcome.plan.edits,
        tableEdits: outcome.plan.tableEdits,
        ...(formulas?.ok ? { formulaValues: formulas.values } : {}),
      })
    } catch (error) {
      if (isLockedError(error)) return { ok: false, code: 'locked' }
      return {
        ok: false,
        code: 'failed',
        detail: short(error instanceof Error ? error.message : String(error)),
      }
    }
    if (formulas?.ok) return { ok: true, tables: outcome.plan.tables, formulas: 'recalculated' }
    return {
      ok: true,
      tables: outcome.plan.tables,
      formulas: 'stale',
      ...(formulas ? { formulasDetail: short(formulas.reason) } : {}),
    }
  } finally {
    running = false
  }
}
