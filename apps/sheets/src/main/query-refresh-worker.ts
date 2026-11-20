/**
 * One built-in query refresh plan in a worker thread of its own (see
 * builtin-refresh.ts): the queries take seconds of CPU, which on the main
 * thread would freeze every window. It answers once and is terminated.
 */
import { readFileSync } from 'node:fs'
import { parentPort, workerData } from 'node:worker_threads'
import { nodeQueryHost } from './power-query/node-host'
import { planWorkbookRefresh, type RefreshOutcome } from './power-query/refresh'

export interface QueryRefreshJob {
  readonly path: string
  readonly systemCulture?: string
}

export function runQueryRefreshJob(job: QueryRefreshJob): RefreshOutcome {
  try {
    return planWorkbookRefresh(readFileSync(job.path), {
      host: nodeQueryHost,
      ...(job.systemCulture === undefined ? {} : { systemCulture: job.systemCulture }),
    })
  } catch (error) {
    return {
      ok: false,
      failure: {
        kind: 'error',
        query: null,
        message: error instanceof Error ? error.message : String(error),
      },
    }
  }
}

/** set in workerData by the starter, so importing this module elsewhere (a test runner's worker) listens to nothing */
export const QUERY_REFRESH_WORKER_DATA = { psOfficeQueryRefresh: true } as const

if ((workerData as { psOfficeQueryRefresh?: unknown } | null)?.psOfficeQueryRefresh === true) {
  parentPort?.once('message', (job: QueryRefreshJob) => {
    parentPort?.postMessage(runQueryRefreshJob(job))
  })
}
