/**
 * One built-in query refresh in a worker thread of its own (see
 * builtin-refresh.ts): computing the queries and recalculating the
 * workbook's formulas take seconds of CPU, which on the main thread would
 * freeze every window. It answers once and is terminated.
 */
import { readFileSync } from 'node:fs'
import { parentPort, workerData } from 'node:worker_threads'
import type { SheetFormulaValues } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { dateSerial, recalculateAfterEdits } from './formula-calc/recalc'
import { nodeQueryHost } from './power-query/node-host'
import { planWorkbookRefresh, type RefreshOutcome } from './power-query/refresh'

export interface QueryRefreshJob {
  readonly path: string
  readonly systemCulture?: string
  /** today as a date serial, for TODAY; default the PC's date */
  readonly today?: number
}

/**
 * The formulas' new results to save with the refreshed tables, or why they
 * are left as stored (a formula the built-in calculation can't compute
 * exactly as Excel does); null when the refresh itself was refused.
 */
export type RefreshedFormulas =
  | { readonly ok: true; readonly values: SheetFormulaValues[]; readonly changed: number }
  | { readonly ok: false; readonly reason: string }

export interface QueryRefreshAnswer {
  readonly outcome: RefreshOutcome
  readonly formulas: RefreshedFormulas | null
}

export function runQueryRefreshJob(job: QueryRefreshJob): QueryRefreshAnswer {
  let bytes: Uint8Array
  let outcome: RefreshOutcome
  try {
    bytes = readFileSync(job.path)
    outcome = planWorkbookRefresh(bytes, {
      host: nodeQueryHost,
      ...(job.systemCulture === undefined ? {} : { systemCulture: job.systemCulture }),
    })
  } catch (error) {
    return {
      outcome: {
        ok: false,
        failure: {
          kind: 'error',
          query: null,
          message: error instanceof Error ? error.message : String(error),
        },
      },
      formulas: null,
    }
  }
  if (!outcome.ok) return { outcome, formulas: null }
  try {
    const recalc = recalculateAfterEdits(bytes, outcome.plan.edits, outcome.plan.tableEdits, {
      today: job.today ?? dateSerial(new Date()),
      systemCulture: job.systemCulture,
    })
    return {
      outcome,
      formulas: recalc.ok
        ? { ok: true, values: recalc.formulaValues, changed: recalc.changed }
        : { ok: false, reason: recalc.reason },
    }
  } catch (error) {
    // the tables are still refreshed; the formulas keep their stored results
    return {
      outcome,
      formulas: { ok: false, reason: error instanceof Error ? error.message : String(error) },
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
