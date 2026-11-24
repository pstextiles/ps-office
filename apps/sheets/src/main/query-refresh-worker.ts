/**
 * One built-in query refresh in a worker thread of its own (see
 * builtin-refresh.ts): computing the queries and recalculating the
 * workbook's formulas take seconds of CPU, which on the main thread would
 * freeze every window. It answers once and is terminated. It also plans
 * the save (PS Office): planning reads and rewrites the touched sheets, which
 * took seconds and up to ~1GB on the main thread; here the main process only
 * writes the planned parts.
 */
import { readFileSync } from 'node:fs'
import { parentPort, workerData } from 'node:worker_threads'
import {
  createBufferEntrySource,
  type SheetFormulaValues,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  planStreamingSave,
  readArchiveEntries,
  type PlannedArchiveSave,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
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
  /** the save, planned from the workbook's bytes; absent when planning failed (the main process plans it then) */
  readonly save?: PlannedArchiveSave
}

export function runQueryRefreshJob(job: QueryRefreshJob): QueryRefreshAnswer {
  return computeRefresh(job).answer
}

/** The refresh, then its save planned in this thread from the same bytes. */
export async function runQueryRefreshJobAndPlanSave(
  job: QueryRefreshJob,
): Promise<QueryRefreshAnswer> {
  const { answer, bytes } = computeRefresh(job)
  if (!answer.outcome.ok || !bytes) return answer
  try {
    const buffer = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    const plan = await planStreamingSave(await createBufferEntrySource(buffer), {
      edits: answer.outcome.plan.edits,
      tableEdits: answer.outcome.plan.tableEdits,
      ...(answer.formulas?.ok ? { formulaValues: answer.formulas.values } : {}),
    })
    return {
      ...answer,
      save: {
        replaced: plan.replaced,
        added: plan.added,
        addedBinary: plan.addedBinary,
        removedEntries: plan.removedEntries,
        addedEntries: plan.addedEntries,
        touchedEntries: plan.touchedEntries,
        sourceEntries: readArchiveEntries(bytes),
      },
    }
  } catch {
    return answer
  }
}

function computeRefresh(job: QueryRefreshJob): {
  answer: QueryRefreshAnswer
  bytes: Uint8Array | null
} {
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
      answer: {
        outcome: {
          ok: false,
          failure: {
            kind: 'error',
            query: null,
            message: error instanceof Error ? error.message : String(error),
          },
        },
        formulas: null,
      },
      bytes: null,
    }
  }
  if (!outcome.ok) return { answer: { outcome, formulas: null }, bytes }
  try {
    const recalc = recalculateAfterEdits(bytes, outcome.plan.edits, outcome.plan.tableEdits, {
      today: job.today ?? dateSerial(new Date()),
      systemCulture: job.systemCulture,
    })
    return {
      answer: {
        outcome,
        formulas: recalc.ok
          ? { ok: true, values: recalc.formulaValues, changed: recalc.changed }
          : { ok: false, reason: recalc.reason },
      },
      bytes,
    }
  } catch (error) {
    // the tables are still refreshed; the formulas keep their stored results
    return {
      answer: {
        outcome,
        formulas: { ok: false, reason: error instanceof Error ? error.message : String(error) },
      },
      bytes,
    }
  }
}

/** set in workerData by the starter, so importing this module elsewhere (a test runner's worker) listens to nothing */
export const QUERY_REFRESH_WORKER_DATA = { psOfficeQueryRefresh: true } as const

if ((workerData as { psOfficeQueryRefresh?: unknown } | null)?.psOfficeQueryRefresh === true) {
  parentPort?.once('message', (job: QueryRefreshJob) => {
    void runQueryRefreshJobAndPlanSave(job).then((answer) => parentPort?.postMessage(answer))
  })
}
