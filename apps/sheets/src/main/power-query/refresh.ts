/**
 * A whole built-in refresh: every query table of the workbook gets its
 * query's new result. All queries are computed first; if any one can't be
 * computed exactly, nothing is written (as with the refresh in Excel, where
 * one failure leaves the file unsaved).
 */
import type { CellEdit, SheetTableEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { evaluateQueries, workbookQueries } from './index'
import type { QueryHost } from './library'
import { findQueryTables, planQueryTableWrite } from './query-tables'
import { EngineLimit, MError, isTable } from './values'
import { readZip } from './zip-read'

export interface RefreshedTable {
  readonly query: string
  readonly sheet: string
  readonly table: string
  readonly rowsBefore: number
  readonly rowsAfter: number
}

export interface RefreshPlan {
  readonly edits: CellEdit[]
  readonly tableEdits: SheetTableEdit[]
  readonly tables: readonly RefreshedTable[]
}

export type RefreshFailure =
  /** the engine can't compute or write this exactly as Excel does; Excel can */
  | { readonly kind: 'limit'; readonly query: string | null; readonly message: string }
  /** the query itself fails (a missing source file, an error in the M code) */
  | { readonly kind: 'error'; readonly query: string | null; readonly message: string }

export type RefreshOutcome =
  | { readonly ok: true; readonly plan: RefreshPlan }
  | { readonly ok: false; readonly failure: RefreshFailure }

export interface RefreshOptions {
  readonly host: QueryHost
  /** the PC's regional format, for a workbook that stores no locale */
  readonly systemCulture?: string
}

export function planWorkbookRefresh(
  workbookBytes: Uint8Array,
  options: RefreshOptions,
): RefreshOutcome {
  let current: string | null = null
  try {
    const doc = workbookQueries(workbookBytes)
    if (!doc) throw new EngineLimit('the workbook has no queries')
    const zip = readZip(workbookBytes)
    const read = (path: string): string | null => zip.text(path)
    const targets = findQueryTables(read)
    if (targets.length === 0) throw new EngineLimit('no query loads into a table of this workbook')
    for (const t of targets) {
      if (!doc.names.includes(t.query))
        throw new EngineLimit(`the query '${t.query}' is not in the workbook's queries`)
    }
    const names = [...new Set(targets.map((t) => t.query))]
    const results = evaluateQueries(doc, names, {
      host: options.host,
      ...(options.systemCulture === undefined ? {} : { systemCulture: options.systemCulture }),
      onQuery: (name) => {
        current = name
      },
    })
    const edits: CellEdit[] = []
    const tableEdits: SheetTableEdit[] = []
    const tables: RefreshedTable[] = []
    for (const target of targets) {
      current = target.query
      const result = results.get(target.query)!
      if (!isTable(result))
        throw new EngineLimit(
          `the query '${target.query}' gives a ${result === null ? 'null' : typeof result}, not a table`,
        )
      const worksheetXml = read(target.worksheetPath)
      if (worksheetXml === null)
        throw new EngineLimit(`the sheet '${target.sheetName}' could not be read`)
      const write = planQueryTableWrite(target, result, worksheetXml)
      for (const edit of write.edits) edits.push(edit)
      if (write.tableEdit) tableEdits.push(write.tableEdit)
      tables.push({
        query: target.query,
        sheet: target.sheetName,
        table: target.tableName,
        rowsBefore: write.rowsBefore,
        rowsAfter: write.rowsAfter,
      })
    }
    return { ok: true, plan: { edits, tableEdits, tables } }
  } catch (error) {
    if (error instanceof EngineLimit)
      return { ok: false, failure: { kind: 'limit', query: current, message: error.message } }
    if (error instanceof MError) {
      return {
        ok: false,
        failure: { kind: 'error', query: current, message: `${error.reason}: ${error.message}` },
      }
    }
    throw error
  }
}
