/**
 * The built-in Power Query engine (PS Office): evaluates a workbook's queries
 * without Excel. Phase 1 computes and compares results; writing them back
 * into the workbook's query tables comes later.
 */
import { cultureNamed, type Culture } from './convert'
import { Evaluator } from './evaluator'
import { buildLibrary, type QueryHost } from './library'
import { readWorkbookMashup } from './mashup'
import { parseSection, type SectionMember } from './parser'
import { EngineLimit, MError, type MValue } from './values'

export type { QueryHost, FolderEntry } from './library'
export { EngineLimit, MError } from './values'
export { readWorkbookMashup, workbookSectionText } from './mashup'

export interface QueryDocument {
  readonly members: readonly SectionMember[]
  /** query names in the order the document defines them */
  readonly names: readonly string[]
  /** the workbook's Power Query culture, when it stores one */
  readonly culture: string | null
}

export function parseQueries(sectionText: string, culture: string | null = null): QueryDocument {
  const members = parseSection(sectionText)
  return { members, names: members.map((m) => m.name), culture }
}

/** The queries of a workbook, or null when it has none. */
export function workbookQueries(workbookBytes: Uint8Array): QueryDocument | null {
  const mashup = readWorkbookMashup(workbookBytes)
  return mashup === null ? null : parseQueries(mashup.section, mashup.culture)
}

export interface EngineOptions {
  readonly host: QueryHost
  /**
   * The culture text is read and written in, overriding the workbook's.
   * Leave it out to do as Excel does: the workbook's own locale (Query
   * Options › Current Workbook › Regional Settings), else the PC's.
   */
  readonly culture?: string
  /** the PC's regional format (e.g. app.getSystemLocale()), used when the workbook stores no locale */
  readonly systemCulture?: string
}

/**
 * The culture Excel would use: the workbook's stored locale, else the PC's
 * regional format. Checked in Excel: a workbook saved with es-GT reads text
 * as es-GT on a ko-KR PC; one with no locale reads it as the PC's ko-KR.
 */
export function queryCulture(doc: QueryDocument, options: Omit<EngineOptions, 'host'>): Culture {
  const name = options.culture ?? doc.culture ?? options.systemCulture
  if (!name) throw new EngineLimit('the workbook stores no locale and the PC locale is unknown')
  return cultureNamed(name)
}

/**
 * Evaluates several queries in one section scope, so a query that others
 * reference is computed once. Raises like evaluateQuery; `onQuery` names the
 * query about to be computed (for reporting which one failed).
 */
export function evaluateQueries(
  doc: QueryDocument,
  names: readonly string[],
  options: EngineOptions & { readonly onQuery?: (name: string) => void },
): Map<string, MValue> {
  const culture = queryCulture(doc, options)
  const evaluator = new Evaluator(buildLibrary(options.host, culture))
  const scope = evaluator.sectionScope(doc.members)
  const out = new Map<string, MValue>()
  for (const name of names) {
    options.onQuery?.(name)
    const query = scope.lookup(name)
    if (!query) throw new MError('Expression.Error', `The query '${name}' was not found.`)
    out.set(name, query())
  }
  return out
}

/** Evaluates one query of the document; raises an MError when it fails, EngineLimit when it can't be exact. */
export function evaluateQuery(doc: QueryDocument, name: string, options: EngineOptions): MValue {
  const culture = queryCulture(doc, options)
  const evaluator = new Evaluator(buildLibrary(options.host, culture))
  const scope = evaluator.sectionScope(doc.members)
  const query = scope.lookup(name)
  if (!query) throw new MError('Expression.Error', `The query '${name}' was not found.`)
  return query()
}
