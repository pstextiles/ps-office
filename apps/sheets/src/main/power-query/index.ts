/**
 * The built-in Power Query engine (PS Office): evaluates a workbook's queries
 * without Excel. Phase 1 computes and compares results; writing them back
 * into the workbook's query tables comes later.
 */
import { cultureNamed } from './convert'
import { Evaluator } from './evaluator'
import { buildLibrary, type QueryHost } from './library'
import { readWorkbookMashup } from './mashup'
import { parseSection, type SectionMember } from './parser'
import { MError, type MValue } from './values'

export type { QueryHost, FolderEntry } from './library'
export { MError } from './values'
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
  /** the culture text is read and written in; default the workbook's own, else en-US */
  readonly culture?: string
}

/** Evaluates one query of the document; raises an MError when it fails. */
export function evaluateQuery(doc: QueryDocument, name: string, options: EngineOptions): MValue {
  const culture = cultureNamed(options.culture ?? doc.culture ?? 'en-US')
  const evaluator = new Evaluator(buildLibrary(options.host, culture))
  const scope = evaluator.sectionScope(doc.members)
  const query = scope.lookup(name)
  if (!query) throw new MError('Expression.Error', `The query '${name}' was not found.`)
  return query()
}
