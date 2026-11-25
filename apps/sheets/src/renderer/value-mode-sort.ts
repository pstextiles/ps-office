/**
 * PS Office: a sort in a large (value-mode) workbook, done here instead of by
 * Univer's sort command.
 *
 * Univer's reorder command runs every command interceptor, and its ref-range
 * service re-derives the range and formula of every conditional-format rule
 * with a formula for the moved rows — 155 of the 160 seconds a sort of the
 * production plan's 1,260 rows took, and not what Excel does: Excel's sort
 * leaves conditional formats where they are. Here the order is computed with
 * Excel's sort order and applied with the reorder-range mutation alone (its
 * inverse goes on the undo stack), so the journal records the moved cells and
 * formulas as for any sort (journalRangeSnapshot) and nothing else moves.
 *
 * Which rows take part follows Univer's sort command: the rows of the range
 * minus a header row, without filtered or hidden rows, and only the first row
 * of each merge inside the range.
 */
import { CellValueType, IUndoRedoService, type ICellData, type IRange } from '@univerjs/core'

import { REORDER_RANGE_MUTATION } from './app-constants'
import type { UniverRuntime, UniverWorksheet } from './univer-state'

export interface SortRule {
  readonly colIndex: number
  readonly type: 'asc' | 'desc'
}

const EXCEL_ERRORS = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#SPILL!',
  '#CALC!',
  '#GETTING_DATA',
])

type SortKey =
  | { readonly kind: 0; readonly number: number }
  | { readonly kind: 1; readonly text: string }
  | { readonly kind: 2; readonly bool: boolean }
  | { readonly kind: 3 }
  | { readonly kind: 4 }

/** a cell's value as Excel sorts it: number < text < logical < error, blank always last */
export function sortKeyOf(cell: ICellData | null | undefined): SortKey {
  if (!cell) return { kind: 4 }
  const rich = cell.p?.body?.dataStream
  const raw: unknown =
    rich !== undefined ? rich.replace(/\r\n$/, '').replace(/[\r\n]$/, '') : cell.v
  if (raw === null || raw === undefined || raw === '') return { kind: 4 }
  if (cell.t === CellValueType.BOOLEAN) {
    return { kind: 2, bool: raw === true || raw === 1 || String(raw).toUpperCase() === 'TRUE' }
  }
  if (typeof raw === 'boolean') return { kind: 2, bool: raw }
  if (typeof raw === 'number') return { kind: 0, number: raw }
  const text = String(raw)
  if (cell.t === CellValueType.NUMBER && text.trim() !== '' && Number.isFinite(Number(text))) {
    return { kind: 0, number: Number(text) }
  }
  if (EXCEL_ERRORS.has(text)) return { kind: 3 }
  return { kind: 1, text }
}

// Excel's default sort order for text: digits, the space and punctuation in
// this order, then letters (case ignored); apostrophes and hyphens are
// ignored, and two strings equal but for a hyphen put the hyphenated one last.
const SYMBOL_ORDER = ' !"#$%&()*,./:;?@[\\]^_`{|}~+<=>'
const LETTER_BASE = 100

function charRank(ch: string): number {
  if (ch >= '0' && ch <= '9') return ch.charCodeAt(0) - 48
  const symbol = SYMBOL_ORDER.indexOf(ch)
  if (symbol >= 0) return 10 + symbol
  const upper = ch.toUpperCase()
  if (upper >= 'A' && upper <= 'Z') return LETTER_BASE + upper.charCodeAt(0) - 65
  // accented Latin letters sort with their base letter
  const base = upper.normalize('NFD').replace(/[̀-ͯ]/g, '')
  if (base.length === 1 && base >= 'A' && base <= 'Z') return LETTER_BASE + base.charCodeAt(0) - 65
  // anything else after the letters, by code point (Hangul syllables are in 가나다 order)
  return 1000 + ch.codePointAt(0)!
}

/** Excel's comparison of two texts for sorting (ascending). */
export function compareSortText(a: string, b: string): number {
  const strip = (s: string): string => s.replace(/['’-]/g, '')
  const left = [...strip(a)]
  const right = [...strip(b)]
  const n = Math.min(left.length, right.length)
  for (let i = 0; i < n; i++) {
    const d = charRank(left[i]!) - charRank(right[i]!)
    if (d !== 0) return d
  }
  if (left.length !== right.length) return left.length - right.length
  // equal but for hyphens and apostrophes: the one with a hyphen sorts last
  const hyphens = (s: string): number => (s.match(/-/g) ?? []).length
  return hyphens(a) - hyphens(b)
}

/** Excel's order of two sort keys; blanks last in both directions. */
export function compareSortKeys(a: SortKey, b: SortKey, type: 'asc' | 'desc'): number {
  if (a.kind === 4 || b.kind === 4) return a.kind === b.kind ? 0 : a.kind === 4 ? 1 : -1
  let order: number
  if (a.kind !== b.kind) order = a.kind - b.kind
  else if (a.kind === 0) order = a.number - (b as { number: number }).number
  else if (a.kind === 1) order = compareSortText(a.text, (b as { text: string }).text)
  else if (a.kind === 2) order = Number(a.bool) - Number((b as { bool: boolean }).bool)
  else order = 0
  return type === 'asc' ? order : -order
}

/**
 * The reorder for a sort: target row → the row that moves there, as Univer's
 * reorder-range mutation takes it. Ties keep their order (Excel's sort is stable).
 */
export function sortOrder(
  worksheet: Pick<UniverWorksheet, 'getSheet'>,
  range: IRange,
  rules: readonly SortRule[],
  hasTitle: boolean,
): Record<number, number> {
  const sheet = worksheet.getSheet()
  const merges = sheet
    .getMergeData()
    .filter(
      (merge) =>
        merge.startRow >= range.startRow &&
        merge.endRow <= range.endRow &&
        merge.startColumn >= range.startColumn &&
        merge.endColumn <= range.endColumn,
    )
  const mergeStarts = new Set(merges.map((merge) => merge.startRow))
  const rows: { index: number; keys: SortKey[] }[] = []
  for (let row = hasTitle ? range.startRow + 1 : range.startRow; row <= range.endRow; row++) {
    if (sheet.getRowFiltered(row)) continue
    if (sheet.getRowRawVisible(row) === false) continue
    if (merges.length > 0 && !mergeStarts.has(row)) continue
    rows.push({
      index: row,
      keys: rules.map((rule) => sortKeyOf(sheet.getCellRaw(row, rule.colIndex))),
    })
  }
  const positions = rows.map((row) => row.index)
  const sorted = rows
    .map((row, position) => ({ ...row, position }))
    .sort((a, b) => {
      for (let i = 0; i < rules.length; i++) {
        const d = compareSortKeys(a.keys[i]!, b.keys[i]!, rules[i]!.type)
        if (d !== 0) return d
      }
      return a.position - b.position
    })
  const order: Record<number, number> = {}
  sorted.forEach((row, position) => {
    order[positions[position]!] = row.index
  })
  return order
}

/** The reorder that undoes `order`. */
export function inverseOrder(order: Readonly<Record<number, number>>): Record<number, number> {
  const inverse: Record<number, number> = {}
  for (const [target, source] of Object.entries(order)) inverse[source] = Number(target)
  return inverse
}

/**
 * Applies a reorder with the mutation alone and puts its inverse on the undo
 * stack. Returns false (nothing changed) when the mutation was refused.
 */
export function applyReorder(
  runtime: UniverRuntime,
  unitId: string,
  subUnitId: string,
  range: IRange,
  order: Record<number, number>,
): boolean {
  const redo = { id: REORDER_RANGE_MUTATION, params: { unitId, subUnitId, range, order } }
  const undo = {
    id: REORDER_RANGE_MUTATION,
    params: { unitId, subUnitId, range, order: inverseOrder(order) },
  }
  if (!runtime.univerAPI.syncExecuteCommand(redo.id, redo.params)) return false
  runtime.univer
    .__getInjector()
    .get(IUndoRedoService)
    .pushUndoRedo({ unitID: unitId, undoMutations: [undo], redoMutations: [redo] })
  return true
}
