/**
 * Recalculating a workbook after edits that are about to be saved (the
 * built-in query refresh): the edits are laid over the workbook as read, every
 * formula is computed, and the cells whose result changed come back as cached
 * values for the same save.
 *
 * All or nothing: if any formula can't be computed exactly as Excel would
 * (FormulaLimit), no value is returned, the file keeps its stored results and
 * the caller says so. Excel recalculates on open either way.
 */
import type {
  CellEdit,
  FormulaCachedValue,
  SheetFormulaValues,
  SheetTableEdit,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { Calculator, address, isArrayValue, type ArrayValue } from './calc'
import {
  XlError,
  cellKey,
  keyColumn,
  keyRow,
  readWorkbookModel,
  type Scalar,
  type SheetModel,
  type WorkbookModel,
} from './model'
import { FormulaLimit } from './parser'

export interface RecalcOptions {
  /** today's date as a serial (TODAY) */
  readonly today: number
  /** the PC's regional format: decides how text such as "8/12" is read as a date */
  readonly systemCulture?: string | undefined
}

export type RecalcOutcome =
  | { readonly ok: true; readonly formulaValues: SheetFormulaValues[]; readonly changed: number }
  | { readonly ok: false; readonly reason: string }

/** Excel's serial for a local calendar date */
export function dateSerial(date: Date): number {
  return Math.round(
    (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - Date.UTC(1899, 11, 30)) /
      86_400_000,
  )
}

/**
 * Whether text like "8/12" has the day first in a locale. Only locales where
 * Excel's reading was checked answer; in any other, such text is a limit.
 */
function dayFirstOf(culture: string | undefined): boolean | null {
  const name = (culture ?? '').toLowerCase()
  if (name === 'ko-kr' || name === 'en-us') return false
  return null
}

/** Lays the edits of a save over the model: cell values and formulas, and resized tables. */
export function applyEditsToModel(
  book: WorkbookModel,
  edits: readonly CellEdit[],
  tableEdits: readonly SheetTableEdit[],
): void {
  for (const edit of edits) {
    if (!edit.writeValue) continue
    const sheet = book.sheetByName.get(edit.sheetName.toLowerCase())
    if (!sheet) throw new FormulaLimit(`an edit on the unknown sheet ${edit.sheetName}`)
    const key = cellKey(edit.row, edit.column)
    if (edit.cell.formula) {
      const text = edit.cell.formula.replace(/^=/, '')
      sheet.cells.set(key, {
        stored: null,
        formula: {
          text,
          baseRow: edit.row,
          baseColumn: edit.column,
          array: null,
          dynamic: false,
          ast: null,
        },
        inStoredSpill: false,
        rich: false,
      })
    } else if (edit.cell.value === null) {
      sheet.cells.delete(key)
      continue
    } else {
      sheet.cells.set(key, {
        stored: edit.cell.value,
        formula: null,
        inStoredSpill: false,
        rich: false,
      })
    }
    if (edit.row > sheet.maxRow) sheet.maxRow = edit.row
    if (edit.column > sheet.maxColumn) sheet.maxColumn = edit.column
  }
  for (const edit of tableEdits) {
    if (!edit.area) continue
    const table = book.tables.get(edit.tableName.toLowerCase())
    if (!table) throw new FormulaLimit(`an edit of the unknown table ${edit.tableName}`)
    table.r1 = edit.area.startRow
    table.c1 = edit.area.startColumn
    table.r2 = edit.area.endRow
    table.c2 = edit.area.endColumn
  }
}

const cached = (value: Scalar): FormulaCachedValue => {
  if (value instanceof XlError) {
    // errors newer than the file format are stored as #VALUE! (Excel adds value metadata, which is not written here)
    return { error: value.code === '#SPILL!' || value.code === '#CALC!' ? '#VALUE!' : value.code }
  }
  return value
}

/** whether a computed value is what the file already stores */
function unchanged(value: Scalar, stored: Scalar, rich: boolean): boolean {
  if (value === stored) return true
  if (value instanceof XlError && stored instanceof XlError) {
    return (
      rich && stored.code === '#VALUE!' && (value.code === '#SPILL!' || value.code === '#CALC!')
    )
  }
  return false
}

type Entry = SheetFormulaValues['cells'][number]

function sheetValues(calc: Calculator, sheet: SheetModel): Entry[] {
  const entries = new Map<number, Entry>()
  const keys = [...sheet.cells.keys()]
    .filter((k) => sheet.cells.get(k)!.formula)
    .sort((a, b) => a - b)
  // first the cells array formulas no longer fill, then what is filled now: a cell one array left may be another's
  const arrays: { key: number; result: ArrayValue | null }[] = []
  for (const key of keys) {
    const cell = sheet.cells.get(key)!
    const row = keyRow(key)
    const column = keyColumn(key)
    const result = calc.formulaResult(sheet, row, column, cell)
    const value = isArrayValue(result) ? (result.data[0] ?? null) : result
    const range = cell.formula!.array
    let arrayRef: string | undefined
    if (range) {
      const rows = isArrayValue(result) ? result.rows : 1
      const cols = isArrayValue(result) ? result.cols : 1
      if (
        cell.formula!.dynamic &&
        (range.r2 !== row + rows - 1 || range.c2 !== column + cols - 1)
      ) {
        arrayRef =
          rows === 1 && cols === 1
            ? address(row, column)
            : `${address(row, column)}:${address(row + rows - 1, column + cols - 1)}`
      }
      arrays.push({ key, result: isArrayValue(result) ? result : null })
    }
    const clearValueMeta = cell.rich && !(value instanceof XlError)
    if (!unchanged(value, cell.stored, cell.rich) || arrayRef !== undefined || clearValueMeta) {
      entries.set(key, {
        row,
        column,
        value: cached(value),
        ...(arrayRef === undefined ? {} : { arrayRef }),
        ...(clearValueMeta ? { clearValueMeta: true as const } : {}),
      })
    }
  }
  for (const { key } of arrays) {
    const range = sheet.cells.get(key)!.formula!.array!
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const k = cellKey(r, c)
        const old = sheet.cells.get(k)
        if (k === key || !old || old.formula || old.stored === null) continue
        entries.set(k, { row: r, column: c, value: null, spill: true })
      }
    }
  }
  for (const { key, result } of arrays) {
    if (!result) continue
    const row = keyRow(key)
    const column = keyColumn(key)
    for (let r = 0; r < result.rows; r++) {
      for (let c = 0; c < result.cols; c++) {
        if (r === 0 && c === 0) continue
        const k = cellKey(row + r, column + c)
        const value = result.data[r * result.cols + c] ?? null
        const old = sheet.cells.get(k)
        if (unchanged(value, old?.stored ?? null, old?.rich ?? false)) entries.delete(k)
        else entries.set(k, { row: row + r, column: column + c, value: cached(value), spill: true })
      }
    }
  }
  return [...entries.values()]
}

/**
 * Recalculates the workbook as it will be after `edits` and `tableEdits`,
 * and returns the cached values to save with them.
 */
export function recalculateAfterEdits(
  workbookBytes: Uint8Array,
  edits: readonly CellEdit[],
  tableEdits: readonly SheetTableEdit[],
  options: RecalcOptions,
): RecalcOutcome {
  try {
    const book = readWorkbookModel(workbookBytes)
    applyEditsToModel(book, edits, tableEdits)
    const calc = new Calculator(book, {
      today: options.today,
      dateSeparator: '/',
      dayFirst: dayFirstOf(options.systemCulture),
    })
    const formulaValues: SheetFormulaValues[] = []
    let changed = 0
    for (const sheet of book.sheets) {
      const cells = sheetValues(calc, sheet)
      if (cells.length === 0) continue
      changed += cells.length
      formulaValues.push({ sheetName: sheet.name, cells })
    }
    return { ok: true, formulaValues, changed }
  } catch (error) {
    if (error instanceof FormulaLimit) return { ok: false, reason: error.message }
    // formulas that depend on each other thousands deep run out of stack
    if (error instanceof RangeError)
      return { ok: false, reason: 'Formula limit: formulas nested too deeply' }
    throw error
  }
}
