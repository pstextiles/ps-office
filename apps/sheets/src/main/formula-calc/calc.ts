/**
 * The built-in recalculation (PS Office): evaluates a workbook's formulas the
 * way Excel does, fast enough for workbooks with hundreds of thousands of
 * them — exact lookups go through an index built once per lookup range, and
 * every formula cell is computed once and remembered.
 *
 * Formulas flagged as array formulas are evaluated with array semantics
 * (operators and scalar functions spread over arrays, results spill); the
 * others with implicit intersection, as Excel treats them. What the engine
 * can't be sure to compute as Excel does (ordering different texts, text
 * that may be a date, an unknown function) raises FormulaLimit.
 */
import {
  ERR,
  XlError,
  cellKey,
  formulaTree,
  keyColumn,
  keyRow,
  type CellModel,
  type Scalar,
  type SheetModel,
  type WorkbookModel,
} from './model'
import { FormulaLimit, MAX_COLUMN, MAX_ROW, type Node } from './parser'

export interface ArrayValue {
  readonly kind: 'array'
  readonly rows: number
  readonly cols: number
  /** row-major */
  readonly data: Scalar[]
  /**
   * The last row stands for the endless empty rows of a whole-column
   * reference. An array that still carries it can't be given a size.
   */
  readonly open?: boolean
}
export interface RefValue {
  readonly kind: 'ref'
  readonly sheet: SheetModel
  readonly r1: number
  readonly c1: number
  readonly r2: number
  readonly c2: number
}
export type Value = Scalar | ArrayValue | RefValue

export const isArrayValue = (v: Value): v is ArrayValue =>
  typeof v === 'object' && v !== null && !(v instanceof XlError) && v.kind === 'array'
export const isRefValue = (v: Value): v is RefValue =>
  typeof v === 'object' && v !== null && !(v instanceof XlError) && v.kind === 'ref'
const isError = (v: unknown): v is XlError => v instanceof XlError

export interface Ctx {
  readonly sheet: SheetModel
  readonly row: number
  readonly column: number
  /** how far the cell is from the cell its formula text was written for (shared formulas) */
  readonly dr: number
  readonly dc: number
  readonly arrayMode: boolean
  readonly vars: ReadonlyMap<string, Value> | null
}

export interface CalcOptions {
  /** today's date as a serial (TODAY) */
  readonly today: number
  /** the date separator TEXT writes for `/` in a date format ("/" or "-") */
  readonly dateSeparator: string
  /**
   * Text like "8/12" read as a date has the day first (d/M) or the month
   * (M/d); null where Excel's reading in the PC's locale was not checked,
   * and such text is then a limit.
   */
  readonly dayFirst: boolean | null
}

/**
 * The options of the calculation in progress: text-to-number conversion
 * depends on them (a text date takes the year of `today` and the locale's
 * day/month order). Set by the Calculator; one calculation runs at a time.
 */
let ACTIVE: CalcOptions = { today: 0, dateSeparator: '/', dayFirst: false }

const array = (rows: number, cols: number, data: Scalar[], open = false): ArrayValue =>
  open ? { kind: 'array', rows, cols, data, open } : { kind: 'array', rows, cols, data }

// ---- coercions

const EPOCH = Date.UTC(1899, 11, 30)

const MONTH_WORD =
  /(?:^|[^a-z])(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|ene|abr|ago|dic)[a-z]*(?:[^a-z]|$)/i
const NUMERIC_TEXT = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/
const GROUPED_TEXT = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d*)?$/

/** text as the number Excel reads it as, null when it is not one; raises FormulaLimit for text Excel may read as a date, time or currency */
export function numericText(text: string): number | null {
  const s = text.trim()
  if (s === '') return null
  if (NUMERIC_TEXT.test(s)) return Number(s)
  if (GROUPED_TEXT.test(s)) return Number(s.replace(/,/g, ''))
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)%$/.test(s)) return Number(s.slice(0, -1)) / 100
  const date = textDate(s)
  if (date !== null) return date
  // digits with date, time or money marks, or with a month name: Excel may read it as a number in some locale
  if (/\d/.test(s) && (/^[\d\s./:,\-+$€₩%()apmAPM년월일시분초]+$/.test(s) || MONTH_WORD.test(s))) {
    throw new FormulaLimit(
      `text that Excel may read as a date or an amount (${JSON.stringify(text)})`,
    )
  }
  return null
}

/** yyyy-M-d, and M/d or d/M (by locale) with the year or, without one, this year: the date serial, or null */
function textDate(s: string): number | null {
  let y: number
  let m: number
  let d: number
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s)
  if (match) {
    y = Number(match[1])
    m = Number(match[2])
    d = Number(match[3])
  } else if ((match = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?$/.exec(s))) {
    const a = Number(match[1])
    const b = Number(match[2])
    if (ACTIVE.dayFirst === null)
      throw new FormulaLimit(
        `text that Excel reads as a date by the PC's locale (${JSON.stringify(s)})`,
      )
    m = ACTIVE.dayFirst ? b : a
    d = ACTIVE.dayFirst ? a : b
    y = match[3] === undefined ? dateParts(ACTIVE.today).y : Number(match[3])
  } else return null
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1901 || y > 9999) return null
  const t = new Date(0)
  t.setUTCFullYear(y, m - 1, d)
  if (t.getUTCMonth() !== m - 1) return null
  return Math.round((t.getTime() - EPOCH) / 86_400_000)
}

function toNumber(v: Scalar): number | XlError {
  if (typeof v === 'number') return v
  if (v === null) return 0
  if (typeof v === 'boolean') return v ? 1 : 0
  if (typeof v === 'string') return numericText(v) ?? ERR.VALUE
  return v
}

/** a number as General-format text (concatenation, TEXTJOIN, LEFT of a number) */
export function generalText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return n === 0 ? '0' : String(n)
  const abs = Math.abs(n)
  if (!Number.isFinite(n) || abs >= 1e15 || abs < 1e-4)
    throw new FormulaLimit(`a number written as text (${n})`)
  return String(Number(n.toPrecision(15)))
}

function toText(v: Scalar): string | XlError {
  if (typeof v === 'string') return v
  if (v === null) return ''
  if (typeof v === 'number') return generalText(v)
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return v
}

function toBool(v: Scalar): boolean | XlError {
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v === null) return false
  if (typeof v === 'string') {
    const upper = v.toUpperCase()
    if (upper === 'TRUE') return true
    if (upper === 'FALSE') return false
    return ERR.VALUE
  }
  return v
}

const typeRank = (v: number | string | boolean): number =>
  typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2

/** -1, 0, 1 as Excel's comparison operators order two values; blanks take the other side's type */
function compareScalars(a: Exclude<Scalar, XlError>, b: Exclude<Scalar, XlError>): number {
  let x = a
  let y = b
  if (x === null && y === null) return 0
  if (x === null) x = typeof y === 'number' ? 0 : typeof y === 'string' ? '' : false
  if (y === null) y = typeof x === 'number' ? 0 : typeof x === 'string' ? '' : false
  const rx = typeRank(x)
  const ry = typeRank(y)
  if (rx !== ry) return rx < ry ? -1 : 1
  if (typeof x === 'number') return x === y ? 0 : x < (y as number) ? -1 : 1
  if (typeof x === 'boolean') return x === y ? 0 : x ? 1 : -1
  return textOrder(x, y as string)
}

/**
 * Two texts in Excel's order. Case is ignored, and so are hyphens and
 * apostrophes (Excel's documented sort rule). The order is decided at the
 * first position where the texts differ, and only when both characters there
 * are digits or English letters (digits first): where another character
 * decides, the order follows the language's collation, which the engine
 * does not reproduce.
 */
function textOrder(a: string, b: string): number {
  const x = a.toLowerCase()
  const y = b.toLowerCase()
  if (x === y) return 0
  if (x === '') return -1
  if (y === '') return 1
  const p = x.replace(/[-']/g, '')
  const q = y.replace(/[-']/g, '')
  const n = Math.min(p.length, q.length)
  for (let i = 0; i < n; i++) {
    const c = p.charCodeAt(i)
    const d = q.charCodeAt(i)
    if (c === d) continue
    if (isAlphanumeric(c) && isAlphanumeric(d)) return c < d ? -1 : 1
    throw new FormulaLimit(`the order of two texts (${JSON.stringify(a)}, ${JSON.stringify(b)})`)
  }
  if (p.length !== q.length) return p.length < q.length ? -1 : 1
  // the same but for hyphens or apostrophes
  throw new FormulaLimit(`the order of two texts (${JSON.stringify(a)}, ${JSON.stringify(b)})`)
}
/** 0-9 or a-z */
const isAlphanumeric = (code: number): boolean =>
  (code >= 48 && code <= 57) || (code >= 97 && code <= 122)

/** whether two values are equal to Excel's `=`: text ignores case, a blank equals 0, "" and FALSE */
function equalScalars(a: Exclude<Scalar, XlError>, b: Exclude<Scalar, XlError>): boolean {
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase()
  if (a === null && typeof b === 'string') return b === ''
  if (b === null && typeof a === 'string') return a === ''
  return compareScalars(a, b) === 0
}

/** a value for exact matching (lookups, UNIQUE): numbers, case-folded text and booleans are distinct kinds */
const matchKey = (v: Scalar): string | null =>
  typeof v === 'number'
    ? `n${v}`
    : typeof v === 'string'
      ? `s${v.toLowerCase()}`
      : typeof v === 'boolean'
        ? `b${v}`
        : null

function dateParts(serial: number): { y: number; m: number; d: number; dow: number } {
  const day = Math.floor(serial)
  if (day < 61) throw new FormulaLimit(`a date before March 1900 (${serial})`)
  const t = new Date(EPOCH + day * 86_400_000)
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate(), dow: t.getUTCDay() }
}

// ---- the calculator

type Fn = (calc: Calculator, args: readonly (Node | null)[], ctx: Ctx) => Value

interface Spill {
  readonly array: ArrayValue
  readonly offset: number
}

export class Calculator {
  private readonly done = new Map<CellModel, Scalar | ArrayValue>()
  private readonly computing = new Set<CellModel>()
  /** per sheet: array formula cells not computed yet */
  private readonly pending = new Map<SheetModel, number[]>()
  /** per sheet: the cells an array formula fills now (not the formula cell itself) */
  private readonly spills = new Map<SheetModel, Map<number, Spill>>()
  private readonly lookups = new Map<string, Map<string, number>>()
  /** per sheet: the last row and column in use, spills included */
  private readonly extent = new Map<SheetModel, { row: number; column: number }>()

  constructor(
    readonly book: WorkbookModel,
    readonly options: CalcOptions,
  ) {
    ACTIVE = options
    for (const sheet of book.sheets) {
      this.pending.set(sheet, [...sheet.arrayMasters])
      this.spills.set(sheet, new Map())
      this.extent.set(sheet, { row: sheet.maxRow, column: sheet.maxColumn })
    }
  }

  // ---- cells

  /** the value a cell shows after recalculation */
  cellValue(sheet: SheetModel, row: number, column: number): Scalar {
    const key = cellKey(row, column)
    const cell = sheet.cells.get(key)
    if (cell?.formula) {
      const v = this.formulaResult(sheet, row, column, cell)
      return isArrayValue(v) ? (v.data[0] ?? null) : v
    }
    this.settle(sheet, row, column)
    const spill = this.spills.get(sheet)!.get(key)
    if (spill) return spill.array.data[spill.offset] ?? null
    if (!cell || cell.inStoredSpill) return null
    return cell.stored
  }

  /** every array formula above and to the left of a cell is computed, so its spill is known */
  private settle(sheet: SheetModel, row: number, column: number): void {
    const pending = this.pending.get(sheet)!
    if (pending.length === 0) return
    for (let i = 0; i < pending.length;) {
      const key = pending[i]!
      const r = keyRow(key)
      const c = keyColumn(key)
      const cell = sheet.cells.get(key)!
      if (r <= row && c <= column && !this.computing.has(cell)) {
        pending.splice(i, 1)
        this.formulaResult(sheet, r, c, cell)
        // computing it may have settled others: start over
        i = 0
      } else i++
    }
  }

  /** the result of a formula cell: an array for an array formula, else a value */
  formulaResult(
    sheet: SheetModel,
    row: number,
    column: number,
    cell: CellModel,
  ): Scalar | ArrayValue {
    const known = this.done.get(cell)
    if (known !== undefined) return known
    if (this.computing.has(cell))
      throw new FormulaLimit(`a circular reference at ${sheet.name}!${address(row, column)}`)
    const formula = cell.formula!
    this.computing.add(cell)
    let result: Scalar | ArrayValue
    try {
      const ctx: Ctx = {
        sheet,
        row,
        column,
        dr: row - formula.baseRow,
        dc: column - formula.baseColumn,
        arrayMode: formula.array !== null,
        vars: null,
      }
      let tree: Node
      try {
        tree = formulaTree(formula)
      } catch (error) {
        if (error instanceof FormulaLimit)
          throw new FormulaLimit(
            `${error.what} at ${sheet.name}!${address(row, column)}`,
            formula.text,
          )
        throw error
      }
      const value = this.eval(tree, ctx)
      result = formula.array
        ? this.spill(sheet, row, column, cell, value)
        : (this.single(value, ctx) ?? 0)
    } finally {
      this.computing.delete(cell)
    }
    this.done.set(cell, result)
    return result
  }

  /** an array formula's result placed on the sheet: a dynamic one takes its own size, a fixed one fills its range */
  private spill(
    sheet: SheetModel,
    row: number,
    column: number,
    cell: CellModel,
    value: Value,
  ): Scalar | ArrayValue {
    const formula = cell.formula!
    const range = formula.array!
    let result = isRefValue(value) ? this.materialize(value) : value
    if (!isArrayValue(result)) {
      if (range.r1 === range.r2 && range.c1 === range.c2) return result ?? 0
      result = array(1, 1, [result])
    }
    if (result.open)
      throw new FormulaLimit(`an array without an end at ${sheet.name}!${address(row, column)}`)
    let rows = result.rows
    let cols = result.cols
    let data: Scalar[]
    if (formula.dynamic) {
      data = result.data.map((v) => v ?? 0)
      // a dynamic array needs its cells free: anything not left by its own earlier spill blocks it
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          if (r === 0 && c === 0) continue
          const other = sheet.cells.get(cellKey(row + r, column + c))
          if (other && (other.formula || (other.stored !== null && !other.inStoredSpill)))
            return ERR.SPILL
          if (this.spills.get(sheet)!.has(cellKey(row + r, column + c))) return ERR.SPILL
        }
      }
    } else {
      rows = range.r2 - range.r1 + 1
      cols = range.c2 - range.c1 + 1
      data = new Array<Scalar>(rows * cols)
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const rr = result.rows === 1 ? 0 : r
          const cc = result.cols === 1 ? 0 : c
          data[r * cols + c] =
            rr < result.rows && cc < result.cols
              ? (result.data[rr * result.cols + cc] ?? 0)
              : ERR.NA
        }
      }
    }
    const placed = array(rows, cols, data)
    const spills = this.spills.get(sheet)!
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (r === 0 && c === 0) continue
        spills.set(cellKey(row + r, column + c), { array: placed, offset: r * cols + c })
      }
    }
    const extent = this.extent.get(sheet)!
    if (row + rows - 1 > extent.row) extent.row = row + rows - 1
    if (column + cols - 1 > extent.column) extent.column = column + cols - 1
    return placed
  }

  /** the cells a formula fills beyond its own (row offset, column offset from the formula cell) */
  spillOf(cell: CellModel): ArrayValue | null {
    const v = this.done.get(cell)
    return v !== undefined && isArrayValue(v) ? v : null
  }

  // ---- values

  /** a reference's cells as an array; a whole-column or whole-row one ends at the sheet's last cell plus one empty row or column */
  materialize(ref: RefValue, lastRow?: number): ArrayValue {
    this.settle(ref.sheet, Math.min(ref.r2, MAX_ROW), Math.min(ref.c2, MAX_COLUMN))
    let r2 = ref.r2
    const c2 = ref.c2
    let open = false
    if (ref.c1 === 0 && ref.c2 === MAX_COLUMN)
      throw new FormulaLimit('a whole row used as an array')
    if (ref.r1 === 0 && ref.r2 === MAX_ROW) {
      r2 = lastRow ?? this.openRow(ref.sheet)
      open = true
    }
    const rows = r2 - ref.r1 + 1
    const cols = c2 - ref.c1 + 1
    if (rows * cols > 20_000_000) throw new FormulaLimit('a very large range used as an array')
    const data = new Array<Scalar>(rows * cols)
    let i = 0
    for (let r = ref.r1; r <= r2; r++) {
      for (let c = ref.c1; c <= c2; c++) data[i++] = this.readSettled(ref.sheet, r, c)
    }
    return array(rows, cols, data, open)
  }

  /** the row a whole-column array of this sheet ends on: one empty row after the last cell in use */
  openRow(sheet: SheetModel): number {
    this.settle(sheet, MAX_ROW, MAX_COLUMN)
    return Math.min(this.extent.get(sheet)!.row + 1, MAX_ROW)
  }

  /** a cell's value when the array formulas around it are already computed */
  private readSettled(sheet: SheetModel, row: number, column: number): Scalar {
    const key = cellKey(row, column)
    const cell = sheet.cells.get(key)
    if (cell?.formula) {
      const v = this.formulaResult(sheet, row, column, cell)
      return isArrayValue(v) ? (v.data[0] ?? null) : v
    }
    const spill = this.spills.get(sheet)!.get(key)
    if (spill) return spill.array.data[spill.offset] ?? null
    if (!cell || cell.inStoredSpill) return null
    return cell.stored
  }

  /** a value where one value is wanted, outside array evaluation: a range gives the cell in the formula's row or column, an array its first element */
  single(value: Value, ctx: Ctx): Scalar {
    if (isArrayValue(value)) return value.data[0] ?? null
    if (!isRefValue(value)) return value
    if (value.r1 === value.r2 && value.c1 === value.c2)
      return this.cellValue(value.sheet, value.r1, value.c1)
    if (value.c1 === value.c2 && ctx.row >= value.r1 && ctx.row <= value.r2)
      return this.cellValue(value.sheet, ctx.row, value.c1)
    if (value.r1 === value.r2 && ctx.column >= value.c1 && ctx.column <= value.c2) {
      return this.cellValue(value.sheet, value.r1, ctx.column)
    }
    return ERR.VALUE
  }

  /** an operand for array evaluation: a one-cell reference is its value, a range its array */
  operand(value: Value): Scalar | ArrayValue {
    if (!isRefValue(value)) return value
    if (value.r1 === value.r2 && value.c1 === value.c2)
      return this.cellValue(value.sheet, value.r1, value.c1)
    return this.materialize(value)
  }

  /** applies a function of scalars across arrays, Excel's way: a single row or column repeats, a missing position is #N/A */
  broadcast(
    values: readonly (Scalar | ArrayValue)[],
    f: (...scalars: Scalar[]) => Scalar,
  ): Scalar | ArrayValue {
    let rows = 0
    let cols = 0
    let open = false
    for (const v of values) {
      if (!isArrayValue(v)) continue
      if (v.rows > rows) rows = v.rows
      if (v.cols > cols) cols = v.cols
      if (v.open) open = true
    }
    if (rows === 0) return f(...(values as Scalar[]))
    if (open) {
      for (const v of values) {
        if (isArrayValue(v) && v.rows !== 1 && (v.rows !== rows || !v.open)) {
          throw new FormulaLimit('whole columns of different sheets used together as arrays')
        }
      }
    }
    const data = new Array<Scalar>(rows * cols)
    const scalars = new Array<Scalar>(values.length)
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let missing = false
        for (let k = 0; k < values.length; k++) {
          const v = values[k]!
          if (!isArrayValue(v)) {
            scalars[k] = v
            continue
          }
          const rr = v.rows === 1 ? 0 : r
          const cc = v.cols === 1 ? 0 : c
          if (rr >= v.rows || cc >= v.cols) missing = true
          else scalars[k] = v.data[rr * v.cols + cc] ?? null
        }
        data[r * cols + c] = missing ? ERR.NA : f(...scalars)
      }
    }
    return array(rows, cols, data, open)
  }

  /** evaluates arguments that each want one value, and applies `f` — across arrays in array evaluation */
  scalars(
    nodes: readonly (Node | null)[],
    ctx: Ctx,
    f: (...scalars: (Scalar | undefined)[]) => Scalar,
  ): Value {
    if (!ctx.arrayMode)
      return f(...nodes.map((n) => (n ? this.single(this.eval(n, ctx), ctx) : undefined)))
    const values = nodes.map((n) => (n ? this.operand(this.eval(n, ctx)) : MISSING))
    return this.broadcast(values as (Scalar | ArrayValue)[], (...s) =>
      f(...s.map((x) => (x === MISSING ? undefined : x))),
    )
  }

  /** an argument that takes a range or an array: evaluated with array semantics whatever the formula */
  range(node: Node | null, ctx: Ctx): Value {
    if (!node) return null
    return this.eval(node, ctx.arrayMode ? ctx : { ...ctx, arrayMode: true })
  }

  /** a range argument as an array (a single value is a 1×1 array) or the error it is */
  arrayOf(value: Value): ArrayValue | XlError {
    const v = isRefValue(value) ? this.materialize(value) : value
    if (isArrayValue(v)) return v
    if (isError(v)) return v
    return array(1, 1, [v])
  }

  // ---- lookups

  /**
   * The position (0-based) of the first cell of a one-dimensional range or
   * array that exactly matches `value`, or -1. A range's positions are
   * indexed once and kept.
   */
  findExact(vector: Value, value: Scalar): number | XlError {
    if (isError(value)) return value
    if (value === null || value === '') throw new FormulaLimit('a lookup of an empty value')
    const key = matchKey(value)!
    if (isRefValue(vector)) {
      if (vector.r1 !== vector.r2 && vector.c1 !== vector.c2)
        throw new FormulaLimit('a lookup in a two-dimensional range')
      const id = `${vector.sheet.index}:${vector.r1}:${vector.c1}:${vector.r2}:${vector.c2}`
      let index = this.lookups.get(id)
      if (!index) {
        index = new Map()
        const sheet = vector.sheet
        this.settle(sheet, Math.min(vector.r2, MAX_ROW), Math.min(vector.c2, MAX_COLUMN))
        const extent = this.extent.get(sheet)!
        const vertical = vector.c1 === vector.c2
        const last = vertical ? Math.min(vector.r2, extent.row) : Math.min(vector.c2, extent.column)
        const first = vertical ? vector.r1 : vector.c1
        for (let p = first; p <= last; p++) {
          const v = vertical
            ? this.readSettled(sheet, p, vector.c1)
            : this.readSettled(sheet, vector.r1, p)
          const k = matchKey(v)
          if (k !== null && k !== 's' && !index.has(k)) index.set(k, p - first)
        }
        this.lookups.set(id, index)
      }
      return index.get(key) ?? -1
    }
    const arr = this.arrayOf(vector)
    if (isError(arr)) return arr
    if (arr.rows !== 1 && arr.cols !== 1)
      throw new FormulaLimit('a lookup in a two-dimensional array')
    for (let i = 0; i < arr.data.length; i++) if (matchKey(arr.data[i] ?? null) === key) return i
    return -1
  }

  /** the element at a position of a one-dimensional range or array */
  elementAt(vector: Value, position: number): Scalar {
    if (isRefValue(vector)) {
      const vertical = vector.c1 === vector.c2
      if (!vertical && vector.r1 !== vector.r2)
        throw new FormulaLimit('a lookup result from a two-dimensional range')
      const r = vertical ? vector.r1 + position : vector.r1
      const c = vertical ? vector.c1 : vector.c1 + position
      if (r > vector.r2 || c > vector.c2) return ERR.NA
      return this.cellValue(vector.sheet, r, c)
    }
    const arr = this.arrayOf(vector)
    if (isError(arr)) return arr
    if (arr.rows !== 1 && arr.cols !== 1)
      throw new FormulaLimit('a lookup result from a two-dimensional array')
    return position < arr.data.length ? (arr.data[position] ?? null) : ERR.NA
  }

  // ---- evaluation

  eval(node: Node, ctx: Ctx): Value {
    switch (node.t) {
      case 'num':
      case 'str':
      case 'bool':
        return node.v
      case 'err':
        return XlError.of(node.v)
      case 'ref': {
        const sheet =
          node.sheet === null ? ctx.sheet : this.book.sheetByName.get(node.sheet.toLowerCase())
        if (!sheet) return ERR.REF
        const r1 = node.r1Abs ? node.r1 : node.r1 + ctx.dr
        const r2 = node.r2Abs ? node.r2 : node.r2 + ctx.dr
        const c1 = node.c1Abs ? node.c1 : node.c1 + ctx.dc
        const c2 = node.c2Abs ? node.c2 : node.c2 + ctx.dc
        if (r1 < 0 || c1 < 0 || r2 > MAX_ROW || c2 > MAX_COLUMN) return ERR.REF
        return {
          kind: 'ref',
          sheet,
          r1: Math.min(r1, r2),
          c1: Math.min(c1, c2),
          r2: Math.max(r1, r2),
          c2: Math.max(c1, c2),
        }
      }
      case 'sref': {
        const table = this.book.tables.get(node.table.toLowerCase())
        if (!table) return ERR.REF
        let c1 = table.c1
        let c2 = table.c2
        if (node.columns) {
          const find = (name: string): number =>
            table.columns.findIndex((c) => c.toLowerCase() === name.toLowerCase())
          const a = find(node.columns[0])
          const b = find(node.columns[1])
          if (a < 0 || b < 0) return ERR.REF
          c1 = table.c1 + Math.min(a, b)
          c2 = table.c1 + Math.max(a, b)
        }
        const dataFirst = table.r1 + table.headerRows
        const dataLast = table.r2 - table.totalsRows
        if (node.thisRow) {
          if (table.sheet !== ctx.sheet || ctx.row < dataFirst || ctx.row > dataLast)
            return ERR.VALUE
          return { kind: 'ref', sheet: table.sheet, r1: ctx.row, c1, r2: ctx.row, c2 }
        }
        if (node.area === 'all')
          return { kind: 'ref', sheet: table.sheet, r1: table.r1, c1, r2: table.r2, c2 }
        if (node.area === 'headers') {
          if (table.headerRows === 0) return ERR.REF
          return { kind: 'ref', sheet: table.sheet, r1: table.r1, c1, r2: table.r1, c2 }
        }
        if (node.area === 'totals') {
          if (table.totalsRows === 0) return ERR.REF
          return { kind: 'ref', sheet: table.sheet, r1: table.r2, c1, r2: table.r2, c2 }
        }
        return { kind: 'ref', sheet: table.sheet, r1: dataFirst, c1, r2: dataLast, c2 }
      }
      case 'name': {
        const bound = ctx.vars?.get(node.name.toLowerCase())
        if (bound !== undefined) return bound
        throw new FormulaLimit(`the name ${node.name}`)
      }
      case 'array': {
        const rows = node.rows.length
        const cols = node.rows[0]!.length
        const data: Scalar[] = []
        for (const row of node.rows)
          for (const item of row) data.push(this.eval(item, ctx) as Scalar)
        return array(rows, cols, data)
      }
      case 'un':
        return this.unary(node.op, node.x, ctx)
      case 'bin':
        return this.binary(node.op, node.a, node.b, ctx)
      case 'call': {
        const fn = FUNCTIONS[node.name]
        if (!fn) throw new FormulaLimit(`the function ${node.name}`)
        return fn(this, node.args, ctx)
      }
    }
  }

  private unary(op: '-' | '+' | '%' | '@' | '#', x: Node, ctx: Ctx): Value {
    if (op === '+') return this.eval(x, ctx)
    if (op === '@') return this.single(this.eval(x, { ...ctx, arrayMode: false }), ctx)
    if (op === '#') {
      const target = this.eval(x, ctx)
      if (isError(target)) return target
      if (!isRefValue(target) || target.r1 !== target.r2 || target.c1 !== target.c2) return ERR.REF
      const cell = target.sheet.cells.get(cellKey(target.r1, target.c1))
      if (!cell?.formula?.array) return ERR.REF
      return this.formulaResult(target.sheet, target.r1, target.c1, cell)
    }
    return this.scalars([x], ctx, (v) => {
      const n = toNumber(v ?? null)
      if (isError(n)) return n
      return op === '-' ? (n === 0 ? 0 : -n) : n / 100
    })
  }

  private binary(op: string, a: Node, b: Node, ctx: Ctx): Value {
    if (op === ':') {
      const x = this.eval(a, ctx)
      const y = this.eval(b, ctx)
      if (isError(x)) return x
      if (isError(y)) return y
      if (!isRefValue(x) || !isRefValue(y) || x.sheet !== y.sheet)
        throw new FormulaLimit('a range between two expressions')
      return {
        kind: 'ref',
        sheet: x.sheet,
        r1: Math.min(x.r1, y.r1),
        c1: Math.min(x.c1, y.c1),
        r2: Math.max(x.r2, y.r2),
        c2: Math.max(x.c2, y.c2),
      }
    }
    return this.scalars([a, b], ctx, (x, y) => binaryScalar(op, x ?? null, y ?? null))
  }
}

/** marks an omitted argument while broadcasting */
const MISSING = Symbol('missing') as unknown as Scalar

function binaryScalar(op: string, x: Scalar, y: Scalar): Scalar {
  if (op === '&') {
    const a = toText(x)
    if (isError(a)) return a
    const b = toText(y)
    if (isError(b)) return b
    return a + b
  }
  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
    if (isError(x)) return x
    if (isError(y)) return y
    if (op === '=') return equalScalars(x, y)
    if (op === '<>') return !equalScalars(x, y)
    const c = compareScalars(x, y)
    switch (op) {
      case '<':
        return c < 0
      case '>':
        return c > 0
      case '<=':
        return c <= 0
      default:
        return c >= 0
    }
  }
  const a = toNumber(x)
  if (isError(a)) return a
  const b = toNumber(y)
  if (isError(b)) return b
  let n: number
  switch (op) {
    case '+':
      n = a + b
      break
    case '-':
      n = a - b
      break
    case '*':
      n = a * b
      break
    case '/':
      if (b === 0) return ERR.DIV0
      n = a / b
      break
    default:
      if (a === 0 && b === 0) return ERR.NUM
      if (a === 0 && b < 0) return ERR.DIV0
      n = Math.pow(a, b)
      break
  }
  if (Number.isNaN(n) || !Number.isFinite(n)) return ERR.NUM
  return n === 0 ? 0 : n
}

const A_CODE = 65
export function address(row: number, column: number): string {
  let letters = ''
  let n = column + 1
  while (n > 0) {
    n--
    letters = String.fromCharCode(A_CODE + (n % 26)) + letters
    n = Math.floor(n / 26)
  }
  return `${letters}${row + 1}`
}

// ---- functions

/** the first error among values, else null */
const firstError = (values: readonly (Scalar | undefined)[]): XlError | null => {
  for (const v of values) if (isError(v)) return v
  return null
}

/** a function of values that fails with the first error argument */
function plain(min: number, max: number, f: (args: (Scalar | undefined)[]) => Scalar): Fn {
  return (calc, nodes, ctx) => {
    if (nodes.length < min || nodes.length > max) return ERR.VALUE
    return calc.scalars(nodes, ctx, (...args) => firstError(args) ?? f(args))
  }
}

/** a function of values that sees error arguments itself */
function raw(min: number, max: number, f: (args: (Scalar | undefined)[]) => Scalar): Fn {
  return (calc, nodes, ctx) => {
    if (nodes.length < min || nodes.length > max) return ERR.VALUE
    return calc.scalars(nodes, ctx, (...args) => f(args))
  }
}

const numberArg = (v: Scalar | undefined, fallback?: number): number | XlError => {
  if (v === undefined) return fallback ?? ERR.VALUE
  return toNumber(v)
}
const textArg = (v: Scalar | undefined): string | XlError => toText(v ?? null)

/** every value of range-or-value arguments in order: [value, came from a range or array] */
function* flatten(
  calc: Calculator,
  nodes: readonly (Node | null)[],
  ctx: Ctx,
): Generator<[Scalar, boolean]> {
  for (const node of nodes) {
    if (!node) {
      yield [null, false]
      continue
    }
    const v = calc.eval(node, ctx)
    if (isRefValue(v)) {
      const arr = calc.materialize(v)
      for (const x of arr.data) yield [x ?? null, true]
    } else if (isArrayValue(v)) {
      for (const x of v.data) yield [x ?? null, true]
    } else yield [v, false]
  }
}

/** the numbers SUM, MIN, MAX and COUNT see: in ranges only numbers count; given directly, numeric text and logical values too */
function numbersOf(
  calc: Calculator,
  nodes: readonly (Node | null)[],
  ctx: Ctx,
  ignoreErrors: boolean,
): number[] | XlError {
  const out: number[] = []
  for (const [v, fromRange] of flatten(calc, nodes, ctx)) {
    if (isError(v)) {
      if (ignoreErrors) continue
      return v
    }
    if (typeof v === 'number') out.push(v)
    else if (!fromRange && v !== null) {
      const n = toNumber(v)
      if (isError(n)) {
        if (ignoreErrors) continue
        return n
      }
      out.push(n)
    } else if (!fromRange && v === null) out.push(0)
  }
  return out
}

function logical(
  calc: Calculator,
  nodes: readonly (Node | null)[],
  ctx: Ctx,
  all: boolean,
): Scalar {
  let seen = false
  let result = all
  for (const node of nodes) {
    if (!node) continue
    const v = calc.eval(node, ctx)
    const fromRange = isRefValue(v) || isArrayValue(v)
    const items = isRefValue(v) ? calc.materialize(v).data : isArrayValue(v) ? v.data : [v]
    for (const item of items) {
      if (isError(item)) return item
      if (item === null || (typeof item === 'string' && fromRange)) continue
      const b = toBool(item ?? null)
      if (isError(b)) return b
      seen = true
      if (all ? !b : b) result = !all
    }
  }
  return seen ? result : ERR.VALUE
}

/** a criterion of SUMIF, COUNTIF and their plural forms as a test of a cell value */
export function criterion(c: Scalar): (cell: Scalar) => boolean {
  if (isError(c)) throw new FormulaLimit('an error value as a criterion')
  if (c === null) c = 0
  if (typeof c === 'boolean') return (cell) => cell === c
  if (typeof c === 'number') {
    const n = c
    return (cell) =>
      typeof cell === 'number'
        ? cell === n
        : typeof cell === 'string'
          ? safeNumeric(cell) === n
          : false
  }
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!
  const op = m[1] ?? '='
  const rest = m[2]!
  if (rest === '') {
    if (m[1] === undefined) return (cell) => cell === null || cell === ''
    if (op === '=') return (cell) => cell === null
    if (op === '<>') return (cell) => cell !== null
    throw new FormulaLimit(`the criterion ${JSON.stringify(c)}`)
  }
  const upper = rest.toUpperCase()
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE'
    if (op === '=') return (cell) => cell === b
    if (op === '<>') return (cell) => cell !== b
    throw new FormulaLimit(`the criterion ${JSON.stringify(c)}`)
  }
  const n = safeNumeric(rest)
  if (n !== null) {
    switch (op) {
      case '=':
        return (cell) =>
          typeof cell === 'number'
            ? cell === n
            : typeof cell === 'string'
              ? safeNumeric(cell) === n
              : false
      case '<>':
        return (cell) =>
          !(typeof cell === 'number'
            ? cell === n
            : typeof cell === 'string'
              ? safeNumeric(cell) === n
              : false)
      case '<':
        return (cell) => typeof cell === 'number' && cell < n
      case '>':
        return (cell) => typeof cell === 'number' && cell > n
      case '<=':
        return (cell) => typeof cell === 'number' && cell <= n
      default:
        return (cell) => typeof cell === 'number' && cell >= n
    }
  }
  if (op !== '=' && op !== '<>')
    throw new FormulaLimit(`a criterion that orders text (${JSON.stringify(c)})`)
  const pattern = wildcard(rest)
  const test = pattern
    ? (cell: Scalar): boolean => typeof cell === 'string' && pattern.test(cell)
    : (cell: Scalar): boolean =>
        typeof cell === 'string' && cell.toLowerCase() === rest.toLowerCase()
  return op === '=' ? test : (cell) => !test(cell)
}

/** text as a number for a criterion; text that may be a date is a limit only when it reaches a comparison */
function safeNumeric(text: string): number | null {
  const s = text.trim()
  if (NUMERIC_TEXT.test(s)) return Number(s)
  if (s === '' || !/\d/.test(s)) return null
  return numericText(s)
}

/** a wildcard pattern (`?`, `*`, `~` escapes) as a whole-text, case-insensitive test; null when the text has no wildcard */
function wildcard(text: string): RegExp | null {
  if (!/[*?~]/.test(text)) return null
  let re = ''
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!
    if (ch === '~' && i + 1 < text.length) {
      re += escapeRe(text[++i]!)
    } else if (ch === '*') re += '[\\s\\S]*'
    else if (ch === '?') re += '[\\s\\S]'
    else re += escapeRe(ch)
  }
  return new RegExp(`^${re}$`, 'i')
}
const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** SUMIFS, COUNTIFS, SUMIF and COUNTIF over (range, criterion) pairs; `sum` is the range to add up, null to count */
function conditional(calc: Calculator, sum: Value, pairs: readonly [Value, Scalar][]): Scalar {
  const tests: { arr: ArrayValue; test: (cell: Scalar) => boolean }[] = []
  let rows = -1
  let cols = -1
  let lastRow = 0
  for (const v of [sum, ...pairs.map((p) => p[0])]) {
    if (isRefValue(v)) lastRow = Math.max(lastRow, calc.openRow(v.sheet))
  }
  const arrayOf = (v: Value): ArrayValue | XlError =>
    isRefValue(v) ? calc.materialize(v, lastRow) : calc.arrayOf(v)
  for (const [range, c] of pairs) {
    if (isError(range)) return range
    if (isError(c)) return c
    const arr = arrayOf(range)
    if (isError(arr)) return arr
    if (rows < 0) {
      rows = arr.rows
      cols = arr.cols
    } else if (arr.rows !== rows || arr.cols !== cols) return ERR.VALUE
    tests.push({ arr, test: criterion(c) })
  }
  let sumArr: ArrayValue | null = null
  if (sum !== null) {
    if (isError(sum)) return sum
    // a sum range takes the criteria range's shape from its top-left cell
    const anchored = isRefValue(sum)
      ? {
          ...sum,
          r2: sum.r1 === 0 && sum.r2 === MAX_ROW ? MAX_ROW : Math.min(sum.r1 + rows - 1, MAX_ROW),
          c2: Math.min(sum.c1 + cols - 1, MAX_COLUMN),
        }
      : sum
    const arr = arrayOf(anchored)
    if (isError(arr)) return arr
    sumArr = arr
  }
  let total = 0
  const n = rows * cols
  outer: for (let i = 0; i < n; i++) {
    for (const t of tests) if (!t.test(t.arr.data[i] ?? null)) continue outer
    if (!sumArr) total++
    else {
      const r = Math.floor(i / cols)
      const c = i % cols
      const v =
        r < sumArr.rows && c < sumArr.cols ? (sumArr.data[r * sumArr.cols + c] ?? null) : null
      if (isError(v)) return v
      if (typeof v === 'number') total += v
    }
  }
  return total
}

/** criteria may be arrays in an array formula: one result per criteria combination */
function conditionalFn(sumIndex: number | null, firstPair: number): Fn {
  return (calc, nodes, ctx) => {
    const sum = sumIndex === null ? null : calc.range(nodes[sumIndex] ?? null, ctx)
    const ranges: Value[] = []
    const criteria: (Node | null)[] = []
    for (let i = firstPair; i + 1 < nodes.length; i += 2) {
      ranges.push(calc.range(nodes[i] ?? null, ctx))
      criteria.push(nodes[i + 1] ?? null)
    }
    if (ranges.length === 0 || (nodes.length - firstPair) % 2 !== 0) return ERR.VALUE
    return calc.scalars(criteria, ctx, (...cs) =>
      conditional(
        calc,
        sum,
        ranges.map((r, i) => [r, cs[i] ?? null] as [Value, Scalar]),
      ),
    )
  }
}

/** the order SORT and SORTBY put values in: numbers, then text, then logical values, then errors; blanks last */
function sortOrder(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar): number =>
    v === null
      ? 4
      : typeof v === 'number'
        ? 0
        : typeof v === 'string'
          ? 1
          : typeof v === 'boolean'
            ? 2
            : 3
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === 'number') return a - (b as number)
  if (typeof a === 'string') {
    return textOrder(a, b as string)
  }
  if (typeof a === 'boolean') return a === b ? 0 : a ? 1 : -1
  return 0
}

const rowOf = (arr: ArrayValue, r: number): Scalar[] =>
  arr.data.slice(r * arr.cols, (r + 1) * arr.cols)

function workday(start: number, days: number, holidays: ReadonlySet<number>): number {
  let d = Math.floor(start)
  let left = Math.abs(Math.trunc(days))
  const step = days < 0 ? -1 : 1
  while (left > 0) {
    d += step
    const dow = (((d + 6) % 7) + 7) % 7
    if (dow !== 0 && dow !== 6 && !holidays.has(d)) left--
  }
  return d
}

function formatText(value: Scalar, format: string, options: CalcOptions): Scalar {
  if (typeof value === 'string') {
    const n = numericText(value)
    if (n === null) return value
    value = n
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  const n = value === null ? 0 : (value as number)
  const pad = (x: number): string => String(x).padStart(2, '0')
  switch (format) {
    case '0':
    case '#,##0': {
      // half away from zero, as Excel rounds
      const whole = Math.sign(n) * Math.floor(Math.abs(n) + 0.5)
      if (Math.abs(whole) >= 1e15) throw new FormulaLimit(`TEXT of a very large number (${n})`)
      const digits = String(Math.abs(whole))
      const body = format === '0' ? digits : digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
      return whole < 0 ? `-${body}` : body
    }
    case 'mm/dd': {
      const p = dateParts(n)
      return `${pad(p.m)}${options.dateSeparator}${pad(p.d)}`
    }
    case 'yyyy-mm': {
      const p = dateParts(n)
      return `${p.y}-${pad(p.m)}`
    }
    default:
      throw new FormulaLimit(`the TEXT format ${JSON.stringify(format)}`)
  }
}

const FUNCTIONS: Record<string, Fn> = {
  IF: (calc, nodes, ctx) => {
    if (nodes.length < 2 || nodes.length > 3 || !nodes[0]) return ERR.VALUE
    const pick = (c: Scalar, whenTrue: () => Value, whenFalse: () => Value): Value => {
      if (isError(c)) return c
      const b = toBool(c)
      if (isError(b)) return b
      return b ? whenTrue() : whenFalse()
    }
    const yes = (): Value => (nodes[1] ? calc.eval(nodes[1], ctx) : 0)
    const no = (): Value => (nodes.length < 3 ? false : nodes[2] ? calc.eval(nodes[2], ctx) : 0)
    const cond = calc.eval(nodes[0], ctx)
    if (!ctx.arrayMode) return pick(calc.single(cond, ctx), yes, no)
    const c = calc.operand(cond)
    if (!isArrayValue(c)) return pick(c, yes, no)
    const a = calc.operand(yes())
    const b = calc.operand(no())
    return calc.broadcast([c, a, b], (x, y, z) => {
      if (isError(x)) return x
      const t = toBool(x)
      if (isError(t)) return t
      return t ? y : z
    })
  },
  IFERROR: (calc, nodes, ctx) => {
    if (nodes.length !== 2) return ERR.VALUE
    const v = nodes[0] ? calc.eval(nodes[0], ctx) : null
    const alt = (): Value => (nodes[1] ? calc.eval(nodes[1], ctx) : 0)
    if (!ctx.arrayMode) {
      const s = isArrayValue(v) || isRefValue(v) ? v : (v as Scalar)
      if (isError(s)) return alt()
      if (isRefValue(s) || isArrayValue(s)) return isError(calc.single(s, ctx)) ? alt() : s
      return s
    }
    const x = calc.operand(v)
    if (!isArrayValue(x)) return isError(x) ? alt() : x
    if (!x.data.some(isError)) return x
    return calc.broadcast([x, calc.operand(alt())], (a, b) => (isError(a) ? b : a))
  },
  IFNA: (calc, nodes, ctx) => {
    if (nodes.length !== 2) return ERR.VALUE
    const v = nodes[0] ? calc.eval(nodes[0], ctx) : null
    const alt = (): Value => (nodes[1] ? calc.eval(nodes[1], ctx) : 0)
    if (!ctx.arrayMode) {
      if (v === ERR.NA) return alt()
      if (isRefValue(v) || isArrayValue(v)) return calc.single(v, ctx) === ERR.NA ? alt() : v
      return v
    }
    const x = calc.operand(v)
    if (!isArrayValue(x)) return x === ERR.NA ? alt() : x
    if (!x.data.includes(ERR.NA)) return x
    return calc.broadcast([x, calc.operand(alt())], (a, b) => (a === ERR.NA ? b : a))
  },
  LET: (calc, nodes, ctx) => {
    if (nodes.length < 3 || nodes.length % 2 === 0) return ERR.VALUE
    const vars = new Map(ctx.vars ?? [])
    let scope: Ctx = { ...ctx, vars }
    for (let i = 0; i + 1 < nodes.length; i += 2) {
      const name = nodes[i]
      if (!name || name.t !== 'name') throw new FormulaLimit('a LET name that is not a name')
      const value = nodes[i + 1] ? calc.eval(nodes[i + 1]!, scope) : null
      vars.set(name.name.toLowerCase(), value)
      scope = { ...ctx, vars }
    }
    const last = nodes[nodes.length - 1]
    return last ? calc.eval(last, scope) : null
  },
  AND: (calc, nodes, ctx) => logical(calc, nodes, ctx, true),
  OR: (calc, nodes, ctx) => logical(calc, nodes, ctx, false),
  ISNUMBER: raw(1, 1, ([v]) => typeof v === 'number'),
  ISERROR: raw(1, 1, ([v]) => isError(v)),
  N: raw(1, 1, ([v]) => (isError(v) ? v : typeof v === 'number' ? v : v === true ? 1 : 0)),

  // ---- text
  TEXT: (calc, nodes, ctx) => {
    if (nodes.length !== 2) return ERR.VALUE
    return calc.scalars(nodes, ctx, (v, f) => {
      const e = firstError([v, f])
      if (e) return e
      const format = toText(f ?? null)
      if (isError(format)) return format
      return formatText(v ?? null, format, calc.options)
    })
  },
  TRIM: plain(1, 1, ([v]) => {
    const s = textArg(v)
    return isError(s) ? s : s.replace(/ +/g, ' ').replace(/^ | $/g, '')
  }),
  LOWER: plain(1, 1, ([v]) => {
    const s = textArg(v)
    return isError(s) ? s : s.toLowerCase()
  }),
  UPPER: plain(1, 1, ([v]) => {
    const s = textArg(v)
    return isError(s) ? s : s.toUpperCase()
  }),
  LEN: plain(1, 1, ([v]) => {
    const s = textArg(v)
    return isError(s) ? s : s.length
  }),
  LEFT: plain(1, 2, ([v, count]) => {
    const s = textArg(v)
    if (isError(s)) return s
    const n = numberArg(count, 1)
    if (isError(n)) return n
    if (n < 0) return ERR.VALUE
    return s.slice(0, Math.trunc(n))
  }),
  MID: plain(3, 3, ([v, start, count]) => {
    const s = textArg(v)
    if (isError(s)) return s
    const from = numberArg(start)
    if (isError(from)) return from
    const n = numberArg(count)
    if (isError(n)) return n
    if (from < 1 || n < 0) return ERR.VALUE
    return s.substr(Math.trunc(from) - 1, Math.trunc(n))
  }),
  SEARCH: plain(2, 3, ([find, within, start]) => {
    const f = textArg(find)
    if (isError(f)) return f
    const w = textArg(within)
    if (isError(w)) return w
    const from = numberArg(start, 1)
    if (isError(from)) return from
    if (from < 1 || from > w.length + 1) return ERR.VALUE
    if (/[*?~]/.test(f)) throw new FormulaLimit(`SEARCH with a wildcard (${JSON.stringify(f)})`)
    const at = w.toLowerCase().indexOf(f.toLowerCase(), Math.trunc(from) - 1)
    return at < 0 ? ERR.VALUE : at + 1
  }),
  CHAR: plain(1, 1, ([v]) => {
    const n = numberArg(v)
    if (isError(n)) return n
    const code = Math.trunc(n)
    if (code === 10 || code === 13 || code === 9 || (code >= 32 && code <= 126))
      return String.fromCharCode(code)
    throw new FormulaLimit(`CHAR(${code})`)
  }),
  TEXTJOIN: (calc, nodes, ctx) => {
    if (nodes.length < 3) return ERR.VALUE
    const delimiter = calc.single(calc.eval(nodes[0] ?? { t: 'str', v: '' }, ctx), ctx)
    const ignore = calc.single(nodes[1] ? calc.eval(nodes[1], ctx) : true, ctx)
    if (isError(delimiter)) return delimiter
    if (isError(ignore)) return ignore
    const d = toText(delimiter) as string
    const skip = toBool(ignore)
    if (isError(skip)) return skip
    const parts: string[] = []
    for (const [v] of flatten(calc, nodes.slice(2), ctx)) {
      if (isError(v)) return v
      const s = toText(v) as string
      if (skip && s === '') continue
      parts.push(s)
    }
    const out = parts.join(d)
    return out.length > 32767 ? ERR.VALUE : out
  },

  // ---- numbers
  SUM: (calc, nodes, ctx) => {
    const ns = numbersOf(calc, nodes, ctx, false)
    if (isError(ns)) return ns
    let total = 0
    for (const n of ns) total += n
    return total
  },
  MIN: (calc, nodes, ctx) => {
    const ns = numbersOf(calc, nodes, ctx, false)
    if (isError(ns)) return ns
    return ns.length === 0 ? 0 : Math.min(...ns)
  },
  MAX: (calc, nodes, ctx) => {
    const ns = numbersOf(calc, nodes, ctx, false)
    if (isError(ns)) return ns
    return ns.length === 0 ? 0 : Math.max(...ns)
  },
  COUNT: (calc, nodes, ctx) => {
    let n = 0
    for (const [v, fromRange] of flatten(calc, nodes, ctx)) {
      if (typeof v === 'number') n++
      else if (
        !fromRange &&
        (typeof v === 'boolean' || (typeof v === 'string' && safeNumeric(v) !== null))
      )
        n++
    }
    return n
  },
  COUNTA: (calc, nodes, ctx) => {
    let n = 0
    for (const [v] of flatten(calc, nodes, ctx)) if (v !== null) n++
    return n
  },
  SUMIF: (calc, nodes, ctx) => {
    if (nodes.length < 2 || nodes.length > 3) return ERR.VALUE
    const range = calc.range(nodes[0] ?? null, ctx)
    const sum = nodes.length === 3 ? calc.range(nodes[2] ?? null, ctx) : range
    return calc.scalars([nodes[1] ?? null], ctx, (c) =>
      conditional(calc, sum, [[range, c ?? null]]),
    )
  },
  SUMIFS: conditionalFn(0, 1),
  COUNTIF: conditionalFn(null, 0),
  COUNTIFS: conditionalFn(null, 0),

  // ---- dates
  TODAY: (calc) => calc.options.today,
  YEAR: plain(1, 1, ([v]) => {
    const n = numberArg(v)
    return isError(n) ? n : n < 0 ? ERR.NUM : dateParts(n).y
  }),
  MONTH: plain(1, 1, ([v]) => {
    const n = numberArg(v)
    return isError(n) ? n : n < 0 ? ERR.NUM : dateParts(n).m
  }),
  WEEKDAY: plain(1, 2, ([v, type]) => {
    const n = numberArg(v)
    if (isError(n)) return n
    const kind = numberArg(type, 1)
    if (isError(kind)) return kind
    if (kind !== 1) throw new FormulaLimit(`WEEKDAY of type ${kind}`)
    return n < 0 ? ERR.NUM : dateParts(n).dow + 1
  }),
  DATE: plain(3, 3, ([y, m, d]) => {
    const year = numberArg(y)
    if (isError(year)) return year
    const month = numberArg(m)
    if (isError(month)) return month
    const day = numberArg(d)
    if (isError(day)) return day
    let yy = Math.trunc(year)
    if (yy < 0 || yy > 9999) return ERR.NUM
    if (yy < 1900) yy += 1900
    const t = new Date(0)
    t.setUTCFullYear(yy, Math.trunc(month) - 1, Math.trunc(day))
    const serial = Math.round((t.getTime() - EPOCH) / 86_400_000)
    if (serial < 61) throw new FormulaLimit('a date before March 1900')
    return serial > 2_958_465 ? ERR.NUM : serial
  }),
  WORKDAY: (calc, nodes, ctx) => {
    if (nodes.length < 2 || nodes.length > 3) return ERR.VALUE
    const holidays = new Set<number>()
    if (nodes[2]) {
      for (const [v] of flatten(calc, [nodes[2]], ctx)) {
        if (isError(v)) return v
        if (typeof v === 'number') holidays.add(Math.floor(v))
        else if (v !== null) {
          const n = toNumber(v)
          if (isError(n)) return n
          holidays.add(Math.floor(n))
        }
      }
    }
    return calc.scalars(nodes.slice(0, 2), ctx, (start, days) => {
      const e = firstError([start, days])
      if (e) return e
      const s = numberArg(start)
      if (isError(s)) return s
      const n = numberArg(days)
      if (isError(n)) return n
      if (s < 0) return ERR.NUM
      if (Math.floor(s) < 61) throw new FormulaLimit('a date before March 1900')
      return workday(s, n, holidays)
    })
  },

  // ---- lookups
  XLOOKUP: (calc, nodes, ctx) => {
    if (nodes.length < 3 || nodes.length > 6) return ERR.VALUE
    for (const i of [4, 5]) {
      if (nodes[i]) {
        const mode = calc.single(calc.eval(nodes[i]!, ctx), ctx)
        if (mode !== (i === 4 ? 0 : 1))
          throw new FormulaLimit('XLOOKUP with a match or search mode')
      }
    }
    const lookIn = calc.range(nodes[1] ?? null, ctx)
    const returns = calc.range(nodes[2] ?? null, ctx)
    if (isError(lookIn)) return lookIn
    if (isError(returns)) return returns
    return calc.scalars([nodes[0] ?? null], ctx, (value) => {
      const at = calc.findExact(lookIn, value ?? null)
      if (isError(at)) return at
      if (at < 0) return nodes[3] ? calc.single(calc.eval(nodes[3], ctx), ctx) : ERR.NA
      return calc.elementAt(returns, at)
    })
  },
  MATCH: (calc, nodes, ctx) => {
    if (nodes.length < 2 || nodes.length > 3) return ERR.VALUE
    const type = nodes[2] ? calc.single(calc.eval(nodes[2], ctx), ctx) : 1
    if (type !== 0) throw new FormulaLimit('MATCH that is not an exact match')
    const lookIn = calc.range(nodes[1] ?? null, ctx)
    if (isError(lookIn)) return lookIn
    return calc.scalars([nodes[0] ?? null], ctx, (value) => {
      if (typeof value === 'string' && /[*?~]/.test(value))
        throw new FormulaLimit('MATCH with a wildcard')
      const at = calc.findExact(lookIn, value ?? null)
      if (isError(at)) return at
      return at < 0 ? ERR.NA : at + 1
    })
  },
  XMATCH: (calc, nodes, ctx) => {
    if (nodes.length !== 2) throw new FormulaLimit('XMATCH with a match or search mode')
    const lookIn = calc.range(nodes[1] ?? null, ctx)
    if (isError(lookIn)) return lookIn
    return calc.scalars([nodes[0] ?? null], ctx, (value) => {
      const at = calc.findExact(lookIn, value ?? null)
      if (isError(at)) return at
      return at < 0 ? ERR.NA : at + 1
    })
  },
  INDEX: (calc, nodes, ctx) => {
    if (nodes.length < 2 || nodes.length > 3) return ERR.VALUE
    const source = calc.range(nodes[0] ?? null, ctx)
    if (isError(source)) return source
    const rows = isRefValue(source)
      ? source.r2 - source.r1 + 1
      : isArrayValue(source)
        ? source.rows
        : 1
    const cols = isRefValue(source)
      ? source.c2 - source.c1 + 1
      : isArrayValue(source)
        ? source.cols
        : 1
    /** the part of the source at a row and column number (0 or omitted: all of them) */
    const part = (
      rowArg: Scalar | undefined,
      colArg: Scalar | undefined,
      wantScalar: boolean,
    ): Value => {
      const e = firstError([rowArg, colArg])
      if (e) return e
      let row = rowArg === undefined ? undefined : toNumber(rowArg)
      let col = colArg === undefined ? undefined : toNumber(colArg)
      if (isError(row)) return row
      if (isError(col)) return col
      // one index into a single row counts columns
      if (nodes.length === 2 && rows === 1 && cols > 1) {
        col = row
        row = 1
      }
      const r = row === undefined ? 0 : Math.trunc(row)
      const c = col === undefined ? 0 : Math.trunc(col)
      if (r < 0 || c < 0 || r > rows || c > cols) return ERR.REF
      if (isRefValue(source)) {
        const ref: RefValue = {
          kind: 'ref',
          sheet: source.sheet,
          r1: r === 0 ? source.r1 : source.r1 + r - 1,
          r2: r === 0 ? source.r2 : source.r1 + r - 1,
          c1: c === 0 ? source.c1 : source.c1 + c - 1,
          c2: c === 0 ? source.c2 : source.c1 + c - 1,
        }
        if (!wantScalar) return ref
        if (ref.r1 !== ref.r2 || ref.c1 !== ref.c2)
          throw new FormulaLimit('INDEX of whole rows or columns by an array of positions')
        return calc.cellValue(ref.sheet, ref.r1, ref.c1)
      }
      const arr = calc.arrayOf(source) as ArrayValue
      const r1 = r === 0 ? 0 : r - 1
      const r2 = r === 0 ? rows - 1 : r - 1
      const c1 = c === 0 ? 0 : c - 1
      const c2 = c === 0 ? cols - 1 : c - 1
      if (r1 === r2 && c1 === c2) return arr.data[r1 * cols + c1] ?? null
      if (wantScalar)
        throw new FormulaLimit('INDEX of whole rows or columns by an array of positions')
      const data: Scalar[] = []
      for (let i = r1; i <= r2; i++)
        for (let j = c1; j <= c2; j++) data.push(arr.data[i * cols + j] ?? null)
      return array(r2 - r1 + 1, c2 - c1 + 1, data, arr.open === true && r === 0)
    }
    const positions = [nodes[1] ?? null, nodes[2] ?? null].map((n) =>
      n ? calc.eval(n, ctx) : undefined,
    )
    if (ctx.arrayMode) {
      const operands = positions.map((v) => (v === undefined ? undefined : calc.operand(v)))
      if (operands.some((v) => v !== undefined && isArrayValue(v))) {
        // positions given as an array: one value for each
        return calc.broadcast(
          operands.map((v) => (v === undefined ? MISSING : v)),
          (x, y) =>
            part(x === MISSING ? undefined : x, y === MISSING ? undefined : y, true) as Scalar,
        )
      }
      return part(operands[0] as Scalar | undefined, operands[1] as Scalar | undefined, false)
    }
    return part(
      positions[0] === undefined ? undefined : calc.single(positions[0], ctx),
      positions[1] === undefined ? undefined : calc.single(positions[1], ctx),
      false,
    )
  },
  CHOOSE: (calc, nodes, ctx) => {
    if (nodes.length < 2) return ERR.VALUE
    if (ctx.arrayMode) {
      const index = calc.operand(nodes[0] ? calc.eval(nodes[0], ctx) : null)
      if (isArrayValue(index)) {
        const choices = nodes.slice(1).map((n) => calc.operand(n ? calc.eval(n, ctx) : 0))
        return calc.broadcast([index, ...choices], (i, ...rest) => {
          if (isError(i)) return i
          const n = toNumber(i)
          if (isError(n)) return n
          const k = Math.trunc(n)
          return k < 1 || k > rest.length ? ERR.VALUE : (rest[k - 1] ?? null)
        })
      }
    }
    const which = calc.single(nodes[0] ? calc.eval(nodes[0], ctx) : null, ctx)
    if (isError(which)) return which
    const n = toNumber(which)
    if (isError(n)) return n
    const i = Math.trunc(n)
    if (i < 1 || i >= nodes.length) return ERR.VALUE
    return nodes[i] ? calc.eval(nodes[i]!, ctx) : 0
  },
  ROWS: (calc, nodes, ctx) => {
    const v = calc.range(nodes[0] ?? null, ctx)
    if (isError(v)) return v
    if (isRefValue(v)) return v.r2 - v.r1 + 1
    if (isArrayValue(v)) {
      if (v.open) throw new FormulaLimit('ROWS of an array without an end')
      return v.rows
    }
    return 1
  },
  COLUMNS: (calc, nodes, ctx) => {
    const v = calc.range(nodes[0] ?? null, ctx)
    if (isError(v)) return v
    return isRefValue(v) ? v.c2 - v.c1 + 1 : isArrayValue(v) ? v.cols : 1
  },

  // ---- dynamic arrays
  SEQUENCE: (calc, nodes, ctx) => {
    const args = nodes.map((n) =>
      n ? calc.single(calc.eval(n, { ...ctx, arrayMode: false }), ctx) : undefined,
    )
    const e = firstError(args)
    if (e) return e
    const [rows, cols, start, step] = [
      numberArg(args[0], 1),
      numberArg(args[1], 1),
      numberArg(args[2], 1),
      numberArg(args[3], 1),
    ]
    if (isError(rows) || isError(cols) || isError(start) || isError(step)) return ERR.VALUE
    const r = Math.trunc(rows)
    const c = Math.trunc(cols)
    if (r < 1 || c < 1) return ERR.CALC
    if (r * c > 5_000_000) throw new FormulaLimit('a very long SEQUENCE')
    const data = new Array<Scalar>(r * c)
    for (let i = 0; i < r * c; i++) data[i] = start + i * step
    return array(r, c, data)
  },
  FILTER: (calc, nodes, ctx) => {
    if (nodes.length < 2 || nodes.length > 3) return ERR.VALUE
    const source = calc.arrayOf(calc.range(nodes[0] ?? null, ctx))
    if (isError(source)) return source
    const include = calc.arrayOf(calc.range(nodes[1] ?? null, ctx))
    if (isError(include)) return include
    if ((source.open || include.open) && source.rows !== include.rows) {
      throw new FormulaLimit('FILTER over whole columns of different sheets')
    }
    const byRow = include.cols === 1 && include.rows === source.rows
    const byColumn = !byRow && include.rows === 1 && include.cols === source.cols
    if (!byRow && !byColumn) return ERR.VALUE
    const keep: number[] = []
    for (let i = 0; i < include.data.length; i++) {
      const v = include.data[i] ?? null
      if (isError(v)) return v
      if (typeof v === 'string') return ERR.VALUE
      if (v === true || (typeof v === 'number' && v !== 0)) keep.push(i)
    }
    if (keep.length === 0) return nodes[2] ? calc.eval(nodes[2], ctx) : ERR.CALC
    const data: Scalar[] = []
    if (byRow) {
      for (const r of keep)
        for (let c = 0; c < source.cols; c++) data.push(source.data[r * source.cols + c] ?? null)
      const open =
        (source.open === true || include.open === true) && keep[keep.length - 1] === source.rows - 1
      return array(keep.length, source.cols, data, open)
    }
    for (let r = 0; r < source.rows; r++)
      for (const c of keep) data.push(source.data[r * source.cols + c] ?? null)
    return array(source.rows, keep.length, data, source.open === true)
  },
  UNIQUE: (calc, nodes, ctx) => {
    if (nodes.length !== 1) throw new FormulaLimit('UNIQUE by column or of values that occur once')
    const source = calc.arrayOf(calc.range(nodes[0] ?? null, ctx))
    if (isError(source)) return source
    const seen = new Set<string>()
    const data: Scalar[] = []
    let rows = 0
    for (let r = 0; r < source.rows; r++) {
      const row = rowOf(source, r)
      const key = row
        .map((v) => (isError(v) ? `e${v.code}` : (matchKey(v) ?? 'blank')))
        .join('\u0001')
      if (seen.has(key)) continue
      seen.add(key)
      data.push(...row)
      rows++
    }
    return array(rows, source.cols, data)
  },
  SORT: (calc, nodes, ctx) => {
    if (nodes.length !== 1) throw new FormulaLimit('SORT with a sort index, order or direction')
    const source = calc.arrayOf(calc.range(nodes[0] ?? null, ctx))
    if (isError(source)) return source
    if (source.open) throw new FormulaLimit('SORT of an array without an end')
    const order = Array.from({ length: source.rows }, (_, i) => i)
    order.sort(
      (a, b) =>
        sortOrder(source.data[a * source.cols] ?? null, source.data[b * source.cols] ?? null) ||
        a - b,
    )
    return array(
      source.rows,
      source.cols,
      order.flatMap((r) => rowOf(source, r)),
    )
  },
  SORTBY: (calc, nodes, ctx) => {
    if (nodes.length < 2) return ERR.VALUE
    const source = calc.arrayOf(calc.range(nodes[0] ?? null, ctx))
    if (isError(source)) return source
    if (source.open) throw new FormulaLimit('SORTBY of an array without an end')
    const keys: { by: ArrayValue; descending: boolean }[] = []
    for (let i = 1; i < nodes.length; i += 2) {
      const by = calc.arrayOf(calc.range(nodes[i] ?? null, ctx))
      if (isError(by)) return by
      if (by.cols !== 1 || by.rows !== source.rows) {
        if (by.rows === 1 && by.cols === source.cols)
          throw new FormulaLimit('SORTBY across columns')
        return ERR.VALUE
      }
      const dir = nodes[i + 1]
        ? calc.single(calc.eval(nodes[i + 1]!, { ...ctx, arrayMode: false }), ctx)
        : 1
      if (dir !== 1 && dir !== -1) return ERR.VALUE
      keys.push({ by, descending: dir === -1 })
    }
    const order = Array.from({ length: source.rows }, (_, i) => i)
    order.sort((a, b) => {
      for (const k of keys) {
        const x = k.by.data[a] ?? null
        const y = k.by.data[b] ?? null
        let c = sortOrder(x, y)
        // blanks stay last in a descending sort
        if (k.descending && x !== null && y !== null) c = -c
        if (c !== 0) return c
      }
      return a - b
    })
    return array(
      source.rows,
      source.cols,
      order.flatMap((r) => rowOf(source, r)),
    )
  },
  CHOOSECOLS: (calc, nodes, ctx) => {
    if (nodes.length < 2) return ERR.VALUE
    const source = calc.arrayOf(calc.range(nodes[0] ?? null, ctx))
    if (isError(source)) return source
    const picks: number[] = []
    for (const n of nodes.slice(1)) {
      const v = calc.arrayOf(calc.range(n, ctx))
      if (isError(v)) return v
      for (const item of v.data) {
        const x = toNumber(item ?? null)
        if (isError(x)) return x
        const i = Math.trunc(x)
        if (i === 0 || Math.abs(i) > source.cols) return ERR.VALUE
        picks.push(i > 0 ? i - 1 : source.cols + i)
      }
    }
    const data: Scalar[] = []
    for (let r = 0; r < source.rows; r++)
      for (const c of picks) data.push(source.data[r * source.cols + c] ?? null)
    return array(source.rows, picks.length, data, source.open === true)
  },
  HSTACK: (calc, nodes, ctx) => stack(calc, nodes, ctx, true),
  VSTACK: (calc, nodes, ctx) => stack(calc, nodes, ctx, false),
}

function stack(
  calc: Calculator,
  nodes: readonly (Node | null)[],
  ctx: Ctx,
  horizontal: boolean,
): Value {
  const parts: ArrayValue[] = []
  for (const n of nodes) {
    const v = calc.arrayOf(calc.range(n, ctx))
    if (isError(v)) return v
    parts.push(v)
  }
  if (parts.length === 0) return ERR.VALUE
  const open = parts.some((p) => p.open)
  if (open && !horizontal) throw new FormulaLimit('VSTACK of an array without an end')
  if (open && !parts.every((p) => p.open && p.rows === parts[0]!.rows)) {
    throw new FormulaLimit('HSTACK of arrays with and without an end')
  }
  if (horizontal) {
    const rows = Math.max(...parts.map((p) => p.rows))
    const cols = parts.reduce((n, p) => n + p.cols, 0)
    const data = new Array<Scalar>(rows * cols)
    let at = 0
    for (const p of parts) {
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < p.cols; c++) {
          data[r * cols + at + c] = r < p.rows ? (p.data[r * p.cols + c] ?? null) : ERR.NA
        }
      }
      at += p.cols
    }
    return array(rows, cols, data, open)
  }
  const cols = Math.max(...parts.map((p) => p.cols))
  const data: Scalar[] = []
  let rows = 0
  for (const p of parts) {
    for (let r = 0; r < p.rows; r++) {
      for (let c = 0; c < cols; c++)
        data.push(c < p.cols ? (p.data[r * p.cols + c] ?? null) : ERR.NA)
    }
    rows += p.rows
  }
  return array(rows, cols, data)
}

/** the names of the functions the calculation knows */
export const KNOWN_FUNCTIONS: readonly string[] = Object.keys(FUNCTIONS)
