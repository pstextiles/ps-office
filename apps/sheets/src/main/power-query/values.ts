/**
 * Power Query (M) values for the built-in query engine (PS Office).
 *
 * The engine covers the part of M the company's workbooks use: reading other
 * workbooks and folders, and shaping tables. Values are plain data; records
 * hold their fields lazily (a field that is never read is never computed, as
 * in M), tables hold rows of already computed cells, and a cell that failed
 * to compute holds an ErrorCell that raises when it is read.
 */

/** An M error: `reason` is the category Power Query reports (Expression.Error, DataFormat.Error…). */
export class MError extends Error {
  constructor(
    readonly reason: string,
    message: string,
    readonly detail: MValue = null,
  ) {
    super(message)
  }
}

/**
 * The built-in engine met something it cannot compute exactly as Power Query
 * does (a text form whose parsing it has not verified, say). Not an M error:
 * `try` and row-level error cells don't catch it, so the whole evaluation
 * stops instead of producing a result that might differ from Excel's.
 */
export class EngineLimit extends Error {
  constructor(
    readonly what: string,
    readonly input: string | null = null,
  ) {
    super(
      input === null
        ? `Built-in engine limit: ${what}`
        : `Built-in engine limit: ${what} (${JSON.stringify(input)})`,
    )
  }
}

/** A date: whole days since 1899-12-30 (the Excel/OLE serial). */
export class MDate {
  constructor(readonly serial: number) {}
}

/** A datetime: days since 1899-12-30, the fraction being the time of day. */
export class MDateTime {
  constructor(readonly serial: number) {}
}

/** A duration in days (date and datetime differences). */
export class MDuration {
  constructor(readonly days: number) {}
}

export interface MList {
  readonly kind: 'list'
  readonly items: readonly MValue[]
}

/** A record keeps field order; values are computed on first read. */
export interface MRecord {
  readonly kind: 'record'
  readonly names: readonly string[]
  readonly get: (index: number) => MValue
}

export interface MTable {
  readonly kind: 'table'
  readonly columns: readonly string[]
  /** rows[r][c]; a cell may be an ErrorCell */
  readonly rows: readonly (readonly MCell[])[]
}

export interface MFunction {
  readonly kind: 'function'
  readonly name: string
  readonly minArgs: number
  readonly maxArgs: number
  readonly call: (args: readonly MValue[]) => MValue
}

/** File contents, read when first needed. */
export interface MBinary {
  readonly kind: 'binary'
  readonly label: string
  readonly bytes: () => Uint8Array
}

export interface MType {
  readonly kind: 'type'
  /** any, number, text, date, datetime, logical, table, list, record, function, binary, null, int64, … */
  readonly name: string
  readonly nullable: boolean
}

/** A table cell whose value failed to compute (raised when read). */
export interface ErrorCell {
  readonly kind: 'error-cell'
  readonly error: MError
}

export type MValue =
  | null
  | boolean
  | number
  | string
  | MDate
  | MDateTime
  | MDuration
  | MList
  | MRecord
  | MTable
  | MFunction
  | MBinary
  | MType

export type MCell = MValue | ErrorCell

export function isErrorCell(cell: MCell): cell is ErrorCell {
  return (
    typeof cell === 'object' && cell !== null && (cell as { kind?: string }).kind === 'error-cell'
  )
}

/** The value of a cell, raising its error when the cell holds one. */
export function cellValue(cell: MCell): MValue {
  if (isErrorCell(cell)) throw cell.error
  return cell
}

export function list(items: readonly MValue[]): MList {
  return { kind: 'list', items }
}

/** A record from already computed values. */
export function record(names: readonly string[], values: readonly MValue[]): MRecord {
  return { kind: 'record', names, get: (i) => values[i] ?? null }
}

/** A record whose fields are computed on first read and then kept. */
export function lazyRecord(names: readonly string[], compute: (index: number) => MValue): MRecord {
  const done = new Map<number, { ok: true; value: MValue } | { ok: false; error: unknown }>()
  return {
    kind: 'record',
    names,
    get: (i) => {
      let slot = done.get(i)
      if (!slot) {
        try {
          slot = { ok: true, value: compute(i) }
        } catch (error) {
          slot = { ok: false, error }
        }
        done.set(i, slot)
      }
      if (!slot.ok) throw slot.error
      return slot.value
    },
  }
}

/** The record of one table row; reading an error cell raises its error. */
export function rowRecord(table: MTable, row: number): MRecord {
  const cells = table.rows[row]!
  return { kind: 'record', names: table.columns, get: (i) => cellValue(cells[i] ?? null) }
}

export function table(columns: readonly string[], rows: readonly (readonly MCell[])[]): MTable {
  return { kind: 'table', columns, rows }
}

export function fn(
  name: string,
  minArgs: number,
  maxArgs: number,
  call: (args: readonly MValue[]) => MValue,
): MFunction {
  return { kind: 'function', name, minArgs, maxArgs, call }
}

export function type(name: string, nullable = false): MType {
  return { kind: 'type', name, nullable }
}

export function kindOf(value: MValue): string {
  if (value === null) return 'null'
  if (typeof value === 'boolean') return 'logical'
  if (typeof value === 'number') return 'number'
  if (typeof value === 'string') return 'text'
  if (value instanceof MDate) return 'date'
  if (value instanceof MDateTime) return 'datetime'
  if (value instanceof MDuration) return 'duration'
  return value.kind
}

export function isList(value: MValue): value is MList {
  return kindOf(value) === 'list'
}
export function isRecord(value: MValue): value is MRecord {
  return kindOf(value) === 'record'
}
export function isTable(value: MValue): value is MTable {
  return kindOf(value) === 'table'
}
export function isFunction(value: MValue): value is MFunction {
  return kindOf(value) === 'function'
}

export function fieldIndex(rec: MRecord, name: string): number {
  return rec.names.indexOf(name)
}

/** A short description of a value for error messages. */
export function describe(value: MValue): string {
  const k = kindOf(value)
  if (k === 'text') return JSON.stringify(value)
  if (k === 'number' || k === 'logical' || k === 'null') return String(value)
  return k
}

export function conversionError(expected: string, value: MValue): MError {
  return new MError(
    'Expression.Error',
    `We cannot convert the value ${describe(value)} to type ${expected}.`,
    value,
  )
}

export function asText(value: MValue, what = 'Text'): string {
  if (typeof value !== 'string') throw conversionError(what, value)
  return value
}
export function asNumber(value: MValue): number {
  if (typeof value !== 'number') throw conversionError('Number', value)
  return value
}
export function asList(value: MValue): MList {
  if (!isList(value)) throw conversionError('List', value)
  return value
}
export function asRecord(value: MValue): MRecord {
  if (!isRecord(value)) throw conversionError('Record', value)
  return value
}
export function asTable(value: MValue): MTable {
  if (!isTable(value)) throw conversionError('Table', value)
  return value
}
export function asFunction(value: MValue): MFunction {
  if (!isFunction(value)) throw conversionError('Function', value)
  return value
}

/** Calls an M function value, checking the argument count. */
export function invoke(f: MValue, args: readonly MValue[]): MValue {
  const g = asFunction(f)
  if (args.length < g.minArgs || args.length > g.maxArgs) {
    throw new MError(
      'Expression.Error',
      `${g.name || 'The function'} expects ${g.minArgs === g.maxArgs ? g.minArgs : `${g.minArgs} to ${g.maxArgs}`} arguments but got ${args.length}.`,
    )
  }
  const padded =
    args.length < g.maxArgs ? [...args, ...Array(g.maxArgs - args.length).fill(null)] : args
  return g.call(padded)
}
