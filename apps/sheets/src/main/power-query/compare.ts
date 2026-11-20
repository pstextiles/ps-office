/** Equality, ordering and type checks with M's rules. */
import type { TypeRef } from './parser'
import {
  MDate,
  MDateTime,
  MDuration,
  MError,
  describe,
  isList,
  isRecord,
  isTable,
  kindOf,
  type MValue,
} from './values'

/** M equality: values of different kinds are simply unequal (no error). */
export function valuesEqual(a: MValue, b: MValue): boolean {
  if (a === null || b === null) return a === b
  const ka = kindOf(a)
  const kb = kindOf(b)
  if ((ka === 'date' || ka === 'datetime') && (kb === 'date' || kb === 'datetime')) {
    return ka === kb && (a as MDate).serial === (b as MDate).serial
  }
  if (ka !== kb) return false
  if (ka === 'number' || ka === 'text' || ka === 'logical') return a === b
  if (a instanceof MDuration) return a.days === (b as MDuration).days
  if (isList(a) && isList(b)) {
    return a.items.length === b.items.length && a.items.every((x, i) => valuesEqual(x, b.items[i]!))
  }
  if (isRecord(a) && isRecord(b)) {
    if (a.names.length !== b.names.length) return false
    return a.names.every((n, i) => {
      const j = b.names.indexOf(n)
      return j >= 0 && valuesEqual(a.get(i), b.get(j))
    })
  }
  if (isTable(a) && isTable(b)) {
    return (
      a.columns.join('\u0000') === b.columns.join('\u0000') &&
      a.rows.length === b.rows.length &&
      a.rows.every((r, i) => r.every((c, j) => valuesEqual(c as MValue, b.rows[i]![j] as MValue)))
    )
  }
  return a === b
}

/** ordinal (code unit) comparison, as Comparer.Ordinal */
export function ordinal(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0
}

/**
 * Orders two values of the same kind. `strict` raises on different kinds
 * (operators); sorting passes false and orders by kind name instead.
 */
export function compareValues(a: MValue, b: MValue, strict: boolean): number {
  if (a === null || b === null) {
    if (a === b) return 0
    return a === null ? -1 : 1 // null sorts first
  }
  const ka = kindOf(a)
  const kb = kindOf(b)
  if (typeof a === 'number' && typeof b === 'number') return a < b ? -1 : a > b ? 1 : 0
  if (typeof a === 'string' && typeof b === 'string') return ordinal(a, b)
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1
  if (
    (a instanceof MDate || a instanceof MDateTime) &&
    (b instanceof MDate || b instanceof MDateTime) &&
    ka === kb
  ) {
    return a.serial < b.serial ? -1 : a.serial > b.serial ? 1 : 0
  }
  if (a instanceof MDuration && b instanceof MDuration)
    return a.days < b.days ? -1 : a.days > b.days ? 1 : 0
  if (strict) {
    throw new MError('Expression.Error', `We cannot compare ${describe(a)} and ${describe(b)}.`)
  }
  return ordinal(ka, kb)
}

const INT_TYPES = new Set(['Int64.Type', 'Int32.Type', 'Int16.Type', 'Int8.Type', 'Byte.Type'])
const NUMBER_TYPES = new Set([
  'number',
  'Number.Type',
  'Currency.Type',
  'Percentage.Type',
  'Double.Type',
  'Single.Type',
  'Decimal.Type',
  ...INT_TYPES,
])

/** `x is T` / `x as T` for primitive types (records, lists and tables loosely) */
export function conformsTo(value: MValue, t: TypeRef): boolean {
  if (value === null) return t.nullable || t.name === 'null' || t.name === 'any'
  const k = kindOf(value)
  if (t.name === 'any') return true
  if (t.name === 'anynonnull') return true
  if (NUMBER_TYPES.has(t.name)) return k === 'number'
  if (t.name === 'Text.Type') return k === 'text'
  if (t.name === 'Date.Type') return k === 'date'
  if (t.name === 'DateTime.Type') return k === 'datetime'
  return k === t.name
}

export function isIntegerType(name: string): boolean {
  return INT_TYPES.has(name)
}
export function isNumberType(name: string): boolean {
  return NUMBER_TYPES.has(name)
}
