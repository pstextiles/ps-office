/**
 * The standard library of the built-in engine: the Table, List, Text,
 * Record, Number, Date and Comparer functions the company's queries use,
 * and the sources Excel.Workbook, File.Contents and Folder.Contents.
 * Behaviour follows Power Query's documentation; anything outside this set
 * fails with "The name … wasn't recognized" so a query is never silently
 * computed differently.
 */
import { compareValues, isIntegerType, isNumberType, ordinal, valuesEqual } from './compare'
import {
  cultureNamed,
  dateFrom,
  dateTimeFrom,
  numberFrom,
  partsToSerial,
  roundHalfEven,
  textFrom,
  type Culture,
} from './convert'
import { readWorkbookSheets } from './excel-source'
import {
  MDate,
  MDateTime,
  MDuration,
  MError,
  asFunction,
  asList,
  asNumber,
  asRecord,
  asTable,
  asText,
  cellValue,
  describe,
  fieldIndex,
  fn,
  invoke,
  isErrorCell,
  isFunction,
  isList,
  isRecord,
  isTable,
  kindOf,
  list,
  record,
  rowRecord,
  table,
  type,
  type ErrorCell,
  type MBinary,
  type MCell,
  type MTable,
  type MType,
  type MValue,
} from './values'

export interface FolderEntry {
  readonly name: string
  /** folder path ending with a backslash, as Folder.Contents reports it */
  readonly folder: string
  readonly modified: Date
  readonly created: Date
  readonly accessed: Date
  readonly hidden: boolean
}

/** What the engine may reach outside itself: files and folder listings. */
export interface QueryHost {
  readFile(path: string): Uint8Array
  listFolder(path: string): readonly FolderEntry[]
}

const MissingField = { Error: 0, Ignore: 1, UseNull: 2 } as const
const Order = { Ascending: 0, Descending: 1 } as const

function errorCellOf(error: unknown): ErrorCell {
  if (error instanceof MError) return { kind: 'error-cell', error }
  throw error
}

/** runs `compute`; an M error becomes an error cell (Power Query's row-level error) */
function cellOf(compute: () => MValue): MCell {
  try {
    return compute()
  } catch (error) {
    return errorCellOf(error)
  }
}

function columnIndexOrThrow(t: MTable, name: string): number {
  const i = t.columns.indexOf(name)
  if (i < 0)
    throw new MError('Expression.Error', `The column '${name}' of the table wasn't found.`, name)
  return i
}

function textList(value: MValue, what: string): string[] {
  if (typeof value === 'string') return [value]
  return asList(value).items.map((v) => asText(v, what))
}

/** a single pair {a, b} or a list of pairs {{a, b}, …} */
function pairs(value: MValue): MValue[][] {
  const outer = asList(value)
  if (outer.items.length > 0 && !isList(outer.items[0]!)) return [outer.items.slice()]
  return outer.items.map((p) => asList(p).items.slice())
}

function dateTimeFromJs(d: Date): MDateTime {
  // local wall-clock time, as Folder.Contents reports it
  const serial =
    (Date.UTC(
      d.getFullYear(),
      d.getMonth(),
      d.getDate(),
      d.getHours(),
      d.getMinutes(),
      d.getSeconds(),
    ) -
      Date.UTC(1899, 11, 30)) /
    86_400_000
  return new MDateTime(serial)
}

function predicateTrue(f: MValue, arg: MValue): boolean {
  return invoke(f, [arg]) === true
}

/** a comparison criteria: a comparer function, or null for the default order */
function comparerOf(criteria: MValue): (a: MValue, b: MValue) => number {
  if (isFunction(criteria)) return (a, b) => asNumber(invoke(criteria, [a, b]))
  return (a, b) => compareValues(a, b, false)
}

function equalityOf(criteria: MValue): (a: MValue, b: MValue) => boolean {
  if (isFunction(criteria))
    return (a, b) => {
      const r = invoke(criteria, [a, b])
      return typeof r === 'boolean' ? r : r === 0
    }
  return valuesEqual
}

export function buildLibrary(host: QueryHost, culture: Culture): Map<string, MValue> {
  const lib = new Map<string, MValue>()
  const def = (
    name: string,
    min: number,
    max: number,
    call: (args: readonly MValue[]) => MValue,
  ): void => {
    lib.set(name, fn(name, min, max, call))
  }
  /** a culture argument: null is the workbook's culture, text or [Culture=…] names one */
  const cultureOf = (arg: MValue): Culture => {
    if (arg === null || arg === undefined) return culture
    let name: MValue = arg
    if (isRecord(arg)) {
      const i = fieldIndex(arg, 'Culture')
      name = i < 0 ? null : arg.get(i)
    }
    if (name === null || name === undefined) return culture
    if (typeof name !== 'string')
      throw new MError('Expression.Error', 'The culture must be text.', arg)
    return cultureNamed(name)
  }

  // ---- constants and types
  lib.set('MissingField.Error', MissingField.Error)
  lib.set('MissingField.Ignore', MissingField.Ignore)
  lib.set('MissingField.UseNull', MissingField.UseNull)
  lib.set('Order.Ascending', Order.Ascending)
  lib.set('Order.Descending', Order.Descending)
  lib.set('GroupKind.Local', 0)
  lib.set('GroupKind.Global', 1)
  lib.set('JoinKind.Inner', 0)
  lib.set('JoinKind.LeftOuter', 1)
  for (const n of [
    'Int64',
    'Int32',
    'Int16',
    'Int8',
    'Byte',
    'Number',
    'Currency',
    'Percentage',
    'Double',
    'Single',
    'Decimal',
    'Text',
    'Date',
    'DateTime',
    'Logical',
    'Any',
  ]) {
    lib.set(`${n}.Type`, type(`${n}.Type`))
  }

  // ---- comparers
  def('Comparer.Ordinal', 2, 2, ([a, b]) => compareValues(a!, b!, false))
  def('Comparer.OrdinalIgnoreCase', 2, 2, ([a, b]) => {
    if (typeof a === 'string' && typeof b === 'string')
      return ordinal(a.toUpperCase(), b.toUpperCase())
    return compareValues(a!, b!, false)
  })

  // ---- sources
  def('File.Contents', 1, 2, ([path]) => {
    const p = asText(path!)
    let cached: Uint8Array | null = null
    const bin: MBinary = {
      kind: 'binary',
      label: p,
      bytes: () => {
        if (cached) return cached
        try {
          cached = host.readFile(p)
        } catch (error) {
          throw new MError('DataSource.NotFound', `File or Folder: Could not find file '${p}'.`, p)
        }
        return cached
      },
    }
    return bin
  })

  def('Folder.Contents', 1, 2, ([path]) => {
    const p = asText(path!)
    let entries: readonly FolderEntry[]
    try {
      entries = host.listFolder(p)
    } catch {
      throw new MError(
        'DataSource.NotFound',
        `File or Folder: We couldn't find the folder '${p}'.`,
        p,
      )
    }
    const columns = [
      'Content',
      'Name',
      'Extension',
      'Date accessed',
      'Date modified',
      'Date created',
      'Attributes',
      'Folder Path',
    ]
    const rows = entries.map((e): MCell[] => {
      const dot = e.name.lastIndexOf('.')
      const contents = lib.get('File.Contents')!
      return [
        invoke(contents, [e.folder + e.name]),
        e.name,
        dot > 0 ? e.name.slice(dot).toLowerCase() : '',
        dateTimeFromJs(e.accessed),
        dateTimeFromJs(e.modified),
        dateTimeFromJs(e.created),
        record(['Hidden'], [e.hidden]),
        e.folder,
      ]
    })
    return table(columns, rows)
  })

  def('Excel.Workbook', 1, 3, ([workbook, useHeaders]) => {
    if (kindOf(workbook!) !== 'binary')
      throw new MError('Expression.Error', 'Excel.Workbook needs the contents of a file.')
    if (useHeaders === true)
      throw new MError(
        'Expression.Error',
        'Excel.Workbook with useHeaders = true is not supported by the built-in engine.',
      )
    const sheets = readWorkbookSheets((workbook as MBinary).bytes())
    const rows = sheets.map((s): MCell[] => {
      const columns = Array.from({ length: s.width }, (_, i) => `Column${i + 1}`)
      return [s.name, table(columns, s.rows), s.name, 'Sheet', s.hidden]
    })
    return table(['Name', 'Data', 'Item', 'Kind', 'Hidden'], rows)
  })

  // ---- conversions
  def('Text.From', 1, 2, ([v, c]) => textFrom(v!, cultureOf(c!)))
  def('Number.From', 1, 2, ([v, c]) => numberFrom(v!, cultureOf(c!)))
  def('Date.From', 1, 2, ([v, c]) => dateFrom(v!, cultureOf(c!)))
  def('DateTime.From', 1, 2, ([v, c]) => dateTimeFrom(v!, cultureOf(c!)))

  // ---- text
  const nullable = (v: MValue, f: (s: string) => MValue): MValue =>
    v === null ? null : f(asText(v))
  def('Text.Trim', 1, 2, ([t, chars]) =>
    nullable(t!, (s) => {
      if (chars === null || chars === undefined) return s.replace(/^\s+|\s+$/gu, '')
      const set = new Set(textList(chars, 'Text'))
      let a = 0
      let b = s.length
      while (a < b && set.has(s[a]!)) a++
      while (b > a && set.has(s[b - 1]!)) b--
      return s.slice(a, b)
    }),
  )
  def('Text.Lower', 1, 2, ([t]) => nullable(t!, (s) => s.toLowerCase()))
  def('Text.Upper', 1, 2, ([t]) => nullable(t!, (s) => s.toUpperCase()))
  const ignoreCase = (c: MValue): boolean =>
    isFunction(c) && c === lib.get('Comparer.OrdinalIgnoreCase')
  def('Text.Contains', 2, 3, ([t, sub, c]) =>
    nullable(t!, (s) =>
      ignoreCase(c!)
        ? s.toUpperCase().includes(asText(sub!).toUpperCase())
        : s.includes(asText(sub!)),
    ),
  )
  def('Text.StartsWith', 2, 3, ([t, sub, c]) =>
    nullable(t!, (s) =>
      ignoreCase(c!)
        ? s.toUpperCase().startsWith(asText(sub!).toUpperCase())
        : s.startsWith(asText(sub!)),
    ),
  )
  def('Text.EndsWith', 2, 3, ([t, sub, c]) =>
    nullable(t!, (s) =>
      ignoreCase(c!)
        ? s.toUpperCase().endsWith(asText(sub!).toUpperCase())
        : s.endsWith(asText(sub!)),
    ),
  )
  def('Text.Select', 2, 2, ([t, chars]) =>
    nullable(t!, (s) => {
      const set = new Set(textList(chars!, 'Text'))
      return Array.from(s)
        .filter((ch) => set.has(ch))
        .join('')
    }),
  )
  def('Text.Remove', 2, 2, ([t, chars]) =>
    nullable(t!, (s) => {
      const set = new Set(textList(chars!, 'Text'))
      return Array.from(s)
        .filter((ch) => !set.has(ch))
        .join('')
    }),
  )
  def('Text.BetweenDelimiters', 3, 5, ([t, start, end]) =>
    nullable(t!, (s) => {
      const a = s.indexOf(asText(start!))
      if (a < 0) return ''
      const from = a + asText(start!).length
      const b = s.indexOf(asText(end!), from)
      return b < 0 ? s.slice(from) : s.slice(from, b)
    }),
  )
  def('Text.Combine', 1, 2, ([texts, sep]) => {
    const items = asList(texts!)
      .items.filter((v) => v !== null)
      .map((v) => asText(v))
    return items.join(sep === null || sep === undefined ? '' : asText(sep))
  })
  def('Text.Length', 1, 1, ([t]) => nullable(t!, (s) => s.length))
  def('Text.Replace', 3, 3, ([t, old, rep]) =>
    nullable(t!, (s) => s.split(asText(old!)).join(asText(rep!))),
  )

  // ---- records
  def('Record.FieldOrDefault', 2, 3, ([r, name, dflt]) => {
    if (r === null) return dflt ?? null
    const rec = asRecord(r!)
    const i = fieldIndex(rec, asText(name!))
    return i < 0 ? (dflt ?? null) : rec.get(i)
  })
  def('Record.FieldValues', 1, 1, ([r]) => {
    const rec = asRecord(r!)
    return list(rec.names.map((_, i) => rec.get(i)))
  })
  def('Record.FieldNames', 1, 1, ([r]) => list(asRecord(r!).names.slice()))
  def('Record.Field', 2, 2, ([r, name]) => {
    const rec = asRecord(r!)
    const i = fieldIndex(rec, asText(name!))
    if (i < 0)
      throw new MError(
        'Expression.Error',
        `The field '${String(name)}' of the record wasn't found.`,
      )
    return rec.get(i)
  })

  // ---- lists
  def('List.Count', 1, 1, ([l]) => asList(l!).items.length)
  def('List.Transform', 2, 2, ([l, f]) => list(asList(l!).items.map((v) => invoke(f!, [v]))))
  def('List.Select', 2, 2, ([l, f]) => list(asList(l!).items.filter((v) => predicateTrue(f!, v))))
  def('List.RemoveNulls', 1, 1, ([l]) => list(asList(l!).items.filter((v) => v !== null)))
  def('List.Contains', 2, 3, ([l, v, c]) => {
    const eq = equalityOf(c!)
    return asList(l!).items.some((x) => eq(x, v!))
  })
  def('List.PositionOf', 2, 4, ([l, v, _occ, c]) => {
    const eq = equalityOf(c!)
    return asList(l!).items.findIndex((x) => eq(x, v!))
  })
  def('List.Distinct', 1, 2, ([l, c]) => {
    const eq = equalityOf(c!)
    const out: MValue[] = []
    for (const v of asList(l!).items) if (!out.some((x) => eq(x, v))) out.push(v)
    return list(out)
  })
  def('List.Sort', 1, 2, ([l, criteria]) => {
    let cmp = (a: MValue, b: MValue): number => compareValues(a, b, false)
    if (criteria === Order.Descending) cmp = (a, b) => compareValues(b, a, false)
    else if (isFunction(criteria!)) cmp = comparerOf(criteria!)
    return list(asList(l!).items.slice().sort(cmp))
  })
  def('List.First', 1, 2, ([l, dflt]) => {
    const items = asList(l!).items
    return items.length > 0 ? items[0]! : (dflt ?? null)
  })
  def('List.Last', 1, 2, ([l, dflt]) => {
    const items = asList(l!).items
    return items.length > 0 ? items[items.length - 1]! : (dflt ?? null)
  })
  const extreme =
    (sign: 1 | -1) =>
    ([l, dflt, criteria]: readonly MValue[]): MValue => {
      const items = asList(l!).items.filter((v) => v !== null)
      if (items.length === 0) return dflt ?? null
      const cmp = comparerOf(criteria ?? null)
      let best = items[0]!
      for (const v of items.slice(1)) if (sign * cmp(v, best) > 0) best = v
      return best
    }
  def('List.Max', 1, 4, extreme(1))
  def('List.Min', 1, 4, extreme(-1))
  def('List.Sum', 1, 2, ([l]) => {
    const items = asList(l!).items.filter((v) => v !== null)
    if (items.length === 0) return null
    let total = 0
    for (const v of items) total += asNumber(v)
    return total
  })

  // ---- tables
  // ---- date, datetime and duration literals
  const whole = (v: MValue, what: string): number => {
    const n = asNumber(v)
    if (!Number.isInteger(n))
      throw new MError('Expression.Error', `The ${what} must be a whole number.`, v)
    return n
  }
  const serialOf = (y: number, m: number, d: number, hh = 0, mm = 0, ss = 0): number => {
    const serial = partsToSerial(y, m, d, hh, mm, ss)
    if (serial === null)
      throw new MError(
        'Expression.Error',
        'The Date operation failed because the resulting value falls outside the range of allowed values.',
      )
    return serial
  }
  def(
    '#date',
    3,
    3,
    ([y, m, d]) => new MDate(serialOf(whole(y!, 'year'), whole(m!, 'month'), whole(d!, 'day'))),
  )
  def('#datetime', 6, 6, ([y, m, d, hh, mm, ss]) => {
    const sec = asNumber(ss!)
    const base = serialOf(
      whole(y!, 'year'),
      whole(m!, 'month'),
      whole(d!, 'day'),
      whole(hh!, 'hour'),
      whole(mm!, 'minute'),
      0,
    )
    if (sec < 0 || sec >= 60)
      throw new MError('Expression.Error', 'The second must be between 0 and 60.', ss!)
    return new MDateTime(base + sec / 86_400)
  })
  def(
    '#duration',
    4,
    4,
    ([d, h, m, sec]) =>
      new MDuration(
        asNumber(d!) + asNumber(h!) / 24 + asNumber(m!) / 1440 + asNumber(sec!) / 86_400,
      ),
  )

  def('#table', 2, 2, ([cols, rows]) => {
    let columns: string[]
    if (typeof cols === 'number') columns = Array.from({ length: cols }, (_, i) => `Column${i + 1}`)
    else if (kindOf(cols!) === 'type') columns = []
    else columns = textList(cols!, 'Text')
    const body = asList(rows!).items.map((r) => {
      const items = asList(r).items
      if (items.length !== columns.length)
        throw new MError('Expression.Error', 'A row of #table has the wrong number of values.')
      return items.slice() as MCell[]
    })
    return table(columns, body)
  })
  def('Table.RowCount', 1, 1, ([t]) => asTable(t!).rows.length)
  def('Table.ColumnNames', 1, 1, ([t]) => list(asTable(t!).columns.slice()))
  def('Table.ToRows', 1, 1, ([t]) => {
    const tb = asTable(t!)
    return list(tb.rows.map((r) => list(r.map((c) => cellValue(c)))))
  })
  def('Table.SelectRows', 2, 2, ([t, f]) => {
    const tb = asTable(t!)
    return table(
      tb.columns,
      tb.rows.filter((_, r) => predicateTrue(f!, rowRecord(tb, r))),
    )
  })
  const skipOrTake =
    (take: boolean) =>
    ([t, countOrCondition]: readonly MValue[]): MValue => {
      const tb = asTable(t!)
      let n: number
      if (countOrCondition === null || countOrCondition === undefined) n = 1
      else if (isFunction(countOrCondition)) {
        n = 0
        while (n < tb.rows.length && predicateTrue(countOrCondition, rowRecord(tb, n))) n++
      } else n = Math.max(0, Math.floor(asNumber(countOrCondition)))
      return table(tb.columns, take ? tb.rows.slice(0, n) : tb.rows.slice(n))
    }
  def('Table.Skip', 1, 2, skipOrTake(false))
  def('Table.FirstN', 2, 2, skipOrTake(true))
  def('Table.PromoteHeaders', 1, 2, ([t, options]) => {
    const tb = asTable(t!)
    if (tb.rows.length === 0) return tb
    let all = false
    if (options !== null && options !== undefined) {
      const o = asRecord(options)
      const i = fieldIndex(o, 'PromoteAllScalars')
      all = i >= 0 && o.get(i) === true
    }
    const first = tb.rows[0]!
    const names: string[] = []
    first.forEach((cell, i) => {
      let name: string | null = null
      if (!isErrorCell(cell) && cell !== null) {
        if (typeof cell === 'string') name = cell
        else if (
          all &&
          (typeof cell === 'number' ||
            typeof cell === 'boolean' ||
            ['date', 'datetime', 'duration'].includes(kindOf(cell)))
        ) {
          name = textFrom(cell, culture)
        }
      }
      if (name === null || name === '') name = `Column${i + 1}`
      names.push(name)
    })
    // repeated names get _1, _2 … after the first
    const seen = new Map<string, number>()
    const unique = names.map((n) => {
      const count = seen.get(n) ?? 0
      seen.set(n, count + 1)
      if (count === 0) return n
      let k = count
      while (names.includes(`${n}_${k}`) || seen.has(`${n}_${k}`)) k++
      seen.set(`${n}_${k}`, 1)
      return `${n}_${k}`
    })
    return table(unique, tb.rows.slice(1))
  })
  def('Table.SelectColumns', 2, 3, ([t, cols, missing]) => {
    const tb = asTable(t!)
    const wanted = textList(cols!, 'Text')
    const mode = missing ?? MissingField.Error
    const picks: { name: string; index: number }[] = []
    for (const name of wanted) {
      const i = tb.columns.indexOf(name)
      if (i >= 0) picks.push({ name, index: i })
      else if (mode === MissingField.UseNull) picks.push({ name, index: -1 })
      else if (mode === MissingField.Error) columnIndexOrThrow(tb, name)
    }
    return table(
      picks.map((p) => p.name),
      tb.rows.map((r) => picks.map((p) => (p.index < 0 ? null : r[p.index]!))),
    )
  })
  def('Table.RemoveColumns', 2, 3, ([t, cols, missing]) => {
    const tb = asTable(t!)
    const drop = new Set(textList(cols!, 'Text'))
    if ((missing ?? MissingField.Error) === MissingField.Error)
      for (const d of drop) columnIndexOrThrow(tb, d)
    const keep = tb.columns.map((c, i) => ({ c, i })).filter((x) => !drop.has(x.c))
    return table(
      keep.map((k) => k.c),
      tb.rows.map((r) => keep.map((k) => r[k.i]!)),
    )
  })
  def('Table.RenameColumns', 2, 3, ([t, renames, missing]) => {
    const tb = asTable(t!)
    const columns = tb.columns.slice()
    for (const pair of pairs(renames!)) {
      const from = asText(pair[0]!)
      const to = asText(pair[1]!)
      const i = columns.indexOf(from)
      if (i < 0) {
        if ((missing ?? MissingField.Error) === MissingField.Error) columnIndexOrThrow(tb, from)
        continue
      }
      columns[i] = to
    }
    return table(columns, tb.rows)
  })
  def('Table.TransformColumnNames', 2, 3, ([t, f]) => {
    const tb = asTable(t!)
    return table(
      tb.columns.map((c) => asText(invoke(f!, [c]))),
      tb.rows,
    )
  })
  def('Table.AddColumn', 3, 4, ([t, name, f]) => {
    const tb = asTable(t!)
    const col = asText(name!)
    if (tb.columns.includes(col))
      throw new MError('Expression.Error', `A column named '${col}' already exists in the table.`)
    return table(
      [...tb.columns, col],
      tb.rows.map((r, i) => [...r, cellOf(() => invoke(f!, [rowRecord(tb, i)]))]),
    )
  })
  def('Table.TransformColumns', 2, 4, ([t, ops, _dflt, missing]) => {
    const tb = asTable(t!)
    const transforms = new Map<number, MValue>()
    for (const op of pairs(ops!)) {
      const name = asText(op[0]!)
      const i = tb.columns.indexOf(name)
      if (i < 0) {
        if ((missing ?? MissingField.Error) === MissingField.Error) columnIndexOrThrow(tb, name)
        continue
      }
      transforms.set(i, asFunction(op[1]!))
    }
    return table(
      tb.columns,
      tb.rows.map((r) =>
        r.map((cell, i) => {
          const f = transforms.get(i)
          if (!f) return cell
          if (isErrorCell(cell)) return cell
          return cellOf(() => invoke(f, [cell]))
        }),
      ),
    )
  })
  def('Table.TransformColumnTypes', 2, 3, ([t, ops, c]) => {
    const tb = asTable(t!)
    const cult = cultureOf(c ?? null)
    const conversions = new Map<number, (v: MValue) => MValue>()
    for (const op of pairs(ops!)) {
      const i = columnIndexOrThrow(tb, asText(op[0]!))
      const ty = op[1] as MType
      conversions.set(i, converterFor(ty, cult))
    }
    return table(
      tb.columns,
      tb.rows.map((r) =>
        r.map((cell, i) => {
          const conv = conversions.get(i)
          if (!conv || isErrorCell(cell)) return cell
          return cellOf(() => conv(cell))
        }),
      ),
    )
  })
  def('Table.Sort', 2, 2, ([t, criteria]) => {
    const tb = asTable(t!)
    const keys: { index: number; desc: boolean }[] = []
    const add = (c: MValue): void => {
      if (typeof c === 'string') keys.push({ index: columnIndexOrThrow(tb, c), desc: false })
      else {
        const p = asList(c).items
        keys.push({ index: columnIndexOrThrow(tb, asText(p[0]!)), desc: p[1] === Order.Descending })
      }
    }
    if (typeof criteria === 'string') add(criteria)
    else {
      const items = asList(criteria!).items
      if (items.length === 2 && typeof items[0] === 'string' && typeof items[1] === 'number')
        add(criteria!)
      else items.forEach(add)
    }
    const order = tb.rows.map((_, i) => i)
    const valueAt = (r: number, c: number): MValue => cellValue(tb.rows[r]![c] ?? null)
    order.sort((a, b) => {
      for (const k of keys) {
        const d = compareValues(valueAt(a, k.index), valueAt(b, k.index), false)
        if (d !== 0) return k.desc ? -d : d
      }
      return a - b
    })
    return table(
      tb.columns,
      order.map((i) => tb.rows[i]!),
    )
  })
  def('Table.Combine', 1, 2, ([tables]) => {
    const parts = asList(tables!).items.map((v) => asTable(v))
    const columns: string[] = []
    for (const p of parts) for (const c of p.columns) if (!columns.includes(c)) columns.push(c)
    const rows: MCell[][] = []
    for (const p of parts) {
      const map = columns.map((c) => p.columns.indexOf(c))
      for (const r of p.rows) rows.push(map.map((i) => (i < 0 ? null : r[i]!)))
    }
    return table(columns, rows)
  })
  def('Table.Group', 3, 5, ([t, keyCols, aggregates]) => {
    const tb = asTable(t!)
    const keys = textList(keyCols!, 'Text').map((k) => columnIndexOrThrow(tb, k))
    let aggs = asList(aggregates!).items
    // a single aggregate may be given without the outer list
    if (aggs.length > 0 && typeof aggs[0] === 'string') aggs = [aggregates!]
    const groups: { key: MValue[]; rows: (readonly MCell[])[] }[] = []
    // simple keys (text, number, logical, date, null) are found by hash; anything else by equality
    const byHash = new Map<string, (typeof groups)[number]>()
    for (const row of tb.rows) {
      const key = keys.map((k) => cellValue(row[k] ?? null))
      const hash = groupHash(key)
      let g =
        hash !== null
          ? byHash.get(hash)
          : groups.find((x) => x.key.every((v, i) => valuesEqual(v, key[i]!)))
      if (!g) {
        groups.push((g = { key, rows: [] }))
        if (hash !== null) byHash.set(hash, g)
      }
      g.rows.push(row)
    }
    const aggNames = aggs.map((a) => asText(asList(a).items[0]!))
    const aggFns = aggs.map((a) => asFunction(asList(a).items[1]!))
    return table(
      [...keys.map((k) => tb.columns[k]!), ...aggNames],
      groups.map((g) => {
        const sub = table(tb.columns, g.rows)
        return [...g.key, ...aggFns.map((f) => cellOf(() => invoke(f, [sub])))]
      }),
    )
  })

  // ---- misc
  def('Value.Is', 2, 2, ([v, ty]) => kindOf(v!) === (ty as MType).name)

  return lib
}

/** a hash that two group keys share exactly when they are equal, or null for keys that need a full comparison */
function groupHash(key: readonly MValue[]): string | null {
  const parts: string[] = []
  for (const v of key) {
    const k = kindOf(v)
    if (k === 'null') parts.push('n')
    else if (k === 'text') parts.push('t' + (v as string))
    else if (k === 'number') parts.push('#' + String(v))
    else if (k === 'logical') parts.push('b' + String(v))
    else if (k === 'date' || k === 'datetime') parts.push(k[0]! + String((v as MDateTime).serial))
    else return null
  }
  return parts.join('\u0000')
}

/** the value conversion Table.TransformColumnTypes applies for a type */
function converterFor(ty: MType, culture: Culture): (v: MValue) => MValue {
  const name = ty.name
  if (name === 'any' || name === 'Any.Type') return (v) => v
  if (name === 'text' || name === 'Text.Type') return (v) => textFrom(v, culture)
  if (name === 'date' || name === 'Date.Type') return (v) => dateFrom(v, culture)
  if (name === 'datetime' || name === 'DateTime.Type') return (v) => dateTimeFrom(v, culture)
  if (name === 'logical' || name === 'Logical.Type') {
    return (v) => {
      if (v === null || typeof v === 'boolean') return v
      if (typeof v === 'number') return v !== 0
      if (typeof v === 'string') {
        const s = v.trim().toLowerCase()
        if (s === 'true') return true
        if (s === 'false') return false
      }
      throw new MError('DataFormat.Error', `We couldn't convert to Logical.`, v)
    }
  }
  if (isIntegerType(name)) {
    return (v) => {
      const n = numberFrom(v, culture)
      return n === null ? null : roundHalfEven(n)
    }
  }
  if (isNumberType(name)) return (v) => numberFrom(v, culture)
  throw new MError(
    'Expression.Error',
    `The built-in engine cannot convert to type ${name}.`,
    describe(null),
  )
}

export { isRecord, isTable }
