/**
 * Conversions between M values (Text.From, Number.From, Date.From and the
 * column types of Table.TransformColumnTypes). Text forms depend on the
 * query culture, as in Power Query; the culture is chosen by the caller.
 *
 * Every rule here was checked against Excel's own Power Query results in the
 * four supported cultures. Power Query parses text with .NET, which accepts
 * far more shapes than these (it reads "1,234.5" as a date in the year 234);
 * text outside the checked shapes raises EngineLimit, so the engine stops
 * rather than give an answer that may differ from Excel's.
 */
import {
  EngineLimit,
  MDate,
  MDateTime,
  MDuration,
  MError,
  describe,
  kindOf,
  type MValue,
} from './values'

export interface Culture {
  /** 'en-US', 'es-GT', 'ko-KR', 'es-ES' */
  readonly name: string
  /** short numeric dates with the year last: day before month (d/M/yyyy) */
  readonly dayFirst: boolean
  readonly decimal: string
  readonly group: string
  /** short numeric dates with a two-digit year read year first (yy/M/d) */
  readonly yearFirstShort: boolean
  readonly dateText: (y: number, m: number, d: number) => string
  readonly timeText: (hh: number, mm: number, ss: number) => string
}

const pad = (n: number, w = 2): string => String(n).padStart(w, '0')
const h12 = (hh: number): number => (hh % 12 === 0 ? 12 : hh % 12)

export const CULTURES: Readonly<Record<string, Culture>> = {
  'en-US': {
    name: 'en-US',
    dayFirst: false,
    decimal: '.',
    group: ',',
    yearFirstShort: false,
    dateText: (y, m, d) => `${m}/${d}/${y}`,
    timeText: (hh, mm, ss) => `${h12(hh)}:${pad(mm)}:${pad(ss)} ${hh < 12 ? 'AM' : 'PM'}`,
  },
  'es-GT': {
    name: 'es-GT',
    dayFirst: true,
    decimal: '.',
    group: ',',
    yearFirstShort: false,
    dateText: (y, m, d) => `${d}/${pad(m)}/${y}`,
    timeText: (hh, mm, ss) => `${pad(hh)}:${pad(mm)}:${pad(ss)}`,
  },
  'es-ES': {
    name: 'es-ES',
    dayFirst: true,
    decimal: ',',
    group: '.',
    yearFirstShort: false,
    dateText: (y, m, d) => `${pad(d)}/${pad(m)}/${y}`,
    timeText: (hh, mm, ss) => `${hh}:${pad(mm)}:${pad(ss)}`,
  },
  'ko-KR': {
    name: 'ko-KR',
    dayFirst: false,
    decimal: '.',
    group: ',',
    yearFirstShort: true,
    dateText: (y, m, d) => `${y}-${pad(m)}-${pad(d)}`,
    timeText: (hh, mm, ss) => `${hh < 12 ? '오전' : '오후'} ${h12(hh)}:${pad(mm)}:${pad(ss)}`,
  },
}

/**
 * The culture of that name. Only the cultures whose text forms the engine
 * knows are accepted; any other raises an error rather than read text in a
 * guessed culture.
 */
export function cultureNamed(name: string): Culture {
  const hit = Object.values(CULTURES).find((c) => c.name.toLowerCase() === name.toLowerCase())
  if (!hit) throw new EngineLimit('culture not supported', name)
  return hit
}

const MS_PER_DAY = 86_400_000
/** 1899-12-30 as a UTC timestamp */
const EPOCH = Date.UTC(1899, 11, 30)

export function serialToParts(serial: number): {
  y: number
  m: number
  d: number
  hh: number
  mm: number
  ss: number
} {
  const ms = Math.round(serial * MS_PER_DAY)
  const t = new Date(EPOCH + ms)
  return {
    y: t.getUTCFullYear(),
    m: t.getUTCMonth() + 1,
    d: t.getUTCDate(),
    hh: t.getUTCHours(),
    mm: t.getUTCMinutes(),
    ss: t.getUTCSeconds(),
  }
}

export function partsToSerial(
  y: number,
  m: number,
  d: number,
  hh = 0,
  mm = 0,
  ss = 0,
): number | null {
  if (y < 1 || y > 9999) return null
  if (m < 1 || m > 12 || d < 1 || d > 31 || hh > 23 || mm > 59 || ss > 59) return null
  // setUTCFullYear, not Date.UTC: Date.UTC reads years 0–99 as 1900–1999
  const check = new Date(0)
  check.setUTCFullYear(y, m - 1, d)
  check.setUTCHours(hh, mm, ss, 0)
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== m - 1 || check.getUTCDate() !== d)
    return null
  return (check.getTime() - EPOCH) / MS_PER_DAY
}

/**
 * Number to text as Power Query writes it: the shortest of 15 or 17
 * significant digits that reads back as the same number, written out in full
 * up to 10^22 and in E notation from 10^-5 down ("1.234E-06").
 */
export function numberText(n: number, culture: Culture): string {
  if (Number.isNaN(n)) return 'NaN'
  if (!Number.isFinite(n)) return n > 0 ? 'Infinity' : '-Infinity'
  if (n === 0) return '0'
  let exp = n.toExponential(14)
  if (Number(exp) !== n) exp = n.toExponential(16)
  const m = /^(-?)(\d)(?:\.(\d+))?e([+-]\d+)$/.exec(exp)!
  const sign = m[1]!
  const digits = (m[2]! + (m[3] ?? '')).replace(/0+$/, '') || '0'
  const e = Number(m[4])
  const dec = culture.decimal
  if (e <= -5) {
    const mantissa = digits.length > 1 ? `${digits[0]}${dec}${digits.slice(1)}` : digits
    return `${sign}${mantissa}E-${pad(-e)}`
  }
  if (e >= 23) throw new EngineLimit('number text beyond 10^22', String(n))
  if (e < 0) return `${sign}0${dec}${'0'.repeat(-e - 1)}${digits}`
  if (digits.length <= e + 1) return `${sign}${digits}${'0'.repeat(e + 1 - digits.length)}`
  return `${sign}${digits.slice(0, e + 1)}${dec}${digits.slice(e + 1)}`
}

function dateTextOf(serial: number, culture: Culture): string {
  const { y, m, d } = serialToParts(serial)
  return culture.dateText(y, m, d)
}

export function textFrom(value: MValue, culture: Culture): string | null {
  if (value === null) return null
  if (typeof value === 'string') return value
  if (typeof value === 'number') return numberText(value, culture)
  if (typeof value === 'boolean') return value ? 'true' : 'false'
  if (value instanceof MDate) return dateTextOf(value.serial, culture)
  if (value instanceof MDateTime) {
    const { hh, mm, ss } = serialToParts(value.serial)
    return `${dateTextOf(value.serial, culture)} ${culture.timeText(hh, mm, ss)}`
  }
  if (value instanceof MDuration) {
    const total = Math.round(Math.abs(value.days) * 86400)
    const days = Math.floor(total / 86400)
    const rest = total % 86400
    return `${value.days < 0 ? '-' : ''}${days}.${pad(Math.floor(rest / 3600))}:${pad(Math.floor((rest % 3600) / 60))}:${pad(rest % 60)}`
  }
  throw new MError(
    'Expression.Error',
    `We cannot convert a value of type ${kindOf(value)} to type Text.`,
    value,
  )
}

function formatError(value: MValue, target: string): MError {
  return new MError('DataFormat.Error', `We couldn't convert to ${target}.`, value)
}

const ASCII_DIGIT = /[0-9]/
const OTHER_DIGIT = /\p{Nd}/u
const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/**
 * Parses culture text as a number. Blank text is null; text with no digits
 * is a format error; the checked shapes are a sign (+5, 5-, (5)), group
 * separators anywhere before the decimal separator, an exponent, and a
 * percent sign before or after.
 */
export function parseNumberText(text: string, culture: Culture): number | null {
  const s = text.trim()
  if (s === '') return null
  if (!ASCII_DIGIT.test(s)) {
    if (OTHER_DIGIT.test(s) || /^nan$/i.test(s)) throw new EngineLimit('number text', text)
    throw formatError(text, 'Number')
  }
  const g = escapeRe(culture.group)
  const d = escapeRe(culture.decimal)
  const body = `((?:\\d+(?:${g}\\d+)*)?)(?:${d}(\\d*))?((?:[eE][+-]?\\d+)?)`
  const shapes: {
    re: RegExp
    negative: (m: RegExpExecArray) => boolean
    percent: boolean
    at: number
  }[] = [
    { re: new RegExp(`^([+-]?)${body}$`), negative: (m) => m[1] === '-', percent: false, at: 2 },
    { re: new RegExp(`^${body}-$`), negative: () => true, percent: false, at: 1 },
    { re: new RegExp(`^\\(${body}\\)$`), negative: () => true, percent: false, at: 1 },
    { re: new RegExp(`^${body} ?%$`), negative: () => false, percent: true, at: 1 },
    { re: new RegExp(`^%${body}$`), negative: () => false, percent: true, at: 1 },
  ]
  for (const shape of shapes) {
    const m = shape.re.exec(s)
    if (!m) continue
    const int = (m[shape.at] ?? '').split(culture.group).join('')
    const frac = m[shape.at + 1]
    const exp = m[shape.at + 2] ?? ''
    if (int === '' && !frac) continue
    let n = Number(`${int || '0'}.${frac ?? ''}0${exp}`)
    if (shape.percent) n /= 100
    if (shape.negative(m)) n = -n
    return n === 0 ? 0 : n
  }
  // date and time text ("05/11/2026", "2026-11-05", "3:04") is never a number
  if (/[/:]/.test(s) || /\d-\d/.test(s)) throw formatError(text, 'Number')
  throw new EngineLimit('number text', text)
}

export function numberFrom(value: MValue, culture: Culture): number | null {
  if (value === null) return null
  if (typeof value === 'number') return value
  if (typeof value === 'boolean') return value ? 1 : 0
  if (value instanceof MDate || value instanceof MDateTime) return value.serial
  if (value instanceof MDuration) return value.days
  if (typeof value === 'string') return parseNumberText(value, culture)
  throw new MError(
    'Expression.Error',
    `We cannot convert the value ${describe(value)} to type Number.`,
    value,
  )
}

// ---- date text

const ENGLISH_MONTHS = [
  'january',
  'february',
  'march',
  'april',
  'may',
  'june',
  'july',
  'august',
  'september',
  'october',
  'november',
  'december',
]
const SPANISH_MONTHS = [
  'enero',
  'febrero',
  'marzo',
  'abril',
  'mayo',
  'junio',
  'julio',
  'agosto',
  'septiembre',
  'octubre',
  'noviembre',
  'diciembre',
]
const SPANISH_ABBR = [
  'ene',
  'feb',
  'mar',
  'abr',
  'may',
  'jun',
  'jul',
  'ago',
  'sep',
  'oct',
  'nov',
  'dic',
]

type MonthName = { month: number; kind: 'english' | 'spanish-full' | 'spanish-abbr' }

function monthName(word: string): MonthName | null {
  const w = word.toLowerCase()
  const en = ENGLISH_MONTHS.findIndex((full) => w === full || w === full.slice(0, 3))
  if (en >= 0) return { month: en + 1, kind: 'english' }
  const full = SPANISH_MONTHS.indexOf(w)
  if (full >= 0) return { month: full + 1, kind: 'spanish-full' }
  const abbr = SPANISH_ABBR.indexOf(w)
  if (abbr >= 0) return { month: abbr + 1, kind: 'spanish-abbr' }
  return null
}

/** a date part followed by an optional time part, as text forms */
const TIME = String.raw`(\d{1,2}):(\d{2})(?::(\d{2}))?`

interface ParsedDateTime {
  readonly serial: number
  readonly hasTime: boolean
}

/** hh:mm[:ss] with AM/PM after it, or 오전/오후 before it (ko-KR) */
function timeOf(text: string, culture: Culture, input: string): number {
  let s = text.trim()
  let meridiem: 'am' | 'pm' | null = null
  const ko = /^(오전|오후)\s*(.*)$/.exec(s)
  if (ko) {
    if (culture.name !== 'ko-KR') throw formatError(input, 'DateTime')
    meridiem = ko[1] === '오전' ? 'am' : 'pm'
    s = ko[2]!
  } else {
    const ap = /^(.*\S)\s+([AaPp][Mm])$/.exec(s)
    if (ap) {
      meridiem = ap[2]!.toLowerCase() === 'am' ? 'am' : 'pm'
      s = ap[1]!
    }
  }
  const m = new RegExp(`^${TIME}$`).exec(s)
  if (!m) throw new EngineLimit('time text', input)
  let hh = Number(m[1])
  const mm = Number(m[2])
  const ss = Number(m[3] ?? 0)
  if (mm > 59 || ss > 59) throw new EngineLimit('time text', input)
  if (meridiem) {
    if (hh < 1 || hh > 12) throw new EngineLimit('time text', input)
    if (hh === 12) hh = meridiem === 'am' ? 0 : 12
    else if (meridiem === 'pm') hh += 12
  } else if (hh > 23) throw new EngineLimit('time text', input)
  return (hh * 3600 + mm * 60 + ss) / 86_400
}

function dayOf(
  input: string,
  y: number,
  m: number,
  d: number,
  time: string | undefined,
  culture: Culture,
): ParsedDateTime {
  const day = partsToSerial(y, m, d)
  if (day === null) throw formatError(input, 'Date')
  if (time === undefined) return { serial: day, hasTime: false }
  return { serial: day + timeOf(time, culture, input), hasTime: true }
}

/** a two-digit year: 00–26 were checked to mean 2000–2026 */
function shortYear(yy: number, input: string): number {
  if (yy > 26) throw new EngineLimit('two-digit year', input)
  return 2000 + yy
}

/**
 * Parses date or datetime text. Blank text is null; text with no digits is
 * a format error; the checked shapes are listed below, anything else raises
 * EngineLimit.
 */
export function parseDateTimeText(text: string, culture: Culture): ParsedDateTime | null {
  const s = text.trim()
  if (s === '') return null
  if (!ASCII_DIGIT.test(s)) {
    if (OTHER_DIGIT.test(s)) throw new EngineLimit('date text', text)
    throw formatError(text, 'Date')
  }
  let m: RegExpExecArray | null
  // yyyy-M-d [time], yyyy/M/d, yyyy.M.d
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:(?:T|\s+)(.+))?$/.exec(s)))
    return dayOf(text, +m[1]!, +m[2]!, +m[3]!, m[4], culture)
  if ((m = /^(\d{4})([/.])(\d{1,2})\2(\d{1,2})$/.exec(s)))
    return dayOf(text, +m[1]!, +m[3]!, +m[4]!, undefined, culture)
  // yyyy-M and M/yyyy: the first of the month
  if ((m = /^(\d{4})-(\d{1,2})$/.exec(s))) return dayOf(text, +m[1]!, +m[2]!, 1, undefined, culture)
  if ((m = /^(\d{1,2})\/(\d{4})$/.exec(s)))
    return dayOf(text, +m[2]!, +m[1]!, 1, undefined, culture)
  // a/b/yyyy [time] and a.b.yyyy, a-b-yyyy: month/day order of the culture
  if (
    (m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(.+))?$/.exec(s)) ||
    (m = /^(\d{1,2})([.-])(\d{1,2})\2(\d{4})$/.exec(s))
  ) {
    const dotted = m.length === 5 && (m[2] === '.' || m[2] === '-')
    const a = +m[1]!
    const b = dotted ? +m[3]! : +m[2]!
    const y = dotted ? +m[4]! : +m[3]!
    const time = dotted ? undefined : m[4]
    return culture.dayFirst
      ? dayOf(text, y, b, a, time, culture)
      : dayOf(text, y, a, b, time, culture)
  }
  // a/b/yy and a-b-yy: d/M/yy, M/d/yy or yy/M/d by culture
  if ((m = /^(\d{1,2})([/-])(\d{1,2})\2(\d{2})$/.exec(s))) {
    const a = +m[1]!
    const b = +m[3]!
    const c = +m[4]!
    if (culture.yearFirstShort) return dayOf(text, shortYear(a, text), b, c, undefined, culture)
    const y = shortYear(c, text)
    return culture.dayFirst
      ? dayOf(text, y, b, a, undefined, culture)
      : dayOf(text, y, a, b, undefined, culture)
  }
  // 2026년 11월 5일
  if ((m = /^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일$/.exec(s)))
    return dayOf(text, +m[1]!, +m[2]!, +m[3]!, undefined, culture)
  // month names with a four-digit year
  const named =
    (m = /^([A-Za-zñÑ]+)\s+(\d{1,2}),?\s+(\d{4})$/.exec(s)) !== null
      ? { word: m[1]!, d: +m[2]!, y: +m[3]!, form: 'mdy' as const }
      : (m = /^(\d{1,2})\s+([A-Za-zñÑ]+)\s+(\d{4})$/.exec(s)) !== null
        ? { word: m[2]!, d: +m[1]!, y: +m[3]!, form: 'd m y' as const }
        : (m = /^(\d{1,2})([-/])([A-Za-zñÑ]+)\2(\d{4})$/.exec(s)) !== null
          ? {
              word: m[3]!,
              d: +m[1]!,
              y: +m[4]!,
              form: m[2] === '-' ? ('d-m-y' as const) : ('d/m/y' as const),
            }
          : null
  if (named) {
    const month = monthName(named.word)
    if (!month) throw new EngineLimit('date text', text)
    const spanish = culture.name === 'es-GT' || culture.name === 'es-ES'
    if (month.kind === 'english') {
      if (named.form === 'd/m/y') throw new EngineLimit('date text', text)
    } else if (!spanish) {
      throw formatError(text, 'Date')
    } else if (culture.name === 'es-GT') {
      if (month.kind === 'spanish-abbr' && named.form === 'd m y')
        throw new EngineLimit('date text', text)
      if (month.kind === 'spanish-full' && (named.form === 'd-m-y' || named.form === 'd/m/y'))
        throw new EngineLimit('date text', text)
    } else {
      // es-ES reads the full Spanish names; its abbreviations were not checked
      if (month.kind === 'spanish-abbr' || named.form === 'd-m-y' || named.form === 'd/m/y')
        throw new EngineLimit('date text', text)
    }
    return dayOf(text, named.y, month.month, named.d, undefined, culture)
  }
  throw new EngineLimit('date text', text)
}

export function dateFrom(value: MValue, culture: Culture): MDate | null {
  if (value === null) return null
  if (value instanceof MDate) return value
  if (value instanceof MDateTime) return new MDate(Math.floor(value.serial))
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value < -657435 || value > 2958465)
      throw formatError(value, 'Date')
    return new MDate(Math.floor(value))
  }
  if (typeof value === 'string') {
    const parsed = parseDateTimeText(value, culture)
    if (parsed === null) return null
    // Date.From refuses text that carries a time of day
    if (parsed.hasTime) throw formatError(value, 'Date')
    return new MDate(parsed.serial)
  }
  throw new MError(
    'Expression.Error',
    `We cannot convert the value ${describe(value)} to type Date.`,
    value,
  )
}

export function dateTimeFrom(value: MValue, culture: Culture): MDateTime | null {
  if (value === null) return null
  if (value instanceof MDateTime) return value
  if (value instanceof MDate) return new MDateTime(value.serial)
  if (typeof value === 'number') return new MDateTime(value)
  if (typeof value === 'string') {
    const parsed = parseDateTimeText(value, culture)
    return parsed === null ? null : new MDateTime(parsed.serial)
  }
  throw new MError(
    'Expression.Error',
    `We cannot convert the value ${describe(value)} to type DateTime.`,
    value,
  )
}

/** Power Query rounds to even when converting to Int64.Type */
export function roundHalfEven(n: number): number {
  const r = Math.round(n)
  return Math.abs(n % 1) === 0.5 && r % 2 !== 0 ? r - 1 : r
}
