/**
 * The workbook as the built-in recalculation sees it: every sheet's cells
 * with their stored values and formulas (shared formulas resolved to their
 * master, array formulas with their range), and the tables structured
 * references point at.
 */
import { readZip } from '../power-query/zip-read'
import { FormulaLimit, MAX_COLUMN } from './parser'

/** an Excel error value; one instance per code, so `===` compares them */
export class XlError {
  private static readonly all = new Map<string, XlError>()
  private constructor(readonly code: string) {}
  static of(code: string): XlError {
    let hit = XlError.all.get(code)
    if (!hit) XlError.all.set(code, (hit = new XlError(code)))
    return hit
  }
}
export const ERR = {
  NA: XlError.of('#N/A'),
  VALUE: XlError.of('#VALUE!'),
  REF: XlError.of('#REF!'),
  DIV0: XlError.of('#DIV/0!'),
  NUM: XlError.of('#NUM!'),
  NAME: XlError.of('#NAME?'),
  NULL: XlError.of('#NULL!'),
  SPILL: XlError.of('#SPILL!'),
  CALC: XlError.of('#CALC!'),
} as const

/** a cell value; null is an empty cell */
export type Scalar = number | string | boolean | null | XlError

export interface FormulaModel {
  readonly text: string
  /** the cell the text was written for: a shared formula's master, else the cell itself */
  readonly baseRow: number
  readonly baseColumn: number
  /** the range of an array formula (this cell is its top-left), else null */
  readonly array: { r1: number; c1: number; r2: number; c2: number } | null
  /** a dynamic array formula (spills to the size of its result), or one flagged for array evaluation */
  readonly dynamic: boolean
}

export interface CellModel {
  /** the value stored in the file */
  readonly stored: Scalar
  readonly formula: FormulaModel | null
  /** the cell held a value spilled from an array formula when the file was saved */
  inStoredSpill: boolean
  /**
   * The stored value carries value metadata (vm): an error newer than the
   * file format (#SPILL!, #CALC!) is stored as #VALUE! with it.
   */
  readonly rich: boolean
}

export interface TableModel {
  readonly name: string
  readonly sheet: SheetModel
  /** 0-based, inclusive, header and totals rows included; a pending resize moves it */
  r1: number
  c1: number
  r2: number
  c2: number
  readonly headerRows: number
  readonly totalsRows: number
  readonly columns: readonly string[]
}

export interface SheetModel {
  readonly name: string
  readonly index: number
  /** key = row * 16384 + column */
  readonly cells: Map<number, CellModel>
  /** the last row and column holding a cell */
  maxRow: number
  maxColumn: number
  /** keys of the array formula cells */
  readonly arrayMasters: number[]
}

export interface WorkbookModel {
  readonly sheets: readonly SheetModel[]
  /** by lower-cased name */
  readonly sheetByName: ReadonlyMap<string, SheetModel>
  /** by lower-cased name */
  readonly tables: ReadonlyMap<string, TableModel>
  readonly date1904: boolean
}

export const cellKey = (row: number, column: number): number => row * (MAX_COLUMN + 1) + column
export const keyRow = (key: number): number => Math.floor(key / (MAX_COLUMN + 1))
export const keyColumn = (key: number): number => key % (MAX_COLUMN + 1)

function decodeXml(raw: string): string {
  // an XML parser reads a literal CR LF (or CR) in content as LF
  const text = raw.includes('\r') ? raw.replace(/\r\n?/g, '\n') : raw
  if (!text.includes('&') && !text.includes('_x')) return text
  return text
    .replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (_, e: string) => {
      switch (e) {
        case 'lt':
          return '<'
        case 'gt':
          return '>'
        case 'amp':
          return '&'
        case 'quot':
          return '"'
        case 'apos':
          return "'"
        default:
          return String.fromCodePoint(
            e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10),
          )
      }
    })
    .replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
}

function attr(attrs: string, name: string): string | null {
  const at = attrs.indexOf(` ${name}="`)
  if (at < 0) return null
  const start = at + name.length + 3
  return decodeXml(attrs.slice(start, attrs.indexOf('"', start)))
}

function columnOf(ref: string): number {
  let n = 0
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i)
    if (c < 65 || c > 90) break
    n = n * 26 + (c - 64)
  }
  return n - 1
}
const rowOf = (ref: string): number => Number(ref.replace(/^[A-Z]+/, '')) - 1

function parseArea(ref: string): { r1: number; c1: number; r2: number; c2: number } {
  const [a, b] = ref.split(':')
  return { r1: rowOf(a!), c1: columnOf(a!), r2: rowOf(b ?? a!), c2: columnOf(b ?? a!) }
}

function richText(xml: string): string {
  const withoutPhonetic = xml.includes('<rPh') ? xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '') : xml
  let out = ''
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(withoutPhonetic)) !== null) out += decodeXml(m[1]!)
  return out
}

function resolveTarget(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = fromPart.split('/').slice(0, -1).concat(target.split('/'))
  const out: string[] = []
  for (const p of parts) {
    if (p === '..') out.pop()
    else if (p !== '.' && p !== '') out.push(p)
  }
  return out.join('/')
}

function relationships(
  read: (path: string) => string | null,
  part: string,
): { id: string; type: string; target: string }[] {
  const slash = part.lastIndexOf('/')
  const xml = read(`${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`)
  if (!xml) return []
  const out: { id: string; type: string; target: string }[] = []
  for (const m of xml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    const id = attr(m[1]!, 'Id')
    const target = attr(m[1]!, 'Target')
    if (id && target && attr(m[1]!, 'TargetMode') !== 'External') {
      out.push({ id, type: attr(m[1]!, 'Type') ?? '', target: resolveTarget(part, target) })
    }
  }
  return out
}

/** Reads the workbook; raises FormulaLimit for what the calculation does not handle (the 1904 date system). */
export function readWorkbookModel(bytes: Uint8Array): WorkbookModel {
  const zip = readZip(bytes)
  const read = (path: string): string | null => zip.text(path)
  const workbookXml = read('xl/workbook.xml')
  if (!workbookXml) throw new FormulaLimit('a file that is not a workbook')
  const date1904 = /<workbookPr\b[^>]*\bdate1904="(?:1|true)"/.test(workbookXml)
  if (date1904) throw new FormulaLimit('the 1904 date system')
  const calcPr = /<calcPr\b([^>]*)\/?>/.exec(workbookXml)?.[1] ?? ''
  if (/\biterate="(?:1|true)"/.test(calcPr)) throw new FormulaLimit('iterative calculation')
  if (/\bfullPrecision="(?:0|false)"/.test(calcPr)) throw new FormulaLimit('precision as displayed')

  const shared: string[] = []
  const sstXml = read('xl/sharedStrings.xml')
  if (sstXml)
    for (const m of sstXml.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g))
      shared.push(richText(m[1] ?? ''))

  const workbookRels = new Map(relationships(read, 'xl/workbook.xml').map((r) => [r.id, r.target]))
  const sheets: SheetModel[] = []
  const tables = new Map<string, TableModel>()
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
    const name = attr(m[1]!, 'name')
    const path = workbookRels.get(attr(m[1]!, 'r:id') ?? '')
    if (!name || !path) continue
    const xml = read(path)
    if (xml === null) continue
    const sheet = readSheet(name, sheets.length, xml, shared)
    sheets.push(sheet)
    for (const rel of relationships(read, path)) {
      if (!/\/table$/.test(rel.type)) continue
      const tableXml = read(rel.target)
      const open = tableXml ? /<table\b([^>]*)>/.exec(tableXml)?.[1] : undefined
      if (!tableXml || !open) continue
      const tableName = attr(open, 'displayName') ?? attr(open, 'name')
      const ref = attr(open, 'ref')
      if (!tableName || !ref) continue
      const columns: string[] = []
      for (const c of tableXml.matchAll(/<tableColumn\b([^>]*?)(?:\/>|>)/g))
        columns.push(attr(c[1]!, 'name') ?? '')
      tables.set(tableName.toLowerCase(), {
        name: tableName,
        sheet,
        ...parseArea(ref),
        headerRows: Number(attr(open, 'headerRowCount') ?? 1),
        totalsRows: Number(attr(open, 'totalsRowCount') ?? 0),
        columns,
      })
    }
  }
  return {
    sheets,
    sheetByName: new Map(sheets.map((s) => [s.name.toLowerCase(), s])),
    tables,
    date1904,
  }
}

function readSheet(
  name: string,
  index: number,
  xml: string,
  shared: readonly string[],
): SheetModel {
  const cells = new Map<number, CellModel>()
  const arrayMasters: number[] = []
  let maxRow = -1
  let maxColumn = -1
  const masters = new Map<string, { text: string; row: number; column: number }>()
  const followers: { key: number; si: string; stored: Scalar }[] = []
  const dataStart = xml.indexOf('<sheetData')
  const dataEnd = xml.lastIndexOf('</sheetData>')
  const data = dataStart < 0 || dataEnd < 0 ? '' : xml.slice(dataStart, dataEnd)
  const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
  let m: RegExpExecArray | null
  while ((m = cellRe.exec(data)) !== null) {
    const attrs = m[1]!
    const ref = attr(attrs, 'r')
    if (!ref) throw new FormulaLimit(`sheet '${name}' has cells without an address`)
    const row = rowOf(ref)
    const column = columnOf(ref)
    const inner = m[2] ?? ''
    const t = attr(attrs, 't')
    let stored: Scalar = null
    const vAt = inner.indexOf('<v>')
    // <v> can carry xml:space; the plain form is the common, fast case
    const raw =
      vAt >= 0
        ? inner.slice(vAt + 3, inner.indexOf('</v>', vAt))
        : inner.includes('<v ')
          ? (/<v\s[^>]*>([\s\S]*?)<\/v>/.exec(inner)?.[1] ?? null)
          : null
    if (t === 'inlineStr') stored = richText(inner)
    else if (t === 'str')
      stored = raw === null ? (inner.includes('<v/>') ? '' : null) : decodeXml(raw)
    else if (raw !== null) {
      if (t === 's') stored = shared[Number(raw)] ?? ''
      else if (t === 'b') stored = raw === '1'
      else if (t === 'e') stored = XlError.of(decodeXml(raw))
      else stored = Number(raw)
    }
    let formula: FormulaModel | null = null
    const fAt = inner.indexOf('<f')
    if (fAt >= 0) {
      const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(inner.slice(fAt))!
      const fAttrs = fm[1]!
      const text = decodeXml(fm[2] ?? '')
      const ft = attr(fAttrs, 't')
      const key = cellKey(row, column)
      if (ft === 'shared') {
        const si = attr(fAttrs, 'si') ?? ''
        if (text !== '') {
          masters.set(si, { text, row, column })
          formula = {
            text,
            baseRow: row,
            baseColumn: column,
            array: null,
            dynamic: false,
          }
        } else {
          followers.push({ key, si, stored })
          continue
        }
      } else if (ft === 'array') {
        const area = parseArea(attr(fAttrs, 'ref') ?? ref)
        formula = {
          text,
          baseRow: row,
          baseColumn: column,
          array: area,
          dynamic: attr(attrs, 'cm') !== null,
        }
        arrayMasters.push(key)
      } else if (ft === 'dataTable') {
        throw new FormulaLimit(`sheet '${name}' has a data table`)
      } else if (text === '') {
        // <f ca="1"/>: a cell filled by a volatile array formula, not a formula of its own
        formula = null
      } else {
        formula = { text, baseRow: row, baseColumn: column, array: null, dynamic: false }
      }
    }
    if (stored === null && formula === null) continue
    cells.set(cellKey(row, column), {
      stored,
      formula,
      inStoredSpill: false,
      rich: attr(attrs, 'vm') !== null,
    })
    if (row > maxRow) maxRow = row
    if (column > maxColumn) maxColumn = column
  }
  for (const f of followers) {
    const master = masters.get(f.si)
    if (!master) throw new FormulaLimit(`sheet '${name}' has a shared formula without its master`)
    const row = keyRow(f.key)
    const column = keyColumn(f.key)
    cells.set(f.key, {
      stored: f.stored,
      formula: {
        text: master.text,
        baseRow: master.row,
        baseColumn: master.column,
        array: null,
        dynamic: false,
      },
      inStoredSpill: false,
      rich: false,
    })
    if (row > maxRow) maxRow = row
    if (column > maxColumn) maxColumn = column
  }
  // the cells an array formula filled when the file was saved
  for (const key of arrayMasters) {
    const area = cells.get(key)!.formula!.array!
    for (let r = area.r1; r <= area.r2; r++) {
      for (let c = area.c1; c <= area.c2; c++) {
        const k = cellKey(r, c)
        if (k === key) continue
        const cell = cells.get(k)
        if (cell) cell.inStoredSpill = true
      }
    }
  }
  return { name, index, cells, maxRow, maxColumn, arrayMasters }
}
