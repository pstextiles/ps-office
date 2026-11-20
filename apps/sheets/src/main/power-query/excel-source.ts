/**
 * Excel.Workbook for the built-in query engine: the sheets of an .xlsx/.xlsm
 * as tables of typed cells (number, text, logical, datetime, error), the
 * way Power Query reads a workbook with untyped (delayed) columns.
 *
 * A formula cell gives the value Excel last calculated and stored. A cell
 * formatted as a date gives a datetime. The table spans the sheet's used
 * range: from the first to the last row and column holding a value.
 */
import { MDateTime, MError, type ErrorCell, type MCell } from './values'
import { readZip } from './zip-read'

export interface SheetContents {
  readonly name: string
  readonly hidden: boolean
  /** rows of the used range; every row has `width` cells */
  readonly rows: readonly (readonly MCell[])[]
  readonly width: number
}

const BUILTIN_DATE_FORMATS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
  52, 53, 54, 55, 56, 57, 58,
])

/** a custom number format is a date or time when a date token survives literals and brackets */
export function isDateFormatCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/_./g, '')
    .replace(/\*./g, '')
  // only the first (positive) section decides
  const first = stripped.split(';')[0] ?? ''
  return /[dmyhs]/i.test(first) && !/^[#0.,%\s]*$/.test(first)
}

function decodeXml(text: string): string {
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
  const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attrs)
  return m ? decodeXml(m[1]!) : null
}

/** text of an <si> or <is>: every <t>, skipping phonetic runs */
function richText(xml: string): string {
  const withoutPhonetic = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '')
  let out = ''
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(withoutPhonetic)) !== null) out += decodeXml(m[1] ?? '')
  return out
}

function columnIndex(ref: string): number {
  let n = 0
  for (const ch of ref) {
    const c = ch.charCodeAt(0)
    if (c < 65 || c > 90) break
    n = n * 26 + (c - 64)
  }
  return n - 1
}

function resolveTarget(target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = ('xl/' + target).split('/')
  const out: string[] = []
  for (const p of parts) {
    if (p === '..') out.pop()
    else if (p !== '.' && p !== '') out.push(p)
  }
  return out.join('/')
}

function errorCell(text: string): ErrorCell {
  return {
    kind: 'error-cell',
    error: new MError('DataFormat.Error', `Invalid cell value '${text}'.`, text),
  }
}

export function readWorkbookSheets(bytes: Uint8Array): SheetContents[] {
  const zip = readZip(bytes)
  const text = (path: string): string | null => zip.text(path)
  const workbookXml = text('xl/workbook.xml')
  if (!workbookXml)
    throw new MError('DataFormat.Error', 'External table is not in the expected format.')
  const date1904 = /<workbookPr\b[^>]*\bdate1904="(?:1|true)"/.test(workbookXml)

  // relationships: r:id → part path
  const relsXml = text('xl/_rels/workbook.xml.rels') ?? ''
  const targets = new Map<string, string>()
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = attr(m[1]!, 'Id')
    const target = attr(m[1]!, 'Target')
    if (id && target) targets.set(id, resolveTarget(target))
  }

  const shared: string[] = []
  const sstXml = text('xl/sharedStrings.xml')
  if (sstXml)
    for (const m of sstXml.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g))
      shared.push(richText(m[1] ?? ''))

  // cell styles (cellXfs) whose number format is a date or time
  const dateStyles = new Set<number>()
  const stylesXml = text('xl/styles.xml')
  if (stylesXml) {
    const custom = new Map<number, string>()
    for (const m of stylesXml.matchAll(/<numFmt\b([^>]*)\/?>/g)) {
      const id = Number(attr(m[1]!, 'numFmtId'))
      const code = attr(m[1]!, 'formatCode')
      if (code !== null) custom.set(id, code)
    }
    const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? ''
    let index = 0
    for (const m of xfs.matchAll(/<xf\b([^>]*?)(?:\/>|>)/g)) {
      const id = Number(attr(m[1]!, 'numFmtId') ?? 0)
      const code = custom.get(id)
      if (code !== undefined ? isDateFormatCode(code) : BUILTIN_DATE_FORMATS.has(id))
        dateStyles.add(index)
      index++
    }
  }

  const sheets: SheetContents[] = []
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const attrs = m[1]!
    const name = attr(attrs, 'name')
    const rid = attr(attrs, 'r:id')
    if (!name || !rid) continue
    const path = targets.get(rid)
    const xml = path ? text(path) : null
    if (xml === null) continue
    const state = attr(attrs, 'state')
    sheets.push({
      name,
      hidden: state === 'hidden' || state === 'veryHidden',
      ...parseSheet(xml, shared, dateStyles, date1904),
    })
  }
  return sheets
}

function parseSheet(
  xml: string,
  shared: readonly string[],
  dateStyles: ReadonlySet<number>,
  date1904: boolean,
): { rows: MCell[][]; width: number } {
  const cells = new Map<number, Map<number, MCell>>()
  let minRow = Infinity
  let maxRow = -1
  let minCol = Infinity
  let maxCol = -1
  const data = /<sheetData\b[^>]*>([\s\S]*)<\/sheetData>/.exec(xml)?.[1] ?? ''
  let rowNumber = -1
  for (const rowMatch of data.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const r = attr(rowMatch[1]!, 'r')
    rowNumber = r !== null ? Number(r) - 1 : rowNumber + 1
    const body = rowMatch[2]
    if (!body) continue
    let colNumber = -1
    for (const c of body.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const cAttrs = c[1]!
      const ref = attr(cAttrs, 'r')
      colNumber = ref !== null ? columnIndex(ref) : colNumber + 1
      const inner = c[2] ?? ''
      const t = attr(cAttrs, 't')
      const s = Number(attr(cAttrs, 's') ?? 0)
      const v = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1]
      let value: MCell | undefined
      if (t === 'inlineStr') {
        const is = /<is>([\s\S]*?)<\/is>/.exec(inner)?.[1]
        if (is !== undefined) value = richText(is)
      } else if (t === 'str' && v === undefined && inner.includes('<v/>')) {
        // a formula whose result is "" stores <v/>; Power Query reads it as empty text
        value = ''
      } else if (v !== undefined) {
        const raw = decodeXml(v)
        if (t === 's') value = shared[Number(raw)] ?? ''
        else if (t === 'str') value = raw
        else if (t === 'b') value = raw === '1' || raw === 'true'
        else if (t === 'e') value = errorCell(raw)
        else if (t === 'd') {
          const ms = Date.parse(raw)
          value = Number.isNaN(ms) ? raw : new MDateTime((ms - Date.UTC(1899, 11, 30)) / 86_400_000)
        } else {
          const n = Number(raw)
          if (dateStyles.has(s) && Number.isFinite(n)) {
            value = new MDateTime(date1904 ? n + 1462 : n)
          } else value = n
        }
      }
      // empty text is a value (""), as Power Query reads it
      if (value === undefined) continue
      let row = cells.get(rowNumber)
      if (!row) cells.set(rowNumber, (row = new Map()))
      row.set(colNumber, value)
      if (rowNumber < minRow) minRow = rowNumber
      if (rowNumber > maxRow) maxRow = rowNumber
      if (colNumber < minCol) minCol = colNumber
      if (colNumber > maxCol) maxCol = colNumber
    }
  }
  if (maxRow < 0) return { rows: [], width: 0 }
  const width = maxCol - minCol + 1
  const rows: MCell[][] = []
  for (let r = minRow; r <= maxRow; r++) {
    const src = cells.get(r)
    const out: MCell[] = new Array(width).fill(null)
    if (src) for (const [c, value] of src) out[c - minCol] = value
    rows.push(out)
  }
  return { rows, width }
}
