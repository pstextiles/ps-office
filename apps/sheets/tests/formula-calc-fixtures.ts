/**
 * Workbooks for the built-in recalculation tests, built from a description
 * of each sheet's cells.
 */
import { zip } from './power-query-fixtures'

export type CellSpec =
  | number
  | string
  | boolean
  /** a formula; `array` is the range of an array formula, `dynamic` marks a spilling one, `stored` its saved value */
  | { f: string; array?: string; dynamic?: boolean; stored?: number | string }
  /** a cell that holds a value spilled by an array formula */
  | { spilled: number | string }

export interface SheetSpec {
  readonly name: string
  readonly cells: Record<string, CellSpec>
  /** a table over a range, with its column names */
  readonly table?: { name: string; ref: string; columns: string[] }
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

function valueXml(v: number | string | boolean | undefined): { t: string; body: string } {
  if (v === undefined) return { t: '', body: '' }
  if (typeof v === 'number') return { t: '', body: `<v>${v}</v>` }
  if (typeof v === 'boolean') return { t: ' t="b"', body: `<v>${v ? 1 : 0}</v>` }
  return { t: ' t="str"', body: `<v>${esc(v)}</v>` }
}

export function workbook(sheets: readonly SheetSpec[]): Buffer {
  const files: Record<string, string> = {
    '[Content_Types].xml':
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    'xl/workbook.xml': `<workbook ${NS} ${R_NS}><sheets>${sheets
      .map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('')}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
      .map(
        (_, i) =>
          `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
      )
      .join('')}</Relationships>`,
  }
  sheets.forEach((sheet, i) => {
    const rows = new Map<number, string[]>()
    const refs = Object.keys(sheet.cells).sort((a, b) => {
      const ra = Number(a.replace(/\D/g, ''))
      const rb = Number(b.replace(/\D/g, ''))
      return (
        ra - rb || a.replace(/\d/g, '').length - b.replace(/\d/g, '').length || a.localeCompare(b)
      )
    })
    for (const ref of refs) {
      const spec = sheet.cells[ref]!
      const row = Number(ref.replace(/\D/g, ''))
      let xml: string
      if (typeof spec === 'object' && 'f' in spec) {
        const v = valueXml(spec.stored)
        const attrs = spec.array ? ` t="array" ref="${spec.array}"` : ''
        xml = `<c r="${ref}"${v.t}${spec.dynamic ? ' cm="1"' : ''}><f${attrs}>${esc(spec.f)}</f>${v.body}</c>`
      } else if (typeof spec === 'object') {
        const v = valueXml(spec.spilled)
        xml = `<c r="${ref}"${v.t}><f ca="1"/>${v.body}</c>`
      } else {
        const v = valueXml(spec)
        xml = `<c r="${ref}"${v.t}>${v.body}</c>`
      }
      rows.set(row, [...(rows.get(row) ?? []), xml])
    }
    files[`xl/worksheets/sheet${i + 1}.xml`] = `<worksheet ${NS} ${R_NS}><sheetData>${[...rows]
      .map(([r, cells]) => `<row r="${r}">${cells.join('')}</row>`)
      .join('')}</sheetData></worksheet>`
    if (sheet.table) {
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] =
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table${i + 1}.xml"/></Relationships>`
      files[`xl/tables/table${i + 1}.xml`] =
        `<table ${NS} id="${i + 1}" name="${sheet.table.name}" displayName="${sheet.table.name}" ref="${sheet.table.ref}"><tableColumns count="${sheet.table.columns.length}">${sheet.table.columns
          .map((c, k) => `<tableColumn id="${k + 1}" name="${esc(c)}"/>`)
          .join('')}</tableColumns></table>`
    }
  })
  return zip(files)
}
