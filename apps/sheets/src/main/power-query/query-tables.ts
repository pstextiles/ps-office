/**
 * Phase 2 of the built-in engine: writing query results into the workbook's
 * query tables (tables with tableType="queryTable"), as Excel's refresh does.
 *
 * Only what is checked to match Excel is written: the table keeps its
 * columns and grows or shrinks to the new row count, new rows take the
 * styles of the row above and the formulas of calculated columns, and rows
 * that are no longer needed are emptied. Anything else (a totals row, a
 * filter or sort, columns that changed, cells in the way below the table)
 * raises EngineLimit, and the refresh is left to Excel.
 */
import type { CellEdit, SheetTableEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  EngineLimit,
  MDate,
  MDateTime,
  isErrorCell,
  kindOf,
  type MCell,
  type MTable,
} from './values'

/** reads a part of the workbook package as text; null when it is missing */
export type PartText = (path: string) => string | null

export interface QueryTableColumn {
  readonly name: string
  /** the query column it is loaded from; null for a column of the table's own (unbound) */
  readonly field: string | null
  /** a calculated column's formula, for unbound columns */
  readonly formula: string | null
}

export interface SortCondition {
  /** 0-based sheet column */
  readonly column: number
  readonly descending: boolean
  /** 1-based last row of the condition's ref */
  readonly endRow: number
}

export interface QueryTableTarget {
  /** the query (connection Location) that loads this table */
  readonly query: string
  readonly sheetName: string
  readonly worksheetPath: string
  readonly tablePath: string
  /** displayName of the table part */
  readonly tableName: string
  /** 0-based, inclusive, header row included */
  readonly area: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  }
  readonly columns: readonly QueryTableColumn[]
  /** the table's sort, which Excel applies again after a refresh */
  readonly sort: readonly SortCondition[]
  /** the hidden sheet-local name (ExternalData_N) that holds the table's range, when the workbook has one */
  readonly localName: { readonly name: string; readonly sheetIndex: number } | null
  /** reasons Excel's refresh would do more than the engine writes */
  readonly unsupported: readonly string[]
}

function decodeXml(text: string): string {
  return text.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (_, e: string) => {
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
}

function attr(attrs: string, name: string): string | null {
  const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attrs)
  return m ? decodeXml(m[1]!) : null
}

function resolvePart(fromPart: string, target: string): string {
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
  read: PartText,
  part: string,
): { id: string; type: string; target: string }[] {
  const slash = part.lastIndexOf('/')
  const relsPath = `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`
  const xml = read(relsPath)
  if (!xml) return []
  const out: { id: string; type: string; target: string }[] = []
  for (const m of xml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    const id = attr(m[1]!, 'Id')
    const target = attr(m[1]!, 'Target')
    if (!id || !target || attr(m[1]!, 'TargetMode') === 'External') continue
    out.push({ id, type: attr(m[1]!, 'Type') ?? '', target: resolvePart(part, target) })
  }
  return out
}

function columnIndex(letters: string): number {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

function parseRef(ref: string): QueryTableTarget['area'] | null {
  const m = /^\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref)
  if (!m) return null
  return {
    startColumn: columnIndex(m[1]!),
    startRow: Number(m[2]) - 1,
    endColumn: columnIndex(m[3]!),
    endRow: Number(m[4]) - 1,
  }
}

/** the query names of connections.xml, by connection id (Location= of a Power Query connection) */
function connectionQueries(read: PartText): Map<string, string> {
  const out = new Map<string, string>()
  const xml = read('xl/connections.xml') ?? ''
  for (const m of xml.matchAll(/<connection\b([^>]*)>([\s\S]*?)<\/connection>/g)) {
    const id = attr(m[1]!, 'id')
    const db = /<dbPr\b([^>]*)\/?>/.exec(m[2]!)
    const connection = db ? attr(db[1]!, 'connection') : null
    if (!id || !connection || !/Provider=Microsoft\.Mashup\.OleDb/i.test(connection)) continue
    const loc = /(?:^|;)\s*Location=(?:"((?:[^"]|"")*)"|([^;]*))/i.exec(connection)
    const name = loc
      ? loc[1] !== undefined
        ? loc[1].replace(/""/g, '"')
        : (loc[2] ?? '').trim()
      : ''
    if (name) out.set(id, name)
  }
  return out
}

/** Every query table of the workbook: the tables a Power Query connection loads into. */
export function findQueryTables(read: PartText): QueryTableTarget[] {
  const queries = connectionQueries(read)
  const workbookXml = read('xl/workbook.xml') ?? ''
  const workbookRels = new Map(relationships(read, 'xl/workbook.xml').map((r) => [r.id, r.target]))
  const out: QueryTableTarget[] = []
  let sheetIndex = -1
  for (const sheet of workbookXml.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
    sheetIndex++
    const sheetName = attr(sheet[1]!, 'name')
    const rid = attr(sheet[1]!, 'r:id')
    const worksheetPath = rid ? workbookRels.get(rid) : undefined
    if (!sheetName || !worksheetPath) continue
    for (const rel of relationships(read, worksheetPath)) {
      if (!/\/table$/.test(rel.type)) continue
      const tableXml = read(rel.target)
      if (!tableXml) continue
      const open = /<table\b([^>]*)>/.exec(tableXml)?.[1] ?? ''
      if (attr(open, 'tableType') !== 'queryTable') continue
      const queryTablePart = relationships(read, rel.target).find((r) =>
        /\/queryTable$/.test(r.type),
      )?.target
      const queryTableXml = queryTablePart ? read(queryTablePart) : null
      if (!queryTableXml) continue
      const qtOpen = /<queryTable\b([^>]*)>/.exec(queryTableXml)?.[1] ?? ''
      const query = queries.get(attr(qtOpen, 'connectionId') ?? '')
      const qtName = attr(qtOpen, 'name')
      const hasLocalName =
        qtName !== null &&
        [...workbookXml.matchAll(/<definedName\b([^>]*)>/g)].some(
          (d) =>
            attr(d[1]!, 'name') === qtName && attr(d[1]!, 'localSheetId') === String(sheetIndex),
        )
      const area = parseRef(attr(open, 'ref') ?? '')
      if (!query || !area) continue

      const unsupported: string[] = []
      if (Number(attr(open, 'totalsRowCount') ?? 0) > 0)
        unsupported.push('the table has a totals row')
      if (Number(attr(open, 'headerRowCount') ?? 1) !== 1)
        unsupported.push('the table has no header row')
      if (/<filterColumn\b/.test(tableXml)) unsupported.push('the table is filtered')
      const sort: SortCondition[] = []
      const sortState = /<sortState\b([^>]*)>([\s\S]*?)<\/sortState>|<sortState\b([^>]*)\/>/.exec(
        tableXml,
      )
      if (sortState) {
        const stateAttrs = sortState[1] ?? sortState[3] ?? ''
        if (attr(stateAttrs, 'columnSort') === '1')
          unsupported.push('the table is sorted left to right')
        if (attr(stateAttrs, 'caseSensitive') === '1')
          unsupported.push('the table is sorted case-sensitively')
        for (const c of (sortState[2] ?? '').matchAll(/<sortCondition\b([^>]*?)\/?>/g)) {
          const a = c[1]!
          const ref = /^\$?([A-Z]{1,3})\$?\d+:\$?[A-Z]{1,3}\$?(\d+)$/.exec(attr(a, 'ref') ?? '')
          const sortBy = attr(a, 'sortBy')
          if (!ref || (sortBy !== null && sortBy !== 'value') || attr(a, 'customList') !== null) {
            unsupported.push('the table is sorted by a custom list, a colour or an icon')
            continue
          }
          sort.push({
            column: columnIndex(ref[1]!),
            descending: attr(a, 'descending') === '1',
            endRow: Number(ref[2]),
          })
        }
      }

      const fields = new Map<string, { name: string | null; bound: boolean }>()
      for (const f of queryTableXml.matchAll(/<queryTableField\b([^>]*?)\/?>/g)) {
        const tableColumnId = attr(f[1]!, 'tableColumnId')
        if (!tableColumnId) continue
        fields.set(tableColumnId, {
          name: attr(f[1]!, 'name'),
          bound: attr(f[1]!, 'dataBound') !== '0',
        })
      }
      const columns: QueryTableColumn[] = []
      for (const c of tableXml.matchAll(
        /<tableColumn\b([^>]*?)(?:\/>|>([\s\S]*?)<\/tableColumn>)/g,
      )) {
        const id = attr(c[1]!, 'id') ?? ''
        const name = attr(c[1]!, 'name') ?? ''
        const field = fields.get(id)
        const formulaXml =
          /<calculatedColumnFormula\b([^>]*)>([\s\S]*?)<\/calculatedColumnFormula>/.exec(c[2] ?? '')
        if (formulaXml && attr(formulaXml[1]!, 'array') === '1')
          unsupported.push(`column '${name}' has an array formula`)
        const formula = formulaXml ? decodeXml(formulaXml[2]!) : null
        if (field?.bound) columns.push({ name, field: field.name ?? name, formula: null })
        else {
          if (formula === null)
            unsupported.push(`column '${name}' holds values of its own, not from the query`)
          columns.push({ name, field: null, formula })
        }
      }
      if (columns.length !== area.endColumn - area.startColumn + 1)
        unsupported.push('the table columns do not match its range')
      out.push({
        query,
        sheetName,
        worksheetPath,
        tablePath: rel.target,
        tableName: attr(open, 'displayName') ?? attr(open, 'name') ?? '',
        area,
        columns,
        sort,
        localName: hasLocalName ? { name: qtName!, sheetIndex } : null,
        unsupported,
      })
    }
  }
  return out
}

interface SheetCell {
  readonly style: number | null
  /** the cell holds a value or a formula */
  readonly filled: boolean
  /** the master of a shared formula */
  readonly sharedMaster: boolean
}

/** the cells of rows [fromRow, toRow] (0-based) in columns [fromColumn, toColumn] */
function sheetCells(
  worksheetXml: string,
  fromRow: number,
  toRow: number,
  fromColumn: number,
  toColumn: number,
): Map<number, Map<number, SheetCell>> {
  const out = new Map<number, Map<number, SheetCell>>()
  const data = /<sheetData\b[^>]*>([\s\S]*)<\/sheetData>/.exec(worksheetXml)?.[1] ?? ''
  for (const row of data.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const r = Number(attr(row[1]!, 'r')) - 1
    if (!(r >= fromRow && r <= toRow) || !row[2]) continue
    for (const c of row[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = attr(c[1]!, 'r') ?? ''
      const col = columnIndex(/^[A-Z]+/.exec(ref)?.[0] ?? '')
      if (col < fromColumn || col > toColumn) continue
      const inner = c[2] ?? ''
      const s = attr(c[1]!, 's')
      const f = /<f\b([^>]*)/.exec(inner)
      let cells = out.get(r)
      if (!cells) out.set(r, (cells = new Map()))
      cells.set(col, {
        style: s === null ? null : Number(s),
        filled: /<f\b|<v\b[^>]*>[\s\S]*?<\/v>|<v\/>|<is>/.test(inner),
        sharedMaster: f !== null && attr(f[1]!, 't') === 'shared' && attr(f[1]!, 'ref') !== null,
      })
    }
  }
  return out
}

/** a query result cell as Excel writes it into a sheet cell */
function sheetValue(cell: MCell, query: string): string | number | boolean | null {
  if (cell === null) return null
  // Excel loads an error value as an empty cell (checked: two '#N/A' source cells)
  if (isErrorCell(cell)) return null
  if (typeof cell === 'string') {
    if (cell.length > 32767)
      throw new EngineLimit(`query '${query}' has text longer than a cell holds`)
    return cell
  }
  if (typeof cell === 'number') {
    if (Number.isNaN(cell)) return null
    if (!Number.isFinite(cell)) throw new EngineLimit(`query '${query}' has an infinite number`)
    return cell
  }
  if (typeof cell === 'boolean') return cell
  if (cell instanceof MDate || cell instanceof MDateTime) return cell.serial
  throw new EngineLimit(`query '${query}' has a ${kindOf(cell)} value in a cell`)
}

export interface QueryTableWrite {
  readonly edits: CellEdit[]
  /** the new table range, when the row count changed */
  readonly tableEdit: SheetTableEdit | null
  readonly rowsBefore: number
  readonly rowsAfter: number
}

/**
 * The cell edits and table resize that load `result` into `target`, as
 * Excel's refresh does. Raises EngineLimit for what the engine can't match.
 */
export function planQueryTableWrite(
  target: QueryTableTarget,
  result: MTable,
  worksheetXml: string,
): QueryTableWrite {
  const where = `query '${target.query}' (table '${target.tableName}')`
  if (target.unsupported.length > 0)
    throw new EngineLimit(`${where}: ${target.unsupported.join('; ')}`)
  const bound = target.columns.filter((c) => c.field !== null).map((c) => c.field!)
  const resultSet = new Set(result.columns)
  if (resultSet.size !== result.columns.length)
    throw new EngineLimit(`${where}: the result has duplicate column names`)
  if (bound.length !== result.columns.length || bound.some((name) => !resultSet.has(name))) {
    throw new EngineLimit(`${where}: the query's columns changed (${result.columns.join(', ')})`)
  }
  const { startRow, startColumn, endRow, endColumn } = target.area
  const rowsBefore = endRow - startRow
  const rowsAfter = Math.max(result.rows.length, 1)
  const lastRow = startRow + Math.max(rowsBefore, rowsAfter)
  const cells = sheetCells(worksheetXml, startRow + 1, lastRow, startColumn, endColumn)

  // rows the table grows into must be empty, as nothing below is moved
  for (let r = endRow + 1; r <= startRow + rowsAfter; r++) {
    for (const cell of cells.get(r)?.values() ?? []) {
      if (cell.filled)
        throw new EngineLimit(`${where}: there are cells below the table that would have to move`)
    }
  }
  // rows that are emptied must not hold the master of a shared formula
  for (let r = startRow + rowsAfter + 1; r <= endRow; r++) {
    for (const cell of cells.get(r)?.values() ?? []) {
      if (cell.sharedMaster)
        throw new EngineLimit(`${where}: a shared formula starts in a row that goes away`)
    }
  }

  if (rowsAfter !== rowsBefore && target.sort.some((c) => c.endRow !== endRow + 1)) {
    throw new EngineLimit(`${where}: the table's sort does not cover all its rows`)
  }
  const resultIndex = new Map(result.columns.map((name, i) => [name, i]))
  const rows = sortedRows(target, result, resultIndex, where)
  const styleAbove = (column: number): number | undefined =>
    rowsBefore > 0 ? (cells.get(endRow)?.get(column)?.style ?? undefined) : undefined
  const edits: CellEdit[] = []
  for (let r = 0; r < rowsAfter; r++) {
    const row = startRow + 1 + r
    const isNew = r >= rowsBefore
    const values = rows[r]
    target.columns.forEach((column, j) => {
      const sheetColumn = startColumn + j
      const styleIndex = isNew ? styleAbove(sheetColumn) : undefined
      const style = styleIndex === undefined ? {} : { styleIndex }
      if (column.field !== null) {
        const value = values
          ? sheetValue(values[resultIndex.get(column.field)!] ?? null, target.query)
          : null
        edits.push({
          sheetName: target.sheetName,
          row,
          column: sheetColumn,
          writeValue: true,
          cell: { value },
          ...style,
        })
      } else if (isNew) {
        edits.push({
          sheetName: target.sheetName,
          row,
          column: sheetColumn,
          writeValue: true,
          cell: { value: null, formula: column.formula! },
          ...style,
        })
      }
    })
  }
  for (let row = startRow + rowsAfter + 1; row <= endRow; row++) {
    for (let column = startColumn; column <= endColumn; column++) {
      if (!cells.get(row)?.has(column)) continue
      edits.push({
        sheetName: target.sheetName,
        row,
        column,
        writeValue: true,
        cell: { value: null },
      })
    }
  }
  const tableEdit: SheetTableEdit | null =
    rowsAfter === rowsBefore
      ? null
      : {
          sheetName: target.sheetName,
          tableName: target.tableName,
          area: { startRow, startColumn, endRow: startRow + rowsAfter, endColumn },
          ...(target.sort.length > 0 ? { sortFollowsRows: true as const } : {}),
          ...(target.localName ? { localName: target.localName } : {}),
        }
  return { edits, tableEdit, rowsBefore, rowsAfter }
}

/**
 * The result rows in the order of the table's sort, which Excel applies
 * again after a refresh: by each condition in turn, numbers (and dates)
 * before text, blanks last, rows with equal keys in the query's order.
 * Excel orders different texts by the language's collation, which the engine
 * does not reproduce: a sort column may hold only one distinct text (such as
 * "cancelado"), and a descending one none; anything else raises EngineLimit.
 */
function sortedRows(
  target: QueryTableTarget,
  result: MTable,
  resultIndex: ReadonlyMap<string, number>,
  where: string,
): readonly (readonly MCell[])[] {
  if (target.sort.length === 0) return result.rows
  const keys = target.sort.map((condition) => {
    const column = target.columns[condition.column - target.area.startColumn]
    if (!column) throw new EngineLimit(`${where}: the table is sorted by a column outside it`)
    if (column.field === null)
      throw new EngineLimit(`${where}: the table is sorted by a calculated column`)
    return {
      index: resultIndex.get(column.field)!,
      descending: condition.descending,
      text: null as string | null,
    }
  })
  /** 0 number, 1 text, 2 blank */
  const keyOf = (row: readonly MCell[], k: (typeof keys)[number]): [number, number] => {
    const value = sheetValue(row[k.index] ?? null, target.query)
    if (value === null) return [2, 0]
    if (typeof value === 'number') return [0, value]
    if (typeof value === 'string' && !k.descending && (k.text === null || k.text === value)) {
      k.text = value
      return [1, 0]
    }
    throw new EngineLimit(
      `${where}: the table is sorted by a column holding texts or true/false the engine can't order`,
    )
  }
  const decorated = result.rows.map((row, position) => ({
    row,
    position,
    values: keys.map((k) => keyOf(row, k)),
  }))
  decorated.sort((a, b) => {
    for (let k = 0; k < keys.length; k++) {
      const [rankA, x] = a.values[k]!
      const [rankB, y] = b.values[k]!
      if (rankA !== rankB) return rankA - rankB
      if (x !== y) return keys[k]!.descending ? y - x : x - y
    }
    return a.position - b.position
  })
  return decorated.map((d) => d.row)
}
