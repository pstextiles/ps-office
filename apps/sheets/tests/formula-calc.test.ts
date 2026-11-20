/**
 * The built-in recalculation (PS Office). The rules pinned here are the ones
 * a real production-plan workbook exercises: its 306,451 formula cells were
 * recalculated and compared with Excel's own values, all equal. A rule the
 * engine is not sure of must raise FormulaLimit, never give another answer.
 */
import { describe, expect, it } from 'vitest'

import { Calculator, isArrayValue, type CalcOptions } from '../src/main/formula-calc/calc'
import {
  ERR,
  XlError,
  cellKey,
  readWorkbookModel,
  type Scalar,
} from '../src/main/formula-calc/model'
import { FormulaLimit, parseFormula } from '../src/main/formula-calc/parser'
import { zip } from './power-query-fixtures'

// ---------- a workbook from a description

type CellSpec =
  | number
  | string
  | boolean
  /** a formula; `array` is the range of an array formula, `dynamic` marks a spilling one, `stored` its saved value */
  | { f: string; array?: string; dynamic?: boolean; stored?: number | string }
  /** a cell that holds a value spilled by an array formula */
  | { spilled: number | string }

interface SheetSpec {
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

function workbook(sheets: readonly SheetSpec[]): Buffer {
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

/** 2026-11-20, a Friday */
const TODAY = 46346
const OPTIONS: CalcOptions = { today: TODAY, dateSeparator: '/', dayFirst: false }

const DATA: SheetSpec = {
  name: 'Raw data',
  table: { name: 'Orders', ref: 'A1:D6', columns: ['PO', 'Qty', "D'LY", 'Line'] },
  cells: {
    A1: 'PO',
    B1: 'Qty',
    C1: "D'LY",
    D1: 'Line',
    A2: '4237995637',
    B2: 100,
    C2: 46331,
    D2: 'GT-005',
    A3: '4238049729',
    B3: 250,
    C3: 46327,
    D3: 'GT-004',
    A4: '4237995637_2',
    B4: 40,
    C4: 'cancelado',
    D4: 'gt-004',
    A5: '4239000001',
    B5: '7',
    C5: 46350,
    D5: 'GT-010',
    A6: '4239000002',
    C6: 46331,
  },
}

/** the values a sheet shows after recalculating formulas given by cell */
function calculate(
  formulas: Record<string, CellSpec>,
  options: CalcOptions = OPTIONS,
): (ref: string) => Scalar {
  const book = readWorkbookModel(workbook([DATA, { name: 'Calc', cells: formulas }]))
  const calc = new Calculator(book, options)
  const sheet = book.sheetByName.get('calc')!
  return (ref) => {
    const column =
      [...ref.replace(/\d/g, '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    return calc.cellValue(sheet, Number(ref.replace(/\D/g, '')) - 1, column)
  }
}
/** one formula's result */
const value = (f: string, options?: CalcOptions): Scalar => calculate({ A1: { f } }, options)('A1')
/** one array formula's result as rows */
function spilled(f: string): Scalar[][] {
  const book = readWorkbookModel(
    workbook([DATA, { name: 'Calc', cells: { A1: { f, array: 'A1', dynamic: true } } }]),
  )
  const calc = new Calculator(book, OPTIONS)
  const sheet = book.sheetByName.get('calc')!
  const result = calc.formulaResult(sheet, 0, 0, sheet.cells.get(cellKey(0, 0))!)
  if (!isArrayValue(result)) return [[result]]
  return Array.from({ length: result.rows }, (_, r) =>
    result.data.slice(r * result.cols, (r + 1) * result.cols),
  )
}

describe('parsing formulas as files store them', () => {
  it("reads operators in Excel's precedence", () => {
    expect(value('-2^2')).toBe(4)
    expect(value('2^3^2')).toBe(64)
    expect(value('1+2*3')).toBe(7)
    expect(value('50%+1')).toBe(1.5)
    expect(value('"a"&1+1')).toBe('a2')
    expect(value('1+1=2')).toBe(true)
  })

  it('refuses what it does not know', () => {
    for (const text of ['A1 B1', '[1]Sheet1!A1', '(A1,B2)', 'A1+']) {
      expect(() => parseFormula(text)).toThrow(FormulaLimit)
    }
    expect(() => value('SUM(Calc:Other!A1)')).toThrow(FormulaLimit)
    expect(() => value('NOSUCHFUNCTION(1)')).toThrow(FormulaLimit)
    expect(() => value('SomeName+1')).toThrow(FormulaLimit)
  })

  it('reads structured references with escaped column names', () => {
    expect(parseFormula("Orders[[#This Row],[D''LY]]")).toMatchObject({
      t: 'sref',
      table: 'Orders',
      thisRow: true,
      columns: ["D'LY", "D'LY"],
    })
    expect(parseFormula('Orders[Qty]')).toMatchObject({
      t: 'sref',
      thisRow: false,
      area: 'data',
      columns: ['Qty', 'Qty'],
    })
    expect(value('SUM(Orders[Qty])')).toBe(390)
  })
})

describe('values and operators', () => {
  it('converts as Excel does', () => {
    expect(value('"3"+1')).toBe(4)
    expect(value('"abc"+1')).toBe(ERR.VALUE)
    expect(value('""+1')).toBe(ERR.VALUE)
    expect(value('TRUE+1')).toBe(2)
    expect(value("'Raw data'!B6+1")).toBe(1)
    expect(value('1/0')).toBe(ERR.DIV0)
    expect(value('1&TRUE&""')).toBe('1TRUE')
    expect(value('0.1+0.2&""')).toBe('0.3')
  })

  it("reads text dates in the locale's order, with this year when none is given", () => {
    // 2026-08-12 is 46246
    expect(value('"8/12"-7')).toBe(46239)
    // 8 December 2026 is 46364
    expect(value('"8/12"-7', { ...OPTIONS, dayFirst: true })).toBe(46357)
    expect(value('"2026-08-12"+0')).toBe(46246)
    expect(value('"PO 123"+1')).toBe(ERR.VALUE)
    // text another locale's Excel might read as a date is not guessed
    expect(() => value('"12 de agosto"*1')).toThrow(FormulaLimit)
    expect(() => value('"8.12.2026"*1')).toThrow(FormulaLimit)
  })

  it('compares: text ignores case, a blank is 0 or "", numbers come before text', () => {
    expect(value('"OK"="ok"')).toBe(true)
    expect(value('\'Raw data\'!B6=""')).toBe(true)
    expect(value("'Raw data'!B6=0")).toBe(true)
    expect(value('1<"a"')).toBe(true)
    expect(value('"a"<TRUE')).toBe(true)
    expect(value('"UA1306443"<"UA1384224"')).toBe(true)
    // hyphens are ignored when ordering text
    expect(value('"GT-004"<"GT-005"')).toBe(true)
    expect(() => value('"a_b"<"a.b"')).toThrow(FormulaLimit)
    // equality never needs an order
    expect(value('"a_b"="a.b"')).toBe(false)
  })

  it('shows 0 for a formula that results in an empty cell', () => {
    expect(value("'Raw data'!B6")).toBe(0)
    expect(value("IF(TRUE,'Raw data'!B6)")).toBe(0)
    expect(value('IF(FALSE,1)')).toBe(false)
    expect(value('IF(TRUE,,1)')).toBe(0)
  })

  it("takes the cell of the formula's row from a range, outside array formulas", () => {
    const at = calculate({ F3: { f: "'Raw data'!B:B+1" }, F9: { f: "'Raw data'!B2:B4*2" } })
    expect(at('F3')).toBe(251)
    expect(at('F9')).toBe(ERR.VALUE)
  })
})

describe('lookups', () => {
  it('XLOOKUP and MATCH match exactly: text ignores case, text never matches a number', () => {
    expect(value("XLOOKUP(\"4238049729\",'Raw data'!A:A,'Raw data'!B:B)")).toBe(250)
    expect(value("XLOOKUP(4238049729,'Raw data'!A:A,'Raw data'!B:B)")).toBe(ERR.NA)
    expect(value('XLOOKUP("nope",\'Raw data\'!A:A,\'Raw data\'!B:B,"")')).toBe('')
    expect(value("XLOOKUP(\"gt-004\",'Raw data'!D:D,'Raw data'!A:A)")).toBe('4238049729')
    expect(value('MATCH("Qty",\'Raw data\'!1:1,0)')).toBe(2)
    expect(
      value(
        "INDEX('Raw data'!A:D,MATCH(\"4239000001\",'Raw data'!A:A,0),MATCH(\"Line\",'Raw data'!1:1,0))",
      ),
    ).toBe('GT-010')
    expect(value('MATCH("zzz",\'Raw data\'!A:A,0)')).toBe(ERR.NA)
    expect(() => value("MATCH(5,'Raw data'!B:B,1)")).toThrow(FormulaLimit)
  })

  it('counts and sums by criteria', () => {
    expect(value("SUMIFS('Raw data'!B:B,'Raw data'!D:D,\"GT-004\")")).toBe(290)
    expect(value("SUMIF('Raw data'!A:A,\"4237995637\",'Raw data'!B:B)")).toBe(100)
    expect(value("COUNTIF('Raw data'!C:C,46331)")).toBe(2)
    expect(value('COUNTIF(\'Raw data\'!D2:D6,"<>GT-004")')).toBe(3)
    expect(value('COUNTIF(\'Raw data\'!D2:D6,"GT-*")')).toBe(4)
    // a number criterion also matches numeric text, and "" matches empty cells
    expect(value("COUNTIF('Raw data'!B2:B6,7)")).toBe(1)
    expect(value('COUNTIF(\'Raw data\'!B2:B6,"")')).toBe(1)
    expect(value('COUNTIFS(\'Raw data\'!B2:B6,">50",\'Raw data\'!C2:C6,"<46330")')).toBe(1)
  })
})

describe('functions', () => {
  it('text', () => {
    expect(value('TEXT(1234.5,"0")')).toBe('1235')
    expect(value('TEXT(-2.5,"0")')).toBe('-3')
    expect(value('TEXT(1234567,"#,##0")')).toBe('1,234,567')
    expect(value('TEXT(46331,"mm/dd")')).toBe('11/05')
    expect(value('TEXT(46331,"yyyy-mm")')).toBe('2026-11')
    expect(value('TEXT("abc","0")')).toBe('abc')
    expect(() => value('TEXT(1,"0.00")')).toThrow(FormulaLimit)
    expect(value('TRIM("  a   b ")')).toBe('a b')
    expect(value('MID("abcdef",2,3)&LEFT("xyz",2)&LEN("abc")')).toBe('bcdxy3')
    expect(value('SEARCH("STOCK","en stock")')).toBe(4)
    expect(value('ISNUMBER(SEARCH("x","abc"))')).toBe(false)
    // ignore_empty left out means TRUE
    expect(value('TEXTJOIN(",",,"a","","b")')).toBe('a,b')
    expect(value('TEXTJOIN(",",FALSE,"a","","b")')).toBe('a,,b')
  })

  it('dates', () => {
    expect(value('TODAY()')).toBe(TODAY)
    expect(value('MONTH(46331)*100+YEAR(46331)')).toBe(1100 + 2026)
    expect(value('WEEKDAY(46331)')).toBe(5)
    expect(value('DATE(2026,13,1)')).toBe(46388)
    // 5 working days after Thursday 2026-11-05, with Monday the 9th a holiday
    expect(value('WORKDAY(46331,5)')).toBe(46338)
    expect(value('WORKDAY(46331,5,46335)')).toBe(46339)
    expect(value('WORKDAY(46331,-4)')).toBe(46325)
  })

  it('logic and errors', () => {
    expect(value('IFERROR(1/0,"x")')).toBe('x')
    expect(value('IFNA(1/0,"x")')).toBe(ERR.DIV0)
    expect(value('IFNA(MATCH(1,{2,3},0),"x")')).toBe('x')
    expect(value('AND(1,TRUE)')).toBe(true)
    expect(value('OR(0,FALSE)')).toBe(false)
    expect(value('LET(v,2,w,v*3,v+w)')).toBe(8)
    expect(value('N("a")+N(TRUE)+N(3)')).toBe(4)
    expect(value('MAX(1,"5",TRUE)')).toBe(5)
    expect(value("MIN('Raw data'!B2:B6)")).toBe(40)
  })
})

describe('array formulas', () => {
  it('spreads operators and functions over arrays and spills the result', () => {
    expect(spilled("'Raw data'!B2:B4*2")).toEqual([[200], [500], [80]])
    // "7" is text, and text is greater than any number
    expect(spilled("FILTER('Raw data'!A2:B6,'Raw data'!B2:B6>50,\"\")")).toEqual([
      ['4237995637', 100],
      ['4238049729', 250],
      ['4239000001', '7'],
    ])
    expect(
      spilled(
        "FILTER('Raw data'!A2:A6,ISNUMBER('Raw data'!B2:B6)*('Raw data'!B2:B6>999),\"none\")",
      ),
    ).toEqual([['none']])
    expect(spilled("UNIQUE('Raw data'!D2:D5)")).toEqual([['GT-005'], ['GT-004'], ['GT-010']])
    expect(spilled('HSTACK({1;2;3},"x")')).toEqual([
      [1, 'x'],
      [2, ERR.NA],
      [3, ERR.NA],
    ])
    expect(spilled('INDEX({"a","b","c"},{3;1})')).toEqual([['c'], ['a']])
    expect(spilled('CHOOSE({2;1},"x","y")')).toEqual([['y'], ['x']])
    // an empty cell in a spilled array shows 0
    expect(spilled("'Raw data'!B5:B6")).toEqual([['7'], [0]])
  })

  it('sorts as Excel does: numbers, then text ignoring case and hyphens, blanks last, ties in order', () => {
    expect(spilled("SORTBY('Raw data'!A2:A6,'Raw data'!C2:C6,1,'Raw data'!B2:B6,-1)")).toEqual([
      ['4238049729'],
      ['4237995637'],
      ['4239000002'],
      ['4239000001'],
      ['4237995637_2'],
    ])
    expect(spilled("SORT(UNIQUE('Raw data'!D2:D5))")).toEqual([['GT-004'], ['GT-005'], ['GT-010']])
    expect(() => spilled('SORT({"a_1";"a.1"})')).toThrow(FormulaLimit)
  })

  it('a formula reads cells an array formula fills, and a shrunken spill leaves its old cells empty', () => {
    const at = calculate({
      A1: {
        f: "FILTER('Raw data'!B2:B6,ISNUMBER('Raw data'!B2:B6)*('Raw data'!B2:B6>50),\"\")",
        array: 'A1:A4',
        dynamic: true,
        stored: 1,
      },
      A2: { spilled: 2 },
      A3: { spilled: 3 },
      A4: { spilled: 4 },
      C1: { f: 'SUM(A1#)' },
      C2: { f: 'A2+A3' },
      C3: { f: 'COUNTA(A1:A4)' },
    })
    expect([at('A1'), at('A2'), at('A3'), at('A4')]).toEqual([100, 250, null, null])
    expect(at('C1')).toBe(350)
    expect(at('C2')).toBe(250)
    expect(at('C3')).toBe(2)
  })

  it('is blocked by a cell in its way', () => {
    const at = calculate({ A1: { f: '{1;2;3}', array: 'A1', dynamic: true }, A3: 'in the way' })
    expect(at('A1')).toBe(ERR.SPILL)
    expect(at('A2')).toBeNull()
  })

  it('a fixed (Ctrl+Shift+Enter) array formula fills its range, padded with #N/A', () => {
    const at = calculate({
      A1: { f: '{1,2}', array: 'A1:C1' },
      B1: { spilled: 0 },
      C1: { spilled: 0 },
    })
    expect([at('A1'), at('B1'), at('C1')]).toEqual([1, 2, ERR.NA])
  })
})

describe('the model', () => {
  it('reads shared formulas, array ranges and newer errors', () => {
    const NS2 =
      'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    const book = readWorkbookModel(
      zip({
        'xl/workbook.xml': `<workbook ${NS2}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
        'xl/worksheets/sheet1.xml':
          `<worksheet ${NS2}><sheetData>` +
          '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*10</f><v>0</v></c></row>' +
          '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>0</v></c></row>' +
          '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c><c r="D3" t="e" vm="1"><v>#VALUE!</v></c></row>' +
          '</sheetData></worksheet>',
      }),
    )
    const sheet = book.sheets[0]!
    const calc = new Calculator(book, OPTIONS)
    expect([0, 1, 2].map((r) => calc.cellValue(sheet, r, 1))).toEqual([10, 20, 30])
    expect(sheet.cells.get(cellKey(2, 3))).toMatchObject({
      stored: XlError.of('#VALUE!'),
      rich: true,
    })
  })

  it('refuses a circular reference', () => {
    const at = calculate({ A1: { f: 'B1+1' }, B1: { f: 'A1+1' } })
    expect(() => at('A1')).toThrow(FormulaLimit)
  })
})
