/**
 * The built-in Power Query engine (PS Office). The conversion answers below
 * are Power Query's own results, taken from Excel in each culture; the engine
 * must give the same answer or stop with EngineLimit, never a different one.
 */
import { describe, expect, it } from 'vitest'

import {
  evaluateQuery,
  parseQueries,
  workbookQueries,
  type QueryHost,
} from '../src/main/power-query'
import {
  cultureNamed,
  dateFrom,
  dateTimeFrom,
  numberFrom,
  partsToSerial,
  textFrom,
} from '../src/main/power-query/convert'
import { readWorkbookSheets } from '../src/main/power-query/excel-source'
import {
  EngineLimit,
  MDate,
  MDateTime,
  MError,
  isErrorCell,
  type MList,
  type MTable,
  type MValue,
} from '../src/main/power-query/values'
import { dataMashupPart, zip } from './power-query-fixtures'

// ---------- a workbook like the factory's raw files

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const SHARED = ['Title', 'PO', 'Qty', 'Fecha', '7']
const DATA_SHEET = `<worksheet ${NS}><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>2</v></c><c r="C2" t="s"><v>3</v></c></row>
<row r="3"><c r="A3"><v>4236534533</v></c><c r="B3"><v>10</v></c><c r="C3" s="1"><v>46331</v></c></row>
<row r="4"><c r="A4" t="str"><f>IFERROR(VLOOKUP(X4,Orden,4,FALSE),"")</f><v/></c><c r="B4"><v>5</v></c><c r="C4" t="inlineStr"><is><t>05/11/2026</t></is></c></row>
<row r="5"><c r="A5"><v>4236534533</v></c><c r="B5" t="s"><v>4</v></c><c r="C5" s="2"><v>46331.5</v></c><c r="D5" t="e"><v>#N/A</v></c></row>
<row r="6"><c r="A6" s="1"/><c r="B6"><v>99</v></c></row>
</sheetData></worksheet>`

function workbookParts(): Record<string, string | Buffer> {
  return {
    '[Content_Types].xml':
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    'xl/workbook.xml': `<workbook ${NS} ${R_NS}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Old" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="worksheet" Target="/xl/worksheets/sheet2.xml"/></Relationships>',
    'xl/sharedStrings.xml': `<sst ${NS}>${SHARED.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`,
    'xl/styles.xml': `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy\\ hh:mm"/></numFmts><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/></cellXfs></styleSheet>`,
    'xl/worksheets/sheet1.xml': DATA_SHEET,
    'xl/worksheets/sheet2.xml': `<worksheet ${NS}><sheetData><row r="3"><c r="C3"><v>1</v></c></row></sheetData></worksheet>`,
  }
}

const QUERY = `section Section1;
shared Cut = let
    NumN = (x) => try Number.From(x) otherwise null,
    DateN = (x) => try Date.From(x) otherwise null,
    Src = Excel.Workbook(File.Contents("book.xlsx"), null, true),
    Data = Src{[Item="Data",Kind="Sheet"]}[Data],
    Head = Table.PromoteHeaders(Table.Skip(Data, 1), [PromoteAllScalars=true]),
    Keep = Table.SelectColumns(Head, {"PO","Qty","Fecha","Missing"}, MissingField.UseNull),
    Clean = Table.SelectRows(Keep, each [PO] <> null),
    G = Table.Group(Clean, {"PO"}, {
        {"Qty", each List.Sum(List.Transform([Qty], NumN)), type number},
        {"First", each List.Min(List.RemoveNulls(List.Transform([Fecha], DateN)), null), type date}
    }),
    K = Table.TransformColumns(G, {{"PO", each Text.From(_), type text}})
in
    K;`

function mashupWorkbook(section: string, culture: string): Buffer {
  return zip({ ...workbookParts(), 'customXml/item1.xml': dataMashupPart(section, culture) })
}

const host = (files: Record<string, Buffer>): QueryHost => ({
  readFile: (path) => {
    const hit = files[path]
    if (!hit) throw new Error(`no file ${path}`)
    return hit
  },
  listFolder: () => [],
})

const serial = (y: number, m: number, d: number, hh = 0, mm = 0, ss = 0): number =>
  partsToSerial(y, m, d, hh, mm, ss)!

/** a table as plain rows: dates become serials, error cells 'ERR' */
function plain(t: MTable): unknown[][] {
  return t.rows.map((row) =>
    row.map((c) =>
      isErrorCell(c) ? 'ERR' : c instanceof MDate || c instanceof MDateTime ? c.serial : c,
    ),
  )
}

describe('Excel.Workbook source', () => {
  it('reads cells as Power Query does: formula "" as empty text, dates by style, errors as error cells', () => {
    const sheets = readWorkbookSheets(zip(workbookParts()))
    expect(sheets.map((s) => [s.name, s.hidden])).toEqual([
      ['Data', false],
      ['Old', true],
    ])
    const data = sheets[0]!
    // the used range starts at A1 and spans to column D (D5 holds an error)
    expect(data.width).toBe(4)
    expect(data.rows).toHaveLength(6)
    expect(data.rows[3]![0]).toBe('')
    expect(data.rows[2]![2]).toEqual(new MDateTime(46331))
    expect(data.rows[4]![2]).toEqual(new MDateTime(46331.5))
    expect(data.rows[4]![1]).toBe('7')
    expect(isErrorCell(data.rows[4]![3]!)).toBe(true)
    // a styled cell without a value is no value
    expect(data.rows[5]![0]).toBeNull()
    // the hidden sheet's used range is just C3
    expect(sheets[1]!.rows).toEqual([[1]])
  })
})

describe('evaluating a query', () => {
  const files = { 'book.xlsx': zip(workbookParts()) }

  it('groups the raw rows like Excel: empty-text PO is its own group, null PO is dropped', () => {
    const doc = parseQueries(QUERY)
    const out = evaluateQuery(doc, 'Cut', { host: host(files), culture: 'en-US' }) as MTable
    expect(out.columns).toEqual(['PO', 'Qty', 'First'])
    expect(plain(out)).toEqual([
      ['4236534533', 17, serial(2026, 11, 5)],
      // "05/11/2026" is May 11 in en-US
      ['', 5, serial(2026, 5, 11)],
    ])
  })

  it('reads text in the locale the workbook stores, even on a PC set to another', () => {
    const book = mashupWorkbook(QUERY, 'es-GT')
    const doc = workbookQueries(book)!
    expect(doc.culture).toBe('es-GT')
    expect(doc.names).toEqual(['Cut'])
    const out = evaluateQuery(doc, 'Cut', { host: host(files), systemCulture: 'ko-KR' }) as MTable
    // "05/11/2026" is 5 November in es-GT
    expect(plain(out)[1]).toEqual(['', 5, serial(2026, 11, 5)])
  })

  it('falls back to the PC locale when the workbook stores none, as Excel does', () => {
    const doc = parseQueries(QUERY)
    expect(doc.culture).toBeNull()
    const gt = evaluateQuery(doc, 'Cut', { host: host(files), systemCulture: 'es-GT' }) as MTable
    expect(plain(gt)[1]).toEqual(['', 5, serial(2026, 11, 5)])
    const ko = evaluateQuery(doc, 'Cut', { host: host(files), systemCulture: 'ko-KR' }) as MTable
    expect(plain(ko)[1]).toEqual(['', 5, serial(2026, 5, 11)])
    // no locale anywhere: no guess
    expect(() => evaluateQuery(doc, 'Cut', { host: host(files) })).toThrow(EngineLimit)
    // a PC locale the engine doesn't know: no guess either
    expect(() => evaluateQuery(doc, 'Cut', { host: host(files), systemCulture: 'fr-FR' })).toThrow(
      EngineLimit,
    )
  })

  it('reports a missing source file as an M error', () => {
    const doc = parseQueries(QUERY)
    expect(() => evaluateQuery(doc, 'Cut', { host: host({}), culture: 'en-US' })).toThrow(MError)
  })
})

describe('the M language', () => {
  const run = (expr: string, culture = 'en-US'): MValue =>
    evaluateQuery(parseQueries(`section S; shared v = ${expr};`), 'v', { host: host({}), culture })
  const items = (v: MValue): readonly MValue[] => (v as MList).items

  it('evaluates let, if, operators, records, lists and each', () => {
    expect(
      items(run('let a = 1, b = a + 2 in {a, b, null + 1, "x" & "y", if a > 0 then "p" else "n"}')),
    ).toEqual([1, 3, null, 'xy', 'p'])
    expect(run('[A = 1, B = A + 1][B]')).toBe(2)
    expect(items(run('List.Transform({1..3}, each _ * 2)'))).toEqual([2, 4, 6])
    expect(run('try error "boom" otherwise 7')).toBe(7)
    expect(run('Text.Combine(List.Sort({"b", "a", "c"}), ", ")')).toBe('a, b, c')
    expect(run('#date(2026, 11, 5)')).toEqual(new MDate(serial(2026, 11, 5)))
  })

  it('stops instead of guessing: an engine limit escapes try', () => {
    expect(() => run('try Number.From("$5") otherwise 0')).toThrow(EngineLimit)
    expect(() =>
      run('Table.AddColumn(#table({"A"}, {{"$5"}}), "N", each Number.From([A]))'),
    ).toThrow(EngineLimit)
    // Power Query keeps a 17-digit literal as written; a double can't
    expect(() => run('Text.From(12345678901234567)')).toThrow(EngineLimit)
    expect(run('Text.From(123456789012345680000)')).toBe('123456789012345680000')
    expect(() => run('Text.From(1, "fr-FR")')).toThrow(EngineLimit)
  })

  it('takes an explicit culture argument', () => {
    expect(run('Number.From("1.234,5", "es-ES")')).toBe(1234.5)
    expect(run('Text.From(#date(2026, 11, 5), "ko-KR")')).toBe('2026-11-05')
  })
})

// ---------- conversions: Excel's answers in each culture

type Answer = number | null | 'ERR' | 'LIMIT'
function outcome(f: () => unknown): Answer {
  try {
    const v = f()
    if (v instanceof MDate || v instanceof MDateTime) return v.serial
    return v as number | null
  } catch (e) {
    if (e instanceof EngineLimit) return 'LIMIT'
    if (e instanceof MError) return 'ERR'
    throw e
  }
}
const ALL = ['en-US', 'es-GT', 'ko-KR', 'es-ES'] as const

describe('Number.From on text', () => {
  const cases: [string, string, Answer][] = [
    ['1,234.5', 'en-US', 1234.5],
    ['1,234.5', 'es-ES', 'LIMIT'],
    ['1.234,5', 'es-ES', 1234.5],
    ['1.234,5', 'en-US', 'LIMIT'],
    ['2026.11.05', 'es-ES', 20261105],
    ['1,2,3', 'en-US', 123],
    ['12,34.5', 'ko-KR', 1234.5],
    ['12%', 'en-US', 0.12],
    ['12 %', 'es-GT', 0.12],
    ['%12', 'ko-KR', 0.12],
    ['(3)', 'en-US', -3],
    ['5-', 'en-US', -5],
    ['+5', 'es-ES', 5],
    ['.5', 'en-US', 0.5],
    [',5', 'es-ES', 0.5],
    ['5.', 'en-US', 5],
    ['1e-3', 'en-US', 0.001],
    ['1E+3', 'es-GT', 1000],
    [' 7 ', 'en-US', 7],
    ['', 'en-US', null],
    ['  ', 'es-GT', null],
    ['abc', 'en-US', 'ERR'],
    ['TRUE', 'en-US', 'ERR'],
    ['Infinity', 'en-US', 'ERR'],
    ['05/11/2026', 'es-GT', 'ERR'],
    ['2026-11-05', 'ko-KR', 'ERR'],
    ['$5', 'en-US', 'LIMIT'],
  ]
  it.each(cases)('%j in %s', (text, culture, expected) => {
    expect(outcome(() => numberFrom(text, cultureNamed(culture)))).toBe(expected)
  })
})

describe('Date.From and DateTime.From on text', () => {
  const cases: [string, string, Answer, Answer][] = [
    // text, culture, Date.From, DateTime.From
    ['05/11/2026', 'en-US', serial(2026, 5, 11), serial(2026, 5, 11)],
    ['05/11/2026', 'ko-KR', serial(2026, 5, 11), serial(2026, 5, 11)],
    ['05/11/2026', 'es-GT', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['05.11.2026', 'es-ES', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['20/11/2026', 'en-US', 'ERR', 'ERR'],
    ['13/13/2026', 'es-GT', 'ERR', 'ERR'],
    ['5/11/26', 'en-US', serial(2026, 5, 11), serial(2026, 5, 11)],
    ['5/11/26', 'es-GT', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['5/11/26', 'ko-KR', serial(2005, 11, 26), serial(2005, 11, 26)],
    ['26-11-05', 'es-ES', serial(2005, 11, 26), serial(2005, 11, 26)],
    ['26-11-05', 'ko-KR', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['26-11-05', 'en-US', 'ERR', 'ERR'],
    ['2026-1-5', 'es-GT', serial(2026, 1, 5), serial(2026, 1, 5)],
    ['2026/11/05', 'en-US', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['2026-11', 'ko-KR', serial(2026, 11, 1), serial(2026, 11, 1)],
    ['11/2026', 'es-ES', serial(2026, 11, 1), serial(2026, 11, 1)],
    ['2026-11-05T10:20:30', 'en-US', 'ERR', serial(2026, 11, 5, 10, 20, 30)],
    ['11/20/2026 3:04:05 PM', 'en-US', 'ERR', serial(2026, 11, 20, 15, 4, 5)],
    ['1/2/2026 12:00:00 AM', 'es-GT', 'ERR', serial(2026, 2, 1)],
    ['2026-11-05 오후 3:04:05', 'ko-KR', 'ERR', serial(2026, 11, 5, 15, 4, 5)],
    ['2026-11-05 오후 3:04:05', 'en-US', 'ERR', 'ERR'],
    ['2026년 11월 5일', 'es-GT', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['Nov 5, 2026', 'ko-KR', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['5 November 2026', 'es-ES', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['5-Nov-2026', 'en-US', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['15-ene-2026', 'es-GT', serial(2026, 1, 15), serial(2026, 1, 15)],
    ['15-ene-2026', 'en-US', 'ERR', 'ERR'],
    ['noviembre 5, 2026', 'es-ES', serial(2026, 11, 5), serial(2026, 11, 5)],
    ['', 'en-US', null, null],
    ['PENDIENTE', 'es-GT', 'ERR', 'ERR'],
    // .NET reads these as dates in odd ways (the year 1234, or this year)
    ['1234.5', 'en-US', 'LIMIT', 'LIMIT'],
    ['5/11', 'en-US', 'LIMIT', 'LIMIT'],
    ['3-Aug-26', 'ko-KR', 'LIMIT', 'LIMIT'],
  ]
  it.each(cases)('%j in %s', (text, culture, date, datetime) => {
    const c = cultureNamed(culture)
    expect(outcome(() => dateFrom(text, c))).toBe(date)
    expect(outcome(() => dateTimeFrom(text, c))).toBe(datetime)
  })
})

describe('Text.From', () => {
  it('writes numbers with 15 or 17 significant digits, in full up to 10^22', () => {
    const t = (n: number, culture = 'en-US'): string | null => textFrom(n, cultureNamed(culture))
    expect(t(12345.678)).toBe('12345.678')
    expect(t(0.1 + 0.2)).toBe('0.30000000000000004')
    expect(t(1 / 3)).toBe('0.33333333333333331')
    expect(t(2 / 3, 'es-ES')).toBe('0,66666666666666663')
    expect(t(1e21)).toBe('1000000000000000000000')
    expect(t(-1e21)).toBe('-1000000000000000000000')
    expect(t(1e16)).toBe('10000000000000000')
    expect(t(0.0001)).toBe('0.0001')
    expect(t(0.00001)).toBe('1E-05')
    expect(t(0.000001234, 'es-ES')).toBe('1,234E-06')
    expect(t(-1234.5, 'es-ES')).toBe('-1234,5')
    expect(() => t(1e100)).toThrow(EngineLimit)
  })

  it('writes dates and times in each culture', () => {
    const date = new MDate(serial(2026, 11, 5))
    const pm = new MDateTime(serial(2026, 11, 5, 15, 4, 5))
    const am = new MDateTime(serial(2026, 11, 5, 9, 0, 0))
    const answers: Record<(typeof ALL)[number], [string, string, string]> = {
      'en-US': ['11/5/2026', '11/5/2026 3:04:05 PM', '11/5/2026 9:00:00 AM'],
      'es-GT': ['5/11/2026', '5/11/2026 15:04:05', '5/11/2026 09:00:00'],
      'ko-KR': ['2026-11-05', '2026-11-05 오후 3:04:05', '2026-11-05 오전 9:00:00'],
      'es-ES': ['05/11/2026', '05/11/2026 15:04:05', '05/11/2026 9:00:00'],
    }
    for (const culture of ALL) {
      const c = cultureNamed(culture)
      expect([textFrom(date, c), textFrom(pm, c), textFrom(am, c)]).toEqual(answers[culture])
    }
  })
})
