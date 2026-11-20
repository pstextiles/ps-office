import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  assembleWithJsZip,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  planCellEditsToXlsx,
  type CellEdit,
  type SheetFormulaValues,
  type SheetTableEdit,
} from '../src/gateway/xlsx-gateway'

/**
 * PS Office: the save-path pieces a built-in query refresh needs, beyond a
 * plain table resize. Excel's own refresh of a query table
 *  - gives each new row the cell styles of the row above,
 *  - moves the table's sort state with the new row count, and
 *  - moves the hidden sheet-local name (ExternalData_N) that repeats the
 *    table's range — Excel refuses to open a workbook where the two differ.
 * All three are opt-in fields; a save that does not pass them behaves as
 * before, which the last test pins.
 */

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const TABLE_PART = 'xl/tables/table1.xml'
const SHEET_PART = 'xl/worksheets/sheet2.xml'

async function workbook(sortCondition = 'B1:B4'): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<workbook ${NS} ${R_NS}><sheets><sheet name="Info" sheetId="1" r:id="rId1"/><sheet name="Raw data" sheetId="2" r:id="rId2"/></sheets>` +
      '<definedNames>' +
      '<definedName name="ExternalData_1" localSheetId="0" hidden="1">Info!$A$1:$B$4</definedName>' +
      '<definedName name="ExternalData_1" localSheetId="1" hidden="1">\'Raw data\'!$A$1:$B$4</definedName>' +
      '<definedName name="Otro">Info!$A$1</definedName>' +
      '</definedNames></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>` +
      `<Relationship Id="rId3" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file('xl/worksheets/sheet1.xml', `<worksheet ${NS}><sheetData/></worksheet>`)
  zip.file(
    SHEET_PART,
    `<worksheet ${NS} ${R_NS}><sheetData>` +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>PO</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c></row>' +
      '<row r="2"><c r="A2" s="1" t="inlineStr"><is><t>a</t></is></c><c r="B2" s="2"><v>1</v></c></row>' +
      '<row r="3"><c r="A3" s="1" t="inlineStr"><is><t>b</t></is></c><c r="B3" s="2"><v>2</v></c></row>' +
      '<row r="4"><c r="A4" s="1" t="inlineStr"><is><t>c</t></is></c><c r="B4" s="2"><v>3</v></c></row>' +
      '</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>',
  )
  zip.file(
    'xl/worksheets/_rels/sheet2.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`,
  )
  zip.file(
    TABLE_PART,
    `<table ${NS} id="1" name="Tbl" displayName="Tbl" ref="A1:B4" tableType="queryTable" totalsRowShown="0">` +
      '<autoFilter ref="A1:B4"/>' +
      `<sortState ref="A2:B4"><sortCondition ref="${sortCondition}"/></sortState>` +
      '<tableColumns count="2"><tableColumn id="1" uniqueName="1" name="PO" queryTableFieldId="1"/><tableColumn id="2" uniqueName="2" name="Qty" queryTableFieldId="2"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>',
  )
  zip.file(
    'xl/styles.xml',
    `<styleSheet ${NS}><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
      '<cellXfs count="3"><xf xfId="0"/><xf xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="3" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function save(
  source: Buffer,
  edits: readonly CellEdit[],
  tableEdits: readonly SheetTableEdit[],
  formulaValues: readonly SheetFormulaValues[] = [],
): Promise<JSZip> {
  const args: unknown[] = [await createBufferEntrySource(source), edits]
  // formulaValues is the 21st and tableEdits the 27th parameter of planCellEditsToXlsx
  while (args.length < 26) args.push(undefined)
  args[20] = formulaValues
  args[26] = tableEdits
  const plan = await (
    planCellEditsToXlsx as (...a: unknown[]) => ReturnType<typeof planCellEditsToXlsx>
  )(...args)
  const mutation = await assembleWithJsZip(source, plan)
  assertOnlyTouchedEntriesChanged(mutation)
  return JSZip.loadAsync(mutation.buffer)
}
const text = async (zip: JSZip, path: string): Promise<string> =>
  (await zip.file(path)!.async('string')) ?? ''

const NEW_ROW: CellEdit[] = [
  {
    sheetName: 'Raw data',
    row: 4,
    column: 0,
    writeValue: true,
    cell: { value: 'd' },
    styleIndex: 1,
  },
  { sheetName: 'Raw data', row: 4, column: 1, writeValue: true, cell: { value: 4 }, styleIndex: 2 },
  { sheetName: 'Raw data', row: 5, column: 0, writeValue: true, cell: { value: 'e' } },
]
const GROWN = { startRow: 0, startColumn: 0, endRow: 4, endColumn: 1 }

describe('a query table refresh through the save path', () => {
  it('gives a new cell the style index it is told to copy', async () => {
    const zip = await save(await workbook(), NEW_ROW, [])
    const sheet = await text(zip, SHEET_PART)
    expect(sheet).toContain(
      '<c r="A5" s="1" t="inlineStr"><is><t xml:space="preserve">d</t></is></c>',
    )
    expect(sheet).toContain('<c r="B5" s="2"><v>4</v></c>')
    // without styleIndex a new cell has no style, as before
    expect(sheet).toContain('<c r="A6" t="inlineStr">')
    // existing rows are untouched
    expect(sheet).toContain('<c r="B4" s="2"><v>3</v></c>')
  })

  it('moves the sort state and the hidden range name with the table', async () => {
    const zip = await save(await workbook(), NEW_ROW.slice(0, 2), [
      {
        sheetName: 'Raw data',
        tableName: 'Tbl',
        area: GROWN,
        sortFollowsRows: true,
        localName: { name: 'ExternalData_1', sheetIndex: 1 },
      },
    ])
    const table = await text(zip, TABLE_PART)
    expect(table).toContain('ref="A1:B5" tableType="queryTable"')
    expect(table).toContain('<autoFilter ref="A1:B5"/>')
    expect(table).toContain('<sortState ref="A2:B5"><sortCondition ref="B1:B5"/></sortState>')
    const book = await text(zip, 'xl/workbook.xml')
    // this sheet's name moved, quoting kept; the other sheet's name of the same name did not
    expect(book).toContain(
      '<definedName name="ExternalData_1" localSheetId="1" hidden="1">\'Raw data\'!$A$1:$B$5</definedName>',
    )
    expect(book).toContain(
      '<definedName name="ExternalData_1" localSheetId="0" hidden="1">Info!$A$1:$B$4</definedName>',
    )
    expect(book).toContain('<definedName name="Otro">Info!$A$1</definedName>')
  })

  it('shrinks the same way', async () => {
    const zip = await save(
      await workbook(),
      [],
      [
        {
          sheetName: 'Raw data',
          tableName: 'Tbl',
          area: { ...GROWN, endRow: 2 },
          sortFollowsRows: true,
          localName: { name: 'ExternalData_1', sheetIndex: 1 },
        },
      ],
    )
    expect(await text(zip, TABLE_PART)).toContain(
      '<sortState ref="A2:B3"><sortCondition ref="B1:B3"/></sortState>',
    )
    expect(await text(zip, 'xl/workbook.xml')).toContain("'Raw data'!$A$1:$B$3<")
  })

  it('refuses a range name that is missing and a sort that does not end on the last row', async () => {
    await expect(
      save(
        await workbook(),
        [],
        [
          {
            sheetName: 'Raw data',
            tableName: 'Tbl',
            area: GROWN,
            localName: { name: 'ExternalData_9', sheetIndex: 1 },
          },
        ],
      ),
    ).rejects.toThrow(/ExternalData_9/)
    await expect(
      save(
        await workbook('B1:B3'),
        [],
        [{ sheetName: 'Raw data', tableName: 'Tbl', area: GROWN, sortFollowsRows: true }],
      ),
    ).rejects.toThrow(/sort condition/)
  })

  it('leaves the sort state and the range name alone on a plain resize', async () => {
    const zip = await save(
      await workbook(),
      [],
      [{ sheetName: 'Raw data', tableName: 'Tbl', area: GROWN }],
    )
    const table = await text(zip, TABLE_PART)
    expect(table).toContain('ref="A1:B5" tableType="queryTable"')
    expect(table).toContain('<sortState ref="A2:B4"><sortCondition ref="B1:B4"/></sortState>')
    expect(await text(zip, 'xl/workbook.xml')).toContain("'Raw data'!$A$1:$B$4<")
  })
})

/**
 * PS Office: what the built-in recalculation writes besides a formula cell's
 * cached value — the new range of a dynamic array formula, the values in
 * the cells it fills (errors included, cells created or emptied as the array
 * grows or shrinks), and no stale "newer error" mark on a value.
 */
describe('recalculated array formulas through the save path', () => {
  const INFO = 'xl/worksheets/sheet1.xml'
  async function arrayWorkbook(): Promise<Buffer> {
    const zip = await JSZip.loadAsync(await workbook())
    zip.file(
      INFO,
      `<worksheet ${NS}><sheetData>` +
        '<row r="1"><c r="A1" t="str" cm="1"><f t="array" ref="A1:A2">_xlfn._xlws.FILTER(E:E,F:F)</f><v>b</v></c>' +
        '<c r="C1" t="e" vm="1"><f>SUM(A1:A9)</f><v>#VALUE!</v></c></row>' +
        '<row r="2"><c r="A2" s="1" t="str"><f ca="1"/><v>c</v></c><c r="B2" s="2"><v>9</v></c></row>' +
        '</sheetData></worksheet>',
    )
    return zip.generateAsync({ type: 'nodebuffer' })
  }
  const values = (cells: SheetFormulaValues['cells']): SheetFormulaValues[] => [
    { sheetName: 'Info', cells },
  ]

  it('moves the array range, writes the filled cells and clears the mark of a newer error', async () => {
    const zip = await save(
      await arrayWorkbook(),
      [],
      [],
      values([
        { row: 0, column: 0, value: 'x', arrayRef: 'A1:A3' },
        { row: 1, column: 0, value: 5, spill: true },
        { row: 2, column: 0, value: { error: '#N/A' }, spill: true },
        { row: 1, column: 1, value: null, spill: true },
        { row: 0, column: 2, value: 7, clearValueMeta: true },
      ]),
    )
    const sheet = await text(zip, INFO)
    expect(sheet).toContain(
      '<c r="A1" cm="1" t="str"><f t="array" ref="A1:A3">_xlfn._xlws.FILTER(E:E,F:F)</f><v>x</v></c>',
    )
    // the filled cell keeps its style and loses the spill marker; a cell the array left keeps only its style
    expect(sheet).toContain('<c r="A2" s="1"><v>5</v></c>')
    expect(sheet).toContain('<c r="B2" s="2"/>')
    // a cell the array now reaches is created, an error value included
    expect(sheet).toContain('<row r="3"><c r="A3" t="e"><v>#N/A</v></c></row>')
    expect(sheet).toContain('<c r="C1"><f>SUM(A1:A9)</f><v>7</v></c>')
  })

  it('keeps the mark while the value is still an error, and leaves other cells alone without the fields', async () => {
    const zip = await save(
      await arrayWorkbook(),
      [],
      [],
      values([
        { row: 0, column: 2, value: { error: '#VALUE!' }, clearValueMeta: true },
        { row: 0, column: 0, value: 'y' },
      ]),
    )
    const sheet = await text(zip, INFO)
    expect(sheet).toContain('<c r="C1" vm="1" t="e"><f>SUM(A1:A9)</f><v>#VALUE!</v></c>')
    expect(sheet).toContain('<f t="array" ref="A1:A2">_xlfn._xlws.FILTER(E:E,F:F)</f><v>y</v>')
    expect(sheet).toContain('<c r="A2" s="1" t="str"><f ca="1"/><v>c</v></c>')
  })

  it('refuses a spilled value on a formula cell and a range that does not start at the cell', async () => {
    await expect(
      save(await arrayWorkbook(), [], [], values([{ row: 0, column: 2, value: 1, spill: true }])),
    ).rejects.toThrow(/holds a formula/)
    await expect(
      save(
        await arrayWorkbook(),
        [],
        [],
        values([{ row: 0, column: 0, value: 'x', arrayRef: 'B1:B3' }]),
      ),
    ).rejects.toThrow(/does not start/)
    await expect(
      save(
        await arrayWorkbook(),
        [],
        [],
        values([{ row: 0, column: 2, value: 1, arrayRef: 'C1:C3' }]),
      ),
    ).rejects.toThrow(/no array formula/)
  })
})
