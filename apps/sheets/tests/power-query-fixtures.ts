/**
 * Workbook fixtures for the built-in Power Query engine tests: a tiny zip
 * writer (stored entries) and the DataMashup part Excel keeps queries in.
 */

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  return c >>> 0
})

function crc32(buf: Buffer): number {
  let c = 0xffffffff
  for (const byte of buf) c = CRC_TABLE[(c ^ byte) & 0xff]! ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

export function zip(files: Record<string, string | Buffer>): Buffer {
  const locals: Buffer[] = []
  const centrals: Buffer[] = []
  let offset = 0
  for (const [name, content] of Object.entries(files)) {
    const data = typeof content === 'string' ? Buffer.from(content, 'utf8') : content
    const nameBytes = Buffer.from(name, 'utf8')
    const crc = crc32(data)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBytes.length, 26)
    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(20, 4)
    central.writeUInt16LE(20, 6)
    central.writeUInt32LE(crc, 16)
    central.writeUInt32LE(data.length, 20)
    central.writeUInt32LE(data.length, 24)
    central.writeUInt16LE(nameBytes.length, 28)
    central.writeUInt32LE(offset, 42)
    locals.push(local, nameBytes, data)
    centrals.push(central, nameBytes)
    offset += 30 + nameBytes.length + data.length
  }
  const centralBytes = Buffer.concat(centrals)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(Object.keys(files).length, 8)
  end.writeUInt16LE(Object.keys(files).length, 10)
  end.writeUInt32LE(centralBytes.length, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, centralBytes, end])
}

/** customXml/item1.xml as Excel writes it: UTF-16 with a BOM, the package in base64 */
export function dataMashupPart(section: string, culture: string | null): Buffer {
  const inner = zip({
    'Formulas/Section1.m': section,
    'Config/Package.xml': `<?xml version="1.0" encoding="utf-8"?><Package><Version>2.0</Version>${
      culture === null ? '' : `<Culture>${culture}</Culture>`
    }</Package>`,
  })
  const header = Buffer.alloc(8)
  header.writeUInt32LE(inner.length, 4)
  const blob = Buffer.concat([header, inner, Buffer.alloc(4)])
  const xml = `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${blob.toString('base64')}</DataMashup>`
  return Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, 'utf16le')])
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface QueryTableWorkbook {
  /** the M expression of the query "Datos", which loads into table "Tbl" on sheet "Raw" */
  readonly query: string
  /** rows of the Raw sheet below the header row, as <row> XML; default three data rows */
  readonly rows?: string
  /** the table's range, default A1:D4 */
  readonly ref?: string
  /** extra attributes of <table> (totalsRowCount="1") */
  readonly tableAttributes?: string
  /** children between <table> and <tableColumns>, default an autoFilter over the range */
  readonly tableHead?: string
  /** <tableColumn> elements, default PO, Qty, Fecha and the calculated column Doble */
  readonly tableColumns?: string
  /** false leaves out the hidden ExternalData_1 name */
  readonly localName?: boolean
}

export const DEFAULT_ROWS = [
  '<row r="2"><c r="A2" s="1" t="inlineStr"><is><t>a</t></is></c><c r="B2" s="2"><v>1</v></c><c r="C2" s="3"><v>46023</v></c><c r="D2" s="2"><f>Tbl[[#This Row],[Qty]]*2</f><v>2</v></c></row>',
  '<row r="3"><c r="A3" s="1" t="inlineStr"><is><t>b</t></is></c><c r="B3" s="2"><v>2</v></c><c r="C3" s="3"><v>46024</v></c><c r="D3" s="2"><f>Tbl[[#This Row],[Qty]]*2</f><v>4</v></c></row>',
  '<row r="4"><c r="A4" s="4" t="inlineStr"><is><t>c</t></is></c><c r="B4" s="5"><v>3</v></c><c r="C4" s="6"><v>46025</v></c><c r="D4" s="7"><f>Tbl[[#This Row],[Qty]]*2</f><v>6</v></c></row>',
].join('')

const DEFAULT_COLUMNS =
  '<tableColumn id="1" uniqueName="1" name="PO" queryTableFieldId="1"/>' +
  '<tableColumn id="2" uniqueName="2" name="Qty" queryTableFieldId="2"/>' +
  '<tableColumn id="3" uniqueName="3" name="Fecha" queryTableFieldId="3"/>' +
  '<tableColumn id="4" uniqueName="4" name="Doble" queryTableFieldId="4"><calculatedColumnFormula>Tbl[[#This Row],[Qty]]*2</calculatedColumnFormula></tableColumn>'

/**
 * A workbook shaped like Excel's: sheet "Info", then sheet "Raw" holding the
 * query table "Tbl" (A1:D4 — PO, Qty, Fecha from the query "Datos", and the
 * calculated column Doble), its queryTable part, the connection and the
 * hidden ExternalData_1 name with the table's range.
 */
export function queryTableWorkbook(options: QueryTableWorkbook): Buffer {
  const ref = options.ref ?? 'A1:D4'
  const columns = options.tableColumns ?? DEFAULT_COLUMNS
  const header =
    '<row r="1"><c r="A1" t="inlineStr"><is><t>PO</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c><c r="C1" t="inlineStr"><is><t>Fecha</t></is></c><c r="D1" t="inlineStr"><is><t>Doble</t></is></c></row>'
  const absolute = ref.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')
  return zip({
    '[Content_Types].xml':
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    'xl/workbook.xml':
      `<workbook ${NS} ${R_NS}><sheets><sheet name="Info" sheetId="1" r:id="rId1"/><sheet name="Raw" sheetId="2" r:id="rId2"/></sheets>` +
      `<definedNames><definedName name="Otro">Info!$A$1</definedName>${
        options.localName === false
          ? ''
          : `<definedName name="ExternalData_1" localSheetId="1" hidden="1">Raw!${absolute}</definedName>`
      }</definedNames></workbook>`,
    'xl/_rels/workbook.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet ${NS}><sheetData/></worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet ${NS} ${R_NS}><sheetData>${header}${options.rows ?? DEFAULT_ROWS}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`,
    'xl/worksheets/_rels/sheet2.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`,
    'xl/tables/table1.xml':
      `<table ${NS} id="1" name="Tbl" displayName="Tbl" ref="${ref}" tableType="queryTable" totalsRowShown="0"${
        options.tableAttributes ? ` ${options.tableAttributes}` : ''
      }>${options.tableHead ?? `<autoFilter ref="${ref}"/>`}` +
      `<tableColumns count="4">${columns}</tableColumns><tableStyleInfo name="TableStyleMedium7" showRowStripes="1"/></table>`,
    'xl/tables/_rels/table1.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>`,
    'xl/queryTables/queryTable1.xml':
      `<queryTable ${NS} name="ExternalData_1" connectionId="1" autoFormatId="16"><queryTableRefresh nextId="5" unboundColumnsRight="1"><queryTableFields count="4">` +
      '<queryTableField id="1" name="PO" tableColumnId="1"/><queryTableField id="2" name="Qty" tableColumnId="2"/>' +
      '<queryTableField id="3" name="Fecha" tableColumnId="3"/><queryTableField id="4" dataBound="0" tableColumnId="4"/>' +
      '</queryTableFields></queryTableRefresh></queryTable>',
    'xl/connections.xml':
      `<connections ${NS}><connection id="1" keepAlive="1" name="Query - Datos" type="5" refreshedVersion="8" background="1" saveData="1">` +
      '<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=&quot;Datos&quot;;Extended Properties=&quot;&quot;" command="SELECT * FROM [Datos]"/></connection></connections>',
    'customXml/item1.xml': dataMashupPart(
      `section Section1;\nshared Datos = ${options.query};`,
      'en-US',
    ),
  })
}
