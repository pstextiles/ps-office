/**
 * The Power Query formulas stored in a workbook: Excel keeps them in a
 * customXml part as base64 "DataMashup" — a small header followed by a zip
 * whose Formulas/Section1.m holds the section document and whose
 * Config/Package.xml names the culture the queries read text in.
 */
import { readZip } from './zip-read'

export interface WorkbookMashup {
  readonly section: string
  /** the workbook's Power Query culture (Config/Package.xml), e.g. "ko-KR"; null when not stored */
  readonly culture: string | null
}

export function readWorkbookMashup(workbookBytes: Uint8Array): WorkbookMashup | null {
  const zip = readZip(workbookBytes)
  for (let n = 1; n <= 64; n++) {
    const name = `customXml/item${n}.xml`
    if (!zip.has(name)) continue
    const raw = zip.bytes(name)
    if (!raw) continue
    const text = decodeXmlPart(raw)
    if (!text.includes('DataMashup')) continue
    const m = /<DataMashup\b[^>]*>([^<]+)<\/DataMashup>/.exec(text)
    if (!m) continue
    const blob = Buffer.from(m[1]!.trim(), 'base64')
    if (blob.length < 8) continue
    const partsLength = blob.readUInt32LE(4)
    const parts = readZip(blob.subarray(8, 8 + partsLength))
    const section = parts.text('Formulas/Section1.m')
    if (section === null) continue
    const packageXml = parts.text('Config/Package.xml') ?? ''
    const culture =
      /<Culture>\s*([^<\s]+)\s*<\/Culture>/.exec(packageXml.replace(/^\uFEFF/, ''))?.[1] ?? null
    return { section: section.replace(/^\uFEFF/, ''), culture }
  }
  return null
}

export function workbookSectionText(workbookBytes: Uint8Array): string | null {
  return readWorkbookMashup(workbookBytes)?.section ?? null
}

/** customXml parts are UTF-16 (with a BOM) or UTF-8 */
function decodeXmlPart(bytes: Buffer): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return bytes.subarray(2).toString('utf16le')
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return bytes.subarray(2).swap16().toString('utf16le')
  return bytes.toString('utf8')
}
