/**
 * A small synchronous zip reader (stored and deflated entries), enough for
 * the parts of an .xlsx. The query engine evaluates synchronously, so it
 * cannot use the promise-based zip libraries.
 */
import { inflateRawSync } from 'node:zlib'
import { MError } from './values'

export interface ZipEntries {
  has(name: string): boolean
  bytes(name: string): Buffer | null
  text(name: string): string | null
}

const MAX_ENTRY_BYTES = 512 * 1024 * 1024

function notAWorkbook(): MError {
  return new MError('DataFormat.Error', 'External table is not in the expected format.')
}

export function readZip(bytes: Uint8Array): ZipEntries {
  const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // end of central directory: signature 0x06054b50 within the last 64 KB + 22
  let eocd = -1
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65_557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i
      break
    }
  }
  if (eocd < 0) throw notAWorkbook()
  let count = buf.readUInt16LE(eocd + 10)
  let dirOffset = buf.readUInt32LE(eocd + 16)
  // zip64: the real figures sit in the zip64 end record
  if (dirOffset === 0xffffffff || count === 0xffff) {
    const locator = eocd - 20
    if (locator >= 0 && buf.readUInt32LE(locator) === 0x07064b50) {
      const rec = Number(buf.readBigUInt64LE(locator + 8))
      if (buf.readUInt32LE(rec) === 0x06064b50) {
        count = Number(buf.readBigUInt64LE(rec + 32))
        dirOffset = Number(buf.readBigUInt64LE(rec + 48))
      }
    }
  }
  const entries = new Map<
    string,
    { method: number; compressed: number; size: number; local: number }
  >()
  let p = dirOffset
  for (let n = 0; n < count; n++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) throw notAWorkbook()
    const method = buf.readUInt16LE(p + 10)
    let compressed = buf.readUInt32LE(p + 20)
    let size = buf.readUInt32LE(p + 24)
    const nameLen = buf.readUInt16LE(p + 28)
    const extraLen = buf.readUInt16LE(p + 30)
    const commentLen = buf.readUInt16LE(p + 32)
    let local = buf.readUInt32LE(p + 42)
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen)
    // zip64 extra field (id 1) carries the 64-bit sizes and offset
    let e = p + 46 + nameLen
    const extraEnd = e + extraLen
    while (e + 4 <= extraEnd) {
      const id = buf.readUInt16LE(e)
      const len = buf.readUInt16LE(e + 2)
      if (id === 1) {
        let q = e + 4
        if (size === 0xffffffff) ((size = Number(buf.readBigUInt64LE(q))), (q += 8))
        if (compressed === 0xffffffff) ((compressed = Number(buf.readBigUInt64LE(q))), (q += 8))
        if (local === 0xffffffff) local = Number(buf.readBigUInt64LE(q))
      }
      e += 4 + len
    }
    entries.set(name, { method, compressed, size, local })
    p = extraEnd + commentLen
  }
  const read = (name: string): Buffer | null => {
    const entry = entries.get(name)
    if (!entry) return null
    if (entry.size > MAX_ENTRY_BYTES)
      throw new MError('DataFormat.Error', `The part ${name} is too large to read.`)
    const lh = entry.local
    if (buf.readUInt32LE(lh) !== 0x04034b50) throw notAWorkbook()
    const start = lh + 30 + buf.readUInt16LE(lh + 26) + buf.readUInt16LE(lh + 28)
    const data = buf.subarray(start, start + entry.compressed)
    if (entry.method === 0) return data
    if (entry.method === 8) return inflateRawSync(data, { maxOutputLength: MAX_ENTRY_BYTES })
    throw new MError('DataFormat.Error', `Unsupported compression in ${name}.`)
  }
  return {
    has: (name) => entries.has(name),
    bytes: (name) => read(name),
    text: (name) => {
      const data = read(name)
      return data === null ? null : data.toString('utf8')
    },
  }
}
