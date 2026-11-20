import { readFileSync, statSync } from 'node:fs'
import { nativeImage } from 'electron'

/**
 * A picture attached to an AI message, made fit to send (PS Office).
 *
 * Phone photos are often 3 to 8 MB, over the 5 MB a message picture may be,
 * and were refused outright. A picture that is too heavy or larger than the
 * model looks at is now shrunk to MODEL_IMAGE_LONG_EDGE on its long side and
 * re-encoded: the vendors downscale to about that size anyway, so a photo of a
 * table reads the same, and a smaller payload is lighter on a slow PC.
 */
export const MODEL_IMAGE_MAX_BYTES = 5 * 1024 * 1024
export const MODEL_IMAGE_LONG_EDGE = 2000
/** decoding a bigger file is not worth the memory on a low-end PC */
export const MODEL_IMAGE_SOURCE_MAX_BYTES = 50 * 1024 * 1024

export type ModelImage =
  | { readonly ok: true; readonly base64: string; readonly mime: string }
  | { readonly ok: false; readonly reason: 'too-large' | 'unreadable' }

interface Decoder {
  createFromBuffer(buffer: Buffer): {
    isEmpty(): boolean
    getSize(): { width: number; height: number }
    resize(options: { width?: number; height?: number; quality?: 'good' | 'better' | 'best' }): {
      toJPEG(quality: number): Buffer
      toPNG(): Buffer
    }
  }
}

/** Electron's decoder; absent (or a throwing stand-in) where Electron is mocked away. */
function electronDecoder(): Decoder | undefined {
  try {
    return nativeImage as unknown as Decoder | undefined
  } catch {
    return undefined
  }
}

/** Reads the picture at `path` (of type `mime`) and returns it ready for a model message. */
export function readImageForModel(
  path: string,
  mime: string,
  decoder: Decoder | undefined = electronDecoder(),
): ModelImage {
  let bytes: Buffer
  try {
    const size = statSync(path).size
    if (size > MODEL_IMAGE_SOURCE_MAX_BYTES) return { ok: false, reason: 'too-large' }
    bytes = readFileSync(path)
  } catch {
    return { ok: false, reason: 'unreadable' }
  }
  const asIs: ModelImage =
    bytes.byteLength <= MODEL_IMAGE_MAX_BYTES
      ? { ok: true, base64: bytes.toString('base64'), mime }
      : { ok: false, reason: 'too-large' }
  // Electron decodes PNG, JPEG and BMP; anything else (GIF, WebP) goes as it is
  if (typeof decoder?.createFromBuffer !== 'function' || !/^image\/(png|jpeg|bmp)$/.test(mime)) {
    return asIs
  }
  const image = decoder.createFromBuffer(bytes)
  if (image.isEmpty()) return asIs
  const { width, height } = image.getSize()
  const longEdge = Math.max(width, height)
  if (longEdge <= MODEL_IMAGE_LONG_EDGE && asIs.ok) return asIs
  const scale = Math.min(1, MODEL_IMAGE_LONG_EDGE / longEdge)
  const resized = image.resize(
    width >= height
      ? { width: Math.max(1, Math.round(width * scale)), quality: 'good' }
      : { height: Math.max(1, Math.round(height * scale)), quality: 'good' },
  )
  // a screenshot keeps its sharp text as PNG when that fits; photos go as JPEG
  if (mime === 'image/png') {
    const png = resized.toPNG()
    if (png.byteLength <= MODEL_IMAGE_MAX_BYTES) {
      return { ok: true, base64: png.toString('base64'), mime: 'image/png' }
    }
  }
  for (const quality of [88, 75, 60]) {
    const jpeg = resized.toJPEG(quality)
    if (jpeg.byteLength <= MODEL_IMAGE_MAX_BYTES) {
      return { ok: true, base64: jpeg.toString('base64'), mime: 'image/jpeg' }
    }
  }
  return asIs
}
