import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({ nativeImage: undefined }))

import {
  MODEL_IMAGE_LONG_EDGE,
  MODEL_IMAGE_MAX_BYTES,
  MODEL_IMAGE_SOURCE_MAX_BYTES,
  readImageForModel,
} from '../src/model-image'

const dir = mkdtempSync(join(tmpdir(), 'model-image-'))
afterAll(() => rmSync(dir, { recursive: true, force: true }))

function file(name: string, bytes: number): string {
  const path = join(dir, name)
  writeFileSync(path, Buffer.alloc(bytes, 1))
  return path
}

/** A decoder standing in for Electron's nativeImage: records the resize asked for. */
function fakeDecoder(
  size: { width: number; height: number },
  encoded: { jpeg: number; png: number },
) {
  const resizes: Array<{ width?: number; height?: number }> = []
  const qualities: number[] = []
  const decoder = {
    createFromBuffer: () => ({
      isEmpty: () => false,
      getSize: () => size,
      resize: (options: { width?: number; height?: number }) => {
        resizes.push(options)
        return {
          toJPEG: (quality: number) => {
            qualities.push(quality)
            return Buffer.alloc(quality === 60 ? Math.min(encoded.jpeg, 1000) : encoded.jpeg, 2)
          },
          toPNG: () => Buffer.alloc(encoded.png, 3),
        }
      },
    }),
  }
  return { decoder, resizes, qualities }
}

describe('readImageForModel', () => {
  it('sends a small picture as it is', () => {
    const path = file('small.jpg', 1000)
    const { decoder, resizes } = fakeDecoder({ width: 800, height: 600 }, { jpeg: 10, png: 10 })
    const result = readImageForModel(path, 'image/jpeg', decoder)
    expect(result).toEqual({
      ok: true,
      base64: Buffer.alloc(1000, 1).toString('base64'),
      mime: 'image/jpeg',
    })
    expect(resizes).toHaveLength(0)
  })

  it('shrinks a heavy phone photo to the long-edge limit as JPEG', () => {
    const path = file('phone.jpg', MODEL_IMAGE_MAX_BYTES + 1)
    const { decoder, resizes, qualities } = fakeDecoder(
      { width: 4000, height: 3000 },
      { jpeg: 800_000, png: 9_000_000 },
    )
    const result = readImageForModel(path, 'image/jpeg', decoder)
    expect(result).toMatchObject({ ok: true, mime: 'image/jpeg' })
    expect(resizes).toEqual([{ width: MODEL_IMAGE_LONG_EDGE, quality: 'good' }])
    expect(qualities).toEqual([88])
  })

  it('scales a portrait photo by its height', () => {
    const path = file('portrait.jpg', 2_000_000)
    const { decoder, resizes } = fakeDecoder(
      { width: 3000, height: 4000 },
      { jpeg: 500_000, png: 1 },
    )
    readImageForModel(path, 'image/jpeg', decoder)
    expect(resizes).toEqual([{ height: MODEL_IMAGE_LONG_EDGE, quality: 'good' }])
  })

  it('keeps a large screenshot as PNG when that fits', () => {
    const path = file('screen.png', 3_000_000)
    const { decoder } = fakeDecoder(
      { width: 3840, height: 2160 },
      { jpeg: 400_000, png: 1_500_000 },
    )
    expect(readImageForModel(path, 'image/png', decoder)).toMatchObject({
      ok: true,
      mime: 'image/png',
    })
  })

  it('lowers the JPEG quality until it fits', () => {
    const path = file('dense.jpg', MODEL_IMAGE_MAX_BYTES + 10)
    const { decoder, qualities } = fakeDecoder(
      { width: 6000, height: 4000 },
      { jpeg: MODEL_IMAGE_MAX_BYTES + 1, png: 1 },
    )
    expect(readImageForModel(path, 'image/jpeg', decoder)).toMatchObject({ ok: true })
    expect(qualities).toEqual([88, 75, 60])
  })

  it('refuses what it cannot decode when it is too heavy to send', () => {
    const path = file('anim.gif', MODEL_IMAGE_MAX_BYTES + 1)
    const { decoder } = fakeDecoder({ width: 10, height: 10 }, { jpeg: 1, png: 1 })
    expect(readImageForModel(path, 'image/gif', decoder)).toEqual({
      ok: false,
      reason: 'too-large',
    })
  })

  it('falls back to the size check where Electron has no decoder', () => {
    expect(readImageForModel(file('a.png', 100), 'image/png')).toMatchObject({ ok: true })
    expect(readImageForModel(file('b.png', MODEL_IMAGE_MAX_BYTES + 1), 'image/png')).toEqual({
      ok: false,
      reason: 'too-large',
    })
  })

  it('reports a missing file as unreadable and a huge one as too large', () => {
    expect(readImageForModel(join(dir, 'gone.jpg'), 'image/jpeg')).toEqual({
      ok: false,
      reason: 'unreadable',
    })
    expect(MODEL_IMAGE_SOURCE_MAX_BYTES).toBeGreaterThan(MODEL_IMAGE_MAX_BYTES)
  })
})
