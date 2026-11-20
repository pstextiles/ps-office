import { describe, expect, it } from 'vitest'

import {
  ROOMY_FREE_BYTES,
  ROOMY_TOTAL_BYTES,
  roomForSpareRenderer,
} from '../src/main/machine-memory'

const GB = 1024 ** 3

describe('roomForSpareRenderer', () => {
  it('holds a spare only on a roomy PC with memory free right now', () => {
    expect(roomForSpareRenderer({ total: 16 * GB, free: 6 * GB })).toBe(true)
    expect(roomForSpareRenderer({ total: ROOMY_TOTAL_BYTES, free: ROOMY_FREE_BYTES })).toBe(true)
  })

  it('skips it on a 4 or 8 GB PC, however idle', () => {
    expect(roomForSpareRenderer({ total: 4 * GB, free: 3 * GB })).toBe(false)
    expect(roomForSpareRenderer({ total: 8 * GB, free: 5 * GB })).toBe(false)
  })

  it('skips it on a big PC that is short of free memory', () => {
    expect(roomForSpareRenderer({ total: 32 * GB, free: 1 * GB })).toBe(false)
  })
})
