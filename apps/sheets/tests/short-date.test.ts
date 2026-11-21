import { describe, expect, it } from 'vitest'

import {
  DEFAULT_SHORT_DATE,
  getSystemShortDate,
  setSystemShortDate,
  shortDateNumFmtId,
  shortDatePatternForSystemLocale,
  shortDatePatternFromWindows,
} from '@genoffice/xlsx-gateway/shared/short-date'

describe('shortDatePatternForSystemLocale', () => {
  it.each([
    ['en-US', 'm/d/yyyy'],
    ['zh-CN', 'yyyy/m/d'],
    // The region decides, not the language: an English UI on a CN-region
    // machine (macOS AppleLocale en_CN) must still produce yyyy/m/d.
    ['en-CN', 'yyyy/m/d'],
    ['en-GB', 'dd/mm/yyyy'],
    ['de-DE', 'd.m.yyyy'],
    ['ja-JP', 'yyyy/m/d'],
  ])('%s → %s', (tag, pattern) => {
    expect(shortDatePatternForSystemLocale(tag)).toBe(pattern)
  })

  it('keeps only date tokens and plain separators', () => {
    // ar-SA emits RTL marks as literals and defaults to the Islamic
    // calendar; both must be neutralized.
    const pattern = shortDatePatternForSystemLocale('ar-SA')
    expect(pattern).toMatch(/^[ymd./\- ,]+$/)
    expect(pattern).toContain('yyyy')
  })

  it('falls back on invalid tags', () => {
    expect(shortDatePatternForSystemLocale('')).toBe(DEFAULT_SHORT_DATE)
  })

  it('defaults the region to US when the tag has none', () => {
    expect(shortDatePatternForSystemLocale('en')).toBe('m/d/yyyy')
  })
})

describe('shortDateNumFmtId', () => {
  it('maps the system short date (and its datetime twin) back to 14/22', () => {
    const previous = getSystemShortDate()
    try {
      setSystemShortDate('yyyy/m/d')
      expect(shortDateNumFmtId('yyyy/m/d')).toBe(14)
      expect(shortDateNumFmtId('yyyy/m/d hh:mm')).toBe(22)
      expect(shortDateNumFmtId('yyyy/m/d h:mm')).toBe(22)
      expect(shortDateNumFmtId('yyyy/mm/dd')).toBeUndefined()
    } finally {
      setSystemShortDate(previous)
    }
  })
})

describe('shortDatePatternFromWindows (PS Office)', () => {
  it.each([
    ['yyyy-MM-dd', 'yyyy-mm-dd'],
    ['yyyy-M-d', 'yyyy-m-d'],
    ['M/d/yyyy', 'm/d/yyyy'],
    ['dd/MM/yyyy', 'dd/mm/yyyy'],
    ['d/MM/yyyy', 'd/mm/yyyy'],
    ['yyyy. M. d.', 'yyyy. m. d.'],
    ['dd.MM.yy', 'dd.mm.yy'],
  ])('reads the Windows short date %s as %s', (setting, pattern) => {
    expect(shortDatePatternFromWindows(setting)).toBe(pattern)
  })

  it('refuses settings Excel would draw with names, weekdays, eras or text', () => {
    for (const setting of [
      'dd-MMM-yy',
      'ddd, M/d/yyyy',
      'gg yyyy-MM-dd',
      "yyyy'年'M'月'd'日'",
      'yyyy-MM',
      'yyyy-MM-dd-dd',
      '',
      'y/M/d',
    ]) {
      expect(shortDatePatternFromWindows(setting)).toBeNull()
    }
  })

  it('is what the save side maps back to 14 once set', () => {
    const before = getSystemShortDate()
    try {
      setSystemShortDate(shortDatePatternFromWindows('yyyy-MM-dd')!)
      expect(shortDateNumFmtId('yyyy-mm-dd')).toBe(14)
    } finally {
      setSystemShortDate(before)
    }
  })
})
