/**
 * PS Office: a large workbook's sort, computed with Excel's sort order and
 * applied as one reorder.
 */
import { CellValueType, type ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  compareSortKeys,
  compareSortText,
  inverseOrder,
  sortKeyOf,
  sortOrder,
} from '../src/renderer/value-mode-sort'

const sortTexts = (texts: string[]): string[] => [...texts].sort(compareSortText)

describe("Excel's text order", () => {
  it('puts digits, then the space and punctuation, then letters, ignoring case', () => {
    expect(sortTexts(['b', 'A', '_x', '1', ' z', 'a', '#1'])).toEqual([
      '1',
      ' z',
      '#1',
      '_x',
      'A',
      'a',
      'b',
    ])
  })

  it('ignores hyphens and apostrophes, and puts the hyphenated twin last', () => {
    expect(sortTexts(['GT-003', 'GT002', 'GT-002'])).toEqual(['GT002', 'GT-002', 'GT-003'])
    expect(compareSortText("O'Neil", 'ONeil')).toBe(0)
  })

  it('sorts accented letters with their base letter and Hangul after Latin', () => {
    expect(sortTexts(['b', 'é', 'a', '가', 'f'])).toEqual(['a', 'b', 'é', 'f', '가'])
    expect(sortTexts(['다', '가', '나'])).toEqual(['가', '나', '다'])
  })
})

describe('kinds of values', () => {
  const key = (cell: ICellData | null) => sortKeyOf(cell)

  it('orders number < text < logical < error ascending, and keeps blanks last both ways', () => {
    const cells: (ICellData | null)[] = [
      { v: '#N/A', t: CellValueType.STRING },
      null,
      { v: 1, t: CellValueType.BOOLEAN },
      { v: 'abc', t: CellValueType.STRING },
      { v: 42, t: CellValueType.NUMBER },
      { v: '' },
    ]
    const asc = [...cells].sort((a, b) => compareSortKeys(key(a), key(b), 'asc'))
    expect(asc.map((c) => (c?.v === '' ? null : (c?.v ?? null)))).toEqual([
      42,
      'abc',
      1,
      '#N/A',
      null,
      null,
    ])
    const desc = [...cells].sort((a, b) => compareSortKeys(key(a), key(b), 'desc'))
    expect(desc.map((c) => (c?.v === '' ? null : (c?.v ?? null)))).toEqual([
      '#N/A',
      1,
      'abc',
      42,
      null,
      null,
    ])
  })

  it('reads a rich-text cell by its text', () => {
    expect(sortKeyOf({ p: { body: { dataStream: 'Castlerock\r\n' } } } as ICellData)).toEqual({
      kind: 1,
      text: 'Castlerock',
    })
  })
})

describe('the reorder', () => {
  // rows 5 (header) .. 9, column 0 holds the key; row 7 is filtered out
  const values: Record<number, ICellData> = {
    6: { v: 3 },
    7: { v: 0 },
    8: { v: 1 },
    9: { v: 3 },
  }
  const worksheet = {
    getSheet: () => ({
      getMergeData: () => [],
      getRowFiltered: (row: number) => row === 7,
      getRowRawVisible: () => true,
      getCellRaw: (row: number) => values[row] ?? null,
    }),
  } as never

  it('moves the visible rows by Excel order, ties in place, filtered rows untouched', () => {
    const order = sortOrder(
      worksheet,
      { startRow: 5, endRow: 9, startColumn: 0, endColumn: 3 },
      [{ colIndex: 0, type: 'asc' }],
      true,
    )
    // target ← source: the 1 (row 8) comes first, then the two 3s in their order
    expect(order).toEqual({ 6: 8, 8: 6, 9: 9 })
    expect(inverseOrder(order)).toEqual({ 8: 6, 6: 8, 9: 9 })
  })
})
