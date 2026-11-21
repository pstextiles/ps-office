/**
 * Excel's in-cell checkboxes (PS Office): which cell formats are checkboxes,
 * read from styles.xml and the feature property bags the way Excel writes
 * them (the bag part below is copied from a real workbook), and how the
 * screen reads a cell's value as a box.
 */
import { CellValueType, HorizontalAlign, VerticalAlign } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  checkboxStyleIndexes,
  featurePropertyBagPart,
  parseFeaturePropertyBags,
} from '../src/main/cell-checkbox'
import {
  CELL_CHECKBOX_KEY,
  boxLayout,
  checkboxState,
  isCheckboxCell,
} from '../src/renderer/cell-checkbox'

const BAGS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag">' +
  '<bag type="Checkbox"/><bag type="XFControls"><bagId k="CellControl">0</bagId></bag>' +
  '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>' +
  '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag>' +
  '</FeaturePropertyBags>'

const EXT = (i: number): string =>
  `<extLst><ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><xfpb:xfComplement i="${i}"/></ext></extLst>`

const STYLES =
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="4">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/>${EXT(0)}</xf>` +
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0">${EXT(0)}</xf>` +
  '</cellXfs></styleSheet>'

describe('finding checkbox formats in a workbook', () => {
  it('follows xfComplement → XFComplement → XFControls → CellControl → Checkbox', () => {
    expect(checkboxStyleIndexes(STYLES, BAGS)).toEqual([1, 3])
  })

  it('reads the bags in order, with their ids and lists', () => {
    const bags = parseFeaturePropertyBags(BAGS)
    expect(bags.map((b) => b.type)).toEqual([
      'Checkbox',
      'XFControls',
      'XFComplement',
      'XFComplements',
    ])
    expect(bags[1]!.ids.get('CellControl')).toBe(0)
    expect(bags[3]!.lists.get('MappedFeaturePropertyBags')).toEqual([2])
  })

  it('ignores complements that lead to another feature, or to nothing', () => {
    const other = BAGS.replace('<bag type="Checkbox"/>', '<bag type="Toggle"/>')
    expect(checkboxStyleIndexes(STYLES, other)).toEqual([])
    expect(checkboxStyleIndexes(STYLES.replaceAll('i="0"', 'i="5"'), BAGS)).toEqual([])
    expect(checkboxStyleIndexes(STYLES, '')).toEqual([])
    expect(checkboxStyleIndexes('<styleSheet/>', BAGS)).toEqual([])
  })

  it('finds the bag part from the workbook relationships', () => {
    const rels = (target: string): string =>
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      `<Relationship Id="rId47" Type="http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag" Target="${target}"/>` +
      '</Relationships>'
    expect(featurePropertyBagPart(rels('featurePropertyBag/featurePropertyBag.xml'))).toBe(
      'xl/featurePropertyBag/featurePropertyBag.xml',
    )
    expect(featurePropertyBagPart(rels('/xl/fpb/bags.xml'))).toBe('xl/fpb/bags.xml')
    expect(featurePropertyBagPart('<Relationships/>')).toBeNull()
  })
})

describe('a checkbox cell on screen', () => {
  it('is a cell whose format carries the flag', () => {
    expect(isCheckboxCell({ v: 1, custom: { [CELL_CHECKBOX_KEY]: true } })).toBe(true)
    expect(isCheckboxCell({ v: 1 })).toBe(false)
    expect(isCheckboxCell(null)).toBe(false)
  })

  it('reads TRUE, FALSE and an empty cell as a box, anything else as its value', () => {
    expect(checkboxState({ v: 1, t: CellValueType.BOOLEAN })).toBe(true)
    expect(checkboxState({ v: 0, t: CellValueType.BOOLEAN })).toBe(false)
    expect(checkboxState({ v: 'TRUE', t: CellValueType.BOOLEAN })).toBe(true)
    expect(checkboxState({ v: true })).toBe(true)
    expect(checkboxState({})).toBe(false)
    expect(checkboxState(null)).toBe(false)
    expect(checkboxState({ v: 'Pendiente', t: CellValueType.STRING })).toBeNull()
    expect(checkboxState({ v: 12 })).toBeNull()
    // a formula's TRUE/FALSE result is drawn too
    expect(checkboxState({ f: '=A1>0', v: 1, t: CellValueType.BOOLEAN })).toBe(true)
  })

  it('places the box by the cell alignment, sized by the font', () => {
    const cell = { startX: 0, startY: 0, endX: 104, endY: 20 }
    expect(
      boxLayout(cell, { ht: HorizontalAlign.CENTER, fs: 10, vt: VerticalAlign.MIDDLE }),
    ).toEqual({
      left: 2 + (100 - 13) / 2,
      top: (20 - 13) / 2,
      size: 13,
    })
    expect(
      boxLayout(cell, { ht: HorizontalAlign.LEFT, fs: 10, vt: VerticalAlign.TOP }),
    ).toMatchObject({ left: 2, top: 0 })
    // Excel's default: bottom
    expect(boxLayout({ startX: 10, startY: 40, endX: 114, endY: 60 }, { fs: 10 })).toMatchObject({
      top: 40 + 20 - 13 - 2,
    })
  })
})
