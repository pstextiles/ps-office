/**
 * The refresh-only run keeps the workbook's own recalculate-on-open setting
 * when every formula was recalculated: an older Excel turns XLOOKUP, LET and
 * dynamic arrays into #NAME? on a forced recalculation.
 */
import { describe, expect, it } from 'vitest'

import {
  fullCalcOnLoadOf,
  isWorkbookPart,
  withSourceFullCalcOnLoad,
} from '../src/main/calc-on-load'

const WB = (calcPr: string): string =>
  `<workbook><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`

describe('withSourceFullCalcOnLoad', () => {
  it('drops the calcPr the save added to a workbook that had none', () => {
    expect(withSourceFullCalcOnLoad(WB('<calcPr fullCalcOnLoad="1"/>'), WB(''))).toBe(WB(''))
  })

  it('takes the forced flag off and keeps the other calculation settings', () => {
    const source = WB('<calcPr calcId="191029"/>')
    expect(
      withSourceFullCalcOnLoad(WB('<calcPr calcId="191029" fullCalcOnLoad="1"/>'), source),
    ).toBe(source)
  })

  it('keeps a flag the workbook had itself', () => {
    const source = WB('<calcPr calcId="191029" fullCalcOnLoad="1"/>')
    expect(
      withSourceFullCalcOnLoad(WB('<calcPr calcId="191029" fullCalcOnLoad="1"/>'), source),
    ).toBe(source)
    expect(fullCalcOnLoadOf(source)).toBe('1')
    expect(fullCalcOnLoadOf(WB('<calcPr calcId="191029"/>'))).toBeNull()
    expect(fullCalcOnLoadOf(WB(''))).toBeNull()
  })

  it('finds the workbook part, not its relationships', () => {
    expect(isWorkbookPart('xl/workbook.xml')).toBe(true)
    expect(isWorkbookPart('xl/_rels/workbook.xml.rels')).toBe(false)
    expect(isWorkbookPart('xl/worksheets/sheet1.xml')).toBe(false)
  })
})
