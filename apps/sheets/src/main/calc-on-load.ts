/**
 * PS Office: the refresh-only run (`--refresh-queries`) is meant for PCs with
 * an older Excel. Every save asks Excel to recalculate on open
 * (`<calcPr fullCalcOnLoad="1">`), which an Excel without XLOOKUP, LET or
 * dynamic arrays turns into #NAME? in every cell that uses them. When the
 * built-in calculation recomputed every formula exactly as Excel does, the
 * saved results are already the ones Excel would compute, so that run keeps
 * the workbook's own setting instead (what the file had before the save).
 */

const FULL_CALC = /\s*\bfullCalcOnLoad="[^"]*"/

/** the workbook part's `fullCalcOnLoad` value, or null when it has none */
export function fullCalcOnLoadOf(workbookXml: string): string | null {
  const calcPr = /<calcPr\b[^>]*>/.exec(workbookXml)?.[0]
  if (!calcPr) return null
  return /\bfullCalcOnLoad="([^"]*)"/.exec(calcPr)?.[1] ?? null
}

/**
 * `planned` (the workbook part a save wrote) with the `fullCalcOnLoad` the
 * source part had: the attribute put back as it was, or dropped; a bare
 * `<calcPr/>` the save added to a part that had no calcPr is dropped too.
 */
export function withSourceFullCalcOnLoad(planned: string, source: string): string {
  const sourceValue = fullCalcOnLoadOf(source)
  const sourceHadCalcPr = /<calcPr\b/.test(source)
  return planned.replace(/<calcPr\b([^>]*?)(\/?>)/, (_full, attributes: string, close: string) => {
    const cleaned = attributes.replace(FULL_CALC, '')
    if (sourceValue !== null) return `<calcPr${cleaned} fullCalcOnLoad="${sourceValue}"${close}`
    if (!sourceHadCalcPr && cleaned.trim() === '' && close === '/>') return ''
    return `<calcPr${cleaned}${close}`
  })
}

/** the workbook part among a save's rewritten entries (`xl/workbook.xml` by convention) */
export function isWorkbookPart(entry: string): boolean {
  return /(^|\/)workbook\d*\.xml$/i.test(entry) && !/(^|\/)_rels\//.test(entry)
}
