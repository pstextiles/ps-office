/**
 * The parts of a workbook a built-in refresh changed (PS Office), so the
 * screen can re-read just those from the saved file instead of reopening the
 * whole workbook — the way Excel's refresh leaves the window as it is and only
 * the query tables and the formulas that use them change.
 *
 * A region is a rectangle in 0-based file coordinates. Cells are grouped by
 * row runs (rows at most GAP apart join one run) and each run spans the
 * columns its cells use: a refreshed table becomes one rectangle, scattered
 * formula results a few bands. Past MAX_RANGES a sheet's runs fold into one
 * bounding rectangle, which only costs a larger re-read.
 */

export interface RefreshRegion {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface SheetRefreshRegions {
  readonly sheet: string
  readonly ranges: readonly RefreshRegion[]
}

const GAP = 4
const MAX_RANGES = 200

interface CellAt {
  readonly row: number
  readonly column: number
}

export function refreshRegions(
  sheets: readonly { readonly sheetName: string; readonly cells: readonly CellAt[] }[],
): SheetRefreshRegions[] {
  const bySheet = new Map<string, Map<number, { min: number; max: number }>>()
  for (const { sheetName, cells } of sheets) {
    let rows = bySheet.get(sheetName)
    if (!rows) {
      rows = new Map()
      bySheet.set(sheetName, rows)
    }
    for (const { row, column } of cells) {
      const span = rows.get(row)
      if (!span) rows.set(row, { min: column, max: column })
      else {
        if (column < span.min) span.min = column
        if (column > span.max) span.max = column
      }
    }
  }
  const out: SheetRefreshRegions[] = []
  for (const [sheet, rows] of bySheet) {
    if (rows.size === 0) continue
    const sorted = [...rows.keys()].sort((a, b) => a - b)
    let ranges: RefreshRegion[] = []
    let run: { startRow: number; endRow: number; startColumn: number; endColumn: number } | null =
      null
    for (const row of sorted) {
      const span = rows.get(row)!
      if (run && row - run.endRow <= GAP) {
        run.endRow = row
        run.startColumn = Math.min(run.startColumn, span.min)
        run.endColumn = Math.max(run.endColumn, span.max)
      } else {
        if (run) ranges.push(run)
        run = { startRow: row, endRow: row, startColumn: span.min, endColumn: span.max }
      }
    }
    if (run) ranges.push(run)
    if (ranges.length > MAX_RANGES) {
      ranges = [
        {
          startRow: ranges[0]!.startRow,
          endRow: ranges[ranges.length - 1]!.endRow,
          startColumn: Math.min(...ranges.map((r) => r.startColumn)),
          endColumn: Math.max(...ranges.map((r) => r.endColumn)),
        },
      ]
    }
    out.push({ sheet, ranges })
  }
  return out
}
