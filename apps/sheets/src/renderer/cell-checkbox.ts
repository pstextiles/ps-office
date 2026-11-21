/**
 * Excel's in-cell checkboxes (PS Office). A cell whose format is an Excel
 * checkbox (the main process marks those formats, src/main/cell-checkbox.ts)
 * carries a flag in ICellData.custom, and here its TRUE/FALSE is drawn as a
 * box instead of text — an empty cell as an empty box, as Excel shows it.
 * A click on the box flips the value through the ordinary set-range-values
 * command, so it is journaled and saved like typing TRUE or FALSE (the cell
 * keeps its format, so Excel shows a checkbox again), undoes like an edit and
 * is refused on a protected sheet like one.
 *
 * Display only: a Style-effect interceptor adds Univer's render-time
 * properties (customRender, fontRenderExtension), which are not written to
 * the workbook, the clipboard or what the AI reads. A cell holding anything
 * other than TRUE, FALSE or nothing shows its value as usual, like Excel.
 */
import {
  CellValueType,
  HorizontalAlign,
  ICommandService,
  InterceptorEffectEnum,
  VerticalAlign,
  type ICellData,
  type IStyleData,
  type Nullable,
} from '@univerjs/core'
import { Transform, fixLineWidthByScale } from '@univerjs/engine-render'
import { INTERCEPTOR_POINT, SetRangeValuesCommand, SheetInterceptorService } from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'

/** ICellData.custom key of a cell formatted as an Excel checkbox */
export const CELL_CHECKBOX_KEY = 'xlsxCheckbox'

const MARGIN = 2
/** an unset font color: Excel's automatic text color */
const AUTOMATIC = '#000000'

/** the box's color: Excel paints a checkbox in the cell's font color (a "Good" cell's dark green, a plain cell's black) */
export function boxColor(style: Nullable<IStyleData>): string {
  const rgb = (style?.cl as { rgb?: unknown } | null | undefined)?.rgb
  return typeof rgb === 'string' && rgb !== '' ? rgb : AUTOMATIC
}

/**
 * Excel's box, drawn at the origin `size` wide: square corners; unchecked a
 * white square with a thin border, checked filled with a white tick, both in
 * the font color.
 */
function drawBox(
  ctx: CanvasRenderingContext2D,
  size: number,
  checked: boolean,
  color: string,
  pixel: number,
): void {
  const line = Math.max(pixel, size / 14)
  if (checked) {
    ctx.fillStyle = color
    ctx.fillRect(0, 0, size, size)
    ctx.strokeStyle = '#ffffff'
    ctx.lineWidth = Math.max(pixel, size / 8)
    ctx.lineCap = 'square'
    ctx.lineJoin = 'miter'
    ctx.beginPath()
    ctx.moveTo(size * 0.24, size * 0.52)
    ctx.lineTo(size * 0.42, size * 0.7)
    ctx.lineTo(size * 0.77, size * 0.3)
    ctx.stroke()
    return
  }
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, size, size)
  ctx.strokeStyle = color
  ctx.lineWidth = line
  ctx.strokeRect(line / 2, line / 2, size - line, size - line)
}

/** the cell's value as a checkbox: checked, unchecked, or null when it is something else */
export function checkboxState(cell: Nullable<ICellData>): boolean | null {
  // a formula cell shows its result, and Excel draws a box for a TRUE/FALSE result too
  if (!cell) return false
  const { v, t } = cell
  if (v === null || v === undefined || v === '') return false
  if (t === CellValueType.BOOLEAN)
    return typeof v === 'number' ? v !== 0 : String(v).toUpperCase() === 'TRUE' || v === '1'
  if (typeof v === 'boolean') return v
  return null
}

export function isCheckboxCell(cell: Nullable<ICellData>): boolean {
  return (cell?.custom as Record<string, unknown> | undefined)?.[CELL_CHECKBOX_KEY] === true
}

interface Bounds {
  readonly startX: number
  readonly startY: number
  readonly endX: number
  readonly endY: number
}

/** where the box sits in a cell: its size follows the font, its place the alignment (Univer's own checkbox layout) */
export function boxLayout(
  cell: Bounds,
  style: Nullable<IStyleData>,
): { left: number; top: number; size: number } {
  const size = (style?.fs ?? 11) * 1.3
  const width = cell.endX - cell.startX - MARGIN * 2
  const height = cell.endY - cell.startY
  const top =
    style?.vt === VerticalAlign.TOP
      ? 0
      : style?.vt === VerticalAlign.MIDDLE
        ? (height - size) / 2
        : height - size - MARGIN
  const left =
    style?.ht === HorizontalAlign.LEFT
      ? MARGIN
      : style?.ht === HorizontalAlign.RIGHT
        ? MARGIN + width - size
        : MARGIN + (width - size) / 2
  return { left: cell.startX + left, top: cell.startY + Math.max(0, top), size }
}

interface RenderInfo {
  readonly primaryWithCoord: Bounds & {
    readonly isMergedMainCell?: boolean
    readonly mergeInfo?: Bounds
    readonly actualRow: number
    readonly actualColumn: number
  }
  readonly style: Nullable<IStyleData>
  readonly worksheet: { getCellRaw(row: number, column: number): Nullable<ICellData> }
  readonly row: number
  readonly col: number
}

const boundsOf = (info: RenderInfo): Bounds =>
  info.primaryWithCoord.isMergedMainCell && info.primaryWithCoord.mergeInfo
    ? info.primaryWithCoord.mergeInfo
    : info.primaryWithCoord

export function installCellCheckboxes(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const render = {
    zIndex: 1,
    drawWith(
      ctx: CanvasRenderingContext2D & { getTransform(): DOMMatrix },
      info: RenderInfo,
    ): void {
      const state = checkboxState(info.worksheet.getCellRaw(info.row, info.col))
      if (state === null) return
      const bounds = boundsOf(info)
      const { left, top, size } = boxLayout(bounds, info.style)
      const { a: scaleX, d: scaleY } = ctx.getTransform()
      const m = Transform.create()
        .composeMatrix({
          left: fixLineWidthByScale(left, scaleX),
          top: fixLineWidthByScale(top, scaleY),
          scaleX: 1,
          scaleY: 1,
          angle: 0,
          skewX: 0,
          skewY: 0,
          flipX: false,
          flipY: false,
        })
        .getMatrix()
      ctx.save()
      ctx.beginPath()
      ctx.rect(
        bounds.startX,
        bounds.startY,
        bounds.endX - bounds.startX,
        bounds.endY - bounds.startY,
      )
      ctx.clip()
      ctx.transform(m[0]!, m[1]!, m[2]!, m[3]!, m[4]!, m[5]!)
      drawBox(ctx, size, state, boxColor(info.style), 1 / Math.max(scaleX, 0.01))
      ctx.restore()
    },
    isHit(evt: { x: number; y: number }, info: RenderInfo): boolean {
      if (checkboxState(info.worksheet.getCellRaw(info.row, info.col)) === null) return false
      const { left, top, size } = boxLayout(boundsOf(info), info.style)
      return evt.x >= left && evt.x <= left + size && evt.y >= top && evt.y <= top + size
    },
    onPointerDown(info: RenderInfo, evt: { button?: number }): void {
      if (evt.button === 2) return
      const raw = info.worksheet.getCellRaw(info.row, info.col)
      const state = checkboxState(raw)
      // a formula decides its own result; only a stored value flips
      if (state === null || raw?.f || raw?.si) return
      const row = info.primaryWithCoord.actualRow
      const column = info.primaryWithCoord.actualColumn
      void commandService.executeCommand(SetRangeValuesCommand.id, {
        range: { startRow: row, endRow: row, startColumn: column, endColumn: column },
        value: { v: state ? 0 : 1, t: CellValueType.BOOLEAN, p: null },
      })
    },
  }
  const interceptorService = injector.get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    effect: InterceptorEffectEnum.Style,
    // after the value and style fixes, before nothing that cares about the text
    priority: 9.2,
    handler: (cell, location, next) => {
      const raw = location.rawData
      if (!isCheckboxCell(raw) || checkboxState(cell ?? raw) === null) return next(cell)
      const base = !cell || cell === raw ? { ...raw } : cell
      return next({
        ...base,
        customRender: [...(base.customRender ?? []), render as never],
        fontRenderExtension: { ...base.fontRenderExtension, isSkip: true },
      })
    },
  })
}
