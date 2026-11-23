import { describe, expect, it } from 'vitest'

import {
  drawExcelFilterButton,
  filterButtonArea,
  filterButtonColors,
  filterButtonUnits,
  patchFilterButtonShapes,
} from '../src/renderer/filter-button-excel-style'

function recordingContext(): { ctx: CanvasRenderingContext2D; calls: string[] } {
  const calls: string[] = []
  const handler: ProxyHandler<object> = {
    get: (_target, name) => {
      if (typeof name !== 'string') return undefined
      return (...args: unknown[]) => {
        calls.push(`${name}(${args.map(String).join(',')})`)
      }
    },
    set: (_target, name, value) => {
      calls.push(`${String(name)}=${String(value)}`)
      return true
    },
  }
  return { ctx: new Proxy({}, handler) as CanvasRenderingContext2D, calls }
}

const theme = { getColorFromTheme: (key: string) => `theme(${key})` }

describe('drawExcelFilterButton', () => {
  it('draws a bordered box with one arrow when the column has no criteria', () => {
    const { ctx, calls } = recordingContext()
    drawExcelFilterButton(
      ctx,
      16,
      { hasCriteria: false, hovered: false },
      filterButtonColors(theme, false),
    )
    expect(calls.filter((c) => c.startsWith('roundRect'))).toHaveLength(1)
    expect(calls).toContain('strokeStyle=theme(gray.300)')
    expect(calls.filter((c) => c === 'closePath()')).toHaveLength(1)
  })

  it('adds the funnel beside a smaller arrow once criteria apply', () => {
    const { ctx, calls } = recordingContext()
    drawExcelFilterButton(
      ctx,
      16,
      { hasCriteria: true, hovered: true },
      filterButtonColors(theme, true),
    )
    expect(calls.filter((c) => c === 'closePath()')).toHaveLength(2)
    expect(calls).toContain('strokeStyle=theme(primary.600)')
    expect(calls).toContain('fillStyle=theme(gray.50)')
  })
})

describe('patchFilterButtonShapes', () => {
  it('replaces _draw on the shared prototype and ignores other scene objects', () => {
    class FakeButton {
      _cellWidth = 80
      _cellHeight = 20
      _hovered = false
      _filterParams = { hasCriteria: false }
      _themeService = theme
      _draw(_ctx: CanvasRenderingContext2D): void {
        throw new Error('stock draw must not run')
      }
    }
    const button = new FakeButton()
    expect(patchFilterButtonShapes([{ _draw: () => undefined }, button])).toBe(true)
    const { ctx, calls } = recordingContext()
    globalThis.Path2D ??= class {
      rect(): void {}
    } as unknown as typeof Path2D
    new FakeButton()._draw(ctx)
    expect(calls.some((c) => c.startsWith('clip('))).toBe(true)
    expect(calls.filter((c) => c.startsWith('roundRect'))).toHaveLength(1)
    expect(patchFilterButtonShapes([])).toBe(false)
  })
})

describe('filter button size when zoomed out (PS Office)', () => {
  it('stays 16 screen pixels like Excel instead of shrinking with the zoom', () => {
    expect(filterButtonUnits(1)).toBe(16)
    expect(filterButtonUnits(2)).toBe(16)
    expect(filterButtonUnits(0.5)).toBe(32)
    expect(filterButtonUnits(0.5) * 0.5).toBe(16)
    expect(filterButtonUnits(0)).toBe(16)
  })

  it('grows up and left from the bottom-right corner, never past the cell', () => {
    expect(filterButtonArea(32, 80, 40)).toEqual({ left: -16, top: -16, size: 32 })
    // a 20-unit-high cell: the button is cut at the cell's top, like its drawing
    expect(filterButtonArea(32, 80, 20)).toEqual({ left: -16, top: -4, size: 32 })
    expect(filterButtonArea(16, 80, 40)).toEqual({ left: 0, top: 0, size: 16 })
  })

  it('takes clicks on the enlarged button at the scene zoom', () => {
    class ZoomedButton {
      _cellWidth = 80
      _cellHeight = 40
      _hovered = false
      _filterParams = { hasCriteria: false }
      _themeService = theme
      _draw(): void {}
      isHit(_coord: unknown): boolean {
        return false
      }
      getInverseCoord(coord: { x: number; y: number }): { x: number; y: number } {
        return coord
      }
      getScene(): { getAncestorScale(): { scaleX: number } } {
        return { getAncestorScale: () => ({ scaleX: 0.5 }) }
      }
    }
    expect(patchFilterButtonShapes([new ZoomedButton()])).toBe(true)
    const button = new ZoomedButton()
    // the arrow's middle at 50%: 32 units from the corner, outside the 16-unit box
    expect(button.isHit({ x: 0, y: 0 } as never)).toBe(true)
    expect(button.isHit({ x: -15, y: -15 } as never)).toBe(true)
    expect(button.isHit({ x: -17, y: 8 } as never)).toBe(false)
    expect(button.isHit({ x: 17, y: 8 } as never)).toBe(false)
  })
})
