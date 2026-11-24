/**
 * PS Office: a bulk load draws the grid a few times a second, and right away
 * after the person scrolls, clicks or types.
 */
import { describe, expect, it, vi } from 'vitest'

import {
  INPUT_GRACE_MS,
  THROTTLED_FRAME_MS,
  shouldDrawFrame,
  throttleGridRendering,
} from '../src/renderer/render-throttle'
import type { UniverRuntime } from '../src/renderer/univer-state'

describe('which frames draw while throttled', () => {
  it('draws once per interval, and every frame shortly after an input', () => {
    expect(shouldDrawFrame(1000, -Infinity, -Infinity)).toBe(true)
    expect(shouldDrawFrame(1000 + THROTTLED_FRAME_MS - 1, 1000, -Infinity)).toBe(false)
    expect(shouldDrawFrame(1000 + THROTTLED_FRAME_MS, 1000, -Infinity)).toBe(true)
    expect(shouldDrawFrame(1100, 1000, 1100 - INPUT_GRACE_MS)).toBe(true)
    expect(shouldDrawFrame(1100, 1000, 1100 - INPUT_GRACE_MS - 1)).toBe(false)
  })
})

function runtimeWith(tasks: (() => void)[], makeDirty: () => void): UniverRuntime {
  const render = { engine: { _renderFrameTasks: tasks }, scene: { makeDirty } }
  return {
    univer: {
      __getInjector: () => ({ get: () => ({ getRenderById: () => render }) }),
    },
  } as unknown as UniverRuntime
}

describe('throttling the render loop', () => {
  it('stands in for the frame tasks and puts them back, drawing once', () => {
    const draw = vi.fn()
    const makeDirty = vi.fn()
    const tasks = [draw]
    const restore = throttleGridRendering(runtimeWith(tasks, makeDirty), 'file-x')
    expect(tasks).toHaveLength(1)
    expect(tasks[0]).not.toBe(draw)
    tasks[0]!()
    tasks[0]!()
    // the first frame draws, the next within the interval does not
    expect(draw).toHaveBeenCalledTimes(1)
    restore()
    expect(tasks).toEqual([draw])
    expect(makeDirty).toHaveBeenCalledWith(true)
    restore()
    expect(makeDirty).toHaveBeenCalledTimes(1)
  })

  it('leaves an unreachable engine alone', () => {
    const broken = {
      univer: {
        __getInjector: () => {
          throw new Error('gone')
        },
      },
    } as unknown as UniverRuntime
    expect(() => throttleGridRendering(broken, 'file-x')()).not.toThrow()
    const tasks: (() => void)[] = []
    throttleGridRendering(runtimeWith(tasks, vi.fn()), 'file-x')()
    expect(tasks).toEqual([])
  })
})
