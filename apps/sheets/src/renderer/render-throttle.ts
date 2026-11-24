/**
 * PS Office: while a bulk load writes the grid (the full load, an in-place
 * refresh), draw it a few times a second instead of every frame.
 *
 * Univer redraws the visible grid on every animation frame after the model
 * changed, and each block a load installs changes it; one redraw of a wide,
 * zoomed-out sheet walks thousands of cells through every draw extension. On
 * the production plan workbook a full load spent 7-25s of its 26-62s drawing
 * frames nobody looked at; with this it took ~10s. A scroll, click or key
 * press draws right away for a moment, so the grid still answers the person.
 */
import { IRenderManagerService } from '@univerjs/engine-render'

import type { UniverRuntime } from './univer-state'

/** at most one redraw this often while throttled (ms) */
export const THROTTLED_FRAME_MS = 400
/** after an input, draw every frame this long (ms) */
export const INPUT_GRACE_MS = 300

let lastInput = -Infinity
let listening = false

function listenForInput(): void {
  if (listening || typeof window === 'undefined') return
  listening = true
  const note = (): void => {
    lastInput = performance.now()
  }
  // not pointermove: hovering does not need a redraw, and a resting hand on
  // the mouse would otherwise undo the throttle
  for (const type of ['wheel', 'pointerdown', 'keydown'] as const) {
    window.addEventListener(type, note, { capture: true, passive: true })
  }
}

/** Whether a throttled frame should draw now. */
export function shouldDrawFrame(now: number, lastDrawn: number, inputAt: number): boolean {
  return now - lastDrawn >= THROTTLED_FRAME_MS || now - inputAt <= INPUT_GRACE_MS
}

interface EngineLike {
  _renderFrameTasks?: (() => void)[]
}

/**
 * Throttles the workbook's render loop until the returned function is called,
 * which restores it and draws once. Does nothing (and restores nothing) when
 * the engine is not reachable.
 */
export function throttleGridRendering(runtime: UniverRuntime, unitId: string): () => void {
  try {
    const render = runtime.univer.__getInjector().get(IRenderManagerService).getRenderById(unitId)
    const engine = render?.engine as unknown as EngineLike | undefined
    const tasks = engine?._renderFrameTasks
    if (!render || !Array.isArray(tasks) || tasks.length === 0) return () => {}
    listenForInput()
    const originals = [...tasks]
    let lastDrawn = -Infinity
    const throttled = (): void => {
      const now = performance.now()
      if (!shouldDrawFrame(now, lastDrawn, lastInput)) return
      lastDrawn = now
      for (const task of originals) task()
    }
    tasks.splice(0, tasks.length, throttled)
    let restored = false
    return () => {
      if (restored) return
      restored = true
      const at = tasks.indexOf(throttled)
      if (at >= 0) tasks.splice(at, 1, ...originals)
      render.scene.makeDirty(true)
    }
  } catch {
    return () => {}
  }
}
