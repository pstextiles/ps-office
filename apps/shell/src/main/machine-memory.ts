import { freemem, totalmem } from 'node:os'

/**
 * PS Office: speed-ups that hold memory ahead of time (the spare Sheets
 * renderer, about 270 MB) pay off on a roomy PC and hurt on the low-end PCs
 * of a factory floor, where every hundred MB decides whether Windows starts
 * swapping. Below these figures the app does the same work on demand instead.
 */
export const ROOMY_TOTAL_BYTES = 12 * 1024 ** 3
export const ROOMY_FREE_BYTES = 2 * 1024 ** 3

export interface MemoryFigures {
  readonly total: number
  readonly free: number
}

export function machineMemory(): MemoryFigures {
  return { total: totalmem(), free: freemem() }
}

/** Whether holding a spare renderer ahead of time is worth it on this PC right now. */
export function roomForSpareRenderer(memory: MemoryFigures = machineMemory()): boolean {
  return memory.total >= ROOMY_TOTAL_BYTES && memory.free >= ROOMY_FREE_BYTES
}
