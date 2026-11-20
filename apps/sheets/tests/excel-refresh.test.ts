import { EventEmitter } from 'node:events'
import { PassThrough } from 'node:stream'
import type { ChildProcess, SpawnOptions } from 'node:child_process'
import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  EXCEL_ORPHAN_CLEANUP_SCRIPT,
  EXCEL_REFRESH_SCRIPT,
  encodePowerShellCommand,
  excelRefreshRunning,
  parseRefreshOutput,
  refreshablePath,
  refreshWorkbookInExcel,
  stopExcelRefreshes,
  cancelExcelRefresh,
  EXCEL_REFRESH_TIMEOUT_MS,
} from '../src/main/excel-refresh'

interface FakeChild {
  readonly child: ChildProcess
  readonly command: string
  readonly args: readonly string[]
  readonly options: SpawnOptions
  print(text: string): void
  close(code?: number): void
}

/** A spawn that records each call and hands back a child the test drives. */
function fakeSpawn() {
  const children: FakeChild[] = []
  let nextPid = 1000
  const spawn = (command: string, args: readonly string[], options: SpawnOptions): ChildProcess => {
    const emitter = new EventEmitter() as ChildProcess & EventEmitter
    const stdout = new PassThrough()
    Object.assign(emitter, { stdout, pid: nextPid++ })
    children.push({
      child: emitter,
      command,
      args,
      options,
      print: (text) => stdout.write(text),
      close: (code = 0) => {
        stdout.end()
        // let the data events land before close, as a real process does
        setTimeout(() => emitter.emit('close', code), 5)
      },
    })
    return emitter
  }
  return { spawn, children }
}

const WIN = { platform: 'win32' as const }

afterEach(() => {
  stopExcelRefreshes()
})

describe('parseRefreshOutput', () => {
  it('reads the Excel process id and a success with its count', () => {
    expect(parseRefreshOutput('PID 4242\r\nRESULT ok 3\r\n')).toEqual({
      excelPid: 4242,
      result: { ok: true, refreshed: 3 },
    })
  })

  it("keeps a failure's own message", () => {
    expect(parseRefreshOutput('PID 7\nRESULT failed Datos: [DataSource.NotFound] gone\n')).toEqual({
      excelPid: 7,
      result: { ok: false, code: 'failed', detail: 'Datos: [DataSource.NotFound] gone' },
    })
  })

  it('reads the codes that come before Excel starts', () => {
    expect(parseRefreshOutput('RESULT locked \n').result).toEqual({ ok: false, code: 'locked' })
    expect(parseRefreshOutput('RESULT no-excel class not registered\n').result).toEqual({
      ok: false,
      code: 'no-excel',
      detail: 'class not registered',
    })
  })

  it('ignores anything that is not one of its lines', () => {
    expect(parseRefreshOutput('hello\nRESULT whatever x\nPID abc\nPID -1\n')).toEqual({})
  })

  it('caps a long message', () => {
    const parsed = parseRefreshOutput(`RESULT failed ${'x'.repeat(1000)}\n`)
    expect(parsed.result).toMatchObject({ ok: false, code: 'failed' })
    expect((parsed.result as { detail: string }).detail).toHaveLength(400)
  })
})

describe('the scripts', () => {
  it('pass as UTF-16LE base64, the form -EncodedCommand takes', () => {
    const encoded = encodePowerShellCommand('Write-Output "é"')
    expect(Buffer.from(encoded, 'base64').toString('utf16le')).toBe('Write-Output "é"')
  })

  it('take the workbook path from the environment only', () => {
    expect(EXCEL_REFRESH_SCRIPT).toContain('$env:PSO_WORKBOOK')
    expect(EXCEL_ORPHAN_CLEANUP_SCRIPT).toContain('$env:PSO_SINCE')
  })

  it('refresh table by table with macros off, never RefreshAll', () => {
    // RefreshAll swallows failures: an unreachable source would keep the old rows silently
    expect(EXCEL_REFRESH_SCRIPT).not.toMatch(/RefreshAll\s*\(/)
    expect(EXCEL_REFRESH_SCRIPT).toContain('.Refresh($false)')
    expect(EXCEL_REFRESH_SCRIPT).toContain('$excel.AutomationSecurity = 3')
    expect(EXCEL_REFRESH_SCRIPT).toContain('$excel.Quit()')
  })

  it('only sweep Excel instances started for automation', () => {
    expect(EXCEL_ORPHAN_CLEANUP_SCRIPT).toContain("-match '/automation'")
  })
})

describe('refreshablePath', () => {
  it('takes .xlsx and .xlsm, nothing else', () => {
    expect(refreshablePath('C:\\Plan\\LINE PLAN.xlsx')).toBe(true)
    expect(refreshablePath('C:\\Plan\\macros.XLSM')).toBe(true)
    expect(refreshablePath('C:\\Plan\\old.xls')).toBe(false)
    expect(refreshablePath('C:\\Plan\\data.csv')).toBe(false)
    expect(refreshablePath('C:\\Plan\\sheet.ods')).toBe(false)
  })
})

describe('refreshWorkbookInExcel', () => {
  it('is not offered outside Windows', async () => {
    const { spawn, children } = fakeSpawn()
    await expect(
      refreshWorkbookInExcel('/home/u/a.xlsx', { platform: 'linux', spawn }),
    ).resolves.toEqual({ ok: false, code: 'unsupported' })
    expect(children).toHaveLength(0)
  })

  it('refuses a file Excel would not save in place', async () => {
    const { spawn, children } = fakeSpawn()
    await expect(refreshWorkbookInExcel('C:\\a.csv', { ...WIN, spawn })).resolves.toEqual({
      ok: false,
      code: 'unsupported',
    })
    expect(children).toHaveLength(0)
  })

  it('runs Windows PowerShell without a shell and keeps the path out of the command line', async () => {
    const { spawn, children } = fakeSpawn()
    const path = 'C:\\Plan\\it"s & $(danger).xlsx'
    const pending = refreshWorkbookInExcel(path, { ...WIN, spawn, isAlive: () => false })
    const run = children[0]!
    expect(run.command).toBe('powershell.exe')
    expect(run.options.shell).toBe(false)
    expect(run.options.windowsHide).toBe(true)
    expect((run.options.env as Record<string, string>).PSO_WORKBOOK).toBe(path)
    expect(run.args.join(' ')).not.toContain('Plan')
    expect(run.args).toContain('-EncodedCommand')
    run.print('PID 55\nRESULT ok 2\n')
    run.close(0)
    await expect(pending).resolves.toEqual({ ok: true, refreshed: 2 })
  })

  it('reports what the script reported when it fails', async () => {
    const { spawn, children } = fakeSpawn()
    const pending = refreshWorkbookInExcel('C:\\a.xlsx', { ...WIN, spawn, isAlive: () => false })
    children[0]!.print('PID 9\nRESULT failed Datos: [DataSource.NotFound] gone\n')
    children[0]!.close(1)
    await expect(pending).resolves.toEqual({
      ok: false,
      code: 'failed',
      detail: 'Datos: [DataSource.NotFound] gone',
    })
  })

  it('treats a run that printed no result as failed', async () => {
    const { spawn, children } = fakeSpawn()
    const pending = refreshWorkbookInExcel('C:\\a.xlsx', { ...WIN, spawn })
    children[0]!.close(1)
    await expect(pending).resolves.toMatchObject({ ok: false, code: 'failed' })
  })

  it('runs one refresh at a time', async () => {
    const { spawn, children } = fakeSpawn()
    const first = refreshWorkbookInExcel('C:\\a.xlsx', { ...WIN, spawn, isAlive: () => false })
    expect(excelRefreshRunning()).toBe(true)
    await expect(refreshWorkbookInExcel('C:\\b.xlsx', { ...WIN, spawn })).resolves.toEqual({
      ok: false,
      code: 'busy',
    })
    expect(children).toHaveLength(1)
    children[0]!.print('RESULT ok 1\n')
    children[0]!.close(0)
    await first
    expect(excelRefreshRunning()).toBe(false)
  })

  it('ends Excel when it outstays the script', async () => {
    const { spawn, children } = fakeSpawn()
    const kill = vi.fn()
    const pending = refreshWorkbookInExcel('C:\\a.xlsx', {
      ...WIN,
      spawn,
      kill,
      isAlive: () => true,
      graceMs: 10,
    })
    children[0]!.print('PID 321\nRESULT ok 1\n')
    children[0]!.close(0)
    await pending
    await new Promise((resolve) => setTimeout(resolve, 40))
    expect(kill).toHaveBeenCalledWith(321)
  })

  it('stops PowerShell and its Excel on timeout', async () => {
    const { spawn, children } = fakeSpawn()
    const kill = vi.fn()
    const pending = refreshWorkbookInExcel('C:\\a.xlsx', {
      ...WIN,
      spawn,
      kill,
      isAlive: () => true,
      timeoutMs: 20,
    })
    children[0]!.print('PID 777\n')
    await expect(pending).resolves.toEqual({ ok: false, code: 'timeout' })
    expect(kill).toHaveBeenCalledWith(children[0]!.child.pid)
    expect(kill).toHaveBeenCalledWith(777)
  })

  it('sweeps automation Excel when stopped before Excel reported itself', async () => {
    const { spawn, children } = fakeSpawn()
    const kill = vi.fn()
    const pending = refreshWorkbookInExcel('C:\\a.xlsx', {
      ...WIN,
      spawn,
      kill,
      isAlive: () => true,
      timeoutMs: 20,
    })
    await expect(pending).resolves.toEqual({ ok: false, code: 'timeout' })
    // the refresh run, then the cleanup run with the start time in its environment
    expect(children).toHaveLength(2)
    expect((children[1]!.options.env as Record<string, string>).PSO_SINCE).toMatch(
      /^\d{4}-\d\d-\d\dT/,
    )
    expect(children[1]!.args).toContain(encodePowerShellCommand(EXCEL_ORPHAN_CLEANUP_SCRIPT))
  })

  it('ends a running refresh when the app quits', async () => {
    const { spawn, children } = fakeSpawn()
    const kill = vi.fn()
    const pending = refreshWorkbookInExcel('C:\\a.xlsx', {
      ...WIN,
      spawn,
      kill,
      isAlive: () => true,
    })
    children[0]!.print('PID 888\n')
    await new Promise((resolve) => setTimeout(resolve, 5))
    stopExcelRefreshes()
    expect(kill).toHaveBeenCalledWith(888)
    children[0]!.close(1)
    await expect(pending).resolves.toEqual({ ok: false, code: 'cancelled' })
  })

  it('ends the run and its Excel when the person cancels', async () => {
    const { spawn, children } = fakeSpawn()
    const kill = vi.fn()
    const pending = refreshWorkbookInExcel('C:\\a.xlsx', {
      ...WIN,
      spawn,
      kill,
      isAlive: () => true,
    })
    children[0]!.print('PID 999\n')
    await new Promise((resolve) => setTimeout(resolve, 5))
    cancelExcelRefresh()
    expect(kill).toHaveBeenCalledWith(children[0]!.child.pid)
    expect(kill).toHaveBeenCalledWith(999)
    // a second press while it winds down does nothing more
    cancelExcelRefresh()
    expect(kill).toHaveBeenCalledTimes(2)
    children[0]!.close(1)
    await expect(pending).resolves.toEqual({ ok: false, code: 'cancelled' })
    expect(excelRefreshRunning()).toBe(false)
  })

  it('allows up to thirty minutes by default', () => {
    expect(EXCEL_REFRESH_TIMEOUT_MS).toBe(30 * 60 * 1000)
  })
})
