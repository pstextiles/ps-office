/**
 * Data › Queries & Connections › Refresh in Excel (PS Office): the app has no
 * query engine, so the refresh is done by the Excel installed on the PC.
 * A fixed PowerShell script opens the workbook in a hidden Excel instance
 * with macros forced off, refreshes every query table one by one, saves, and
 * quits. The renderer then reopens the file.
 *
 * - The workbook path reaches the script through the environment
 *   (PSO_WORKBOOK), never through the script text or the command line, so a
 *   file name cannot change what runs.
 * - RefreshAll is not used: it swallows failures (an unreachable source keeps
 *   the old rows and reports nothing). QueryTable.Refresh(false) raises the
 *   source's own error; any failure closes the workbook without saving.
 * - The run is bounded (timeout) and ends with the Excel it started: the
 *   script quits Excel, and this module kills that Excel process if it is
 *   still alive afterwards, on timeout, or when the app quits.
 *
 * Electron-free so the tests run it with a fake spawn.
 */
import { spawn as nodeSpawn, type ChildProcess, type SpawnOptions } from 'node:child_process'

export type ExcelRefreshResult =
  | { readonly ok: true; readonly refreshed: number }
  | {
      readonly ok: false
      readonly code: ExcelRefreshFailure
      /** Excel's or the data source's own message, for the "failed" case */
      readonly detail?: string
    }

export type ExcelRefreshFailure =
  'unsupported' | 'busy' | 'no-excel' | 'locked' | 'readonly' | 'failed' | 'timeout'

/** Ten minutes: large network sources are slow, a hung prompt is not worth waiting longer for. */
export const EXCEL_REFRESH_TIMEOUT_MS = 10 * 60 * 1000
/** How long Excel gets to exit on its own after the script quit it. */
const EXCEL_EXIT_GRACE_MS = 8000
const MAX_DETAIL = 400

/** The script run by Windows PowerShell 5.1 (present on every supported Windows). */
export const EXCEL_REFRESH_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$path = $env:PSO_WORKBOOK
$excel = $null
$book = $null
function Out-Result([string]$code, [string]$detail = '') {
  $clean = ($detail -replace '[\r\n]+', ' ').Trim()
  if ($clean.Length -gt 400) { $clean = $clean.Substring(0, 400) }
  [Console]::Out.WriteLine("RESULT $code $clean")
  [Console]::Out.Flush()
}
try {
  if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { Out-Result 'failed' 'file not found'; exit 1 }
  try { $probe = [IO.File]::Open($path, 'Open', 'ReadWrite', 'None'); $probe.Close() }
  catch { Out-Result 'locked'; exit 3 }
  try { $excel = New-Object -ComObject Excel.Application }
  catch { Out-Result 'no-excel' $_.Exception.Message; exit 2 }
  Add-Type -Namespace PsOffice -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);'
  $excelPid = [uint32]0
  [void][PsOffice.Win]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
  [Console]::Out.WriteLine("PID $excelPid"); [Console]::Out.Flush()
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false
  $excel.ScreenUpdating = $false
  $excel.AutomationSecurity = 3
  $book = $excel.Workbooks.Open($path, 0, $false)
  if ($book.ReadOnly) { Out-Result 'readonly'; exit 3 }
  $failures = @()
  $refreshed = 0
  function Update-Table($qt, [string]$label) {
    $background = $null
    try { $background = [bool]$qt.BackgroundQuery; $qt.BackgroundQuery = $false } catch { }
    try { [void]$qt.Refresh($false); $script:refreshed++ }
    catch { $script:failures += "$($label): $($_.Exception.Message)" }
    finally { if ($null -ne $background) { try { $qt.BackgroundQuery = $background } catch { } } }
  }
  foreach ($ws in @($book.Worksheets)) {
    foreach ($lo in @($ws.ListObjects)) {
      if ($lo.SourceType -eq 3) { Update-Table $lo.QueryTable $lo.Name }
    }
    foreach ($qt in @($ws.QueryTables)) { Update-Table $qt $qt.Name }
  }
  foreach ($c in @($book.Connections)) {
    $ranges = 0
    try { $ranges = $c.Ranges.Count } catch { }
    if ($ranges -gt 0) { continue }
    try { $c.Refresh(); $refreshed++ } catch { $failures += "$($c.Name): $($_.Exception.Message)" }
  }
  $excel.CalculateUntilAsyncQueriesDone()
  if ($failures.Count -gt 0) {
    $book.Close($false)
    $book = $null
    Out-Result 'failed' ($failures -join '; ')
    exit 1
  }
  foreach ($cache in @($book.PivotCaches())) { try { [void]$cache.Refresh() } catch { } }
  $book.Save()
  $book.Close($false)
  $book = $null
  Out-Result 'ok' "$refreshed"
  exit 0
} catch {
  Out-Result 'failed' $_.Exception.Message
  exit 1
} finally {
  if ($null -ne $book) { try { $book.Close($false) } catch { } }
  if ($null -ne $excel) {
    try { $excel.Quit() } catch { }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
`

/**
 * Fallback cleanup when a run is stopped before the script reported Excel's
 * process id (Excel still starting): end the Excel instances COM started for
 * automation (`/automation -Embedding`) since this run began. An Excel the
 * person opened themselves never has that command line.
 */
export const EXCEL_ORPHAN_CLEANUP_SCRIPT = String.raw`
$since = [DateTime]::Parse($env:PSO_SINCE, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToLocalTime()
Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" |
  Where-Object { $_.CommandLine -match '/automation' -and $_.CreationDate -ge $since } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
`

/** -EncodedCommand takes the script as base64 of its UTF-16LE text. */
export function encodePowerShellCommand(script: string): string {
  return Buffer.from(script, 'utf16le').toString('base64')
}

export interface ParsedRefreshOutput {
  readonly excelPid?: number
  readonly result?: ExcelRefreshResult
}

const RESULT_CODES = new Set(['ok', 'no-excel', 'locked', 'readonly', 'failed'])

/** Reads the script's "PID n" and "RESULT code detail" lines out of its standard output. */
export function parseRefreshOutput(stdout: string): ParsedRefreshOutput {
  let excelPid: number | undefined
  let result: ExcelRefreshResult | undefined
  for (const raw of stdout.split(/\r?\n/)) {
    const line = raw.trim()
    const pid = /^PID (\d+)$/.exec(line)
    if (pid) {
      const value = Number(pid[1])
      if (Number.isSafeInteger(value) && value > 0) excelPid = value
      continue
    }
    const found = /^RESULT (\S+)(?: (.*))?$/.exec(line)
    if (!found || !RESULT_CODES.has(found[1]!)) continue
    const code = found[1]!
    const detail = (found[2] ?? '').trim().slice(0, MAX_DETAIL)
    if (code === 'ok') {
      const count = Number(detail)
      result = { ok: true, refreshed: Number.isSafeInteger(count) && count >= 0 ? count : 0 }
    } else {
      result = {
        ok: false,
        code: code as ExcelRefreshFailure,
        ...(detail ? { detail } : {}),
      }
    }
  }
  return { ...(excelPid !== undefined ? { excelPid } : {}), ...(result ? { result } : {}) }
}

/** Only an existing .xlsx/.xlsm the user opened as itself can be refreshed in place. */
export function refreshablePath(path: string): boolean {
  return /\.xls[xm]$/i.test(path)
}

type SpawnLike = (command: string, args: readonly string[], options: SpawnOptions) => ChildProcess

export interface RefreshDeps {
  readonly platform?: NodeJS.Platform
  readonly spawn?: SpawnLike
  /** process liveness check and kill, injectable for tests */
  readonly isAlive?: (pid: number) => boolean
  readonly kill?: (pid: number) => void
  readonly timeoutMs?: number
  readonly graceMs?: number
}

function defaultIsAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}

function defaultKill(pid: number): void {
  try {
    process.kill(pid)
  } catch {
    // already gone
  }
}

/** Runs in flight, so a second click waits and the app can end them on quit. */
const running = new Set<{ stop: () => void }>()

/** Ends every refresh still running (the app is quitting). */
export function stopExcelRefreshes(): void {
  for (const run of [...running]) run.stop()
}

export function excelRefreshRunning(): boolean {
  return running.size > 0
}

/**
 * Refreshes the workbook at `path` with the installed Excel and saves it in
 * place. Resolves with the outcome; never rejects.
 */
export function refreshWorkbookInExcel(
  path: string,
  deps: RefreshDeps = {},
): Promise<ExcelRefreshResult> {
  const platform = deps.platform ?? process.platform
  if (platform !== 'win32') return Promise.resolve({ ok: false, code: 'unsupported' })
  if (!refreshablePath(path)) return Promise.resolve({ ok: false, code: 'unsupported' })
  if (running.size > 0) return Promise.resolve({ ok: false, code: 'busy' })
  const spawn = deps.spawn ?? (nodeSpawn as SpawnLike)
  const isAlive = deps.isAlive ?? defaultIsAlive
  const kill = deps.kill ?? defaultKill
  const timeoutMs = deps.timeoutMs ?? EXCEL_REFRESH_TIMEOUT_MS
  const graceMs = deps.graceMs ?? EXCEL_EXIT_GRACE_MS

  return new Promise((resolve) => {
    let stdout = ''
    let settled = false
    let timedOut = false
    let child: ChildProcess
    // a second before the start, so an Excel launched in the same second is not missed
    const since = new Date(Date.now() - 1000).toISOString()
    const cleanupOrphans = (): void => {
      try {
        const cleanup = spawn(
          'powershell.exe',
          [
            '-NoLogo',
            '-NoProfile',
            '-NonInteractive',
            '-ExecutionPolicy',
            'Bypass',
            '-EncodedCommand',
            encodePowerShellCommand(EXCEL_ORPHAN_CLEANUP_SCRIPT),
          ],
          {
            env: { ...process.env, PSO_SINCE: since },
            shell: false,
            windowsHide: true,
            stdio: 'ignore',
          },
        )
        cleanup.on('error', () => undefined)
      } catch {
        // nothing more to do: the Excel instance ends when the PC restarts
      }
    }
    const endExcel = (pid: number | undefined, wait: number): void => {
      if (pid === undefined) return
      if (wait <= 0) {
        if (isAlive(pid)) kill(pid)
        return
      }
      setTimeout(() => {
        if (isAlive(pid)) kill(pid)
      }, wait).unref?.()
    }
    const run = {
      stop: () => {
        timedOut = true
        if (child.pid !== undefined) kill(child.pid)
        const excelPid = parseRefreshOutput(stdout).excelPid
        if (excelPid !== undefined) endExcel(excelPid, 0)
        else {
          // stopped while Excel was still starting: sweep now and once more after it started
          cleanupOrphans()
          setTimeout(cleanupOrphans, 5000).unref?.()
        }
      },
    }
    const finish = (result: ExcelRefreshResult): void => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      running.delete(run)
      resolve(result)
    }
    try {
      child = spawn(
        'powershell.exe',
        [
          '-NoLogo',
          '-NoProfile',
          '-NonInteractive',
          '-ExecutionPolicy',
          'Bypass',
          '-EncodedCommand',
          encodePowerShellCommand(EXCEL_REFRESH_SCRIPT),
        ],
        {
          env: { ...process.env, PSO_WORKBOOK: path },
          shell: false,
          windowsHide: true,
          stdio: ['ignore', 'pipe', 'ignore'],
        },
      )
    } catch (err) {
      resolve({ ok: false, code: 'failed', detail: String(err).slice(0, MAX_DETAIL) })
      return
    }
    running.add(run)
    const timer = setTimeout(() => {
      run.stop()
      finish({ ok: false, code: 'timeout' })
    }, timeoutMs)
    child.stdout?.setEncoding('utf8')
    child.stdout?.on('data', (chunk: string) => {
      // the script writes two short lines; cap anything else it might print
      if (stdout.length < 64 * 1024) stdout += chunk
    })
    child.on('error', (err) => {
      finish({ ok: false, code: 'failed', detail: String(err).slice(0, MAX_DETAIL) })
    })
    child.on('close', () => {
      const parsed = parseRefreshOutput(stdout)
      // Excel lingers a few seconds after Quit; end it if it outstays that
      endExcel(parsed.excelPid, graceMs)
      if (timedOut) return finish({ ok: false, code: 'timeout' })
      finish(parsed.result ?? { ok: false, code: 'failed', detail: 'no result from Excel' })
    })
  })
}
