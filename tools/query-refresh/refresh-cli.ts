/**
 * PS Query Refresh — the standalone program (one .exe, no install): the
 * Power Query tables of each workbook dropped on it are refreshed and its
 * formulas recalculated by PS Office's built-in engine, and the file is
 * saved. For PCs too slow for PS Office or with an Excel too old to refresh
 * the queries. It is the same code as the app's refresh-only run
 * (apps/sheets/src/main/refresh-file.ts); only the window differs: this
 * program runs in a console window, in the PC's language.
 *
 *   PSQueryRefresh.exe <workbook.xlsx> [<workbook.xlsm> …]
 *
 * Started with no file (a double-click), it puts itself in Explorer's
 * "Send to" menu and says how to use it. Exit code 0 when every file was
 * refreshed, 1 for no files, 3 when any file was not. `--no-wait` skips the
 * "press Enter" at the end (a script, a scheduled task) and `--json` prints
 * one JSON line with each file's result.
 */
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'
import { Worker } from 'node:worker_threads'

import {
  refreshRunLabels,
  refreshRunReport,
  summarizeRefreshRun,
  describeRefreshResult,
  type HeadlessRefreshResult,
} from '../../apps/shell/src/main/query-refresh-run'
import {
  saveWorkbookViaSidecar,
  writePlannedSaveViaSidecar,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import { refreshFileInPlace } from '../../apps/sheets/src/main/refresh-file'
import { XlsxSidecarClient } from '../../apps/sheets/src/main/xlsx-sidecar-client'
import {
  QUERY_REFRESH_WORKER_DATA,
  type QueryRefreshAnswer,
  type QueryRefreshJob,
} from '../../apps/sheets/src/main/query-refresh-worker'

/** the refresh worker gets the stack the app's worker has (formulas that chain row after row recurse deeply) */
const WORKER_STACK_MB = 64
const WORKER_TIMEOUT_MS = 15 * 60_000

interface SeaApi {
  isSea(): boolean
  getAsset(key: string, encoding?: string): string | ArrayBuffer
}

function sea(): SeaApi | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const api = require('node:sea') as SeaApi
    return api.isSea() ? api : null
  } catch {
    return null
  }
}

/** an asset of the single executable, or the file of that name beside this script (a plain `node` run) */
function asset(name: string): Buffer {
  const api = sea()
  if (api) return Buffer.from(api.getAsset(name) as ArrayBuffer)
  return readFileSync(join(__dirname, name))
}

/** the xlsx sidecar, unpacked once per version into the user's local app data */
function sidecarPath(): string {
  const bytes = asset('xlsx-sidecar.exe')
  const digest = createHash('sha256').update(bytes).digest('hex').slice(0, 12)
  const base = process.env.LOCALAPPDATA ?? join(process.env.USERPROFILE ?? '.', 'AppData', 'Local')
  const dir = join(base, 'PSQueryRefresh', digest)
  const path = join(dir, 'xlsx-sidecar.exe')
  if (!existsSync(path)) {
    mkdirSync(dir, { recursive: true })
    writeFileSync(path, bytes)
  }
  return path
}

/** the refresh computed in a worker with the app worker's stack (its code is the worker asset) */
function planInWorker(job: QueryRefreshJob): Promise<QueryRefreshAnswer> {
  const code = asset('refresh-worker.cjs').toString('utf8')
  return new Promise((resolvePlan, reject) => {
    const worker = new Worker(code, {
      eval: true,
      workerData: QUERY_REFRESH_WORKER_DATA,
      resourceLimits: { stackSizeMb: WORKER_STACK_MB },
    })
    const timer = setTimeout(() => {
      void worker.terminate()
      reject(new Error('The refresh took longer than 15 minutes.'))
    }, WORKER_TIMEOUT_MS)
    const finish = (): void => {
      clearTimeout(timer)
      void worker.terminate()
    }
    worker.once('message', (answer: QueryRefreshAnswer) => {
      finish()
      resolvePlan(answer)
    })
    worker.once('error', (error) => {
      finish()
      reject(error)
    })
    worker.once('exit', (code) => {
      clearTimeout(timer)
      reject(new Error(`The refresh worker stopped (exit code ${code}).`))
    })
    worker.postMessage(job)
  })
}

/** the PC's language and culture (es-GT → Spanish labels, es-GT culture for the engine) */
function systemLocale(): string {
  return Intl.DateTimeFormat().resolvedOptions().locale || 'en-US'
}

/** this program in Explorer's Send to menu (per user, no admin): a file sent there is refreshed */
function ensureSendToShortcut(): void {
  const appData = process.env.APPDATA
  if (!appData || process.platform !== 'win32') return
  const link = join(appData, 'Microsoft', 'Windows', 'SendTo', 'PS Query Refresh.lnk')
  const exe = process.execPath
  const script = [
    '$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:PSQR_LINK)',
    '$s.TargetPath = $env:PSQR_EXE',
    '$s.WorkingDirectory = Split-Path $env:PSQR_EXE',
    '$s.IconLocation = "$($env:PSQR_EXE),0"',
    '$s.Save()',
  ].join('; ')
  try {
    execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], {
      env: { ...process.env, PSQR_LINK: link, PSQR_EXE: exe },
      stdio: 'ignore',
      windowsHide: true,
      timeout: 20_000,
    })
  } catch {
    // the menu entry is a convenience; dropping files on the program still works
  }
}

function waitForEnter(): Promise<void> {
  return new Promise((done) => {
    if (!process.stdin.isTTY) {
      done()
      return
    }
    process.stdin.resume()
    process.stdin.once('data', () => done())
  })
}

async function main(): Promise<number> {
  // a single executable repeats its own path as argv[1]; a plain `node` run has the script there
  const args = process.argv.slice(1)
  const self = process.execPath.toLowerCase()
  while (args.length > 0 && (args[0]!.toLowerCase() === self || /\.c?js$/i.test(args[0]!))) {
    args.shift()
  }
  const noWait = args.includes('--no-wait')
  const json = args.includes('--json')
  const paths = [...new Set(args.filter((arg) => !arg.startsWith('--') && arg.trim() !== ''))]
  const locale = systemLocale()
  const labels = refreshRunLabels(locale)
  const say = (text: string): void => {
    if (!json) process.stdout.write(`${text}\n`)
  }
  process.title = labels.title
  say(labels.title)
  say('')

  if (paths.length === 0) ensureSendToShortcut()
  const outcomes: { fileName: string; result: HeadlessRefreshResult }[] = []
  if (paths.length > 0) {
    say(labels.patience)
    const client = new XlsxSidecarClient(sidecarPath())
    client.start()
    try {
      for (const [index, path] of paths.entries()) {
        const fileName = basename(path)
        say('')
        say(
          labels.running
            .replace('{index}', String(index + 1))
            .replace('{total}', String(paths.length)),
        )
        say(`  ${fileName}`)
        const started = Date.now()
        let result: HeadlessRefreshResult
        try {
          result = await refreshFileInPlace(resolve(path), {
            client,
            systemCulture: locale,
            deps: {
              plan: planInWorker,
              save: saveWorkbookViaSidecar,
              write: writePlannedSaveViaSidecar,
            },
          })
        } catch (error) {
          result = {
            ok: false,
            code: 'failed',
            detail: error instanceof Error ? error.message : String(error),
          }
        }
        outcomes.push({ fileName, result })
        say(
          `  ${describeRefreshResult(fileName, result, labels)} (${Math.round((Date.now() - started) / 1000)} s)`,
        )
      }
    } finally {
      client.stop()
    }
  }
  const summary = summarizeRefreshRun(outcomes, labels)
  say('')
  say(paths.length === 0 ? summary.message : summary.title)
  if (json) process.stdout.write(`${JSON.stringify(refreshRunReport(outcomes))}\n`)
  if (!noWait && !json) {
    say('')
    say(labels.pressEnter)
    await waitForEnter()
  }
  return summary.exitCode
}

void main().then(
  (code) => process.exit(code),
  (error: unknown) => {
    process.stderr.write(
      `${error instanceof Error ? (error.stack ?? error.message) : String(error)}\n`,
    )
    process.exit(3)
  },
)
