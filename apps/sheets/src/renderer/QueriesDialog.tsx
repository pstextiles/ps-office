import { useEffect, useRef, useState } from 'react'

import type { WorkbookQueriesResult } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

type Query = WorkbookQueriesResult[number]

const KIND_KEY = {
  powerQuery: 'dlgQueriesKindPowerQuery',
  web: 'dlgQueriesKindWeb',
  text: 'dlgQueriesKindText',
  database: 'dlgQueriesKindDatabase',
  other: 'dlgQueriesKindOther',
} as const

/// `'My Sheet'!A1:C21`, the form Go To resolves
function qualifiedRange(destination: NonNullable<Query['destination']>): string {
  return `'${destination.sheet.replaceAll("'", "''")}'!${destination.range}`
}

/// Data › Queries & Connections (genoffice#2051): the workbook's Power Query
/// queries and data connections, and where each one's rows are loaded.
/// Clicking a loaded query selects its rows. Refresh (PS Office) has the
/// built-in engine compute the queries and write their tables; on Windows,
/// Refresh in Excel has the installed Excel refresh and save the file. Either
/// way the workbook is reopened.
export function QueriesDialog({
  load,
  onRefreshBuiltIn,
  onRefresh,
  onCancelRefresh,
  onGo,
  onClose,
}: {
  readonly load: () => Promise<WorkbookQueriesResult>
  /// Built-in refresh; absent in builds without it. What to tell the person once
  /// the workbook is reopened (formulas elsewhere keep their previous results), or an error.
  readonly onRefreshBuiltIn?:
    (() => Promise<{ readonly done: string } | { readonly error: string }>) | undefined
  /// Refresh in Excel; absent where it cannot run. An error message, or null once reopened.
  readonly onRefresh?: (() => Promise<string | null>) | undefined
  /// Stops a running Refresh in Excel; its onRefresh then resolves with the cancelled message.
  readonly onCancelRefresh?: (() => void) | undefined
  /// Returns an error message, or null on success (Go To's resolver).
  readonly onGo: (ref: string) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [queries, setQueries] = useState<WorkbookQueriesResult | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [refreshing, setRefreshing] = useState(false)
  // which refresh is running: only the one in Excel can be cancelled
  const [builtIn, setBuiltIn] = useState(false)
  // the built-in refresh's result note, kept on screen: the status bar cuts long lines
  const [done, setDone] = useState<string | null>(null)
  const loadRef = useRef(load)
  loadRef.current = load
  // seconds since Refresh in Excel started: a slow PC takes minutes, and a
  // ticking clock tells the person it is still working
  const [elapsed, setElapsed] = useState(0)
  const [cancelling, setCancelling] = useState(false)

  useEffect(() => {
    if (!refreshing) return
    const started = Date.now()
    setElapsed(0)
    const timer = setInterval(() => setElapsed(Math.floor((Date.now() - started) / 1000)), 1000)
    return () => clearInterval(timer)
  }, [refreshing])

  useEffect(() => {
    let alive = true
    load().then(
      (result) => alive && setQueries(result),
      (err: unknown) => alive && setError(err instanceof Error ? err.message : String(err)),
    )
    return () => {
      alive = false
    }
    // load once per open; the dialog is remounted for the next workbook
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const go = (destination: NonNullable<Query['destination']>): void => {
    const failure = onGo(qualifiedRange(destination))
    if (failure === null) onClose()
    else setError(failure)
  }

  const refreshBuiltIn = async (): Promise<void> => {
    if (!onRefreshBuiltIn || refreshing) return
    setError(null)
    setDone(null)
    setBuiltIn(true)
    setRefreshing(true)
    const result = await onRefreshBuiltIn().catch((err: unknown) => ({
      error: t('dlgQueriesBuiltInFailed', {
        error: err instanceof Error ? err.message : String(err),
      }),
    }))
    setRefreshing(false)
    setBuiltIn(false)
    if ('error' in result) {
      setError(result.error)
      return
    }
    setDone(result.done)
    // the workbook was reopened: show its tables' new row counts
    setQueries(await loadRef.current().catch(() => null))
  }

  const refresh = async (): Promise<void> => {
    if (!onRefresh || refreshing) return
    setError(null)
    setDone(null)
    setCancelling(false)
    setRefreshing(true)
    const failure = await onRefresh().catch((err: unknown) =>
      t('dlgQueriesRefreshFailed', { error: err instanceof Error ? err.message : String(err) }),
    )
    setRefreshing(false)
    // success reopened the workbook: the list on screen belongs to the old file
    if (failure === null) onClose()
    else setError(failure)
  }

  // closing mid-refresh would hide the progress; the run ends on its own (timeout included)
  const close = (): void => {
    if (!refreshing) onClose()
  }
  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgQueriesTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgQueriesTitle')}</header>
        <section className="dialog-body">
          {queries?.length === 0 && <p className="dialog-note">{t('dlgQueriesEmpty')}</p>}
          {queries && queries.length > 0 && (
            <ul className="goto-name-list">
              {queries.map((query, index) => {
                const destination = query.destination
                const detail = [
                  t(KIND_KEY[query.kind]),
                  destination
                    ? t('dlgQueriesLoaded', {
                        rows: destination.rows,
                        target: `${destination.table ?? destination.sheet} (${qualifiedRange(destination)})`,
                      })
                    : t('dlgQueriesConnectionOnly'),
                  query.source,
                ]
                  .filter(Boolean)
                  .join(' · ')
                return (
                  <li key={`${index}:${query.name}`}>
                    <button
                      type="button"
                      disabled={!destination}
                      data-tip={destination ? t('dlgQueriesGoHint') : query.description}
                      onClick={() => destination && go(destination)}
                    >
                      <strong>{query.name}</strong>
                      <span>{detail}</span>
                    </button>
                  </li>
                )
              })}
            </ul>
          )}
          {queries && queries.length > 0 && (
            <p className="dialog-note">
              {t(
                onRefreshBuiltIn
                  ? 'dlgQueriesNoteBuiltIn'
                  : onRefresh
                    ? 'dlgQueriesNoteRefresh'
                    : 'dlgQueriesNote',
              )}
            </p>
          )}
          {done && !refreshing && (
            <p className="dialog-note" role="status">
              {done}
            </p>
          )}
          {refreshing && (
            <p className="dialog-note" role="status">
              {t(builtIn ? 'dlgQueriesBuiltInRunning' : 'dlgQueriesRefreshing')}{' '}
              {Math.floor(elapsed / 60)}:{String(elapsed % 60).padStart(2, '0')}
            </p>
          )}
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {t('dlgQueriesError', { error })}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          {onRefreshBuiltIn && queries && queries.length > 0 && (
            <button
              type="button"
              disabled={refreshing}
              data-tip={t('dlgQueriesBuiltInTip')}
              onClick={() => void refreshBuiltIn()}
            >
              {t('dlgQueriesBuiltIn')}
            </button>
          )}
          {onRefresh && queries && queries.length > 0 && (
            <button
              type="button"
              disabled={refreshing}
              data-tip={t('dlgQueriesRefreshTip')}
              onClick={() => void refresh()}
            >
              {t('dlgQueriesRefreshExcel')}
            </button>
          )}
          {refreshing && !builtIn && onCancelRefresh && (
            <button
              type="button"
              disabled={cancelling}
              onClick={() => {
                setCancelling(true)
                onCancelRefresh()
              }}
            >
              {t('dlgCancel')}
            </button>
          )}
          <button className="primary-action" autoFocus disabled={refreshing} onClick={onClose}>
            {t('dlgClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
