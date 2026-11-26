import { describe, expect, it } from 'vitest'

import { LANGS } from '@genoffice/i18n'

import {
  describeRefreshResult,
  parseRefreshQueriesArgv,
  progressPageHtml,
  refreshRunLabels,
  summarizeRefreshRun,
  type HeadlessRefreshResult,
} from '../src/main/query-refresh-run'

/**
 * PS Office's refresh-only run (`--refresh-queries`, src/main/query-refresh-run.ts):
 * the files Explorer passes, and the closing message the person reads.
 */
describe('parseRefreshQueriesArgv', () => {
  it('takes the files Explorer appends after the shortcut, once each', () => {
    expect(
      parseRefreshQueriesArgv([
        'C:\\PSOffice\\PSOffice.exe',
        '--refresh-queries',
        'C:\\plan\\Production Plan.xlsx',
        'C:\\plan\\otro.xlsm',
        'C:\\plan\\Production Plan.xlsx',
      ]),
    ).toEqual({
      kind: 'run',
      paths: ['C:\\plan\\Production Plan.xlsx', 'C:\\plan\\otro.xlsm'],
    })
  })

  it('runs with no files (the icon opened by itself) and ignores other switches', () => {
    expect(parseRefreshQueriesArgv(['app.exe', '--refresh-queries'])).toEqual({
      kind: 'run',
      paths: [],
    })
    expect(
      parseRefreshQueriesArgv(['app.exe', '--refresh-queries', '--allow-file-access', 'a.xlsx']),
    ).toEqual({ kind: 'run', paths: ['a.xlsx'] })
    expect(parseRefreshQueriesArgv(['app.exe', 'a.xlsx'])).toEqual({ kind: 'none' })
  })
})

const ok = (formulas: 'recalculated' | 'stale', detail?: string): HeadlessRefreshResult => ({
  ok: true,
  tables: [{}, {}, {}, {}, {}],
  formulas,
  ...(detail ? { formulasDetail: detail } : {}),
  recalcOnOpen: formulas === 'stale',
})

describe('the closing message', () => {
  const es = refreshRunLabels('es')

  it('says what happened to each file, in usted for Spanish', () => {
    expect(describeRefreshResult('plan.xlsx', ok('recalculated'), es)).toBe(
      'plan.xlsx: se actualizaron 5 tablas de consulta, se recalcularon las fórmulas y se guardó.',
    )
    expect(describeRefreshResult('plan.xlsx', { ok: false, code: 'locked' }, es)).toContain(
      'Ciérrelo e inténtelo de nuevo.',
    )
    expect(
      describeRefreshResult(
        'plan.xlsx',
        { ok: false, code: 'limit', query: 'Cut', detail: 'Table.Buffer' },
        es,
      ),
    ).toContain('La consulta Cut usa algo')
    // no empty brackets when there is no detail
    expect(describeRefreshResult('plan.xlsx', { ok: false, code: 'failed' }, es)).toBe(
      'plan.xlsx: no se pudo actualizar.',
    )
  })

  it('is an error when any file failed, a warning when formulas are stale, else information', () => {
    const files = (results: HeadlessRefreshResult[]) =>
      results.map((result, index) => ({ fileName: `f${index}.xlsx`, result }))
    expect(summarizeRefreshRun(files([ok('recalculated')]), es)).toMatchObject({
      type: 'info',
      title: 'Actualización terminada',
      exitCode: 0,
    })
    expect(
      summarizeRefreshRun(files([ok('recalculated'), ok('stale', 'OFFSET')]), es),
    ).toMatchObject({
      type: 'warning',
      exitCode: 0,
    })
    const mixed = summarizeRefreshRun(
      files([ok('recalculated'), { ok: false, code: 'not-file' }]),
      es,
    )
    expect(mixed).toMatchObject({ type: 'error', exitCode: 3 })
    expect(mixed.message.split('\n\n')).toHaveLength(2)
    expect(summarizeRefreshRun([], es)).toMatchObject({ exitCode: 1, message: es.noFiles })
  })
})

describe('labels', () => {
  it('has every text in every screen language, with the placeholders the messages use', () => {
    const en = refreshRunLabels('en')
    for (const lang of LANGS) {
      const labels = refreshRunLabels(lang)
      if (lang !== 'en') expect(labels.running, lang).not.toBe(en.running)
      for (const [key, text] of Object.entries(labels)) {
        expect(text.trim(), `${lang}.${key}`).not.toBe('')
        const wanted = (en as unknown as Record<string, string>)[key]!.match(/\{\w+\}/g) ?? []
        for (const placeholder of wanted) expect(text, `${lang}.${key}`).toContain(placeholder)
      }
    }
  })

  it('escapes the file name in the progress page and keeps scripts out', () => {
    const page = progressPageHtml(
      refreshRunLabels('ko'),
      { index: 1, total: 2, fileName: '<b>plan</b>.xlsx' },
      false,
    )
    expect(page).toContain('&lt;b&gt;plan&lt;/b&gt;.xlsx')
    expect(page).toContain('(1/2)')
    expect(page).not.toContain('<script')
  })
})
