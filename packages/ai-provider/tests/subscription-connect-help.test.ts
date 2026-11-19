import { describe, expect, it } from 'vitest'
import { subscriptionConnectHelpUrl } from '../src/browser'

describe('subscriptionConnectHelpUrl', () => {
  it('opens the section of each path the connect guide offers', () => {
    expect(subscriptionConnectHelpUrl('claude-code')).toBe(
      'https://github.com/pstextiles/ps-office/blob/main/CONNECT.md#claude',
    )
    expect(subscriptionConnectHelpUrl('codex')).toBe(
      'https://github.com/pstextiles/ps-office/blob/main/CONNECT.md#codex',
    )
    expect(subscriptionConnectHelpUrl('free')).toBe(
      'https://github.com/pstextiles/ps-office/blob/main/CONNECT.md#free',
    )
  })

  it('opens the top of the page for anything else a renderer sends', () => {
    for (const path of [null, undefined, '', 'toString', '#evil', 'https://example.com', 42]) {
      expect(subscriptionConnectHelpUrl(path)).toBe(
        'https://github.com/pstextiles/ps-office/blob/main/CONNECT.md',
      )
    }
  })
})
