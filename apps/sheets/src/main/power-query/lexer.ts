/**
 * Tokens of the M language, as far as the built-in engine reads it.
 * Identifiers may be dotted (`Table.SelectRows`, `Int64.Type`) or quoted
 * (`#"Changed Type"`); text literals use "" for a quote and #(lf) escapes.
 */
import { MError } from './values'

export type TokenKind = 'ident' | 'quoted' | 'keyword' | 'hash' | 'number' | 'text' | 'op' | 'eof'

export interface Token {
  readonly kind: TokenKind
  readonly value: string
  readonly pos: number
}

const KEYWORDS = new Set([
  'and',
  'as',
  'each',
  'else',
  'error',
  'false',
  'if',
  'in',
  'is',
  'let',
  'meta',
  'not',
  'null',
  'or',
  'otherwise',
  'section',
  'shared',
  'then',
  'true',
  'try',
  'type',
])

// longest first
const OPERATORS = [
  '...',
  '=>',
  '<=',
  '>=',
  '<>',
  '..',
  '=',
  '<',
  '>',
  '+',
  '-',
  '*',
  '/',
  '&',
  '?',
  ',',
  ';',
  '(',
  ')',
  '[',
  ']',
  '{',
  '}',
  '@',
  '!',
]

const ESCAPES: Record<string, string> = { cr: '\r', lf: '\n', tab: '\t', '#': '#' }

function isIdentStart(ch: string): boolean {
  return /[\p{L}_]/u.test(ch)
}
function isIdentPart(ch: string): boolean {
  return /[\p{L}\p{N}_\p{Mn}\p{Mc}\p{Pc}]/u.test(ch)
}

/** #(cr,lf), #(tab), #(0041) inside a text literal */
function readEscape(src: string, at: number): { text: string; end: number } | null {
  if (src[at] !== '#' || src[at + 1] !== '(') return null
  const close = src.indexOf(')', at + 2)
  if (close < 0) return null
  const parts = src
    .slice(at + 2, close)
    .split(',')
    .map((p) => p.trim())
  let out = ''
  for (const part of parts) {
    const named = ESCAPES[part.toLowerCase()]
    if (named !== undefined) out += named
    else if (/^[0-9a-fA-F]{4,8}$/.test(part)) out += String.fromCodePoint(parseInt(part, 16))
    else return null
  }
  return { text: out, end: close + 1 }
}

export function tokenize(src: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  while (i < src.length) {
    const ch = src[i]!
    // whitespace and comments
    if (/\s/.test(ch) || ch === '\uFEFF') {
      i++
      continue
    }
    if (src.startsWith('//', i)) {
      while (i < src.length && src[i] !== '\n') i++
      continue
    }
    if (src.startsWith('/*', i)) {
      const end = src.indexOf('*/', i + 2)
      i = end < 0 ? src.length : end + 2
      continue
    }
    const start = i
    // text literal
    if (ch === '"') {
      let text = ''
      i++
      for (;;) {
        if (i >= src.length)
          throw new MError('Expression.SyntaxError', 'Unterminated text literal.')
        const c = src[i]!
        if (c === '"') {
          if (src[i + 1] === '"') {
            text += '"'
            i += 2
            continue
          }
          i++
          break
        }
        const esc = c === '#' ? readEscape(src, i) : null
        if (esc) {
          text += esc.text
          i = esc.end
          continue
        }
        text += c
        i++
      }
      tokens.push({ kind: 'text', value: text, pos: start })
      continue
    }
    // #"quoted identifier" and #keywords (#table, #date, …)
    if (ch === '#') {
      if (src[i + 1] === '"') {
        let name = ''
        i += 2
        for (;;) {
          if (i >= src.length)
            throw new MError('Expression.SyntaxError', 'Unterminated quoted identifier.')
          const c = src[i]!
          if (c === '"') {
            if (src[i + 1] === '"') {
              name += '"'
              i += 2
              continue
            }
            i++
            break
          }
          name += c
          i++
        }
        tokens.push({ kind: 'quoted', value: name, pos: start })
        continue
      }
      let j = i + 1
      while (j < src.length && /[A-Za-z]/.test(src[j]!)) j++
      if (j > i + 1) {
        tokens.push({ kind: 'hash', value: src.slice(i, j), pos: start })
        i = j
        continue
      }
    }
    // number: 12, 1.5, .5, 1e3, 0x1F
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] ?? ''))) {
      const hex = /^0[xX][0-9a-fA-F]+/.exec(src.slice(i))
      if (hex) {
        tokens.push({ kind: 'number', value: String(parseInt(hex[0], 16)), pos: start })
        i += hex[0].length
        continue
      }
      // a ".." range operator must not be read as a decimal point
      const m = /^(?:[0-9]+(?:\.(?!\.)[0-9]+)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/.exec(src.slice(i))!
      tokens.push({ kind: 'number', value: m[0], pos: start })
      i += m[0].length
      continue
    }
    // identifier, possibly dotted: Table.SelectRows, Int64.Type
    if (isIdentStart(ch)) {
      let j = i + 1
      for (;;) {
        while (j < src.length && isIdentPart(src[j]!)) j++
        if (src[j] === '.' && j + 1 < src.length && isIdentStart(src[j + 1]!)) {
          j++
          continue
        }
        break
      }
      const word = src.slice(i, j)
      tokens.push({ kind: KEYWORDS.has(word) ? 'keyword' : 'ident', value: word, pos: start })
      i = j
      continue
    }
    const op = OPERATORS.find((o) => src.startsWith(o, i))
    if (op) {
      tokens.push({ kind: 'op', value: op, pos: start })
      i += op.length
      continue
    }
    throw new MError('Expression.SyntaxError', `Unexpected character "${ch}" at ${i}.`)
  }
  tokens.push({ kind: 'eof', value: '', pos: src.length })
  return tokens
}
