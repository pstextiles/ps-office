/**
 * Parser for the subset of M the built-in engine runs: section documents with
 * shared queries, let/if/each/try/error, function literals, records, lists
 * (with "a".."z" ranges), field and item access, and the usual operators
 * with M's precedence.
 */
import { tokenize, type Token } from './lexer'
import { MError } from './values'

export type Node =
  | { readonly t: 'lit'; readonly value: null | boolean | number | string }
  /** something the built-in engine can't evaluate exactly; evaluating it raises EngineLimit */
  | { readonly t: 'limit'; readonly what: string; readonly input: string }
  | { readonly t: 'id'; readonly name: string }
  | {
      readonly t: 'list'
      readonly items: readonly (
        Node | { readonly t: 'range'; readonly from: Node; readonly to: Node }
      )[]
    }
  | {
      readonly t: 'record'
      readonly fields: readonly { readonly name: string; readonly value: Node }[]
    }
  | {
      readonly t: 'field'
      readonly target: Node | null
      readonly name: string
      readonly optional: boolean
    }
  | { readonly t: 'item'; readonly target: Node; readonly key: Node; readonly optional: boolean }
  | { readonly t: 'call'; readonly callee: Node; readonly args: readonly Node[] }
  | {
      readonly t: 'fn'
      readonly params: readonly { readonly name: string; readonly optional: boolean }[]
      readonly body: Node
    }
  | {
      readonly t: 'let'
      readonly bindings: readonly { readonly name: string; readonly value: Node }[]
      readonly body: Node
    }
  | { readonly t: 'if'; readonly cond: Node; readonly then: Node; readonly else: Node }
  | { readonly t: 'try'; readonly body: Node; readonly otherwise: Node | null }
  | { readonly t: 'error'; readonly value: Node }
  | { readonly t: 'not'; readonly value: Node }
  | { readonly t: 'neg'; readonly value: Node }
  | { readonly t: 'bin'; readonly op: string; readonly left: Node; readonly right: Node }
  | { readonly t: 'as'; readonly value: Node; readonly type: TypeRef }
  | { readonly t: 'is'; readonly value: Node; readonly type: TypeRef }
  | { readonly t: 'type'; readonly type: TypeRef }

export interface TypeRef {
  readonly name: string
  readonly nullable: boolean
}

export interface SectionMember {
  readonly name: string
  readonly value: Node
}

const PRIMITIVE_TYPES = new Set([
  'any',
  'anynonnull',
  'binary',
  'date',
  'datetime',
  'datetimezone',
  'duration',
  'function',
  'list',
  'logical',
  'none',
  'null',
  'number',
  'record',
  'table',
  'text',
  'time',
  'type',
])

class Parser {
  private pos = 0
  constructor(private readonly tokens: readonly Token[]) {}

  private peek(offset = 0): Token {
    return this.tokens[Math.min(this.pos + offset, this.tokens.length - 1)]!
  }
  private next(): Token {
    const tok = this.peek()
    if (tok.kind !== 'eof') this.pos++
    return tok
  }
  private isOp(value: string, offset = 0): boolean {
    const tok = this.peek(offset)
    return tok.kind === 'op' && tok.value === value
  }
  private isKw(value: string, offset = 0): boolean {
    const tok = this.peek(offset)
    return tok.kind === 'keyword' && tok.value === value
  }
  private expectOp(value: string): void {
    if (!this.isOp(value)) this.fail(`"${value}"`)
    this.pos++
  }
  private expectKw(value: string): void {
    if (!this.isKw(value)) this.fail(`"${value}"`)
    this.pos++
  }
  private fail(expected: string): never {
    const tok = this.peek()
    throw new MError(
      'Expression.SyntaxError',
      `Expected ${expected} but found ${tok.kind === 'eof' ? 'the end' : `"${tok.value}"`} at ${tok.pos}.`,
    )
  }

  /** identifier, quoted identifier, or a keyword used as a name ([Kind], [Type]) */
  private name(allowKeyword = false): string {
    const tok = this.peek()
    if (tok.kind === 'ident' || tok.kind === 'quoted' || (allowKeyword && tok.kind === 'keyword')) {
      this.pos++
      return tok.value
    }
    return this.fail('a name')
  }

  section(): SectionMember[] {
    const members: SectionMember[] = []
    if (this.isKw('section')) {
      this.pos++
      if (!this.isOp(';')) this.name()
      this.expectOp(';')
    }
    while (this.peek().kind !== 'eof') {
      if (this.isOp('[')) this.skipAttributes()
      if (this.isKw('shared')) this.pos++
      const name = this.name()
      this.expectOp('=')
      const value = this.expression()
      this.expectOp(';')
      members.push({ name, value })
    }
    return members
  }

  /** member attributes like [ Description = "…" ] are skipped */
  private skipAttributes(): void {
    let depth = 0
    do {
      const tok = this.next()
      if (tok.kind === 'op' && tok.value === '[') depth++
      if (tok.kind === 'op' && tok.value === ']') depth--
      if (tok.kind === 'eof') this.fail('"]"')
    } while (depth > 0)
  }

  expression(): Node {
    if (this.isKw('let')) return this.letExpr()
    if (this.isKw('if')) return this.ifExpr()
    if (this.isKw('each')) {
      this.pos++
      return { t: 'fn', params: [{ name: '_', optional: false }], body: this.expression() }
    }
    if (this.isKw('try')) {
      this.pos++
      const body = this.expression()
      if (this.isKw('otherwise')) {
        this.pos++
        return { t: 'try', body, otherwise: this.expression() }
      }
      return { t: 'try', body, otherwise: null }
    }
    if (this.isKw('error')) {
      this.pos++
      return { t: 'error', value: this.expression() }
    }
    if (this.isOp('(') && this.looksLikeFunction()) return this.functionExpr()
    return this.logicalOr()
  }

  /** after a "(" — a parameter list is followed by "=>" or "as <type> =>" */
  private looksLikeFunction(): boolean {
    let depth = 0
    let i = this.pos
    for (; i < this.tokens.length; i++) {
      const tok = this.tokens[i]!
      if (tok.kind === 'eof') return false
      if (tok.kind === 'op' && (tok.value === '(' || tok.value === '[' || tok.value === '{'))
        depth++
      if (tok.kind === 'op' && (tok.value === ')' || tok.value === ']' || tok.value === '}')) {
        depth--
        if (depth === 0) break
      }
    }
    const after = this.tokens[i + 1]
    if (!after) return false
    if (after.kind === 'op' && after.value === '=>') return true
    if (after.kind === 'keyword' && after.value === 'as') {
      // "as nullable number =>"
      let j = i + 2
      while (this.tokens[j] && !(this.tokens[j]!.kind === 'op' && this.tokens[j]!.value === '=>')) {
        const tok = this.tokens[j]!
        if (tok.kind === 'op' && tok.value !== '.') return false
        j++
      }
      return Boolean(this.tokens[j])
    }
    return false
  }

  private functionExpr(): Node {
    this.expectOp('(')
    const params: { name: string; optional: boolean }[] = []
    while (!this.isOp(')')) {
      let optional = false
      if (
        this.peek().kind === 'ident' &&
        this.peek().value === 'optional' &&
        this.peek(1).kind !== 'op'
      ) {
        this.pos++
        optional = true
      }
      params.push({ name: this.name(), optional })
      if (this.isKw('as')) {
        this.pos++
        this.typeRef()
      }
      if (this.isOp(',')) this.pos++
      else break
    }
    this.expectOp(')')
    if (this.isKw('as')) {
      this.pos++
      this.typeRef()
    }
    this.expectOp('=>')
    return { t: 'fn', params, body: this.expression() }
  }

  private letExpr(): Node {
    this.expectKw('let')
    const bindings: { name: string; value: Node }[] = []
    for (;;) {
      const name = this.name()
      this.expectOp('=')
      bindings.push({ name, value: this.expression() })
      if (this.isOp(',')) {
        this.pos++
        continue
      }
      break
    }
    this.expectKw('in')
    return { t: 'let', bindings, body: this.expression() }
  }

  private ifExpr(): Node {
    this.expectKw('if')
    const cond = this.expression()
    this.expectKw('then')
    const then = this.expression()
    this.expectKw('else')
    return { t: 'if', cond, then, else: this.expression() }
  }

  private logicalOr(): Node {
    let left = this.logicalAnd()
    while (this.isKw('or')) {
      this.pos++
      left = { t: 'bin', op: 'or', left, right: this.logicalAnd() }
    }
    return left
  }

  private logicalAnd(): Node {
    let left = this.isExpr()
    while (this.isKw('and')) {
      this.pos++
      left = { t: 'bin', op: 'and', left, right: this.isExpr() }
    }
    return left
  }

  private isExpr(): Node {
    let left = this.asExpr()
    while (this.isKw('is')) {
      this.pos++
      left = { t: 'is', value: left, type: this.typeRef() }
    }
    return left
  }

  private asExpr(): Node {
    let left = this.equality()
    while (this.isKw('as')) {
      this.pos++
      left = { t: 'as', value: left, type: this.typeRef() }
    }
    return left
  }

  private equality(): Node {
    let left = this.relational()
    while (this.isOp('=') || this.isOp('<>')) {
      const op = this.next().value
      left = { t: 'bin', op, left, right: this.relational() }
    }
    return left
  }

  private relational(): Node {
    let left = this.additive()
    while (this.isOp('<') || this.isOp('>') || this.isOp('<=') || this.isOp('>=')) {
      const op = this.next().value
      left = { t: 'bin', op, left, right: this.additive() }
    }
    return left
  }

  private additive(): Node {
    let left = this.multiplicative()
    while (this.isOp('+') || this.isOp('-') || this.isOp('&')) {
      const op = this.next().value
      left = { t: 'bin', op, left, right: this.multiplicative() }
    }
    return left
  }

  private multiplicative(): Node {
    let left = this.metadata()
    while (this.isOp('*') || this.isOp('/')) {
      const op = this.next().value
      left = { t: 'bin', op, left, right: this.metadata() }
    }
    return left
  }

  private metadata(): Node {
    const value = this.unary()
    if (this.isKw('meta')) {
      this.pos++
      this.unary() // metadata is not used by the engine
    }
    return value
  }

  private unary(): Node {
    if (this.isOp('-')) {
      this.pos++
      return { t: 'neg', value: this.unary() }
    }
    if (this.isOp('+')) {
      this.pos++
      return this.unary()
    }
    if (this.isKw('not')) {
      this.pos++
      return { t: 'not', value: this.unary() }
    }
    if (this.isKw('type')) {
      this.pos++
      return { t: 'type', type: this.typeRef() }
    }
    return this.postfix(this.primary())
  }

  /** a primitive type, `nullable <type>`, or a table/list/record type (read loosely) */
  private typeRef(): TypeRef {
    let nullable = false
    if (this.peek().kind === 'ident' && this.peek().value === 'nullable') {
      this.pos++
      nullable = true
    }
    const tok = this.peek()
    if (tok.kind === 'keyword' && (tok.value === 'null' || tok.value === 'type')) {
      this.pos++
      return { name: tok.value, nullable }
    }
    if (tok.kind === 'ident') {
      this.pos++
      const lower = tok.value.toLowerCase()
      if ((lower === 'table' || lower === 'record') && this.isOp('[')) this.skipBalanced('[', ']')
      // Int64.Type, Number.Type, Currency.Type, Percentage.Type …
      const name = PRIMITIVE_TYPES.has(lower) ? lower : tok.value
      return { name, nullable }
    }
    if (this.isOp('{')) {
      this.skipBalanced('{', '}')
      return { name: 'list', nullable }
    }
    if (this.isOp('[')) {
      this.skipBalanced('[', ']')
      return { name: 'record', nullable }
    }
    return this.fail('a type')
  }

  private skipBalanced(open: string, close: string): void {
    let depth = 0
    do {
      const tok = this.next()
      if (tok.kind === 'op' && tok.value === open) depth++
      if (tok.kind === 'op' && tok.value === close) depth--
      if (tok.kind === 'eof') this.fail(`"${close}"`)
    } while (depth > 0)
  }

  private primary(): Node {
    const tok = this.peek()
    switch (tok.kind) {
      case 'number':
        this.pos++
        return numberLiteral(tok.value)
      case 'text':
        this.pos++
        return { t: 'lit', value: tok.value }
      case 'ident':
      case 'quoted':
        this.pos++
        return { t: 'id', name: tok.value }
      case 'hash':
        this.pos++
        return { t: 'id', name: tok.value }
      case 'keyword':
        if (tok.value === 'null') {
          this.pos++
          return { t: 'lit', value: null }
        }
        if (tok.value === 'true' || tok.value === 'false') {
          this.pos++
          return { t: 'lit', value: tok.value === 'true' }
        }
        return this.fail('an expression')
      case 'op':
        if (tok.value === '(') {
          this.pos++
          const inner = this.expression()
          this.expectOp(')')
          return inner
        }
        if (tok.value === '{') return this.listExpr()
        if (tok.value === '[') return this.recordOrImplicitField()
        if (tok.value === '@') {
          this.pos++
          return { t: 'id', name: this.name() }
        }
        return this.fail('an expression')
      default:
        return this.fail('an expression')
    }
  }

  private listExpr(): Node {
    this.expectOp('{')
    const items: (Node | { t: 'range'; from: Node; to: Node })[] = []
    while (!this.isOp('}')) {
      const from = this.expression()
      if (this.isOp('..')) {
        this.pos++
        items.push({ t: 'range', from, to: this.expression() })
      } else items.push(from)
      if (this.isOp(',')) this.pos++
      else break
    }
    this.expectOp('}')
    return { t: 'list', items }
  }

  /** `[a = 1, b = 2]` is a record; `[Name]` alone is a field of `_` */
  private recordOrImplicitField(): Node {
    const nameTok = this.peek(1)
    const isName =
      nameTok.kind === 'ident' || nameTok.kind === 'quoted' || nameTok.kind === 'keyword'
    if (this.isOp(']', 1)) {
      this.pos += 2
      return { t: 'record', fields: [] }
    }
    // an implicit field access may have a multi-word generalized name: [PCS Cortadas]
    if (isName) {
      let j = 1
      while (this.tokens[this.pos + j] && !['op', 'eof'].includes(this.tokens[this.pos + j]!.kind))
        j++
      const after = this.peek(j)
      if (after.kind === 'op' && after.value === ']') {
        const words: string[] = []
        for (let k = 1; k < j; k++) words.push(this.peek(k).value)
        this.pos += j + 1
        const optional = this.isOp('?')
        if (optional) this.pos++
        return { t: 'field', target: null, name: this.joinGeneralized(words), optional }
      }
    }
    this.expectOp('[')
    const fields: { name: string; value: Node }[] = []
    while (!this.isOp(']')) {
      const words: string[] = []
      while (!this.isOp('=')) words.push(this.next().value)
      this.expectOp('=')
      fields.push({ name: this.joinGeneralized(words), value: this.expression() })
      if (this.isOp(',')) this.pos++
      else break
    }
    this.expectOp(']')
    return { t: 'record', fields }
  }

  /** a generalized identifier keeps its words separated by single spaces */
  private joinGeneralized(words: readonly string[]): string {
    return words.join(' ')
  }

  private postfix(target: Node): Node {
    for (;;) {
      if (this.isOp('(')) {
        this.pos++
        const args: Node[] = []
        while (!this.isOp(')')) {
          args.push(this.expression())
          if (this.isOp(',')) this.pos++
          else break
        }
        this.expectOp(')')
        target = { t: 'call', callee: target, args }
        continue
      }
      if (this.isOp('{')) {
        this.pos++
        const key = this.expression()
        this.expectOp('}')
        const optional = this.isOp('?')
        if (optional) this.pos++
        target = { t: 'item', target, key, optional }
        continue
      }
      if (this.isOp('[')) {
        const field = this.recordOrImplicitField()
        if (field.t !== 'field') this.fail('a field name')
        target = { t: 'field', target, name: field.name, optional: field.optional }
        continue
      }
      return target
    }
  }
}

/** Reads a section document (`section Section1; shared Q = …;`). */
export function parseSection(src: string): SectionMember[] {
  return new Parser(tokenize(src)).section()
}

/** Reads one expression. */
export function parseExpression(src: string): Node {
  const parser = new Parser(tokenize(src))
  return parser.expression()
}

/**
 * Power Query keeps a number literal with more than 15 significant digits as
 * written (Text.From(12345678901234567) is "12345678901234567"), which a
 * double can't hold; such a literal is an engine limit unless the double
 * reads back as the same digits.
 */
function numberLiteral(text: string): Node {
  const value = Number(text)
  const digits = (t: string): string => {
    const m = /^(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(t)
    if (!m) return t
    return ((m[1] ?? '') + (m[2] ?? '')).replace(/^0+/, '').replace(/0+$/, '')
  }
  const written = digits(text)
  if (written.length <= 15 || /^0x/i.test(text)) return { t: 'lit', value }
  const exact = value.toPrecision(Math.min(written.length, 100))
  const back = digits(exact.replace(/^(\d)\.?(\d*)e[+-]\d+$/, '$1$2'))
  return back === written
    ? { t: 'lit', value }
    : { t: 'limit', what: 'number literal beyond double precision', input: text }
}
