/**
 * Excel formula text, as stored in a workbook file, to a syntax tree (PS
 * Office's built-in recalculation). File formulas are always in the English,
 * comma-separated form: `_xlfn.` marks newer functions, `_xlpm.` the
 * parameters of LET and LAMBDA, `_xlfn.ANCHORARRAY(A1)` is `A1#` and
 * `_xlfn.SINGLE(x)` is `@x`.
 *
 * Whatever the parser does not know (an intersection by space, a 3-D or
 * external reference, a union) raises FormulaLimit: the calculation then
 * leaves the workbook's stored results alone rather than guess.
 */

/** something the built-in calculation can't do exactly as Excel does */
export class FormulaLimit extends Error {
  constructor(
    readonly what: string,
    readonly formula: string | null = null,
  ) {
    super(
      formula === null
        ? `Formula limit: ${what}`
        : `Formula limit: ${what} in ${formula.slice(0, 200)}`,
    )
  }
}

export const MAX_ROW = 1_048_575
export const MAX_COLUMN = 16_383

export interface RefNode {
  readonly t: 'ref'
  /** null: the sheet of the formula */
  readonly sheet: string | null
  /** 0-based, inclusive; a whole-column reference spans every row, a whole-row one every column */
  readonly r1: number
  readonly c1: number
  readonly r2: number
  readonly c2: number
  /** `$` on each coordinate; a relative one moves with a shared formula */
  readonly r1Abs: boolean
  readonly c1Abs: boolean
  readonly r2Abs: boolean
  readonly c2Abs: boolean
  readonly whole: 'none' | 'columns' | 'rows'
}

export interface StructuredRefNode {
  readonly t: 'sref'
  readonly table: string
  /** [#This Row] / @: the row of the formula */
  readonly thisRow: boolean
  readonly area: 'data' | 'all' | 'headers' | 'totals'
  /** first and last column name; null is every column */
  readonly columns: readonly [string, string] | null
}

export type Node =
  | { readonly t: 'num'; readonly v: number }
  | { readonly t: 'str'; readonly v: string }
  | { readonly t: 'bool'; readonly v: boolean }
  | { readonly t: 'err'; readonly v: string }
  | RefNode
  | StructuredRefNode
  /** a defined name, or a LET/LAMBDA parameter (`param`) */
  | { readonly t: 'name'; readonly name: string; readonly param: boolean }
  /** an omitted argument is null */
  | { readonly t: 'call'; readonly name: string; readonly args: readonly (Node | null)[] }
  | { readonly t: 'un'; readonly op: '-' | '+' | '%' | '@' | '#'; readonly x: Node }
  | { readonly t: 'bin'; readonly op: BinaryOp; readonly a: Node; readonly b: Node }
  | { readonly t: 'array'; readonly rows: readonly (readonly Node[])[] }

export type BinaryOp =
  '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=' | ':'

type Token =
  | { readonly k: 'num'; readonly v: number }
  | { readonly k: 'str'; readonly v: string }
  | { readonly k: 'err'; readonly v: string }
  | { readonly k: 'ref'; readonly node: RefNode }
  | { readonly k: 'sref'; readonly node: StructuredRefNode }
  | { readonly k: 'func'; readonly name: string }
  | { readonly k: 'name'; readonly name: string; readonly param: boolean }
  | { readonly k: 'op'; readonly v: string }
  | { readonly k: 'end' }

const NAME_START = String.raw`A-Za-z_\\\u0080-￿`
const NAME_REST = String.raw`\w.\\\u0080-￿`
const AFTER_REF = String.raw`(?![A-Za-z0-9_.(\u0080-￿])`

const RE = {
  space: /[ \t\r\n]+/y,
  string: /"((?:[^"]|"")*)"/y,
  error: /#(?:N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!|GETTING_DATA)/y,
  quotedSheet: /'((?:[^']|'')+)'!/y,
  plainSheet: new RegExp(`([${NAME_START}][${NAME_REST}]*)!`, 'y'),
  area: new RegExp(
    String.raw`(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Z]{1,3})(\$?)(\d{1,7}))?${AFTER_REF}`,
    'y',
  ),
  columns: new RegExp(String.raw`(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})${AFTER_REF}`, 'y'),
  rows: new RegExp(String.raw`(\$?)(\d{1,7}):(\$?)(\d{1,7})${AFTER_REF}`, 'y'),
  table: new RegExp(`([${NAME_START}][${NAME_REST}]*)\\[`, 'y'),
  func: new RegExp(String.raw`((?:_xlfn\.|_xlws\.|_xludf\.)*)([A-Za-z_][\w.]*)\(`, 'y'),
  name: new RegExp(`(_xlpm\\.)?([${NAME_START}][${NAME_REST}]*)`, 'y'),
  number: /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y,
  op: /<=|>=|<>|[-+*/^&=<>%:,(){};@#]/y,
}

function columnNumber(letters: string): number {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

function tokenize(text: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  const match = (re: RegExp): RegExpExecArray | null => {
    re.lastIndex = i
    return re.exec(text)
  }
  /** a reference at i, with the sheet already read; null when there is none */
  const reference = (sheet: string | null): RefNode | null => {
    let m = match(RE.area)
    if (m) {
      const c1 = columnNumber(m[2]!)
      const r1 = Number(m[4]) - 1
      const c2 = m[6] === undefined ? c1 : columnNumber(m[6])
      const r2 = m[8] === undefined ? r1 : Number(m[8]) - 1
      if (
        c1 <= MAX_COLUMN &&
        c2 <= MAX_COLUMN &&
        r1 >= 0 &&
        r2 >= 0 &&
        r1 <= MAX_ROW &&
        r2 <= MAX_ROW
      ) {
        i = RE.area.lastIndex
        const single = m[6] === undefined
        return {
          t: 'ref',
          sheet,
          r1,
          c1,
          r2,
          c2,
          r1Abs: m[3] === '$',
          c1Abs: m[1] === '$',
          r2Abs: single ? m[3] === '$' : m[7] === '$',
          c2Abs: single ? m[1] === '$' : m[5] === '$',
          whole: 'none',
        }
      }
    }
    m = match(RE.columns)
    if (m && columnNumber(m[2]!) <= MAX_COLUMN && columnNumber(m[4]!) <= MAX_COLUMN) {
      i = RE.columns.lastIndex
      return {
        t: 'ref',
        sheet,
        r1: 0,
        c1: columnNumber(m[2]!),
        r2: MAX_ROW,
        c2: columnNumber(m[4]!),
        r1Abs: true,
        c1Abs: m[1] === '$',
        r2Abs: true,
        c2Abs: m[3] === '$',
        whole: 'columns',
      }
    }
    m = match(RE.rows)
    if (
      m &&
      Number(m[2]) >= 1 &&
      Number(m[4]) >= 1 &&
      Number(m[2]) - 1 <= MAX_ROW &&
      Number(m[4]) - 1 <= MAX_ROW
    ) {
      i = RE.rows.lastIndex
      return {
        t: 'ref',
        sheet,
        r1: Number(m[2]) - 1,
        c1: 0,
        r2: Number(m[4]) - 1,
        c2: MAX_COLUMN,
        r1Abs: m[1] === '$',
        c1Abs: true,
        r2Abs: m[3] === '$',
        c2Abs: true,
        whole: 'rows',
      }
    }
    return null
  }

  while (i < text.length) {
    let m = match(RE.space)
    if (m) {
      i = RE.space.lastIndex
      continue
    }
    if ((m = match(RE.string))) {
      tokens.push({ k: 'str', v: m[1]!.replace(/""/g, '"') })
      i = RE.string.lastIndex
      continue
    }
    if ((m = match(RE.error))) {
      tokens.push({ k: 'err', v: m[0] })
      i = RE.error.lastIndex
      continue
    }
    const quoted = match(RE.quotedSheet)
    const plain = quoted ? null : match(RE.plainSheet)
    if (quoted || plain) {
      const sheet = quoted ? quoted[1]!.replace(/''/g, "'") : plain![1]!
      if (/[[\]]/.test(sheet) || sheet.includes(':'))
        throw new FormulaLimit('an external or 3-D reference', text)
      const save = i
      i = quoted ? RE.quotedSheet.lastIndex : RE.plainSheet.lastIndex
      const ref = reference(sheet)
      if (!ref) {
        i = save
        throw new FormulaLimit('a sheet-qualified name', text)
      }
      tokens.push({ k: 'ref', node: ref })
      continue
    }
    if ((m = match(RE.table))) {
      const table = m[1]!
      const parsed = structuredReference(text, RE.table.lastIndex, table)
      tokens.push({ k: 'sref', node: parsed.node })
      i = parsed.end
      continue
    }
    if ((m = match(RE.func))) {
      tokens.push({ k: 'func', name: m[2]!.toUpperCase() })
      i = RE.func.lastIndex
      continue
    }
    // a number before a reference: "1:3" is a row range, "12" a number, "1E5" a number
    const rowRange = match(RE.rows)
    if (rowRange) {
      const ref = reference(null)
      if (ref) {
        tokens.push({ k: 'ref', node: ref })
        continue
      }
    }
    if ((m = match(RE.number))) {
      // "1E5" is a number, but "A1" style text never starts with a digit
      tokens.push({ k: 'num', v: Number(m[0]) })
      i = RE.number.lastIndex
      continue
    }
    const ref = reference(null)
    if (ref) {
      tokens.push({ k: 'ref', node: ref })
      continue
    }
    if ((m = match(RE.name))) {
      tokens.push({ k: 'name', name: m[2]!, param: m[1] !== undefined })
      i = RE.name.lastIndex
      continue
    }
    if ((m = match(RE.op))) {
      tokens.push({ k: 'op', v: m[0] })
      i = RE.op.lastIndex
      continue
    }
    throw new FormulaLimit(`an unknown character '${text[i]}'`, text)
  }
  tokens.push({ k: 'end' })
  return tokens
}

/** `Table[…]` from just after the opening bracket: the node and the index after the closing bracket */
function structuredReference(
  text: string,
  start: number,
  table: string,
): { node: StructuredRefNode; end: number } {
  let i = start
  /** an item's text up to its unescaped `]`; `'` escapes the next character */
  const item = (): string => {
    let out = ''
    while (i < text.length && text[i] !== ']') {
      if (text[i] === "'" && i + 1 < text.length) {
        out += text[i + 1]
        i += 2
      } else if (text[i] === '[') {
        throw new FormulaLimit('a structured reference', text)
      } else {
        out += text[i]
        i++
      }
    }
    if (text[i] !== ']') throw new FormulaLimit('an unclosed structured reference', text)
    i++
    return out
  }
  let thisRow = false
  let area: StructuredRefNode['area'] = 'data'
  const columns: string[] = []
  const take = (raw: string): void => {
    const special = /^#(This Row|All|Data|Headers|Totals)$/i.exec(raw.trim())
    if (special) {
      const which = special[1]!.toLowerCase()
      if (which === 'this row') thisRow = true
      else if (which === 'all') area = 'all'
      else if (which === 'headers') area = 'headers'
      else if (which === 'totals') area = 'totals'
      return
    }
    if (raw.startsWith('@')) {
      thisRow = true
      if (raw.length > 1) columns.push(raw.slice(1))
      return
    }
    columns.push(raw)
  }
  if (text[i] === '[') {
    // a list of bracketed items: [[#This Row],[Col]], [[A]:[B]]
    let rangeNext = false
    for (;;) {
      while (text[i] === ' ') i++
      if (text[i] === '@') {
        thisRow = true
        i++
        continue
      }
      if (text[i] !== '[') throw new FormulaLimit('a structured reference', text)
      i++
      const before = columns.length
      take(item())
      if (rangeNext && columns.length !== before + 1)
        throw new FormulaLimit('a structured reference', text)
      rangeNext = false
      while (text[i] === ' ') i++
      if (text[i] === ',') {
        i++
        continue
      }
      if (text[i] === ':') {
        rangeNext = true
        i++
        continue
      }
      if (text[i] === ']') {
        i++
        break
      }
      throw new FormulaLimit('a structured reference', text)
    }
  } else {
    take(item())
  }
  if (columns.length > 2)
    throw new FormulaLimit('a structured reference with several columns', text)
  return {
    node: {
      t: 'sref',
      table,
      thisRow,
      area,
      columns: columns.length === 0 ? null : [columns[0]!, columns[columns.length - 1]!],
    },
    end: i,
  }
}

const COMPARISON = new Set(['=', '<>', '<', '>', '<=', '>='])

class Parser {
  private pos = 0
  constructor(
    private readonly tokens: readonly Token[],
    private readonly text: string,
  ) {}

  private peek(): Token {
    return this.tokens[this.pos]!
  }
  private isOp(v: string): boolean {
    const tok = this.peek()
    return tok.k === 'op' && tok.v === v
  }
  private fail(what: string): never {
    throw new FormulaLimit(what, this.text)
  }

  parse(): Node {
    const node = this.comparison()
    if (this.peek().k !== 'end') this.fail('unexpected text after the formula')
    return node
  }

  private comparison(): Node {
    let left = this.concat()
    for (;;) {
      const tok = this.peek()
      if (tok.k !== 'op' || !COMPARISON.has(tok.v)) return left
      this.pos++
      left = { t: 'bin', op: tok.v as BinaryOp, a: left, b: this.concat() }
    }
  }
  private concat(): Node {
    let left = this.additive()
    while (this.isOp('&')) {
      this.pos++
      left = { t: 'bin', op: '&', a: left, b: this.additive() }
    }
    return left
  }
  private additive(): Node {
    let left = this.multiplicative()
    for (;;) {
      const op = this.isOp('+') ? '+' : this.isOp('-') ? '-' : null
      if (!op) return left
      this.pos++
      left = { t: 'bin', op, a: left, b: this.multiplicative() }
    }
  }
  private multiplicative(): Node {
    let left = this.power()
    for (;;) {
      const op = this.isOp('*') ? '*' : this.isOp('/') ? '/' : null
      if (!op) return left
      this.pos++
      left = { t: 'bin', op, a: left, b: this.power() }
    }
  }
  /** `^` is left-associative, and a sign binds tighter: -2^2 is 4 */
  private power(): Node {
    let left = this.unary()
    while (this.isOp('^')) {
      this.pos++
      left = { t: 'bin', op: '^', a: left, b: this.unary() }
    }
    return left
  }
  private unary(): Node {
    if (this.isOp('-') || this.isOp('+')) {
      const op = (this.peek() as { v: '-' | '+' }).v
      this.pos++
      return { t: 'un', op, x: this.unary() }
    }
    if (this.isOp('@')) {
      this.pos++
      return { t: 'un', op: '@', x: this.unary() }
    }
    return this.postfix()
  }
  private postfix(): Node {
    let node = this.range()
    while (this.isOp('%')) {
      this.pos++
      node = { t: 'un', op: '%', x: node }
    }
    return node
  }
  private range(): Node {
    let node = this.primary()
    for (;;) {
      if (this.isOp(':')) {
        this.pos++
        node = { t: 'bin', op: ':', a: node, b: this.primary() }
      } else if (this.isOp('#')) {
        this.pos++
        node = { t: 'un', op: '#', x: node }
      } else return node
    }
  }

  private primary(): Node {
    const tok = this.peek()
    switch (tok.k) {
      case 'num':
        this.pos++
        return { t: 'num', v: tok.v }
      case 'str':
        this.pos++
        return { t: 'str', v: tok.v }
      case 'err':
        this.pos++
        return { t: 'err', v: tok.v }
      case 'ref':
        this.pos++
        return tok.node
      case 'sref':
        this.pos++
        return tok.node
      case 'name': {
        this.pos++
        const upper = tok.name.toUpperCase()
        if (!tok.param && (upper === 'TRUE' || upper === 'FALSE'))
          return { t: 'bool', v: upper === 'TRUE' }
        return { t: 'name', name: tok.name, param: tok.param }
      }
      case 'func': {
        this.pos++
        const args: (Node | null)[] = []
        if (this.isOp(')')) {
          this.pos++
        } else {
          for (;;) {
            if (this.isOp(',')) {
              args.push(null)
              this.pos++
              continue
            }
            if (this.isOp(')')) {
              args.push(null)
              this.pos++
              break
            }
            args.push(this.comparison())
            if (this.isOp(',')) {
              this.pos++
              if (this.isOp(')')) {
                args.push(null)
                this.pos++
                break
              }
              continue
            }
            if (this.isOp(')')) {
              this.pos++
              break
            }
            this.fail('a function call that does not close')
          }
        }
        if (tok.name === 'ANCHORARRAY' && args.length === 1 && args[0])
          return { t: 'un', op: '#', x: args[0] }
        if (tok.name === 'SINGLE' && args.length === 1 && args[0])
          return { t: 'un', op: '@', x: args[0] }
        return { t: 'call', name: tok.name, args }
      }
      case 'op':
        if (tok.v === '(') {
          this.pos++
          const inner = this.comparison()
          if (this.isOp(',')) this.fail('a union of references')
          if (!this.isOp(')')) this.fail('a parenthesis that does not close')
          this.pos++
          return inner
        }
        if (tok.v === '{') return this.arrayConstant()
        return this.fail(`an unexpected '${tok.v}'`)
      default:
        return this.fail('a formula that ends early')
    }
  }

  private arrayConstant(): Node {
    this.pos++
    const rows: Node[][] = [[]]
    for (;;) {
      rows[rows.length - 1]!.push(this.constant())
      if (this.isOp(',')) this.pos++
      else if (this.isOp(';')) {
        this.pos++
        rows.push([])
      } else if (this.isOp('}')) {
        this.pos++
        break
      } else this.fail('an array constant')
    }
    const width = rows[0]!.length
    if (rows.some((row) => row.length !== width)) this.fail('an array constant with uneven rows')
    return { t: 'array', rows }
  }
  private constant(): Node {
    let sign = 1
    if (this.isOp('-') || this.isOp('+')) {
      if (this.isOp('-')) sign = -1
      this.pos++
      const num = this.peek()
      if (num.k !== 'num') this.fail('an array constant')
      this.pos++
      return { t: 'num', v: sign * (num as { v: number }).v }
    }
    const tok = this.peek()
    this.pos++
    if (tok.k === 'num') return { t: 'num', v: tok.v }
    if (tok.k === 'str') return { t: 'str', v: tok.v }
    if (tok.k === 'err') return { t: 'err', v: tok.v }
    if (tok.k === 'name' && /^(true|false)$/i.test(tok.name))
      return { t: 'bool', v: /^true$/i.test(tok.name) }
    return this.fail('an array constant')
  }
}

/** Parses a formula as stored in a file (no leading `=`). Raises FormulaLimit for syntax it does not know. */
export function parseFormula(text: string): Node {
  const body = text.startsWith('=') ? text.slice(1) : text
  return new Parser(tokenize(body), body).parse()
}
