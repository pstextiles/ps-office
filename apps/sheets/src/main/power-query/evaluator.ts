/**
 * Evaluates parsed M. Bindings (let, section members, record fields) are
 * lazy and memoized, so a step that is never needed never runs and never
 * fails, as in Power Query.
 */
import type { Node, SectionMember } from './parser'
import {
  EngineLimit,
  MDate,
  MDateTime,
  MDuration,
  MError,
  asList,
  asRecord,
  asTable,
  describe,
  fieldIndex,
  fn,
  invoke,
  isList,
  isRecord,
  isTable,
  kindOf,
  lazyRecord,
  list,
  record,
  rowRecord,
  type,
  type MValue,
} from './values'
import { conformsTo, valuesEqual, compareValues } from './compare'

type Thunk = () => MValue

export class Scope {
  private readonly slots = new Map<string, Thunk>()
  constructor(private readonly parent: Scope | null) {}

  /** binds `name` to a lazily computed, memoized value */
  define(name: string, compute: Thunk): void {
    let state: { done: true; value: MValue } | { done: false; error: unknown } | null = null
    let running = false
    this.slots.set(name, () => {
      if (state) {
        if (state.done) return state.value
        throw state.error
      }
      if (running)
        throw new MError(
          'Expression.Error',
          `A cyclic reference was encountered while evaluating ${name}.`,
        )
      running = true
      try {
        state = { done: true, value: compute() }
      } catch (error) {
        state = { done: false, error }
      } finally {
        running = false
      }
      if (state.done) return state.value
      throw state.error
    })
  }

  value(name: string, value: MValue): void {
    this.slots.set(name, () => value)
  }

  lookup(name: string): Thunk | undefined {
    return this.slots.get(name) ?? this.parent?.lookup(name)
  }
}

export class Evaluator {
  constructor(private readonly globals: ReadonlyMap<string, MValue>) {}

  /** a scope holding every member of a section, each evaluated on demand */
  sectionScope(members: readonly SectionMember[]): Scope {
    const scope = new Scope(null)
    for (const member of members) scope.define(member.name, () => this.eval(member.value, scope))
    return scope
  }

  eval(node: Node, scope: Scope): MValue {
    switch (node.t) {
      case 'lit':
        return node.value
      case 'limit':
        throw new EngineLimit(node.what, node.input)
      case 'id': {
        const local = scope.lookup(node.name)
        if (local) return local()
        if (this.globals.has(node.name)) return this.globals.get(node.name)!
        throw new MError('Expression.Error', `The name '${node.name}' wasn't recognized.`)
      }
      case 'list': {
        const items: MValue[] = []
        for (const item of node.items) {
          if ('t' in item && item.t === 'range') {
            items.push(...rangeItems(this.eval(item.from, scope), this.eval(item.to, scope)))
          } else items.push(this.eval(item as Node, scope))
        }
        return list(items)
      }
      case 'record': {
        // fields see each other (and the enclosing scope), lazily
        const inner = new Scope(scope)
        const names = node.fields.map((f) => f.name)
        for (const f of node.fields) inner.define(f.name, () => this.eval(f.value, inner))
        return lazyRecord(names, (i) => inner.lookup(names[i]!)!())
      }
      case 'field': {
        const target =
          node.target === null ? this.implicitRow(scope) : this.eval(node.target, scope)
        return fieldOf(target, node.name, node.optional)
      }
      case 'item':
        return itemOf(this.eval(node.target, scope), this.eval(node.key, scope), node.optional)
      case 'call': {
        const callee = this.eval(node.callee, scope)
        const args = node.args.map((a) => this.eval(a, scope))
        return invoke(callee, args)
      }
      case 'fn':
        return fn('', node.params.filter((p) => !p.optional).length, node.params.length, (args) => {
          const inner = new Scope(scope)
          node.params.forEach((p, i) => inner.value(p.name, args[i] ?? null))
          return this.eval(node.body, inner)
        })
      case 'let': {
        const inner = new Scope(scope)
        for (const b of node.bindings) inner.define(b.name, () => this.eval(b.value, inner))
        return this.eval(node.body, inner)
      }
      case 'if': {
        const cond = this.eval(node.cond, scope)
        if (typeof cond !== 'boolean') {
          throw new MError(
            'Expression.Error',
            `We cannot convert the value ${describe(cond)} to type Logical.`,
          )
        }
        return this.eval(cond ? node.then : node.else, scope)
      }
      case 'try': {
        try {
          const value = this.eval(node.body, scope)
          return node.otherwise ? value : record(['HasError', 'Value'], [false, value])
        } catch (error) {
          if (!(error instanceof MError)) throw error
          if (node.otherwise) return this.eval(node.otherwise, scope)
          return record(['HasError', 'Error'], [true, errorRecord(error)])
        }
      }
      case 'error': {
        const value = this.eval(node.value, scope)
        if (typeof value === 'string') throw new MError('Expression.Error', value)
        const rec = asRecord(value)
        const get = (n: string): MValue => {
          const i = fieldIndex(rec, n)
          return i < 0 ? null : rec.get(i)
        }
        throw new MError(
          String(get('Reason') ?? 'Expression.Error'),
          String(get('Message') ?? ''),
          get('Detail'),
        )
      }
      case 'not': {
        const v = this.eval(node.value, scope)
        if (v === null) return null
        if (typeof v !== 'boolean')
          throw new MError('Expression.Error', `We cannot apply operator not to ${describe(v)}.`)
        return !v
      }
      case 'neg': {
        const v = this.eval(node.value, scope)
        if (v === null) return null
        if (typeof v === 'number') return -v
        if (v instanceof MDuration) return new MDuration(-v.days)
        throw new MError('Expression.Error', `We cannot apply operator - to ${describe(v)}.`)
      }
      case 'bin':
        return this.binary(node.op, node.left, node.right, scope)
      case 'as': {
        const v = this.eval(node.value, scope)
        if (!conformsTo(v, node.type)) {
          throw new MError(
            'Expression.Error',
            `We cannot convert the value ${describe(v)} to type ${node.type.name}.`,
          )
        }
        return v
      }
      case 'is':
        return conformsTo(this.eval(node.value, scope), node.type)
      case 'type':
        return type(node.type.name, node.type.nullable)
    }
  }

  /** `[Field]` inside `each`: a field of `_` */
  private implicitRow(scope: Scope): MValue {
    const underscore = scope.lookup('_')
    if (!underscore)
      throw new MError('Expression.Error', 'A field access needs a record (no "_" in scope).')
    return underscore()
  }

  private binary(op: string, leftNode: Node, rightNode: Node, scope: Scope): MValue {
    if (op === 'and' || op === 'or') {
      const left = this.eval(leftNode, scope)
      if (left !== null && typeof left !== 'boolean') throw logicalError(op, left)
      // short circuit: false and _ = false, true or _ = true
      if (op === 'and' && left === false) return false
      if (op === 'or' && left === true) return true
      const right = this.eval(rightNode, scope)
      if (right !== null && typeof right !== 'boolean') throw logicalError(op, right)
      if (op === 'and')
        return right === false ? false : left === null || right === null ? null : true
      return right === true ? true : left === null || right === null ? null : false
    }
    const left = this.eval(leftNode, scope)
    const right = this.eval(rightNode, scope)
    switch (op) {
      case '=':
        return valuesEqual(left, right)
      case '<>':
        return !valuesEqual(left, right)
      case '<':
      case '>':
      case '<=':
      case '>=': {
        if (left === null || right === null) return null
        const c = compareValues(left, right, true)
        return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0
      }
      case '&':
        return concatenate(left, right)
      default:
        return arithmetic(op, left, right)
    }
  }
}

function logicalError(op: string, v: MValue): MError {
  return new MError('Expression.Error', `We cannot apply operator ${op} to ${describe(v)}.`)
}

export function errorRecord(error: MError): MValue {
  return record(['Reason', 'Message', 'Detail'], [error.reason, error.message, error.detail])
}

function fieldOf(target: MValue, name: string, optional: boolean): MValue {
  if (isRecord(target)) {
    const i = fieldIndex(target, name)
    if (i < 0) {
      if (optional) return null
      throw new MError('Expression.Error', `The field '${name}' of the record wasn't found.`)
    }
    return target.get(i)
  }
  if (isTable(target)) {
    // a column as a list
    const c = target.columns.indexOf(name)
    if (c < 0) {
      if (optional) return null
      throw new MError('Expression.Error', `The column '${name}' of the table wasn't found.`)
    }
    return list(target.rows.map((_, r) => rowRecord(target, r).get(c)))
  }
  if (target === null && optional) return null
  throw new MError(
    'Expression.Error',
    `We cannot apply field access to the type ${kindOf(target)}.`,
  )
}

function itemOf(target: MValue, key: MValue, optional: boolean): MValue {
  if (isList(target)) {
    if (typeof key !== 'number')
      throw new MError('Expression.Error', 'A list item index must be a number.')
    if (key < 0 || key >= target.items.length || !Number.isInteger(key)) {
      if (optional) return null
      throw new MError(
        'Expression.Error',
        "There weren't enough elements in the enumeration to complete the operation.",
      )
    }
    return target.items[key]!
  }
  if (isTable(target)) {
    if (typeof key === 'number') {
      if (key < 0 || key >= target.rows.length || !Number.isInteger(key)) {
        if (optional) return null
        throw new MError(
          'Expression.Error',
          "There weren't enough elements in the enumeration to complete the operation.",
        )
      }
      return rowRecord(target, key)
    }
    // {[Item="DATA", Kind="Sheet"]}: the one row whose fields equal the key's
    const k = asRecord(key)
    const cols = k.names.map((n) => {
      const c = target.columns.indexOf(n)
      if (c < 0)
        throw new MError('Expression.Error', `The column '${n}' of the table wasn't found.`)
      return c
    })
    const wanted = k.names.map((_, i) => k.get(i))
    const hits: number[] = []
    for (let r = 0; r < target.rows.length; r++) {
      const row = rowRecord(target, r)
      if (cols.every((c, i) => valuesEqual(row.get(c), wanted[i]!))) hits.push(r)
    }
    if (hits.length === 1) return rowRecord(target, hits[0]!)
    if (hits.length === 0 && optional) return null
    throw new MError(
      'Expression.Error',
      hits.length === 0
        ? "The key didn't match any rows in the table."
        : 'The key matched more than one row in the table.',
    )
  }
  throw new MError('Expression.Error', `We cannot apply item access to the type ${kindOf(target)}.`)
}

/** "a".."z" and 1..5 */
function rangeItems(from: MValue, to: MValue): MValue[] {
  if (typeof from === 'number' && typeof to === 'number') {
    const out: number[] = []
    for (let n = from; n <= to; n++) out.push(n)
    return out
  }
  if (typeof from === 'string' && typeof to === 'string' && from.length === 1 && to.length === 1) {
    const out: string[] = []
    for (let c = from.codePointAt(0)!; c <= to.codePointAt(0)!; c++)
      out.push(String.fromCodePoint(c))
    return out
  }
  throw new MError('Expression.Error', 'A list range needs two numbers or two single characters.')
}

function concatenate(left: MValue, right: MValue): MValue {
  if (left === null || right === null) return null
  if (typeof left === 'string' && typeof right === 'string') return left + right
  if (isList(left) && isList(right)) return list([...left.items, ...right.items])
  if (isRecord(left) && isRecord(right)) {
    const names = [...left.names]
    const getters: Array<() => MValue> = left.names.map((_, i) => () => left.get(i))
    right.names.forEach((n, i) => {
      const at = names.indexOf(n)
      const get = (): MValue => right.get(i)
      if (at >= 0) getters[at] = get
      else {
        names.push(n)
        getters.push(get)
      }
    })
    return lazyRecord(names, (i) => getters[i]!())
  }
  if (isTable(left) && isTable(right)) {
    const columns = [...left.columns]
    for (const c of right.columns) if (!columns.includes(c)) columns.push(c)
    const rows = [
      ...left.rows.map((r) =>
        columns.map((c) => (left.columns.includes(c) ? r[left.columns.indexOf(c)]! : null)),
      ),
      ...right.rows.map((r) =>
        columns.map((c) => (right.columns.includes(c) ? r[right.columns.indexOf(c)]! : null)),
      ),
    ]
    return { kind: 'table', columns, rows }
  }
  throw new MError(
    'Expression.Error',
    `We cannot apply operator & to types ${kindOf(left)} and ${kindOf(right)}.`,
  )
}

function arithmetic(op: string, left: MValue, right: MValue): MValue {
  if (left === null || right === null) return null
  if (typeof left === 'number' && typeof right === 'number') {
    switch (op) {
      case '+':
        return left + right
      case '-':
        return left - right
      case '*':
        return left * right
      case '/':
        return right === 0 ? (left === 0 ? NaN : left > 0 ? Infinity : -Infinity) : left / right
    }
  }
  // date/datetime ± duration, date - date
  if ((left instanceof MDate || left instanceof MDateTime) && right instanceof MDuration) {
    const serial =
      op === '+' ? left.serial + right.days : op === '-' ? left.serial - right.days : NaN
    if (!Number.isNaN(serial))
      return left instanceof MDate ? new MDate(Math.floor(serial)) : new MDateTime(serial)
  }
  if (
    op === '-' &&
    (left instanceof MDate || left instanceof MDateTime) &&
    (right instanceof MDate || right instanceof MDateTime)
  ) {
    return new MDuration(left.serial - right.serial)
  }
  throw new MError(
    'Expression.Error',
    `We cannot apply operator ${op} to types ${kindOf(left)} and ${kindOf(right)}.`,
  )
}

export { asList, asTable }
