/**
 * Excel's in-cell checkboxes (PS Office): Insert › Checkbox in Excel stores
 * no control, only a cell format. The format (an xf of styles.xml) carries
 * an extension `<xfpb:xfComplement i="n"/>` that points, through the
 * workbook's feature property bags (xl/featurePropertyBag/*.xml), at a bag
 * of type "Checkbox". The cell itself holds TRUE or FALSE.
 *
 * This reads which xf indices are such checkbox formats, so the screen can
 * draw those cells as checkboxes. Nothing is written: the file keeps its
 * formats, and toggling a box is an ordinary TRUE/FALSE edit of the cell.
 */

interface Bag {
  readonly type: string
  /** `<bagId k="…">n</bagId>` children */
  readonly ids: ReadonlyMap<string, number>
  /** `<a k="…"><bagId>n</bagId>…</a>` children */
  readonly lists: ReadonlyMap<string, readonly number[]>
}

function attr(attrs: string, name: string): string | null {
  return new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attrs)?.[1] ?? null
}

/** The bags of a feature property bag part, in order (a bag's id is its position). */
export function parseFeaturePropertyBags(xml: string): Bag[] {
  const bags: Bag[] = []
  for (const m of xml.matchAll(/<(?:\w+:)?bag\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?bag>)/g)) {
    const body = m[2] ?? ''
    const ids = new Map<string, number>()
    const lists = new Map<string, number[]>()
    // lists first, then the bagIds outside them
    const outside = body.replace(
      /<(?:\w+:)?a\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?a>/g,
      (_whole, a: string, inner: string) => {
        const key = attr(a, 'k')
        if (key !== null) {
          lists.set(
            key,
            [...inner.matchAll(/<(?:\w+:)?bagId\b[^>]*>\s*(\d+)\s*</g)].map((x) => Number(x[1])),
          )
        }
        return ''
      },
    )
    for (const id of outside.matchAll(/<(?:\w+:)?bagId\b([^>]*)>\s*(\d+)\s*</g)) {
      const key = attr(id[1]!, 'k')
      if (key !== null) ids.set(key, Number(id[2]))
    }
    bags.push({ type: attr(m[1]!, 'type') ?? '', ids, lists })
  }
  return bags
}

/** whether an xfComplement bag leads to a checkbox: XFComplement → XFControls → CellControl → Checkbox */
function isCheckboxComplement(bags: readonly Bag[], index: number | undefined): boolean {
  const complement = index === undefined ? undefined : bags[index]
  if (complement?.type !== 'XFComplement') return false
  const controls = bags[complement.ids.get('XFControls') ?? -1]
  if (controls?.type !== 'XFControls') return false
  return bags[controls.ids.get('CellControl') ?? -1]?.type === 'Checkbox'
}

/** The indices (cellXfs order) of the cell formats that are Excel checkboxes. */
export function checkboxStyleIndexes(stylesXml: string, bagsXml: string): number[] {
  const bags = parseFeaturePropertyBags(bagsXml)
  const mapper = bags.find((bag) => bag.type === 'XFComplements')
  const mapped = mapper?.lists.get('MappedFeaturePropertyBags') ?? []
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? ''
  const out: number[] = []
  let index = 0
  for (const m of xfs.matchAll(/<xf\b[^>]*?(?:\/>|>([\s\S]*?)<\/xf>)/g)) {
    const complement = /<(?:\w+:)?xfComplement\b([^>]*?)\/?>/.exec(m[1] ?? '')
    const i = complement ? attr(complement[1]!, 'i') : null
    if (i !== null && isCheckboxComplement(bags, mapped[Number(i)])) out.push(index)
    index++
  }
  return out
}

/** The feature property bag part of a workbook, from its relationships; null when it has none. */
export function featurePropertyBagPart(workbookRelsXml: string): string | null {
  for (const m of workbookRelsXml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    if (!/\/FeaturePropertyBag$/i.test(attr(m[1]!, 'Type') ?? '')) continue
    const target = attr(m[1]!, 'Target')
    if (!target) continue
    if (target.startsWith('/')) return target.slice(1)
    const parts = `xl/${target}`.split('/')
    const out: string[] = []
    for (const p of parts) {
      if (p === '..') out.pop()
      else if (p !== '.' && p !== '') out.push(p)
    }
    return out.join('/')
  }
  return null
}
