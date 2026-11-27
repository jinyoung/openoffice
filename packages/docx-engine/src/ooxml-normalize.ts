/**
 * ISO Strict OOXML (purl.oclc.org URIs) and non-canonical namespace prefixes,
 * normalized to the transitional URIs + canonical prefixes that the rest of
 * the engine matches literally (`w:p`, `wps:wsp`, rel-type strings...). Runs
 * once per XML part at zip load, so parse offsets, patch splices and saved
 * output all see the same bytes.
 */

const T = 'http://schemas.openxmlformats.org/'
const S = 'http://purl.oclc.org/ooxml/'
const MS = 'http://schemas.microsoft.com/office/'
const XML_NS = 'http://www.w3.org/XML/1998/namespace'

/** ISO 29500 Strict namespace -> ECMA transitional equivalent (official 1:1 pairs) */
const STRICT_TO_TRANSITIONAL: ReadonlyMap<string, string> = new Map([
  [`${S}wordprocessingml/main`, `${T}wordprocessingml/2006/main`],
  [`${S}officeDocument/relationships`, `${T}officeDocument/2006/relationships`],
  [`${S}officeDocument/math`, `${T}officeDocument/2006/math`],
  [`${S}officeDocument/extendedProperties`, `${T}officeDocument/2006/extended-properties`],
  [`${S}officeDocument/customProperties`, `${T}officeDocument/2006/custom-properties`],
  [`${S}officeDocument/docPropsVTypes`, `${T}officeDocument/2006/docPropsVTypes`],
  [`${S}officeDocument/bibliography`, `${T}officeDocument/2006/bibliography`],
  [`${S}officeDocument/characteristics`, `${T}officeDocument/2006/characteristics`],
  [`${S}officeDocument/customXml`, `${T}officeDocument/2006/customXml`],
  [`${S}schemaLibrary/main`, `${T}schemaLibrary/2006/main`],
  [`${S}drawingml/main`, `${T}drawingml/2006/main`],
  [`${S}drawingml/wordprocessingDrawing`, `${T}drawingml/2006/wordprocessingDrawing`],
  [`${S}drawingml/spreadsheetDrawing`, `${T}drawingml/2006/spreadsheetDrawing`],
  [`${S}drawingml/picture`, `${T}drawingml/2006/picture`],
  [`${S}drawingml/chart`, `${T}drawingml/2006/chart`],
  [`${S}drawingml/chartDrawing`, `${T}drawingml/2006/chartDrawing`],
  [`${S}drawingml/diagram`, `${T}drawingml/2006/diagram`],
  [`${S}drawingml/lockedCanvas`, `${T}drawingml/2006/lockedCanvas`],
  [`${S}drawingml/compatibility`, `${T}drawingml/2006/compatibility`],
])

const STRICT_REL_BASE = `${S}officeDocument/relationships/`
const TRANSITIONAL_REL_BASE = `${T}officeDocument/2006/relationships/`
/** relationship-type tails that differ between the strict and transitional families */
const REL_TAIL: Readonly<Record<string, string>> = {
  extendedProperties: 'extended-properties',
  customProperties: 'custom-properties',
}

const W_URI = `${T}wordprocessingml/2006/main`
const A_URI = `${T}drawingml/2006/main`
const WP_URI = `${T}drawingml/2006/wordprocessingDrawing`
const WPS_URI = `${MS}word/2010/wordprocessingShape`
const MC_URI = `${T}markup-compatibility/2006`

/**
 * Prefixes the engine matches literally, keyed by (transitional) namespace URI.
 * A binding of one of these URIs to any other prefix gets renamed — matching
 * by URI + localName, never by localName alone, so unrelated namespaces that
 * happen to reuse a local name are left untouched.
 */
const CANONICAL_PREFIX_BY_URI: ReadonlyMap<string, string> = new Map([
  [W_URI, 'w'],
  [`${T}officeDocument/2006/relationships`, 'r'],
  [`${T}officeDocument/2006/math`, 'm'],
  [A_URI, 'a'],
  [WP_URI, 'wp'],
  [`${T}drawingml/2006/picture`, 'pic'],
  [`${T}drawingml/2006/chart`, 'c'],
  [`${T}drawingml/2006/diagram`, 'dgm'],
  [`${T}drawingml/2006/lockedCanvas`, 'lc'],
  [MC_URI, 'mc'],
  [WPS_URI, 'wps'],
  [`${MS}word/2010/wordprocessingGroup`, 'wpg'],
  [`${MS}word/2010/wordprocessingCanvas`, 'wpc'],
  [`${MS}word/2010/wordml`, 'w14'],
  [`${MS}word/2012/wordml`, 'w15'],
  ['urn:schemas-microsoft-com:vml', 'v'],
  ['urn:schemas-microsoft-com:office:office', 'o'],
  ['urn:schemas-microsoft-com:office:word', 'w10'],
])

// Strict "universal measures" ("612pt", "2.54cm"...) -> the plain numeric unit
// each attribute is typed with (twips / half-points / eighth-points / points).
const PT_PER_UNIT: Readonly<Record<string, number>> = {
  pt: 1,
  in: 72,
  pc: 12,
  pi: 12,
  cm: 72 / 2.54,
  mm: 72 / 25.4,
}
const MEASURE_RE = /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/
const TWIPS = 20
const HALF_POINTS = 2
const EIGHTH_POINTS = 8
const POINTS = 1

type AttrFactors = Readonly<Record<string, number>>
const MARGIN_OR_BORDER: AttrFactors = { w: TWIPS, sz: EIGHTH_POINTS, space: POINTS }
/** element localName -> attr localName -> target units per point (w: namespace only) */
const MEASURE_ATTRS: ReadonlyMap<string, AttrFactors> = new Map<string, AttrFactors>([
  ['pgSz', { w: TWIPS, h: TWIPS }],
  [
    'pgMar',
    {
      top: TWIPS,
      right: TWIPS,
      bottom: TWIPS,
      left: TWIPS,
      header: TWIPS,
      footer: TWIPS,
      gutter: TWIPS,
    },
  ],
  [
    'ind',
    { left: TWIPS, right: TWIPS, start: TWIPS, end: TWIPS, firstLine: TWIPS, hanging: TWIPS },
  ],
  ['spacing', { before: TWIPS, after: TWIPS, line: TWIPS, val: TWIPS }],
  ['tab', { pos: TWIPS }],
  ['defaultTabStop', { val: TWIPS }],
  ['cols', { space: TWIPS }],
  ['col', { w: TWIPS, space: TWIPS }],
  ['tblInd', { w: TWIPS }],
  ['tblW', { w: TWIPS }],
  ['tcW', { w: TWIPS }],
  ['tblCellSpacing', { w: TWIPS }],
  ['object', { dxaOrig: TWIPS, dyaOrig: TWIPS }],
  ['trHeight', { val: TWIPS }],
  ['framePr', { w: TWIPS, h: TWIPS, x: TWIPS, y: TWIPS, hSpace: TWIPS, vSpace: TWIPS }],
  [
    'tblpPr',
    {
      leftFromText: TWIPS,
      rightFromText: TWIPS,
      topFromText: TWIPS,
      bottomFromText: TWIPS,
      tblpX: TWIPS,
      tblpY: TWIPS,
    },
  ],
  ['sz', { val: HALF_POINTS }],
  ['szCs', { val: HALF_POINTS }],
  ['kern', { val: HALF_POINTS }],
  ['position', { val: HALF_POINTS }],
  ['hps', { val: HALF_POINTS }],
  ['hpsRaise', { val: HALF_POINTS }],
  ['hpsBaseText', { val: HALF_POINTS }],
  // cell margins (w) and borders (sz/space) share these element names
  ['top', MARGIN_OR_BORDER],
  ['bottom', MARGIN_OR_BORDER],
  ['left', MARGIN_OR_BORDER],
  ['right', MARGIN_OR_BORDER],
  ['start', MARGIN_OR_BORDER],
  ['end', MARGIN_OR_BORDER],
  ['between', MARGIN_OR_BORDER],
  ['bar', MARGIN_OR_BORDER],
  ['insideH', MARGIN_OR_BORDER],
  ['insideV', MARGIN_OR_BORDER],
  ['tl2br', MARGIN_OR_BORDER],
  ['tr2bl', MARGIN_OR_BORDER],
])

function convertMeasure(value: string, factor: number): string {
  const m = MEASURE_RE.exec(value)
  if (!m) return value
  // A hostile file can declare astronomical measures ("9".repeat(400) + "in"):
  // the product overflows to Infinity, which would serialize verbatim into
  // the normalized attribute. Keep the original bytes instead.
  const converted = parseFloat(m[1]) * PT_PER_UNIT[m[2]]! * factor
  if (!Number.isFinite(converted)) return value
  return String(Math.round(converted))
}

function mapUriValue(value: string): string {
  const direct = STRICT_TO_TRANSITIONAL.get(value)
  if (direct) return direct
  if (value.startsWith(STRICT_REL_BASE)) {
    const tail = value.slice(STRICT_REL_BASE.length)
    return TRANSITIONAL_REL_BASE + (REL_TAIL[tail] ?? tail)
  }
  return value
}

const XMLNS_DECL_RE = /xmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(["'])([^"']*)\2/g

/**
 * True when a start tag declares a strict namespace, names a strict relationship
 * type, or binds a known namespace to an unconventional prefix. The same text in
 * element content, comments, CDATA or another attribute value does not count.
 */
export function needsOoxmlNormalization(xml: string): boolean {
  // cheap text pre-filter: most parts carry neither marker anywhere
  let hint = xml.includes(S)
  if (!hint) {
    XMLNS_DECL_RE.lastIndex = 0
    let m: RegExpExecArray | null
    while ((m = XMLNS_DECL_RE.exec(xml)) !== null) {
      const canonical = CANONICAL_PREFIX_BY_URI.get(m[3]!)
      if (canonical !== undefined && canonical !== (m[1] ?? '')) {
        hint = true
        break
      }
    }
  }
  if (!hint) return false
  for (const t of startTags(xml)) {
    if (!t.includes('xmlns') && !t.includes('Type')) continue
    ATTR_RE.lastIndex = TAG_NAME_RE.exec(t)?.[0].length ?? t.length
    let am: RegExpExecArray | null
    while ((am = ATTR_RE.exec(t)) !== null) {
      const name = am[1]!
      const value = am[2] ?? am[3] ?? ''
      if (name === 'Type' && value.startsWith(STRICT_REL_BASE)) return true
      if (name !== 'xmlns' && !name.startsWith('xmlns:')) continue
      if (value.startsWith(S)) return true
      const canonical = CANONICAL_PREFIX_BY_URI.get(value)
      if (canonical !== undefined && canonical !== (name === 'xmlns' ? '' : name.slice(6))) {
        return true
      }
    }
  }
  return false
}

/**
 * Start tags in document order: comments, CDATA, processing instructions and the
 * doctype are skipped whole, and a '>' inside a quoted attribute value does not
 * end a tag.
 */
function* startTags(xml: string): Generator<string> {
  let cursor = 0
  while (cursor < xml.length) {
    const start = xml.indexOf('<', cursor)
    if (start < 0) return
    const specialEnd = xml.startsWith('<!--', start)
      ? '-->'
      : xml.startsWith('<![CDATA[', start)
        ? ']]>'
        : xml.startsWith('<?', start)
          ? '?>'
          : xml.startsWith('<!', start)
            ? '>'
            : null
    if (specialEnd !== null) {
      const at = xml.indexOf(specialEnd, start + 2)
      if (at < 0) return
      cursor = at + specialEnd.length
      continue
    }
    let quote = ''
    let end = start + 1
    for (; end < xml.length; end += 1) {
      const ch = xml[end]!
      if (quote) {
        if (ch === quote) quote = ''
      } else if (ch === '"' || ch === "'") {
        quote = ch
      } else if (ch === '>') {
        break
      }
    }
    if (end >= xml.length) return
    if (xml[start + 1] !== '/') yield xml.slice(start, end + 1)
    cursor = end + 1
  }
}

interface Attr {
  name: string
  value: string
}

interface Frame {
  /** rewritten name to emit for the matching end tag */
  name: string
  wspScope: boolean
}

interface Binding {
  uri: string
  /** the prefix this binding is written with in the output */
  out: string
}

/**
 * In-scope namespace bindings as one stack per prefix, so a lookup is O(1)
 * however deep the document nests (a frame-by-frame walk went quadratic on a
 * declaration at every level). Each open element logs what it pushed.
 */
class NsScope {
  private readonly src = new Map<string, Binding[]>()
  private readonly out = new Map<string, string[]>()
  private readonly log: Array<Array<() => void>> = [[]]

  constructor() {
    this.bind('xml', XML_NS, 'xml')
    this.bindOut('xml', XML_NS)
  }

  open(): void {
    this.log.push([])
  }

  close(): void {
    const undo = this.log.pop() ?? []
    for (let i = undo.length - 1; i >= 0; i--) undo[i]!()
  }

  /** source prefix -> namespace + output spelling */
  bind(prefix: string, uri: string, out: string): void {
    let stack = this.src.get(prefix)
    if (!stack) this.src.set(prefix, (stack = []))
    stack.push({ uri, out })
    this.log[this.log.length - 1]!.push(() => stack.pop())
  }

  /** what an output prefix means from here down */
  bindOut(prefix: string, uri: string): void {
    let stack = this.out.get(prefix)
    if (!stack) this.out.set(prefix, (stack = []))
    stack.push(uri)
    this.log[this.log.length - 1]!.push(() => stack.pop())
  }

  top(prefix: string): Binding | undefined {
    const stack = this.src.get(prefix)
    return stack?.[stack.length - 1]
  }

  outUri(prefix: string): string | undefined {
    const stack = this.out.get(prefix)
    return stack?.[stack.length - 1]
  }
}

const TAG_NAME_RE = /^<\/?\s*([^\s/>]+)/
const ATTR_RE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

function splitQName(name: string): [prefix: string, local: string] {
  const i = name.indexOf(':')
  return i < 0 ? ['', name] : [name.slice(0, i), name.slice(i + 1)]
}

export function normalizeOoxmlXml(xml: string): string {
  const scope = new NsScope()
  /** last number handed out per base of a fresh prefix ('ns' for the default namespace) */
  const fresh = new Map<string, number>()
  /** an output prefix unbound here; one counter per base only climbs, so no
   *  search restarts from the first name */
  const freshPrefix = (source: string): string => {
    const base = source || 'ns'
    let k = fresh.get(base) ?? 0
    let next: string
    do next = `${base}-${++k}`
    while (scope.outUri(next) !== undefined)
    fresh.set(base, k)
    return next
  }
  const frames: Frame[] = []
  let wspDepth = 0

  let out = ''
  let cursor = 0
  while (cursor < xml.length) {
    const start = xml.indexOf('<', cursor)
    if (start < 0) break
    out += xml.slice(cursor, start)

    const specialEnd = xml.startsWith('<!--', start)
      ? '-->'
      : xml.startsWith('<![CDATA[', start)
        ? ']]>'
        : xml.startsWith('<?', start)
          ? '?>'
          : xml.startsWith('<!', start)
            ? '>'
            : null
    if (specialEnd !== null) {
      const at = xml.indexOf(specialEnd, start + 2)
      const end = at < 0 ? xml.length : at + specialEnd.length
      out += xml.slice(start, end)
      cursor = end
      continue
    }

    // find the tag end, tolerating '>' inside quoted attribute values
    let quote = ''
    let end = start + 1
    for (; end < xml.length; end += 1) {
      const ch = xml[end]!
      if (quote) {
        if (ch === quote) quote = ''
      } else if (ch === '"' || ch === "'") {
        quote = ch
      } else if (ch === '>') {
        break
      }
    }
    if (end >= xml.length) {
      out += xml.slice(start)
      cursor = xml.length
      continue
    }
    const tag = xml.slice(start, end + 1)
    cursor = end + 1

    if (tag.startsWith('</')) {
      const frame = frames.pop()
      if (frame) {
        out += `</${frame.name}>`
        scope.close()
        if (frame.wspScope) wspDepth -= 1
      } else {
        out += tag
      }
      continue
    }

    const selfClosing = /\/\s*>$/.test(tag)
    const name = TAG_NAME_RE.exec(tag)?.[1]
    if (!name) {
      out += tag
      continue
    }
    const attrs: Attr[] = []
    ATTR_RE.lastIndex = TAG_NAME_RE.exec(tag)![0].length
    let am: RegExpExecArray | null
    while ((am = ATTR_RE.exec(tag)) !== null) {
      attrs.push({ name: am[1]!, value: am[2] ?? am[3] ?? '' })
    }

    // namespace declarations; strict URIs map to transitional
    scope.open()
    const declByPrefix = new Map<string, Attr>()
    for (const attr of attrs) {
      if (attr.name !== 'xmlns' && !attr.name.startsWith('xmlns:')) continue
      attr.value = mapUriValue(attr.value)
      declByPrefix.set(attr.name === 'xmlns' ? '' : attr.name.slice(6), attr)
    }
    if (declByPrefix.size > 0) {
      const drop = new Set<Attr>()
      const inject: Attr[] = []
      /** output prefixes a rename or injection declared on this element */
      const renamedHere = new Set<string>()
      // declarations that keep their prefix claim it in the output first, so a
      // rename can neither steal it nor be blocked by a prefix that is itself
      // being renamed away
      for (const [prefix, decl] of declByPrefix) {
        const canonical = CANONICAL_PREFIX_BY_URI.get(decl.value)
        if (canonical !== undefined && canonical !== prefix) continue
        scope.bind(prefix, decl.value, prefix)
        scope.bindOut(prefix, decl.value)
      }
      for (const [prefix, decl] of declByPrefix) {
        const uri = decl.value
        const canonical = CANONICAL_PREFIX_BY_URI.get(uri)
        if (canonical === undefined || canonical === prefix) continue
        const bound = scope.outUri(canonical)
        if (bound !== undefined && bound !== uri) {
          // canonical prefix taken by another namespace: keep as-is, unless a
          // rename on this element already took this name
          const taken = prefix !== '' && renamedHere.has(prefix)
          const out = taken ? freshPrefix(prefix) : prefix
          if (taken) decl.name = `xmlns:${out}`
          scope.bind(prefix, uri, out)
          scope.bindOut(out, uri)
          continue
        }
        scope.bind(prefix, uri, canonical)
        // by captured reference: finding by name could hit a declaration whose
        // name an earlier rename already rewrote to this prefix
        if (bound === uri) {
          if (prefix !== '') drop.add(decl) // canonical binding already exists
        } else {
          if (prefix === '') {
            inject.push({ name: `xmlns:${canonical}`, value: uri }) // keep xmlns= for other content
          } else {
            decl.name = `xmlns:${canonical}`
          }
          scope.bindOut(canonical, uri)
          renamedHere.add(canonical)
        }
        if (prefix === '') scope.bindOut('', uri)
      }
      if (drop.size > 0 || inject.length > 0) {
        attrs.splice(0, attrs.length, ...attrs.filter((a) => !drop.has(a)), ...inject)
      }
    }

    // A reference whose output prefix a declaration between its binding and here
    // rebinds to another namespace (<x xmlns:w="urn:other"> under a renamed
    // ns0:) gets a declaration on this element: its own prefix when that is free
    // here, a fresh one otherwise.
    const resolve = (source: string): string => {
      const b = scope.top(source)
      if (!b || b.uri === '') return source
      if (scope.outUri(b.out) === b.uri) return b.out
      if (source === '' && scope.outUri('') === b.uri) {
        scope.bind('', b.uri, '') // the default namespace still covers it
        return ''
      }
      const next =
        source === '' || (scope.outUri(source) ?? b.uri) !== b.uri ? freshPrefix(source) : source
      if (scope.outUri(next) !== b.uri) {
        attrs.push({ name: `xmlns:${next}`, value: b.uri })
        scope.bindOut(next, b.uri)
      }
      scope.bind(source, b.uri, next)
      return next
    }

    const [prefix, local] = splitQName(name)
    const elementUri = scope.top(prefix)?.uri

    // Word 2013 strict writes wordprocessingShape children in the wp (drawing)
    // namespace inside <a:graphicData uri="...wordprocessingShape">; the engine
    // matches wps:*, so shift those elements (only there) into the wps prefix.
    let targetPrefix: string
    if (wspDepth > 0 && elementUri === WP_URI) {
      targetPrefix = 'wps'
      if (scope.outUri('wps') !== WPS_URI) {
        attrs.push({ name: 'xmlns:wps', value: WPS_URI })
        scope.bindOut('wps', WPS_URI)
      }
    } else {
      targetPrefix = resolve(prefix)
    }
    const newName = targetPrefix === '' ? local : `${targetPrefix}:${local}`

    const inWNs = elementUri === W_URI
    const factors = inWNs ? MEASURE_ATTRS.get(local) : undefined
    const kept: Attr[] = []
    for (const attr of attrs) {
      if (attr.name === 'xmlns' || attr.name.startsWith('xmlns:')) {
        kept.push(attr)
        continue
      }
      const [ap, al] = splitQName(attr.name)
      if (inWNs && local === 'document' && al === 'conformance') continue
      const attrUri = ap === '' ? undefined : scope.top(ap)?.uri
      if (ap !== '' && ap !== 'xml') {
        const np = resolve(ap)
        if (np !== ap) attr.name = `${np}:${al}`
      }
      // MCE prefix lists must keep referring to the renamed prefixes
      const mceList =
        (al === 'Ignorable' || al === 'MustUnderstand') && attrUri === MC_URI
          ? true
          : al === 'Requires' && ap === '' && elementUri === MC_URI
      // ...and so must the qualified-name lists ("w14:*", "v:shape")
      const mceQNames =
        (al === 'ProcessContent' || al === 'PreserveElements' || al === 'PreserveAttributes') &&
        attrUri === MC_URI
      if (mceList) {
        attr.value = attr.value
          .split(/\s+/)
          .map((p) => (p ? resolve(p) : p))
          .join(' ')
      } else if (mceQNames) {
        attr.value = attr.value
          .split(/\s+/)
          .map((q) => {
            const i = q.indexOf(':')
            return i < 0 ? q : `${resolve(q.slice(0, i))}${q.slice(i)}`
          })
          .join(' ')
      } else {
        attr.value = mapUriValue(attr.value)
        const factor = factors?.[al]
        if (factor !== undefined && (ap === '' || attrUri === W_URI)) {
          attr.value = convertMeasure(attr.value, factor)
        }
      }
      kept.push(attr)
    }

    const attrText = kept.map((a) => ` ${a.name}="${a.value.replace(/"/g, '&quot;')}"`).join('')
    out += `<${newName}${attrText}${selfClosing ? '/>' : '>'}`

    const wspScope =
      !selfClosing && local === 'graphicData' && elementUri === A_URI
        ? attrs.some((a) => a.name === 'uri' && a.value === WPS_URI)
        : false
    if (wspScope) wspDepth += 1

    if (selfClosing) scope.close()
    else frames.push({ name: newName, wspScope })
  }
  out += xml.slice(cursor)
  return out
}
