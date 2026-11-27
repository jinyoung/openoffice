/**
 * Namespace-prefix canonicalization for pptx XML parts.
 *
 * The engine reads DrawingML / PresentationML by qualified name ('a:p', 'p:sp',
 * '@_r:embed', `<a:t>` regexes...), which only holds while every namespace is
 * bound to its conventional prefix. XML lets a writer pick any prefix: lxml
 * serializes DrawingML it did not get a prefix for as ns0:, ns1:..., each
 * declared on the element itself, and such a deck opened with every text box
 * empty (only pictures and tables drew).
 *
 * canonicalizePrefixes resolves each element and attribute prefix to its
 * namespace URI through the in-scope xmlns declarations, and respells it with
 * the prefix the engine matches. A renamed declaration is renamed where it
 * stands (never dropped or moved), so every element keeps the declarations it
 * had and a copied fragment stays self-contained. Where a descendant rebinds an
 * output prefix to another namespace, the references below it that still mean
 * the outer namespace get a declaration of their own. The result is the same
 * XML (same expanded names, same text). Parts already on the conventional
 * prefixes come back as the identical string.
 *
 * PackageArchive.readText applies it, so parsing, byte anchors and patch splices
 * all see one text, while archive.entries keeps the original bytes (an unedited
 * save stays byte-identical).
 */

const T = 'http://schemas.openxmlformats.org/'
const MS = 'http://schemas.microsoft.com/office/'
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const MC_URI = `${T}markup-compatibility/2006`
const A_URI = `${T}drawingml/2006/main`
const P_URI = `${T}presentationml/2006/main`
const R_URI = `${T}officeDocument/2006/relationships`

/** The prefixes the engine matches literally, keyed by namespace URI. */
const CANONICAL_PREFIX_BY_URI: ReadonlyMap<string, string> = new Map([
  [A_URI, 'a'],
  [P_URI, 'p'],
  [R_URI, 'r'],
  [MC_URI, 'mc'],
  [`${T}drawingml/2006/chart`, 'c'],
  [`${T}drawingml/2006/chartDrawing`, 'cdr'],
  [`${T}drawingml/2006/diagram`, 'dgm'],
  [`${MS}drawing/2008/diagram`, 'dsp'],
  [`${MS}drawing/2014/chartex`, 'cx'],
  [`${MS}drawing/2010/main`, 'a14'],
  [`${MS}drawing/2014/main`, 'a16'],
  [`${MS}drawing/2012/chart`, 'c15'],
  [`${MS}drawing/2016/SVG/main`, 'asvg'],
  [`${MS}powerpoint/2010/main`, 'p14'],
  ['urn:schemas-microsoft-com:vml', 'v'],
  ['urn:schemas-microsoft-com:office:office', 'o'],
])

/** Prefixes the slide generators write into a part without declaring them. */
const GENERATED_PREFIXES: ReadonlyArray<readonly [prefix: string, uri: string]> = [
  ['a', A_URI],
  ['r', R_URI],
  ['p', P_URI],
]

/** markup-compatibility attributes listing prefixes ("a14 p14") */
const MCE_PREFIX_LISTS = new Set(['Ignorable', 'MustUnderstand'])
/** markup-compatibility attributes listing qualified names ("p14:*", "v:shape") */
const MCE_QNAME_LISTS = new Set(['ProcessContent', 'PreserveElements', 'PreserveAttributes'])

const XMLNS_HINT_RE = /\bxmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(["'])(.*?)\2/g
const TAG_NAME_RE = /^<\s*([^\s/>]+)/
const ATTR_RE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
const SELF_CLOSING_RE = /\/\s*>$/

interface Attr {
  name: string
  value: string
}

interface TagToken {
  start: number
  /** exclusive */
  end: number
  close: boolean
}

const isDecl = (name: string): boolean => name === 'xmlns' || name.startsWith('xmlns:')
const declPrefix = (name: string): string => (name === 'xmlns' ? '' : name.slice(6))

function splitQName(name: string): [prefix: string, local: string] {
  const i = name.indexOf(':')
  return i < 0 ? ['', name] : [name.slice(0, i), name.slice(i + 1)]
}

/**
 * Start and end tags in document order. Comments, CDATA, processing
 * instructions and the doctype are skipped whole, and a '>' inside a quoted
 * attribute value does not end a tag. An unterminated tag ends the walk.
 */
function* tagTokens(xml: string): Generator<TagToken> {
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
    yield { start, end: end + 1, close: xml[start + 1] === '/' }
    cursor = end + 1
  }
}

function attrsOf(tag: string, from: number): Attr[] {
  const attrs: Attr[] = []
  ATTR_RE.lastIndex = from
  let m: RegExpExecArray | null
  while ((m = ATTR_RE.exec(tag)) !== null) attrs.push({ name: m[1]!, value: m[2] ?? m[3] ?? '' })
  return attrs
}

/**
 * True when some start tag binds a known namespace to a prefix the engine does
 * not match. Only real declarations count: the same text in element content,
 * comments, CDATA or an attribute value does not.
 */
export function needsPrefixCanonicalization(xml: string): boolean {
  // cheap text pre-filter first: most parts have no such declaration anywhere
  XMLNS_HINT_RE.lastIndex = 0
  let hint = false
  let m: RegExpExecArray | null
  while ((m = XMLNS_HINT_RE.exec(xml)) !== null) {
    const canonical = CANONICAL_PREFIX_BY_URI.get(m[3]!)
    if (canonical !== undefined && canonical !== (m[1] ?? '')) {
      hint = true
      break
    }
  }
  if (!hint) return false
  for (const t of tagTokens(xml)) {
    if (t.close) continue
    const tag = xml.slice(t.start, t.end)
    if (!tag.includes('xmlns')) continue
    const name = TAG_NAME_RE.exec(tag)
    if (!name) continue
    for (const a of attrsOf(tag, name[0].length)) {
      if (!isDecl(a.name)) continue
      const canonical = CANONICAL_PREFIX_BY_URI.get(a.value)
      if (canonical !== undefined && canonical !== declPrefix(a.name)) return true
    }
  }
  return false
}

interface Binding {
  uri: string
  /** the prefix this binding is written with in the output */
  out: string
}

/**
 * In-scope namespace bindings as one stack per prefix, so a lookup is O(1)
 * however deep the document nests. Each open element logs what it pushed.
 */
class Scope {
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

  /** source prefix → namespace + output spelling */
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

/** Respell known namespaces with their canonical prefixes (see the module comment). */
export function canonicalizePrefixes(xml: string): string {
  if (!needsPrefixCanonicalization(xml)) return xml

  const scope = new Scope()
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
  const frames: Array<{ name: string; outName: string }> = []
  let out = ''
  let cursor = 0
  for (const t of tagTokens(xml)) {
    out += xml.slice(cursor, t.start)
    cursor = t.end
    const tag = xml.slice(t.start, t.end)

    if (t.close) {
      const frame = frames.pop()
      if (!frame) {
        out += tag
        continue
      }
      scope.close()
      out += frame.outName === frame.name ? tag : `</${frame.outName}>`
      continue
    }

    const nameMatch = TAG_NAME_RE.exec(tag)
    if (!nameMatch) {
      out += tag
      continue
    }
    const name = nameMatch[1]!
    const attrs = attrsOf(tag, nameMatch[0].length)
    const selfClosing = SELF_CLOSING_RE.test(tag)
    let changed = false
    scope.open()

    // declarations: those keeping their prefix bind first, so a rename can neither
    // take a prefix this element binds to something else nor be blocked by one
    // that is itself renamed away
    const declared = new Map<string, string>() // output prefix -> URI, on this element
    const drop = new Set<Attr>()
    for (const attr of attrs) {
      if (!isDecl(attr.name)) continue
      const prefix = declPrefix(attr.name)
      const canonical = CANONICAL_PREFIX_BY_URI.get(attr.value)
      if (canonical !== undefined && canonical !== prefix) continue
      scope.bind(prefix, attr.value, prefix)
      scope.bindOut(prefix, attr.value)
      declared.set(prefix, attr.value)
    }
    for (const attr of attrs) {
      if (!isDecl(attr.name)) continue
      const prefix = declPrefix(attr.name)
      const uri = attr.value
      const canonical = CANONICAL_PREFIX_BY_URI.get(uri)
      if (canonical === undefined || canonical === prefix) continue
      const here = declared.get(canonical)
      const current = here ?? scope.outUri(canonical)
      if (current !== undefined && current !== uri) {
        // the canonical prefix means another namespace here: keep this one as
        // written, unless a rename on this element already took its name
        const taken = prefix !== '' && (declared.get(prefix) ?? uri) !== uri
        const out = taken ? freshPrefix(prefix) : prefix
        if (taken) {
          attr.name = `xmlns:${out}`
          changed = true
        }
        scope.bind(prefix, uri, out)
        scope.bindOut(out, uri)
        declared.set(out, uri)
        continue
      }
      scope.bind(prefix, uri, canonical)
      if (prefix === '') {
        // the default namespace keeps covering unprefixed content; the canonical
        // prefix is declared beside it
        scope.bindOut('', uri)
        declared.set('', uri)
        if (here !== uri) {
          attrs.push({ name: `xmlns:${canonical}`, value: uri })
          scope.bindOut(canonical, uri)
          declared.set(canonical, uri)
          changed = true
        }
      } else if (here === uri) {
        drop.add(attr) // this element already declares the canonical prefix for it
        changed = true
      } else {
        attr.name = `xmlns:${canonical}`
        scope.bindOut(canonical, uri)
        declared.set(canonical, uri)
        changed = true
      }
    }

    // A reference whose output prefix a declaration between its binding and here
    // rebinds to another namespace gets a declaration on this element: its own
    // prefix when that is free here, a fresh one otherwise.
    const resolve = (prefix: string): string => {
      const b = scope.top(prefix)
      if (!b || b.uri === '') return prefix
      if (scope.outUri(b.out) === b.uri) return b.out
      if (prefix === '' && scope.outUri('') === b.uri) {
        scope.bind('', b.uri, '') // the default namespace still covers it
        changed = true
        return ''
      }
      const next =
        prefix === '' || (scope.outUri(prefix) ?? b.uri) !== b.uri ? freshPrefix(prefix) : prefix
      if (scope.outUri(next) !== b.uri) {
        attrs.push({ name: `xmlns:${next}`, value: b.uri })
        scope.bindOut(next, b.uri)
        declared.set(next, b.uri)
      }
      scope.bind(prefix, b.uri, next)
      changed = true
      return next
    }
    const respell = (qname: string): string => {
      if (!qname.includes(':')) return qname
      const [p, local] = splitQName(qname)
      const np = resolve(p)
      return np === p ? qname : `${np}:${local}`
    }

    const [prefix, local] = splitQName(name)
    const elementUri = scope.top(prefix)?.uri
    const outPrefix = resolve(prefix)
    const newName = outPrefix === '' ? local : `${outPrefix}:${local}`
    if (newName !== name) changed = true

    const own = attrs.length
    for (let i = 0; i < own; i++) {
      const attr = attrs[i]!
      if (isDecl(attr.name) || drop.has(attr)) continue
      const [ap, al] = splitQName(attr.name)
      const attrUri = ap === '' ? undefined : scope.top(ap)?.uri
      if (ap !== '') {
        const renamed = respell(attr.name)
        if (renamed !== attr.name) {
          attr.name = renamed
          changed = true
        }
      }
      // markup-compatibility lists keep naming the prefixes as written in the output
      let value = attr.value
      if (attrUri === MC_URI && MCE_PREFIX_LISTS.has(al)) {
        value = mapList(attr.value, resolve)
      } else if (attrUri === MC_URI && MCE_QNAME_LISTS.has(al)) {
        value = mapList(attr.value, respell)
      } else if (ap === '' && al === 'Requires' && elementUri === MC_URI) {
        value = mapList(attr.value, resolve)
      }
      if (value !== attr.value) {
        attr.value = value
        changed = true
      }
    }

    if (changed) {
      const attrText = attrs
        .filter((a) => !drop.has(a))
        .map((a) => ` ${a.name}="${a.value.replace(/"/g, '&quot;')}"`)
        .join('')
      out += `<${newName}${attrText}${selfClosing ? '/>' : '>'}`
    } else {
      out += tag
    }

    if (selfClosing) scope.close()
    else frames.push({ name, outName: newName })
  }
  out += xml.slice(cursor)
  return out
}

function mapList(value: string, map: (entry: string) => string): string {
  return value
    .split(/(\s+)/)
    .map((part) => (part === '' || /^\s+$/.test(part) ? part : map(part)))
    .join('')
}

/**
 * The part with its document element declaring a:, r: and p: where it lacks
 * them. The slide generators write those prefixes undeclared, so a part that
 * declares DrawingML only locally (lxml output) needs them on the root once an
 * edit rewrites it. A root binding one of them to another namespace keeps it.
 */
export function declareGeneratedPrefixes(xml: string): string {
  for (const t of tagTokens(xml)) {
    if (t.close) return xml
    const tag = xml.slice(t.start, t.end)
    const name = TAG_NAME_RE.exec(tag)
    if (!name) return xml
    const declared = new Set(attrsOf(tag, name[0].length).map((a) => a.name))
    const missing = GENERATED_PREFIXES.filter(([prefix]) => !declared.has(`xmlns:${prefix}`))
    if (missing.length === 0) return xml
    const insertAt = t.start + tag.search(/\s*\/?\s*>$/)
    return (
      xml.slice(0, insertAt) +
      missing.map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`).join('') +
      xml.slice(insertAt)
    )
  }
  return xml
}
