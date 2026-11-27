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
 * the prefix the engine matches. Declarations are renamed along with it, so the
 * result is the same XML (same expanded names, same text), only spelled the way
 * PowerPoint writes it. Parts already on the conventional prefixes come back as
 * the identical string. PackageArchive.readText applies it, so parsing, byte
 * anchors and patch splices all see one text, while archive.entries keeps the
 * original bytes (an unedited save stays byte-identical).
 */

const T = 'http://schemas.openxmlformats.org/'
const MS = 'http://schemas.microsoft.com/office/'
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const MC_URI = `${T}markup-compatibility/2006`
const P_URI = `${T}presentationml/2006/main`

/** The prefixes the engine matches literally, keyed by namespace URI. */
const CANONICAL_PREFIX_BY_URI: ReadonlyMap<string, string> = new Map([
  [`${T}drawingml/2006/main`, 'a'],
  [P_URI, 'p'],
  [`${T}officeDocument/2006/relationships`, 'r'],
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

/** Prefixes the slide generators write into presentation parts without declaring them. */
const GENERATED_PREFIXES: ReadonlyArray<readonly [prefix: string, uri: string]> = [
  ['a', `${T}drawingml/2006/main`],
  ['r', `${T}officeDocument/2006/relationships`],
  ['p', P_URI],
]

const XMLNS_DECL_RE = /\bxmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(["'])(.*?)\2/g
const TAG_NAME_RE = /^<\s*([^\s/>]+)/
const ATTR_RE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

/** Cheap pre-check: true when some known namespace is bound to a prefix the engine does not match. */
export function needsPrefixCanonicalization(xml: string): boolean {
  XMLNS_DECL_RE.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = XMLNS_DECL_RE.exec(xml)) !== null) {
    const canonical = CANONICAL_PREFIX_BY_URI.get(m[3]!)
    if (canonical !== undefined && canonical !== (m[1] ?? '')) return true
  }
  return false
}

interface Attr {
  name: string
  value: string
}

interface Frame {
  /** end-tag name as written in the source, and as it goes out */
  source: string
  name: string
  scoped: boolean
}

function splitQName(name: string): [prefix: string, local: string] {
  const i = name.indexOf(':')
  return i < 0 ? ['', name] : [name.slice(0, i), name.slice(i + 1)]
}

const isDecl = (name: string): boolean => name === 'xmlns' || name.startsWith('xmlns:')

/** Respell known namespaces with their canonical prefixes (see the module comment). */
export function canonicalizePrefixes(xml: string): string {
  if (!needsPrefixCanonicalization(xml)) return xml

  // scope stacks; lookups walk from the top so inner declarations shadow outer ones
  const bindings: Array<Map<string, string>> = [new Map([['xml', XML_NS]])] // source prefix -> URI
  const renames: Array<Map<string, string>> = [new Map()] // source prefix -> output prefix
  const outBindings: Array<Map<string, string>> = [new Map([['xml', XML_NS]])] // output prefix -> URI
  const frames: Frame[] = []
  const lookup = (stack: Array<Map<string, string>>, key: string): string | undefined => {
    for (let i = stack.length - 1; i >= 0; i--) {
      const v = stack[i]!.get(key)
      if (v !== undefined) return v
    }
    return undefined
  }
  const renamePrefix = (prefix: string): string => lookup(renames, prefix) ?? prefix

  let rootSeen = false
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

    // the tag ends at the first '>' outside a quoted attribute value
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
        out += frame.name === frame.source ? tag : `</${frame.name}>`
        if (frame.scoped) {
          bindings.pop()
          renames.pop()
          outBindings.pop()
        }
      } else {
        out += tag
      }
      continue
    }

    const nameMatch = TAG_NAME_RE.exec(tag)
    if (!nameMatch) {
      out += tag
      continue
    }
    const name = nameMatch[1]!
    const selfClosing = /\/\s*>$/.test(tag)
    const attrs: Attr[] = []
    ATTR_RE.lastIndex = nameMatch[0].length
    let am: RegExpExecArray | null
    while ((am = ATTR_RE.exec(tag)) !== null) {
      attrs.push({ name: am[1]!, value: am[2] ?? am[3] ?? '' })
    }
    let changed = false

    // The document element declares the prefixes renamed anywhere below it, and on a
    // presentation part the ones the slide generators write undeclared, so edited
    // fragments spliced in later stay namespace-valid.
    if (!rootSeen) {
      rootSeen = true
      const declared = new Set(attrs.filter((a) => isDecl(a.name)).map((a) => a.name))
      const wanted = new Map<string, string>()
      XMLNS_DECL_RE.lastIndex = 0
      let dm: RegExpExecArray | null
      while ((dm = XMLNS_DECL_RE.exec(xml)) !== null) {
        const uri = dm[3]!
        const canonical = CANONICAL_PREFIX_BY_URI.get(uri)
        if (canonical !== undefined && canonical !== (dm[1] ?? '')) wanted.set(canonical, uri)
      }
      const rootPrefix = splitQName(name)[0]
      const rootDecl = rootPrefix === '' ? 'xmlns' : `xmlns:${rootPrefix}`
      if (attrs.find((a) => a.name === rootDecl)?.value === P_URI) {
        for (const [prefix, uri] of GENERATED_PREFIXES) wanted.set(prefix, uri)
      }
      for (const [prefix, uri] of wanted) {
        if (declared.has(`xmlns:${prefix}`)) continue
        attrs.push({ name: `xmlns:${prefix}`, value: uri })
        changed = true
      }
    }

    // namespace declarations open a scope
    const newBindings = new Map<string, string>()
    const declByPrefix = new Map<string, Attr>()
    for (const attr of attrs) {
      if (!isDecl(attr.name)) continue
      const declPrefix = attr.name === 'xmlns' ? '' : attr.name.slice(6)
      newBindings.set(declPrefix, attr.value)
      declByPrefix.set(declPrefix, attr)
    }
    let scoped = false
    if (newBindings.size > 0) {
      scoped = true
      bindings.push(newBindings)
      const newRenames = new Map<string, string>()
      renames.push(newRenames)
      const newOut = new Map<string, string>()
      outBindings.push(newOut)
      const drop = new Set<Attr>()
      // declarations keeping their prefix claim it in the output first, so a rename
      // can neither steal it nor be blocked by a prefix that is itself renamed away
      // (a kept declaration also shadows an outer rename of the same prefix)
      for (const [prefix, uri] of newBindings) {
        const canonical = CANONICAL_PREFIX_BY_URI.get(uri)
        if (canonical === undefined || canonical === prefix) {
          newOut.set(prefix, uri)
          newRenames.set(prefix, prefix)
        }
      }
      for (const [prefix, uri] of newBindings) {
        const canonical = CANONICAL_PREFIX_BY_URI.get(uri)
        if (canonical === undefined || canonical === prefix) continue
        const bound = lookup(outBindings, canonical)
        if (bound !== undefined && bound !== uri) {
          // the canonical prefix means another namespace here: keep this one as written
          newOut.set(prefix, uri)
          newRenames.set(prefix, prefix)
          continue
        }
        newRenames.set(prefix, canonical)
        const decl = declByPrefix.get(prefix)!
        if (prefix === '') {
          // a default namespace keeps covering unprefixed content; the canonical
          // prefix joins it unless already in scope
          newOut.set('', uri)
          if (bound !== uri) {
            attrs.push({ name: `xmlns:${canonical}`, value: uri })
            changed = true
          }
        } else if (bound === uri) {
          drop.add(decl) // the canonical binding is already in scope
          changed = true
        } else {
          decl.name = `xmlns:${canonical}`
          changed = true
        }
        newOut.set(canonical, uri)
      }
      if (drop.size > 0) attrs.splice(0, attrs.length, ...attrs.filter((a) => !drop.has(a)))
    }

    const [prefix, local] = splitQName(name)
    const targetPrefix = lookup(bindings, prefix) === undefined ? prefix : renamePrefix(prefix)
    const newName = targetPrefix === '' ? local : `${targetPrefix}:${local}`
    if (newName !== name) changed = true

    const elementUri = lookup(bindings, prefix)
    for (const attr of attrs) {
      if (isDecl(attr.name)) continue
      const [ap, al] = splitQName(attr.name)
      if (ap !== '' && ap !== 'xml') {
        const np = renamePrefix(ap)
        if (np !== ap) {
          attr.name = `${np}:${al}`
          changed = true
        }
      }
      // markup-compatibility prefix lists keep naming the renamed prefixes
      const mceList =
        ((al === 'Ignorable' || al === 'MustUnderstand' || al === 'ProcessContent') &&
          ap !== '' &&
          lookup(bindings, ap) === MC_URI) ||
        (al === 'Requires' && ap === '' && elementUri === MC_URI)
      if (mceList) {
        const renamed = attr.value
          .split(/(\s+)/)
          .map((part) => (/^\s*$/.test(part) ? part : renameListEntry(part, renamePrefix)))
          .join('')
        if (renamed !== attr.value) {
          attr.value = renamed
          changed = true
        }
      }
    }

    if (changed) {
      const attrText = attrs.map((a) => ` ${a.name}="${a.value.replace(/"/g, '&quot;')}"`).join('')
      out += `<${newName}${attrText}${selfClosing ? '/>' : '>'}`
    } else {
      out += tag
    }

    if (selfClosing) {
      if (scoped) {
        bindings.pop()
        renames.pop()
        outBindings.pop()
      }
    } else {
      frames.push({ source: name, name: newName, scoped })
    }
  }
  out += xml.slice(cursor)
  return out
}

/** One entry of an MCE list: a bare prefix ("a14") or a qualified name ("p14:*"). */
function renameListEntry(entry: string, renamePrefix: (prefix: string) => string): string {
  const [prefix, local] = splitQName(entry)
  if (!entry.includes(':')) return renamePrefix(entry)
  const np = renamePrefix(prefix)
  return np === prefix ? entry : `${np}:${local}`
}
