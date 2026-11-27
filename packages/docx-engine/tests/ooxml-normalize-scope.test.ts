/**
 * @vitest-environment jsdom
 */
/// <reference lib="dom" />
// DOMParser reads the normalized XML with namespaces: an unbound prefix is an
// error and every element and attribute must keep its expanded name.

/**
 * Prefix renames under nested declarations. A descendant that binds the canonical
 * prefix to another namespace (<x xmlns:w="urn:other"> under a renamed ns0:)
 * used to capture the renamed references below it, moving them into urn:other.
 */
import { describe, expect, it } from 'vitest'
import { normalizeOoxmlXml } from '../src/ooxml-normalize'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'

function parseStrictly(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  const error = doc.getElementsByTagName('parsererror')[0]
  if (error) throw new Error(`not namespace-well-formed: ${error.textContent}`)
  return doc
}

function expandedNames(xml: string): string[] {
  const out: string[] = []
  const walk = (el: Element): void => {
    out.push(`{${el.namespaceURI ?? ''}}${el.localName}`)
    for (const a of Array.from(el.attributes)) {
      if (a.name === 'xmlns' || a.prefix === 'xmlns') continue
      out.push(`@{${a.namespaceURI ?? ''}}${a.localName}=${a.value}`)
    }
    for (const child of Array.from(el.children)) walk(child)
  }
  walk(parseStrictly(xml).documentElement)
  return out
}

describe('normalizeOoxmlXml under nested declarations', () => {
  it('keeps renamed elements in their namespace below a rebinding of the canonical prefix', () => {
    const xml =
      `<ns0:document xmlns:ns0="${W}"><ns0:body><x xmlns:w="urn:other">` +
      '<ns0:p><w:own/></ns0:p></x><ns0:p/></ns0:body></ns0:document>'
    const out = normalizeOoxmlXml(xml)
    expect(expandedNames(out)).toEqual(expandedNames(xml))
    expect(out.startsWith(`<w:document xmlns:w="${W}">`)).toBe(true)
  })

  it('keeps renamed attributes in their namespace below a rebinding', () => {
    const xml =
      `<ns0:document xmlns:ns0="${W}" xmlns:ns1="${R}"><ns0:body>` +
      '<x xmlns:r="urn:other"><ns0:p ns1:id="rId1" r:id="mine"/></x></ns0:body></ns0:document>'
    const out = normalizeOoxmlXml(xml)
    expect(expandedNames(out)).toEqual(expandedNames(xml))
  })

  it('keeps an Ignorable list naming the prefix that is bound where it stands', () => {
    const xml =
      `<ns0:document xmlns:ns0="${W}" xmlns:ns2="${MC}" xmlns:ns3="${W14}" ns2:Ignorable="ns3">` +
      `<x xmlns:w14="urn:other"><y ns2:Ignorable="ns3"/></x></ns0:document>`
    const out = normalizeOoxmlXml(xml)
    expect(expandedNames(out.replace(/ \w+:Ignorable="[^"]*"/g, ''))).toEqual(
      expandedNames(xml.replace(/ ns2:Ignorable="[^"]*"/g, '')),
    )
    const doc = parseStrictly(out)
    // each list names a prefix bound to W14 at that element
    for (const el of [doc.documentElement, doc.getElementsByTagName('y')[0]!]) {
      const listed = el.getAttributeNS(MC, 'Ignorable')!
      expect(el.lookupNamespaceURI(listed)).toBe(W14)
    }
  })

  it('renames the prefixes in markup-compatibility qualified-name lists', () => {
    const xml =
      `<ns0:document xmlns:ns0="${W}" xmlns:ns2="${MC}" xmlns:ns3="${W14}" ` +
      'ns2:ProcessContent="ns3:x" ns2:PreserveElements="ns3:*" ns2:PreserveAttributes="ns3:k"/>'
    const out = normalizeOoxmlXml(xml)
    expect(out).toContain('mc:ProcessContent="w14:x"')
    expect(out).toContain('mc:PreserveElements="w14:*"')
    expect(out).toContain('mc:PreserveAttributes="w14:k"')
  })

  it('gives fresh prefixes from a climbing counter (20,000 siblings)', () => {
    // r: is WordprocessingML here and renamed to w:; each sibling rebinds w: and
    // the name r: is the relationships namespace in the output, so every <r:t/>
    // needs a fresh prefix
    const doc = (siblings: number) =>
      `<r:document xmlns:r="${W}" xmlns:ns9="${R}">` +
      '<x xmlns:w="urn:other"><r:t/><ns9:k/></x>'.repeat(siblings) +
      '</r:document>'
    const started = performance.now()
    const out = normalizeOoxmlXml(doc(20_000))
    expect(performance.now() - started).toBeLessThan(1000)
    expect(out.match(/<r-\d+:t xmlns:r-\d+=/g)).toHaveLength(20_000)
    const small = doc(50)
    expect(expandedNames(normalizeOoxmlXml(small))).toEqual(expandedNames(small))
  })

  it('never declares one prefix twice on an element', () => {
    // ns0: takes w: by rename; w: itself (relationships, whose canonical r: is
    // taken by another namespace here) then has to move aside
    for (const decls of [
      `xmlns:ns0="${W}" xmlns:w="${R}" xmlns:r="urn:other"`,
      `xmlns:w="${R}" xmlns:r="urn:other" xmlns:ns0="${W}"`,
    ]) {
      const xml = `<root ${decls}><ns0:p/><w:id/><r:own/></root>`
      const out = normalizeOoxmlXml(xml)
      const root = out.slice(0, out.indexOf('>') + 1)
      const names = [...root.matchAll(/\s(xmlns:[^=\s]+)=/g)].map((m) => m[1])
      expect(new Set(names).size).toBe(names.length)
      expect(expandedNames(out)).toEqual(expandedNames(xml))
    }
  })

  it('stays linear on deep nesting', () => {
    // a declaration on each of 60,000 levels: the old frame-by-frame lookup was quadratic
    const depth = 60_000
    const xml =
      `<ns0:document xmlns:ns0="${W}">` +
      '<x xmlns:q="urn:q">'.repeat(depth) +
      '<ns0:p/>' +
      '</x>'.repeat(depth) +
      '</ns0:document>'
    const started = performance.now()
    const out = normalizeOoxmlXml(xml)
    expect(performance.now() - started).toBeLessThan(1000)
    expect(out).toContain('<w:p/>')
  })
})
