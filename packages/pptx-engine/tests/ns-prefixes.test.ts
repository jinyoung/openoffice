/**
 * @vitest-environment jsdom
 */
/// <reference lib="dom" />
// The package's own tsconfig omits the DOM lib (headless engine); DOMParser is
// pulled in here to read the output with a namespace-aware parser.

/**
 * Namespaces bound to unconventional prefixes. lxml serializes DrawingML it got no
 * prefix for as ns0:, ns1:... declared on each element, so a deck written that way
 * opened with every text box empty: the parser looked for 'a:p' by name. The
 * archive now respells such parts with the canonical prefixes when it reads them.
 *
 * Every rewrite is checked with a namespace-aware parser: the output must parse
 * (no unbound prefix) and give each element and attribute the same expanded name
 * as the input.
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import {
  addElement,
  copyElementData,
  copySlide,
  createBlankPptx,
  ensureBuiltinLayout,
  openPptx,
  parseMasterPart,
  pasteElements,
  pasteSlide,
  patchSlideXml,
  savePptx,
} from '../src/index'
import { notesPathForSlide, setSlideNotes } from '../src/notes'
import {
  canonicalizePrefixes,
  declareGeneratedPrefixes,
  needsPrefixCanonicalization,
} from '../src/ns-prefixes'
import type { SlideElement, TableElement, TextElement } from '../src/types'

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main'
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
const VML = 'urn:schemas-microsoft-com:vml'
const VML_PPT = 'urn:schemas-microsoft-com:office:powerpoint'

/** parse as XML with namespaces; an unbound prefix is an error, as in PowerPoint */
function parseStrictly(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  const error = doc.getElementsByTagName('parsererror')[0]
  if (error) throw new Error(`not namespace-well-formed: ${error.textContent}`)
  return doc
}

/** every element and attribute as {namespace}local (declarations left out), in order */
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

/** the rewrite parses and means exactly what the input meant */
function expectSameXml(input: string, output: string): void {
  expect(expandedNames(output)).toEqual(expandedNames(input))
}

const para = (prefix: string, text: string) =>
  `<${prefix}:p xmlns:${prefix}="${A}"><${prefix}:pPr algn="ctr"/>` +
  `<${prefix}:r><${prefix}:rPr sz="2400" b="1"><${prefix}:solidFill><${prefix}:srgbClr val="67519E"/></${prefix}:solidFill></${prefix}:rPr>` +
  `<${prefix}:t>${text}</${prefix}:t></${prefix}:r></${prefix}:p>`

const HIDDEN_FILL_EXT =
  `<a:extLst xmlns:a="${A}"><a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}">` +
  `<ns5:hiddenFill xmlns:ns5="${A14}"><ns6:solidFill xmlns:ns6="${A}"><ns6:srgbClr val="FFFFFF"/></ns6:solidFill></ns5:hiddenFill>` +
  '</a:ext></a:extLst>'

/** A slide shaped like lxml output: only p: on the root, DrawingML declared where used. */
const LXML_SLIDE =
  `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<p:sld xmlns:p="${P}"><p:cSld><p:spTree>` +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  `<p:grpSpPr><a:xfrm xmlns:a="${A}"/></p:grpSpPr>` +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  `<p:spPr><a:xfrm xmlns:a="${A}"><a:off x="781050" y="2524125"/><a:ext cx="8096250" cy="514350"/></a:xfrm>` +
  `<a:prstGeom xmlns:a="${A}" prst="rect"><a:avLst/></a:prstGeom>${HIDDEN_FILL_EXT}</p:spPr>` +
  `<p:txBody><a:bodyPr xmlns:a="${A}" anchor="t"/><a:lstStyle xmlns:a="${A}"/>` +
  `${para('ns0', 'Kicker line')}${para('ns1', 'Second line')}${para('ns0', 'Third line')}` +
  '</p:txBody></p:sp>' +
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  `<p:xfrm><a:off xmlns:a="${A}" x="0" y="3000000"/><a:ext xmlns:a="${A}" cx="4000000" cy="400000"/></p:xfrm>` +
  `<ns2:graphic xmlns:ns2="${A}"><ns2:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
  '<ns2:tbl><ns2:tblPr/><ns2:tblGrid><ns2:gridCol w="2000000"/><ns2:gridCol w="2000000"/></ns2:tblGrid>' +
  '<ns2:tr h="400000">' +
  '<ns2:tc><ns2:txBody><ns2:bodyPr/><ns2:p><ns2:r><ns2:t>Cell A</ns2:t></ns2:r></ns2:p></ns2:txBody><ns2:tcPr/></ns2:tc>' +
  '<ns2:tc><ns2:txBody><ns2:bodyPr/><ns2:p><ns2:r><ns2:t>Cell B</ns2:t></ns2:r></ns2:p></ns2:txBody><ns2:tcPr/></ns2:tc>' +
  '</ns2:tr></ns2:tbl></ns2:graphicData></ns2:graphic></p:graphicFrame>' +
  '</p:spTree></p:cSld></p:sld>'

/** a conventional part respelled with another prefix for one namespace (as lxml may write it) */
const respell = (xml: string, from: string, to: string) =>
  xml
    .replace(new RegExp(`xmlns:${from}=`, 'g'), `xmlns:${to}=`)
    .replace(new RegExp(`<(/?)${from}:`, 'g'), `<$1${to}:`)

const MASTER = 'ppt/slideMasters/slideMaster1.xml'
const LAYOUT = 'ppt/slideLayouts/slideLayout1.xml'
const THEME = 'ppt/theme/theme1.xml'

async function deckWithSlide(
  slideXml: string,
  opts?: { lxmlChain?: boolean },
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await createBlankPptx())
  zip.file('ppt/slides/slide1.xml', slideXml)
  if (opts?.lxmlChain) {
    for (const path of [MASTER, LAYOUT, THEME]) {
      zip.file(path, respell(await zip.file(path)!.async('string'), 'a', 'ns0'))
    }
  }
  return zip.generateAsync({ type: 'uint8array' })
}

const textOf = (el: SlideElement | undefined) =>
  ((el as TextElement | undefined)?.text?.paragraphs ?? []).map((p) =>
    p.runs.map((r) => r.text).join(''),
  )
const firstText = (els: SlideElement[]) =>
  els.find((e) => e.type === 'text' || e.type === 'shape') as TextElement

describe('canonicalizePrefixes', () => {
  it('leaves parts on the conventional prefixes untouched (same string)', () => {
    const xml = `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><a:p><a:r><a:t>x</a:t></a:r></a:p></p:sld>`
    expect(needsPrefixCanonicalization(xml)).toBe(false)
    expect(canonicalizePrefixes(xml)).toBe(xml)
  })

  it('respells ns0:/ns1: DrawingML as a: on the element that declares it', () => {
    const out = canonicalizePrefixes(LXML_SLIDE)
    expectSameXml(LXML_SLIDE, out)
    expect(out).not.toMatch(/<\/?ns\d+:|xmlns:ns\d+=/)
    // the root is not touched: declarations stay where they stood
    expect(out).toContain(`<p:sld xmlns:p="${P}"><p:cSld>`)
    expect(out).toContain(`<a:p xmlns:a="${A}"><a:pPr algn="ctr"/><a:r><a:rPr sz="2400" b="1">`)
    // tags that needed nothing keep their bytes
    expect(out).toContain(`<a:bodyPr xmlns:a="${A}" anchor="t"/>`)
  })

  it('keeps every fragment self-contained', () => {
    const out = canonicalizePrefixes(LXML_SLIDE)
    const shape = out.slice(out.indexOf('<p:sp>'), out.indexOf('</p:sp>') + '</p:sp>'.length)
    // the shape alone parses once p: is bound: no declaration was hoisted out of it
    const alone = `<p:spTree xmlns:p="${P}">${shape}</p:spTree>`
    const doc = parseStrictly(alone)
    expect(doc.getElementsByTagNameNS(A14, 'hiddenFill')).toHaveLength(1)
    expect(shape).toContain(`<a14:hiddenFill xmlns:a14="${A14}">`)
  })

  it('maps a default DrawingML namespace and attribute prefixes', () => {
    const xml =
      `<p:sld xmlns:p="${P}"><p xmlns="${A}"><r><t>Default</t></r></p>` +
      `<p:pic><ns3:blip xmlns:ns3="${A}" xmlns:ns4="${R}" ns4:embed="rId2"/></p:pic></p:sld>`
    const out = canonicalizePrefixes(xml)
    expectSameXml(xml, out)
    expect(out).toContain(`<a:p xmlns="${A}" xmlns:a="${A}"><a:r><a:t>Default</a:t></a:r></a:p>`)
    expect(out).toContain(`<a:blip xmlns:a="${A}" xmlns:r="${R}" r:embed="rId2"/>`)
  })

  it('keeps markup-compatibility lists pointing at the renamed prefixes', () => {
    const xml =
      `<ns9:sld xmlns:ns9="${P}" xmlns:ns8="${MC}" xmlns:ns7="${P14}" xmlns:ns6="${VML}" ` +
      'ns8:Ignorable="ns7" ns8:MustUnderstand="ns7" ns8:ProcessContent="ns7:x" ' +
      'ns8:PreserveElements="ns6:shape ns7:*" ns8:PreserveAttributes="ns7:k">' +
      `<ns8:AlternateContent><ns8:Choice Requires="ns7"><ns7:x/></ns8:Choice></ns8:AlternateContent></ns9:sld>`
    const out = canonicalizePrefixes(xml)
    expectSameXml(
      xml.replace(/ns8:\w+="[^"]*" ?/g, '').replace(' Requires="ns7"', ''),
      out.replace(/mc:\w+="[^"]*" ?/g, '').replace(' Requires="p14"', ''),
    )
    expect(out).toContain('mc:Ignorable="p14"')
    expect(out).toContain('mc:MustUnderstand="p14"')
    expect(out).toContain('mc:ProcessContent="p14:x"')
    expect(out).toContain('mc:PreserveElements="v:shape p14:*"')
    expect(out).toContain('mc:PreserveAttributes="p14:k"')
    expect(out).toContain('<mc:Choice Requires="p14"><p14:x/></mc:Choice>')
  })

  it('does not take a canonical prefix that is bound to another namespace in scope', () => {
    const xml = `<root xmlns:a="urn:other"><a:keep/><ns0:p xmlns:ns0="${A}"><ns0:t>x</ns0:t></ns0:p></root>`
    const out = canonicalizePrefixes(xml)
    expectSameXml(xml, out)
    expect(out).toBe(xml)
  })

  it('an inner rebinding of a renamed source prefix shadows the rename', () => {
    const xml = `<x:root xmlns:x="${P}"><x:inner xmlns:x="urn:other"><x:leaf/></x:inner></x:root>`
    const out = canonicalizePrefixes(xml)
    expectSameXml(xml, out)
    expect(out).toContain('<x:inner xmlns:x="urn:other"><x:leaf/></x:inner></p:root>')
  })

  describe('a descendant binding the output prefix to another namespace', () => {
    it('does not capture renamed elements below it', () => {
      const xml = `<ns0:sp xmlns:ns0="${A}"><x xmlns:a="urn:other"><ns0:t/><a:own/></x><ns0:t/></ns0:sp>`
      const out = canonicalizePrefixes(xml)
      expectSameXml(xml, out)
      expect(out.startsWith(`<a:sp xmlns:a="${A}">`)).toBe(true)
    })

    it('does not capture renamed attributes below it', () => {
      const xml = `<ns5:r xmlns:ns5="${R}"><x xmlns:r="urn:other"><y ns5:id="rId1" r:id="mine"/></x></ns5:r>`
      const out = canonicalizePrefixes(xml)
      expectSameXml(xml, out)
    })

    it('does not capture PresentationML under a VML shape that binds p: to its own namespace', () => {
      const xml =
        `<ns9:sld xmlns:ns9="${P}"><v:shape xmlns:v="${VML}" xmlns:p="${VML_PPT}">` +
        '<p:x/><ns9:y/></v:shape><ns9:z/></ns9:sld>'
      const out = canonicalizePrefixes(xml)
      expectSameXml(xml, out)
      expect(out).toContain('<p:x/>')
    })
  })

  it('skips comments, CDATA, processing instructions and attribute values', () => {
    const decoys =
      `<!-- <ns0:p xmlns:ns0="${A}"> -->` +
      `<p:x descr='a > b xmlns:ns0="${A}"' title="ns0:t />">` +
      `<![CDATA[<ns0:t xmlns:ns0="${A}">]]><?pi xmlns:ns0="${A}"?>t xmlns:ns0="${A}"</p:x>`
    const only = `<p:sld xmlns:p="${P}">${decoys}</p:sld>`
    expect(needsPrefixCanonicalization(only)).toBe(false)
    expect(canonicalizePrefixes(only)).toBe(only)
    const mixed = `<p:sld xmlns:p="${P}">${decoys}${para('ns0', 'Real')}</p:sld>`
    const out = canonicalizePrefixes(mixed)
    expect(out).toContain(decoys)
    expect(out).toContain('<a:t>Real</a:t>')
    expectSameXml(mixed, out)
  })

  it('gives fresh prefixes from a climbing counter (20,000 siblings)', () => {
    // p: is DrawingML here and renamed to a:; each sibling rebinds a: and the
    // name p: is PresentationML in the output, so every <p:t/> needs a fresh prefix
    const doc = (siblings: number) =>
      `<p:r xmlns:p="${A}" xmlns:ns9="${P}">` +
      '<x xmlns:a="urn:other"><p:t/><ns9:k/></x>'.repeat(siblings) +
      '</p:r>'
    const big = doc(20_000)
    const started = performance.now()
    const out = canonicalizePrefixes(big)
    expect(performance.now() - started).toBeLessThan(1000)
    expect(out.match(/<p-\d+:t xmlns:p-\d+=/g)).toHaveLength(20_000)
    // the namespace-aware comparison on a sample (DOM parsing 20,000 elements is the slow part)
    const small = doc(50)
    expectSameXml(small, canonicalizePrefixes(small))
  })

  it('never declares one prefix twice on an element', () => {
    // ns0: takes a: by rename; a: itself (PresentationML, whose canonical p: is
    // taken by another namespace here) then has to move aside
    for (const decls of [
      `xmlns:ns0="${A}" xmlns:a="${P}" xmlns:p="urn:other"`,
      `xmlns:a="${P}" xmlns:p="urn:other" xmlns:ns0="${A}"`,
    ]) {
      const xml = `<root ${decls}><ns0:t/><a:sld/><p:own/></root>`
      const out = canonicalizePrefixes(xml)
      const root = out.slice(0, out.indexOf('>') + 1)
      const names = [...root.matchAll(/\s(xmlns:[^=\s]+)=/g)].map((m) => m[1])
      expect(new Set(names).size).toBe(names.length)
      expectSameXml(xml, out)
    }
  })

  it('stays linear on deep nesting', () => {
    // a declaration on each of 60,000 levels: the old frame-by-frame lookup took 12 s
    const depth = 60_000
    const xml =
      `<ns0:r xmlns:ns0="${A}">` +
      `<x xmlns:q="urn:q">`.repeat(depth) +
      '<ns0:t/>' +
      '</x>'.repeat(depth) +
      '</ns0:r>'
    const started = performance.now()
    const out = canonicalizePrefixes(xml)
    expect(performance.now() - started).toBeLessThan(1000)
    expect(out).toContain('<a:t/>')
  })
})

describe('declareGeneratedPrefixes', () => {
  it('adds the a:/r:/p: declarations a root lacks and keeps a root that has them', () => {
    expect(declareGeneratedPrefixes(`<?xml version="1.0"?><p:sld xmlns:p="${P}"><p:cSld>`)).toBe(
      `<?xml version="1.0"?><p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld>`,
    )
    const full = `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld>`
    expect(declareGeneratedPrefixes(full)).toBe(full)
  })
})

describe('a deck written with ns0:-style prefixes', () => {
  it('opens with its text boxes and table cells', async () => {
    const opened = await openPptx(await deckWithSlide(LXML_SLIDE, { lxmlChain: true }))
    const els = opened.deck.slides[0]!.elements
    const box = firstText(els)
    expect(textOf(box)).toEqual(['Kicker line', 'Second line', 'Third line'])
    expect(box.text!.paragraphs[0]!.align).toBe('center')
    expect(box.text!.paragraphs[0]!.runs[0]!.bold).toBe(true)
    const table = els.find((e) => e.type === 'table') as TableElement
    expect(table.rows[0]!.map((c) => c.text?.paragraphs[0]?.runs[0]?.text)).toEqual([
      'Cell A',
      'Cell B',
    ])
  })

  it('saves unedited with every part byte-identical, the respelled layout, master and theme too', async () => {
    const bytes = await deckWithSlide(LXML_SLIDE, { lxmlChain: true })
    const before = await JSZip.loadAsync(bytes)
    const after = await JSZip.loadAsync(await savePptx(await openPptx(bytes)))
    for (const path of Object.keys(before.files)) {
      if (before.files[path]!.dir) continue
      const a = await before.file(path)!.async('uint8array')
      const b = await after.file(path)!.async('uint8array')
      expect(Buffer.from(b).equals(Buffer.from(a)), path).toBe(true)
    }
    expect(await after.file(MASTER)!.async('string')).toContain('<ns0:')
  })

  it('an edited text box saves namespace-valid XML and reopens with the edit', async () => {
    const opened = await openPptx(await deckWithSlide(LXML_SLIDE))
    const box = firstText(opened.deck.slides[0]!.elements)
    box.text!.paragraphs[0]!.runs[0]!.text = 'Edited kicker'
    box.dirty = true
    const saved = await JSZip.loadAsync(await savePptx(opened))
    const xml = await saved.file('ppt/slides/slide1.xml')!.async('string')
    const doc = parseStrictly(xml)
    const texts = Array.from(doc.getElementsByTagNameNS(A, 't')).map((t) => t.textContent)
    expect(texts).toContain('Edited kicker')
    expect(doc.getElementsByTagNameNS(A14, 'hiddenFill')).toHaveLength(1)
  })

  it('a pasted shape stays namespace-valid in a conventional deck', async () => {
    const source = await openPptx(await deckWithSlide(LXML_SLIDE))
    const slide = source.deck.slides[0]!
    const item = copyElementData(source, slide, firstText(slide.elements))
    const target = await openPptx(await createBlankPptx())
    expect(pasteElements(target, 0, [item], { dx: 0, dy: 0 })).not.toBeNull()
    const saved = await JSZip.loadAsync(await savePptx(target))
    const doc = parseStrictly(await saved.file('ppt/slides/slide1.xml')!.async('string'))
    expect(doc.getElementsByTagNameNS(A14, 'hiddenFill')).toHaveLength(1)
    expect(Array.from(doc.getElementsByTagNameNS(A, 't')).map((t) => t.textContent)).toEqual([
      'Kicker line',
      'Second line',
      'Third line',
    ])
  })

  it('"keep source formatting" reuses its respelled master instead of copying it again', async () => {
    const opened = await openPptx(await deckWithSlide(LXML_SLIDE, { lxmlChain: true }))
    const bundle = copySlide(opened, 0)!
    expect(pasteSlide(opened, 0, bundle, { keepSourceFormatting: true })).not.toBeNull()
    const masters = [...opened.archive.entries.keys()].filter((p) =>
      /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p),
    )
    expect(masters).toEqual([MASTER])
  })
})

describe('a slide declaring DrawingML only locally, on conventional prefixes', () => {
  // nothing to respell, but the root lacks a: — generated fragments need it there
  const LOCAL_ONLY = [
    ['ns0', 'a'],
    ['ns1', 'a'],
    ['ns2', 'a'],
    ['ns5', 'a14'],
    ['ns6', 'a'],
  ].reduce((xml, [from, to]) => respell(xml, from!, to!), LXML_SLIDE)

  it('rewrites its text body into namespace-valid XML', async () => {
    expect(needsPrefixCanonicalization(LOCAL_ONLY)).toBe(false)
    const opened = await openPptx(await deckWithSlide(LOCAL_ONLY))
    const slide = opened.deck.slides[0]!
    const box = firstText(slide.elements)
    // a run-count change takes the full txBody rebuild (fresh a: markup, no declarations)
    box.text!.paragraphs[0]!.runs.push({ text: ' added' })
    box.dirty = true
    const xml = patchSlideXml(slide)
    expect(xml).toContain(`xmlns:a="${A}"`)
    const doc = parseStrictly(xml)
    expect(Array.from(doc.getElementsByTagNameNS(A, 't')).map((t) => t.textContent)).toContain(
      ' added',
    )
    const reopened = await openPptx(await savePptx(opened))
    expect(textOf(firstText(reopened.deck.slides[0]!.elements))[0]).toBe('Kicker line added')
  })
})

describe('masters, layouts and notes declaring DrawingML only locally', () => {
  /** a part with a:/r: dropped from its root and declared on each element instead (lxml-like) */
  const localOnly = (xml: string) =>
    xml
      .replace(` xmlns:a="${A}"`, '')
      .replace(` xmlns:r="${R}"`, '')
      .replace(/<a:(\w+)/g, `<a:$1 xmlns:a="${A}"`)
      .replace(/<(p:\w+)((?:\s[^>]*?)?\sr:\w+=)/g, `<$1 xmlns:r="${R}"$2`)

  async function localOnlyDeck(): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(await createBlankPptx())
    for (const path of [MASTER, LAYOUT]) {
      const xml = localOnly(await zip.file(path)!.async('string'))
      parseStrictly(xml)
      expect(xml.slice(0, xml.indexOf('<p:cSld'))).not.toContain('xmlns:a=')
      zip.file(path, xml)
    }
    return zip.generateAsync({ type: 'uint8array' })
  }

  it('a master edit writes the generated declarations on its root', async () => {
    const opened = await openPptx(await localOnlyDeck())
    const master = parseMasterPart(opened.archive, MASTER)!
    addElement(master, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
      paragraphs: [{ runs: [{ text: 'Master note' }] }],
    })
    const doc = parseStrictly(patchSlideXml(master))
    expect(Array.from(doc.getElementsByTagNameNS(A, 't')).map((t) => t.textContent)).toContain(
      'Master note',
    )
  })

  it('a layout registered on such a master keeps the master namespace-valid', async () => {
    const opened = await openPptx(await localOnlyDeck())
    expect(ensureBuiltinLayout(opened.archive, opened.deck.size, 'titleContent')).not.toBeNull()
    const master = Buffer.from(opened.archive.entries.get(MASTER)!).toString('utf8')
    expect(parseStrictly(master).getElementsByTagNameNS(P, 'sldLayoutId').length).toBe(2)
  })

  it('rewritten speaker notes stay namespace-valid', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(setSlideNotes(opened, 0, 'first')).toBe(true)
    const notesPath = notesPathForSlide(opened.archive, opened.deck.slides[0]!.path)!
    const notes = Buffer.from(opened.archive.entries.get(notesPath)!).toString('utf8')
    opened.archive.entries.set(notesPath, Buffer.from(localOnly(notes), 'utf8'))
    expect(setSlideNotes(opened, 0, 'second line')).toBe(true)
    const doc = parseStrictly(Buffer.from(opened.archive.entries.get(notesPath)!).toString('utf8'))
    expect(Array.from(doc.getElementsByTagNameNS(A, 't')).map((t) => t.textContent)).toEqual([
      'second line',
    ])
  })
})
