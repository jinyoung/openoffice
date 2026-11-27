/**
 * Namespaces bound to unconventional prefixes. lxml serializes DrawingML it got no
 * prefix for as ns0:, ns1:... declared on each element, so a deck written that way
 * opened with every text box empty: the parser looked for 'a:p' by name. The
 * archive now respells such parts with the canonical prefixes when it reads them.
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { createBlankPptx, openPptx, patchSlideXml, savePptx } from '../src/index'
import { canonicalizePrefixes, needsPrefixCanonicalization } from '../src/ns-prefixes'
import type { TableElement, TextElement } from '../src/types'

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'

const para = (prefix: string, text: string) =>
  `<${prefix}:p xmlns:${prefix}="${A}"><${prefix}:pPr algn="ctr"/>` +
  `<${prefix}:r><${prefix}:rPr sz="2400" b="1"><${prefix}:solidFill><${prefix}:srgbClr val="67519E"/></${prefix}:solidFill></${prefix}:rPr>` +
  `<${prefix}:t>${text}</${prefix}:t></${prefix}:r></${prefix}:p>`

/** A slide shaped like lxml output: only p: on the root, DrawingML declared where used. */
const LXML_SLIDE =
  `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<p:sld xmlns:p="${P}"><p:cSld><p:spTree>` +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  `<p:grpSpPr><a:xfrm xmlns:a="${A}"/></p:grpSpPr>` +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  `<p:spPr><a:xfrm xmlns:a="${A}"><a:off x="781050" y="2524125"/><a:ext cx="8096250" cy="514350"/></a:xfrm>` +
  `<a:prstGeom xmlns:a="${A}" prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
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

async function deckWithSlide(slideXml: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await createBlankPptx())
  zip.file('ppt/slides/slide1.xml', slideXml)
  return zip.generateAsync({ type: 'uint8array' })
}

const textOf = (el: TextElement | undefined) =>
  (el?.text?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join(''))

describe('canonicalizePrefixes', () => {
  it('leaves parts on the conventional prefixes untouched (same string)', () => {
    const xml = `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><a:p><a:r><a:t>x</a:t></a:r></a:p></p:sld>`
    expect(needsPrefixCanonicalization(xml)).toBe(false)
    expect(canonicalizePrefixes(xml)).toBe(xml)
  })

  it('respells ns0:/ns1: DrawingML as a: and declares a: (and r:) on the root', () => {
    const out = canonicalizePrefixes(LXML_SLIDE)
    expect(out).not.toMatch(/<\/?ns\d+:|xmlns:ns\d+=/)
    expect(out).toContain(`<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}">`)
    // the per-paragraph declaration is redundant under the root one and is dropped
    expect(out).toContain('<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2400" b="1">')
    expect(out).toContain('<a:t>Kicker line</a:t></a:r></a:p>')
    // tags that needed nothing keep their bytes, local declarations included
    expect(out).toContain(`<a:bodyPr xmlns:a="${A}" anchor="t"/>`)
  })

  it('maps a default DrawingML namespace and attribute prefixes', () => {
    const xml =
      `<p:sld xmlns:p="${P}"><p xmlns="${A}"><r><t>Default</t></r></p>` +
      `<p:pic><ns3:blip xmlns:ns3="${A}" xmlns:ns4="${R}" ns4:embed="rId2"/></p:pic></p:sld>`
    const out = canonicalizePrefixes(xml)
    expect(out).toContain(`<a:p xmlns="${A}"><a:r><a:t>Default</a:t></a:r></a:p>`)
    expect(out).toContain('<a:blip r:embed="rId2"/>')
  })

  it('keeps markup-compatibility prefix lists pointing at the renamed prefixes', () => {
    const xml =
      `<ns9:sld xmlns:ns9="${P}" xmlns:ns8="${MC}" xmlns:ns7="${P14}" ns8:Ignorable="ns7">` +
      `<ns8:AlternateContent><ns8:Choice Requires="ns7"><ns7:x/></ns8:Choice></ns8:AlternateContent></ns9:sld>`
    const out = canonicalizePrefixes(xml)
    expect(out).toContain('mc:Ignorable="p14"')
    expect(out).toContain('<mc:Choice Requires="p14"><p14:x/></mc:Choice>')
    expect(out.startsWith('<p:sld ')).toBe(true)
    expect(out.endsWith('</mc:AlternateContent></p:sld>')).toBe(true)
  })

  it('does not take a canonical prefix that is bound to another namespace in scope', () => {
    const xml = `<root xmlns:a="urn:other"><a:keep/><ns0:p xmlns:ns0="${A}"><ns0:t>x</ns0:t></ns0:p></root>`
    const out = canonicalizePrefixes(xml)
    expect(out).toContain('<a:keep/>')
    expect(out).toContain(`<ns0:p xmlns:ns0="${A}"><ns0:t>x</ns0:t></ns0:p>`)
  })

  it('an inner rebinding of a renamed prefix shadows the rename', () => {
    const xml = `<x:root xmlns:x="${P}"><x:inner xmlns:x="urn:other"><x:leaf/></x:inner></x:root>`
    const out = canonicalizePrefixes(xml)
    expect(out).toContain('<p:root ')
    expect(out).toContain('<x:inner xmlns:x="urn:other"><x:leaf/></x:inner></p:root>')
  })
})

describe('opening a deck written with ns0:-style prefixes', () => {
  it('reads the text boxes and the table cells', async () => {
    const opened = await openPptx(await deckWithSlide(LXML_SLIDE))
    const els = opened.deck.slides[0]!.elements
    const box = els.find((e) => e.type === 'text' || e.type === 'shape') as TextElement
    expect(textOf(box)).toEqual(['Kicker line', 'Second line', 'Third line'])
    expect(box.text!.paragraphs[0]!.align).toBe('center')
    expect(box.text!.paragraphs[0]!.runs[0]!.bold).toBe(true)
    const table = els.find((e) => e.type === 'table') as TableElement
    expect(table.rows[0]!.map((c) => c.text?.paragraphs[0]?.runs[0]?.text)).toEqual([
      'Cell A',
      'Cell B',
    ])
  })

  it('saves an unedited deck with the slide bytes unchanged', async () => {
    const bytes = await deckWithSlide(LXML_SLIDE)
    const opened = await openPptx(bytes)
    const out = await JSZip.loadAsync(await savePptx(opened))
    expect(await out.file('ppt/slides/slide1.xml')!.async('string')).toBe(LXML_SLIDE)
  })

  it('an edited text box saves into namespace-valid XML and reopens with the edit', async () => {
    const opened = await openPptx(await deckWithSlide(LXML_SLIDE))
    const slide = opened.deck.slides[0]!
    const box = slide.elements.find((e) => e.type === 'text' || e.type === 'shape') as TextElement
    box.text!.paragraphs[0]!.runs[0]!.text = 'Edited kicker'
    box.dirty = true
    const xml = patchSlideXml(slide)
    expect(xml).toContain(`xmlns:a="${A}"`)
    expect(xml).not.toMatch(/<\/?ns\d+:/)
    const reopened = await openPptx(await savePptx(opened))
    const again = reopened.deck.slides[0]!.elements.find(
      (e) => e.type === 'text' || e.type === 'shape',
    ) as TextElement
    expect(textOf(again)[0]).toBe('Edited kicker')
  })
})
