/**
 * The package writer, checked as a package rather than as a string.
 *
 * String assertions on generated XML pass while the file remains unopenable, so
 * everything here goes through a real ZIP round-trip and a real XML parser:
 * `unzip` unpacks it and `xmllint` parses every part. Between them they catch
 * the two failures that actually happen — a malformed archive and an unescaped
 * `&` in an event name — neither of which any amount of `toContain` would see.
 *
 * `xmllint` is on this machine via MacPorts and in most CI images. Where it is
 * missing the parse checks skip rather than fail: an absent tool is not a bug in
 * this code, and a suite that cannot run anywhere is worse than one that runs
 * with a gap it declares.
 */

import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'

import { buildPptx, safeBox, scaledSz, layoutBoxes, deckFileName, SlideSizeError } from '../pptx'
import type { DeckSpec } from '../pptx'
import { slideFromResolution } from '../slides'
import { OFFICE_COLOURS } from '../theme'
import type { Brand, SafeArea, Wall } from '../../types'

const dirs: string[] = []
afterAll(() => dirs.forEach((d) => rmSync(d, { recursive: true, force: true })))

function scratch(): string {
  const d = mkdtempSync(join(tmpdir(), 'wallslide-'))
  dirs.push(d)
  return d
}

function has(tool: string): boolean {
  try {
    // `which` rather than a shelled-out `command -v`: passing args alongside
    // `shell: true` concatenates them unescaped, which Node now warns about.
    execFileSync('/usr/bin/which', [tool], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}
const HAS_UNZIP = has('unzip')
const HAS_XMLLINT = has('xmllint')

const BRAND: Brand = {
  colours: OFFICE_COLOURS,
  majorFont: 'Calibri Light',
  minorFont: 'Calibri',
  logo: null,
  logoCorner: 'none',
  logoWidthPct: 12,
  background: null,
  useSolidBackground: false,
  backgroundColour: 'FFFFFF',
}

const NO_SAFE: SafeArea = { enabled: false, topPx: 0, rightPx: 0, bottomPx: 0, leftPx: 0 }

function spec(wall: Wall, over: Partial<DeckSpec> = {}): DeckSpec {
  const slide = slideFromResolution(wall.widthPx, wall.heightPx, 96)
  if (!slide) throw new Error('no slide spec')
  return {
    wall,
    slide,
    brand: BRAND,
    safeArea: NO_SAFE,
    slides: [{ kind: 'title', title: 'Load-in briefing', subtitle: 'Stage left wall' }],
    fileKind: 'pptx',
    exportNote: 'Export at 96 dpi.',
    now: new Date('2026-08-27T09:00:00Z'),
    ...over,
  }
}

function unpack(bytes: Uint8Array): { dir: string; names: string[] } {
  const dir = scratch()
  const file = join(dir, 'deck.pptx')
  writeFileSync(file, bytes)
  execFileSync('unzip', ['-o', '-q', file, '-d', join(dir, 'out')])
  const names = execFileSync('find', [join(dir, 'out'), '-type', 'f'], { encoding: 'utf8' })
    .trim()
    .split('\n')
    .map((p) => p.replace(join(dir, 'out') + '/', ''))
    .sort()
  return { dir: join(dir, 'out'), names }
}

describe('buildPptx', () => {
  it.skipIf(!HAS_UNZIP)('produces an archive real tools can unpack', () => {
    const { names } = unpack(buildPptx(spec({ name: 'Main', widthPx: 3840, heightPx: 1080 })))
    for (const required of [
      '[Content_Types].xml',
      '_rels/.rels',
      'docProps/core.xml',
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
      'ppt/presProps.xml',
      'ppt/theme/theme1.xml',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
    ]) {
      expect(names, `missing ${required}`).toContain(required)
    }
  })

  it.skipIf(!HAS_UNZIP || !HAS_XMLLINT)('emits well-formed XML in every part', () => {
    const { dir, names } = unpack(
      buildPptx(
        spec({ name: 'Main', widthPx: 3840, heightPx: 1080 }, {
          slides: [
            // An ampersand and an angle bracket in content, which is the escape
            // bug this test exists to catch.
            { kind: 'title', title: 'Rock & Roll <Hall>', subtitle: 'Q4 "review"' },
            { kind: 'heading', title: 'Cues', body: ['GO 1', 'GO 2 & 3'] },
            { kind: 'blank' },
          ],
        }),
      ),
    )
    const xml = names.filter((n) => n.endsWith('.xml') || n.endsWith('.rels'))
    expect(xml.length).toBeGreaterThan(8)
    // A part that will not parse is a package PowerPoint reports as damaged with
    // no further detail, so naming the part in the failure is the whole value of
    // this test.
    for (const name of xml) {
      const parse = () => execFileSync('xmllint', ['--noout', join(dir, name)], { stdio: 'pipe' })
      expect(parse, `${name} is not well-formed`).not.toThrow()
    }
  })

  it.skipIf(!HAS_UNZIP)('writes the build size, not the native size, into sldSz', () => {
    // 7680 px at 96 dpi is 80 inches, which PowerPoint refuses. slides.ts halves
    // it; the package must carry the halved size or the file will not open.
    const wall = { name: 'Ultrawide', widthPx: 7680, heightPx: 1080 }
    const s = slideFromResolution(wall.widthPx, wall.heightPx, 96)
    expect(s?.buildScale).toBe(0.5)

    const { dir } = unpack(buildPptx(spec(wall)))
    const pres = readFileSync(join(dir, 'ppt/presentation.xml'), 'utf8')
    // 40 inches x 914400.
    expect(pres).toContain('<p:sldSz cx="36576000" cy="5143500"/>')
  })

  it('refuses a slide it cannot legally express rather than emitting one', () => {
    // 60:1. slides.ts finds no whole scale that is both under 56" and over 1".
    const wall = { name: 'Ticker', widthPx: 12000, heightPx: 200 }
    expect(() => buildPptx(spec(wall))).toThrow(SlideSizeError)
  })

  it.skipIf(!HAS_UNZIP)('marks a template with the template content type', () => {
    const { dir } = unpack(
      buildPptx(spec({ name: 'Main', widthPx: 1920, heightPx: 1080 }, { fileKind: 'potx' })),
    )
    const ct = readFileSync(join(dir, '[Content_Types].xml'), 'utf8')
    expect(ct).toContain('presentationml.template.main+xml')
    expect(ct).not.toContain('presentationml.presentation.main+xml')
  })

  it.skipIf(!HAS_UNZIP)('hides master shapes behind a full-bleed pattern', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const { dir } = unpack(
      buildPptx(
        spec({ name: 'Main', widthPx: 1920, heightPx: 1080 }, {
          slides: [{ kind: 'full-bleed', image: { bytes: png, mime: 'image/png' }, name: 'Grid' }],
        }),
      ),
    )
    expect(readFileSync(join(dir, 'ppt/slides/slide1.xml'), 'utf8')).toContain('showMasterSp="0"')
    expect(readFileSync(join(dir, 'ppt/media/image1.png'), null)).toEqual(Buffer.from(png))
  })

  it('is byte-identical for two builds of the same deck', () => {
    const wall = { name: 'Main', widthPx: 1920, heightPx: 1080 }
    expect(Buffer.from(buildPptx(spec(wall)))).toEqual(Buffer.from(buildPptx(spec(wall))))
  })
})

describe('safeBox', () => {
  it('scales wall pixels into slide EMU through the slide, not the DPI', () => {
    // A half-scale build: 100 px of a 1000 px wall is a tenth of the slide
    // whatever the slide's absolute size.
    const box = safeBox(1_000_000, 500_000, { name: '', widthPx: 1000, heightPx: 500 }, {
      enabled: true,
      leftPx: 100,
      rightPx: 0,
      topPx: 50,
      bottomPx: 0,
    })
    expect(box.x).toBe(100_000)
    expect(box.y).toBe(50_000)
    expect(box.cx).toBe(900_000)
    expect(box.cy).toBe(450_000)
  })

  it('ignores a safe area that would swallow the slide', () => {
    const box = safeBox(1000, 1000, { name: '', widthPx: 100, heightPx: 100 }, {
      enabled: true,
      leftPx: 60,
      rightPx: 60,
      topPx: 0,
      bottomPx: 0,
    })
    expect(box).toEqual({ x: 0, y: 0, cx: 1000, cy: 1000 })
  })
})

describe('scaledSz', () => {
  it('scales type with the slide', () => {
    expect(scaledSz(44, 1)).toBe(4400)
    expect(scaledSz(44, 3)).toBe(13200)
  })
  it("clamps to PowerPoint's own 1pt–4000pt range", () => {
    expect(scaledSz(44, 1000)).toBe(400000)
    expect(scaledSz(1, 0.0001)).toBe(100)
  })
})

describe('layoutBoxes', () => {
  it('keeps every box inside the safe area it was given', () => {
    const safe = { x: 200, y: 100, cx: 800, cy: 400 }
    const b = layoutBoxes(safe)
    for (const [name, box] of Object.entries(b)) {
      expect(box.x, name).toBeGreaterThanOrEqual(safe.x)
      expect(box.y, name).toBeGreaterThanOrEqual(safe.y)
      expect(box.x + box.cx, name).toBeLessThanOrEqual(safe.x + safe.cx)
      expect(box.y + box.cy, name).toBeLessThanOrEqual(safe.y + safe.cy)
    }
  })
})

describe('deckFileName', () => {
  it('strips characters a file system will not take', () => {
    expect(deckFileName({ name: 'Main/Wall: 1', widthPx: 3840, heightPx: 1080 }, 'potx')).toBe(
      'Main-Wall- 1 3840x1080.potx',
    )
  })
})
