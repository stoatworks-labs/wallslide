/**
 * Write sample decks to a directory so they can be opened in the real
 * applications.
 *
 * The unit tests prove the package is a well-formed ZIP of well-formed XML.
 * They cannot prove PowerPoint will open it, and the gap between those two
 * statements is where every OOXML bug lives. This script closes it by hand:
 *
 *   node scripts/emit-samples.ts /tmp/out
 *   open /tmp/out/*.pptx
 *
 * Run with plain node — Node strips the types itself, no build step.
 */

import { deflateSync } from 'node:zlib'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { buildPptx, deckFileName, type DeckSpec } from '../src/lib/pptx'
import { slideFromResolution } from '../src/lib/slides'
import { OFFICE_COLOURS } from '../src/lib/theme'
import type { Brand, SafeArea, Wall } from '../src/types'

/**
 * A real PNG, built here rather than read off disk so the script has no fixture
 * to lose. Two flat halves and a border — enough to see at a glance whether
 * PowerPoint placed it full-bleed and the right way up.
 */
function makePng(w: number, h: number): Uint8Array {
  const raw = Buffer.alloc((w * 3 + 1) * h)
  let o = 0
  for (let y = 0; y < h; y++) {
    raw[o++] = 0 // filter: none
    for (let x = 0; x < w; x++) {
      const edge = x < 2 || y < 2 || x > w - 3 || y > h - 3
      const top = y < h / 2
      raw[o++] = edge ? 255 : top ? 200 : 20
      raw[o++] = edge ? 255 : top ? 30 : 120
      raw[o++] = edge ? 255 : top ? 40 : 220
    }
  }
  const chunk = (type: string, data: Buffer): Buffer => {
    const len = Buffer.alloc(4)
    len.writeUInt32BE(data.length)
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
    const crcBuf = Buffer.alloc(4)
    crcBuf.writeUInt32BE(crc32(body) >>> 0)
    return Buffer.concat([len, body, crcBuf])
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0)
  ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // colour type: truecolour
  return new Uint8Array(
    Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      chunk('IHDR', ihdr),
      chunk('IDAT', deflateSync(raw)),
      chunk('IEND', Buffer.alloc(0)),
    ]),
  )
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256)
  for (let i = 0; i < 256; i++) {
    let c = i
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[i] = c >>> 0
  }
  return t
})()

function crc32(b: Buffer): number {
  let c = 0xffffffff
  for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]!) & 0xff]! ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

const outDir = process.argv[2]
if (!outDir) {
  console.error('usage: node scripts/emit-samples.ts <output-dir>')
  process.exit(1)
}
mkdirSync(outDir, { recursive: true })

const brand: Brand = {
  colours: { ...OFFICE_COLOURS, accent1: 'C8102E', dk2: '1D2B3A' },
  majorFont: 'Helvetica Neue',
  minorFont: 'Helvetica Neue',
  logo: null,
  logoCorner: 'none',
  logoWidthPct: 12,
  background: null,
  useSolidBackground: true,
  backgroundColour: '0D1B2A',
}

const noSafe: SafeArea = { enabled: false, topPx: 0, rightPx: 0, bottomPx: 0, leftPx: 0 }

function make(wall: Wall, over: Partial<DeckSpec> = {}) {
  const slide = slideFromResolution(wall.widthPx, wall.heightPx, 96)
  if (!slide) throw new Error(`no slide spec for ${wall.name}`)
  const spec: DeckSpec = {
    wall,
    slide,
    brand,
    safeArea: noSafe,
    slides: [
      { kind: 'title', title: wall.name, subtitle: `${wall.widthPx} x ${wall.heightPx}` },
      {
        kind: 'heading',
        title: 'Export settings',
        body: [
          `Build scale ${slide.buildScale}`,
          `Export at ${slide.buildDpi} dpi`,
          'Rock & Roll <angle> "quoted"',
        ],
      },
      { kind: 'blank' },
    ],
    fileKind: 'pptx',
    exportNote: `Export at ${slide.buildDpi} dpi to land on ${wall.widthPx} x ${wall.heightPx}.`,
    now: new Date('2026-08-27T09:00:00Z'),
    ...over,
  }
  const name = deckFileName(wall, spec.fileKind)
  writeFileSync(join(outDir, name), buildPptx(spec))
  console.log(
    `${name.padEnd(38)} slide ${(slide.buildWidthMm / 25.4).toFixed(2)}" x ` +
      `${(slide.buildHeightMm / 25.4).toFixed(2)}"  scale ${slide.buildScale}  ` +
      `export ${slide.buildDpi} dpi  typeScale ${slide.typeScale.toFixed(2)}`,
  )
}

// 1:1 — fits inside 56" at 96 dpi with room to spare.
make({ name: 'Standard HD wall', widthPx: 1920, heightPx: 1080 })
// Over 56": must be halved and exported at 192 dpi.
make({ name: 'Ultrawide wall', widthPx: 7680, heightPx: 1080 })
// A tall portrait totem, and a template rather than a presentation.
make({ name: 'Portrait totem', widthPx: 1080, heightPx: 3840 }, { fileKind: 'potx' })
// Full-bleed picture slides: the path the pattern deck uses, and the only one
// PowerPoint had not been shown before this script existed.
make(
  { name: 'Pattern deck', widthPx: 3840, heightPx: 1080 },
  {
    slides: [
      { kind: 'title', title: 'Pattern deck', subtitle: '3840 x 1080' },
      { kind: 'full-bleed', image: { bytes: makePng(3840, 1080), mime: 'image/png' }, name: 'Grid' },
      { kind: 'full-bleed', image: { bytes: makePng(1920, 540), mime: 'image/png' }, name: 'Half raster' },
    ],
  },
)
// Safe area: the bottom 240 px is behind the band.
make(
  { name: 'Safe area wall', widthPx: 3840, heightPx: 1080 },
  { safeArea: { enabled: true, topPx: 0, rightPx: 0, bottomPx: 240, leftPx: 120 } },
)
