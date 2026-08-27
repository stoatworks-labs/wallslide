/**
 * Test patterns, rendered at the wall's native raster and embedded full-bleed.
 *
 * THIS IS NOT test-card, AND SHOULD NOT GROW INTO IT
 * ==================================================
 * test-card is the serious pattern generator: its SMPTE RP 219 is measured
 * against ffmpeg's own `smptehdbars` at four resolutions, it imports real
 * cabinet maps out of Pixel Peeker, and it writes PNGs at the exact raster of
 * every physical output. If somebody needs patterns, that is where they should
 * go, and the UI says so.
 *
 * What lives here is the subset that is useful INSIDE A DECK, which is a
 * different question, because a deck is a lossy path to a wall:
 *
 *   PNG -> PowerPoint's renderer -> screen scaler -> LED processor
 *
 * RP 219 is deliberately absent. Its entire value is that the fifteen colours
 * are exactly right, and a pattern whose point is exact colour should not be
 * sent down a path that may colour-manage it. Shipping it here would put
 * test-card's carefully measured name on a number this app cannot stand behind.
 * The bars below are plain 75% bars and are labelled as such.
 *
 * THE DISTINCTION EVERY PATTERN CARRIES
 * =====================================
 * `reads` says what a failure of this pattern actually tells you. Most patterns
 * diagnose the WALL. The 1px checkerboard and the line bursts diagnose the
 * CHAIN — if they come out as flat grey, something between PowerPoint and the
 * panels is scaling, which is worth knowing and is not a fault in the wall. A
 * deck that does not make that distinction will have somebody re-terminating a
 * healthy wall at midnight.
 */

import type { PatternKind, SafeArea, Wall } from '../types'

/** Conservative canvas ceilings, as measured for test-card. Browsers return a
 *  BLANK canvas past these rather than throwing, so they must be checked up
 *  front — a folder of empty PNGs looks fine in a file listing. */
export const MAX_SIDE_PX = 16384
export const MAX_AREA_PX = 268_435_456

export function canvasLimitProblem(width: number, height: number): string | null {
  if (!(width > 0) || !(height > 0) || !Number.isFinite(width) || !Number.isFinite(height)) {
    return 'The raster must be a positive number of pixels on both edges.'
  }
  if (width > MAX_SIDE_PX || height > MAX_SIDE_PX) {
    return `${width}×${height} is over the ${MAX_SIDE_PX} px per-side canvas limit; the browser would return a blank image rather than an error.`
  }
  if (width * height > MAX_AREA_PX) {
    return `${width}×${height} is ${((width * height) / 1e6).toFixed(0)} MP, over the ~${(
      MAX_AREA_PX / 1e6
    ).toFixed(0)} MP canvas area limit; the browser would return a blank image rather than an error.`
  }
  return null
}

export interface PatternDef {
  id: PatternKind
  label: string
  purpose: string
  /** What a failure of this pattern is evidence of. */
  reads: 'wall' | 'chain'
}

export const PATTERNS: PatternDef[] = [
  {
    id: 'grid',
    label: 'Grid',
    purpose: 'Geometry and scaling. Squares that are not square mean the aspect is wrong.',
    reads: 'wall',
  },
  {
    id: 'alignment',
    label: 'Alignment',
    purpose: 'Edges, centre and corners, with the pixel counts written on them.',
    reads: 'wall',
  },
  {
    id: 'bars',
    label: '75% bars',
    purpose: 'Eight bars for a quick colour and order check. Not a standards-grade pattern.',
    reads: 'wall',
  },
  {
    id: 'greyscale',
    label: 'Greyscale wedge',
    purpose: 'Banding, crush and clipping across the range.',
    reads: 'wall',
  },
  {
    id: 'pixel-check',
    label: 'Pixel check',
    purpose:
      'A 1px checkerboard and line bursts. If this renders as flat grey, something in the chain is scaling — this is a test of the path, not of the wall.',
    reads: 'chain',
  },
  { id: 'safe-area', label: 'Safe area', purpose: 'Draws the masked margins to scale.', reads: 'wall' },
  { id: 'solid-white', label: 'Solid white', purpose: 'Uniformity and dead pixels.', reads: 'wall' },
  { id: 'solid-black', label: 'Solid black', purpose: 'Black level and stuck pixels.', reads: 'wall' },
  { id: 'solid-red', label: 'Solid red', purpose: 'Per-channel dead pixels.', reads: 'wall' },
  { id: 'solid-green', label: 'Solid green', purpose: 'Per-channel dead pixels.', reads: 'wall' },
  { id: 'solid-blue', label: 'Solid blue', purpose: 'Per-channel dead pixels.', reads: 'wall' },
]

export function patternDef(id: PatternKind): PatternDef {
  const found = PATTERNS.find((p) => p.id === id)
  if (!found) throw new Error(`Unknown pattern ${id}`)
  return found
}

type Ctx = CanvasRenderingContext2D
const MONO = 'ui-monospace, "SF Mono", Menlo, Consolas, monospace'

/** Type size that stays readable whatever the raster. */
function baseFont(w: number, h: number): number {
  return Math.max(14, Math.round(Math.min(w, h) / 28))
}

function drawGrid(ctx: Ctx, w: number, h: number) {
  const step = Math.max(16, Math.round(Math.min(w, h) / 16))
  ctx.fillStyle = '#000'
  ctx.fillRect(0, 0, w, h)
  ctx.strokeStyle = '#3a3a3a'
  ctx.lineWidth = 1
  for (let x = 0; x <= w; x += step) {
    ctx.beginPath()
    ctx.moveTo(x + 0.5, 0)
    ctx.lineTo(x + 0.5, h)
    ctx.stroke()
  }
  for (let y = 0; y <= h; y += step) {
    ctx.beginPath()
    ctx.moveTo(0, y + 0.5)
    ctx.lineTo(w, y + 0.5)
    ctx.stroke()
  }
  // Heavy line every fifth, so the eye can count without following every line.
  ctx.strokeStyle = '#8a8a8a'
  ctx.lineWidth = 2
  for (let i = 0, x = 0; x <= w; i++, x += step * 5) {
    ctx.beginPath()
    ctx.moveTo(x, 0)
    ctx.lineTo(x, h)
    ctx.stroke()
  }
  for (let i = 0, y = 0; y <= h; i++, y += step * 5) {
    ctx.beginPath()
    ctx.moveTo(0, y)
    ctx.lineTo(w, y)
    ctx.stroke()
  }
  ctx.strokeStyle = '#ffffff'
  ctx.lineWidth = 2
  ctx.strokeRect(1, 1, w - 2, h - 2)
  const f = baseFont(w, h)
  ctx.font = `${f}px ${MONO}`
  ctx.fillStyle = '#ffffff'
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  ctx.fillText(`${step} px grid`, w / 2, h / 2)
}

function drawAlignment(ctx: Ctx, w: number, h: number) {
  ctx.fillStyle = '#000'
  ctx.fillRect(0, 0, w, h)

  // A one-pixel border. If any edge of this is missing on the wall, the picture
  // is being cropped or the panels are not where the map says they are.
  ctx.strokeStyle = '#ffffff'
  ctx.lineWidth = 1
  ctx.strokeRect(0.5, 0.5, w - 1, h - 1)

  const arm = Math.round(Math.min(w, h) / 8)
  ctx.strokeStyle = '#00ff00'
  ctx.lineWidth = Math.max(2, Math.round(Math.min(w, h) / 300))
  const corners: [number, number, number, number][] = [
    [0, 0, 1, 1],
    [w, 0, -1, 1],
    [0, h, 1, -1],
    [w, h, -1, -1],
  ]
  for (const c of corners) {
    const [cx, cy, sx, sy] = c
    ctx.beginPath()
    ctx.moveTo(cx, cy + sy * arm)
    ctx.lineTo(cx, cy)
    ctx.lineTo(cx + sx * arm, cy)
    ctx.stroke()
  }

  // Centre cross and circles: a circle that is an ellipse is the fastest read
  // there is on a wrong aspect ratio.
  ctx.strokeStyle = '#ffffff'
  ctx.beginPath()
  ctx.moveTo(w / 2, h / 2 - arm)
  ctx.lineTo(w / 2, h / 2 + arm)
  ctx.moveTo(w / 2 - arm, h / 2)
  ctx.lineTo(w / 2 + arm, h / 2)
  ctx.stroke()
  const r = Math.min(w, h) / 2
  for (const frac of [0.4, 0.7, 0.95]) {
    ctx.beginPath()
    ctx.arc(w / 2, h / 2, r * frac, 0, Math.PI * 2)
    ctx.stroke()
  }

  const f = baseFont(w, h)
  ctx.font = `${f}px ${MONO}`
  ctx.fillStyle = '#00ff00'
  ctx.textBaseline = 'top'
  ctx.textAlign = 'left'
  ctx.fillText(`0,0`, arm * 0.2, arm * 0.2)
  ctx.textAlign = 'right'
  ctx.fillText(`${w},0`, w - arm * 0.2, arm * 0.2)
  ctx.textBaseline = 'bottom'
  ctx.textAlign = 'left'
  ctx.fillText(`0,${h}`, arm * 0.2, h - arm * 0.2)
  ctx.textAlign = 'right'
  ctx.fillText(`${w},${h}`, w - arm * 0.2, h - arm * 0.2)
}

function drawBars(ctx: Ctx, w: number, h: number) {
  // 75% bars, descending luminance. Not RP 219 and not labelled as it.
  const bars = ['#bfbfbf', '#bfbf00', '#00bfbf', '#00bf00', '#bf00bf', '#bf0000', '#0000bf', '#000000']
  const bw = w / bars.length
  bars.forEach((c, i) => {
    ctx.fillStyle = c
    // Math.ceil on the width so rounding never leaves a one-pixel black seam
    // between bars, which reads on a wall as a dead column.
    ctx.fillRect(Math.floor(i * bw), 0, Math.ceil(bw) + 1, h)
  })
}

function drawGreyscale(ctx: Ctx, w: number, h: number) {
  const steps = 11
  const sw = w / steps
  for (let i = 0; i < steps; i++) {
    const v = Math.round((255 * i) / (steps - 1))
    ctx.fillStyle = `rgb(${v},${v},${v})`
    ctx.fillRect(Math.floor(i * sw), 0, Math.ceil(sw) + 1, h * 0.7)
  }
  // A continuous ramp under the wedge: steps show banding, the ramp shows where.
  const grad = ctx.createLinearGradient(0, 0, w, 0)
  grad.addColorStop(0, '#000')
  grad.addColorStop(1, '#fff')
  ctx.fillStyle = grad
  ctx.fillRect(0, h * 0.7, w, h * 0.3)
}

function drawPixelCheck(ctx: Ctx, w: number, h: number) {
  const img = ctx.createImageData(w, Math.min(h, 4096))
  const d = img.data
  for (let y = 0; y < img.height; y++) {
    for (let x = 0; x < w; x++) {
      const on = (x + y) % 2 === 0
      const i = (y * w + x) * 4
      const v = on ? 255 : 0
      d[i] = v
      d[i + 1] = v
      d[i + 2] = v
      d[i + 3] = 255
    }
  }
  for (let y = 0; y < h; y += img.height) ctx.putImageData(img, 0, y)

  // A plate saying what a failure means, because this is the one pattern whose
  // failure is not the wall's fault.
  const f = baseFont(w, h)
  const lines = [
    '1 px checkerboard',
    'Flat grey here means the chain is scaling,',
    'not that the wall is faulty.',
  ]
  ctx.font = `${f}px ${MONO}`
  const width = Math.max(...lines.map((l) => ctx.measureText(l).width)) + f
  const height = lines.length * f * 1.4 + f
  const x = (w - width) / 2
  const y = (h - height) / 2
  ctx.fillStyle = 'rgba(0,0,0,0.88)'
  ctx.fillRect(x, y, width, height)
  ctx.strokeStyle = '#ff0000'
  ctx.lineWidth = 2
  ctx.strokeRect(x, y, width, height)
  ctx.fillStyle = '#ffffff'
  ctx.textAlign = 'left'
  ctx.textBaseline = 'top'
  lines.forEach((l, i) => ctx.fillText(l, x + f / 2, y + f / 2 + i * f * 1.4))
}

function drawSafeArea(ctx: Ctx, w: number, h: number, safe: SafeArea) {
  ctx.fillStyle = '#101010'
  ctx.fillRect(0, 0, w, h)
  const x = safe.leftPx
  const y = safe.topPx
  const sw = w - safe.leftPx - safe.rightPx
  const sh = h - safe.topPx - safe.bottomPx

  // The masked region in red, the safe region left dark: what you see is what
  // the audience sees.
  ctx.fillStyle = 'rgba(200,16,46,0.45)'
  ctx.fillRect(0, 0, w, y)
  ctx.fillRect(0, y + sh, w, h - y - sh)
  ctx.fillRect(0, y, x, sh)
  ctx.fillRect(x + sw, y, w - x - sw, sh)

  ctx.strokeStyle = '#00ff00'
  ctx.setLineDash([24, 16])
  ctx.lineWidth = Math.max(2, Math.round(Math.min(w, h) / 300))
  ctx.strokeRect(x, y, sw, sh)
  ctx.setLineDash([])

  const f = baseFont(w, h)
  ctx.font = `${f}px ${MONO}`
  ctx.fillStyle = '#00ff00'
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  ctx.fillText(`safe ${sw} × ${sh}`, x + sw / 2, y + sh / 2)
  ctx.fillText(`of ${w} × ${h}`, x + sw / 2, y + sh / 2 + f * 1.5)
}

const SOLIDS: Record<string, string> = {
  'solid-white': '#ffffff',
  'solid-black': '#000000',
  'solid-red': '#ff0000',
  'solid-green': '#00ff00',
  'solid-blue': '#0000ff',
}

/** The burn-in. Kept out of the corners so a masked edge does not eat it. */
function burnIn(ctx: Ctx, w: number, h: number, lines: string[]) {
  if (!lines.length) return
  const f = baseFont(w, h)
  ctx.font = `${f}px ${MONO}`
  const width = Math.max(...lines.map((l) => ctx.measureText(l).width)) + f
  const height = lines.length * f * 1.35 + f * 0.7
  const x = (w - width) / 2
  const y = h - height - f
  ctx.fillStyle = 'rgba(0,0,0,0.85)'
  ctx.fillRect(x, y, width, height)
  ctx.strokeStyle = '#ffffff'
  ctx.lineWidth = 1
  ctx.strokeRect(x + 0.5, y + 0.5, width - 1, height - 1)
  ctx.fillStyle = '#ffffff'
  ctx.textAlign = 'left'
  ctx.textBaseline = 'top'
  lines.forEach((l, i) => ctx.fillText(l, x + f / 2, y + f * 0.35 + i * f * 1.35))
}

export function renderPattern(args: {
  kind: PatternKind
  wall: Wall
  safeArea: SafeArea
  burnIn: boolean
}): HTMLCanvasElement {
  const { kind, wall, safeArea } = args
  const w = wall.widthPx
  const h = wall.heightPx
  const problem = canvasLimitProblem(w, h)
  if (problem) throw new Error(problem)

  const canvas = document.createElement('canvas')
  canvas.width = w
  canvas.height = h
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Could not get a 2D canvas context.')

  const solid = SOLIDS[kind]
  if (solid) {
    ctx.fillStyle = solid
    ctx.fillRect(0, 0, w, h)
  } else {
    switch (kind) {
      case 'grid':
        drawGrid(ctx, w, h)
        break
      case 'alignment':
        drawAlignment(ctx, w, h)
        break
      case 'bars':
        drawBars(ctx, w, h)
        break
      case 'greyscale':
        drawGreyscale(ctx, w, h)
        break
      case 'pixel-check':
        drawPixelCheck(ctx, w, h)
        break
      case 'safe-area':
        drawSafeArea(ctx, w, h, safeArea)
        break
      default:
        break
    }
  }

  // Never on the solid fields: a burn-in on a uniformity test is a defect in
  // the middle of the thing being measured.
  if (args.burnIn && !solid) {
    burnIn(ctx, w, h, [
      wall.name || 'LED wall',
      `${w} × ${h}`,
      `${patternDef(kind).label} — reads the ${patternDef(kind).reads}`,
    ])
  }
  return canvas
}

export function canvasToPng(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error('The canvas produced no image — the raster is probably over a browser limit.'))
        return
      }
      blob.arrayBuffer().then((b) => resolve(new Uint8Array(b)), reject)
    }, 'image/png')
  })
}
