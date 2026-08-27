/**
 * Turning what the user typed into a finished file.
 *
 * The one piece of real logic here is the ORDER of the pattern slides, and the
 * spec slide's content. Everything else is assembly, and lives apart from
 * `pptx.ts` so that the package writer can be tested without a browser: this
 * module needs a canvas and that one does not.
 */

import type { SlideSpec } from './slides'
import type { UiState } from '../types'
import { buildPptx, deckFileName, type DeckSpec, type SlideContent } from './pptx'
import { canvasToPng, patternDef, renderPattern } from './patterns'
import type { KeynoteSpec } from './keynote'

/**
 * The line that has to survive the file being emailed twice and renamed once.
 *
 * It goes into the package's Comments field, which is what PowerPoint shows in
 * File > Info, and onto the spec slide. A deck built at half size with no record
 * of that fact is a deck that will be exported at 96 dpi and played out at half
 * the wall's resolution, and it will look "a bit soft" to everyone and wrong to
 * nobody until it is too late.
 */
export function exportNote(slide: SlideSpec, wall: { widthPx: number; heightPx: number }): string {
  const target = `${wall.widthPx} × ${wall.heightPx}`
  if (slide.buildScale === 1) {
    return `Slide is at full size. Export at ${trim(slide.dpi)} dpi to land on ${target} px.`
  }
  const n = Math.round(1 / slide.buildScale)
  return (
    `Slide is built at ${slide.buildScale < 1 ? `1/${n}` : `${slide.buildScale}×`} of full size ` +
    `because ${target} px is outside PowerPoint's 1"–56" slide range. ` +
    `Export at ${trim(slide.buildDpi)} dpi (or print at ${Math.round(100 / slide.buildScale)}%) ` +
    `to land back on ${target} px exactly.`
  )
}

function trim(n: number): string {
  const s = n.toFixed(2)
  return s.includes('.') ? s.replace(/\.?0+$/, '') : s
}

function specSlide(state: UiState, slide: SlideSpec, keynote: KeynoteSpec | null): SlideContent {
  const { wall } = state
  const body = [
    `Wall: ${wall.widthPx} × ${wall.heightPx} px`,
    `Slide: ${trim(slide.buildWidthMm / 25.4)}″ × ${trim(slide.buildHeightMm / 25.4)}″ ` +
      `(${trim((slide.buildWidthMm / 25.4) * 72)} × ${trim((slide.buildHeightMm / 25.4) * 72)} pt)`,
    `Export at ${trim(slide.buildDpi)} dpi`,
  ]
  if (slide.buildScale !== 1) {
    body.push(`Built at ${slide.buildScale}× — PowerPoint caps a slide edge at 56″`)
  }
  if (state.safeArea.enabled) {
    const { topPx, rightPx, bottomPx, leftPx } = state.safeArea
    body.push(`Safe area insets: ${topPx} top, ${rightPx} right, ${bottomPx} bottom, ${leftPx} left`)
  }
  if (keynote?.beatsPowerPoint) {
    body.push(`Keynote would hold this wall at full size — its limit is 8192 pt, PowerPoint's is 4032`)
  }
  return { kind: 'heading', title: 'Deck specification', body }
}

export interface BuiltDeck {
  bytes: Uint8Array
  fileName: string
}

/**
 * Build the file. Async only because rendering a pattern to a PNG is.
 *
 * Patterns are rendered one at a time rather than in parallel: each one is a
 * canvas at the wall's full raster, and holding eight 8-megapixel canvases at
 * once is how a browser tab gets killed halfway through a download.
 */
export async function buildDeck(
  state: UiState,
  slide: SlideSpec,
  keynote: KeynoteSpec | null,
): Promise<BuiltDeck> {
  const slides: SlideContent[] = []

  if (state.deck.specSlide) slides.push(specSlide(state, slide, keynote))

  slides.push({
    kind: 'title',
    title: state.wall.name || 'LED wall',
    subtitle: `${state.wall.widthPx} × ${state.wall.heightPx}`,
  })

  for (const kind of state.deck.patterns) {
    const canvas = renderPattern({
      kind,
      wall: state.wall,
      safeArea: state.safeArea,
      burnIn: state.deck.burnIn,
    })
    const bytes = await canvasToPng(canvas)
    // Release the raster before the next one is allocated.
    canvas.width = 0
    canvas.height = 0
    slides.push({
      kind: 'full-bleed',
      image: { bytes, mime: 'image/png' },
      name: patternDef(kind).label,
    })
  }

  const spec: DeckSpec = {
    wall: state.wall,
    slide,
    brand: state.brand,
    safeArea: state.safeArea,
    slides,
    fileKind: state.fileKind,
    exportNote: exportNote(slide, state.wall),
  }

  return { bytes: buildPptx(spec), fileName: deckFileName(state.wall, state.fileKind) }
}

/** Hand the file to the browser. No server, so this is the whole delivery path. */
export function download(bytes: Uint8Array, fileName: string) {
  const blob = new Blob([bytes as unknown as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  a.remove()
  // Revoking immediately races the download in Safari; a tick is enough.
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
