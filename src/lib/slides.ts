/**
 * VENDORED from aspect-calc (`src/lib/slides.ts`).
 *
 * The only edit is the `Problem` import, which points at this app's `problem.ts`
 * rather than aspect-calc's `solve.ts`. Everything else is byte-identical and
 * should stay that way — the 56-inch cap, the whole-fraction build scale and the
 * 100 MP export ceiling are pinned by 100 tests over there. Fix bugs in
 * aspect-calc and re-copy.
 *
 * Wallslide uses it in one direction only (`slideFromResolution`), but the
 * module came across whole rather than filleted: a partial copy diverges from
 * its origin faster than a complete one, and the unused half costs nothing.
 */
/**
 * PowerPoint slide geometry.
 *
 * A slide is a display like any other and this is `solve.ts` again with the
 * pitch written the other way up:
 *
 *     slide size (in)  x  export DPI  =  resolution (px)
 *
 * `dpi = 25.4 / pitch_mm`, so a slide exported at PowerPoint's default 96 dpi
 * is a panel of 0.2646 mm pitch. It would be tempting to fold this in as a
 * fourth mode of the main solver. It lives apart because PowerPoint adds three
 * constraints an LED wall does not have, and those constraints are the entire
 * reason anyone needs this calculation done for them:
 *
 *   1. NO SLIDE EDGE MAY EXCEED 56 INCHES. A 7680 px wide wall at 96 dpi wants
 *      an 80 inch slide and PowerPoint simply refuses. The answer is to build
 *      at half size and export at 2x, which is arithmetic PowerPoint will not
 *      do for you and which people get wrong in the other direction.
 *   2. NO EDGE MAY BE UNDER 1 INCH either, which bites on ticker strips.
 *   3. NO EXPORTED BITMAP MAY EXCEED 100 MEGAPIXELS, so the export DPI has its
 *      own ceiling of sqrt(1e8 / (w x h)) with the slide measured in inches.
 *
 * Everything here is millimetres on the way in and out, same as the rest of the
 * app. Inches only appear because PowerPoint's own limits are stated in them.
 */

import type { Problem } from './problem'
import { MM_PER_INCH } from './units'

/** Hard limits from PowerPoint's Slide Size dialog. 142.24 cm and 2.54 cm. */
export const PPT_MAX_IN = 56
export const PPT_MIN_IN = 1
export const PPT_MAX_MM = PPT_MAX_IN * MM_PER_INCH
export const PPT_MIN_MM = PPT_MIN_IN * MM_PER_INCH

/** PowerPoint renders and exports at 96 dpi unless the registry says otherwise. */
export const DEFAULT_DPI = 96

/** English Metric Units — what a .pptx actually stores in `<p:sldSz>`. */
export const EMU_PER_INCH = 914400
export const PT_PER_INCH = 72

/** PowerPoint will not write a bitmap bigger than this, whatever the DPI. */
export const EXPORT_PIXEL_CAP = 100_000_000

/**
 * `<p:sldSz cx="12192000" cy="6858000"/>` — the Widescreen default.
 *
 * Taken from the EMU rather than the dialog's rounded "13.333 in" because the
 * EMU figure is the real one and is exactly 16:9. The dialog's decimal is not.
 */
export const DECK_WIDESCREEN_EMU = { cx: 12192000, cy: 6858000 }
export const DEFAULT_DECK_WIDTH_IN = DECK_WIDESCREEN_EMU.cx / EMU_PER_INCH
export const DEFAULT_DECK_HEIGHT_IN = DECK_WIDESCREEN_EMU.cy / EMU_PER_INCH

export interface SlideSpec {
  /** The slide at NATIVE size — one slide pixel per target pixel at `dpi`. */
  widthMm: number
  heightMm: number
  hPixels: number
  vPixels: number
  /** Unrounded pixel counts, present only when the pixels were derived. */
  rawHPixels: number | null
  rawVPixels: number | null
  dpi: number

  /**
   * Multiply the native size by this to get a slide PowerPoint will accept.
   * 1 when it already fits, 0.5 when it has to be halved, 2 when it is too
   * small to be a slide at all. Whole numbers and whole reciprocals only:
   * "build at half and export at 200%" is an instruction a person can follow
   * on site. "build at 1/2.37" is not.
   */
  buildScale: number
  buildWidthMm: number
  buildHeightMm: number
  /** Export DPI that turns the built slide back into the native pixel count. */
  buildDpi: number
  /** PowerPoint's own ceiling for a slide this size, from the 100 MP bitmap cap. */
  maxExportDpi: number

  /**
   * Built slide width against the 13.333 in Widescreen default. Every type size
   * and every template measurement scales by this, and it is the number that
   * decides whether a native-size slide is a good idea or a miserable one.
   */
  typeScale: number

  /**
   * When the target ratio matches Widescreen you can leave the deck alone and
   * raise the export DPI instead, which keeps every template, every font size
   * and every bit of muscle memory intact. Null when the ratio differs, and
   * null when the required DPI is not a whole number — `ExportBitmapResolution`
   * is a DWORD and there is no way to ask for 287.7 dpi.
   */
  standardDeckDpi: number | null

  problems: Problem[]
}

const pos = (n: number | null | undefined): n is number =>
  typeof n === 'number' && Number.isFinite(n) && n > 0

/** Trim a computed length for prose, where 40.0000001 in helps nobody. */
function inches(mm: number): number {
  return mm / MM_PER_INCH
}

function fmtIn(mm: number): string {
  const v = inches(mm)
  const s = v.toFixed(3)
  return `${s.includes('.') ? s.replace(/\.?0+$/, '') : s}"`
}

/**
 * The scale that brings both edges inside PowerPoint's range, preferring 1.
 *
 * The constraint is an interval: `s x longest <= 56` and `s x shortest >= 1`,
 * so `s` lives in [1/shortest, 56/longest]. That interval is empty exactly when
 * the slide is steeper than 56:1, which no amount of scaling fixes.
 */
function fitScale(wIn: number, hIn: number): number | null {
  const longest = Math.max(wIn, hIn)
  const shortest = Math.min(wIn, hIn)
  const lower = PPT_MIN_IN / shortest
  const upper = PPT_MAX_IN / longest
  if (lower > upper) return null
  if (lower <= 1 && 1 <= upper) return 1
  // Too big: halve, third, quarter. Too small: double, triple.
  if (upper < 1) {
    const n = Math.ceil(1 / upper)
    return 1 / n >= lower ? 1 / n : null
  }
  const m = Math.ceil(lower)
  return m <= upper ? m : null
}

function describe(args: {
  widthMm: number
  heightMm: number
  hPixels: number
  vPixels: number
  rawHPixels: number | null
  rawVPixels: number | null
  dpi: number
}): SlideSpec {
  const { widthMm, heightMm, hPixels, vPixels, dpi } = args
  const problems: Problem[] = []

  const wIn = inches(widthMm)
  const hIn = inches(heightMm)

  const scale = fitScale(wIn, hIn)
  const buildScale = scale ?? 1
  const buildWidthMm = widthMm * buildScale
  const buildHeightMm = heightMm * buildScale
  const buildDpi = dpi / buildScale

  const maxExportDpi = Math.floor(
    Math.sqrt(EXPORT_PIXEL_CAP / (inches(buildWidthMm) * inches(buildHeightMm))),
  )

  if (scale === null) {
    problems.push({
      level: 'error',
      text: `${fmtIn(widthMm)} x ${fmtIn(heightMm)} is steeper than 56:1, and PowerPoint caps an edge at 56" while requiring at least 1". No single scale satisfies both — this shape cannot be a slide. Split it across two slides, or drive it from something that is not PowerPoint.`,
    })
  } else if (buildScale < 1) {
    const n = Math.round(1 / buildScale)
    problems.push({
      level: 'warn',
      text: `${fmtIn(widthMm)} x ${fmtIn(heightMm)} is over PowerPoint's 56" limit. Build the slide at ${fmtIn(buildWidthMm)} x ${fmtIn(buildHeightMm)} — ${fraction(n)} of full size — and export at ${trimNum(buildDpi)} dpi, or print at ${n * 100}%. That lands back on ${hPixels.toLocaleString('en-GB')} x ${vPixels.toLocaleString('en-GB')} px exactly.`,
    })
  } else if (buildScale > 1) {
    problems.push({
      level: 'warn',
      text: `${fmtIn(widthMm)} x ${fmtIn(heightMm)} is under PowerPoint's 1" minimum on at least one edge. Build the slide at ${fmtIn(buildWidthMm)} x ${fmtIn(buildHeightMm)} — ${buildScale}x full size — and export at ${trimNum(buildDpi)} dpi to land back on ${hPixels.toLocaleString('en-GB')} x ${vPixels.toLocaleString('en-GB')} px.`,
    })
  }

  // A pixel count is an integer or it is fiction — same rule as solve.ts. Only
  // reachable from `resolutionFromSlide`; the other direction is exact by
  // construction because the size is computed from the pixels.
  if (args.rawHPixels != null && args.rawVPixels != null) {
    const dh = Math.abs(args.rawHPixels - hPixels)
    const dv = Math.abs(args.rawVPixels - vPixels)
    // A QUARTER PIXEL, not solve.ts's 0.005. There the input is a tape measure
    // and any rounding at all is worth reporting; here the input is a number
    // typed into a dialog that only shows three decimals, so a few thousandths
    // of a pixel is the dialog's own rounding coming back and saying so is
    // noise. 13.333" lands 0.03 px off 1280 and that is not news.
    if (dh > 0.25 || dv > 0.25) {
      problems.push({
        level: 'info',
        text: `${trimNum(inches(widthMm))}" x ${trimNum(inches(heightMm))}" is ${args.rawHPixels.toFixed(2)} x ${args.rawVPixels.toFixed(2)} px at ${trimNum(dpi)} dpi, so the export lands on ${hPixels} x ${vPixels}. A slide size typed in whole centimetres rarely gives a pixel count anyone would have chosen — go the other way round and let the slide size fall out of the resolution.`,
      })
    }
  }

  if (hPixels * vPixels > EXPORT_PIXEL_CAP) {
    problems.push({
      level: 'warn',
      text: `${(hPixels * vPixels / 1e6).toFixed(1)} MP is past the 100 MP ceiling on a PowerPoint bitmap export, so no export DPI reaches this. ${maxExportDpi} dpi is the most this slide will give you. Export in sections, or render the deck somewhere else.`,
    })
  }

  const typeScale = inches(buildWidthMm) / DEFAULT_DECK_WIDTH_IN

  // Only offer the standard-deck route when the ratio genuinely matches and the
  // DPI comes out whole. ExportBitmapResolution is a DWORD; 287.7 is not on
  // offer and a rounded 288 would quietly give you the wrong pixel count.
  const targetRatio = hPixels / vPixels
  const deckRatio = DEFAULT_DECK_WIDTH_IN / DEFAULT_DECK_HEIGHT_IN
  const ratioOff = Math.abs(targetRatio - deckRatio) / deckRatio
  const deckDpi = hPixels / DEFAULT_DECK_WIDTH_IN

  // Nothing to suggest when the slide already IS the Widescreen deck. The
  // tolerance is a thousandth of an inch, which is finer than the Slide Size
  // dialog will even display, so a deck typed as 13.333 still counts as one.
  const alreadyStandard =
    Math.abs(inches(buildWidthMm) - DEFAULT_DECK_WIDTH_IN) < 0.001 &&
    Math.abs(inches(buildHeightMm) - DEFAULT_DECK_HEIGHT_IN) < 0.001

  const standardDeckDpi =
    !alreadyStandard &&
    ratioOff <= 0.005 &&
    Math.abs(deckDpi - Math.round(deckDpi)) < 1e-6 &&
    Math.round(deckDpi) >= 1 &&
    Math.round(deckDpi) <=
      Math.floor(Math.sqrt(EXPORT_PIXEL_CAP / (DEFAULT_DECK_WIDTH_IN * DEFAULT_DECK_HEIGHT_IN)))
      ? Math.round(deckDpi)
      : null

  return {
    widthMm,
    heightMm,
    hPixels,
    vPixels,
    rawHPixels: args.rawHPixels,
    rawVPixels: args.rawVPixels,
    dpi,
    buildScale,
    buildWidthMm,
    buildHeightMm,
    buildDpi,
    maxExportDpi,
    typeScale,
    standardDeckDpi,
    problems,
  }
}

function trimNum(n: number): string {
  const s = n.toFixed(2)
  return s.includes('.') ? s.replace(/\.?0+$/, '') : s
}

/** "one third", not "one 3th". Only ever called with small whole divisors. */
function fraction(n: number): string {
  const NAMES: Record<number, string> = { 2: 'one half', 3: 'one third', 4: 'one quarter' }
  return NAMES[n] ?? `one ${n}th`
}

/** Slide size for a target resolution. The size is exact — pixels are the input. */
export function slideFromResolution(
  hPixels: number | null,
  vPixels: number | null,
  dpi: number | null,
): SlideSpec | null {
  if (!pos(hPixels) || !pos(vPixels) || !pos(dpi)) return null
  return describe({
    widthMm: (hPixels / dpi) * MM_PER_INCH,
    heightMm: (vPixels / dpi) * MM_PER_INCH,
    hPixels,
    vPixels,
    rawHPixels: null,
    rawVPixels: null,
    dpi,
  })
}

/** Resolution of a given slide size. Pixels round; the rounding is reported. */
export function resolutionFromSlide(
  widthMm: number | null,
  heightMm: number | null,
  dpi: number | null,
): SlideSpec | null {
  if (!pos(widthMm) || !pos(heightMm) || !pos(dpi)) return null
  const rawH = inches(widthMm) * dpi
  const rawV = inches(heightMm) * dpi
  return describe({
    widthMm,
    heightMm,
    hPixels: Math.max(1, Math.round(rawH)),
    vPixels: Math.max(1, Math.round(rawV)),
    rawHPixels: rawH,
    rawVPixels: rawV,
    dpi,
  })
}

/** Points, which is what a .pptx measures type and shape positions in. */
export function toPoints(mm: number): number {
  return inches(mm) * PT_PER_INCH
}

/** EMU, which is what `<p:sldSz>` actually holds. Whole numbers, always. */
export function toEmu(mm: number): number {
  return Math.round(inches(mm) * EMU_PER_INCH)
}

export interface SlidePreset {
  label: string
  /**
   * Inches, EXACT — not the three decimals PowerPoint's dialog displays.
   *
   * Widescreen is 40/3 in, and the dialog's "13.333" is a different slide:
   * 12,191,695 EMU against the real 12,192,000. Nobody would ever see that
   * difference on a screen, but the tool reports EMU and reporting a near-miss
   * of the canonical value as though it were the canonical value is exactly the
   * sort of quiet lie the rest of this app refuses to tell. `slideFieldText`
   * writes enough places to land back on the right EMU.
   */
  wIn: number
  hIn: number
  note?: string
}

/**
 * A preset dimension as text for an input box: six decimals, trimmed.
 *
 * Six because that is what it takes for 40/3 to round-trip through EMU —
 * 13.333333 x 914400 = 12,191,999.7, which rounds to 12,192,000 exactly. Five
 * places does not.
 */
export function slideFieldText(inchesValue: number): string {
  const s = inchesValue.toFixed(6)
  return s.includes('.') ? s.replace(/\.?0+$/, '') : s
}

/** The friendly three-decimal form PowerPoint's own dialog shows. */
export function slideLabelText(inchesValue: number): string {
  const s = inchesValue.toFixed(3)
  return s.includes('.') ? s.replace(/\.?0+$/, '') : s
}

export interface SlidePresetGroup {
  group: string
  items: SlidePreset[]
}

/**
 * PowerPoint's "Slides sized for" list, verbatim.
 *
 * The pair worth staring at is the two 16:9 entries. Widescreen is 13.333 x 7.5
 * and On-screen Show (16:9) is 10 x 5.625. Both are exactly 16:9 and they are a
 * third apart in absolute size, so a deck built in one and resized to the other
 * has every point size on every slide wrong by 33%. This is the whole reason
 * the tool reports a size and never just a ratio.
 */
export const SLIDE_PRESETS: SlidePresetGroup[] = [
  {
    group: 'PowerPoint',
    items: [
      {
        label: 'Widescreen',
        wIn: DEFAULT_DECK_WIDTH_IN,
        hIn: DEFAULT_DECK_HEIGHT_IN,
        note: 'The modern default: exactly 16:9, 960 × 540 pt, 12192000 × 6858000 EMU. The dialog rounds the width to 13.333″ but stores 40/3.',
      },
      {
        label: 'On-screen Show (16:9)',
        wIn: 10,
        hIn: 5.625,
        note: 'Also exactly 16:9, but three quarters the size of Widescreen — the same deck at the same point sizes looks a third bigger here.',
      },
      { label: 'On-screen Show (4:3)', wIn: 10, hIn: 7.5 },
      { label: 'On-screen Show (16:10)', wIn: 10, hIn: 6.25 },
      { label: 'A4 Paper', wIn: 10.833, hIn: 7.5 },
      { label: 'A3 Paper', wIn: 14, hIn: 10.5 },
      { label: 'Letter Paper', wIn: 10, hIn: 7.5 },
      { label: 'Ledger Paper', wIn: 13.319, hIn: 9.99 },
      { label: 'B4 (ISO) Paper', wIn: 11.84, hIn: 8.88 },
      { label: 'B5 (ISO) Paper', wIn: 7.84, hIn: 5.88 },
      { label: '35 mm Slides', wIn: 11.25, hIn: 7.5 },
      { label: 'Overhead', wIn: 10, hIn: 7.5 },
      { label: 'Banner', wIn: 8, hIn: 1, note: 'Sits exactly on the 1" minimum edge.' },
    ],
  },
  {
    group: 'Google Slides',
    items: [
      {
        label: 'Widescreen 16:9',
        wIn: 10,
        hIn: 5.625,
        note: "Google's default matches PowerPoint's On-screen Show (16:9), not Widescreen. A deck round-tripped between the two keeps its ratio and loses its type scale.",
      },
      { label: 'Standard 4:3', wIn: 10, hIn: 7.5 },
    ],
  },
]

/** Export DPIs worth offering. 96 is the default; the rest are ExportBitmapResolution values. */
export const DPI_PRESETS = [96, 120, 144, 150, 192, 200, 240, 288, 300, 384, 600]
