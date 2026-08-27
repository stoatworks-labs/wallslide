/**
 * Keynote slide sizing.
 *
 * WHY THIS IS NUMBERS AND NOT A FILE
 * ==================================
 * Since Keynote 6 (2013) a `.key` is an IWA package — Snappy-compressed
 * Protocol Buffers in an undocumented schema. There is no writer for it worth
 * depending on and this app does not attempt one. Keynote is served the two
 * ways that actually work:
 *
 *   1. Open the generated `.pptx`. Keynote imports it and KEEPS THE SLIDE SIZE.
 *      Verified on this machine: a deck written at 2880 x 405 pt imported as
 *      2880 x 405 pt with all three slides intact.
 *   2. Set the size by hand in Document Setup, for which this module produces
 *      the two numbers to type.
 *
 * KEYNOTE'S LIMITS ARE NOT POWERPOINT'S, AND THAT MATTERS
 * ======================================================
 * Measured by asking Keynote to accept sizes until it refused. Its own error
 * text: "width values must be between 200 and 8,192" — points, inclusive, and
 * the same on both axes.
 *
 *              min        max
 *   PowerPoint  72 pt      4032 pt     (1" .. 56")
 *   Keynote    200 pt      8192 pt     (2.78" .. 113.78")
 *
 * So Keynote takes a slide MORE THAN TWICE as wide as PowerPoint will. A
 * 7680 px wall at 96 dpi is 5760 pt: PowerPoint must build it at half size and
 * export at 2x, and Keynote simply holds it 1:1. That is a real reason to
 * choose one over the other for a given wall, and it is the sort of thing
 * nobody finds out until the afternoon of the show.
 *
 * The floor runs the other way. Keynote will not go below 200 pt on either
 * axis, so a 1080 x 160 px ticker strip — 810 x 120 pt, perfectly legal in
 * PowerPoint — cannot be a Keynote slide at native size at all.
 */

import type { Problem } from './problem'
import { PT_PER_INCH } from './slides'

/** Measured from Keynote itself, not from documentation. Points, inclusive. */
export const KEYNOTE_MIN_PT = 200
export const KEYNOTE_MAX_PT = 8192

export interface KeynoteSpec {
  /** The slide at native size: one slide point per target pixel at `dpi`. */
  widthPt: number
  heightPt: number
  /** Whole-number scale that brings both edges into Keynote's range. */
  buildScale: number
  buildWidthPt: number
  buildHeightPt: number
  /** True when Keynote can hold the wall at 1:1 and PowerPoint cannot. */
  beatsPowerPoint: boolean
  problems: Problem[]
}

/**
 * The scale bringing both edges inside Keynote's range, preferring 1.
 *
 * Same shape as `slides.ts`'s `fitScale` and same reasoning — whole numbers and
 * whole reciprocals only, because "build at half and export at 200%" is an
 * instruction a person can follow at 2 a.m. and "build at 1/2.37" is not. The
 * interval is empty exactly when the slide is steeper than 8192:200, or 40.96:1.
 */
function fitScale(wPt: number, hPt: number): number | null {
  const longest = Math.max(wPt, hPt)
  const shortest = Math.min(wPt, hPt)
  const lower = KEYNOTE_MIN_PT / shortest
  const upper = KEYNOTE_MAX_PT / longest
  if (lower > upper) return null
  if (lower <= 1 && 1 <= upper) return 1
  if (upper < 1) {
    const n = Math.ceil(1 / upper)
    return 1 / n >= lower ? 1 / n : null
  }
  const m = Math.ceil(lower)
  return m <= upper ? m : null
}

function trim(n: number): string {
  const s = n.toFixed(2)
  return s.includes('.') ? s.replace(/\.?0+$/, '') : s
}

export function keynoteFromResolution(
  hPixels: number,
  vPixels: number,
  dpi: number,
  /** PowerPoint's own build scale, so the two can be compared in one sentence. */
  pptBuildScale: number,
): KeynoteSpec | null {
  if (!(hPixels > 0) || !(vPixels > 0) || !(dpi > 0)) return null

  const widthPt = (hPixels / dpi) * PT_PER_INCH
  const heightPt = (vPixels / dpi) * PT_PER_INCH
  const problems: Problem[] = []

  const scale = fitScale(widthPt, heightPt)
  const buildScale = scale ?? 1
  const buildWidthPt = widthPt * buildScale
  const buildHeightPt = heightPt * buildScale

  if (scale === null) {
    problems.push({
      level: 'error',
      text: `${trim(widthPt)} x ${trim(heightPt)} pt is steeper than 40.96:1, and Keynote caps an edge at ${KEYNOTE_MAX_PT} pt while requiring at least ${KEYNOTE_MIN_PT}. No scale satisfies both, so this shape cannot be a Keynote slide.`,
    })
  } else if (buildScale < 1) {
    const n = Math.round(1 / buildScale)
    problems.push({
      level: 'warn',
      text: `Over Keynote's ${KEYNOTE_MAX_PT} pt limit. Set Document Setup to ${trim(buildWidthPt)} x ${trim(buildHeightPt)} pt — one ${n === 2 ? 'half' : `${n}th`} of full size — and export at ${n}x.`,
    })
  } else if (buildScale > 1) {
    problems.push({
      level: 'warn',
      text: `Under Keynote's ${KEYNOTE_MIN_PT} pt minimum on at least one edge — a narrow strip is more constrained in Keynote than in PowerPoint, whose floor is 72 pt. Set Document Setup to ${trim(buildWidthPt)} x ${trim(buildHeightPt)} pt, ${buildScale}x full size, and export at 1/${buildScale}.`,
    })
  }

  // The headline comparison, and the only reason this module reports on
  // PowerPoint at all: when Keynote holds the wall 1:1 and PowerPoint cannot,
  // the choice of application is a real decision rather than a preference.
  const beatsPowerPoint = buildScale === 1 && pptBuildScale !== 1
  if (beatsPowerPoint) {
    problems.push({
      level: 'info',
      text: `Keynote takes this wall at full size where PowerPoint cannot: its slide limit is ${KEYNOTE_MAX_PT} pt against PowerPoint's 4032. If the deck can be built in Keynote, it needs no build scale and no export multiplier at all.`,
    })
  }

  return {
    widthPt,
    heightPt,
    buildScale,
    buildWidthPt,
    buildHeightPt,
    beatsPowerPoint,
    problems,
  }
}

/** Document Setup takes points, and this is the text to type into it. */
export function keynoteFieldText(pt: number): string {
  // Keynote's own field rounds to whole points, so offering more is a lie about
  // the precision available.
  return String(Math.round(pt))
}
