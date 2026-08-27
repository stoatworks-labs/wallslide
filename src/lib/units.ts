/**
 * Length conversion. Millimetres in, millimetres out.
 *
 * Vendored down from aspect-calc's `units.ts`, which also parses feet-and-inches
 * input. Nothing here takes a length from a person — every millimetre in this
 * app is derived from a pixel count and a DPI — so the parser did not come with
 * it. Inches exist only because PowerPoint's own limits are stated in them.
 */

export const MM_PER_INCH = 25.4

export function mmToIn(mm: number): number {
  return mm / MM_PER_INCH
}

export function inToMm(inches: number): number {
  return inches * MM_PER_INCH
}
