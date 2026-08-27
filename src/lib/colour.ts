/**
 * Hex colour handling and the contrast check.
 *
 * The internal form is SIX UPPERCASE HEX DIGITS, NO HASH — `<a:srgbClr val="">`
 * takes exactly that, and keeping the hash out of the stored value means no code
 * path can put one into the XML. `withHash` is the only way back to CSS.
 */

/** Six uppercase hex digits, no hash. Null when the input is not a colour. */
export function normaliseHex(input: string): string | null {
  const s = input.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}$/.test(s)) return s.toUpperCase()
  // Three-digit shorthand is what people type by hand, and rejecting it as
  // "not a colour" is a baffling thing for a form to do.
  if (/^[0-9a-fA-F]{3}$/.test(s)) {
    return s
      .toUpperCase()
      .split('')
      .map((c) => c + c)
      .join('')
  }
  return null
}

export function withHash(hex: string): string {
  return `#${hex}`
}

function channel(hex: string, i: number): number {
  return parseInt(hex.slice(i, i + 2), 16) / 255
}

/**
 * WCAG relative luminance. sRGB in, 0..1 out.
 *
 * This is the perceptual weighting, not a plain average — a saturated blue and a
 * saturated yellow have wildly different luminance at the same "brightness", and
 * an average would call both of them mid-grey.
 */
export function relativeLuminance(hex: string): number {
  const lin = (c: number) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4)
  return 0.2126 * lin(channel(hex, 0)) + 0.7152 * lin(channel(hex, 2)) + 0.0722 * lin(channel(hex, 4))
}

/** WCAG contrast ratio, 1..21. Order of the arguments does not matter. */
export function contrastRatio(a: string, b: string): number {
  const la = relativeLuminance(a)
  const lb = relativeLuminance(b)
  const [hi, lo] = la > lb ? [la, lb] : [lb, la]
  return (hi + 0.05) / (lo + 0.05)
}

/**
 * Whichever of black or white is more readable on `hex`.
 *
 * Used to pick the burn-in colour over a brand background, and to pick a default
 * body-text colour when the user sets a dark background and nothing else.
 */
export function readableOn(hex: string): string {
  return contrastRatio(hex, 'FFFFFF') >= contrastRatio(hex, '000000') ? 'FFFFFF' : '000000'
}

/** CSS `rgb()` — canvas wants this, and it never needs the alpha form. */
export function cssRgb(hex: string): string {
  return `#${hex}`
}
