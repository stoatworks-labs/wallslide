/**
 * `ppt/theme/theme1.xml` — the part that makes a template a template.
 *
 * A theme is four lists: colours, fonts, and the format scheme's fill/line/
 * effect/background styles. Everything on every slide that is not explicitly
 * overridden resolves through here, which is why editing the theme is the whole
 * of "let the event organiser brand it" and why editing individual slides is not.
 *
 * TWO THINGS ARE DELIBERATELY UNLIKE AN OFFICE-GENERATED THEME
 * ============================================================
 *
 * 1. NO SCRIPT FALLBACK TABLE. Office writes about forty `<a:font script="...">`
 *    entries into every fontScheme — Japanese, Devanagari, Khmer, Syriac and so
 *    on — whether or not the deck contains a character of any of them. That
 *    table is exactly what makes pptx-font-manager's job hard: searching a real
 *    deck for `typeface=` finds 39 fonts where it uses two. A theme generated
 *    here names the fonts it actually uses and nothing else, so a tech opening
 *    it in that tool gets a two-line answer.
 *
 *    The cost is real and worth stating: a deck with no fallback table renders
 *    CJK and Indic text in whatever the platform picks rather than in Office's
 *    curated choice. For an event title slide in Latin script that is no cost at
 *    all. For a deck that will carry Japanese, it is, and the answer there is to
 *    set the theme fonts to a face that covers it.
 *
 * 2. NO EFFECTS. All three `<a:effectStyle>` entries are empty. Office's ship
 *    with soft shadows, and a soft shadow on an LED wall is a grey smear at
 *    2.6 mm pitch and a source of banding on a low-bit-depth processor. A
 *    template that quietly adds one to every shape is a template that has to be
 *    fought. If somebody wants a shadow they can add one.
 */

import type { ThemeColours } from '../types'
import { XML_DECL, text } from './xml'

/**
 * Microsoft's Office theme colours, used as the starting point in the UI.
 *
 * Not chosen for looks — they are here because they are what every stock deck
 * already uses, so a template seeded with them looks *identical* to an untouched
 * PowerPoint until the organiser changes something, and nothing surprises them.
 */
export const OFFICE_COLOURS: ThemeColours = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
}

/** The order `<a:clrScheme>` requires. Not alphabetical, and not negotiable. */
const CLR_ORDER: (keyof ThemeColours)[] = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]

function clrScheme(name: string, c: ThemeColours): string {
  const entries = CLR_ORDER.map(
    (k) => `<a:${k}><a:srgbClr val="${c[k]}"/></a:${k}>`,
  ).join('')
  return `<a:clrScheme name="${text(name)}">${entries}</a:clrScheme>`
}

function fontScheme(name: string, major: string, minor: string): string {
  // Empty `ea` and `cs` rather than absent: the elements are required by the
  // schema, and an empty typeface means "no preference", which is the honest
  // answer. See the header note on the fallback table.
  const face = (t: string) =>
    `<a:latin typeface="${text(t)}"/><a:ea typeface=""/><a:cs typeface=""/>`
  return (
    `<a:fontScheme name="${text(name)}">` +
    `<a:majorFont>${face(major)}</a:majorFont>` +
    `<a:minorFont>${face(minor)}</a:minorFont>` +
    `</a:fontScheme>`
  )
}

/**
 * The format scheme. Three entries in each list, which the schema requires.
 *
 * `phClr` is the placeholder colour — whatever colour the shape is actually
 * using resolves into these templates. The three fills are subtle / moderate /
 * intense, and all three here are flat: Office's moderate and intense are
 * gradients, and a gradient across a 7 m wall at 8 bits per channel is visible
 * banding. Flat is not a limitation, it is the correct answer for this medium.
 */
function fmtScheme(name: string): string {
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
  const tinted = (lumMod: number, lumOff: number) =>
    `<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="${lumMod}"/><a:lumOff val="${lumOff}"/></a:schemeClr></a:solidFill>`

  const line = (w: number) =>
    `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`

  const noEffect = '<a:effectStyle><a:effectLst/></a:effectStyle>'

  return (
    `<a:fmtScheme name="${text(name)}">` +
    `<a:fillStyleLst>${solid}${tinted(110000, 0)}${tinted(90000, 0)}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    `<a:effectStyleLst>${noEffect}${noEffect}${noEffect}</a:effectStyleLst>` +
    `<a:bgFillStyleLst>${solid}${tinted(95000, 5000)}${tinted(105000, 0)}</a:bgFillStyleLst>` +
    `</a:fmtScheme>`
  )
}

export function themeXml(args: {
  name: string
  colours: ThemeColours
  majorFont: string
  minorFont: string
}): string {
  const { name, colours, majorFont, minorFont } = args
  return (
    XML_DECL +
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${text(name)}">` +
    `<a:themeElements>` +
    clrScheme(name, colours) +
    fontScheme(name, majorFont, minorFont) +
    fmtScheme(name) +
    `</a:themeElements>` +
    // Both required by the schema, both legitimately empty: no per-object
    // defaults and no alternate colour schemes.
    `<a:objectDefaults/>` +
    `<a:extraClrSchemeLst/>` +
    `</a:theme>`
  )
}
