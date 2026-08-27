/**
 * What the user has told the app, and the shapes the generators consume.
 *
 * Colours are stored as SIX HEX DIGITS WITHOUT A LEADING HASH, because that is
 * what `<a:srgbClr val="...">` holds and converting at every use site is how one
 * of them ends up with a stray `#` in the XML. The UI adds the hash for
 * `<input type="color">` and strips it again on the way back in — that boundary
 * is `normaliseHex` in `lib/colour.ts` and it is the only place either form is
 * allowed to cross.
 */

export interface Wall {
  /** What this wall is called on the plan. Ends up in the file name and burn-ins. */
  name: string
  widthPx: number
  heightPx: number
}

export type ImageMime = 'image/png' | 'image/jpeg'

export interface ImageAsset {
  /** Original file name, kept for the part name's extension and for the UI. */
  name: string
  mime: ImageMime
  bytes: Uint8Array
  /** Intrinsic pixel size, needed to place the logo without distorting it. */
  widthPx: number
  heightPx: number
  /** Object URL for the preview. Owned by the UI, which must revoke it. */
  previewUrl: string
}

/**
 * The twelve theme colours, in the order `<a:clrScheme>` requires them.
 *
 * dk1/lt1 and dk2/lt2 are pairs, not four independent colours: the master's
 * `<p:clrMap>` decides which of each pair is "background" and which is "text",
 * and this app maps bg1=lt1 / tx1=dk1 like every Office theme. Swapping the
 * values rather than the mapping is what gives you a deck whose text is
 * invisible in exactly one placeholder.
 */
export interface ThemeColours {
  dk1: string
  lt1: string
  dk2: string
  lt2: string
  accent1: string
  accent2: string
  accent3: string
  accent4: string
  accent5: string
  accent6: string
  hlink: string
  folHlink: string
}

export type LogoCorner = 'none' | 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right'

export interface Brand {
  colours: ThemeColours
  /** Theme font names. Not embedded — see AGENTS.md on why that is deliberate. */
  majorFont: string
  minorFont: string
  logo: ImageAsset | null
  logoCorner: LogoCorner
  /** Logo width as a percentage of slide width. Height follows the aspect ratio. */
  logoWidthPct: number
  background: ImageAsset | null
  /** Slide background when there is no image: theme lt1, or an explicit colour. */
  useSolidBackground: boolean
  backgroundColour: string
}

/**
 * A margin the picture must stay inside, in PIXELS of the wall.
 *
 * Pixels rather than a percentage because the reason for it is physical — the
 * bottom two cabinets are behind the band, the left edge is masked by the
 * proscenium — and those are measured in cabinets, which are measured in pixels.
 * A percentage would have to be recomputed by hand every time the wall changed.
 */
export interface SafeArea {
  enabled: boolean
  topPx: number
  rightPx: number
  bottomPx: number
  leftPx: number
}

export type PatternKind =
  | 'grid'
  | 'alignment'
  | 'bars'
  | 'greyscale'
  | 'pixel-check'
  | 'solid-white'
  | 'solid-black'
  | 'solid-red'
  | 'solid-green'
  | 'solid-blue'
  | 'safe-area'

export interface DeckOptions {
  /** Include the human-readable spec slide as slide 1. */
  specSlide: boolean
  /** Patterns to render, one slide each, in this order. */
  patterns: PatternKind[]
  /** Burn the wall name and resolution into each pattern. */
  burnIn: boolean
}

export type FileKind = 'pptx' | 'potx'

export interface UiState {
  wall: Wall
  brand: Brand
  safeArea: SafeArea
  deck: DeckOptions
  fileKind: FileKind
  /** Export DPI the slide is sized against. 96 unless the user knows better. */
  dpi: number
}
