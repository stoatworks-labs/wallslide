/**
 * The .pptx / .potx writer.
 *
 * A PresentationML package is a ZIP of XML parts wired together by relationship
 * files. Nothing here is clever; what it has to be is *exactly right*, because
 * the failure mode for a malformed package is PowerPoint's "found a problem with
 * content" dialog, which names no part, no element and no reason.
 *
 * THE PART GRAPH THIS WRITES
 * ==========================
 *
 *   [Content_Types].xml          every part's MIME type, by extension or by name
 *   _rels/.rels                  -> docProps/core.xml, ppt/presentation.xml
 *   docProps/core.xml            title, and the export instructions as Comments
 *   ppt/presentation.xml         <p:sldSz> lives here. THE POINT OF THE APP
 *   ppt/_rels/presentation.xml.rels  -> master, slides, presProps, theme
 *   ppt/presProps.xml            empty, but PowerPoint expects the relationship
 *   ppt/theme/theme1.xml         colours + fonts (theme.ts)
 *   ppt/slideMasters/slideMaster1.xml   background, logo, text styles, clrMap
 *   ppt/slideLayouts/slideLayout{1..4}.xml
 *   ppt/slides/slide{1..n}.xml
 *   ppt/media/image{1..m}.{png,jpeg}
 *
 * Deliberately absent, having confirmed against a real Office-opened package
 * that none of them is required: `viewProps.xml`, `tableStyles.xml`,
 * `docProps/app.xml`, and any notes master or notes slides. Each one is a part
 * that could be malformed, and none earns its risk here.
 *
 * `viewProps.xml` is the one worth naming, because it is where PowerPoint's
 * drawing guides live and guides are the obvious way to mark a safe area. They
 * are not written, because `<p:guide pos="">` is in eighth-points and this app
 * has no PowerPoint to confirm that against — a safe-area marker in the wrong
 * place is worse than no marker at all. The safe area is expressed the two ways
 * that cannot be wrong instead: every placeholder is positioned inside it, and
 * the test deck can carry a slide that draws it.
 *
 * ID RULES THAT ARE NOT OPTIONAL
 * ==============================
 *   <p:sldMasterId id>  >= 2147483648
 *   <p:sldLayoutId id>  >= 2147483648, unique
 *   <p:sldId id>        256 .. 2147483647
 *   <p:cNvPr id>        unique within a slide's shape tree, and never 0
 */

import type { SlideSpec } from './slides'
import { toEmu } from './slides'
import type { Brand, FileKind, ImageMime, SafeArea, ThemeColours, Wall } from '../types'
import { readableOn } from './colour'
import { themeXml } from './theme'
import { XML_DECL, relsPart, rel, text, utf8 } from './xml'
import { buildZip, type ZipEntry } from './zip'

// ---------------------------------------------------------------------------
// Relationship type URIs. Spelled out, never abbreviated behind a short name:
// a wrong one of these is the single most common cause of a package that opens
// to an error message with nothing in it.
// ---------------------------------------------------------------------------
const OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const SLIDE_MASTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const SLIDE_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
const PRES_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
const IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

/** Namespace declarations for the presentationml part roots. */
const NS_PML =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface DeckImage {
  bytes: Uint8Array
  mime: ImageMime
}

/** One slide's worth of content. Layout choice follows from the kind. */
export type SlideContent =
  | { kind: 'title'; title: string; subtitle: string }
  | { kind: 'heading'; title: string; body: string[] }
  | { kind: 'blank' }
  /**
   * A full-bleed picture with the master's own shapes suppressed.
   *
   * `showMasterSp="0"` is the important half: a test pattern with the event
   * logo sitting on top of it is not a test pattern, it is a picture of a logo.
   */
  | { kind: 'full-bleed'; image: DeckImage; name: string }

export interface DeckSpec {
  wall: Wall
  slide: SlideSpec
  brand: Brand
  safeArea: SafeArea
  slides: SlideContent[]
  fileKind: FileKind
  /** Shown in File > Info, and the only place the export DPI survives the file. */
  exportNote: string
  /** Fixed so a rebuild of the same deck is byte-identical. Tests depend on it. */
  now?: Date
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

export interface Box {
  x: number
  y: number
  cx: number
  cy: number
}

/**
 * The rectangle a slide's content must stay inside, in EMU.
 *
 * The safe area is given in WALL pixels and the slide is in EMU at BUILD size,
 * so the conversion runs through the whole-slide ratio rather than through the
 * DPI. That way it stays correct when `buildScale` is a half or a third: the
 * slide is a scale model of the wall, and a proportion of it is the same
 * proportion whatever the model's scale.
 */
export function safeBox(slideCx: number, slideCy: number, wall: Wall, safe: SafeArea): Box {
  if (!safe.enabled) return { x: 0, y: 0, cx: slideCx, cy: slideCy }
  const px = slideCx / wall.widthPx
  const py = slideCy / wall.heightPx
  const left = Math.max(0, safe.leftPx) * px
  const right = Math.max(0, safe.rightPx) * px
  const top = Math.max(0, safe.topPx) * py
  const bottom = Math.max(0, safe.bottomPx) * py
  const cx = slideCx - left - right
  const cy = slideCy - top - bottom
  // A safe area that swallows the slide is a typo, not an instruction. Fall back
  // to the full slide rather than emitting a negative <a:ext>, which PowerPoint
  // rejects outright.
  if (cx <= 0 || cy <= 0) return { x: 0, y: 0, cx: slideCx, cy: slideCy }
  return { x: Math.round(left), y: Math.round(top), cx: Math.round(cx), cy: Math.round(cy) }
}

function xfrm(b: Box): string {
  return `<a:xfrm><a:off x="${Math.round(b.x)}" y="${Math.round(b.y)}"/><a:ext cx="${Math.round(
    b.cx,
  )}" cy="${Math.round(b.cy)}"/></a:xfrm>`
}

/**
 * Font size in hundredths of a point, scaled to the slide and clamped.
 *
 * `typeScale` is the built slide's width against PowerPoint's 13.333" default,
 * and without it a template for a 40-inch slide arrives with 44pt titles that
 * are, relative to the slide, a third of the size anyone intended. The clamp is
 * PowerPoint's own: 1pt to 4000pt.
 */
export function scaledSz(basePt: number, typeScale: number): number {
  return Math.max(100, Math.min(400000, Math.round(basePt * 100 * typeScale)))
}

// ---------------------------------------------------------------------------
// Text style lists
// ---------------------------------------------------------------------------

/** `<a:lvl1pPr>` .. `<a:lvl9pPr>`, which is what every text style list wants. */
function levelList(args: {
  sizes: number[]
  typeScale: number
  bullets: boolean
  fontRef: 'mj' | 'mn'
  colour: string
}): string {
  const { sizes, typeScale, bullets, fontRef, colour } = args
  let out = ''
  for (let i = 0; i < 9; i++) {
    const basePt = sizes[Math.min(i, sizes.length - 1)] ?? 18
    const indent = bullets ? Math.round(342900 * typeScale * i) : 0
    const marL = bullets ? Math.round(342900 * typeScale * (i + 1)) : 0
    const bullet = bullets
      ? `<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="&#8226;"/>`
      : '<a:buNone/>'
    out +=
      `<a:lvl${i + 1}pPr marL="${marL}" indent="${indent ? -Math.round(342900 * typeScale) : 0}" algn="l" ` +
      `defTabSz="${Math.round(914400 * typeScale)}" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
      `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>` +
      `<a:spcBef><a:spcPts val="${Math.round(1000 * typeScale)}"/></a:spcBef>` +
      bullet +
      `<a:defRPr sz="${scaledSz(basePt, typeScale)}" kern="1200">` +
      `<a:solidFill><a:schemeClr val="${colour}"/></a:solidFill>` +
      `<a:latin typeface="+${fontRef}-lt"/><a:ea typeface="+${fontRef}-ea"/><a:cs typeface="+${fontRef}-cs"/>` +
      `</a:defRPr></a:lvl${i + 1}pPr>`
  }
  return out
}

function txStyles(typeScale: number): string {
  return (
    `<p:txStyles>` +
    `<p:titleStyle>${levelList({
      sizes: [44],
      typeScale,
      bullets: false,
      fontRef: 'mj',
      colour: 'tx1',
    })}</p:titleStyle>` +
    `<p:bodyStyle>${levelList({
      sizes: [28, 24, 20, 18, 18, 18, 18, 18, 18],
      typeScale,
      bullets: true,
      fontRef: 'mn',
      colour: 'tx1',
    })}</p:bodyStyle>` +
    `<p:otherStyle>${levelList({
      sizes: [18],
      typeScale,
      bullets: false,
      fontRef: 'mn',
      colour: 'tx1',
    })}</p:otherStyle>` +
    `</p:txStyles>`
  )
}

/** `<p:defaultTextStyle>` on the presentation part. Same shape, different tag. */
function defaultTextStyle(typeScale: number): string {
  return `<p:defaultTextStyle>${levelList({
    sizes: [18],
    typeScale,
    bullets: false,
    fontRef: 'mn',
    colour: 'tx1',
  })}</p:defaultTextStyle>`
}

// ---------------------------------------------------------------------------
// Shapes
// ---------------------------------------------------------------------------

/** A placeholder shape. `idx` is required for body placeholders, absent for title. */
function placeholder(args: {
  id: number
  name: string
  phType: string
  idx?: number
  box: Box
  /** Runs to put in it. Empty means PowerPoint supplies its own prompt text. */
  paragraphs?: string[]
  anchor?: 'b' | 'ctr' | 't'
  sz?: number
}): string {
  const { id, name, phType, idx, box, paragraphs, anchor, sz } = args
  const ph = `<p:ph type="${phType}"${idx === undefined ? '' : ` idx="${idx}"`}/>`
  const body =
    paragraphs && paragraphs.length
      ? paragraphs
          .map(
            (p) =>
              `<a:p><a:r><a:rPr lang="en-GB" dirty="0"${sz ? ` sz="${sz}"` : ''}/><a:t>${text(
                p,
              )}</a:t></a:r></a:p>`,
          )
          .join('')
      : `<a:p><a:endParaRPr lang="en-GB" dirty="0"/></a:p>`
  return (
    `<p:sp><p:nvSpPr>` +
    `<p:cNvPr id="${id}" name="${text(name)}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr>${ph}</p:nvPr>` +
    `</p:nvSpPr>` +
    `<p:spPr>${xfrm(box)}</p:spPr>` +
    `<p:txBody><a:bodyPr${anchor ? ` anchor="${anchor}"` : ''}/><a:lstStyle/>${body}</p:txBody>` +
    `</p:sp>`
  )
}

/** A picture. `rId` must already exist in the owning part's rels. */
function picture(args: { id: number; name: string; rId: string; box: Box }): string {
  const { id, name, rId, box } = args
  return (
    `<p:pic><p:nvPicPr>` +
    `<p:cNvPr id="${id}" name="${text(name)}"/>` +
    `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
    `<p:nvPr/>` +
    `</p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `</p:pic>`
  )
}

/** An empty shape tree, which every `<p:cSld>` needs even when it holds nothing. */
function spTree(shapes: string): string {
  return (
    `<p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    shapes +
    `</p:spTree>`
  )
}

// ---------------------------------------------------------------------------
// Media registry
// ---------------------------------------------------------------------------

/**
 * Collects the images the package will carry and names their parts.
 *
 * Part names use only `.png` and `.jpeg` so that `[Content_Types].xml` needs
 * exactly two `<Default>` entries. A file the user picked called `logo.JPG` is
 * still written as `image1.jpeg` — the part name has nothing to do with the
 * original file name, and matching it would mean a Default per extension the
 * user happens to own.
 */
class Media {
  readonly parts: { partName: string; bytes: Uint8Array; mime: ImageMime }[] = []

  add(image: DeckImage): string {
    const ext = image.mime === 'image/png' ? 'png' : 'jpeg'
    const partName = `image${this.parts.length + 1}.${ext}`
    this.parts.push({ partName, bytes: image.bytes, mime: image.mime })
    return partName
  }

  get hasPng(): boolean {
    return this.parts.some((p) => p.mime === 'image/png')
  }

  get hasJpeg(): boolean {
    return this.parts.some((p) => p.mime === 'image/jpeg')
  }
}

// ---------------------------------------------------------------------------
// Parts
// ---------------------------------------------------------------------------

function contentTypes(args: {
  layoutCount: number
  slideCount: number
  media: Media
  fileKind: FileKind
}): string {
  const { layoutCount, slideCount, media, fileKind } = args
  const defaults = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    media.hasPng ? '<Default Extension="png" ContentType="image/png"/>' : '',
    media.hasJpeg ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : '',
  ].join('')

  // The ONE difference between a .pptx and a .potx, other than the file name.
  // A template opens as a new untitled deck; a presentation opens as itself.
  const mainType =
    fileKind === 'potx'
      ? 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'
      : 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'

  const o = (part: string, type: string) => `<Override PartName="${part}" ContentType="${type}"/>`
  const overrides = [
    o('/ppt/presentation.xml', mainType),
    o(
      '/ppt/presProps.xml',
      'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
    ),
    o('/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'),
    o(
      '/ppt/slideMasters/slideMaster1.xml',
      'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    ),
    ...Array.from({ length: layoutCount }, (_, i) =>
      o(
        `/ppt/slideLayouts/slideLayout${i + 1}.xml`,
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
      ),
    ),
    ...Array.from({ length: slideCount }, (_, i) =>
      o(
        `/ppt/slides/slide${i + 1}.xml`,
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      ),
    ),
    o('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
  ].join('')

  return (
    XML_DECL +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    defaults +
    overrides +
    `</Types>`
  )
}

function coreProps(args: { title: string; note: string; now: Date }): string {
  const stamp = args.now.toISOString().replace(/\.\d{3}Z$/, 'Z')
  return (
    XML_DECL +
    `<cp:coreProperties ` +
    `xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
    `xmlns:dcterms="http://purl.org/dc/terms/" ` +
    `xmlns:dcmitype="http://purl.org/dc/dcmitype/" ` +
    `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${text(args.title)}</dc:title>` +
    `<dc:creator>Wallslide</dc:creator>` +
    `<cp:lastModifiedBy>Wallslide</cp:lastModifiedBy>` +
    // The export instructions ride in Comments because it is the only field
    // PowerPoint shows in File > Info without opening anything. A deck that has
    // been emailed twice and renamed once still knows what DPI it wants.
    `<dc:description>${text(args.note)}</dc:description>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>` +
    `</cp:coreProperties>`
  )
}

function presentationXml(args: {
  cx: number
  cy: number
  slideCount: number
  typeScale: number
}): string {
  const { cx, cy, slideCount, typeScale } = args
  const sldIds = Array.from(
    { length: slideCount },
    (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`,
  ).join('')
  return (
    XML_DECL +
    `<p:presentation ${NS_PML} saveSubsetFonts="1">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    (slideCount ? `<p:sldIdLst>${sldIds}</p:sldIdLst>` : '') +
    // THE LINE THE WHOLE APP EXISTS TO WRITE.
    `<p:sldSz cx="${cx}" cy="${cy}"/>` +
    // Notes pages are ordinary A4-ish paper whatever the slide is. Scaling this
    // with the wall would give a 40-inch notes page, which is not a thing.
    `<p:notesSz cx="6858000" cy="9144000"/>` +
    defaultTextStyle(typeScale) +
    `</p:presentation>`
  )
}

function presProps(): string {
  return XML_DECL + `<p:presentationPr ${NS_PML}/>`
}

// ---------------------------------------------------------------------------
// Master and layouts
// ---------------------------------------------------------------------------

/**
 * Where the placeholders go.
 *
 * All five boxes are derived from the SAFE box, not the slide, so that turning
 * a safe area on moves every placeholder in every layout at once. The 5% pad
 * inside it is ordinary slide margin and is a proportion of the safe width, so
 * a wall with a 300 px masked strip down one side does not end up with a
 * lopsided margin on the other.
 */
export function layoutBoxes(safe: Box) {
  const pad = Math.round(safe.cx * 0.05)
  const inner: Box = {
    x: safe.x + pad,
    y: safe.y + pad,
    cx: safe.cx - pad * 2,
    cy: safe.cy - pad * 2,
  }
  const at = (fy: number, fh: number): Box => ({
    x: inner.x,
    y: Math.round(inner.y + inner.cy * fy),
    cx: inner.cx,
    cy: Math.round(inner.cy * fh),
  })
  return {
    inner,
    title: at(0, 0.22),
    body: at(0.28, 0.72),
    ctrTitle: at(0.26, 0.32),
    subTitle: at(0.6, 0.24),
  }
}

/** The logo's box, sized by percentage of slide width and pinned to a corner. */
export function logoBox(args: {
  slideCx: number
  slideCy: number
  safe: Box
  aspect: number
  widthPct: number
  corner: Exclude<Brand['logoCorner'], 'none'>
}): Box {
  const { safe, aspect, widthPct, corner } = args
  const cx = Math.round((safe.cx * Math.max(1, Math.min(100, widthPct))) / 100)
  const cy = Math.round(cx / (aspect || 1))
  // The logo sits inside the safe area for the same reason the text does: a
  // sponsor logo behind the PA stack is a conversation nobody wants to have.
  const margin = Math.round(safe.cx * 0.025)
  const left = safe.x + margin
  const right = safe.x + safe.cx - margin - cx
  const top = safe.y + margin
  const bottom = safe.y + safe.cy - margin - cy
  const x = corner === 'top-left' || corner === 'bottom-left' ? left : right
  const y = corner === 'top-left' || corner === 'top-right' ? top : bottom
  return { x, y, cx, cy }
}

function background(args: { brand: Brand; bgRid: string | null }): string {
  const { brand, bgRid } = args
  if (bgRid) {
    return (
      `<p:bg><p:bgPr>` +
      `<a:blipFill rotWithShape="1"><a:blip r:embed="${bgRid}"/>` +
      `<a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
      `<a:effectLst/></p:bgPr></p:bg>`
    )
  }
  if (brand.useSolidBackground) {
    return (
      `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${brand.backgroundColour}"/></a:solidFill>` +
      `<a:effectLst/></p:bgPr></p:bg>`
    )
  }
  // Falls through to the theme, which is the right default for a template: the
  // organiser changes one colour in the theme and the whole deck follows.
  return `<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>`
}

function slideMasterXml(args: {
  brand: Brand
  boxes: ReturnType<typeof layoutBoxes>
  safe: Box
  slideCx: number
  slideCy: number
  layoutCount: number
  typeScale: number
  bgRid: string | null
  logoRid: string | null
}): string {
  const { brand, boxes, safe, slideCx, slideCy, layoutCount, typeScale, bgRid, logoRid } = args

  const shapes =
    placeholder({ id: 2, name: 'Title Placeholder 1', phType: 'title', box: boxes.title }) +
    placeholder({
      id: 3,
      name: 'Text Placeholder 2',
      phType: 'body',
      idx: 1,
      box: boxes.body,
    }) +
    (logoRid && brand.logo && brand.logoCorner !== 'none'
      ? picture({
          id: 4,
          name: 'Logo',
          rId: logoRid,
          box: logoBox({
            slideCx,
            slideCy,
            safe,
            aspect: brand.logo.widthPx / brand.logo.heightPx,
            widthPct: brand.logoWidthPct,
            corner: brand.logoCorner,
          }),
        })
      : '')

  const layoutIds = Array.from(
    { length: layoutCount },
    (_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`,
  ).join('')

  return (
    XML_DECL +
    `<p:sldMaster ${NS_PML}>` +
    `<p:cSld>${background({ brand, bgRid })}${spTree(shapes)}</p:cSld>` +
    // Office's standard mapping. Swapping bg2/tx2 here (as some exporters do) is
    // legal and is how a deck ends up with one placeholder whose text matches
    // its own background.
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ` +
    `accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" ` +
    `hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst>` +
    txStyles(typeScale) +
    `</p:sldMaster>`
  )
}

interface LayoutDef {
  /** `<p:sldLayout type="">`. */
  type: string
  name: string
  shapes: string
}

function layoutDefs(boxes: ReturnType<typeof layoutBoxes>): LayoutDef[] {
  return [
    {
      type: 'title',
      name: 'Title Slide',
      shapes:
        placeholder({
          id: 2,
          name: 'Title 1',
          phType: 'ctrTitle',
          box: boxes.ctrTitle,
          anchor: 'b',
        }) +
        placeholder({
          id: 3,
          name: 'Subtitle 2',
          phType: 'subTitle',
          idx: 1,
          box: boxes.subTitle,
        }),
    },
    {
      type: 'obj',
      name: 'Title and Content',
      shapes:
        placeholder({ id: 2, name: 'Title 1', phType: 'title', box: boxes.title }) +
        placeholder({ id: 3, name: 'Content Placeholder 2', phType: 'body', idx: 1, box: boxes.body }),
    },
    {
      type: 'titleOnly',
      name: 'Title Only',
      shapes: placeholder({ id: 2, name: 'Title 1', phType: 'title', box: boxes.title }),
    },
    // Blank exists so the pattern slides have somewhere to live that inherits no
    // placeholder geometry at all — and so an organiser has one clean canvas.
    { type: 'blank', name: 'Blank', shapes: '' },
  ]
}

function slideLayoutXml(def: LayoutDef): string {
  return (
    XML_DECL +
    `<p:sldLayout ${NS_PML} type="${def.type}" preserve="1">` +
    `<p:cSld name="${text(def.name)}">${spTree(def.shapes)}</p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    `</p:sldLayout>`
  )
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

/** Which layout (1-based) each slide kind is built on. */
function layoutIndexFor(content: SlideContent): number {
  switch (content.kind) {
    case 'title':
      return 1
    case 'heading':
      return 2
    case 'blank':
      return 4
    case 'full-bleed':
      return 4
  }
}

function slideXml(args: {
  content: SlideContent
  boxes: ReturnType<typeof layoutBoxes>
  slideCx: number
  slideCy: number
  imageRid: string | null
}): string {
  const { content, boxes, slideCx, slideCy, imageRid } = args

  let shapes = ''
  let showMasterSp = ''

  switch (content.kind) {
    case 'title':
      shapes =
        placeholder({
          id: 2,
          name: 'Title 1',
          phType: 'ctrTitle',
          box: boxes.ctrTitle,
          anchor: 'b',
          paragraphs: [content.title],
        }) +
        placeholder({
          id: 3,
          name: 'Subtitle 2',
          phType: 'subTitle',
          idx: 1,
          box: boxes.subTitle,
          paragraphs: content.subtitle ? [content.subtitle] : [],
        })
      break
    case 'heading':
      shapes =
        placeholder({
          id: 2,
          name: 'Title 1',
          phType: 'title',
          box: boxes.title,
          paragraphs: [content.title],
        }) +
        placeholder({
          id: 3,
          name: 'Content Placeholder 2',
          phType: 'body',
          idx: 1,
          box: boxes.body,
          paragraphs: content.body,
        })
      break
    case 'blank':
      break
    case 'full-bleed':
      if (!imageRid) throw new Error('A full-bleed slide was built without its image relationship.')
      shapes = picture({
        id: 2,
        name: content.name,
        rId: imageRid,
        box: { x: 0, y: 0, cx: slideCx, cy: slideCy },
      })
      // Suppress the master's shapes. A test pattern with the event logo on top
      // of it measures the logo, not the wall.
      showMasterSp = ' showMasterSp="0"'
      break
  }

  return (
    XML_DECL +
    `<p:sld ${NS_PML}${showMasterSp}>` +
    `<p:cSld>${spTree(shapes)}</p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    `</p:sld>`
  )
}

// ---------------------------------------------------------------------------
// Assembly
// ---------------------------------------------------------------------------

/** PowerPoint's slide-size range, in EMU. `ST_SlideSizeCoordinate`, 1" to 56". */
export const SLD_SZ_MIN = 914400
export const SLD_SZ_MAX = 51206400

export class SlideSizeError extends Error {}

function part(name: string, content: string | Uint8Array): ZipEntry {
  return { name, data: typeof content === 'string' ? utf8(content) : content }
}

/**
 * Build the package.
 *
 * Throws `SlideSizeError` rather than emitting an out-of-range `<p:sldSz>`. That
 * case is reachable — `slides.ts` reports "steeper than 56:1" as a problem and
 * leaves `buildScale` at 1 rather than inventing a scale — and a package with a
 * 60-inch slide in it opens to a repair prompt that mentions nothing useful.
 * The UI checks the same condition first and never gets here; this is the
 * backstop that makes the module safe to call from anywhere else.
 */
export function buildPptx(spec: DeckSpec): Uint8Array {
  const { wall, slide, brand, safeArea, slides, fileKind } = spec
  const now = spec.now ?? new Date()

  const cx = toEmu(slide.buildWidthMm)
  const cy = toEmu(slide.buildHeightMm)
  for (const [axis, v] of [
    ['width', cx],
    ['height', cy],
  ] as const) {
    if (v < SLD_SZ_MIN || v > SLD_SZ_MAX) {
      throw new SlideSizeError(
        `Slide ${axis} of ${(v / 914400).toFixed(2)}" is outside PowerPoint's 1"–56" range. ` +
          `${wall.widthPx}×${wall.heightPx} cannot be expressed as a slide at any whole build scale.`,
      )
    }
  }

  const media = new Media()
  const safe = safeBox(cx, cy, wall, safeArea)
  const boxes = layoutBoxes(safe)

  // Master media first, so image1 is the background and image2 the logo in every
  // package that has them. Stable part numbering makes two builds of the same
  // deck diffable, which is how the tests check anything at all.
  const bgPart = brand.background ? media.add(brand.background) : null
  const logoPart = brand.logo && brand.logoCorner !== 'none' ? media.add(brand.logo) : null

  const layouts = layoutDefs(boxes)

  // --- master rels: layouts, then theme, then its own media ---
  const masterRels: string[] = layouts.map((_, i) =>
    rel(`rId${i + 1}`, SLIDE_LAYOUT, `../slideLayouts/slideLayout${i + 1}.xml`),
  )
  let masterRid = layouts.length
  masterRels.push(rel(`rId${++masterRid}`, THEME, '../theme/theme1.xml'))
  const bgRid = bgPart ? `rId${++masterRid}` : null
  if (bgPart) masterRels.push(rel(bgRid as string, IMAGE, `../media/${bgPart}`))
  const logoRid = logoPart ? `rId${++masterRid}` : null
  if (logoPart) masterRels.push(rel(logoRid as string, IMAGE, `../media/${logoPart}`))

  const parts: ZipEntry[] = []

  parts.push(
    part(
      'ppt/slideMasters/slideMaster1.xml',
      slideMasterXml({
        brand,
        boxes,
        safe,
        slideCx: cx,
        slideCy: cy,
        layoutCount: layouts.length,
        typeScale: slide.typeScale,
        bgRid,
        logoRid,
      }),
    ),
    part('ppt/slideMasters/_rels/slideMaster1.xml.rels', relsPart(masterRels)),
  )

  layouts.forEach((def, i) => {
    parts.push(
      part(`ppt/slideLayouts/slideLayout${i + 1}.xml`, slideLayoutXml(def)),
      part(
        `ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`,
        relsPart([rel('rId1', SLIDE_MASTER, '../slideMasters/slideMaster1.xml')]),
      ),
    )
  })

  slides.forEach((content, i) => {
    const n = i + 1
    const layoutIndex = layoutIndexFor(content)
    const slideRels = [
      rel('rId1', SLIDE_LAYOUT, `../slideLayouts/slideLayout${layoutIndex}.xml`),
    ]
    let imageRid: string | null = null
    if (content.kind === 'full-bleed') {
      const partName = media.add(content.image)
      imageRid = 'rId2'
      slideRels.push(rel(imageRid, IMAGE, `../media/${partName}`))
    }
    parts.push(
      part(
        `ppt/slides/slide${n}.xml`,
        slideXml({ content, boxes, slideCx: cx, slideCy: cy, imageRid }),
      ),
      part(`ppt/slides/_rels/slide${n}.xml.rels`, relsPart(slideRels)),
    )
  })

  // --- presentation rels: master, slides, presProps, theme ---
  const presRels = [rel('rId1', SLIDE_MASTER, 'slideMasters/slideMaster1.xml')]
  slides.forEach((_, i) => presRels.push(rel(`rId${i + 2}`, SLIDE, `slides/slide${i + 1}.xml`)))
  presRels.push(rel(`rId${slides.length + 2}`, PRES_PROPS, 'presProps.xml'))
  presRels.push(rel(`rId${slides.length + 3}`, THEME, 'theme/theme1.xml'))

  parts.push(
    part(
      'ppt/presentation.xml',
      presentationXml({ cx, cy, slideCount: slides.length, typeScale: slide.typeScale }),
    ),
    part('ppt/_rels/presentation.xml.rels', relsPart(presRels)),
    part('ppt/presProps.xml', presProps()),
    part(
      'ppt/theme/theme1.xml',
      themeXml({
        name: wall.name || 'Wallslide',
        colours: brand.colours,
        majorFont: brand.majorFont,
        minorFont: brand.minorFont,
      }),
    ),
  )

  for (const m of media.parts) parts.push(part(`ppt/media/${m.partName}`, m.bytes))

  parts.push(
    part(
      '[Content_Types].xml',
      contentTypes({
        layoutCount: layouts.length,
        slideCount: slides.length,
        media,
        fileKind,
      }),
    ),
    part(
      '_rels/.rels',
      relsPart([
        rel('rId1', OFFICE_DOC, 'ppt/presentation.xml'),
        rel('rId2', CORE_PROPS, 'docProps/core.xml'),
      ]),
    ),
    part(
      'docProps/core.xml',
      coreProps({
        title: `${wall.name || 'LED wall'} — ${wall.widthPx}×${wall.heightPx}`,
        note: spec.exportNote,
        now,
      }),
    ),
  )

  return buildZip(parts, now)
}

/**
 * The theme colour a burn-in should use so it stays readable on this brand.
 *
 * Exported because the pattern renderer needs the same answer and computing it
 * twice from two different ideas of "the background" is how the two drift.
 */
export function burnInColour(brand: Brand, colours: ThemeColours): string {
  if (brand.background) return 'FFFFFF'
  return readableOn(brand.useSolidBackground ? brand.backgroundColour : colours.lt1)
}

/** `Main Wall 7680x1080.potx`, with anything a file system dislikes removed. */
export function deckFileName(wall: Wall, kind: FileKind, suffix = ''): string {
  const base = (wall.name || 'LED wall').replace(/[\\/:*?"<>|]/g, '-').trim()
  return `${base} ${wall.widthPx}x${wall.heightPx}${suffix}.${kind}`
}
