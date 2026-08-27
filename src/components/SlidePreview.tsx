/**
 * The slide, drawn to scale, with the things that are easy to get wrong on it.
 *
 * One SVG in a NORMALISED coordinate space 1000 units wide, not in the slide's
 * own EMU. The EMU would be the obvious choice and it does not work: a 40-inch
 * slide is 36,576,000 EMU across, and at that magnitude the browser silently
 * mis-renders SVG geometry — a full-width rect came back from
 * getBoundingClientRect at 46% of the width it asked for. Nothing errors; the
 * picture is simply wrong, which is the worst way for a preview to fail.
 *
 * So the real geometry is computed in EMU by the same `safeBox`, `layoutBoxes`
 * and `logoBox` the package writer uses — the preview is never a second opinion
 * about where things go — and every box is mapped through one scale factor on
 * the way onto the canvas.
 *
 * It shows the safe area and the placeholder boxes because those are the two
 * things a person cannot check any other way until the file is open in
 * PowerPoint, on a machine that may not be to hand.
 */

import { layoutBoxes, logoBox, safeBox, type Box } from '../lib/pptx'
import type { Brand, SafeArea, Wall } from '../types'

export function SlidePreview({
  slideCx,
  slideCy,
  wall,
  safeArea,
  brand,
}: {
  slideCx: number
  slideCy: number
  wall: Wall
  safeArea: SafeArea
  brand: Brand
}) {
  if (!(slideCx > 0) || !(slideCy > 0)) return null

  // One thousand units wide, and the height follows the slide's ratio.
  const VW = 1000
  const k = VW / slideCx
  const VH = slideCy * k
  const map = (b: Box): Box => ({ x: b.x * k, y: b.y * k, cx: b.cx * k, cy: b.cy * k })

  const safeEmu = safeBox(slideCx, slideCy, wall, safeArea)
  const boxesEmu = layoutBoxes(safeEmu)
  const safe = map(safeEmu)
  const boxes = {
    title: map(boxesEmu.title),
    body: map(boxesEmu.body),
  }
  const bg = brand.useSolidBackground ? `#${brand.backgroundColour}` : `#${brand.colours.lt1}`
  const ink = `#${brand.colours.dk1}`
  const accent = `#${brand.colours.accent1}`

  const logo =
    brand.logo && brand.logoCorner !== 'none'
      ? map(
          logoBox({
            slideCx,
            slideCy,
            safe: safeEmu,
            aspect: brand.logo.widthPx / brand.logo.heightPx,
            widthPct: brand.logoWidthPct,
            corner: brand.logoCorner,
          }),
        )
      : null

  // One hairline whatever the slide's absolute size, in the normalised space.
  const hair = Math.max(VW, VH) / 400

  const rect = (b: Box, props: Record<string, string | number>) => (
    <rect x={b.x} y={b.y} width={b.cx} height={b.cy} {...props} />
  )

  return (
    <figure className="preview">
      <svg viewBox={`0 0 ${VW} ${VH}`} className="preview__svg" role="img"
        aria-label={`Slide preview, ${wall.widthPx} by ${wall.heightPx} pixels`}>
        <rect x={0} y={0} width={VW} height={VH} fill={bg} />
        {brand.background ? (
          <image
            href={brand.background.previewUrl}
            x={0}
            y={0}
            width={VW}
            height={VH}
            preserveAspectRatio="xMidYMid slice"
          />
        ) : null}

        {safeArea.enabled ? (
          <>
            {/* The masked region, drawn as what it is: not yours to use. */}
            <path
              d={`M0,0 H${VW} V${VH} H0 Z M${safe.x},${safe.y} v${safe.cy} h${safe.cx} v${-safe.cy} Z`}
              fill="rgba(200,16,46,0.28)"
              fillRule="evenodd"
            />
            {rect(safe, {
              fill: 'none',
              stroke: '#00d26a',
              'stroke-width': hair,
              'stroke-dasharray': `${hair * 6} ${hair * 4}`,
            })}
          </>
        ) : null}

        {rect(boxes.title, { fill: ink, opacity: 0.16 })}
        {rect(boxes.body, { fill: ink, opacity: 0.08 })}
        <rect
          x={boxes.title.x}
          y={boxes.title.y + boxes.title.cy}
          width={boxes.title.cx * 0.35}
          height={hair * 2}
          fill={accent}
        />

        {logo ? (
          brand.logo ? (
            <image
              href={brand.logo.previewUrl}
              x={logo.x}
              y={logo.y}
              width={logo.cx}
              height={logo.cy}
              preserveAspectRatio="xMidYMid meet"
            />
          ) : null
        ) : null}

        <rect
          x={hair / 2}
          y={hair / 2}
          width={VW - hair}
          height={VH - hair}
          fill="none"
          stroke={ink}
          strokeWidth={hair}
          opacity={0.35}
        />
      </svg>
      <figcaption>
        Title and body placeholders sit inside the safe area, so a masked edge moves the text
        rather than hiding it.
      </figcaption>
    </figure>
  )
}
