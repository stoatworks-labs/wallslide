/**
 * Turning a picked file into an `ImageAsset`.
 *
 * The bytes are read straight through and embedded unchanged — no canvas
 * round-trip, no re-encode. That matters for a logo: re-encoding a PNG with an
 * alpha channel through a canvas is how a crisp logo acquires a grey fringe,
 * and re-encoding a JPEG loses a generation for nothing. The only thing decoded
 * is the intrinsic size, which is needed to place the logo without stretching
 * it, and that comes from an `Image` the bytes are never taken back out of.
 *
 * Nothing here uploads anything. There is no server to upload to — the whole
 * app is static assets — and for an event organiser's unannounced client logo
 * that is the entire point rather than a technicality.
 */

import type { ImageAsset, ImageMime } from '../types'

/** What PowerPoint will take without a re-encode, and `[Content_Types]` declares. */
const ACCEPTED: Record<string, ImageMime> = {
  'image/png': 'image/png',
  'image/jpeg': 'image/jpeg',
  'image/jpg': 'image/jpeg',
}

export const ACCEPT_ATTR = 'image/png,image/jpeg'

/** 8 MB. A logo past this is a scan of a logo, and it will bloat every deck. */
export const MAX_IMAGE_BYTES = 8 * 1024 * 1024

export class ImageError extends Error {}

export async function loadImage(file: File): Promise<ImageAsset> {
  const mime = ACCEPTED[file.type]
  if (!mime) {
    throw new ImageError(
      `${file.name} is ${file.type || 'of an unknown type'}. PowerPoint takes PNG and JPEG here; ` +
        `an SVG or a PDF logo needs exporting to one of those first.`,
    )
  }
  if (file.size > MAX_IMAGE_BYTES) {
    throw new ImageError(
      `${file.name} is ${(file.size / 1e6).toFixed(1)} MB. The limit is ${(
        MAX_IMAGE_BYTES / 1e6
      ).toFixed(0)} MB — every slide in the deck carries this file, so a large one makes a deck nobody can email.`,
    )
  }

  const bytes = new Uint8Array(await file.arrayBuffer())
  const previewUrl = URL.createObjectURL(file)

  const { widthPx, heightPx } = await new Promise<{ widthPx: number; heightPx: number }>(
    (resolve, reject) => {
      const img = new Image()
      img.onload = () => resolve({ widthPx: img.naturalWidth, heightPx: img.naturalHeight })
      img.onerror = () => {
        URL.revokeObjectURL(previewUrl)
        reject(new ImageError(`${file.name} could not be decoded — it may be truncated or misnamed.`))
      }
      img.src = previewUrl
    },
  )

  return { name: file.name, mime, bytes, widthPx, heightPx, previewUrl }
}

export function releaseImage(asset: ImageAsset | null) {
  if (asset) URL.revokeObjectURL(asset.previewUrl)
}
