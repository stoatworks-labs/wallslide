/**
 * The small amount of XML machinery an OOXML package needs.
 *
 * There is no DOM here and no builder API. Every part in `pptx.ts` is written as
 * a template literal, because OOXML parts are not really documents you assemble
 * — they are fixed forms with a handful of holes in them, and a builder makes
 * the shape harder to read against the spec rather than easier. What that
 * approach does need is a disciplined escape function on every hole, which is
 * this file.
 */

/** The XML declaration Office writes at the top of every part. Byte-for-byte. */
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

/**
 * Escape for element content and for double-quoted attribute values alike.
 *
 * `>` is escaped although only `<` and `&` strictly must be, because the one
 * sequence that genuinely bites is `]]>` and escaping every `>` removes it
 * without needing to detect it.
 */
export function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

/**
 * Strip characters XML 1.0 cannot represent AT ALL.
 *
 * Escaping does not help here: `&#x0;` is as illegal as a raw NUL. A user can
 * paste anything into an event name — text dragged out of a PDF arrives with
 * control characters in it more often than you would think — and a single stray
 * byte makes the whole package unopenable with an error naming a line number,
 * which is an impossible thing to debug on site.
 *
 * Lone surrogates go too: they survive a JS string but are not valid XML and
 * PowerPoint rejects the part.
 */
export function xmlSafe(s: string): string {
  let out = ''
  for (const ch of s) {
    const c = ch.codePointAt(0) as number
    if (c === 0x9 || c === 0xa || c === 0xd) {
      out += ch
    } else if (c < 0x20) {
      continue
    } else if (c >= 0xd800 && c <= 0xdfff) {
      continue
    } else if (c === 0xfffe || c === 0xffff) {
      continue
    } else {
      out += ch
    }
  }
  return out
}

/** Both passes, in the only order that is correct. */
export function text(s: string): string {
  return esc(xmlSafe(s))
}

/** UTF-8 bytes of a part, which is what the ZIP writer wants. */
export function utf8(s: string): Uint8Array {
  return new TextEncoder().encode(s)
}

/**
 * A `<Relationship>` line.
 *
 * Relationship types are the full schema URIs and are spelled out at each call
 * site rather than abbreviated behind constants — a wrong relationship type is
 * one of the few OOXML mistakes that produces "PowerPoint found a problem with
 * content" and nothing else, and being able to read the URI in place is worth
 * the repetition.
 */
export function rel(id: string, type: string, target: string, mode?: 'External'): string {
  return `<Relationship Id="${esc(id)}" Type="${esc(type)}" Target="${esc(target)}"${
    mode ? ` TargetMode="${mode}"` : ''
  }/>`
}

export function relsPart(relationships: string[]): string {
  return `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join(
    '',
  )}</Relationships>`
}
