# Notes — Wallslide

Repo-specific working notes. Fleet-wide things belong in fleet-notes.

## Why a new repo rather than a panel in aspect-calc

aspect-calc already computes the slide size and has a PowerPoint section. The theming UI is
substantially larger than that section and would swamp a calculator whose whole shape is
"three groups, solve for one". test-card was the other candidate and is worse: it is a raster
tool, and a document format does not belong in it. So: a third repo that vendors one file
from each.

The fleet has no shared package and shares code by copying, so that is what this does. The
copies carry their provenance in a header comment; see `CLAUDE.md` for the rule.

## Deployment

Static-assets Worker on Cloudflare, port 4355 in dev. `pages_build_output_dir` does **not**
work — see the comment in `wrangler.toml`, which is the same note aspect-calc carries.

## The verification session, 2026-08-27

Both PowerPoint and Keynote are installed on the build machine, which made real verification
possible and is worth repeating whenever the writer changes:

```bash
npm run samples -- /tmp/wallslide-samples
```

then drive them with `osascript`. Findings worth keeping:

- **`slide height of page setup` errors in PowerPoint's AppleScript dictionary** while
  `slide width` works. Do not read anything into it; the height is right in the file.
- **`presentation 1` is the first-opened document**, not the frontmost. Query by name.
- **PowerPoint's PDF export needs a "Grant File Access" sandbox prompt** that a script cannot
  answer. Export verification needs a human at the keyboard, or a folder PowerPoint already
  has access to.
- **Keynote's slide range is 200–8192 pt**, both axes, inclusive, from its own error text.
  This is now in `keynote.ts` and is the most useful thing the session produced.

## Things deliberately not built

- **`.key` output.** IWA, undocumented. See the README.
- **Drawing guides in `viewProps.xml`.** `<p:guide pos="">` is in eighth-points as far as I
  can tell, but "as far as I can tell" is not good enough for a safe-area marker — one in the
  wrong place is worse than none. The safe area is expressed by moving the placeholders,
  which cannot be wrong, and by an optional pattern slide that draws it.
- **Font embedding.** Licence-encumbered, and Mac PowerPoint ignores it.
- **SMPTE RP 219.** Its value is exact colour; a deck is not an exact-colour path.
