# AGENTS.md — bringing an LLM up to speed on Wallslide

Orientation for an AI assistant (or a new human) picking this up cold. `CLAUDE.md` holds the
command reference; this explains the model and the traps.

---

## 1. What this is

A generator for **PowerPoint templates sized to an LED wall**, with the event's branding in
the theme and slide master, plus an optional deck of test patterns. Browser-only, no backend:
React + TypeScript + Vite, built to a static `dist/` and served by a Cloudflare Worker with
static assets. State is in `localStorage`.

## 2. Layout

```
src/
  types.ts                UiState and the shapes the generators consume
  lib/problem.ts          the {level, text} advisory type
  lib/units.ts            mm <-> inch. Nothing else
  lib/colour.ts           hex handling + WCAG contrast
  lib/slides.ts           VENDORED from aspect-calc. PowerPoint's limits
  lib/keynote.ts          Keynote's DIFFERENT limits, measured from Keynote
  lib/zip.ts              VENDORED from test-card. Store-only ZIP writer
  lib/xml.ts              escaping, and the XML declaration
  lib/theme.ts            ppt/theme/theme1.xml
  lib/pptx.ts             THE PACKAGE WRITER. The interesting module
  lib/patterns.ts         canvas test patterns, for the deck only
  lib/image.ts            File -> ImageAsset, no re-encode
  lib/deck.ts             UiState -> finished file
  components/SlidePreview.tsx   the picture. Normalised SVG space
  components/ui.tsx       Field / Panel / Segmented / Stat / Swatch / Toggle
  App.tsx                 wiring and all the state
```

## 3. The relation everything turns on

```
slide size (inches)  ×  export DPI  =  resolution (pixels)
```

This is aspect-calc's `res × pitch = size` with the pitch written the other way up, and
`slides.ts` is that module, vendored. What makes it non-trivial is that PowerPoint adds three
constraints an LED wall does not have — a 56″ ceiling, a 1″ floor, and a 100 MP export cap —
and those constraints are the entire reason anyone needs this done for them.

Keynote adds a **different** pair of constraints (200 pt / 8192 pt), which is why
`keynote.ts` exists as a separate calculator rather than a flag on the first one. Folding
them together would produce a single "slide size" that is correct for neither application.

## 4. Traps

**A malformed package names nothing.** PowerPoint says "found a problem with content" and
stops. There is no line number, no part name, no reason. This is why the tests unpack the
archive and run `xmllint` over every part rather than asserting on strings, and why
`scripts/emit-samples.ts` exists at all — the only real proof is opening the file.

**`presentation 1` in AppleScript is the first-opened document, not the front one.** Verifying
several decks in a loop and reading `presentation 1` each time measures the same deck four
times and reports success. Query by name.

**SVG at EMU magnitudes renders wrong and does not error.** A `<rect>` asking for the full
36,576,000-unit width came back from `getBoundingClientRect` at 46% of it. The preview works
in a normalised 1000-unit space for this reason and must keep doing so.

**The build size is not the native size.** `<p:sldSz>` carries `slide.buildWidthMm`, never the
native. Writing the native size is the one bug that produces a file which looks plausible and
will not open.

**Media part numbering is stable on purpose.** Background is `image1`, logo is `image2`, then
patterns in slide order. Two builds of the same deck are byte-identical, and a test asserts
it — that is what makes any other assertion about the bytes meaningful.

**`showMasterSp="0"` on pattern slides is load-bearing**, not cosmetic. Without it the event
logo sits on top of every test pattern.

## 5. What has and has not been verified

Verified against the real applications on macOS: packages open in PowerPoint with no repair
prompt; slide widths read back exactly; `.potx` opens as a new untitled presentation; escaped
text round-trips; a full-bleed picture measures the full slide; Keynote imports a `.pptx` and
keeps the slide size; Keynote's 200–8192 pt limits come from Keynote's own error text.

Not verified: **nothing produced here has been played out to an LED processor or a wall.**
PowerPoint's PDF/image export path is also unverified — the export needs a sandbox permission
grant that a script cannot give itself. If you get a chance to run a deck onto real hardware,
that is the disclaimer in the README to come back and replace.
