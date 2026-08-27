# Wallslide

> **AI-assisted project.** This codebase was created with [Claude Code](https://claude.com/claude-code)
> (Anthropic), directed and reviewed by a human author. The generated packages have been
> **opened in a real Microsoft PowerPoint and a real Keynote** on macOS, which is the only
> check that means anything for a file format whose failure mode is a dialog that names no
> reason: four decks opened without a repair prompt, PowerPoint reported their slide widths
> as 40.00″, 20.00″, 11.25″ and 40.00″ — exactly the computed values — the `.potx` opened as
> a new untitled presentation as a template should, deliberately hostile text came back
> through the XML unharmed, and a full-bleed picture measured 2880 × 810 pt against a
> 2880 × 810 pt slide. Keynote's own slide limits were **measured by asking Keynote** until
> it refused, not taken from documentation. Every XML part is parsed by `xmllint` in CI.
> **No deck produced by this has been played out to a real LED processor or wall and
> measured.** The slide is the right size; what a playback machine and a processor do with
> the picture is not something this can tell you.

Type an LED wall's resolution. Get a PowerPoint template that is actually the right size,
themed with the event's own colours, fonts, logo and background — and, if you want one, a
deck of test patterns rendered at the wall's exact raster.

Everything happens in the browser. There is no backend, so a client's unannounced logo has
nowhere to be uploaded to.

---

## The problem it exists for

A 7680 × 1080 wall at 96 dpi wants an 80-inch slide. **PowerPoint will not make one** — the
format caps a slide edge at 56 inches, and it is a schema limit (`ST_SlideSizeCoordinate`
tops out at 51,206,400 EMU), not a dialog being fussy. So the deck gets built at 16:9
instead and stretched, or built at "close enough" and letterboxed, and the first anyone
knows is at the tech rehearsal.

The answer is to build at half size and export at 2×. That is arithmetic PowerPoint will not
do for you, and it is the arithmetic people get wrong in the other direction.

| Wall | Slide PowerPoint accepts | Export at |
|---|---|---|
| 1920 × 1080 | 20″ × 11.25″ | 96 dpi |
| 3840 × 1080 | 40″ × 11.25″ | 96 dpi |
| 7680 × 1080 | 40″ × 5.625″ — **half size** | 192 dpi |
| 1080 × 3840 | 11.25″ × 40″ | 96 dpi |

The build scale is always a whole number or a whole reciprocal. "Build at half and export at
200%" is an instruction a person can follow at 2 a.m.; "build at 1/2.37" is not.

## Keynote is not PowerPoint, and it is not close

Measured from Keynote itself — it names its own range in the error it raises:

| | Minimum | Maximum |
|---|---|---|
| PowerPoint | 72 pt (1″) | **4032 pt** (56″) |
| Keynote | **200 pt** (2.78″) | **8192 pt** (113.78″) |

Keynote takes a slide **more than twice as wide** as PowerPoint will. A 7680 px wall is
5760 pt, which PowerPoint must halve and Keynote simply holds at 1:1 — no build scale, no
export multiplier, nothing to get wrong. That is a real reason to choose one application
over the other for a given wall.

The floor runs the other way: Keynote will not go below 200 pt on either edge, so a
1080 × 160 px ticker strip is legal in PowerPoint and impossible in Keynote at native size.

**A `.key` file is not generated, and will not be.** Since 2013 Keynote's format has been
IWA — Snappy-compressed protobuf in an undocumented schema. Wallslide serves Keynote the two
ways that work: open the generated `.pptx` in it (verified — Keynote imports and keeps the
slide size), or type the two numbers Wallslide gives you into Document Setup.

## What goes in the file

**The slide size**, as `<p:sldSz>`, at the build size with the export DPI recorded in the
package's Comments field so the deck still knows what it wants after it has been emailed
twice and renamed once.

**A theme** — twelve colours and two font names, in `ppt/theme/theme1.xml`. An organiser
changes the theme and the whole deck follows, which is the entire reason to ship a template
rather than a deck.

**A slide master and four layouts** — Title Slide, Title and Content, Title Only, Blank —
with the event's logo and background on the master where they can be edited in one place.

**Type scaled to the slide.** A 40-inch slide with 44 pt titles has titles a third of the
size anyone intended. Master text styles are multiplied by the slide's width against
PowerPoint's 13.333″ default.

**A safe area, if part of the wall is masked.** Insets in wall pixels — because the reason is
physical, and "the bottom two cabinets are behind the band" is measured in cabinets. Every
placeholder in every layout moves inside them.

**Test patterns, optionally**, rendered at the wall's native raster and placed full-bleed
with the master's shapes suppressed, because a test pattern with a logo on it is a picture
of a logo.

## Two kinds of test pattern, and the difference matters

Most patterns diagnose **the wall**. Two diagnose **the chain**:

> **Pixel check** — a 1 px checkerboard and line bursts. If this comes out as flat grey,
> something between PowerPoint and the panels is scaling. That is worth knowing and it is
> **not a fault in the wall.**

A deck that does not make that distinction is how somebody spends midnight re-terminating a
healthy wall. Every pattern in the UI is labelled with which of the two it reads.

### This is not Test Card, and should not become it

[Test Card](https://github.com/stoatworks-labs/test-card) is the serious pattern generator:
SMPTE RP 219 measured against ffmpeg's own `smptehdbars`, real cabinet maps imported from
Pixel Peeker, PNGs at the exact raster of every physical output. If you want patterns, go
there.

What is here is the subset that survives a trip through a deck. **RP 219 is deliberately
absent**: its whole value is that the fifteen colours are exactly right, and a pattern whose
point is exact colour should not be sent down a path that may colour-manage it. The bars
here are plain 75% bars and are labelled as such.

## Fonts are named, not embedded

The theme asks for a typeface; it does not carry one. Font embedding is licence-encumbered
and Mac PowerPoint does not honour it, so a template that claimed to embed would be lying on
one of the two platforms an event runs on. Check the playback machine with
[pptx-font-manager](https://github.com/stoatworks-labs/pptx-font-manager) before the show.

Relatedly, the themes generated here carry **no script fallback table**. Office writes about
forty `<a:font script="...">` entries into every theme whether or not the deck contains a
character of any of them — which is why searching a real deck for `typeface=` finds 39 fonts
where it uses two. A Wallslide theme names the fonts it uses and nothing else. The cost is
real: a deck that will carry Japanese or Devanagari should have its theme fonts set to a face
that covers it, because there is no curated fallback to rescue it.

## Where the engine came from

Two of the three hard parts were already written and are **vendored unmodified**, with the
originals named in each file's header:

- **[Aspect Calc](https://github.com/stoatworks-labs/aspect-calc)** — `slides.ts`. The 56-inch
  cap, the 1-inch floor, the 100 MP bitmap-export ceiling, the whole-fraction build scale and
  the type scale. Pinned by 100 tests over there.
- **[Test Card](https://github.com/stoatworks-labs/test-card)** — `zip.ts`. A store-only ZIP
  writer checked against the published CRC-32 vector and unpacked with the system `unzip`.
  A `.pptx` *is* a ZIP, and OPC permits stored members, so it needed no changes at all.

Fix bugs in those repos and re-copy; do not improve the copies in place.

## Development

```bash
npm install
npm run dev
npm test
npm run build
```

`node scripts/emit-samples.ts <dir>` — via `npm run samples` — writes sample decks to open in
the real applications. The unit tests prove the package is a well-formed ZIP of well-formed
XML; they cannot prove PowerPoint will open it, and the gap between those two statements is
where every OOXML bug lives.

## Licence

MIT. See [LICENSE](LICENSE).
