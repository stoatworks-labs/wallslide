# Attributions

## Vendored source

- **`src/lib/slides.ts`** — from [Aspect Calc](https://github.com/stoatworks-labs/aspect-calc),
  MIT, same author. One edit: the `Problem` import points at this repo's `problem.ts`.
- **`src/lib/zip.ts`** and **`src/lib/__tests__/zip.test.ts`** — from
  [Test Card](https://github.com/stoatworks-labs/test-card), MIT, same author. Unmodified.

## Specifications

- **ECMA-376** (Office Open XML), Parts 1 and 2 — the PresentationML part graph, the
  relationship types, `ST_SlideSizeCoordinate`'s 914400–51206400 EMU range, and OPC's
  permission to store package members uncompressed.
- **PKWARE APPNOTE 6.3.x** — the ZIP structures, via test-card's writer.
- **WCAG 2.1** — the relative-luminance and contrast-ratio formulae in `colour.ts`.

## Measured, not cited

Keynote's 200–8192 pt slide range was obtained by driving a real Keynote until it refused,
and is quoted from its own error text. It appears in no documentation I could find.

## Runtime dependencies

React and React DOM (MIT). Nothing else ships to the browser.
