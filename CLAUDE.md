# CLAUDE.md — Wallslide

Command reference. For the model, the invariants and the traps, read
[AGENTS.md](AGENTS.md) first.

## Commands

```bash
npm install
npm run dev          # vite dev server on :4355
npm test             # vitest
npm run test:watch
npm run build        # tsc -b && vite build -> dist/
npm run samples      # write sample decks to open in real PowerPoint / Keynote
npm run serve:dist   # serve dist/ WITH _headers applied — use this to check the CSP
npx tsc -b           # typecheck only
```

## Deploy

Static-assets Worker, not Cloudflare Pages.

**Automatic.** `.github/workflows/deploy.yml` tests, builds, smoke-checks and
deploys on every push to `main` that touches something a visitor sees, then
verifies the live `<head>` hash matches the build. Doc-only pushes are skipped.

Needs `CLOUDFLARE_API_TOKEN` (repo secret) and `CLOUDFLARE_ACCOUNT_ID` (repo
variable). The workflow fails in one second with instructions if the token is
missing, rather than after a build.

Manual fallback, for when Actions is down:

```bash
cf-run npx wrangler deploy
```

**Do not connect a Cloudflare dashboard build.** None was ever connected; the
comment that used to claim one is why this repo went weeks without shipping.

## Ground rules

- **`slides.ts` and `zip.ts` are VENDORED COPIES.** Fix bugs in aspect-calc and test-card
  respectively, then re-copy. Do not improve them here; a local edit forks them away from
  the tests that make them trustworthy.
- All internal lengths are **EMU** once past the UI. `slides.ts` speaks millimetres because
  aspect-calc does; convert once, at the boundary, with `toEmu`.
- Never emit a `<p:sldSz>` outside 914400–51206400 EMU. `buildPptx` throws `SlideSizeError`
  rather than writing one — keep it that way, and do not "clamp" instead.
- Colours are stored as **six hex digits with no leading hash**. The hash exists only inside
  the UI's `<input type="color">`.
- Every string that reaches the XML goes through `text()` from `xml.ts`. Not `esc()` alone —
  `xmlSafe()` strips characters XML cannot represent at all, which escaping cannot fix.
- The SVG preview works in a **normalised 1000-unit space, never in EMU**. Browsers
  mis-render SVG geometry at EMU magnitudes, silently. See the comment in `SlidePreview.tsx`.
- Patterns must declare whether they read the **wall** or the **chain**. A pattern that
  cannot survive PowerPoint and does not say so is worse than no pattern.
- Do not add SMPTE RP 219 here. See the README on why.
