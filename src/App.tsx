/**
 * All the state and the wiring.
 *
 * Numeric fields are held as TEXT, the same choice aspect-calc makes and for the
 * same reason: a field parsed to a number on every keystroke cannot hold "192"
 * on its way to "1920", and clearing one to type a new value fights the user.
 * Parsing happens once, here, at the boundary into the engines.
 *
 * The images are the exception to the persistence rule — everything else is
 * written to localStorage, and a logo is not, because it is somebody's
 * unannounced client and because an ObjectURL does not survive a reload anyway.
 */

import { useEffect, useMemo, useRef, useState } from 'react'

import { Field, Panel, Problems, Segmented, Stat, Swatch, Toggle } from './components/ui'
import { SlidePreview } from './components/SlidePreview'
import { buildDeck, download, exportNote } from './lib/deck'
import { keynoteFieldText, keynoteFromResolution, KEYNOTE_MAX_PT, KEYNOTE_MIN_PT } from './lib/keynote'
import { ACCEPT_ATTR, ImageError, loadImage, releaseImage } from './lib/image'
import { PATTERNS } from './lib/patterns'
import { canvasLimitProblem } from './lib/patterns'
import { PT_PER_INCH, slideFromResolution, toEmu } from './lib/slides'
import { OFFICE_COLOURS } from './lib/theme'
import { contrastRatio } from './lib/colour'
import type { Brand, DeckOptions, FileKind, ImageAsset, PatternKind, SafeArea } from './types'

declare const __APP_VERSION__: string

const WALL_PRESETS: { label: string; w: number; h: number; note?: string }[] = [
  { label: '1920 × 1080', w: 1920, h: 1080, note: 'One HD feed' },
  { label: '2560 × 1440', w: 2560, h: 1440 },
  { label: '3840 × 1080', w: 3840, h: 1080, note: 'Two HD feeds side by side' },
  { label: '3840 × 2160', w: 3840, h: 2160, note: 'UHD' },
  { label: '5760 × 1080', w: 5760, h: 1080, note: 'Three HD feeds' },
  { label: '7680 × 1080', w: 7680, h: 1080, note: 'Four HD feeds — over PowerPoint’s limit' },
  { label: '1080 × 1920', w: 1080, h: 1920, note: 'Portrait totem' },
  { label: '2048 × 1152', w: 2048, h: 1152 },
]

const STORAGE_KEY = 'wallslide.v1'

interface Persisted {
  name: string
  width: string
  height: string
  dpi: string
  safe: { enabled: boolean; top: string; right: string; bottom: string; left: string }
  brand: Omit<Brand, 'logo' | 'background'>
  deck: DeckOptions
  fileKind: FileKind
}

const DEFAULTS: Persisted = {
  name: 'Main wall',
  width: '3840',
  height: '1080',
  dpi: '96',
  safe: { enabled: false, top: '0', right: '0', bottom: '0', left: '0' },
  brand: {
    colours: OFFICE_COLOURS,
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
    logoCorner: 'top-right',
    logoWidthPct: 12,
    useSolidBackground: false,
    backgroundColour: 'FFFFFF',
  },
  deck: { specSlide: true, patterns: [], burnIn: true },
  fileKind: 'potx',
}

function load(): Persisted {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return DEFAULTS
    const parsed = JSON.parse(raw) as Partial<Persisted>
    // Merge rather than replace: a stored state from an older version is missing
    // whatever was added since, and spreading defaults under it is the
    // difference between a new field appearing and the app failing to start.
    return {
      ...DEFAULTS,
      ...parsed,
      safe: { ...DEFAULTS.safe, ...parsed.safe },
      brand: { ...DEFAULTS.brand, ...parsed.brand, colours: { ...OFFICE_COLOURS, ...parsed.brand?.colours } },
      deck: { ...DEFAULTS.deck, ...parsed.deck },
    }
  } catch {
    return DEFAULTS
  }
}

const num = (s: string): number => {
  const n = Number(s.replace(/[^0-9.\-]/g, ''))
  return Number.isFinite(n) ? n : 0
}

export default function App() {
  const [s, setS] = useState<Persisted>(load)
  const [logo, setLogo] = useState<ImageAsset | null>(null)
  const [background, setBackground] = useState<ImageAsset | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const logoInput = useRef<HTMLInputElement>(null)
  const bgInput = useRef<HTMLInputElement>(null)

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(s))
    } catch {
      // A full or disabled localStorage is not a reason to stop working.
    }
  }, [s])

  // Object URLs outlive the component unless something revokes them.
  useEffect(() => () => {
    releaseImage(logo)
    releaseImage(background)
  }, [logo, background])

  const wall = {
    name: s.name,
    widthPx: Math.round(num(s.width)),
    heightPx: Math.round(num(s.height)),
  }
  const dpi = num(s.dpi) || 96

  const brand: Brand = { ...s.brand, logo, background }
  const safeArea: SafeArea = {
    enabled: s.safe.enabled,
    topPx: Math.round(num(s.safe.top)),
    rightPx: Math.round(num(s.safe.right)),
    bottomPx: Math.round(num(s.safe.bottom)),
    leftPx: Math.round(num(s.safe.left)),
  }

  const slide = useMemo(
    () => slideFromResolution(wall.widthPx, wall.heightPx, dpi),
    [wall.widthPx, wall.heightPx, dpi],
  )
  const keynote = useMemo(
    () =>
      slide ? keynoteFromResolution(wall.widthPx, wall.heightPx, dpi, slide.buildScale) : null,
    [wall.widthPx, wall.heightPx, dpi, slide],
  )

  const fatal = slide?.problems.find((p) => p.level === 'error') ?? null
  const rasterProblem = s.deck.patterns.length
    ? canvasLimitProblem(wall.widthPx, wall.heightPx)
    : null

  // The contrast check exists because an LED wall is viewed at distance, where a
  // 3:1 title over its background is not "a bit low", it is unreadable.
  const bgHex = s.brand.useSolidBackground ? s.brand.backgroundColour : s.brand.colours.lt1
  const textContrast = contrastRatio(s.brand.colours.dk1, bgHex)

  async function pick(kind: 'logo' | 'background', file: File | undefined) {
    if (!file) return
    setError(null)
    try {
      const asset = await loadImage(file)
      if (kind === 'logo') {
        releaseImage(logo)
        setLogo(asset)
      } else {
        releaseImage(background)
        setBackground(asset)
      }
    } catch (e) {
      setError(e instanceof ImageError ? e.message : String(e))
    }
  }

  async function generate() {
    if (!slide || fatal) return
    setBusy(true)
    setError(null)
    try {
      const built = await buildDeck(
        { wall, brand, safeArea, deck: s.deck, fileKind: s.fileKind, dpi },
        slide,
        keynote,
      )
      download(built.bytes, built.fileName)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  const togglePattern = (id: PatternKind) =>
    setS((v) => ({
      ...v,
      deck: {
        ...v.deck,
        patterns: v.deck.patterns.includes(id)
          ? v.deck.patterns.filter((p) => p !== id)
          : [...v.deck.patterns, id],
      },
    }))

  const inches = (mm: number) => mm / 25.4

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <strong>Wallslide</strong>
          <span className="brand__sub">PowerPoint templates sized for an LED wall</span>
        </div>
        <div className="spacer" />
        <Segmented<FileKind>
          label="File kind"
          value={s.fileKind}
          onChange={(fileKind) => setS((v) => ({ ...v, fileKind }))}
          options={[
            { id: 'potx', label: '.potx template', title: 'Opens as a new untitled deck' },
            { id: 'pptx', label: '.pptx deck', title: 'Opens as itself' },
          ]}
        />
        <button className="primary" onClick={generate} disabled={busy || !slide || !!fatal}>
          {busy ? 'Building…' : 'Download'}
        </button>
      </header>

      {error ? <div className="banner banner--error">{error}</div> : null}

      <main className="cols">
        <div className="col">
          <Panel title="The wall">
            <div className="row">
              <Field label="Name" value={s.name} inputMode="text" onChange={(name) => setS((v) => ({ ...v, name }))} />
            </div>
            <div className="row row--2">
              <Field label="Width" suffix="px" value={s.width} onChange={(width) => setS((v) => ({ ...v, width }))} />
              <Field label="Height" suffix="px" value={s.height} onChange={(height) => setS((v) => ({ ...v, height }))} />
            </div>
            <div className="chips">
              {WALL_PRESETS.map((p) => (
                <button
                  key={p.label}
                  type="button"
                  className="chip chip--btn"
                  title={p.note}
                  onClick={() => setS((v) => ({ ...v, width: String(p.w), height: String(p.h) }))}
                >
                  {p.label}
                </button>
              ))}
            </div>
            <div className="row">
              <Field
                label="Export DPI"
                value={s.dpi}
                onChange={(d) => setS((v) => ({ ...v, dpi: d }))}
                hint="96 is what PowerPoint exports at unless the registry says otherwise. Change this only if you know yours differs."
              />
            </div>
          </Panel>

          <Panel title="Safe area">
            <Toggle
              label="Part of this wall is masked or hidden"
              checked={s.safe.enabled}
              onChange={(enabled) => setS((v) => ({ ...v, safe: { ...v.safe, enabled } }))}
              hint="Insets in wall pixels. Every placeholder in every layout moves inside them."
            />
            {s.safe.enabled ? (
              <div className="row row--4">
                {(['top', 'right', 'bottom', 'left'] as const).map((k) => (
                  <Field
                    key={k}
                    label={k[0]!.toUpperCase() + k.slice(1)}
                    suffix="px"
                    value={s.safe[k]}
                    onChange={(val) => setS((v) => ({ ...v, safe: { ...v.safe, [k]: val } }))}
                  />
                ))}
              </div>
            ) : null}
          </Panel>

          <Panel title="Theme">
            <div className="row row--2">
              <Field
                label="Heading font"
                inputMode="text"
                value={s.brand.majorFont}
                onChange={(majorFont) => setS((v) => ({ ...v, brand: { ...v.brand, majorFont } }))}
              />
              <Field
                label="Body font"
                inputMode="text"
                value={s.brand.minorFont}
                onChange={(minorFont) => setS((v) => ({ ...v, brand: { ...v.brand, minorFont } }))}
              />
            </div>
            <p className="note">
              Fonts are <em>named</em>, not embedded — the deck asks for them and the playback
              machine must have them. Check that with{' '}
              <a href="https://github.com/stoatworks-labs/pptx-font-manager">pptx-font-manager</a>{' '}
              before the show.
            </p>
            <div className="swatches">
              {(
                [
                  ['dk1', 'Text'],
                  ['lt1', 'Background'],
                  ['accent1', 'Accent 1'],
                  ['accent2', 'Accent 2'],
                  ['accent3', 'Accent 3'],
                  ['accent4', 'Accent 4'],
                  ['accent5', 'Accent 5'],
                  ['accent6', 'Accent 6'],
                ] as const
              ).map(([k, label]) => (
                <Swatch
                  key={k}
                  label={label}
                  value={s.brand.colours[k]}
                  onChange={(hex) =>
                    setS((v) => ({ ...v, brand: { ...v.brand, colours: { ...v.brand.colours, [k]: hex } } }))
                  }
                />
              ))}
            </div>
            {textContrast < 4.5 ? (
              <p className="note note--warn">
                Text on background is {textContrast.toFixed(1)}:1. On a wall seen from the back of a
                room, anything under about 4.5:1 stops being readable rather than merely looking
                subtle.
              </p>
            ) : null}
          </Panel>

          <Panel title="Logo and background">
            <div className="row row--2">
              <div>
                <span className="field__label">Logo</span>
                <div className="filerow">
                  <button className="ghost" onClick={() => logoInput.current?.click()}>
                    {logo ? 'Replace…' : 'Choose…'}
                  </button>
                  {logo ? (
                    <>
                      <img className="thumb" src={logo.previewUrl} alt="" />
                      <button
                        className="ghost"
                        onClick={() => {
                          releaseImage(logo)
                          setLogo(null)
                        }}
                      >
                        Remove
                      </button>
                    </>
                  ) : null}
                </div>
                <input
                  ref={logoInput}
                  type="file"
                  accept={ACCEPT_ATTR}
                  hidden
                  onChange={(e) => pick('logo', e.target.files?.[0])}
                />
              </div>
              <div>
                <span className="field__label">Background</span>
                <div className="filerow">
                  <button className="ghost" onClick={() => bgInput.current?.click()}>
                    {background ? 'Replace…' : 'Choose…'}
                  </button>
                  {background ? (
                    <>
                      <img className="thumb" src={background.previewUrl} alt="" />
                      <button
                        className="ghost"
                        onClick={() => {
                          releaseImage(background)
                          setBackground(null)
                        }}
                      >
                        Remove
                      </button>
                    </>
                  ) : null}
                </div>
                <input
                  ref={bgInput}
                  type="file"
                  accept={ACCEPT_ATTR}
                  hidden
                  onChange={(e) => pick('background', e.target.files?.[0])}
                />
              </div>
            </div>
            {logo ? (
              <div className="row row--2">
                <label className="field">
                  <span className="field__label">Logo corner</span>
                  <select
                    className="input"
                    value={s.brand.logoCorner}
                    onChange={(e) =>
                      setS((v) => ({ ...v, brand: { ...v.brand, logoCorner: e.target.value as Brand['logoCorner'] } }))
                    }
                  >
                    <option value="none">Do not place</option>
                    <option value="top-left">Top left</option>
                    <option value="top-right">Top right</option>
                    <option value="bottom-left">Bottom left</option>
                    <option value="bottom-right">Bottom right</option>
                  </select>
                </label>
                <Field
                  label="Logo width"
                  suffix="% of slide"
                  value={String(s.brand.logoWidthPct)}
                  onChange={(val) =>
                    setS((v) => ({ ...v, brand: { ...v.brand, logoWidthPct: num(val) || 12 } }))
                  }
                />
              </div>
            ) : null}
            <Toggle
              label="Use a solid background colour"
              checked={s.brand.useSolidBackground}
              onChange={(useSolidBackground) =>
                setS((v) => ({ ...v, brand: { ...v.brand, useSolidBackground } }))
              }
              hint="Off means the slide follows the theme, which is what a template usually wants."
            />
            {s.brand.useSolidBackground ? (
              <div className="swatches">
                <Swatch
                  label="Background"
                  value={s.brand.backgroundColour}
                  onChange={(backgroundColour) =>
                    setS((v) => ({ ...v, brand: { ...v.brand, backgroundColour } }))
                  }
                />
              </div>
            ) : null}
            <p className="note">
              Nothing here is uploaded. The whole app is static files — there is no server to
              upload a client’s logo to.
            </p>
          </Panel>

          <Panel title="Test deck">
            <Toggle
              label="Specification slide"
              checked={s.deck.specSlide}
              onChange={(specSlide) => setS((v) => ({ ...v, deck: { ...v.deck, specSlide } }))}
              hint="The build scale and export DPI, on a slide, so the deck carries its own instructions."
            />
            <Toggle
              label="Burn the wall name and resolution into each pattern"
              checked={s.deck.burnIn}
              onChange={(burnIn) => setS((v) => ({ ...v, deck: { ...v.deck, burnIn } }))}
            />
            <div className="patterns">
              {PATTERNS.map((p) => (
                <label key={p.id} className="pattern">
                  <input
                    type="checkbox"
                    checked={s.deck.patterns.includes(p.id)}
                    onChange={() => togglePattern(p.id)}
                  />
                  <span>
                    <strong>{p.label}</strong>
                    <em className={p.reads === 'chain' ? 'reads reads--chain' : 'reads'}>
                      reads the {p.reads}
                    </em>
                    <small>{p.purpose}</small>
                  </span>
                </label>
              ))}
            </div>
            {rasterProblem ? <p className="note note--warn">{rasterProblem}</p> : null}
            <p className="note">
              These are a subset, chosen because they survive a trip through PowerPoint. For
              standards-grade patterns as PNGs at the exact raster of every output — including SMPTE
              RP 219 and real cabinet maps — use{' '}
              <a href="https://github.com/stoatworks-labs/test-card">Test Card</a> instead.
            </p>
          </Panel>
        </div>

        <div className="col col--out">
          {slide ? (
            <>
              <Panel title="The slide">
                <div className="stats">
                  <Stat
                    label="Slide size"
                    tone="accent"
                    value={`${inches(slide.buildWidthMm).toFixed(2)}″ × ${inches(slide.buildHeightMm).toFixed(2)}″`}
                    sub={`${(inches(slide.buildWidthMm) * PT_PER_INCH).toFixed(0)} × ${(
                      inches(slide.buildHeightMm) * PT_PER_INCH
                    ).toFixed(0)} pt · ${toEmu(slide.buildWidthMm).toLocaleString('en-GB')} × ${toEmu(
                      slide.buildHeightMm,
                    ).toLocaleString('en-GB')} EMU`}
                  />
                  <Stat
                    label="Build scale"
                    value={slide.buildScale === 1 ? 'Full size' : `${slide.buildScale}×`}
                    tone={slide.buildScale === 1 ? undefined : 'warn'}
                  />
                  <Stat label="Export at" value={`${slide.buildDpi} dpi`} sub={`max ${slide.maxExportDpi} dpi`} />
                  <Stat
                    label="Type scale"
                    value={`${slide.typeScale.toFixed(2)}×`}
                    sub="Master text sizes are multiplied by this so type looks right on the slide"
                  />
                </div>
                <Problems items={slide.problems} />
                <p className="note">{exportNote(slide, wall)}</p>
              </Panel>

              <Panel title="Preview">
                <SlidePreview
                  slideCx={toEmu(slide.buildWidthMm)}
                  slideCy={toEmu(slide.buildHeightMm)}
                  wall={wall}
                  safeArea={safeArea}
                  brand={brand}
                />
              </Panel>

              <Panel
                title="Keynote"
                aside={<span className="chip chip--out">{KEYNOTE_MIN_PT}–{KEYNOTE_MAX_PT} pt</span>}
              >
                {keynote ? (
                  <>
                    <div className="stats">
                      <Stat
                        label="Document Setup"
                        tone={keynote.beatsPowerPoint ? 'accent' : undefined}
                        value={`${keynoteFieldText(keynote.buildWidthPt)} × ${keynoteFieldText(
                          keynote.buildHeightPt,
                        )} pt`}
                        sub={keynote.buildScale === 1 ? 'Full size' : `${keynote.buildScale}× — export to match`}
                      />
                    </div>
                    <Problems items={keynote.problems} />
                  </>
                ) : null}
                <p className="note">
                  A <code>.key</code> cannot be generated — since 2013 Keynote’s format has been
                  undocumented compressed protobuf. Open the <code>.pptx</code> above in Keynote
                  instead: it imports and keeps the slide size, which was checked on a real Keynote
                  rather than assumed. Or type the numbers above into Document Setup.
                </p>
              </Panel>
            </>
          ) : (
            <Panel title="The slide">
              <p className="note">Type a width and height to see the slide size.</p>
            </Panel>
          )}
        </div>
      </main>

      <footer className="foot">
        <span>
          Slide sizing from{' '}
          <a href="https://github.com/stoatworks-labs/aspect-calc">Aspect Calc</a>; ZIP writer from{' '}
          <a href="https://github.com/stoatworks-labs/test-card">Test Card</a>.
        </span>
        <span className="spacer" />
        <span className="dim">{typeof __APP_VERSION__ === 'string' ? __APP_VERSION__ : ''}</span>
      </footer>
    </div>
  )
}
