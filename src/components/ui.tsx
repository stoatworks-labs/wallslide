import type { ReactNode } from 'react'
import type { Problem } from '../lib/problem'

export function Segmented<T extends string>({
  value,
  options,
  onChange,
  label,
}: {
  value: T
  options: { id: T; label: string; title?: string }[]
  onChange: (v: T) => void
  label?: string
}) {
  return (
    <div className="segmented" role="group" aria-label={label}>
      {options.map((o) => (
        <button
          key={o.id}
          type="button"
          title={o.title}
          className={o.id === value ? 'seg on' : 'seg'}
          aria-pressed={o.id === value}
          onClick={() => onChange(o.id)}
        >
          {o.label}
        </button>
      ))}
    </div>
  )
}

export function Field({
  label,
  value,
  onChange,
  suffix,
  placeholder,
  invalid,
  hint,
  inputMode = 'decimal',
}: {
  label: string
  value: string
  onChange?: (v: string) => void
  suffix?: ReactNode
  placeholder?: string
  invalid?: boolean
  hint?: string
  inputMode?: 'decimal' | 'numeric' | 'text'
}) {
  return (
    <label className="field">
      <span className="field__label">{label}</span>
      <span className="field__box">
        <input
          className={invalid ? 'input bad' : 'input'}
          value={value}
          placeholder={placeholder}
          inputMode={inputMode}
          onChange={(e) => onChange?.(e.target.value)}
          onFocus={(e) => e.currentTarget.select()}
        />
        {suffix ? <span className="field__suffix">{suffix}</span> : null}
      </span>
      {hint ? <span className="field__hint">{hint}</span> : null}
    </label>
  )
}

export function Panel({
  title,
  children,
  aside,
}: {
  title: string
  children: ReactNode
  aside?: ReactNode
}) {
  return (
    <section className="panel">
      <header className="panel__head">
        <h2>{title}</h2>
        {aside}
      </header>
      <div className="panel__body">{children}</div>
    </section>
  )
}

export function Stat({
  label,
  value,
  sub,
  wide,
  tone,
}: {
  label: string
  value: ReactNode
  sub?: ReactNode
  wide?: boolean
  tone?: 'accent' | 'warn'
}) {
  return (
    <div className={`stat${wide ? ' stat--wide' : ''}${tone ? ` stat--${tone}` : ''}`}>
      <div className="stat__label">{label}</div>
      <div className="stat__value">{value}</div>
      {sub ? <div className="stat__sub">{sub}</div> : null}
    </div>
  )
}

export function Problems({ items }: { items: Problem[] }) {
  if (!items.length) return null
  return (
    <ul className="problems">
      {items.map((p, i) => (
        <li key={i} className={`problem problem--${p.level}`}>
          {p.text}
        </li>
      ))}
    </ul>
  )
}

export function Toggle({
  label,
  checked,
  onChange,
  hint,
}: {
  label: string
  checked: boolean
  onChange: (v: boolean) => void
  hint?: string
}) {
  return (
    <label className="toggle">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span>
        {label}
        {hint ? <em className="toggle__hint">{hint}</em> : null}
      </span>
    </label>
  )
}

export function Swatch({
  label,
  value,
  onChange,
}: {
  label: string
  value: string
  onChange: (hex: string) => void
}) {
  return (
    <label className="swatch">
      <input
        type="color"
        value={`#${value}`}
        onChange={(e) => onChange(e.target.value.replace('#', '').toUpperCase())}
        aria-label={label}
      />
      <span className="swatch__label">{label}</span>
      <span className="swatch__hex">{value}</span>
    </label>
  )
}
