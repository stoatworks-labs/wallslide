/**
 * A thing the user should be told about a result.
 *
 * In aspect-calc this lives in `solve.ts` alongside the engine that raises most
 * of them. There is no equivalent engine here — `slides.ts` is the only module
 * that produces them — so it gets a file of its own rather than an import from
 * a solver this app does not have.
 */
export interface Problem {
  level: 'info' | 'warn' | 'error'
  text: string
}
