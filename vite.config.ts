import { defineConfig, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'

import { readFileSync } from 'node:fs'

const pkg = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'))

/**
 * Stamp the version this build produced onto the support-footer script tag.
 *
 * Same rationale as aspect-calc: a literal in index.html goes stale the moment a
 * release is tagged, and a feedback report naming the wrong build is worse than
 * one naming no build at all.
 */
function supportFooterVersion(): Plugin {
  const tag = /<script\s[^>]*\bsrc="[^"]*support-footer\.js"/
  return {
    name: 'stoatworks-support-footer-version',
    transformIndexHtml: {
      order: 'post',
      handler(html: string) {
        if (!tag.test(html)) {
          throw new Error('no support-footer.js tag in index.html — nothing to stamp')
        }
        return html.replace(tag, (m) => `${m} data-version="v${pkg.version}"`)
      },
    },
  }
}

// Static SPA, no backend. dist/ is what the Cloudflare Worker serves as assets.
export default defineConfig({
  define: { __APP_VERSION__: JSON.stringify(`v${pkg.version}`) },
  plugins: [react(), supportFooterVersion()],
  base: './',
  build: {
    outDir: 'dist',
    sourcemap: true,
  },
})
