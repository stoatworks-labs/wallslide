import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'

import App from './App'
// Vendored from stoatworks-backend/theme by scripts/sync-theme.sh: loaded first so the tool's own CSS layers over it.
import './stoatworks-theme.css'
import './styles.css'
import { registerServiceWorker } from './register-sw'

const root = document.getElementById('root')
if (!root) throw new Error('no #root in the document')
createRoot(root).render(
  <StrictMode>
    <App />
  </StrictMode>,
)

registerServiceWorker()
