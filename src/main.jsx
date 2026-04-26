import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App.jsx'
import './index.css'

// Prototype-pollution hardening for the bundled SheetJS (xlsx@0.18.5).
// Freezes the prototype chain so a crafted XLSX cannot inject inherited
// properties at runtime. Must run before any user data is parsed.
try {
  Object.freeze(Object.prototype);
  Object.freeze(Array.prototype);
  Object.freeze(Function.prototype);
} catch { /* environments without configurable prototype - no-op */ }

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
