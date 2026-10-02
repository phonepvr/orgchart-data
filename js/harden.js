// Prototype-pollution hardening for the vendored SheetJS (xlsx@0.18.5).
// Freezes the prototype chain so a crafted XLSX cannot inject inherited
// properties at runtime. Loaded as the first script, before the XLSX vendor
// file and before any user data is parsed.
try {
  Object.freeze(Object.prototype);
  Object.freeze(Array.prototype);
  Object.freeze(Function.prototype);
} catch (e) { /* environments without configurable prototype - no-op */ }
