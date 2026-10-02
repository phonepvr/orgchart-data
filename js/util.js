(function () {
'use strict';
const OS = window.OrgSense = window.OrgSense || {};
// Small shared helpers for building HTML strings safely.

const ESC_MAP = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

// Escape a value for interpolation into HTML text or attribute values.
const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => ESC_MAP[c]);

Object.assign(OS, { esc });
})();
