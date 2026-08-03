// Vendors the shared libraries into the Custom UI static resource folder.
// Unlike the browser-converter and office-addin apps, this page never
// generates a separate downloadable HTML file — the Custom UI iframe IS the
// interactive viewer, rendered directly on the Confluence page — so
// SheetJS and HyperFormula are loaded live here, not embedded-then-re-emitted.
const fs = require('fs');
const path = require('path');

const SIBLING_DIR = path.join(__dirname, '..', 'browser-converter');
const CORE_DIR = path.join(SIBLING_DIR, 'core');
const VENDOR_DIR = path.join(__dirname, 'static', 'main', 'vendor');

fs.mkdirSync(VENDOR_DIR, { recursive: true });

fs.copyFileSync(
  path.join(CORE_DIR, 'xlsx.core.min.js'),
  path.join(VENDOR_DIR, 'xlsx.core.min.js')
);
fs.copyFileSync(
  path.join(CORE_DIR, 'hyperformula.bundle.js'),
  path.join(VENDOR_DIR, 'hyperformula.bundle.js')
);
fs.copyFileSync(
  path.join(CORE_DIR, 'convert-core.js'),
  path.join(VENDOR_DIR, 'convert-core.js')
);
fs.copyFileSync(
  path.join(CORE_DIR, 'render-core.js'),
  path.join(VENDOR_DIR, 'render-core.js')
);

console.log('Vendored xlsx.core.min.js, hyperformula.bundle.js, convert-core.js, and render-core.js into static/main/vendor/');
console.log('(forge-bridge.bundle.js is built separately via esbuild — see package.json)');
