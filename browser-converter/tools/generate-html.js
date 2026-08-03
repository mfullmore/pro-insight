// Converts budget-test.xlsx into a single self-contained interactive HTML
// file: embeds the parsed sheet data + formulas, inlines SheetJS and
// HyperFormula so recalculation happens entirely client-side (no server,
// no network calls — this is the constraint that lets it live on a locked-
// down Confluence/SharePoint page), and wires up live editing + a
// "save as new version" download that writes back a real .xlsx.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { convertWorkbook } = require('../core/convert-core.js');

const ROOT = path.join(__dirname, '..');
const CORE = path.join(ROOT, 'core');
const SRC_FILE = path.join(ROOT, 'fixtures', 'budget-test-3.xlsx');
const SRC_FILE_NAME = path.basename(SRC_FILE);
const OUT_FILE = path.join(ROOT, 'dist', 'budget-prototype.html');

const wb = XLSX.readFile(SRC_FILE, { cellFormula: true, cellNF: true });
const payload = convertWorkbook(wb, SRC_FILE_NAME);
console.log(`Detected ${payload.unsupportedFormulas.length} unsupported formula(s):`, payload.unsupportedFormulas);

// A cell's text could contain literal "</script>"; escaping "<" keeps the
// JSON valid while making it impossible for that text to break out of the
// script tag it's embedded in.
function safeJsonForScriptTag(payload) {
  return JSON.stringify(payload).replace(/</g, '\\u003c');
}

const xlsxLib = fs.readFileSync(path.join(CORE, 'xlsx.core.min.js'), 'utf8');
const hfLib = fs.readFileSync(path.join(CORE, 'hyperformula.bundle.js'), 'utf8');
const renderCoreLib = fs.readFileSync(path.join(CORE, 'render-core.js'), 'utf8');
const appCss = fs.readFileSync(path.join(CORE, 'app.css'), 'utf8');
const appJs = fs.readFileSync(path.join(CORE, 'app.js'), 'utf8');

const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>${SRC_FILE_NAME} — Interactive Plan</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
${appCss}
</style>
</head>
<body>
<div id="root"></div>
<script id="budget-data" type="application/json">${safeJsonForScriptTag(payload)}</script>
<script>
${xlsxLib}
</script>
<script>
${hfLib}
</script>
<script>
${renderCoreLib}
</script>
<script>
${appJs}
</script>
</body>
</html>
`;

fs.writeFileSync(OUT_FILE, html);
console.log(`Wrote ${OUT_FILE} (${(html.length / 1024 / 1024).toFixed(2)} MB)`);

// Artifact-preview variant: same content, no outer html/head/body — the
// Artifact host wraps that itself. Not the product deliverable, just a
// convenient way to interact with this prototype live in the conversation.
const previewBody = `<title>${SRC_FILE_NAME} — Interactive Plan</title>
<style>
${appCss}
</style>
<div id="root"></div>
<script id="budget-data" type="application/json">${safeJsonForScriptTag(payload)}</script>
<script>
${xlsxLib}
</script>
<script>
${hfLib}
</script>
<script>
${renderCoreLib}
</script>
<script>
${appJs}
</script>
`;
const PREVIEW_FILE = path.join(ROOT, 'dist', 'artifact-preview.html');
fs.writeFileSync(PREVIEW_FILE, previewBody);
console.log(`Wrote ${PREVIEW_FILE}`);
