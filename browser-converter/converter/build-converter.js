// Builds converter.html: a single self-contained page that IS the "option 1"
// authoring tool — no extension, no install. A department head drags their
// .xlsx onto it, conversion runs entirely client-side (SheetJS parse ->
// same department/formula analysis as generate-html.js), and it downloads
// a ready-to-embed interactive HTML file. The output file's own libraries
// (SheetJS, HyperFormula) and app code are embedded here twice: once live
// (so this page can parse the upload) and once inert (so this page can
// re-emit them into the file it generates) — see converter.js.
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const CORE = path.join(ROOT, 'core');

const xlsxLib = fs.readFileSync(path.join(CORE, 'xlsx.core.min.js'), 'utf8');
const hfLib = fs.readFileSync(path.join(CORE, 'hyperformula.bundle.js'), 'utf8');
const convertCoreLib = fs.readFileSync(path.join(CORE, 'convert-core.js'), 'utf8');
const renderCoreLib = fs.readFileSync(path.join(CORE, 'render-core.js'), 'utf8');
const appCss = fs.readFileSync(path.join(CORE, 'app.css'), 'utf8');
const appJs = fs.readFileSync(path.join(CORE, 'app.js'), 'utf8');
const converterCss = fs.readFileSync(path.join(__dirname, 'converter.css'), 'utf8');
const converterJs = fs.readFileSync(path.join(__dirname, 'converter.js'), 'utf8');

// Sanity check: none of the re-emitted blobs may contain a literal
// "</script" — that would terminate their <script type="text/plain">
// holder early and corrupt everything after it in this page.
for (const [label, src] of [['xlsxLib', xlsxLib], ['hfLib', hfLib], ['convertCoreLib', convertCoreLib], ['renderCoreLib', renderCoreLib], ['appCss', appCss], ['appJs', appJs], ['converterJs', converterJs]]) {
  if (/<\/script/i.test(src)) throw new Error(`${label} contains a literal "</script" — cannot embed safely as-is.`);
}

const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Budget Plan Converter</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
${converterCss}
</style>
</head>
<body>
<div id="page">
  <h1>Budget Plan Converter</h1>
  <div class="subtitle">Turn a department budget spreadsheet into an interactive page. Runs entirely in your browser — nothing is uploaded anywhere.</div>

  <div id="dropzone" tabindex="0" role="button" aria-label="Choose or drop a budget spreadsheet">
    <div class="dropzone-title">Drop any budget .xlsx here, or click to browse</div>
    <div class="dropzone-sub">Works on any workbook — no required sheet names or layout</div>
  </div>
  <input id="file-input" type="file" accept=".xlsx" style="display:none" />
  <div id="status" class="status"></div>
  <div id="result"></div>

  <div class="notes">
    <h2>What this does</h2>
    <ul>
      <li>Parses your workbook's formulas and data locally — the file never leaves this browser tab.</li>
      <li>Any cell with a formula stays read-only and recalculates live; any plain number becomes editable — nothing about your layout is assumed.</li>
      <li>Flags any formula it can't safely convert (e.g. INDIRECT, Excel Tables) instead of guessing.</li>
      <li>Produces one self-contained HTML file you can attach or embed on a Confluence or SharePoint page.</li>
    </ul>
  </div>
</div>

<script type="text/plain" id="lib-xlsx-source">${xlsxLib}</script>
<script type="text/plain" id="lib-hf-source">${hfLib}</script>
<script type="text/plain" id="render-core-source">${renderCoreLib}</script>
<script type="text/plain" id="app-css-source">${appCss}</script>
<script type="text/plain" id="app-js-source">${appJs}</script>

<script>
${xlsxLib}
</script>
<script>
${convertCoreLib}
</script>
<script>
${converterJs}
</script>
</body>
</html>
`;

const OUT_FILE = path.join(ROOT, 'converter.html');
fs.writeFileSync(OUT_FILE, html);
console.log(`Wrote ${OUT_FILE} (${(html.length / 1024 / 1024).toFixed(2)} MB)`);
