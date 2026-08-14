// Assembles taskpane.html: the actual UI is small (a button + status text),
// but it needs the SAME re-emittable library/app sources as converter.html
// so the file it exports is the same self-contained interactive page —
// unlike converter.html, none of these need to run live in this page itself
// (Excel already computed every formula; SheetJS/HyperFormula are only
// needed inside the exported viewer), so each is embedded inert exactly once.
const fs = require('fs');
const path = require('path');

const PROTOTYPE_DIR = path.join(__dirname, '..', 'browser-converter');
const CORE_DIR = path.join(PROTOTYPE_DIR, 'core');

const xlsxLib = fs.readFileSync(path.join(CORE_DIR, 'xlsx.core.min.js'), 'utf8');
const hfLib = fs.readFileSync(path.join(CORE_DIR, 'hyperformula.bundle.js'), 'utf8');
const renderCoreLib = fs.readFileSync(path.join(CORE_DIR, 'render-core.js'), 'utf8');
const appCss = fs.readFileSync(path.join(CORE_DIR, 'app.css'), 'utf8');
const appJs = fs.readFileSync(path.join(CORE_DIR, 'app.js'), 'utf8');
const taskpaneCss = fs.readFileSync(path.join(__dirname, 'taskpane.css'), 'utf8');
const taskpaneJs = fs.readFileSync(path.join(__dirname, 'taskpane.js'), 'utf8');

for (const [label, src] of [['xlsxLib', xlsxLib], ['hfLib', hfLib], ['renderCoreLib', renderCoreLib], ['appCss', appCss], ['appJs', appJs], ['taskpaneJs', taskpaneJs]]) {
  if (/<\/script/i.test(src)) throw new Error(`${label} contains a literal "</script" — cannot embed safely as-is.`);
}

const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Interactive Plan Exporter</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"><\/script>
<style>
${taskpaneCss}
</style>
</head>
<body>
  <h1>Interactive Plan Exporter</h1>
  <div class="subtitle">Reads this workbook and produces a self-contained interactive page — nothing is uploaded anywhere.</div>

  <div class="card">
    <button id="export-btn" class="btn" type="button" disabled>Export as Interactive Plan</button>
    <div id="status" class="status"></div>
  </div>

  <div class="notes">
    <ul>
      <li>Reads values, formulas, and number formats directly from this open workbook.</li>
      <li>Flags any formula the exported page can't safely recalculate (e.g. INDIRECT, Excel Tables) instead of guessing.</li>
      <li>Downloads one HTML file you can attach or embed on a Confluence or SharePoint page.</li>
    </ul>
  </div>

<script type="text/plain" id="lib-xlsx-source">${xlsxLib}</script>
<script type="text/plain" id="lib-hf-source">${hfLib}</script>
<script type="text/plain" id="render-core-source">${renderCoreLib}</script>
<script type="text/plain" id="app-css-source">${appCss}</script>
<script type="text/plain" id="app-js-source">${appJs}</script>

<script>
${taskpaneJs}
</script>
</body>
</html>
`;

const OUT_FILE = path.join(__dirname, 'taskpane.html');
fs.writeFileSync(OUT_FILE, html);
console.log(`Wrote ${OUT_FILE} (${(html.length / 1024 / 1024).toFixed(2)} MB)`);
