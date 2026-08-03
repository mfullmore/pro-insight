(function () {
  'use strict';

  function getSource(id) {
    return document.getElementById(id).textContent;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  }

  // A cell's text (vendor name, a note, anything) could contain a literal
  // script-closing tag — escaping every "<" keeps the JSON valid while
  // making it impossible for that text to break out of the script tag it's
  // embedded in, however the workbook was authored.
  function safeJsonForScriptTag(payload) {
    return JSON.stringify(payload).replace(/</g, '\\u003c');
  }

  function buildOutputHtml(payload) {
    const xlsxLib = getSource('lib-xlsx-source');
    const hfLib = getSource('lib-hf-source');
    const renderCoreLib = getSource('render-core-source');
    const appCss = getSource('app-css-source');
    const appJs = getSource('app-js-source');

    return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(payload.sourceFile)} — Interactive Plan</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
${appCss}
</style>
</head>
<body>
<div id="root"></div>
<script id="budget-data" type="application/json">${safeJsonForScriptTag(payload)}<\/script>
<script>
${xlsxLib}
<\/script>
<script>
${hfLib}
<\/script>
<script>
${renderCoreLib}
<\/script>
<script>
${appJs}
<\/script>
</body>
</html>
`;
  }

  // ---- UI wiring ----
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const statusEl = document.getElementById('status');
  const resultBox = document.getElementById('result');

  function setStatus(msg, isError) {
    statusEl.textContent = msg;
    statusEl.className = 'status' + (isError ? ' error' : '');
  }

  async function handleFile(file) {
    resultBox.innerHTML = '';
    if (!/\.xlsx$/i.test(file.name)) {
      setStatus('Please choose a .xlsx file.', true);
      return;
    }
    setStatus(`Reading ${file.name}…`, false);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true });
      const payload = InteractivePlanConvert.convertWorkbook(wb, file.name);
      const html = buildOutputHtml(payload);
      const outName = file.name.replace(/\.xlsx$/i, '') + '-interactive.html';

      const blob = new Blob([html], { type: 'text/html' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = outName;
      link.className = 'btn';
      link.textContent = `Download ${outName}`;
      resultBox.appendChild(link);

      const note = document.createElement('div');
      note.className = 'result-note';
      const n = payload.unsupportedFormulas.length;
      note.textContent = n
        ? `Converted with ${n} formula${n === 1 ? '' : 's'} flagged as unsupported — the generated page will show exactly which ones and why.`
        : 'Converted cleanly — no unsupported formulas detected.';
      resultBox.appendChild(note);

      setStatus(`Done (${(html.length / 1024 / 1024).toFixed(2)} MB). Nothing was uploaded anywhere — this ran entirely in your browser.`, false);
    } catch (err) {
      console.error(err);
      setStatus('Conversion failed: ' + err.message, true);
    }
  }

  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });
  ['dragenter', 'dragover'].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add('drag'); })
  );
  ['dragleave', 'drop'].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove('drag'); })
  );
  dropzone.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
})();
