// NOTE ON CONFIDENCE: the Excel JS API calls below (Excel.run, getUsedRangeOrNullObject,
// range.values/formulas/numberFormat, workbook.names) are written against documented,
// stable Excel JS API surface, but this file has NOT been run inside Excel — there was
// no Office install available to sideload and verify against. Treat this as a reviewed
// first draft, not a tested implementation. Likely places to double-check first:
//   - getUsedRangeOrNullObject's exact null-handling behavior on a truly empty sheet
//   - whether NamedItem.formula always starts with "=" the way workbook Names.Ref does in SheetJS
//   - Office.context.document.url availability before the workbook has ever been saved

(function () {
  'use strict';

  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      document.getElementById('export-btn').disabled = false;
      document.getElementById('export-btn').addEventListener('click', runExport);
    } else {
      setStatus('This add-in only runs inside Excel.', 'bad');
    }
  });

  // Same detection logic as converter.js — flags formula patterns known (from
  // fidelity testing against the browser-upload tool) not to survive the
  // HyperFormula recalculation engine used by the exported viewer page.
  const UNSUPPORTED_PATTERNS = [
    { test: /\bINDIRECT\s*\(/i, reason: "Uses INDIRECT() to look up a value dynamically — this tool can't follow dynamic references yet." },
    { test: /[A-Za-z_][A-Za-z0-9_.]*\[[^\]]+\]/, reason: "Uses an Excel Table reference (like Table1[Column]) — this tool doesn't support Excel Tables yet." },
  ];

  function colLetter(n) {
    let s = '', i = n + 1;
    while (i > 0) {
      const rem = (i - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  function setStatus(msg, kind) {
    const el = document.getElementById('status');
    el.textContent = msg;
    el.className = 'status visible' + (kind ? ' ' + kind : '');
  }

  // ---- Extraction: reads the live workbook via the Excel JS API. Excel has
  // already computed every formula, so — unlike the browser-upload tool —
  // this never has to re-derive a value the recalculation engine can't
  // produce; it only needs to re-derive values AFTER a viewer-side edit. ----
  async function readWorkbook() {
    return await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items/name');
      await context.sync();

      const sheetNames = worksheets.items.map((s) => s.name);
      const rangeBySheet = {};
      worksheets.items.forEach((sheet) => {
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load(['values', 'formulas', 'numberFormat', 'rowCount', 'columnCount', 'isNullObject']);
        rangeBySheet[sheet.name] = used;
      });
      await context.sync();

      const sheetsData = {};
      const formulaMap = {};
      const formatMap = {};
      const cachedValues = {};
      const dims = {};

      for (const name of sheetNames) {
        const range = rangeBySheet[name];
        formulaMap[name] = {};
        formatMap[name] = {};
        cachedValues[name] = {};

        if (range.isNullObject || range.rowCount === 0 || range.columnCount === 0) {
          sheetsData[name] = [];
          dims[name] = { rows: 0, cols: 0 };
          continue;
        }

        const grid = [];
        for (let r = 0; r < range.rowCount; r++) {
          const row = [];
          for (let c = 0; c < range.columnCount; c++) {
            const formulaCell = range.formulas[r][c];
            const valueCell = range.values[r][c];
            const fmt = range.numberFormat[r][c];
            if (fmt && fmt !== 'General') formatMap[name][`${r},${c}`] = fmt;

            const isFormula = typeof formulaCell === 'string' && formulaCell.charAt(0) === '=';
            if (isFormula) {
              const formulaText = formulaCell.slice(1);
              formulaMap[name][`${r},${c}`] = formulaText;
              // Excel's own last-computed value, kept so a formula this tool's
              // engine can't evaluate can still show something real (frozen)
              // instead of a bare error.
              cachedValues[name][`${r},${c}`] = valueCell;
              row.push('=' + formulaText);
            } else if (valueCell === '' || valueCell === null || valueCell === undefined) {
              row.push(null);
            } else {
              row.push(valueCell);
            }
          }
          grid.push(row);
        }
        sheetsData[name] = grid;
        dims[name] = { rows: range.rowCount, cols: range.columnCount };
      }

      // Workbook-scoped named ranges/expressions. Not filtered by NamedItem.type
      // since the exact enum values weren't verifiable here — every name is
      // passed through, and app.js already tolerates names that fail to
      // register (wraps hf.addNamedExpression in try/catch).
      const names = context.workbook.names;
      names.load('items/name,items/formula');
      await context.sync();
      const definedNames = names.items.map((n) => ({ name: n.name, ref: String(n.formula).replace(/^=/, '') }));

      return { sheetNames, sheetsData, formulaMap, formatMap, cachedValues, dims, definedNames };
    });
  }

  function detectUnsupported(sheetNames, formulaMap) {
    const unsupported = [];
    for (const name of sheetNames) {
      for (const [key, formula] of Object.entries(formulaMap[name])) {
        const match = UNSUPPORTED_PATTERNS.find((p) => p.test.test(formula));
        if (match) {
          const [r, c] = key.split(',').map(Number);
          unsupported.push({ sheet: name, row: r, col: c, addr: colLetter(c) + (r + 1), formula, reason: match.reason });
        }
      }
    }
    return unsupported;
  }

  function getWorkbookFileName() {
    try {
      const url = Office.context.document.url || '';
      const base = url.split(/[\\/]/).pop();
      return base && base.length ? base : 'workbook.xlsx';
    } catch (err) {
      return 'workbook.xlsx';
    }
  }

  function getSource(id) {
    return document.getElementById(id).textContent;
  }

  // A cell's text could contain a literal script-closing tag; escaping "<"
  // keeps the JSON valid while making it impossible for that text to break
  // out of the script tag it's embedded in.
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
<title>${payload.sourceFile.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))} — Interactive Plan</title>
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

  function downloadFile(html, filename) {
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  async function runExport() {
    const btn = document.getElementById('export-btn');
    btn.disabled = true;
    setStatus('Reading workbook…');
    try {
      const fileName = getWorkbookFileName();
      const extracted = await readWorkbook();
      const unsupportedFormulas = detectUnsupported(extracted.sheetNames, extracted.formulaMap);
      const payload = Object.assign({}, extracted, { unsupportedFormulas, sourceFile: fileName });
      const html = buildOutputHtml(payload);
      const outName = fileName.replace(/\.xlsx$/i, '') + '-interactive.html';
      downloadFile(html, outName);
      const n = unsupportedFormulas.length;
      setStatus(
        n ? `Exported (${(html.length / 1024 / 1024).toFixed(2)} MB). ${n} formula${n === 1 ? '' : 's'} flagged as unsupported — the exported page will show exactly which.`
          : `Exported cleanly (${(html.length / 1024 / 1024).toFixed(2)} MB). No unsupported formulas detected.`,
        'good'
      );
    } catch (err) {
      console.error(err);
      setStatus('Export failed: ' + (err && err.message ? err.message : String(err)), 'bad');
    } finally {
      btn.disabled = false;
    }
  }
})();
