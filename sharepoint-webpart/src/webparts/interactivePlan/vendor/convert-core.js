// Shared workbook-parsing logic — the single place all four apps' SheetJS-based
// extraction goes through (office-addin uses Excel's own API instead, so it
// doesn't need this). UMD-style: works as a plain <script> global (browser
// tools, and side-effect-imported into the SharePoint webpack build) or via
// require('./convert-core.js') from Node (generate-html.js).
//
// Browser is checked first (not "has module.exports", the usual UMD check)
// because a webpack bundle — SharePoint's build, when this file is pulled in
// as a side-effect import — always provides a real `module` object even
// though it's building for the browser, so a module.exports-first check
// would wrongly take the Node branch there. Worse, that branch's
// require('xlsx') would then fail to bundle: 'xlsx' isn't a real dependency
// of the SharePoint project (only vendored as a raw file), so webpack can't
// resolve it — and webpack statically detects (and tries to resolve) even
// an aliased `var r = require; r('xlsx')`. Routing it through eval is the
// one form its static analyzer genuinely can't see into, while still
// behaving as a normal require() at runtime in real Node.
(function (root, factory) {
  if (typeof window !== 'undefined') {
    root.InteractivePlanConvert = factory(root.XLSX);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(eval('require')('xlsx'));
  }
})(typeof self !== 'undefined' ? self : this, function (XLSX) {
  'use strict';

  // No assumptions about sheet names, column meaning, or layout — a cell is
  // editable if it's a plain number with no formula, read-only+recalculated
  // if it has a formula, and a label otherwise. That's the only distinction
  // any consumer relies on, so it works on any workbook, not just one template.
  const UNSUPPORTED_PATTERNS = [
    { test: /\bINDIRECT\s*\(/i, reason: "Uses INDIRECT() to look up a value dynamically — this tool can't follow dynamic references yet." },
    { test: /[A-Za-z_][A-Za-z0-9_.]*\[[^\]]+\]/, reason: "Uses an Excel Table reference (like Table1[Column]) — this tool doesn't support Excel Tables yet." },
  ];

  function convertWorkbook(wb, sourceFileName) {
    const sheetNames = wb.SheetNames;
    const sheetsData = {};
    const formulaMap = {}; // sheet -> "r,c" -> formula string (no leading '=')
    const formatMap = {}; // sheet -> "r,c" -> Excel number-format code (e.g. "0%", "$#,##0")
    const cachedValues = {}; // sheet -> "r,c" -> Excel's last-computed value for formula cells
    const dims = {}; // sheet -> {rows, cols}

    for (const name of sheetNames) {
      const ws = wb.Sheets[name];
      formulaMap[name] = {};
      formatMap[name] = {};
      cachedValues[name] = {};
      if (!ws['!ref']) { sheetsData[name] = []; dims[name] = { rows: 0, cols: 0 }; continue; }
      const range = XLSX.utils.decode_range(ws['!ref']);
      const grid = [];
      for (let r = range.s.r; r <= range.e.r; r++) {
        const row = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = ws[addr];
          if (!cell) { row.push(null); continue; }
          if (cell.z && cell.z !== 'General') formatMap[name][`${r},${c}`] = cell.z;
          if (cell.f) {
            formulaMap[name][`${r},${c}`] = cell.f;
            // Excel's own last-computed value, kept so a formula this tool's
            // engine can't evaluate can still show something real (frozen)
            // instead of a bare error.
            cachedValues[name][`${r},${c}`] = cell.v;
            row.push('=' + cell.f);
          } else {
            row.push(cell.v);
          }
        }
        grid.push(row);
      }
      sheetsData[name] = grid;
      dims[name] = { rows: range.e.r - range.s.r + 1, cols: range.e.c - range.s.c + 1 };
    }

    const rawDefinedNames = (wb.Workbook && wb.Workbook.Names) || [];

    const unsupportedFormulas = [];
    for (const name of sheetNames) {
      for (const key of Object.keys(formulaMap[name])) {
        const formula = formulaMap[name][key];
        const match = UNSUPPORTED_PATTERNS.find((p) => p.test.test(formula));
        if (match) {
          const parts = key.split(',');
          const r = Number(parts[0]), c = Number(parts[1]);
          unsupportedFormulas.push({ sheet: name, row: r, col: c, addr: XLSX.utils.encode_cell({ r, c }), formula, reason: match.reason });
        }
      }
    }

    return {
      sheetNames, sheetsData, formulaMap, formatMap, cachedValues, dims,
      definedNames: rawDefinedNames.map((n) => ({ name: n.Name, ref: n.Ref })),
      unsupportedFormulas,
      sourceFile: sourceFileName,
    };
  }

  return { convertWorkbook, UNSUPPORTED_PATTERNS };
});
