// Shared rendering logic — the single place all three DOM-based apps
// (browser-converter, confluence-forge-app, sharepoint-webpart) build their
// HyperFormula engine and draw the tabbed grid from. office-addin reuses this
// too, via the same generated-output pipeline as browser-converter. Browser
// global only — no consumer needs this from Node, so unlike convert-core.js
// this isn't UMD-wrapped (SPFx's webpack bundle also counts as "browser":
// it always provides a `module` global, so a Node-style `typeof module ===
// 'object'` UMD check would wrongly take the CommonJS branch there).
(function (root, factory) {
  root.InteractivePlanRender = factory(root.XLSX, root.HyperFormula);
})(typeof self !== 'undefined' ? self : this, function (XLSX, HyperFormula) {
  'use strict';

  const isErr = (v) => typeof v === 'string' && v.startsWith('#');

  function colLetter(n) {
    let s = '', i = n + 1;
    while (i > 0) {
      const rem = (i - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    if (attrs) for (const k of Object.keys(attrs)) {
      const v = attrs[k];
      if (k === 'text') node.textContent = v;
      else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
      else node.setAttribute(k, v);
    }
    (children || []).forEach((c) => c && node.appendChild(c));
    return node;
  }

  // ---- Recalculation engine, built fresh per workbook ----
  function createEngine(payload) {
    const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
    hf.addNamedExpression('TRUE', true);
    hf.addNamedExpression('FALSE', false);

    const sheetIdByName = {};
    for (const name of payload.sheetNames) {
      hf.addSheet(name);
      sheetIdByName[name] = hf.getSheetId(name);
    }
    for (const name of payload.sheetNames) {
      hf.setSheetContent(sheetIdByName[name], payload.sheetsData[name]);
    }
    for (const dn of payload.definedNames) {
      try {
        hf.addNamedExpression(dn.name, '=' + dn.ref.replace(/^=/, ''));
      } catch (err) {
        console.warn('named expression failed:', dn.name, err.message);
      }
    }

    function getVal(sheet, row, col) {
      // HyperFormula normally reports a bad formula (e.g. INDIRECT) as an
      // error VALUE object, which the frozen-value fallback below handles —
      // but for some formulas it throws a real JS exception instead. Treating
      // that the same way (as an error string) keeps one bad cell from
      // aborting the whole grid render.
      let v;
      try {
        v = hf.getCellValue({ sheet: sheetIdByName[sheet], row, col });
      } catch (err) {
        return '#ERROR!';
      }
      if (v && typeof v === 'object' && 'value' in v) return v.value; // HF error object
      return v;
    }
    function setVal(sheet, row, col, value) {
      hf.setCellContents({ sheet: sheetIdByName[sheet], row, col }, [[value]]);
    }
    function hasFormula(sheet, row, col) {
      return Object.prototype.hasOwnProperty.call(payload.formulaMap[sheet], `${row},${col}`);
    }
    function formatOf(sheet, row, col) {
      return payload.formatMap[sheet][`${row},${col}`];
    }
    function fmtCell(sheet, row, col, value) {
      if (isErr(value)) return value;
      if (typeof value !== 'number') return value == null ? '' : String(value);
      const fmt = formatOf(sheet, row, col);
      if (fmt) {
        try { return XLSX.SSF.format(fmt, value); } catch (err) { /* fall through */ }
      }
      return Number.isInteger(value) ? value.toLocaleString('en-US') : value.toLocaleString('en-US', { maximumFractionDigits: 4 });
    }

    return { getVal, setVal, hasFormula, formatOf, fmtCell };
  }

  function cellType(v) {
    if (typeof v === 'boolean') return 'b';
    if (typeof v === 'number') return 'n';
    return 's';
  }

  // Rebuilds every sheet from the engine's current (possibly edited) values
  // into a real .xlsx, formulas and formats intact. Returns an ArrayBuffer —
  // what a consumer does with those bytes (trigger a browser download,
  // base64-encode for a Forge invoke(), PUT to a SharePoint REST endpoint) is
  // environment-specific and left to the caller.
  function buildWorkbookBytes(payload, engine, opts) {
    const newWb = XLSX.utils.book_new();

    for (const name of payload.sheetNames) {
      const grid = payload.sheetsData[name];
      const ws = {};
      let maxR = 0, maxC = 0;
      for (let r = 0; r < grid.length; r++) {
        for (let c = 0; c < grid[r].length; c++) {
          if (grid[r][c] === null || grid[r][c] === undefined) continue;
          const value = engine.getVal(name, r, c);
          const formula = payload.formulaMap[name][`${r},${c}`];
          const fmt = payload.formatMap[name][`${r},${c}`];
          const addr = XLSX.utils.encode_cell({ r, c });
          const cellObj = formula ? { t: cellType(value), v: value, f: formula } : { t: cellType(value), v: value };
          if (fmt) cellObj.z = fmt;
          ws[addr] = cellObj;
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        }
      }
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
      XLSX.utils.book_append_sheet(newWb, ws, name);
    }

    if (opts && opts.versionInfo) {
      const { author, stamp, basedOn, note } = opts.versionInfo;
      const infoSheet = XLSX.utils.aoa_to_sheet([
        ['Version Info'],
        ['Author', author],
        ['Saved', stamp],
        ['Based on', basedOn],
        ['Note', note || 'Formulas preserved; values reflect edits made in the interactive plan.'],
      ]);
      XLSX.utils.book_append_sheet(newWb, infoSheet, 'Version Info');
    }

    return XLSX.write(newWb, { type: 'array', bookType: 'xlsx' });
  }

  // ---- Rendering ----
  // opts: { root, payload, engine, sourceLine, footerText, onSave, saveLabel }
  // onSave is optional — omit it to render without a save button (a consumer
  // not ready to save yet, or one that puts the button elsewhere).
  //
  // Deliberately generic, same as convert-core.js: no assumption about sheet
  // names, column meaning, or layout. The sidebar shows sheet names plus how
  // many formulas are flagged per sheet (always knowable, always accurate) —
  // never a computed amount, since summing arbitrary cells isn't reliably
  // meaningful on a workbook this code doesn't understand the shape of.
  function renderApp(opts) {
    const { root, payload, engine } = opts;
    root.innerHTML = '';

    const app = el('div', { class: 'app' });
    const rail = el('aside', { class: 'rail' });
    const content = el('main', { class: 'content' });
    app.appendChild(rail);
    app.appendChild(content);
    root.appendChild(app);

    rail.appendChild(el('div', { class: 'rail-head' }, [
      el('h1', { text: 'Interactive Plan' }),
      el('div', { class: 'source', text: opts.sourceLine || `Source: ${payload.sourceFile} · recalculated live in your browser` }),
    ]));

    const unsupported = payload.unsupportedFormulas || [];
    const unsupportedByAddr = new Map();
    unsupported.forEach((u) => unsupportedByAddr.set(`${u.sheet}|${u.row},${u.col}`, u));
    const unsupportedBySheet = new Map();
    unsupported.forEach((u) => unsupportedBySheet.set(u.sheet, (unsupportedBySheet.get(u.sheet) || 0) + 1));

    if (unsupported.length) {
      const n = unsupported.length;
      const details = el('details', { class: 'flag-disclosure' });
      details.appendChild(el('summary', { text: `${n} formula${n === 1 ? '' : 's'} flagged` }));
      const body = el('div', { class: 'flag-detail' });
      body.appendChild(el('div', {
        class: 'flag-detail-intro',
        text: "Each cell below shows Excel's last saved value, marked as frozen (✱). Anything that depends on it still shows as an error until the source formula is fixed.",
      }));
      const list = el('ul', { class: 'flag-list' });
      unsupported.forEach((u) => {
        list.appendChild(el('li', {}, [
          el('span', { class: 'num flag-loc', text: `${u.sheet}!${u.addr}` }),
          el('span', { text: ' — ' + u.reason }),
        ]));
      });
      body.appendChild(list);
      details.appendChild(body);
      rail.appendChild(details);
    }

    const nav = el('nav', { class: 'deptnav' });
    rail.appendChild(nav);

    const state = { activeSheet: payload.sheetNames[0] };

    function renderNav() {
      nav.innerHTML = '';
      payload.sheetNames.forEach((name) => {
        const flagCount = unsupportedBySheet.get(name) || 0;
        nav.appendChild(el('button', {
          type: 'button',
          class: 'nav-item' + (name === state.activeSheet ? ' active' : ''),
          onclick: () => { state.activeSheet = name; renderNav(); renderPanel(); },
        }, [
          el('span', { class: 'nav-row' }, [
            el('span', { class: 'name', text: name }),
            flagCount ? el('span', {
              class: 'nav-flag',
              title: `${flagCount} formula${flagCount === 1 ? '' : 's'} flagged in this sheet`,
              text: String(flagCount),
            }) : null,
          ]),
        ]));
      });
    }

    rail.appendChild(el('div', { class: 'rail-spacer' }));
    if (opts.onSave) {
      rail.appendChild(el('button', { class: 'btn', type: 'button', text: opts.saveLabel || 'Save as new version…', onclick: opts.onSave }));
    }

    const panelWrap = el('div', { class: 'panel-wrap' });
    content.appendChild(panelWrap);

    function renderPanel() {
      panelWrap.innerHTML = '';
      const activeSheet = state.activeSheet;
      const { rows, cols } = payload.dims[activeSheet] || { rows: 0, cols: 0 };

      panelWrap.appendChild(el('div', { class: 'panel-head' }, [
        el('h2', { text: activeSheet }),
        el('div', { class: 'panel-meta', text: rows && cols ? `${rows} row${rows === 1 ? '' : 's'} × ${cols} column${cols === 1 ? '' : 's'}` : '' }),
      ]));

      const card = el('div', { class: 'card' });
      if (rows === 0 || cols === 0) {
        card.appendChild(el('div', { class: 'empty-sheet', text: 'This sheet is empty.' }));
        panelWrap.appendChild(card);
        return;
      }

      const scroll = el('div', { class: 'table-scroll' });
      const table = el('table', { class: 'grid-table' });

      const thead = el('thead');
      const headRow = el('tr');
      headRow.appendChild(el('th', { class: 'gutter' }));
      for (let c = 0; c < cols; c++) headRow.appendChild(el('th', { class: 'num col-letter', text: colLetter(c) }));
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = el('tbody');
      for (let r = 0; r < rows; r++) {
        const tr = el('tr');
        tr.appendChild(el('td', { class: 'gutter num', text: String(r + 1) }));
        for (let c = 0; c < cols; c++) {
          const raw = payload.sheetsData[activeSheet][r] ? payload.sheetsData[activeSheet][r][c] : null;
          const isFormula = engine.hasFormula(activeSheet, r, c);

          if (raw === null || raw === undefined) {
            tr.appendChild(el('td', { class: 'empty-cell' }));
            continue;
          }

          const fmt = engine.formatOf(activeSheet, r, c);
          const isDateCell = !isFormula && typeof raw === 'number' && fmt && XLSX.SSF.is_date(fmt);

          if (isFormula) {
            const value = engine.getVal(activeSheet, r, c);
            const flagged = isErr(value) && unsupportedByAddr.get(`${activeSheet}|${r},${c}`);
            if (flagged) {
              const cachedValues = payload.cachedValues && payload.cachedValues[activeSheet];
              const cached = cachedValues ? cachedValues[`${r},${c}`] : undefined;
              const display = engine.fmtCell(activeSheet, r, c, cached);
              tr.appendChild(el('td', {
                class: 'num readonly frozen',
                text: '✱ ' + display,
                title: "Static value from Excel, not live — " + flagged.reason,
              }));
            } else {
              const display = engine.fmtCell(activeSheet, r, c, value);
              tr.appendChild(el('td', { class: 'num readonly' + (isErr(value) ? ' error' : ''), text: display }));
            }
          } else if (isDateCell) {
            // A raw Excel date serial (e.g. 46204) is meaningless as an editable
            // number input, so date-formatted constants render as read-only text
            // instead — same as a formula cell, just without recalculation.
            tr.appendChild(el('td', { class: 'num readonly', text: engine.fmtCell(activeSheet, r, c, raw) }));
          } else if (typeof raw === 'number') {
            // The engine's current value, not the workbook's original `raw` —
            // after an edit, renderGrid() rebuilds this input from scratch, so
            // seeding it from `raw` would always snap back to the un-edited
            // number instead of showing what the user just typed.
            const input = el('input', {
              type: 'number', class: 'num', value: engine.getVal(activeSheet, r, c),
              onchange: (e) => {
                const n = parseFloat(e.target.value);
                if (!isNaN(n)) { engine.setVal(activeSheet, r, c, n); renderPanel(); }
              },
            });
            tr.appendChild(el('td', { class: 'num input-cell' }, [input]));
          } else {
            tr.appendChild(el('td', { class: 'label-cell', text: String(raw) }));
          }
        }
        tbody.appendChild(tr);
      }
      table.appendChild(tbody);
      scroll.appendChild(table);
      card.appendChild(scroll);
      panelWrap.appendChild(card);
    }

    renderNav();
    renderPanel();

    content.appendChild(el('div', {
      class: 'footer',
      text: opts.footerText || `Generated from ${payload.sourceFile}. Formulas recalculate client-side via HyperFormula — no data leaves the browser.`,
    }));
  }

  return { createEngine, buildWorkbookBytes, renderApp, isErr, colLetter, el };
});
