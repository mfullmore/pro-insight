// Simulates the real product pipeline: parse an .xlsx with SheetJS
// (as a browser extension would), rebuild the workbook in HyperFormula
// (a JS recalculation engine), and check whether its recomputed values
// match what Excel itself cached when the file was saved.
const XLSX = require('xlsx');
const { HyperFormula } = require('hyperformula');
const path = require('path');

const filePath = path.join(__dirname, '..', 'fixtures', 'budget-test.xlsx');
const wb = XLSX.readFile(filePath, { cellFormula: true, cellNF: false });

const sheetNames = wb.SheetNames;
const sheetsData = {};
const expected = {}; // { "Sheet!A1": cachedValue }

for (const name of sheetNames) {
  const ws = wb.Sheets[name];
  const ref = ws['!ref'];
  const range = XLSX.utils.decode_range(ref);
  const grid = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell) {
        row.push(null);
        continue;
      }
      if (cell.f) {
        row.push('=' + cell.f);
        expected[`${name}!${addr}`] = cell.v;
      } else {
        row.push(cell.v);
      }
    }
    grid.push(row);
  }
  sheetsData[name] = grid;
}

// Named ranges: SheetJS exposes workbook-level defined names here.
const definedNames = (wb.Workbook && wb.Workbook.Names) || [];
console.log('Parsed named ranges from file:', definedNames.map((n) => `${n.Name} -> ${n.Ref}`));

const options = { licenseKey: 'gpl-v3' };
const hf = HyperFormula.buildEmpty(options);

// Workaround: HyperFormula's lexer has no token for bare TRUE/FALSE
// literals (only the TRUE()/FALSE() functions), so unmodified Excel
// formulas using bare TRUE/FALSE throw #NAME?. Registering them as
// global named expressions patches this.
hf.addNamedExpression('TRUE', true);
hf.addNamedExpression('FALSE', false);

const sheetIdByName = {};
for (const name of sheetNames) {
  const sheetId = hf.addSheet(name);
  sheetIdByName[name] = hf.getSheetId(name);
}

for (const name of sheetNames) {
  hf.setSheetContent(sheetIdByName[name], sheetsData[name]);
}

for (const dn of definedNames) {
  // dn.Ref looks like "Summary!$B$6" (may include leading '=' or workbook scoping)
  let ref = '=' + dn.Ref.replace(/^=/, '');
  try {
    hf.addNamedExpression(dn.Name, ref);
  } catch (err) {
    console.log(`  ! failed to register named range ${dn.Name} (${ref}): ${err.message}`);
  }
}

console.log(`\nComparing ${Object.keys(expected).length} formula cells (Excel cached value vs HyperFormula recompute):\n`);

let pass = 0;
let fail = 0;
for (const [addr, expectedValue] of Object.entries(expected)) {
  const [sheet, cellRef] = addr.split('!');
  const { r: row, c: col } = XLSX.utils.decode_cell(cellRef);
  const cellAddress = { sheet: sheetIdByName[sheet], row, col };
  let actual;
  try {
    actual = hf.getCellValue(cellAddress);
  } catch (err) {
    actual = `ERROR: ${err.message}`;
  }

  const numericMatch = typeof expectedValue === 'number' && typeof actual === 'number' && Math.abs(expectedValue - actual) < 0.01;
  const exactMatch = expectedValue === actual;
  const ok = numericMatch || exactMatch;
  ok ? pass++ : fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${addr.padEnd(20)} expected=${JSON.stringify(expectedValue).padEnd(20)} actual=${JSON.stringify(actual)}`);
}

console.log(`\n${pass} passed, ${fail} failed out of ${pass + fail}`);
