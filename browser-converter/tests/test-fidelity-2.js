const XLSX = require('xlsx');
const { HyperFormula } = require('hyperformula');
const path = require('path');

const filePath = path.join(__dirname, '..', 'fixtures', 'budget-test-2.xlsx');
const wb = XLSX.readFile(filePath, { cellFormula: true, cellNF: false });

const sheetNames = wb.SheetNames;
const sheetsData = {};

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
      row.push(cell.f ? '=' + cell.f : cell.v);
    }
    grid.push(row);
  }
  sheetsData[name] = grid;
}

const options = { licenseKey: 'gpl-v3' };
const hf = HyperFormula.buildEmpty(options);
hf.addNamedExpression('TRUE', true);
hf.addNamedExpression('FALSE', false);

const sheetIdByName = {};
for (const name of sheetNames) {
  hf.addSheet(name);
  sheetIdByName[name] = hf.getSheetId(name);
}
for (const name of sheetNames) {
  hf.setSheetContent(sheetIdByName[name], sheetsData[name]);
}

console.log('=== FnTests sheet ===\n');
const ftGrid = sheetsData['FnTests'];
let pass = 0;
let fail = 0;
for (let r = 1; r < ftGrid.length; r++) {
  const label = ftGrid[r][0];
  const expectedRaw = wb.Sheets['FnTests'][XLSX.utils.encode_cell({ r, c: 1 })];
  const expected = expectedRaw ? expectedRaw.v : undefined;
  let actual;
  try {
    actual = hf.getCellValue({ sheet: sheetIdByName['FnTests'], row: r, col: 1 });
  } catch (err) {
    actual = `THROW: ${err.message}`;
  }

  let ok;
  if (expected === '#REF!') {
    // Controlled-failure case: we WANT an error object here, not a plausible number.
    ok = actual && typeof actual === 'object' && actual.value && String(actual.value).startsWith('#');
  } else {
    const numericMatch = typeof expected === 'number' && typeof actual === 'number' && Math.abs(expected - actual) < 0.01;
    ok = numericMatch || expected === actual;
  }
  ok ? pass++ : fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label.padEnd(45)} expected=${JSON.stringify(expected).padEnd(15)} actual=${JSON.stringify(actual)}`);
}

console.log('\n=== CircularTest sheet ===\n');
const circA1 = hf.getCellValue({ sheet: sheetIdByName['CircularTest'], row: 0, col: 0 });
const circA2 = hf.getCellValue({ sheet: sheetIdByName['CircularTest'], row: 1, col: 0 });
console.log('A1:', JSON.stringify(circA1));
console.log('A2:', JSON.stringify(circA2));
const circDetected = [circA1, circA2].some((v) => v && typeof v === 'object' && String(v.value).includes('CYCLE'));
console.log(circDetected ? 'PASS  circular reference correctly detected as an error, not a silent number' : 'FAIL  circular reference was NOT flagged as an error');
circDetected ? pass++ : fail++;

console.log(`\n${pass} passed, ${fail} failed out of ${pass + fail}`);
