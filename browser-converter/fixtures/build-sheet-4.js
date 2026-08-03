// Deliberately adversarial workbook for pressure-testing the generic
// converter: merged cells, multiple disjoint tables on one sheet, a sheet
// with zero formulas, a genuinely empty sheet, a text-only notes sheet,
// and real Excel number formats (currency/percent/date) so we can finally
// see XLSX.SSF do something visible.
const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();

  // --- Merged title cell + a normal table below with real number formats ---
  const merged = wb.addWorksheet('Merged');
  merged.mergeCells('A1:C1');
  merged.getCell('A1').value = 'Q3 Regional Spend';
  merged.getCell('A1').font = { bold: true };

  merged.addRow(['Region', 'Amount', 'Date']);
  const regionRows = [
    ['Northeast', 42500.5, new Date(Date.UTC(2026, 6, 1))],
    ['Southeast', 31200, new Date(Date.UTC(2026, 6, 8))],
    ['West', -1500, new Date(Date.UTC(2026, 6, 15))], // a correction/refund, negative on purpose
  ];
  regionRows.forEach((r) => merged.addRow(r));
  for (let i = 0; i < regionRows.length; i++) {
    const row = 3 + i; // row 2 is the header, data starts row 3
    merged.getCell(`B${row}`).numFmt = '$#,##0.00';
    merged.getCell(`C${row}`).numFmt = 'm/d/yyyy';
  }
  const totalRow = 3 + regionRows.length;
  const totalSum = regionRows.reduce((s, r) => s + r[1], 0);
  merged.getCell(`A${totalRow}`).value = 'Total';
  merged.getCell(`B${totalRow}`).value = { formula: `SUM(B3:B${totalRow - 1})`, result: totalSum };
  merged.getCell(`B${totalRow}`).numFmt = '$#,##0.00';

  // --- Two side-by-side tables plus a third stacked below with a gap ---
  const multi = wb.addWorksheet('MultiTable');
  multi.addRow(['Vendor', 'Cost', null, 'Category', 'Rate']);
  multi.addRow(['Acme', 1200, null, 'Travel', 0.08]);
  multi.addRow(['Globex', 3400, null, 'Meals', 0.12]);
  multi.addRow(['Initech', 900, null, 'Lodging', 0.05]);
  multi.getCell('B5').value = { formula: 'SUM(B2:B4)', result: 5500 };
  multi.getCell('E2').numFmt = '0%';
  multi.getCell('E3').numFmt = '0%';
  multi.getCell('E4').numFmt = '0%';
  // blank row 6 as a gap, then a third table stacked below
  multi.addRow([]);
  multi.addRow(['Quarter', 'Headcount']);
  multi.addRow(['Q1', 42]);
  multi.addRow(['Q2', 45]);
  multi.getCell('B10').value = { formula: 'B8+B9', result: 87 };

  // --- A sheet with plain data and zero formulas ---
  const numbers = wb.addWorksheet('Numbers');
  numbers.addRow(['Item', 'Value']);
  numbers.addRow(['Widget count', 128]);
  numbers.addRow(['Refund', -45.5]);
  numbers.addRow(['Adjustment factor', 1.0375]);
  numbers.getCell('B4').numFmt = '0.00%';

  // --- A text-only notes sheet: no numbers, no formulas ---
  const notes = wb.addWorksheet('Notes');
  notes.addRow(['Planning Notes']);
  notes.addRow(['All figures pending final sign-off from regional directors before the Q3 close. This is a deliberately long note to check how the grid handles text that is much wider than a typical cell, since real spreadsheets often have exactly this kind of stray commentary column.']);
  notes.addRow(['Contact: budgeting@example.com']);

  // --- A genuinely empty sheet ---
  wb.addWorksheet('Empty');

  const outPath = __dirname + '/budget-test-4.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log(`Wrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
