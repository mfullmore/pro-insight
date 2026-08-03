// Round 2 formula-coverage stress test: SUMIFS/COUNTIFS/AVERAGEIF rollups,
// IFERROR/ISERROR, ROUND family, DATE/EOMONTH/quarter-bucketing, OFFSET,
// INDIRECT, Excel Table structured references, and two "should fail loudly,
// not silently" cases: an external-workbook reference and a circular ref.
const ExcelJS = require('exceljs');

function excelSerial(y, m, d) {
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

async function main() {
  const wb = new ExcelJS.Workbook();

  const data = wb.addWorksheet('Data');
  data.addRow(['Category', 'Dept', 'Quarter', 'Amount']);
  const rows = [
    ['Cloud Infra', 'Engineering', 'Q1', 8000],
    ['Cloud Infra', 'Engineering', 'Q2', 8000],
    ['Cloud Infra', 'Engineering', 'Q3', 8000],
    ['Cloud Infra', 'Engineering', 'Q4', 8000],
    ['Contractor', 'Engineering', 'Q1', 45000],
    ['Digital Ads', 'Marketing', 'Q1', 5000],
    ['Digital Ads', 'Marketing', 'Q2', 5000],
    ['CRM', 'Sales', 'Q1', 48000],
  ];
  rows.forEach((r) => data.addRow(r));

  // Excel Table (structured references), common in modern budget templates.
  data.addTable({
    name: 'BudgetTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { showRowStripes: true },
    columns: [{ name: 'Category' }, { name: 'Dept' }, { name: 'Quarter' }, { name: 'Amount' }],
    rows: rows,
  });

  const engineeringAmounts = rows.filter((r) => r[1] === 'Engineering').map((r) => r[3]);
  const engQ1 = rows.filter((r) => r[1] === 'Engineering' && r[2] === 'Q1').reduce((s, r) => s + r[3], 0);
  const marketingCount = rows.filter((r) => r[1] === 'Marketing').length;
  const salesSum = rows.filter((r) => r[1] === 'Sales').reduce((s, r) => s + r[3], 0);
  const engAvg = engineeringAmounts.reduce((s, x) => s + x, 0) / engineeringAmounts.length;
  const tableSum = rows.reduce((s, r) => s + r[3], 0);

  const t = wb.addWorksheet('FnTests');
  t.addRow(['Test', 'Result']);

  const cases = [];
  const add = (label, formula, result) => cases.push({ label, formula, result });

  add('SUMIFS Engineering+Q1', 'SUMIFS(Data!D:D,Data!B:B,"Engineering",Data!C:C,"Q1")', engQ1);
  add('COUNTIFS Marketing', 'COUNTIFS(Data!B:B,"Marketing")', marketingCount);
  add('SUMIF Sales', 'SUMIF(Data!B:B,"Sales",Data!D:D)', salesSum);
  add('AVERAGEIF Engineering', 'AVERAGEIF(Data!B:B,"Engineering",Data!D:D)', engAvg);
  add('IFERROR on failed VLOOKUP', 'IFERROR(VLOOKUP("Nope",Data!A:A,1,FALSE),"NOT FOUND")', 'NOT FOUND');
  add('IFERROR on div/0', 'IFERROR(1/0,"DIV ERROR")', 'DIV ERROR');
  add('ISERROR on div/0', 'ISERROR(1/0)', true);
  add('ROUND', 'ROUND(1234.5678,2)', 1234.57);
  add('ROUNDUP', 'ROUNDUP(1234.561,2)', 1234.57);
  add('ROUNDDOWN', 'ROUNDDOWN(1234.569,2)', 1234.56);
  add('DATE serial', 'DATE(2026,3,15)', excelSerial(2026, 3, 15));
  add('EOMONTH', 'EOMONTH(DATE(2026,3,15),0)', excelSerial(2026, 3, 31));
  add('YEAR', 'YEAR(DATE(2026,3,15))', 2026);
  add('MONTH', 'MONTH(DATE(2026,3,15))', 3);
  add('Quarter bucketing', 'ROUNDUP(MONTH(DATE(2026,3,15))/3,0)', 1);
  add('OFFSET into Data', 'OFFSET(Data!A1,1,3)', 8000);
  add('INDIRECT ref', 'INDIRECT("Data!D2")', 8000);
  add('Structured table ref SUM', 'SUM(BudgetTable[Amount])', tableSum);
  add('External workbook ref (should fail loudly)', "'[OtherBudget.xlsx]Sheet1'!A1", '__EXPECT_CONTROLLED_FAILURE__');

  cases.forEach(({ label, formula, result }) => {
    t.addRow([label, { formula, result: result === '__EXPECT_CONTROLLED_FAILURE__' ? '#REF!' : result }]);
  });

  // Circular reference test on its own sheet: A1 depends on A2, A2 depends on A1.
  const circ = wb.addWorksheet('CircularTest');
  circ.addRow([{ formula: 'A2+1', result: '#CYCLE!' }]);
  circ.addRow([{ formula: 'A1+1', result: '#CYCLE!' }]);

  console.log('Expected values (ground truth, hand/programmatically computed):');
  cases.forEach((c) => console.log(`  ${c.label.padEnd(45)} ${JSON.stringify(c.result)}`));

  const outPath = __dirname + '/budget-test-2.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log(`\nWrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
