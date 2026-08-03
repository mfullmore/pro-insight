// A messier, more realistic budget workbook for hardening the interactive
// prototype: departments with uneven line-item counts (Marketing has 5,
// not 4), and one line that uses INDIRECT instead of VLOOKUP — a plausible
// thing a power-user analyst actually writes, and one of the two known
// HyperFormula gaps found during fidelity testing. Also formats the Rates
// sheet as a real Excel Table (structured references), the other gap.
const ExcelJS = require('exceljs');

const RATES = { ENG: 0.15, MKT: 0.12, SLS: 0.18 };

const DEPARTMENTS = {
  Engineering: {
    code: 'ENG',
    items: [
      ['Cloud Infrastructure', 'AWS', 12, 8000],
      ['Contractor Staffing', 'Acme Contracting', 3, 45000],
      ['Software Licenses', 'JetBrains', 50, 200],
      ['Hardware', 'Dell', 20, 1800],
    ],
  },
  Marketing: {
    code: 'MKT',
    items: [
      ['Digital Ads', 'Google Ads', 12, 5000],
      ['Content Agency', 'Contently', 12, 3000],
      ['Events & Conferences', 'Various', 4, 15000],
      ['Marketing Software', 'HubSpot', 1, 24000],
      ['Regional Ad Buy', 'Various Local', 1, 30000], // uses INDIRECT below
    ],
  },
  Sales: {
    code: 'SLS',
    items: [
      ['CRM Software', 'Salesforce', 1, 48000],
      ['Sales Commission Pool', 'N/A', 1, 220000],
      ['Travel & Entertainment', 'Various', 1, 35000],
      ['Sales Enablement Tools', 'Gong', 1, 18000],
    ],
  },
};

const APPROVED_BUDGET = 850000;

function buildWorkbook() {
  const wb = new ExcelJS.Workbook();

  const ratesRows = Object.entries(RATES);
  const rates = wb.addWorksheet('Rates');
  rates.addRow(['Code', 'Overhead Rate']);
  ratesRows.forEach(([code, rate]) => rates.addRow([code, rate]));
  rates.addTable({
    name: 'RatesTable',
    ref: 'A1',
    headerRow: true,
    columns: [{ name: 'Code' }, { name: 'Overhead Rate' }],
    rows: ratesRows,
  });

  const subtotals = {};

  for (const [dept, { code, items }] of Object.entries(DEPARTMENTS)) {
    const ws = wb.addWorksheet(dept);
    ws.addRow(['Category', 'Vendor', 'Qty', 'Unit Cost', 'Cost Center Code', 'Overhead Rate', 'Line Total']);

    let subtotal = 0;
    items.forEach(([category, vendor, qty, unitCost], i) => {
      const row = i + 2;
      const rate = RATES[code];
      const isIndirectRow = dept === 'Marketing' && i === items.length - 1;
      const lineTotal = qty * unitCost * (1 + rate);
      // The INDIRECT row's "expected" value is intentionally left as a
      // real number here (what a correct engine *would* produce) so the
      // fidelity report can show "computed vs. what it should have been."
      subtotal += lineTotal;

      const overheadFormula = isIndirectRow
        ? `INDIRECT("Rates!B"&MATCH(E${row},Rates!A:A,0))`
        : `VLOOKUP(E${row},Rates!A:B,2,FALSE)`;

      ws.addRow([
        category,
        vendor,
        qty,
        unitCost,
        code,
        { formula: overheadFormula, result: rate },
        { formula: `C${row}*D${row}*(1+F${row})`, result: Math.round(lineTotal * 100) / 100 },
      ]);
    });

    const lastItemRow = items.length + 1;
    const subtotalRow = lastItemRow + 1;
    ws.addRow(['Subtotal', '', '', '', '', '', { formula: `SUM(G2:G${lastItemRow})`, result: Math.round(subtotal * 100) / 100 }]);
    subtotals[dept] = { cell: `G${subtotalRow}`, value: Math.round(subtotal * 100) / 100 };
  }

  const summary = wb.addWorksheet('Summary');
  summary.addRow(['Budget Summary']);
  summary.addRow(['Department', 'Subtotal']);

  let rowNum = 3;
  for (const dept of Object.keys(DEPARTMENTS)) {
    summary.addRow([dept, { formula: `${dept}!${subtotals[dept].cell}`, result: subtotals[dept].value }]);
    rowNum++;
  }
  const firstDeptRow = 3;
  const lastDeptRow = rowNum - 1;

  const totalBudget = Object.values(subtotals).reduce((s, x) => s + x.value, 0);
  const totalRow = rowNum++;
  summary.addRow(['Total Budget', { formula: `SUM(B${firstDeptRow}:B${lastDeptRow})`, result: Math.round(totalBudget * 100) / 100 }]);

  const approvedRow = rowNum++;
  summary.addRow(['Approved Budget', APPROVED_BUDGET]);

  const varianceRow = rowNum++;
  const variance = totalBudget - APPROVED_BUDGET;
  summary.addRow(['Variance', { formula: `TotalBudget-ApprovedBudget`, result: Math.round(variance * 100) / 100 }]);

  const variancePctRow = rowNum++;
  const variancePct = variance / APPROVED_BUDGET;
  summary.addRow(['Variance %', { formula: `B${varianceRow}/ApprovedBudget`, result: variancePct }]);

  const statusRow = rowNum++;
  const status = totalBudget > APPROVED_BUDGET ? 'OVER BUDGET' : 'WITHIN BUDGET';
  summary.addRow(['Status', { formula: `IF(TotalBudget>ApprovedBudget,"OVER BUDGET","WITHIN BUDGET")`, result: status }]);

  const highestRow = rowNum++;
  const highestDept = Object.entries(subtotals).sort((a, b) => b[1].value - a[1].value)[0][0];
  summary.addRow(['Highest Spending Dept', { formula: `INDEX(A${firstDeptRow}:A${lastDeptRow},MATCH(MAX(B${firstDeptRow}:B${lastDeptRow}),B${firstDeptRow}:B${lastDeptRow},0))`, result: highestDept }]);

  const avgRateRow = rowNum++;
  summary.addRow(['Avg Overhead Rate (structured ref)', { formula: `AVERAGE(RatesTable[Overhead Rate])`, result: Object.values(RATES).reduce((a, b) => a + b, 0) / Object.values(RATES).length }]);

  wb.definedNames.add(`Summary!$B$${totalRow}`, 'TotalBudget');
  wb.definedNames.add(`Summary!$B$${approvedRow}`, 'ApprovedBudget');

  console.log('Note: Marketing has 5 line items (uneven vs. Engineering/Sales at 4); its 5th row uses INDIRECT (unsupported).');
  console.log('Note: Summary!B' + avgRateRow + ' uses a structured table reference (unsupported).');
  console.log('Expected-if-everything-converted totals:', { subtotals: Object.fromEntries(Object.entries(subtotals).map(([k, v]) => [k, v.value])), totalBudget, variance, variancePct, status, highestDept });

  return wb;
}

async function main() {
  const wb = buildWorkbook();
  const outPath = __dirname + '/budget-test-3.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log(`\nWrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
