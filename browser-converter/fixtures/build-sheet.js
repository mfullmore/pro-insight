// Builds a representative bank department-budget workbook to test
// Excel -> JS conversion fidelity (cross-sheet refs, VLOOKUP, IF,
// INDEX/MATCH, named ranges). All formula results below are computed
// by hand so they act as ground truth "cached" values, exactly like
// what Excel itself would store after a real recalculation.
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

  const rates = wb.addWorksheet('Rates');
  rates.addRow(['Code', 'Overhead Rate']);
  Object.entries(RATES).forEach(([code, rate]) => rates.addRow([code, rate]));

  const subtotals = {};

  for (const [dept, { code, items }] of Object.entries(DEPARTMENTS)) {
    const ws = wb.addWorksheet(dept);
    ws.addRow(['Category', 'Vendor', 'Qty', 'Unit Cost', 'Cost Center Code', 'Overhead Rate', 'Line Total']);

    let subtotal = 0;
    items.forEach(([category, vendor, qty, unitCost], i) => {
      const row = i + 2;
      const rate = RATES[code];
      const lineTotal = qty * unitCost * (1 + rate);
      subtotal += lineTotal;
      ws.addRow([
        category,
        vendor,
        qty,
        unitCost,
        code,
        { formula: `VLOOKUP(E${row},Rates!A:B,2,FALSE)`, result: rate },
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
  const deptRows = {};
  for (const dept of Object.keys(DEPARTMENTS)) {
    summary.addRow([dept, { formula: `${dept}!${subtotals[dept].cell}`, result: subtotals[dept].value }]);
    deptRows[dept] = rowNum;
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

  // Named ranges, since real budget workbooks lean on these heavily.
  wb.definedNames.add(`Summary!$B$${totalRow}`, 'TotalBudget');
  wb.definedNames.add(`Summary!$B$${approvedRow}`, 'ApprovedBudget');

  console.log('Expected values (ground truth, hand-computed):');
  console.log({ subtotals: Object.fromEntries(Object.entries(subtotals).map(([k, v]) => [k, v.value])), totalBudget, variance, variancePct, status, highestDept });

  return wb;
}

async function main() {
  const wb = buildWorkbook();
  const outPath = __dirname + '/budget-test.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log(`\nWrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
