// Generates a realistic large-enterprise FY26 annual operating budget workbook
// for testing the Interactive Plan tools: an Assumptions sheet, a Headcount
// Plan, six department budgets, and a Summary rollup, all linked with real
// cross-sheet formulas (no INDIRECT / Table references — this file is meant
// to convert cleanly end-to-end; budget-test-3.xlsx already covers the
// unsupported-formula edge cases).
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'browser-converter', 'node_modules', 'xlsx'));

const OUT_FILE = path.join(__dirname, '..', 'enterprise-budget-fy26.xlsx');

const CUR = '$#,##0';
const PCT = '0.0%';
const INT = '#,##0';

function A(r, c) { return XLSX.utils.encode_cell({ r, c }); }
function setNum(ws, r, c, v, z) { ws[A(r, c)] = z ? { t: 'n', v, z } : { t: 'n', v }; }
function setFormula(ws, r, c, f, v, z) { ws[A(r, c)] = z ? { t: 'n', v, f, z } : { t: 'n', v, f }; }
function setText(ws, r, c, v) { ws[A(r, c)] = { t: 's', v }; }
function setRange(ws, rTop, rBot, cLeft, cRight) {
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: rTop, c: cLeft }, e: { r: rBot, c: cRight } });
}

// ---------------------------------------------------------------------------
// Assumptions
// ---------------------------------------------------------------------------
const ASM = {
  fiscalYear: 3, benefitsLoad: 4, meritIncrease: 5, salaryNonTech: 6, salaryTech: 7,
  tePerFte: 8, facilitiesPerFte: 9, softwarePerFte: 10, contractorRate: 11, telecomPerFte: 12,
};
const ASM_VALUES = {
  fiscalYear: 2026,
  benefitsLoad: 0.32,
  meritIncrease: 0.035,
  salaryNonTech: 118000,
  salaryTech: 165000,
  tePerFte: 850,
  facilitiesPerFte: 2100,
  softwarePerFte: 1200,
  contractorRate: 145,
  telecomPerFte: 180,
};

function buildAssumptionsSheet() {
  const ws = {};
  setText(ws, 0, 0, 'Northbridge Financial Group — FY26 Budget Assumptions');
  setText(ws, 2, 0, 'Assumption'); setText(ws, 2, 1, 'Value'); setText(ws, 2, 2, 'Notes');

  setText(ws, ASM.fiscalYear, 0, 'Fiscal Year');
  setNum(ws, ASM.fiscalYear, 1, ASM_VALUES.fiscalYear, INT);

  setText(ws, ASM.benefitsLoad, 0, 'Benefits Load % (of Salary)');
  setNum(ws, ASM.benefitsLoad, 1, ASM_VALUES.benefitsLoad, PCT);
  setText(ws, ASM.benefitsLoad, 2, 'Health, retirement match, payroll taxes');

  setText(ws, ASM.meritIncrease, 0, 'Standard Merit Increase %');
  setNum(ws, ASM.meritIncrease, 1, ASM_VALUES.meritIncrease, PCT);

  setText(ws, ASM.salaryNonTech, 0, 'Avg Fully-Loaded Salary — Non-Technology');
  setNum(ws, ASM.salaryNonTech, 1, ASM_VALUES.salaryNonTech, CUR);

  setText(ws, ASM.salaryTech, 0, 'Avg Fully-Loaded Salary — Technology');
  setNum(ws, ASM.salaryTech, 1, ASM_VALUES.salaryTech, CUR);

  setText(ws, ASM.tePerFte, 0, 'Travel & Entertainment per FTE (Quarterly)');
  setNum(ws, ASM.tePerFte, 1, ASM_VALUES.tePerFte, CUR);

  setText(ws, ASM.facilitiesPerFte, 0, 'Facilities Cost per FTE (Quarterly)');
  setNum(ws, ASM.facilitiesPerFte, 1, ASM_VALUES.facilitiesPerFte, CUR);

  setText(ws, ASM.softwarePerFte, 0, 'Software & Technology Allocation per FTE (Quarterly)');
  setNum(ws, ASM.softwarePerFte, 1, ASM_VALUES.softwarePerFte, CUR);

  setText(ws, ASM.contractorRate, 0, 'Contractor Blended Rate (Hourly)');
  setNum(ws, ASM.contractorRate, 1, ASM_VALUES.contractorRate, CUR);
  setText(ws, ASM.contractorRate, 2, 'Reference only — not wired to a formula');

  setText(ws, ASM.telecomPerFte, 0, 'Telecom & Communications per FTE (Quarterly)');
  setNum(ws, ASM.telecomPerFte, 1, ASM_VALUES.telecomPerFte, CUR);

  setRange(ws, 0, ASM.telecomPerFte, 0, 2);
  return ws;
}
const ASM_SHEET_NAME = 'Assumptions';

// ---------------------------------------------------------------------------
// Departments — the model driving everything downstream
// ---------------------------------------------------------------------------
const DEPTS = [
  {
    sheet: 'Technology', hcRow: 3, salaryRate: 'salaryTech', growth: 0.08,
    fte: [142, 145, 148, 152],
    programLabel: 'Cloud Infrastructure & Hosting',
    program: [410000, 425000, 440000, 460000],
    contractors: [180000, 195000, 210000, 200000],
    training: [22000, 18000, 25000, 30000],
    officeSupplies: [8000, 7500, 8200, 9000],
    other: [12000, 10000, 15000, 11000],
    capex: [650000, 120000, 300000, 180000],
  },
  {
    sheet: 'Retail Banking Ops', hcRow: 4, salaryRate: 'salaryNonTech', growth: 0.03,
    fte: [890, 895, 900, 905],
    programLabel: 'Branch Operations & Cash Services',
    program: [520000, 530000, 540000, 550000],
    contractors: [65000, 60000, 70000, 68000],
    training: [45000, 40000, 42000, 48000],
    officeSupplies: [35000, 33000, 34000, 36000],
    other: [20000, 18000, 22000, 19000],
    capex: [90000, 60000, 45000, 120000],
  },
  {
    sheet: 'Risk & Compliance', hcRow: 5, salaryRate: 'salaryNonTech', growth: 0.12,
    fte: [78, 82, 86, 90],
    programLabel: 'Regulatory Filings & Audit Fees',
    program: [310000, 300000, 340000, 360000],
    contractors: [140000, 150000, 160000, 155000],
    training: [15000, 14000, 16000, 18000],
    officeSupplies: [4000, 3800, 4200, 4500],
    other: [9000, 8500, 10000, 9500],
    capex: [25000, 15000, 10000, 30000],
  },
  {
    sheet: 'Marketing', hcRow: 6, salaryRate: 'salaryNonTech', growth: 0.06,
    fte: [54, 56, 58, 60],
    programLabel: 'Marketing Programs (Campaigns, Media & Events)',
    program: [680000, 720000, 900000, 1100000],
    contractors: [95000, 100000, 110000, 130000],
    training: [12000, 10000, 11000, 13000],
    officeSupplies: [3000, 2800, 3200, 3500],
    other: [15000, 12000, 14000, 20000],
    capex: [20000, 15000, 10000, 25000],
  },
  {
    sheet: 'Human Resources', hcRow: 7, salaryRate: 'salaryNonTech', growth: 0.04,
    fte: [46, 47, 48, 49],
    programLabel: 'Recruiting & Talent Acquisition',
    program: [220000, 200000, 240000, 260000],
    contractors: [40000, 35000, 45000, 50000],
    training: [60000, 55000, 58000, 65000],
    officeSupplies: [3500, 3200, 3400, 3700],
    other: [11000, 9500, 10500, 12000],
    capex: [8000, 5000, 6000, 9000],
  },
  {
    sheet: 'Corporate Services', hcRow: 8, salaryRate: 'salaryNonTech', growth: 0.05,
    fte: [63, 64, 65, 66],
    programLabel: 'Corporate Insurance, Legal & Facilities',
    program: [380000, 390000, 400000, 410000],
    contractors: [70000, 65000, 72000, 75000],
    training: [10000, 9000, 9500, 11000],
    officeSupplies: [12000, 11500, 11800, 12500],
    other: [18000, 16000, 17000, 19000],
    capex: [35000, 20000, 25000, 40000],
  },
];

// ---------------------------------------------------------------------------
// Headcount Plan — FTE counts (editable) + computed quarterly salary $
// ---------------------------------------------------------------------------
const HC_SHEET_NAME = 'Headcount Plan';
function buildHeadcountSheet() {
  const ws = {};
  setText(ws, 0, 0, 'Northbridge Financial Group — FY26 Headcount Plan');
  setText(ws, 2, 0, 'Department');
  ['Q1 FTE', 'Q2 FTE', 'Q3 FTE', 'Q4 FTE', 'Q1 Salary $', 'Q2 Salary $', 'Q3 Salary $', 'Q4 Salary $'].forEach((h, i) => setText(ws, 2, 1 + i, h));

  const salaryByDeptQtr = {}; // sheet -> [q1,q2,q3,q4] computed salary $
  DEPTS.forEach((d) => {
    const r = d.hcRow;
    setText(ws, r, 0, d.sheet);
    const rate = ASM_VALUES[d.salaryRate];
    salaryByDeptQtr[d.sheet] = [];
    for (let q = 0; q < 4; q++) {
      setNum(ws, r, 1 + q, d.fte[q], INT);
      const salary = Math.round((d.fte[q] * rate) / 4);
      salaryByDeptQtr[d.sheet].push(salary);
      const rateAddr = `${ASM_SHEET_NAME}!$B$${d.salaryRate === 'salaryTech' ? ASM.salaryTech + 1 : ASM.salaryNonTech + 1}`;
      setFormula(ws, r, 5 + q, `${A(r, 1 + q)}*${rateAddr}/4`, salary, CUR);
    }
  });

  setRange(ws, 0, DEPTS[DEPTS.length - 1].hcRow, 0, 8);
  return { ws, salaryByDeptQtr };
}

// ---------------------------------------------------------------------------
// Department sheet — 11 line items + OpEx subtotal + CapEx + Total
// ---------------------------------------------------------------------------
const ROW = {
  salaries: 3, benefits: 4, contractors: 5, software: 6, travel: 7, training: 8,
  facilities: 9, telecom: 10, officeSupplies: 11, program: 12, other: 13,
  opexTotal: 14, capex: 15, deptTotal: 16,
};
const COL = { item: 0, q1: 1, q2: 2, q3: 3, q4: 4, fy26: 5, fy25: 6, yoy: 7, notes: 8 };

function buildDeptSheet(d, salaryByDeptQtr) {
  const ws = {};
  setText(ws, 0, 0, `Northbridge Financial Group — FY26 Budget — ${d.sheet}`);
  ['Line Item', 'Q1 Budget', 'Q2 Budget', 'Q3 Budget', 'Q4 Budget', 'FY26 Total', 'FY25 Actual', 'YoY Change %', 'Notes']
    .forEach((h, i) => setText(ws, 2, i, h));

  const hcRow1 = d.hcRow + 1; // 1-indexed Excel row in Headcount Plan
  const fy26 = {}; // row -> [q1..q4] computed values, for FY26 Total + FY25 back-calc

  function rowLabel(r, label) { setText(ws, r, COL.item, label); }
  function fillQuarterly(r, values, isFormulaFn) {
    fy26[r] = values;
    for (let q = 0; q < 4; q++) {
      const c = COL.q1 + q;
      if (isFormulaFn) {
        const { f } = isFormulaFn(q);
        setFormula(ws, r, c, f, values[q], CUR);
      } else {
        setNum(ws, r, c, values[q], CUR);
      }
    }
  }
  function fillTotalsAndYoy(r, fy25Value) {
    setFormula(ws, r, COL.fy26, `SUM(${A(r, COL.q1)}:${A(r, COL.q4)})`, fy26[r].reduce((a, b) => a + b, 0), CUR);
    setNum(ws, r, COL.fy25, fy25Value, CUR);
    const fy26Total = fy26[r].reduce((a, b) => a + b, 0);
    setFormula(ws, r, COL.yoy, `(${A(r, COL.fy26)}-${A(r, COL.fy25)})/${A(r, COL.fy25)}`, (fy26Total - fy25Value) / fy25Value, PCT);
  }

  // Salaries & Wages — from Headcount Plan
  rowLabel(ROW.salaries, 'Salaries & Wages');
  fillQuarterly(ROW.salaries, salaryByDeptQtr[d.sheet], (q) => ({
    f: `'${HC_SHEET_NAME}'!${A(hcRow1 - 1, 5 + q)}`,
  }));
  fillTotalsAndYoy(ROW.salaries, Math.round(fy26[ROW.salaries].reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Payroll Taxes & Benefits — Salaries * Benefits Load %
  rowLabel(ROW.benefits, 'Payroll Taxes & Benefits');
  const benefitsVals = fy26[ROW.salaries].map((v) => Math.round(v * ASM_VALUES.benefitsLoad));
  fillQuarterly(ROW.benefits, benefitsVals, (q) => ({ f: `${A(ROW.salaries, COL.q1 + q)}*${ASM_SHEET_NAME}!$B$${ASM.benefitsLoad + 1}` }));
  fillTotalsAndYoy(ROW.benefits, Math.round(benefitsVals.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Contractors & Professional Services — editable
  rowLabel(ROW.contractors, 'Contractors & Professional Services');
  fillQuarterly(ROW.contractors, d.contractors);
  fillTotalsAndYoy(ROW.contractors, Math.round(d.contractors.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Software & Technology Subscriptions — FTE * per-FTE rate
  rowLabel(ROW.software, 'Software & Technology Subscriptions');
  const softwareVals = d.fte.map((f) => Math.round(f * ASM_VALUES.softwarePerFte));
  fillQuarterly(ROW.software, softwareVals, (q) => ({ f: `'${HC_SHEET_NAME}'!${A(hcRow1 - 1, 1 + q)}*${ASM_SHEET_NAME}!$B$${ASM.softwarePerFte + 1}` }));
  fillTotalsAndYoy(ROW.software, Math.round(softwareVals.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Travel & Entertainment — FTE * per-FTE rate
  rowLabel(ROW.travel, 'Travel & Entertainment');
  const travelVals = d.fte.map((f) => Math.round(f * ASM_VALUES.tePerFte));
  fillQuarterly(ROW.travel, travelVals, (q) => ({ f: `'${HC_SHEET_NAME}'!${A(hcRow1 - 1, 1 + q)}*${ASM_SHEET_NAME}!$B$${ASM.tePerFte + 1}` }));
  fillTotalsAndYoy(ROW.travel, Math.round(travelVals.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Training & Development — editable
  rowLabel(ROW.training, 'Training & Development');
  fillQuarterly(ROW.training, d.training);
  fillTotalsAndYoy(ROW.training, Math.round(d.training.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Facilities Allocation — FTE * per-FTE rate
  rowLabel(ROW.facilities, 'Facilities Allocation');
  const facilitiesVals = d.fte.map((f) => Math.round(f * ASM_VALUES.facilitiesPerFte));
  fillQuarterly(ROW.facilities, facilitiesVals, (q) => ({ f: `'${HC_SHEET_NAME}'!${A(hcRow1 - 1, 1 + q)}*${ASM_SHEET_NAME}!$B$${ASM.facilitiesPerFte + 1}` }));
  fillTotalsAndYoy(ROW.facilities, Math.round(facilitiesVals.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Telecom & Communications — FTE * per-FTE rate
  rowLabel(ROW.telecom, 'Telecom & Communications');
  const telecomVals = d.fte.map((f) => Math.round(f * ASM_VALUES.telecomPerFte));
  fillQuarterly(ROW.telecom, telecomVals, (q) => ({ f: `'${HC_SHEET_NAME}'!${A(hcRow1 - 1, 1 + q)}*${ASM_SHEET_NAME}!$B$${ASM.telecomPerFte + 1}` }));
  fillTotalsAndYoy(ROW.telecom, Math.round(telecomVals.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Office Supplies & Equipment — editable
  rowLabel(ROW.officeSupplies, 'Office Supplies & Equipment');
  fillQuarterly(ROW.officeSupplies, d.officeSupplies);
  fillTotalsAndYoy(ROW.officeSupplies, Math.round(d.officeSupplies.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Department-specific program line — editable
  rowLabel(ROW.program, d.programLabel);
  fillQuarterly(ROW.program, d.program);
  fillTotalsAndYoy(ROW.program, Math.round(d.program.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Other / Miscellaneous — editable
  rowLabel(ROW.other, 'Other / Miscellaneous');
  fillQuarterly(ROW.other, d.other);
  fillTotalsAndYoy(ROW.other, Math.round(d.other.reduce((a, b) => a + b, 0) / (1 + d.growth)));

  // Total OpEx — sum of the 11 line items above, per quarter
  rowLabel(ROW.opexTotal, 'Total OpEx');
  const opexRows = [ROW.salaries, ROW.benefits, ROW.contractors, ROW.software, ROW.travel, ROW.training, ROW.facilities, ROW.telecom, ROW.officeSupplies, ROW.program, ROW.other];
  const opexVals = [0, 1, 2, 3].map((q) => opexRows.reduce((sum, r) => sum + fy26[r][q], 0));
  fy26[ROW.opexTotal] = opexVals;
  for (let q = 0; q < 4; q++) {
    setFormula(ws, ROW.opexTotal, COL.q1 + q, `SUM(${A(ROW.salaries, COL.q1 + q)}:${A(ROW.other, COL.q1 + q)})`, opexVals[q], CUR);
  }
  const opexFy25 = opexRows.reduce((sum, r) => sum + parseInt(ws[A(r, COL.fy25)].v, 10), 0);
  setFormula(ws, ROW.opexTotal, COL.fy26, `SUM(${A(ROW.opexTotal, COL.q1)}:${A(ROW.opexTotal, COL.q4)})`, opexVals.reduce((a, b) => a + b, 0), CUR);
  setFormula(ws, ROW.opexTotal, COL.fy25, `SUM(${A(ROW.salaries, COL.fy25)}:${A(ROW.other, COL.fy25)})`, opexFy25, CUR);
  const opexFy26Total = opexVals.reduce((a, b) => a + b, 0);
  setFormula(ws, ROW.opexTotal, COL.yoy, `(${A(ROW.opexTotal, COL.fy26)}-${A(ROW.opexTotal, COL.fy25)})/${A(ROW.opexTotal, COL.fy25)}`, (opexFy26Total - opexFy25) / opexFy25, PCT);

  // Capital Expenditures — editable
  rowLabel(ROW.capex, 'Capital Expenditures');
  fillQuarterly(ROW.capex, d.capex);
  const capexFy25 = Math.round(d.capex.reduce((a, b) => a + b, 0) / (1 + d.growth));
  fillTotalsAndYoy(ROW.capex, capexFy25);

  // Total Department Budget — OpEx + CapEx
  rowLabel(ROW.deptTotal, 'Total Department Budget');
  for (let q = 0; q < 4; q++) {
    const c = COL.q1 + q;
    const val = opexVals[q] + d.capex[q];
    setFormula(ws, ROW.deptTotal, c, `${A(ROW.opexTotal, c)}+${A(ROW.capex, c)}`, val, CUR);
  }
  const deptTotalFy26 = opexFy26Total + d.capex.reduce((a, b) => a + b, 0);
  setFormula(ws, ROW.deptTotal, COL.fy26, `SUM(${A(ROW.deptTotal, COL.q1)}:${A(ROW.deptTotal, COL.q4)})`, deptTotalFy26, CUR);
  const deptTotalFy25 = opexFy25 + capexFy25;
  setFormula(ws, ROW.deptTotal, COL.fy25, `${A(ROW.opexTotal, COL.fy25)}+${A(ROW.capex, COL.fy25)}`, deptTotalFy25, CUR);
  setFormula(ws, ROW.deptTotal, COL.yoy, `(${A(ROW.deptTotal, COL.fy26)}-${A(ROW.deptTotal, COL.fy25)})/${A(ROW.deptTotal, COL.fy25)}`, (deptTotalFy26 - deptTotalFy25) / deptTotalFy25, PCT);

  setRange(ws, 0, ROW.deptTotal, 0, 8);
  return { ws, deptTotalFy26, deptTotalFy25 };
}

// A few realistic free-text notes, sprinkled in (not every row needs one).
const NOTES = {
  Technology: { [ROW.salaries]: 'Includes planned Q3 platform-engineering team expansion', [ROW.capex]: 'Q1 data-center hardware refresh cycle' },
  'Retail Banking Ops': { [ROW.program]: 'Branch cash-handling vendor contract renews in Q4' },
  'Risk & Compliance': { [ROW.contractors]: 'External audit firm engagement, annual exam cycle', [ROW.salaries]: 'New regulatory-reporting hires per FY26 exam findings' },
  Marketing: { [ROW.program]: 'Q4 spend reflects year-end deposit-growth campaign' },
  'Human Resources': { [ROW.training]: 'Manager development program launches company-wide in Q2' },
  'Corporate Services': { [ROW.other]: 'General liability insurance premium true-up' },
};

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------
function buildSummarySheet(deptTotals) {
  const ws = {};
  setText(ws, 0, 0, 'Northbridge Financial Group — FY26 Budget Summary');
  ['Department', 'FY26 Budget', 'FY25 Actual', 'YoY Change %', '% of Total Budget'].forEach((h, i) => setText(ws, 2, i, h));

  const firstRow = 3;
  const grandTotalRow = firstRow + DEPTS.length;
  DEPTS.forEach((d, i) => {
    const r = firstRow + i;
    setText(ws, r, 0, d.sheet);
    const t = deptTotals[d.sheet];
    setFormula(ws, r, 1, `'${d.sheet}'!${A(ROW.deptTotal, COL.fy26)}`, t.deptTotalFy26, CUR);
    setFormula(ws, r, 2, `'${d.sheet}'!${A(ROW.deptTotal, COL.fy25)}`, t.deptTotalFy25, CUR);
    setFormula(ws, r, 3, `(${A(r, 1)}-${A(r, 2)})/${A(r, 2)}`, (t.deptTotalFy26 - t.deptTotalFy25) / t.deptTotalFy25, PCT);
    setFormula(ws, r, 4, `${A(r, 1)}/${A(grandTotalRow, 1)}`, 0, PCT); // filled after grand total known
  });

  const grandFy26 = DEPTS.reduce((sum, d) => sum + deptTotals[d.sheet].deptTotalFy26, 0);
  const grandFy25 = DEPTS.reduce((sum, d) => sum + deptTotals[d.sheet].deptTotalFy25, 0);
  setText(ws, grandTotalRow, 0, 'Grand Total');
  setFormula(ws, grandTotalRow, 1, `SUM(${A(firstRow, 1)}:${A(grandTotalRow - 1, 1)})`, grandFy26, CUR);
  setFormula(ws, grandTotalRow, 2, `SUM(${A(firstRow, 2)}:${A(grandTotalRow - 1, 2)})`, grandFy25, CUR);
  setFormula(ws, grandTotalRow, 3, `(${A(grandTotalRow, 1)}-${A(grandTotalRow, 2)})/${A(grandTotalRow, 2)}`, (grandFy26 - grandFy25) / grandFy25, PCT);

  // Now backfill each department's % of total with its real cached value.
  DEPTS.forEach((d, i) => {
    const r = firstRow + i;
    const t = deptTotals[d.sheet];
    ws[A(r, 4)].v = t.deptTotalFy26 / grandFy26;
  });

  // OpEx vs CapEx breakdown
  const sectionRow = grandTotalRow + 2;
  setText(ws, sectionRow, 0, 'OpEx vs CapEx Breakdown');
  setText(ws, sectionRow + 1, 0, 'Category'); setText(ws, sectionRow + 1, 1, 'FY26 Budget');

  const opexRow = sectionRow + 2, capexRow = sectionRow + 3, totalRow = sectionRow + 4;
  setText(ws, opexRow, 0, 'Total OpEx');
  const opexFormula = DEPTS.map((d) => `'${d.sheet}'!${A(ROW.opexTotal, COL.fy26)}`).join('+');
  const grandOpex = grandFy26 - DEPTS.reduce((sum, d) => sum + d.capex.reduce((a, b) => a + b, 0), 0);
  setFormula(ws, opexRow, 1, opexFormula, grandOpex, CUR);

  setText(ws, capexRow, 0, 'Total CapEx');
  const capexFormula = DEPTS.map((d) => `'${d.sheet}'!${A(ROW.capex, COL.fy26)}`).join('+');
  const grandCapex = DEPTS.reduce((sum, d) => sum + d.capex.reduce((a, b) => a + b, 0), 0);
  setFormula(ws, capexRow, 1, capexFormula, grandCapex, CUR);

  setText(ws, totalRow, 0, 'Total FY26 Budget');
  setFormula(ws, totalRow, 1, `${A(opexRow, 1)}+${A(capexRow, 1)}`, grandOpex + grandCapex, CUR);

  setRange(ws, 0, totalRow, 0, 4);
  return ws;
}

// ---------------------------------------------------------------------------
// Assemble workbook
// ---------------------------------------------------------------------------
const wb = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb, buildAssumptionsSheet(), ASM_SHEET_NAME);

const { ws: hcWs, salaryByDeptQtr } = buildHeadcountSheet();
XLSX.utils.book_append_sheet(wb, hcWs, HC_SHEET_NAME);

const deptTotals = {};
DEPTS.forEach((d) => {
  const { ws, deptTotalFy26, deptTotalFy25 } = buildDeptSheet(d, salaryByDeptQtr);
  const notes = NOTES[d.sheet] || {};
  Object.keys(notes).forEach((r) => setText(ws, Number(r), COL.notes, notes[r]));
  deptTotals[d.sheet] = { deptTotalFy26, deptTotalFy25 };
  XLSX.utils.book_append_sheet(wb, ws, d.sheet);
});

XLSX.utils.book_append_sheet(wb, buildSummarySheet(deptTotals), 'Summary');

XLSX.writeFile(wb, OUT_FILE);
console.log(`Wrote ${OUT_FILE}`);
console.log('Sheets:', wb.SheetNames.join(', '));
console.log('Grand Total FY26 Budget:', DEPTS.reduce((sum, d) => sum + deptTotals[d.sheet].deptTotalFy26, 0).toLocaleString());
