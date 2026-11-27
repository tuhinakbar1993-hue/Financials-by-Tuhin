import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import type * as XlsxType from 'xlsx';

const require = createRequire(import.meta.url);
const xlsxCjs = require('xlsx') as typeof XlsxType;
const capturedWorkbooks: { filename: string; wb: XlsxType.WorkBook }[] = [];
const origWriteFile = xlsxCjs.writeFile;
xlsxCjs.writeFile = ((wb: XlsxType.WorkBook, filename: string) => {
  capturedWorkbooks.push({ filename, wb });
}) as typeof XlsxType.writeFile;

const engine = await import('./src/services/financialEngine');
const exportsMod = await import('./src/services/exportService');
import type { StatementConfigItem, StatementCellOverrides, StatementLabelOverrides } from './src/types/financial';

function runPipeline(
  config: StatementConfigItem[],
  labelOverrides: StatementLabelOverrides = {},
  cellOverrides: StatementCellOverrides = {}
) {
  const baseYear = engine.INITIAL_HISTORICAL_INPUTS[engine.INITIAL_HISTORICAL_INPUTS.length - 1];
  const projectedInputs = engine.generateProjectedYears(baseYear, engine.DEFAULT_PROJECTION_ASSUMPTIONS);
  const allInputs = [...engine.INITIAL_HISTORICAL_INPUTS, ...projectedInputs].sort((a, b) => a.year - b.year);
  const opts = { statementConfig: config, cellOverrides, labelOverrides };
  const fa = engine.calculateScheduleOfFixedAssets(allInputs, opts);
  const inc = engine.calculateIncomeStatement(allInputs, fa, opts);
  const cf = engine.calculateCashFlowStatement(allInputs, inc, fa, opts);
  const eq = engine.calculateEquityStatement(allInputs, inc, opts);
  const bs = engine.calculateBalanceSheet(allInputs, eq, fa, cf, opts, inc);
  const notes = engine.generateNotesToFinancialStatements(
    engine.INITIAL_COMPANY_PROFILE,
    allInputs.map(y => y.year),
    inc,
    bs,
    cf,
    fa,
    opts
  );
  const dscr = engine.calculateDebtServiceCoverage(allInputs, inc, bs, cf);
  const prev = engine.buildPreviousFormatData(allInputs, inc, bs, cf, eq, fa, opts);
  const proj = engine.buildProjectedFormatData(allInputs, inc, bs, cf, eq, fa, opts);
  return { allInputs, fa, inc, cf, eq, bs, notes, dscr, prev, proj, labelOverrides, cellOverrides };
}

function validateFormulaCells(wb: XlsxType.WorkBook) {
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const ref = ws['!ref'];
    if (!ref) continue;
    const range = xlsxCjs.utils.decode_range(ref);
    for (const addr of Object.keys(ws)) {
      if (addr.startsWith('!')) continue;
      const cell = ws[addr] as XlsxType.CellObject;
      if (cell && cell.f) {
        assert(!cell.f.includes('NaN') && !cell.f.includes('undefined') && !cell.f.includes('#REF!'), `Invalid formula ${cell.f} in ${sheetName}!${addr}`);
        const sumMatches = cell.f.matchAll(/SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/g);
        for (const m of sumMatches) {
          const r1 = Number(m[2]);
          const r2 = Number(m[4]);
          assert(r1 <= r2 && r1 >= 1 && r2 <= range.e.r + 1, `Out-of-bounds/inverted SUM ${m[0]} in ${sheetName}!${addr} (max row ${range.e.r + 1})`);
        }
        const cellRefs = cell.f.matchAll(/\b([A-Z]+)(\d+)\b/g);
        for (const m of cellRefs) {
          const r = Number(m[2]);
          assert(r >= 1 && r <= range.e.r + 1, `Row out of bounds ${m[0]} in ${sheetName}!${addr} (max row ${range.e.r + 1})`);
        }
      }
    }
  }
}

// 1. Baseline parity test
const baseline = runPipeline(engine.DEFAULT_STATEMENT_CONFIG_ITEMS);
baseline.bs.forEach(b => {
  assert(Math.abs(b.balanceCheck) < 0.05, `Baseline BS unbalanced in ${b.year}: ${b.balanceCheck}`);
});
baseline.cf.forEach(c => {
  assert.equal(c.reconciliationStatus, 'RECONCILED', `Baseline CF not reconciled in ${c.year}`);
});

// 2. Customization test: deselect items across all 5 statements + rename labels + add custom items + cellOverrides
const customConfig: StatementConfigItem[] = engine.DEFAULT_STATEMENT_CONFIG_ITEMS.map(item => {
  if (
    item.id === 'inc_costOfSales' ||
    item.id === 'inc_financeCost' ||
    item.id === 'bs_tradeReceivables' ||
    item.id === 'cf_dep' ||
    item.id === 'eq_dividendsPaid' ||
    item.id === 'fa_office_equipment'
  ) {
    return { ...item, isVisible: false };
  }
  return item;
});
customConfig.push(
  {
    id: 'custom_income_royalty',
    statement: 'income',
    label: 'Royalty Grant Income',
    section: 'Non-Operating & Tax',
    isMandatory: false,
    isVisible: true,
    order: 200,
    isCustom: true,
    sign: 'positive',
    customValues: Object.fromEntries(baseline.allInputs.map(y => [y.year, 50000])),
  },
  {
    id: 'custom_income_opex_extra',
    statement: 'income',
    label: 'Custom CSR Levy',
    section: 'Operating Expenses',
    isMandatory: false,
    isVisible: true,
    order: 201,
    isCustom: true,
    sign: 'negative',
    customValues: Object.fromEntries(baseline.allInputs.map(y => [y.year, 1200])),
  },
  {
    id: 'custom_income_rev_bonus',
    statement: 'income',
    label: 'Export Rebate Revenue',
    section: 'Revenue & Gross Margin',
    isMandatory: false,
    isVisible: true,
    order: 199,
    isCustom: true,
    sign: 'positive',
    customValues: Object.fromEntries(baseline.allInputs.map(y => [y.year, 2500])),
  }
);

const labelOverrides: StatementLabelOverrides = {
  income_revenue: 'Turnover & Gross Sales',
  bs_cashAndEquivalents: 'Cash at Bank & Hand',
};

// Include both an active cell override (income_revenue) and a stale cell override on a deselected element (bs_tradeReceivables)
const cellOverrides: StatementCellOverrides = {
  income_revenue: { 2025: 90000 },
  bs_tradeReceivables: { 2025: 999999 },
};

const customized = runPipeline(customConfig, labelOverrides, cellOverrides);

const base2025 = baseline.inc.find(d => d.year === 2025)!;
const cust2025 = customized.inc.find(d => d.year === 2025)!;
assert.equal(cust2025.costOfSales, 0, 'Deselected COGS should be 0');
assert.equal(cust2025.financeCost, 0, 'Deselected Finance Cost should be 0');
assert.equal(cust2025.revenue, 92500, 'Revenue should combine override (90000) + custom revenue (2500)');
assert(cust2025.grossProfit > base2025.grossProfit, 'Gross Profit should increase when COGS deselected');
assert(cust2025.netProfit > base2025.netProfit, 'Net Profit should adjust dynamically');

// Verify Notes match Income & Balance Sheet for latest year (2040) and historical-scoped Notes (2025) even with stale override on hidden BS Trade Receivables + custom Revenue
const cust2040 = customized.inc[customized.inc.length - 1];
const note2 = customized.notes.find(n => n.noteNumber === 2)!;
assert.equal(note2.tableData?.totalRow?.values[0], cust2040.revenue.toLocaleString(), 'Note 2 current-year (2040) total must match Income Statement revenue');
const note7 = customized.notes.find(n => n.noteNumber === 7)!;
assert.equal(note7.tableData?.rows[0].values[0], '0', 'Note 7 hidden Trade Receivables must stay 0 despite stale cellOverride');

const histNotes = engine.generateNotesToFinancialStatements(
  engine.INITIAL_COMPANY_PROFILE,
  customized.inc.filter(y => !y.isProjected).map(y => y.year),
  customized.inc.filter(y => !y.isProjected),
  customized.bs.filter(y => !y.isProjected),
  customized.cf.filter(y => !y.isProjected),
  customized.fa.filter(y => !y.isProjected),
  { statementConfig: customConfig, labelOverrides, cellOverrides }
);
const histNote2 = histNotes.find(n => n.noteNumber === 2)!;
assert.equal(histNote2.tableData?.totalRow?.values[0], cust2025.revenue.toLocaleString(), 'Scoped historical Note 2 (2025) total must match Income Statement revenue (override 90000 + custom 2500)');
const histNote7 = histNotes.find(n => n.noteNumber === 7)!;
assert.equal(histNote7.tableData?.rows[0].values[0], '0', 'Scoped historical Note 7 (2025) hidden Trade Receivables must stay 0 despite stale cellOverride');

const baseBs2025 = baseline.bs.find(d => d.year === 2025)!;
const custBs2025 = customized.bs.find(d => d.year === 2025)!;
assert.equal(custBs2025.tradeReceivables, 0, 'Deselected Trade Receivables should be 0');
assert(
  Math.abs(custBs2025.balanceCheck + baseBs2025.tradeReceivables) < 0.05,
  `Expected BS balanceCheck to reflect deselected Trade Receivables (-${baseBs2025.tradeReceivables}), got ${custBs2025.balanceCheck}`
);

// Verify Note 12.00 in buildProjectedQuantNoteRows matches Projected P&L Admin & Selling Expenses
const quantNotes = engine.buildProjectedQuantNoteRows(customized.proj, { statementConfig: customConfig, labelOverrides });
const qNote12 = quantNotes.find(r => r.noteNo === '12.00')!;
const firstProj = customized.proj[0];
const expectedPlAdmin = firstProj.adminAndSellingExpenses + engine.getVisibleCustomItemsSum(customConfig, 'income', firstProj.year, 'Operating Expenses');
assert.equal(qNote12.values[0], expectedPlAdmin, 'Projected Note 12.00 must match Projected P&L Admin & Selling Expenses including custom Operating Expenses');

// Verify BS remains balanced when toggling Income Statement + FA category without dropping BS lines
const incomeOnlyCustom = engine.DEFAULT_STATEMENT_CONFIG_ITEMS.map(item =>
  item.id === 'inc_costOfSales' || item.id === 'fa_office_equipment' ? { ...item, isVisible: false } : item
);
const incomeRes = runPipeline(incomeOnlyCustom);
incomeRes.bs.forEach(b => {
  assert(Math.abs(b.balanceCheck) < 0.05, `Income-only customized BS unbalanced in ${b.year}: ${b.balanceCheck}`);
});

const custCf2025 = customized.cf.find(d => d.year === 2025)!;
assert.equal(custCf2025.adjDepreciation, 0, 'Deselected Depreciation add-back should be 0');
assert.equal(custCf2025.reconciliationStatus, 'RECONCILED', 'Customized CF should reconcile cleanly');

const custFa2025 = customized.fa.find(d => d.year === 2025)!;
assert(!custFa2025.categories.some(c => c.name === 'Office Equipment & Fixtures'), 'Deselected FA category must be excluded');
assert(custFa2025.totalNetBookValue < baseline.fa.find(d => d.year === 2025)!.totalNetBookValue, 'FA Total NBV must decrease');

// 3. Export workbook synchronization & formula validation
const syncOpts = { statementConfig: customConfig, labelOverrides, cellOverrides };
exportsMod.exportCompletePackageExcel(
  engine.INITIAL_COMPANY_PROFILE,
  customized.inc,
  customized.bs,
  customized.cf,
  customized.eq,
  customized.fa,
  customized.notes,
  'test_complete.xlsx',
  customized.dscr,
  syncOpts
);
exportsMod.exportPreviousFinancialsExcel(engine.INITIAL_COMPANY_PROFILE, customized.prev.slice(0, 2), customized.notes.slice(0, 5), syncOpts);
exportsMod.exportProjectedFinancialsExcel(engine.INITIAL_COMPANY_PROFILE, customized.proj, syncOpts);

assert.equal(capturedWorkbooks.length, 3, `Expected 3 captured workbooks, got ${capturedWorkbooks.length}`);

for (const { filename, wb } of capturedWorkbooks) {
  validateFormulaCells(wb);
  if (filename === 'test_complete.xlsx') {
    const incCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['1. Income Statement']);
    const bsCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['2. Balance Sheet']);
    const cfCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['3. Cash Flows']);
    const eqCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['4. Changes in Equity']);
    const faCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['5. Fixed Assets']);
    assert(!incCsv.includes('Cost of Sales'), 'Complete Package Income Statement included deselected Cost of Sales');
    assert(!incCsv.includes('Finance Cost'), 'Complete Package Income Statement included deselected Finance Cost');
    assert(incCsv.includes('Turnover & Gross Sales'), 'Complete Package Income Statement missing custom label');
    assert(incCsv.includes('Royalty Grant Income'), 'Complete Package Income Statement missing custom row');
    assert(incCsv.includes('Custom CSR Levy'), 'Complete Package Income Statement missing custom opex row');
    assert(incCsv.includes('Export Rebate Revenue'), 'Complete Package Income Statement missing custom revenue row');
    assert(!bsCsv.includes('Trade Receivables'), 'Complete Package BS included deselected Trade Receivables');
    assert(bsCsv.includes('Cash at Bank & Hand'), 'Complete Package BS missing custom BS label');
    assert(!cfCsv.includes('Adjustments for Depreciation'), 'Complete Package CF included deselected Depreciation');
    assert(!eqCsv.includes('Dividends Paid'), 'Complete Package Equity included deselected Dividends Paid');
    assert(eqCsv.includes('TOTAL OPENING EQUITY'), 'Complete Package Equity missing TOTAL OPENING EQUITY');
    assert(!faCsv.includes('Office Equipment & Fixtures'), 'Complete Package FA included deselected FA category');
    assert(faCsv.includes('Dep. Disposals'), 'Complete Package FA missing Dep. Disposals header');
  } else if (filename.includes('Previous_Financials')) {
    const bsCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['1. Financial Position']);
    const isCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['2. Comprehensive Income']);
    const cfCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['3. Cash Flows']);
    const eqCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['4. Changes in Equity']);
    assert(isCsv.includes('Turnover & Gross Sales'), 'Previous Excel missing custom income label');
    assert(bsCsv.includes('Cash at Bank & Hand'), 'Previous Excel missing custom BS label');
    assert(!isCsv.includes('Less- Cost of Goods Sold'), 'Previous Excel included deselected COGS');
    assert(!isCsv.includes('Less: Financial Exp.'), 'Previous Excel included deselected Financial Exp');
    assert(!bsCsv.includes('Bills Receivable'), 'Previous Excel included deselected Trade Receivables');
    assert(!cfCsv.includes('Depreciation of Tangible Fixed Assets'), 'Previous Excel included deselected CF Depreciation');
    assert(!eqCsv.includes('Dividends Paid'), 'Previous Excel included deselected Dividends Paid');
  } else if (filename.includes('Projected_Financials')) {
    const bsCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['1. Financial Position']);
    const plCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['2. Profit or Loss']);
    const cfCsv = xlsxCjs.utils.sheet_to_csv(wb.Sheets['3. Cash Flows']);
    const plRowsJson = xlsxCjs.utils.sheet_to_json<(string | number)[]>(wb.Sheets['2. Profit or Loss'], { header: 1 });
    const qnRowsJson = xlsxCjs.utils.sheet_to_json<(string | number)[]>(wb.Sheets['4. Notes to Financials'], { header: 1 });
    const pl12Row = plRowsJson.find(r => r[1] === '12.00');
    const qn12Row = qnRowsJson.find(r => r[0] === '12.00');
    assert(pl12Row && qn12Row, 'Projected Excel missing Note 12.00 row');
    assert.equal(qn12Row[2], expectedPlAdmin, 'Projected Notes sheet 12.00 value mismatch with expectedPlAdmin');
    assert.equal(pl12Row[2], qn12Row[2], 'Projected P&L 12.00 value out of sync with Notes sheet 12.00');
    assert(plCsv.includes('Turnover & Gross Sales'), 'Projected Excel missing custom income label');
    assert(bsCsv.includes('Cash at Bank & Hand'), 'Projected Excel missing custom BS label');
    assert(!plCsv.includes('Cost of Goods Sold'), 'Projected Excel included deselected COGS');
    assert(!plCsv.includes('Less: Financial Expenses'), 'Projected Excel included deselected Finance Cost');
    assert(!cfCsv.includes('\nDepreciation,'), 'Projected Excel included deselected CF Depreciation');
  }
}

// 3. First Projected Year 'Loan Received' = 0 verification (Baseline & Repayment Schedule Rescheduling)
const baselineFirstProjCF = baseline.cf.find(c => c.isProjected)!;
const baselineSecondProjCF = baseline.cf.filter(c => c.isProjected)[1];
assert.equal(baselineFirstProjCF.year, 2026);
assert.equal(baselineFirstProjCF.financingProceedsBorrowings, 0, 'First projected year Section 76 proceeds must be 0');
assert(baselineSecondProjCF.financingProceedsBorrowings > 0, 'Second projected year should retain normal proceeds logic');
const baselineProjAsc = [...baseline.proj].sort((a, b) => a.year - b.year);
assert.equal(baselineProjAsc[0].year, 2026);
assert.equal(baselineProjAsc[0].cfLoanReceived, 0, 'First projected year Projected Format cfLoanReceived must be 0');
assert.equal(baselineProjAsc[0].cfRedeemLongTermDebt, -baselineFirstProjCF.repaymentOfBorrowings, 'First projected year debt redemption must equal -repaymentOfBorrowings');
assert.equal(baselineProjAsc[0].cfNetCashFinancing, baselineFirstProjCF.netCashFromFinancing, 'First projected year financing cash flow must tie out across formats');

// Verify baseline Excel Projected Financials sheet has 0 for 2026 'Loan Received'
capturedWorkbooks.length = 0;
exportsMod.exportProjectedFinancialsExcel(engine.INITIAL_COMPANY_PROFILE, baseline.proj);
const baselineProjWb = capturedWorkbooks[0].wb;
validateFormulaCells(baselineProjWb);
const baseCfJson = xlsxCjs.utils.sheet_to_json<(string | number)[]>(baselineProjWb.Sheets['3. Cash Flows'], { header: 1 });
const baseCfHeader = baseCfJson.find(r => r[0] === '1. CASH FLOW FROM OPERATING ACTIVITIES:')!;
const baseCol2026 = baseCfHeader.indexOf(2026);
const baseLoanRecRow = baseCfJson.find(r => r[0] === 'Loan Received')!;
assert(baseCol2026 > 0 && baseLoanRecRow, 'Baseline Projected Excel missing 2026 column or Loan Received row');
assert.equal(baseLoanRecRow[baseCol2026], 0, 'Baseline Projected Excel 2026 Loan Received must be 0');

// Verify Repayment Schedule implemented in first projected year (existing loan proposed for rescheduling)
const reschedHistStore = Object.fromEntries(engine.INITIAL_HISTORICAL_INPUTS.map(i => [i.year, i]));
const reschedResult = engine.implementRepaymentScheduleIntoModel({
  schedule: {
    fileName: 'Rescheduling_Schedule.pdf',
    fileType: 'pdf',
    rows: [],
    annualSummaries: [
      {
        financialYear: 2026,
        installmentCount: 12,
        totalProfit: 280000,
        totalPrincipal: 300000,
        totalInstallment: 580000,
        openingBalance: 2430778.97,
        closingBalance: 2130778.97,
        disbursementAmount: 2430778.97,
      },
      {
        financialYear: 2027,
        installmentCount: 12,
        totalProfit: 250000,
        totalPrincipal: 350000,
        totalInstallment: 600000,
        openingBalance: 2130778.97,
        closingBalance: 1780778.97,
        disbursementAmount: 0,
      },
    ],
    metadata: {
      investmentMade: 2430778.97,
      disbursementDate: '01/01/2026',
      firstInstallmentDate: '01/02/2026',
      expiryDate: '01/01/2032',
      interval: 'Monthly',
      numberOfInstallments: 96,
      dateFormatDetected: 'DD/MM/YYYY',
    },
    grandTotalProfit: 530000,
    grandTotalPrincipal: 650000,
    grandTotalInstallment: 1180000,
    isValid: true,
    validationWarnings: [],
  },
  historicalYearsStore: reschedHistStore,
  customProjectedOverrides: {},
  assumptions: engine.DEFAULT_PROJECTION_ASSUMPTIONS,
  autoIncludeScheduleYears: true,
  alignHistoricalScale: true,
});
const reschedOrdered = [
  ...(reschedResult.assumptions.selectedHistoricalYears || []).map(y => reschedResult.historicalYearsStore[y]),
  ...(reschedResult.assumptions.selectedProjectedYears || []).map(y => reschedResult.customProjectedOverrides[y]),
];
const rFa = engine.calculateScheduleOfFixedAssets(reschedOrdered);
const rInc = engine.calculateIncomeStatement(reschedOrdered, rFa);
const rCf = engine.calculateCashFlowStatement(reschedOrdered, rInc, rFa);
const rEq = engine.calculateEquityStatement(reschedOrdered, rInc);
const rBs = engine.calculateBalanceSheet(reschedOrdered, rEq, rFa, rCf, undefined, rInc);
const rProj = engine.buildProjectedFormatData(reschedOrdered, rInc, rBs, rCf, rEq, rFa);
const rProjAsc = [...rProj].sort((a, b) => a.year - b.year);
const rFirstCF = rCf.find(c => c.isProjected)!;
assert.equal(rFirstCF.year, 2026);
assert.equal(rFirstCF.financingProceedsBorrowings, 0, 'Rescheduled Year 1 Section 76 proceeds must be 0');
assert.equal(rFirstCF.reconciliationStatus, 'RECONCILED', 'Rescheduled Year 1 CF must be reconciled');
assert(Math.abs(rBs.find(b => b.year === 2026)!.balanceCheck) < 0.05, 'Rescheduled Year 1 BS must balance');
assert.equal(rProjAsc[0].cfLoanReceived, 0, 'Rescheduled Year 1 Projected Format cfLoanReceived must be 0');
assert.equal(rProjAsc[0].cfRedeemLongTermDebt, -300000, 'Rescheduled Year 1 debt redemption must match exact scheduled principal');
assert.equal(rProjAsc[0].cfNetCashFinancing, rFirstCF.netCashFromFinancing, 'Rescheduled Year 1 financing CF must tie out');
assert.equal(rProjAsc[0].totalAssets, rProjAsc[0].totalEquityAndLiabilities, 'Rescheduled Year 1 Projected BS must balance');

// 4. Realistic and non-uniform proportional allocation & deterministic repeatability checks
const baseYearInput = engine.INITIAL_HISTORICAL_INPUTS[engine.INITIAL_HISTORICAL_INPUTS.length - 1];
const projRepeat1 = engine.generateProjectedYears(baseYearInput, engine.DEFAULT_PROJECTION_ASSUMPTIONS);
const projRepeat2 = engine.generateProjectedYears(baseYearInput, engine.DEFAULT_PROJECTION_ASSUMPTIONS);
assert.deepEqual(projRepeat1, projRepeat2, 'Projected engine output must be 100% deterministic across calls');

const pctGrowth = (curr: number, prev: number) => Number((((curr - prev) / Math.max(1, Math.abs(prev))) * 100).toFixed(2));
const baseAsc = [...baseline.allInputs].sort((a, b) => a.year - b.year);
const revYoyRates = new Set<number>();
const cosYoyRates = new Set<number>();
const opexYoyRates = new Set<number>();
const invYoyRates = new Set<number>();
const recYoyRates = new Set<number>();
const payYoyRates = new Set<number>();

for (let i = 1; i < baseAsc.length; i++) {
  const c = baseAsc[i];
  const p = baseAsc[i - 1];
  if (!c.isProjected) continue;
  const rPct = pctGrowth(c.salesRevenue, p.salesRevenue);
  const cPct = pctGrowth(c.costOfSales, p.costOfSales);
  const oPct = pctGrowth(c.operatingExpenses, p.operatingExpenses);
  const iPct = pctGrowth(c.closingInventory, p.closingInventory);
  const dPct = pctGrowth(c.tradeReceivables, p.tradeReceivables);
  const pPct = pctGrowth(c.tradePayables, p.tradePayables);
  revYoyRates.add(rPct);
  cosYoyRates.add(cPct);
  opexYoyRates.add(oPct);
  invYoyRates.add(iPct);
  recYoyRates.add(dPct);
  payYoyRates.add(pPct);

  if (c.year >= 2028 && c.year <= 2035) {
    const distinctComponentRates = new Set([rPct, cPct, oPct, iPct, dPct, pPct]);
    assert(
      distinctComponentRates.size >= 4,
      `Expected varied component YoY rates in ${c.year}, got ${Array.from(distinctComponentRates).join(', ')}`
    );
  }
}

assert(revYoyRates.size >= 5, `Expected varied annual revenue growth rates, got ${Array.from(revYoyRates).join(', ')}`);
assert(cosYoyRates.size >= 5, `Expected varied annual COGS growth rates, got ${Array.from(cosYoyRates).join(', ')}`);
assert(opexYoyRates.size >= 5, `Expected varied annual OpEx growth rates, got ${Array.from(opexYoyRates).join(', ')}`);
assert(invYoyRates.size >= 5, `Expected varied annual inventory growth rates, got ${Array.from(invYoyRates).join(', ')}`);
assert(recYoyRates.size >= 5, `Expected varied annual receivable growth rates, got ${Array.from(recYoyRates).join(', ')}`);
assert(payYoyRates.size >= 5, `Expected varied annual payable growth rates, got ${Array.from(payYoyRates).join(', ')}`);

// Fixed Asset Schedule category-level horizontal roll-forward & vertical sum integrity
for (const faRow of baseline.fa) {
  let sumOp = 0;
  let sumAdd = 0;
  let sumDisp = 0;
  let sumCl = 0;
  let sumDepOp = 0;
  let sumDepChg = 0;
  let sumDepDisp = 0;
  let sumDepCl = 0;
  for (const cat of faRow.categories) {
    assert(Math.abs(cat.costOpening + cat.additions - cat.disposals - cat.costClosing) < 0.05, `FA horizontal cost tie failed in ${faRow.year} (${cat.name})`);
    assert(Math.abs(cat.depOpening + cat.depCharge - cat.depDisposals - cat.depClosing) < 0.05, `FA horizontal dep tie failed in ${faRow.year} (${cat.name})`);
    sumOp += cat.costOpening;
    sumAdd += cat.additions;
    sumDisp += cat.disposals;
    sumCl += cat.costClosing;
    sumDepOp += cat.depOpening;
    sumDepChg += cat.depCharge;
    sumDepDisp += cat.depDisposals;
    sumDepCl += cat.depClosing;
  }
  assert(Math.abs(sumOp - faRow.totalCostOpening) < 0.05, `FA vertical opening cost mismatch in ${faRow.year}`);
  assert(Math.abs(sumAdd - faRow.totalAdditions) < 0.05, `FA vertical additions mismatch in ${faRow.year}`);
  assert(Math.abs(sumDisp - faRow.totalDisposals) < 0.05, `FA vertical disposals mismatch in ${faRow.year}`);
  assert(Math.abs(sumCl - faRow.totalCostClosing) < 0.05, `FA vertical closing cost mismatch in ${faRow.year}`);
  assert(Math.abs(sumDepOp - faRow.totalDepOpening) < 0.05, `FA vertical opening dep mismatch in ${faRow.year}`);
  assert(Math.abs(sumDepChg - faRow.totalDepCharge) < 0.05, `FA vertical dep charge mismatch in ${faRow.year}`);
  assert(Math.abs(sumDepDisp - faRow.totalDepDisposals) < 0.05, `FA vertical dep disposals mismatch in ${faRow.year}`);
  assert(Math.abs(sumDepCl - faRow.totalDepClosing) < 0.05, `FA vertical closing dep mismatch in ${faRow.year}`);
}

// Notes disaggregation sums cleanly & varies weights between current and prior year
for (const nn of [2, 3, 4, 6, 7, 8]) {
  const note = baseline.notes.find(n => n.noteNumber === nn)!;
  const rows = note.tableData?.rows || [];
  const tot = note.tableData?.totalRow!;
  const parseNum = (s: string) => Number(String(s).replace(/,/g, ''));
  const sum0 = rows.reduce((acc, r) => acc + parseNum(r.values[0]), 0);
  const sum1 = rows.reduce((acc, r) => acc + parseNum(r.values[1]), 0);
  const tot0 = parseNum(tot.values[0]);
  const tot1 = parseNum(tot.values[1]);
  assert.equal(sum0, tot0, `Note ${nn} current-year breakdown sum must match totalRow`);
  assert.equal(sum1, tot1, `Note ${nn} prior-year breakdown sum must match totalRow`);
  const firstShare0 = parseNum(rows[0].values[0]) / Math.max(1, tot0);
  const firstShare1 = parseNum(rows[0].values[1]) / Math.max(1, tot1);
  assert(Math.abs(firstShare0 - firstShare1) > 0.001, `Note ${nn} current and prior year disaggregation shares should vary naturally`);
}

// Rescheduled model multi-year integrity & DSCR headroom
assert(reschedResult.minEbitdaDscr >= 1.25, `Expected healthy EBITDA DSCR >= 1.25, got ${reschedResult.minEbitdaDscr}`);
rBs.forEach(b => assert(Math.abs(b.balanceCheck) < 0.05, `Rescheduled BS unbalanced in ${b.year}: ${b.balanceCheck}`));
rCf.forEach(c => assert.equal(c.reconciliationStatus, 'RECONCILED', `Rescheduled CF unreconciled in ${c.year}`));

// 5. Mandatory inclusion of 'Repayment of Borrowings & Financing' & Automatic omission of nil-value fields
const baseCfCsv = xlsxCjs.utils.sheet_to_csv(baselineProjWb.Sheets['3. Cash Flows']);
const baseProjBsCsv = xlsxCjs.utils.sheet_to_csv(baselineProjWb.Sheets['1. Financial Position']);
assert(
  baseCfCsv.includes('Repayment of Borrowings & Financing'),
  'Baseline Projected Excel Cash Flows must include Repayment of Borrowings & Financing'
);
assert(
  !baseCfCsv.includes('To Redeem Long Term Debt'),
  'Legacy label To Redeem Long Term Debt must not appear in Projected Excel Cash Flows'
);
const baseRepayRow = baseCfJson.find(r => r[0] === 'Repayment of Borrowings & Financing')!;
assert(baseRepayRow, 'Missing Repayment of Borrowings & Financing row in Projected Excel');
assert.equal(
  baseRepayRow[baseCol2026],
  -baselineFirstProjCF.repaymentOfBorrowings,
  'Projected Excel 2026 Repayment of Borrowings & Financing must equal -repaymentOfBorrowings'
);
const baseCol2027 = baseCfHeader.indexOf(2027);
assert.equal(
  baseRepayRow[baseCol2027],
  -baselineSecondProjCF.repaymentOfBorrowings,
  'Projected Excel 2027 Repayment of Borrowings & Financing must be gross (-repaymentOfBorrowings), not netted against loan proceeds'
);
assert.equal(
  baseLoanRecRow[baseCol2027],
  baselineSecondProjCF.financingProceedsBorrowings,
  'Projected Excel 2027 Loan Received must show gross proceeds'
);

// Nil-value fields must be omitted in Projected Excel
for (const nilLabel of [
  'Increase/(Decrease) in Current Liabilities',
  'Preliminary Expenses',
  'Purchase Investment In Debt Or Securities',
  'Issuance Of Debt (Bond & Notes)',
  'Issuance Of Bonds Payable',
  'To Stockholders as Dividend',
  'Additional Capital',
]) {
  assert(!baseCfCsv.includes(nilLabel), `Nil-value Cash Flow row "${nilLabel}" should be omitted in Projected Excel`);
}
assert(!baseProjBsCsv.includes('Additional Capital'), 'Nil-value Additional Capital should be omitted in Projected Balance Sheet');

// Verify Word/PDF row builders match Excel rules
const { cfRows: baseProjTextCf, bsRows: baseProjTextBs, quantNotesRows: baseQuantNotes } =
  exportsMod.buildProjectedTextTables(baseline.proj);
assert(
  baseProjTextCf.some(r => r[0] === 'Repayment of Borrowings & Financing'),
  'Projected Word/PDF cfRows must include Repayment of Borrowings & Financing'
);
assert(
  !baseProjTextCf.some(r => String(r[0]).includes('To Redeem Long Term Debt')),
  'Projected Word/PDF cfRows must not contain legacy label'
);
assert(
  !baseProjTextCf.some(r => r[0] === 'Issuance Of Debt (Bond & Notes)' || r[0] === 'Preliminary Expenses'),
  'Projected Word/PDF cfRows must omit nil rows'
);
assert(
  !baseProjTextBs.some(r => r[0] === 'Additional Capital'),
  'Projected Word/PDF bsRows must omit nil Additional Capital'
);
// Quantitative notes filter nil detail rows (exclude header/spacer rows whose value cells are '')
assert(
  !baseQuantNotes.some(r => r.slice(2).some(v => v !== '') && r.slice(2).every(v => exportsMod.isNilExportValue(v))),
  'Quantitative Notes Word/PDF/Excel should omit completely nil numeric lines'
);

// Verify Section 76 Complete Package (baseline) omits nil detail lines while preserving Repayment of Borrowings & Financing
capturedWorkbooks.length = 0;
exportsMod.exportCompletePackageExcel(
  engine.INITIAL_COMPANY_PROFILE,
  baseline.inc,
  baseline.bs,
  baseline.cf,
  baseline.eq,
  baseline.fa,
  baseline.notes,
  'baseline_complete.xlsx',
  baseline.dscr
);
const baseCompWb = capturedWorkbooks[0].wb;
validateFormulaCells(baseCompWb);
const baseCompCfCsv = xlsxCjs.utils.sheet_to_csv(baseCompWb.Sheets['3. Cash Flows']);
assert(
  baseCompCfCsv.includes('Repayment of Borrowings & Financing'),
  'Section 76 Complete Excel must include Repayment of Borrowings & Financing'
);
assert(
  !baseCompCfCsv.includes('Purchase of Intangible Assets') && !baseCompCfCsv.includes('Dividends Received'),
  'Section 76 Complete Excel must omit all-zero detail lines'
);
const syncCfRows = exportsMod.buildSyncCashFlowTextRows(baseline.cf);
assert(
  syncCfRows.some(r => r[0] === 'Repayment of Borrowings & Financing'),
  'Section 76 Word/PDF cfRows must include Repayment of Borrowings & Financing'
);
assert(
  !syncCfRows.some(r => r[0] === 'Purchase of Intangible Assets' || r[0] === 'Dividends Received'),
  'Section 76 Word/PDF cfRows must omit all-zero detail lines'
);

// Verify Rescheduled Projected Export reflects exact rescheduled -300000 repayment
capturedWorkbooks.length = 0;
exportsMod.exportProjectedFinancialsExcel(engine.INITIAL_COMPANY_PROFILE, rProj);
const reschedProjWb = capturedWorkbooks[0].wb;
validateFormulaCells(reschedProjWb);
const reschedCfJson = xlsxCjs.utils.sheet_to_json<(string | number)[]>(reschedProjWb.Sheets['3. Cash Flows'], { header: 1 });
const reschedHdr = reschedCfJson.find(r => r[0] === '1. CASH FLOW FROM OPERATING ACTIVITIES:')!;
const reschedCol2026 = reschedHdr.indexOf(2026);
const reschedRepayRow = reschedCfJson.find(r => r[0] === 'Repayment of Borrowings & Financing')!;
assert.equal(reschedRepayRow[reschedCol2026], -300000, 'Rescheduled Projected Excel 2026 repayment must equal -300000');

// 6. PRECEDENCE TEST: Even when Repayment of Borrowings & Financing is 0 across all exported years (or deselected),
// it MUST still appear in downloaded Cash Flows, while nil-value items (e.g., 1-year 2026 Loan Received = 0 & empty Inflows header) ARE omitted!
const singleZeroProj = baseline.proj
  .filter(d => d.year === 2026)
  .map(d => ({
    ...d,
    cfLoanReceived: 0,
    cfRedeemLongTermDebt: 0,
    cfNetCashFinancing: 0,
    cfSalePPE: 0,
    cfPurchasePPE: 0,
    cfNetCashInvesting: 0,
  }));
capturedWorkbooks.length = 0;
exportsMod.exportProjectedFinancialsExcel(engine.INITIAL_COMPANY_PROFILE, singleZeroProj);
const zeroPrecedenceWb = capturedWorkbooks[0].wb;
validateFormulaCells(zeroPrecedenceWb);
const zeroCfJson = xlsxCjs.utils.sheet_to_json<(string | number)[]>(zeroPrecedenceWb.Sheets['3. Cash Flows'], { header: 1 });
const zeroCfCsv = xlsxCjs.utils.sheet_to_csv(zeroPrecedenceWb.Sheets['3. Cash Flows']);
const zeroRepayRow = zeroCfJson.find(r => r[0] === 'Repayment of Borrowings & Financing');
assert(zeroRepayRow, 'PRECEDENCE FAILED: Repayment of Borrowings & Financing must appear even when all values are 0');
assert.equal(zeroRepayRow[1], 0, 'Zero-repayment row must display 0 when nil');
assert(!zeroCfCsv.includes('Loan Received'), 'Nil 1-year 2026 Loan Received (0) must be omitted');
// Since Investing has no inflows or outflows and Financing has no inflows in singleZeroProj, 'Inflows:' must not appear at all!
assert(!zeroCfCsv.includes('Inflows:'), 'Empty Inflows: subheader must be omitted when no inflow items exist');

const { cfRows: zeroTextCf } = exportsMod.buildProjectedTextTables(singleZeroProj);
assert(
  zeroTextCf.some(r => r[0] === 'Repayment of Borrowings & Financing' && r[1] === 0),
  'Projected Word/PDF cfRows must retain Repayment of Borrowings & Financing even when value is 0'
);
assert(!zeroTextCf.some(r => r[0] === 'Loan Received' || r[0] === 'Inflows:'), 'Projected Word/PDF cfRows must omit nil Loan Received and empty Inflows:');

// Precedence even when cf_borrow_repay is deselected in StatementConfig
const deselectedRepayCfg = engine.DEFAULT_STATEMENT_CONFIG_ITEMS.map(item =>
  item.id === 'cf_borrow_repay' ? { ...item, isVisible: false } : item
);
const deselectedRes = runPipeline(deselectedRepayCfg);
const syncZeroRows = exportsMod.buildSyncCashFlowTextRows(deselectedRes.cf, { statementConfig: deselectedRepayCfg });
assert(
  syncZeroRows.some(r => r[0] === 'Repayment of Borrowings & Financing'),
  'PRECEDENCE FAILED: Section 76 Word/PDF must include Repayment of Borrowings & Financing even when deselected'
);

// 7. Irregular & Non-Uniform Proportional Variations (>= 1.2 percentage point minimum spread & natural ups/downs)
const baseYoyPct = (curr: number, prev: number) => ((curr - prev) / Math.max(1, Math.abs(prev))) * 100;
const basePairs: [string, (i: number) => number, (i: number) => number][] = [
  ['Revenue vs COGS', i => baseYoyPct(baseline.inc[i].revenue, baseline.inc[i - 1].revenue), i => baseYoyPct(baseline.inc[i].costOfSales, baseline.inc[i - 1].costOfSales)],
  ['Revenue vs OpEx', i => baseYoyPct(baseline.inc[i].revenue, baseline.inc[i - 1].revenue), i => baseYoyPct(baseline.inc[i].operatingExpenses, baseline.inc[i - 1].operatingExpenses)],
  ['COGS vs OpEx', i => baseYoyPct(baseline.inc[i].costOfSales, baseline.inc[i - 1].costOfSales), i => baseYoyPct(baseline.inc[i].operatingExpenses, baseline.inc[i - 1].operatingExpenses)],
  ['Revenue vs Net Profit', i => baseYoyPct(baseline.inc[i].revenue, baseline.inc[i - 1].revenue), i => baseYoyPct(baseline.inc[i].netProfit, baseline.inc[i - 1].netProfit)],
  ['Inventory vs COGS', i => baseYoyPct(baseline.bs[i].closingInventory, baseline.bs[i - 1].closingInventory), i => baseYoyPct(baseline.inc[i].costOfSales, baseline.inc[i - 1].costOfSales)],
  ['Inventory vs Trade Receivables', i => baseYoyPct(baseline.bs[i].closingInventory, baseline.bs[i - 1].closingInventory), i => baseYoyPct(baseline.bs[i].tradeReceivables, baseline.bs[i - 1].tradeReceivables)],
  ['Trade Receivables vs Other Receivables', i => baseYoyPct(baseline.bs[i].tradeReceivables, baseline.bs[i - 1].tradeReceivables), i => baseYoyPct(baseline.bs[i].otherReceivables, baseline.bs[i - 1].otherReceivables)],
  ['Trade Receivables vs Other Current Assets', i => baseYoyPct(baseline.bs[i].tradeReceivables, baseline.bs[i - 1].tradeReceivables), i => baseYoyPct(baseline.bs[i].otherCurrentAssets, baseline.bs[i - 1].otherCurrentAssets)],
  ['Other Receivables vs Other Current Assets', i => baseYoyPct(baseline.bs[i].otherReceivables, baseline.bs[i - 1].otherReceivables), i => baseYoyPct(baseline.bs[i].otherCurrentAssets, baseline.bs[i - 1].otherCurrentAssets)],
  ['Trade Payables vs Other Payables', i => baseYoyPct(baseline.bs[i].tradePayables, baseline.bs[i - 1].tradePayables), i => baseYoyPct(baseline.bs[i].otherPayables, baseline.bs[i - 1].otherPayables)],
  ['Trade Payables vs COGS', i => baseYoyPct(baseline.bs[i].tradePayables, baseline.bs[i - 1].tradePayables), i => baseYoyPct(baseline.inc[i].costOfSales, baseline.inc[i - 1].costOfSales)],
  ['Finance Cost vs Repayment of Borrowings', i => baseYoyPct(baseline.inc[i].financeCost, baseline.inc[i - 1].financeCost), i => baseYoyPct(baseAsc[i].repaymentOfBorrowings, baseAsc[i - 1].repaymentOfBorrowings)],
  ['Net Profit vs Dividends Paid', i => baseYoyPct(baseline.inc[i].netProfit, baseline.inc[i - 1].netProfit), i => baseYoyPct(baseAsc[i].dividendsPaid, baseAsc[i - 1].dividendsPaid)],
  ['Capex Additions vs Disposals', i => baseYoyPct(baseline.fa[i].totalAdditions, baseline.fa[i - 1].totalAdditions), i => baseYoyPct(baseline.fa[i].totalDisposals, baseline.fa[i - 1].totalDisposals)],
  ['Gross Fixed Assets vs Depreciation', i => baseYoyPct(baseAsc[i].fixedAssetsGross, baseAsc[i - 1].fixedAssetsGross), i => baseYoyPct(baseline.inc[i].depreciationExpense, baseline.inc[i - 1].depreciationExpense)],
];

for (let i = 3; i < baseAsc.length; i++) {
  assert(baseAsc[i].isProjected, `Expected projected year at index ${i}`);
  for (const [label, fnA, fnB] of basePairs) {
    const spread = Math.abs(fnA(i) - fnB(i));
    assert(
      spread >= 1.2,
      `Baseline ${label} in ${baseAsc[i].year} spread ${spread.toFixed(2)}pp < 1.2pp`
    );
  }
}

// Verify natural ups and downs (both positive increases and negative decreases) across projected years
const projIndicesBase = baseAsc.map((y, idx) => ({ y, idx })).filter(e => e.y.isProjected).map(e => e.idx);
const checkHasBothUpAndDown = (name: string, seriesFn: (i: number) => number) => {
  const rates = projIndicesBase.map(i => seriesFn(i));
  assert(rates.some(r => r > 1.0), `${name} must include realistic growth years`);
  assert(rates.some(r => r < -0.5), `${name} must include realistic contraction/reduction years`);
};
checkHasBothUpAndDown('Sales Revenue', i => baseYoyPct(baseline.inc[i].revenue, baseline.inc[i - 1].revenue));
checkHasBothUpAndDown('Cost of Sales', i => baseYoyPct(baseline.inc[i].costOfSales, baseline.inc[i - 1].costOfSales));
checkHasBothUpAndDown('Operating Expenses', i => baseYoyPct(baseline.inc[i].operatingExpenses, baseline.inc[i - 1].operatingExpenses));
checkHasBothUpAndDown('Net Profit', i => baseYoyPct(baseline.inc[i].netProfit, baseline.inc[i - 1].netProfit));
checkHasBothUpAndDown('Closing Inventory', i => baseYoyPct(baseline.bs[i].closingInventory, baseline.bs[i - 1].closingInventory));
checkHasBothUpAndDown('Trade Receivables', i => baseYoyPct(baseline.bs[i].tradeReceivables, baseline.bs[i - 1].tradeReceivables));
checkHasBothUpAndDown('Other Receivables', i => baseYoyPct(baseline.bs[i].otherReceivables, baseline.bs[i - 1].otherReceivables));
checkHasBothUpAndDown('Other Current Assets', i => baseYoyPct(baseline.bs[i].otherCurrentAssets, baseline.bs[i - 1].otherCurrentAssets));
checkHasBothUpAndDown('Trade Payables', i => baseYoyPct(baseline.bs[i].tradePayables, baseline.bs[i - 1].tradePayables));
checkHasBothUpAndDown('Dividends Paid', i => baseYoyPct(baseAsc[i].dividendsPaid, baseAsc[i - 1].dividendsPaid));
checkHasBothUpAndDown('Capex Additions', i => baseYoyPct(baseline.fa[i].totalAdditions, baseline.fa[i - 1].totalAdditions));

// Projected Note 11 direct expenses also maintain >= 1.2pp spread between components
for (let i = 1; i < baselineProjAsc.length; i++) {
  const c = baselineProjAsc[i];
  const p = baselineProjAsc[i - 1];
  const directRates = [
    baseYoyPct(c.note11LabourWages, p.note11LabourWages),
    baseYoyPct(c.note11ElectricityBill, p.note11ElectricityBill),
    baseYoyPct(c.note11CarriageInward, p.note11CarriageInward),
    baseYoyPct(c.note11OtherDirectExpenses, p.note11OtherDirectExpenses),
  ];
  for (let a = 0; a < directRates.length; a++) {
    for (let b = a + 1; b < directRates.length; b++) {
      assert(
        Math.abs(directRates[a] - directRates[b]) >= 1.2,
        `Note 11 direct expenses spread < 1.2pp in ${c.year}: ${directRates.map(r => r.toFixed(2)).join(', ')}`
      );
    }
  }
}

// Verify Rescheduled model preserves exact schedule locks AND >= 1.2pp spread across non-locked pairs
for (let i = 3; i < reschedOrdered.length; i++) {
  const rPairs: [string, number, number][] = [
    ['Revenue vs COGS', baseYoyPct(rInc[i].revenue, rInc[i - 1].revenue), baseYoyPct(rInc[i].costOfSales, rInc[i - 1].costOfSales)],
    ['Revenue vs OpEx', baseYoyPct(rInc[i].revenue, rInc[i - 1].revenue), baseYoyPct(rInc[i].operatingExpenses, rInc[i - 1].operatingExpenses)],
    ['COGS vs OpEx', baseYoyPct(rInc[i].costOfSales, rInc[i - 1].costOfSales), baseYoyPct(rInc[i].operatingExpenses, rInc[i - 1].operatingExpenses)],
    ['Inventory vs COGS', baseYoyPct(rBs[i].closingInventory, rBs[i - 1].closingInventory), baseYoyPct(rInc[i].costOfSales, rInc[i - 1].costOfSales)],
    ['Trade Receivables vs Other Receivables', baseYoyPct(rBs[i].tradeReceivables, rBs[i - 1].tradeReceivables), baseYoyPct(rBs[i].otherReceivables, rBs[i - 1].otherReceivables)],
    ['Trade Payables vs Other Payables', baseYoyPct(rBs[i].tradePayables, rBs[i - 1].tradePayables), baseYoyPct(rBs[i].otherPayables, rBs[i - 1].otherPayables)],
    ['Capex Additions vs Disposals', baseYoyPct(rFa[i].totalAdditions, rFa[i - 1].totalAdditions), baseYoyPct(rFa[i].totalDisposals, rFa[i - 1].totalDisposals)],
    ['Gross FA vs Depreciation', baseYoyPct(reschedOrdered[i].fixedAssetsGross, reschedOrdered[i - 1].fixedAssetsGross), baseYoyPct(rInc[i].depreciationExpense, rInc[i - 1].depreciationExpense)],
  ];
  for (const [label, a, b] of rPairs) {
    assert(
      Math.abs(a - b) >= 1.2,
      `Rescheduled ${label} in ${reschedOrdered[i].year} spread ${Math.abs(a - b).toFixed(2)}pp < 1.2pp`
    );
  }
}

// 8. Enhancement Verifications:
// (A) Two Financial Year-End Options ('June 30' & 'December 31')
const juneProfile = { ...engine.INITIAL_COMPANY_PROFILE, financialYearEnd: 'June 30' as const };
const decProfile = { ...engine.INITIAL_COMPANY_PROFILE, financialYearEnd: 'December 31' as const };
const junePrev = engine.buildPreviousFormatData(
  baseline.allInputs,
  baseline.inc,
  baseline.bs,
  baseline.cf,
  baseline.eq,
  baseline.fa,
  undefined,
  'June 30'
);
assert(junePrev.every(p => p.dateLabel.startsWith('30.06.')), 'June 30 Previous Format must use 30.06.YYYY date labels');
const decPrev = engine.buildPreviousFormatData(
  baseline.allInputs,
  baseline.inc,
  baseline.bs,
  baseline.cf,
  baseline.eq,
  baseline.fa,
  undefined,
  'December 31'
);
assert(decPrev.every(p => p.dateLabel.startsWith('31.12.')), 'December 31 Previous Format must use 31.12.YYYY date labels');

const { rebuildScheduleSummaries } = await import('./src/services/repaymentScheduleParser');
const sampleRows = [
  { id: 'r1', installmentNo: 1, eventType: 'Installment', emiDateRaw: '15/09/2025', emiDateISO: '2025-09-15', financialYear: 2025, instPrincipal: 100000, instProfit: 50000, gestationPayment: 0, overdueGestation: 0, overdueProfit: 0, balance: 900000, totalInstallment: 150000 },
  { id: 'r2', installmentNo: 2, eventType: 'Installment', emiDateRaw: '15/05/2026', emiDateISO: '2026-05-15', financialYear: 2026, instPrincipal: 105000, instProfit: 45000, gestationPayment: 0, overdueGestation: 0, overdueProfit: 0, balance: 795000, totalInstallment: 150000 },
];
const sampleMeta = {
  sourceFileName: 'test_schedule.pdf',
  parserMethod: 'pdf-text' as const,
  investmentMade: 1000000,
  disbursementDate: '01/07/2025',
  firstInstallmentDate: '15/09/2025',
  expiryDate: '15/05/2026',
  interval: 'Monthly',
  numberOfInstallments: 2,
  dateFormatDetected: 'DD/MM/YYYY' as const,
  uploadedAt: '2026-01-01T00:00:00.000Z',
};
const juneRebuilt = rebuildScheduleSummaries(sampleRows, sampleMeta, 'June 30', [2025, 2026]);
assert.equal(juneRebuilt.annualSummaries.length, 1, 'June 30 FY should combine Sep 2025 and May 2026 into FY 2026');
assert.equal(juneRebuilt.annualSummaries[0].financialYear, 2026);
assert.equal(juneRebuilt.annualSummaries[0].totalPrincipal, 205000);
const decRebuilt = rebuildScheduleSummaries(sampleRows, sampleMeta, 'December 31', [2025, 2026]);
assert.equal(decRebuilt.annualSummaries.length, 2, 'December 31 FY should separate Sep 2025 (FY 2025) and May 2026 (FY 2026)');

const juneNotes = engine.generateNotesToFinancialStatements(
  juneProfile,
  baseline.allInputs.map(y => y.year),
  baseline.inc,
  baseline.bs,
  baseline.cf,
  baseline.fa
);
capturedWorkbooks.length = 0;
exportsMod.exportPreviousFinancialsExcel(juneProfile, junePrev.slice(0, 2), juneNotes);
exportsMod.exportProjectedFinancialsExcel(juneProfile, baseline.proj);
exportsMod.exportCompletePackageExcel(juneProfile, baseline.inc, baseline.bs, baseline.cf, baseline.eq, baseline.fa, juneNotes, 'june_complete.xlsx');
for (const { filename, wb } of capturedWorkbooks) {
  validateFormulaCells(wb);
  const allSheetsText = wb.SheetNames.map(n => xlsxCjs.utils.sheet_to_csv(wb.Sheets[n])).join('\n');
  assert(!allSheetsText.includes('31 December') && !allSheetsText.includes('31st December') && !allSheetsText.includes('31.12.'), `June 30 export ${filename} leaked December dates`);
  assert(allSheetsText.includes('June') || allSheetsText.includes('30.06.'), `June 30 export ${filename} missing June 30 dates`);
}

// (B) Year-Wise Editable Long-Term Borrowings (Historical + Projected + Zero value)
const ltModifiedInputs = baseline.allInputs.map((y, idx) => {
  if (y.year === 2024) {
    return engine.deriveDependentFinancials(
      y,
      idx > 0 ? baseline.allInputs[idx - 1] : null,
      engine.DEFAULT_PROJECTION_ASSUMPTIONS,
      { ...y, longTermBorrowings: 1234567, hasCustomLongTermBorrowings: true }
    );
  }
  if (y.year === 2027) {
    return engine.deriveDependentFinancials(
      y,
      baseline.allInputs[idx - 1],
      engine.DEFAULT_PROJECTION_ASSUMPTIONS,
      { ...y, longTermBorrowings: 0, hasCustomLongTermBorrowings: true }
    );
  }
  return y;
});
const ltFa = engine.calculateScheduleOfFixedAssets(ltModifiedInputs);
const ltInc = engine.calculateIncomeStatement(ltModifiedInputs, ltFa);
const ltCf = engine.calculateCashFlowStatement(ltModifiedInputs, ltInc, ltFa);
const ltEq = engine.calculateEquityStatement(ltModifiedInputs, ltInc);
const ltBs = engine.calculateBalanceSheet(ltModifiedInputs, ltEq, ltFa, ltCf, undefined, ltInc);
const ltProj = engine.buildProjectedFormatData(ltModifiedInputs, ltInc, ltBs, ltCf, ltEq, ltFa);
const ltNotes = engine.generateNotesToFinancialStatements(decProfile, ltModifiedInputs.map(y => y.year), ltInc, ltBs, ltCf, ltFa);
const ltQuantNotes = engine.buildProjectedQuantNoteRows(ltProj);
assert.equal(ltBs.find(b => b.year === 2024)!.longTermBorrowings, 1234567, 'Historical 2024 Long-Term Borrowings must reflect custom user input');
assert.equal(ltBs.find(b => b.year === 2027)!.longTermBorrowings, 0, 'Projected 2027 Long-Term Borrowings must reflect zero custom input');
assert.equal(ltProj.find(p => p.year === 2027)!.longTermLoan, 0, 'Projected Format 2027 longTermLoan must reflect 0');
const ltNote7Idx = ltProj.findIndex(p => p.year === 2027);
assert.equal(ltQuantNotes.find(r => r.noteNo === '7.00')!.values[ltNote7Idx], 0, 'Projected Quant Note 7.00 must match user-edited LT borrowings');
ltBs.forEach(b => assert(Math.abs(b.balanceCheck) < 0.05, `LT-edited BS unbalanced in ${b.year}: ${b.balanceCheck}`));
ltCf.forEach(c => assert.equal(c.reconciliationStatus, 'RECONCILED', `LT-edited CF unreconciled in ${c.year}`));
assert(ltNotes.length > 0, 'Notes generation succeeded after LT borrowings edit');

// (C) New Case Clean Reset Integrity (zeroed financial data + empty company profile)
const cleanProfile = engine.createCleanCaseCompanyProfile('June 30');
const cleanHist = engine.createZeroHistoricalInputs([2023, 2024, 2025]);
const cleanProj = engine.generateProjectedYears(cleanHist[cleanHist.length - 1], engine.DEFAULT_PROJECTION_ASSUMPTIONS);
const cleanAll = [...cleanHist, ...cleanProj];
cleanAll.forEach(y => {
  assert.equal(y.salesRevenue, 0, `New Case year ${y.year} salesRevenue must start at 0`);
  assert.equal(y.longTermBorrowings, 0, `New Case year ${y.year} longTermBorrowings must start at 0`);
});
const cFa = engine.calculateScheduleOfFixedAssets(cleanAll);
const cInc = engine.calculateIncomeStatement(cleanAll, cFa);
const cCf = engine.calculateCashFlowStatement(cleanAll, cInc, cFa);
const cEq = engine.calculateEquityStatement(cleanAll, cInc);
const cBs = engine.calculateBalanceSheet(cleanAll, cEq, cFa, cCf, undefined, cInc);
const cNotes = engine.generateNotesToFinancialStatements(cleanProfile, cleanAll.map(y => y.year), cInc, cBs, cCf, cFa);
const cPrev = engine.buildPreviousFormatData(cleanAll, cInc, cBs, cCf, cEq, cFa, undefined, cleanProfile.financialYearEnd);
const cProjFmt = engine.buildProjectedFormatData(cleanAll, cInc, cBs, cCf, cEq, cFa);
cBs.forEach(b => assert(Math.abs(b.balanceCheck) < 0.05, `New Case BS unbalanced in ${b.year}: ${b.balanceCheck}`));
cCf.forEach(c => assert.equal(c.reconciliationStatus, 'RECONCILED', `New Case CF unreconciled in ${c.year}`));

capturedWorkbooks.length = 0;
exportsMod.exportCompletePackageExcel(cleanProfile, cInc, cBs, cCf, cEq, cFa, cNotes);
exportsMod.exportPreviousFinancialsExcel(cleanProfile, cPrev.slice(0, 2), cNotes);
exportsMod.exportProjectedFinancialsExcel(cleanProfile, cProjFmt);
for (const { filename, wb } of capturedWorkbooks) {
  assert(!filename.startsWith('_'), `Empty companyName export filename "${filename}" must use clean fallback prefix`);
  validateFormulaCells(wb);
}

// ============================================================================
// 6. UPLOADED REPAYMENT SCHEDULE AUTOMATION, HISTORICAL INABILITY VS PROJECTED CAPACITY, & UNEVEN RATIOS (>= 1.2pp)
// ============================================================================

// (A) Verify Historical INITIAL_HISTORICAL_INPUTS & Computed Previous Format have uneven YoY spreads >= 1.2pp
const histAsc = [...engine.INITIAL_HISTORICAL_INPUTS].sort((a, b) => a.year - b.year);
const histPrev = engine.buildPreviousFormatData(
  histAsc,
  baseline.inc.filter(i => !i.isProjected),
  baseline.bs.filter(b => !b.isProjected),
  baseline.cf.filter(c => !c.isProjected),
  baseline.eq.filter(e => !e.isProjected),
  baseline.fa.filter(f => !f.isProjected)
).sort((a, b) => a.year - b.year);

const pctDelta = (curr: number, prev: number) => ((curr - prev) / Math.max(1, Math.abs(prev))) * 100;
const histSpreadPairs: [string, number, number][] = [
  ['salesRevenue', pctDelta(histAsc[1].salesRevenue, histAsc[0].salesRevenue), pctDelta(histAsc[2].salesRevenue, histAsc[1].salesRevenue)],
  ['costOfSales', pctDelta(histAsc[1].costOfSales, histAsc[0].costOfSales), pctDelta(histAsc[2].costOfSales, histAsc[1].costOfSales)],
  ['operatingExpenses', pctDelta(histAsc[1].operatingExpenses, histAsc[0].operatingExpenses), pctDelta(histAsc[2].operatingExpenses, histAsc[1].operatingExpenses)],
  ['closingInventory', pctDelta(histAsc[1].closingInventory, histAsc[0].closingInventory), pctDelta(histAsc[2].closingInventory, histAsc[1].closingInventory)],
  ['tradeReceivables', pctDelta(histAsc[1].tradeReceivables, histAsc[0].tradeReceivables), pctDelta(histAsc[2].tradeReceivables, histAsc[1].tradeReceivables)],
  ['salaryAndAllowances', pctDelta(histPrev[1].salaryAndAllowances, histPrev[0].salaryAndAllowances), pctDelta(histPrev[2].salaryAndAllowances, histPrev[1].salaryAndAllowances)],
  ['electricCharges', pctDelta(histPrev[1].electricCharges, histPrev[0].electricCharges), pctDelta(histPrev[2].electricCharges, histPrev[1].electricCharges)],
];
for (const [label, g1, g2] of histSpreadPairs) {
  assert(
    Math.abs(g1 - g2) >= 1.2,
    `Historical YoY spread for ${label} must be >= 1.2pp (got ${g1.toFixed(2)}% vs ${g2.toFixed(2)}%)`
  );
}

// Verify Custom Element Uneven Ratio Series produces >= 1.2pp consecutive spreads
const customUnevenSeries = engine.generateCustomElementUnevenSeries(
  350000,
  [2023, 2024, 2025, 2026, 2027, 2028],
  2023,
  8.0,
  'Sukuk Profit Reserve'
);
const customYears = [2023, 2024, 2025, 2026, 2027, 2028];
for (let i = 2; i < customYears.length; i++) {
  const gPrev = pctDelta(customUnevenSeries[customYears[i - 1]], customUnevenSeries[customYears[i - 2]]);
  const gCurr = pctDelta(customUnevenSeries[customYears[i]], customUnevenSeries[customYears[i - 1]]);
  assert(
    Math.abs(gCurr - gPrev) >= 1.2,
    `Custom element uneven series spread in ${customYears[i]} must be >= 1.2pp (got ${gPrev.toFixed(2)}% vs ${gCurr.toFixed(2)}%)`
  );
}

// (B) Automated Schedule Generation, Running Balance Auto-Calc, and Recalculation
const {
  generateStandardRepaymentSchedule,
  recalculateRepaymentSchedule,
  getAttachedReferenceRepaymentSchedule,
} = await import('./src/services/repaymentScheduleParser');

const genSched = generateStandardRepaymentSchedule({
  investmentMade: 2500000,
  annualProfitRatePct: 14.5,
  numberOfInstallments: 48,
  interval: 'Monthly',
  disbursementDateRaw: '01/01/2026',
  firstInstallmentDateRaw: '01/01/2026',
  financialYearEnd: '31 December',
  selectedModelYears: [2023, 2024, 2025, 2026, 2027, 2028, 2029],
});
assert(genSched.isValid, 'Generated standard repayment schedule must be valid');
assert.equal(genSched.closingBalance, 0, 'Generated schedule closingBalance must amortize cleanly to 0');
assert(Math.abs(genSched.grandTotalPrincipal - 2500000) < 0.05, 'Generated grandTotalPrincipal must equal Investment Made');

// Inline edit row 1 principal + gestation and confirm recalculateRepaymentSchedule updates running balance and totals
const editedSched = recalculateRepaymentSchedule(
  {
    ...genSched,
    rows: genSched.rows.map((r, idx) =>
      idx === 1 ? { ...r, instPrincipal: r.instPrincipal + 1000, gestationPayment: 500 } : r
    ),
  },
  '31 December',
  [2023, 2024, 2025, 2026, 2027, 2028, 2029]
);
assert.equal(editedSched.grandTotalGestationPayment, 500, 'Recalculated schedule must aggregate Gestation Payment');
assert(Math.abs(editedSched.grandTotalDebtService - (editedSched.grandTotalPrincipal + editedSched.grandTotalProfit + 500)) < 0.05, 'Grand total debt service must include gestation payments');
assert(Math.abs(editedSched.rows[1].balance - (2500000 - editedSched.rows[1].instPrincipal)) < 0.05, 'Recalculated row 1 balance must reflect edited principal');

// (C) Client ABABIL Schedule Implementation: Historical Pre-Rescheduling Inability (< 1.0x DSCR) vs Projected Servicing Sufficiency (>= 1.25x DSCR)
const ababilSched = getAttachedReferenceRepaymentSchedule('31 December', [2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032]);
assert.equal(ababilSched.closingBalance, 0, 'ABABIL schedule closingBalance must be 0');
assert.equal(ababilSched.metadata.gestationPeriod, '0 Days');
assert.equal(ababilSched.metadata.lastInstallmentMinimized, 'No');

const ababilHistStore = Object.fromEntries(engine.INITIAL_HISTORICAL_INPUTS.map(i => [i.year, i]));
const ababilResult = engine.implementRepaymentScheduleIntoModel({
  schedule: ababilSched,
  historicalYearsStore: ababilHistStore,
  customProjectedOverrides: {},
  assumptions: engine.DEFAULT_PROJECTION_ASSUMPTIONS,
  autoIncludeScheduleYears: true,
  alignHistoricalScale: true,
});
const ababilOrdered = [
  ...(ababilResult.assumptions.selectedHistoricalYears || []).map(y => ababilResult.historicalYearsStore[y]),
  ...(ababilResult.assumptions.selectedProjectedYears || []).map(y => ababilResult.customProjectedOverrides[y]),
].sort((a, b) => a.year - b.year);

const abFa = engine.calculateScheduleOfFixedAssets(ababilOrdered);
const abInc = engine.calculateIncomeStatement(ababilOrdered, abFa);
const abCf = engine.calculateCashFlowStatement(ababilOrdered, abInc, abFa);
const abEq = engine.calculateEquityStatement(ababilOrdered, abInc);
const abBs = engine.calculateBalanceSheet(ababilOrdered, abEq, abFa, abCf, undefined, abInc);
const abDscr = engine.calculateDebtServiceCoverage(ababilOrdered, abInc, abBs, abCf);
const abProj = engine.buildProjectedFormatData(ababilOrdered, abInc, abBs, abCf, abEq, abFa);

abBs.forEach(b => assert(Math.abs(b.balanceCheck) < 0.05, `ABABIL BS unbalanced in ${b.year}: ${b.balanceCheck}`));
abCf.forEach(c => assert.equal(c.reconciliationStatus, 'RECONCILED', `ABABIL CF not reconciled in ${c.year}`));

const histAbDscr = abDscr.filter(d => !d.isProjected);
const projSchedAbDscr = abDscr.filter(d => d.isProjected && d.isScheduleCovered);
assert(histAbDscr.length >= 3 && projSchedAbDscr.length >= 5, 'ABABIL model must cover both historical and scheduled projected years');
histAbDscr.forEach(h => {
  assert(h.ebitdaDscr < 1.0, `Historical ${h.year} EBITDA DSCR must show inability (< 1.0x) to justify rescheduling, got ${h.ebitdaDscr}`);
  assert.equal(h.capacityStatus, 'DEFICIT', `Historical ${h.year} capacityStatus must be DEFICIT`);
  assert((h.reschedulingContext || '').length > 10, `Historical ${h.year} missing reschedulingContext`);
});
projSchedAbDscr.forEach(p => {
  assert(p.ebitdaDscr >= 1.25, `Projected scheduled ${p.year} EBITDA DSCR must be >= 1.25x, got ${p.ebitdaDscr}`);
  assert(p.cashDscr >= 1.15, `Projected scheduled ${p.year} Cash DSCR must be >= 1.15x, got ${p.cashDscr}`);
  assert(p.closingCash > 0, `Projected scheduled ${p.year} Closing Cash must be positive, got ${p.closingCash}`);
  assert(p.capacityStatus === 'STRONG' || p.capacityStatus === 'ADEQUATE', `Projected scheduled ${p.year} status must be STRONG/ADEQUATE`);
});

// (D) Export Workbooks containing 9-Column Repayment Schedule & DSCR Justification pass formula validation
const abNotes = engine.generateNotesToFinancialStatements(
  engine.INITIAL_COMPANY_PROFILE,
  ababilOrdered.map(y => y.year),
  abInc,
  abBs,
  abCf,
  abFa
);
capturedWorkbooks.length = 0;
exportsMod.exportRepaymentScheduleExcel(engine.INITIAL_COMPANY_PROFILE, ababilSched, abDscr);
exportsMod.exportCompletePackageExcel(
  engine.INITIAL_COMPANY_PROFILE,
  abInc,
  abBs,
  abCf,
  abEq,
  abFa,
  abNotes,
  'ababil_complete_with_schedule.xlsx',
  abDscr,
  undefined,
  ababilSched
);
exportsMod.exportProjectedFinancialsExcel(
  engine.INITIAL_COMPANY_PROFILE,
  abProj,
  undefined,
  undefined,
  ababilSched,
  abDscr
);
assert.equal(capturedWorkbooks.length, 3, 'Expected 3 schedule-integrated Excel exports');
for (const { filename, wb } of capturedWorkbooks) {
  validateFormulaCells(wb);
  const hasScheduleSheet = wb.SheetNames.some(s => s.includes('Repayment Schedule'));
  assert(hasScheduleSheet, `Export ${filename} must include Repayment Schedule worksheet`);
}

// ============================================================================
// 9. CHARTERED ACCOUNTANCY LOAN RESCHEDULING, AUDIT ENGINE & PDF-PARITY EXPORTS
// ============================================================================

// (A) Editable Loan Interest Rate (15.50% default + cascading recalculation)
assert.equal(
  engine.DEFAULT_PROJECTION_ASSUMPTIONS.loanInterestRatePct,
  15.5,
  'Default loanInterestRatePct must be 15.50%'
);

const rateSampleIdx = baseline.allInputs.findIndex(y => y.year === 2024);
const rateLowYear = engine.recalculateYearOnLoanOrRateChange(
  baseline.allInputs[rateSampleIdx],
  baseline.allInputs[rateSampleIdx - 1],
  { ...engine.DEFAULT_PROJECTION_ASSUMPTIONS, loanInterestRatePct: 12.5 },
  { loanRateOverridePct: 12.5, recalculateFinanceCostFromDebt: true }
);
const rateHighYear = engine.recalculateYearOnLoanOrRateChange(
  baseline.allInputs[rateSampleIdx],
  baseline.allInputs[rateSampleIdx - 1],
  { ...engine.DEFAULT_PROJECTION_ASSUMPTIONS, loanInterestRatePct: 18.0 },
  { loanRateOverridePct: 18.0, recalculateFinanceCostFromDebt: true }
);
assert(
  rateHighYear.financeCost > rateLowYear.financeCost,
  `Higher loan interest rate (18%) must increase Finance Cost (${rateHighYear.financeCost}) vs 12.5% (${rateLowYear.financeCost})`
);
assert.equal(
  rateHighYear.operatingExpenses,
  rateLowYear.operatingExpenses,
  'Operating Expenses must remain preserved when loan interest rate changes'
);

const rateHighInputs = baseline.allInputs.map((y, idx) =>
  idx === rateSampleIdx ? rateHighYear : y
);
const rateFa = engine.calculateScheduleOfFixedAssets(rateHighInputs);
const rateInc = engine.calculateIncomeStatement(rateHighInputs, rateFa);
const rateCf = engine.calculateCashFlowStatement(rateHighInputs, rateInc, rateFa);
const rateEq = engine.calculateEquityStatement(rateHighInputs, rateInc);
const rateBs = engine.calculateBalanceSheet(rateHighInputs, rateEq, rateFa, rateCf, undefined, rateInc);
rateBs.forEach(b =>
  assert(Math.abs(b.balanceCheck) < 0.05, `Loan-rate-recalculated BS unbalanced in ${b.year}: ${b.balanceCheck}`)
);
rateCf.forEach(c =>
  assert.equal(c.reconciliationStatus, 'RECONCILED', `Loan-rate-recalculated CF unreconciled in ${c.year}`)
);

// (B) Editable Historical & Projected Short-Term + Long-Term Borrowings Propagation
const debtModifiedInputs = baseline.allInputs.map((y, idx) => {
  if (y.year === 2024) {
    return engine.deriveDependentFinancials(
      y,
      idx > 0 ? baseline.allInputs[idx - 1] : null,
      engine.DEFAULT_PROJECTION_ASSUMPTIONS,
      {
        ...y,
        shortTermBorrowings: 456789,
        hasCustomShortTermBorrowings: true,
        longTermBorrowings: 2345678,
        hasCustomLongTermBorrowings: true,
      }
    );
  }
  if (y.year === 2027) {
    return engine.deriveDependentFinancials(
      y,
      baseline.allInputs[idx - 1],
      engine.DEFAULT_PROJECTION_ASSUMPTIONS,
      {
        ...y,
        shortTermBorrowings: 321000,
        hasCustomShortTermBorrowings: true,
        longTermBorrowings: 1650000,
        hasCustomLongTermBorrowings: true,
      }
    );
  }
  return y;
});
const dmFa = engine.calculateScheduleOfFixedAssets(debtModifiedInputs);
const dmInc = engine.calculateIncomeStatement(debtModifiedInputs, dmFa);
const dmCf = engine.calculateCashFlowStatement(debtModifiedInputs, dmInc, dmFa);
const dmEq = engine.calculateEquityStatement(debtModifiedInputs, dmInc);
const dmBs = engine.calculateBalanceSheet(debtModifiedInputs, dmEq, dmFa, dmCf, undefined, dmInc);
const dmProj = engine.buildProjectedFormatData(debtModifiedInputs, dmInc, dmBs, dmCf, dmEq, dmFa);
const dmQuantNotes = engine.buildProjectedQuantNoteRows(dmProj);

assert.equal(dmBs.find(b => b.year === 2024)!.shortTermBorrowings, 456789, '2024 Short-Term Borrowings must match custom edit');
assert.equal(dmBs.find(b => b.year === 2024)!.longTermBorrowings, 2345678, '2024 Long-Term Borrowings must match custom edit');
assert.equal(dmProj.find(p => p.year === 2027)!.shortTermLoan, 321000, 'Projected 2027 shortTermLoan must match custom edit');
assert.equal(dmProj.find(p => p.year === 2027)!.longTermLoan, 1650000, 'Projected 2027 longTermLoan must match custom edit');
const idx2027Proj = dmProj.findIndex(p => p.year === 2027);
assert.equal(dmQuantNotes.find(r => r.noteNo === '7.00')!.values[idx2027Proj], 1650000, 'Quant Note 7.00 must reflect edited LT loan');
assert.equal(
  dmQuantNotes.find(r => r.configId === 'bs_shortTermBorrowings' && !r.isSectionHeader)!.values[idx2027Proj],
  321000,
  'Quant Note 10.00 Bank Loan row must reflect edited ST loan'
);
dmBs.forEach(b => assert(Math.abs(b.balanceCheck) < 0.05, `ST+LT-edited BS unbalanced in ${b.year}: ${b.balanceCheck}`));
dmCf.forEach(c => assert.equal(c.reconciliationStatus, 'RECONCILED', `ST+LT-edited CF unreconciled in ${c.year}`));

// (C) Exclusive statementConfig Calculation (deselecting cf_finance_paid & inc_interestIncome keeps BS balanced; deselecting bs_otherCurrentAssets zeros it cleanly)
const exclusiveIncCfConfig: StatementConfigItem[] = engine.DEFAULT_STATEMENT_CONFIG_ITEMS.map(item =>
  item.id === 'cf_finance_paid' || item.id === 'inc_interestIncome'
    ? { ...item, isVisible: false }
    : item
);
const exclIncCfRes = runPipeline(exclusiveIncCfConfig);
exclIncCfRes.inc.forEach(i => {
  assert.equal(i.interestIncome, 0, `Deselected inc_interestIncome must be 0 in ${i.year}`);
});
exclIncCfRes.bs.forEach(b => {
  assert(Math.abs(b.balanceCheck) < 0.05, `Exclusive Inc/CF-config BS unbalanced in ${b.year}: ${b.balanceCheck}`);
});
exclIncCfRes.cf.forEach(c => {
  assert.equal(c.financeCostPaid, 0, `Deselected cf_finance_paid must be 0 in ${c.year}`);
  if (c.year === 2023) {
    assert.equal(c.reconciliationStatus, 'VARIANCE_IDENTIFIED', 'Year 1 CF must transparently flag variance when CF line hidden');
  } else {
    assert.equal(c.reconciliationStatus, 'RECONCILED', `Exclusive Inc/CF-config CF unreconciled in ${c.year}`);
  }
});

const exclusiveBsConfig: StatementConfigItem[] = engine.DEFAULT_STATEMENT_CONFIG_ITEMS.map(item =>
  item.id === 'bs_otherCurrentAssets' ? { ...item, isVisible: false } : item
);
const exclBsRes = runPipeline(exclusiveBsConfig);
exclBsRes.bs.forEach((b, idx) => {
  const baseOtherCA = baseline.bs[idx].otherCurrentAssets;
  assert.equal(b.otherCurrentAssets, 0, `Deselected bs_otherCurrentAssets must be 0 in ${b.year}`);
  assert(
    Math.abs(b.balanceCheck + baseOtherCA) < 0.05,
    `Deselecting bs_otherCurrentAssets in ${b.year} must reflect exact hidden difference (-${baseOtherCA}), got ${b.balanceCheck}`
  );
});
exclBsRes.cf.forEach(c => {
  if (c.year === 2023) {
    assert.equal(c.reconciliationStatus, 'VARIANCE_IDENTIFIED', 'Year 1 CF must transparently flag variance when BS working capital line hidden');
  } else {
    assert.equal(c.reconciliationStatus, 'RECONCILED', `Exclusive BS-config CF unreconciled in ${c.year}`);
  }
});

// (D) Chartered Accountancy Audit & Instant Rectification Engine
const auditMod = await import('./src/services/charteredAuditEngine');
const cleanAbabilAudit = auditMod.runComprehensiveCharteredAudit({
  inputs: ababilOrdered,
  incomeStatements: abInc,
  balanceSheets: abBs,
  cashFlowStatements: abCf,
  equityStatements: abEq,
  fixedAssetsSchedule: abFa,
  debtServiceCoverage: abDscr,
  assumptions: ababilResult.assumptions,
  companyProfile: engine.INITIAL_COMPANY_PROFILE,
  cellOverrides: {},
  statementConfig: engine.DEFAULT_STATEMENT_CONFIG_ITEMS,
  parsedSchedule: ababilSched,
  isScheduleImplemented: true,
});
assert.equal(
  cleanAbabilAudit.criticalCount,
  0,
  `Integrated ABABIL baseline must have 0 critical audit errors (found ${cleanAbabilAudit.criticalCount}: ${cleanAbabilAudit.findings.filter(f => f.severity === 'critical').map(f => f.title).join('; ')})`
);

// Inject synthetic errors: BS override variance + CF cash mismatch + Year 1 new loan proceeds
const brokenInputs = ababilOrdered.map((y, idx) => {
  if (idx === 3 && y.isProjected) {
    return {
      ...y,
      financingProceedsBorrowings: 500000,
      cashFlowFromFinancing: (y.cashFlowFromFinancing || 0) + 500000,
    };
  }
  return y;
});
const brokenOverrides: StatementCellOverrides = {
  bs_tradeReceivables: { 2024: abBs.find(b => b.year === 2024)!.tradeReceivables + 95000 },
  bs_cashAndEquivalents: { 2025: abBs.find(b => b.year === 2025)!.cashAndEquivalents + 42000 },
};
const brokenOpts = { statementConfig: engine.DEFAULT_STATEMENT_CONFIG_ITEMS, cellOverrides: brokenOverrides };
const brFa = engine.calculateScheduleOfFixedAssets(brokenInputs, brokenOpts);
const brInc = engine.calculateIncomeStatement(brokenInputs, brFa, brokenOpts);
const brCf = engine.calculateCashFlowStatement(brokenInputs, brInc, brFa, brokenOpts);
const brEq = engine.calculateEquityStatement(brokenInputs, brInc, brokenOpts);
const brBs = engine.calculateBalanceSheet(brokenInputs, brEq, brFa, brCf, brokenOpts, brInc);
const brDscr = engine.calculateDebtServiceCoverage(brokenInputs, brInc, brBs, brCf, ababilSched, brokenOpts);

const preRectAudit = auditMod.runComprehensiveCharteredAudit({
  inputs: brokenInputs,
  incomeStatements: brInc,
  balanceSheets: brBs,
  cashFlowStatements: brCf,
  equityStatements: brEq,
  fixedAssetsSchedule: brFa,
  debtServiceCoverage: brDscr,
  assumptions: ababilResult.assumptions,
  companyProfile: engine.INITIAL_COMPANY_PROFILE,
  cellOverrides: brokenOverrides,
  statementConfig: engine.DEFAULT_STATEMENT_CONFIG_ITEMS,
  parsedSchedule: ababilSched,
  isScheduleImplemented: true,
});
assert(preRectAudit.criticalCount > 0, 'Audit must detect injected BS/CF critical discrepancies');
assert(preRectAudit.autoRectifiableCount > 0, 'Audit must classify injected discrepancies as auto-rectifiable');

const rectified = auditMod.applyInstantAuditRectifications({
  allYearsInputs: brokenInputs,
  historicalYearsStore: Object.fromEntries(brokenInputs.filter(y => !y.isProjected).map(y => [y.year, y])),
  customProjectedOverrides: Object.fromEntries(brokenInputs.filter(y => y.isProjected).map(y => [y.year, y])),
  assumptions: ababilResult.assumptions,
  cellOverrides: brokenOverrides,
  statementConfig: engine.DEFAULT_STATEMENT_CONFIG_ITEMS,
  companyProfile: engine.INITIAL_COMPANY_PROFILE,
  parsedSchedule: ababilSched,
  isScheduleImplemented: true,
});
assert(rectified.appliedFindingIds.length > 0, 'Instant Rectification must apply fixes');
assert.equal(
  rectified.postAuditReport.criticalCount,
  0,
  `Post-rectification criticalCount must drop to 0 (got ${rectified.postAuditReport.criticalCount}: ${rectified.postAuditReport.findings.filter(f => f.severity === 'critical').map(f => `${f.id}:${f.title}`).join('; ')})`
);
assert.equal(
  rectified.postAuditReport.autoRectifiableCount,
  0,
  `Post-rectification autoRectifiableCount must drop to 0 (got ${rectified.postAuditReport.autoRectifiableCount}: ${rectified.postAuditReport.findings.filter(f => f.isAutoRectifiable).map(f => `${f.id}:${f.title}`).join('; ')})`
);

// (E) PDF-Parity Excel Styling, Print Setup, Real Styled .xlsx Binary Write & Word .docx Exports
const xlsxStyleMod = require('xlsx-js-style') as {
  write: (wb: XlsxType.WorkBook, opts: { type: 'buffer'; bookType: 'xlsx' }) => Buffer;
};

capturedWorkbooks.length = 0;
exportsMod.exportCompletePackageExcel(
  engine.INITIAL_COMPANY_PROFILE,
  abInc,
  abBs,
  abCf,
  abEq,
  abFa,
  abNotes,
  'parity_complete.xlsx',
  abDscr,
  undefined,
  ababilSched
);
exportsMod.exportPreviousFinancialsExcel(
  engine.INITIAL_COMPANY_PROFILE,
  baseline.prev.slice(0, 2),
  baseline.notes
);
exportsMod.exportProjectedFinancialsExcel(
  engine.INITIAL_COMPANY_PROFILE,
  abProj,
  undefined,
  undefined,
  ababilSched,
  abDscr
);
exportsMod.exportRepaymentScheduleExcel(engine.INITIAL_COMPANY_PROFILE, ababilSched, abDscr);

assert.equal(capturedWorkbooks.length, 4, 'Expected 4 parity Excel workbooks captured');
for (const { filename, wb } of capturedWorkbooks) {
  validateFormulaCells(wb);
  const binBuf = xlsxStyleMod.write(wb, { type: 'buffer', bookType: 'xlsx' });
  assert(Buffer.isBuffer(binBuf) && binBuf.length > 5000, `Styled .xlsx binary export failed/too small for ${filename}`);

  // Every worksheet must carry A4 pageSetup & margins matching PDF layout
  for (const sName of wb.SheetNames) {
    const ws = wb.Sheets[sName] as any;
    assert.equal(ws['!pageSetup']?.paperSize, 9, `${filename}!${sName} must target A4 paperSize (9)`);
    assert(ws['!margins'] && typeof ws['!margins'].top === 'number', `${filename}!${sName} must declare !margins`);
  }
}

// Verify Section 76 Cover sheet, Helvetica font & navy banner fill
const compWb = capturedWorkbooks[0].wb;
assert(compWb.SheetNames[0].includes('Cover'), `Section 76 workbook must start with Cover sheet, got ${compWb.SheetNames[0]}`);
const incWs = compWb.Sheets['1. Income Statement'] as any;
assert(
  String(incWs['A1']?.s?.fill?.fgColor?.rgb || '').toUpperCase().endsWith('0F172A'),
  `Section 76 banner row must have PDF navy fill #0F172A, got ${incWs['A1']?.s?.fill?.fgColor?.rgb}`
);
assert.equal(incWs['A1']?.s?.font?.name, 'Helvetica', 'Section 76 must use Helvetica font matching PDF');

// Verify Previous Format uses Times New Roman matching PDF
const prevWb = capturedWorkbooks[1].wb;
const prevBsWs = prevWb.Sheets['1. Financial Position'] as any;
assert.equal(prevBsWs['A1']?.s?.font?.name, 'Times New Roman', 'Previous Format must use Times New Roman matching PDF');

// Verify Projected Format Landscape A4 + Schedule-A multi-row Cost merge
const projWb = capturedWorkbooks[2].wb;
const projBsWs = projWb.Sheets['1. Financial Position'] as any;
assert.equal(projBsWs['!pageSetup']?.orientation, 'landscape', 'Projected Format BS must use Landscape orientation matching PDF');
assert.equal(projBsWs['A1']?.s?.font?.name, 'Helvetica', 'Projected Format must use Helvetica matching PDF');
const projSchedAWs = projWb.Sheets['5. Schedule-A PPE'] as any;
assert(
  (projSchedAWs['!merges'] || []).some((m: XlsxType.Range) => m.s.c === 2 && m.e.c === 5),
  'Projected Schedule A must include merged Cost group header spanning C..F'
);

// Verify all 4 Word (.docx) PDF-parity exports generate real binary .docx documents cleanly
exportsMod.clearCapturedDocxExportsForTest();
await exportsMod.exportCompletePackageWord(
  engine.INITIAL_COMPANY_PROFILE,
  abInc,
  abBs,
  abCf,
  abEq,
  abFa,
  abNotes,
  'parity_complete.docx',
  abDscr,
  undefined,
  ababilSched
);
await exportsMod.exportPreviousFinancialsWord(
  engine.INITIAL_COMPANY_PROFILE,
  baseline.prev.slice(0, 2),
  baseline.notes
);
await exportsMod.exportProjectedFinancialsWord(
  engine.INITIAL_COMPANY_PROFILE,
  abProj,
  undefined,
  undefined,
  ababilSched,
  abDscr
);
await exportsMod.exportRepaymentScheduleWord(engine.INITIAL_COMPANY_PROFILE, ababilSched, abDscr);

const capturedDocx = exportsMod.getCapturedDocxExportsForTest();
assert.equal(capturedDocx.length, 4, `Expected 4 captured Word .docx exports, got ${capturedDocx.length}`);
for (const docItem of capturedDocx) {
  assert(docItem.size > 5000, `Word .docx export ${docItem.filename} produced suspiciously small buffer (${docItem.size} bytes)`);
}

xlsxCjs.writeFile = origWriteFile;

// ============================================================================
// 10. SAVED USER DEFAULT CONFIG, HISTORICAL UNPAID LOANS & FINAL YEAR LT = NIL
// ============================================================================

// (A) Save User-Selected Financial Statement Items as Default (mock globalThis.localStorage)
const memStorage: Record<string, string> = {};
(globalThis as unknown as { localStorage: Storage }).localStorage = {
  getItem: (k: string) => (k in memStorage ? memStorage[k] : null),
  setItem: (k: string, v: string) => {
    memStorage[k] = String(v);
  },
  removeItem: (k: string) => {
    delete memStorage[k];
  },
  clear: () => {
    Object.keys(memStorage).forEach(k => delete memStorage[k]);
  },
  key: (i: number) => Object.keys(memStorage)[i] ?? null,
  get length() {
    return Object.keys(memStorage).length;
  },
} as Storage;

const customSavedCfg: StatementConfigItem[] = engine.DEFAULT_STATEMENT_CONFIG_ITEMS.map(item =>
  item.id === 'bs_otherNonCurrentAssets'
    ? { ...item, isVisible: false }
    : item.id === 'inc_revenue'
    ? { ...item, label: 'Turnover & Net Sales' }
    : item
).concat([
  {
    id: 'custom_test_default_line',
    statement: 'income',
    section: 'Operating Expenses',
    label: 'Saved Custom Audit Fee',
    isMandatory: false,
    isVisible: true,
    order: 99,
    isCustom: true,
    sign: 'negative',
    customValues: { 2025: 12345 },
  },
]);

const savedRes = engine.saveStatementConfigAsDefault(customSavedCfg);
assert.equal(savedRes.ok, true, 'saveStatementConfigAsDefault must succeed when storage is available');

const reloadedCfg = engine.loadDefaultStatementConfig();
assert.equal(reloadedCfg.hasSavedDefault, true, 'loadDefaultStatementConfig must report hasSavedDefault=true');
assert.equal(
  reloadedCfg.items.find(i => i.id === 'bs_otherNonCurrentAssets')!.isVisible,
  false,
  'Deselected item must persist across reloads'
);
assert.equal(
  reloadedCfg.items.find(i => i.id === 'inc_revenue')!.label,
  'Turnover & Net Sales',
  'Customized label must persist across reloads'
);
const reloadedCustom = reloadedCfg.items.find(i => i.id === 'custom_test_default_line');
assert(reloadedCustom, 'Saved custom element template must persist across reloads');
assert.deepEqual(reloadedCustom!.customValues, {}, 'Saved default template must clear case-specific customValues');

const factoryRestored = engine.clearSavedDefaultStatementConfig();
assert.equal(
  factoryRestored.find(i => i.id === 'bs_otherNonCurrentAssets')!.isVisible,
  true,
  'clearSavedDefaultStatementConfig must restore factory defaults'
);
assert.equal(engine.loadDefaultStatementConfig().hasSavedDefault, false, 'Cleared default must report hasSavedDefault=false');

// (B) Recognition of Previous Unpaid Loans under Long-Term Borrowings & Finance Cost (ABABIL schedule)
const abHist2023 = ababilOrdered.find(y => y.year === 2023)!;
const abHist2024 = ababilOrdered.find(y => y.year === 2024)!;
const abHist2025 = ababilOrdered.find(y => y.year === 2025)!;
[abHist2023, abHist2024, abHist2025].forEach(h => {
  assert.equal(h.repaymentOfBorrowings, 0, `Unpaid historical year ${h.year} must carry 0 principal repayment`);
  assert.equal(h.financingProceedsBorrowings, 0, `Unpaid historical year ${h.year} must carry 0 new borrowing proceeds`);
  assert.equal(h.shortTermBorrowings, 0, `Unpaid historical year ${h.year} must recognize loan under Long-Term Borrowings (ST=0)`);
  assert(h.longTermBorrowings > 0, `Unpaid historical year ${h.year} must carry positive Long-Term Borrowings`);
  assert(h.financeCost > 0, `Unpaid historical year ${h.year} must recognize Finance Cost on unpaid debt`);
});
assert(
  Math.abs(abHist2025.longTermBorrowings - ababilSched.metadata.investmentMade) < 0.05,
  `2025 Long-Term Borrowings (${abHist2025.longTermBorrowings}) must match unpaid loan balance (${ababilSched.metadata.investmentMade})`
);

// (C) Avoidance of 2027 Disbursement Double-Counting & Complete LT Settlement to NIL (0) by Final Projected Year
const ab2027 = ababilOrdered.find(y => y.year === 2027)!;
const abFinal = ababilOrdered[ababilOrdered.length - 1];
assert(
  Math.abs(ab2027.longTermBorrowings - 2113842.29) < 1,
  `Projected 2027 LT Borrowings must equal schedule closing balance 2113842.29 (no double counting), got ${ab2027.longTermBorrowings}`
);
assert.equal(
  abFinal.longTermBorrowings,
  0,
  `Final projected year (${abFinal.year}) Long-Term Borrowings must be NIL (0), got ${abFinal.longTermBorrowings}`
);
assert.equal(
  abBs[abBs.length - 1].longTermBorrowings,
  0,
  `Final projected Balance Sheet (${abFinal.year}) Long-Term Borrowings must be NIL (0)`
);
assert(
  cleanAbabilAudit.passedChecks.some(p => p.id === 'pass-final-lt-nil'),
  'Audit report must verify pass-final-lt-nil passed check when final projected LT debt is 0'
);

// (D) Shorter 7-Year Horizon Ending 2032 Also Settles LT Borrowings to NIL (0)
const shortHorizonRes = engine.implementRepaymentScheduleIntoModel({
  schedule: ababilSched,
  historicalYearsStore: ababilHistStore,
  customProjectedOverrides: {},
  assumptions: {
    ...engine.DEFAULT_PROJECTION_ASSUMPTIONS,
    selectedProjectedYears: [2026, 2027, 2028, 2029, 2030, 2031, 2032],
    projectedYearsCount: 7,
  },
  autoIncludeScheduleYears: true,
  alignHistoricalScale: true,
});
const shortHorizonFinal =
  shortHorizonRes.customProjectedOverrides[
    shortHorizonRes.assumptions.selectedProjectedYears![
      shortHorizonRes.assumptions.selectedProjectedYears!.length - 1
    ]
  ];
assert.equal(
  shortHorizonFinal.longTermBorrowings,
  0,
  `Schedule terminal year (${shortHorizonFinal.year}) Long-Term Borrowings must be NIL (0), got ${shortHorizonFinal.longTermBorrowings}`
);

// (E) Automatic Projection Horizon (Without Uploaded Schedule) Also Settles Long-Term Borrowings to NIL (0) by Final Projected Year
const baselineFinalBs = baseline.bs[baseline.bs.length - 1];
assert.equal(
  baselineFinalBs.longTermBorrowings,
  0,
  `Baseline automatic projection final year (${baselineFinalBs.year}) Long-Term Borrowings must be NIL (0), got ${baselineFinalBs.longTermBorrowings}`
);

// (F) Audit Detection & Instant Rectification of Unsettled Final-Year Long-Term Borrowings
const unsettledInputs = ababilOrdered.map((y, idx) =>
  idx === ababilOrdered.length - 1
    ? { ...y, longTermBorrowings: 125000, hasCustomLongTermBorrowings: true }
    : y
);
const usFa = engine.calculateScheduleOfFixedAssets(unsettledInputs);
const usInc = engine.calculateIncomeStatement(unsettledInputs, usFa);
const usCf = engine.calculateCashFlowStatement(unsettledInputs, usInc, usFa);
const usEq = engine.calculateEquityStatement(unsettledInputs, usInc);
const usBs = engine.calculateBalanceSheet(unsettledInputs, usEq, usFa, usCf, undefined, usInc);
const usDscr = engine.calculateDebtServiceCoverage(unsettledInputs, usInc, usBs, usCf, ababilSched);
const unsettledAudit = auditMod.runComprehensiveCharteredAudit({
  inputs: unsettledInputs,
  incomeStatements: usInc,
  balanceSheets: usBs,
  cashFlowStatements: usCf,
  equityStatements: usEq,
  fixedAssetsSchedule: usFa,
  debtServiceCoverage: usDscr,
  assumptions: ababilResult.assumptions,
  companyProfile: engine.INITIAL_COMPANY_PROFILE,
  cellOverrides: {},
  statementConfig: engine.DEFAULT_STATEMENT_CONFIG_ITEMS,
  parsedSchedule: ababilSched,
  isScheduleImplemented: true,
});
assert(
  unsettledAudit.findings.some(f => f.rectificationAction === 'SETTLE_FINAL_YEAR_LT_BORROWINGS'),
  'Audit engine must detect unsettled LT borrowings in final projected year under implemented schedule'
);
const settledRect = auditMod.applyInstantAuditRectifications({
  allYearsInputs: unsettledInputs,
  historicalYearsStore: Object.fromEntries(unsettledInputs.filter(y => !y.isProjected).map(y => [y.year, y])),
  customProjectedOverrides: Object.fromEntries(unsettledInputs.filter(y => y.isProjected).map(y => [y.year, y])),
  assumptions: ababilResult.assumptions,
  cellOverrides: {},
  statementConfig: engine.DEFAULT_STATEMENT_CONFIG_ITEMS,
  companyProfile: engine.INITIAL_COMPANY_PROFILE,
  parsedSchedule: ababilSched,
  isScheduleImplemented: true,
});
assert.equal(settledRect.postAuditReport.criticalCount, 0, 'Instant rectification must resolve unsettled final-year LT borrowings cleanly');
assert.equal(
  settledRect.customProjectedOverrides[abFinal.year].longTermBorrowings,
  0,
  'Instant rectification must settle final projected year LT borrowings to 0'
);

console.log('ALL VERIFICATION CHECKS PASSED! Workbooks:', capturedWorkbooks.map(c => c.filename), 'Docx:', capturedDocx.map(d => `${d.filename}(${d.size}B)`));




