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
import type { StatementConfigItem, StatementLabelOverrides } from './src/types/financial';

function runPipeline(config: StatementConfigItem[], labelOverrides: StatementLabelOverrides = {}) {
  const baseYear = engine.INITIAL_HISTORICAL_INPUTS[engine.INITIAL_HISTORICAL_INPUTS.length - 1];
  const projectedInputs = engine.generateProjectedYears(baseYear, engine.DEFAULT_PROJECTION_ASSUMPTIONS);
  const allInputs = [...engine.INITIAL_HISTORICAL_INPUTS, ...projectedInputs].sort((a, b) => a.year - b.year);
  const opts = { statementConfig: config };
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
    fa
  );
  const dscr = engine.calculateDebtServiceCoverage(allInputs, bs, inc, cf);
  const prev = engine.buildPreviousFormatData(allInputs, inc, bs, cf, eq, fa, opts);
  const proj = engine.buildProjectedFormatData(allInputs, inc, bs, cf, eq, fa, opts);
  return { allInputs, fa, inc, cf, eq, bs, notes, dscr, prev, proj, labelOverrides };
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

// 2. Customization test: deselect items across all 5 statements + rename labels + add custom item
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
customConfig.push({
  id: 'custom_income_royalty',
  statement: 'income',
  label: 'Royalty Grant Income',
  section: 'Non-Operating & Tax',
  isVisible: true,
  isCustom: true,
  customSign: 'add',
  customValuesByYear: Object.fromEntries(baseline.allInputs.map(y => [y.year, 50000])),
});

const labelOverrides: StatementLabelOverrides = {
  income: { revenue: 'Turnover & Gross Sales' },
  bs: { cashAndEquivalents: 'Cash at Bank & Hand' },
};

const customized = runPipeline(customConfig, labelOverrides);

const base2025 = baseline.inc.find(d => d.year === 2025)!;
const cust2025 = customized.inc.find(d => d.year === 2025)!;
assert.equal(cust2025.costOfSales, 0, 'Deselected COGS should be 0');
assert.equal(cust2025.financeCost, 0, 'Deselected Finance Cost should be 0');
assert(cust2025.grossProfit > base2025.grossProfit, 'Gross Profit should increase when COGS deselected');
assert(cust2025.netProfit > base2025.netProfit, 'Net Profit should adjust dynamically');

const baseBs2025 = baseline.bs.find(d => d.year === 2025)!;
const custBs2025 = customized.bs.find(d => d.year === 2025)!;
assert.equal(custBs2025.tradeReceivables, 0, 'Deselected Trade Receivables should be 0');
assert(
  Math.abs(custBs2025.balanceCheck + baseBs2025.tradeReceivables) < 0.05,
  `Expected BS balanceCheck to reflect deselected Trade Receivables (-${baseBs2025.tradeReceivables}), got ${custBs2025.balanceCheck}`
);

// Also verify BS remains balanced (balanceCheck === 0) when toggling Income Statement + FA category without dropping BS lines
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
assert(!custFa2025.categories.some(c => c.name === 'Office Equipment & IT'), 'Deselected FA category must be excluded');
assert(custFa2025.totalNetBookValue < baseline.fa.find(d => d.year === 2025)!.totalNetBookValue, 'FA Total NBV must decrease');

// 3. Export workbook synchronization & formula validation
const syncOpts = { statementConfig: customConfig, labelOverrides };
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
    assert(!bsCsv.includes('Trade Receivables'), 'Complete Package BS included deselected Trade Receivables');
    assert(bsCsv.includes('Cash at Bank & Hand'), 'Complete Package BS missing custom BS label');
    assert(!cfCsv.includes('Add: Depreciation'), 'Complete Package CF included deselected Depreciation');
    assert(!eqCsv.includes('Dividends Paid'), 'Complete Package Equity included deselected Dividends Paid');
    assert(!faCsv.includes('Office Equipment & IT'), 'Complete Package FA included deselected FA category');
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
    assert(plCsv.includes('Turnover & Gross Sales'), 'Projected Excel missing custom income label');
    assert(bsCsv.includes('Cash at Bank & Hand'), 'Projected Excel missing custom BS label');
    assert(!plCsv.includes('Cost of Goods Sold'), 'Projected Excel included deselected COGS');
    assert(!plCsv.includes('Less: Financial Expenses'), 'Projected Excel included deselected Finance Cost');
    assert(!cfCsv.includes('\nDepreciation,'), 'Projected Excel included deselected CF Depreciation');
  }
}

xlsxCjs.writeFile = origWriteFile;
console.log('ALL VERIFICATION CHECKS PASSED! Workbooks validated:', capturedWorkbooks.map(c => c.filename));
