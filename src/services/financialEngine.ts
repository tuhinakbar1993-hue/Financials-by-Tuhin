import {
  CompanyProfile,
  YearDataInput,
  IncomeStatementYear,
  BalanceSheetYear,
  CashFlowStatementYear,
  EquityStatementYear,
  ScheduleOfFixedAssetsYear,
  FixedAssetCategory,
  FinancialNoteItem,
  ProjectionAssumptions,
  StatementConfigItem,
  EntityType,
  PrimaryDriverKey,
  DriverAutoFillConfig,
  AutoFillConfigs,
  StatementCellOverrides,
  StatementLabelOverrides,
  FinancialYearEndOption,
} from '../types/financial';

export interface StatementCalculationOptions {
  statementConfig?: StatementConfigItem[];
  cellOverrides?: StatementCellOverrides;
}

export const FIXED_ASSET_CATEGORY_CONFIG_IDS = [
  'fa_land_buildings',
  'fa_plant_machinery',
  'fa_office_equipment',
  'fa_motor_vehicles',
  'fa_computer_it',
] as const;

export const STATEMENT_CONFIG_TO_OVERRIDE_KEY: Record<string, { stmt: string; field: string }> = {
  // Income Statement
  inc_revenue: { stmt: 'income', field: 'revenue' },
  inc_costOfSales: { stmt: 'income', field: 'costOfSales' },
  inc_operatingExpenses: { stmt: 'income', field: 'operatingExpenses' },
  inc_depreciationExpense: { stmt: 'income', field: 'depreciationExpense' },
  inc_interestIncome: { stmt: 'income', field: 'interestIncome' },
  inc_financeCost: { stmt: 'income', field: 'financeCost' },
  inc_taxExpense: { stmt: 'income', field: 'taxExpense' },
  // Balance Sheet
  bs_cashAndEquivalents: { stmt: 'bs', field: 'cashAndEquivalents' },
  bs_tradeReceivables: { stmt: 'bs', field: 'tradeReceivables' },
  bs_otherReceivables: { stmt: 'bs', field: 'otherReceivables' },
  bs_closingInventory: { stmt: 'bs', field: 'closingInventory' },
  bs_otherCurrentAssets: { stmt: 'bs', field: 'otherCurrentAssets' },
  bs_propertyPlantEquipment: { stmt: 'bs', field: 'propertyPlantEquipment' },
  bs_intangibleAssets: { stmt: 'bs', field: 'intangibleAssets' },
  bs_otherNonCurrentAssets: { stmt: 'bs', field: 'otherNonCurrentAssets' },
  bs_tradePayables: { stmt: 'bs', field: 'tradePayables' },
  bs_otherPayables: { stmt: 'bs', field: 'otherPayables' },
  bs_shortTermBorrowings: { stmt: 'bs', field: 'shortTermBorrowings' },
  bs_currentTaxPayable: { stmt: 'bs', field: 'currentTaxPayable' },
  bs_longTermBorrowings: { stmt: 'bs', field: 'longTermBorrowings' },
  bs_otherNonCurrentLiabilities: { stmt: 'bs', field: 'otherNonCurrentLiabilities' },
  bs_shareCapital: { stmt: 'bs', field: 'shareCapital' },
  bs_retainedEarnings: { stmt: 'bs', field: 'retainedEarnings' },
  bs_reserves: { stmt: 'bs', field: 'reserves' },
  // Cash Flow
  cf_pbt: { stmt: 'cf', field: 'profitBeforeTax' },
  cf_dep: { stmt: 'cf', field: 'adjDepreciation' },
  cf_amort: { stmt: 'cf', field: 'adjAmortization' },
  cf_finance: { stmt: 'cf', field: 'adjFinanceCost' },
  cf_interest_inc: { stmt: 'cf', field: 'adjInterestIncome' },
  cf_disposal_gain_loss: { stmt: 'cf', field: 'adjLossGainDisposal' },
  cf_wc_inventory: { stmt: 'cf', field: 'changeInInventory' },
  cf_wc_receivables: { stmt: 'cf', field: 'changeInReceivables' },
  cf_wc_other_rec: { stmt: 'cf', field: 'changeInOtherReceivables' },
  cf_wc_other_ca: { stmt: 'cf', field: 'changeInOtherCA' },
  cf_wc_payables: { stmt: 'cf', field: 'changeInPayables' },
  cf_wc_other_pay: { stmt: 'cf', field: 'changeInOtherPayables' },
  cf_tax_paid: { stmt: 'cf', field: 'incomeTaxPaid' },
  cf_finance_paid: { stmt: 'cf', field: 'financeCostPaid' },
  cf_pur_ppe: { stmt: 'cf', field: 'purchasePPE' },
  cf_sale_ppe: { stmt: 'cf', field: 'salePPE' },
  cf_pur_intangible: { stmt: 'cf', field: 'purchaseIntangibles' },
  cf_interest_rec: { stmt: 'cf', field: 'interestReceived' },
  cf_dividends_rec: { stmt: 'cf', field: 'dividendsReceived' },
  cf_direct_financing: { stmt: 'cf', field: 'directlyEnteredFinancingCashFlow' },
  cf_shares_issue: { stmt: 'cf', field: 'financingProceedsShares' },
  cf_borrow_proceeds: { stmt: 'cf', field: 'financingProceedsBorrowings' },
  cf_borrow_repay: { stmt: 'cf', field: 'repaymentOfBorrowings' },
  cf_dividends_paid: { stmt: 'cf', field: 'financingDividendsPaid' },
  // Equity
  eq_openingShareCapital: { stmt: 'eq', field: 'openingShareCapital' },
  eq_shareCapitalIssued: { stmt: 'eq', field: 'shareCapitalIssued' },
  eq_openingRetainedEarnings: { stmt: 'eq', field: 'openingRetainedEarnings' },
  eq_netProfitForYear: { stmt: 'eq', field: 'netProfitForYear' },
  eq_dividendsPaid: { stmt: 'eq', field: 'dividendsPaid' },
  eq_openingReserves: { stmt: 'eq', field: 'openingReserves' },
  eq_otherComprehensiveIncome: { stmt: 'eq', field: 'otherComprehensiveIncome' },
};

export function buildConfigVisibilityMap(statementConfig?: StatementConfigItem[]): Map<string, StatementConfigItem> {
  const map = new Map<string, StatementConfigItem>();
  (statementConfig || []).forEach(item => map.set(item.id, item));
  return map;
}

export function isConfigElementVisible(
  configMap: Map<string, StatementConfigItem>,
  id: string
): boolean {
  const item = configMap.get(id);
  return item ? item.isVisible : true;
}

export function getCellOverride(
  cellOverrides: StatementCellOverrides | undefined,
  stmt: string,
  field: string,
  year: number
): number | undefined {
  const key = `${stmt}_${field}`;
  const val = cellOverrides?.[key]?.[year];
  return val !== undefined && Number.isFinite(val) ? val : undefined;
}

export function resolveEffectiveLabel(
  stmt: string,
  field: string,
  defaultLabel: string,
  labelOverrides?: StatementLabelOverrides,
  statementConfig?: StatementConfigItem[]
): string {
  const key = `${stmt}_${field}`;
  if (labelOverrides && labelOverrides[key] !== undefined) {
    return labelOverrides[key];
  }
  if (statementConfig) {
    let matchedId = Object.entries(STATEMENT_CONFIG_TO_OVERRIDE_KEY).find(
      ([, v]) => v.stmt === stmt && v.field === field
    )?.[0];
    if (!matchedId && stmt === 'fa') {
      const idxNum = Number(field);
      if (Number.isInteger(idxNum)) {
        matchedId = FIXED_ASSET_CATEGORY_CONFIG_IDS[idxNum];
      }
    }
    if (matchedId) {
      const cfgItem = statementConfig.find(c => c.id === matchedId);
      const defItem = DEFAULT_STATEMENT_CONFIG_ITEMS.find(c => c.id === matchedId);
      if (cfgItem && defItem && cfgItem.label !== defItem.label) {
        return cfgItem.label;
      }
    }
  }
  return defaultLabel;
}

export function getCustomItemYearValue(item: StatementConfigItem, year: number): number {
  const raw = item.customValues?.[year] ?? 0;
  if (!Number.isFinite(raw)) return 0;
  return item.sign === 'negative' && raw > 0 ? -raw : raw;
}

export function getVisibleCustomItemsSum(
  statementConfig: StatementConfigItem[] | undefined,
  statement: StatementConfigItem['statement'],
  year: number,
  section?: string
): number {
  if (!statementConfig) return 0;
  return statementConfig
    .filter(
      item =>
        item.isCustom &&
        item.isVisible &&
        item.statement === statement &&
        (!section || item.section === section)
    )
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
    .reduce((sum, item) => sum + getCustomItemYearValue(item, year), 0);
}

export function isJuneYearEnd(fyEnd?: string): boolean {
  return /june|jun|\b06\b/i.test(fyEnd || '');
}

export function normalizeFinancialYearEnd(fyEnd?: string): FinancialYearEndOption {
  return isJuneYearEnd(fyEnd) ? 'June 30' : 'December 31';
}

export function formatFyEndDateCode(year: number, fyEnd?: string): string {
  return `${isJuneYearEnd(fyEnd) ? '30.06' : '31.12'}.${year}`;
}

export function formatFyStartDateCode(year: number, fyEnd?: string): string {
  return isJuneYearEnd(fyEnd) ? `01.07.${year - 1}` : `01.01.${year}`;
}

export function formatFyAsAtUpper(year: number, fyEnd?: string): string {
  return `AS AT ${isJuneYearEnd(fyEnd) ? '30 JUNE' : '31 DECEMBER'}, ${year}`;
}

export function formatFyYearEndedTitle(year: number, fyEnd?: string): string {
  return `For the year ended ${isJuneYearEnd(fyEnd) ? '30 June' : '31 December'}, ${year}`;
}

export function formatFyYearEndedNoComma(year: number, fyEnd?: string): string {
  return `For the year ended ${isJuneYearEnd(fyEnd) ? '30 June' : '31 December'} ${year}`;
}

export function formatFyProjectedAsAt(year: number, fyEnd?: string): string {
  return `As at ${isJuneYearEnd(fyEnd) ? '30th June' : '31st December'}, ${year}(Projected)`;
}

export function formatFyProjectedYearEnded(year: number, fyEnd?: string): string {
  return `For the year ended ${isJuneYearEnd(fyEnd) ? '30th June' : '31st December'}, ${year}(Projected)`;
}

export const INITIAL_COMPANY_PROFILE: CompanyProfile = {
  companyName: 'Apex Industrial Holdings Ltd.',
  tradingName: 'Apex Global Manufacturing',
  registrationNumber: 'C-784920/2012',
  taxId: 'TIN-9821-4402-99',
  registeredAddress: 'Plot 42, Avenue of Financial Architecture, High-Tech Industrial Zone, Dhaka 1230',
  reportingCurrency: '$',
  reportingUnit: 'Thousands',
  entityType: 'manufacturing',
  accountingFramework: 'IFRS',
  financialYearEnd: 'December 31',
  auditorName: 'Rahman Rahman Huq & Co. (Chartered Accountants)',
  preparedBy: 'Mohammad Towhidul Akbar Tuhin, Officer, Sammilito Islami Bank PLC (Transferee of Assets & Liabilities of Union Bank PLC) | BBA, Heriot Watt University, UK · MBA, University of Chittagong',
};

export const DEFAULT_PROJECTION_ASSUMPTIONS: ProjectionAssumptions = {
  historicalYearsCount: 3,
  projectedYearsCount: 15,
  selectedHistoricalYears: [2023, 2024, 2025],
  selectedProjectedYears: Array.from({ length: 15 }, (_, i) => 2026 + i),
  revenueGrowthRate: 8.5,
  grossMarginPct: 38.0,
  opexPctOfRevenue: 17.5,
  taxRatePct: 25.0,
  capexPctOfRevenue: 6.2,
  depreciationRatePct: 10.0,
  receivableDays: 45,
  inventoryDays: 60,
  payableDays: 42,
  dividendPayoutRatio: 30.0,
  projectedFinancingCashFlowGrowth: -4.0,
  loanInterestRatePct: 15.5,
};

export function createCleanCaseCompanyProfile(currentFyEnd?: string): CompanyProfile {
  return {
    companyName: '',
    tradingName: '',
    registrationNumber: '',
    taxId: '',
    registeredAddress: '',
    reportingCurrency: '$',
    reportingUnit: 'Thousands',
    entityType: 'manufacturing',
    accountingFramework: 'IFRS',
    financialYearEnd: normalizeFinancialYearEnd(currentFyEnd),
    auditorName: '',
    preparedBy: INITIAL_COMPANY_PROFILE.preparedBy,
  };
}

export function createZeroHistoricalInputs(
  years: number[] = DEFAULT_PROJECTION_ASSUMPTIONS.selectedHistoricalYears || [2023, 2024, 2025]
): YearDataInput[] {
  return years.map((year) => ({
    year,
    isProjected: false,
    salesRevenue: 0,
    netProfit: 0,
    financeCost: 0,
    cashFlowFromFinancing: 0,
    repaymentOfBorrowings: 0,
    fixedAssetsGross: 0,
    fixedAssetsAdditions: 0,
    fixedAssetsDisposals: 0,
    depreciationRate: DEFAULT_PROJECTION_ASSUMPTIONS.depreciationRatePct,
    costOfSales: 0,
    operatingExpenses: 0,
    taxRate: DEFAULT_PROJECTION_ASSUMPTIONS.taxRatePct,
    closingInventory: 0,
    tradeReceivables: 0,
    otherReceivables: 0,
    tradePayables: 0,
    otherPayables: 0,
    cashAndEquivalentsBS: 0,
    otherCurrentAssets: 0,
    intangibleAssets: 0,
    otherNonCurrentAssets: 0,
    shortTermBorrowings: 0,
    longTermBorrowings: 0,
    hasCustomLongTermBorrowings: false,
    shareCapital: 0,
    retainedEarningsOpening: 0,
    reservesOpening: 0,
    dividendsPaid: 0,
    financingProceedsShares: 0,
    financingProceedsBorrowings: 0,
    financingRepaymentBorrowings: 0,
    financingDividendsPaid: 0,
    isScheduleLocked: false,
  }));
}

// Default 3 Historical Years (2023, 2024, 2025) with authentic uneven YoY ratios (>= 1.2pp distinct across all items)
export const INITIAL_HISTORICAL_INPUTS: YearDataInput[] = [
  {
    year: 2023,
    isProjected: false,
    salesRevenue: 48500,
    netProfit: 4620,
    financeCost: 1150, // Mandatory year-wise separate direct entry
    cashFlowFromFinancing: -1850, // Mandatory year-wise separate direct entry (Net repayment/dividend)
    repaymentOfBorrowings: 2150, // Mandatory Year-Wise User Input: Repayment of Borrowings & Financing
    fixedAssetsGross: 32000,
    fixedAssetsAdditions: 3800,
    fixedAssetsDisposals: 450,
    depreciationRate: 10,
    costOfSales: 30500,
    operatingExpenses: 8400,
    taxRate: 25,
    closingInventory: 7800,
    tradeReceivables: 6100,
    otherReceivables: 950,
    tradePayables: 4200,
    otherPayables: 1100,
    cashAndEquivalentsBS: 4120,
    otherCurrentAssets: 680,
    intangibleAssets: 1500,
    otherNonCurrentAssets: 4542,
    shortTermBorrowings: 3200,
    longTermBorrowings: 9400,
    shareCapital: 12000,
    retainedEarningsOpening: 6500,
    dividendsPaid: 1200,
    financingProceedsShares: 0,
    financingProceedsBorrowings: 1500,
    financingRepaymentBorrowings: -2150,
    financingDividendsPaid: -1200,
  },
  {
    year: 2024,
    isProjected: false,
    salesRevenue: 54200, // +11.75%
    netProfit: 5050, // +9.31%
    financeCost: 1340, // +16.52% (Mandatory year-wise separate direct entry)
    cashFlowFromFinancing: -2060, // Mandatory year-wise separate direct entry
    repaymentOfBorrowings: 1780, // -17.21% (Mandatory Year-Wise User Input)
    fixedAssetsGross: 35930, // +12.28%
    fixedAssetsAdditions: 4550, // +19.74%
    fixedAssetsDisposals: 620, // +37.78%
    depreciationRate: 9.5, // Dep ~3227 (+6.40%)
    costOfSales: 34400,
    operatingExpenses: 8770,
    taxRate: 25,
    closingInventory: 8360, // +7.18%
    tradeReceivables: 7180, // +17.70%
    otherReceivables: 1040, // +9.47%
    tradePayables: 4690, // +11.67%
    otherPayables: 1310, // +19.09%
    cashAndEquivalentsBS: 6098,
    otherCurrentAssets: 705, // +3.68%
    intangibleAssets: 1400,
    otherNonCurrentAssets: 3906,
    shortTermBorrowings: 2900,
    longTermBorrowings: 8820,
    shareCapital: 12000,
    retainedEarningsOpening: 9920,
    dividendsPaid: 1480, // +23.33%
    financingProceedsShares: 0,
    financingProceedsBorrowings: 1200,
    financingRepaymentBorrowings: -1780,
    financingDividendsPaid: -1480,
  },
  {
    year: 2025,
    isProjected: false,
    salesRevenue: 61800, // +14.02%
    netProfit: 6290, // +24.55%
    financeCost: 1290, // -3.73% (Mandatory year-wise separate direct entry)
    cashFlowFromFinancing: -2240, // Mandatory year-wise separate direct entry
    repaymentOfBorrowings: 1520, // -14.61% (Mandatory Year-Wise User Input)
    fixedAssetsGross: 40320, // +12.22%
    fixedAssetsAdditions: 4920, // +8.13%
    fixedAssetsDisposals: 530, // -14.52%
    depreciationRate: 10.4, // Dep ~3965 (+22.87%)
    costOfSales: 37975,
    operatingExpenses: 10475,
    taxRate: 25,
    closingInventory: 9840, // +17.70%
    tradeReceivables: 7750, // +7.94%
    otherReceivables: 1290, // +24.04%
    tradePayables: 5590, // +19.19%
    otherPayables: 1440, // +9.92%
    cashAndEquivalentsBS: 8703,
    otherCurrentAssets: 830, // +17.73%
    intangibleAssets: 1300,
    otherNonCurrentAssets: 3191,
    shortTermBorrowings: 2600,
    longTermBorrowings: 8120,
    shareCapital: 12000,
    retainedEarningsOpening: 13490,
    dividendsPaid: 1720, // +16.22%
    financingProceedsShares: 0,
    financingProceedsBorrowings: 1000,
    financingRepaymentBorrowings: -1520,
    financingDividendsPaid: -1720,
  },
];

/**
 * Financial Ratios Suite
 * Covers Profitability, Liquidity, Solvency, and Efficiency ratios
 */
export interface FinancialRatios {
  grossMarginPct: number;
  operatingMarginPct: number;
  netMarginPct: number;
  ebitdaMarginPct: number;
  returnOnEquityPct: number;
  returnOnCapitalEmployedPct: number;
  currentRatio: number;
  quickRatio: number;
  cashRatio: number;
  workingCapital: number;
  debtToEquity: number;
  debtToAssets: number;
  interestCoverageRatio: number;
  debtServiceCoverageRatio: number;
  receivableDays: number;
  inventoryDays: number;
  payableDays: number;
  assetTurnover: number;
}

export function calculateFinancialRatios(
  input: YearDataInput,
  income: IncomeStatementYear,
  balance: BalanceSheetYear,
  cashFlow: CashFlowStatementYear
): FinancialRatios {
  const hasRev = income.revenue > 0;
  const rev = Math.max(1, income.revenue);
  const gp = income.grossProfit;
  const ebit = income.operatingProfit;
  const ebitda = ebit + income.depreciationExpense;
  const np = income.netProfit;
  const hasEq = balance.totalEquity > 0;
  const eq = Math.max(1, balance.totalEquity);
  const hasAssets = balance.totalAssets > 0;
  const assets = Math.max(1, balance.totalAssets);
  const ca = balance.totalCurrentAssets;
  const hasCl = balance.totalCurrentLiabilities > 0;
  const cl = Math.max(1, balance.totalCurrentLiabilities);
  const inv = balance.closingInventory;
  const cash = balance.cashAndEquivalents;
  const totalDebt = balance.shortTermBorrowings + balance.longTermBorrowings;
  const hasFc = income.financeCost > 0;
  const fc = Math.max(1, income.financeCost);
  const hasCos = income.costOfSales > 0;
  const cos = Math.max(1, income.costOfSales);
  const debtService = income.financeCost + (cashFlow.repaymentOfBorrowings || 0);

  return {
    grossMarginPct: hasRev ? Number(((gp / rev) * 100).toFixed(1)) : 0,
    operatingMarginPct: hasRev ? Number(((ebit / rev) * 100).toFixed(1)) : 0,
    netMarginPct: hasRev ? Number(((np / rev) * 100).toFixed(1)) : 0,
    ebitdaMarginPct: hasRev ? Number(((ebitda / rev) * 100).toFixed(1)) : 0,
    returnOnEquityPct: hasEq ? Number(((np / eq) * 100).toFixed(1)) : 0,
    returnOnCapitalEmployedPct: assets - cl > 0 ? Number(((ebit / (assets - cl)) * 100).toFixed(1)) : 0,
    currentRatio: hasCl ? Number((ca / cl).toFixed(2)) : 0,
    quickRatio: hasCl ? Number(((ca - inv) / cl).toFixed(2)) : 0,
    cashRatio: hasCl ? Number((cash / cl).toFixed(2)) : 0,
    workingCapital: ca - balance.totalCurrentLiabilities,
    debtToEquity: hasEq ? Number((totalDebt / eq).toFixed(2)) : 0,
    debtToAssets: hasAssets ? Number((totalDebt / assets).toFixed(2)) : 0,
    interestCoverageRatio: hasFc ? Number((ebit / fc).toFixed(2)) : 0,
    debtServiceCoverageRatio: debtService > 0 ? Number((ebitda / debtService).toFixed(2)) : 0,
    receivableDays: hasRev ? Math.round((balance.tradeReceivables / rev) * 365) : 0,
    inventoryDays: hasCos ? Math.round((inv / cos) * 365) : 0,
    payableDays: hasCos ? Math.round((balance.tradePayables / cos) * 365) : 0,
    assetTurnover: hasAssets ? Number((income.revenue / assets).toFixed(2)) : 0,
  };
}

/**
 * Deterministic bounded cyclical wave in [-1, 1] seeded by year and component channel.
 * Combines 3 incommensurate frequencies for irregular, non-repetitive annual variations
 * while guaranteeing 100% repeatable output across UI renders, previews, and exports.
 */
function deterministicAnnualWave(year: number, channel: number): number {
  const a = Math.sin(year * 1.43 + channel * 2.17);
  const b = Math.cos(year * 0.83 + channel * 1.53 + 0.41);
  const c = Math.sin(year * 2.29 + channel * 3.11 + 0.92);
  return Math.min(1, Math.max(-1, a * 0.48 + b * 0.31 + c * 0.21));
}

function clampVal(val: number, minVal: number, maxVal: number): number {
  return Math.min(maxVal, Math.max(minVal, val));
}

function yoyPct(curr: number, prev: number): number {
  if (!Number.isFinite(prev) || Math.abs(prev) < 1e-6) return 0;
  return ((curr - prev) / Math.abs(prev)) * 100;
}

function varyBounded(
  baseVal: number,
  year: number,
  channel: number,
  amplitude: number,
  minVal: number,
  maxVal: number
): number {
  return clampVal(baseVal + deterministicAnnualWave(year, channel) * amplitude, minVal, maxVal);
}

/**
 * Adjusts a candidate YoY percentage change so it differs by at least `minSpreadPct`
 * percentage points from all `referenceGrowthPcts`, strictly within `[minAllowedPct, maxAllowedPct]`.
 */
function adjustGrowthPctForMinSpread(
  candidateGrowthPct: number,
  referenceGrowthPcts: number[],
  minSpreadPct: number,
  minAllowedPct: number,
  maxAllowedPct: number,
  preferredDir: number = 1
): number {
  const lo = Math.min(minAllowedPct, maxAllowedPct);
  const hi = Math.max(minAllowedPct, maxAllowedPct);
  const clampedCand = clampVal(candidateGrowthPct, lo, hi);
  const refs = referenceGrowthPcts.filter(r => Number.isFinite(r));
  if (refs.length === 0) return clampedCand;

  const minDiffOf = (g: number): number => Math.min(...refs.map(r => Math.abs(g - r)));
  if (minDiffOf(clampedCand) >= minSpreadPct) return clampedCand;

  const candidates: number[] = [clampedCand, lo, hi];
  for (const ref of refs) {
    candidates.push(ref + minSpreadPct + 0.08);
    candidates.push(ref - minSpreadPct - 0.08);
    candidates.push(ref + minSpreadPct + 0.55);
    candidates.push(ref - minSpreadPct - 0.55);
  }
  const span = Math.max(0.05, hi - lo);
  for (let i = 0; i <= 28; i++) {
    candidates.push(lo + (span * i) / 28);
  }

  const bounded = candidates.map(c => clampVal(c, lo, hi));
  const valid = bounded.filter(c => minDiffOf(c) >= minSpreadPct);

  if (valid.length > 0) {
    valid.sort((a, b) => {
      const da = Math.abs(a - clampedCand);
      const db = Math.abs(b - clampedCand);
      if (Math.abs(da - db) > 1e-4) return da - db;
      return preferredDir >= 0 ? b - a : a - b;
    });
    return valid[0];
  }

  bounded.sort((a, b) => {
    const diff = minDiffOf(b) - minDiffOf(a);
    if (Math.abs(diff) > 1e-4) return diff;
    return Math.abs(a - clampedCand) - Math.abs(b - clampedCand);
  });
  return bounded[0];
}

/**
 * Computes a bounded non-uniform integer/decimal value whose YoY growth rate
 * stays within `[minBoundVal, maxBoundVal]` while keeping at least `minSpreadPct` pp
 * separation from `referenceGrowthPcts`.
 */
function computeSpreadValue(
  targetVal: number,
  prevVal: number | null | undefined,
  minBoundVal: number,
  maxBoundVal: number,
  referenceGrowthPcts: number[],
  year: number,
  channel: number,
  minSpreadPct: number = 1.48,
  roundFn: (n: number) => number = Math.round
): number {
  const lo = Math.max(0, Math.min(minBoundVal, maxBoundVal));
  const hi = Math.max(lo, maxBoundVal);
  const clampedTarget = clampVal(targetVal, lo, hi);
  if (!prevVal || prevVal <= 0) return roundFn(clampedTarget);

  const candG = yoyPct(clampedTarget, prevVal);
  const minG = yoyPct(lo, prevVal);
  const maxG = yoyPct(hi, prevVal);
  const prefDir = deterministicAnnualWave(year, channel + 31) >= 0 ? 1 : -1;
  const chosenG = adjustGrowthPctForMinSpread(candG, referenceGrowthPcts, minSpreadPct, minG, maxG, prefDir);
  return roundFn(clampVal(prevVal * (1 + chosenG / 100), lo, hi));
}

/**
 * Computes irregular annual revenue growth % with natural expansions, slowdowns,
 * and occasional mild cyclical contractions while maintaining overall assumption trajectory.
 */
function computeProjectedRevenueGrowthPct(
  yearNumber: number,
  baseGrowthPct: number,
  stepIndex: number
): number {
  const wave1 = deterministicAnnualWave(yearNumber, 1);
  const wave2 = deterministicAnnualWave(yearNumber, 27);
  const maturityScale = stepIndex > 10 ? 0.87 : stepIndex > 6 ? 0.94 : 1.0;

  if (baseGrowthPct <= 0) {
    const amp = Math.max(1.8, Math.abs(baseGrowthPct) * 0.45);
    return clampVal(baseGrowthPct + wave1 * amp, -12, 6);
  }

  // Natural cyclical inflection (occasional isolated single-year mild slowdown + subsequent rebound)
  const cycleTrigger = Math.sin(stepIndex * 1.18 + 0.55);
  const prevCycleTrigger = stepIndex > 1 ? Math.sin((stepIndex - 1) * 1.18 + 0.55) : 0;
  const isIsolatedDip = stepIndex >= 4 && cycleTrigger < -0.78 && !(stepIndex - 1 >= 4 && prevCycleTrigger < -0.78);
  const prevWasDip = stepIndex >= 5 && prevCycleTrigger < -0.78;

  if (baseGrowthPct <= 11.5 && isIsolatedDip) {
    // Realistic mild single-year contraction (-1.4% to -3.6%) during periodic industrial slowdown
    return clampVal(-1.6 - Math.abs(wave2) * 1.9, -4.2, -1.25);
  }

  const amp = Math.max(1.85, Math.abs(baseGrowthPct) * 0.42);
  const reboundBoost = baseGrowthPct <= 11.5 && prevWasDip ? 2.6 : 0;
  return clampVal(baseGrowthPct * maturityScale + reboundBoost + wave1 * amp, -5.0, 28.0);
}

function computeProjectedDepRatePct(
  baseDepRatePct: number,
  yearNumber: number,
  grossGrowthPct: number | null,
  prevDepExpense: number | null,
  prevGross: number,
  currGross: number
): number {
  const compositeCenter = Number((baseDepRatePct * 0.935).toFixed(1));
  const minR = Math.max(2.0, Number((baseDepRatePct - 1.8).toFixed(1)));
  const maxR = Math.min(35.0, Number((baseDepRatePct + 0.5).toFixed(1)));
  const waveRate = Number(
    varyBounded(
      compositeCenter,
      yearNumber,
      24,
      Math.min(1.05, baseDepRatePct * 0.105),
      minR,
      maxR
    ).toFixed(1)
  );
  if (grossGrowthPct === null || !prevDepExpense || prevDepExpense <= 0) return waveRate;

  const avgGross = (prevGross + currGross) / 2;
  if (avgGross <= 0) return waveRate;

  let bestRate = waveRate;
  let bestDist = Infinity;
  let bestSpread = -1;

  for (let rTenths = Math.round(minR * 10); rTenths <= Math.round(maxR * 10); rTenths++) {
    const rCand = rTenths / 10;
    const depCand = Math.round(avgGross * (rCand / 100));
    const depG = yoyPct(depCand, prevDepExpense);
    const spread = Math.abs(depG - grossGrowthPct);
    const dist = Math.abs(rCand - waveRate);
    if (spread >= 1.50) {
      if (bestSpread < 1.50 || dist < bestDist) {
        bestRate = rCand;
        bestDist = dist;
        bestSpread = spread;
      }
    } else if (bestSpread < 1.50 && spread > bestSpread) {
      bestRate = rCand;
      bestSpread = spread;
    }
  }
  return Number(bestRate.toFixed(1));
}

/**
 * Solves `costOfSales` inside realistic ratio bounds `[loCos, hiCos]` such that
 * `operatingExpenses = totalOperatingCosts - costOfSales` stays inside `[loOpex, hiOpex]`
 * and YoY rates of CoS and OpEx stay >= `minSpreadPct` apart from `refs` and from each other.
 */
function allocateOperatingCostsWithSpread(
  totalOperatingCosts: number,
  targetCos: number,
  salesRevenue: number,
  targetCosPct: number,
  targetOpexPct: number,
  prevYear: YearDataInput | null,
  refsForBoth: number[],
  minSpreadPct: number = 1.48
): { costOfSales: number; operatingExpenses: number } {
  const minCosRatio = clampVal((targetCosPct - 3.8) / 100, 0.14, 0.90);
  const maxCosRatio = clampVal((targetCosPct + 3.8) / 100, minCosRatio + 0.02, 0.93);
  const minOpexRatio = clampVal((targetOpexPct - 2.8) / 100, 0.025, 0.50);
  const maxOpexRatio = clampVal((targetOpexPct + 2.8) / 100, minOpexRatio + 0.01, 0.56);

  const minCos = Math.round(salesRevenue * minCosRatio);
  const maxCos = Math.round(salesRevenue * maxCosRatio);
  const minOpex = Math.round(salesRevenue * minOpexRatio);
  const maxOpex = Math.round(salesRevenue * maxOpexRatio);

  let loCos = Math.max(0, minCos, totalOperatingCosts - maxOpex);
  let hiCos = Math.min(totalOperatingCosts, maxCos, totalOperatingCosts - minOpex);
  if (loCos > hiCos) {
    loCos = Math.max(0, Math.round(totalOperatingCosts * 0.55));
    hiCos = Math.min(totalOperatingCosts, Math.round(totalOperatingCosts * 0.90));
  }

  const baseCos = Math.round(clampVal(targetCos, loCos, hiCos));
  if (!prevYear || prevYear.costOfSales <= 0 || prevYear.operatingExpenses <= 0) {
    return {
      costOfSales: baseCos,
      operatingExpenses: Math.max(0, totalOperatingCosts - baseCos),
    };
  }

  const evalMinSpread = (cosCand: number): number => {
    const opexCand = totalOperatingCosts - cosCand;
    const gCos = yoyPct(cosCand, prevYear.costOfSales);
    const gOpex = yoyPct(opexCand, prevYear.operatingExpenses);
    let m = Math.abs(gCos - gOpex);
    for (const r of refsForBoth) {
      m = Math.min(m, Math.abs(gCos - r), Math.abs(gOpex - r));
    }
    return m;
  };

  if (evalMinSpread(baseCos) >= minSpreadPct) {
    return {
      costOfSales: baseCos,
      operatingExpenses: Math.max(0, totalOperatingCosts - baseCos),
    };
  }

  let bestCos = baseCos;
  let bestDist = Infinity;
  let bestSpread = evalMinSpread(baseCos);
  const steps = 64;
  const span = Math.max(1, hiCos - loCos);

  for (let i = 0; i <= steps; i++) {
    const cand = Math.round(loCos + (span * i) / steps);
    const sp = evalMinSpread(cand);
    const dist = Math.abs(cand - baseCos);
    if (sp >= minSpreadPct) {
      if (bestSpread < minSpreadPct || dist < bestDist) {
        bestCos = cand;
        bestDist = dist;
        bestSpread = sp;
      }
    } else if (bestSpread < minSpreadPct && sp > bestSpread) {
      bestCos = cand;
      bestSpread = sp;
    }
  }

  return {
    costOfSales: bestCos,
    operatingExpenses: Math.max(0, totalOperatingCosts - bestCos),
  };
}

export function computeInterestIncomeRate(year: number, isProjected: boolean): number {
  const amp = isProjected ? 0.00085 : 0.00045;
  return Number(clampVal(0.005 + deterministicAnnualWave(year, 10) * amp, 0.0034, 0.0068).toFixed(6));
}

export function computeDepDisposalsRatio(year: number, isProjected: boolean): number {
  const amp = isProjected ? 0.085 : 0.045;
  return Number(clampVal(0.65 + deterministicAnnualWave(year, 16) * amp, 0.48, 0.82).toFixed(4));
}

function splitTotalNonUniform(
  total: number,
  year: number,
  channel: number,
  baseWeights: number[],
  amplitudes: number[],
  prevParts?: number[],
  minSpreadPct: number = 1.35
): number[] {
  if (total === 0) return baseWeights.map(() => 0);

  const buildFromWeights = (weights: number[]): number[] => {
    const wSum = weights.reduce((acc, v) => acc + v, 0);
    const out: number[] = [];
    let running = 0;
    for (let i = 0; i < weights.length - 1; i++) {
      const val = Math.round(total * (weights[i] / wSum));
      out.push(val);
      running += val;
    }
    out.push(r2(total - running));
    return out;
  };

  const baseAdjusted = baseWeights.map((w, idx) =>
    Math.max(0.015, w + deterministicAnnualWave(year, channel + idx * 5) * (amplitudes[idx] ?? 0.03))
  );
  let bestParts = buildFromWeights(baseAdjusted);

  if (!prevParts || prevParts.length !== baseWeights.length || prevParts.some(p => p <= 0)) {
    return bestParts;
  }

  const prevTotal = prevParts.reduce((a, b) => a + b, 0);
  const pairSpread = (parts: number[]): number => {
    const rates = parts.map((v, idx) => yoyPct(v, prevParts[idx]));
    let minDiff = Infinity;
    for (let i = 0; i < rates.length; i++) {
      for (let j = i + 1; j < rates.length; j++) {
        minDiff = Math.min(minDiff, Math.abs(rates[i] - rates[j]));
      }
    }
    const shareShift = prevTotal > 0 && total !== 0 ? Math.abs(parts[0] / total - prevParts[0] / prevTotal) : 1;
    if (shareShift < 0.0035) {
      return Math.min(minDiff, shareShift * 100);
    }
    return minDiff;
  };

  let bestMinDiff = pairSpread(bestParts);
  if (bestMinDiff >= minSpreadPct) return bestParts;

  // Deterministically explore bounded local weight perturbations to separate clustered sub-item YoY rates
  for (let attempt = 1; attempt <= 28; attempt++) {
    const trialWeights = baseAdjusted.map((w, idx) => {
      const nudge = deterministicAnnualWave(year + attempt * 7, channel + idx * 11 + attempt) * ((amplitudes[idx] ?? 0.03) * 0.85);
      return Math.max(0.015, w + nudge);
    });
    const trialParts = buildFromWeights(trialWeights);
    const diff = pairSpread(trialParts);
    if (diff > bestMinDiff) {
      bestMinDiff = diff;
      bestParts = trialParts;
      if (bestMinDiff >= minSpreadPct) break;
    }
  }

  return bestParts;
}

/**
 * Automatically generates realistic, internally consistent year-wise figures
 * for the 6 primary financial drivers based on appropriate financial ratios,
 * historical trends, irregular annual business cycles, and logical interconnections.
 */
export function generateAutomaticPrimaryDrivers(
  prevYear: YearDataInput,
  assumptions: ProjectionAssumptions,
  yearNumber: number,
  stepIndex: number = 1
): {
  year: number;
  isProjected: boolean;
  salesRevenue: number;
  closingInventory: number;
  financeCost: number;
  netProfit: number;
  repaymentOfBorrowings: number;
  fixedAssetsGross: number;
} {
  if (
    prevYear.salesRevenue === 0 &&
    prevYear.netProfit === 0 &&
    prevYear.closingInventory === 0 &&
    prevYear.fixedAssetsGross === 0 &&
    prevYear.longTermBorrowings === 0 &&
    prevYear.shortTermBorrowings === 0 &&
    prevYear.financeCost === 0 &&
    prevYear.repaymentOfBorrowings === 0
  ) {
    return {
      year: yearNumber,
      isProjected: true,
      salesRevenue: 0,
      closingInventory: 0,
      financeCost: 0,
      netProfit: 0,
      repaymentOfBorrowings: 0,
      fixedAssetsGross: 0,
    };
  }

  // 1. Sales Revenue: irregular business cycle centered on revenueGrowthRate with natural expansions & occasional mild contractions
  const baseGrowthPct = assumptions.revenueGrowthRate;
  const effectiveGrowthPct = computeProjectedRevenueGrowthPct(yearNumber, baseGrowthPct, stepIndex);
  const salesRevenue = prevYear.salesRevenue > 0
    ? Math.max(100, Math.round(prevYear.salesRevenue * (1 + effectiveGrowthPct / 100)))
    : 0;
  const gRev = yoyPct(salesRevenue, prevYear.salesRevenue);

  // Cost of Sales fluctuating naturally around target gross margin while staying >= 1.5pp distinct from Revenue growth
  const targetCosPct = 100 - assumptions.grossMarginPct;
  const prevCosPct = prevYear.salesRevenue > 0 ? (prevYear.costOfSales / prevYear.salesRevenue) * 100 : targetCosPct;
  const rawCosPct = varyBounded(
    prevCosPct * 0.25 + targetCosPct * 0.75,
    yearNumber,
    2,
    2.1,
    Math.max(15, targetCosPct - 3.6),
    Math.min(92, targetCosPct + 3.6)
  );
  const minCosVal = Math.round(salesRevenue * (Math.max(15, targetCosPct - 3.6) / 100));
  const maxCosVal = Math.round(salesRevenue * (Math.min(92, targetCosPct + 3.6) / 100));
  const costOfSales = computeSpreadValue(
    salesRevenue * (rawCosPct / 100),
    prevYear.costOfSales,
    minCosVal,
    maxCosVal,
    [gRev],
    yearNumber,
    2,
    1.55
  );
  const gCos = yoyPct(costOfSales, prevYear.costOfSales);

  // 2. Closing Inventory: varied DIO cycles separated >= 1.5pp from Revenue & COGS growth
  const targetDIO = Math.max(5, assumptions.inventoryDays);
  const prevDIO = prevYear.costOfSales > 0 ? (prevYear.closingInventory / prevYear.costOfSales) * 365 : targetDIO;
  const blendedDIO = prevDIO * 0.30 + targetDIO * 0.70;
  const dioAmp = Math.max(3.0, targetDIO * 0.095);
  const minDIO = Math.max(5, targetDIO * 0.84);
  const maxDIO = Math.min(240, targetDIO * 1.16);
  const yearDIO = varyBounded(blendedDIO, yearNumber, 4, dioAmp, minDIO, maxDIO);
  const closingInventory = computeSpreadValue(
    (costOfSales / 365) * yearDIO,
    prevYear.closingInventory,
    Math.round((costOfSales / 365) * minDIO),
    Math.round((costOfSales / 365) * maxDIO),
    [gRev, gCos],
    yearNumber,
    4,
    1.55
  );

  // 3. Finance Cost: converges smoothly toward configured loan interest rate (default 15.50%) with realistic servicing timing variation
  const targetLoanRate = clampVal((assumptions.loanInterestRatePct ?? 15.5) / 100, 0, 0.38);
  const hasPrevDebt = prevYear.longTermBorrowings > 0 || prevYear.shortTermBorrowings > 0 || prevYear.financeCost > 0;
  const prevTotalDebt = hasPrevDebt
    ? Math.max(1, prevYear.longTermBorrowings + prevYear.shortTermBorrowings)
    : 0;
  let financeCost = 0;
  if (prevTotalDebt > 0 && targetLoanRate > 0) {
    const minRateBound = Math.max(0.01, targetLoanRate * 0.72);
    const maxRateBound = Math.min(0.40, targetLoanRate * 1.28);
    const rawPrevYield = prevYear.financeCost > 0 ? prevYear.financeCost / prevTotalDebt : targetLoanRate;
    const clampedPrevYield = clampVal(rawPrevYield, minRateBound, maxRateBound);
    const blendedYield = prevYear.isProjected
      ? clampedPrevYield * 0.18 + targetLoanRate * 0.82
      : clampedPrevYield * 0.35 + targetLoanRate * 0.65;
    const effectiveInterestRate = varyBounded(
      blendedYield,
      yearNumber,
      14,
      Math.max(0.006, targetLoanRate * 0.055),
      minRateBound,
      maxRateBound
    );
    const midYearDeleveragingAdj = varyBounded(0.95, yearNumber, 18, 0.028, 0.89, 0.995);
    financeCost = Math.max(120, Math.round(prevTotalDebt * effectiveInterestRate * midYearDeleveragingAdj));
  }
  const gFin = yoyPct(financeCost, prevYear.financeCost);

  // 4. Fixed Assets Additions, Disposals & Asset-Mix Depreciation
  const prevCapexPct =
    prevYear.salesRevenue > 0 && prevYear.fixedAssetsAdditions > 0
      ? clampVal((prevYear.fixedAssetsAdditions / prevYear.salesRevenue) * 100, 2.0, 15.0)
      : assumptions.capexPctOfRevenue;
  const blendedCapexPct = prevCapexPct * 0.22 + assumptions.capexPctOfRevenue * 0.78;
  const minCapexPct = Math.max(1.5, assumptions.capexPctOfRevenue * 0.58);
  const maxCapexPct = Math.min(24.0, assumptions.capexPctOfRevenue * 1.48);
  const yearCapexPct = clampVal(
    blendedCapexPct * (1 + deterministicAnnualWave(yearNumber, 11) * 0.34),
    minCapexPct,
    maxCapexPct
  );
  const rawCapexAdditions = Math.round(salesRevenue * (yearCapexPct / 100));
  const capexAdditions = computeSpreadValue(
    rawCapexAdditions,
    prevYear.fixedAssetsAdditions,
    Math.round(salesRevenue * (minCapexPct / 100)),
    Math.round(salesRevenue * (maxCapexPct / 100)),
    [gRev],
    yearNumber,
    11,
    1.6
  );
  const gAdd = yoyPct(capexAdditions, prevYear.fixedAssetsAdditions);

  const rawDispRate = varyBounded(0.012, yearNumber, 12, 0.0062, 0.0045, 0.024);
  const disposals = computeSpreadValue(
    prevYear.fixedAssetsGross * rawDispRate,
    prevYear.fixedAssetsDisposals,
    Math.round(prevYear.fixedAssetsGross * 0.0045),
    Math.round(prevYear.fixedAssetsGross * 0.024),
    [gAdd, gRev],
    yearNumber,
    12,
    1.6
  );
  const fixedAssetsGross = prevYear.fixedAssetsGross > 0
    ? Math.max(100, prevYear.fixedAssetsGross + capexAdditions - disposals)
    : capexAdditions;
  const gGrossFA = yoyPct(fixedAssetsGross, prevYear.fixedAssetsGross);
  const prevOpeningGrossEst = Math.max(0, prevYear.fixedAssetsGross - (prevYear.fixedAssetsAdditions || 0) + (prevYear.fixedAssetsDisposals || 0));
  const prevDepEst = prevYear.fixedAssetsGross > 0
    ? Math.round(((prevOpeningGrossEst + prevYear.fixedAssetsGross) / 2) * ((prevYear.depreciationRate || 10) / 100))
    : null;
  const projDepRate = computeProjectedDepRatePct(
    assumptions.depreciationRatePct || 10,
    yearNumber,
    gGrossFA,
    prevDepEst,
    prevYear.fixedAssetsGross,
    fixedAssetsGross
  );
  const depEstimate = Math.round(((prevYear.fixedAssetsGross + fixedAssetsGross) / 2) * (projDepRate / 100));

  // 5. Operating Expenses & Net Profit After Tax: maintain realistic spread across Rev, CoS, OpEx, and NP
  const grossProfit = salesRevenue - costOfSales;
  const prevOpexPct = prevYear.salesRevenue > 0 ? (prevYear.operatingExpenses / prevYear.salesRevenue) * 100 : assumptions.opexPctOfRevenue;
  const rawOpexPct = varyBounded(
    prevOpexPct * 0.25 + assumptions.opexPctOfRevenue * 0.75,
    yearNumber,
    3,
    1.75,
    Math.max(3, assumptions.opexPctOfRevenue - 2.6),
    Math.min(55, assumptions.opexPctOfRevenue + 2.6)
  );
  const minOpexVal = Math.round(salesRevenue * (Math.max(3, assumptions.opexPctOfRevenue - 2.6) / 100));
  const maxOpexVal = Math.round(salesRevenue * (Math.min(55, assumptions.opexPctOfRevenue + 2.6) / 100));
  let operatingExpenses = computeSpreadValue(
    salesRevenue * (rawOpexPct / 100),
    prevYear.operatingExpenses,
    minOpexVal,
    maxOpexVal,
    [gRev, gCos],
    yearNumber,
    3,
    1.55
  );

  const interestIncome = Math.round(salesRevenue * computeInterestIncomeRate(yearNumber, true));
  const calcNpFromOpex = (opexVal: number): number => {
    const opProfit = grossProfit - opexVal - depEstimate;
    const pbt = opProfit + interestIncome - financeCost;
    const tax = Math.round(Math.max(0, pbt) * (assumptions.taxRatePct / 100));
    return pbt - tax;
  };

  let netProfit = calcNpFromOpex(operatingExpenses);
  if (prevYear.netProfit > 0) {
    const gOpexCurr = yoyPct(operatingExpenses, prevYear.operatingExpenses);
    const gNpCurr = yoyPct(netProfit, prevYear.netProfit);
    const minNpDiff = Math.min(Math.abs(gNpCurr - gRev), Math.abs(gNpCurr - gCos), Math.abs(gNpCurr - gOpexCurr));
    if (minNpDiff < 1.48) {
      let bestOpex = operatingExpenses;
      let bestDist = Infinity;
      let bestMinSpread = minNpDiff;
      const span = Math.max(1, maxOpexVal - minOpexVal);
      for (let i = 0; i <= 40; i++) {
        const candOpex = Math.round(minOpexVal + (span * i) / 40);
        const gO = yoyPct(candOpex, prevYear.operatingExpenses);
        const npCand = calcNpFromOpex(candOpex);
        const gN = yoyPct(npCand, prevYear.netProfit);
        const sp = Math.min(
          Math.abs(gO - gRev),
          Math.abs(gO - gCos),
          Math.abs(gN - gRev),
          Math.abs(gN - gCos),
          Math.abs(gN - gO)
        );
        const dist = Math.abs(candOpex - operatingExpenses);
        if (sp >= 1.48) {
          if (bestMinSpread < 1.48 || dist < bestDist) {
            bestOpex = candOpex;
            bestDist = dist;
            bestMinSpread = sp;
          }
        } else if (bestMinSpread < 1.48 && sp > bestMinSpread) {
          bestOpex = candOpex;
          bestMinSpread = sp;
        }
      }
      operatingExpenses = bestOpex;
      netProfit = calcNpFromOpex(operatingExpenses);
    }
  }

  // 6. Repayment of Borrowings: non-uniform facility reduction pacing distinct from Finance Cost & NP
  const prevLT = prevYear.longTermBorrowings;
  const gNp = yoyPct(netProfit, prevYear.netProfit);
  let repaymentOfBorrowings = 0;
  if (prevLT > 0 || prevYear.repaymentOfBorrowings > 0) {
    const baseLT = Math.max(500, prevLT);
    const rawPrevRepayRatio = prevYear.repaymentOfBorrowings > 0 ? prevYear.repaymentOfBorrowings / baseLT : 0.12;
    const clampedPrevRepay = clampVal(rawPrevRepayRatio, 0.065, 0.23);
    const yearRepayRatio = varyBounded(clampedPrevRepay * 0.38 + 0.125 * 0.62, yearNumber, 15, 0.022, 0.065, 0.24);
    repaymentOfBorrowings = computeSpreadValue(
      Math.max(400, baseLT * yearRepayRatio),
      prevYear.repaymentOfBorrowings,
      Math.max(400, Math.round(baseLT * 0.065)),
      Math.max(450, Math.round(baseLT * 0.24)),
      [gFin, gNp],
      yearNumber,
      15,
      1.55
    );
  }

  return {
    year: yearNumber,
    isProjected: true,
    salesRevenue,
    closingInventory,
    financeCost,
    netProfit,
    repaymentOfBorrowings,
    fixedAssetsGross,
  };
}

export interface DriverAutoFillMetadata {
  key: PrimaryDriverKey;
  number: number;
  title: string;
  badge: string;
  defaultGrowthOrTrendPct: number;
  trendLabel: string;
  rationale: string;
  color: string;
}

export const PRIMARY_DRIVERS_METADATA: DriverAutoFillMetadata[] = [
  {
    key: 'salesRevenue',
    number: 1,
    title: 'Sales Revenue',
    badge: 'Top-Line Driver',
    defaultGrowthOrTrendPct: 8.5,
    trendLabel: 'Annual Revenue Growth',
    rationale: 'Compound annual sales growth based on corporate capacity and market expansion. Prior years are back-discounted at compound rate; future years compound forward.',
    color: 'emerald',
  },
  {
    key: 'closingInventory',
    number: 2,
    title: 'Closing Inventory',
    badge: 'Working Capital Driver',
    defaultGrowthOrTrendPct: 7.2,
    trendLabel: 'Annual Inventory Scaling Rate',
    rationale: 'Directly linked to Days Inventory Outstanding (DIO ~60 days) and Cost of Sales throughput, reflecting operational scale and efficiency optimizations.',
    color: 'teal',
  },
  {
    key: 'financeCost',
    number: 3,
    title: 'Finance Cost',
    badge: 'Debt Servicing Driver',
    defaultGrowthOrTrendPct: -4.2,
    trendLabel: 'Forward Debt Deleveraging / Interest Trend',
    rationale: 'Reflects scheduled term debt amortization. Backward years model higher debt balances prior to repayments; forward years model scheduled interest savings.',
    color: 'amber',
  },
  {
    key: 'netProfit',
    number: 4,
    title: 'Net Profit After Tax',
    badge: 'Bottom-Line & Equity Driver',
    defaultGrowthOrTrendPct: 9.0,
    trendLabel: 'Annual Net Earnings Growth',
    rationale: 'Bottom-line profitability with operational leverage. Future years expand margins via scale economies; backward years reflect lower operational throughput.',
    color: 'sky',
  },
  {
    key: 'repaymentOfBorrowings',
    number: 5,
    title: 'Repayment from Borrowings and Financing',
    badge: 'Debt Amortization Driver',
    defaultGrowthOrTrendPct: -5.5,
    trendLabel: 'Facility Amortization Schedule Trend',
    rationale: 'Principal debt service cash outflow. Scheduled facility reductions taper forward as debt principal amortizes down; prior years reflect structured repayment tranches.',
    color: 'rose',
  },
  {
    key: 'fixedAssetsGross',
    number: 6,
    title: 'Fixed Assets (Gross Asset Base)',
    badge: 'Capital Asset Driver',
    defaultGrowthOrTrendPct: 5.6,
    trendLabel: 'Annual Gross Asset Base Expansion',
    rationale: 'PP&E gross productive base expanding via net capex reinvestment (capex additions of 6% of sales less 1.2% asset retirements), maintaining steady additions.',
    color: 'indigo',
  },
];

/**
 * Calculates a professional, realistic year-wise sequence for any of the 6 primary financial drivers
 * based on a user's selected base year and base value.
 * Populates all previous historical years before the base year,
 * preserves the user-entered base year value strictly,
 * and populates all subsequent historical or projected years after the base year.
 */
export function calculateDriverAutoFillSeries(
  driver: PrimaryDriverKey,
  baseYear: number,
  baseValue: number,
  allYears: number[],
  assumptions: ProjectionAssumptions,
  existingYearsData?: YearDataInput[],
  customGrowthRatePct?: number
): Record<number, number> {
  const result: Record<number, number> = {};

  const meta = PRIMARY_DRIVERS_METADATA.find(m => m.key === driver);
  const growthRate = customGrowthRatePct !== undefined 
    ? customGrowthRatePct / 100 
    : (driver === 'salesRevenue' ? assumptions.revenueGrowthRate / 100 : (meta?.defaultGrowthOrTrendPct || 8.0) / 100);

  // Map of existing sales if needed for ratio-based derivation (e.g. inventory)
  const salesMap: Record<number, number> = {};
  if (existingYearsData) {
    existingYearsData.forEach(y => {
      salesMap[y.year] = y.salesRevenue;
    });
  }

  const sortedYears = [...allYears].sort((a, b) => a - b);
  const stepGrowthFor = (yr: number, rate: number, direction: 1 | -1, stepNum: number, channel: number): number => {
    if (channel === 1 && direction === 1) {
      return computeProjectedRevenueGrowthPct(yr, rate * 100, stepNum) / 100;
    }
    const absAmp = Math.max(0.012, Math.abs(rate) * 0.36);
    const wave = deterministicAnnualWave(yr, channel) * absAmp;
    if (direction === 1) {
      const maturity = stepNum > 10 ? 0.87 : stepNum > 6 ? 0.94 : 1.0;
      return clampVal(rate * maturity + wave, -0.06, 0.35);
    }
    return clampVal(rate + wave, -0.06, 0.35);
  };

  sortedYears.forEach(year => {
    if (year === baseYear) {
      // STRICTLY PRESERVE USER ENTERED BASE VALUE
      result[year] = Math.round(baseValue);
      return;
    }

    const diff = year - baseYear;

    switch (driver) {
      case 'salesRevenue': {
        if (diff > 0) {
          let val = baseValue;
          for (let step = 1; step <= diff; step++) {
            const yrStep = baseYear + step;
            const g = stepGrowthFor(yrStep, growthRate, 1, step, 1);
            val = val * (1 + g);
          }
          result[year] = Math.round(val);
        } else {
          let val = baseValue;
          const backRate = Math.max(0.04, growthRate);
          for (let step = 1; step <= Math.abs(diff); step++) {
            const yrStep = baseYear - step + 1;
            const g = Math.max(0.02, stepGrowthFor(yrStep, backRate, -1, step, 1));
            val = val / (1 + g);
          }
          result[year] = Math.round(val);
        }
        break;
      }

      case 'closingInventory': {
        const baseSales = salesMap[baseYear];
        const targetSales = salesMap[year];
        if (baseSales && targetSales && baseSales > 0 && targetSales > 0) {
          const baseRatio = baseValue / baseSales;
          const efficiency = Math.pow(1 - 0.004, diff);
          const dioWave = 1 + deterministicAnnualWave(year, 4) * 0.095;
          result[year] = Math.max(0, Math.round(targetSales * baseRatio * efficiency * dioWave));
        } else {
          const rate = growthRate || 0.072;
          let val = baseValue;
          if (diff > 0) {
            for (let step = 1; step <= diff; step++) {
              val = val * (1 + stepGrowthFor(baseYear + step, rate, 1, step, 4));
            }
          } else {
            for (let step = 1; step <= Math.abs(diff); step++) {
              val = val / (1 + Math.max(0.02, stepGrowthFor(baseYear - step + 1, rate, -1, step, 4)));
            }
          }
          result[year] = Math.max(0, Math.round(val));
        }
        break;
      }

      case 'financeCost': {
        if (diff > 0) {
          const reductionRate = Math.abs(growthRate) || 0.042;
          let val = baseValue;
          for (let step = 1; step <= diff; step++) {
            const yrStep = baseYear + step;
            const r = clampVal(reductionRate + deterministicAnnualWave(yrStep, 14) * 0.038, -0.04, 0.14);
            val = val * (1 - r);
          }
          const minInterestFloor = Math.round(baseValue * 0.30);
          result[year] = Math.max(minInterestFloor, Math.round(val));
        } else {
          let val = baseValue;
          for (let step = 1; step <= Math.abs(diff); step++) {
            const yrStep = baseYear - step + 1;
            const r = clampVal(0.052 + deterministicAnnualWave(yrStep, 14) * 0.018, 0.02, 0.10);
            val = val * (1 + r);
          }
          result[year] = Math.round(val);
        }
        break;
      }

      case 'netProfit': {
        const profitGrowth = Math.max(0.05, growthRate || 0.090);
        let val = baseValue;
        if (diff > 0) {
          for (let step = 1; step <= diff; step++) {
            val = val * (1 + stepGrowthFor(baseYear + step, profitGrowth, 1, step, 21));
          }
        } else {
          for (let step = 1; step <= Math.abs(diff); step++) {
            val = val / (1 + Math.max(0.02, stepGrowthFor(baseYear - step + 1, profitGrowth, -1, step, 21)));
          }
        }
        result[year] = Math.round(val);
        break;
      }

      case 'repaymentOfBorrowings': {
        if (diff > 0) {
          const taperRate = Math.abs(growthRate) || 0.055;
          let val = baseValue;
          for (let step = 1; step <= diff; step++) {
            const yrStep = baseYear + step;
            const r = clampVal(taperRate + deterministicAnnualWave(yrStep, 15) * 0.048, -0.05, 0.16);
            val = val * (1 - r);
          }
          const minRepaymentFloor = Math.max(350, Math.round(baseValue * 0.30));
          result[year] = Math.max(minRepaymentFloor, Math.round(val));
        } else {
          let val = baseValue;
          for (let step = 1; step <= Math.abs(diff); step++) {
            const yrStep = baseYear - step + 1;
            const r = clampVal(0.068 + deterministicAnnualWave(yrStep, 15) * 0.024, 0.025, 0.14);
            val = val * (1 + r);
          }
          result[year] = Math.round(val);
        }
        break;
      }

      case 'fixedAssetsGross': {
        const assetGrowth = Math.max(0.03, growthRate || 0.056);
        let val = baseValue;
        if (diff > 0) {
          for (let step = 1; step <= diff; step++) {
            val = val * (1 + stepGrowthFor(baseYear + step, assetGrowth, 1, step, 11));
          }
        } else {
          for (let step = 1; step <= Math.abs(diff); step++) {
            val = val / (1 + Math.max(0.02, stepGrowthFor(baseYear - step + 1, assetGrowth, -1, step, 11)));
          }
        }
        result[year] = Math.round(val);
        break;
      }
    }
  });

  return result;
}

/**
 * Generates an uneven, realistic multi-year series for any custom user-selected financial
 * statement element, ensuring YoY ratios differ cleanly (>= 1.35pp spread) across years
 * with natural ups and downs rather than flat uniform percentages.
 */
export function generateCustomElementUnevenSeries(
  baseValue: number,
  years: number[],
  baseYear?: number,
  growthRatePct = 6.5,
  seedChannel = 77
): Record<number, number> {
  const sorted = [...years].sort((a, b) => a - b);
  if (sorted.length === 0) return {};
  const anchorYr = baseYear && sorted.includes(baseYear) ? baseYear : sorted[0];
  const anchorVal = Math.round(Math.abs(baseValue) || 1000);
  const out: Record<number, number> = { [anchorYr]: anchorVal };
  const anchorIdx = sorted.indexOf(anchorYr);
  const gBase = growthRatePct / 100;

  let prevRate = growthRatePct;
  let running = anchorVal;
  for (let i = anchorIdx + 1; i < sorted.length; i++) {
    const yr = sorted[i];
    const wave = deterministicAnnualWave(yr, seedChannel + (i % 7) * 3);
    const rawTarget = running * (1 + clampVal(gBase + wave * Math.max(0.032, Math.abs(gBase) * 0.65), -0.09, 0.24));
    const val = computeSpreadValue(
      rawTarget,
      running,
      Math.max(10, Math.round(running * 0.90)),
      Math.round(running * 1.25),
      [prevRate],
      yr,
      seedChannel + (i % 5),
      1.45
    );
    prevRate = yoyPct(val, running);
    running = val;
    out[yr] = val;
  }

  let backPrevRate = -growthRatePct;
  let backRunning = anchorVal;
  for (let i = anchorIdx - 1; i >= 0; i--) {
    const yr = sorted[i];
    const wave = deterministicAnnualWave(yr, seedChannel + 19 + (i % 5) * 3);
    const rawTarget = backRunning / (1 + clampVal(gBase + wave * Math.max(0.03, Math.abs(gBase) * 0.6), -0.08, 0.22));
    const val = computeSpreadValue(
      rawTarget,
      backRunning,
      Math.max(10, Math.round(backRunning * 0.78)),
      Math.round(backRunning * 1.12),
      [backPrevRate],
      yr,
      seedChannel + 19,
      1.45
    );
    backPrevRate = yoyPct(val, backRunning);
    backRunning = val;
    out[yr] = val;
  }

  return out;
}

/**
 * Takes the current full set of year data inputs, updates a specific driver with
 * the new year-wise values, and re-derives all dependent financial metrics
 * in chronological order, maintaining full double-entry balance and cash flow integrity.
 */
export function recalculateAllYearsWithUpdatedDriver(
  currentYears: YearDataInput[],
  driver: PrimaryDriverKey | 'longTermBorrowings',
  driverValuesByYear: Record<number, number>,
  assumptions: ProjectionAssumptions
): YearDataInput[] {
  const result: YearDataInput[] = [];
  
  for (let i = 0; i < currentYears.length; i++) {
    const existing = currentYears[i];
    const prevYearObj = i > 0 ? result[i - 1] : null;
    const hasUpdatedVal = driverValuesByYear[existing.year] !== undefined;
    const newVal = hasUpdatedVal
      ? driverValuesByYear[existing.year]
      : existing[driver];

    const updatedDrivers = {
      year: existing.year,
      isProjected: existing.isProjected,
      salesRevenue: driver === 'salesRevenue' ? newVal : existing.salesRevenue,
      closingInventory: driver === 'closingInventory' ? newVal : existing.closingInventory,
      financeCost: driver === 'financeCost' ? newVal : existing.financeCost,
      netProfit: driver === 'netProfit' ? newVal : existing.netProfit,
      repaymentOfBorrowings: driver === 'repaymentOfBorrowings' ? newVal : existing.repaymentOfBorrowings,
      fixedAssetsGross: driver === 'fixedAssetsGross' ? newVal : existing.fixedAssetsGross,
    };

    const existingOverride: Partial<YearDataInput> =
      driver === 'longTermBorrowings' && hasUpdatedVal
        ? {
            ...existing,
            longTermBorrowings: Math.max(0, newVal),
            hasCustomLongTermBorrowings: true,
          }
        : existing;

    const recomputed = deriveDependentFinancials(updatedDrivers, prevYearObj, assumptions, existingOverride);
    result.push(recomputed);
  }

  return result;
}

/**
 * Derives and interconnects all dependent financial metrics across all statements
 * whenever ANY of the six primary input items is updated by the user or system:
 * 1. Sales Revenue
 * 2. Closing Inventory
 * 3. Finance Cost
 * 4. Net Profit After Tax
 * 5. Repayment from Borrowings and Financing
 * 6. Fixed Assets
 */
export function deriveDependentFinancials(
  primaryDrivers: {
    year: number;
    isProjected: boolean;
    salesRevenue: number;
    closingInventory: number;
    financeCost: number;
    netProfit: number;
    repaymentOfBorrowings: number;
    fixedAssetsGross: number;
  },
  prevYear: YearDataInput | null,
  assumptions: ProjectionAssumptions,
  existingFullYear?: Partial<YearDataInput>
): YearDataInput {
  const {
    year,
    isProjected,
    salesRevenue,
    closingInventory,
    financeCost,
    netProfit,
    repaymentOfBorrowings,
    fixedAssetsGross,
  } = primaryDrivers;

  const yearStep = prevYear ? Math.max(1, year - prevYear.year) : 1;
  const taxRate = existingFullYear?.taxRate ?? assumptions.taxRatePct ?? 25;
  const gRev = prevYear && prevYear.salesRevenue > 0 ? yoyPct(salesRevenue, prevYear.salesRevenue) : 0;
  const gInv = prevYear && prevYear.closingInventory > 0 ? yoyPct(closingInventory, prevYear.closingInventory) : 0;
  const gFin = prevYear && prevYear.financeCost > 0 ? yoyPct(financeCost, prevYear.financeCost) : 0;
  const gNp = prevYear && prevYear.netProfit > 0 ? yoyPct(netProfit, prevYear.netProfit) : 0;
  const gRepay = prevYear && prevYear.repaymentOfBorrowings > 0 ? yoyPct(repaymentOfBorrowings, prevYear.repaymentOfBorrowings) : 0;

  // 1. Reconcile Profit Before Tax & Tax Expense strictly from fixed Net Profit After Tax (NPAT)
  const profitBeforeTax = netProfit >= 0
    ? Math.round(netProfit / Math.max(0.01, 1 - taxRate / 100))
    : netProfit;
  const taxExpense = profitBeforeTax - netProfit;
  const interestIncome = Math.round(salesRevenue * computeInterestIncomeRate(year, isProjected));

  // 2. Fixed Assets Additions, Disposals & asset-mix composite Depreciation preserving fixedAssetsGross strictly
  const prevGross = prevYear ? prevYear.fixedAssetsGross : Math.round(fixedAssetsGross * 0.9);
  const gGrossFA = prevYear && prevGross > 0 ? yoyPct(fixedAssetsGross, prevGross) : null;
  const baseDepRatePct = existingFullYear?.depreciationRate ?? assumptions.depreciationRatePct ?? 10;
  const prevOpeningGross = prevYear
    ? Math.max(0, prevGross - (prevYear.fixedAssetsAdditions || 0) + (prevYear.fixedAssetsDisposals || 0))
    : 0;
  const prevDepCharge = prevYear && prevGross > 0
    ? Math.round(((prevOpeningGross + prevGross) / 2) * ((prevYear.depreciationRate || 10) / 100))
    : null;
  const depreciationRate =
    isProjected && existingFullYear?.depreciationRate === undefined
      ? computeProjectedDepRatePct(
          baseDepRatePct,
          year,
          gGrossFA,
          prevDepCharge,
          prevGross,
          fixedAssetsGross
        )
      : baseDepRatePct;

  const rawDispRate = varyBounded(0.012, year, 12, 0.0062, 0.0045, 0.024);
  const baselineDisposals = prevYear
    ? computeSpreadValue(
        prevGross * rawDispRate * yearStep,
        prevYear.fixedAssetsDisposals,
        Math.round(prevGross * 0.0045 * yearStep),
        Math.round(prevGross * 0.024 * yearStep),
        [gRev],
        year,
        12,
        1.55
      )
    : Math.round(prevGross * rawDispRate * yearStep);

  let fixedAssetsAdditions = 0;
  let fixedAssetsDisposals = baselineDisposals;
  if (fixedAssetsGross >= prevGross - baselineDisposals) {
    fixedAssetsAdditions = fixedAssetsGross - prevGross + baselineDisposals;
    if (prevYear && prevYear.fixedAssetsAdditions > 0 && prevYear.fixedAssetsDisposals > 0) {
      const gAddNow = yoyPct(fixedAssetsAdditions, prevYear.fixedAssetsAdditions);
      const gDispNow = yoyPct(fixedAssetsDisposals, prevYear.fixedAssetsDisposals);
      if (Math.abs(gAddNow - gDispNow) < 1.45) {
        const adjDisp = computeSpreadValue(
          prevGross * rawDispRate * yearStep,
          prevYear.fixedAssetsDisposals,
          Math.round(prevGross * 0.0045 * yearStep),
          Math.round(prevGross * 0.024 * yearStep),
          [gAddNow, gRev],
          year,
          12,
          1.6
        );
        if (fixedAssetsGross >= prevGross - adjDisp) {
          fixedAssetsDisposals = adjDisp;
          fixedAssetsAdditions = fixedAssetsGross - prevGross + adjDisp;
        }
      }
    }
  } else {
    fixedAssetsAdditions = 0;
    fixedAssetsDisposals = Math.max(0, prevGross - fixedAssetsGross);
  }
  const depCharge = Math.round(((prevGross + fixedAssetsGross) / 2) * (depreciationRate / 100));

  // 3. Operating Profit (EBIT) = PBT - interestIncome + financeCost
  const operatingProfit = profitBeforeTax - interestIncome + financeCost;

  // 4. Total Operating Costs = Sales - EBIT - depCharge, allocated using non-uniform CoS vs OpEx mix with >= 1.2pp spread
  const totalOperatingCosts = Math.max(0, salesRevenue - operatingProfit - depCharge);
  const targetCosPct = 100 - (assumptions.grossMarginPct || 38);
  const targetOpexPct = assumptions.opexPctOfRevenue || 17.5;
  const prevCosPct = prevYear && prevYear.salesRevenue > 0 ? (prevYear.costOfSales / prevYear.salesRevenue) * 100 : targetCosPct;
  const yearCosPct = varyBounded(
    prevCosPct * 0.25 + targetCosPct * 0.75,
    year,
    2,
    2.1,
    Math.max(15, targetCosPct - 3.6),
    Math.min(92, targetCosPct + 3.6)
  );
  const prevOpexPct = prevYear && prevYear.salesRevenue > 0 ? (prevYear.operatingExpenses / prevYear.salesRevenue) * 100 : targetOpexPct;
  const yearOpexPct = varyBounded(
    prevOpexPct * 0.25 + targetOpexPct * 0.75,
    year,
    3,
    1.75,
    Math.max(3, targetOpexPct - 2.6),
    Math.min(55, targetOpexPct + 2.6)
  );
  const targetCoSWeight = Math.max(0.1, yearCosPct / 100);
  const targetOpExWeight = Math.max(0.05, yearOpexPct / 100);
  const cosShare = targetCoSWeight / (targetCoSWeight + targetOpExWeight);
  const rawCos = Math.round(totalOperatingCosts * cosShare);

  const refsCost = prevYear ? [gRev, ...(prevYear.closingInventory > 0 ? [gInv] : []), ...(prevYear.netProfit > 0 ? [gNp] : [])] : [];
  const { costOfSales, operatingExpenses } = allocateOperatingCostsWithSpread(
    totalOperatingCosts,
    rawCos,
    salesRevenue,
    targetCosPct,
    targetOpexPct,
    prevYear,
    refsCost,
    1.48
  );
  const gCos = prevYear && prevYear.costOfSales > 0 ? yoyPct(costOfSales, prevYear.costOfSales) : 0;

  // 5. Working Capital balances derived using varied turnover days & spread separation from related items
  const baseDSO = assumptions.receivableDays || 45;
  const prevDSO = prevYear && prevYear.salesRevenue > 0 ? (prevYear.tradeReceivables / prevYear.salesRevenue) * 365 : baseDSO;
  const blendedDSO = isProjected && prevYear ? prevDSO * 0.30 + baseDSO * 0.70 : baseDSO;
  const minDSO = Math.max(5, baseDSO * 0.84);
  const maxDSO = Math.min(180, baseDSO * 1.16);
  const yearDSO = varyBounded(blendedDSO, year, 5, Math.max(2.2, baseDSO * 0.095), minDSO, maxDSO);
  const tradeReceivables = computeSpreadValue(
    (salesRevenue / 365) * yearDSO,
    prevYear?.tradeReceivables,
    Math.round((salesRevenue / 365) * minDSO),
    Math.round((salesRevenue / 365) * maxDSO),
    prevYear ? [gRev, gInv, gCos] : [],
    year,
    5,
    1.50
  );
  const gTrRec = prevYear && prevYear.tradeReceivables > 0 ? yoyPct(tradeReceivables, prevYear.tradeReceivables) : 0;

  const prevOthRecRatio = prevYear && prevYear.tradeReceivables > 0 ? clampVal(prevYear.otherReceivables / prevYear.tradeReceivables, 0.11, 0.23) : 0.16;
  const yearOthRecRatio = varyBounded(isProjected && prevYear ? prevOthRecRatio * 0.35 + 0.16 * 0.65 : 0.16, year, 7, 0.028, 0.11, 0.23);
  const otherReceivables = computeSpreadValue(
    tradeReceivables * yearOthRecRatio,
    prevYear?.otherReceivables,
    Math.round(tradeReceivables * 0.105),
    Math.round(tradeReceivables * 0.235),
    prevYear ? [gTrRec, gInv, gRev] : [],
    year,
    7,
    1.50
  );
  const gOthRec = prevYear && prevYear.otherReceivables > 0 ? yoyPct(otherReceivables, prevYear.otherReceivables) : 0;

  const baseDPO = assumptions.payableDays || 42;
  const prevDPO = prevYear && prevYear.costOfSales > 0 ? (prevYear.tradePayables / prevYear.costOfSales) * 365 : baseDPO;
  const blendedDPO = isProjected && prevYear ? prevDPO * 0.32 + baseDPO * 0.68 : baseDPO;
  const minDPO = Math.max(5, baseDPO * 0.84);
  const maxDPO = Math.min(180, baseDPO * 1.16);
  const yearDPO = varyBounded(blendedDPO, year, 6, Math.max(2.2, baseDPO * 0.095), minDPO, maxDPO);
  const tradePayables = computeSpreadValue(
    (costOfSales / 365) * yearDPO,
    prevYear?.tradePayables,
    Math.round((costOfSales / 365) * minDPO),
    Math.round((costOfSales / 365) * maxDPO),
    prevYear ? [gCos, gInv, gTrRec] : [],
    year,
    6,
    1.50
  );
  const gTrPay = prevYear && prevYear.tradePayables > 0 ? yoyPct(tradePayables, prevYear.tradePayables) : 0;

  const prevOthPayRatio = prevYear && prevYear.tradePayables > 0 ? clampVal(prevYear.otherPayables / prevYear.tradePayables, 0.18, 0.33) : 0.25;
  const yearOthPayRatio = varyBounded(isProjected && prevYear ? prevOthPayRatio * 0.35 + 0.25 * 0.65 : 0.25, year, 8, 0.032, 0.18, 0.33);
  const otherPayables = computeSpreadValue(
    tradePayables * yearOthPayRatio,
    prevYear?.otherPayables,
    Math.round(tradePayables * 0.175),
    Math.round(tradePayables * 0.335),
    prevYear ? [gTrPay, gCos] : [],
    year,
    8,
    1.50
  );

  const ocaGrowth = varyBounded(0.024, year, 9, 0.056, -0.045, 0.088);
  const otherCurrentAssets = prevYear
    ? computeSpreadValue(
        prevYear.otherCurrentAssets * Math.pow(1 + ocaGrowth, yearStep),
        prevYear.otherCurrentAssets,
        prevYear.otherCurrentAssets > 0 ? Math.max(100, Math.round(prevYear.otherCurrentAssets * Math.pow(0.955, yearStep))) : 0,
        Math.round(prevYear.otherCurrentAssets * Math.pow(1.088, yearStep)),
        [gTrRec, gOthRec, gInv],
        year,
        9,
        1.50
      )
    : Math.round(salesRevenue * varyBounded(0.015, year, 9, 0.004, 0.008, 0.025));

  // 6. Equity & Dividends / Drawings with non-uniform annual payout pacing separated from NP/Repay/FinCost
  const baseDivPct = assumptions.dividendPayoutRatio ?? 30;
  const yearDivPct = baseDivPct > 0 ? varyBounded(baseDivPct, year, 13, Math.min(6.8, baseDivPct * 0.24), 2, 75) : 0;
  const dividendsPaid =
    baseDivPct > 0 && netProfit > 0
      ? computeSpreadValue(
          netProfit * (yearDivPct / 100),
          prevYear?.dividendsPaid,
          Math.round(netProfit * (Math.max(2, baseDivPct * 0.72) / 100)),
          Math.round(netProfit * (Math.min(75, baseDivPct * 1.28) / 100)),
          prevYear ? [gNp, gRepay, gFin] : [],
          year,
          13,
          1.50
        )
      : 0;
  const isZeroSeed =
    !prevYear &&
    salesRevenue === 0 &&
    closingInventory === 0 &&
    financeCost === 0 &&
    netProfit === 0 &&
    repaymentOfBorrowings === 0 &&
    fixedAssetsGross === 0;
  const hasCustomLT = Boolean(
    existingFullYear?.hasCustomLongTermBorrowings && existingFullYear.longTermBorrowings !== undefined
  );
  const hasCustomST = Boolean(
    existingFullYear?.hasCustomShortTermBorrowings && existingFullYear.shortTermBorrowings !== undefined
  );

  const shareCapital = prevYear
    ? prevYear.shareCapital
    : (existingFullYear?.shareCapital ?? (isZeroSeed ? 0 : 12000));

  // 7. Borrowings & Financing Cash Flow preserving fixed repaymentOfBorrowings & custom LT/ST Borrowings
  const absRepayment = Math.max(0, repaymentOfBorrowings);
  const financingProceedsShares = existingFullYear?.financingProceedsShares ?? 0;

  // First projected year represents existing loans proposed for rescheduling/restructuring (no fresh loan proceeds)
  const isFirstProjectedYear = Boolean(isProjected && (!prevYear || !prevYear.isProjected));

  let financingProceedsBorrowings = 0;
  let longTermBorrowings = 0;
  let shortTermBorrowings = 0;

  if (isZeroSeed && !hasCustomLT && !hasCustomST) {
    longTermBorrowings = Math.max(0, existingFullYear?.longTermBorrowings ?? 0);
    shortTermBorrowings = Math.max(0, existingFullYear?.shortTermBorrowings ?? 0);
    financingProceedsBorrowings = existingFullYear?.financingProceedsBorrowings ?? 0;
  } else {
    const prevLT = prevYear
      ? prevYear.longTermBorrowings
      : (hasCustomLT ? Math.max(0, existingFullYear!.longTermBorrowings!) : Math.max(8000, absRepayment * 4));
    const prevST = prevYear
      ? prevYear.shortTermBorrowings
      : (hasCustomST ? Math.max(0, existingFullYear!.shortTermBorrowings!) : Math.max(2500, Math.round(salesRevenue * 0.06)));
    const prevTotalDebt = prevLT + prevST;

    // Proportionate working-capital facility proceeds with varied annual utilization; sustain operational credit headroom as business scales
    const proceedsFactor = varyBounded(0.24, year, 19, 0.065, 0.12, 0.36);
    const wcFacilityDraw = prevYear && prevTotalDebt < salesRevenue * 0.065
      ? Math.round(salesRevenue * varyBounded(0.009, year, 29, 0.0035, 0.003, 0.016))
      : 0;
    financingProceedsBorrowings = isFirstProjectedYear
      ? 0
      : prevYear
      ? Math.round(prevST * proceedsFactor) + wcFacilityDraw
      : Math.round(salesRevenue * varyBounded(0.02, year, 19, 0.004, 0.01, 0.03));
    if (!isFirstProjectedYear && prevTotalDebt + financingProceedsBorrowings < absRepayment) {
      financingProceedsBorrowings = absRepayment - prevTotalDebt + Math.round(salesRevenue * varyBounded(0.03, year, 20, 0.006, 0.018, 0.045));
    }

    if (hasCustomLT && hasCustomST) {
      longTermBorrowings = Math.max(0, existingFullYear!.longTermBorrowings!);
      shortTermBorrowings = Math.max(0, existingFullYear!.shortTermBorrowings!);
    } else if (hasCustomLT) {
      longTermBorrowings = Math.max(0, existingFullYear!.longTermBorrowings!);
      shortTermBorrowings = prevYear
        ? Math.max(0, Math.round(prevYear.shortTermBorrowings * (1 + clampVal(gRev / 100, -0.08, 0.15))))
        : (existingFullYear?.shortTermBorrowings ?? Math.max(2500, Math.round(salesRevenue * 0.06)));
    } else if (hasCustomST) {
      shortTermBorrowings = Math.max(0, existingFullYear!.shortTermBorrowings!);
      longTermBorrowings = prevYear
        ? Math.max(0, prevYear.longTermBorrowings - absRepayment + Math.round(financingProceedsBorrowings * 0.4))
        : (existingFullYear?.longTermBorrowings ?? Math.max(8000, absRepayment * 4));
    } else {
      const netDebtChange = financingProceedsBorrowings - absRepayment;
      const targetTotalDebt = Math.max(0, prevTotalDebt + netDebtChange);
      const baseLtShare = prevTotalDebt > 0 ? prevLT / prevTotalDebt : 0.72;
      const ltShare = varyBounded(baseLtShare, year, 17, 0.032, 0.54, 0.86);
      longTermBorrowings =
        prevYear && prevYear.longTermBorrowings > 0 && targetTotalDebt > 0
          ? computeSpreadValue(
              targetTotalDebt * ltShare,
              prevYear.longTermBorrowings,
              Math.round(targetTotalDebt * 0.54),
              Math.round(targetTotalDebt * 0.86),
              [gRepay, gFin],
              year,
              17,
              1.45
            )
          : Math.round(targetTotalDebt * ltShare);
      shortTermBorrowings = Math.max(0, targetTotalDebt - longTermBorrowings);
    }
  }

  const financingRepaymentBorrowings = -absRepayment;
  const financingDividendsPaid = -dividendsPaid;
  const cashFlowFromFinancing =
    financingProceedsShares + financingProceedsBorrowings - absRepayment + financingDividendsPaid;

  const amortStep = Math.round(varyBounded(100, year, 15, 28, 55, 155)) * yearStep;
  const intangibleAssets = prevYear
    ? Math.max(0, prevYear.intangibleAssets - Math.min(prevYear.intangibleAssets, amortStep))
    : (existingFullYear?.intangibleAssets ?? (isZeroSeed ? 0 : Math.round(salesRevenue * 0.03)));
  const otherNonCurrentAssets = prevYear
    ? prevYear.otherNonCurrentAssets
    : (existingFullYear?.otherNonCurrentAssets ?? (isZeroSeed ? 0 : Math.round(salesRevenue * 0.08)));

  // Estimated Cash & Equivalents (actual Closing Cash is computed cleanly by Cash Flow Engine)
  const cashAndEquivalentsBS = prevYear
    ? Math.max(0, Math.round(prevYear.cashAndEquivalentsBS + netProfit * 0.35))
    : (existingFullYear?.cashAndEquivalentsBS ?? (isZeroSeed ? 0 : Math.round(salesRevenue * 0.085)));

  return {
    ...(existingFullYear || {}),
    year,
    isProjected,
    salesRevenue,
    closingInventory,
    financeCost,
    netProfit,
    repaymentOfBorrowings: absRepayment,
    fixedAssetsGross,
    cashFlowFromFinancing,
    fixedAssetsAdditions,
    fixedAssetsDisposals,
    depreciationRate,
    costOfSales,
    operatingExpenses,
    taxRate,
    tradeReceivables,
    otherReceivables,
    tradePayables,
    otherPayables,
    cashAndEquivalentsBS,
    otherCurrentAssets,
    intangibleAssets,
    otherNonCurrentAssets,
    shortTermBorrowings,
    longTermBorrowings,
    hasCustomLongTermBorrowings: hasCustomLT,
    hasCustomShortTermBorrowings: hasCustomST,
    shareCapital,
    dividendsPaid,
    financingProceedsShares,
    financingProceedsBorrowings,
    financingRepaymentBorrowings,
    financingDividendsPaid,
  };
}

/**
 * Chartered-Accountancy rational recalculation when Long-Term Borrowings, Short-Term Borrowings,
 * or the Loan Interest Rate (% p.a.) changes for a year: preserves Revenue/COGS/OpEx (EBITDA)
 * while dynamically adjusting Finance Cost -> PBT -> Tax Expense -> Net Profit -> Equity -> Cash Flow.
 */
export function recalculateYearOnLoanOrRateChange(
  yearInput: YearDataInput,
  prevYear: YearDataInput | null,
  assumptions: ProjectionAssumptions,
  overrides?: {
    longTermBorrowings?: number;
    shortTermBorrowings?: number;
    hasCustomLongTermBorrowings?: boolean;
    hasCustomShortTermBorrowings?: boolean;
    loanRateOverridePct?: number;
    recalculateFinanceCostFromDebt?: boolean;
  }
): YearDataInput {
  const newLT = Math.max(0, overrides?.longTermBorrowings ?? yearInput.longTermBorrowings);
  const newST = Math.max(0, overrides?.shortTermBorrowings ?? yearInput.shortTermBorrowings);
  const totalDebt = r2(newLT + newST);
  const hasCustomLT = overrides?.hasCustomLongTermBorrowings ?? yearInput.hasCustomLongTermBorrowings;
  const hasCustomST = overrides?.hasCustomShortTermBorrowings ?? yearInput.hasCustomShortTermBorrowings;

  const ratePct = clampVal(overrides?.loanRateOverridePct ?? assumptions.loanInterestRatePct ?? 15.5, 0, 45);
  const shouldRecalcInterest =
    !yearInput.isScheduleLocked && (overrides?.recalculateFinanceCostFromDebt ?? true);

  let nextFinanceCost = yearInput.financeCost;
  let nextNetProfit = yearInput.netProfit;

  if (shouldRecalcInterest) {
    nextFinanceCost = totalDebt > 0 && ratePct > 0 ? Math.round(totalDebt * (ratePct / 100)) : 0;
    const taxRate = yearInput.taxRate || assumptions.taxRatePct || 25;
    const oldTax =
      yearInput.netProfit >= 0
        ? Math.round(yearInput.netProfit / Math.max(0.01, 1 - taxRate / 100) - yearInput.netProfit)
        : 0;
    const oldPBT = r2(yearInput.netProfit + oldTax);
    const deltaFin = nextFinanceCost - yearInput.financeCost;
    const newPBT = r2(oldPBT - deltaFin);
    const newTax = newPBT > 0 ? Math.round(newPBT * (taxRate / 100)) : 0;
    nextNetProfit = r2(newPBT - newTax);
  }

  return deriveDependentFinancials(
    {
      year: yearInput.year,
      isProjected: yearInput.isProjected,
      salesRevenue: yearInput.salesRevenue,
      closingInventory: yearInput.closingInventory,
      financeCost: nextFinanceCost,
      netProfit: nextNetProfit,
      repaymentOfBorrowings: yearInput.repaymentOfBorrowings,
      fixedAssetsGross: yearInput.fixedAssetsGross,
    },
    prevYear,
    assumptions,
    {
      ...yearInput,
      longTermBorrowings: newLT,
      shortTermBorrowings: newST,
      hasCustomLongTermBorrowings: hasCustomLT,
      hasCustomShortTermBorrowings: hasCustomST,
      financeCost: nextFinanceCost,
      netProfit: nextNetProfit,
    }
  );
}

/**
 * Generates proportionate, financially logical YearDataInput for any individual year
 * (historical or projected) anchored to the closest existing year.
 */
export function generateYearDataForTargetYear(
  targetYear: number,
  isProjected: boolean,
  existingYears: YearDataInput[],
  assumptions: ProjectionAssumptions
): YearDataInput {
  const sorted = [...existingYears].sort((a, b) => a.year - b.year);
  if (sorted.length === 0) {
    return { ...createZeroHistoricalInputs([targetYear])[0], isProjected };
  }

  // Find closest preceding year or earliest year
  const preceding = [...sorted].reverse().find(y => y.year < targetYear);
  const anchor = preceding || sorted[0];
  if (
    anchor.salesRevenue === 0 &&
    anchor.closingInventory === 0 &&
    anchor.financeCost === 0 &&
    anchor.netProfit === 0 &&
    anchor.repaymentOfBorrowings === 0 &&
    anchor.fixedAssetsGross === 0 &&
    anchor.longTermBorrowings === 0
  ) {
    return { ...createZeroHistoricalInputs([targetYear])[0], isProjected };
  }
  const diff = targetYear - anchor.year;

  const revRate = (assumptions.revenueGrowthRate || 8.5) / 100;
  const stepAccum = (baseR: number, channel: number, dirInvert = false): number => {
    let mult = 1;
    const steps = Math.abs(diff);
    for (let s = 1; s <= steps; s++) {
      const yr = diff >= 0 ? anchor.year + s : anchor.year - s + 1;
      const r =
        channel === 1 && diff >= 0 && !dirInvert
          ? computeProjectedRevenueGrowthPct(yr, baseR * 100, s) / 100
          : clampVal(baseR + deterministicAnnualWave(yr, channel) * Math.max(0.012, Math.abs(baseR) * 0.34), -0.15, 0.35);
      if ((diff >= 0 && !dirInvert) || (diff < 0 && dirInvert)) {
        mult *= 1 + r;
      } else {
        mult /= 1 + Math.max(0.02, r);
      }
    }
    return mult;
  };

  const salesScale = stepAccum(diff >= 0 ? revRate : Math.max(0.04, revRate), 1);
  const invScale = stepAccum(diff >= 0 ? revRate * 0.88 : Math.max(0.035, revRate * 0.88), 4);

  const salesRevenue = Math.max(100, Math.round(anchor.salesRevenue * salesScale));
  const gRevTarget = preceding && preceding.salesRevenue > 0 ? yoyPct(salesRevenue, preceding.salesRevenue) : revRate * 100;

  const rawInv = Math.max(0, Math.round(anchor.closingInventory * invScale));
  const closingInventory = preceding && preceding.closingInventory > 0
    ? computeSpreadValue(
        rawInv,
        preceding.closingInventory,
        Math.round(rawInv * 0.84),
        Math.round(rawInv * 1.16),
        [gRevTarget],
        targetYear,
        4,
        1.45
      )
    : rawInv;
  const gInvTarget = preceding && preceding.closingInventory > 0 ? yoyPct(closingInventory, preceding.closingInventory) : gRevTarget;

  const rawFin = Math.max(
    50,
    Math.round(anchor.financeCost * stepAccum(diff >= 0 ? -0.04 : 0.05, 14))
  );
  const financeCost = preceding && preceding.financeCost > 0
    ? computeSpreadValue(
        rawFin,
        preceding.financeCost,
        Math.max(50, Math.round(rawFin * 0.84)),
        Math.round(rawFin * 1.16),
        [gRevTarget],
        targetYear,
        14,
        1.45
      )
    : rawFin;

  const rawNp = Math.round(anchor.netProfit * stepAccum(Math.max(0.04, revRate * 1.04), 21));
  const netProfit = preceding && preceding.netProfit > 0
    ? computeSpreadValue(
        rawNp,
        preceding.netProfit,
        Math.round(rawNp * 0.82),
        Math.round(rawNp * 1.18),
        [gRevTarget, gInvTarget],
        targetYear,
        21,
        1.45
      )
    : rawNp;

  const rawRepay = Math.max(
    0,
    Math.round(anchor.repaymentOfBorrowings * stepAccum(diff >= 0 ? -0.055 : 0.06, 15))
  );
  const gFinTarget = preceding && preceding.financeCost > 0 ? yoyPct(financeCost, preceding.financeCost) : 0;
  const repaymentOfBorrowings = preceding && preceding.repaymentOfBorrowings > 0
    ? computeSpreadValue(
        rawRepay,
        preceding.repaymentOfBorrowings,
        Math.max(0, Math.round(rawRepay * 0.82)),
        Math.round(rawRepay * 1.18),
        [gFinTarget, gRevTarget],
        targetYear,
        15,
        1.45
      )
    : rawRepay;

  const rawFA = Math.max(
    100,
    Math.round(anchor.fixedAssetsGross * stepAccum(0.055, 11))
  );
  const fixedAssetsGross = preceding && preceding.fixedAssetsGross > 0
    ? computeSpreadValue(
        rawFA,
        preceding.fixedAssetsGross,
        Math.max(100, Math.round(rawFA * 0.90)),
        Math.round(rawFA * 1.12),
        [gRevTarget],
        targetYear,
        11,
        1.45
      )
    : rawFA;

  return deriveDependentFinancials(
    {
      year: targetYear,
      isProjected,
      salesRevenue,
      closingInventory,
      financeCost,
      netProfit,
      repaymentOfBorrowings,
      fixedAssetsGross,
    },
    preceding || null,
    assumptions
  );
}

// Generate Projected Years based on selectedProjectedYears (or projectedYearsCount)
export function generateProjectedYears(
  baseInput: YearDataInput,
  assumptions: ProjectionAssumptions
): YearDataInput[] {
  const targetYears = assumptions.selectedProjectedYears && assumptions.selectedProjectedYears.length > 0
    ? [...assumptions.selectedProjectedYears].sort((a, b) => a - b)
    : Array.from({ length: Math.max(1, assumptions.projectedYearsCount) }, (_, i) => baseInput.year + i + 1);

  const projected: YearDataInput[] = [];
  let prevYear = baseInput;
  let stepCounter = 0;

  for (const yearNum of targetYears) {
    const step = Math.max(1, yearNum - prevYear.year);
    let tempPrev = prevYear;
    if (step > 1) {
      for (let s = 1; s < step; s++) {
        stepCounter += 1;
        const midDrivers = generateAutomaticPrimaryDrivers(tempPrev, assumptions, prevYear.year + s, stepCounter);
        tempPrev = deriveDependentFinancials(midDrivers, tempPrev, assumptions);
      }
    }
    stepCounter += 1;
    const finalDrivers = generateAutomaticPrimaryDrivers(tempPrev, assumptions, yearNum, stepCounter);
    const fullYear = deriveDependentFinancials(finalDrivers, prevYear, assumptions);
    projected.push(fullYear);
    prevYear = fullYear;
  }

  return projected;
}

// 1. Calculate Schedule of Fixed Assets (Strictly preserves input.fixedAssetsGross as Closing Cost by default)
export function calculateScheduleOfFixedAssets(
  inputs: YearDataInput[],
  options?: StatementCalculationOptions
): ScheduleOfFixedAssetsYear[] {
  const result: ScheduleOfFixedAssetsYear[] = [];
  const firstGross = inputs[0]?.fixedAssetsGross ?? 0;
  const baseAccumDep = Math.round(firstGross * 0.35);
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const cellOverrides = options?.cellOverrides;

  const assetClasses = [
    { name: 'Freehold Land & Industrial Buildings', depRate: 2.5, initCostShare: 0.40, initDepShare: 0.18 },
    { name: 'Plant, Machinery & Automated Systems', depRate: 10.0, initCostShare: 0.35, initDepShare: 0.45 },
    { name: 'Office Equipment & Fixtures', depRate: 15.0, initCostShare: 0.10, initDepShare: 0.15 },
    { name: 'Motor Vehicles & Logistics Fleet', depRate: 20.0, initCostShare: 0.08, initDepShare: 0.12 },
    { name: 'Computer, IT & Telecommunications', depRate: 25.0, initCostShare: 0.07, initDepShare: 0.10 },
  ];

  let prevBaseClosingCost = 0;
  let prevBaseClosingDep = 0;
  let prevBaseCategories: FixedAssetCategory[] | null = null;

  inputs.forEach((input, index) => {
    const isFirst = index === 0;
    const closingCost = input.fixedAssetsGross; // STRICTLY FIXED USER INPUT
    const openingCost = isFirst
      ? Math.max(0, closingCost - input.fixedAssetsAdditions + input.fixedAssetsDisposals)
      : prevBaseClosingCost;

    let additions = input.fixedAssetsAdditions;
    let disposals = input.fixedAssetsDisposals;

    // Ensure openingCost + additions - disposals === closingCost (user's fixedAssetsGross)
    if (!isFirst) {
      const expectedClosing = openingCost + additions - disposals;
      if (expectedClosing !== closingCost) {
        if (closingCost >= openingCost - disposals) {
          additions = closingCost - openingCost + disposals;
        } else {
          additions = 0;
          disposals = Math.max(0, openingCost - closingCost);
        }
      }
    }

    const openingDep = isFirst ? baseAccumDep : prevBaseClosingDep;
    const depCharge = Math.round(((openingCost + closingCost) / 2) * (input.depreciationRate / 100));
    const depDisposals = Math.round(disposals * computeDepDisposalsRatio(input.year, input.isProjected));
    const closingDep = Math.max(0, openingDep + depCharge - depDisposals);
    const netBookValue = Math.max(0, closingCost - closingDep);

    prevBaseClosingCost = closingCost;
    prevBaseClosingDep = closingDep;

    // Roll forward category openings cleanly from prior year's closings
    const catCostOpenings: number[] = [];
    const catDepOpenings: number[] = [];
    if (isFirst || !prevBaseCategories) {
      let remCost = openingCost;
      let remDep = openingDep;
      assetClasses.forEach((cls, idx) => {
        if (idx === assetClasses.length - 1) {
          catCostOpenings.push(r2(remCost));
          catDepOpenings.push(r2(remDep));
        } else {
          const c = Math.round(openingCost * cls.initCostShare);
          const d = Math.round(openingDep * cls.initDepShare);
          catCostOpenings.push(c);
          catDepOpenings.push(d);
          remCost -= c;
          remDep -= d;
        }
      });
    } else {
      assetClasses.forEach((_, idx) => {
        catCostOpenings.push(prevBaseCategories![idx].costClosing);
        catDepOpenings.push(prevBaseCategories![idx].depClosing);
      });
    }

    // Allocate Additions & Disposals with non-uniform year-specific capital investment weights & spread separation
    const prevAddParts = prevBaseCategories ? prevBaseCategories.map(c => c.additions) : undefined;
    const prevDispParts = prevBaseCategories ? prevBaseCategories.map(c => c.disposals) : undefined;
    const catAdditionsArr = splitTotalNonUniform(
      additions,
      input.year,
      30,
      [0.22, 0.44, 0.13, 0.11, 0.10],
      [0.055, 0.065, 0.032, 0.03, 0.026],
      prevAddParts,
      1.35
    );
    const rawDisposalsArr = splitTotalNonUniform(
      disposals,
      input.year,
      45,
      [0.08, 0.46, 0.16, 0.18, 0.12],
      [0.03, 0.06, 0.035, 0.035, 0.03],
      prevDispParts,
      1.35
    );
    // Cap each class disposal at available gross cost + additions; redistribute any excess to larger classes
    const catDisposalsArr = [...rawDisposalsArr];
    let excessDisp = 0;
    for (let i = 0; i < assetClasses.length; i++) {
      const maxAvail = Math.max(0, catCostOpenings[i] + catAdditionsArr[i]);
      if (catDisposalsArr[i] > maxAvail) {
        excessDisp += catDisposalsArr[i] - maxAvail;
        catDisposalsArr[i] = maxAvail;
      }
    }
    if (excessDisp > 0) {
      for (let i = 1; i < assetClasses.length && excessDisp > 0; i++) {
        const room = Math.max(0, catCostOpenings[i] + catAdditionsArr[i] - catDisposalsArr[i]);
        const addHere = Math.min(room, excessDisp);
        catDisposalsArr[i] = r2(catDisposalsArr[i] + addHere);
        excessDisp = r2(excessDisp - addHere);
      }
    }

    const catCostClosings = assetClasses.map((_, i) =>
      r2(catCostOpenings[i] + catAdditionsArr[i] - catDisposalsArr[i])
    );

    // Rate-weighted depreciation charge scaled cleanly to exact totalDepCharge
    const rawDepWeights = assetClasses.map((cls, i) => {
      const avgGross = Math.max(0, (catCostOpenings[i] + catCostClosings[i]) / 2);
      const jitter = 1 + deterministicAnnualWave(input.year, 55 + i * 2) * 0.04;
      return Math.max(0, avgGross * (cls.depRate / 100) * jitter);
    });
    const rawDepSum = rawDepWeights.reduce((acc, v) => acc + v, 0);
    const catDepCharges: number[] = [];
    let remDepCharge = depCharge;
    for (let i = 0; i < assetClasses.length; i++) {
      if (i === assetClasses.length - 1) {
        catDepCharges.push(Math.max(0, r2(remDepCharge)));
      } else {
        const chg = rawDepSum > 0 ? Math.round(depCharge * (rawDepWeights[i] / rawDepSum)) : 0;
        const clampedChg = Math.min(chg, Math.max(0, remDepCharge));
        catDepCharges.push(clampedChg);
        remDepCharge -= clampedChg;
      }
    }

    // Allocate depDisposals proportional to category disposals & accumulated dep progress
    const rawDepDispWeights = assetClasses.map((_, i) =>
      Math.max(0, catDisposalsArr[i] * Math.max(0.1, catDepOpenings[i] / Math.max(1, catCostOpenings[i])))
    );
    const rawDepDispSum = rawDepDispWeights.reduce((acc, v) => acc + v, 0);
    const catDepDisposalsArr: number[] = [];
    let remDepDisp = depDisposals;
    for (let i = 0; i < assetClasses.length; i++) {
      const maxDepAvail = Math.max(0, catDepOpenings[i] + catDepCharges[i]);
      if (i === assetClasses.length - 1) {
        catDepDisposalsArr.push(Math.min(maxDepAvail, Math.max(0, r2(remDepDisp))));
      } else {
        const d = rawDepDispSum > 0 ? Math.round(depDisposals * (rawDepDispWeights[i] / rawDepDispSum)) : 0;
        const clampedD = Math.min(maxDepAvail, Math.max(0, Math.min(d, remDepDisp)));
        catDepDisposalsArr.push(clampedD);
        remDepDisp -= clampedD;
      }
    }
    const depDispShortfall = r2(depDisposals - catDepDisposalsArr.reduce((a, b) => a + b, 0));
    if (Math.abs(depDispShortfall) > 0.001) {
      for (let i = 0; i < assetClasses.length; i++) {
        const maxDepAvail = Math.max(0, catDepOpenings[i] + catDepCharges[i]);
        const room = maxDepAvail - catDepDisposalsArr[i];
        if (room >= depDispShortfall) {
          catDepDisposalsArr[i] = r2(catDepDisposalsArr[i] + depDispShortfall);
          break;
        }
      }
    }

    const baseCategories: FixedAssetCategory[] = assetClasses.map((cls, i) => {
      const costOpeningCat = catCostOpenings[i];
      const addCat = catAdditionsArr[i];
      const dispCat = catDisposalsArr[i];
      const costClosingCat = catCostClosings[i];
      const depOpeningCat = catDepOpenings[i];
      const depChargeCat = catDepCharges[i];
      const depDisposalsCat = catDepDisposalsArr[i];
      const depClosingCat = Math.max(0, r2(depOpeningCat + depChargeCat - depDisposalsCat));
      return {
        name: cls.name,
        costOpening: costOpeningCat,
        additions: addCat,
        disposals: dispCat,
        costClosing: costClosingCat,
        depOpening: depOpeningCat,
        depCharge: depChargeCat,
        depDisposals: depDisposalsCat,
        depClosing: depClosingCat,
        netBookValue: r2(costClosingCat - depClosingCat),
        depRate: cls.depRate,
      };
    });
    prevBaseCategories = baseCategories;

    let hasCategoryOverride = false;
    const processedCategories: FixedAssetCategory[] = baseCategories.map((cat, catIdx) => {
      const prefix = `fa_${catIdx}`;
      const ovCostOpening = getCellOverride(cellOverrides, prefix, 'costOpening', input.year);
      const ovAdditions = getCellOverride(cellOverrides, prefix, 'additions', input.year);
      const ovDisposals = getCellOverride(cellOverrides, prefix, 'disposals', input.year);
      const ovCostClosing = getCellOverride(cellOverrides, prefix, 'costClosing', input.year);
      const ovDepOpening = getCellOverride(cellOverrides, prefix, 'depOpening', input.year);
      const ovDepCharge = getCellOverride(cellOverrides, prefix, 'depCharge', input.year);
      const ovDepDisposals = getCellOverride(cellOverrides, prefix, 'depDisposals', input.year);
      const ovDepClosing = getCellOverride(cellOverrides, prefix, 'depClosing', input.year);
      const ovDepRate = getCellOverride(cellOverrides, prefix, 'depRate', input.year);

      const catEdited =
        ovCostOpening !== undefined ||
        ovAdditions !== undefined ||
        ovDisposals !== undefined ||
        ovCostClosing !== undefined ||
        ovDepOpening !== undefined ||
        ovDepCharge !== undefined ||
        ovDepDisposals !== undefined ||
        ovDepClosing !== undefined ||
        ovDepRate !== undefined;

      if (!catEdited) {
        return cat;
      }

      hasCategoryOverride = true;
      const costOpening = ovCostOpening ?? cat.costOpening;
      const catAdditions = ovAdditions ?? cat.additions;
      const catDisposals = ovDisposals ?? cat.disposals;
      const catCostClosing = ovCostClosing ?? r2(costOpening + catAdditions - catDisposals);
      const catDepOpening = ovDepOpening ?? cat.depOpening;
      const catDepRate = ovDepRate ?? cat.depRate;
      const catDepCharge =
        ovDepCharge !== undefined
          ? ovDepCharge
          : ovDepRate !== undefined
          ? Math.round(((costOpening + catCostClosing) / 2) * (catDepRate / 100))
          : cat.depCharge;
      const catDepDisposals = ovDepDisposals ?? cat.depDisposals;
      const catDepClosing =
        ovDepClosing ?? Math.max(0, r2(catDepOpening + catDepCharge - catDepDisposals));

      return {
        ...cat,
        costOpening,
        additions: catAdditions,
        disposals: catDisposals,
        costClosing: catCostClosing,
        depOpening: catDepOpening,
        depCharge: catDepCharge,
        depDisposals: catDepDisposals,
        depClosing: catDepClosing,
        netBookValue: r2(catCostClosing - catDepClosing),
        depRate: catDepRate,
      };
    });

    const customFACategories: FixedAssetCategory[] = (options?.statementConfig || [])
      .filter(item => item.isCustom && item.isVisible && item.statement === 'fixedAssets')
      .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
      .map(item => {
        const val = getCustomItemYearValue(item, input.year);
        return {
          name: item.label,
          costOpening: val,
          additions: 0,
          disposals: 0,
          costClosing: val,
          depOpening: 0,
          depCharge: 0,
          depDisposals: 0,
          depClosing: 0,
          netBookValue: val,
          depRate: 0,
        };
      });

    const visibleStandardCategories = processedCategories.filter((_, catIdx) => {
      const cfgId = FIXED_ASSET_CATEGORY_CONFIG_IDS[catIdx];
      return cfgId ? isConfigElementVisible(configMap, cfgId) : true;
    });

    const allHiddenOrCustom =
      visibleStandardCategories.length !== processedCategories.length || customFACategories.length > 0;
    const activeCategories = [...visibleStandardCategories, ...customFACategories];

    if (!allHiddenOrCustom && !hasCategoryOverride) {
      result.push({
        year: input.year,
        isProjected: input.isProjected,
        categories: activeCategories,
        allCategories: processedCategories,
        totalCostOpening: openingCost,
        totalAdditions: additions,
        totalDisposals: disposals,
        totalCostClosing: closingCost,
        totalDepOpening: openingDep,
        totalDepCharge: depCharge,
        totalDepDisposals: depDisposals,
        totalDepClosing: closingDep,
        totalNetBookValue: netBookValue,
      });
    } else {
      const sum = (fn: (c: FixedAssetCategory) => number) =>
        r2(activeCategories.reduce((acc, c) => acc + fn(c), 0));
      result.push({
        year: input.year,
        isProjected: input.isProjected,
        categories: activeCategories,
        allCategories: processedCategories,
        totalCostOpening: sum(c => c.costOpening),
        totalAdditions: sum(c => c.additions),
        totalDisposals: sum(c => c.disposals),
        totalCostClosing: sum(c => c.costClosing),
        totalDepOpening: sum(c => c.depOpening),
        totalDepCharge: sum(c => c.depCharge),
        totalDepDisposals: sum(c => c.depDisposals),
        totalDepClosing: sum(c => c.depClosing),
        totalNetBookValue: sum(c => c.netBookValue),
      });
    }
  });

  return result;
}

const r2 = (n: number): number => Math.round((n + Number.EPSILON) * 100) / 100;

// 2. Calculate Income Statement (Strictly preserves salesRevenue, financeCost, and netProfit by default)
export function calculateIncomeStatement(
  inputs: YearDataInput[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  options?: StatementCalculationOptions
): IncomeStatementYear[] {
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const cellOverrides = options?.cellOverrides;

  // Baseline un-filtered FA depreciation so Cost of Sales & OpEx stay anchored when FA categories are deselected
  const firstGross = inputs[0]?.fixedAssetsGross ?? 0;
  let prevBaseClosingCost = 0;
  let prevBaseClosingDep = 0;
  const baseDepCharges: number[] = inputs.map((input, index) => {
    const isFirst = index === 0;
    const closingCost = input.fixedAssetsGross;
    const openingCost = isFirst
      ? Math.max(0, closingCost - input.fixedAssetsAdditions + input.fixedAssetsDisposals)
      : prevBaseClosingCost;
    let disposals = input.fixedAssetsDisposals;
    if (!isFirst && closingCost < openingCost - disposals) {
      disposals = Math.max(0, openingCost - closingCost);
    }
    const openingDep = isFirst ? Math.round(firstGross * 0.35) : prevBaseClosingDep;
    const depCharge = Math.round(((openingCost + closingCost) / 2) * (input.depreciationRate / 100));
    const depDisposals = Math.round(disposals * computeDepDisposalsRatio(input.year, input.isProjected));
    const closingDep = Math.max(0, openingDep + depCharge - depDisposals);
    prevBaseClosingCost = closingCost;
    prevBaseClosingDep = closingDep;
    return depCharge;
  });

  const allIncKeys = [
    'inc_revenue',
    'inc_costOfSales',
    'inc_operatingExpenses',
    'inc_depreciationExpense',
    'inc_interestIncome',
    'inc_financeCost',
    'inc_taxExpense',
  ] as const;

  return inputs.map((input, index) => {
    const baseRevenue = input.salesRevenue;
    const baseFinanceCost = input.financeCost;
    const baseNetProfit = input.netProfit;
    const hasCents =
      !Number.isInteger(baseRevenue) ||
      !Number.isInteger(baseFinanceCost) ||
      !Number.isInteger(baseNetProfit) ||
      !Number.isInteger(input.repaymentOfBorrowings);
    const r = (n: number) => (hasCents ? r2(n) : Math.round(n));

    const baselineDepCharge = baseDepCharges[index] ?? r(input.fixedAssetsGross * 0.1);
    const activeFADep = fixedAssetsSchedule[index]?.totalDepCharge ?? baselineDepCharge;

    const taxRate = input.taxRate || 25;
    const baseTaxExpense =
      baseNetProfit >= 0 ? r(baseNetProfit / Math.max(0.01, 1 - taxRate / 100) - baseNetProfit) : 0;
    const basePBT = r2(baseNetProfit + baseTaxExpense);
    let baseInterestIncome = r(baseRevenue * computeInterestIncomeRate(input.year, input.isProjected));

    const baseOperatingProfit = r2(basePBT - baseInterestIncome + baseFinanceCost);
    const requiredOperatingCosts = r2(baseRevenue - baseOperatingProfit - baselineDepCharge);
    let baseCostOfSales = input.costOfSales;
    let baseOperatingExpenses = input.operatingExpenses;

    if (requiredOperatingCosts >= 0) {
      const existingSum = Math.max(1, input.costOfSales + input.operatingExpenses);
      const cosRatio = input.costOfSales > 0 ? input.costOfSales / existingSum : 0.75;
      baseCostOfSales = r(requiredOperatingCosts * cosRatio);
      baseOperatingExpenses = r2(requiredOperatingCosts - baseCostOfSales);
    } else {
      baseCostOfSales = 0;
      baseOperatingExpenses = 0;
      baseInterestIncome = r2(baseInterestIncome + Math.abs(requiredOperatingCosts));
    }

    const baseGrossProfit = r2(baseRevenue - baseCostOfSales);

    const anyIncHidden = allIncKeys.some(k => !isConfigElementVisible(configMap, k));
    const hasIncOverride =
      ['revenue', 'costOfSales', 'grossProfit', 'operatingExpenses', 'depreciationExpense', 'operatingProfit', 'interestIncome', 'financeCost', 'profitBeforeTax', 'taxExpense', 'netProfit']
        .some(f => getCellOverride(cellOverrides, 'income', f, input.year) !== undefined);
    const customOpRev = getVisibleCustomItemsSum(options?.statementConfig, 'income', input.year, 'Revenue & Gross Margin');
    const customOpExp = getVisibleCustomItemsSum(options?.statementConfig, 'income', input.year, 'Operating Expenses');
    const customNonOp = getVisibleCustomItemsSum(options?.statementConfig, 'income', input.year, 'Non-Operating & Tax');
    const customUnscoped = (options?.statementConfig || [])
      .filter(
        item =>
          item.isCustom &&
          item.isVisible &&
          item.statement === 'income' &&
          item.section !== 'Revenue & Gross Margin' &&
          item.section !== 'Operating Expenses' &&
          item.section !== 'Non-Operating & Tax'
      )
      .reduce((acc, item) => acc + getCustomItemYearValue(item, input.year), 0);

    const hasCustomIncome =
      (options?.statementConfig || []).some(
        item => item.isCustom && item.isVisible && item.statement === 'income'
      );
    const faDepChanged = Math.abs(activeFADep - baselineDepCharge) > 0.01;

    if (!anyIncHidden && !hasIncOverride && !hasCustomIncome && !faDepChanged) {
      return {
        year: input.year,
        isProjected: input.isProjected,
        revenue: baseRevenue,
        costOfSales: baseCostOfSales,
        grossProfit: baseGrossProfit,
        operatingExpenses: baseOperatingExpenses,
        depreciationExpense: baselineDepCharge,
        operatingProfit: baseOperatingProfit,
        interestIncome: baseInterestIncome,
        financeCost: baseFinanceCost,
        profitBeforeTax: basePBT,
        taxExpense: baseTaxExpense,
        netProfit: baseNetProfit,
      };
    }

    const baseRevPart = isConfigElementVisible(configMap, 'inc_revenue')
      ? getCellOverride(cellOverrides, 'income', 'revenue', input.year) ?? baseRevenue
      : 0;
    const revenue = r2(baseRevPart + customOpRev);
    const costOfSales = isConfigElementVisible(configMap, 'inc_costOfSales')
      ? getCellOverride(cellOverrides, 'income', 'costOfSales', input.year) ?? baseCostOfSales
      : 0;
    const grossProfit =
      getCellOverride(cellOverrides, 'income', 'grossProfit', input.year) ?? r2(revenue - costOfSales);
    const baseOpexPart = isConfigElementVisible(configMap, 'inc_operatingExpenses')
      ? getCellOverride(cellOverrides, 'income', 'operatingExpenses', input.year) ?? baseOperatingExpenses
      : 0;
    const operatingExpenses = r2(baseOpexPart - customOpExp);
    const depreciationExpense = isConfigElementVisible(configMap, 'inc_depreciationExpense')
      ? getCellOverride(cellOverrides, 'income', 'depreciationExpense', input.year) ?? activeFADep
      : 0;
    const operatingProfit =
      getCellOverride(cellOverrides, 'income', 'operatingProfit', input.year) ??
      r2(grossProfit - operatingExpenses - depreciationExpense);
    const baseIntIncPart = isConfigElementVisible(configMap, 'inc_interestIncome')
      ? getCellOverride(cellOverrides, 'income', 'interestIncome', input.year) ?? baseInterestIncome
      : 0;
    const interestIncome = r2(baseIntIncPart + customNonOp + customUnscoped);
    const financeCost = isConfigElementVisible(configMap, 'inc_financeCost')
      ? getCellOverride(cellOverrides, 'income', 'financeCost', input.year) ?? baseFinanceCost
      : 0;
    const profitBeforeTax =
      getCellOverride(cellOverrides, 'income', 'profitBeforeTax', input.year) ??
      r2(operatingProfit + interestIncome - financeCost);

    const ovTax = getCellOverride(cellOverrides, 'income', 'taxExpense', input.year);
    let taxExpense = 0;
    if (isConfigElementVisible(configMap, 'inc_taxExpense')) {
      if (ovTax !== undefined) {
        taxExpense = ovTax;
      } else if (Math.abs(profitBeforeTax - basePBT) > 0.01) {
        taxExpense = profitBeforeTax > 0 ? r(profitBeforeTax * (taxRate / 100)) : 0;
      } else {
        taxExpense = baseTaxExpense;
      }
    }

    const netProfit =
      getCellOverride(cellOverrides, 'income', 'netProfit', input.year) ?? r2(profitBeforeTax - taxExpense);

    return {
      year: input.year,
      isProjected: input.isProjected,
      revenue,
      costOfSales,
      grossProfit,
      operatingExpenses,
      depreciationExpense,
      operatingProfit,
      interestIncome,
      financeCost,
      profitBeforeTax,
      taxExpense,
      netProfit,
    };
  });
}

// 3. Calculate Statement of Cash Flows (Indirect Method - Section 76 Compliant)
export function calculateCashFlowStatement(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  options?: StatementCalculationOptions
): CashFlowStatementYear[] {
  const result: CashFlowStatementYear[] = [];
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const cellOverrides = options?.cellOverrides;

  // Track unfiltered base first-year Net Change in Cash so Year 0 opening cash stays anchored when elements are deselected
  let baseFirstNetChangeInCash = 0;
  const firstProjectedIdx = inputs.findIndex(y => y.isProjected);

  inputs.forEach((input, index) => {
    const isFirst = index === 0;
    const isFirstProjectedYear = firstProjectedIdx !== -1 && index === firstProjectedIdx;
    const prevInput = isFirst ? null : inputs[index - 1];
    const income = incomeStatements[index];
    const fixedAssets = fixedAssetsSchedule[index];

    if (isFirst) {
      const baseTaxRate = input.taxRate || 25;
      const hasCents =
        !Number.isInteger(input.salesRevenue) ||
        !Number.isInteger(input.financeCost) ||
        !Number.isInteger(input.netProfit) ||
        !Number.isInteger(input.repaymentOfBorrowings);
      const r = (n: number) => (hasCents ? r2(n) : Math.round(n));
      const baseTaxExp =
        input.netProfit >= 0
          ? r(input.netProfit / Math.max(0.01, 1 - baseTaxRate / 100) - input.netProfit)
          : 0;
      const basePBT = r2(input.netProfit + baseTaxExp);
      const firstGross = inputs[0]?.fixedAssetsGross ?? 0;
      const baseOpeningCost = Math.max(0, input.fixedAssetsGross - input.fixedAssetsAdditions + input.fixedAssetsDisposals);
      const baseDepCharge = Math.round(((baseOpeningCost + input.fixedAssetsGross) / 2) * (input.depreciationRate / 100));
      const baseDepDisposals = Math.round(input.fixedAssetsDisposals * computeDepDisposalsRatio(input.year, input.isProjected));
      const baseIntInc = r(input.salesRevenue * computeInterestIncomeRate(input.year, input.isProjected));
      const baseIntangAmort = input.intangibleAssets > 0 ? 100 : 0;
      const baseOpProfitBeforeWC =
        basePBT + baseDepCharge + baseIntangAmort + input.financeCost - baseIntInc - baseDepDisposals;
      const baseWC =
        -Math.round(input.closingInventory * 0.08) -
        Math.round(input.tradeReceivables * 0.08) -
        Math.round(input.otherReceivables * 0.05) -
        Math.round(input.otherCurrentAssets * 0.04) +
        Math.round(input.tradePayables * 0.07) +
        Math.round(input.otherPayables * 0.04);
      const baseOpCash = baseOpProfitBeforeWC + baseWC - baseTaxExp - input.financeCost;
      const baseInvCash = -input.fixedAssetsAdditions + input.fixedAssetsDisposals + baseIntInc;
      baseFirstNetChangeInCash = r2(baseOpCash + baseInvCash + input.cashFlowFromFinancing);
      void firstGross;
    }

    const vis = (id: string) => isConfigElementVisible(configMap, id);

    const profitBeforeTax = vis('cf_pbt')
      ? getCellOverride(cellOverrides, 'cf', 'profitBeforeTax', input.year) ?? income.profitBeforeTax
      : 0;

    // Depreciation add-back tracks P&L depreciationExpense so toggling off P&L depreciation doesn't double-count
    const depAddBackDefault = vis('inc_depreciationExpense') ? income.depreciationExpense : 0;
    const adjDepreciation = vis('cf_dep')
      ? getCellOverride(cellOverrides, 'cf', 'adjDepreciation', input.year) ?? depAddBackDefault
      : 0;

    const baseAmort = prevInput
      ? Math.max(0, prevInput.intangibleAssets - input.intangibleAssets)
      : (input.intangibleAssets > 0 ? 100 : 0);
    const adjAmortization =
      vis('cf_amort') && vis('bs_intangibleAssets')
        ? getCellOverride(cellOverrides, 'cf', 'adjAmortization', input.year) ?? baseAmort
        : 0;

    const adjFinanceCost =
      vis('cf_finance') && vis('inc_financeCost')
        ? getCellOverride(cellOverrides, 'cf', 'adjFinanceCost', input.year) ?? income.financeCost
        : 0;

    const adjInterestIncome =
      vis('cf_interest_inc') && vis('inc_interestIncome')
        ? getCellOverride(cellOverrides, 'cf', 'adjInterestIncome', input.year) ?? -income.interestIncome
        : 0;

    const adjLossGainDisposal = vis('cf_disposal_gain_loss')
      ? getCellOverride(cellOverrides, 'cf', 'adjLossGainDisposal', input.year) ?? -fixedAssets.totalDepDisposals
      : 0;

    const operatingProfitBeforeWC =
      getCellOverride(cellOverrides, 'cf', 'operatingProfitBeforeWC', input.year) ??
      r2(
        profitBeforeTax + adjDepreciation + adjAmortization + adjFinanceCost + adjInterestIncome + adjLossGainDisposal
      );

    const baseInvChg = prevInput
      ? -(input.closingInventory - prevInput.closingInventory)
      : -Math.round(input.closingInventory * 0.08);
    const changeInInventory =
      vis('cf_wc_inventory') && vis('bs_closingInventory')
        ? getCellOverride(cellOverrides, 'cf', 'changeInInventory', input.year) ?? baseInvChg
        : 0;

    const baseRecChg = prevInput
      ? -(input.tradeReceivables - prevInput.tradeReceivables)
      : -Math.round(input.tradeReceivables * 0.08);
    const changeInReceivables =
      vis('cf_wc_receivables') && vis('bs_tradeReceivables')
        ? getCellOverride(cellOverrides, 'cf', 'changeInReceivables', input.year) ?? baseRecChg
        : 0;

    const baseOtherRecChg = prevInput
      ? -(input.otherReceivables - prevInput.otherReceivables)
      : -Math.round(input.otherReceivables * 0.05);
    const changeInOtherReceivables =
      vis('cf_wc_other_rec') && vis('bs_otherReceivables')
        ? getCellOverride(cellOverrides, 'cf', 'changeInOtherReceivables', input.year) ?? baseOtherRecChg
        : 0;

    const baseOtherCAChg = prevInput
      ? -(input.otherCurrentAssets - prevInput.otherCurrentAssets)
      : -Math.round(input.otherCurrentAssets * 0.04);
    const changeInOtherCA =
      vis('cf_wc_other_ca') && vis('bs_otherCurrentAssets')
        ? getCellOverride(cellOverrides, 'cf', 'changeInOtherCA', input.year) ?? baseOtherCAChg
        : 0;

    const basePayChg = prevInput
      ? input.tradePayables - prevInput.tradePayables
      : Math.round(input.tradePayables * 0.07);
    const changeInPayables =
      vis('cf_wc_payables') && vis('bs_tradePayables')
        ? getCellOverride(cellOverrides, 'cf', 'changeInPayables', input.year) ?? basePayChg
        : 0;

    const baseOtherPayChg = prevInput
      ? input.otherPayables - prevInput.otherPayables
      : Math.round(input.otherPayables * 0.04);
    const changeInOtherPayables =
      vis('cf_wc_other_pay') && vis('bs_otherPayables')
        ? getCellOverride(cellOverrides, 'cf', 'changeInOtherPayables', input.year) ?? baseOtherPayChg
        : 0;

    const currTaxPayable = vis('bs_currentTaxPayable')
      ? getCellOverride(cellOverrides, 'bs', 'currentTaxPayable', input.year) ??
        Math.round(Math.max(0, income.taxExpense) * 0.45)
      : 0;
    const prevTaxPayable = prevInput
      ? vis('bs_currentTaxPayable')
        ? getCellOverride(cellOverrides, 'bs', 'currentTaxPayable', prevInput.year) ??
          Math.round(Math.max(0, incomeStatements[index - 1].taxExpense) * 0.45)
        : 0
      : currTaxPayable;
    const changeInTaxPayable = currTaxPayable - prevTaxPayable;

    const customOperating = getVisibleCustomItemsSum(
      options?.statementConfig,
      'cashFlow',
      input.year,
      'Operating Activities'
    );

    const cashGeneratedFromOperations =
      getCellOverride(cellOverrides, 'cf', 'cashGeneratedFromOperations', input.year) ??
      r2(
        operatingProfitBeforeWC +
          changeInInventory +
          changeInReceivables +
          changeInOtherReceivables +
          changeInOtherCA +
          changeInPayables +
          changeInOtherPayables +
          customOperating
      );

    const incomeTaxPaid = vis('cf_tax_paid')
      ? getCellOverride(cellOverrides, 'cf', 'incomeTaxPaid', input.year) ?? (-income.taxExpense + changeInTaxPayable)
      : 0;
    const financeCostPaid =
      vis('cf_finance_paid') && vis('inc_financeCost')
        ? getCellOverride(cellOverrides, 'cf', 'financeCostPaid', input.year) ?? -income.financeCost
        : 0;

    const netCashFromOperating =
      getCellOverride(cellOverrides, 'cf', 'netCashFromOperating', input.year) ??
      r2(cashGeneratedFromOperations + incomeTaxPaid + financeCostPaid);

    // B. Investing Activities
    const purchasePPE = vis('cf_pur_ppe')
      ? getCellOverride(cellOverrides, 'cf', 'purchasePPE', input.year) ?? -fixedAssets.totalAdditions
      : 0;
    const salePPE = vis('cf_sale_ppe')
      ? getCellOverride(cellOverrides, 'cf', 'salePPE', input.year) ?? fixedAssets.totalDisposals
      : 0;
    const basePurInt = prevInput ? -Math.max(0, input.intangibleAssets - prevInput.intangibleAssets) : 0;
    const purchaseIntangibles =
      vis('cf_pur_intangible') && vis('bs_intangibleAssets')
        ? getCellOverride(cellOverrides, 'cf', 'purchaseIntangibles', input.year) ?? basePurInt
        : 0;
    const interestReceived =
      vis('cf_interest_rec') && vis('inc_interestIncome')
        ? getCellOverride(cellOverrides, 'cf', 'interestReceived', input.year) ?? income.interestIncome
        : 0;
    const dividendsReceived = vis('cf_dividends_rec')
      ? getCellOverride(cellOverrides, 'cf', 'dividendsReceived', input.year) ?? 0
      : 0;

    const customInvesting = getVisibleCustomItemsSum(
      options?.statementConfig,
      'cashFlow',
      input.year,
      'Investing Activities'
    );

    const netCashFromInvesting =
      getCellOverride(cellOverrides, 'cf', 'netCashFromInvesting', input.year) ??
      r2(purchasePPE + salePPE + purchaseIntangibles + interestReceived + dividendsReceived + customInvesting);

    // C. Financing Activities
    const rawBorrowProc = input.financingProceedsBorrowings || 0;
    const staleFirstProjBorrowProc = isFirstProjectedYear ? rawBorrowProc : 0;
    const rawDirectFinancing =
      getCellOverride(cellOverrides, 'cf', 'directlyEnteredFinancingCashFlow', input.year) ??
      r2(input.cashFlowFromFinancing - staleFirstProjBorrowProc);
    const directlyEnteredFinancingCashFlow = vis('cf_direct_financing') ? rawDirectFinancing : 0;

    const fullShares = input.financingProceedsShares || 0;
    const fullBorrowProc = isFirstProjectedYear ? 0 : rawBorrowProc;
    const fullRepayment =
      input.repaymentOfBorrowings !== undefined
        ? input.repaymentOfBorrowings
        : input.financingRepaymentBorrowings
        ? Math.abs(input.financingRepaymentBorrowings)
        : 0;
    const fullDivs = -input.dividendsPaid;
    const fullBreakdownSum = fullShares + fullBorrowProc - fullRepayment + fullDivs;

    const financingProceedsShares = vis('cf_shares_issue')
      ? getCellOverride(cellOverrides, 'cf', 'financingProceedsShares', input.year) ?? fullShares
      : 0;
    const financingProceedsBorrowings = isFirstProjectedYear
      ? 0
      : vis('cf_borrow_proceeds')
      ? getCellOverride(cellOverrides, 'cf', 'financingProceedsBorrowings', input.year) ?? fullBorrowProc
      : 0;
    const ovRepayPos = getCellOverride(cellOverrides, 'cf', 'repaymentOfBorrowings', input.year);
    const ovRepayNeg = getCellOverride(cellOverrides, 'cf', 'financingRepaymentBorrowings', input.year);
    const repaymentOfBorrowings = vis('cf_borrow_repay')
      ? ovRepayPos !== undefined
        ? ovRepayPos
        : ovRepayNeg !== undefined
        ? Math.abs(ovRepayNeg)
        : fullRepayment
      : 0;
    const financingRepaymentBorrowings = vis('cf_borrow_repay')
      ? ovRepayNeg !== undefined
        ? ovRepayNeg
        : -repaymentOfBorrowings
      : 0;
    const financingDividendsPaid = vis('cf_dividends_paid')
      ? getCellOverride(cellOverrides, 'cf', 'financingDividendsPaid', input.year) ?? fullDivs
      : 0;

    const customFinancing = getVisibleCustomItemsSum(
      options?.statementConfig,
      'cashFlow',
      input.year,
      'Financing Activities'
    );

    const activeBreakdownSum =
      financingProceedsShares +
      financingProceedsBorrowings +
      financingRepaymentBorrowings +
      financingDividendsPaid +
      customFinancing;
    const isFinancingBreakdownComplete = Math.abs(fullBreakdownSum - rawDirectFinancing) < 10;

    const allFinBreakdownDefault =
      vis('cf_shares_issue') &&
      vis('cf_borrow_proceeds') &&
      vis('cf_borrow_repay') &&
      vis('cf_dividends_paid') &&
      ovRepayPos === undefined &&
      ovRepayNeg === undefined &&
      getCellOverride(cellOverrides, 'cf', 'financingProceedsShares', input.year) === undefined &&
      (isFirstProjectedYear ||
        getCellOverride(cellOverrides, 'cf', 'financingProceedsBorrowings', input.year) === undefined) &&
      getCellOverride(cellOverrides, 'cf', 'financingDividendsPaid', input.year) === undefined &&
      customFinancing === 0;

    const computedNetCashFromFinancing = vis('cf_direct_financing')
      ? allFinBreakdownDefault
        ? directlyEnteredFinancingCashFlow
        : r2(directlyEnteredFinancingCashFlow + (activeBreakdownSum - fullBreakdownSum))
      : r2(activeBreakdownSum);

    const netCashFromFinancing =
      getCellOverride(cellOverrides, 'cf', 'netCashFromFinancing', input.year) ??
      computedNetCashFromFinancing;

    // D. Net Increase / Decrease in Cash and Cash Equivalents
    const netChangeInCash =
      getCellOverride(cellOverrides, 'cf', 'netChangeInCash', input.year) ??
      r2(netCashFromOperating + netCashFromInvesting + netCashFromFinancing);

    const ovOpeningCash = getCellOverride(cellOverrides, 'cf', 'openingCash', input.year);
    let openingCash = ovOpeningCash !== undefined
      ? ovOpeningCash
      : isFirst
      ? r2(input.cashAndEquivalentsBS - baseFirstNetChangeInCash)
      : result[index - 1].closingCashCalculated;
    if (ovOpeningCash === undefined && openingCash < 0 && isFirst) {
      openingCash =
        input.salesRevenue > 0 || baseFirstNetChangeInCash !== 0
          ? Math.max(200, Math.round(input.salesRevenue * 0.05))
          : 0;
    }

    const closingCashCalculated =
      getCellOverride(cellOverrides, 'cf', 'closingCashCalculated', input.year) ??
      r2(openingCash + netChangeInCash);

    const ovCFBSCash = getCellOverride(cellOverrides, 'cf', 'balanceSheetCash', input.year);
    const ovBSCash = getCellOverride(cellOverrides, 'bs', 'cashAndEquivalents', input.year);
    const balanceSheetCash = ovCFBSCash !== undefined
      ? ovCFBSCash
      : !vis('bs_cashAndEquivalents')
      ? 0
      : ovBSCash !== undefined
      ? ovBSCash
      : isFirst && input.cashAndEquivalentsBS !== undefined && !input.isProjected
      ? Math.abs(closingCashCalculated - input.cashAndEquivalentsBS) < 5 &&
        Math.abs(netChangeInCash - baseFirstNetChangeInCash) < 0.01
        ? closingCashCalculated
        : input.cashAndEquivalentsBS
      : closingCashCalculated;

    const reconciliationDifference = r2(closingCashCalculated - balanceSheetCash);
    const isReconciled = Math.abs(reconciliationDifference) < 0.01;

    const reconciliationExplanation = isReconciled
      ? 'Fully reconciled with Balance Sheet Cash & Cash Equivalents.'
      : `Variance of ${reconciliationDifference > 0 ? '+' : ''}${reconciliationDifference.toLocaleString()} identified between Cash Flow ending cash and Balance Sheet stated cash. The system preserves raw user entries without fabricating offsetting transactions.`;

    result.push({
      year: input.year,
      isProjected: input.isProjected,
      profitBeforeTax,
      adjDepreciation,
      adjAmortization,
      adjFinanceCost,
      adjInterestIncome,
      adjLossGainDisposal,
      operatingProfitBeforeWC,
      changeInInventory,
      changeInReceivables,
      changeInOtherReceivables,
      changeInOtherCA,
      changeInPayables,
      changeInOtherPayables,
      cashGeneratedFromOperations,
      incomeTaxPaid,
      financeCostPaid,
      netCashFromOperating,
      purchasePPE,
      salePPE,
      purchaseIntangibles,
      interestReceived,
      dividendsReceived,
      netCashFromInvesting,
      directlyEnteredFinancingCashFlow,
      financingProceedsShares,
      financingProceedsBorrowings,
      financingRepaymentBorrowings,
      repaymentOfBorrowings,
      financingDividendsPaid,
      netCashFromFinancing,
      isFinancingBreakdownComplete,
      netChangeInCash,
      openingCash,
      closingCashCalculated,
      balanceSheetCash,
      reconciliationDifference,
      reconciliationStatus: isReconciled ? 'RECONCILED' : 'VARIANCE_IDENTIFIED',
      reconciliationExplanation,
    });
  });

  return result;
}

// 4. Calculate Statement of Changes in Equity
export function calculateEquityStatement(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  options?: StatementCalculationOptions
): EquityStatementYear[] {
  const result: EquityStatementYear[] = [];
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const cellOverrides = options?.cellOverrides;
  const vis = (id: string) => isConfigElementVisible(configMap, id);

  inputs.forEach((input, index) => {
    const isFirst = index === 0;
    const income = incomeStatements[index];

    const openingShareCapital = vis('eq_openingShareCapital')
      ? getCellOverride(cellOverrides, 'eq', 'openingShareCapital', input.year) ??
        (isFirst ? input.shareCapital : result[index - 1].closingShareCapital)
      : 0;
    const shareCapitalIssued = vis('eq_shareCapitalIssued')
      ? getCellOverride(cellOverrides, 'eq', 'shareCapitalIssued', input.year) ??
        (input.financingProceedsShares || 0)
      : 0;
    const closingShareCapital = r2(openingShareCapital + shareCapitalIssued);

    const isZeroFirst =
      isFirst &&
      input.salesRevenue === 0 &&
      input.shareCapital === 0 &&
      input.netProfit === 0 &&
      input.longTermBorrowings === 0;

    const openingRetainedEarnings = vis('eq_openingRetainedEarnings')
      ? getCellOverride(cellOverrides, 'eq', 'openingRetainedEarnings', input.year) ??
        (isFirst ? input.retainedEarningsOpening ?? (isZeroFirst ? 0 : 6500) : result[index - 1].closingRetainedEarnings)
      : 0;
    const netProfitForYear = vis('eq_netProfitForYear')
      ? getCellOverride(cellOverrides, 'eq', 'netProfitForYear', input.year) ?? income.netProfit
      : 0;
    const dividendsPaid = vis('eq_dividendsPaid')
      ? getCellOverride(cellOverrides, 'eq', 'dividendsPaid', input.year) ?? input.dividendsPaid
      : 0;
    const closingRetainedEarnings = r2(openingRetainedEarnings + netProfitForYear - dividendsPaid);

    const openingReserves = vis('eq_openingReserves')
      ? getCellOverride(cellOverrides, 'eq', 'openingReserves', input.year) ??
        (isFirst ? (input.reservesOpening ?? (isZeroFirst ? 0 : 1800)) : result[index - 1].closingReserves)
      : 0;
    const otherComprehensiveIncome = vis('eq_otherComprehensiveIncome')
      ? getCellOverride(cellOverrides, 'eq', 'otherComprehensiveIncome', input.year) ?? 0
      : 0;
    const closingReserves = r2(openingReserves + otherComprehensiveIncome);

    const customEquity = getVisibleCustomItemsSum(options?.statementConfig, 'equity', input.year);
    const totalOpeningEquity = r2(openingShareCapital + openingRetainedEarnings + openingReserves);
    const totalClosingEquity = r2(
      closingShareCapital + closingRetainedEarnings + closingReserves + customEquity
    );

    result.push({
      year: input.year,
      isProjected: input.isProjected,
      openingShareCapital,
      shareCapitalIssued,
      closingShareCapital,
      openingRetainedEarnings,
      netProfitForYear,
      dividendsPaid,
      closingRetainedEarnings,
      openingReserves,
      otherComprehensiveIncome,
      closingReserves,
      totalOpeningEquity,
      totalClosingEquity,
    });
  });

  return result;
}

// 5. Calculate Statement of Financial Position (Balance Sheet)
export function calculateBalanceSheet(
  inputs: YearDataInput[],
  equityStatements: EquityStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  cashFlowStatements?: CashFlowStatementYear[],
  options?: StatementCalculationOptions,
  incomeStatements?: IncomeStatementYear[]
): BalanceSheetYear[] {
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const cellOverrides = options?.cellOverrides;
  const vis = (id: string) => isConfigElementVisible(configMap, id);

  return inputs.map((input, index) => {
    const equity = equityStatements[index];
    const fixedAssets = fixedAssetsSchedule[index];
    const income = incomeStatements?.[index];

    const hasCents = !Number.isInteger(input.financeCost) || !Number.isInteger(input.netProfit);
    const rawBaseTax = Math.max(
      0,
      input.netProfit / Math.max(0.01, 1 - (input.taxRate || 25) / 100) - input.netProfit
    );
    const baseTaxPayable = hasCents
      ? r2(rawBaseTax * 0.45)
      : Math.round(Math.round(rawBaseTax) * 0.45);
    const activeTaxPayable =
      income && (!vis('inc_taxExpense') || getCellOverride(cellOverrides, 'income', 'taxExpense', input.year) !== undefined || Math.abs(income.taxExpense - rawBaseTax) > 1)
        ? Math.round(Math.max(0, income.taxExpense) * 0.45)
        : baseTaxPayable;

    // Baseline plug for otherNonCurrentAssets computed BEFORE applying bs_* deselections/overrides
    const baseCash =
      cashFlowStatements && cashFlowStatements[index]
        ? cashFlowStatements[index].closingCashCalculated
        : input.cashAndEquivalentsBS;
    const baseCurrentAssets = r2(
      baseCash +
        input.tradeReceivables +
        input.otherReceivables +
        input.closingInventory +
        input.otherCurrentAssets
    );
    const baseCurrentLiabilities = r2(
      input.tradePayables + input.otherPayables + input.shortTermBorrowings + activeTaxPayable
    );
    const baseLiabilitiesAndEquity = r2(
      baseCurrentLiabilities +
        input.longTermBorrowings +
        (equity.closingShareCapital + equity.closingRetainedEarnings + equity.closingReserves)
    );
    const rawNonCurrentGap = r2(
      baseLiabilitiesAndEquity - (baseCurrentAssets + fixedAssets.totalNetBookValue + input.intangibleAssets)
    );
    const baseOtherNonCurrentAssets = rawNonCurrentGap >= 0 ? rawNonCurrentGap : 0;
    const baseOtherNonCurrentLiab = rawNonCurrentGap < 0 ? Math.abs(rawNonCurrentGap) : 0;

    // Apply bs_* visibility, inline cell overrides, and custom items
    const cashAndEquivalents = vis('bs_cashAndEquivalents')
      ? getCellOverride(cellOverrides, 'bs', 'cashAndEquivalents', input.year) ?? baseCash
      : 0;
    const tradeReceivables = vis('bs_tradeReceivables')
      ? getCellOverride(cellOverrides, 'bs', 'tradeReceivables', input.year) ?? input.tradeReceivables
      : 0;
    const otherReceivables = vis('bs_otherReceivables')
      ? getCellOverride(cellOverrides, 'bs', 'otherReceivables', input.year) ?? input.otherReceivables
      : 0;
    const closingInventory = vis('bs_closingInventory')
      ? getCellOverride(cellOverrides, 'bs', 'closingInventory', input.year) ?? input.closingInventory
      : 0;
    const otherCurrentAssets = vis('bs_otherCurrentAssets')
      ? getCellOverride(cellOverrides, 'bs', 'otherCurrentAssets', input.year) ?? input.otherCurrentAssets
      : 0;
    const customCurrentAssets = getVisibleCustomItemsSum(
      options?.statementConfig,
      'balanceSheet',
      input.year,
      'Current Assets'
    );
    const totalCurrentAssets = r2(
      cashAndEquivalents +
        tradeReceivables +
        otherReceivables +
        closingInventory +
        otherCurrentAssets +
        customCurrentAssets
    );

    const propertyPlantEquipment = vis('bs_propertyPlantEquipment')
      ? getCellOverride(cellOverrides, 'bs', 'propertyPlantEquipment', input.year) ?? fixedAssets.totalNetBookValue
      : 0;
    const intangibleAssets = vis('bs_intangibleAssets')
      ? getCellOverride(cellOverrides, 'bs', 'intangibleAssets', input.year) ?? input.intangibleAssets
      : 0;
    const otherNonCurrentAssets = vis('bs_otherNonCurrentAssets')
      ? getCellOverride(cellOverrides, 'bs', 'otherNonCurrentAssets', input.year) ?? baseOtherNonCurrentAssets
      : 0;
    const customNonCurrentAssets = getVisibleCustomItemsSum(
      options?.statementConfig,
      'balanceSheet',
      input.year,
      'Non-Current Assets'
    );
    const totalNonCurrentAssets = r2(
      propertyPlantEquipment + intangibleAssets + otherNonCurrentAssets + customNonCurrentAssets
    );
    const totalAssets = r2(totalCurrentAssets + totalNonCurrentAssets);

    const tradePayables = vis('bs_tradePayables')
      ? getCellOverride(cellOverrides, 'bs', 'tradePayables', input.year) ?? input.tradePayables
      : 0;
    const otherPayables = vis('bs_otherPayables')
      ? getCellOverride(cellOverrides, 'bs', 'otherPayables', input.year) ?? input.otherPayables
      : 0;
    const shortTermBorrowings = vis('bs_shortTermBorrowings')
      ? getCellOverride(cellOverrides, 'bs', 'shortTermBorrowings', input.year) ?? input.shortTermBorrowings
      : 0;
    const currentTaxPayable = vis('bs_currentTaxPayable')
      ? getCellOverride(cellOverrides, 'bs', 'currentTaxPayable', input.year) ?? activeTaxPayable
      : 0;
    const customCurrentLiab = getVisibleCustomItemsSum(
      options?.statementConfig,
      'balanceSheet',
      input.year,
      'Current Liabilities'
    );
    const totalCurrentLiabilities = r2(
      tradePayables + otherPayables + shortTermBorrowings + currentTaxPayable + customCurrentLiab
    );

    const longTermBorrowings = vis('bs_longTermBorrowings')
      ? getCellOverride(cellOverrides, 'bs', 'longTermBorrowings', input.year) ?? input.longTermBorrowings
      : 0;
    const otherNonCurrentLiabilities = vis('bs_otherNonCurrentLiabilities')
      ? getCellOverride(cellOverrides, 'bs', 'otherNonCurrentLiabilities', input.year) ?? baseOtherNonCurrentLiab
      : 0;
    const customNonCurrentLiab = getVisibleCustomItemsSum(
      options?.statementConfig,
      'balanceSheet',
      input.year,
      'Non-Current Liabilities'
    );
    const totalNonCurrentLiabilities = r2(
      longTermBorrowings + otherNonCurrentLiabilities + customNonCurrentLiab
    );

    const totalLiabilities = r2(totalCurrentLiabilities + totalNonCurrentLiabilities);

    const shareCapital = vis('bs_shareCapital')
      ? getCellOverride(cellOverrides, 'bs', 'shareCapital', input.year) ?? equity.closingShareCapital
      : 0;
    const retainedEarnings = vis('bs_retainedEarnings')
      ? getCellOverride(cellOverrides, 'bs', 'retainedEarnings', input.year) ?? equity.closingRetainedEarnings
      : 0;
    const reserves = vis('bs_reserves')
      ? getCellOverride(cellOverrides, 'bs', 'reserves', input.year) ?? equity.closingReserves
      : 0;
    const customEquityBS = getVisibleCustomItemsSum(
      options?.statementConfig,
      'balanceSheet',
      input.year,
      'Equity'
    );
    const totalEquity = r2(shareCapital + retainedEarnings + reserves + customEquityBS);

    const totalLiabilitiesAndEquity = r2(totalLiabilities + totalEquity);
    const balanceCheck = r2(totalAssets - totalLiabilitiesAndEquity);

    return {
      year: input.year,
      isProjected: input.isProjected,
      cashAndEquivalents,
      tradeReceivables,
      otherReceivables,
      closingInventory,
      otherCurrentAssets,
      totalCurrentAssets,
      propertyPlantEquipment,
      intangibleAssets,
      otherNonCurrentAssets,
      totalNonCurrentAssets,
      totalAssets,
      tradePayables,
      otherPayables,
      shortTermBorrowings,
      currentTaxPayable,
      totalCurrentLiabilities,
      longTermBorrowings,
      otherNonCurrentLiabilities,
      totalNonCurrentLiabilities,
      totalLiabilities,
      shareCapital,
      retainedEarnings,
      reserves,
      totalEquity,
      totalLiabilitiesAndEquity,
      balanceCheck,
    };
  });
}

// 6. Generate Statutory Notes to the Financial Statements (Notes 1 to 12)
export function generateNotesToFinancialStatements(
  company: CompanyProfile,
  years: number[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  options?: {
    statementConfig?: StatementConfigItem[];
    labelOverrides?: StatementLabelOverrides;
    cellOverrides?: StatementCellOverrides;
  }
): FinancialNoteItem[] {
  if (years.length === 0) return [];
  const latestIdx = years.length - 1;
  const currYear = years[latestIdx];
  const prevYear = years.length > 1 ? years[latestIdx - 1] : currYear;
  const cIS = incomeStatements[latestIdx];
  const pIS = incomeStatements[latestIdx > 0 ? latestIdx - 1 : 0];
  const cBS = balanceSheets[latestIdx];
  const pBS = balanceSheets[latestIdx > 0 ? latestIdx - 1 : 0];
  const cCF = cashFlowStatements[latestIdx];
  const pCF = cashFlowStatements[latestIdx > 0 ? latestIdx - 1 : 0];
  const cFA = fixedAssetsSchedule[latestIdx];

  const cfg = options?.statementConfig;
  const lblOv = options?.labelOverrides;
  const cellOv = options?.cellOverrides;
  const configMap = buildConfigVisibilityMap(cfg);
  const vis = (id: string) => isConfigElementVisible(configMap, id);
  const nLbl = (stmt: string, field: string, def: string) =>
    resolveEffectiveLabel(stmt, field, def, lblOv, cfg);

  const noteVal = (
    stmt: string,
    field: string,
    year: number,
    fallback: number,
    cfgId?: string,
    customSum = 0
  ) => {
    if (cfgId && !vis(cfgId)) return r2(customSum);
    const ov = getCellOverride(cellOv, stmt, field, year);
    return ov !== undefined ? r2(ov + customSum) : fallback;
  };

  // Resolve Note 5 PPE categories using visibility + label overrides + custom fixed asset classes
  const stdFACats = (cFA.allCategories || cFA.categories.slice(0, FIXED_ASSET_CATEGORY_CONFIG_IDS.length))
    .map((cat, idx) => {
      const cfgId = FIXED_ASSET_CATEGORY_CONFIG_IDS[idx];
      return {
        cat: { ...cat, name: nLbl('fa', String(idx), cat.name) },
        visible: cfgId ? vis(cfgId) : true,
      };
    })
    .filter(e => e.visible)
    .map(e => e.cat);
  const customFACats: FixedAssetCategory[] = (cfg || [])
    .filter(item => item.isCustom && item.isVisible && item.statement === 'fixedAssets')
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
    .map(item => {
      const val = getCustomItemYearValue(item, currYear);
      return {
        name: item.label,
        costOpening: val,
        additions: 0,
        disposals: 0,
        costClosing: val,
        depOpening: 0,
        depCharge: 0,
        depDisposals: 0,
        depClosing: 0,
        netBookValue: val,
        depRate: 0,
      };
    });
  const note5Categories = [...stdFACats, ...customFACats];

  const customEquityRows = (cfg || [])
    .filter(item => item.isCustom && item.isVisible && item.statement === 'balanceSheet' && item.section === 'Equity')
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
    .map(item => ({
      label: item.label,
      values: [
        getCustomItemYearValue(item, currYear).toLocaleString(),
        getCustomItemYearValue(item, prevYear).toLocaleString(),
      ],
    }));

  const cCustomRev = getVisibleCustomItemsSum(cfg, 'income', currYear, 'Revenue & Gross Margin');
  const pCustomRev = getVisibleCustomItemsSum(cfg, 'income', prevYear, 'Revenue & Gross Margin');
  const cCustomEqBs = getVisibleCustomItemsSum(cfg, 'balanceSheet', currYear, 'Equity');
  const pCustomEqBs = getVisibleCustomItemsSum(cfg, 'balanceSheet', prevYear, 'Equity');

  const cRevenue = noteVal('income', 'revenue', currYear, cIS.revenue, 'inc_revenue', cCustomRev);
  const pRevenue = noteVal('income', 'revenue', prevYear, pIS.revenue, 'inc_revenue', pCustomRev);
  const cCogs = noteVal('income', 'costOfSales', currYear, cIS.costOfSales, 'inc_costOfSales');
  const pCogs = noteVal('income', 'costOfSales', prevYear, pIS.costOfSales, 'inc_costOfSales');
  const cFinCost = noteVal('income', 'financeCost', currYear, cIS.financeCost, 'inc_financeCost');
  const pFinCost = noteVal('income', 'financeCost', prevYear, pIS.financeCost, 'inc_financeCost');

  const cInv = noteVal('bs', 'closingInventory', currYear, cBS.closingInventory, 'bs_closingInventory');
  const pInv = noteVal('bs', 'closingInventory', prevYear, pBS.closingInventory, 'bs_closingInventory');
  const cTrRec = noteVal('bs', 'tradeReceivables', currYear, cBS.tradeReceivables, 'bs_tradeReceivables');
  const pTrRec = noteVal('bs', 'tradeReceivables', prevYear, pBS.tradeReceivables, 'bs_tradeReceivables');
  const cOthRec = noteVal('bs', 'otherReceivables', currYear, cBS.otherReceivables, 'bs_otherReceivables');
  const pOthRec = noteVal('bs', 'otherReceivables', prevYear, pBS.otherReceivables, 'bs_otherReceivables');
  const cCashBS = noteVal('bs', 'cashAndEquivalents', currYear, cBS.cashAndEquivalents, 'bs_cashAndEquivalents');
  const pCashBS = noteVal('bs', 'cashAndEquivalents', prevYear, pBS.cashAndEquivalents, 'bs_cashAndEquivalents');
  const cSC = noteVal('bs', 'shareCapital', currYear, cBS.shareCapital, 'bs_shareCapital');
  const pSC = noteVal('bs', 'shareCapital', prevYear, pBS.shareCapital, 'bs_shareCapital');
  const cRE = noteVal('bs', 'retainedEarnings', currYear, cBS.retainedEarnings, 'bs_retainedEarnings');
  const pRE = noteVal('bs', 'retainedEarnings', prevYear, pBS.retainedEarnings, 'bs_retainedEarnings');
  const cRes = noteVal('bs', 'reserves', currYear, cBS.reserves, 'bs_reserves');
  const pRes = noteVal('bs', 'reserves', prevYear, pBS.reserves, 'bs_reserves');
  const cTotEq = noteVal('bs', 'totalEquity', currYear, cBS.totalEquity, undefined, 0);
  const pTotEq = noteVal('bs', 'totalEquity', prevYear, pBS.totalEquity, undefined, 0);
  void cCustomEqBs;
  void pCustomEqBs;
  const cLTB = noteVal('bs', 'longTermBorrowings', currYear, cBS.longTermBorrowings, 'bs_longTermBorrowings');
  const pLTB = noteVal('bs', 'longTermBorrowings', prevYear, pBS.longTermBorrowings, 'bs_longTermBorrowings');
  const cSTB = noteVal('bs', 'shortTermBorrowings', currYear, cBS.shortTermBorrowings, 'bs_shortTermBorrowings');
  const pSTB = noteVal('bs', 'shortTermBorrowings', prevYear, pBS.shortTermBorrowings, 'bs_shortTermBorrowings');
  const cTrPay = noteVal('bs', 'tradePayables', currYear, cBS.tradePayables, 'bs_tradePayables');
  const pTrPay = noteVal('bs', 'tradePayables', prevYear, pBS.tradePayables, 'bs_tradePayables');
  const cOthPay = noteVal('bs', 'otherPayables', currYear, cBS.otherPayables, 'bs_otherPayables');
  const pOthPay = noteVal('bs', 'otherPayables', prevYear, pBS.otherPayables, 'bs_otherPayables');

  const cCfOp = noteVal('cf', 'netCashFromOperating', currYear, cCF.netCashFromOperating);
  const pCfOp = noteVal('cf', 'netCashFromOperating', prevYear, pCF.netCashFromOperating);
  const cCfInv = noteVal('cf', 'netCashFromInvesting', currYear, cCF.netCashFromInvesting);
  const pCfInv = noteVal('cf', 'netCashFromInvesting', prevYear, pCF.netCashFromInvesting);
  const cCfFinDirect = noteVal('cf', 'directlyEnteredFinancingCashFlow', currYear, cCF.directlyEnteredFinancingCashFlow, 'cf_direct_financing');
  const pCfFinDirect = noteVal('cf', 'directlyEnteredFinancingCashFlow', prevYear, pCF.directlyEnteredFinancingCashFlow, 'cf_direct_financing');
  const cCfChange = noteVal('cf', 'netChangeInCash', currYear, cCF.netChangeInCash);
  const pCfChange = noteVal('cf', 'netChangeInCash', prevYear, pCF.netChangeInCash);
  const cCfOpen = noteVal('cf', 'openingCash', currYear, cCF.openingCash);
  const pCfOpen = noteVal('cf', 'openingCash', prevYear, pCF.openingCash);
  const cCfClose = noteVal('cf', 'closingCashCalculated', currYear, cCF.closingCashCalculated);
  const pCfClose = noteVal('cf', 'closingCashCalculated', prevYear, pCF.closingCashCalculated);
  const cCfStatedBS = noteVal('cf', 'balanceSheetCash', currYear, cCF.balanceSheetCash);
  const pCfStatedBS = noteVal('cf', 'balanceSheetCash', prevYear, pCF.balanceSheetCash);
  const cCfDiff = r2(cCfClose - cCfStatedBS);
  const pCfDiff = r2(pCfClose - pCfStatedBS);
  const cCfReconciled = Math.abs(cCfDiff) < 0.01;

  const pRevSplit = splitTotalNonUniform(pRevenue, prevYear, 70, [0.72, 0.20, 0.08], [0.032, 0.026, 0.016]);
  const cRevSplit = splitTotalNonUniform(cRevenue, currYear, 70, [0.72, 0.20, 0.08], [0.032, 0.026, 0.016], pRevSplit, 1.35);
  const pCogsSplit = splitTotalNonUniform(pCogs, prevYear, 80, [0.58, 0.22, 0.20], [0.028, 0.022, 0.020]);
  const cCogsSplit = splitTotalNonUniform(cCogs, currYear, 80, [0.58, 0.22, 0.20], [0.028, 0.022, 0.020], pCogsSplit, 1.35);
  const pFinSplit = splitTotalNonUniform(pFinCost, prevYear, 90, [0.65, 0.25, 0.10], [0.035, 0.026, 0.018]);
  const cFinSplit = splitTotalNonUniform(cFinCost, currYear, 90, [0.65, 0.25, 0.10], [0.035, 0.026, 0.018], pFinSplit, 1.35);
  const pInvSplit = splitTotalNonUniform(pInv, prevYear, 100, [0.45, 0.20, 0.35], [0.032, 0.024, 0.028]);
  const cInvSplit = splitTotalNonUniform(cInv, currYear, 100, [0.45, 0.20, 0.35], [0.032, 0.024, 0.028], pInvSplit, 1.35);
  const pRecSplit = splitTotalNonUniform(pTrRec, prevYear, 110, [0.65, 0.25, 0.10], [0.036, 0.026, 0.018]);
  const cRecSplit = splitTotalNonUniform(cTrRec, currYear, 110, [0.65, 0.25, 0.10], [0.036, 0.026, 0.018], pRecSplit, 1.35);
  const pCashSplit = splitTotalNonUniform(pCashBS, prevYear, 120, [0.06, 0.69, 0.25], [0.018, 0.036, 0.030]);
  const cCashSplit = splitTotalNonUniform(cCashBS, currYear, 120, [0.06, 0.69, 0.25], [0.018, 0.036, 0.030], pCashSplit, 1.35);

  return [
    {
      noteNumber: 1,
      title: 'Company Information & Significant Accounting Policies',
      category: 'General Disclosures',
      textDescription: `${company.companyName?.trim() || 'The Company'} is a limited liability company incorporated under applicable commercial law, with its registered corporate office situated at ${company.registeredAddress?.trim() || 'its registered corporate office'}. The financial statements are prepared on a consistent annual cycle ending ${normalizeFinancialYearEnd(company.financialYearEnd)} in accordance with ${company.accountingFramework} and applicable statutory frameworks. The Statement of Cash Flows is presented using the Indirect Method in compliance with IAS 7 / ASC 230, distinguishing between operating, investing, and financing activities. Cash and cash equivalents comprise cash in hand, call deposits with scheduled banks, and liquid money-market placements with original maturities of three months or less.`,
    },
    {
      noteNumber: 2,
      title: `${nLbl('income', 'revenue', 'Sales Revenue')} & Disaggregation of Contracts`,
      category: 'Income Statement',
      textDescription: 'Revenue is recognized upon transfer of promised goods or services to customers at an amount that reflects the consideration to which the entity expects to be entitled in exchange for those goods or services.',
      tableData: {
        headers: ['Revenue Streams', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Commercial Contract Deliverables', values: [cRevSplit[0].toLocaleString(), pRevSplit[0].toLocaleString()] },
          { label: 'Institutional & Direct Enterprise Sales', values: [cRevSplit[1].toLocaleString(), pRevSplit[1].toLocaleString()] },
          { label: 'Technical Services & Ancillary Revenue', values: [cRevSplit[2].toLocaleString(), pRevSplit[2].toLocaleString()] },
        ],
        totalRow: { label: `Total ${nLbl('income', 'revenue', 'Revenue from Contracts')}`, values: [cRevenue.toLocaleString(), pRevenue.toLocaleString()] },
      },
    },
    {
      noteNumber: 3,
      title: `${nLbl('income', 'costOfSales', 'Cost of Sales')} & Operating Overheads`,
      category: 'Income Statement',
      textDescription: 'Cost of sales reflects direct raw materials consumed, conversion costs, production plant utilities, direct labor, and manufacturing depreciation allocated in accordance with production throughput.',
      tableData: {
        headers: ['Component Breakdown', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Raw Materials & Components Consumed', values: [cCogsSplit[0].toLocaleString(), pCogsSplit[0].toLocaleString()] },
          { label: 'Direct Labor & Industrial Staff Costs', values: [cCogsSplit[1].toLocaleString(), pCogsSplit[1].toLocaleString()] },
          { label: 'Factory Overheads, Fuel & Power', values: [cCogsSplit[2].toLocaleString(), pCogsSplit[2].toLocaleString()] },
        ],
        totalRow: { label: `Total ${nLbl('income', 'costOfSales', 'Cost of Sales')}`, values: [cCogs.toLocaleString(), pCogs.toLocaleString()] },
      },
    },
    {
      noteNumber: 4,
      title: `${nLbl('income', 'financeCost', 'Finance Cost')} & Borrowing Disclosures`,
      category: 'Financing',
      textDescription: 'Finance costs encompass interest charges on term borrowings, bank overdraft facilities, letter of credit commissions, and Islamic financing profit rates where applicable. In accordance with Section 76.4, Finance Cost is a mandatory user-entered parameter across all historical and projected reporting cycles.',
      tableData: {
        headers: ['Financing Facility Charges', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Interest on Long-Term Credit Facilities', values: [cFinSplit[0].toLocaleString(), pFinSplit[0].toLocaleString()] },
          { label: 'Interest on Working Capital & Short-Term Debt', values: [cFinSplit[1].toLocaleString(), pFinSplit[1].toLocaleString()] },
          { label: 'Bank Charges, LC Fees & Financing Commissions', values: [cFinSplit[2].toLocaleString(), pFinSplit[2].toLocaleString()] },
        ],
        totalRow: { label: `Total ${nLbl('income', 'financeCost', 'Finance Cost')}`, values: [cFinCost.toLocaleString(), pFinCost.toLocaleString()] },
      },
    },
    {
      noteNumber: 5,
      title: 'Property, Plant and Equipment (PPE)',
      category: 'Balance Sheet',
      textDescription: 'Items of PPE are stated at cost less accumulated depreciation and impairment losses. Depreciation is computed on the straight-line method based on estimated economic useful lives. Capital additions and disposals are linked directly to Cash Flows from Investing Activities.',
      tableData: {
        headers: ['PPE Category', 'Gross Cost', 'Accumulated Dep.', 'Net Book Value'],
        rows: note5Categories.map(cat => ({
          label: cat.name,
          values: [cat.costClosing.toLocaleString(), cat.depClosing.toLocaleString(), cat.netBookValue.toLocaleString()],
        })),
        totalRow: {
          label: 'Total Fixed Assets Net Book Value',
          values: [cFA.totalCostClosing.toLocaleString(), cFA.totalDepClosing.toLocaleString(), cFA.totalNetBookValue.toLocaleString()],
        },
      },
    },
    {
      noteNumber: 6,
      title: nLbl('bs', 'closingInventory', 'Closing Inventories'),
      category: 'Working Capital',
      textDescription: 'Inventories are valued at the lower of cost and net realizable value (NRV). Cost is determined using the weighted average method. Provisions are made for slow-moving and obsolete lines.',
      tableData: {
        headers: ['Inventory Classes', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Raw Materials & Imported Spares', values: [cInvSplit[0].toLocaleString(), pInvSplit[0].toLocaleString()] },
          { label: 'Work-in-Progress (WIP)', values: [cInvSplit[1].toLocaleString(), pInvSplit[1].toLocaleString()] },
          { label: 'Finished Goods Inventory', values: [cInvSplit[2].toLocaleString(), pInvSplit[2].toLocaleString()] },
        ],
        totalRow: { label: `Total ${nLbl('bs', 'closingInventory', 'Closing Inventory')}`, values: [cInv.toLocaleString(), pInv.toLocaleString()] },
      },
    },
    {
      noteNumber: 7,
      title: 'Trade and Other Receivables',
      category: 'Working Capital',
      textDescription: 'Trade receivables are recognized initially at fair value and subsequently measured at amortized cost less provision for impairment using the IFRS 9 simplified expected credit loss matrix.',
      tableData: {
        headers: ['Receivable Category', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: `${nLbl('bs', 'tradeReceivables', 'Trade Debtors')} (within 30 days)`, values: [cRecSplit[0].toLocaleString(), pRecSplit[0].toLocaleString()] },
          { label: `${nLbl('bs', 'tradeReceivables', 'Trade Debtors')} (31 - 90 days)`, values: [cRecSplit[1].toLocaleString(), pRecSplit[1].toLocaleString()] },
          { label: `${nLbl('bs', 'tradeReceivables', 'Trade Debtors')} (over 90 days, secured)`, values: [cRecSplit[2].toLocaleString(), pRecSplit[2].toLocaleString()] },
          { label: nLbl('bs', 'otherReceivables', 'Advance Deposits & Prepayments'), values: [cOthRec.toLocaleString(), pOthRec.toLocaleString()] },
        ],
        totalRow: { label: 'Total Receivables', values: [(cTrRec + cOthRec).toLocaleString(), (pTrRec + pOthRec).toLocaleString()] },
      },
    },
    {
      noteNumber: 8,
      title: `${nLbl('bs', 'cashAndEquivalents', 'Cash and Cash Equivalents')} & Bank Accounts`,
      category: 'Liquidity',
      textDescription: 'Cash and cash equivalents comprise operational petty cash floats and bank current/savings accounts with scheduled commercial financial institutions.',
      tableData: {
        headers: ['Liquid Asset Classes', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Cash in Hand (Corporate & Branches)', values: [cCashSplit[0].toLocaleString(), pCashSplit[0].toLocaleString()] },
          { label: 'Balances with Commercial Banks', values: [cCashSplit[1].toLocaleString(), pCashSplit[1].toLocaleString()] },
          { label: 'Short-Term Liquid Placements (< 90 Days)', values: [cCashSplit[2].toLocaleString(), pCashSplit[2].toLocaleString()] },
        ],
        totalRow: { label: `Total Balance Sheet ${nLbl('bs', 'cashAndEquivalents', 'Cash & Cash Equivalents')}`, values: [cCashBS.toLocaleString(), pCashBS.toLocaleString()] },
      },
    },
    {
      noteNumber: 9,
      title: 'Share Capital & Reserves',
      category: 'Equity',
      textDescription: 'The authorized share capital of the company is 20,000,000 ordinary shares of par value 1.00 each. The issued, subscribed and paid-up capital represents fully paid ordinary equity shares.',
      tableData: {
        headers: ['Equity Component', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: nLbl('bs', 'shareCapital', 'Paid-Up Ordinary Share Capital'), values: [cSC.toLocaleString(), pSC.toLocaleString()] },
          { label: nLbl('bs', 'retainedEarnings', 'Retained Earnings Balance'), values: [cRE.toLocaleString(), pRE.toLocaleString()] },
          { label: nLbl('bs', 'reserves', 'Capital & Asset Revaluation Reserves'), values: [cRes.toLocaleString(), pRes.toLocaleString()] },
          ...customEquityRows,
        ],
        totalRow: { label: 'Total Shareholders Equity', values: [cTotEq.toLocaleString(), pTotEq.toLocaleString()] },
      },
    },
    {
      noteNumber: 10,
      title: 'Borrowings & Financing Facilities',
      category: 'Liabilities',
      textDescription: 'Long-term borrowings represent syndication term facilities secured against fixed charges over industrial land, factory buildings, and plant machinery. Short-term facilities support ongoing raw material inventory imports.',
      tableData: {
        headers: ['Credit Facility Type', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: nLbl('bs', 'longTermBorrowings', 'Secured Long-Term Term Loans'), values: [cLTB.toLocaleString(), pLTB.toLocaleString()] },
          { label: nLbl('bs', 'shortTermBorrowings', 'Short-Term Working Capital & Overdraft Facilities'), values: [cSTB.toLocaleString(), pSTB.toLocaleString()] },
        ],
        totalRow: { label: 'Total Financial Debt Obligations', values: [(cLTB + cSTB).toLocaleString(), (pLTB + pSTB).toLocaleString()] },
      },
    },
    {
      noteNumber: 11,
      title: 'Trade and Other Payables',
      category: 'Working Capital',
      textDescription: 'Trade payables are liabilities for goods and services provided to the entity prior to the end of the financial year which are unpaid. All amounts are non-interest bearing and are settled within customary supplier terms.',
      tableData: {
        headers: ['Payable Type', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: nLbl('bs', 'tradePayables', 'Trade Creditors & Raw Material Suppliers'), values: [cTrPay.toLocaleString(), pTrPay.toLocaleString()] },
          { label: nLbl('bs', 'otherPayables', 'Accrued Expenses & Statutory Withholdings'), values: [cOthPay.toLocaleString(), pOthPay.toLocaleString()] },
        ],
        totalRow: { label: 'Total Current Payables', values: [(cTrPay + cOthPay).toLocaleString(), (pTrPay + pOthPay).toLocaleString()] },
      },
    },
    {
      noteNumber: 12,
      title: 'Cash Flow Disclosures & Balance Sheet Reconciliation',
      category: 'Cash Flow Reconciliation',
      textDescription: `Pursuant to Section 76.3 and 76.4, the Statement of Cash Flows directly incorporates the user-entered Cash Flow from Financing Activities (${cCfFinDirect.toLocaleString()} ${company.reportingCurrency}) without automatic derivations or assumptions. Closing cash is reconciled with the Statement of Financial Position: Closing Calculated Cash is ${cCfClose.toLocaleString()} vs Balance Sheet Cash of ${cCfStatedBS.toLocaleString()}. Audit status: ${cCfReconciled ? 'COMPLETELY RECONCILED (Variance 0)' : `VARIANCE OF ${cCfDiff.toLocaleString()} IDENTIFIED`}. ${cCfReconciled ? 'Fully reconciled with Balance Sheet Cash & Cash Equivalents.' : cCF.reconciliationExplanation}`,
      tableData: {
        headers: ['Reconciliation Parameter', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Net Cash Generated from Operating Activities', values: [cCfOp.toLocaleString(), pCfOp.toLocaleString()] },
          { label: 'Net Cash Generated from / (Used in) Investing Activities', values: [cCfInv.toLocaleString(), pCfInv.toLocaleString()] },
          { label: 'Directly Entered Cash Flow from Financing Activities', values: [cCfFinDirect.toLocaleString(), pCfFinDirect.toLocaleString()] },
          { label: 'Net Change in Cash and Cash Equivalents', values: [cCfChange.toLocaleString(), pCfChange.toLocaleString()] },
          { label: 'Opening Cash and Cash Equivalents', values: [cCfOpen.toLocaleString(), pCfOpen.toLocaleString()] },
          { label: 'Closing Cash Derived from Cash Flow Statement', values: [cCfClose.toLocaleString(), pCfClose.toLocaleString()] },
          { label: 'Stated Cash & Equivalents in Statement of Financial Position', values: [cCfStatedBS.toLocaleString(), pCfStatedBS.toLocaleString()] },
        ],
        totalRow: {
          label: 'Audit Reconciliation Variance',
          values: [cCfDiff.toLocaleString(), pCfDiff.toLocaleString()],
        },
      },
    },
  ];
}

// Default Configuration Items for the Financial Statement Configuration Module across all 5 statements
export const DEFAULT_STATEMENT_CONFIG_ITEMS: StatementConfigItem[] = [
  // Income Statement
  { id: 'inc_revenue', statement: 'income', section: 'Revenue & Gross Margin', label: 'Sales Revenue / Turnover', isMandatory: true, isVisible: true, order: 1 },
  { id: 'inc_costOfSales', statement: 'income', section: 'Revenue & Gross Margin', label: 'Cost of Sales', isMandatory: false, isVisible: true, order: 2 },
  { id: 'inc_operatingExpenses', statement: 'income', section: 'Operating Expenses', label: 'Operating & Administrative Expenses', isMandatory: false, isVisible: true, order: 3 },
  { id: 'inc_depreciationExpense', statement: 'income', section: 'Operating Expenses', label: 'Depreciation Expense (Linked to Fixed Assets)', isMandatory: false, isVisible: true, order: 4 },
  { id: 'inc_interestIncome', statement: 'income', section: 'Non-Operating & Tax', label: 'Interest & Investment Income', isMandatory: false, isVisible: true, order: 5 },
  { id: 'inc_financeCost', statement: 'income', section: 'Non-Operating & Tax', label: 'Finance Cost', isMandatory: true, isVisible: true, order: 6 },
  { id: 'inc_taxExpense', statement: 'income', section: 'Non-Operating & Tax', label: 'Income Tax Expense', isMandatory: false, isVisible: true, order: 7 },

  // Balance Sheet - Current & Non-Current Assets
  { id: 'bs_cashAndEquivalents', statement: 'balanceSheet', section: 'Current Assets', label: 'Cash and Cash Equivalents', isMandatory: true, isVisible: true, order: 1 },
  { id: 'bs_tradeReceivables', statement: 'balanceSheet', section: 'Current Assets', label: 'Trade Receivables', isMandatory: false, isVisible: true, order: 2 },
  { id: 'bs_otherReceivables', statement: 'balanceSheet', section: 'Current Assets', label: 'Other Receivables & Prepayments', isMandatory: false, isVisible: true, order: 3 },
  { id: 'bs_closingInventory', statement: 'balanceSheet', section: 'Current Assets', label: 'Closing Inventory', isMandatory: true, isVisible: true, order: 4 },
  { id: 'bs_otherCurrentAssets', statement: 'balanceSheet', section: 'Current Assets', label: 'Other Current Assets', isMandatory: false, isVisible: true, order: 5 },
  { id: 'bs_propertyPlantEquipment', statement: 'balanceSheet', section: 'Non-Current Assets', label: 'Property, Plant & Equipment (WDV)', isMandatory: true, isVisible: true, order: 6 },
  { id: 'bs_intangibleAssets', statement: 'balanceSheet', section: 'Non-Current Assets', label: 'Intangible Assets', isMandatory: false, isVisible: true, order: 7 },
  { id: 'bs_otherNonCurrentAssets', statement: 'balanceSheet', section: 'Non-Current Assets', label: 'Other Non-Current Assets', isMandatory: false, isVisible: true, order: 8 },
  // Balance Sheet - Liabilities & Equity
  { id: 'bs_tradePayables', statement: 'balanceSheet', section: 'Current Liabilities', label: 'Trade Payables', isMandatory: false, isVisible: true, order: 9 },
  { id: 'bs_otherPayables', statement: 'balanceSheet', section: 'Current Liabilities', label: 'Other Payables & Accruals', isMandatory: false, isVisible: true, order: 10 },
  { id: 'bs_shortTermBorrowings', statement: 'balanceSheet', section: 'Current Liabilities', label: 'Short-Term Borrowings', isMandatory: false, isVisible: true, order: 11 },
  { id: 'bs_currentTaxPayable', statement: 'balanceSheet', section: 'Current Liabilities', label: 'Current Tax Payable', isMandatory: false, isVisible: true, order: 12 },
  { id: 'bs_longTermBorrowings', statement: 'balanceSheet', section: 'Non-Current Liabilities', label: 'Long-Term Borrowings', isMandatory: false, isVisible: true, order: 13 },
  { id: 'bs_otherNonCurrentLiabilities', statement: 'balanceSheet', section: 'Non-Current Liabilities', label: 'Other Non-Current Liabilities', isMandatory: false, isVisible: true, order: 14 },
  { id: 'bs_shareCapital', statement: 'balanceSheet', section: 'Equity', label: 'Share Capital', isMandatory: true, isVisible: true, order: 15 },
  { id: 'bs_retainedEarnings', statement: 'balanceSheet', section: 'Equity', label: 'Retained Earnings', isMandatory: true, isVisible: true, order: 16 },
  { id: 'bs_reserves', statement: 'balanceSheet', section: 'Equity', label: 'Reserves & Surplus', isMandatory: false, isVisible: true, order: 17 },

  // Cash Flow Configuration Items (Section 76.6)
  { id: 'cf_pbt', statement: 'cashFlow', section: 'Operating Activities', label: 'Net Profit Before Tax', isMandatory: true, isVisible: true, order: 1 },
  { id: 'cf_dep', statement: 'cashFlow', section: 'Operating Activities', label: 'Adjustments for Depreciation', isMandatory: true, isVisible: true, order: 2 },
  { id: 'cf_amort', statement: 'cashFlow', section: 'Operating Activities', label: 'Adjustments for Amortization', isMandatory: false, isVisible: true, order: 3 },
  { id: 'cf_finance', statement: 'cashFlow', section: 'Operating Activities', label: 'Finance Cost Adjustment', isMandatory: true, isVisible: true, order: 4 },
  { id: 'cf_interest_inc', statement: 'cashFlow', section: 'Operating Activities', label: 'Interest & Investment Income', isMandatory: false, isVisible: true, order: 5 },
  { id: 'cf_disposal_gain_loss', statement: 'cashFlow', section: 'Operating Activities', label: 'Loss / (Gain) on Disposal of Fixed Assets', isMandatory: false, isVisible: true, order: 6 },
  { id: 'cf_wc_inventory', statement: 'cashFlow', section: 'Operating Activities', label: '(Increase) / Decrease in Inventory', isMandatory: false, isVisible: true, order: 7 },
  { id: 'cf_wc_receivables', statement: 'cashFlow', section: 'Operating Activities', label: '(Increase) / Decrease in Trade Receivables', isMandatory: false, isVisible: true, order: 8 },
  { id: 'cf_wc_other_rec', statement: 'cashFlow', section: 'Operating Activities', label: '(Increase) / Decrease in Other Receivables', isMandatory: false, isVisible: true, order: 9 },
  { id: 'cf_wc_other_ca', statement: 'cashFlow', section: 'Operating Activities', label: '(Increase) / Decrease in Other Current Assets', isMandatory: false, isVisible: true, order: 10 },
  { id: 'cf_wc_payables', statement: 'cashFlow', section: 'Operating Activities', label: 'Increase / (Decrease) in Trade Payables', isMandatory: false, isVisible: true, order: 11 },
  { id: 'cf_wc_other_pay', statement: 'cashFlow', section: 'Operating Activities', label: 'Increase / (Decrease) in Other Payables', isMandatory: false, isVisible: true, order: 12 },
  { id: 'cf_tax_paid', statement: 'cashFlow', section: 'Operating Activities', label: 'Income Tax Paid', isMandatory: true, isVisible: true, order: 13 },
  { id: 'cf_finance_paid', statement: 'cashFlow', section: 'Operating Activities', label: 'Finance Cost Paid', isMandatory: true, isVisible: true, order: 14 },

  // Investing Activities
  { id: 'cf_pur_ppe', statement: 'cashFlow', section: 'Investing Activities', label: 'Purchase of Property, Plant and Equipment', isMandatory: true, isVisible: true, order: 15 },
  { id: 'cf_sale_ppe', statement: 'cashFlow', section: 'Investing Activities', label: 'Proceeds from Sale of Fixed Assets', isMandatory: false, isVisible: true, order: 16 },
  { id: 'cf_pur_intangible', statement: 'cashFlow', section: 'Investing Activities', label: 'Purchase of Intangible Assets', isMandatory: false, isVisible: true, order: 17 },
  { id: 'cf_interest_rec', statement: 'cashFlow', section: 'Investing Activities', label: 'Interest Received', isMandatory: false, isVisible: true, order: 18 },
  { id: 'cf_dividends_rec', statement: 'cashFlow', section: 'Investing Activities', label: 'Dividends Received', isMandatory: false, isVisible: true, order: 19 },

  // Financing Activities
  { id: 'cf_direct_financing', statement: 'cashFlow', section: 'Financing Activities', label: 'Net Cash Flow from Financing Activities (Direct User Entry)', isMandatory: true, isVisible: true, order: 20 },
  { id: 'cf_shares_issue', statement: 'cashFlow', section: 'Financing Activities', label: 'Proceeds from Issue of Share Capital', isMandatory: false, isVisible: true, order: 21 },
  { id: 'cf_borrow_proceeds', statement: 'cashFlow', section: 'Financing Activities', label: 'Proceeds from Borrowings & Financing', isMandatory: false, isVisible: true, order: 22 },
  { id: 'cf_borrow_repay', statement: 'cashFlow', section: 'Financing Activities', label: 'Repayment of Borrowings & Financing', isMandatory: false, isVisible: true, order: 23 },
  { id: 'cf_dividends_paid', statement: 'cashFlow', section: 'Financing Activities', label: 'Dividends Paid to Shareholders', isMandatory: false, isVisible: true, order: 24 },

  // Statement of Changes in Equity
  { id: 'eq_openingShareCapital', statement: 'equity', section: 'Share Capital', label: 'Opening Share Capital', isMandatory: true, isVisible: true, order: 1 },
  { id: 'eq_shareCapitalIssued', statement: 'equity', section: 'Share Capital', label: 'Share Capital Issued', isMandatory: false, isVisible: true, order: 2 },
  { id: 'eq_openingRetainedEarnings', statement: 'equity', section: 'Retained Earnings', label: 'Opening Retained Earnings', isMandatory: true, isVisible: true, order: 3 },
  { id: 'eq_netProfitForYear', statement: 'equity', section: 'Retained Earnings', label: 'Net Profit for the Year', isMandatory: true, isVisible: true, order: 4 },
  { id: 'eq_dividendsPaid', statement: 'equity', section: 'Retained Earnings', label: 'Dividends Paid / Drawings', isMandatory: false, isVisible: true, order: 5 },
  { id: 'eq_openingReserves', statement: 'equity', section: 'Reserves & OCI', label: 'Opening Reserves', isMandatory: false, isVisible: true, order: 6 },
  { id: 'eq_otherComprehensiveIncome', statement: 'equity', section: 'Reserves & OCI', label: 'Other Comprehensive Income', isMandatory: false, isVisible: true, order: 7 },

  // Schedule of Fixed Assets Categories
  { id: 'fa_land_buildings', statement: 'fixedAssets', section: 'Asset Categories', label: 'Freehold Land & Industrial Buildings', isMandatory: false, isVisible: true, order: 1 },
  { id: 'fa_plant_machinery', statement: 'fixedAssets', section: 'Asset Categories', label: 'Plant, Machinery & Automated Systems', isMandatory: false, isVisible: true, order: 2 },
  { id: 'fa_office_equipment', statement: 'fixedAssets', section: 'Asset Categories', label: 'Office Equipment & Fixtures', isMandatory: false, isVisible: true, order: 3 },
  { id: 'fa_motor_vehicles', statement: 'fixedAssets', section: 'Asset Categories', label: 'Motor Vehicles & Logistics Fleet', isMandatory: false, isVisible: true, order: 4 },
  { id: 'fa_computer_it', statement: 'fixedAssets', section: 'Asset Categories', label: 'Computer, IT & Telecommunications', isMandatory: false, isVisible: true, order: 5 },
];

export const COMPANY_TYPE_TEMPLATES: Record<EntityType, { title: string; desc: string; operatingCustomLines: string[]; financingCustomLines: string[] }> = {
  manufacturing: {
    title: 'Manufacturing & Industrial Production',
    desc: 'Heavy inventory cycle (Raw, WIP, Finished Goods), capital-intensive fixed asset schedule, and supplier working capital.',
    operatingCustomLines: ['Raw Materials Working Capital Movement', 'Factory Overheads Non-Cash Allocation'],
    financingCustomLines: ['Industrial Term Loan Financing', 'Syndicated Machinery Financing'],
  },
  trading: {
    title: 'Wholesale & Retail Trading',
    desc: 'High velocity inventory turnover, mercantile debtor terms, and merchant banking facilities.',
    operatingCustomLines: ['Merchandise Inventory Velocity Adjustment', 'Supplier Trade Rebate Adjustments'],
    financingCustomLines: ['Short-Term Commercial Paper', 'Revolving Working Capital Lines'],
  },
  service: {
    title: 'Professional & Technical Services',
    desc: 'Low capex, zero manufacturing inventory, high receivables and personnel expense focus.',
    operatingCustomLines: ['Unbilled Work-in-Progress (Services)', 'Deferred Service Revenue Changes'],
    financingCustomLines: ['Partner Capital Contributions', 'Retained Earnings Distribution'],
  },
  banking_conventional: {
    title: 'Commercial Banking (Conventional)',
    desc: 'Statutory cash reserve deposits, loans and advances to customers, customer term deposits, and interbank repo.',
    operatingCustomLines: ['Net (Increase)/Decrease in Loans and Advances to Customers', 'Net Increase/(Decrease) in Customer Deposit Liabilities', 'Statutory Cash Reserve Ratio (CRR) Adjustments with Central Bank'],
    financingCustomLines: ['Subordinated Tier-II Debt Instruments', 'Perpetual Bond Financing'],
  },
  banking_islamic: {
    title: 'Islamic & Shariah Compliant Banking',
    desc: 'Murabaha, Ijarah, Mudaraba, and Musharaka asset disbursements, Mudaraba profit sharing deposits.',
    operatingCustomLines: ['Movement in Murabaha & Ijarah Financing Assets', 'Movement in Mudaraba & Musharaka Investment Accounts', 'Zakat & Qard Hasan Disbursements'],
    financingCustomLines: ['Sukuk Issuance Proceeds', 'Tier-I Mudaraba Capital Certificates'],
  },
  nbfi: {
    title: 'Non-Bank Financial Institution (NBFI)',
    desc: 'Leasing operations, consumer credit, bridge financing, and institutional term borrowings.',
    operatingCustomLines: ['Lease Receivables Portfolio Movement', 'Consumer Credit Allowances'],
    financingCustomLines: ['Institutional Term Borrowings', 'Securitized Debt Facilities'],
  },
  real_estate: {
    title: 'Real Estate & Infrastructure Construction',
    desc: 'Long cycle construction WIP, project milestone billing, customer construction advances, and land acquisition capex.',
    operatingCustomLines: ['Real Estate Development Inventory Under Construction', 'Customer Advances on Property Bookings', 'Contract Retention Money Adjustments'],
    financingCustomLines: ['Project Finance Debt Facilities', 'Real Estate Investment Trust (REIT) Capital'],
  },
  healthcare: {
    title: 'Healthcare & Hospital Networks',
    desc: 'Medical equipment depreciation, medical supplies inventory, insurance claim receivables.',
    operatingCustomLines: ['Medical & Pharmaceutical Consumables Adjustment', 'Health Insurance Receivables Settlement'],
    financingCustomLines: ['Diagnostic Equipment Leasing Obligations', 'Hospital Expansion Term Debt'],
  },
  education: {
    title: 'Educational Institutions & Universities',
    desc: 'Deferred tuition fees, academic infrastructure capex, research endowment grants.',
    operatingCustomLines: ['Deferred Tuition & Enrollment Fees', 'Academic Grant Adjustments'],
    financingCustomLines: ['Endowment Capital Inflows', 'Campus Infrastructure Development Financing'],
  },
  ngo: {
    title: 'Non-Governmental Organization (NGO / Non-Profit)',
    desc: 'Donor grant funding receipts, restricted program funds, zero commercial equity, social program disbursements.',
    operatingCustomLines: ['Unearned Grant Receipts Adjustment', 'Program Beneficiary Aid Disbursements'],
    financingCustomLines: ['Donor Capital Endowment Contributions', 'Institutional Soft Loans'],
  },
};

// ---------------- PREVIOUS FINANCIALS FORMAT BUILDER (Exact Compliance) ---------------- //
export function buildPreviousFormatData(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  equityStatements: EquityStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  options?: StatementCalculationOptions,
  financialYearEnd?: string
): import('../types/financial').PreviousFormatYearData[] {
  const histIndices = inputs
    .map((inp, idx) => ({ inp, idx }))
    .filter(item => !item.inp.isProjected);
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const vis = (id: string) => isConfigElementVisible(configMap, id);

  const prevResults: import('../types/financial').PreviousFormatYearData[] = [];
  let prevExpenseSplit: number[] | undefined = undefined;

  histIndices.forEach(({ inp, idx }, hIdx) => {
    const inc = incomeStatements[idx];
    const bs = balanceSheets[idx];
    const cf = cashFlowStatements[idx];
    const eq = equityStatements[idx];
    const fa = fixedAssetsSchedule[idx];

    // Recalculated drivers synchronized with Section 76 engines
    const sales = inc.revenue;
    const stockAndStores = bs.closingInventory;
    const financialExp = inc.financeCost;
    const netProfitAfterIncomeTax = inc.netProfit;
    const repaymentOfBorrowings = cf.repaymentOfBorrowings;
    const faGrossAssetBase = fa.totalCostClosing;

    // Income Statement exact breakdown
    const costOfGoodsSold = inc.costOfSales;
    const grossProfit = inc.grossProfit;
    const depreciation = inc.depreciationExpense;
    const customOpRev = getVisibleCustomItemsSum(options?.statementConfig, 'income', inp.year, 'Revenue & Gross Margin');
    const customOpExp = getVisibleCustomItemsSum(options?.statementConfig, 'income', inp.year, 'Operating Expenses');
    const nonDepOpEx = vis('inc_operatingExpenses') ? Math.max(0, r2(inc.operatingExpenses + customOpExp)) : 0;

    // Allocate non-depreciation Operating Expenses with realistic non-uniform annual cost shares across all 12 lines
    const baseWeights = [
      0.281, // Salary & Allowances
      0.018, // Telephone & Mobile Bill
      0.002, // Postage & Courier
      0.032, // Electric Charges
      0.040, // Entertainment
      0.012, // Printing & Stationery
      0.082, // General Expenses
      0.011, // Licence & Renewals
      0.003, // Audit & Professional Fee
      0.042, // Travelling & Conveyance
      0.002, // News Paper & Periodicals
      0.475, // Selling Exp. (residual)
    ];
    const expenseSplit = splitTotalNonUniform(
      nonDepOpEx,
      inp.year,
      140,
      baseWeights,
      [0.024, 0.004, 0.0008, 0.0055, 0.0065, 0.0028, 0.012, 0.0028, 0.0009, 0.0068, 0.0007, 0.025],
      prevExpenseSplit,
      1.35
    );
    prevExpenseSplit = expenseSplit;
    const salaryAndAllowances = expenseSplit[0];
    const telephoneAndMobileBill = expenseSplit[1];
    const postageAndCourier = expenseSplit[2];
    const electricCharges = expenseSplit[3];
    const entertainment = expenseSplit[4];
    const printingAndStationery = expenseSplit[5];
    const generalExpenses = expenseSplit[6];
    const licenceAndRenewals = expenseSplit[7];
    const auditAndProfessionalFee = expenseSplit[8];
    const travellingAndConveyance = expenseSplit[9];
    const newspaperAndPeriodicals = expenseSplit[10];
    const partialSum =
      salaryAndAllowances +
      telephoneAndMobileBill +
      postageAndCourier +
      electricCharges +
      entertainment +
      printingAndStationery +
      generalExpenses +
      licenceAndRenewals +
      auditAndProfessionalFee +
      travellingAndConveyance +
      newspaperAndPeriodicals;
    const sellingExp = Math.max(0, r2(nonDepOpEx - partialSum));
    const adminAndSellingTotal = r2(partialSum + sellingExp - customOpExp + depreciation);
    const netOperatingProfit = r2(grossProfit - adminAndSellingTotal);
    const otherIncome = inc.interestIncome;
    const operatingSubtotal = r2(netOperatingProfit + otherIncome);
    const incomeTaxProvision = inc.taxExpense;
    void customOpRev;

    // Balance Sheet exact structure synchronized with BalanceSheetYear
    const fixedAssetsWDV = bs.propertyPlantEquipment;
    const customCurrentAssetsBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Current Assets');
    const customNonCurrentAssetsBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Non-Current Assets');
    const customCurrentLiabBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Current Liabilities');
    const customNonCurrentLiabBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Non-Current Liabilities');
    const customEquityBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Equity');
    const advancesDepositsPrepayments = r2(bs.otherReceivables + bs.otherCurrentAssets);
    const billsReceivable = bs.tradeReceivables;
    const cashAndCashEquivalents = bs.cashAndEquivalents;
    const totalCurrentAssets = r2(
      advancesDepositsPrepayments + stockAndStores + billsReceivable + cashAndCashEquivalents + customCurrentAssetsBS
    );
    const totalAssetsTaka = r2(fixedAssetsWDV + customNonCurrentAssetsBS + totalCurrentAssets);

    const forExpenses = bs.otherPayables;
    const accountsPayable = bs.tradePayables;
    const shortTermLoan = r2(bs.shortTermBorrowings + bs.longTermBorrowings);
    const provisionForTaxation = bs.currentTaxPayable;
    const totalCurrentLiabilities = r2(
      forExpenses + accountsPayable + shortTermLoan + provisionForTaxation + customCurrentLiabBS
    );

    const capitalAfterDrawings = r2(totalAssetsTaka - totalCurrentLiabilities - customNonCurrentLiabBS - customEquityBS);
    const drawings = eq.dividendsPaid;
    const capitalSubtotal = r2(capitalAfterDrawings + drawings);
    const openingCapital = r2(capitalSubtotal - netProfitAfterIncomeTax);

    // Cash Flow exact structure matching Previous Financials PDF & exact ending cash
    const prevBS = hIdx > 0 ? prevResults[hIdx - 1] : null;
    const cfNetIncome = vis('cf_pbt') ? netProfitAfterIncomeTax : 0;
    const cfDepreciation = vis('cf_dep') ? depreciation : 0;
    const nonCashCA = r2(advancesDepositsPrepayments + stockAndStores + billsReceivable + customCurrentAssetsBS);
    const prevNonCashCA = prevBS
      ? r2(prevBS.advancesDepositsPrepayments + prevBS.stockAndStores + prevBS.billsReceivable)
      : Math.round(nonCashCA * 0.92);
    const cfChangeInCurrentAssets = vis('cf_wc_receivables') ? r2(-(nonCashCA - prevNonCashCA)) : 0;
    const cfChangeInCWIP = 0;

    // Working capital liabilities + Short Term Loan changes
    const cfChangeInCurrentLiabilities = vis('cf_wc_payables')
      ? prevBS
        ? r2(totalCurrentLiabilities - prevBS.totalCurrentLiabilities)
        : Math.round(totalCurrentLiabilities * 0.05)
      : 0;
    const customCfOp = getVisibleCustomItemsSum(options?.statementConfig, 'cashFlow', inp.year, 'Operating Activities');
    const customCfInv = getVisibleCustomItemsSum(options?.statementConfig, 'cashFlow', inp.year, 'Investing Activities');
    const customCfFin = getVisibleCustomItemsSum(options?.statementConfig, 'cashFlow', inp.year, 'Financing Activities');
    const cfNetCashOperating = r2(
      cfNetIncome + cfDepreciation + cfChangeInCurrentAssets + cfChangeInCWIP + cfChangeInCurrentLiabilities + customCfOp
    );

    const cfAcquisitionPPE = cf.purchasePPE;
    const cfDisposalIntangible = vis('cf_pur_intangible') ? r2(fa.totalDisposals - fa.totalDepDisposals) : 0;
    const cfNetCashInvesting = r2(cfAcquisitionPPE + cfDisposalIntangible + customCfInv);

    const cfDrawings = vis('cf_dividends_paid') ? -drawings : 0;
    const cfOpeningBalance = hIdx === 0
      ? 0
      : r2(capitalAfterDrawings - (prevBS!.capitalAfterDrawings + netProfitAfterIncomeTax - drawings));
    const cfNetCashFinancing = r2(cfDrawings + cfOpeningBalance + customCfFin);

    const cfNetIncreaseCash = r2(cfNetCashOperating + cfNetCashInvesting + cfNetCashFinancing);
    const cfCashBeginning = prevBS
      ? prevBS.cfCashEnding
      : r2(cashAndCashEquivalents - cfNetIncreaseCash);
    const cfCashEnding = r2(cfCashBeginning + cfNetIncreaseCash);

    prevResults.push({
      year: inp.year,
      dateLabel: formatFyEndDateCode(inp.year, financialYearEnd),
      fixedAssetsWDV,
      advancesDepositsPrepayments,
      stockAndStores,
      billsReceivable,
      cashAndCashEquivalents,
      totalCurrentAssets,
      totalAssetsTaka,
      openingCapital,
      netProfitForYear: netProfitAfterIncomeTax,
      capitalSubtotal,
      drawings,
      capitalAfterDrawings,
      forExpenses,
      accountsPayable,
      shortTermLoan,
      provisionForTaxation,
      totalCurrentLiabilities,
      totalCapitalAndLiabilities: r2(capitalAfterDrawings + customEquityBS + customNonCurrentLiabBS + totalCurrentLiabilities),
      sales,
      costOfGoodsSold,
      grossProfit,
      salaryAndAllowances,
      telephoneAndMobileBill,
      postageAndCourier,
      electricCharges,
      entertainment,
      printingAndStationery,
      generalExpenses,
      licenceAndRenewals,
      auditAndProfessionalFee,
      travellingAndConveyance,
      newspaperAndPeriodicals,
      sellingExp,
      depreciation,
      adminAndSellingTotal,
      netOperatingProfit,
      otherIncome,
      operatingSubtotal,
      financialExp,
      incomeTaxProvision,
      netProfitAfterIncomeTax,
      cfNetIncome,
      cfDepreciation,
      cfChangeInCurrentAssets,
      cfChangeInCWIP,
      cfChangeInCurrentLiabilities,
      cfNetCashOperating,
      cfAcquisitionPPE,
      cfDisposalIntangible,
      cfNetCashInvesting,
      cfDrawings,
      cfRepaymentOfBorrowings: repaymentOfBorrowings,
      cfOpeningBalance,
      cfNetCashFinancing,
      cfNetIncreaseCash,
      cfCashBeginning,
      cfCashEnding,
      eqOpeningShareCapital: eq.openingShareCapital,
      eqShareCapitalIssued: eq.shareCapitalIssued,
      eqClosingShareCapital: eq.closingShareCapital,
      eqOpeningRetainedEarnings: eq.openingRetainedEarnings,
      eqNetProfitForYear: eq.netProfitForYear,
      eqDividendsPaid: eq.dividendsPaid,
      eqClosingRetainedEarnings: eq.closingRetainedEarnings,
      eqClosingReserves: eq.closingReserves,
      eqTotalClosingEquity: eq.totalClosingEquity,
      faOpening: fa.totalCostOpening,
      faAddition: fa.totalAdditions,
      faDisposal: fa.totalDisposals,
      faTotal: fa.totalCostClosing,
      faRatePct: inp.depreciationRate,
      faDepreciationCharged: fa.totalDepCharge,
      faWDV: fa.totalNetBookValue,
      faGrossAssetBase: faGrossAssetBase,
    });
  });

  // Sort descending by year (e.g. 2025, 2024, 2023) to match Previous_Financials PDF column order
  return prevResults.sort((a, b) => b.year - a.year);
}

/**
 * Splits selected previous financial years into consecutive 2-year comparative pairs.
 * Example: [2023, 2024, 2025] -> Pair 1: (2023 & 2024) [31.12.2024 | 31.12.2023],
 *                                 Pair 2: (2024 & 2025) [31.12.2025 | 31.12.2024].
 * Works dynamically for any number of selected historical years.
 */
export function buildPreviousComparativePairs(
  prevData: import('../types/financial').PreviousFormatYearData[]
): import('../types/financial').PreviousComparativePair[] {
  if (prevData.length === 0) return [];
  const asc = [...prevData].sort((a, b) => a.year - b.year);
  if (asc.length === 1) {
    const only = asc[0];
    return [
      {
        id: `${only.year}`,
        label: `${only.year}`,
        earlierYear: only.year,
        laterYear: only.year,
        data: [only],
      },
    ];
  }

  const pairs: import('../types/financial').PreviousComparativePair[] = [];
  for (let i = 0; i < asc.length - 1; i++) {
    const earlier = asc[i];
    const later = asc[i + 1];
    pairs.push({
      id: `${earlier.year}_${later.year}`,
      label: `${earlier.year} & ${later.year}`,
      earlierYear: earlier.year,
      laterYear: later.year,
      data: [later, earlier], // Descending column order: later year left, earlier year right
    });
  }
  return pairs;
}

/**
 * Builds Statutory Notes (1 to 12) strictly scoped to a specific Previous Comparative Pair
 * (e.g. 2023 & 2024 or 2024 & 2025) while preserving any user-customized note text.
 */
export function buildPreviousPairScopedNotes(
  company: CompanyProfile,
  pairYearsDesc: number[],
  allYearsInputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  baseNotes?: FinancialNoteItem[],
  defaultNotesRef?: FinancialNoteItem[],
  options?: {
    statementConfig?: StatementConfigItem[];
    labelOverrides?: StatementLabelOverrides;
    cellOverrides?: StatementCellOverrides;
  }
): FinancialNoteItem[] {
  const pairAsc = [...pairYearsDesc].sort((a, b) => a - b);
  const indices = pairAsc
    .map(yr => allYearsInputs.findIndex(y => y.year === yr))
    .filter(idx => idx !== -1);

  if (indices.length === 0) return baseNotes || [];

  const scopedNotes = generateNotesToFinancialStatements(
    company,
    indices.map(i => allYearsInputs[i].year),
    indices.map(i => incomeStatements[i]),
    indices.map(i => balanceSheets[i]),
    indices.map(i => cashFlowStatements[i]),
    indices.map(i => fixedAssetsSchedule[i]),
    options
  );

  // Preserve custom user edits to title/textDescription if user modified baseNotes
  if (baseNotes && defaultNotesRef && baseNotes !== defaultNotesRef) {
    return scopedNotes.map(n => {
      const custom = baseNotes.find(b => b.noteNumber === n.noteNumber);
      const origDef = defaultNotesRef.find(d => d.noteNumber === n.noteNumber);
      if (!custom) return n;
      const titleEdited = origDef && custom.title !== origDef.title;
      const descEdited = origDef && custom.textDescription !== origDef.textDescription;
      return {
        ...n,
        title: titleEdited ? custom.title : n.title,
        textDescription: descEdited ? custom.textDescription : n.textDescription,
      };
    });
  }

  return scopedNotes;
}

// ---------------- PROJECTED FINANCIALS FORMAT BUILDER (Exact Compliance) ---------------- //
export function buildProjectedFormatData(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  equityStatements: EquityStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  options?: StatementCalculationOptions
): import('../types/financial').ProjectedFormatYearData[] {
  const projIndices = inputs
    .map((inp, idx) => ({ inp, idx }))
    .filter(item => item.inp.isProjected);
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const vis = (id: string) => isConfigElementVisible(configMap, id);

  // Anchor opening year to last historical year if available
  const firstProjIdx = projIndices[0]?.idx ?? 0;
  const anchorHistIdx = firstProjIdx > 0 ? firstProjIdx - 1 : -1;

  const projResults: import('../types/financial').ProjectedFormatYearData[] = [];

  projIndices.forEach(({ inp, idx }, pIdx) => {
    const inc = incomeStatements[idx];
    const bs = balanceSheets[idx];
    const cf = cashFlowStatements[idx];
    const eq = equityStatements[idx];
    const fa = fixedAssetsSchedule[idx];

    const prevProj = pIdx > 0 ? projResults[pIdx - 1] : null;
    const prevBS = anchorHistIdx >= 0 && pIdx === 0 ? balanceSheets[anchorHistIdx] : null;
    const prevFA = anchorHistIdx >= 0 && pIdx === 0 ? fixedAssetsSchedule[anchorHistIdx] : null;

    // Recalculated figures synchronized with Section 76 engines
    const sales = inc.revenue;
    const inventories = bs.closingInventory;
    const financialExpenses = inc.financeCost;
    const netProfitAfterTax = inc.netProfit;
    const repaymentOfBorrowings = cf.repaymentOfBorrowings;
    const schedGrossAssetBase = fa.totalCostClosing;

    // P&L
    const costOfGoodsSold = inc.costOfSales;
    const grossProfit = inc.grossProfit;
    const othersIncome = inc.interestIncome;
    const grossProfitWithOtherIncome = r2(grossProfit + othersIncome);
    const adminAndSellingExpenses = r2(inc.operatingExpenses + inc.depreciationExpense);
    const totalOperatingProfit = r2(grossProfitWithOtherIncome - adminAndSellingExpenses);
    const netProfitBeforeTax = r2(totalOperatingProfit - financialExpenses);
    const provisionForIncomeTax = inc.taxExpense;

    // Balance Sheet Assets
    const customCurrentAssetsBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Current Assets');
    const customNonCurrentAssetsBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Non-Current Assets');
    const customCurrentLiabBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Current Liabilities');
    const customNonCurrentLiabBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Non-Current Liabilities');
    const propertyPlantEquipments = r2(bs.propertyPlantEquipment + customNonCurrentAssetsBS);
    const tradeAndOtherReceivables = bs.tradeReceivables;
    const advanceDepositPrepayments = r2(bs.otherReceivables + bs.otherCurrentAssets);

    // Liabilities
    const longTermLoan = r2(bs.longTermBorrowings + customNonCurrentLiabBS);
    const tradeAndOtherPayables = bs.tradePayables;
    const provisionsForTax = bs.currentTaxPayable;
    const shortTermLoan = bs.shortTermBorrowings;
    const liabilitiesForExp = bs.otherPayables;
    const totalCurrentLiabilities = r2(
      tradeAndOtherPayables + provisionsForTax + shortTermLoan + liabilitiesForExp + customCurrentLiabBS
    );

    // Cash Flow Statement (exact roll-forward matching Projected_Financials PDF + element toggles)
    const prevInv = prevProj ? prevProj.inventories : (prevBS ? prevBS.closingInventory : Math.round(inventories * 0.9));
    const prevRec = prevProj ? prevProj.tradeAndOtherReceivables : (prevBS ? prevBS.tradeReceivables : Math.round(tradeAndOtherReceivables * 0.9));
    const prevAdv = prevProj ? prevProj.advanceDepositPrepayments : (prevBS ? prevBS.otherReceivables + prevBS.otherCurrentAssets : Math.round(advanceDepositPrepayments * 0.9));
    const prevPay = prevProj ? prevProj.tradeAndOtherPayables : (prevBS ? prevBS.tradePayables : Math.round(tradeAndOtherPayables * 0.9));
    const prevTax = prevProj ? prevProj.provisionsForTax : (prevBS ? prevBS.currentTaxPayable : Math.round(provisionsForTax * 0.9));
    const prevExp = prevProj ? prevProj.liabilitiesForExp : (prevBS ? prevBS.otherPayables : Math.round(liabilitiesForExp * 0.9));

    const cfNetIncome = vis('cf_pbt') ? netProfitAfterTax : 0;
    const cfChangeInventory =
      vis('cf_wc_inventory') && vis('bs_closingInventory') ? r2(-(inventories - prevInv)) : 0;
    const cfChangeReceivables =
      vis('cf_wc_receivables') && vis('bs_tradeReceivables') ? r2(-(tradeAndOtherReceivables - prevRec)) : 0;
    const cfChangeAdvanceDeposit =
      vis('cf_wc_other_rec') && vis('bs_otherReceivables') ? r2(-(advanceDepositPrepayments - prevAdv)) : 0;
    const cfChangeCurrentAssets = 0;
    const cfChangeCurrentLiabilities = 0;
    const cfChangeTradePayables =
      vis('cf_wc_payables') && vis('bs_tradePayables') ? r2(tradeAndOtherPayables - prevPay) : 0;
    const cfChangeProvisionsTax =
      vis('cf_tax_paid') && vis('bs_currentTaxPayable') ? r2(provisionsForTax - prevTax) : 0;
    const cfChangeLiabilitiesExp =
      vis('cf_wc_other_pay') && vis('bs_otherPayables') ? r2(liabilitiesForExp - prevExp) : 0;
    const cfDepreciation = vis('cf_dep') ? inc.depreciationExpense : 0;
    const customCfOp = getVisibleCustomItemsSum(options?.statementConfig, 'cashFlow', inp.year, 'Operating Activities');
    const customCfInv = getVisibleCustomItemsSum(options?.statementConfig, 'cashFlow', inp.year, 'Investing Activities');
    const customCfFin = getVisibleCustomItemsSum(options?.statementConfig, 'cashFlow', inp.year, 'Financing Activities');

    const cfNetCashOperating = r2(
      cfNetIncome +
      cfChangeInventory +
      cfChangeReceivables +
      cfChangeAdvanceDeposit +
      cfChangeCurrentAssets +
      cfChangeCurrentLiabilities +
      cfChangeTradePayables +
      cfChangeProvisionsTax +
      cfChangeLiabilitiesExp +
      cfDepreciation +
      customCfOp
    );

    // Investing
    const cfSalePPE = vis('cf_sale_ppe') ? r2(fa.totalDisposals - fa.totalDepDisposals) : 0;
    const cfInvestingCapitalInflow = 0;
    const cfPreliminaryExpenses = 0;
    const cfPurchasePPE = cf.purchasePPE;
    const cfPurchaseInvestment = 0;
    const cfNetCashInvesting = r2(
      cfSalePPE + cfInvestingCapitalInflow + cfPreliminaryExpenses + cfPurchasePPE + cfPurchaseInvestment + customCfInv
    );

    // Financing (keeps gross Loan Received and Repayment of Borrowings & Financing separate without netting;
    // first projected year recognizes no fresh Loan Received since it models existing loan rescheduling)
    const isFirstProjectedYear = pIdx === 0;
    const cfLoanReceived = isFirstProjectedYear ? 0 : cf.financingProceedsBorrowings;
    const cfAdditionalCapital = cf.financingProceedsShares;
    const cfIssuanceOfDebt = 0;
    const cfDividendsToStockholders = 0;
    const drawings = eq.dividendsPaid;
    const cfDrawingsOutflow = cf.financingDividendsPaid;
    const cfRedeemLongTermDebt = cf.financingRepaymentBorrowings;
    const cfIssuanceBondsPayable = 0;
    const cfNetCashFinancing = r2(
      cfLoanReceived +
      cfAdditionalCapital +
      cfIssuanceOfDebt +
      cfDividendsToStockholders +
      cfDrawingsOutflow +
      cfRedeemLongTermDebt +
      cfIssuanceBondsPayable +
      customCfFin
    );

    const cfNetCashIncrease = r2(cfNetCashOperating + cfNetCashInvesting + cfNetCashFinancing);
    const cfOpeningCash = prevProj
      ? prevProj.cfClosingCash
      : (prevBS ? prevBS.cashAndEquivalents : Math.round(sales * 0.08));
    const cfClosingCash = r2(cfOpeningCash + cfNetCashIncrease);

    const cashAndCashEquivalents = vis('bs_cashAndEquivalents') ? cfClosingCash : 0;
    const totalCurrentAssets = r2(
      inventories + tradeAndOtherReceivables + advanceDepositPrepayments + cashAndCashEquivalents + customCurrentAssetsBS
    );
    const totalAssets = r2(propertyPlantEquipments + totalCurrentAssets);

    // Capital Roll-forward ensuring Total Assets === Total Equity & Liabilities (0.00 variance!)
    const customEquityBS = getVisibleCustomItemsSum(options?.statementConfig, 'balanceSheet', inp.year, 'Equity');
    const closingCapital = r2(totalAssets - longTermLoan - totalCurrentLiabilities - customEquityBS);
    const equitySubtotal = r2(closingCapital + drawings);
    const capital = prevProj
      ? prevProj.closingCapital
      : r2(equitySubtotal - cfAdditionalCapital - netProfitAfterTax);
    const additionalCapital = r2(equitySubtotal - capital - netProfitAfterTax);

    // Notes 3.00 - 13.00 Breakdown (non-uniform annual operational turnover & direct expense splits preserving exact roll-forward identities)
    const note4CreditShare = varyBounded(0.25, inp.year, 180, 0.045, 0.18, 0.33);
    const note4ReceivablesAdded = Math.max(0, tradeAndOtherReceivables - prevRec) + Math.round(sales * note4CreditShare);
    const note4ReceivablesReceived = prevRec + note4ReceivablesAdded - tradeAndOtherReceivables;

    const note5AdvTurnover = varyBounded(0.105, inp.year, 185, 0.028, 0.06, 0.16);
    const note5AdvanceAdded = Math.max(0, advanceDepositPrepayments - prevAdv) + Math.round(advanceDepositPrepayments * note5AdvTurnover);
    const note5AdvanceAdjusted = prevAdv + note5AdvanceAdded - advanceDepositPrepayments;

    const note8PayTurnover = varyBounded(0.15, inp.year, 190, 0.032, 0.10, 0.22);
    const note8PayablesAdded = Math.max(0, tradeAndOtherPayables - prevPay) + Math.round(costOfGoodsSold * note8PayTurnover);
    const note8PayablesPaid = prevPay + note8PayablesAdded - tradeAndOtherPayables;

    const note9TaxAdded = provisionForIncomeTax;
    const note9TaxAdjusted = prevTax + note9TaxAdded - provisionsForTax;

    const note11OpeningStock = prevInv;
    const note11ClosingStock = inventories;
    const directShareRate = varyBounded(0.0155, inp.year, 195, 0.0042, 0.0095, 0.024);
    const directTotal = Math.round(costOfGoodsSold * directShareRate);
    const prevDirectParts = prevProj
      ? [prevProj.note11LabourWages, prevProj.note11ElectricityBill, prevProj.note11CarriageInward, prevProj.note11OtherDirectExpenses]
      : undefined;
    const directSplit = splitTotalNonUniform(
      directTotal,
      inp.year,
      200,
      [0.18, 0.36, 0.21, 0.25],
      [0.038, 0.055, 0.038, 0.044],
      prevDirectParts,
      1.45
    );
    const note11LabourWages = directSplit[0];
    const note11ElectricityBill = directSplit[1];
    const note11CarriageInward = directSplit[2];
    const note11OtherDirectExpenses = Math.max(0, directSplit[3]);
    const note11DirectExpensesTotal = note11LabourWages + note11ElectricityBill + note11CarriageInward + note11OtherDirectExpenses;
    const materialConsumed = costOfGoodsSold - note11DirectExpensesTotal;
    const note11Purchase = materialConsumed + note11ClosingStock - note11OpeningStock;

    projResults.push({
      year: inp.year,
      propertyPlantEquipments,
      inventories,
      tradeAndOtherReceivables,
      advanceDepositPrepayments,
      cashAndCashEquivalents,
      totalCurrentAssets,
      totalAssets,
      capital,
      additionalCapital,
      profitDuringYear: netProfitAfterTax,
      equitySubtotal,
      drawings,
      closingCapital,
      longTermLoan,
      tradeAndOtherPayables,
      provisionsForTax,
      shortTermLoan,
      liabilitiesForExp,
      totalCurrentLiabilities,
      totalEquityAndLiabilities: r2(closingCapital + customEquityBS + longTermLoan + totalCurrentLiabilities),
      sales,
      costOfGoodsSold,
      grossProfit,
      othersIncome,
      grossProfitWithOtherIncome,
      adminAndSellingExpenses,
      totalOperatingProfit,
      financialExpenses,
      netProfitBeforeTax,
      provisionForIncomeTax,
      netProfitAfterTax,
      cfNetIncome,
      cfChangeInventory,
      cfChangeReceivables,
      cfChangeAdvanceDeposit,
      cfChangeCurrentAssets,
      cfChangeCurrentLiabilities,
      cfChangeTradePayables,
      cfChangeProvisionsTax,
      cfChangeLiabilitiesExp,
      cfDepreciation,
      cfNetCashOperating,
      cfSalePPE,
      cfInvestingCapitalInflow,
      cfPreliminaryExpenses,
      cfPurchasePPE,
      cfPurchaseInvestment,
      cfNetCashInvesting,
      cfLoanReceived,
      cfAdditionalCapital,
      cfIssuanceOfDebt,
      cfDividendsToStockholders,
      cfDrawingsOutflow,
      cfRedeemLongTermDebt,
      cfIssuanceBondsPayable,
      cfNetCashFinancing,
      cfNetCashIncrease,
      cfOpeningCash,
      cfClosingCash,
      note4ReceivablesOpening: prevRec,
      note4ReceivablesAdded,
      note4ReceivablesReceived,
      note5AdvanceOpening: prevAdv,
      note5AdvanceAdded,
      note5AdvanceAdjusted,
      note8PayablesOpening: prevPay,
      note8PayablesAdded,
      note8PayablesPaid,
      note9TaxOpening: prevTax,
      note9TaxAdded,
      note9TaxAdjusted,
      note11OpeningStock,
      note11Purchase,
      note11ClosingStock,
      note11LabourWages,
      note11ElectricityBill,
      note11CarriageInward,
      note11OtherDirectExpenses,
      note11DirectExpensesTotal,
      schedOpeningBalance: prevFA ? prevFA.totalCostClosing : fa.totalCostOpening,
      schedAddition: fa.totalAdditions,
      schedDisposals: fa.totalDisposals,
      schedBalanceClosing: fa.totalCostClosing,
      schedRatePct: inp.depreciationRate,
      schedDepreciation: fa.totalDepCharge,
      schedWDV: fa.totalNetBookValue,
      schedGrossAssetBase,
    });
  });

  // Sort descending by year (e.g. 2030, 2029, ..., 2026) to match Projected_Financials PDF column order
  return projResults.sort((a, b) => b.year - a.year);
}

export interface ProjectedQuantNoteRow {
  noteNo: string;
  label: string;
  configId?: string;
  isSectionHeader?: boolean;
  isBold?: boolean;
  isSubtotal?: boolean;
  indent?: boolean;
  values: (number | null)[];
}

export function buildProjectedQuantNoteRows(
  projData: import('../types/financial').ProjectedFormatYearData[],
  options?: {
    statementConfig?: StatementConfigItem[];
    labelOverrides?: StatementLabelOverrides;
    includeHidden?: boolean;
  }
): ProjectedQuantNoteRow[] {
  const cfg = options?.statementConfig;
  const lblOv = options?.labelOverrides;
  const configMap = buildConfigVisibilityMap(cfg);
  const vis = (id: string) => isConfigElementVisible(configMap, id);
  const nLbl = (stmt: string, field: string, def: string) =>
    resolveEffectiveLabel(stmt, field, def, lblOv, cfg);

  const invLbl = nLbl('bs', 'closingInventory', 'Inventories');
  const trRecLbl = nLbl('bs', 'tradeReceivables', 'Trade & Other Receivables');
  const advLbl = nLbl('bs', 'otherReceivables', 'Advance, Deposit & Prepayments');
  const cashLbl = nLbl('bs', 'cashAndEquivalents', 'Cash & Cash Equivalents');
  const ltLoanLbl = nLbl('bs', 'longTermBorrowings', 'Long Term Loan');
  const trPayLbl = nLbl('bs', 'tradePayables', 'Trade & Other Payables');
  const taxPayLbl = nLbl('bs', 'currentTaxPayable', 'Provisions For Tax');
  const stLoanLbl = nLbl('bs', 'shortTermBorrowings', 'Short Term Loan');
  const cogsLbl = nLbl('income', 'costOfSales', 'Cost of Goods Sold');
  const opexLbl = nLbl('income', 'operatingExpenses', 'Administrative & Selling Expenses');
  const finCostLbl = nLbl('income', 'financeCost', 'Financial Expenses');

  const includeHidden = Boolean(options?.includeHidden);
  const keep = (id: string) => includeHidden || vis(id);
  const nulls = () => projData.map(() => null);

  const rows: ProjectedQuantNoteRow[] = [];

  if (keep('bs_closingInventory')) {
    rows.push({
      noteNo: '3.00',
      label: invLbl,
      configId: 'bs_closingInventory',
      isBold: true,
      isSubtotal: true,
      values: projData.map(d => d.inventories),
    });
  }

  if (keep('bs_tradeReceivables')) {
    rows.push(
      { noteNo: '4.00', label: trRecLbl, configId: 'bs_tradeReceivables', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'As Per Last Year', configId: 'bs_tradeReceivables', indent: true, values: projData.map(d => d.note4ReceivablesOpening) },
      { noteNo: '', label: 'Add: Credit Sales During the Year', configId: 'bs_tradeReceivables', indent: true, values: projData.map(d => d.note4ReceivablesAdded) },
      { noteNo: '', label: 'Subtotal', configId: 'bs_tradeReceivables', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.note4ReceivablesOpening + d.note4ReceivablesAdded) },
      { noteNo: '', label: 'Less: Received During the Year', configId: 'bs_tradeReceivables', indent: true, values: projData.map(d => d.note4ReceivablesReceived) },
      { noteNo: '', label: `Closing ${trRecLbl}`, configId: 'bs_tradeReceivables', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.tradeAndOtherReceivables) },
    );
  }

  if (keep('bs_otherReceivables') || keep('bs_otherCurrentAssets')) {
    rows.push(
      { noteNo: '5.00', label: advLbl, configId: 'bs_otherReceivables', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'As Per Last Year', configId: 'bs_otherReceivables', indent: true, values: projData.map(d => d.note5AdvanceOpening) },
      { noteNo: '', label: 'Add: During the Year', configId: 'bs_otherReceivables', indent: true, values: projData.map(d => d.note5AdvanceAdded) },
      { noteNo: '', label: 'Subtotal', configId: 'bs_otherReceivables', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.note5AdvanceOpening + d.note5AdvanceAdded) },
      { noteNo: '', label: 'Less: Adjustment', configId: 'bs_otherReceivables', indent: true, values: projData.map(d => d.note5AdvanceAdjusted) },
      { noteNo: '', label: `Closing ${advLbl}`, configId: 'bs_otherReceivables', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.advanceDepositPrepayments) },
    );
  }

  if (keep('bs_cashAndEquivalents')) {
    rows.push(
      { noteNo: '6.00', label: cashLbl, configId: 'bs_cashAndEquivalents', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'Cash In Hand & Bank', configId: 'bs_cashAndEquivalents', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.cashAndCashEquivalents) },
    );
  }

  if (keep('bs_longTermBorrowings')) {
    rows.push({
      noteNo: '7.00',
      label: ltLoanLbl,
      configId: 'bs_longTermBorrowings',
      isBold: true,
      isSubtotal: true,
      values: projData.map(d => d.longTermLoan),
    });
  }

  if (keep('bs_tradePayables')) {
    rows.push(
      { noteNo: '8.00', label: trPayLbl, configId: 'bs_tradePayables', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'As Per Last Year', configId: 'bs_tradePayables', indent: true, values: projData.map(d => d.note8PayablesOpening) },
      { noteNo: '', label: 'Add: During the Year', configId: 'bs_tradePayables', indent: true, values: projData.map(d => d.note8PayablesAdded) },
      { noteNo: '', label: 'Subtotal', configId: 'bs_tradePayables', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.note8PayablesOpening + d.note8PayablesAdded) },
      { noteNo: '', label: 'Less: Payment During the Year', configId: 'bs_tradePayables', indent: true, values: projData.map(d => d.note8PayablesPaid) },
      { noteNo: '', label: `Closing ${trPayLbl}`, configId: 'bs_tradePayables', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.tradeAndOtherPayables) },
    );
  }

  if (keep('bs_currentTaxPayable')) {
    rows.push(
      { noteNo: '9.00', label: taxPayLbl, configId: 'bs_currentTaxPayable', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'As Per Last Year', configId: 'bs_currentTaxPayable', indent: true, values: projData.map(d => d.note9TaxOpening) },
      { noteNo: '', label: 'Add: During the Year', configId: 'bs_currentTaxPayable', indent: true, values: projData.map(d => d.note9TaxAdded) },
      { noteNo: '', label: 'Subtotal', configId: 'bs_currentTaxPayable', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.note9TaxOpening + d.note9TaxAdded) },
      { noteNo: '', label: 'Less: Adjustment', configId: 'bs_currentTaxPayable', indent: true, values: projData.map(d => d.note9TaxAdjusted) },
      { noteNo: '', label: `Closing ${taxPayLbl}`, configId: 'bs_currentTaxPayable', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.provisionsForTax) },
    );
  }

  if (keep('bs_shortTermBorrowings')) {
    rows.push(
      { noteNo: '10.00', label: stLoanLbl, configId: 'bs_shortTermBorrowings', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'Bank Loan', configId: 'bs_shortTermBorrowings', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.shortTermLoan) },
    );
  }

  if (keep('inc_costOfSales')) {
    rows.push(
      { noteNo: '11.00', label: cogsLbl, configId: 'inc_costOfSales', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'Opening Stock', configId: 'inc_costOfSales', indent: true, values: projData.map(d => d.note11OpeningStock) },
      { noteNo: '', label: 'Add: Purchase', configId: 'inc_costOfSales', indent: true, values: projData.map(d => d.note11Purchase) },
      { noteNo: '', label: 'Subtotal', configId: 'inc_costOfSales', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.note11OpeningStock + d.note11Purchase) },
      { noteNo: '', label: 'Less: Closing Stock', configId: 'inc_costOfSales', indent: true, values: projData.map(d => d.note11ClosingStock) },
      { noteNo: '', label: 'Material Consumed', configId: 'inc_costOfSales', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.note11OpeningStock + d.note11Purchase - d.note11ClosingStock) },
      { noteNo: '', label: 'Add: Direct Expenses', configId: 'inc_costOfSales', isSectionHeader: true, isBold: true, indent: true, values: nulls() },
      { noteNo: '', label: 'Labour & Wages', configId: 'inc_costOfSales', indent: true, values: projData.map(d => d.note11LabourWages) },
      { noteNo: '', label: 'Electricity Bill', configId: 'inc_costOfSales', indent: true, values: projData.map(d => d.note11ElectricityBill) },
      { noteNo: '', label: 'Carriage Inward', configId: 'inc_costOfSales', indent: true, values: projData.map(d => d.note11CarriageInward) },
      { noteNo: '', label: 'Other Direct Expenses', configId: 'inc_costOfSales', indent: true, values: projData.map(d => d.note11OtherDirectExpenses) },
      { noteNo: '', label: 'Total Direct Expenses', configId: 'inc_costOfSales', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.note11DirectExpensesTotal) },
      { noteNo: '', label: cogsLbl, configId: 'inc_costOfSales', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.costOfGoodsSold) },
    );
  }

  if (includeHidden || vis('inc_operatingExpenses') || vis('inc_depreciationExpense')) {
    rows.push({
      noteNo: '12.00',
      label: opexLbl,
      configId: 'inc_operatingExpenses',
      isBold: true,
      isSubtotal: true,
      values: projData.map(d =>
        r2(d.adminAndSellingExpenses + getVisibleCustomItemsSum(cfg, 'income', d.year, 'Operating Expenses'))
      ),
    });
  }

  if (keep('inc_financeCost')) {
    rows.push(
      { noteNo: '13.00', label: finCostLbl, configId: 'inc_financeCost', isSectionHeader: true, isBold: true, values: nulls() },
      { noteNo: '', label: 'Bank Charges & Interest', configId: 'inc_financeCost', isBold: true, isSubtotal: true, indent: true, values: projData.map(d => d.financialExpenses) },
    );
  }

  return rows;
}

// ---------------- MODEL VALIDATION SUITE (Requirement 5) ---------------- //
export function validateFinancialModel(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  equityStatements: EquityStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[],
  options?: StatementCalculationOptions
): import('../types/financial').FinancialValidationIssue[] {
  const issues: import('../types/financial').FinancialValidationIssue[] = [];
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const cellOverrides = options?.cellOverrides;

  const hasDeselectedIncome = (options?.statementConfig || []).some(
    c => (c.statement === 'income' || c.statement === 'fixedAssets') && !c.isVisible
  );
  const hasDeselectedFA = (options?.statementConfig || []).some(
    c => c.statement === 'fixedAssets' && !c.isVisible
  );
  const hasDeselectedCF = (options?.statementConfig || []).some(
    c => c.statement === 'cashFlow' && !c.isVisible
  );

  const histCount = inputs.filter(y => !y.isProjected).length;
  const projCount = inputs.filter(y => y.isProjected).length;

  if (histCount === 0) {
    issues.push({
      id: 'no-hist',
      severity: 'error',
      title: 'No Previous Financial Years Selected',
      message: 'Select at least one previous (historical) financial year to anchor calculations.',
    });
  }
  if (projCount === 0) {
    issues.push({
      id: 'no-proj',
      severity: 'warning',
      title: 'No Projected Years Selected',
      message: 'Select at least one projected financial year to generate forecast statements.',
    });
  }

  inputs.forEach((inp, idx) => {
    const inc = incomeStatements[idx];
    const bs = balanceSheets[idx];
    const cf = cashFlowStatements[idx];
    const eq = equityStatements[idx];
    const fa = fixedAssetsSchedule[idx];

    if (inp.salesRevenue < 0 || inp.closingInventory < 0 || inp.financeCost < 0 || inp.fixedAssetsGross < 0) {
      issues.push({
        id: `neg-driver-${inp.year}`,
        year: inp.year,
        severity: 'error',
        title: `Negative Asset/Revenue Input (${inp.year})`,
        message: 'Sales Revenue, Closing Inventory, Finance Cost, and Fixed Assets cannot be negative.',
      });
    }

    const hasIncomeYrOverride = Object.keys(cellOverrides || {}).some(
      k => (k.startsWith('income_') || k.startsWith('fa_')) && cellOverrides?.[k]?.[inp.year] !== undefined
    );
    if (!hasDeselectedIncome && !hasIncomeYrOverride && inc && Math.abs(inc.netProfit - inp.netProfit) > 0.01) {
      issues.push({
        id: `npat-mismatch-${inp.year}`,
        year: inp.year,
        severity: 'error',
        title: `NPAT Preservation Alert (${inp.year})`,
        message: `Income Statement Net Profit (${inc.netProfit}) differs from user input (${inp.netProfit}).`,
      });
    }

    const hasFAYrOverride = Object.keys(cellOverrides || {}).some(
      k => k.startsWith('fa_') && cellOverrides?.[k]?.[inp.year] !== undefined
    );
    if (!hasDeselectedFA && !hasFAYrOverride && fa && Math.abs(fa.totalCostClosing - inp.fixedAssetsGross) > 0.01) {
      issues.push({
        id: `fa-mismatch-${inp.year}`,
        year: inp.year,
        severity: 'error',
        title: `Fixed Assets Preservation Alert (${inp.year})`,
        message: `Schedule Closing Gross Cost (${fa.totalCostClosing}) differs from user input (${inp.fixedAssetsGross}).`,
      });
    }

    if (bs && Math.abs(bs.balanceCheck) > 1) {
      issues.push({
        id: `bs-imbalance-${inp.year}`,
        year: inp.year,
        severity: 'error',
        title: `Balance Sheet Imbalance (${inp.year})`,
        message: `Total Assets differ from Total Liabilities & Equity by ${bs.balanceCheck.toLocaleString()}.`,
      });
    }

    const hasBSorEqCustomization = (options?.statementConfig || []).some(
      c => (c.statement === 'balanceSheet' || c.statement === 'equity') && !c.isVisible
    );
    if (!hasBSorEqCustomization && eq && bs && Math.abs(eq.totalClosingEquity - bs.totalEquity) > 1) {
      issues.push({
        id: `eq-mismatch-${inp.year}`,
        year: inp.year,
        severity: 'error',
        title: `Equity Tie-Out Variance (${inp.year})`,
        message: `Statement of Changes in Equity (${eq.totalClosingEquity}) does not match Balance Sheet Equity (${bs.totalEquity}).`,
      });
    }

    if (cf && Math.abs(cf.reconciliationDifference) >= 0.01) {
      issues.push({
        id: `cf-variance-${inp.year}`,
        year: inp.year,
        severity: 'warning',
        title: `Cash Flow Reconciliation Variance (${inp.year})`,
        message: `Cash Flow closing cash differs from stated Balance Sheet cash by ${cf.reconciliationDifference.toLocaleString()}.`,
      });
    }

    if (inp.isScheduleLocked) {
      if (
        isConfigElementVisible(configMap, 'cf_borrow_repay') &&
        !hasDeselectedCF &&
        inp.schedulePrincipalLocked !== undefined &&
        cf &&
        Math.abs(cf.repaymentOfBorrowings - inp.schedulePrincipalLocked) > 0.01
      ) {
        issues.push({
          id: `sched-principal-lock-${inp.year}`,
          year: inp.year,
          severity: 'error',
          title: `Repayment Schedule Principal Mismatch (${inp.year})`,
          message: `Cash Flow Repayment of Borrowings (${cf.repaymentOfBorrowings}) differs from exact Schedule Inst. Principal (${inp.schedulePrincipalLocked}).`,
        });
      }
      if (
        isConfigElementVisible(configMap, 'inc_financeCost') &&
        inp.scheduleFinanceCostLocked !== undefined &&
        inc &&
        Math.abs(inc.financeCost - inp.scheduleFinanceCostLocked) > 0.01
      ) {
        issues.push({
          id: `sched-profit-lock-${inp.year}`,
          year: inp.year,
          severity: 'error',
          title: `Repayment Schedule Finance Cost Mismatch (${inp.year})`,
          message: `Income Statement Finance Cost (${inc.financeCost}) differs from exact Schedule Inst. Profit (${inp.scheduleFinanceCostLocked}).`,
        });
      }
    }
  });

  return issues;
}

// ---------------- DEBT SERVICING & REPAYMENT CAPACITY ANALYSIS (DSCR) ---------------- //
export function calculateDebtServiceCoverage(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  schedule?: import('../types/financial').ParsedRepaymentSchedule | null,
  options?: StatementCalculationOptions
): import('../types/financial').DebtServiceCoverageYear[] {
  const schedMap = new Map(
    (schedule?.annualSummaries || []).map(s => [s.financialYear, s])
  );
  const configMap = buildConfigVisibilityMap(options?.statementConfig);
  const vis = (id: string) => isConfigElementVisible(configMap, id);

  return inputs.map((inp, idx) => {
    const inc = incomeStatements[idx];
    const cf = cashFlowStatements[idx];
    const sched = schedMap.get(inp.year);

    const scheduledPrincipal = vis('cf_borrow_repay')
      ? (sched ? sched.totalPrincipal : (cf?.repaymentOfBorrowings ?? inp.repaymentOfBorrowings ?? 0))
      : 0;
    const scheduledFinanceCost = vis('inc_financeCost')
      ? (sched ? sched.totalProfit : (inc?.financeCost ?? inp.financeCost ?? 0))
      : 0;
    const totalDebtService = r2(scheduledPrincipal + scheduledFinanceCost);

    const ebitda = r2((inc?.operatingProfit ?? 0) + (inc?.depreciationExpense ?? 0));
    const operatingCashFlowBeforeFinance = r2((cf?.netCashFromOperating ?? 0) - (cf?.financeCostPaid ?? 0));
    const ebitdaDscr = totalDebtService > 0 ? Number((ebitda / totalDebtService).toFixed(2)) : 99.99;
    const cashDscr = totalDebtService > 0 ? Number((operatingCashFlowBeforeFinance / totalDebtService).toFixed(2)) : 99.99;
    const interestCoverage = scheduledFinanceCost > 0
      ? Number(((inc?.operatingProfit ?? 0) / scheduledFinanceCost).toFixed(2))
      : 99.99;

    const closingCash = cf?.closingCashCalculated ?? balanceSheets[idx]?.cashAndEquivalents ?? 0;
    const netCashAfterDebtService = r2(operatingCashFlowBeforeFinance - totalDebtService);

    let capacityStatus: 'STRONG' | 'ADEQUATE' | 'TIGHT' | 'DEFICIT' = 'STRONG';
    if (ebitdaDscr < 1.0 || cashDscr < 1.0 || closingCash < 0) {
      capacityStatus = 'DEFICIT';
    } else if (ebitdaDscr < 1.25 || cashDscr < 1.15) {
      capacityStatus = 'TIGHT';
    } else if (ebitdaDscr < 1.5 || cashDscr < 1.35) {
      capacityStatus = 'ADEQUATE';
    }

    const hasAnyScheduleContext = Boolean(schedule?.annualSummaries?.length) || inputs.some(y => y.isScheduleLocked);
    let reschedulingContext: string | undefined;
    if (hasAnyScheduleContext) {
      if (!inp.isProjected) {
        reschedulingContext =
          capacityStatus === 'DEFICIT' || capacityStatus === 'TIGHT'
            ? 'Pre-rescheduling operating cash / EBITDA shortfall under existing debt burden — justifies loan rescheduling'
            : 'Historical pre-rescheduling servicing baseline';
      } else {
        reschedulingContext = Boolean(sched || inp.isScheduleLocked)
          ? 'Post-rescheduling projected cash flow & profitability comfortably satisfy rescheduled installments'
          : 'Post-rescheduling projected repayment capacity headroom';
      }
    }

    return {
      year: inp.year,
      isProjected: inp.isProjected,
      scheduledPrincipal,
      scheduledFinanceCost,
      totalDebtService,
      ebitda,
      operatingCashFlowBeforeFinance,
      ebitdaDscr,
      cashDscr,
      interestCoverage,
      closingCash,
      netCashAfterDebtService,
      isScheduleCovered: Boolean(sched || inp.isScheduleLocked),
      capacityStatus,
      reschedulingContext,
    };
  });
}

// ---------------- INTELLIGENT REPAYMENT SCHEDULE IMPLEMENTATION & OPTIMIZATION ---------------- //
export interface ImplementScheduleParams {
  schedule: import('../types/financial').ParsedRepaymentSchedule;
  historicalYearsStore: Record<number, YearDataInput>;
  customProjectedOverrides: Record<number, YearDataInput>;
  assumptions: ProjectionAssumptions;
  autoIncludeScheduleYears?: boolean;
  alignHistoricalScale?: boolean;
}

export interface ImplementScheduleResult {
  historicalYearsStore: Record<number, YearDataInput>;
  customProjectedOverrides: Record<number, YearDataInput>;
  assumptions: ProjectionAssumptions;
  scaleFactorApplied: number;
  coveredYears: number[];
  minEbitdaDscr: number;
  minCashDscr: number;
}

/**
 * Implements an uploaded Repayment Schedule across Previous & Projected Financial Statements:
 * 1. Injects exact annual Inst. Principal -> Repayment from Borrowings & Financing (never altered)
 * 2. Injects exact annual Inst. Profit -> Finance Cost (never altered)
 * 3. Synchronizes selected years so every schedule year is covered
 * 4. Intelligently & proportionately optimizes operating performance & debt balances
 *    when needed so projected business cash flows sustainably cover all scheduled obligations.
 */
export function implementRepaymentScheduleIntoModel({
  schedule,
  historicalYearsStore,
  customProjectedOverrides,
  assumptions,
  autoIncludeScheduleYears = true,
  alignHistoricalScale = true,
}: ImplementScheduleParams): ImplementScheduleResult {
  const schedSummaries = schedule.annualSummaries;
  const schedMap = new Map(schedSummaries.map(s => [s.financialYear, s]));
  const schedYears = schedSummaries.map(s => s.financialYear).sort((a, b) => a - b);

  // 1. Resolve Historical & Projected Year sets
  let histYears = assumptions.selectedHistoricalYears && assumptions.selectedHistoricalYears.length > 0
    ? [...assumptions.selectedHistoricalYears].sort((a, b) => a - b)
    : INITIAL_HISTORICAL_INPUTS.map(h => h.year);
  const maxHistYear = histYears[histYears.length - 1] || 2025;

  let projYears = assumptions.selectedProjectedYears && assumptions.selectedProjectedYears.length > 0
    ? [...assumptions.selectedProjectedYears].sort((a, b) => a - b)
    : Array.from({ length: Math.max(1, assumptions.projectedYearsCount) }, (_, i) => maxHistYear + 1 + i);

  if (autoIncludeScheduleYears && schedYears.length > 0) {
    const histSet = new Set(histYears);
    const projSet = new Set(projYears);
    for (const sy of schedYears) {
      if (sy <= maxHistYear) {
        histSet.add(sy);
      } else {
        projSet.add(sy);
      }
    }
    histYears = Array.from(histSet).sort((a, b) => a - b);
    projYears = Array.from(projSet).filter(y => !histSet.has(y)).sort((a, b) => a - b);
  }

  // Ensure every historical year exists in store
  const nextHistStore: Record<number, YearDataInput> = { ...historicalYearsStore };
  const seedList = Object.values(nextHistStore).sort((a, b) => a.year - b.year);
  for (const hy of histYears) {
    if (!nextHistStore[hy]) {
      nextHistStore[hy] = generateYearDataForTargetYear(hy, false, seedList, assumptions);
      seedList.push(nextHistStore[hy]);
      seedList.sort((a, b) => a.year - b.year);
    }
  }

  // 2. Assess sustainable operating scale required to cover scheduled debt service cleanly
  const revRate = (assumptions.revenueGrowthRate || 8.5) / 100;
  const gmPct = Math.max(20, assumptions.grossMarginPct || 38.0);
  const opexPct = Math.max(5, assumptions.opexPctOfRevenue || 17.5);
  const taxPct = Math.max(0, assumptions.taxRatePct || 25.0);
  const depPct = Math.max(5, assumptions.depreciationRatePct || 10.0);
  const dso = assumptions.receivableDays || 45;
  const dio = assumptions.inventoryDays || 60;
  const dpo = assumptions.payableDays || 42;
  const capexPct = assumptions.capexPctOfRevenue || 6.2;
  const divPayout = Math.min(25, assumptions.dividendPayoutRatio || 30);

  const ebitdaMargin = Math.max(0.14, (gmPct - opexPct) / 100);
  const firstProjYear = projYears[0] || (maxHistYear + 1);
  const lastHistInput = nextHistStore[histYears[histYears.length - 1]] || INITIAL_HISTORICAL_INPUTS[INITIAL_HISTORICAL_INPUTS.length - 1];
  const projRevRateForYear = (yr: number): number => {
    const step = Math.max(1, yr - firstProjYear + 1);
    return computeProjectedRevenueGrowthPct(yr, assumptions.revenueGrowthRate || 8.5, step) / 100;
  };
  const cumRevMult = (fromYr: number, toYr: number): number => {
    if (toYr === fromYr) return 1;
    let mult = 1;
    if (toYr > fromYr) {
      for (let y = fromYr + 1; y <= toYr; y++) {
        mult *= 1 + projRevRateForYear(y);
      }
    } else {
      for (let y = toYr + 1; y <= fromYr; y++) {
        mult /= 1 + Math.max(0.02, projRevRateForYear(y));
      }
    }
    return mult;
  };
  const baselineProjRev = lastHistInput.salesRevenue * (1 + projRevRateForYear(firstProjYear));

  const investmentMade = schedule.metadata.investmentMade || schedule.grandTotalPrincipal || 0;
  let requiredAnchorProjRev = baselineProjRev;

  for (const s of schedSummaries) {
    const annualDS = s.totalPrincipal + s.totalProfit;
    // Sustainable business EBITDA must comfortably cover 1.55x Debt Service plus working capital / capex cushion
    const taxAllowance = Math.max(0, (annualDS * 1.48 - s.totalProfit) * (taxPct / 100));
    const reqEBITDA = Math.max(annualDS * 1.55 + taxAllowance * 0.45, investmentMade * 0.30);
    // Account for realistic cyclical GM & OpEx margin swings (-0.042 worst-case safety buffer)
    const reqRevThisYear = reqEBITDA / Math.max(0.10, ebitdaMargin - 0.042);
    const backDiscountedRev = reqRevThisYear / Math.max(0.1, cumRevMult(firstProjYear, s.financialYear));
    if (backDiscountedRev > requiredAnchorProjRev) {
      requiredAnchorProjRev = backDiscountedRev;
    }
  }

  const rawScaleRatio = baselineProjRev > 0 ? requiredAnchorProjRev / baselineProjRev : 1;
  const needsOptimizationLift = rawScaleRatio > 1.02;
  const targetAnchorProjRev = needsOptimizationLift ? requiredAnchorProjRev * 1.08 : baselineProjRev;
  const scaleFactorApplied = baselineProjRev > 0 ? Number((targetAnchorProjRev / baselineProjRev).toFixed(4)) : 1;

  // Optionally align historical baseline proportionally when there is a >1.5x unit/capacity scale jump
  // (e.g., default sample data in Thousands vs a client schedule in exact Units of BDT)
  const shouldScaleHist = alignHistoricalScale && scaleFactorApplied > 1.5;
  const hasFutureReschedule = schedSummaries.some(s => s.financialYear > maxHistYear);
  const firstSchedDS = schedSummaries[0] ? schedSummaries[0].totalPrincipal + schedSummaries[0].totalProfit : 0;
  // Constrained pre-rescheduling historical scale (~64% of turnaround anchor so historical operations reflect strained repayment ability before rescheduling recovery)
  const histRevenueScale = shouldScaleHist ? Math.max(1, scaleFactorApplied * 0.64) : 1;

  const buildOptimizedYear = (
    yr: number,
    isProjected: boolean,
    prevYear: YearDataInput | null,
    existingInput?: YearDataInput
  ): YearDataInput => {
    const sched = schedMap.get(yr);
    const nextYrSched = schedMap.get(yr + 1);
    const isDistressedHistorical = !isProjected && !sched && hasFutureReschedule;

    // Determine Sales Revenue on irregular non-uniform business curve
    let salesRevenue: number;
    if (!isProjected) {
      if (shouldScaleHist) {
        const baseHistSales = (existingInput?.salesRevenue || lastHistInput.salesRevenue) * histRevenueScale;
        salesRevenue = prevYear && prevYear.salesRevenue > 0
          ? computeSpreadValue(
              baseHistSales,
              prevYear.salesRevenue,
              Math.round(baseHistSales * 0.88),
              Math.round(baseHistSales * 1.12),
              [],
              yr,
              1,
              1.45
            )
          : Math.round(baseHistSales);
      } else {
        salesRevenue = existingInput?.salesRevenue || lastHistInput.salesRevenue;
      }
    } else {
      salesRevenue = Math.round(targetAnchorProjRev * cumRevMult(firstProjYear, yr));
      if (!needsOptimizationLift && customProjectedOverrides[yr] && !sched) {
        salesRevenue = customProjectedOverrides[yr].salesRevenue;
      }
    }
    const gRev = prevYear && prevYear.salesRevenue > 0 ? yoyPct(salesRevenue, prevYear.salesRevenue) : 0;

    // Productive Fixed Assets base scales with revenue & non-uniform capex/disposal cycles
    const prevGross = prevYear
      ? prevYear.fixedAssetsGross
      : Math.round(salesRevenue * (lastHistInput.fixedAssetsGross / Math.max(1, lastHistInput.salesRevenue)));

    let fixedAssetsGross: number;
    let fixedAssetsAdditions: number;
    let fixedAssetsDisposals: number;

    const isFirstProjYear = isProjected && yr === firstProjYear;

    if (!isProjected && existingInput && !shouldScaleHist) {
      fixedAssetsGross = existingInput.fixedAssetsGross;
      fixedAssetsAdditions = existingInput.fixedAssetsAdditions;
      fixedAssetsDisposals = existingInput.fixedAssetsDisposals;
    } else if (!isProjected && existingInput && shouldScaleHist) {
      fixedAssetsGross = Math.round(existingInput.fixedAssetsGross * histRevenueScale);
      fixedAssetsAdditions = Math.round(existingInput.fixedAssetsAdditions * histRevenueScale);
      fixedAssetsDisposals = Math.round(existingInput.fixedAssetsDisposals * histRevenueScale);
    } else {
      const disbBoost = !isFirstProjYear && sched && sched.disbursementAmount > 0 ? Math.round(sched.disbursementAmount * 0.45) : 0;
      const minCapexPct = Math.max(1.5, capexPct * 0.58);
      const maxCapexPct = Math.min(22, capexPct * 1.48);
      const yearCapexPct = clampVal(capexPct * (1 + deterministicAnnualWave(yr, 11) * 0.34), minCapexPct, maxCapexPct);
      const capexBase = prevYear && prevYear.fixedAssetsAdditions > 0
        ? computeSpreadValue(
            salesRevenue * (yearCapexPct / 100),
            prevYear.fixedAssetsAdditions,
            Math.round(salesRevenue * (minCapexPct / 100)),
            Math.round(salesRevenue * (maxCapexPct / 100)),
            [gRev],
            yr,
            11,
            1.6
          )
        : Math.round(salesRevenue * (yearCapexPct / 100));
      fixedAssetsAdditions = capexBase + disbBoost;
      const gAdd = prevYear && prevYear.fixedAssetsAdditions > 0 ? yoyPct(fixedAssetsAdditions, prevYear.fixedAssetsAdditions) : 0;
      const yearDispRate = varyBounded(0.011, yr, 12, 0.0058, 0.0045, 0.023);
      fixedAssetsDisposals = prevYear && prevYear.fixedAssetsDisposals > 0
        ? computeSpreadValue(
            prevGross * yearDispRate,
            prevYear.fixedAssetsDisposals,
            Math.round(prevGross * 0.0045),
            Math.round(prevGross * 0.023),
            [gAdd, gRev],
            yr,
            12,
            1.6
          )
        : Math.round(prevGross * yearDispRate);
      fixedAssetsGross = Math.max(100, prevGross + fixedAssetsAdditions - fixedAssetsDisposals);
    }

    const gGrossFA = prevYear && prevGross > 0 ? yoyPct(fixedAssetsGross, prevGross) : null;
    const prevOpeningGrossOpt = prevYear
      ? Math.max(0, prevGross - (prevYear.fixedAssetsAdditions || 0) + (prevYear.fixedAssetsDisposals || 0))
      : 0;
    const prevDepChargeOpt = prevYear && prevGross > 0
      ? Math.round(((prevOpeningGrossOpt + prevGross) / 2) * ((prevYear.depreciationRate || depPct) / 100))
      : null;
    const effectiveDepPct = isProjected
      ? computeProjectedDepRatePct(
          depPct,
          yr,
          gGrossFA,
          prevDepChargeOpt,
          prevGross,
          fixedAssetsGross
        )
      : (existingInput?.depreciationRate ?? depPct);

    const depCharge = Math.round(((prevGross + fixedAssetsGross) / 2) * (effectiveDepPct / 100));

    // Pre-rescheduling historical years reflect margin compression & higher cost burden vs post-rescheduling recovery
    const targetCosPct = isDistressedHistorical ? clampVal((100 - gmPct) + 7.2, 50, 79.5) : (100 - gmPct);
    const prevCosPct = prevYear && prevYear.salesRevenue > 0 ? (prevYear.costOfSales / prevYear.salesRevenue) * 100 : targetCosPct;
    const rawCosPct = varyBounded(
      prevYear ? prevCosPct * 0.25 + targetCosPct * 0.75 : targetCosPct,
      yr,
      2,
      2.1,
      Math.max(18, targetCosPct - 3.4),
      Math.min(84, targetCosPct + 3.4)
    );
    const costOfSales = prevYear
      ? computeSpreadValue(
          salesRevenue * (rawCosPct / 100),
          prevYear.costOfSales,
          Math.round(salesRevenue * (Math.max(18, targetCosPct - 3.4) / 100)),
          Math.round(salesRevenue * (Math.min(84, targetCosPct + 3.4) / 100)),
          [gRev],
          yr,
          2,
          1.55
        )
      : Math.round(salesRevenue * (rawCosPct / 100));
    const gCos = prevYear && prevYear.costOfSales > 0 ? yoyPct(costOfSales, prevYear.costOfSales) : 0;

    const targetOpexPct = isDistressedHistorical ? clampVal(opexPct + 2.8, 8, 34) : opexPct;
    const prevOpexPct = prevYear && prevYear.salesRevenue > 0 ? (prevYear.operatingExpenses / prevYear.salesRevenue) * 100 : targetOpexPct;
    const rawOpexPct = varyBounded(
      prevYear ? prevOpexPct * 0.25 + targetOpexPct * 0.75 : targetOpexPct,
      yr,
      3,
      1.75,
      Math.max(3, targetOpexPct - 2.4),
      Math.min(50, targetOpexPct + 2.4)
    );
    const minOpexVal = Math.round(salesRevenue * (Math.max(3, targetOpexPct - 2.4) / 100));
    const maxOpexVal = Math.round(salesRevenue * (Math.min(50, targetOpexPct + 2.4) / 100));
    let operatingExpenses = prevYear
      ? computeSpreadValue(
          salesRevenue * (rawOpexPct / 100),
          prevYear.operatingExpenses,
          minOpexVal,
          maxOpexVal,
          [gRev, gCos],
          yr,
          3,
          1.55
        )
      : Math.round(salesRevenue * (rawOpexPct / 100));
    const interestIncome = Math.round(salesRevenue * computeInterestIncomeRate(yr, isProjected));

    // Exact Schedule Locks for Principal & Profit on covered years; distressed existing obligations on prior historical years
    let exactFinanceCost: number;
    let exactRepayment: number;
    if (sched) {
      exactFinanceCost = sched.totalProfit;
      exactRepayment = sched.totalPrincipal;
    } else if (isDistressedHistorical) {
      // Strained pre-rescheduling debt obligations exceeding historical EBITDA (EBITDA DSCR ~0.72x - 0.90x)
      const histEBITDA = Math.max(1000, salesRevenue - costOfSales - operatingExpenses);
      const targetDistressRatio = varyBounded(1.26, yr, 19, 0.08, 1.15, 1.42);
      const preReschedFloorDS = Math.round(Math.max(histEBITDA * targetDistressRatio, firstSchedDS * 0.75));
      const finShare = varyBounded(0.41, yr, 14, 0.035, 0.34, 0.48);
      const rawDistFin = Math.max(100, Math.round(preReschedFloorDS * finShare));
      const rawDistRepay = Math.max(100, preReschedFloorDS - rawDistFin);
      exactFinanceCost = prevYear && prevYear.financeCost > 0
        ? computeSpreadValue(
            rawDistFin,
            prevYear.financeCost,
            Math.round(rawDistFin * 0.85),
            Math.round(rawDistFin * 1.15),
            [gRev, gCos],
            yr,
            14,
            1.45
          )
        : rawDistFin;
      const gHistFin = prevYear && prevYear.financeCost > 0 ? yoyPct(exactFinanceCost, prevYear.financeCost) : 0;
      exactRepayment = prevYear && prevYear.repaymentOfBorrowings > 0
        ? computeSpreadValue(
            rawDistRepay,
            prevYear.repaymentOfBorrowings,
            Math.round(rawDistRepay * 0.85),
            Math.round(rawDistRepay * 1.15),
            [gHistFin, gRev],
            yr,
            15,
            1.45
          )
        : rawDistRepay;
    } else if (existingInput && !isProjected) {
      exactFinanceCost = shouldScaleHist ? Math.round(existingInput.financeCost * histRevenueScale) : existingInput.financeCost;
      exactRepayment = shouldScaleHist ? Math.round(existingInput.repaymentOfBorrowings * histRevenueScale) : existingInput.repaymentOfBorrowings;
    } else {
      const targetProjRate = clampVal((assumptions.loanInterestRatePct ?? 15.5) / 100, 0, 0.38);
      const debtBase = prevYear ? prevYear.longTermBorrowings + prevYear.shortTermBorrowings : salesRevenue * 0.15;
      exactFinanceCost =
        debtBase > 0 && targetProjRate > 0
          ? Math.max(
              100,
              Math.round(
                debtBase *
                  varyBounded(
                    targetProjRate,
                    yr,
                    14,
                    Math.max(0.006, targetProjRate * 0.055),
                    Math.max(0.01, targetProjRate * 0.72),
                    Math.min(0.40, targetProjRate * 1.28)
                  )
              )
            )
          : 0;
      const inPreScheduleMoratorium = Boolean(
        isProjected && hasFutureReschedule && schedSummaries[0] && yr < schedSummaries[0].financialYear
      );
      exactRepayment = inPreScheduleMoratorium
        ? 0
        : Math.round((prevYear ? prevYear.longTermBorrowings : salesRevenue * 0.1) * varyBounded(0.12, yr, 15, 0.022, 0.065, 0.22));
    }

    // P&L forward calculation incorporating exact Finance Cost, keeping NP spread distinct from Rev/CoS/OpEx
    const calcNp = (opexArg: number): { pbt: number; tax: number; np: number } => {
      const opProf = salesRevenue - costOfSales - opexArg - depCharge;
      const pbtVal = r2(opProf + interestIncome - exactFinanceCost);
      const taxVal = pbtVal > 0 ? r2(pbtVal * (taxPct / 100)) : 0;
      const npVal = r2(pbtVal - taxVal);
      return { pbt: pbtVal, tax: taxVal, np: npVal };
    };

    let { np: computedNp } = calcNp(operatingExpenses);
    if (prevYear && prevYear.netProfit !== 0) {
      const gOp0 = yoyPct(operatingExpenses, prevYear.operatingExpenses);
      const gNp0 = yoyPct(computedNp, prevYear.netProfit);
      const minSp = Math.min(Math.abs(gNp0 - gRev), Math.abs(gNp0 - gCos), Math.abs(gNp0 - gOp0));
      if (minSp < 1.48) {
        let bestOpex = operatingExpenses;
        let bestDist = Infinity;
        let bestMinSp = minSp;
        const span = Math.max(1, maxOpexVal - minOpexVal);
        for (let i = 0; i <= 40; i++) {
          const candOpex = Math.round(minOpexVal + (span * i) / 40);
          const gO = yoyPct(candOpex, prevYear.operatingExpenses);
          const { np: candNp } = calcNp(candOpex);
          const gN = yoyPct(candNp, prevYear.netProfit);
          const sp = Math.min(
            Math.abs(gO - gRev),
            Math.abs(gO - gCos),
            Math.abs(gN - gRev),
            Math.abs(gN - gCos),
            Math.abs(gN - gO)
          );
          const dist = Math.abs(candOpex - operatingExpenses);
          if (sp >= 1.48) {
            if (bestMinSp < 1.48 || dist < bestDist) {
              bestOpex = candOpex;
              bestDist = dist;
              bestMinSp = sp;
            }
          } else if (bestMinSp < 1.48 && sp > bestMinSp) {
            bestOpex = candOpex;
            bestMinSp = sp;
          }
        }
        operatingExpenses = bestOpex;
        computedNp = calcNp(operatingExpenses).np;
      }
    }

    const netProfit = sched || isProjected || shouldScaleHist || isDistressedHistorical
      ? computedNp
      : (existingInput?.netProfit ?? computedNp);

    // Working capital items aligned logically to Sales / CoS with non-uniform turnover & >= 1.2pp spreads
    const prevDio = prevYear && prevYear.costOfSales > 0 ? (prevYear.closingInventory / prevYear.costOfSales) * 365 : dio;
    const minDio = Math.max(5, dio * 0.84);
    const maxDio = Math.min(220, dio * 1.16);
    const yearDio = varyBounded(prevYear ? prevDio * 0.30 + dio * 0.70 : dio, yr, 4, Math.max(3.0, dio * 0.095), minDio, maxDio);
    const closingInventory = (!isProjected && existingInput && !shouldScaleHist && !isDistressedHistorical)
      ? existingInput.closingInventory
      : prevYear
      ? computeSpreadValue(
          (costOfSales / 365) * yearDio,
          prevYear.closingInventory,
          Math.round((costOfSales / 365) * minDio),
          Math.round((costOfSales / 365) * maxDio),
          [gRev, gCos],
          yr,
          4,
          1.50
        )
      : Math.round((costOfSales / 365) * yearDio);
    const gInv = prevYear && prevYear.closingInventory > 0 ? yoyPct(closingInventory, prevYear.closingInventory) : 0;

    const prevDso = prevYear && prevYear.salesRevenue > 0 ? (prevYear.tradeReceivables / prevYear.salesRevenue) * 365 : dso;
    const minDso = Math.max(5, dso * 0.84);
    const maxDso = Math.min(180, dso * 1.16);
    const yearDso = varyBounded(prevYear ? prevDso * 0.30 + dso * 0.70 : dso, yr, 5, Math.max(2.2, dso * 0.095), minDso, maxDso);
    const tradeReceivables = prevYear
      ? computeSpreadValue(
          (salesRevenue / 365) * yearDso,
          prevYear.tradeReceivables,
          Math.round((salesRevenue / 365) * minDso),
          Math.round((salesRevenue / 365) * maxDso),
          [gRev, gInv, gCos],
          yr,
          5,
          1.50
        )
      : Math.round((salesRevenue / 365) * yearDso);
    const gTrRec = prevYear && prevYear.tradeReceivables > 0 ? yoyPct(tradeReceivables, prevYear.tradeReceivables) : 0;

    const otherRecRatio = varyBounded(0.16, yr, 7, 0.028, 0.11, 0.23);
    const otherReceivables = prevYear
      ? computeSpreadValue(
          tradeReceivables * otherRecRatio,
          prevYear.otherReceivables,
          Math.round(tradeReceivables * 0.105),
          Math.round(tradeReceivables * 0.235),
          [gTrRec, gInv, gRev],
          yr,
          7,
          1.50
        )
      : Math.round(tradeReceivables * otherRecRatio);
    const gOthRec = prevYear && prevYear.otherReceivables > 0 ? yoyPct(otherReceivables, prevYear.otherReceivables) : 0;

    const prevDpo = prevYear && prevYear.costOfSales > 0 ? (prevYear.tradePayables / prevYear.costOfSales) * 365 : dpo;
    const minDpo = Math.max(5, dpo * 0.84);
    const maxDpo = Math.min(180, dpo * 1.16);
    const yearDpo = varyBounded(prevYear ? prevDpo * 0.32 + dpo * 0.68 : dpo, yr, 6, Math.max(2.2, dpo * 0.095), minDpo, maxDpo);
    const tradePayables = prevYear
      ? computeSpreadValue(
          (costOfSales / 365) * yearDpo,
          prevYear.tradePayables,
          Math.round((costOfSales / 365) * minDpo),
          Math.round((costOfSales / 365) * maxDpo),
          [gCos, gInv, gTrRec],
          yr,
          6,
          1.50
        )
      : Math.round((costOfSales / 365) * yearDpo);
    const gTrPay = prevYear && prevYear.tradePayables > 0 ? yoyPct(tradePayables, prevYear.tradePayables) : 0;

    const otherPayRatio = varyBounded(0.25, yr, 8, 0.032, 0.18, 0.33);
    const otherPayables = prevYear
      ? computeSpreadValue(
          tradePayables * otherPayRatio,
          prevYear.otherPayables,
          Math.round(tradePayables * 0.175),
          Math.round(tradePayables * 0.335),
          [gTrPay, gCos],
          yr,
          8,
          1.50
        )
      : Math.round(tradePayables * otherPayRatio);

    const ocaGrowth = varyBounded(0.024, yr, 9, 0.056, -0.045, 0.088);
    const otherCurrentAssets = prevYear
      ? computeSpreadValue(
          prevYear.otherCurrentAssets * (1 + ocaGrowth),
          prevYear.otherCurrentAssets,
          Math.max(100, Math.round(prevYear.otherCurrentAssets * 0.955)),
          Math.round(prevYear.otherCurrentAssets * 1.088),
          [gTrRec, gOthRec, gInv],
          yr,
          9,
          1.50
        )
      : Math.round(salesRevenue * varyBounded(0.015, yr, 9, 0.004, 0.008, 0.025));

    // Borrowings & Financing Cash Flow synchronized cleanly with Schedule Balance & Disbursement
    const prevLT = prevYear
      ? prevYear.longTermBorrowings
      : Math.max(investmentMade || 0, Math.round((existingInput?.longTermBorrowings || 8000) * (shouldScaleHist ? histRevenueScale : 1)));
    const prevST = prevYear
      ? prevYear.shortTermBorrowings
      : Math.round((existingInput?.shortTermBorrowings || 2500) * (shouldScaleHist ? histRevenueScale : 1));
    const prevTotalDebt = r2(prevLT + prevST);

    let financingProceedsBorrowings = 0;
    if (isFirstProjYear) {
      // First projected year models existing loan rescheduling/restructuring with no fresh loan proceeds
      financingProceedsBorrowings = 0;
    } else if (sched) {
      if (sched.disbursementAmount > 0) {
        financingProceedsBorrowings = r2(sched.disbursementAmount + Math.round(salesRevenue * varyBounded(0.015, yr, 19, 0.004, 0.008, 0.023)));
      } else {
        financingProceedsBorrowings = 0;
      }
    } else if (isDistressedHistorical) {
      // Historical pre-rescheduling working capital / overdraft support prevents negative cash while DSCR remains in DEFICIT
      const cashNeedForPrincipalAndCapex = Math.round(exactRepayment * 0.95 + Math.max(0, fixedAssetsAdditions - fixedAssetsDisposals) * 0.5);
      financingProceedsBorrowings = Math.max(
        Math.round((existingInput?.financingProceedsBorrowings || 0) * (shouldScaleHist ? histRevenueScale : 1)),
        cashNeedForPrincipalAndCapex
      );
    } else if (existingInput && !isProjected) {
      financingProceedsBorrowings = Math.round((existingInput.financingProceedsBorrowings || 0) * (shouldScaleHist ? histRevenueScale : 1));
    } else {
      financingProceedsBorrowings = Math.round(prevST * varyBounded(0.12, yr, 19, 0.042, 0.055, 0.21));
    }

    // Ensure total debt never drops below remaining scheduled balance + stays non-negative
    if (!isFirstProjYear && prevTotalDebt + financingProceedsBorrowings < exactRepayment) {
      financingProceedsBorrowings = r2(exactRepayment - prevTotalDebt + (sched?.closingBalance || 0));
    }
    const preserveCustomHistLT = Boolean(!isProjected && existingInput && !shouldScaleHist && existingInput.hasCustomLongTermBorrowings);
    const preserveCustomHistST = Boolean(!isProjected && existingInput && !shouldScaleHist && existingInput.hasCustomShortTermBorrowings);
    const targetTotalDebt = Math.max(sched?.closingBalance ?? 0, r2(prevTotalDebt + financingProceedsBorrowings - exactRepayment));
    const currentMatOfLT = nextYrSched ? nextYrSched.totalPrincipal : 0;
    const stSalesFloor = Math.round(salesRevenue * varyBounded(0.03, yr, 17, 0.009, 0.015, 0.045));
    const shortTermBorrowings = preserveCustomHistST
      ? Math.max(0, existingInput!.shortTermBorrowings)
      : r2(Math.min(targetTotalDebt, Math.max(currentMatOfLT, stSalesFloor)));
    const longTermBorrowings = preserveCustomHistLT
      ? Math.max(0, existingInput!.longTermBorrowings)
      : r2(Math.max(0, targetTotalDebt - shortTermBorrowings));

    const yearDivPayout = isDistressedHistorical
      ? 0
      : (divPayout > 0 ? varyBounded(divPayout, yr, 13, Math.min(5.5, divPayout * 0.22), 2, 40) : 0);
    const gNp = prevYear && prevYear.netProfit > 0 ? yoyPct(netProfit, prevYear.netProfit) : 0;
    const gRepay = prevYear && prevYear.repaymentOfBorrowings > 0 ? yoyPct(exactRepayment, prevYear.repaymentOfBorrowings) : 0;
    const gFin = prevYear && prevYear.financeCost > 0 ? yoyPct(exactFinanceCost, prevYear.financeCost) : 0;
    const dividendsPaid =
      !isDistressedHistorical && divPayout > 0 && netProfit > 0
        ? prevYear && prevYear.dividendsPaid > 0
          ? computeSpreadValue(
              netProfit * (yearDivPayout / 100),
              prevYear.dividendsPaid,
              Math.round(netProfit * (Math.max(2, divPayout * 0.72) / 100)),
              Math.round(netProfit * (Math.min(40, divPayout * 1.28) / 100)),
              [gNp, gRepay, gFin],
              yr,
              13,
              1.50
            )
          : Math.round(netProfit * (yearDivPayout / 100))
        : 0;
    const shareCapital = prevYear
      ? prevYear.shareCapital
      : Math.round((existingInput?.shareCapital || 12000) * (shouldScaleHist ? histRevenueScale : 1));

    const cashFlowFromFinancing = r2(financingProceedsBorrowings - exactRepayment - dividendsPaid);
    // For the first chronological year, ensure cashAndEquivalentsBS >= estimated netChangeInCash so openingCash >= 0 & Cash Flow stays 100% RECONCILED
    const estFirstOcf = r2(Math.max(0, netProfit) + depCharge - Math.max(0, netProfit) * 0.12);
    const estFirstCapex = r2(fixedAssetsAdditions - fixedAssetsDisposals);
    const estFirstNetChange = r2(estFirstOcf - estFirstCapex + cashFlowFromFinancing);
    const cashAndEquivalentsBS = prevYear
      ? Math.max(0, r2(prevYear.cashAndEquivalentsBS + (isDistressedHistorical ? -Math.round(salesRevenue * 0.008) : Math.max(0, netProfit) * 0.35)))
      : Math.max(
          Math.round((existingInput?.cashAndEquivalentsBS || Math.round(salesRevenue * 0.085)) * (shouldScaleHist ? histRevenueScale : 1)),
          Math.round(estFirstNetChange + salesRevenue * 0.055),
          Math.round(salesRevenue * 0.05)
        );

    return {
      ...(existingInput || {}),
      year: yr,
      isProjected,
      salesRevenue,
      closingInventory,
      financeCost: exactFinanceCost,
      netProfit,
      repaymentOfBorrowings: exactRepayment,
      fixedAssetsGross,
      cashFlowFromFinancing,
      fixedAssetsAdditions,
      fixedAssetsDisposals,
      depreciationRate: effectiveDepPct,
      costOfSales,
      operatingExpenses,
      taxRate: taxPct,
      tradeReceivables,
      otherReceivables,
      tradePayables,
      otherPayables,
      cashAndEquivalentsBS,
      otherCurrentAssets,
      intangibleAssets: prevYear
        ? Math.max(0, prevYear.intangibleAssets - Math.min(prevYear.intangibleAssets, Math.round(salesRevenue * varyBounded(0.002, yr, 15, 0.0007, 0.0008, 0.0034))))
        : Math.round((existingInput?.intangibleAssets || 1400) * (shouldScaleHist ? histRevenueScale : 1)),
      otherNonCurrentAssets: prevYear
        ? prevYear.otherNonCurrentAssets
        : Math.round((existingInput?.otherNonCurrentAssets || 3900) * (shouldScaleHist ? histRevenueScale : 1)),
      shortTermBorrowings,
      longTermBorrowings,
      shareCapital,
      retainedEarningsOpening: prevYear
        ? undefined
        : existingInput?.retainedEarningsOpening !== undefined
        ? Math.round(existingInput.retainedEarningsOpening * (shouldScaleHist ? histRevenueScale : 1))
        : (shouldScaleHist || salesRevenue > 0 ? Math.round(6500 * (shouldScaleHist ? histRevenueScale : 1)) : 0),
      dividendsPaid,
      financingProceedsShares: 0,
      financingProceedsBorrowings,
      financingRepaymentBorrowings: -exactRepayment,
      financingDividendsPaid: -dividendsPaid,
      hasCustomLongTermBorrowings: existingInput?.hasCustomLongTermBorrowings,
      hasCustomShortTermBorrowings: existingInput?.hasCustomShortTermBorrowings,
      isScheduleLocked: Boolean(sched),
      schedulePrincipalLocked: sched ? sched.totalPrincipal : undefined,
      scheduleFinanceCostLocked: sched ? sched.totalProfit : undefined,
    };
  };

  // Assemble all years chronologically
  const updatedHistStore: Record<number, YearDataInput> = { ...nextHistStore };
  const updatedProjOverrides: Record<number, YearDataInput> = {};

  let prevBuilt: YearDataInput | null = null;
  for (const hy of histYears) {
    const built = buildOptimizedYear(hy, false, prevBuilt, nextHistStore[hy]);
    updatedHistStore[hy] = built;
    prevBuilt = built;
  }
  for (const py of projYears) {
    const existing = customProjectedOverrides[py];
    const built = buildOptimizedYear(py, true, prevBuilt, existing);
    updatedProjOverrides[py] = built;
    prevBuilt = built;
  }

  // Prudent Chartered Accountancy Dividend Restraint: if a transition projected year dips into negative closing cash while paying dividends, suspend/trim discretionary dividends in that year
  let allOrdered = [
    ...histYears.map(y => updatedHistStore[y]),
    ...projYears.map(y => updatedProjOverrides[y]),
  ];
  let fa = calculateScheduleOfFixedAssets(allOrdered);
  let inc = calculateIncomeStatement(allOrdered, fa);
  let cf = calculateCashFlowStatement(allOrdered, inc, fa);

  let dividendAdjusted = false;
  let cumCashGain = 0;
  for (let i = 0; i < allOrdered.length; i++) {
    const yObj = allOrdered[i];
    if (!yObj.isProjected) continue;
    const cfYear = cf[i];
    if (!cfYear) continue;
    const projectedCashWithPrevSavings = r2(cfYear.closingCashCalculated + cumCashGain);
    if (projectedCashWithPrevSavings < 0 && yObj.dividendsPaid > 0) {
      const cashNeeded = Math.abs(projectedCashWithPrevSavings) + Math.max(1000, Math.round(yObj.salesRevenue * 0.005));
      const newDiv = cashNeeded >= yObj.dividendsPaid ? 0 : Math.max(0, r2(yObj.dividendsPaid - cashNeeded));
      const saved = r2(yObj.dividendsPaid - newDiv);
      cumCashGain = r2(cumCashGain + saved);
      updatedProjOverrides[yObj.year] = {
        ...updatedProjOverrides[yObj.year],
        dividendsPaid: newDiv,
        financingDividendsPaid: -newDiv,
        cashFlowFromFinancing: r2(
          (updatedProjOverrides[yObj.year].financingProceedsBorrowings || 0) -
            (updatedProjOverrides[yObj.year].repaymentOfBorrowings || 0) -
            newDiv
        ),
      };
      dividendAdjusted = true;
    }
  }

  if (dividendAdjusted) {
    allOrdered = [
      ...histYears.map(y => updatedHistStore[y]),
      ...projYears.map(y => updatedProjOverrides[y]),
    ];
    fa = calculateScheduleOfFixedAssets(allOrdered);
    inc = calculateIncomeStatement(allOrdered, fa);
    cf = calculateCashFlowStatement(allOrdered, inc, fa);
  }

  const eq = calculateEquityStatement(allOrdered, inc);
  const bs = calculateBalanceSheet(allOrdered, eq, fa, cf);
  const dscrList = calculateDebtServiceCoverage(allOrdered, inc, bs, cf, schedule);
  const coveredList = dscrList.filter(d => d.isScheduleCovered);
  const minEbitdaDscr = coveredList.length > 0 ? Math.min(...coveredList.map(d => d.ebitdaDscr)) : 0;
  const minCashDscr = coveredList.length > 0 ? Math.min(...coveredList.map(d => d.cashDscr)) : 0;

  return {
    historicalYearsStore: updatedHistStore,
    customProjectedOverrides: updatedProjOverrides,
    assumptions: {
      ...assumptions,
      historicalYearsCount: histYears.length,
      selectedHistoricalYears: histYears,
      selectedProjectedYears: projYears,
      projectedYearsCount: projYears.length,
    },
    scaleFactorApplied,
    coveredYears: schedYears,
    minEbitdaDscr,
    minCashDscr,
  };
}


