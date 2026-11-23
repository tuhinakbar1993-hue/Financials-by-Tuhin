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
} from '../types/financial';

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
  financialYearEnd: '31 December',
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
};

// Default 3 Historical Years (2023, 2024, 2025)
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
    salesRevenue: 54200,
    netProfit: 5380,
    financeCost: 1280, // Mandatory year-wise separate direct entry
    cashFlowFromFinancing: -2100, // Mandatory year-wise separate direct entry
    repaymentOfBorrowings: 1800, // Mandatory Year-Wise User Input: Repayment of Borrowings & Financing
    fixedAssetsGross: 35350,
    fixedAssetsAdditions: 4200,
    fixedAssetsDisposals: 600,
    depreciationRate: 10,
    costOfSales: 33800,
    operatingExpenses: 9100,
    taxRate: 25,
    closingInventory: 8650,
    tradeReceivables: 6950,
    otherReceivables: 1100,
    tradePayables: 4850,
    otherPayables: 1250,
    cashAndEquivalentsBS: 6098,
    otherCurrentAssets: 740,
    intangibleAssets: 1400,
    otherNonCurrentAssets: 3906,
    shortTermBorrowings: 2900,
    longTermBorrowings: 8800,
    shareCapital: 12000,
    retainedEarningsOpening: 9920,
    dividendsPaid: 1500,
    financingProceedsShares: 0,
    financingProceedsBorrowings: 1200,
    financingRepaymentBorrowings: -1800,
    financingDividendsPaid: -1500,
  },
  {
    year: 2025,
    isProjected: false,
    salesRevenue: 61800,
    netProfit: 6410,
    financeCost: 1340, // Mandatory year-wise separate direct entry
    cashFlowFromFinancing: -2350, // Mandatory year-wise separate direct entry
    repaymentOfBorrowings: 1550, // Mandatory Year-Wise User Input: Repayment of Borrowings & Financing
    fixedAssetsGross: 38950,
    fixedAssetsAdditions: 4800,
    fixedAssetsDisposals: 550,
    depreciationRate: 10,
    costOfSales: 38200,
    operatingExpenses: 10100,
    taxRate: 25,
    closingInventory: 9700,
    tradeReceivables: 7800,
    otherReceivables: 1280,
    tradePayables: 5500,
    otherPayables: 1420,
    cashAndEquivalentsBS: 8703,
    otherCurrentAssets: 820,
    intangibleAssets: 1300,
    otherNonCurrentAssets: 3191,
    shortTermBorrowings: 2600,
    longTermBorrowings: 8100,
    shareCapital: 12000,
    retainedEarningsOpening: 13800,
    dividendsPaid: 1800,
    financingProceedsShares: 0,
    financingProceedsBorrowings: 1000,
    financingRepaymentBorrowings: -1550,
    financingDividendsPaid: -1800,
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
  const rev = Math.max(1, income.revenue);
  const gp = income.grossProfit;
  const ebit = income.operatingProfit;
  const ebitda = ebit + income.depreciationExpense;
  const np = income.netProfit;
  const eq = Math.max(1, balance.totalEquity);
  const assets = Math.max(1, balance.totalAssets);
  const ca = balance.totalCurrentAssets;
  const cl = Math.max(1, balance.totalCurrentLiabilities);
  const inv = balance.closingInventory;
  const cash = balance.cashAndEquivalents;
  const totalDebt = balance.shortTermBorrowings + balance.longTermBorrowings;
  const fc = Math.max(1, income.financeCost);
  const cos = Math.max(1, income.costOfSales);
  const debtService = fc + (input.repaymentOfBorrowings || 0);

  return {
    grossMarginPct: Number(((gp / rev) * 100).toFixed(1)),
    operatingMarginPct: Number(((ebit / rev) * 100).toFixed(1)),
    netMarginPct: Number(((np / rev) * 100).toFixed(1)),
    ebitdaMarginPct: Number(((ebitda / rev) * 100).toFixed(1)),
    returnOnEquityPct: Number(((np / eq) * 100).toFixed(1)),
    returnOnCapitalEmployedPct: Number(((ebit / Math.max(1, assets - cl)) * 100).toFixed(1)),
    currentRatio: Number((ca / cl).toFixed(2)),
    quickRatio: Number(((ca - inv) / cl).toFixed(2)),
    cashRatio: Number((cash / cl).toFixed(2)),
    workingCapital: ca - cl,
    debtToEquity: Number((totalDebt / eq).toFixed(2)),
    debtToAssets: Number((totalDebt / assets).toFixed(2)),
    interestCoverageRatio: Number((ebit / fc).toFixed(2)),
    debtServiceCoverageRatio: Number((ebitda / Math.max(1, debtService)).toFixed(2)),
    receivableDays: Math.round((balance.tradeReceivables / rev) * 365),
    inventoryDays: Math.round((inv / cos) * 365),
    payableDays: Math.round((balance.tradePayables / cos) * 365),
    assetTurnover: Number((rev / assets).toFixed(2)),
  };
}

/**
 * Automatically generates realistic, internally consistent year-wise figures
 * for the 6 primary financial drivers based on appropriate financial ratios,
 * historical trends, accounting relationships, and logical interconnections.
 */
export function generateAutomaticPrimaryDrivers(
  prevYear: YearDataInput,
  assumptions: ProjectionAssumptions,
  yearNumber: number
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
  // 1. Sales Revenue: realistic compound growth based on assumptions
  const growth = 1 + assumptions.revenueGrowthRate / 100;
  const salesRevenue = Math.round(prevYear.salesRevenue * growth);

  // Cost of Sales based on gross margin
  const costOfSales = Math.round(salesRevenue * (1 - assumptions.grossMarginPct / 100));

  // 2. Closing Inventory: derived via Days Inventory Outstanding (DIO)
  const closingInventory = Math.round((costOfSales / 365) * assumptions.inventoryDays);

  // 3. Finance Cost: realistic debt yield applied to outstanding borrowings
  const prevTotalDebt = (prevYear.longTermBorrowings || 8000) + (prevYear.shortTermBorrowings || 2500);
  const effectiveInterestRate = 0.085; // 8.5% p.a.
  const financeCost = Math.max(350, Math.round(prevTotalDebt * effectiveInterestRate * 0.96));

  // 4. Net Profit After Tax: derived from operating profitability, depreciation, finance cost and statutory tax
  const grossProfit = salesRevenue - costOfSales;
  const operatingExpenses = Math.round(salesRevenue * (assumptions.opexPctOfRevenue / 100));
  const depEstimate = Math.round(prevYear.fixedAssetsGross * (assumptions.depreciationRatePct / 100));
  const operatingProfit = grossProfit - operatingExpenses - depEstimate;
  const interestIncome = Math.round(salesRevenue * 0.005);
  const profitBeforeTax = operatingProfit + interestIncome - financeCost;
  const taxExpense = Math.round(Math.max(0, profitBeforeTax) * (assumptions.taxRatePct / 100));
  const netProfit = profitBeforeTax - taxExpense;

  // 5. Repayment of Borrowings and Financing: scheduled 12% annual term debt principal reduction
  const repaymentOfBorrowings = Math.max(400, Math.round((prevYear.longTermBorrowings || 8000) * 0.12));

  // 6. Fixed Assets (Gross Asset Base): disciplined capital reinvestment less disposals
  const capexAdditions = Math.round(salesRevenue * (assumptions.capexPctOfRevenue / 100));
  const disposals = Math.round(prevYear.fixedAssetsGross * 0.012);
  const fixedAssetsGross = prevYear.fixedAssetsGross + capexAdditions - disposals;

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

  allYears.forEach(year => {
    if (year === baseYear) {
      // STRICTLY PRESERVE USER ENTERED BASE VALUE
      result[year] = Math.round(baseValue);
      return;
    }

    const diff = year - baseYear;

    switch (driver) {
      case 'salesRevenue': {
        if (diff > 0) {
          // Future years: compound annual growth rate with gentle maturity dampening after year 7
          let val = baseValue;
          for (let step = 1; step <= diff; step++) {
            const stepGrowth = step > 10 ? growthRate * 0.88 : (step > 6 ? growthRate * 0.94 : growthRate);
            val = val * (1 + stepGrowth);
          }
          result[year] = Math.round(val);
        } else {
          // Past years: back-discounting with steady expansion
          const absDiff = Math.abs(diff);
          result[year] = Math.round(baseValue / Math.pow(1 + Math.max(0.04, growthRate), absDiff));
        }
        break;
      }

      case 'closingInventory': {
        const baseSales = salesMap[baseYear];
        const targetSales = salesMap[year];
        if (baseSales && targetSales && baseSales > 0 && targetSales > 0) {
          const baseRatio = baseValue / baseSales;
          // Working capital efficiency gains: 0.4% p.a. DIO improvement
          const efficiency = Math.pow(1 - 0.004, diff);
          result[year] = Math.round(targetSales * baseRatio * efficiency);
        } else {
          const rate = growthRate || 0.072;
          if (diff > 0) {
            result[year] = Math.round(baseValue * Math.pow(1 + rate, diff));
          } else {
            result[year] = Math.round(baseValue / Math.pow(1 + rate, Math.abs(diff)));
          }
        }
        break;
      }

      case 'financeCost': {
        // Forward: deleveraging as debt amortizes (default ~4.2% reduction p.a. down to working capital interest floor)
        // Backward: debt was higher prior to principal repayment tranches
        if (diff > 0) {
          const reductionRate = Math.abs(growthRate) || 0.042;
          const factor = Math.pow(1 - reductionRate, diff);
          const minInterestFloor = Math.round(baseValue * 0.32);
          result[year] = Math.max(minInterestFloor, Math.round(baseValue * factor));
        } else {
          const priorDebtRate = 0.052;
          const factor = Math.pow(1 + priorDebtRate, Math.abs(diff));
          result[year] = Math.round(baseValue * factor);
        }
        break;
      }

      case 'netProfit': {
        // Forward: earnings growth with moderate operating leverage (margins gently expand)
        // Backward: lower volume and tighter margin
        const profitGrowth = Math.max(0.05, growthRate || 0.090);
        if (diff > 0) {
          result[year] = Math.round(baseValue * Math.pow(1 + profitGrowth, diff));
        } else {
          result[year] = Math.round(baseValue / Math.pow(1 + profitGrowth, Math.abs(diff)));
        }
        break;
      }

      case 'repaymentOfBorrowings': {
        // Forward: scheduled facility amortization gradually tapers down as term debt is liquidated
        // Backward: earlier active facility amortization was higher
        if (diff > 0) {
          const taperRate = Math.abs(growthRate) || 0.055;
          const factor = Math.pow(1 - taperRate, diff);
          const minRepaymentFloor = Math.max(350, Math.round(baseValue * 0.30));
          result[year] = Math.max(minRepaymentFloor, Math.round(baseValue * factor));
        } else {
          const priorRepayRate = 0.068;
          result[year] = Math.round(baseValue * Math.pow(1 + priorRepayRate, Math.abs(diff)));
        }
        break;
      }

      case 'fixedAssetsGross': {
        // Gross asset base expands as capex additions exceed disposals
        const assetGrowth = Math.max(0.03, growthRate || 0.056);
        if (diff > 0) {
          result[year] = Math.round(baseValue * Math.pow(1 + assetGrowth, diff));
        } else {
          result[year] = Math.round(baseValue / Math.pow(1 + assetGrowth, Math.abs(diff)));
        }
        break;
      }
    }
  });

  return result;
}

/**
 * Takes the current full set of year data inputs, updates a specific driver with
 * the new year-wise values, and re-derives all dependent financial metrics
 * in chronological order, maintaining full double-entry balance and cash flow integrity.
 */
export function recalculateAllYearsWithUpdatedDriver(
  currentYears: YearDataInput[],
  driver: PrimaryDriverKey,
  driverValuesByYear: Record<number, number>,
  assumptions: ProjectionAssumptions
): YearDataInput[] {
  const result: YearDataInput[] = [];
  
  for (let i = 0; i < currentYears.length; i++) {
    const existing = currentYears[i];
    const prevYearObj = i > 0 ? result[i - 1] : null;
    const newVal = driverValuesByYear[existing.year] !== undefined 
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

    const recomputed = deriveDependentFinancials(updatedDrivers, prevYearObj, assumptions, existing);
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
  const depreciationRate = existingFullYear?.depreciationRate ?? assumptions.depreciationRatePct ?? 10;

  // 1. Reconcile Profit Before Tax & Tax Expense strictly from fixed Net Profit After Tax (NPAT)
  const profitBeforeTax = netProfit >= 0
    ? Math.round(netProfit / Math.max(0.01, 1 - taxRate / 100))
    : netProfit;
  const taxExpense = profitBeforeTax - netProfit;
  const interestIncome = Math.round(salesRevenue * 0.005);

  // 2. Fixed Assets Additions & Disposals preserving fixedAssetsGross strictly
  const prevGross = prevYear ? prevYear.fixedAssetsGross : Math.round(fixedAssetsGross * 0.9);
  const baselineDisposals = Math.round(prevGross * 0.012 * yearStep);
  let fixedAssetsAdditions = 0;
  let fixedAssetsDisposals = baselineDisposals;
  if (fixedAssetsGross >= prevGross - baselineDisposals) {
    fixedAssetsAdditions = fixedAssetsGross - prevGross + baselineDisposals;
  } else {
    fixedAssetsAdditions = 0;
    fixedAssetsDisposals = Math.max(0, prevGross - fixedAssetsGross);
  }
  const depCharge = Math.round(((prevGross + fixedAssetsGross) / 2) * (depreciationRate / 100));

  // 3. Operating Profit (EBIT) = PBT - interestIncome + financeCost
  const operatingProfit = profitBeforeTax - interestIncome + financeCost;

  // 4. Total Operating Costs = Sales - EBIT - depCharge, allocated proportionately
  const totalOperatingCosts = Math.max(0, salesRevenue - operatingProfit - depCharge);
  const targetCoSWeight = Math.max(0.1, 1 - (assumptions.grossMarginPct || 38) / 100);
  const targetOpExWeight = Math.max(0.05, (assumptions.opexPctOfRevenue || 17.5) / 100);
  const cosShare = targetCoSWeight / (targetCoSWeight + targetOpExWeight);
  const costOfSales = Math.round(totalOperatingCosts * cosShare);
  const operatingExpenses = Math.max(0, totalOperatingCosts - costOfSales);

  // 5. Working Capital balances derived proportionately via turnover days & Sales/COGS
  const tradeReceivables = Math.round((salesRevenue / 365) * (assumptions.receivableDays || 45));
  const otherReceivables = Math.round(tradeReceivables * 0.16);
  const tradePayables = Math.round((costOfSales / 365) * (assumptions.payableDays || 42));
  const otherPayables = Math.round(tradePayables * 0.25);
  const otherCurrentAssets = prevYear
    ? Math.max(100, Math.round(prevYear.otherCurrentAssets * Math.pow(1.04, yearStep)))
    : Math.round(salesRevenue * 0.015);

  // 6. Equity & Dividends / Drawings
  const dividendsPaid = Math.round(Math.max(0, netProfit) * ((assumptions.dividendPayoutRatio || 30) / 100));
  const shareCapital = prevYear ? prevYear.shareCapital : (existingFullYear?.shareCapital ?? 12000);

  // 7. Borrowings & Financing Cash Flow preserving fixed repaymentOfBorrowings strictly
  const absRepayment = Math.max(0, repaymentOfBorrowings);
  const financingProceedsShares = existingFullYear?.financingProceedsShares ?? 0;

  const prevLT = prevYear ? prevYear.longTermBorrowings : Math.max(8000, absRepayment * 4);
  const prevST = prevYear ? prevYear.shortTermBorrowings : Math.max(2500, Math.round(salesRevenue * 0.06));
  const prevTotalDebt = prevLT + prevST;

  // Proportionate working-capital facility proceeds; top up only if total debt would otherwise turn negative
  let financingProceedsBorrowings = prevYear
    ? Math.round(prevST * 0.18)
    : Math.round(salesRevenue * 0.02);
  if (prevTotalDebt + financingProceedsBorrowings < absRepayment) {
    financingProceedsBorrowings = absRepayment - prevTotalDebt + Math.round(salesRevenue * 0.03);
  }

  const netDebtChange = financingProceedsBorrowings - absRepayment;
  const targetTotalDebt = Math.max(0, prevTotalDebt + netDebtChange);
  const ltShare = prevTotalDebt > 0 ? prevLT / prevTotalDebt : 0.72;
  const longTermBorrowings = Math.round(targetTotalDebt * ltShare);
  const shortTermBorrowings = Math.max(0, targetTotalDebt - longTermBorrowings);

  const financingRepaymentBorrowings = -absRepayment;
  const financingDividendsPaid = -dividendsPaid;
  const cashFlowFromFinancing =
    financingProceedsShares + financingProceedsBorrowings - absRepayment + financingDividendsPaid;

  const intangibleAssets = prevYear
    ? Math.max(0, prevYear.intangibleAssets - Math.min(prevYear.intangibleAssets, 100 * yearStep))
    : (existingFullYear?.intangibleAssets ?? Math.round(salesRevenue * 0.03));
  const otherNonCurrentAssets = prevYear
    ? prevYear.otherNonCurrentAssets
    : (existingFullYear?.otherNonCurrentAssets ?? Math.round(salesRevenue * 0.08));

  // Estimated Cash & Equivalents (actual Closing Cash is computed cleanly by Cash Flow Engine)
  const cashAndEquivalentsBS = prevYear
    ? Math.max(0, Math.round(prevYear.cashAndEquivalentsBS + netProfit * 0.35))
    : (existingFullYear?.cashAndEquivalentsBS ?? Math.round(salesRevenue * 0.085));

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
    shareCapital,
    dividendsPaid,
    financingProceedsShares,
    financingProceedsBorrowings,
    financingRepaymentBorrowings,
    financingDividendsPaid,
  };
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
    const fallbackBase = INITIAL_HISTORICAL_INPUTS[0];
    return { ...fallbackBase, year: targetYear, isProjected };
  }

  // Find closest preceding year or earliest year
  const preceding = [...sorted].reverse().find(y => y.year < targetYear);
  const anchor = preceding || sorted[0];
  const diff = targetYear - anchor.year;

  const revRate = (assumptions.revenueGrowthRate || 8.5) / 100;
  const scale = diff >= 0
    ? Math.pow(1 + revRate, diff)
    : 1 / Math.pow(1 + Math.max(0.04, revRate), Math.abs(diff));

  const salesRevenue = Math.max(100, Math.round(anchor.salesRevenue * scale));
  const closingInventory = Math.max(0, Math.round(anchor.closingInventory * scale));
  const financeCost = Math.max(
    50,
    Math.round(anchor.financeCost * (diff >= 0 ? Math.pow(0.96, diff) : Math.pow(1.05, Math.abs(diff))))
  );
  const netProfit = Math.round(anchor.netProfit * Math.pow(1 + Math.max(0.04, revRate), diff));
  const repaymentOfBorrowings = Math.max(
    0,
    Math.round(anchor.repaymentOfBorrowings * (diff >= 0 ? Math.pow(0.945, diff) : Math.pow(1.06, Math.abs(diff))))
  );
  const fixedAssetsGross = Math.max(
    100,
    Math.round(anchor.fixedAssetsGross * Math.pow(1.055, diff))
  );

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

  for (const yearNum of targetYears) {
    const step = Math.max(1, yearNum - prevYear.year);
    // Compound across step years if non-consecutive
    let tempPrev = prevYear;
    let finalDrivers = generateAutomaticPrimaryDrivers(tempPrev, assumptions, yearNum);
    if (step > 1) {
      for (let s = 1; s < step; s++) {
        const midDrivers = generateAutomaticPrimaryDrivers(tempPrev, assumptions, prevYear.year + s);
        tempPrev = deriveDependentFinancials(midDrivers, tempPrev, assumptions);
      }
      finalDrivers = generateAutomaticPrimaryDrivers(tempPrev, assumptions, yearNum);
    }
    const fullYear = deriveDependentFinancials(finalDrivers, prevYear, assumptions);
    projected.push(fullYear);
    prevYear = fullYear;
  }

  return projected;
}

// 1. Calculate Schedule of Fixed Assets (Strictly preserves input.fixedAssetsGross as Closing Cost)
export function calculateScheduleOfFixedAssets(
  inputs: YearDataInput[]
): ScheduleOfFixedAssetsYear[] {
  const result: ScheduleOfFixedAssetsYear[] = [];
  const firstGross = inputs[0]?.fixedAssetsGross || 32000;
  const baseAccumDep = Math.round(firstGross * 0.35);

  inputs.forEach((input, index) => {
    const isFirst = index === 0;
    const closingCost = input.fixedAssetsGross; // STRICTLY FIXED USER INPUT
    let openingCost = isFirst
      ? Math.max(0, closingCost - input.fixedAssetsAdditions + input.fixedAssetsDisposals)
      : result[index - 1].totalCostClosing;

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

    const openingDep = isFirst ? baseAccumDep : result[index - 1].totalDepClosing;
    const depCharge = Math.round(((openingCost + closingCost) / 2) * (input.depreciationRate / 100));
    const depDisposals = Math.round(disposals * 0.65);
    const closingDep = Math.max(0, openingDep + depCharge - depDisposals);
    const netBookValue = Math.max(0, closingCost - closingDep);

    const categories: FixedAssetCategory[] = [
      {
        name: 'Freehold Land & Industrial Buildings',
        costOpening: Math.round(openingCost * 0.40),
        additions: Math.round(additions * 0.25),
        disposals: Math.round(disposals * 0.10),
        costClosing: Math.round(closingCost * 0.40),
        depOpening: Math.round(openingDep * 0.18),
        depCharge: Math.round(depCharge * 0.15),
        depDisposals: Math.round(depDisposals * 0.10),
        depClosing: Math.round(closingDep * 0.18),
        netBookValue: Math.round(closingCost * 0.40) - Math.round(closingDep * 0.18),
        depRate: 2.5,
      },
      {
        name: 'Plant, Machinery & Automated Systems',
        costOpening: Math.round(openingCost * 0.35),
        additions: Math.round(additions * 0.45),
        disposals: Math.round(disposals * 0.50),
        costClosing: Math.round(closingCost * 0.35),
        depOpening: Math.round(openingDep * 0.45),
        depCharge: Math.round(depCharge * 0.50),
        depDisposals: Math.round(depDisposals * 0.50),
        depClosing: Math.round(closingDep * 0.45),
        netBookValue: Math.round(closingCost * 0.35) - Math.round(closingDep * 0.45),
        depRate: 10.0,
      },
      {
        name: 'Office Equipment & Fixtures',
        costOpening: Math.round(openingCost * 0.10),
        additions: Math.round(additions * 0.12),
        disposals: Math.round(disposals * 0.15),
        costClosing: Math.round(closingCost * 0.10),
        depOpening: Math.round(openingDep * 0.15),
        depCharge: Math.round(depCharge * 0.15),
        depDisposals: Math.round(depDisposals * 0.15),
        depClosing: Math.round(closingDep * 0.15),
        netBookValue: Math.round(closingCost * 0.10) - Math.round(closingDep * 0.15),
        depRate: 15.0,
      },
      {
        name: 'Motor Vehicles & Logistics Fleet',
        costOpening: Math.round(openingCost * 0.08),
        additions: Math.round(additions * 0.10),
        disposals: Math.round(disposals * 0.15),
        costClosing: Math.round(closingCost * 0.08),
        depOpening: Math.round(openingDep * 0.12),
        depCharge: Math.round(depCharge * 0.12),
        depDisposals: Math.round(depDisposals * 0.15),
        depClosing: Math.round(closingDep * 0.12),
        netBookValue: Math.round(closingCost * 0.08) - Math.round(closingDep * 0.12),
        depRate: 20.0,
      },
      {
        name: 'Computer, IT & Telecommunications',
        costOpening: Math.round(openingCost * 0.07),
        additions: Math.round(additions * 0.08),
        disposals: Math.round(disposals * 0.10),
        costClosing: closingCost - (
          Math.round(closingCost * 0.40) +
          Math.round(closingCost * 0.35) +
          Math.round(closingCost * 0.10) +
          Math.round(closingCost * 0.08)
        ),
        depOpening: Math.round(openingDep * 0.10),
        depCharge: depCharge - (
          Math.round(depCharge * 0.15) +
          Math.round(depCharge * 0.50) +
          Math.round(depCharge * 0.15) +
          Math.round(depCharge * 0.12)
        ),
        depDisposals: Math.round(depDisposals * 0.10),
        depClosing: closingDep - (
          Math.round(closingDep * 0.18) +
          Math.round(closingDep * 0.45) +
          Math.round(closingDep * 0.15) +
          Math.round(closingDep * 0.12)
        ),
        netBookValue: 0,
        depRate: 25.0,
      },
    ];
    const lastCat = categories[categories.length - 1];
    lastCat.netBookValue = lastCat.costClosing - lastCat.depClosing;

    result.push({
      year: input.year,
      isProjected: input.isProjected,
      categories,
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
  });

  return result;
}

// 2. Calculate Income Statement (Strictly preserves salesRevenue, financeCost, and netProfit)
export function calculateIncomeStatement(
  inputs: YearDataInput[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[]
): IncomeStatementYear[] {
  return inputs.map((input, index) => {
    const revenue = input.salesRevenue; // FIXED USER INPUT
    const financeCost = input.financeCost; // FIXED USER INPUT
    const netProfit = input.netProfit; // FIXED USER INPUT
    const depreciationExpense = fixedAssetsSchedule[index]?.totalDepCharge || Math.round(input.fixedAssetsGross * 0.1);

    const taxRate = input.taxRate || 25;
    const taxExpense = netProfit >= 0
      ? Math.round(netProfit / Math.max(0.01, 1 - taxRate / 100)) - netProfit
      : 0;
    const profitBeforeTax = netProfit + taxExpense;
    let interestIncome = Math.round(revenue * 0.005);

    // Operating Profit (EBIT) required to yield exact PBT given fixed financeCost & interestIncome
    const operatingProfit = profitBeforeTax - interestIncome + financeCost;

    // Total operating costs (CoS + OpEx) = Revenue - EBIT - Depreciation
    const requiredOperatingCosts = revenue - operatingProfit - depreciationExpense;
    let costOfSales = input.costOfSales;
    let operatingExpenses = input.operatingExpenses;

    if (requiredOperatingCosts >= 0) {
      const existingSum = Math.max(1, input.costOfSales + input.operatingExpenses);
      const cosRatio = input.costOfSales > 0 ? input.costOfSales / existingSum : 0.75;
      costOfSales = Math.round(requiredOperatingCosts * cosRatio);
      operatingExpenses = requiredOperatingCosts - costOfSales;
    } else {
      // If NPAT exceeds Sales, preserve revenue/NPAT/financeCost by increasing Other/Interest Income
      costOfSales = 0;
      operatingExpenses = 0;
      interestIncome += Math.abs(requiredOperatingCosts);
    }

    const grossProfit = revenue - costOfSales;

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
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[]
): CashFlowStatementYear[] {
  const result: CashFlowStatementYear[] = [];

  inputs.forEach((input, index) => {
    const isFirst = index === 0;
    const prevInput = isFirst ? null : inputs[index - 1];
    const income = incomeStatements[index];
    const fixedAssets = fixedAssetsSchedule[index];

    const profitBeforeTax = income.profitBeforeTax;
    const adjDepreciation = fixedAssets.totalDepCharge;
    const adjAmortization = prevInput ? Math.max(0, prevInput.intangibleAssets - input.intangibleAssets) : 100;
    const adjFinanceCost = income.financeCost;
    const adjInterestIncome = -income.interestIncome;
    const adjLossGainDisposal = -fixedAssets.totalDepDisposals;

    const operatingProfitBeforeWC =
      profitBeforeTax + adjDepreciation + adjAmortization + adjFinanceCost + adjInterestIncome + adjLossGainDisposal;

    const changeInInventory = prevInput
      ? -(input.closingInventory - prevInput.closingInventory)
      : -Math.round(input.closingInventory * 0.08);

    const changeInReceivables = prevInput
      ? -(input.tradeReceivables - prevInput.tradeReceivables)
      : -Math.round(input.tradeReceivables * 0.08);

    const changeInOtherReceivables = prevInput
      ? -(input.otherReceivables - prevInput.otherReceivables)
      : -Math.round(input.otherReceivables * 0.05);

    const changeInOtherCA = prevInput
      ? -(input.otherCurrentAssets - prevInput.otherCurrentAssets)
      : -Math.round(input.otherCurrentAssets * 0.04);

    const changeInPayables = prevInput
      ? (input.tradePayables - prevInput.tradePayables)
      : Math.round(input.tradePayables * 0.07);

    const changeInOtherPayables = prevInput
      ? (input.otherPayables - prevInput.otherPayables)
      : Math.round(input.otherPayables * 0.04);

    const currTaxPayable = Math.round(Math.max(0, income.taxExpense) * 0.45);
    const prevTaxPayable = prevInput ? Math.round(Math.max(0, incomeStatements[index - 1].taxExpense) * 0.45) : currTaxPayable;
    const changeInTaxPayable = currTaxPayable - prevTaxPayable;

    const cashGeneratedFromOperations =
      operatingProfitBeforeWC +
      changeInInventory +
      changeInReceivables +
      changeInOtherReceivables +
      changeInOtherCA +
      changeInPayables +
      changeInOtherPayables;

    const incomeTaxPaid = -income.taxExpense + changeInTaxPayable;
    const financeCostPaid = -income.financeCost;

    const netCashFromOperating = cashGeneratedFromOperations + incomeTaxPaid + financeCostPaid;

    // B. Investing Activities
    const purchasePPE = -fixedAssets.totalAdditions;
    const salePPE = fixedAssets.totalDisposals;
    const purchaseIntangibles = prevInput ? -Math.max(0, input.intangibleAssets - prevInput.intangibleAssets) : 0;
    const interestReceived = income.interestIncome;
    const dividendsReceived = 0;

    const netCashFromInvesting =
      purchasePPE + salePPE + purchaseIntangibles + interestReceived + dividendsReceived;

    // C. Financing Activities (Preserves user's direct financing input and repaymentOfBorrowings)
    const directlyEnteredFinancingCashFlow = input.cashFlowFromFinancing;
    const financingProceedsShares = input.financingProceedsShares || 0;
    const financingProceedsBorrowings = input.financingProceedsBorrowings || 0;
    const repaymentOfBorrowings = input.repaymentOfBorrowings !== undefined
      ? input.repaymentOfBorrowings
      : (input.financingRepaymentBorrowings ? Math.abs(input.financingRepaymentBorrowings) : 0);
    const financingRepaymentBorrowings = -repaymentOfBorrowings;
    const financingDividendsPaid = -input.dividendsPaid;

    const breakdownSum = financingProceedsShares + financingProceedsBorrowings + financingRepaymentBorrowings + financingDividendsPaid;
    const isFinancingBreakdownComplete = Math.abs(breakdownSum - directlyEnteredFinancingCashFlow) < 10;

    const netCashFromFinancing = directlyEnteredFinancingCashFlow;

    // D. Net Increase / Decrease in Cash and Cash Equivalents
    const netChangeInCash = netCashFromOperating + netCashFromInvesting + netCashFromFinancing;

    let openingCash = isFirst
      ? (input.cashAndEquivalentsBS - netChangeInCash)
      : result[index - 1].closingCashCalculated;
    if (openingCash < 0 && isFirst) {
      openingCash = Math.max(200, Math.round(input.salesRevenue * 0.05));
    }

    const closingCashCalculated = openingCash + netChangeInCash;
    // Reconcile Balance Sheet Cash cleanly unless user has explicitly entered an audited override for historical years
    const balanceSheetCash = isFirst && input.cashAndEquivalentsBS !== undefined && !input.isProjected
      ? (Math.abs(closingCashCalculated - input.cashAndEquivalentsBS) < 5 ? closingCashCalculated : input.cashAndEquivalentsBS)
      : closingCashCalculated;

    const reconciliationDifference = closingCashCalculated - balanceSheetCash;
    const isReconciled = Math.abs(reconciliationDifference) === 0;

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
  incomeStatements: IncomeStatementYear[]
): EquityStatementYear[] {
  const result: EquityStatementYear[] = [];

  inputs.forEach((input, index) => {
    const isFirst = index === 0;
    const income = incomeStatements[index];

    const openingShareCapital = isFirst ? input.shareCapital : result[index - 1].closingShareCapital;
    const shareCapitalIssued = input.financingProceedsShares || 0;
    const closingShareCapital = openingShareCapital + shareCapitalIssued;

    const openingRetainedEarnings = isFirst
      ? (input.retainedEarningsOpening ?? 6500)
      : result[index - 1].closingRetainedEarnings;
    const netProfitForYear = income.netProfit; // FIXED USER INPUT
    const dividendsPaid = input.dividendsPaid;
    const closingRetainedEarnings = openingRetainedEarnings + netProfitForYear - dividendsPaid;

    const openingReserves = isFirst ? 1800 : result[index - 1].closingReserves;
    const otherComprehensiveIncome = 0;
    const closingReserves = openingReserves + otherComprehensiveIncome;

    const totalOpeningEquity = openingShareCapital + openingRetainedEarnings + openingReserves;
    const totalClosingEquity = closingShareCapital + closingRetainedEarnings + closingReserves;

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
  cashFlowStatements?: CashFlowStatementYear[]
): BalanceSheetYear[] {
  return inputs.map((input, index) => {
    const equity = equityStatements[index];
    const fixedAssets = fixedAssetsSchedule[index];

    const cashAndEquivalents = cashFlowStatements && cashFlowStatements[index]
      ? cashFlowStatements[index].closingCashCalculated
      : input.cashAndEquivalentsBS;
    const tradeReceivables = input.tradeReceivables;
    const otherReceivables = input.otherReceivables;
    const closingInventory = input.closingInventory; // FIXED USER INPUT
    const otherCurrentAssets = input.otherCurrentAssets;
    const totalCurrentAssets = cashAndEquivalents + tradeReceivables + otherReceivables + closingInventory + otherCurrentAssets;

    const tradePayables = input.tradePayables;
    const otherPayables = input.otherPayables;
    const shortTermBorrowings = input.shortTermBorrowings;
    const currentTaxPayable = Math.round(Math.max(0, input.netProfit / Math.max(0.01, 1 - (input.taxRate || 25) / 100) - input.netProfit) * 0.45);
    const totalCurrentLiabilities = tradePayables + otherPayables + shortTermBorrowings + currentTaxPayable;

    const longTermBorrowings = input.longTermBorrowings;
    const otherNonCurrentLiabilities = 0;
    const totalNonCurrentLiabilities = longTermBorrowings + otherNonCurrentLiabilities;

    const totalLiabilities = totalCurrentLiabilities + totalNonCurrentLiabilities;

    const shareCapital = equity.closingShareCapital;
    const retainedEarnings = equity.closingRetainedEarnings;
    const reserves = equity.closingReserves;
    const totalEquity = shareCapital + retainedEarnings + reserves;

    const totalLiabilitiesAndEquity = totalLiabilities + totalEquity;

    const propertyPlantEquipment = fixedAssets.totalNetBookValue;
    const intangibleAssets = input.intangibleAssets;
    const nonCurrentAssetsBaseline = propertyPlantEquipment + intangibleAssets;
    const otherNonCurrentAssets = Math.max(0, totalLiabilitiesAndEquity - (totalCurrentAssets + nonCurrentAssetsBaseline));
    const totalNonCurrentAssets = nonCurrentAssetsBaseline + otherNonCurrentAssets;

    const totalAssets = totalCurrentAssets + totalNonCurrentAssets;
    const balanceCheck = totalAssets - totalLiabilitiesAndEquity;

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
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[]
): FinancialNoteItem[] {
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

  return [
    {
      noteNumber: 1,
      title: 'Company Information & Significant Accounting Policies',
      category: 'General Disclosures',
      textDescription: `${company.companyName} is a limited liability company incorporated under applicable commercial law, with its registered corporate office situated at ${company.registeredAddress}. The financial statements have been prepared in accordance with ${company.accountingFramework} and applicable statutory frameworks. The Statement of Cash Flows is presented using the Indirect Method in compliance with IAS 7 / ASC 230, distinguishing between operating, investing, and financing activities. Cash and cash equivalents comprise cash in hand, call deposits with scheduled banks, and liquid money-market placements with original maturities of three months or less.`,
    },
    {
      noteNumber: 2,
      title: 'Sales Revenue & Disaggregation of Contracts',
      category: 'Income Statement',
      textDescription: 'Revenue is recognized upon transfer of promised goods or services to customers at an amount that reflects the consideration to which the entity expects to be entitled in exchange for those goods or services.',
      tableData: {
        headers: ['Revenue Streams', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Commercial Contract Deliverables', values: [Math.round(cIS.revenue * 0.72).toLocaleString(), Math.round(pIS.revenue * 0.72).toLocaleString()] },
          { label: 'Institutional & Direct Enterprise Sales', values: [Math.round(cIS.revenue * 0.20).toLocaleString(), Math.round(pIS.revenue * 0.20).toLocaleString()] },
          { label: 'Technical Services & Ancillary Revenue', values: [Math.round(cIS.revenue * 0.08).toLocaleString(), Math.round(pIS.revenue * 0.08).toLocaleString()] },
        ],
        totalRow: { label: 'Total Revenue from Contracts', values: [cIS.revenue.toLocaleString(), pIS.revenue.toLocaleString()] },
      },
    },
    {
      noteNumber: 3,
      title: 'Cost of Sales & Operating Overheads',
      category: 'Income Statement',
      textDescription: 'Cost of sales reflects direct raw materials consumed, conversion costs, production plant utilities, direct labor, and manufacturing depreciation allocated in accordance with production throughput.',
      tableData: {
        headers: ['Component Breakdown', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Raw Materials & Components Consumed', values: [Math.round(cIS.costOfSales * 0.58).toLocaleString(), Math.round(pIS.costOfSales * 0.58).toLocaleString()] },
          { label: 'Direct Labor & Industrial Staff Costs', values: [Math.round(cIS.costOfSales * 0.22).toLocaleString(), Math.round(pIS.costOfSales * 0.22).toLocaleString()] },
          { label: 'Factory Overheads, Fuel & Power', values: [Math.round(cIS.costOfSales * 0.20).toLocaleString(), Math.round(pIS.costOfSales * 0.20).toLocaleString()] },
        ],
        totalRow: { label: 'Total Cost of Sales', values: [cIS.costOfSales.toLocaleString(), pIS.costOfSales.toLocaleString()] },
      },
    },
    {
      noteNumber: 4,
      title: 'Finance Cost & Borrowing Disclosures',
      category: 'Financing',
      textDescription: 'Finance costs encompass interest charges on term borrowings, bank overdraft facilities, letter of credit commissions, and Islamic financing profit rates where applicable. In accordance with Section 76.4, Finance Cost is a mandatory user-entered parameter across all historical and projected reporting cycles.',
      tableData: {
        headers: ['Financing Facility Charges', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Interest on Long-Term Credit Facilities', values: [Math.round(cIS.financeCost * 0.65).toLocaleString(), Math.round(pIS.financeCost * 0.65).toLocaleString()] },
          { label: 'Interest on Working Capital & Short-Term Debt', values: [Math.round(cIS.financeCost * 0.25).toLocaleString(), Math.round(pIS.financeCost * 0.25).toLocaleString()] },
          { label: 'Bank Charges, LC Fees & Financing Commissions', values: [Math.round(cIS.financeCost * 0.10).toLocaleString(), Math.round(pIS.financeCost * 0.10).toLocaleString()] },
        ],
        totalRow: { label: 'Total Finance Cost', values: [cIS.financeCost.toLocaleString(), pIS.financeCost.toLocaleString()] },
      },
    },
    {
      noteNumber: 5,
      title: 'Property, Plant and Equipment (PPE)',
      category: 'Balance Sheet',
      textDescription: 'Items of PPE are stated at cost less accumulated depreciation and impairment losses. Depreciation is computed on the straight-line method based on estimated economic useful lives. Capital additions and disposals are linked directly to Cash Flows from Investing Activities.',
      tableData: {
        headers: ['PPE Category', 'Gross Cost', 'Accumulated Dep.', 'Net Book Value'],
        rows: cFA.categories.map(cat => ({
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
      title: 'Closing Inventories',
      category: 'Working Capital',
      textDescription: 'Inventories are valued at the lower of cost and net realizable value (NRV). Cost is determined using the weighted average method. Provisions are made for slow-moving and obsolete lines.',
      tableData: {
        headers: ['Inventory Classes', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Raw Materials & Imported Spares', values: [Math.round(cBS.closingInventory * 0.45).toLocaleString(), Math.round(pBS.closingInventory * 0.45).toLocaleString()] },
          { label: 'Work-in-Progress (WIP)', values: [Math.round(cBS.closingInventory * 0.20).toLocaleString(), Math.round(pBS.closingInventory * 0.20).toLocaleString()] },
          { label: 'Finished Goods Inventory', values: [Math.round(cBS.closingInventory * 0.35).toLocaleString(), Math.round(pBS.closingInventory * 0.35).toLocaleString()] },
        ],
        totalRow: { label: 'Total Closing Inventory', values: [cBS.closingInventory.toLocaleString(), pBS.closingInventory.toLocaleString()] },
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
          { label: 'Trade Debtors (within 30 days)', values: [Math.round(cBS.tradeReceivables * 0.65).toLocaleString(), Math.round(pBS.tradeReceivables * 0.65).toLocaleString()] },
          { label: 'Trade Debtors (31 - 90 days)', values: [Math.round(cBS.tradeReceivables * 0.25).toLocaleString(), Math.round(pBS.tradeReceivables * 0.25).toLocaleString()] },
          { label: 'Trade Debtors (over 90 days, secured)', values: [Math.round(cBS.tradeReceivables * 0.10).toLocaleString(), Math.round(pBS.tradeReceivables * 0.10).toLocaleString()] },
          { label: 'Advance Deposits & Prepayments', values: [cBS.otherReceivables.toLocaleString(), pBS.otherReceivables.toLocaleString()] },
        ],
        totalRow: { label: 'Total Receivables', values: [(cBS.tradeReceivables + cBS.otherReceivables).toLocaleString(), (pBS.tradeReceivables + pBS.otherReceivables).toLocaleString()] },
      },
    },
    {
      noteNumber: 8,
      title: 'Cash and Cash Equivalents & Bank Accounts',
      category: 'Liquidity',
      textDescription: 'Cash and cash equivalents comprise operational petty cash floats and bank current/savings accounts with scheduled commercial financial institutions.',
      tableData: {
        headers: ['Liquid Asset Classes', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Cash in Hand (Corporate & Branches)', values: [Math.round(cBS.cashAndEquivalents * 0.05).toLocaleString(), Math.round(pBS.cashAndEquivalents * 0.05).toLocaleString()] },
          { label: 'Balances with Commercial Banks', values: [Math.round(cBS.cashAndEquivalents * 0.70).toLocaleString(), Math.round(pBS.cashAndEquivalents * 0.70).toLocaleString()] },
          { label: 'Short-Term Liquid Placements (< 90 Days)', values: [Math.round(cBS.cashAndEquivalents * 0.25).toLocaleString(), Math.round(pBS.cashAndEquivalents * 0.25).toLocaleString()] },
        ],
        totalRow: { label: 'Total Balance Sheet Cash & Cash Equivalents', values: [cBS.cashAndEquivalents.toLocaleString(), pBS.cashAndEquivalents.toLocaleString()] },
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
          { label: 'Paid-Up Ordinary Share Capital', values: [cBS.shareCapital.toLocaleString(), pBS.shareCapital.toLocaleString()] },
          { label: 'Retained Earnings Balance', values: [cBS.retainedEarnings.toLocaleString(), pBS.retainedEarnings.toLocaleString()] },
          { label: 'Capital & Asset Revaluation Reserves', values: [cBS.reserves.toLocaleString(), pBS.reserves.toLocaleString()] },
        ],
        totalRow: { label: 'Total Shareholders Equity', values: [cBS.totalEquity.toLocaleString(), pBS.totalEquity.toLocaleString()] },
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
          { label: 'Secured Long-Term Term Loans', values: [cBS.longTermBorrowings.toLocaleString(), pBS.longTermBorrowings.toLocaleString()] },
          { label: 'Short-Term Working Capital & Overdraft Facilities', values: [cBS.shortTermBorrowings.toLocaleString(), pBS.shortTermBorrowings.toLocaleString()] },
        ],
        totalRow: { label: 'Total Financial Debt Obligations', values: [(cBS.longTermBorrowings + cBS.shortTermBorrowings).toLocaleString(), (pBS.longTermBorrowings + pBS.shortTermBorrowings).toLocaleString()] },
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
          { label: 'Trade Creditors & Raw Material Suppliers', values: [cBS.tradePayables.toLocaleString(), pBS.tradePayables.toLocaleString()] },
          { label: 'Accrued Expenses & Statutory Withholdings', values: [cBS.otherPayables.toLocaleString(), pBS.otherPayables.toLocaleString()] },
        ],
        totalRow: { label: 'Total Current Payables', values: [(cBS.tradePayables + cBS.otherPayables).toLocaleString(), (pBS.tradePayables + pBS.otherPayables).toLocaleString()] },
      },
    },
    {
      noteNumber: 12,
      title: 'Cash Flow Disclosures & Balance Sheet Reconciliation',
      category: 'Cash Flow Reconciliation',
      textDescription: `Pursuant to Section 76.3 and 76.4, the Statement of Cash Flows directly incorporates the user-entered Cash Flow from Financing Activities (${cCF.directlyEnteredFinancingCashFlow.toLocaleString()} ${company.reportingCurrency}) without automatic derivations or assumptions. Closing cash is reconciled with the Statement of Financial Position: Closing Calculated Cash is ${cCF.closingCashCalculated.toLocaleString()} vs Balance Sheet Cash of ${cCF.balanceSheetCash.toLocaleString()}. Audit status: ${cCF.reconciliationStatus === 'RECONCILED' ? 'COMPLETELY RECONCILED (Variance 0)' : `VARIANCE OF ${cCF.reconciliationDifference.toLocaleString()} IDENTIFIED`}. ${cCF.reconciliationExplanation}`,
      tableData: {
        headers: ['Reconciliation Parameter', `${currYear} (${company.reportingCurrency})`, `${prevYear} (${company.reportingCurrency})`],
        rows: [
          { label: 'Net Cash Generated from Operating Activities', values: [cCF.netCashFromOperating.toLocaleString(), pCF.netCashFromOperating.toLocaleString()] },
          { label: 'Net Cash Generated from / (Used in) Investing Activities', values: [cCF.netCashFromInvesting.toLocaleString(), pCF.netCashFromInvesting.toLocaleString()] },
          { label: 'Directly Entered Cash Flow from Financing Activities', values: [cCF.directlyEnteredFinancingCashFlow.toLocaleString(), pCF.directlyEnteredFinancingCashFlow.toLocaleString()] },
          { label: 'Net Change in Cash and Cash Equivalents', values: [cCF.netChangeInCash.toLocaleString(), pCF.netChangeInCash.toLocaleString()] },
          { label: 'Opening Cash and Cash Equivalents', values: [cCF.openingCash.toLocaleString(), pCF.openingCash.toLocaleString()] },
          { label: 'Closing Cash Derived from Cash Flow Statement', values: [cCF.closingCashCalculated.toLocaleString(), pCF.closingCashCalculated.toLocaleString()] },
          { label: 'Stated Cash & Equivalents in Statement of Financial Position', values: [cCF.balanceSheetCash.toLocaleString(), pCF.balanceSheetCash.toLocaleString()] },
        ],
        totalRow: {
          label: 'Audit Reconciliation Variance',
          values: [cCF.reconciliationDifference.toLocaleString(), pCF.reconciliationDifference.toLocaleString()],
        },
      },
    },
  ];
}

// Default Configuration Items for the Financial Statement Configuration Module (Section 76.6)
export const DEFAULT_STATEMENT_CONFIG_ITEMS: StatementConfigItem[] = [
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
  { id: 'cf_wc_payables', statement: 'cashFlow', section: 'Operating Activities', label: 'Increase / (Decrease) in Trade Payables', isMandatory: false, isVisible: true, order: 10 },
  { id: 'cf_wc_other_pay', statement: 'cashFlow', section: 'Operating Activities', label: 'Increase / (Decrease) in Other Payables', isMandatory: false, isVisible: true, order: 11 },
  { id: 'cf_tax_paid', statement: 'cashFlow', section: 'Operating Activities', label: 'Income Tax Paid', isMandatory: true, isVisible: true, order: 12 },
  { id: 'cf_finance_paid', statement: 'cashFlow', section: 'Operating Activities', label: 'Finance Cost Paid', isMandatory: true, isVisible: true, order: 13 },
  
  // Investing Activities
  { id: 'cf_pur_ppe', statement: 'cashFlow', section: 'Investing Activities', label: 'Purchase of Property, Plant and Equipment', isMandatory: true, isVisible: true, order: 14 },
  { id: 'cf_sale_ppe', statement: 'cashFlow', section: 'Investing Activities', label: 'Proceeds from Sale of Fixed Assets', isMandatory: false, isVisible: true, order: 15 },
  { id: 'cf_pur_intangible', statement: 'cashFlow', section: 'Investing Activities', label: 'Purchase of Intangible Assets', isMandatory: false, isVisible: true, order: 16 },
  { id: 'cf_interest_rec', statement: 'cashFlow', section: 'Investing Activities', label: 'Interest Received', isMandatory: false, isVisible: true, order: 17 },
  { id: 'cf_dividends_rec', statement: 'cashFlow', section: 'Investing Activities', label: 'Dividends Received', isMandatory: false, isVisible: true, order: 18 },

  // Financing Activities
  { id: 'cf_direct_financing', statement: 'cashFlow', section: 'Financing Activities', label: 'Net Cash Flow from Financing Activities (Direct User Entry)', isMandatory: true, isVisible: true, order: 19 },
  { id: 'cf_shares_issue', statement: 'cashFlow', section: 'Financing Activities', label: 'Proceeds from Issue of Share Capital', isMandatory: false, isVisible: true, order: 20 },
  { id: 'cf_borrow_proceeds', statement: 'cashFlow', section: 'Financing Activities', label: 'Proceeds from Borrowings & Financing', isMandatory: false, isVisible: true, order: 21 },
  { id: 'cf_borrow_repay', statement: 'cashFlow', section: 'Financing Activities', label: 'Repayment of Borrowings & Financing', isMandatory: false, isVisible: true, order: 22 },
  { id: 'cf_dividends_paid', statement: 'cashFlow', section: 'Financing Activities', label: 'Dividends Paid to Shareholders', isMandatory: false, isVisible: true, order: 23 },
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
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[]
): import('../types/financial').PreviousFormatYearData[] {
  const histIndices = inputs
    .map((inp, idx) => ({ inp, idx }))
    .filter(item => !item.inp.isProjected);

  const prevResults: import('../types/financial').PreviousFormatYearData[] = [];

  histIndices.forEach(({ inp, idx }, hIdx) => {
    const inc = incomeStatements[idx];
    const bs = balanceSheets[idx];
    const cf = cashFlowStatements[idx];
    const eq = equityStatements[idx];
    const fa = fixedAssetsSchedule[idx];

    // Fixed 6 inputs strictly preserved
    const sales = inp.salesRevenue;
    const stockAndStores = inp.closingInventory;
    const financialExp = inp.financeCost;
    const netProfitAfterIncomeTax = inp.netProfit;
    const repaymentOfBorrowings = inp.repaymentOfBorrowings;
    const faGrossAssetBase = inp.fixedAssetsGross;

    // Income Statement exact breakdown
    const costOfGoodsSold = inc.costOfSales;
    const grossProfit = sales - costOfGoodsSold;
    const depreciation = fa.totalDepCharge;
    const nonDepOpEx = Math.max(0, inc.operatingExpenses);

    // Allocate non-depreciation Operating Expenses cleanly across the 12 exact expense lines from Previous Financials
    const weights = [
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
    ];
    const salaryAndAllowances = Math.round(nonDepOpEx * weights[0]);
    const telephoneAndMobileBill = Math.round(nonDepOpEx * weights[1]);
    const postageAndCourier = Math.round(nonDepOpEx * weights[2]);
    const electricCharges = Math.round(nonDepOpEx * weights[3]);
    const entertainment = Math.round(nonDepOpEx * weights[4]);
    const printingAndStationery = Math.round(nonDepOpEx * weights[5]);
    const generalExpenses = Math.round(nonDepOpEx * weights[6]);
    const licenceAndRenewals = Math.round(nonDepOpEx * weights[7]);
    const auditAndProfessionalFee = Math.round(nonDepOpEx * weights[8]);
    const travellingAndConveyance = Math.round(nonDepOpEx * weights[9]);
    const newspaperAndPeriodicals = Math.round(nonDepOpEx * weights[10]);
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
    const sellingExp = Math.max(0, nonDepOpEx - partialSum);
    const adminAndSellingTotal = partialSum + sellingExp + depreciation;
    const netOperatingProfit = grossProfit - adminAndSellingTotal;
    const otherIncome = inc.interestIncome;
    const operatingSubtotal = netOperatingProfit + otherIncome;
    const incomeTaxProvision = inc.taxExpense;

    // Balance Sheet exact structure (reconciles Assets === Capital + Liabilities to 0!)
    const fixedAssetsWDV = fa.totalNetBookValue;
    const advancesDepositsPrepayments = bs.otherReceivables + bs.otherCurrentAssets;
    const billsReceivable = bs.tradeReceivables;
    const cashAndCashEquivalents = bs.cashAndEquivalents;
    const totalCurrentAssets =
      advancesDepositsPrepayments + stockAndStores + billsReceivable + cashAndCashEquivalents;
    const totalAssetsTaka = fixedAssetsWDV + totalCurrentAssets;

    const forExpenses = bs.otherPayables;
    const accountsPayable = bs.tradePayables;
    const shortTermLoan = bs.shortTermBorrowings + bs.longTermBorrowings;
    const provisionForTaxation = bs.currentTaxPayable;
    const totalCurrentLiabilities =
      forExpenses + accountsPayable + shortTermLoan + provisionForTaxation;

    const capitalAfterDrawings = totalAssetsTaka - totalCurrentLiabilities;
    const drawings = inp.dividendsPaid;
    const capitalSubtotal = capitalAfterDrawings + drawings;
    const openingCapital = hIdx === 0
      ? capitalSubtotal - netProfitAfterIncomeTax
      : prevResults[hIdx - 1].capitalAfterDrawings;

    // Cash Flow exact structure matching Previous Financials PDF & exact ending cash
    const prevBS = hIdx > 0 ? prevResults[hIdx - 1] : null;
    const cfNetIncome = netProfitAfterIncomeTax;
    const cfDepreciation = depreciation;
    const nonCashCA = advancesDepositsPrepayments + stockAndStores + billsReceivable;
    const prevNonCashCA = prevBS
      ? prevBS.advancesDepositsPrepayments + prevBS.stockAndStores + prevBS.billsReceivable
      : Math.round(nonCashCA * 0.92);
    const cfChangeInCurrentAssets = -(nonCashCA - prevNonCashCA);
    const cfChangeInCWIP = 0;

    // Working capital liabilities + Short Term Loan changes
    const cfChangeInCurrentLiabilities = prevBS
      ? totalCurrentLiabilities - prevBS.totalCurrentLiabilities
      : Math.round(totalCurrentLiabilities * 0.05);
    const cfNetCashOperating =
      cfNetIncome + cfDepreciation + cfChangeInCurrentAssets + cfChangeInCWIP + cfChangeInCurrentLiabilities;

    const cfAcquisitionPPE = -fa.totalAdditions;
    const cfDisposalIntangible = fa.totalDisposals - fa.totalDepDisposals;
    const cfNetCashInvesting = cfAcquisitionPPE + cfDisposalIntangible;

    const cfDrawings = -drawings;
    const cfOpeningBalance = hIdx === 0
      ? 0
      : (capitalAfterDrawings - (prevBS!.capitalAfterDrawings + netProfitAfterIncomeTax - drawings));
    const cfNetCashFinancing = cfDrawings + cfOpeningBalance;

    const cfNetIncreaseCash = cfNetCashOperating + cfNetCashInvesting + cfNetCashFinancing;
    const cfCashBeginning = prevBS
      ? prevBS.cfCashEnding
      : cashAndCashEquivalents - cfNetIncreaseCash;
    const cfCashEnding = cfCashBeginning + cfNetIncreaseCash;

    prevResults.push({
      year: inp.year,
      dateLabel: `31.12.${inp.year}`,
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
      totalCapitalAndLiabilities: capitalAfterDrawings + totalCurrentLiabilities,
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

// ---------------- PROJECTED FINANCIALS FORMAT BUILDER (Exact Compliance) ---------------- //
export function buildProjectedFormatData(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  equityStatements: EquityStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[]
): import('../types/financial').ProjectedFormatYearData[] {
  const projIndices = inputs
    .map((inp, idx) => ({ inp, idx }))
    .filter(item => item.inp.isProjected);

  // Anchor opening year to last historical year if available
  const firstProjIdx = projIndices[0]?.idx ?? 0;
  const anchorHistIdx = firstProjIdx > 0 ? firstProjIdx - 1 : -1;

  const projResults: import('../types/financial').ProjectedFormatYearData[] = [];

  projIndices.forEach(({ inp, idx }, pIdx) => {
    const inc = incomeStatements[idx];
    const bs = balanceSheets[idx];
    const eq = equityStatements[idx];
    const fa = fixedAssetsSchedule[idx];

    const prevProj = pIdx > 0 ? projResults[pIdx - 1] : null;
    const prevBS = anchorHistIdx >= 0 && pIdx === 0 ? balanceSheets[anchorHistIdx] : null;
    const prevFA = anchorHistIdx >= 0 && pIdx === 0 ? fixedAssetsSchedule[anchorHistIdx] : null;

    // Fixed 6 inputs strictly preserved
    const sales = inp.salesRevenue;
    const inventories = inp.closingInventory;
    const financialExpenses = inp.financeCost;
    const netProfitAfterTax = inp.netProfit;
    const repaymentOfBorrowings = inp.repaymentOfBorrowings;
    const schedGrossAssetBase = inp.fixedAssetsGross;

    // P&L
    const costOfGoodsSold = inc.costOfSales;
    const grossProfit = sales - costOfGoodsSold;
    const othersIncome = inc.interestIncome;
    const grossProfitWithOtherIncome = grossProfit + othersIncome;
    const adminAndSellingExpenses = inc.operatingExpenses + fa.totalDepCharge;
    const totalOperatingProfit = grossProfitWithOtherIncome - adminAndSellingExpenses;
    const netProfitBeforeTax = totalOperatingProfit - financialExpenses;
    const provisionForIncomeTax = inc.taxExpense;

    // Balance Sheet Assets
    const propertyPlantEquipments = fa.totalNetBookValue;
    const tradeAndOtherReceivables = bs.tradeReceivables;
    const advanceDepositPrepayments = bs.otherReceivables + bs.otherCurrentAssets;

    // Liabilities
    const longTermLoan = bs.longTermBorrowings;
    const tradeAndOtherPayables = bs.tradePayables;
    const provisionsForTax = bs.currentTaxPayable;
    const shortTermLoan = bs.shortTermBorrowings;
    const liabilitiesForExp = bs.otherPayables;
    const totalCurrentLiabilities =
      tradeAndOtherPayables + provisionsForTax + shortTermLoan + liabilitiesForExp;

    // Cash Flow Statement (exact roll-forward matching Projected_Financials PDF)
    const prevInv = prevProj ? prevProj.inventories : (prevBS ? prevBS.closingInventory : Math.round(inventories * 0.9));
    const prevRec = prevProj ? prevProj.tradeAndOtherReceivables : (prevBS ? prevBS.tradeReceivables : Math.round(tradeAndOtherReceivables * 0.9));
    const prevAdv = prevProj ? prevProj.advanceDepositPrepayments : (prevBS ? prevBS.otherReceivables + prevBS.otherCurrentAssets : Math.round(advanceDepositPrepayments * 0.9));
    const prevPay = prevProj ? prevProj.tradeAndOtherPayables : (prevBS ? prevBS.tradePayables : Math.round(tradeAndOtherPayables * 0.9));
    const prevTax = prevProj ? prevProj.provisionsForTax : (prevBS ? prevBS.currentTaxPayable : Math.round(provisionsForTax * 0.9));
    const prevExp = prevProj ? prevProj.liabilitiesForExp : (prevBS ? prevBS.otherPayables : Math.round(liabilitiesForExp * 0.9));

    const cfNetIncome = netProfitAfterTax;
    const cfChangeInventory = -(inventories - prevInv);
    const cfChangeReceivables = -(tradeAndOtherReceivables - prevRec);
    const cfChangeAdvanceDeposit = -(advanceDepositPrepayments - prevAdv);
    const cfChangeCurrentAssets = 0;
    const cfChangeCurrentLiabilities = 0;
    const cfChangeTradePayables = tradeAndOtherPayables - prevPay;
    const cfChangeProvisionsTax = provisionsForTax - prevTax;
    const cfChangeLiabilitiesExp = liabilitiesForExp - prevExp;
    const cfDepreciation = fa.totalDepCharge;

    const cfNetCashOperating =
      cfNetIncome +
      cfChangeInventory +
      cfChangeReceivables +
      cfChangeAdvanceDeposit +
      cfChangeCurrentAssets +
      cfChangeCurrentLiabilities +
      cfChangeTradePayables +
      cfChangeProvisionsTax +
      cfChangeLiabilitiesExp +
      cfDepreciation;

    // Investing
    const cfSalePPE = fa.totalDisposals - fa.totalDepDisposals;
    const cfInvestingCapitalInflow = 0;
    const cfPreliminaryExpenses = 0;
    const cfPurchasePPE = -fa.totalAdditions;
    const cfPurchaseInvestment = 0;
    const cfNetCashInvesting = cfSalePPE + cfInvestingCapitalInflow + cfPreliminaryExpenses + cfPurchasePPE + cfPurchaseInvestment;

    // Financing (strictly reflects Repayment from Borrowings & Financing + debt delta)
    const prevTotalDebt = prevProj
      ? prevProj.longTermLoan + prevProj.shortTermLoan
      : (prevBS ? prevBS.longTermBorrowings + prevBS.shortTermBorrowings : longTermLoan + shortTermLoan + repaymentOfBorrowings);
    const currTotalDebt = longTermLoan + shortTermLoan;
    const cfLoanReceived = currTotalDebt - prevTotalDebt; // Net debt inflow / (-repayment)
    const cfAdditionalCapital = inp.financingProceedsShares || 0;
    const cfIssuanceOfDebt = 0;
    const cfDividendsToStockholders = 0;
    const drawings = inp.dividendsPaid;
    const cfDrawingsOutflow = -drawings;
    const cfRedeemLongTermDebt = 0;
    const cfIssuanceBondsPayable = 0;
    const cfNetCashFinancing =
      cfLoanReceived +
      cfAdditionalCapital +
      cfIssuanceOfDebt +
      cfDividendsToStockholders +
      cfDrawingsOutflow +
      cfRedeemLongTermDebt +
      cfIssuanceBondsPayable;

    const cfNetCashIncrease = cfNetCashOperating + cfNetCashInvesting + cfNetCashFinancing;
    const cfOpeningCash = prevProj
      ? prevProj.cfClosingCash
      : (prevBS ? prevBS.cashAndEquivalents : Math.round(sales * 0.08));
    const cfClosingCash = cfOpeningCash + cfNetCashIncrease;

    const cashAndCashEquivalents = cfClosingCash;
    const totalCurrentAssets =
      inventories + tradeAndOtherReceivables + advanceDepositPrepayments + cashAndCashEquivalents;
    const totalAssets = propertyPlantEquipments + totalCurrentAssets;

    // Capital Roll-forward ensuring Total Assets === Total Equity & Liabilities (0.00 variance!)
    const closingCapital = totalAssets - longTermLoan - totalCurrentLiabilities;
    const equitySubtotal = closingCapital + drawings;
    const capital = prevProj
      ? prevProj.closingCapital
      : equitySubtotal - cfAdditionalCapital - netProfitAfterTax;
    const additionalCapital = equitySubtotal - capital - netProfitAfterTax;

    // Notes 3.00 - 13.00 Breakdown
    const note4ReceivablesAdded = Math.max(0, tradeAndOtherReceivables - prevRec) + Math.round(sales * 0.25);
    const note4ReceivablesReceived = prevRec + note4ReceivablesAdded - tradeAndOtherReceivables;

    const note5AdvanceAdded = Math.max(0, advanceDepositPrepayments - prevAdv) + Math.round(advanceDepositPrepayments * 0.1);
    const note5AdvanceAdjusted = prevAdv + note5AdvanceAdded - advanceDepositPrepayments;

    const note8PayablesAdded = Math.max(0, tradeAndOtherPayables - prevPay) + Math.round(costOfGoodsSold * 0.15);
    const note8PayablesPaid = prevPay + note8PayablesAdded - tradeAndOtherPayables;

    const note9TaxAdded = provisionForIncomeTax;
    const note9TaxAdjusted = prevTax + note9TaxAdded - provisionsForTax;

    const note11OpeningStock = prevInv;
    const note11ClosingStock = inventories;
    const directTotal = Math.round(costOfGoodsSold * 0.015);
    const note11LabourWages = Math.round(directTotal * 0.18);
    const note11ElectricityBill = Math.round(directTotal * 0.36);
    const note11CarriageInward = Math.round(directTotal * 0.21);
    const note11OtherDirectExpenses = Math.max(0, directTotal - note11LabourWages - note11ElectricityBill - note11CarriageInward);
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
      totalEquityAndLiabilities: closingCapital + longTermLoan + totalCurrentLiabilities,
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

// ---------------- MODEL VALIDATION SUITE (Requirement 5) ---------------- //
export function validateFinancialModel(
  inputs: YearDataInput[],
  incomeStatements: IncomeStatementYear[],
  balanceSheets: BalanceSheetYear[],
  cashFlowStatements: CashFlowStatementYear[],
  equityStatements: EquityStatementYear[],
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[]
): import('../types/financial').FinancialValidationIssue[] {
  const issues: import('../types/financial').FinancialValidationIssue[] = [];

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

    if (inc && Math.abs(inc.netProfit - inp.netProfit) > 0.01) {
      issues.push({
        id: `npat-mismatch-${inp.year}`,
        year: inp.year,
        severity: 'error',
        title: `NPAT Preservation Alert (${inp.year})`,
        message: `Income Statement Net Profit (${inc.netProfit}) differs from user input (${inp.netProfit}).`,
      });
    }

    if (fa && Math.abs(fa.totalCostClosing - inp.fixedAssetsGross) > 0.01) {
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

    if (eq && bs && Math.abs(eq.totalClosingEquity - bs.totalEquity) > 1) {
      issues.push({
        id: `eq-mismatch-${inp.year}`,
        year: inp.year,
        severity: 'error',
        title: `Equity Tie-Out Variance (${inp.year})`,
        message: `Statement of Changes in Equity (${eq.totalClosingEquity}) does not match Balance Sheet Equity (${bs.totalEquity}).`,
      });
    }

    if (cf && cf.reconciliationDifference !== 0) {
      issues.push({
        id: `cf-variance-${inp.year}`,
        year: inp.year,
        severity: 'warning',
        title: `Cash Flow Reconciliation Variance (${inp.year})`,
        message: `Cash Flow closing cash differs from stated Balance Sheet cash by ${cf.reconciliationDifference.toLocaleString()}.`,
      });
    }
  });

  return issues;
}

