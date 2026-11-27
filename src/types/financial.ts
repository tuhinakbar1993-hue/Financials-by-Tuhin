/**
 * Financial Types & Data Contracts for Financials with Tuhin
 * Section 76 Compliant Architecture
 */

export type EntityType = 
  | 'manufacturing'
  | 'trading'
  | 'service'
  | 'banking_conventional'
  | 'banking_islamic'
  | 'nbfi'
  | 'real_estate'
  | 'healthcare'
  | 'education'
  | 'ngo';

export type AccountingFramework = 'IFRS' | 'US_GAAP' | 'LOCAL_GAAP';

export type CurrencySymbol = '$' | '€' | '£' | '৳' | 'AED' | '₹' | '¥';

export type MultiplierUnit = 'Units' | 'Thousands' | 'Millions' | 'Billions';

export type InputMode = 'manual' | 'automatic';

export type FinancialYearEndOption = 'June 30' | 'December 31';

export interface CompanyProfile {
  companyName: string;
  tradingName?: string;
  registrationNumber: string;
  taxId: string;
  registeredAddress: string;
  reportingCurrency: CurrencySymbol;
  reportingUnit: MultiplierUnit;
  entityType: EntityType;
  accountingFramework: AccountingFramework;
  financialYearEnd: string; // "June 30" or "December 31"
  auditorName: string;
  preparedBy: string; // "Tuhin Akbar, Lead Financial Architect"
}

export type PrimaryDriverKey = 
  | 'salesRevenue'
  | 'closingInventory'
  | 'financeCost'
  | 'netProfit'
  | 'repaymentOfBorrowings'
  | 'fixedAssetsGross';

export interface DriverAutoFillConfig {
  driver: PrimaryDriverKey;
  baseYear: number;
  baseValue: number;
  isActive: boolean;
  growthRatePct?: number;
  description?: string;
}

export type AutoFillConfigs = Partial<Record<PrimaryDriverKey, DriverAutoFillConfig>>;

export interface YearDataInput {
  year: number;
  isProjected: boolean;
  
  // THE SIX PRIMARY FINANCIAL INPUT DRIVERS (Section 76 Core Architecture)
  salesRevenue: number;          // 1. Sales Revenue
  closingInventory: number;      // 2. Closing Inventory
  financeCost: number;           // 3. Finance Cost
  netProfit: number;             // 4. Net Profit After Tax
  repaymentOfBorrowings: number; // 5. Repayment from Borrowings and Financing
  fixedAssetsGross: number;      // 6. Fixed Assets (Gross Asset Base)
  
  // Directly entered financing cash flow total (retained for Section 76.3 compliance)
  cashFlowFromFinancing: number;
  fixedAssetsAdditions: number;
  fixedAssetsDisposals: number;
  depreciationRate: number; // in percentage e.g. 10
  
  // Working Capital & Operating Inputs
  costOfSales: number;
  operatingExpenses: number;
  taxRate: number; // percentage e.g. 25
  tradeReceivables: number;
  otherReceivables: number;
  tradePayables: number;
  otherPayables: number;
  
  // Cash & Balance Sheet Inputs
  cashAndEquivalentsBS: number; // Stated in Balance Sheet
  otherCurrentAssets: number;
  intangibleAssets: number;
  otherNonCurrentAssets: number;
  shortTermBorrowings: number;
  longTermBorrowings: number;
  hasCustomLongTermBorrowings?: boolean;
  hasCustomShortTermBorrowings?: boolean;
  shareCapital: number;
  retainedEarningsOpening?: number;
  reservesOpening?: number;
  dividendsPaid: number;
  
  // Sector-specific inputs
  bankingDeposits?: number;
  bankingLoansAndAdvances?: number;
  islamicFinancingAssets?: number;
  mudarabaDeposits?: number;

  // Optional Financing detailed breakdown (if user enters them)
  financingProceedsShares?: number;
  financingProceedsBorrowings?: number;
  financingRepaymentBorrowings?: number;
  financingDividendsPaid?: number;
  financingLeasePaid?: number;

  // Repayment Schedule lock metadata (ensures exact extracted Principal & Finance Cost remain untouched)
  isScheduleLocked?: boolean;
  schedulePrincipalLocked?: number;
  scheduleFinanceCostLocked?: number;
}

export interface FixedAssetCategory {
  name: string;
  costOpening: number;
  additions: number;
  disposals: number;
  costClosing: number;
  depOpening: number;
  depCharge: number;
  depDisposals: number;
  depClosing: number;
  netBookValue: number;
  depRate: number;
}

export interface ScheduleOfFixedAssetsYear {
  year: number;
  isProjected: boolean;
  categories: FixedAssetCategory[];
  allCategories?: FixedAssetCategory[];
  totalCostOpening: number;
  totalAdditions: number;
  totalDisposals: number;
  totalCostClosing: number;
  totalDepOpening: number;
  totalDepCharge: number;
  totalDepDisposals: number;
  totalDepClosing: number;
  totalNetBookValue: number;
}

export interface IncomeStatementYear {
  year: number;
  isProjected: boolean;
  revenue: number;
  costOfSales: number;
  grossProfit: number;
  operatingExpenses: number;
  depreciationExpense: number;
  operatingProfit: number; // EBIT
  interestIncome: number;
  financeCost: number; // Mandatory user entered
  profitBeforeTax: number; // EBT
  taxExpense: number;
  netProfit: number; // Net Profit after Tax
}

export interface BalanceSheetYear {
  year: number;
  isProjected: boolean;
  
  // Assets
  cashAndEquivalents: number;
  tradeReceivables: number;
  otherReceivables: number;
  closingInventory: number;
  otherCurrentAssets: number;
  totalCurrentAssets: number;
  
  propertyPlantEquipment: number; // NBV from Schedule
  intangibleAssets: number;
  otherNonCurrentAssets: number;
  totalNonCurrentAssets: number;
  totalAssets: number;
  
  // Liabilities
  tradePayables: number;
  otherPayables: number;
  shortTermBorrowings: number;
  currentTaxPayable: number;
  totalCurrentLiabilities: number;
  
  longTermBorrowings: number;
  otherNonCurrentLiabilities: number;
  totalNonCurrentLiabilities: number;
  totalLiabilities: number;
  
  // Equity
  shareCapital: number;
  retainedEarnings: number;
  reserves: number;
  totalEquity: number;
  
  totalLiabilitiesAndEquity: number;
  balanceCheck: number; // totalAssets - totalLiabilitiesAndEquity
}

export interface CashFlowStatementYear {
  year: number;
  isProjected: boolean;
  
  // A. Operating Activities (Indirect Method)
  profitBeforeTax: number;
  adjDepreciation: number;
  adjAmortization: number;
  adjFinanceCost: number;
  adjInterestIncome: number;
  adjLossGainDisposal: number;
  operatingProfitBeforeWC: number;
  
  changeInInventory: number;
  changeInReceivables: number;
  changeInOtherReceivables: number;
  changeInOtherCA?: number;
  changeInPayables: number;
  changeInOtherPayables: number;
  cashGeneratedFromOperations: number;
  
  incomeTaxPaid: number;
  financeCostPaid: number;
  netCashFromOperating: number;
  
  // B. Investing Activities
  purchasePPE: number;
  salePPE: number;
  purchaseIntangibles: number;
  interestReceived: number;
  dividendsReceived: number;
  netCashFromInvesting: number;
  
  // C. Financing Activities (User Directly Entered preserved)
  directlyEnteredFinancingCashFlow: number; // Mandatory user-entered
  financingProceedsShares: number;
  financingProceedsBorrowings: number;
  financingRepaymentBorrowings: number;
  repaymentOfBorrowings: number; // Repayment of Borrowings & Financing (Cash outflow)
  financingDividendsPaid: number;
  netCashFromFinancing: number;
  isFinancingBreakdownComplete: boolean;
  
  // D. Net Increase / Decrease in Cash & Reconciliation
  netChangeInCash: number;
  openingCash: number;
  closingCashCalculated: number;
  balanceSheetCash: number;
  reconciliationDifference: number; // closingCashCalculated - balanceSheetCash
  reconciliationStatus: 'RECONCILED' | 'VARIANCE_IDENTIFIED';
  reconciliationExplanation?: string;
}

export interface EquityStatementYear {
  year: number;
  isProjected: boolean;
  openingShareCapital: number;
  shareCapitalIssued: number;
  closingShareCapital: number;
  
  openingRetainedEarnings: number;
  netProfitForYear: number;
  dividendsPaid: number;
  closingRetainedEarnings: number;
  
  openingReserves: number;
  otherComprehensiveIncome: number;
  closingReserves: number;
  
  totalOpeningEquity: number;
  totalClosingEquity: number;
}

export interface FinancialNoteItem {
  noteNumber: number;
  title: string;
  category: string;
  textDescription: string;
  tableData?: {
    headers: string[];
    rows: { label: string; values: (string | number)[] }[];
    totalRow?: { label: string; values: (string | number)[] };
  };
}

export interface StatementConfigItem {
  id: string;
  statement: 'income' | 'balanceSheet' | 'cashFlow' | 'equity' | 'fixedAssets';
  section: string;
  label: string;
  isMandatory: boolean;
  isVisible: boolean;
  order: number;
  isCustom?: boolean;
  customValues?: Record<number, number>;
  sign?: 'positive' | 'negative';
}

export interface ProjectionAssumptions {
  historicalYearsCount: number; // default 3
  projectedYearsCount: number; // default 15
  selectedHistoricalYears?: number[]; // individual previous years selected by user
  selectedProjectedYears?: number[]; // individual projected years selected by user
  revenueGrowthRate: number; // % e.g. 8.5
  grossMarginPct: number; // % e.g. 38.0
  opexPctOfRevenue: number; // % e.g. 17.5
  taxRatePct: number; // % e.g. 25.0
  capexPctOfRevenue: number; // % e.g. 6.2
  depreciationRatePct: number; // % e.g. 10.0
  receivableDays: number; // DSO e.g. 45
  inventoryDays: number; // DIO e.g. 60
  payableDays: number; // DPO e.g. 42
  dividendPayoutRatio: number; // % e.g. 30
  projectedFinancingCashFlowGrowth: number; // % e.g. -4% (deleveraging)
  loanInterestRatePct?: number; // % p.a. default 15.50
}

export type StatementCellOverrides = Record<string, Record<number, number>>;
export type StatementLabelOverrides = Record<string, string>;

export type StatementFormatMode = 'section76' | 'previous_format' | 'projected_format';

export interface FinancialValidationIssue {
  id: string;
  year?: number;
  severity: 'error' | 'warning' | 'info';
  title: string;
  message: string;
}

// ---------------- PREVIOUS FINANCIALS FORMAT (Exact PDF Compliance) ---------------- //
export interface PreviousFormatYearData {
  year: number;
  dateLabel: string; // e.g. "31.12.2023"

  // Statement of Financial Position
  fixedAssetsWDV: number;
  advancesDepositsPrepayments: number;
  stockAndStores: number; // exact closingInventory
  billsReceivable: number;
  cashAndCashEquivalents: number;
  totalCurrentAssets: number;
  totalAssetsTaka: number;

  openingCapital: number;
  netProfitForYear: number; // exact netProfit
  capitalSubtotal: number;
  drawings: number;
  capitalAfterDrawings: number;
  forExpenses: number;
  accountsPayable: number;
  shortTermLoan: number;
  provisionForTaxation: number;
  totalCurrentLiabilities: number;
  totalCapitalAndLiabilities: number;

  // Statement of Comprehensive Income
  sales: number; // exact salesRevenue
  costOfGoodsSold: number;
  grossProfit: number;
  salaryAndAllowances: number;
  telephoneAndMobileBill: number;
  postageAndCourier: number;
  electricCharges: number;
  entertainment: number;
  printingAndStationery: number;
  generalExpenses: number;
  licenceAndRenewals: number;
  auditAndProfessionalFee: number;
  travellingAndConveyance: number;
  newspaperAndPeriodicals: number;
  sellingExp: number;
  depreciation: number;
  adminAndSellingTotal: number;
  netOperatingProfit: number;
  otherIncome: number;
  operatingSubtotal: number;
  financialExp: number; // exact financeCost
  incomeTaxProvision: number;
  netProfitAfterIncomeTax: number; // exact netProfit

  // Statement of Cash Flows
  cfNetIncome: number;
  cfDepreciation: number;
  cfChangeInCurrentAssets: number;
  cfChangeInCWIP: number;
  cfChangeInCurrentLiabilities: number;
  cfNetCashOperating: number;
  cfAcquisitionPPE: number;
  cfDisposalIntangible: number;
  cfNetCashInvesting: number;
  cfDrawings: number;
  cfRepaymentOfBorrowings: number; // exact repaymentOfBorrowings (reflected cleanly)
  cfOpeningBalance: number;
  cfNetCashFinancing: number;
  cfNetIncreaseCash: number;
  cfCashBeginning: number;
  cfCashEnding: number;

  // Statement of Changes in Equity
  eqOpeningShareCapital: number;
  eqShareCapitalIssued: number;
  eqClosingShareCapital: number;
  eqOpeningRetainedEarnings: number;
  eqNetProfitForYear: number;
  eqDividendsPaid: number;
  eqClosingRetainedEarnings: number;
  eqClosingReserves: number;
  eqTotalClosingEquity: number;

  // 1.00 Fixed Assets Schedule
  faOpening: number;
  faAddition: number;
  faDisposal: number;
  faTotal: number;
  faRatePct: number;
  faDepreciationCharged: number;
  faWDV: number;
  faGrossAssetBase: number; // exact fixedAssetsGross
}

export interface PreviousComparativePair {
  id: string;           // e.g. "2023_2024"
  label: string;        // e.g. "2023 & 2024"
  earlierYear: number;  // e.g. 2023
  laterYear: number;    // e.g. 2024
  data: PreviousFormatYearData[]; // [laterYearData, earlierYearData] descending
}

// ---------------- PROJECTED FINANCIALS FORMAT (Exact PDF Compliance) ---------------- //
export interface ProjectedFormatYearData {
  year: number;

  // Statement of Financial Position
  propertyPlantEquipments: number; // Schedule-A WDV
  inventories: number; // Note 3.00 (exact closingInventory)
  tradeAndOtherReceivables: number; // Note 4.00
  advanceDepositPrepayments: number; // Note 5.00
  cashAndCashEquivalents: number; // Note 6.00
  totalCurrentAssets: number;
  totalAssets: number;

  capital: number;
  additionalCapital: number;
  profitDuringYear: number; // exact netProfit
  equitySubtotal: number;
  drawings: number;
  closingCapital: number;

  longTermLoan: number; // Note 7.00
  tradeAndOtherPayables: number; // Note 8.00
  provisionsForTax: number; // Note 9.00
  shortTermLoan: number; // Note 10.00
  liabilitiesForExp: number;
  totalCurrentLiabilities: number;
  totalEquityAndLiabilities: number;

  // Statement of Profit or Loss & Other Comprehensive Income
  sales: number; // exact salesRevenue
  costOfGoodsSold: number; // Note 11.00
  grossProfit: number;
  othersIncome: number;
  grossProfitWithOtherIncome: number;
  adminAndSellingExpenses: number; // Note 12.00
  totalOperatingProfit: number;
  financialExpenses: number; // Note 13.00 (exact financeCost)
  netProfitBeforeTax: number;
  provisionForIncomeTax: number;
  netProfitAfterTax: number; // exact netProfit

  // Statement of Cash Flows
  cfNetIncome: number;
  cfChangeInventory: number;
  cfChangeReceivables: number;
  cfChangeAdvanceDeposit: number;
  cfChangeCurrentAssets: number;
  cfChangeCurrentLiabilities: number;
  cfChangeTradePayables: number;
  cfChangeProvisionsTax: number;
  cfChangeLiabilitiesExp: number;
  cfDepreciation: number;
  cfNetCashOperating: number;

  cfSalePPE: number;
  cfInvestingCapitalInflow: number;
  cfPreliminaryExpenses: number;
  cfPurchasePPE: number;
  cfPurchaseInvestment: number;
  cfNetCashInvesting: number;

  cfLoanReceived: number; // Net debt inflow / (-repaymentOfBorrowings)
  cfAdditionalCapital: number;
  cfIssuanceOfDebt: number;
  cfDividendsToStockholders: number;
  cfDrawingsOutflow: number;
  cfRedeemLongTermDebt: number;
  cfIssuanceBondsPayable: number;
  cfNetCashFinancing: number;

  cfNetCashIncrease: number;
  cfOpeningCash: number;
  cfClosingCash: number;

  // Notes Breakdown (Notes 3.00 - 13.00)
  note4ReceivablesOpening: number;
  note4ReceivablesAdded: number;
  note4ReceivablesReceived: number;
  note5AdvanceOpening: number;
  note5AdvanceAdded: number;
  note5AdvanceAdjusted: number;
  note8PayablesOpening: number;
  note8PayablesAdded: number;
  note8PayablesPaid: number;
  note9TaxOpening: number;
  note9TaxAdded: number;
  note9TaxAdjusted: number;
  note11OpeningStock: number;
  note11Purchase: number;
  note11ClosingStock: number;
  note11LabourWages: number;
  note11ElectricityBill: number;
  note11CarriageInward: number;
  note11OtherDirectExpenses: number;
  note11DirectExpensesTotal: number;

  // Schedule-A PPE
  schedOpeningBalance: number;
  schedAddition: number;
  schedDisposals: number;
  schedBalanceClosing: number;
  schedRatePct: number;
  schedDepreciation: number;
  schedWDV: number;
  schedGrossAssetBase: number; // exact fixedAssetsGross
}

// ---------------- REPAYMENT SCHEDULE INTEGRATION & DEBT SERVICING TYPES ---------------- //
export interface RepaymentInstallmentRow {
  id: string;
  installmentNo: number;
  eventType: string;
  emiDateRaw: string;
  emiDateISO: string;
  financialYear: number;
  instPrincipal: number;
  instProfit: number;
  gestationPayment: number;
  overdueGestation: number;
  overdueProfit: number;
  balance: number;
  totalInstallment: number;
  isDisbursement?: boolean;
}

export interface RepaymentScheduleMetadata {
  sourceFileName: string;
  parserMethod: 'pdf-text' | 'excel' | 'csv' | 'gemini-ocr' | 'manual-paste' | 'ababil-reference' | 'generated';
  investmentMade?: number;
  interval?: string;
  numberOfInstallments?: number;
  gestationPeriod?: string;
  disbursementDate?: string;
  firstInstallmentDate?: string;
  expiryDate?: string;
  lastInstallmentMinimized?: string;
  preparedByNote?: string;
  annualProfitRatePct?: number;
  statedTotalPrincipal?: number;
  statedTotalProfit?: number;
  dateFormatDetected: 'DD/MM/YYYY' | 'MM/DD/YYYY' | 'YYYY-MM-DD' | 'EXCEL_SERIAL';
  uploadedAt: string;
}

export interface AnnualRepaymentSummary {
  financialYear: number;
  installmentCount: number;
  totalPrincipal: number; // Exact sum of Inst. Principal -> Repayment from Borrowings and Financing
  totalProfit: number; // Exact sum of Inst. Profit -> Finance Cost
  totalGestationPayment?: number;
  totalOverdueGestation?: number;
  totalOverdueProfit?: number;
  totalDebtService: number;
  disbursementAmount: number;
  closingBalance: number;
  isSelectedInModel: boolean;
}

export interface ParsedRepaymentSchedule {
  metadata: RepaymentScheduleMetadata;
  rows: RepaymentInstallmentRow[];
  annualSummaries: AnnualRepaymentSummary[];
  grandTotalPrincipal: number;
  grandTotalProfit: number;
  grandTotalGestationPayment?: number;
  grandTotalOverdueGestation?: number;
  grandTotalOverdueProfit?: number;
  grandTotalDebtService: number;
  closingBalance?: number;
  validationErrors: string[];
  validationWarnings: string[];
  isValid: boolean;
}

export interface DebtServiceCoverageYear {
  year: number;
  isProjected: boolean;
  scheduledPrincipal: number;
  scheduledFinanceCost: number;
  totalDebtService: number;
  ebitda: number;
  operatingCashFlowBeforeFinance: number;
  ebitdaDscr: number;
  cashDscr: number;
  interestCoverage: number;
  closingCash: number;
  netCashAfterDebtService: number;
  isScheduleCovered: boolean;
  capacityStatus: 'STRONG' | 'ADEQUATE' | 'TIGHT' | 'DEFICIT';
  reschedulingContext?: string;
}

export interface RepaymentScheduleSnapshot {
  capturedAt: string;
  historicalYearsStore: Record<number, YearDataInput>;
  customProjectedOverrides: Record<number, YearDataInput>;
  assumptions: ProjectionAssumptions;
  autoFillConfigs: AutoFillConfigs;
  companyProfile: CompanyProfile;
  cellOverrides: StatementCellOverrides;
}

// ---------------- CHARTERED ACCOUNTANCY AUDIT & INSTANT RECTIFICATION TYPES ---------------- //
export type AuditFindingSeverity = 'critical' | 'warning' | 'advisory' | 'info';

export type AuditFindingCategory =
  | 'Double-Entry & Mathematical Integrity'
  | 'Cash Flow & Liquidity Reconciliation'
  | 'Fixed Assets & Equity Continuity'
  | 'Loan Rescheduling & Debt Servicing'
  | 'Financial Proportions & Realism'
  | 'Statutory & Case Disclosures';

export type AuditRectificationActionId =
  | 'CLEAR_BS_OVERRIDES'
  | 'SYNC_CF_CASH_TIE'
  | 'SYNC_EQUITY_CONTINUITY'
  | 'SYNC_FA_ROLLFORWARD'
  | 'CLEAR_INCOME_OVERRIDES'
  | 'ALIGN_FINANCE_COST_TO_LOAN_RATE'
  | 'RESTORE_SCHEDULE_LOCK'
  | 'ZERO_FIRST_PROJ_LOAN_PROCEEDS'
  | 'RECALIBRATE_UNIFORM_RATIOS'
  | 'PRUNE_EXCESSIVE_DIVIDEND_DRAIN'
  | 'NORMALIZE_ZERO_COGS_OR_WC'
  | 'SETTLE_FINAL_YEAR_LT_BORROWINGS';

export interface CharteredAuditFinding {
  id: string;
  year?: number;
  severity: AuditFindingSeverity;
  category: AuditFindingCategory;
  standardRef: string; // e.g. "IAS 1 / IAS 7", "BRPD Loan Rescheduling Norm"
  title: string;
  description: string;
  impactSummary: string;
  currentValue?: string;
  expectedValue?: string;
  isAutoRectifiable: boolean;
  requiresManualVerification: boolean;
  manualGuidance?: string;
  rectificationAction?: AuditRectificationActionId;
  rectificationPreview?: string;
}

export interface CharteredAuditPassedCheck {
  id: string;
  category: AuditFindingCategory;
  standardRef: string;
  label: string;
  detail: string;
}

export interface CharteredAuditReport {
  auditedAt: string;
  auditedYearsCount: number;
  overallOpinion: 'UNQUALIFIED_CLEAN' | 'QUALIFIED_ACTION_REQUIRED' | 'ADVERSE_MATHEMATICAL_ERROR';
  autoRectifiableCount: number;
  manualVerificationCount: number;
  criticalCount: number;
  warningCount: number;
  findings: CharteredAuditFinding[];
  passedChecks: CharteredAuditPassedCheck[];
}

export interface InstantRectificationSnapshot {
  capturedAt: string;
  appliedFindingIds: string[];
  historicalYearsStore: Record<number, YearDataInput>;
  customProjectedOverrides: Record<number, YearDataInput>;
  assumptions: ProjectionAssumptions;
  cellOverrides: StatementCellOverrides;
}


