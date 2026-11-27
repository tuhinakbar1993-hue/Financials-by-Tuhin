import {
  YearDataInput,
  IncomeStatementYear,
  BalanceSheetYear,
  CashFlowStatementYear,
  EquityStatementYear,
  ScheduleOfFixedAssetsYear,
  DebtServiceCoverageYear,
  ProjectionAssumptions,
  CompanyProfile,
  StatementCellOverrides,
  StatementConfigItem,
  ParsedRepaymentSchedule,
  CharteredAuditFinding,
  CharteredAuditPassedCheck,
  CharteredAuditReport,
} from '../types/financial';
import {
  DEFAULT_PROJECTION_ASSUMPTIONS,
  calculateScheduleOfFixedAssets,
  calculateIncomeStatement,
  calculateCashFlowStatement,
  calculateEquityStatement,
  calculateBalanceSheet,
  calculateDebtServiceCoverage,
  deriveDependentFinancials,
  recalculateYearOnLoanOrRateChange,
  generateAutomaticPrimaryDrivers,
} from './financialEngine';

export interface RunAuditParams {
  inputs: YearDataInput[];
  incomeStatements: IncomeStatementYear[];
  balanceSheets: BalanceSheetYear[];
  cashFlowStatements: CashFlowStatementYear[];
  equityStatements: EquityStatementYear[];
  fixedAssetsSchedule: ScheduleOfFixedAssetsYear[];
  debtServiceCoverage: DebtServiceCoverageYear[];
  assumptions: ProjectionAssumptions;
  companyProfile: CompanyProfile;
  cellOverrides: StatementCellOverrides;
  statementConfig: StatementConfigItem[];
  parsedSchedule?: ParsedRepaymentSchedule | null;
  isScheduleImplemented?: boolean;
}

const r2 = (n: number): number => Math.round((n + Number.EPSILON) * 100) / 100;
const yoyPct = (curr: number, prev: number): number =>
  Math.abs(prev) < 0.01 ? 0 : ((curr - prev) / Math.abs(prev)) * 100;

function fmtNum(val: number): string {
  if (!Number.isFinite(val)) return '0';
  const hasCents = Math.abs(val - Math.round(val)) >= 0.005;
  return val.toLocaleString('en-US', {
    minimumFractionDigits: hasCents ? 2 : 0,
    maximumFractionDigits: hasCents ? 2 : 0,
  });
}

export function runComprehensiveCharteredAudit(params: RunAuditParams): CharteredAuditReport {
  const {
    inputs,
    incomeStatements,
    balanceSheets,
    cashFlowStatements,
    equityStatements,
    fixedAssetsSchedule,
    debtServiceCoverage,
    assumptions,
    companyProfile,
    cellOverrides,
    statementConfig,
    parsedSchedule,
    isScheduleImplemented,
  } = params;

  const findings: CharteredAuditFinding[] = [];
  const passedChecks: CharteredAuditPassedCheck[] = [];

  const targetLoanRate = assumptions.loanInterestRatePct ?? DEFAULT_PROJECTION_ASSUMPTIONS.loanInterestRatePct ?? 15.5;
  const hiddenConfigCount = statementConfig.filter(c => !c.isVisible).length;
  const isElementActive = (id: string) => {
    const item = statementConfig.find(c => c.id === id);
    return item ? item.isVisible : true;
  };

  // Track category-level pass flags
  let bsImbalanceCount = 0;
  let cfVarianceCount = 0;
  let equityGapCount = 0;
  let faGapCount = 0;
  let incomeGapCount = 0;
  let schedLockGapCount = 0;
  let rateMismatchCount = 0;
  let uniformityCount = 0;

  // 1. Statutory & Case Metadata Completeness Checks (Manual Verification)
  const missingEntityFields: string[] = [];
  if (!companyProfile.companyName?.trim()) missingEntityFields.push('Entity Name');
  if (!companyProfile.registrationNumber?.trim()) missingEntityFields.push('Registration Number');
  if (!companyProfile.taxId?.trim()) missingEntityFields.push('Tax ID / TIN');
  if (!companyProfile.auditorName?.trim()) missingEntityFields.push('Reporting Auditor / Firm');

  if (missingEntityFields.length > 0) {
    findings.push({
      id: 'audit-missing-entity-profile',
      severity: 'advisory',
      category: 'Statutory & Case Disclosures',
      standardRef: 'IAS 1 §51 (Identification of Financial Statements)',
      title: 'Incomplete Statutory Entity & Auditor Disclosures',
      description: `Missing statutory profile details: ${missingEntityFields.join(', ')}. Complete client entity disclosures are mandatory on formal Chartered Accountancy reports and bank rescheduling submissions.`,
      impactSummary: 'Statutory headers and cover sheets will show placeholder/blank identification fields.',
      currentValue: `${4 - missingEntityFields.length}/4 completed`,
      expectedValue: '4/4 statutory fields completed',
      isAutoRectifiable: false,
      requiresManualVerification: true,
      manualGuidance: 'Open Entity Setup from the top bar and enter the client company name, registration number, TIN, and auditor details.',
    });
  } else {
    passedChecks.push({
      id: 'pass-entity-profile',
      category: 'Statutory & Case Disclosures',
      standardRef: 'IAS 1 §51',
      label: 'Statutory Entity Identification Verified',
      detail: `${companyProfile.companyName} (${companyProfile.registrationNumber}) under ${companyProfile.accountingFramework} · FY Ending ${companyProfile.financialYearEnd}.`,
    });
  }

  // Check if entire model is blank zero state
  const totalModelRevenue = inputs.reduce((acc, y) => acc + Math.abs(y.salesRevenue || 0), 0);
  if (totalModelRevenue === 0) {
    findings.push({
      id: 'audit-zero-revenue-case',
      severity: 'info',
      category: 'Statutory & Case Disclosures',
      standardRef: 'CA Engagement Readiness',
      title: 'Zero-Base Clean Case (No Operating Data Entered Yet)',
      description: 'All financial years currently carry zero revenue and zero debt balances. Enter historical client financials and rescheduling parameters to activate full ratio and servicing analysis.',
      impactSummary: 'Statements balance mathematically at zero; business viability ratios await user inputs.',
      isAutoRectifiable: false,
      requiresManualVerification: true,
      manualGuidance: 'Enter client historical figures in Financial Data Entry and apply or upload a Repayment Schedule.',
    });
  }

  // Check unimplemented uploaded Repayment Schedule
  if (parsedSchedule && parsedSchedule.isValid && !isScheduleImplemented) {
    findings.push({
      id: 'audit-unimplemented-schedule',
      severity: 'warning',
      category: 'Loan Rescheduling & Debt Servicing',
      standardRef: 'Loan Rescheduling Proposal Alignment',
      title: 'Uploaded Repayment Schedule Pending Model Integration',
      description: `A valid repayment schedule (${parsedSchedule.metadata.sourceFileName}, ${parsedSchedule.rows.filter(r => !r.isDisbursement).length} installments totaling ${fmtNum(parsedSchedule.grandTotalPrincipal)} principal) is loaded in the studio but not yet locked into projected statements.`,
      impactSummary: 'Projected Finance Cost and Principal Repayments currently reflect model estimates instead of exact bank schedule commitments.',
      currentValue: 'Not Implemented',
      expectedValue: 'Implemented & Schedule-Locked',
      isAutoRectifiable: false,
      requiresManualVerification: true,
      manualGuidance: 'Review the schedule in Financial Data Entry -> Repayment Schedule Studio and click "Implement Schedule into Financial Model".',
    });
  }

  const hasIncomeCellOvForYear = (yr: number) =>
    Object.keys(cellOverrides).some(k => k.startsWith('income_') && cellOverrides[k]?.[yr] !== undefined);
  const hasBsCellOvForYear = (yr: number) =>
    Object.keys(cellOverrides).some(k => k.startsWith('bs_') && cellOverrides[k]?.[yr] !== undefined);
  const hasCfCellOvForYear = (yr: number) =>
    Object.keys(cellOverrides).some(k => k.startsWith('cf_') && cellOverrides[k]?.[yr] !== undefined);
  const hasEqCellOvForYear = (yr: number) =>
    Object.keys(cellOverrides).some(k => k.startsWith('eq_') && cellOverrides[k]?.[yr] !== undefined);
  const hasFaCellOvForYear = (yr: number) =>
    Object.keys(cellOverrides).some(k => k.startsWith('fa_') && cellOverrides[k]?.[yr] !== undefined);

  const schedMap = new Map((parsedSchedule?.annualSummaries || []).map(s => [s.financialYear, s]));

  // 2. Per-Year Mathematical, Double-Entry, Continuity & Rescheduling Audits
  inputs.forEach((inp, idx) => {
    const yr = inp.year;
    const prevInp = idx > 0 ? inputs[idx - 1] : null;
    const inc = incomeStatements[idx];
    const bs = balanceSheets[idx];
    const cf = cashFlowStatements[idx];
    const eq = equityStatements[idx];
    const prevEq = idx > 0 ? equityStatements[idx - 1] : null;
    const fa = fixedAssetsSchedule[idx];
    const prevFa = idx > 0 ? fixedAssetsSchedule[idx - 1] : null;
    const dscr = debtServiceCoverage[idx];

    // (A) Balance Sheet Accounting Equilibrium (Assets = Liabilities + Equity)
    if (bs && Math.abs(bs.balanceCheck) > 0.05) {
      bsImbalanceCount++;
      findings.push({
        id: `audit-bs-imbalance-${yr}`,
        year: yr,
        severity: 'critical',
        category: 'Double-Entry & Mathematical Integrity',
        standardRef: 'IAS 1 §54 (Accounting Equation Equilibrium)',
        title: `Balance Sheet Imbalance (${yr})`,
        description: `Total Assets (${fmtNum(bs.totalAssets)}) do not equal Total Liabilities & Equity (${fmtNum(bs.totalLiabilitiesAndEquity)}), creating an unexplained out-of-balance variance of ${fmtNum(bs.balanceCheck)}.`,
        impactSummary: 'Statements will be rejected under statutory audit and banking credit appraisal.',
        currentValue: fmtNum(bs.balanceCheck),
        expectedValue: '0.00 Balanced',
        isAutoRectifiable: hasBsCellOvForYear(yr) || hiddenConfigCount === 0,
        requiresManualVerification: !(hasBsCellOvForYear(yr) || hiddenConfigCount === 0),
        rectificationAction: 'CLEAR_BS_OVERRIDES',
        rectificationPreview: hasBsCellOvForYear(yr)
          ? `Clear conflicting inline Balance Sheet cell overrides for ${yr} and re-derive equilibrium.`
          : `Re-derive ${yr} dependent non-current / working capital balancing items cleanly.`,
        manualGuidance: hiddenConfigCount > 0 ? 'Some Balance Sheet line items were hidden in Cash Flow Config; verify custom element totals.' : undefined,
      });
    }

    // (B) Cash Flow Closing Cash vs Balance Sheet Cash Tie-Out
    if (cf && Math.abs(cf.reconciliationDifference) >= 0.01) {
      cfVarianceCount++;
      findings.push({
        id: `audit-cf-variance-${yr}`,
        year: yr,
        severity: 'critical',
        category: 'Cash Flow & Liquidity Reconciliation',
        standardRef: 'IAS 7 §45 (Cash & Cash Equivalents Tie-Out)',
        title: `Statement of Cash Flows Tie-Out Variance (${yr})`,
        description: `Calculated Closing Cash (${fmtNum(cf.closingCashCalculated)}) diverges from Stated Balance Sheet Cash (${fmtNum(cf.balanceSheetCash)}) by ${fmtNum(cf.reconciliationDifference)}.`,
        impactSummary: 'Indirect Cash Flow fails reconciliation audit matrix.',
        currentValue: fmtNum(cf.reconciliationDifference),
        expectedValue: '0.00 Reconciled',
        isAutoRectifiable: true,
        requiresManualVerification: false,
        rectificationAction: 'SYNC_CF_CASH_TIE',
        rectificationPreview: hasCfCellOvForYear(yr) || cellOverrides.bs_cashAndEquivalents?.[yr] !== undefined
          ? `Clear overriding CF/Cash cells for ${yr} and synchronize Closing Cash to ${fmtNum(cf.closingCashCalculated)}.`
          : `Synchronize stated Balance Sheet opening/closing cash anchor for ${yr} so Derived Cash = Stated Cash (0.00 variance).`,
      });
    }

    // (C) Equity Continuity & Balance Sheet Tie-Out
    if (eq && bs) {
      const openingMismatch = prevEq && Math.abs(eq.openingRetainedEarnings - prevEq.closingRetainedEarnings) > 0.05;
      const shareOpeningMismatch = prevEq && Math.abs(eq.openingShareCapital - prevEq.closingShareCapital) > 0.05;
      const equityTieMismatch = hiddenConfigCount === 0 && Math.abs(eq.totalClosingEquity - bs.totalEquity) > 1;
      if (openingMismatch || shareOpeningMismatch || equityTieMismatch) {
        equityGapCount++;
        findings.push({
          id: `audit-eq-continuity-${yr}`,
          year: yr,
          severity: 'critical',
          category: 'Fixed Assets & Equity Continuity',
          standardRef: 'IAS 1 §106 (Statement of Changes in Equity)',
          title: `Equity Roll-Forward / Tie-Out Break (${yr})`,
          description: openingMismatch
            ? `Opening Retained Earnings (${fmtNum(eq.openingRetainedEarnings)}) does not roll forward from prior Closing Retained Earnings (${fmtNum(prevEq!.closingRetainedEarnings)}).`
            : shareOpeningMismatch
            ? `Opening Share Capital (${fmtNum(eq.openingShareCapital)}) breaks continuity with prior Closing Share Capital (${fmtNum(prevEq!.closingShareCapital)}).`
            : `Closing Equity (${fmtNum(eq.totalClosingEquity)}) does not tie to Balance Sheet Equity (${fmtNum(bs.totalEquity)}).`,
          impactSummary: 'Inter-period equity reconciliation error compromises retained profit integrity.',
          currentValue: fmtNum(openingMismatch ? eq.openingRetainedEarnings : eq.totalClosingEquity),
          expectedValue: fmtNum(openingMismatch ? prevEq!.closingRetainedEarnings : bs.totalEquity),
          isAutoRectifiable: true,
          requiresManualVerification: false,
          rectificationAction: 'SYNC_EQUITY_CONTINUITY',
          rectificationPreview: `Restore continuous opening equity roll-forward and clear broken Equity overrides for ${yr}.`,
        });
      }
    }

    // (D) Fixed Assets Cost & Depreciation Roll-Forward Integrity
    if (fa) {
      const internalCostGap = Math.abs(fa.totalCostClosing - r2(fa.totalCostOpening + fa.totalAdditions - fa.totalDisposals)) > 0.05;
      const interYearOpenGap = prevFa && Math.abs(fa.totalCostOpening - prevFa.totalCostClosing) > 0.05;
      const driverMismatch = isElementActive('fa_land') && !hasFaCellOvForYear(yr) && Math.abs(fa.totalCostClosing - inp.fixedAssetsGross) > 0.05;
      if (internalCostGap || interYearOpenGap || driverMismatch) {
        faGapCount++;
        findings.push({
          id: `audit-fa-rollforward-${yr}`,
          year: yr,
          severity: 'critical',
          category: 'Fixed Assets & Equity Continuity',
          standardRef: 'IAS 16 §73(e) (PPE Gross Cost & Depreciation Roll-Forward)',
          title: `Fixed Assets Schedule Continuity Error (${yr})`,
          description: interYearOpenGap
            ? `Opening PPE Gross Cost (${fmtNum(fa.totalCostOpening)}) differs from prior year Closing Gross Cost (${fmtNum(prevFa!.totalCostClosing)}).`
            : `PPE Opening + Additions - Disposals (${fmtNum(fa.totalCostOpening + fa.totalAdditions - fa.totalDisposals)}) does not match Closing Cost (${fmtNum(fa.totalCostClosing)}).`,
          impactSummary: 'Investing Cash Flow (PPE Capex/Disposals) and Balance Sheet NBV become inconsistent.',
          isAutoRectifiable: true,
          requiresManualVerification: false,
          rectificationAction: 'SYNC_FA_ROLLFORWARD',
          rectificationPreview: `Clear broken Fixed Asset category overrides and align PPE Additions for ${yr} so Closing = Opening + Additions - Disposals.`,
        });
      }
    }

    // (E) Income Statement Internal Arithmetic (Revenue -> GP -> EBIT -> PBT -> NPAT)
    if (inc && hiddenConfigCount === 0) {
      const gpGap = Math.abs(inc.grossProfit - r2(inc.revenue - inc.costOfSales)) > 0.05;
      const ebitGap = Math.abs(inc.operatingProfit - r2(inc.grossProfit - inc.operatingExpenses - inc.depreciationExpense)) > 0.05;
      const pbtGap = Math.abs(inc.profitBeforeTax - r2(inc.operatingProfit + inc.interestIncome - inc.financeCost)) > 0.05;
      const npGap = Math.abs(inc.netProfit - r2(inc.profitBeforeTax - inc.taxExpense)) > 0.05;
      const missingTax = inc.profitBeforeTax > 100 && (inp.taxRate || assumptions.taxRatePct) > 0 && inc.taxExpense === 0 && isElementActive('inc_taxExpense');
      if (gpGap || ebitGap || pbtGap || npGap || missingTax) {
        incomeGapCount++;
        findings.push({
          id: `audit-income-arithmetic-${yr}`,
          year: yr,
          severity: 'critical',
          category: 'Double-Entry & Mathematical Integrity',
          standardRef: 'IAS 1 §82 (Statement of Profit or Loss Vertical Arithmetic)',
          title: `Income Statement Arithmetic / Tax Anomaly (${yr})`,
          description: missingTax
            ? `Profit Before Tax is ${fmtNum(inc.profitBeforeTax)} while Corporate Tax Expense is 0 despite an active ${inp.taxRate || assumptions.taxRatePct}% statutory tax rate.`
            : `Vertical P&L arithmetic (Revenue - COGS - OpEx - Dep + Int.Inc - Finance Cost - Tax = Net Profit) contains an override break in ${yr}.`,
          impactSummary: 'Tax provisioning or net profit derivation does not reconcile vertically.',
          isAutoRectifiable: hasIncomeCellOvForYear(yr) || missingTax,
          requiresManualVerification: false,
          rectificationAction: 'CLEAR_INCOME_OVERRIDES',
          rectificationPreview: `Restore vertical P&L arithmetic & ${inp.taxRate || assumptions.taxRatePct}% corporate tax consistency for ${yr}.`,
        });
      }
    }

    // (F) Zero Cost of Sales or Implausible DSO (> 450 days) in Operating Entities
    if (inp.salesRevenue > 0 && !inp.isScheduleLocked) {
      const zeroCogs = (inp.costOfSales <= 0 || (inc && inc.costOfSales <= 0)) && isElementActive('inc_costOfSales') && companyProfile.entityType !== 'service' && companyProfile.entityType !== 'banking_conventional' && companyProfile.entityType !== 'banking_islamic';
      const dsoDays = (inp.tradeReceivables / Math.max(1, inp.salesRevenue)) * 365;
      const excessiveDso = dsoDays > 450 && isElementActive('bs_tradeReceivables');
      if (zeroCogs || excessiveDso) {
        findings.push({
          id: `audit-wc-realism-${yr}`,
          year: yr,
          severity: 'warning',
          category: 'Financial Proportions & Realism',
          standardRef: 'CA Prudence & Working Capital Turnover Benchmark',
          title: zeroCogs ? `Zero Cost of Goods Sold (${yr})` : `Excessive Receivable Days (${Math.round(dsoDays)} Days in ${yr})`,
          description: zeroCogs
            ? `Revenue of ${fmtNum(inp.salesRevenue)} is reported with zero Cost of Goods Sold for a ${companyProfile.entityType} business.`
            : `Trade Receivables (${fmtNum(inp.tradeReceivables)}) represent ${Math.round(dsoDays)} days of sales, far above prudent commercial credit terms.`,
          impactSummary: 'Distorts Gross Profit Margin, Working Capital Cash Flow, and Liquidity ratios.',
          currentValue: zeroCogs ? 'COGS = 0' : `${Math.round(dsoDays)} DSO Days`,
          expectedValue: zeroCogs ? `~${100 - assumptions.grossMarginPct}% of Sales` : `${assumptions.receivableDays}–180 Days`,
          isAutoRectifiable: true,
          requiresManualVerification: false,
          rectificationAction: 'NORMALIZE_ZERO_COGS_OR_WC',
          rectificationPreview: `Align ${yr} COGS & Trade Receivables to prudent industry proportions (${assumptions.grossMarginPct}% target margin / ${assumptions.receivableDays} DSO).`,
        });
      }
    }

    // (G) Repayment Schedule Lock Integrity & First Projected Year No-Fresh-Borrowing Rule
    if (inp.isScheduleLocked) {
      const expectedP = inp.schedulePrincipalLocked ?? 0;
      const expectedI = inp.scheduleFinanceCostLocked ?? 0;
      const schedMatch = schedMap.get(yr);
      const refP = isScheduleImplemented && schedMatch ? schedMatch.totalPrincipal : expectedP;
      const refI = isScheduleImplemented && schedMatch ? schedMatch.totalProfit : expectedI;
      if (Math.abs(inp.repaymentOfBorrowings - refP) > 0.05 || Math.abs(inp.financeCost - refI) > 0.05) {
        schedLockGapCount++;
        findings.push({
          id: `audit-sched-lock-${yr}`,
          year: yr,
          severity: 'critical',
          category: 'Loan Rescheduling & Debt Servicing',
          standardRef: 'Banking Rescheduling Sanction Schedule Compliance',
          title: `Repayment Schedule Lock Deviation (${yr})`,
          description: `Model Debt Service (Principal ${fmtNum(inp.repaymentOfBorrowings)}, Interest ${fmtNum(inp.financeCost)}) differs from exact Schedule obligations (Principal ${fmtNum(refP)}, Interest ${fmtNum(refI)}).`,
          impactSummary: 'Projected financial statements do not match the bank Repayment Schedule.',
          currentValue: `${fmtNum(inp.repaymentOfBorrowings)} / ${fmtNum(inp.financeCost)}`,
          expectedValue: `${fmtNum(refP)} / ${fmtNum(refI)}`,
          isAutoRectifiable: true,
          requiresManualVerification: false,
          rectificationAction: 'RESTORE_SCHEDULE_LOCK',
          rectificationPreview: `Restore exact schedule Principal (${fmtNum(refP)}) and Interest (${fmtNum(refI)}) for ${yr} and re-derive P&L & Cash Flow.`,
        });
      }
    }

    const firstProjIdx = inputs.findIndex(y => y.isProjected);
    if (idx === firstProjIdx && (inp.financingProceedsBorrowings || 0) > 0.01) {
      findings.push({
        id: `audit-year1-fresh-loan-${yr}`,
        year: yr,
        severity: 'warning',
        category: 'Loan Rescheduling & Debt Servicing',
        standardRef: 'Loan Rescheduling Non-Disbursement Principle',
        title: `Fresh Borrowing Proceeds in First Rescheduling Year (${yr})`,
        description: `Year ${yr} records ${fmtNum(inp.financingProceedsBorrowings || 0)} of fresh loan proceeds. Since Year 1 of projections restructures existing loans, fresh loan proceeds should be 0.`,
        impactSummary: 'Inflates Financing Cash Flow and overstates closing liquidity in Year 1.',
        currentValue: fmtNum(inp.financingProceedsBorrowings || 0),
        expectedValue: '0',
        isAutoRectifiable: true,
        requiresManualVerification: false,
        rectificationAction: 'ZERO_FIRST_PROJ_LOAN_PROCEEDS',
        rectificationPreview: `Zero out Year ${yr} fresh Loan Proceeds and recalculate Financing Cash Flow & Closing Borrowings.`,
      });
    }

    // (H) Finance Cost vs Debt Balance Alignment at Configured Loan Interest Rate (15.50% default)
    const debtBal = r2((inp.longTermBorrowings || 0) + (inp.shortTermBorrowings || 0));
    const prevDebtBal = prevInp ? r2((prevInp.longTermBorrowings || 0) + (prevInp.shortTermBorrowings || 0)) : debtBal;
    const referenceDebt = inp.isProjected && prevDebtBal > 0 ? prevDebtBal : debtBal;
    if (!inp.isScheduleLocked && referenceDebt > 0 && targetLoanRate > 0) {
      const impliedRatePct = (inp.financeCost / referenceDebt) * 100;
      const devPp = Math.abs(impliedRatePct - targetLoanRate);
      const isProjDeviation = inp.isProjected && devPp > Math.max(4.5, targetLoanRate * 0.35);
      const isCustomHistZeroInterest = !inp.isProjected && (inp.hasCustomLongTermBorrowings || inp.hasCustomShortTermBorrowings) && inp.financeCost === 0;
      if (isProjDeviation || isCustomHistZeroInterest) {
        rateMismatchCount++;
        const expectedFinCost = Math.round(referenceDebt * (targetLoanRate / 100));
        findings.push({
          id: `audit-rate-mismatch-${yr}`,
          year: yr,
          severity: 'warning',
          category: 'Loan Rescheduling & Debt Servicing',
          standardRef: `IFRS 9 Effective Interest Audit (${targetLoanRate.toFixed(2)}% Loan Rate)`,
          title: `Interest Rate vs Loan Balance Discrepancy (${yr})`,
          description: isCustomHistZeroInterest
            ? `User-entered existing loan balance (${fmtNum(debtBal)}) carries 0 Finance Cost despite an active ${targetLoanRate.toFixed(2)}% interest rate.`
            : `Implied effective borrowing rate in ${yr} is ${impliedRatePct.toFixed(2)}% (Finance Cost ${fmtNum(inp.financeCost)} on Debt ${fmtNum(referenceDebt)}), deviating significantly from configured ${targetLoanRate.toFixed(2)}%.`,
          impactSummary: 'Understates/overstates debt servicing burden, Interest Coverage, and Taxable Profit.',
          currentValue: `${impliedRatePct.toFixed(2)}% (${fmtNum(inp.financeCost)})`,
          expectedValue: `~${targetLoanRate.toFixed(2)}% (~${fmtNum(expectedFinCost)})`,
          isAutoRectifiable: true,
          requiresManualVerification: false,
          rectificationAction: 'ALIGN_FINANCE_COST_TO_LOAN_RATE',
          rectificationPreview: `Recalculate ${yr} Finance Cost at ${targetLoanRate.toFixed(2)}% (${fmtNum(expectedFinCost)}) and cascade PBT, Tax, Net Profit, and Cash Flow.`,
        });
      }
    }

    // (I) Projected Negative Cash / Excessive Dividend Drain / DSCR Adequacy
    if (inp.isProjected && cf && cf.closingCashCalculated < 0) {
      if ((inp.dividendsPaid || 0) > 0) {
        findings.push({
          id: `audit-neg-cash-div-${yr}`,
          year: yr,
          severity: 'critical',
          category: 'Cash Flow & Liquidity Reconciliation',
          standardRef: 'Corporate Solvency & Dividend Restraint Norm',
          title: `Projected Cash Deficit with Active Dividends (${yr})`,
          description: `Projected Closing Cash turns negative (${fmtNum(cf.closingCashCalculated)}) while distributing ${fmtNum(inp.dividendsPaid)} in Dividends/Drawings.`,
          impactSummary: 'Liquidity shortfall renders proposed repayment plan unbankable.',
          currentValue: `Cash ${fmtNum(cf.closingCashCalculated)} (Div ${fmtNum(inp.dividendsPaid)})`,
          expectedValue: 'Suspend dividends to restore liquidity',
          isAutoRectifiable: true,
          requiresManualVerification: false,
          rectificationAction: 'PRUNE_EXCESSIVE_DIVIDEND_DRAIN',
          rectificationPreview: `Suspend discretionary Dividends/Drawings in ${yr} to preserve cash for loan servicing.`,
        });
      } else {
        findings.push({
          id: `audit-neg-cash-manual-${yr}`,
          year: yr,
          severity: 'critical',
          category: 'Cash Flow & Liquidity Reconciliation',
          standardRef: 'Cash Flow Solvency Assessment',
          title: `Projected Cash Deficit (${yr})`,
          description: `Projected Closing Cash in ${yr} is negative (${fmtNum(cf.closingCashCalculated)}). Operating cash flows do not cover scheduled Capex and Debt Service.`,
          impactSummary: 'Bank credit committee will require extended moratorium/tenor or equity injection.',
          currentValue: fmtNum(cf.closingCashCalculated),
          expectedValue: '> 0 Positive Liquidity',
          isAutoRectifiable: false,
          requiresManualVerification: true,
          manualGuidance: 'Increase loan tenure / moratorium in Repayment Schedule Studio, adjust Capex % in Projections Engine, or verify working capital days.',
        });
      }
    } else if (inp.isProjected && dscr && dscr.totalDebtService > 0 && dscr.capacityStatus === 'DEFICIT') {
      findings.push({
        id: `audit-proj-dscr-deficit-${yr}`,
        year: yr,
        severity: 'warning',
        category: 'Loan Rescheduling & Debt Servicing',
        standardRef: 'Bank Credit Risk & DSCR Threshold (>= 1.25x)',
        title: `Projected Debt Service Shortfall (${yr}: DSCR ${dscr.ebitdaDscr.toFixed(2)}x)`,
        description: `EBITDA DSCR (${dscr.ebitdaDscr.toFixed(2)}x) or Cash DSCR (${dscr.cashDscr.toFixed(2)}x) falls below 1.00x coverage in projected year ${yr}.`,
        impactSummary: 'Projected post-rescheduling cash generation is tight against installment obligations.',
        currentValue: `EBITDA DSCR ${dscr.ebitdaDscr.toFixed(2)}x`,
        expectedValue: '>= 1.25x Recommended',
        isAutoRectifiable: false,
        requiresManualVerification: true,
        manualGuidance: 'Review Repayment Schedule tenure/installments or calibrate projected Revenue Growth & Gross Margin in Projections Engine.',
      });
    }

    // (J) Mechanical / Uniform YoY Growth Spread Check (>= 1.2pp realism across projected years)
    if (inp.isProjected && prevInp && prevInp.salesRevenue > 0 && inp.salesRevenue > 0 && !inp.isScheduleLocked) {
      const gRev = yoyPct(inp.salesRevenue, prevInp.salesRevenue);
      const gCos = yoyPct(inp.costOfSales, prevInp.costOfSales);
      const gOpex = yoyPct(inp.operatingExpenses, prevInp.operatingExpenses);
      const gInv = yoyPct(inp.closingInventory, prevInp.closingInventory);
      const gRec = yoyPct(inp.tradeReceivables, prevInp.tradeReceivables);
      const spreads = [
        { name: 'Revenue vs COGS', diff: Math.abs(gRev - gCos) },
        { name: 'Revenue vs OpEx', diff: Math.abs(gRev - gOpex) },
        { name: 'COGS vs OpEx', diff: Math.abs(gCos - gOpex) },
        { name: 'Inventory vs COGS', diff: Math.abs(gInv - gCos) },
        { name: 'Receivables vs Inventory', diff: Math.abs(gRec - gInv) },
      ];
      const uniformOnes = spreads.filter(s => s.diff < 1.15);
      if (uniformOnes.length > 0 && totalModelRevenue > 0) {
        uniformityCount++;
        findings.push({
          id: `audit-uniform-spread-${yr}`,
          year: yr,
          severity: 'warning',
          category: 'Financial Proportions & Realism',
          standardRef: 'CA Financial Modeling Realism (>= 1.20% YoY Variance)',
          title: `Mechanical / Uniform Year-Over-Year Pattern (${yr})`,
          description: `${uniformOnes.map(u => `${u.name} (${u.diff.toFixed(2)}pp)`).join(', ')} show near-identical percentage changes (< 1.20% spread), appearing artificially flat for bank appraisal.`,
          impactSummary: 'Lenders flag uniform percentage-scaled projections as non-commercial.',
          currentValue: `${Math.min(...uniformOnes.map(u => u.diff)).toFixed(2)}pp min spread`,
          expectedValue: '>= 1.20pp natural variance',
          isAutoRectifiable: true,
          requiresManualVerification: false,
          rectificationAction: 'RECALIBRATE_UNIFORM_RATIOS',
          rectificationPreview: `Apply realistic commercial variation (>= 1.2% spread) across ${yr} operating and working capital ratios while preserving user primary inputs.`,
        });
      }
    }
  });

  // Record clean Audit Assertions in Passed Checks
  if (bsImbalanceCount === 0) {
    passedChecks.push({
      id: 'pass-bs-equilibrium',
      category: 'Double-Entry & Mathematical Integrity',
      standardRef: 'IAS 1 §54',
      label: 'Balance Sheet Equilibrium Verified (0.00 Variance)',
      detail: `Total Assets equal Total Liabilities & Equity across all ${inputs.length} audited years.`,
    });
  }
  if (cfVarianceCount === 0) {
    passedChecks.push({
      id: 'pass-cf-tie',
      category: 'Cash Flow & Liquidity Reconciliation',
      standardRef: 'IAS 7 §45',
      label: 'Cash Flow Indirect Reconciliation Clean',
      detail: `Closing Cash Derived matches Stated Balance Sheet Cash across all ${inputs.length} periods.`,
    });
  }
  if (equityGapCount === 0 && faGapCount === 0) {
    passedChecks.push({
      id: 'pass-continuity',
      category: 'Fixed Assets & Equity Continuity',
      standardRef: 'IAS 1 / IAS 16',
      label: 'PPE & Equity Inter-Period Continuity Verified',
      detail: `Closing-to-opening roll-forwards for Property, Plant & Equipment and Retained Earnings hold cleanly across all periods.`,
    });
  }
  if (incomeGapCount === 0) {
    passedChecks.push({
      id: 'pass-income-vertical',
      category: 'Double-Entry & Mathematical Integrity',
      standardRef: 'IAS 1 §82',
      label: 'Income Statement Vertical Arithmetic & Tax Logic Passed',
      detail: `Gross Profit, Operating Profit, PBT, Corporate Tax, and Net Profit After Tax tie out across every period.`,
    });
  }
  if (rateMismatchCount === 0 && schedLockGapCount === 0) {
    passedChecks.push({
      id: 'pass-debt-rate',
      category: 'Loan Rescheduling & Debt Servicing',
      standardRef: `IFRS 9 @ ${targetLoanRate.toFixed(2)}%`,
      label: 'Rescheduling Debt, Interest Rate & Schedule Locks Synchronized',
      detail: `Loan balances, ${targetLoanRate.toFixed(2)}% interest rate calibration, and any locked Repayment Schedule obligations are fully aligned.`,
    });
  }
  if (uniformityCount === 0 && totalModelRevenue > 0) {
    passedChecks.push({
      id: 'pass-realism',
      category: 'Financial Proportions & Realism',
      standardRef: 'CA Appraisal Prudence (>= 1.2% Spread)',
      label: 'Realistic Uneven Financial Ratios Confirmed',
      detail: `Projected revenues, costs, operating expenses, and working capital turnover exhibit realistic uneven commercial variation (>= 1.20%).`,
    });
  }

  // Informational verification of historical pre-rescheduling justification when present
  const histDistressed = debtServiceCoverage.filter(d => !d.isProjected && d.totalDebtService > 0 && d.ebitdaDscr < 1.0);
  const projViable = debtServiceCoverage.filter(d => d.isProjected && d.totalDebtService > 0 && d.ebitdaDscr >= 1.25);
  if (histDistressed.length > 0 && projViable.length > 0) {
    passedChecks.push({
      id: 'pass-resched-justification',
      category: 'Loan Rescheduling & Debt Servicing',
      standardRef: 'Credit Rescheduling Viability Proof',
      label: 'Pre-Rescheduling Distress vs. Post-Rescheduling Viability Confirmed',
      detail: `Historical servicing deficit (DSCR < 1.00x in ${histDistressed.map(h => h.year).join(', ')}) transitions cleanly to sustainable projected coverage (>= 1.25x).`,
    });
  }

  const criticalCount = findings.filter(f => f.severity === 'critical').length;
  const warningCount = findings.filter(f => f.severity === 'warning').length;
  const autoRectifiableCount = findings.filter(f => f.isAutoRectifiable).length;
  const manualVerificationCount = findings.filter(f => f.requiresManualVerification).length;

  const overallOpinion: CharteredAuditReport['overallOpinion'] =
    criticalCount > 0
      ? 'ADVERSE_MATHEMATICAL_ERROR'
      : warningCount > 0 || manualVerificationCount > 0
      ? 'QUALIFIED_ACTION_REQUIRED'
      : 'UNQUALIFIED_CLEAN';

  return {
    auditedAt: new Date().toISOString(),
    auditedYearsCount: inputs.length,
    overallOpinion,
    autoRectifiableCount,
    manualVerificationCount,
    criticalCount,
    warningCount,
    findings,
    passedChecks,
  };
}

export interface InstantRectificationParams {
  findingIdsToApply?: string[];
  allYearsInputs: YearDataInput[];
  historicalYearsStore: Record<number, YearDataInput>;
  customProjectedOverrides: Record<number, YearDataInput>;
  assumptions: ProjectionAssumptions;
  cellOverrides: StatementCellOverrides;
  statementConfig: StatementConfigItem[];
  companyProfile: CompanyProfile;
  parsedSchedule?: ParsedRepaymentSchedule | null;
  isScheduleImplemented?: boolean;
}

export interface InstantRectificationResult {
  historicalYearsStore: Record<number, YearDataInput>;
  customProjectedOverrides: Record<number, YearDataInput>;
  assumptions: ProjectionAssumptions;
  cellOverrides: StatementCellOverrides;
  appliedFindingIds: string[];
  postAuditReport: CharteredAuditReport;
}

function clearOverridesMatching(
  overrides: StatementCellOverrides,
  year: number,
  prefixes: string[],
  exactKeys: string[] = []
): StatementCellOverrides {
  const next: StatementCellOverrides = {};
  for (const [k, yrMap] of Object.entries(overrides)) {
    const matches = prefixes.some(p => k.startsWith(p)) || exactKeys.includes(k);
    if (!matches || yrMap[year] === undefined) {
      next[k] = yrMap;
      continue;
    }
    const copy = { ...yrMap };
    delete copy[year];
    if (Object.keys(copy).length > 0) {
      next[k] = copy;
    }
  }
  return next;
}

export function applyInstantAuditRectifications(
  params: InstantRectificationParams
): InstantRectificationResult {
  const {
    findingIdsToApply,
    allYearsInputs,
    assumptions,
    statementConfig,
    companyProfile,
    parsedSchedule,
    isScheduleImplemented,
  } = params;

  const nextHistStore: Record<number, YearDataInput> = JSON.parse(JSON.stringify(params.historicalYearsStore));
  const nextProjOverrides: Record<number, YearDataInput> = JSON.parse(JSON.stringify(params.customProjectedOverrides));
  let nextOverrides: StatementCellOverrides = JSON.parse(JSON.stringify(params.cellOverrides));
  const nextAssumptions: ProjectionAssumptions = { ...assumptions };

  // Build initial chronological working list of years
  const workingYears: YearDataInput[] = allYearsInputs.map(y => JSON.parse(JSON.stringify(y)));
  const targetIds = findingIdsToApply ? new Set(findingIdsToApply) : null;
  const appliedFindingIds: string[] = [];
  const appliedSet = new Set<string>();
  const touchedYears = new Set<number>();
  const schedMap = new Map((parsedSchedule?.annualSummaries || []).map(s => [s.financialYear, s]));

  // Run up to 3 convergence passes when applying full Instant Rectification so secondary cascades resolve cleanly
  const maxPasses = targetIds ? 1 : 3;
  for (let pass = 0; pass < maxPasses; pass++) {
    const calcOpts = { statementConfig, cellOverrides: nextOverrides };
    const faCur = calculateScheduleOfFixedAssets(workingYears, calcOpts);
    const incCur = calculateIncomeStatement(workingYears, faCur, calcOpts);
    const cfCur = calculateCashFlowStatement(workingYears, incCur, faCur, calcOpts);
    const eqCur = calculateEquityStatement(workingYears, incCur, calcOpts);
    const bsCur = calculateBalanceSheet(workingYears, eqCur, faCur, cfCur, calcOpts, incCur);
    const dscrCur = calculateDebtServiceCoverage(workingYears, incCur, bsCur, cfCur, isScheduleImplemented ? parsedSchedule : null, calcOpts);

    const passAudit = runComprehensiveCharteredAudit({
      inputs: workingYears,
      incomeStatements: incCur,
      balanceSheets: bsCur,
      cashFlowStatements: cfCur,
      equityStatements: eqCur,
      fixedAssetsSchedule: faCur,
      debtServiceCoverage: dscrCur,
      assumptions: nextAssumptions,
      companyProfile,
      cellOverrides: nextOverrides,
      statementConfig,
      parsedSchedule,
      isScheduleImplemented,
    });

    const actionableFindings = passAudit.findings.filter(
      f => f.isAutoRectifiable && (!targetIds || targetIds.has(f.id))
    );
    if (actionableFindings.length === 0) break;

    let madeProgress = false;
    for (const finding of actionableFindings) {
      const yr = finding.year;
      if (yr === undefined) continue;
      const idx = workingYears.findIndex(y => y.year === yr);
      if (idx === -1) continue;
      const prevYr = idx > 0 ? workingYears[idx - 1] : null;
      const currYr = workingYears[idx];
      const action = finding.rectificationAction;
      if (!action) continue;

      if (action === 'CLEAR_BS_OVERRIDES') {
        const hadBsOv = Object.keys(nextOverrides).some(
          k => k.startsWith('bs_') && nextOverrides[k]?.[yr] !== undefined
        );
        nextOverrides = clearOverridesMatching(nextOverrides, yr, ['bs_']);
        if (!hadBsOv) {
          workingYears[idx] = deriveDependentFinancials(currYr, prevYr, nextAssumptions, currYr);
        }
      } else if (action === 'CLEAR_INCOME_OVERRIDES') {
        const hadIncOv = Object.keys(nextOverrides).some(
          k => k.startsWith('income_') && nextOverrides[k]?.[yr] !== undefined
        );
        nextOverrides = clearOverridesMatching(nextOverrides, yr, ['income_']);
        if (currYr.taxRate <= 0 && (nextAssumptions.taxRatePct || 25) > 0) {
          workingYears[idx] = { ...currYr, taxRate: nextAssumptions.taxRatePct || 25 };
        }
        if (!hadIncOv) {
          workingYears[idx] = deriveDependentFinancials(workingYears[idx], prevYr, nextAssumptions, workingYears[idx]);
        }
      } else if (action === 'SYNC_EQUITY_CONTINUITY') {
        nextOverrides = clearOverridesMatching(nextOverrides, yr, ['eq_']);
        const cleaned = idx > 0 ? { ...currYr, retainedEarningsOpening: undefined, reservesOpening: undefined } : currYr;
        workingYears[idx] = deriveDependentFinancials(cleaned, prevYr, nextAssumptions, cleaned);
      } else if (action === 'SYNC_FA_ROLLFORWARD') {
        nextOverrides = clearOverridesMatching(nextOverrides, yr, ['fa_']);
        const openingGross = prevYr ? prevYr.fixedAssetsGross : Math.max(0, currYr.fixedAssetsGross - currYr.fixedAssetsAdditions + currYr.fixedAssetsDisposals);
        const fixedAdditions = Math.max(0, r2(currYr.fixedAssetsGross - openingGross + currYr.fixedAssetsDisposals));
        const updated = { ...currYr, fixedAssetsAdditions: fixedAdditions };
        workingYears[idx] = deriveDependentFinancials(updated, prevYr, nextAssumptions, updated);
      } else if (action === 'RESTORE_SCHEDULE_LOCK') {
        const s = isScheduleImplemented ? schedMap.get(yr) : undefined;
        const lockedP = s ? s.totalPrincipal : (currYr.schedulePrincipalLocked ?? currYr.repaymentOfBorrowings);
        const lockedI = s ? s.totalProfit : (currYr.scheduleFinanceCostLocked ?? currYr.financeCost);
        nextOverrides = clearOverridesMatching(nextOverrides, yr, [], ['income_financeCost', 'cf_repaymentOfBorrowings', 'cf_financingRepaymentBorrowings', 'cf_adjFinanceCost', 'cf_financeCostPaid']);
        const updated = deriveDependentFinancials(
          { ...currYr, repaymentOfBorrowings: lockedP, financeCost: lockedI },
          prevYr,
          nextAssumptions,
          {
            ...currYr,
            repaymentOfBorrowings: lockedP,
            financeCost: lockedI,
            schedulePrincipalLocked: lockedP,
            scheduleFinanceCostLocked: lockedI,
          }
        );
        workingYears[idx] = updated;
      } else if (action === 'ZERO_FIRST_PROJ_LOAN_PROCEEDS') {
        const removed = currYr.financingProceedsBorrowings || 0;
        const updated: YearDataInput = {
          ...currYr,
          financingProceedsBorrowings: 0,
          cashFlowFromFinancing: r2((currYr.cashFlowFromFinancing || 0) - removed),
          longTermBorrowings: Math.max(0, r2(currYr.longTermBorrowings - removed)),
        };
        nextOverrides = clearOverridesMatching(nextOverrides, yr, [], ['cf_financingProceedsBorrowings']);
        workingYears[idx] = updated;
      } else if (action === 'ALIGN_FINANCE_COST_TO_LOAN_RATE') {
        nextOverrides = clearOverridesMatching(nextOverrides, yr, [], ['income_financeCost', 'cf_adjFinanceCost', 'cf_financeCostPaid']);
        workingYears[idx] = recalculateYearOnLoanOrRateChange(currYr, prevYr, nextAssumptions, {
          loanRateOverridePct: nextAssumptions.loanInterestRatePct ?? 15.5,
          recalculateFinanceCostFromDebt: true,
        });
      } else if (action === 'NORMALIZE_ZERO_COGS_OR_WC') {
        const targetCogs = Math.round(currYr.salesRevenue * ((100 - nextAssumptions.grossMarginPct) / 100));
        const targetRec = Math.round((currYr.salesRevenue / 365) * (nextAssumptions.receivableDays || 45));
        const updated: YearDataInput = {
          ...currYr,
          costOfSales: currYr.costOfSales <= 0 ? targetCogs : currYr.costOfSales,
          tradeReceivables:
            (currYr.tradeReceivables / Math.max(1, currYr.salesRevenue)) * 365 > 450
              ? targetRec
              : currYr.tradeReceivables,
        };
        workingYears[idx] = deriveDependentFinancials(updated, prevYr, nextAssumptions, updated);
      } else if (action === 'PRUNE_EXCESSIVE_DIVIDEND_DRAIN') {
        const updated: YearDataInput = {
          ...currYr,
          dividendsPaid: 0,
          financingDividendsPaid: 0,
          cashFlowFromFinancing: r2(
            (currYr.financingProceedsShares || 0) +
              (currYr.financingProceedsBorrowings || 0) -
              (currYr.repaymentOfBorrowings || 0)
          ),
        };
        nextOverrides = clearOverridesMatching(nextOverrides, yr, [], ['eq_dividendsPaid', 'cf_financingDividendsPaid']);
        workingYears[idx] = updated;
      } else if (action === 'RECALIBRATE_UNIFORM_RATIOS' && prevYr) {
        const autoDrivers = generateAutomaticPrimaryDrivers(prevYr, nextAssumptions, yr);
        const blendedDrivers = {
          year: yr,
          isProjected: currYr.isProjected,
          salesRevenue: currYr.salesRevenue,
          closingInventory: autoDrivers.closingInventory,
          financeCost: currYr.isScheduleLocked ? currYr.financeCost : autoDrivers.financeCost,
          netProfit: currYr.isScheduleLocked ? currYr.netProfit : autoDrivers.netProfit,
          repaymentOfBorrowings: currYr.isScheduleLocked ? currYr.repaymentOfBorrowings : currYr.repaymentOfBorrowings,
          fixedAssetsGross: currYr.fixedAssetsGross,
        };
        workingYears[idx] = deriveDependentFinancials(blendedDrivers, prevYr, nextAssumptions, {
          ...currYr,
          ...blendedDrivers,
        });
      } else if (action === 'SYNC_CF_CASH_TIE') {
        nextOverrides = clearOverridesMatching(nextOverrides, yr, ['cf_'], ['bs_cashAndEquivalents']);
        const tmpOpts = { statementConfig, cellOverrides: nextOverrides };
        const tFa = calculateScheduleOfFixedAssets(workingYears, tmpOpts);
        const tInc = calculateIncomeStatement(workingYears, tFa, tmpOpts);
        const tCf = calculateCashFlowStatement(workingYears, tInc, tFa, tmpOpts);
        const targetClosingCash = tCf[idx]?.closingCashCalculated ?? currYr.cashAndEquivalentsBS;
        workingYears[idx] = { ...workingYears[idx], cashAndEquivalentsBS: targetClosingCash };
      } else {
        continue;
      }

      touchedYears.add(yr);
      madeProgress = true;
      if (!appliedSet.has(finding.id)) {
        appliedSet.add(finding.id);
        appliedFindingIds.push(finding.id);
      }
    }

    if (!madeProgress) break;
  }

  // Persist changes into historical store & custom projected overrides
  for (const y of workingYears) {
    if (!y.isProjected) {
      nextHistStore[y.year] = { ...y, isProjected: false };
    } else {
      const hadOverride = Boolean(nextProjOverrides[y.year]);
      const yrWasTouched = touchedYears.has(y.year);
      if (hadOverride || yrWasTouched || isScheduleImplemented) {
        nextProjOverrides[y.year] = { ...y, isProjected: true };
      }
    }
  }

  // Re-run Audit to verify post-rectification cleanliness
  const postOpts = { statementConfig, cellOverrides: nextOverrides };
  const pFa = calculateScheduleOfFixedAssets(workingYears, postOpts);
  const pInc = calculateIncomeStatement(workingYears, pFa, postOpts);
  const pCf = calculateCashFlowStatement(workingYears, pInc, pFa, postOpts);
  const pEq = calculateEquityStatement(workingYears, pInc, postOpts);
  const pBs = calculateBalanceSheet(workingYears, pEq, pFa, pCf, postOpts, pInc);
  const pDscr = calculateDebtServiceCoverage(workingYears, pInc, pBs, pCf, isScheduleImplemented ? parsedSchedule : null, postOpts);

  const postAuditReport = runComprehensiveCharteredAudit({
    inputs: workingYears,
    incomeStatements: pInc,
    balanceSheets: pBs,
    cashFlowStatements: pCf,
    equityStatements: pEq,
    fixedAssetsSchedule: pFa,
    debtServiceCoverage: pDscr,
    assumptions: nextAssumptions,
    companyProfile,
    cellOverrides: nextOverrides,
    statementConfig,
    parsedSchedule,
    isScheduleImplemented,
  });

  return {
    historicalYearsStore: nextHistStore,
    customProjectedOverrides: nextProjOverrides,
    assumptions: nextAssumptions,
    cellOverrides: nextOverrides,
    appliedFindingIds,
    postAuditReport,
  };
}
