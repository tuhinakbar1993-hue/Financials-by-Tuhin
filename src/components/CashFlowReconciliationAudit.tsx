import React from 'react';
import { Scale, CheckCircle2, AlertTriangle, ShieldCheck, CalendarCheck2, ArrowUpRight } from 'lucide-react';
import {
  CashFlowStatementYear,
  BalanceSheetYear,
  CompanyProfile,
  FinancialValidationIssue,
  DebtServiceCoverageYear,
  ParsedRepaymentSchedule,
} from '../types/financial';

function fmtAuditNum(val: number): string {
  if (Math.abs(val) < 0.005) return '0';
  const absVal = Math.abs(val);
  const hasCents = Math.abs(absVal - Math.round(absVal)) >= 0.005;
  const s = hasCents
    ? absVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : Math.round(absVal).toLocaleString('en-US');
  return val < 0 ? `-${s}` : s;
}

interface CashFlowReconciliationAuditProps {
  cashFlowYears: CashFlowStatementYear[];
  balanceSheetYears: BalanceSheetYear[];
  companyProfile: CompanyProfile;
  validationIssues?: FinancialValidationIssue[];
  debtServiceCoverage?: DebtServiceCoverageYear[];
  isScheduleImplemented?: boolean;
  parsedSchedule?: ParsedRepaymentSchedule | null;
  onJumpToRepaymentSchedule?: () => void;
}

export const CashFlowReconciliationAudit: React.FC<CashFlowReconciliationAuditProps> = ({
  cashFlowYears,
  balanceSheetYears,
  companyProfile,
  validationIssues = [],
  debtServiceCoverage = [],
  isScheduleImplemented = false,
  parsedSchedule = null,
  onJumpToRepaymentSchedule,
}) => {
  const reconciledCount = cashFlowYears.filter(cf => cf.reconciliationDifference === 0).length;
  const isAllReconciled = reconciledCount === cashFlowYears.length;
  const balancedBSCount = balanceSheetYears.filter(bs => Math.abs(bs.balanceCheck) <= 1).length;
  const scheduleCoverageYears = debtServiceCoverage.filter(d => d.totalDebtService > 0);

  return (
    <div className="space-y-6">
      {/* Executive Header Banner */}
      <div className="p-5 sm:p-6 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl shadow-sm">
        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-teal-50 border border-teal-200 flex items-center justify-center text-teal-700">
                <Scale className="w-4 h-4" />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-lg font-bold text-slate-900 tracking-tight">
                    Cash Flow & Balance Sheet Reconciliation Matrix
                  </h2>
                  <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-teal-800 bg-teal-50 border border-teal-200 px-2 py-0.5 rounded shadow-xs">
                    Section 76.3 & 76.4
                  </span>
                </div>
                <p className="text-xs text-slate-600 mt-0.5">
                  Audited mathematical reconciliation between Statement of Cash Flows closing cash and Statement of Financial Position stated cash.
                </p>
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            {onJumpToRepaymentSchedule && (
              <button
                type="button"
                onClick={onJumpToRepaymentSchedule}
                className="px-3 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-colors shadow-sm"
              >
                <CalendarCheck2 className="w-3.5 h-3.5" />
                <span>{isScheduleImplemented ? 'Repayment Schedule Active' : 'Repayment Schedule'}</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            )}
            <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 px-3.5 py-2 rounded-xl">
              <div className={`w-2.5 h-2.5 rounded-full ${isAllReconciled ? 'bg-emerald-600' : 'bg-amber-500'}`} />
              <span className="text-xs font-mono font-bold text-slate-900">
                {reconciledCount} / {cashFlowYears.length} Periods Reconciled · BS {balancedBSCount}/{balanceSheetYears.length}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Model-Wide Validation & Fixed-Driver Integrity Banner */}
      <div className={`p-4 rounded-2xl border text-xs flex flex-col gap-2 backdrop-blur-md ${
        validationIssues.length === 0
          ? 'bg-emerald-50/90 border-emerald-200 text-emerald-900'
          : 'bg-amber-50/90 border-amber-300 text-amber-900'
      }`}>
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 font-bold">
            {validationIssues.length === 0 ? (
              <ShieldCheck className="w-4 h-4 text-emerald-600 shrink-0" />
            ) : (
              <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
            )}
            <span>
              {validationIssues.length === 0
                ? 'Model Validation Passed: All 6 Fixed Primary Inputs Strictly Preserved · Zero Accounting Variance Across All Selected Years'
                : `${validationIssues.length} Validation Check(s) Identified`}
            </span>
          </div>
          <span className="text-[10px] font-mono uppercase px-2 py-0.5 rounded bg-white border border-slate-200 text-slate-700 font-semibold">
            {cashFlowYears.length} Years Audited
          </span>
        </div>
        {validationIssues.length > 0 && (
          <ul className="space-y-1 pl-6 list-disc text-[11px] text-amber-800">
            {validationIssues.slice(0, 6).map(issue => (
              <li key={issue.id}>
                <strong>{issue.title}:</strong> {issue.message}
              </li>
            ))}
          </ul>
        )}
      </div>

      {/* Accounting Methodology Callout Strip */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
        <div className="p-4 bg-white/90 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-900 font-bold flex items-center gap-2">
            <div className="w-5 h-5 rounded bg-sky-50 border border-sky-200 flex items-center justify-center text-sky-700 text-[10px] font-mono font-bold">
              1
            </div>
            <span>Operating Cash Flow</span>
          </div>
          <p className="text-slate-600 leading-relaxed text-[11px]">
            Derived via the <strong>Indirect Method</strong> from Net Profit Before Tax, adjusting for non-cash depreciation, finance cost, and working capital shifts (Inventory, Receivables, Payables).
          </p>
        </div>

        <div className="p-4 bg-white/90 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-900 font-bold flex items-center gap-2">
            <div className="w-5 h-5 rounded bg-indigo-50 border border-indigo-200 flex items-center justify-center text-indigo-700 text-[10px] font-mono font-bold">
              2
            </div>
            <span>Investing Cash Flow</span>
          </div>
          <p className="text-slate-600 leading-relaxed text-[11px]">
            Tied directly to the <strong>Schedule of Fixed Assets</strong> capital expenditure additions, asset disposals, and financial investment allocations.
          </p>
        </div>

        <div className="p-4 bg-white/90 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-900 font-bold flex items-center gap-2">
            <div className="w-5 h-5 rounded bg-amber-50 border border-amber-200 flex items-center justify-center text-amber-800 text-[10px] font-mono font-bold">
              3
            </div>
            <span>Financing Cash Flow</span>
          </div>
          <p className="text-slate-600 leading-relaxed text-[11px]">
            Preserved as a <strong>mandatory direct user entry</strong> separately for every historical and projected financial year. Never derived or forced.
          </p>
        </div>
      </div>

      {/* Reconciliation Audit Table */}
      <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
        <div className="p-4 border-b border-slate-200 bg-slate-50/90 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-mono uppercase tracking-widest text-emerald-700 font-bold">
                {companyProfile.companyName}
              </span>
              <span className="text-slate-400">·</span>
              <span className="text-[10px] font-mono text-slate-600">
                Audited Schedule
              </span>
            </div>
            <h3 className="text-sm font-bold text-slate-900 tracking-tight mt-0.5">
              Multi-Year Cash Flow Audit Reconciliation Schedule
            </h3>
            <p className="text-[11px] text-slate-600 mt-0.5">
              All monetary figures in {companyProfile.reportingCurrency} ({companyProfile.reportingUnit}) · Complete {cashFlowYears.length}-Year Horizon
            </p>
          </div>
          <span className="text-xs font-mono font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded shadow-xs self-start sm:self-auto">
            {cashFlowYears.length} Financial Years Audited
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs font-mono">
            <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
              <tr>
                <th className="py-3 px-4 font-sans font-semibold text-slate-900 tracking-wide">
                  Audit Parameter / Cash Flow Classification
                </th>
                {cashFlowYears.map(cf => (
                  <th key={cf.year} className={`py-2.5 px-3 text-right ${cf.isProjected ? 'bg-sky-50/90 text-sky-900' : 'text-slate-800'}`}>
                    <div className="flex flex-col items-end">
                      <span className="font-bold text-xs">{cf.year}</span>
                      <span className={`text-[9px] font-sans uppercase tracking-wider font-semibold ${
                        cf.isProjected ? 'text-sky-700' : 'text-slate-500'
                      }`}>
                        {cf.isProjected ? 'Forecast' : 'Audited'}
                      </span>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200/70 tabular-nums">
              <tr className="hover:bg-slate-50 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-800 font-medium">A. Operating Cash Flow (Calculated)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-medium ${cf.netCashFromOperating >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                    {fmtAuditNum(cf.netCashFromOperating)}
                  </td>
                ))}
              </tr>

              <tr className="hover:bg-slate-50 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-800 font-medium">B. Investing Cash Flow (Calculated)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-medium ${cf.netCashFromInvesting >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                    {fmtAuditNum(cf.netCashFromInvesting)}
                  </td>
                ))}
              </tr>

              <tr className="bg-amber-50/75 font-semibold hover:bg-amber-50 transition-colors">
                <td className="py-2.5 px-4 font-sans text-amber-900 flex items-center gap-2">
                  <span>C. Financing Cash Flow (Direct User Entry)</span>
                  <span className="text-[9px] font-mono uppercase bg-amber-100 text-amber-900 border border-amber-300 px-1.5 py-0.2 rounded">
                    Mandatory
                  </span>
                </td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-amber-800 font-bold">
                    {fmtAuditNum(cf.directlyEnteredFinancingCashFlow)}
                  </td>
                ))}
              </tr>

              <tr className="border-t border-slate-300 font-semibold bg-slate-100/80">
                <td className="py-2.5 px-4 font-sans text-slate-900">Net Change in Cash (A + B + C)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-bold ${cf.netChangeInCash >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                    {fmtAuditNum(cf.netChangeInCash)}
                  </td>
                ))}
              </tr>

              <tr className="hover:bg-slate-50 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-600">Opening Cash & Cash Equivalents</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-slate-700">
                    {fmtAuditNum(cf.openingCash)}
                  </td>
                ))}
              </tr>

              {/* Closing Cash Derived with Statutory Double Underline */}
              <tr className="font-bold text-emerald-800 bg-emerald-50/85 border-b-4 border-double border-emerald-600 border-t border-emerald-200">
                <td className="py-3 px-4 font-sans text-sm">Closing Cash Derived (Opening + Net Change)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-3 px-3 text-right text-emerald-800 font-bold text-sm">
                    {fmtAuditNum(cf.closingCashCalculated)}
                  </td>
                ))}
              </tr>

              <tr className="font-semibold text-slate-800 hover:bg-sky-50/40 transition-colors">
                <td className="py-2.5 px-4 font-sans text-sky-800">Stated Cash in Statement of Financial Position</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-sky-800 font-semibold">
                    {fmtAuditNum(cf.balanceSheetCash)}
                  </td>
                ))}
              </tr>

              <tr className="font-bold border-t border-slate-200 bg-slate-50/90">
                <td className="py-3 px-4 font-sans text-slate-900 flex items-center gap-1.5">
                  <span>Reconciliation Difference</span>
                  <span className="text-[10px] text-slate-500 font-normal font-sans">(Derived - Stated)</span>
                </td>
                {cashFlowYears.map(cf => {
                  const isZero = Math.abs(cf.reconciliationDifference) < 0.005;
                  return (
                    <td key={cf.year} className={`py-3 px-3 text-right ${isZero ? 'text-emerald-700 font-bold' : 'text-amber-700 font-bold'}`}>
                      {cf.reconciliationDifference > 0 ? `+${fmtAuditNum(cf.reconciliationDifference)}` : fmtAuditNum(cf.reconciliationDifference)}
                    </td>
                  );
                })}
              </tr>

              <tr className="bg-white font-sans">
                <td className="py-2.5 px-4 text-slate-600 text-[11px] font-medium">Audit Verdict</td>
                {cashFlowYears.map(cf => {
                  const isZero = Math.abs(cf.reconciliationDifference) < 0.005;
                  return (
                    <td key={cf.year} className="py-2.5 px-3 text-right">
                      {isZero ? (
                        <span className="inline-flex items-center gap-1 text-[10px] font-bold text-emerald-800 bg-emerald-50 border border-emerald-300 px-2 py-0.5 rounded">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" /> Reconciled
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 text-[10px] font-bold text-amber-900 bg-amber-50 border border-amber-300 px-2 py-0.5 rounded">
                          <AlertTriangle className="w-3.5 h-3.5 text-amber-600" /> Variance
                        </span>
                      )}
                    </td>
                  );
                })}
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {/* Debt Servicing & Repayment Capacity Audit Matrix */}
      {scheduleCoverageYears.length > 0 && (
        <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-4 border-b border-slate-200 bg-slate-50/90 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-mono uppercase tracking-widest text-emerald-700 font-bold">
                  {parsedSchedule?.metadata?.sourceFileName || companyProfile.companyName}
                </span>
                {parsedSchedule?.metadata?.investmentMade && (
                  <>
                    <span className="text-slate-400">·</span>
                    <span className="text-[10px] font-mono text-slate-600">
                      Principal Facility {fmtAuditNum(parsedSchedule.metadata.investmentMade)}
                    </span>
                  </>
                )}
              </div>
              <h3 className="text-sm font-bold text-slate-900 tracking-tight mt-0.5">
                Debt Servicing &amp; Repayment Capacity Audit (DSCR)
              </h3>
              <p className="text-[11px] text-slate-600 mt-0.5">
                Annual verification that Operating Cash Flow and EBITDA sustain Scheduled Installment Principal + Profit obligations.
              </p>
            </div>
            <span className={`text-xs font-mono font-bold px-2.5 py-1 rounded border ${
              isScheduleImplemented
                ? 'text-emerald-800 bg-emerald-50 border-emerald-300'
                : 'text-sky-800 bg-sky-50 border-sky-300'
            }`}>
              {isScheduleImplemented ? 'Schedule Locked & Implemented' : 'Model Debt Servicing'}
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs font-mono">
              <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
                <tr>
                  <th className="py-2.5 px-4 font-sans font-semibold text-slate-900">Year</th>
                  <th className="py-2.5 px-3 text-right font-semibold">EBITDA</th>
                  <th className="py-2.5 px-3 text-right font-semibold">Pre-Fin OCF</th>
                  <th className="py-2.5 px-3 text-right text-amber-800 font-semibold">Principal</th>
                  <th className="py-2.5 px-3 text-right text-sky-800 font-semibold">Finance Cost</th>
                  <th className="py-2.5 px-3 text-right text-slate-900 font-semibold">Total Debt Service</th>
                  <th className="py-2.5 px-3 text-right font-semibold">EBITDA DSCR</th>
                  <th className="py-2.5 px-3 text-right font-semibold">Cash DSCR</th>
                  <th className="py-2.5 px-3 text-right font-semibold">Closing Cash</th>
                  <th className="py-2.5 px-4 text-right font-semibold">Capacity Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/70 tabular-nums">
                {debtServiceCoverage.map(row => (
                  <tr key={row.year} className="hover:bg-slate-50">
                    <td className="py-2.5 px-4 font-bold text-slate-900">
                      {row.year}{' '}
                      <span className="text-[9px] font-normal text-slate-500">
                        {row.isProjected ? '(P)' : '(H)'}{row.isScheduleCovered ? ' · Covered' : ''}
                      </span>
                    </td>
                    <td className="py-2.5 px-3 text-right text-slate-800">{fmtAuditNum(row.ebitda)}</td>
                    <td className="py-2.5 px-3 text-right text-slate-800">{fmtAuditNum(row.operatingCashFlowBeforeFinance)}</td>
                    <td className="py-2.5 px-3 text-right text-amber-800 font-semibold">{fmtAuditNum(row.scheduledPrincipal)}</td>
                    <td className="py-2.5 px-3 text-right text-sky-800 font-semibold">{fmtAuditNum(row.scheduledFinanceCost)}</td>
                    <td className="py-2.5 px-3 text-right text-slate-900 font-bold">{fmtAuditNum(row.totalDebtService)}</td>
                    <td className={`py-2.5 px-3 text-right font-bold ${
                      row.ebitdaDscr >= 1.25 ? 'text-emerald-700' : row.ebitdaDscr >= 1.0 ? 'text-amber-800' : 'text-rose-700'
                    }`}>
                      {row.totalDebtService > 0 ? `${row.ebitdaDscr.toFixed(2)}x` : '-'}
                    </td>
                    <td className={`py-2.5 px-3 text-right font-bold ${
                      row.cashDscr >= 1.15 ? 'text-emerald-700' : row.cashDscr >= 1.0 ? 'text-amber-800' : 'text-rose-700'
                    }`}>
                      {row.totalDebtService > 0 ? `${row.cashDscr.toFixed(2)}x` : '-'}
                    </td>
                    <td className="py-2.5 px-3 text-right text-emerald-700 font-medium">{fmtAuditNum(row.closingCash)}</td>
                    <td className="py-2.5 px-4 text-right font-sans">
                      <span className={`inline-block text-[10px] font-bold px-2 py-0.5 rounded border ${
                        row.capacityStatus === 'STRONG'
                          ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                          : row.capacityStatus === 'ADEQUATE'
                          ? 'bg-sky-50 text-sky-800 border-sky-300'
                          : row.capacityStatus === 'TIGHT'
                          ? 'bg-amber-50 text-amber-900 border-amber-300'
                          : 'bg-rose-50 text-rose-800 border-rose-300'
                      }`}>
                        {row.capacityStatus}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
