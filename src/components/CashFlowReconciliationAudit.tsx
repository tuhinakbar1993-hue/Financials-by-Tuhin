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
      <div className="p-5 sm:p-6 bg-gradient-to-br from-slate-900/95 via-slate-900 to-slate-950 border border-slate-800/90 rounded-2xl shadow-xl shadow-black/20">
        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-teal-950/80 border border-teal-800/60 flex items-center justify-center text-teal-400">
                <Scale className="w-4 h-4" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h2 className="text-lg font-bold text-white tracking-tight">
                    Cash Flow & Balance Sheet Reconciliation Matrix
                  </h2>
                  <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-teal-300 bg-teal-950/90 border border-teal-800/70 px-2 py-0.5 rounded shadow-sm">
                    Section 76.3 & 76.4
                  </span>
                </div>
                <p className="text-xs text-slate-400 mt-0.5">
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
                className="px-3 py-2 bg-emerald-600/90 hover:bg-emerald-500 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-colors shadow-sm"
              >
                <CalendarCheck2 className="w-3.5 h-3.5" />
                <span>{isScheduleImplemented ? 'Repayment Schedule Active' : 'Repayment Schedule'}</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            )}
            <div className="flex items-center gap-2 bg-slate-950/80 border border-slate-800 px-3.5 py-2 rounded-xl">
              <div className={`w-2.5 h-2.5 rounded-full ${isAllReconciled ? 'bg-emerald-400 animate-pulse' : 'bg-amber-400'}`} />
              <span className="text-xs font-mono font-bold text-white">
                {reconciledCount} / {cashFlowYears.length} Periods Reconciled · BS {balancedBSCount}/{balanceSheetYears.length}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Model-Wide Validation & Fixed-Driver Integrity Banner */}
      <div className={`p-4 rounded-2xl border text-xs flex flex-col gap-2 ${
        validationIssues.length === 0
          ? 'bg-emerald-950/20 border-emerald-800/60 text-emerald-200'
          : 'bg-amber-950/20 border-amber-800/60 text-amber-200'
      }`}>
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 font-bold">
            {validationIssues.length === 0 ? (
              <ShieldCheck className="w-4 h-4 text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />
            )}
            <span>
              {validationIssues.length === 0
                ? 'Model Validation Passed: All 6 Fixed Primary Inputs Strictly Preserved · Zero Accounting Variance Across All Selected Years'
                : `${validationIssues.length} Validation Check(s) Identified`}
            </span>
          </div>
          <span className="text-[10px] font-mono uppercase px-2 py-0.5 rounded bg-slate-900 border border-slate-700">
            {cashFlowYears.length} Years Audited
          </span>
        </div>
        {validationIssues.length > 0 && (
          <ul className="space-y-1 pl-6 list-disc text-[11px] text-amber-300">
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
        <div className="p-4 bg-slate-900/90 border border-slate-800/80 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-200 font-bold flex items-center gap-2">
            <div className="w-5 h-5 rounded bg-sky-950/80 border border-sky-800/60 flex items-center justify-center text-sky-400 text-[10px] font-mono font-bold">
              1
            </div>
            <span>Operating Cash Flow</span>
          </div>
          <p className="text-slate-400 leading-relaxed text-[11px]">
            Derived via the <strong>Indirect Method</strong> from Net Profit Before Tax, adjusting for non-cash depreciation, finance cost, and working capital shifts (Inventory, Receivables, Payables).
          </p>
        </div>

        <div className="p-4 bg-slate-900/90 border border-slate-800/80 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-200 font-bold flex items-center gap-2">
            <div className="w-5 h-5 rounded bg-indigo-950/80 border border-indigo-800/60 flex items-center justify-center text-indigo-400 text-[10px] font-mono font-bold">
              2
            </div>
            <span>Investing Cash Flow</span>
          </div>
          <p className="text-slate-400 leading-relaxed text-[11px]">
            Tied directly to the <strong>Schedule of Fixed Assets</strong> capital expenditure additions, asset disposals, and financial investment allocations.
          </p>
        </div>

        <div className="p-4 bg-slate-900/90 border border-slate-800/80 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-200 font-bold flex items-center gap-2">
            <div className="w-5 h-5 rounded bg-amber-950/80 border border-amber-800/60 flex items-center justify-center text-amber-400 text-[10px] font-mono font-bold">
              3
            </div>
            <span>Financing Cash Flow</span>
          </div>
          <p className="text-slate-400 leading-relaxed text-[11px]">
            Preserved as a <strong>mandatory direct user entry</strong> separately for every historical and projected financial year. Never derived or forced.
          </p>
        </div>
      </div>

      {/* Reconciliation Audit Table */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
        <div className="p-4 border-b border-slate-800 bg-slate-850/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-mono uppercase tracking-widest text-emerald-400 font-bold">
                {companyProfile.companyName}
              </span>
              <span className="text-slate-600">·</span>
              <span className="text-[10px] font-mono text-slate-400">
                Audited Schedule
              </span>
            </div>
            <h3 className="text-sm font-bold text-white tracking-tight mt-0.5">
              Multi-Year Cash Flow Audit Reconciliation Schedule
            </h3>
            <p className="text-[11px] text-slate-400 mt-0.5">
              All monetary figures in {companyProfile.reportingCurrency} ({companyProfile.reportingUnit}) · Complete {cashFlowYears.length}-Year Horizon
            </p>
          </div>
          <span className="text-xs font-mono font-bold text-emerald-400 bg-emerald-950/80 border border-emerald-800/60 px-2.5 py-1 rounded shadow-sm self-start sm:self-auto">
            {cashFlowYears.length} Financial Years Audited
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs font-mono">
            <thead className="bg-slate-800/90 text-slate-300 border-b border-slate-700 text-[11px]">
              <tr>
                <th className="py-3 px-4 font-sans font-semibold text-white tracking-wide">
                  Audit Parameter / Cash Flow Classification
                </th>
                {cashFlowYears.map(cf => (
                  <th key={cf.year} className={`py-2.5 px-3 text-right ${cf.isProjected ? 'bg-sky-950/20 text-sky-300' : 'text-slate-200'}`}>
                    <div className="flex flex-col items-end">
                      <span className="font-bold text-xs">{cf.year}</span>
                      <span className={`text-[9px] font-sans uppercase tracking-wider font-semibold ${
                        cf.isProjected ? 'text-sky-400' : 'text-slate-400'
                      }`}>
                        {cf.isProjected ? 'Forecast' : 'Audited'}
                      </span>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/60 tabular-nums">
              <tr className="hover:bg-slate-800/20 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-200 font-medium">A. Operating Cash Flow (Calculated)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-medium ${cf.netCashFromOperating >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                    {fmtAuditNum(cf.netCashFromOperating)}
                  </td>
                ))}
              </tr>

              <tr className="hover:bg-slate-800/20 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-200 font-medium">B. Investing Cash Flow (Calculated)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-medium ${cf.netCashFromInvesting >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                    {fmtAuditNum(cf.netCashFromInvesting)}
                  </td>
                ))}
              </tr>

              <tr className="bg-amber-950/15 font-semibold hover:bg-amber-950/25 transition-colors">
                <td className="py-2.5 px-4 font-sans text-amber-300 flex items-center gap-2">
                  <span>C. Financing Cash Flow (Direct User Entry)</span>
                  <span className="text-[9px] font-mono uppercase bg-amber-950 text-amber-300 border border-amber-800/70 px-1.5 py-0.2 rounded">
                    Mandatory
                  </span>
                </td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-amber-300 font-bold">
                    {fmtAuditNum(cf.directlyEnteredFinancingCashFlow)}
                  </td>
                ))}
              </tr>

              <tr className="border-t border-slate-700 font-semibold bg-slate-850/60">
                <td className="py-2.5 px-4 font-sans text-white">Net Change in Cash (A + B + C)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-bold ${cf.netChangeInCash >= 0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                    {fmtAuditNum(cf.netChangeInCash)}
                  </td>
                ))}
              </tr>

              <tr className="hover:bg-slate-800/20 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-400">Opening Cash & Cash Equivalents</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-slate-300">
                    {fmtAuditNum(cf.openingCash)}
                  </td>
                ))}
              </tr>

              {/* Closing Cash Derived with Statutory Double Underline */}
              <tr className="font-bold text-emerald-400 bg-emerald-950/20 border-b-4 border-double border-emerald-500/80 border-t border-emerald-800/60">
                <td className="py-3 px-4 font-sans text-sm">Closing Cash Derived (Opening + Net Change)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-3 px-3 text-right text-emerald-400 font-bold text-sm">
                    {fmtAuditNum(cf.closingCashCalculated)}
                  </td>
                ))}
              </tr>

              <tr className="font-semibold text-slate-200 hover:bg-slate-800/20 transition-colors">
                <td className="py-2.5 px-4 font-sans text-sky-400">Stated Cash in Statement of Financial Position</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-sky-400 font-semibold">
                    {fmtAuditNum(cf.balanceSheetCash)}
                  </td>
                ))}
              </tr>

              <tr className="font-bold border-t border-slate-700 bg-slate-850/40">
                <td className="py-3 px-4 font-sans text-white flex items-center gap-1.5">
                  <span>Reconciliation Difference</span>
                  <span className="text-[10px] text-slate-400 font-normal font-sans">(Derived - Stated)</span>
                </td>
                {cashFlowYears.map(cf => {
                  const isZero = Math.abs(cf.reconciliationDifference) < 0.005;
                  return (
                    <td key={cf.year} className={`py-3 px-3 text-right ${isZero ? 'text-emerald-400 font-bold' : 'text-amber-400 font-bold'}`}>
                      {cf.reconciliationDifference > 0 ? `+${fmtAuditNum(cf.reconciliationDifference)}` : fmtAuditNum(cf.reconciliationDifference)}
                    </td>
                  );
                })}
              </tr>

              <tr className="bg-slate-900/90 font-sans">
                <td className="py-2.5 px-4 text-slate-400 text-[11px] font-medium">Audit Verdict</td>
                {cashFlowYears.map(cf => {
                  const isZero = Math.abs(cf.reconciliationDifference) < 0.005;
                  return (
                    <td key={cf.year} className="py-2.5 px-3 text-right">
                      {isZero ? (
                        <span className="inline-flex items-center gap-1 text-[10px] font-bold text-emerald-400 bg-emerald-950/60 border border-emerald-800/50 px-2 py-0.5 rounded">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" /> Reconciled
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 text-[10px] font-bold text-amber-400 bg-amber-950/60 border border-amber-800/50 px-2 py-0.5 rounded">
                          <AlertTriangle className="w-3.5 h-3.5 text-amber-400" /> Variance
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
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-4 border-b border-slate-800 bg-slate-850/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-mono uppercase tracking-widest text-emerald-400 font-bold">
                  {parsedSchedule?.metadata?.sourceFileName || companyProfile.companyName}
                </span>
                {parsedSchedule?.metadata?.investmentMade && (
                  <>
                    <span className="text-slate-600">·</span>
                    <span className="text-[10px] font-mono text-slate-400">
                      Principal Facility {fmtAuditNum(parsedSchedule.metadata.investmentMade)}
                    </span>
                  </>
                )}
              </div>
              <h3 className="text-sm font-bold text-white tracking-tight mt-0.5">
                Debt Servicing &amp; Repayment Capacity Audit (DSCR)
              </h3>
              <p className="text-[11px] text-slate-400 mt-0.5">
                Annual verification that Operating Cash Flow and EBITDA sustain Scheduled Installment Principal + Profit obligations.
              </p>
            </div>
            <span className={`text-xs font-mono font-bold px-2.5 py-1 rounded border ${
              isScheduleImplemented
                ? 'text-emerald-300 bg-emerald-950/80 border-emerald-700/70'
                : 'text-sky-300 bg-sky-950/80 border-sky-700/70'
            }`}>
              {isScheduleImplemented ? 'Schedule Locked & Implemented' : 'Model Debt Servicing'}
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs font-mono">
              <thead className="bg-slate-800/90 text-slate-300 border-b border-slate-700 text-[11px]">
                <tr>
                  <th className="py-2.5 px-4 font-sans font-semibold text-white">Year</th>
                  <th className="py-2.5 px-3 text-right">EBITDA</th>
                  <th className="py-2.5 px-3 text-right">Pre-Fin OCF</th>
                  <th className="py-2.5 px-3 text-right text-amber-300">Principal</th>
                  <th className="py-2.5 px-3 text-right text-sky-300">Finance Cost</th>
                  <th className="py-2.5 px-3 text-right text-white">Total Debt Service</th>
                  <th className="py-2.5 px-3 text-right">EBITDA DSCR</th>
                  <th className="py-2.5 px-3 text-right">Cash DSCR</th>
                  <th className="py-2.5 px-3 text-right">Closing Cash</th>
                  <th className="py-2.5 px-4 text-right">Capacity Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60 tabular-nums">
                {debtServiceCoverage.map(row => (
                  <tr key={row.year} className="hover:bg-slate-800/20">
                    <td className="py-2.5 px-4 font-bold text-white">
                      {row.year}{' '}
                      <span className="text-[9px] font-normal text-slate-400">
                        {row.isProjected ? '(P)' : '(H)'}{row.isScheduleCovered ? ' · Covered' : ''}
                      </span>
                    </td>
                    <td className="py-2.5 px-3 text-right text-slate-200">{fmtAuditNum(row.ebitda)}</td>
                    <td className="py-2.5 px-3 text-right text-slate-200">{fmtAuditNum(row.operatingCashFlowBeforeFinance)}</td>
                    <td className="py-2.5 px-3 text-right text-amber-300 font-semibold">{fmtAuditNum(row.scheduledPrincipal)}</td>
                    <td className="py-2.5 px-3 text-right text-sky-300 font-semibold">{fmtAuditNum(row.scheduledFinanceCost)}</td>
                    <td className="py-2.5 px-3 text-right text-white font-bold">{fmtAuditNum(row.totalDebtService)}</td>
                    <td className={`py-2.5 px-3 text-right font-bold ${
                      row.ebitdaDscr >= 1.25 ? 'text-emerald-400' : row.ebitdaDscr >= 1.0 ? 'text-amber-300' : 'text-rose-400'
                    }`}>
                      {row.totalDebtService > 0 ? `${row.ebitdaDscr.toFixed(2)}x` : '-'}
                    </td>
                    <td className={`py-2.5 px-3 text-right font-bold ${
                      row.cashDscr >= 1.15 ? 'text-emerald-400' : row.cashDscr >= 1.0 ? 'text-amber-300' : 'text-rose-400'
                    }`}>
                      {row.totalDebtService > 0 ? `${row.cashDscr.toFixed(2)}x` : '-'}
                    </td>
                    <td className="py-2.5 px-3 text-right text-emerald-300">{fmtAuditNum(row.closingCash)}</td>
                    <td className="py-2.5 px-4 text-right font-sans">
                      <span className={`inline-block text-[10px] font-bold px-2 py-0.5 rounded border ${
                        row.capacityStatus === 'STRONG'
                          ? 'bg-emerald-950/70 text-emerald-300 border-emerald-700/60'
                          : row.capacityStatus === 'ADEQUATE'
                          ? 'bg-sky-950/70 text-sky-300 border-sky-700/60'
                          : row.capacityStatus === 'TIGHT'
                          ? 'bg-amber-950/70 text-amber-300 border-amber-700/60'
                          : 'bg-rose-950/70 text-rose-300 border-rose-700/60'
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
