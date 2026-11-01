import React from 'react';
import { Scale, CheckCircle2, AlertTriangle, Info, ShieldCheck, Check, ArrowRight } from 'lucide-react';
import { CashFlowStatementYear, BalanceSheetYear, CompanyProfile } from '../types/financial';

interface CashFlowReconciliationAuditProps {
  cashFlowYears: CashFlowStatementYear[];
  balanceSheetYears: BalanceSheetYear[];
  companyProfile: CompanyProfile;
}

export const CashFlowReconciliationAudit: React.FC<CashFlowReconciliationAuditProps> = ({
  cashFlowYears,
  balanceSheetYears,
  companyProfile,
}) => {
  const reconciledCount = cashFlowYears.filter(cf => cf.reconciliationDifference === 0).length;
  const isAllReconciled = reconciledCount === cashFlowYears.length;

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

          <div className="flex items-center gap-2 bg-slate-950/80 border border-slate-800 px-3.5 py-2 rounded-xl shrink-0">
            <div className={`w-2.5 h-2.5 rounded-full ${isAllReconciled ? 'bg-emerald-400 animate-pulse' : 'bg-amber-400'}`} />
            <span className="text-xs font-mono font-bold text-white">
              {reconciledCount} / {cashFlowYears.length} Periods Reconciled
            </span>
          </div>
        </div>
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
                    {cf.netCashFromOperating.toLocaleString()}
                  </td>
                ))}
              </tr>

              <tr className="hover:bg-slate-800/20 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-200 font-medium">B. Investing Cash Flow (Calculated)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-medium ${cf.netCashFromInvesting >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                    {cf.netCashFromInvesting.toLocaleString()}
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
                    {cf.directlyEnteredFinancingCashFlow.toLocaleString()}
                  </td>
                ))}
              </tr>

              <tr className="border-t border-slate-700 font-semibold bg-slate-850/60">
                <td className="py-2.5 px-4 font-sans text-white">Net Change in Cash (A + B + C)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className={`py-2.5 px-3 text-right font-bold ${cf.netChangeInCash >= 0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                    {cf.netChangeInCash.toLocaleString()}
                  </td>
                ))}
              </tr>

              <tr className="hover:bg-slate-800/20 transition-colors">
                <td className="py-2.5 px-4 font-sans text-slate-400">Opening Cash & Cash Equivalents</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-slate-300">
                    {cf.openingCash.toLocaleString()}
                  </td>
                ))}
              </tr>

              {/* Closing Cash Derived with Statutory Double Underline */}
              <tr className="font-bold text-emerald-400 bg-emerald-950/20 border-b-4 border-double border-emerald-500/80 border-t border-emerald-800/60">
                <td className="py-3 px-4 font-sans text-sm">Closing Cash Derived (Opening + Net Change)</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-3 px-3 text-right text-emerald-400 font-bold text-sm">
                    {cf.closingCashCalculated.toLocaleString()}
                  </td>
                ))}
              </tr>

              <tr className="font-semibold text-slate-200 hover:bg-slate-800/20 transition-colors">
                <td className="py-2.5 px-4 font-sans text-sky-400">Stated Cash in Statement of Financial Position</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-sky-400 font-semibold">
                    {cf.balanceSheetCash.toLocaleString()}
                  </td>
                ))}
              </tr>

              <tr className="font-bold border-t border-slate-700 bg-slate-850/40">
                <td className="py-3 px-4 font-sans text-white flex items-center gap-1.5">
                  <span>Reconciliation Difference</span>
                  <span className="text-[10px] text-slate-400 font-normal font-sans">(Derived - Stated)</span>
                </td>
                {cashFlowYears.map(cf => {
                  const isZero = cf.reconciliationDifference === 0;
                  return (
                    <td key={cf.year} className={`py-3 px-3 text-right ${isZero ? 'text-emerald-400 font-bold' : 'text-amber-400 font-bold'}`}>
                      {cf.reconciliationDifference > 0 ? `+${cf.reconciliationDifference.toLocaleString()}` : cf.reconciliationDifference.toLocaleString()}
                    </td>
                  );
                })}
              </tr>

              <tr className="bg-slate-900/90 font-sans">
                <td className="py-2.5 px-4 text-slate-400 text-[11px] font-medium">Audit Verdict</td>
                {cashFlowYears.map(cf => {
                  const isZero = cf.reconciliationDifference === 0;
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
    </div>
  );
};
