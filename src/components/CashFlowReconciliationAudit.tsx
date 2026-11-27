import React, { useState, useEffect, useMemo } from 'react';
import {
  Scale,
  CheckCircle2,
  AlertTriangle,
  ShieldCheck,
  CalendarCheck2,
  ArrowUpRight,
  Wrench,
  RotateCcw,
  FileSearch,
  ClipboardCheck,
  AlertCircle,
  Building2,
} from 'lucide-react';
import {
  CashFlowStatementYear,
  BalanceSheetYear,
  CompanyProfile,
  FinancialValidationIssue,
  DebtServiceCoverageYear,
  ParsedRepaymentSchedule,
  CharteredAuditReport,
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
  auditReport?: CharteredAuditReport;
  lastAuditRunAt?: string | null;
  lastRectificationAppliedCount?: number | null;
  hasRectificationSnapshot?: boolean;
  onRunAudit?: () => void;
  onInstantRectification?: (findingIds?: string[]) => void;
  onUndoLastRectification?: () => void;
  onOpenEntitySetup?: () => void;
}

type AuditViewFilter = 'all' | 'auto' | 'manual' | 'passed';

export const CashFlowReconciliationAudit: React.FC<CashFlowReconciliationAuditProps> = ({
  cashFlowYears,
  balanceSheetYears,
  companyProfile,
  validationIssues = [],
  debtServiceCoverage = [],
  isScheduleImplemented = false,
  parsedSchedule = null,
  onJumpToRepaymentSchedule,
  auditReport,
  lastAuditRunAt,
  lastRectificationAppliedCount,
  hasRectificationSnapshot = false,
  onRunAudit,
  onInstantRectification,
  onUndoLastRectification,
  onOpenEntitySetup,
}) => {
  const reconciledCount = cashFlowYears.filter(cf => cf.reconciliationDifference === 0).length;
  const balancedBSCount = balanceSheetYears.filter(bs => Math.abs(bs.balanceCheck) <= 1).length;
  const scheduleCoverageYears = debtServiceCoverage.filter(d => d.totalDebtService > 0);
  const histDscr = debtServiceCoverage.filter(d => !d.isProjected && d.totalDebtService > 0);
  const projDscr = debtServiceCoverage.filter(d => d.isProjected && d.totalDebtService > 0);
  const minHistEbitdaDscr = histDscr.length > 0 ? Math.min(...histDscr.map(d => d.ebitdaDscr)) : null;
  const minProjEbitdaDscr = projDscr.length > 0 ? Math.min(...projDscr.map(d => d.ebitdaDscr)) : null;

  const [auditFilter, setAuditFilter] = useState<AuditViewFilter>('all');
  const autoRectifiableFindings = useMemo(
    () => (auditReport?.findings || []).filter(f => f.isAutoRectifiable),
    [auditReport]
  );
  const manualFindings = useMemo(
    () => (auditReport?.findings || []).filter(f => f.requiresManualVerification),
    [auditReport]
  );

  const [selectedFindingIds, setSelectedFindingIds] = useState<Record<string, boolean>>({});

  useEffect(() => {
    const nextMap: Record<string, boolean> = {};
    autoRectifiableFindings.forEach(f => {
      nextMap[f.id] = selectedFindingIds[f.id] !== undefined ? selectedFindingIds[f.id] : true;
    });
    setSelectedFindingIds(nextMap);
  }, [autoRectifiableFindings]);

  const selectedIdsList = useMemo(
    () => autoRectifiableFindings.filter(f => selectedFindingIds[f.id] ?? true).map(f => f.id),
    [autoRectifiableFindings, selectedFindingIds]
  );

  const toggleFindingSelection = (id: string) => {
    setSelectedFindingIds(prev => ({
      ...prev,
      [id]: !(prev[id] ?? true),
    }));
  };

  const opinionLabel =
    auditReport?.overallOpinion === 'UNQUALIFIED_CLEAN'
      ? 'Unqualified Clean Audit Opinion'
      : auditReport?.overallOpinion === 'ADVERSE_MATHEMATICAL_ERROR'
      ? 'Critical Discrepancies Identified'
      : 'Qualified — Review / Action Advisory';

  const opinionBadgeClass =
    auditReport?.overallOpinion === 'UNQUALIFIED_CLEAN'
      ? 'bg-emerald-50 border-emerald-300 text-emerald-900'
      : auditReport?.overallOpinion === 'ADVERSE_MATHEMATICAL_ERROR'
      ? 'bg-rose-50 border-rose-300 text-rose-900'
      : 'bg-amber-50 border-amber-300 text-amber-900';

  const visibleFindings = useMemo(() => {
    if (!auditReport) return [];
    if (auditFilter === 'auto') return autoRectifiableFindings;
    if (auditFilter === 'manual') return manualFindings;
    return auditReport.findings;
  }, [auditReport, auditFilter, autoRectifiableFindings, manualFindings]);

  return (
    <div className="space-y-6">
      {/* Executive Chartered Accountancy Audit & Instant Rectification Console */}
      <div className="p-5 sm:p-6 bg-white/94 backdrop-blur-md border border-slate-200/90 rounded-2xl shadow-sm space-y-5">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-start gap-3">
            <div className="w-10 h-10 rounded-xl bg-teal-50 border border-teal-200 flex items-center justify-center text-teal-700 shrink-0">
              <Scale className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900 tracking-tight">
                Chartered Accountancy Audit &amp; Instant Rectification Suite
              </h2>
              <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600 mt-0.5">
                <span>{companyProfile.companyName?.trim() || 'Active Case Entity'}</span>
                <span aria-hidden="true">·</span>
                <span>{companyProfile.accountingFramework} Engagement Audit</span>
                <span aria-hidden="true">·</span>
                <span>Periods Reconciled {reconciledCount}/{cashFlowYears.length}</span>
                <span aria-hidden="true">·</span>
                <span>Balance Sheet {balancedBSCount}/{balanceSheetYears.length}</span>
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            {onRunAudit && (
              <button
                type="button"
                onClick={onRunAudit}
                className="px-4 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold inline-flex items-center gap-2 transition-colors shadow-sm whitespace-nowrap shrink-0"
              >
                <FileSearch className="w-3.5 h-3.5" />
                <span>Audit</span>
              </button>
            )}

            {onInstantRectification && (
              <button
                type="button"
                disabled={selectedIdsList.length === 0}
                onClick={() => onInstantRectification(selectedIdsList)}
                className={`px-4 py-2 rounded-xl text-xs font-bold inline-flex items-center gap-2 transition-colors shadow-sm whitespace-nowrap shrink-0 ${
                  selectedIdsList.length > 0
                    ? 'bg-emerald-600 hover:bg-emerald-700 text-white'
                    : 'bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed'
                }`}
              >
                <Wrench className="w-3.5 h-3.5" />
                <span>
                  Instant Rectification
                  {selectedIdsList.length > 0 ? ` (${selectedIdsList.length})` : ''}
                </span>
              </button>
            )}

            {hasRectificationSnapshot && onUndoLastRectification && (
              <button
                type="button"
                onClick={onUndoLastRectification}
                className="px-3.5 py-2 rounded-xl bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-semibold inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span>Undo Rectification</span>
              </button>
            )}

            {onJumpToRepaymentSchedule && (
              <button
                type="button"
                onClick={onJumpToRepaymentSchedule}
                className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-200 rounded-xl text-xs font-semibold inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0"
              >
                <CalendarCheck2 className="w-3.5 h-3.5 text-emerald-700" />
                <span>{isScheduleImplemented ? 'Repayment Schedule Active' : 'Repayment Schedule'}</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>

        {/* Audit Opinion & Summary KPI Strip */}
        {auditReport && (
          <>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-3 text-xs pt-2 border-t border-slate-200/80">
              <div className={`p-3.5 rounded-xl border ${opinionBadgeClass}`}>
                <span className="text-[11px] font-medium block opacity-80">Audit Opinion</span>
                <span className="font-bold text-sm mt-0.5 block">{opinionLabel}</span>
                {lastAuditRunAt && (
                  <span className="text-[10px] opacity-75 block mt-1">
                    Verified {new Date(lastAuditRunAt).toLocaleTimeString()}
                  </span>
                )}
              </div>

              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70">
                <span className="text-[11px] text-slate-600 font-medium block">Auto-Rectifiable Items</span>
                <span className="font-mono font-bold text-lg text-slate-900 tabular-nums mt-0.5 block">
                  {auditReport.autoRectifiableCount}
                </span>
                <span className="text-[10px] text-slate-500">Selected: {selectedIdsList.length} ready</span>
              </div>

              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70">
                <span className="text-[11px] text-slate-600 font-medium block">Manual Review Required</span>
                <span className="font-mono font-bold text-lg text-slate-900 tabular-nums mt-0.5 block">
                  {auditReport.manualVerificationCount}
                </span>
                <span className="text-[10px] text-slate-500">Preserves user judgment</span>
              </div>

              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70">
                <span className="text-[11px] text-slate-600 font-medium block">Assertions Verified Clean</span>
                <span className="font-mono font-bold text-lg text-emerald-700 tabular-nums mt-0.5 block">
                  {auditReport.passedChecks.length}
                </span>
                <span className="text-[10px] text-slate-500">{auditReport.auditedYearsCount} Years Audited</span>
              </div>
            </div>

            {lastRectificationAppliedCount != null && (
              <div className="p-3.5 rounded-xl bg-emerald-50 border border-emerald-200 text-xs text-emerald-900 flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                  <span>
                    {lastRectificationAppliedCount > 0
                      ? `Instant Rectification completed: ${lastRectificationAppliedCount} issue(s) resolved and all five statements, notes, ratios, and exports re-audited.`
                      : 'Re-audit complete: no additional auto-rectifiable discrepancies remained.'}
                  </span>
                </div>
                {hasRectificationSnapshot && onUndoLastRectification && (
                  <button
                    type="button"
                    onClick={onUndoLastRectification}
                    className="text-emerald-800 hover:text-emerald-950 font-semibold underline"
                  >
                    Revert Last Rectification
                  </button>
                )}
              </div>
            )}

            {/* Filter Bar */}
            <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
              <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-xl border border-slate-200 text-xs">
                <button
                  type="button"
                  onClick={() => setAuditFilter('all')}
                  className={`px-3 py-1.5 rounded-lg font-medium transition-colors whitespace-nowrap shrink-0 ${
                    auditFilter === 'all' ? 'bg-white text-slate-900 font-bold shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  All Findings ({auditReport.findings.length})
                </button>
                <button
                  type="button"
                  onClick={() => setAuditFilter('auto')}
                  className={`px-3 py-1.5 rounded-lg font-medium transition-colors whitespace-nowrap shrink-0 ${
                    auditFilter === 'auto' ? 'bg-white text-slate-900 font-bold shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Auto-Rectifiable ({autoRectifiableFindings.length})
                </button>
                <button
                  type="button"
                  onClick={() => setAuditFilter('manual')}
                  className={`px-3 py-1.5 rounded-lg font-medium transition-colors whitespace-nowrap shrink-0 ${
                    auditFilter === 'manual' ? 'bg-white text-slate-900 font-bold shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Manual Verification ({manualFindings.length})
                </button>
                <button
                  type="button"
                  onClick={() => setAuditFilter('passed')}
                  className={`px-3 py-1.5 rounded-lg font-medium transition-colors whitespace-nowrap shrink-0 ${
                    auditFilter === 'passed' ? 'bg-white text-slate-900 font-bold shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Verified Passed ({auditReport.passedChecks.length})
                </button>
              </div>

              {autoRectifiableFindings.length > 0 && auditFilter !== 'passed' && (
                <div className="flex items-center gap-3 text-[11px] text-slate-600">
                  <button
                    type="button"
                    onClick={() => {
                      const next: Record<string, boolean> = {};
                      autoRectifiableFindings.forEach(f => { next[f.id] = true; });
                      setSelectedFindingIds(next);
                    }}
                    className="hover:text-emerald-700 font-semibold underline"
                  >
                    Select All Auto-Fixes
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const next: Record<string, boolean> = {};
                      autoRectifiableFindings.forEach(f => { next[f.id] = false; });
                      setSelectedFindingIds(next);
                    }}
                    className="hover:text-slate-900 underline"
                  >
                    Clear Selection
                  </button>
                </div>
              )}
            </div>

            {/* Findings List */}
            {auditFilter !== 'passed' ? (
              visibleFindings.length === 0 ? (
                <div className="p-6 rounded-xl bg-emerald-50/70 border border-emerald-200 text-center space-y-1.5">
                  <ShieldCheck className="w-7 h-7 text-emerald-700 mx-auto" />
                  <h3 className="text-sm font-bold text-emerald-950">
                    No Discrepancies Detected in Current View
                  </h3>
                  <p className="text-xs text-emerald-800 max-w-xl mx-auto">
                    All double-entry balances, indirect cash flows, fixed asset schedules, equity roll-forwards, debt servicing rates, and natural YoY proportions pass Chartered Accountancy verification.
                  </p>
                </div>
              ) : (
                <div className="divide-y divide-slate-200 border border-slate-200 rounded-xl overflow-hidden bg-white">
                  {visibleFindings.map(finding => {
                    const isChecked = selectedFindingIds[finding.id] ?? true;
                    const sevColor =
                      finding.severity === 'critical'
                        ? 'text-rose-800 bg-rose-50 border-rose-200'
                        : finding.severity === 'warning'
                        ? 'text-amber-900 bg-amber-50 border-amber-200'
                        : finding.severity === 'advisory'
                        ? 'text-sky-900 bg-sky-50 border-sky-200'
                        : 'text-slate-700 bg-slate-100 border-slate-200';
                    const sevText =
                      finding.severity === 'critical'
                        ? 'Critical'
                        : finding.severity === 'warning'
                        ? 'Warning'
                        : finding.severity === 'advisory'
                        ? 'Advisory'
                        : 'Info';

                    return (
                      <div key={finding.id} className="p-4 hover:bg-slate-50/60 transition-colors flex flex-col lg:flex-row lg:items-start justify-between gap-4 text-xs">
                        <div className="flex items-start gap-3 flex-1">
                          {finding.isAutoRectifiable ? (
                            <input
                              type="checkbox"
                              checked={isChecked}
                              onChange={() => toggleFindingSelection(finding.id)}
                              aria-label={`Select ${finding.title} for instant rectification`}
                              className="mt-1 w-4 h-4 accent-emerald-600 rounded shrink-0 cursor-pointer"
                            />
                          ) : (
                            <AlertCircle className="w-4 h-4 text-amber-600 mt-0.5 shrink-0" />
                          )}

                          <div className="space-y-1 flex-1">
                            <div className="flex flex-wrap items-center gap-2">
                              <span className={`px-2 py-0.5 rounded border text-[10px] font-bold uppercase ${sevColor}`}>
                                {sevText}
                              </span>
                              <span className="font-bold text-slate-900 text-sm">{finding.title}</span>
                            </div>

                            <div className="flex flex-wrap items-center gap-2 text-[11px] text-slate-500">
                              <span>{finding.category}</span>
                              <span aria-hidden="true">·</span>
                              <span>{finding.standardRef}</span>
                              {finding.currentValue && (
                                <>
                                  <span aria-hidden="true">·</span>
                                  <span className="font-mono text-slate-700">Current: {finding.currentValue}</span>
                                </>
                              )}
                              {finding.expectedValue && (
                                <>
                                  <span aria-hidden="true">·</span>
                                  <span className="font-mono text-emerald-700 font-semibold">Expected: {finding.expectedValue}</span>
                                </>
                              )}
                            </div>

                            <p className="text-slate-700 leading-relaxed pt-0.5">{finding.description}</p>

                            {finding.rectificationPreview && finding.isAutoRectifiable && (
                              <p className="text-[11px] text-emerald-800 font-medium pt-0.5">
                                Proposed Rectification: {finding.rectificationPreview}
                              </p>
                            )}
                            {finding.manualGuidance && finding.requiresManualVerification && (
                              <p className="text-[11px] text-amber-900 font-medium pt-0.5">
                                Manual Verification Required: {finding.manualGuidance}
                              </p>
                            )}
                          </div>
                        </div>

                        <div className="flex items-center gap-2 shrink-0 lg:pt-1">
                          {finding.isAutoRectifiable && onInstantRectification && (
                            <button
                              type="button"
                              onClick={() => onInstantRectification([finding.id])}
                              className="px-3 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-600 text-emerald-800 hover:text-white border border-emerald-300 font-semibold text-xs inline-flex items-center gap-1.5 transition-colors whitespace-nowrap"
                            >
                              <Wrench className="w-3.5 h-3.5" />
                              <span>Rectify</span>
                            </button>
                          )}

                          {finding.id === 'audit-missing-entity-profile' && onOpenEntitySetup && (
                            <button
                              type="button"
                              onClick={onOpenEntitySetup}
                              className="px-3 py-1.5 rounded-lg bg-white hover:bg-slate-100 text-slate-800 border border-slate-300 font-semibold text-xs inline-flex items-center gap-1.5 transition-colors whitespace-nowrap"
                            >
                              <Building2 className="w-3.5 h-3.5" />
                              <span>Open Entity Setup</span>
                            </button>
                          )}

                          {finding.id === 'audit-unimplemented-schedule' && onJumpToRepaymentSchedule && (
                            <button
                              type="button"
                              onClick={onJumpToRepaymentSchedule}
                              className="px-3 py-1.5 rounded-lg bg-white hover:bg-slate-100 text-slate-800 border border-slate-300 font-semibold text-xs inline-flex items-center gap-1.5 transition-colors whitespace-nowrap"
                            >
                              <CalendarCheck2 className="w-3.5 h-3.5" />
                              <span>Review Schedule</span>
                            </button>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )
            ) : (
              <div className="divide-y divide-slate-200 border border-slate-200 rounded-xl overflow-hidden bg-white">
                {auditReport.passedChecks.map(chk => (
                  <div key={chk.id} className="p-3.5 flex items-start gap-3 text-xs">
                    <ClipboardCheck className="w-4 h-4 text-emerald-700 mt-0.5 shrink-0" />
                    <div className="space-y-0.5">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="font-bold text-slate-900">{chk.label}</span>
                        <span className="text-[11px] text-slate-500">· {chk.standardRef}</span>
                      </div>
                      <p className="text-slate-600">{chk.detail}</p>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </div>

      {/* Model-Wide Validation & Fixed-Driver Integrity Banner (Preserved) */}
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
                ? 'Model Validation Passed: Fixed Primary Inputs Strictly Preserved · Zero Accounting Variance Across All Selected Years'
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
          <div className="text-slate-900 font-bold">1. Operating Cash Flow</div>
          <p className="text-slate-600 leading-relaxed text-[11px]">
            Derived via the <strong>Indirect Method</strong> from Net Profit Before Tax, adjusting for non-cash depreciation, finance cost, and working capital shifts (Inventory, Receivables, Payables).
          </p>
        </div>

        <div className="p-4 bg-white/90 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-900 font-bold">2. Investing Cash Flow</div>
          <p className="text-slate-600 leading-relaxed text-[11px]">
            Tied directly to the <strong>Schedule of Fixed Assets</strong> capital expenditure additions, asset disposals, and financial investment allocations.
          </p>
        </div>

        <div className="p-4 bg-white/90 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1.5 shadow-sm">
          <div className="text-slate-900 font-bold">3. Financing Cash Flow</div>
          <p className="text-slate-600 leading-relaxed text-[11px]">
            Preserves <strong>mandatory direct financing cash flows</strong> separately for every historical and projected financial year.
          </p>
        </div>
      </div>

      {/* Reconciliation Audit Table */}
      <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
        <div className="p-4 border-b border-slate-200 bg-slate-50/90 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2 text-[11px] text-slate-600">
              <span className="font-bold text-emerald-700">
                {companyProfile.companyName || 'Active Entity'}
              </span>
              <span aria-hidden="true">·</span>
              <span>Audited Schedule</span>
            </div>
            <h3 className="text-sm font-bold text-slate-900 tracking-tight mt-0.5">
              Multi-Year Cash Flow Audit Reconciliation Schedule
            </h3>
            <p className="text-[11px] text-slate-600 mt-0.5">
              All monetary figures in {companyProfile.reportingCurrency} ({companyProfile.reportingUnit}) · Complete {cashFlowYears.length}-Year Horizon
            </p>
          </div>
          <span className="text-xs font-mono font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded self-start sm:self-auto">
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
                <td className="py-2.5 px-4 font-sans text-amber-900">
                  C. Financing Cash Flow (Direct User Entry)
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
                <td className="py-2.5 px-4 font-sans text-slate-600">Opening Cash &amp; Cash Equivalents</td>
                {cashFlowYears.map(cf => (
                  <td key={cf.year} className="py-2.5 px-3 text-right text-slate-700">
                    {fmtAuditNum(cf.openingCash)}
                  </td>
                ))}
              </tr>

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
                <td className="py-3 px-4 font-sans text-slate-900">
                  Reconciliation Difference <span className="text-[10px] text-slate-500 font-normal">(Derived - Stated)</span>
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

      {/* Debt Servicing & Repayment Capacity Audit Matrix (Preserved) */}
      {scheduleCoverageYears.length > 0 && (
        <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-4 border-b border-slate-200 bg-slate-50/90 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <div className="flex items-center gap-2 text-[11px] text-slate-600">
                <span className="font-bold text-emerald-700">
                  {parsedSchedule?.metadata?.sourceFileName || companyProfile.companyName || 'Active Facility'}
                </span>
                {parsedSchedule?.metadata?.investmentMade && (
                  <>
                    <span aria-hidden="true">·</span>
                    <span className="font-mono">
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

          {minHistEbitdaDscr !== null && minProjEbitdaDscr !== null && minHistEbitdaDscr < 1.0 && minProjEbitdaDscr >= 1.25 && (
            <div className="px-4 py-2.5 bg-gradient-to-r from-rose-50/80 via-slate-50 to-emerald-50/80 border-b border-slate-200 grid grid-cols-1 md:grid-cols-2 gap-3 text-[11px]">
              <div className="text-rose-900">
                <strong className="font-bold">Historical Pre-Rescheduling Distress:</strong>{' '}
                Compressed operating margins and legacy debt burden yielded Historical EBITDA DSCR of{' '}
                <span className="font-mono font-bold">{minHistEbitdaDscr.toFixed(2)}x (&lt; 1.00x)</span>, demonstrating inability to service unrescheduled obligations.
              </div>
              <div className="text-emerald-900">
                <strong className="font-bold">Projected Post-Rescheduling Viability:</strong>{' '}
                Rescheduled repayment terms lift Projected EBITDA DSCR to{' '}
                <span className="font-mono font-bold">&ge; {minProjEbitdaDscr.toFixed(2)}x</span> with positive liquidity across scheduled periods.
              </div>
            </div>
          )}

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
                  <tr key={row.year} title={row.reschedulingContext || ''} className="hover:bg-slate-50">
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
                      <span
                        title={row.reschedulingContext || ''}
                        className={`inline-block text-[10px] font-bold px-2 py-0.5 rounded border ${
                          row.capacityStatus === 'STRONG'
                            ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                            : row.capacityStatus === 'ADEQUATE'
                            ? 'bg-sky-50 text-sky-800 border-sky-300'
                            : row.capacityStatus === 'TIGHT'
                            ? 'bg-amber-50 text-amber-900 border-amber-300'
                            : 'bg-rose-50 text-rose-800 border-rose-300'
                        }`}
                      >
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
