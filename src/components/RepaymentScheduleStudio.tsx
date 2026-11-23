import React, { useState, useRef } from 'react';
import {
  Upload,
  FileSpreadsheet,
  FileText,
  CheckCircle2,
  AlertTriangle,
  Play,
  RotateCcw,
  Lock,
  ChevronDown,
  ChevronUp,
  ClipboardPaste,
  ShieldCheck,
  Sparkles,
  Trash2,
} from 'lucide-react';
import {
  CompanyProfile,
  DebtServiceCoverageYear,
  ParsedRepaymentSchedule,
  RepaymentInstallmentRow,
  RepaymentScheduleSnapshot,
} from '../types/financial';
import { ImplementScheduleResult } from '../services/financialEngine';
import {
  getAttachedReferenceRepaymentSchedule,
  parseRepaymentScheduleText,
  parseScheduleNumber,
  parseUploadedRepaymentSchedule,
  rebuildScheduleSummaries,
} from '../services/repaymentScheduleParser';

interface RepaymentScheduleStudioProps {
  companyProfile: CompanyProfile;
  allModelYears: number[];
  parsedSchedule: ParsedRepaymentSchedule | null;
  onSetParsedSchedule: (schedule: ParsedRepaymentSchedule | null) => void;
  scheduleSnapshot: RepaymentScheduleSnapshot | null;
  isScheduleImplemented: boolean;
  lastImplementationResult: ImplementScheduleResult | null;
  debtServiceCoverage: DebtServiceCoverageYear[];
  onImplementRepaymentSchedule: (options: {
    autoIncludeScheduleYears: boolean;
    alignHistoricalScale: boolean;
  }) => void;
  onRevertScheduleSnapshot: () => void;
}

function fmtAmt(val: number): string {
  if (!Number.isFinite(val) || Math.abs(val) < 0.005) return '0';
  const hasCents = Math.abs(val - Math.round(val)) >= 0.005;
  return val.toLocaleString(undefined, hasCents
    ? { minimumFractionDigits: 2, maximumFractionDigits: 2 }
    : { maximumFractionDigits: 0 }
  );
}

const SAMPLE_PASTE_PLACEHOLDER = `Event Type\tEMI Date\tInst. Principal\tInst. Profit\tBalance\tTotal
Disbursement\t01/01/2027\t0.00\t0.00\t24,30,778.97\t0.00
Repayment 1\t01/02/2027\t26,297.85\t32,444.15\t24,04,481.12\t58,742.00
Repayment 2\t01/03/2027\t29,754.64\t28,987.36\t23,74,726.48\t58,742.00`;

export const RepaymentScheduleStudio: React.FC<RepaymentScheduleStudioProps> = ({
  companyProfile,
  allModelYears,
  parsedSchedule,
  onSetParsedSchedule,
  scheduleSnapshot,
  isScheduleImplemented,
  lastImplementationResult,
  debtServiceCoverage,
  onImplementRepaymentSchedule,
  onRevertScheduleSnapshot,
}) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [parseError, setParseError] = useState<string | null>(null);
  const [isDragActive, setIsDragActive] = useState(false);
  const [showPasteDrawer, setShowPasteDrawer] = useState(false);
  const [pasteText, setPasteText] = useState('');
  const [showInstallmentLedger, setShowInstallmentLedger] = useState(false);
  const [autoIncludeScheduleYears, setAutoIncludeScheduleYears] = useState(true);
  const [alignHistoricalScale, setAlignHistoricalScale] = useState(true);

  const handleFileSelected = async (file: File) => {
    setIsParsing(true);
    setParseError(null);
    try {
      const parsed = await parseUploadedRepaymentSchedule(
        file,
        companyProfile.financialYearEnd,
        allModelYears
      );
      onSetParsedSchedule(parsed);
      if (parsed.rows.filter(r => !r.isDisbursement).length === 0) {
        setParseError(
          'No repayment installment rows could be identified. Verify headers (EMI Date, Inst. Principal, Inst. Profit) or use the Paste Schedule Table fallback.'
        );
      }
    } catch (err) {
      setParseError(err instanceof Error ? err.message : 'Unable to process the uploaded file.');
    } finally {
      setIsParsing(false);
    }
  };

  const handleInputFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFileSelected(file);
    }
    e.target.value = '';
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragActive(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      handleFileSelected(file);
    }
  };

  const handleLoadAttachedSchedule = () => {
    setParseError(null);
    const ref = getAttachedReferenceRepaymentSchedule(
      companyProfile.financialYearEnd,
      allModelYears
    );
    onSetParsedSchedule(ref);
  };

  const handleParsePastedText = () => {
    if (!pasteText.trim()) {
      setParseError('Paste tab-separated or CSV schedule rows before parsing.');
      return;
    }
    setParseError(null);
    const parsed = parseRepaymentScheduleText(
      pasteText,
      companyProfile.financialYearEnd,
      allModelYears,
      'Pasted Client Schedule'
    );
    onSetParsedSchedule(parsed);
    if (parsed.isValid) {
      setShowPasteDrawer(false);
    }
  };

  const handleDateFormatChange = (fmt: 'DD/MM/YYYY' | 'MM/DD/YYYY') => {
    if (!parsedSchedule) return;
    const rebuilt = rebuildScheduleSummaries(
      parsedSchedule.rows,
      parsedSchedule.metadata,
      companyProfile.financialYearEnd,
      allModelYears,
      fmt
    );
    onSetParsedSchedule(rebuilt);
  };

  const handleUpdateInstallmentCell = (
    rowId: string,
    field: 'instPrincipal' | 'instProfit' | 'emiDateRaw',
    rawVal: string
  ) => {
    if (!parsedSchedule) return;
    const nextRows: RepaymentInstallmentRow[] = parsedSchedule.rows.map(r => {
      if (r.id !== rowId) return r;
      if (field === 'instPrincipal') {
        const p = Math.max(0, parseScheduleNumber(rawVal));
        return { ...r, instPrincipal: p, totalInstallment: Number((p + r.instProfit).toFixed(2)) };
      }
      if (field === 'instProfit') {
        const pr = Math.max(0, parseScheduleNumber(rawVal));
        return { ...r, instProfit: pr, totalInstallment: Number((r.instPrincipal + pr).toFixed(2)) };
      }
      return { ...r, emiDateRaw: rawVal.trim() };
    });

    const rebuilt = rebuildScheduleSummaries(
      nextRows,
      parsedSchedule.metadata,
      companyProfile.financialYearEnd,
      allModelYears,
      parsedSchedule.metadata.dateFormatDetected === 'MM/DD/YYYY' ? 'MM/DD/YYYY' : 'DD/MM/YYYY'
    );
    onSetParsedSchedule(rebuilt);
  };

  const coveredYearsCount = debtServiceCoverage.filter(d => d.isScheduleCovered).length;

  return (
    <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl p-5 space-y-5 shadow-sm">
      {/* Studio Header */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-slate-200 pb-4">
        <div className="space-y-1">
          <div className="flex flex-wrap items-center gap-2">
            <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
            <h3 className="text-base font-bold text-slate-900 tracking-tight">
              Repayment Schedule Upload, Extraction &amp; Statement Optimization
            </h3>
            {isScheduleImplemented && (
              <span className="text-[10px] font-mono font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-300 inline-flex items-center gap-1">
                <Lock className="w-3 h-3" />
                Implemented ({coveredYearsCount}Y Locked)
              </span>
            )}
          </div>
          <p className="text-xs text-slate-600">
            Upload a client repayment schedule (PDF, Excel, CSV, or scanned image). Exact annual{' '}
            <strong className="text-slate-900">Inst. Principal</strong> populates{' '}
            <strong className="text-emerald-700">Repayment from Borrowings &amp; Financing</strong> and exact annual{' '}
            <strong className="text-slate-900">Inst. Profit</strong> populates{' '}
            <strong className="text-amber-800">Finance Cost</strong>, then optimizes projected cash flows for sustainable servicing.
          </p>
        </div>

        {/* Safe Fallbacks & Revert Action Strip */}
        <div className="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={handleLoadAttachedSchedule}
            className="px-3 py-1.5 rounded-lg text-xs font-semibold bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-300 transition-colors flex items-center gap-1.5"
          >
            <FileText className="w-3.5 h-3.5" />
            Load Client ABABIL Schedule (2027–2032)
          </button>

          <button
            type="button"
            onClick={() => setShowPasteDrawer(!showPasteDrawer)}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold border transition-colors flex items-center gap-1.5 ${
              showPasteDrawer
                ? 'bg-indigo-600 text-white border-indigo-600'
                : 'bg-white hover:bg-slate-50 text-slate-700 border-slate-300'
            }`}
          >
            <ClipboardPaste className="w-3.5 h-3.5" />
            Paste Schedule Table
          </button>

          {scheduleSnapshot && (
            <button
              type="button"
              onClick={onRevertScheduleSnapshot}
              className="px-3 py-1.5 rounded-lg text-xs font-semibold bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 transition-colors flex items-center gap-1.5"
              title={`Revert financial model to snapshot captured at ${new Date(scheduleSnapshot.capturedAt).toLocaleTimeString()}`}
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Revert to Pre-Schedule Snapshot
            </button>
          )}
        </div>
      </div>

      {/* File Upload Dropzone */}
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setIsDragActive(true);
        }}
        onDragLeave={() => setIsDragActive(false)}
        onDrop={handleDrop}
        className={`border-2 border-dashed rounded-xl p-5 transition-all flex flex-col sm:flex-row items-center justify-between gap-4 ${
          isDragActive
            ? 'border-emerald-500 bg-emerald-50/70'
            : 'border-slate-300 bg-slate-50/80 hover:border-slate-400'
        }`}
      >
        <div className="flex items-center gap-3.5">
          <div className="w-10 h-10 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700 shrink-0">
            <Upload className="w-5 h-5" />
          </div>
          <div className="text-xs">
            <div className="font-bold text-slate-900">
              {isParsing
                ? 'Reading & extracting repayment schedule rows...'
                : 'Drop Client Repayment Schedule here or browse file'}
            </div>
            <div className="text-slate-600 mt-0.5">
              Supports <span className="font-mono text-slate-800 font-medium">.pdf</span>,{' '}
              <span className="font-mono text-slate-800 font-medium">.xlsx / .xls</span>,{' '}
              <span className="font-mono text-slate-800 font-medium">.csv / .txt</span>, and scanned{' '}
              <span className="font-mono text-slate-800 font-medium">.png / .jpg</span> schedules (Fiscal Year End:{' '}
              <strong className="text-slate-900">{companyProfile.financialYearEnd}</strong>)
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,.xlsx,.xls,.csv,.txt,.png,.jpg,.jpeg,.webp"
            onChange={handleInputFileChange}
            className="hidden"
          />
          <button
            type="button"
            disabled={isParsing}
            onClick={() => fileInputRef.current?.click()}
            className="px-4 py-2 rounded-lg text-xs font-bold bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white transition-colors shadow-sm"
          >
            {isParsing ? 'Extracting...' : 'Upload Repayment Schedule'}
          </button>

          {parsedSchedule && (
            <button
              type="button"
              onClick={() => {
                onSetParsedSchedule(null);
                setParseError(null);
              }}
              className="p-2 rounded-lg text-slate-500 hover:text-rose-600 hover:bg-rose-50 bg-white border border-slate-300 transition-colors"
              title="Clear uploaded schedule preview"
            >
              <Trash2 className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* Paste Table / CSV Safe Fallback Drawer */}
      {showPasteDrawer && (
        <div className="p-4 bg-indigo-50/50 border border-indigo-200 rounded-xl space-y-3 text-xs">
          <div className="flex items-center justify-between">
            <span className="font-bold text-indigo-900">
              Safe Fallback: Paste Repayment Schedule Table (Tab-Separated or CSV)
            </span>
            <span className="text-[11px] text-slate-600">
              Requires EMI Date, Inst. Principal &amp; Inst. Profit columns
            </span>
          </div>
          <textarea
            value={pasteText}
            onChange={(e) => setPasteText(e.target.value)}
            rows={5}
            placeholder={SAMPLE_PASTE_PLACEHOLDER}
            className="w-full bg-white border border-slate-300 rounded-lg p-3 text-xs font-mono text-slate-900 placeholder-slate-400 focus:outline-none focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
          />
          <div className="flex items-center justify-end gap-2">
            <button
              type="button"
              onClick={() => setShowPasteDrawer(false)}
              className="px-3 py-1.5 rounded-lg text-xs text-slate-700 hover:text-slate-900 bg-white hover:bg-slate-50 border border-slate-300"
            >
              Close
            </button>
            <button
              type="button"
              onClick={handleParsePastedText}
              className="px-4 py-1.5 rounded-lg text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white"
            >
              Parse Pasted Schedule
            </button>
          </div>
        </div>
      )}

      {/* Parse Error Banner */}
      {parseError && (
        <div className="p-3.5 rounded-xl bg-rose-50 border border-rose-300 text-rose-800 text-xs flex items-start gap-2.5">
          <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
          <div>
            <span className="font-bold">Extraction Issue: </span>
            <span>{parseError}</span>
          </div>
        </div>
      )}

      {/* Extracted Schedule Preview, Validation & Implement CTA */}
      {parsedSchedule && (
        <div className="space-y-4 border-t border-slate-200 pt-4">
          {/* Metadata KPI Bar */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 text-xs">
            <div className="p-3 bg-slate-50/90 border border-slate-200 rounded-xl">
              <div className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Source File</div>
              <div className="font-semibold text-slate-900 truncate mt-0.5" title={parsedSchedule.metadata.sourceFileName}>
                {parsedSchedule.metadata.sourceFileName}
              </div>
              <div className="text-[10px] text-emerald-700 font-mono font-semibold mt-0.5">
                {parsedSchedule.rows.filter(r => !r.isDisbursement).length} Installments
              </div>
            </div>

            <div className="p-3 bg-slate-50/90 border border-slate-200 rounded-xl">
              <div className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Facility / Principal</div>
              <div className="font-mono font-bold text-slate-900 tabular-nums mt-0.5">
                {companyProfile.reportingCurrency} {fmtAmt(parsedSchedule.metadata.investmentMade || parsedSchedule.grandTotalPrincipal)}
              </div>
              <div className="text-[10px] text-slate-500 mt-0.5">
                {parsedSchedule.metadata.interval || 'Monthly'}
              </div>
            </div>

            <div className="p-3 bg-rose-50/75 border border-rose-200 rounded-xl">
              <div className="text-[10px] text-rose-800 uppercase tracking-wider font-semibold">Total Inst. Principal</div>
              <div className="font-mono font-bold text-rose-700 tabular-nums mt-0.5">
                {companyProfile.reportingCurrency} {fmtAmt(parsedSchedule.grandTotalPrincipal)}
              </div>
              <div className="text-[10px] text-slate-600 mt-0.5">→ Borrowings Repayment</div>
            </div>

            <div className="p-3 bg-amber-50/80 border border-amber-200 rounded-xl">
              <div className="text-[10px] text-amber-900 uppercase tracking-wider font-semibold">Total Inst. Profit</div>
              <div className="font-mono font-bold text-amber-800 tabular-nums mt-0.5">
                {companyProfile.reportingCurrency} {fmtAmt(parsedSchedule.grandTotalProfit)}
              </div>
              <div className="text-[10px] text-slate-600 mt-0.5">→ Finance Cost</div>
            </div>

            <div className="p-3 bg-sky-50/75 border border-sky-200 rounded-xl">
              <div className="text-[10px] text-sky-800 uppercase tracking-wider font-semibold">Schedule Horizon</div>
              <div className="font-mono font-bold text-sky-800 mt-0.5">
                {parsedSchedule.annualSummaries[0]?.financialYear ?? '-'} –{' '}
                {parsedSchedule.annualSummaries[parsedSchedule.annualSummaries.length - 1]?.financialYear ?? '-'}
              </div>
              <div className="text-[10px] text-slate-600 mt-0.5">
                {parsedSchedule.annualSummaries.length} Financial Years
              </div>
            </div>

            <div className="p-3 bg-slate-50/90 border border-slate-200 rounded-xl">
              <label className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold block">Date Format</label>
              <select
                value={parsedSchedule.metadata.dateFormatDetected === 'MM/DD/YYYY' ? 'MM/DD/YYYY' : 'DD/MM/YYYY'}
                onChange={(e) => handleDateFormatChange(e.target.value as 'DD/MM/YYYY' | 'MM/DD/YYYY')}
                className="mt-1 w-full bg-white border border-slate-300 rounded px-2 py-1 text-xs font-mono text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500"
              >
                <option value="DD/MM/YYYY">DD/MM/YYYY</option>
                <option value="MM/DD/YYYY">MM/DD/YYYY</option>
              </select>
            </div>
          </div>

          {/* Validation Errors / Warnings */}
          {(parsedSchedule.validationErrors.length > 0 || parsedSchedule.validationWarnings.length > 0) && (
            <div className="space-y-2 text-xs">
              {parsedSchedule.validationErrors.map((err, idx) => (
                <div key={`err-${idx}`} className="p-3 rounded-xl bg-rose-50 border border-rose-300 text-rose-800 flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
                  <span>{err}</span>
                </div>
              ))}
              {parsedSchedule.validationWarnings.map((warn, idx) => (
                <div key={`warn-${idx}`} className="p-3 rounded-xl bg-amber-50 border border-amber-300 text-amber-900 flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                  <span>{warn}</span>
                </div>
              ))}
            </div>
          )}

          {/* Year-Wise Extracted Schedule Summary Table */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="text-xs font-bold text-slate-900">
                Year-Wise Scheduled Aggregation (Exact Principal &amp; Finance Cost Mapping)
              </div>
              <button
                type="button"
                onClick={() => setShowInstallmentLedger(!showInstallmentLedger)}
                className="text-xs text-sky-700 hover:text-sky-800 font-semibold inline-flex items-center gap-1"
              >
                {showInstallmentLedger ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                {showInstallmentLedger
                  ? 'Hide Installment-by-Installment Ledger'
                  : `Inspect & Edit ${parsedSchedule.rows.length} Installment Rows`}
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-xs font-mono">
                <thead className="bg-slate-100 text-slate-700 border-b border-slate-200 text-[11px]">
                  <tr>
                    <th className="py-2.5 px-4 text-left font-sans font-semibold text-slate-900">Financial Year</th>
                    <th className="py-2.5 px-3 text-right font-sans font-semibold">Installments</th>
                    <th className="py-2.5 px-3 text-right font-sans font-semibold text-rose-700">
                      Inst. Principal → Borrowings Repayment
                    </th>
                    <th className="py-2.5 px-3 text-right font-sans font-semibold text-amber-800">
                      Inst. Profit → Finance Cost
                    </th>
                    <th className="py-2.5 px-3 text-right font-sans font-semibold">Total Debt Service</th>
                    <th className="py-2.5 px-3 text-right font-sans font-semibold">Closing Balance</th>
                    <th className="py-2.5 px-4 text-right font-sans font-semibold">Model Horizon Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200/80 tabular-nums">
                  {parsedSchedule.annualSummaries.map(row => (
                    <tr key={row.financialYear} className="hover:bg-slate-50">
                      <td className="py-2 px-4 font-bold text-slate-900">FY {row.financialYear}</td>
                      <td className="py-2 px-3 text-right text-slate-700">{row.installmentCount}</td>
                      <td className="py-2 px-3 text-right font-bold text-rose-700">
                        {fmtAmt(row.totalPrincipal)}
                      </td>
                      <td className="py-2 px-3 text-right font-bold text-amber-800">
                        {fmtAmt(row.totalProfit)}
                      </td>
                      <td className="py-2 px-3 text-right font-bold text-slate-900">
                        {fmtAmt(row.totalDebtService)}
                      </td>
                      <td className="py-2 px-3 text-right text-slate-700">
                        {fmtAmt(row.closingBalance)}
                      </td>
                      <td className="py-2 px-4 text-right font-sans">
                        {row.isSelectedInModel ? (
                          <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 px-2 py-0.5 rounded">
                            In Model Horizon
                          </span>
                        ) : (
                          <span className="text-[10px] font-semibold text-sky-800 bg-sky-50 border border-sky-300 px-2 py-0.5 rounded">
                            Will Auto-Include
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                  <tr className="bg-emerald-50/75 font-bold text-emerald-800 border-t border-emerald-200">
                    <td className="py-2.5 px-4 font-sans">TOTAL SCHEDULE AGGREGATE</td>
                    <td className="py-2.5 px-3 text-right">
                      {parsedSchedule.annualSummaries.reduce((s, a) => s + a.installmentCount, 0)}
                    </td>
                    <td className="py-2.5 px-3 text-right text-rose-700">
                      {fmtAmt(parsedSchedule.grandTotalPrincipal)}
                    </td>
                    <td className="py-2.5 px-3 text-right text-amber-800">
                      {fmtAmt(parsedSchedule.grandTotalProfit)}
                    </td>
                    <td className="py-2.5 px-3 text-right text-slate-900">
                      {fmtAmt(parsedSchedule.grandTotalDebtService)}
                    </td>
                    <td className="py-2.5 px-3 text-right">0</td>
                    <td className="py-2.5 px-4 text-right font-sans text-[11px]">
                      100% Exact Preservation
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Expandable Installment-Level Editable Ledger */}
          {showInstallmentLedger && (
            <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
              <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-xs">
                <span className="font-semibold text-slate-900">
                  Full Installment Ledger (Editable Verification Fallback)
                </span>
                <span className="text-[11px] text-slate-600">
                  Editing any cell automatically recalculates year-wise totals
                </span>
              </div>
              <div className="max-h-72 overflow-y-auto overflow-x-auto">
                <table className="w-full text-xs font-mono">
                  <thead className="bg-slate-100 sticky top-0 text-slate-700 border-b border-slate-200 text-[10px]">
                    <tr>
                      <th className="py-2 px-3 text-left font-sans font-semibold">Event</th>
                      <th className="py-2 px-3 text-left font-sans font-semibold">EMI Date</th>
                      <th className="py-2 px-3 text-center font-sans font-semibold">FY</th>
                      <th className="py-2 px-3 text-right font-sans font-semibold">Inst. Principal</th>
                      <th className="py-2 px-3 text-right font-sans font-semibold">Inst. Profit</th>
                      <th className="py-2 px-3 text-right font-sans font-semibold">Total</th>
                      <th className="py-2 px-3 text-right font-sans font-semibold">Balance</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200/70 tabular-nums">
                    {parsedSchedule.rows.map(r => (
                      <tr key={r.id} className={r.isDisbursement ? 'bg-sky-50/70 text-sky-900' : 'hover:bg-slate-50'}>
                        <td className="py-1.5 px-3 font-sans text-slate-800">{r.eventType}</td>
                        <td className="py-1.5 px-3">
                          <input
                            type="text"
                            value={r.emiDateRaw}
                            onChange={(e) => handleUpdateInstallmentCell(r.id, 'emiDateRaw', e.target.value)}
                            className="w-24 bg-white border border-slate-300 rounded px-1.5 py-0.5 text-xs text-slate-900 focus:outline-none focus:border-emerald-600"
                          />
                        </td>
                        <td className="py-1.5 px-3 text-center text-slate-700">{r.financialYear}</td>
                        <td className="py-1.5 px-3 text-right">
                          {r.isDisbursement ? (
                            '0.00'
                          ) : (
                            <input
                              type="number"
                              step="0.01"
                              value={r.instPrincipal}
                              onChange={(e) => handleUpdateInstallmentCell(r.id, 'instPrincipal', e.target.value)}
                              className="w-28 bg-white border border-slate-300 rounded px-1.5 py-0.5 text-right text-xs text-rose-700 font-semibold focus:outline-none focus:border-emerald-600"
                            />
                          )}
                        </td>
                        <td className="py-1.5 px-3 text-right">
                          {r.isDisbursement ? (
                            '0.00'
                          ) : (
                            <input
                              type="number"
                              step="0.01"
                              value={r.instProfit}
                              onChange={(e) => handleUpdateInstallmentCell(r.id, 'instProfit', e.target.value)}
                              className="w-28 bg-white border border-slate-300 rounded px-1.5 py-0.5 text-right text-xs text-amber-800 font-semibold focus:outline-none focus:border-emerald-600"
                            />
                          )}
                        </td>
                        <td className="py-1.5 px-3 text-right text-slate-900 font-medium">{fmtAmt(r.totalInstallment)}</td>
                        <td className="py-1.5 px-3 text-right text-slate-600">{fmtAmt(r.balance)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Implementation Controls & Primary CTA */}
          <div className="p-4 bg-emerald-50/60 border border-emerald-200 rounded-xl flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div className="flex flex-wrap items-center gap-4 text-xs text-slate-700">
              <label className="inline-flex items-center gap-2 cursor-pointer">
                <input
                  type="checkbox"
                  checked={autoIncludeScheduleYears}
                  onChange={(e) => setAutoIncludeScheduleYears(e.target.checked)}
                  className="rounded border-slate-300 bg-white text-emerald-600 focus:ring-emerald-500"
                />
                <span>Auto-include all schedule years in selected horizon</span>
              </label>

              <label className="inline-flex items-center gap-2 cursor-pointer">
                <input
                  type="checkbox"
                  checked={alignHistoricalScale}
                  onChange={(e) => setAlignHistoricalScale(e.target.checked)}
                  className="rounded border-slate-300 bg-white text-emerald-600 focus:ring-emerald-500"
                />
                <span>Proportionately align business scale if capacity lift &gt; 1.5x</span>
              </label>
            </div>

            <button
              type="button"
              disabled={!parsedSchedule.isValid}
              onClick={() =>
                onImplementRepaymentSchedule({
                  autoIncludeScheduleYears,
                  alignHistoricalScale,
                })
              }
              className="px-5 py-2.5 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white shadow-sm transition-all flex items-center justify-center gap-2 shrink-0"
            >
              <Play className="w-4 h-4 fill-white" />
              Implement Repayment Schedule
            </button>
          </div>
        </div>
      )}

      {/* Implementation Summary Feedback */}
      {isScheduleImplemented && lastImplementationResult && (
        <div className="p-4 rounded-xl bg-emerald-50 border border-emerald-300 text-xs text-emerald-900 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex items-start gap-2.5">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
            <div>
              <div className="font-bold text-slate-900">
                Repayment Schedule Successfully Implemented &amp; Financial Statements Optimized
              </div>
              <p className="text-emerald-800 mt-0.5">
                Locked exact annual Principal &amp; Finance Cost across years{' '}
                <strong className="font-mono">{lastImplementationResult.coveredYears.join(', ')}</strong>. Minimum achieved{' '}
                <strong className="font-mono">EBITDA DSCR: {lastImplementationResult.minEbitdaDscr}x</strong> ·{' '}
                <strong className="font-mono">Cash DSCR: {lastImplementationResult.minCashDscr}x</strong>
                {lastImplementationResult.scaleFactorApplied > 1.02 && (
                  <span> · Operating scale calibrated by {lastImplementationResult.scaleFactorApplied.toFixed(2)}x for sustainable repayment coverage</span>
                )}.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Debt Servicing & Repayment Capacity Analysis (DSCR) Matrix */}
      {debtServiceCoverage.length > 0 && (
        <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
          <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-emerald-600" />
              <span className="text-xs font-bold text-slate-900">
                Debt Servicing &amp; Repayment Capacity Analysis (DSCR)
              </span>
              {isScheduleImplemented && (
                <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 px-2 py-0.5 rounded inline-flex items-center gap-1">
                  <Sparkles className="w-3 h-3" />
                  Schedule-Optimized
                </span>
              )}
            </div>
            <span className="text-[11px] text-slate-600 font-mono">
              Target Sustainability Threshold: EBITDA DSCR ≥ 1.25x · Positive Closing Cash
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs font-mono">
              <thead className="bg-slate-100 text-slate-700 border-b border-slate-200 text-[11px]">
                <tr>
                  <th className="py-2.5 px-3 text-left font-sans font-semibold text-slate-900">Year</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold">EBITDA</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold">Pre-Fin OCF</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold text-rose-700">Principal</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold text-amber-800">Finance Cost</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold">Debt Service</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold">EBITDA DSCR</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold">Cash DSCR</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold">Closing Cash</th>
                  <th className="py-2.5 px-3 text-right font-sans font-semibold">Capacity</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/70 tabular-nums">
                {debtServiceCoverage.map(row => (
                  <tr
                    key={row.year}
                    className={row.isScheduleCovered ? 'bg-emerald-50/45 hover:bg-emerald-50/80' : 'hover:bg-slate-50'}
                  >
                    <td className="py-2 px-3 font-bold text-slate-900 whitespace-nowrap">
                      {row.year}{' '}
                      <span className="text-[9px] font-sans font-normal text-slate-500">
                        {row.isProjected ? '(P)' : '(H)'}
                      </span>
                      {row.isScheduleCovered && (
                        <Lock className="w-2.5 h-2.5 text-emerald-600 inline ml-1" />
                      )}
                    </td>
                    <td className="py-2 px-3 text-right text-slate-800">{fmtAmt(row.ebitda)}</td>
                    <td className="py-2 px-3 text-right text-slate-800">{fmtAmt(row.operatingCashFlowBeforeFinance)}</td>
                    <td className="py-2 px-3 text-right text-rose-700 font-medium">{fmtAmt(row.scheduledPrincipal)}</td>
                    <td className="py-2 px-3 text-right text-amber-800 font-medium">{fmtAmt(row.scheduledFinanceCost)}</td>
                    <td className="py-2 px-3 text-right font-bold text-slate-900">{fmtAmt(row.totalDebtService)}</td>
                    <td className={`py-2 px-3 text-right font-bold ${
                      row.ebitdaDscr >= 1.25 ? 'text-emerald-700' : row.ebitdaDscr >= 1.0 ? 'text-amber-800' : 'text-rose-700'
                    }`}>
                      {row.ebitdaDscr.toFixed(2)}x
                    </td>
                    <td className={`py-2 px-3 text-right font-bold ${
                      row.cashDscr >= 1.15 ? 'text-emerald-700' : row.cashDscr >= 1.0 ? 'text-amber-800' : 'text-rose-700'
                    }`}>
                      {row.cashDscr.toFixed(2)}x
                    </td>
                    <td className={`py-2 px-3 text-right font-semibold ${
                      row.closingCash >= 0 ? 'text-emerald-700' : 'text-rose-700'
                    }`}>
                      {fmtAmt(row.closingCash)}
                    </td>
                    <td className="py-2 px-3 text-right font-sans">
                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded border ${
                        row.capacityStatus === 'STRONG'
                          ? 'text-emerald-800 bg-emerald-50 border-emerald-300'
                          : row.capacityStatus === 'ADEQUATE'
                          ? 'text-sky-800 bg-sky-50 border-sky-300'
                          : row.capacityStatus === 'TIGHT'
                          ? 'text-amber-900 bg-amber-50 border-amber-300'
                          : 'text-rose-800 bg-rose-50 border-rose-300'
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
