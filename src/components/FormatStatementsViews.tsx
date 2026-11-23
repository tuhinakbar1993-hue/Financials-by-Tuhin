import React, { useState } from 'react';
import {
  FileSpreadsheet,
  FileText,
  Download,
  Lock,
  Landmark,
  TrendingUp,
  Scale,
  Layers,
  Building2,
  BookOpen,
} from 'lucide-react';
import {
  CompanyProfile,
  FinancialNoteItem,
  PreviousFormatYearData,
  ProjectedFormatYearData,
  YearDataInput,
} from '../types/financial';
import {
  exportPreviousFinancialsExcel,
  exportPreviousFinancialsWord,
  exportPreviousFinancialsPdf,
  exportProjectedFinancialsExcel,
  exportProjectedFinancialsWord,
  exportProjectedFinancialsPdf,
} from '../services/exportService';

function fmtCell(val: number): string {
  if (val === 0) return '-';
  if (val < 0) return `(${Math.abs(val).toLocaleString()})`;
  return val.toLocaleString();
}

interface InlineDriverInputProps {
  value: number;
  isEditMode: boolean;
  onSave: (val: number) => void;
  title?: string;
}

const InlineDriverCell: React.FC<InlineDriverInputProps> = ({ value, isEditMode, onSave, title }) => {
  const [localVal, setLocalVal] = useState(String(value));
  const [editing, setEditing] = useState(false);

  React.useEffect(() => {
    setLocalVal(String(value));
  }, [value]);

  const commit = () => {
    setEditing(false);
    const num = Number(localVal);
    if (!Number.isNaN(num)) onSave(num);
  };

  if (isEditMode || editing) {
    return (
      <input
        type="number"
        value={localVal}
        onChange={e => setLocalVal(e.target.value)}
        onBlur={commit}
        onKeyDown={e => e.key === 'Enter' && commit()}
        autoFocus={editing}
        className="w-24 bg-slate-950 border border-emerald-500/80 rounded px-1.5 py-0.5 text-right font-mono text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-400"
      />
    );
  }

  return (
    <span
      onClick={() => setEditing(true)}
      title={title || 'Click to edit fixed primary user input'}
      className="cursor-pointer hover:text-emerald-300 hover:underline font-mono tabular-nums"
    >
      {fmtCell(value)}
    </span>
  );
};

// ============================================================================
// PREVIOUS FINANCIALS EXACT-FORMAT SUITE VIEW
// Matches Previous_Financials_aec5c4ae.pdf line-for-line
// ============================================================================

interface PreviousFormatSuiteViewProps {
  companyProfile: CompanyProfile;
  prevData: PreviousFormatYearData[];
  notes: FinancialNoteItem[];
  isEditMode: boolean;
  onUpdatePrimaryDriver?: (year: number, driver: keyof YearDataInput, val: number) => void;
}

export const PreviousFormatSuiteView: React.FC<PreviousFormatSuiteViewProps> = ({
  companyProfile,
  prevData,
  notes,
  isEditMode,
  onUpdatePrimaryDriver,
}) => {
  const [subTab, setSubTab] = useState<'all' | 'bs' | 'is' | 'cf' | 'eq' | 'fa'>('all');

  if (prevData.length === 0) {
    return (
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-8 text-center text-xs text-slate-400">
        Select at least one previous financial year to render the Previous Financials statements.
      </div>
    );
  }

  const maxYear = prevData[0].year;
  const ccyLabel = companyProfile.reportingCurrency === '৳' ? 'BDT' : companyProfile.reportingCurrency;
  const unitCaption = `(Figure in ${ccyLabel} · ${companyProfile.reportingUnit})`;
  const faAsc = [...prevData].sort((a, b) => a.year - b.year);

  const subTabs = [
    { id: 'all', label: 'All Previous Statements', icon: BookOpen },
    { id: 'bs', label: 'Financial Position', icon: Landmark },
    { id: 'is', label: 'Comprehensive Income', icon: TrendingUp },
    { id: 'cf', label: 'Cash Flows', icon: Scale },
    { id: 'eq', label: 'Changes in Equity', icon: Layers },
    { id: 'fa', label: '1.00 Fixed Assets & Notes', icon: Building2 },
  ] as const;

  return (
    <div className="space-y-5">
      {/* Control & Export Strip */}
      <div className="bg-slate-900 border border-emerald-900/50 rounded-2xl p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div className="space-y-1">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-[10px] font-mono font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-emerald-950 text-emerald-300 border border-emerald-700/60">
              Exact Format: Previous Financials
            </span>
            <span className="text-xs font-bold text-white">
              {prevData.length} Previous Years Selected ({faAsc[0].year}–{maxYear})
            </span>
          </div>
          <p className="text-[11px] text-slate-400">
            Reproduces exact headings, particulars, row sequence, descending 31.12.YYYY columns, Fixed Assets Schedule 1.00, and Statutory Notes.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => exportPreviousFinancialsExcel(companyProfile, prevData, notes)}
            className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            Previous Excel (.xlsx)
          </button>
          <button
            type="button"
            onClick={() => exportPreviousFinancialsWord(companyProfile, prevData, notes)}
            className="px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <FileText className="w-3.5 h-3.5" />
            Previous Word (.docx)
          </button>
          <button
            type="button"
            onClick={() => exportPreviousFinancialsPdf(companyProfile, prevData, notes)}
            className="px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            Previous PDF (.pdf)
          </button>
        </div>
      </div>

      {/* Sub-Tabs */}
      <div className="flex flex-wrap gap-1.5">
        {subTabs.map(t => {
          const Icon = t.icon;
          return (
            <button
              key={t.id}
              type="button"
              onClick={() => setSubTab(t.id)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors border ${
                subTab === t.id
                  ? 'bg-emerald-600 text-white border-emerald-400/50'
                  : 'bg-slate-900 text-slate-300 border-slate-800 hover:bg-slate-800'
              }`}
            >
              <Icon className="w-3.5 h-3.5" />
              {t.label}
            </button>
          );
        })}
      </div>

      {/* 1. STATEMENT OF FINANCIAL POSITION */}
      {(subTab === 'all' || subTab === 'bs') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-400 uppercase tracking-wide">
              STATEMENT OF FINANCIAL POSITION
            </div>
            <div className="text-xs font-semibold text-slate-300">
              AS AT 31 DECEMBER, {maxYear}
            </div>
            <div className="text-[11px] text-slate-400 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-700/80 min-w-[290px]">
                    Particulars
                  </th>
                  {prevData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-700/60 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/70 tabular-nums">
                <tr className="bg-slate-800/40 font-bold text-emerald-300">
                  <td colSpan={prevData.length + 1} className="py-2 px-4">PROPERTIES &amp; ASSETS</td>
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 text-slate-200 border-r border-slate-800/60">Fixed Assets (at cost less depreciation)</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-200 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.fixedAssetsWDV)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/25 font-semibold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">Current Assets</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.totalCurrentAssets)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 pl-7 text-slate-300 border-r border-slate-800/60">Advances, Deposits &amp; Pre-payments</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.advancesDepositsPrepayments)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20 bg-emerald-950/10">
                  <td className="py-2 px-4 pl-7 text-emerald-300 font-medium border-r border-slate-800/60 flex items-center gap-1.5">
                    <span>Stock &amp; Stores</span>
                    <Lock className="w-3 h-3 text-emerald-400" />
                  </td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-emerald-300 font-semibold border-r border-slate-800/40 last:border-r-0">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.stockAndStores}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'closingInventory', val)}
                          title="Fixed User Input: Closing Inventory"
                        />
                      ) : fmtCell(d.stockAndStores)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 pl-7 text-slate-300 border-r border-slate-800/60">Bills Receivable</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.billsReceivable)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 pl-7 text-slate-300 border-r border-slate-800/60">Cash &amp; Cash Equivalents</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.cashAndCashEquivalents)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-emerald-950/20 font-bold text-emerald-400 border-t border-b-2 border-emerald-700/70">
                  <td className="py-2.5 px-4 text-right border-r border-slate-800/60">Taka</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.totalAssetsTaka)}
                    </td>
                  ))}
                </tr>

                <tr className="bg-slate-800/40 font-bold text-sky-300">
                  <td colSpan={prevData.length + 1} className="py-2 px-4">CAPITAL &amp; LIABILITIES</td>
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Opening Capital</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.openingCapital)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Add. Net Profit for the year</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-emerald-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.netProfitForYear)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/25 font-semibold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">Subtotal</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.capitalSubtotal)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Less: Drawings</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-rose-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.drawings)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">Capital after drawings</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.capitalAfterDrawings)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/25 font-semibold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">Current Liabilities</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.totalCurrentLiabilities)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 pl-7 text-slate-300 border-r border-slate-800/60">For Expenses</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.forExpenses)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 pl-7 text-slate-300 border-r border-slate-800/60">Accounts Payable</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.accountsPayable)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 pl-7 text-slate-300 border-r border-slate-800/60">Short Term Loan</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.shortTermLoan)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-800/20">
                  <td className="py-2 px-4 pl-7 text-slate-300 border-r border-slate-800/60">Provision for Taxation</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.provisionForTaxation)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-emerald-950/20 font-bold text-emerald-400 border-t border-b-4 border-double border-emerald-500/80">
                  <td className="py-2.5 px-4 text-right border-r border-slate-800/60">Total</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.totalCapitalAndLiabilities)}
                    </td>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 2. STATEMENT OF COMPREHENSIVE INCOME */}
      {(subTab === 'all' || subTab === 'is') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-400">
              Statement of Comprehensive Income
            </div>
            <div className="text-xs font-semibold text-slate-300">
              For the year ended 31 December, {maxYear}
            </div>
            <div className="text-[11px] text-slate-400 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-700/80 min-w-[290px]">Particulars</th>
                  {prevData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-700/60 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/70 tabular-nums">
                <tr className="bg-emerald-950/10 font-semibold text-emerald-300">
                  <td className="py-2 px-4 border-r border-slate-800/60 flex items-center gap-1.5">
                    <span>Sales</span>
                    <Lock className="w-3 h-3 text-emerald-400" />
                  </td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.sales}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'salesRevenue', val)}
                          title="Fixed User Input: Sales Revenue"
                        />
                      ) : fmtCell(d.sales)}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Less- Cost of Goods Sold</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.costOfGoodsSold)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">Gross Profit</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.grossProfit)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/25 font-semibold text-slate-100">
                  <td className="py-2 px-4 border-r border-slate-800/60">Less- Administrative &amp; Selling Expenses</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.adminAndSellingTotal)}
                    </td>
                  ))}
                </tr>
                {([
                  ['Salary & Allowances', 'salaryAndAllowances'],
                  ['Telephone & Mobile Bill', 'telephoneAndMobileBill'],
                  ['Postage & Courier', 'postageAndCourier'],
                  ['Electric Charges', 'electricCharges'],
                  ['Entertainment', 'entertainment'],
                  ['Printing & Stationery', 'printingAndStationery'],
                  ['General Expenses', 'generalExpenses'],
                  ['Licence & Renewals', 'licenceAndRenewals'],
                  ['Audit & Professional Fee', 'auditAndProfessionalFee'],
                  ['Travelling & Conveyance', 'travellingAndConveyance'],
                  ['News Paper & Periodicals', 'newspaperAndPeriodicals'],
                  ['Selling Exp.', 'sellingExp'],
                  ['Depreciation', 'depreciation'],
                ] as const).map(([label, key]) => (
                  <tr key={key} className="hover:bg-slate-800/20">
                    <td className="py-1.5 px-4 pl-7 text-slate-400 border-r border-slate-800/60">{label}</td>
                    {prevData.map(d => (
                      <td key={d.year} className="py-1.5 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                        {fmtCell(d[key])}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-800/35 font-bold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">Net Operating Profit</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.netOperatingProfit)}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Add: Other Income</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.otherIncome)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/25 font-semibold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">Subtotal</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.operatingSubtotal)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-amber-950/10">
                  <td className="py-2 px-4 text-amber-300 border-r border-slate-800/60 flex items-center gap-1.5">
                    <span>Less: Financial Exp.</span>
                    <Lock className="w-3 h-3 text-amber-400" />
                  </td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-amber-300 border-r border-slate-800/40 last:border-r-0">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.financialExp}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'financeCost', val)}
                          title="Fixed User Input: Finance Cost"
                        />
                      ) : fmtCell(d.financialExp)}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Less- Income Tax Provision</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.incomeTaxProvision)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-emerald-950/20 font-bold text-emerald-400 border-t border-b-4 border-double border-emerald-500/80">
                  <td className="py-2.5 px-4 border-r border-slate-800/60 flex items-center gap-1.5">
                    <span>Net Profit After Income Tax</span>
                    <Lock className="w-3 h-3 text-emerald-400" />
                  </td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.netProfitAfterIncomeTax}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'netProfit', val)}
                          title="Fixed User Input: Net Profit After Tax"
                        />
                      ) : fmtCell(d.netProfitAfterIncomeTax)}
                    </td>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 3. STATEMENT OF CASH FLOWS */}
      {(subTab === 'all' || subTab === 'cf') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-400 uppercase tracking-wide">
              STATEMENT OF CASH FLOWS
            </div>
            <div className="text-xs font-semibold text-slate-300">
              For the year ended 31 December {maxYear}
            </div>
            <div className="text-[11px] text-slate-400 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-700/80 min-w-[340px]">Particulars</th>
                  {prevData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-700/60 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/70 tabular-nums">
                <tr className="bg-slate-800/40 font-bold text-emerald-300">
                  <td colSpan={prevData.length + 1} className="py-2 px-4">a) CASH FLOWS FROM OPERATING ACTIVITIES:</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-200 border-r border-slate-800/60">Net Income</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-200 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfNetIncome)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Add: Depreciation of Tangible Fixed Assets</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfDepreciation)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Add: (Increase) or decrease in Current Assets</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfChangeInCurrentAssets)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Add: (Increase) or decrease in Construction Working Progress</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfChangeInCWIP)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Add: Increase or (decrease) in Current Liabilities</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfChangeInCurrentLiabilities)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-emerald-400">
                  <td className="py-2 px-4 border-r border-slate-800/60">Net Cash from operating activities</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfNetCashOperating)}</td>)}
                </tr>

                <tr className="bg-slate-800/40 font-bold text-sky-300">
                  <td colSpan={prevData.length + 1} className="py-2 px-4">b) CASH FLOWS FROM INVESTING ACTIVITIES:</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Acquisition of Tangible Fixed Assets</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfAcquisitionPPE)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Disposal of Intangible Assets</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfDisposalIntangible)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-sky-400">
                  <td className="py-2 px-4 border-r border-slate-800/60">Net Cash used in investing Activities</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfNetCashInvesting)}</td>)}
                </tr>

                <tr className="bg-slate-800/40 font-bold text-amber-300">
                  <td colSpan={prevData.length + 1} className="py-2 px-4">c) CASH FLOWS FROM FINANCING ACTIVITIES:</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Drawings</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfDrawings)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">Opening Balance</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfOpeningBalance)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-amber-400">
                  <td className="py-2 px-4 border-r border-slate-800/60">Net Cash used in Financing Activities</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfNetCashFinancing)}</td>)}
                </tr>

                <tr className="font-semibold text-white">
                  <td className="py-2 px-4 border-r border-slate-800/60">d) Net Increase decrease in cash &amp; cash equivalents (A+B+C)</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfNetIncreaseCash)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300 border-r border-slate-800/60">e) Cash &amp; Cash Equivalents at beginning of period</td>
                  {prevData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-300 border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfCashBeginning)}</td>)}
                </tr>
                <tr className="bg-emerald-950/20 font-bold text-emerald-400 border-t border-b-4 border-double border-emerald-500/80">
                  <td className="py-2.5 px-4 border-r border-slate-800/60">f) Cash &amp; Cash Equivalents at end of period (D+E)</td>
                  {prevData.map(d => <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">{fmtCell(d.cfCashEnding)}</td>)}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 4. STATEMENT OF CHANGES IN EQUITY */}
      {(subTab === 'all' || subTab === 'eq') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-400 uppercase tracking-wide">
              STATEMENT OF CHANGES IN EQUITY
            </div>
            <div className="text-xs font-semibold text-slate-300">
              For the year ended 31 December {maxYear}
            </div>
            <div className="text-[11px] text-slate-400 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-700/80 min-w-[290px]">Particulars</th>
                  {prevData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-700/60 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/70 tabular-nums">
                {([
                  ['Opening Share Capital', 'eqOpeningShareCapital'],
                  ['Share Capital Issued', 'eqShareCapitalIssued'],
                  ['Closing Share Capital', 'eqClosingShareCapital'],
                  ['Opening Retained Earnings', 'eqOpeningRetainedEarnings'],
                  ['Net Profit for the Year', 'eqNetProfitForYear'],
                  ['Dividends Paid', 'eqDividendsPaid'],
                  ['Closing Retained Earnings', 'eqClosingRetainedEarnings'],
                  ['Closing Reserves', 'eqClosingReserves'],
                ] as const).map(([label, key]) => (
                  <tr key={key} className="hover:bg-slate-800/20">
                    <td className="py-2 px-4 text-slate-200 border-r border-slate-800/60">{label}</td>
                    {prevData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-200 border-r border-slate-800/40 last:border-r-0">
                        {key === 'eqDividendsPaid' ? fmtCell(-d[key]) : fmtCell(d[key])}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-emerald-950/20 font-bold text-emerald-400 border-t border-b-4 border-double border-emerald-500/80">
                  <td className="py-2.5 px-4 border-r border-slate-800/60">TOTAL CLOSING EQUITY</td>
                  {prevData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-800/40 last:border-r-0">
                      {fmtCell(d.eqTotalClosingEquity)}
                    </td>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 5. 1.00 FIXED ASSETS SCHEDULE & NOTES TO THE FINANCIAL STATEMENTS */}
      {(subTab === 'all' || subTab === 'fa') && (
        <div className="space-y-5">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
              <h3 className="text-base font-extrabold text-white tracking-wide uppercase">
                {companyProfile.companyName}
              </h3>
              <div className="text-xs font-semibold text-slate-300">
                For the Year ended 31 December {maxYear}
              </div>
              <div className="text-left text-xs font-bold text-emerald-400 pt-2">
                1.00 FIXED ASSETS SCHEDULE
              </div>
            </div>

            <div className="p-4 space-y-4 overflow-x-auto">
              {faAsc.map(d => (
                <table key={d.year} className="w-full text-xs border border-slate-700 border-collapse tabular-nums">
                  <thead>
                    <tr className="bg-slate-800/90 text-slate-100 border-b border-slate-700">
                      <th className="py-2 px-3 text-left border-r border-slate-700">Particulars</th>
                      <th className="py-2 px-3 text-right border-r border-slate-700">Opening at 01.01.{d.year}</th>
                      <th className="py-2 px-3 text-right border-r border-slate-700">Addition during the year</th>
                      <th className="py-2 px-3 text-right border-r border-slate-700">Total at 31.12.{d.year}</th>
                      <th className="py-2 px-3 text-center border-r border-slate-700">Rate %</th>
                      <th className="py-2 px-3 text-right border-r border-slate-700">Depreciation charged for the year</th>
                      <th className="py-2 px-3 text-right">W.D.V. at 31.12.{d.year}</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="border-b border-slate-800">
                      <td className="py-2 px-3 text-slate-200 border-r border-slate-800">Fixed Assets</td>
                      <td className="py-2 px-3 text-right font-mono text-slate-200 border-r border-slate-800">{fmtCell(d.faOpening)}</td>
                      <td className="py-2 px-3 text-right font-mono text-slate-200 border-r border-slate-800">{fmtCell(d.faAddition - d.faDisposal)}</td>
                      <td className="py-2 px-3 text-right font-mono text-emerald-300 font-semibold border-r border-slate-800">
                        {onUpdatePrimaryDriver ? (
                          <InlineDriverCell
                            value={d.faGrossAssetBase}
                            isEditMode={isEditMode}
                            onSave={val => onUpdatePrimaryDriver(d.year, 'fixedAssetsGross', val)}
                            title="Fixed User Input: Fixed Assets (Gross Asset Base)"
                          />
                        ) : fmtCell(d.faTotal)}
                      </td>
                      <td className="py-2 px-3 text-center font-mono text-slate-300 border-r border-slate-800">{d.faRatePct}</td>
                      <td className="py-2 px-3 text-right font-mono text-slate-200 border-r border-slate-800">{fmtCell(d.faDepreciationCharged)}</td>
                      <td className="py-2 px-3 text-right font-mono text-white font-semibold">{fmtCell(d.faWDV)}</td>
                    </tr>
                    <tr className="bg-slate-800/40 font-bold text-emerald-400">
                      <td className="py-2 px-3 border-r border-slate-800">Taka</td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.faOpening)}</td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.faAddition - d.faDisposal)}</td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.faTotal)}</td>
                      <td className="py-2 px-3 border-r border-slate-800"></td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.faDepreciationCharged)}</td>
                      <td className="py-2 px-3 text-right font-mono">{fmtCell(d.faWDV)}</td>
                    </tr>
                  </tbody>
                </table>
              ))}
            </div>
          </div>

          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
            <h4 className="text-sm font-bold text-white border-b border-slate-800 pb-2">
              NOTES TO THE FINANCIAL STATEMENTS:
            </h4>
            <div className="space-y-3 text-xs">
              {notes.map(note => (
                <div key={note.noteNumber} className="space-y-1">
                  <div className="font-bold text-emerald-400">
                    Note {note.noteNumber}: {note.title}
                  </div>
                  <p className="text-slate-300 leading-relaxed">{note.textDescription}</p>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// ============================================================================
// PROJECTED FINANCIALS EXACT-FORMAT SUITE VIEW
// Matches Projected_Financials_a13a1940.pdf line-for-line
// ============================================================================

interface ProjectedFormatSuiteViewProps {
  companyProfile: CompanyProfile;
  projData: ProjectedFormatYearData[];
  isEditMode: boolean;
  onUpdatePrimaryDriver?: (year: number, driver: keyof YearDataInput, val: number) => void;
}

export const ProjectedFormatSuiteView: React.FC<ProjectedFormatSuiteViewProps> = ({
  companyProfile,
  projData,
  isEditMode,
  onUpdatePrimaryDriver,
}) => {
  const [subTab, setSubTab] = useState<'all' | 'bs' | 'pl' | 'cf' | 'notes' | 'sched'>('all');

  if (projData.length === 0) {
    return (
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-8 text-center text-xs text-slate-400">
        Select at least one projected financial year to render the Projected Financials statements.
      </div>
    );
  }

  const maxYear = projData[0].year; // Descending order
  const minYear = projData[projData.length - 1].year;
  const amtLabel = companyProfile.reportingCurrency === '৳' ? 'Amount In Taka' : `Amount In ${companyProfile.reportingCurrency}`;
  const projAsc = [...projData].sort((a, b) => a.year - b.year);

  const subTabs = [
    { id: 'all', label: 'All Projected Statements', icon: BookOpen },
    { id: 'bs', label: 'Financial Position', icon: Landmark },
    { id: 'pl', label: 'Profit or Loss', icon: TrendingUp },
    { id: 'cf', label: 'Cash Flows', icon: Scale },
    { id: 'notes', label: 'Notes (1 - 13.00)', icon: FileText },
    { id: 'sched', label: 'Schedule-A PPE', icon: Building2 },
  ] as const;

  return (
    <div className="space-y-5">
      {/* Control & Export Strip */}
      <div className="bg-slate-900 border border-sky-900/50 rounded-2xl p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div className="space-y-1">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-[10px] font-mono font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-sky-950 text-sky-300 border border-sky-700/60">
              Exact Format: Projected Financials
            </span>
            <span className="text-xs font-bold text-white">
              {projData.length} Projected Years Selected ({minYear}–{maxYear})
            </span>
          </div>
          <p className="text-[11px] text-slate-400">
            Reproduces exact layout, Notes column, two-row Amount In Taka / Year headers, Proprietor signoff, Notes 1–13.00, and Schedule-A PPE.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => exportProjectedFinancialsExcel(companyProfile, projData)}
            className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            Projected Excel (.xlsx)
          </button>
          <button
            type="button"
            onClick={() => exportProjectedFinancialsWord(companyProfile, projData)}
            className="px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <FileText className="w-3.5 h-3.5" />
            Projected Word (.docx)
          </button>
          <button
            type="button"
            onClick={() => exportProjectedFinancialsPdf(companyProfile, projData)}
            className="px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            Projected PDF (.pdf)
          </button>
        </div>
      </div>

      {/* Sub-Tabs */}
      <div className="flex flex-wrap gap-1.5">
        {subTabs.map(t => {
          const Icon = t.icon;
          return (
            <button
              key={t.id}
              type="button"
              onClick={() => setSubTab(t.id)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors border ${
                subTab === t.id
                  ? 'bg-sky-600 text-white border-sky-400/50'
                  : 'bg-slate-900 text-slate-300 border-slate-800 hover:bg-slate-800'
              }`}
            >
              <Icon className="w-3.5 h-3.5" />
              {t.label}
            </button>
          );
        })}
      </div>

      {/* 1. STATEMENT OF FINANCIAL POSITION */}
      {(subTab === 'all' || subTab === 'bs') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-sm font-bold text-sky-400">Statement of Financial Position</div>
            <div className="text-xs text-slate-300">As at 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th rowSpan={2} className="py-2 px-4 text-center font-bold border-r border-slate-700 min-w-[270px]">Particulars</th>
                  <th rowSpan={2} className="py-2 px-3 text-center font-bold border-r border-slate-700 w-24">Notes</th>
                  {projData.map(d => (
                    <th key={`amt-${d.year}`} className="py-1.5 px-3 text-center font-semibold border-r border-b border-slate-700/80 last:border-r-0 min-w-[125px] text-[11px]">
                      {amtLabel}
                    </th>
                  ))}
                </tr>
                <tr className="bg-slate-800/80 border-b border-slate-700 text-sky-300 font-mono">
                  {projData.map(d => (
                    <th key={`yr-${d.year}`} className="py-1.5 px-3 text-center font-bold border-r border-slate-700/80 last:border-r-0">
                      {d.year}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60">
                <tr className="bg-slate-800/40 font-bold text-white underline">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Assets</td>
                </tr>
                <tr className="font-semibold text-slate-200">
                  <td colSpan={projData.length + 2} className="py-1.5 px-4">Non Current Assets</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Property Plant &amp; Equiptments</td>
                  <td className="py-2 px-3 text-center font-mono text-sky-300">Schedule-A</td>
                  {projData.map(d => (
                    <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.propertyPlantEquipments)}</td>
                  ))}
                </tr>
                <tr className="bg-slate-800/25 font-bold text-white border-t border-b-2 border-slate-600">
                  <td className="py-1.5 px-4"></td>
                  <td></td>
                  {projData.map(d => (
                    <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.propertyPlantEquipments)}</td>
                  ))}
                </tr>

                <tr className="font-semibold text-slate-200">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Current Assets</td>
                </tr>
                <tr className="bg-emerald-950/10">
                  <td className="py-2 px-4 text-emerald-300 flex items-center gap-1.5">
                    <span>Inventories</span>
                    <Lock className="w-3 h-3 text-emerald-400" />
                  </td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">3.00</td>
                  {projData.map(d => (
                    <td key={d.year} className="py-2 px-3 text-right font-mono text-emerald-300 font-semibold">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.inventories}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'closingInventory', val)}
                          title="Fixed User Input: Closing Inventory"
                        />
                      ) : fmtCell(d.inventories)}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Trade &amp; Other Receivables</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">4.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.tradeAndOtherReceivables)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Advance, Deposit &amp; Prepayments</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">5.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.advanceDepositPrepayments)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Cash &amp; Cash Equivalents</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">6.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.cashAndCashEquivalents)}</td>)}
                </tr>
                <tr className="bg-slate-800/25 font-bold text-white border-t border-slate-600">
                  <td className="py-1.5 px-4"></td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.totalCurrentAssets)}</td>)}
                </tr>
                <tr className="bg-sky-950/25 font-bold text-sky-300 border-t border-b-4 border-double border-sky-500/70">
                  <td className="py-2.5 px-4 text-right">Total Assets</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2.5 px-3 text-right font-mono">{fmtCell(d.totalAssets)}</td>)}
                </tr>

                <tr className="bg-slate-800/40 font-bold text-white underline">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Equity &amp; Liabilities</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Capital</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.capital)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Additional Capital</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.additionalCapital)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Profit during the year</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-emerald-300">{fmtCell(d.profitDuringYear)}</td>)}
                </tr>
                <tr className="bg-slate-800/25 font-bold text-white border-t border-slate-600">
                  <td className="py-1.5 px-4"></td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.equitySubtotal)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Drawings</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-rose-300">{fmtCell(d.drawings)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-white border-t border-b-2 border-slate-600">
                  <td className="py-1.5 px-4"></td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.closingCapital)}</td>)}
                </tr>

                <tr className="font-semibold text-slate-200">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Non-Current Liabilities</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Long Term Loan</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">7.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.longTermLoan)}</td>)}
                </tr>

                <tr className="font-semibold text-slate-200">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Current Liabilities</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Trade &amp; Other Payables</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">8.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.tradeAndOtherPayables)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Provisions For Tax</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">9.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.provisionsForTax)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Short Term Loan</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">10.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.shortTermLoan)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Liabilities for exp.</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.liabilitiesForExp)}</td>)}
                </tr>
                <tr className="bg-slate-800/25 font-bold text-white border-t border-slate-600">
                  <td className="py-1.5 px-4"></td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.totalCurrentLiabilities)}</td>)}
                </tr>
                <tr className="bg-sky-950/25 font-bold text-sky-300 border-t border-b-4 border-double border-sky-500/70">
                  <td className="py-2.5 px-4 text-right">Total Equity &amp; Liabilities</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2.5 px-3 text-right font-mono">{fmtCell(d.totalEquityAndLiabilities)}</td>)}
                </tr>
              </tbody>
            </table>
          </div>

          <div className="px-6 py-4 bg-slate-950/60 border-t border-slate-800 flex flex-col sm:flex-row items-start sm:items-end justify-between gap-4 text-xs text-slate-300">
            <div>
              <div className="text-slate-500">...............................</div>
              <div className="font-bold text-white">Proprietor</div>
            </div>
            <div className="italic text-slate-400">
              Signed in terms of our separate report of even date.
            </div>
          </div>
        </div>
      )}

      {/* 2. STATEMENT OF PROFIT OR LOSS & OTHER COMPREHENSIVE INCOME */}
      {(subTab === 'all' || subTab === 'pl') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-sm font-bold text-sky-400">Statement of Profit or Loss &amp; Other Comprehensive Income</div>
            <div className="text-xs text-slate-300">For the year ended 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th rowSpan={2} className="py-2 px-4 text-center font-bold border-r border-slate-700 min-w-[270px]">Particulars</th>
                  <th rowSpan={2} className="py-2 px-3 text-center font-bold border-r border-slate-700 w-24">Notes</th>
                  {projData.map(d => (
                    <th key={`pl-amt-${d.year}`} className="py-1.5 px-3 text-center font-semibold border-r border-b border-slate-700/80 last:border-r-0 min-w-[125px] text-[11px]">
                      {amtLabel}
                    </th>
                  ))}
                </tr>
                <tr className="bg-slate-800/80 border-b border-slate-700 text-sky-300 font-mono">
                  {projData.map(d => (
                    <th key={`pl-yr-${d.year}`} className="py-1.5 px-3 text-center font-bold border-r border-slate-700/80 last:border-r-0">
                      {d.year}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60">
                <tr className="bg-emerald-950/10 font-semibold text-emerald-300">
                  <td className="py-2 px-4 flex items-center gap-1.5">
                    <span>Sales</span>
                    <Lock className="w-3 h-3 text-emerald-400" />
                  </td>
                  <td></td>
                  {projData.map(d => (
                    <td key={d.year} className="py-2 px-3 text-right font-mono">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.sales}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'salesRevenue', val)}
                          title="Fixed User Input: Sales Revenue"
                        />
                      ) : fmtCell(d.sales)}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Less: Cost of Goods Sold</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">11.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.costOfGoodsSold)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-white border-t border-slate-600">
                  <td className="py-2 px-4">Gross Profit</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.grossProfit)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Add; Others Income</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.othersIncome)}</td>)}
                </tr>
                <tr className="bg-slate-800/25 font-bold text-white border-t border-slate-600">
                  <td className="py-1.5 px-4"></td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.grossProfitWithOtherIncome)}</td>)}
                </tr>
                <tr className="font-semibold text-slate-200">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Less: Operating Expenses:</td>
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Administrative &amp; Selling Expenses</td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">12.00</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.adminAndSellingExpenses)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-white border-t border-slate-600">
                  <td className="py-2 px-4">Total Operating Profit/(Loss)</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.totalOperatingProfit)}</td>)}
                </tr>
                <tr className="bg-amber-950/10">
                  <td className="py-2 px-4 text-amber-300 flex items-center gap-1.5">
                    <span>Less: Financial Expenses</span>
                    <Lock className="w-3 h-3 text-amber-400" />
                  </td>
                  <td className="py-2 px-3 text-center font-mono text-slate-400">13.00</td>
                  {projData.map(d => (
                    <td key={d.year} className="py-2 px-3 text-right font-mono text-amber-300">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.financialExpenses}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'financeCost', val)}
                          title="Fixed User Input: Finance Cost"
                        />
                      ) : fmtCell(d.financialExpenses)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-800/25 font-bold text-white border-t border-slate-600">
                  <td className="py-2 px-4">Net Profit/(Loss) Before Income Tax</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.netProfitBeforeTax)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Less: Provision For Income Tax</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200">{fmtCell(d.provisionForIncomeTax)}</td>)}
                </tr>
                <tr className="bg-sky-950/25 font-bold text-sky-300 border-t border-b-4 border-double border-sky-500/70">
                  <td className="py-2.5 px-4 flex items-center gap-1.5">
                    <span>Net Profit/(Loss) After Income Tax</span>
                    <Lock className="w-3 h-3 text-emerald-400" />
                  </td>
                  <td></td>
                  {projData.map(d => (
                    <td key={d.year} className="py-2.5 px-3 text-right font-mono">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.netProfitAfterTax}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'netProfit', val)}
                          title="Fixed User Input: Net Profit After Tax"
                        />
                      ) : fmtCell(d.netProfitAfterTax)}
                    </td>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>

          <div className="px-6 py-4 bg-slate-950/60 border-t border-slate-800 flex flex-col sm:flex-row items-start sm:items-end justify-between gap-4 text-xs text-slate-300">
            <div>
              <div className="text-slate-500">...............................</div>
              <div className="font-bold text-white">Proprietor</div>
            </div>
            <div className="italic text-slate-400">
              Signed in terms of our separate report of even date.
            </div>
          </div>
        </div>
      )}

      {/* 3. STATEMENT OF CASH FLOWS */}
      {(subTab === 'all' || subTab === 'cf') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-[11px] text-slate-400">{companyProfile.registeredAddress}</div>
            <div className="text-sm font-bold text-sky-400">Statement Of Cash Flows</div>
            <div className="text-xs text-slate-300">For the year ended 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th className="py-2.5 px-4 text-left font-bold min-w-[330px]">1. CASH FLOW FROM OPERATING ACTIVITIES:</th>
                  {projData.map(d => (
                    <th key={d.year} className="py-2.5 px-3 text-center font-mono font-bold border-l border-slate-700/70 min-w-[120px]">
                      {d.year}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60">
                {([
                  ['Net Income', 'cfNetIncome'],
                  ['(Increase)/Decrease In Inventory', 'cfChangeInventory'],
                  ['(Increase)/Decrease In Trade & Other Receivables', 'cfChangeReceivables'],
                  ['(Increase)/Decrease in Advance, Deposit & Prepayments', 'cfChangeAdvanceDeposit'],
                  ['Increase/(Decrease) in Current Assets', 'cfChangeCurrentAssets'],
                  ['Increase/(Decrease) in Current Liabilities', 'cfChangeCurrentLiabilities'],
                  ['Increase/(Decrease) in Trade & Other Payables', 'cfChangeTradePayables'],
                  ['Increase/(Decrease) In Provisions For Tax', 'cfChangeProvisionsTax'],
                  ['Increase/(Decrease) Liabilities For Expense', 'cfChangeLiabilitiesExp'],
                  ['Depreciation', 'cfDepreciation'],
                ] as const).map(([label, key]) => (
                  <tr key={key}>
                    <td className="py-1.5 px-4 text-slate-300">{label}</td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d[key])}</td>)}
                  </tr>
                ))}
                <tr className="bg-slate-800/35 font-bold text-emerald-400">
                  <td className="py-2 px-4">Net Cash Used by Operating Activities</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-800/50">{fmtCell(d.cfNetCashOperating)}</td>)}
                </tr>

                <tr className="bg-slate-800/40 font-bold text-sky-300">
                  <td colSpan={projData.length + 1} className="py-2 px-4 underline">2. CASH FLOW FROM INVESTING ACTIVITIES:</td>
                </tr>
                <tr className="font-semibold text-slate-300">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Inflows:</td>
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Sale Of Property, Plant &amp; Equipment</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfSalePPE)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Capital</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfInvestingCapitalInflow)}</td>)}
                </tr>
                <tr className="font-semibold text-slate-300">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Outflows:</td>
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Preliminary Expenses</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfPreliminaryExpenses)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Purchase Property, Plant &amp; Equipment</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfPurchasePPE)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Purchase Investment In Debt Or Securities</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfPurchaseInvestment)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-sky-400">
                  <td className="py-2 px-4">Net Cash Used by Investing Activities</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-800/50">{fmtCell(d.cfNetCashInvesting)}</td>)}
                </tr>

                <tr className="bg-slate-800/40 font-bold text-amber-300">
                  <td colSpan={projData.length + 1} className="py-2 px-4 underline">3. CASH FLOW FROM FINANCIAL ACTIVITIES:</td>
                </tr>
                <tr className="font-semibold text-slate-300">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Inflows:</td>
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Loan Received</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfLoanReceived)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Additional Capital</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfAdditionalCapital)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Issuance Of Debt (Bond &amp; Notes)</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfIssuanceOfDebt)}</td>)}
                </tr>
                <tr className="font-semibold text-slate-300">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Outflows:</td>
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">To Stockholders as Dividend</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfDividendsToStockholders)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Drawings / Dividend</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfDrawingsOutflow)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">To Redeem Long Term Debt (Treasury Stock)</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfRedeemLongTermDebt)}</td>)}
                </tr>
                <tr>
                  <td className="py-1.5 px-4 text-slate-300">Issuance Of Bonds Payable</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfIssuanceBondsPayable)}</td>)}
                </tr>
                <tr className="bg-slate-800/35 font-bold text-amber-400">
                  <td className="py-2 px-4">Net Cash Provided by Financial Activities</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-800/50">{fmtCell(d.cfNetCashFinancing)}</td>)}
                </tr>

                <tr className="font-bold text-white">
                  <td className="py-2 px-4">Net Cash Increase</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-800/50">{fmtCell(d.cfNetCashIncrease)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-300">Opening Cash &amp; Bank Balance</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-200 border-l border-slate-800/50">{fmtCell(d.cfOpeningCash)}</td>)}
                </tr>
                <tr className="bg-sky-950/25 font-bold text-sky-300 border-t border-b-4 border-double border-sky-500/70">
                  <td className="py-2.5 px-4">Closing Cash &amp; Bank Balance</td>
                  {projData.map(d => <td key={d.year} className="py-2.5 px-3 text-right font-mono border-l border-slate-800/50">{fmtCell(d.cfClosingCash)}</td>)}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 4. NOTES TO THE FINANCIAL STATEMENTS (1, 2.1-2.5, 3.00-13.00) */}
      {(subTab === 'all' || subTab === 'notes') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-[11px] text-slate-400">{companyProfile.registeredAddress}</div>
            <div className="text-sm font-bold text-sky-400">Notes to the Financial Statements</div>
            <div className="text-xs text-slate-300">For the year ended 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="p-5 space-y-4 text-xs border-b border-slate-800">
            <div>
              <div className="font-bold text-white">1 Reporting Entity:</div>
              <p className="text-slate-300 mt-1">
                {companyProfile.companyName} is a {companyProfile.entityType.replace(/_/g, ' ')} entity and its office is situated at {companyProfile.registeredAddress}.
              </p>
            </div>
            <div className="space-y-2">
              <div className="font-bold text-white">2 SIGNIFICANT ACCOUNTING POLICIES:</div>
              <p className="text-slate-300"><strong>2.1 Basis of Accounting:</strong> The Financial Statements have been prepared on a going concern basis using accrual method of accounting in accordance with {companyProfile.accountingFramework}.</p>
              <p className="text-slate-300"><strong>2.2 Accounting Policies:</strong> Selected and applied consistently for significant transactions and events to all periods presented.</p>
              <p className="text-slate-300"><strong>2.3 Property, Plant and Equipment:</strong> Stated at cost less accumulated depreciation. Depreciation is computed over estimated useful lives consistently.</p>
              <p className="text-slate-300"><strong>2.4 Reporting Period:</strong> Financial Statements cover the annual reporting cycle ending {companyProfile.financialYearEnd}.</p>
              <p className="text-slate-300"><strong>2.5 Reporting currency and level of precision:</strong> Figures represent {companyProfile.reportingCurrency} ({companyProfile.reportingUnit}) rounded cleanly.</p>
              <p className="text-slate-300"><strong>2.6 Comparative Information:</strong> Disclosed for all selected years for numerical information and descriptive narratives.</p>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-800/90 border-b border-slate-700 text-slate-100">
                  <th rowSpan={2} className="py-2 px-3 text-center font-bold border-r border-slate-700 w-20">Note No.</th>
                  <th rowSpan={2} className="py-2 px-4 text-left font-bold border-r border-slate-700 min-w-[260px]">Particulars</th>
                  {projData.map(d => (
                    <th key={`nt-amt-${d.year}`} className="py-1.5 px-3 text-center font-semibold border-r border-b border-slate-700/80 last:border-r-0 min-w-[120px] text-[11px]">
                      {amtLabel}
                    </th>
                  ))}
                </tr>
                <tr className="bg-slate-800/80 border-b border-slate-700 text-sky-300 font-mono">
                  {projData.map(d => (
                    <th key={`nt-yr-${d.year}`} className="py-1.5 px-3 text-center font-bold border-r border-slate-700/80 last:border-r-0">{d.year}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60">
                <tr className="font-bold text-white bg-slate-800/30">
                  <td className="py-2 px-3 text-center font-mono">3.00</td>
                  <td className="py-2 px-4">Inventories</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.inventories)}</td>)}
                </tr>
                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">4.00</td>
                  <td colSpan={projData.length + 1} className="py-2 px-4">Trade &amp; Other Receivables</td>
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">As Per Last Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note4ReceivablesOpening)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Add: Credit Sales During the Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note4ReceivablesAdded)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Less: Received During the Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note4ReceivablesReceived)}</td>)}
                </tr>
                <tr className="font-bold text-white bg-slate-800/30">
                  <td></td><td className="py-1.5 px-4 pl-6">Closing Trade &amp; Other Receivables</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.tradeAndOtherReceivables)}</td>)}
                </tr>

                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">5.00</td>
                  <td colSpan={projData.length + 1} className="py-2 px-4">Advance, Deposit &amp; Prepayments</td>
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">As Per Last Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note5AdvanceOpening)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Add: During the Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note5AdvanceAdded)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Less: Adjustment</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note5AdvanceAdjusted)}</td>)}
                </tr>
                <tr className="font-bold text-white bg-slate-800/30">
                  <td></td><td className="py-1.5 px-4 pl-6">Closing Advance, Deposit &amp; Prepayments</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.advanceDepositPrepayments)}</td>)}
                </tr>

                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">6.00</td>
                  <td className="py-2 px-4">Cash &amp; Equivalent — Cash In Hand &amp; Bank</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-white">{fmtCell(d.cashAndCashEquivalents)}</td>)}
                </tr>
                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">7.00</td>
                  <td className="py-2 px-4">Long Term Loan</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-white">{fmtCell(d.longTermLoan)}</td>)}
                </tr>

                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">8.00</td>
                  <td colSpan={projData.length + 1} className="py-2 px-4">Trade &amp; Other Payables</td>
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">As Per Last Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note8PayablesOpening)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Add: During the Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note8PayablesAdded)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Less: Payment During the Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note8PayablesPaid)}</td>)}
                </tr>
                <tr className="font-bold text-white bg-slate-800/30">
                  <td></td><td className="py-1.5 px-4 pl-6">Closing Trade &amp; Other Payables</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.tradeAndOtherPayables)}</td>)}
                </tr>

                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">9.00</td>
                  <td colSpan={projData.length + 1} className="py-2 px-4">Provisions For Tax</td>
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">As Per Last Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note9TaxOpening)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Add: During the Year</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note9TaxAdded)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Less: Adjustment</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note9TaxAdjusted)}</td>)}
                </tr>
                <tr className="font-bold text-white bg-slate-800/30">
                  <td></td><td className="py-1.5 px-4 pl-6">Closing Provisions For Tax</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.provisionsForTax)}</td>)}
                </tr>

                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">10.00</td>
                  <td className="py-2 px-4">Short Term Loan — Bank Loan</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-white">{fmtCell(d.shortTermLoan)}</td>)}
                </tr>

                <tr className="font-bold text-sky-300 bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">11.00</td>
                  <td colSpan={projData.length + 1} className="py-2 px-4">Cost of Goods Sold</td>
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Opening Stock</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note11OpeningStock)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Add: Purchase</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note11Purchase)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Less: Closing Stock</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note11ClosingStock)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Labour &amp; Wages</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note11LabourWages)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Electricity Bill</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note11ElectricityBill)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Carriage Inward</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note11CarriageInward)}</td>)}
                </tr>
                <tr>
                  <td></td><td className="py-1.5 px-4 pl-6 text-slate-400">Other Direct Expenses</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-300">{fmtCell(d.note11OtherDirectExpenses)}</td>)}
                </tr>
                <tr className="font-bold text-white bg-slate-800/30">
                  <td></td><td className="py-1.5 px-4 pl-6">Cost of Goods Sold</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.costOfGoodsSold)}</td>)}
                </tr>

                <tr className="font-bold text-white bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">12.00</td>
                  <td className="py-2 px-4">Administrative &amp; Selling Expenses</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.adminAndSellingExpenses)}</td>)}
                </tr>
                <tr className="font-bold text-white bg-slate-800/20">
                  <td className="py-2 px-3 text-center font-mono">13.00</td>
                  <td className="py-2 px-4">Financial Expenses — Bank Charges &amp; Interest</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.financialExpenses)}</td>)}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 5. SCHEDULE OF PROPERTY PLANT & EQUIPMENTS (Schedule-A) */}
      {(subTab === 'all' || subTab === 'sched') && (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-800 text-center space-y-1 bg-slate-900/95">
            <h3 className="text-base font-extrabold text-white tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-[11px] text-slate-400">{companyProfile.registeredAddress}</div>
            <div className="text-sm font-bold text-sky-400">Schedule of Property Plant &amp; Equipments</div>
            <div className="text-xs text-slate-300">As at 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="p-4 space-y-4 overflow-x-auto">
            {projAsc.map(d => (
              <table key={d.year} className="w-full text-xs border border-slate-700 border-collapse tabular-nums">
                <thead>
                  <tr className="bg-slate-800/90 text-slate-100 border-b border-slate-700">
                    <th rowSpan={2} className="py-2 px-2 text-center border-r border-slate-700 w-14">SL No.</th>
                    <th rowSpan={2} className="py-2 px-3 text-left border-r border-slate-700 min-w-[200px]">Particulars</th>
                    <th colSpan={4} className="py-1.5 px-3 text-center border-r border-b border-slate-700">Cost</th>
                    <th rowSpan={2} className="py-2 px-2 text-center border-r border-slate-700 w-24">Rate of Dep. (%)</th>
                    <th rowSpan={2} className="py-2 px-3 text-right border-r border-slate-700">Depreciation Charged During the Year</th>
                    <th rowSpan={2} className="py-2 px-3 text-right">W.D.V as on 31.12.{d.year}</th>
                  </tr>
                  <tr className="bg-slate-800/75 text-slate-300 border-b border-slate-700 text-[11px]">
                    <th className="py-1.5 px-3 text-right border-r border-slate-700">Opening Balance</th>
                    <th className="py-1.5 px-3 text-right border-r border-slate-700">Addition During the Year</th>
                    <th className="py-1.5 px-3 text-right border-r border-slate-700">Disposals During the Year</th>
                    <th className="py-1.5 px-3 text-right border-r border-slate-700">Balance as on 31.12.{d.year}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr className="border-b border-slate-800">
                    <td className="py-2 px-2 text-center font-mono text-slate-300 border-r border-slate-800">1</td>
                    <td className="py-2 px-3 text-slate-200 border-r border-slate-800">Property Plant &amp; Equipment</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-200 border-r border-slate-800">{fmtCell(d.schedOpeningBalance)}</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-200 border-r border-slate-800">{fmtCell(d.schedAddition)}</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-200 border-r border-slate-800">{fmtCell(d.schedDisposals)}</td>
                    <td className="py-2 px-3 text-right font-mono text-emerald-300 font-semibold border-r border-slate-800">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.schedGrossAssetBase}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'fixedAssetsGross', val)}
                          title="Fixed User Input: Fixed Assets (Gross Asset Base)"
                        />
                      ) : fmtCell(d.schedBalanceClosing)}
                    </td>
                    <td className="py-2 px-2 text-center font-mono text-slate-300 border-r border-slate-800">{d.schedRatePct}%</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-200 border-r border-slate-800">{fmtCell(d.schedDepreciation)}</td>
                    <td className="py-2 px-3 text-right font-mono text-white font-semibold">{fmtCell(d.schedWDV)}</td>
                  </tr>
                  <tr className="bg-slate-800/40 font-bold text-sky-300">
                    <td className="py-2 px-2 border-r border-slate-800"></td>
                    <td className="py-2 px-3 text-center border-r border-slate-800">Total</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.schedOpeningBalance)}</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.schedAddition)}</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.schedDisposals)}</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.schedBalanceClosing)}</td>
                    <td className="py-2 px-2 border-r border-slate-800"></td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-800">{fmtCell(d.schedDepreciation)}</td>
                    <td className="py-2 px-3 text-right font-mono">{fmtCell(d.schedWDV)}</td>
                  </tr>
                </tbody>
              </table>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
