import React, { useState, useMemo, useEffect } from 'react';
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
  CheckSquare,
  Square,
  RotateCcw,
} from 'lucide-react';
import {
  CompanyProfile,
  FinancialNoteItem,
  PreviousFormatYearData,
  PreviousComparativePair,
  ProjectedFormatYearData,
  StatementConfigItem,
  StatementLabelOverrides,
  YearDataInput,
} from '../types/financial';
import { resolveEffectiveLabel, getCustomItemYearValue, buildProjectedQuantNoteRows } from '../services/financialEngine';
import {
  exportPreviousFinancialsExcel,
  exportPreviousFinancialsWord,
  exportPreviousFinancialsPdf,
  exportAllPreviousComparativePairsExcel,
  exportAllPreviousComparativePairsWord,
  exportAllPreviousComparativePairsPdf,
  exportProjectedFinancialsExcel,
  exportProjectedFinancialsWord,
  exportProjectedFinancialsPdf,
} from '../services/exportService';

function fmtCell(val: number): string {
  if (Math.abs(val) < 0.005) return '-';
  const absVal = Math.abs(val);
  const hasCents = Math.abs(absVal - Math.round(absVal)) >= 0.005;
  const formatted = hasCents
    ? absVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : Math.round(absVal).toLocaleString('en-US');
  return val < 0 ? `(${formatted})` : formatted;
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

  useEffect(() => {
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
        className="w-24 bg-white border border-emerald-600 rounded px-1.5 py-0.5 text-right font-mono text-xs text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500 shadow-xs"
      />
    );
  }

  return (
    <span
      onClick={() => setEditing(true)}
      title={title || 'Click to edit fixed primary user input'}
      className="cursor-pointer hover:text-emerald-700 hover:underline font-mono tabular-nums"
    >
      {fmtCell(value)}
    </span>
  );
};

// ============================================================================
// PREVIOUS FINANCIALS EXACT-FORMAT SUITE VIEW
// Matches Previous_Financials_aec5c4ae.pdf & splits into overlapping 2-year comparative sets
// ============================================================================

interface PreviousFormatSuiteViewProps {
  companyProfile: CompanyProfile;
  prevData: PreviousFormatYearData[];
  comparativePairs?: PreviousComparativePair[];
  getPairScopedNotes?: (pairData: PreviousFormatYearData[]) => FinancialNoteItem[];
  notes: FinancialNoteItem[];
  isEditMode: boolean;
  onUpdatePrimaryDriver?: (year: number, driver: keyof YearDataInput, val: number) => void;
  statementConfig?: StatementConfigItem[];
  onUpdateStatementConfig?: (items: StatementConfigItem[]) => void;
  onResetStatementConfig?: () => void;
  labelOverrides?: StatementLabelOverrides;
}

export const PreviousFormatSuiteView: React.FC<PreviousFormatSuiteViewProps> = ({
  companyProfile,
  prevData,
  comparativePairs = [],
  getPairScopedNotes,
  notes,
  isEditMode,
  onUpdatePrimaryDriver,
  statementConfig = [],
  onUpdateStatementConfig,
  onResetStatementConfig,
  labelOverrides = {},
}) => {
  const [subTab, setSubTab] = useState<'all' | 'bs' | 'is' | 'cf' | 'eq' | 'fa'>('all');
  const [selectedPairId, setSelectedPairId] = useState<string>(() => comparativePairs[0]?.id || 'all');
  const [isCustomizeMode, setIsCustomizeMode] = useState<boolean>(false);

  const isElemVisible = (id: string): boolean => {
    const found = statementConfig.find(c => c.id === id);
    return found ? found.isVisible : true;
  };

  const toggleElem = (id: string) => {
    if (!onUpdateStatementConfig || statementConfig.length === 0) return;
    onUpdateStatementConfig(
      statementConfig.map(c => (c.id === id ? { ...c, isVisible: !c.isVisible } : c))
    );
  };

  const handleSetAllConfigVisibility = (select: boolean) => {
    if (!onUpdateStatementConfig || statementConfig.length === 0) return;
    onUpdateStatementConfig(
      statementConfig.map(c => ({ ...c, isVisible: select }))
    );
  };

  const renderCheck = (id: string) =>
    isCustomizeMode && onUpdateStatementConfig ? (
      <input
        type="checkbox"
        checked={isElemVisible(id)}
        onChange={() => toggleElem(id)}
        title="Include or exclude this element from calculation and downloads"
        className="w-3.5 h-3.5 accent-emerald-600 rounded cursor-pointer shrink-0"
      />
    ) : null;

  const lbl = (stmt: string, field: string, def: string) =>
    resolveEffectiveLabel(stmt, field, def, labelOverrides, statementConfig);

  const stdIncomeSections = ['Revenue & Gross Margin', 'Operating Expenses', 'Non-Operating & Tax'];
  const getCustom = (stmt: StatementConfigItem['statement'], section?: string) =>
    statementConfig.filter(c => c.isCustom && c.statement === stmt && (!section || c.section === section));
  const customBsNca = getCustom('balanceSheet', 'Non-Current Assets');
  const customBsCa = getCustom('balanceSheet', 'Current Assets');
  const customBsNcl = getCustom('balanceSheet', 'Non-Current Liabilities');
  const customBsCl = getCustom('balanceSheet', 'Current Liabilities');
  const customBsEq = getCustom('balanceSheet', 'Equity');
  const customIncRev = getCustom('income', 'Revenue & Gross Margin');
  const customIncOpex = getCustom('income', 'Operating Expenses');
  const customIncNonOp = statementConfig.filter(
    c => c.isCustom && c.statement === 'income' && (c.section === 'Non-Operating & Tax' || !stdIncomeSections.includes(c.section))
  );
  const visIncRev = customIncRev.filter(c => c.isVisible);
  const visIncOpex = customIncOpex.filter(c => c.isVisible);
  const visIncNonOp = customIncNonOp.filter(c => c.isVisible);
  const customCfOp = getCustom('cashFlow', 'Operating Activities');
  const customCfInv = getCustom('cashFlow', 'Investing Activities');
  const customCfFin = getCustom('cashFlow', 'Financing Activities');
  const customEq = getCustom('equity');

  const exportSyncOptions = { statementConfig, labelOverrides };

  useEffect(() => {
    if (selectedPairId !== 'all' && !comparativePairs.some(p => p.id === selectedPairId)) {
      setSelectedPairId(comparativePairs[0]?.id || 'all');
    }
  }, [comparativePairs, selectedPairId]);

  const activePair = useMemo(
    () => comparativePairs.find(p => p.id === selectedPairId) || null,
    [comparativePairs, selectedPairId]
  );

  const displayedData = useMemo(
    () => (activePair ? activePair.data : prevData),
    [activePair, prevData]
  );

  const displayedNotes = useMemo(() => {
    if (activePair && getPairScopedNotes) {
      return getPairScopedNotes(activePair.data);
    }
    return notes;
  }, [activePair, getPairScopedNotes, notes]);

  if (prevData.length === 0 || displayedData.length === 0) {
    return (
      <div className="bg-white/92 backdrop-blur-md border border-slate-200 rounded-2xl p-8 text-center text-xs text-slate-600 shadow-sm">
        Select at least one previous financial year to render the Previous Financials statements.
      </div>
    );
  }

  const maxYear = displayedData[0].year;
  const allAsc = [...prevData].sort((a, b) => a.year - b.year);
  const faAsc = [...displayedData].sort((a, b) => a.year - b.year);
  const ccyLabel = companyProfile.reportingCurrency === '৳' ? 'BDT' : companyProfile.reportingCurrency;
  const unitCaption = companyProfile.reportingUnit === 'Units'
    ? `(Figure in ${ccyLabel})`
    : `(Figure in ${ccyLabel} · ${companyProfile.reportingUnit})`;

  const subTabs = [
    { id: 'all', label: 'All Previous Statements', icon: BookOpen },
    { id: 'bs', label: 'Financial Position', icon: Landmark },
    { id: 'is', label: 'Comprehensive Income', icon: TrendingUp },
    { id: 'cf', label: 'Cash Flows', icon: Scale },
    { id: 'eq', label: 'Changes in Equity', icon: Layers },
    { id: 'fa', label: '1.00 Fixed Assets & Notes', icon: Building2 },
  ] as const;

  const resolveNotes = (pairData: PreviousFormatYearData[]) =>
    getPairScopedNotes ? getPairScopedNotes(pairData) : notes;

  return (
    <div className="space-y-5">
      {/* Control & Comparative Pair Downloads Strip */}
      <div className="bg-white/94 backdrop-blur-md border border-emerald-200 rounded-2xl p-4 sm:p-5 space-y-4 shadow-sm">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-[11px] font-mono font-bold text-emerald-800 uppercase tracking-wider">
                Exact Format: Previous Financials
              </span>
              <span className="text-slate-300" aria-hidden="true">·</span>
              <span className="text-xs font-bold text-slate-900">
                {allAsc.length} Previous Years ({allAsc[0].year}–{allAsc[allAsc.length - 1].year})
              </span>
            </div>
            <p className="text-[11px] text-slate-600">
              Automatically splits selected previous years into overlapping 2-year comparative downloads ({comparativePairs.map(p => p.label).join(', ')}). Viewing:{' '}
              <strong className="text-emerald-800 font-mono">
                {activePair ? `Comparative Statement (${activePair.label})` : `All Selected (${allAsc[0].year}–${allAsc[allAsc.length - 1].year})`}
              </strong>
            </p>
          </div>

          {/* All Comparative Pairs Batch Download Buttons + Element Selection Toggle */}
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            {onUpdateStatementConfig && statementConfig.length > 0 && (
              <button
                type="button"
                onClick={() => setIsCustomizeMode(!isCustomizeMode)}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 border transition-colors ${
                  isCustomizeMode
                    ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                    : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                }`}
              >
                <CheckSquare className="w-3.5 h-3.5" />
                <span>{isCustomizeMode ? 'Done Selecting' : 'Customize Elements'}</span>
              </button>
            )}
            {isCustomizeMode && onUpdateStatementConfig && (
              <>
                <button
                  type="button"
                  onClick={() => handleSetAllConfigVisibility(true)}
                  className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-emerald-50 text-emerald-800 border border-slate-300 hover:border-emerald-300 text-xs font-medium flex items-center gap-1"
                >
                  <CheckSquare className="w-3.5 h-3.5 text-emerald-600" />
                  Select All
                </button>
                <button
                  type="button"
                  onClick={() => handleSetAllConfigVisibility(false)}
                  className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-rose-50 text-rose-700 border border-slate-300 hover:border-rose-300 text-xs font-medium flex items-center gap-1"
                >
                  <Square className="w-3.5 h-3.5 text-rose-500" />
                  Deselect All
                </button>
              </>
            )}
            {isCustomizeMode && onResetStatementConfig && (
              <button
                type="button"
                onClick={onResetStatementConfig}
                className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-medium flex items-center gap-1"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                Reset Defaults
              </button>
            )}
            {comparativePairs.length > 0 && (
              <>
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsExcel(companyProfile, comparativePairs, resolveNotes, exportSyncOptions)}
                  className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  {comparativePairs.length > 1 ? `Download All ${comparativePairs.length} Excel (.xlsx)` : 'Previous Excel (.xlsx)'}
                </button>
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsWord(companyProfile, comparativePairs, resolveNotes, exportSyncOptions)}
                  className="px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
                >
                  <FileText className="w-3.5 h-3.5" />
                  {comparativePairs.length > 1 ? `Download All ${comparativePairs.length} Word (.docx)` : 'Previous Word (.docx)'}
                </button>
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsPdf(companyProfile, comparativePairs, resolveNotes, exportSyncOptions)}
                  className="px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
                >
                  <Download className="w-3.5 h-3.5" />
                  {comparativePairs.length > 1 ? `Download All ${comparativePairs.length} PDF (.pdf)` : 'Previous PDF (.pdf)'}
                </button>
              </>
            )}
          </div>
        </div>

        {/* Comparative Set Selector + Individual Pair Download Cards */}
        {comparativePairs.length > 0 && (
          <div className="pt-3 border-t border-slate-200 flex flex-col gap-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-[11px] font-bold text-slate-700 mr-1">Preview Comparative Set:</span>
                {comparativePairs.map(pair => (
                  <button
                    key={pair.id}
                    type="button"
                    onClick={() => setSelectedPairId(pair.id)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-mono font-bold transition-all border ${
                      selectedPairId === pair.id
                        ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                        : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                    }`}
                  >
                    Previous Financials ({pair.label})
                  </button>
                ))}
                {comparativePairs.length > 1 && (
                  <button
                    type="button"
                    onClick={() => setSelectedPairId('all')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-mono font-bold transition-all border ${
                      selectedPairId === 'all'
                        ? 'bg-slate-900 text-white border-slate-900 shadow-sm'
                        : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                    }`}
                  >
                    Combined Preview ({allAsc[0].year}–{allAsc[allAsc.length - 1].year})
                  </button>
                )}
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2.5">
              {comparativePairs.map((pair, idx) => {
                const pairNotes = resolveNotes(pair.data);
                return (
                  <div
                    key={`dl-${pair.id}`}
                    className={`rounded-xl border p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 transition-all ${
                      selectedPairId === pair.id
                        ? 'bg-emerald-50/85 border-emerald-300'
                        : 'bg-slate-50/90 border-slate-200'
                    }`}
                  >
                    <div>
                      <div className="text-xs font-bold text-slate-900 font-mono">
                        Set {idx + 1}: ({pair.label})
                      </div>
                      <div className="text-[10px] text-slate-600 font-mono">
                        Cols: 31.12.{pair.laterYear} &amp; 31.12.{pair.earlierYear}
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      <button
                        type="button"
                        onClick={() => exportPreviousFinancialsExcel(companyProfile, pair.data, pairNotes, exportSyncOptions)}
                        className="px-2.5 py-1 rounded-md bg-white hover:bg-emerald-600 text-emerald-700 hover:text-white border border-emerald-300 font-mono text-[11px] font-bold transition-colors"
                      >
                        .XLSX
                      </button>
                      <button
                        type="button"
                        onClick={() => void exportPreviousFinancialsWord(companyProfile, pair.data, pairNotes, exportSyncOptions)}
                        className="px-2.5 py-1 rounded-md bg-white hover:bg-blue-600 text-blue-700 hover:text-white border border-blue-300 font-mono text-[11px] font-bold transition-colors"
                      >
                        .DOCX
                      </button>
                      <button
                        type="button"
                        onClick={() => exportPreviousFinancialsPdf(companyProfile, pair.data, pairNotes, exportSyncOptions)}
                        className="px-2.5 py-1 rounded-md bg-white hover:bg-rose-600 text-rose-700 hover:text-white border border-rose-300 font-mono text-[11px] font-bold transition-colors"
                      >
                        .PDF
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
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
                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                  : 'bg-white/92 text-slate-700 border-slate-200 hover:bg-slate-50 hover:text-slate-900'
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
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-white">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-800 uppercase tracking-wide">
              STATEMENT OF FINANCIAL POSITION
            </div>
            <div className="text-xs font-semibold text-slate-700">
              AS AT 31 DECEMBER, {maxYear}
            </div>
            <div className="text-[11px] text-slate-500 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-300 text-slate-900">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-300 min-w-[290px]">
                    Particulars
                  </th>
                  {displayedData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-300 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 tabular-nums">
                <tr className="bg-slate-50 font-bold text-slate-900">
                  <td colSpan={displayedData.length + 1} className="py-2 px-4">PROPERTIES &amp; ASSETS</td>
                </tr>
                {(isElemVisible('bs_propertyPlantEquipment') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_propertyPlantEquipment') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 text-slate-800 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_propertyPlantEquipment')}
                        <span>{lbl('bs', 'propertyPlantEquipment', 'Fixed Assets (at cost less depreciation)')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.fixedAssetsWDV)}
                      </td>
                    ))}
                  </tr>
                )}
                {customBsNca.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-50 font-semibold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">Current Assets</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.totalCurrentAssets)}
                    </td>
                  ))}
                </tr>
                {(isElemVisible('bs_otherReceivables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_otherReceivables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 pl-7 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_otherReceivables')}
                        <span>{lbl('bs', 'otherReceivables', 'Advances, Deposits & Pre-payments')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.advancesDepositsPrepayments)}
                      </td>
                    ))}
                  </tr>
                )}
                {(isElemVisible('bs_closingInventory') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_closingInventory') ? 'opacity-45 bg-slate-50' : 'bg-emerald-50/55 hover:bg-emerald-50'}>
                    <td className="py-2 px-4 pl-7 text-emerald-900 font-medium border-r border-slate-200">
                      <div className="flex items-center gap-1.5">
                        {renderCheck('bs_closingInventory')}
                        <span>{lbl('bs', 'closingInventory', 'Stock & Stores')}</span>
                        <Lock className="w-3 h-3 text-emerald-700" />
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-emerald-800 font-semibold border-r border-slate-200 last:border-r-0">
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
                )}
                {(isElemVisible('bs_tradeReceivables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_tradeReceivables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 pl-7 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_tradeReceivables')}
                        <span>{lbl('bs', 'tradeReceivables', 'Bills Receivable')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.billsReceivable)}
                      </td>
                    ))}
                  </tr>
                )}
                {(isElemVisible('bs_cashAndEquivalents') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_cashAndEquivalents') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 pl-7 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_cashAndEquivalents')}
                        <span>{lbl('bs', 'cashAndEquivalents', 'Cash & Cash Equivalents')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.cashAndCashEquivalents)}
                      </td>
                    ))}
                  </tr>
                )}
                {customBsCa.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 pl-7 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-emerald-50 font-bold text-emerald-950 border-t border-b-2 border-emerald-700">
                  <td className="py-2.5 px-4 text-right border-r border-slate-200">Taka</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.totalAssetsTaka)}
                    </td>
                  ))}
                </tr>

                <tr className="bg-slate-50 font-bold text-slate-900">
                  <td colSpan={displayedData.length + 1} className="py-2 px-4">CAPITAL &amp; LIABILITIES</td>
                </tr>
                <tr className="hover:bg-slate-50/70">
                  <td className="py-2 px-4 text-slate-700 border-r border-slate-200">Opening Capital</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.openingCapital)}
                    </td>
                  ))}
                </tr>
                <tr className="hover:bg-slate-50/70">
                  <td className="py-2 px-4 text-slate-700 border-r border-slate-200">Add. Net Profit for the year</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono text-emerald-700 border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.netProfitForYear)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-50 font-semibold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">Subtotal</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.capitalSubtotal)}
                    </td>
                  ))}
                </tr>
                {(isElemVisible('eq_dividendsPaid') || isCustomizeMode) && (
                  <tr className={!isElemVisible('eq_dividendsPaid') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('eq_dividendsPaid')}
                        <span>Less: {lbl('eq', 'dividendsPaid', 'Drawings')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-rose-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.drawings)}
                      </td>
                    ))}
                  </tr>
                )}
                <tr className="bg-slate-100/80 font-bold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">Capital after drawings</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.capitalAfterDrawings)}
                    </td>
                  ))}
                </tr>
                {customBsEq.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                {customBsNcl.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-50 font-semibold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">Current Liabilities</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.totalCurrentLiabilities)}
                    </td>
                  ))}
                </tr>
                {(isElemVisible('bs_otherPayables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_otherPayables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 pl-7 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_otherPayables')}
                        <span>{lbl('bs', 'otherPayables', 'For Expenses')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.forExpenses)}
                      </td>
                    ))}
                  </tr>
                )}
                {(isElemVisible('bs_tradePayables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_tradePayables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 pl-7 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_tradePayables')}
                        <span>{lbl('bs', 'tradePayables', 'Accounts Payable')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.accountsPayable)}
                      </td>
                    ))}
                  </tr>
                )}
                {(isElemVisible('bs_shortTermBorrowings') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_shortTermBorrowings') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 pl-7 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_shortTermBorrowings')}
                        <span>{lbl('bs', 'shortTermBorrowings', 'Short Term Loan')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.shortTermLoan)}
                      </td>
                    ))}
                  </tr>
                )}
                {(isElemVisible('bs_currentTaxPayable') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_currentTaxPayable') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                    <td className="py-2 px-4 pl-7 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_currentTaxPayable')}
                        <span>{lbl('bs', 'currentTaxPayable', 'Provision for Taxation')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.provisionForTaxation)}
                      </td>
                    ))}
                  </tr>
                )}
                {customBsCl.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 pl-7 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-emerald-50 font-bold text-emerald-950 border-t border-b-4 border-double border-emerald-700">
                  <td className="py-2.5 px-4 text-right border-r border-slate-200">Total</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
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
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-white">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-800">
              Statement of Comprehensive Income
            </div>
            <div className="text-xs font-semibold text-slate-700">
              For the year ended 31 December, {maxYear}
            </div>
            <div className="text-[11px] text-slate-500 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-300 text-slate-900">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-300 min-w-[290px]">Particulars</th>
                  {displayedData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-300 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 tabular-nums">
                {(isElemVisible('inc_revenue') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_revenue') ? 'opacity-45 bg-slate-50' : 'bg-emerald-50/55 font-semibold text-emerald-900'}>
                    <td className="py-2 px-4 border-r border-slate-200">
                      <div className="flex items-center gap-1.5">
                        {renderCheck('inc_revenue')}
                        <span>{lbl('income', 'revenue', 'Sales')}</span>
                        <Lock className="w-3 h-3 text-emerald-700" />
                      </div>
                    </td>
                    {displayedData.map(d => {
                      const custSum = visIncRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
                      const baseSales = Math.round((d.sales - custSum) * 100) / 100;
                      return (
                        <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                          {onUpdatePrimaryDriver ? (
                            <InlineDriverCell
                              value={baseSales}
                              isEditMode={isEditMode}
                              onSave={val => onUpdatePrimaryDriver(d.year, 'salesRevenue', val)}
                              title="Fixed User Input: Sales Revenue"
                            />
                          ) : fmtCell(baseSales)}
                        </td>
                      );
                    })}
                  </tr>
                )}
                {customIncRev.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 pl-7 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                {(isElemVisible('inc_costOfSales') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_costOfSales') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('inc_costOfSales')}
                        <span>Less- {lbl('income', 'costOfSales', 'Cost of Goods Sold')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.costOfGoodsSold)}
                      </td>
                    ))}
                  </tr>
                )}
                <tr className="bg-slate-100/80 font-bold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">Gross Profit</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.grossProfit)}
                    </td>
                  ))}
                </tr>
                <tr className="bg-slate-50 font-semibold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">
                    <div className="flex items-center gap-2">
                      {renderCheck('inc_operatingExpenses')}
                      <span>Less- Administrative &amp; Selling Expenses</span>
                    </div>
                  </td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.adminAndSellingTotal)}
                    </td>
                  ))}
                </tr>
                {([
                  ['Salary & Allowances', 'salaryAndAllowances', 'inc_operatingExpenses'],
                  ['Telephone & Mobile Bill', 'telephoneAndMobileBill', 'inc_operatingExpenses'],
                  ['Postage & Courier', 'postageAndCourier', 'inc_operatingExpenses'],
                  ['Electric Charges', 'electricCharges', 'inc_operatingExpenses'],
                  ['Entertainment', 'entertainment', 'inc_operatingExpenses'],
                  ['Printing & Stationery', 'printingAndStationery', 'inc_operatingExpenses'],
                  ['General Expenses', 'generalExpenses', 'inc_operatingExpenses'],
                  ['Licence & Renewals', 'licenceAndRenewals', 'inc_operatingExpenses'],
                  ['Audit & Professional Fee', 'auditAndProfessionalFee', 'inc_operatingExpenses'],
                  ['Travelling & Conveyance', 'travellingAndConveyance', 'inc_operatingExpenses'],
                  ['News Paper & Periodicals', 'newspaperAndPeriodicals', 'inc_operatingExpenses'],
                  ['Selling Exp.', 'sellingExp', 'inc_operatingExpenses'],
                  [lbl('income', 'depreciationExpense', 'Depreciation'), 'depreciation', 'inc_depreciationExpense'],
                ] as const).map(([label, key, cfgId]) => {
                  const vis = isElemVisible(cfgId);
                  if (!vis && !isCustomizeMode) return null;
                  return (
                    <tr key={key} className={!vis ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                      <td className="py-1.5 px-4 pl-7 text-slate-600 border-r border-slate-200">
                        <div className="flex items-center gap-2">
                          {cfgId === 'inc_depreciationExpense' && renderCheck(cfgId)}
                          <span>{label}</span>
                        </div>
                      </td>
                      {displayedData.map(d => (
                        <td key={d.year} className="py-1.5 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                          {fmtCell(d[key])}
                        </td>
                      ))}
                    </tr>
                  );
                })}
                {customIncOpex.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-1.5 px-4 pl-7 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-1.5 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(-getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-100/80 font-bold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">Net Operating Profit</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.netOperatingProfit)}
                    </td>
                  ))}
                </tr>
                {(isElemVisible('inc_interestIncome') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_interestIncome') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('inc_interestIncome')}
                        <span>Add: {lbl('income', 'interestIncome', 'Other Income')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => {
                      const custSum = visIncNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
                      return (
                        <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                          {fmtCell(Math.round((d.otherIncome - custSum) * 100) / 100)}
                        </td>
                      );
                    })}
                  </tr>
                )}
                {customIncNonOp.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 pl-7 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-50 font-semibold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">Subtotal</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
                      {fmtCell(d.operatingSubtotal)}
                    </td>
                  ))}
                </tr>
                {(isElemVisible('inc_financeCost') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_financeCost') ? 'opacity-45 bg-slate-50' : 'bg-amber-50/65'}>
                    <td className="py-2 px-4 text-amber-900 font-medium border-r border-slate-200">
                      <div className="flex items-center gap-1.5">
                        {renderCheck('inc_financeCost')}
                        <span>Less: {lbl('income', 'financeCost', 'Financial Exp.')}</span>
                        <Lock className="w-3 h-3 text-amber-700" />
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-amber-800 font-semibold border-r border-slate-200 last:border-r-0">
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
                )}
                {(isElemVisible('inc_taxExpense') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_taxExpense') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('inc_taxExpense')}
                        <span>Less- {lbl('income', 'taxExpense', 'Income Tax Provision')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">
                        {fmtCell(d.incomeTaxProvision)}
                      </td>
                    ))}
                  </tr>
                )}
                <tr className="bg-emerald-50 font-bold text-emerald-950 border-t border-b-4 border-double border-emerald-700">
                  <td className="py-2.5 px-4 border-r border-slate-200 flex items-center gap-1.5">
                    <span>Net Profit After Income Tax</span>
                    <Lock className="w-3 h-3 text-emerald-700" />
                  </td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
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
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-white">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-800 uppercase tracking-wide">
              STATEMENT OF CASH FLOWS
            </div>
            <div className="text-xs font-semibold text-slate-700">
              For the year ended 31 December {maxYear}
            </div>
            <div className="text-[11px] text-slate-500 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-300 text-slate-900">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-300 min-w-[340px]">Particulars</th>
                  {displayedData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-300 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 tabular-nums">
                <tr className="bg-emerald-50/70 font-bold text-emerald-900">
                  <td colSpan={displayedData.length + 1} className="py-2 px-4">a) CASH FLOWS FROM OPERATING ACTIVITIES:</td>
                </tr>
                {(isElemVisible('cf_pbt') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_pbt') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-800 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_pbt')}
                        <span>{lbl('cf', 'profitBeforeTax', 'Net Income')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-800 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfNetIncome)}</td>)}
                  </tr>
                )}
                {(isElemVisible('cf_dep') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_dep') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_dep')}
                        <span>{lbl('cf', 'adjDepreciation', 'Add: Depreciation of Tangible Fixed Assets')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfDepreciation)}</td>)}
                  </tr>
                )}
                {(isElemVisible('cf_wc_receivables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_wc_receivables') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_wc_receivables')}
                        <span>{lbl('cf', 'changeInReceivables', 'Add: (Increase) or decrease in Current Assets')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfChangeInCurrentAssets)}</td>)}
                  </tr>
                )}
                <tr>
                  <td className="py-2 px-4 text-slate-700 border-r border-slate-200">Add: (Increase) or decrease in Construction Working Progress</td>
                  {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfChangeInCWIP)}</td>)}
                </tr>
                {(isElemVisible('cf_wc_payables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_wc_payables') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_wc_payables')}
                        <span>{lbl('cf', 'changeInPayables', 'Add: Increase or (decrease) in Current Liabilities')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfChangeInCurrentLiabilities)}</td>)}
                  </tr>
                )}
                {customCfOp.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-emerald-50/75 font-bold text-emerald-900">
                  <td className="py-2 px-4 border-r border-slate-200">Net Cash from operating activities</td>
                  {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">{fmtCell(d.cfNetCashOperating)}</td>)}
                </tr>

                <tr className="bg-sky-50/70 font-bold text-sky-900">
                  <td colSpan={displayedData.length + 1} className="py-2 px-4">b) CASH FLOWS FROM INVESTING ACTIVITIES:</td>
                </tr>
                {(isElemVisible('cf_pur_ppe') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_pur_ppe') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_pur_ppe')}
                        <span>{lbl('cf', 'purchasePPE', 'Acquisition of Tangible Fixed Assets')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfAcquisitionPPE)}</td>)}
                  </tr>
                )}
                {(isElemVisible('cf_pur_intangible') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_pur_intangible') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_pur_intangible')}
                        <span>{lbl('cf', 'purchaseIntangibles', 'Disposal of Intangible Assets')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfDisposalIntangible)}</td>)}
                  </tr>
                )}
                {customCfInv.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-sky-50/75 font-bold text-sky-900">
                  <td className="py-2 px-4 border-r border-slate-200">Net Cash used in investing Activities</td>
                  {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">{fmtCell(d.cfNetCashInvesting)}</td>)}
                </tr>

                <tr className="bg-amber-50/70 font-bold text-amber-900">
                  <td colSpan={displayedData.length + 1} className="py-2 px-4">c) CASH FLOWS FROM FINANCING ACTIVITIES:</td>
                </tr>
                {(isElemVisible('cf_dividends_paid') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_dividends_paid') ? 'opacity-45 bg-slate-50' : ''}>
                    <td className="py-2 px-4 text-slate-700 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_dividends_paid')}
                        <span>{lbl('cf', 'financingDividendsPaid', 'Drawings')}</span>
                      </div>
                    </td>
                    {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfDrawings)}</td>)}
                  </tr>
                )}
                <tr>
                  <td className="py-2 px-4 text-slate-700 border-r border-slate-200">Opening Balance</td>
                  {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfOpeningBalance)}</td>)}
                </tr>
                {customCfFin.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-amber-50/75 font-bold text-amber-900">
                  <td className="py-2 px-4 border-r border-slate-200">Net Cash used in Financing Activities</td>
                  {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">{fmtCell(d.cfNetCashFinancing)}</td>)}
                </tr>

                <tr className="font-semibold text-slate-900">
                  <td className="py-2 px-4 border-r border-slate-200">d) Net Increase decrease in cash &amp; cash equivalents (A+B+C)</td>
                  {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">{fmtCell(d.cfNetIncreaseCash)}</td>)}
                </tr>
                <tr>
                  <td className="py-2 px-4 text-slate-700 border-r border-slate-200">e) Cash &amp; Cash Equivalents at beginning of period</td>
                  {displayedData.map(d => <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-700 border-r border-slate-200 last:border-r-0">{fmtCell(d.cfCashBeginning)}</td>)}
                </tr>
                <tr className="bg-emerald-50 font-bold text-emerald-950 border-t border-b-4 border-double border-emerald-700">
                  <td className="py-2.5 px-4 border-r border-slate-200">f) Cash &amp; Cash Equivalents at end of period (D+E)</td>
                  {displayedData.map(d => <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">{fmtCell(d.cfCashEnding)}</td>)}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 4. STATEMENT OF CHANGES IN EQUITY */}
      {(subTab === 'all' || subTab === 'eq') && (
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-white">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide uppercase">
              {companyProfile.companyName}
            </h3>
            <div className="text-sm font-bold text-emerald-800 uppercase tracking-wide">
              STATEMENT OF CHANGES IN EQUITY
            </div>
            <div className="text-xs font-semibold text-slate-700">
              For the year ended 31 December {maxYear}
            </div>
            <div className="text-[11px] text-slate-500 text-right font-mono pt-1">
              {unitCaption}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-300 text-slate-900">
                  <th className="py-2.5 px-4 text-left font-bold border-r border-slate-300 min-w-[290px]">Particulars</th>
                  {displayedData.map(d => (
                    <th key={d.year} className="py-2.5 px-4 text-right font-mono font-bold border-r border-slate-300 last:border-r-0 min-w-[130px]">
                      {d.dateLabel}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 tabular-nums">
                {([
                  [lbl('eq', 'openingShareCapital', 'Opening Share Capital'), 'eqOpeningShareCapital', 'eq_openingShareCapital'],
                  [lbl('eq', 'shareCapitalIssued', 'Share Capital Issued'), 'eqShareCapitalIssued', 'eq_shareCapitalIssued'],
                  [lbl('eq', 'closingShareCapital', 'Closing Share Capital'), 'eqClosingShareCapital', null],
                  [lbl('eq', 'openingRetainedEarnings', 'Opening Retained Earnings'), 'eqOpeningRetainedEarnings', 'eq_openingRetainedEarnings'],
                  [lbl('eq', 'netProfitForYear', 'Net Profit for the Year'), 'eqNetProfitForYear', 'eq_netProfitForYear'],
                  [lbl('eq', 'dividendsPaid', 'Dividends Paid'), 'eqDividendsPaid', 'eq_dividendsPaid'],
                  [lbl('eq', 'closingRetainedEarnings', 'Closing Retained Earnings'), 'eqClosingRetainedEarnings', null],
                  [lbl('eq', 'closingReserves', 'Closing Reserves'), 'eqClosingReserves', 'eq_openingReserves'],
                ] as const).map(([label, key, cfgId]) => {
                  const vis = cfgId ? isElemVisible(cfgId) : true;
                  if (!vis && !isCustomizeMode) return null;
                  return (
                    <tr key={key} className={!vis ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/70'}>
                      <td className="py-2 px-4 text-slate-800 border-r border-slate-200">
                        <div className="flex items-center gap-2">
                          {cfgId && renderCheck(cfgId)}
                          <span>{label}</span>
                        </div>
                      </td>
                      {displayedData.map(d => (
                        <td key={d.year} className="py-2 px-4 text-right font-mono text-slate-800 border-r border-slate-200 last:border-r-0">
                          {key === 'eqDividendsPaid' ? fmtCell(-d[key]) : fmtCell(d[key])}
                        </td>
                      ))}
                    </tr>
                  );
                })}
                {customEq.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900 border-r border-slate-200">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {displayedData.map(d => (
                      <td key={d.year} className="py-2 px-4 text-right font-mono text-indigo-800 border-r border-slate-200 last:border-r-0">
                        {fmtCell(getCustomItemYearValue(ci, d.year))}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-emerald-50 font-bold text-emerald-950 border-t border-b-4 border-double border-emerald-700">
                  <td className="py-2.5 px-4 border-r border-slate-200">{lbl('eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY')}</td>
                  {displayedData.map(d => (
                    <td key={d.year} className="py-2.5 px-4 text-right font-mono border-r border-slate-200 last:border-r-0">
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
          <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-white">
              <h3 className="text-base font-extrabold text-slate-900 tracking-wide uppercase">
                {companyProfile.companyName}
              </h3>
              <div className="text-xs font-semibold text-slate-700">
                For the Year ended 31 December {maxYear}
              </div>
              <div className="text-left text-xs font-bold text-emerald-800 pt-2">
                1.00 FIXED ASSETS SCHEDULE
              </div>
            </div>

            <div className="p-4 space-y-4 overflow-x-auto">
              {faAsc.map(d => (
                <table key={d.year} className="w-full text-xs border border-slate-300 border-collapse tabular-nums">
                  <thead>
                    <tr className="bg-slate-100 text-slate-900 border-b border-slate-300">
                      <th className="py-2 px-3 text-left border-r border-slate-300">Particulars</th>
                      <th className="py-2 px-3 text-right border-r border-slate-300">Opening at 01.01.{d.year}</th>
                      <th className="py-2 px-3 text-right border-r border-slate-300">Addition during the year</th>
                      <th className="py-2 px-3 text-right border-r border-slate-300">Total at 31.12.{d.year}</th>
                      <th className="py-2 px-3 text-center border-r border-slate-300">Rate %</th>
                      <th className="py-2 px-3 text-right border-r border-slate-300">Depreciation charged for the year</th>
                      <th className="py-2 px-3 text-right">W.D.V. at 31.12.{d.year}</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="border-b border-slate-200">
                      <td className="py-2 px-3 text-slate-800 border-r border-slate-200">Fixed Assets</td>
                      <td className="py-2 px-3 text-right font-mono text-slate-800 border-r border-slate-200">{fmtCell(d.faOpening)}</td>
                      <td className="py-2 px-3 text-right font-mono text-slate-800 border-r border-slate-200">{fmtCell(d.faAddition - d.faDisposal)}</td>
                      <td className="py-2 px-3 text-right font-mono text-emerald-800 font-semibold border-r border-slate-200">
                        {onUpdatePrimaryDriver ? (
                          <InlineDriverCell
                            value={d.faGrossAssetBase}
                            isEditMode={isEditMode}
                            onSave={val => onUpdatePrimaryDriver(d.year, 'fixedAssetsGross', val)}
                            title="Fixed User Input: Fixed Assets (Gross Asset Base)"
                          />
                        ) : fmtCell(d.faTotal)}
                      </td>
                      <td className="py-2 px-3 text-center font-mono text-slate-700 border-r border-slate-200">{d.faRatePct}</td>
                      <td className="py-2 px-3 text-right font-mono text-slate-800 border-r border-slate-200">{fmtCell(d.faDepreciationCharged)}</td>
                      <td className="py-2 px-3 text-right font-mono text-slate-900 font-semibold">{fmtCell(d.faWDV)}</td>
                    </tr>
                    <tr className="bg-emerald-50 font-bold text-emerald-950">
                      <td className="py-2 px-3 border-r border-slate-200">Taka</td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.faOpening)}</td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.faAddition - d.faDisposal)}</td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.faTotal)}</td>
                      <td className="py-2 px-3 border-r border-slate-200"></td>
                      <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.faDepreciationCharged)}</td>
                      <td className="py-2 px-3 text-right font-mono">{fmtCell(d.faWDV)}</td>
                    </tr>
                  </tbody>
                </table>
              ))}
            </div>
          </div>

          <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl p-5 space-y-4 shadow-sm">
            <h4 className="text-sm font-bold text-slate-900 border-b border-slate-200 pb-2">
              NOTES TO THE FINANCIAL STATEMENTS:
            </h4>
            <div className="space-y-4 text-xs">
              {displayedNotes.map(note => (
                <div key={note.noteNumber} className="space-y-2">
                  <div className="font-bold text-emerald-800">
                    Note {note.noteNumber}: {note.title}
                  </div>
                  <p className="text-slate-700 leading-relaxed">{note.textDescription}</p>
                  {note.tableData && (
                    <div className="overflow-x-auto">
                      <table className="w-full text-[11px] border border-slate-200 border-collapse tabular-nums">
                        <thead>
                          <tr className="bg-slate-100 text-slate-800 border-b border-slate-200">
                            {note.tableData.headers.map((h, i) => (
                              <th key={i} className={`py-1.5 px-3 font-semibold ${i === 0 ? 'text-left' : 'text-right'}`}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {note.tableData.rows.map((r, rIdx) => (
                            <tr key={rIdx}>
                              <td className="py-1 px-3 text-slate-700">{r.label}</td>
                              {r.values.map((v, vIdx) => (
                                <td key={vIdx} className="py-1 px-3 text-right font-mono text-slate-800">{String(v)}</td>
                              ))}
                            </tr>
                          ))}
                          {note.tableData.totalRow && (
                            <tr className="bg-slate-50 font-bold text-slate-900">
                              <td className="py-1.5 px-3">{note.tableData.totalRow.label}</td>
                              {note.tableData.totalRow.values.map((v, vIdx) => (
                                <td key={vIdx} className="py-1.5 px-3 text-right font-mono">{String(v)}</td>
                              ))}
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  )}
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
  statementConfig?: StatementConfigItem[];
  onUpdateStatementConfig?: (items: StatementConfigItem[]) => void;
  onResetStatementConfig?: () => void;
  labelOverrides?: StatementLabelOverrides;
}

export const ProjectedFormatSuiteView: React.FC<ProjectedFormatSuiteViewProps> = ({
  companyProfile,
  projData,
  isEditMode,
  onUpdatePrimaryDriver,
  statementConfig = [],
  onUpdateStatementConfig,
  onResetStatementConfig,
  labelOverrides = {},
}) => {
  const [subTab, setSubTab] = useState<'all' | 'bs' | 'pl' | 'cf' | 'notes' | 'sched'>('all');
  const [isCustomizeMode, setIsCustomizeMode] = useState<boolean>(false);

  const isElemVisible = (id: string): boolean => {
    const found = statementConfig.find(c => c.id === id);
    return found ? found.isVisible : true;
  };

  const toggleElem = (id: string) => {
    if (!onUpdateStatementConfig || statementConfig.length === 0) return;
    onUpdateStatementConfig(
      statementConfig.map(c => (c.id === id ? { ...c, isVisible: !c.isVisible } : c))
    );
  };

  const renderCheck = (id: string) =>
    isCustomizeMode && onUpdateStatementConfig ? (
      <input
        type="checkbox"
        checked={isElemVisible(id)}
        onChange={() => toggleElem(id)}
        title="Include or exclude this element from calculation and downloads"
        className="w-3.5 h-3.5 accent-sky-600 rounded cursor-pointer shrink-0"
      />
    ) : null;

  const handleSetAllConfigVisibility = (select: boolean) => {
    if (!onUpdateStatementConfig || statementConfig.length === 0) return;
    onUpdateStatementConfig(
      statementConfig.map(c => ({ ...c, isVisible: select }))
    );
  };

  const lbl = (stmt: string, field: string, def: string) =>
    resolveEffectiveLabel(stmt, field, def, labelOverrides, statementConfig);

  const stdIncomeSections = ['Revenue & Gross Margin', 'Operating Expenses', 'Non-Operating & Tax'];
  const getCustom = (stmt: StatementConfigItem['statement'], section?: string) =>
    statementConfig.filter(c => c.isCustom && c.statement === stmt && (!section || c.section === section));
  const customBsNca = getCustom('balanceSheet', 'Non-Current Assets');
  const customBsCa = getCustom('balanceSheet', 'Current Assets');
  const customBsNcl = getCustom('balanceSheet', 'Non-Current Liabilities');
  const customBsCl = getCustom('balanceSheet', 'Current Liabilities');
  const customBsEq = getCustom('balanceSheet', 'Equity');
  const visBsNca = customBsNca.filter(c => c.isVisible);
  const visBsNcl = customBsNcl.filter(c => c.isVisible);

  const customIncRev = getCustom('income', 'Revenue & Gross Margin');
  const customIncOpex = getCustom('income', 'Operating Expenses');
  const customIncNonOp = statementConfig.filter(
    c => c.isCustom && c.statement === 'income' && (c.section === 'Non-Operating & Tax' || !stdIncomeSections.includes(c.section))
  );
  const visIncRev = customIncRev.filter(c => c.isVisible);
  const visIncOpex = customIncOpex.filter(c => c.isVisible);
  const visIncNonOp = customIncNonOp.filter(c => c.isVisible);
  const showAdminOpexRow = isElemVisible('inc_operatingExpenses') || isElemVisible('inc_depreciationExpense');

  const customCfOp = getCustom('cashFlow', 'Operating Activities');
  const customCfInv = getCustom('cashFlow', 'Investing Activities');
  const customCfFin = getCustom('cashFlow', 'Financing Activities');

  const exportSyncOptions = { statementConfig, labelOverrides };

  if (projData.length === 0) {
    return (
      <div className="bg-white/92 backdrop-blur-md border border-slate-200 rounded-2xl p-8 text-center text-xs text-slate-600 shadow-sm">
        Select at least one projected financial year to render the Projected Financials statements.
      </div>
    );
  }

  const maxYear = projData[0].year; // Descending order
  const minYear = projData[projData.length - 1].year;
  const rangeStr = minYear === maxYear ? `${maxYear}` : `${minYear}–${maxYear}`;
  const amtLabel = companyProfile.reportingCurrency === '৳' ? 'Amount In Taka' : `Amount In ${companyProfile.reportingCurrency}`;
  const unitNote = companyProfile.reportingUnit !== 'Units' ? ` · ${companyProfile.reportingUnit}` : '';
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
      <div className="bg-white/92 backdrop-blur-md border border-sky-200 rounded-2xl p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-4 shadow-sm">
        <div className="space-y-1">
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <span className="font-mono font-bold uppercase tracking-wider text-sky-700">
              Projected Financials Format
            </span>
            <span className="text-slate-300" aria-hidden="true">·</span>
            <span className="font-bold text-slate-900">
              {projData.length} Projected {projData.length === 1 ? 'Year' : 'Years'} Selected ({rangeStr})
            </span>
          </div>
          <p className="text-[11px] text-slate-600">
            Reproduces exact landscape layout, Notes column, two-row {amtLabel} / Year headers, Proprietor signoff, Notes 1–13.00, and Schedule-A PPE{unitNote}.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {onUpdateStatementConfig && statementConfig.length > 0 && (
            <button
              type="button"
              onClick={() => setIsCustomizeMode(!isCustomizeMode)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 border transition-colors whitespace-nowrap shrink-0 ${
                isCustomizeMode
                  ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
              }`}
            >
              <CheckSquare className="w-3.5 h-3.5" />
              <span>{isCustomizeMode ? 'Done Selecting' : 'Customize Elements'}</span>
            </button>
          )}
          {isCustomizeMode && onUpdateStatementConfig && (
            <>
              <button
                type="button"
                onClick={() => handleSetAllConfigVisibility(true)}
                className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-emerald-50 text-emerald-800 border border-slate-300 hover:border-emerald-300 text-xs font-medium flex items-center gap-1 whitespace-nowrap shrink-0"
              >
                <CheckSquare className="w-3.5 h-3.5 text-emerald-600" />
                Select All
              </button>
              <button
                type="button"
                onClick={() => handleSetAllConfigVisibility(false)}
                className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-rose-50 text-rose-700 border border-slate-300 hover:border-rose-300 text-xs font-medium flex items-center gap-1 whitespace-nowrap shrink-0"
              >
                <Square className="w-3.5 h-3.5 text-rose-500" />
                Deselect All
              </button>
            </>
          )}
          {isCustomizeMode && onResetStatementConfig && (
            <button
              type="button"
              onClick={onResetStatementConfig}
              className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-medium flex items-center gap-1 whitespace-nowrap shrink-0"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Reset Defaults
            </button>
          )}
          <button
            type="button"
            onClick={() => exportProjectedFinancialsExcel(companyProfile, projData, exportSyncOptions)}
            className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors whitespace-nowrap shrink-0"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            Projected Excel (.xlsx)
          </button>
          <button
            type="button"
            onClick={() => exportProjectedFinancialsWord(companyProfile, projData, exportSyncOptions)}
            className="px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors whitespace-nowrap shrink-0"
          >
            <FileText className="w-3.5 h-3.5" />
            Projected Word (.docx)
          </button>
          <button
            type="button"
            onClick={() => exportProjectedFinancialsPdf(companyProfile, projData, exportSyncOptions)}
            className="px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors whitespace-nowrap shrink-0"
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
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors border whitespace-nowrap shrink-0 ${
                subTab === t.id
                  ? 'bg-sky-600 text-white border-sky-600 shadow-sm'
                  : 'bg-white/90 text-slate-700 border-slate-200 hover:bg-slate-50 hover:text-slate-900'
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
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-slate-50/75">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-sm font-bold text-sky-800">Statement of Financial Position</div>
            <div className="text-xs text-slate-600">As at 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-100/95 border-b border-slate-200 text-slate-900">
                  <th rowSpan={2} className="py-2 px-4 text-center font-bold border-r border-slate-200 min-w-[270px]">Particulars</th>
                  <th rowSpan={2} className="py-2 px-3 text-center font-bold border-r border-slate-200 w-24">Notes</th>
                  {projData.map(d => (
                    <th key={`amt-${d.year}`} className="py-1.5 px-3 text-center font-semibold border-r border-b border-slate-200 last:border-r-0 min-w-[125px] text-[11px]">
                      {amtLabel}
                    </th>
                  ))}
                </tr>
                <tr className="bg-slate-50/90 border-b border-slate-200 text-sky-800 font-mono">
                  {projData.map(d => (
                    <th key={`yr-${d.year}`} className="py-1.5 px-3 text-center font-bold border-r border-slate-200 last:border-r-0">
                      {d.year}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/80 text-slate-800">
                <tr className="bg-slate-50 font-bold text-slate-900 underline">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Assets</td>
                </tr>
                <tr className="font-semibold text-slate-900">
                  <td colSpan={projData.length + 2} className="py-1.5 px-4">Non Current Assets</td>
                </tr>
                {(isElemVisible('bs_propertyPlantEquipment') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_propertyPlantEquipment') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_propertyPlantEquipment')}
                        <span>{lbl('bs', 'propertyPlantEquipment', 'Property Plant & Equiptments')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-sky-700 font-semibold">Schedule-A</td>
                    {projData.map(d => {
                      const custSum = visBsNca.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
                      return (
                        <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">
                          {fmtCell(Math.round((d.propertyPlantEquipments - custSum) * 100) / 100)}
                        </td>
                      );
                    })}
                  </tr>
                )}
                {customBsNca.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-100/70 font-bold text-slate-900 border-t border-b-2 border-slate-300">
                  <td className="py-1.5 px-4">Total Non-Current Assets</td>
                  <td></td>
                  {projData.map(d => (
                    <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.propertyPlantEquipments)}</td>
                  ))}
                </tr>

                <tr className="font-semibold text-slate-900">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Current Assets</td>
                </tr>
                {(isElemVisible('bs_closingInventory') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_closingInventory') ? 'opacity-45 bg-slate-50' : 'bg-emerald-50/60'}>
                    <td className="py-2 px-4 text-emerald-900 font-medium">
                      <div className="flex items-center gap-1.5">
                        {renderCheck('bs_closingInventory')}
                        <span>{lbl('bs', 'closingInventory', 'Inventories')}</span>
                        <Lock className="w-3 h-3 text-emerald-600" />
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">3.00</td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-emerald-800 font-semibold">
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
                )}
                {(isElemVisible('bs_tradeReceivables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_tradeReceivables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_tradeReceivables')}
                        <span>{lbl('bs', 'tradeReceivables', 'Trade & Other Receivables')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">4.00</td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.tradeAndOtherReceivables)}</td>)}
                  </tr>
                )}
                {(isElemVisible('bs_otherReceivables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_otherReceivables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_otherReceivables')}
                        <span>{lbl('bs', 'otherReceivables', 'Advance, Deposit & Prepayments')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">5.00</td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.advanceDepositPrepayments)}</td>)}
                  </tr>
                )}
                {(isElemVisible('bs_cashAndEquivalents') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_cashAndEquivalents') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_cashAndEquivalents')}
                        <span>{lbl('bs', 'cashAndEquivalents', 'Cash & Cash Equivalents')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">6.00</td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.cashAndCashEquivalents)}</td>)}
                  </tr>
                )}
                {customBsCa.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-100/70 font-bold text-slate-900 border-t border-slate-300">
                  <td className="py-1.5 px-4">Total Current Assets</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.totalCurrentAssets)}</td>)}
                </tr>
                <tr className="bg-sky-50/80 font-bold text-sky-900 border-t border-b-4 border-double border-sky-600">
                  <td className="py-2.5 px-4 text-right">Total Assets</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2.5 px-3 text-right font-mono">{fmtCell(d.totalAssets)}</td>)}
                </tr>

                <tr className="bg-slate-50 font-bold text-slate-900 underline">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Equity &amp; Liabilities</td>
                </tr>
                <tr className="hover:bg-slate-50/80">
                  <td className="py-2 px-4 text-slate-700">Capital</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.capital)}</td>)}
                </tr>
                <tr className="hover:bg-slate-50/80">
                  <td className="py-2 px-4 text-slate-700">Additional Capital</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.additionalCapital)}</td>)}
                </tr>
                <tr className="hover:bg-slate-50/80">
                  <td className="py-2 px-4 text-slate-700">Profit during the year</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-emerald-700 font-semibold">{fmtCell(d.profitDuringYear)}</td>)}
                </tr>
                <tr className="bg-slate-100/70 font-bold text-slate-900 border-t border-slate-300">
                  <td className="py-1.5 px-4">Subtotal</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.equitySubtotal)}</td>)}
                </tr>
                {(isElemVisible('eq_dividendsPaid') || isCustomizeMode) && (
                  <tr className={!isElemVisible('eq_dividendsPaid') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('eq_dividendsPaid')}
                        <span>{lbl('eq', 'dividendsPaid', 'Drawings')}</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-rose-700">{fmtCell(d.drawings)}</td>)}
                  </tr>
                )}
                <tr className="bg-slate-100/80 font-bold text-slate-900 border-t border-b-2 border-slate-300">
                  <td className="py-1.5 px-4">{lbl('eq', 'closingShareCapital', 'Closing Capital / Equity')}</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.closingCapital)}</td>)}
                </tr>
                {customBsEq.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}

                <tr className="font-semibold text-slate-900">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Non-Current Liabilities</td>
                </tr>
                {(isElemVisible('bs_longTermBorrowings') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_longTermBorrowings') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_longTermBorrowings')}
                        <span>{lbl('bs', 'longTermBorrowings', 'Long Term Loan')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">7.00</td>
                    {projData.map(d => {
                      const custSum = visBsNcl.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
                      return (
                        <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">
                          {fmtCell(Math.round((d.longTermLoan - custSum) * 100) / 100)}
                        </td>
                      );
                    })}
                  </tr>
                )}
                {customBsNcl.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}

                <tr className="font-semibold text-slate-900">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Current Liabilities</td>
                </tr>
                {(isElemVisible('bs_tradePayables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_tradePayables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_tradePayables')}
                        <span>{lbl('bs', 'tradePayables', 'Trade & Other Payables')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">8.00</td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.tradeAndOtherPayables)}</td>)}
                  </tr>
                )}
                {(isElemVisible('bs_currentTaxPayable') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_currentTaxPayable') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_currentTaxPayable')}
                        <span>{lbl('bs', 'currentTaxPayable', 'Provisions For Tax')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">9.00</td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.provisionsForTax)}</td>)}
                  </tr>
                )}
                {(isElemVisible('bs_shortTermBorrowings') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_shortTermBorrowings') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_shortTermBorrowings')}
                        <span>{lbl('bs', 'shortTermBorrowings', 'Short Term Loan')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">10.00</td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.shortTermLoan)}</td>)}
                  </tr>
                )}
                {(isElemVisible('bs_otherPayables') || isCustomizeMode) && (
                  <tr className={!isElemVisible('bs_otherPayables') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('bs_otherPayables')}
                        <span>{lbl('bs', 'otherPayables', 'Liabilities for exp.')}</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.liabilitiesForExp)}</td>)}
                  </tr>
                )}
                {customBsCl.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-100/70 font-bold text-slate-900 border-t border-slate-300">
                  <td className="py-1.5 px-4">Total Current Liabilities</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.totalCurrentLiabilities)}</td>)}
                </tr>
                <tr className="bg-sky-50/80 font-bold text-sky-900 border-t border-b-4 border-double border-sky-600">
                  <td className="py-2.5 px-4 text-right">Total Equity &amp; Liabilities</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2.5 px-3 text-right font-mono">{fmtCell(d.totalEquityAndLiabilities)}</td>)}
                </tr>
              </tbody>
            </table>
          </div>

          <div className="px-6 py-4 bg-slate-50/85 border-t border-slate-200 flex flex-col sm:flex-row items-start sm:items-end justify-between gap-4 text-xs text-slate-700">
            <div>
              <div className="text-slate-400">...............................</div>
              <div className="font-bold text-slate-900">Proprietor</div>
            </div>
            <div className="italic text-slate-600">
              Signed in terms of our separate report of even date.
            </div>
          </div>
        </div>
      )}

      {/* 2. STATEMENT OF PROFIT OR LOSS & OTHER COMPREHENSIVE INCOME */}
      {(subTab === 'all' || subTab === 'pl') && (
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-slate-50/75">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-sm font-bold text-sky-800">Statement of Profit or Loss &amp; Other Comprehensive Income</div>
            <div className="text-xs text-slate-600">For the year ended 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-100/95 border-b border-slate-200 text-slate-900">
                  <th rowSpan={2} className="py-2 px-4 text-center font-bold border-r border-slate-200 min-w-[270px]">Particulars</th>
                  <th rowSpan={2} className="py-2 px-3 text-center font-bold border-r border-slate-200 w-24">Notes</th>
                  {projData.map(d => (
                    <th key={`pl-amt-${d.year}`} className="py-1.5 px-3 text-center font-semibold border-r border-b border-slate-200 last:border-r-0 min-w-[125px] text-[11px]">
                      {amtLabel}
                    </th>
                  ))}
                </tr>
                <tr className="bg-slate-50/90 border-b border-slate-200 text-sky-800 font-mono">
                  {projData.map(d => (
                    <th key={`pl-yr-${d.year}`} className="py-1.5 px-3 text-center font-bold border-r border-slate-200 last:border-r-0">
                      {d.year}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/80 text-slate-800">
                {(isElemVisible('inc_revenue') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_revenue') ? 'opacity-45 bg-slate-50' : 'bg-emerald-50/60 font-semibold text-emerald-900'}>
                    <td className="py-2 px-4">
                      <div className="flex items-center gap-1.5">
                        {renderCheck('inc_revenue')}
                        <span>{lbl('income', 'revenue', 'Sales')}</span>
                        <Lock className="w-3 h-3 text-emerald-600" />
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => {
                      const custSum = visIncRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
                      const baseSales = Math.round((d.sales - custSum) * 100) / 100;
                      return (
                        <td key={d.year} className="py-2 px-3 text-right font-mono text-emerald-800">
                          {onUpdatePrimaryDriver ? (
                            <InlineDriverCell
                              value={baseSales}
                              isEditMode={isEditMode}
                              onSave={val => onUpdatePrimaryDriver(d.year, 'salesRevenue', val)}
                              title="Fixed User Input: Sales Revenue"
                            />
                          ) : fmtCell(baseSales)}
                        </td>
                      );
                    })}
                  </tr>
                )}
                {customIncRev.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 pl-7 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                {(isElemVisible('inc_costOfSales') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_costOfSales') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('inc_costOfSales')}
                        <span>Less: {lbl('income', 'costOfSales', 'Cost of Goods Sold')}</span>
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">11.00</td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.costOfGoodsSold)}</td>)}
                  </tr>
                )}
                <tr className="bg-slate-100/80 font-bold text-slate-900 border-t border-slate-300">
                  <td className="py-2 px-4">Gross Profit</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.grossProfit)}</td>)}
                </tr>
                {(isElemVisible('inc_interestIncome') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_interestIncome') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('inc_interestIncome')}
                        <span>Add; {lbl('income', 'interestIncome', 'Others Income')}</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => {
                      const custSum = visIncNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
                      return (
                        <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">
                          {fmtCell(Math.round((d.othersIncome - custSum) * 100) / 100)}
                        </td>
                      );
                    })}
                  </tr>
                )}
                {customIncNonOp.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 pl-7 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-100/60 font-bold text-slate-900 border-t border-slate-300">
                  <td className="py-1.5 px-4">Total Gross Profit &amp; Other Income</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono">{fmtCell(d.grossProfitWithOtherIncome)}</td>)}
                </tr>
                <tr className="font-semibold text-slate-900">
                  <td colSpan={projData.length + 2} className="py-2 px-4">Less: Operating Expenses:</td>
                </tr>
                {(showAdminOpexRow || isCustomizeMode) && (
                  <tr className={!showAdminOpexRow ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex flex-wrap items-center gap-2">
                        {renderCheck('inc_operatingExpenses')}
                        <span>{lbl('income', 'operatingExpenses', 'Administrative & Selling Expenses')}</span>
                        {isCustomizeMode && (
                          <label className="inline-flex items-center gap-1 text-[10px] text-slate-600 ml-1 cursor-pointer">
                            {renderCheck('inc_depreciationExpense')}
                            <span>(incl. {lbl('income', 'depreciationExpense', 'Depreciation')})</span>
                          </label>
                        )}
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">12.00</td>
                    {projData.map(d => {
                      const custSum = visIncOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
                      return (
                        <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">
                          {fmtCell(Math.round((d.adminAndSellingExpenses + custSum) * 100) / 100)}
                        </td>
                      );
                    })}
                  </tr>
                )}
                {customIncOpex.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2 px-4 pl-7 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-indigo-800">{fmtCell(-getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-slate-100/80 font-bold text-slate-900 border-t border-slate-300">
                  <td className="py-2 px-4">Total Operating Profit/(Loss)</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.totalOperatingProfit)}</td>)}
                </tr>
                {(isElemVisible('inc_financeCost') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_financeCost') ? 'opacity-45 bg-slate-50' : 'bg-amber-50/75'}>
                    <td className="py-2 px-4 text-amber-900 font-medium">
                      <div className="flex items-center gap-1.5">
                        {renderCheck('inc_financeCost')}
                        <span>Less: {lbl('income', 'financeCost', 'Financial Expenses')}</span>
                        <Lock className="w-3 h-3 text-amber-600" />
                      </div>
                    </td>
                    <td className="py-2 px-3 text-center font-mono text-slate-600">13.00</td>
                    {projData.map(d => (
                      <td key={d.year} className="py-2 px-3 text-right font-mono text-amber-800 font-semibold">
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
                )}
                <tr className="bg-slate-100/70 font-bold text-slate-900 border-t border-slate-300">
                  <td className="py-2 px-4">Net Profit/(Loss) Before Income Tax</td>
                  <td></td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono">{fmtCell(d.netProfitBeforeTax)}</td>)}
                </tr>
                {(isElemVisible('inc_taxExpense') || isCustomizeMode) && (
                  <tr className={!isElemVisible('inc_taxExpense') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-2 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('inc_taxExpense')}
                        <span>Less: {lbl('income', 'taxExpense', 'Provision For Income Tax')}</span>
                      </div>
                    </td>
                    <td></td>
                    {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900">{fmtCell(d.provisionForIncomeTax)}</td>)}
                  </tr>
                )}
                <tr className="bg-sky-50/85 font-bold text-sky-900 border-t border-b-4 border-double border-sky-600">
                  <td className="py-2.5 px-4 flex items-center gap-1.5">
                    <span>Net Profit/(Loss) After Income Tax</span>
                    <Lock className="w-3 h-3 text-emerald-600" />
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

          <div className="px-6 py-4 bg-slate-50/85 border-t border-slate-200 flex flex-col sm:flex-row items-start sm:items-end justify-between gap-4 text-xs text-slate-700">
            <div>
              <div className="text-slate-400">...............................</div>
              <div className="font-bold text-slate-900">Proprietor</div>
            </div>
            <div className="italic text-slate-600">
              Signed in terms of our separate report of even date.
            </div>
          </div>
        </div>
      )}

      {/* 3. STATEMENT OF CASH FLOWS */}
      {(subTab === 'all' || subTab === 'cf') && (
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-slate-50/75">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-[11px] text-slate-600">{companyProfile.registeredAddress}</div>
            <div className="text-sm font-bold text-sky-800">Statement Of Cash Flows</div>
            <div className="text-xs text-slate-600">For the year ended 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-100/95 border-b border-slate-200 text-slate-900">
                  <th className="py-2.5 px-4 text-left font-bold min-w-[330px]">1. CASH FLOW FROM OPERATING ACTIVITIES:</th>
                  {projData.map(d => (
                    <th key={d.year} className="py-2.5 px-3 text-center font-mono font-bold border-l border-slate-200 min-w-[120px]">
                      {d.year}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/80 text-slate-800">
                {([
                  [lbl('cf', 'profitBeforeTax', 'Net Income'), 'cfNetIncome', 'cf_pbt'],
                  [lbl('cf', 'changeInInventory', '(Increase)/Decrease In Inventory'), 'cfChangeInventory', 'cf_wc_inventory'],
                  [lbl('cf', 'changeInReceivables', '(Increase)/Decrease In Trade & Other Receivables'), 'cfChangeReceivables', 'cf_wc_receivables'],
                  [lbl('cf', 'changeInOtherReceivables', '(Increase)/Decrease in Advance, Deposit & Prepayments'), 'cfChangeAdvanceDeposit', 'cf_wc_other_rec'],
                  [lbl('cf', 'changeInOtherCA', 'Increase/(Decrease) in Current Assets'), 'cfChangeCurrentAssets', 'cf_wc_other_ca'],
                  ['Increase/(Decrease) in Current Liabilities', 'cfChangeCurrentLiabilities', null],
                  [lbl('cf', 'changeInPayables', 'Increase/(Decrease) in Trade & Other Payables'), 'cfChangeTradePayables', 'cf_wc_payables'],
                  [lbl('cf', 'incomeTaxPaid', 'Increase/(Decrease) In Provisions For Tax'), 'cfChangeProvisionsTax', 'cf_tax_paid'],
                  [lbl('cf', 'changeInOtherPayables', 'Increase/(Decrease) Liabilities For Expense'), 'cfChangeLiabilitiesExp', 'cf_wc_other_pay'],
                  [lbl('cf', 'adjDepreciation', 'Depreciation'), 'cfDepreciation', 'cf_dep'],
                ] as const).map(([label, key, cfgId]) => {
                  const vis = cfgId ? isElemVisible(cfgId) : true;
                  if (!vis && !isCustomizeMode) return null;
                  return (
                    <tr key={key} className={!vis ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                      <td className="py-1.5 px-4 text-slate-700">
                        <div className="flex items-center gap-2">
                          {cfgId && renderCheck(cfgId)}
                          <span>{label}</span>
                        </div>
                      </td>
                      {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d[key])}</td>)}
                    </tr>
                  );
                })}
                {customCfOp.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-1.5 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {projData.map(d => (
                      <td key={d.year} className="py-1.5 px-3 text-right font-mono text-indigo-800 border-l border-slate-200/80">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-emerald-50/75 font-bold text-emerald-900">
                  <td className="py-2 px-4">Net Cash Used by Operating Activities</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-200/80">{fmtCell(d.cfNetCashOperating)}</td>)}
                </tr>

                <tr className="bg-slate-50 font-bold text-sky-900">
                  <td colSpan={projData.length + 1} className="py-2 px-4 underline">2. CASH FLOW FROM INVESTING ACTIVITIES:</td>
                </tr>
                <tr className="font-semibold text-slate-800">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Inflows:</td>
                </tr>
                {(isElemVisible('cf_sale_ppe') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_sale_ppe') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-1.5 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_sale_ppe')}
                        <span>{lbl('cf', 'salePPE', 'Sale Of Property, Plant & Equipment')}</span>
                      </div>
                    </td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfSalePPE)}</td>)}
                  </tr>
                )}
                <tr className="hover:bg-slate-50/80">
                  <td className="py-1.5 px-4 text-slate-700">Capital</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfInvestingCapitalInflow)}</td>)}
                </tr>
                <tr className="font-semibold text-slate-800">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Outflows:</td>
                </tr>
                <tr className="hover:bg-slate-50/80">
                  <td className="py-1.5 px-4 text-slate-700">Preliminary Expenses</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfPreliminaryExpenses)}</td>)}
                </tr>
                {(isElemVisible('cf_pur_ppe') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_pur_ppe') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-1.5 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_pur_ppe')}
                        <span>{lbl('cf', 'purchasePPE', 'Purchase Property, Plant & Equipment')}</span>
                      </div>
                    </td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfPurchasePPE)}</td>)}
                  </tr>
                )}
                <tr className="hover:bg-slate-50/80">
                  <td className="py-1.5 px-4 text-slate-700">Purchase Investment In Debt Or Securities</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfPurchaseInvestment)}</td>)}
                </tr>
                {customCfInv.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-1.5 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {projData.map(d => (
                      <td key={d.year} className="py-1.5 px-3 text-right font-mono text-indigo-800 border-l border-slate-200/80">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-sky-50/75 font-bold text-sky-900">
                  <td className="py-2 px-4">Net Cash Used by Investing Activities</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-200/80">{fmtCell(d.cfNetCashInvesting)}</td>)}
                </tr>

                <tr className="bg-slate-50 font-bold text-amber-900">
                  <td colSpan={projData.length + 1} className="py-2 px-4 underline">3. CASH FLOW FROM FINANCIAL ACTIVITIES:</td>
                </tr>
                <tr className="font-semibold text-slate-800">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Inflows:</td>
                </tr>
                {(isElemVisible('cf_borrow_proceeds') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_borrow_proceeds') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-1.5 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_borrow_proceeds')}
                        <span>{lbl('cf', 'financingProceedsBorrowings', 'Loan Received')}</span>
                      </div>
                    </td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfLoanReceived)}</td>)}
                  </tr>
                )}
                {(isElemVisible('cf_shares_issue') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_shares_issue') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-1.5 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_shares_issue')}
                        <span>Additional Capital</span>
                      </div>
                    </td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfAdditionalCapital)}</td>)}
                  </tr>
                )}
                <tr className="hover:bg-slate-50/80">
                  <td className="py-1.5 px-4 text-slate-700">Issuance Of Debt (Bond &amp; Notes)</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfIssuanceOfDebt)}</td>)}
                </tr>
                <tr className="font-semibold text-slate-800">
                  <td colSpan={projData.length + 1} className="py-1 px-4 underline">Outflows:</td>
                </tr>
                {(isElemVisible('cf_dividends_paid') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_dividends_paid') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-1.5 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_dividends_paid')}
                        <span>To Stockholders as Dividend</span>
                      </div>
                    </td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfDividendsToStockholders)}</td>)}
                  </tr>
                )}
                {(isElemVisible('cf_dividends_paid') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_dividends_paid') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-1.5 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_dividends_paid')}
                        <span>{lbl('cf', 'financingDividendsPaid', 'Drawings / Dividend')}</span>
                      </div>
                    </td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfDrawingsOutflow)}</td>)}
                  </tr>
                )}
                {(isElemVisible('cf_borrow_repay') || isCustomizeMode) && (
                  <tr className={!isElemVisible('cf_borrow_repay') ? 'opacity-45 bg-slate-50' : 'hover:bg-slate-50/80'}>
                    <td className="py-1.5 px-4 text-slate-700">
                      <div className="flex items-center gap-2">
                        {renderCheck('cf_borrow_repay')}
                        <span>{lbl('cf', 'repaymentOfBorrowings', 'To Redeem Long Term Debt (Treasury Stock)')}</span>
                      </div>
                    </td>
                    {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfRedeemLongTermDebt)}</td>)}
                  </tr>
                )}
                <tr className="hover:bg-slate-50/80">
                  <td className="py-1.5 px-4 text-slate-700">Issuance Of Bonds Payable</td>
                  {projData.map(d => <td key={d.year} className="py-1.5 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfIssuanceBondsPayable)}</td>)}
                </tr>
                {customCfFin.filter(ci => ci.isVisible || isCustomizeMode).map(ci => (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-1.5 px-4 text-indigo-900">
                      <div className="flex items-center gap-2">
                        {renderCheck(ci.id)}
                        <span>{ci.label}</span>
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 py-0.5 rounded">Custom</span>
                      </div>
                    </td>
                    {projData.map(d => (
                      <td key={d.year} className="py-1.5 px-3 text-right font-mono text-indigo-800 border-l border-slate-200/80">{fmtCell(getCustomItemYearValue(ci, d.year))}</td>
                    ))}
                  </tr>
                ))}
                <tr className="bg-amber-50/75 font-bold text-amber-900">
                  <td className="py-2 px-4">Net Cash Provided by Financial Activities</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-200/80">{fmtCell(d.cfNetCashFinancing)}</td>)}
                </tr>

                <tr className="font-bold text-slate-900 bg-slate-50/80">
                  <td className="py-2 px-4">Net Cash Increase</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono border-l border-slate-200/80">{fmtCell(d.cfNetCashIncrease)}</td>)}
                </tr>
                <tr className="hover:bg-slate-50/80">
                  <td className="py-2 px-4 text-slate-700">Opening Cash &amp; Bank Balance</td>
                  {projData.map(d => <td key={d.year} className="py-2 px-3 text-right font-mono text-slate-900 border-l border-slate-200/80">{fmtCell(d.cfOpeningCash)}</td>)}
                </tr>
                <tr className="bg-sky-50/85 font-bold text-sky-900 border-t border-b-4 border-double border-sky-600">
                  <td className="py-2.5 px-4">Closing Cash &amp; Bank Balance</td>
                  {projData.map(d => <td key={d.year} className="py-2.5 px-3 text-right font-mono border-l border-slate-200/80">{fmtCell(d.cfClosingCash)}</td>)}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 4. NOTES TO THE FINANCIAL STATEMENTS (1, 2.1-2.6, 3.00-13.00) */}
      {(subTab === 'all' || subTab === 'notes') && (
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-slate-50/75">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-[11px] text-slate-600">{companyProfile.registeredAddress}</div>
            <div className="text-sm font-bold text-sky-800">Notes to the Financial Statements</div>
            <div className="text-xs text-slate-600">For the year ended 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="p-5 space-y-4 text-xs border-b border-slate-200 text-slate-700">
            <div>
              <div className="font-bold text-slate-900">1 Reporting Entity:</div>
              <p className="text-slate-600 mt-1">
                {companyProfile.companyName} is a {companyProfile.entityType.replace(/_/g, ' ')} entity and its office is situated at {companyProfile.registeredAddress}.
              </p>
            </div>
            <div className="space-y-2">
              <div className="font-bold text-slate-900">2 SIGNIFICANT ACCOUNTING POLICIES:</div>
              <p><strong className="text-slate-900">2.1 Basis of Accounting:</strong> The Financial Statements have been prepared on a going concern basis using accrual method of accounting in accordance with {companyProfile.accountingFramework}.</p>
              <p><strong className="text-slate-900">2.2 Accounting Policies:</strong> Selected and applied consistently for significant transactions and events to all periods presented.</p>
              <p><strong className="text-slate-900">2.3 Property, Plant and Equipment:</strong> Stated at cost less accumulated depreciation. Depreciation is computed over estimated useful lives consistently.</p>
              <p><strong className="text-slate-900">2.4 Reporting Period:</strong> Financial Statements cover the annual reporting cycle ending {companyProfile.financialYearEnd}.</p>
              <p><strong className="text-slate-900">2.5 Reporting currency and level of precision:</strong> Figures represent {companyProfile.reportingCurrency} ({companyProfile.reportingUnit}) rounded cleanly.</p>
              <p><strong className="text-slate-900">2.6 Comparative Information:</strong> Disclosed for all selected years for numerical information and descriptive narratives.</p>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-100/95 border-b border-slate-200 text-slate-900">
                  <th rowSpan={2} className="py-2 px-3 text-center font-bold border-r border-slate-200 w-20">Note No.</th>
                  <th rowSpan={2} className="py-2 px-4 text-left font-bold border-r border-slate-200 min-w-[260px]">Particulars</th>
                  {projData.map(d => (
                    <th key={`nt-amt-${d.year}`} className="py-1.5 px-3 text-center font-semibold border-r border-b border-slate-200 last:border-r-0 min-w-[120px] text-[11px]">
                      {amtLabel}
                    </th>
                  ))}
                </tr>
                <tr className="bg-slate-50/90 border-b border-slate-200 text-sky-800 font-mono">
                  {projData.map(d => (
                    <th key={`nt-yr-${d.year}`} className="py-1.5 px-3 text-center font-bold border-r border-slate-200 last:border-r-0">{d.year}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/80 text-slate-800">
                {buildProjectedQuantNoteRows(projData, {
                  statementConfig,
                  labelOverrides,
                  includeHidden: isCustomizeMode,
                }).map((row, idx) => {
                  const visible = row.configId
                    ? row.noteNo === '12.00'
                      ? isElemVisible('inc_operatingExpenses') || isElemVisible('inc_depreciationExpense')
                      : isElemVisible(row.configId)
                    : true;
                  const rowCls = !visible
                    ? 'opacity-45 bg-slate-50'
                    : row.isSectionHeader
                      ? 'font-bold text-sky-900 bg-sky-50/50'
                      : row.isSubtotal
                        ? 'font-bold text-slate-900 bg-slate-100/75'
                        : 'hover:bg-slate-50/80';
                  return (
                    <tr key={`qnote-${idx}-${row.noteNo}-${row.label}`} className={rowCls}>
                      <td className="py-1.5 px-3 text-center font-mono">{row.noteNo}</td>
                      {row.isSectionHeader && row.values.every(v => v === null) ? (
                        <td colSpan={projData.length + 1} className={`py-1.5 px-4 ${row.indent ? 'pl-6' : ''}`}>
                          <div className="flex items-center gap-2">
                            {row.noteNo !== '' && row.configId && renderCheck(row.configId)}
                            <span>{row.label}</span>
                          </div>
                        </td>
                      ) : (
                        <>
                          <td className={`py-1.5 px-4 ${row.indent ? 'pl-6' : ''} ${!row.isBold ? 'text-slate-600' : ''}`}>
                            <div className="flex items-center gap-2">
                              {row.noteNo !== '' && row.configId && renderCheck(row.configId)}
                              <span>{row.label}</span>
                            </div>
                          </td>
                          {row.values.map((v, yIdx) => (
                            <td
                              key={projData[yIdx]?.year ?? yIdx}
                              className={`py-1.5 px-3 text-right font-mono ${row.isBold ? 'text-slate-900' : 'text-slate-800'}`}
                            >
                              {v === null ? '' : fmtCell(v)}
                            </td>
                          ))}
                        </>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 5. SCHEDULE OF PROPERTY PLANT & EQUIPMENTS (Schedule-A) */}
      {(subTab === 'all' || subTab === 'sched') && (
        <div className="bg-white/96 backdrop-blur-md border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
          <div className="p-5 border-b border-slate-200 text-center space-y-1 bg-slate-50/75">
            <h3 className="text-base font-extrabold text-slate-900 tracking-wide">{companyProfile.companyName}</h3>
            <div className="text-[11px] text-slate-600">{companyProfile.registeredAddress}</div>
            <div className="text-sm font-bold text-sky-800">Schedule of Property Plant &amp; Equipments</div>
            <div className="text-xs text-slate-600">As at 31st December, {maxYear}(Projected)</div>
          </div>

          <div className="p-4 space-y-4 overflow-x-auto">
            {projAsc.map(d => (
              <table key={d.year} className="w-full text-xs border border-slate-200 border-collapse tabular-nums">
                <thead>
                  <tr className="bg-slate-100/95 text-slate-900 border-b border-slate-200">
                    <th rowSpan={2} className="py-2 px-2 text-center border-r border-slate-200 w-14">SL No.</th>
                    <th rowSpan={2} className="py-2 px-3 text-left border-r border-slate-200 min-w-[200px]">Particulars</th>
                    <th colSpan={4} className="py-1.5 px-3 text-center border-r border-b border-slate-200">Cost</th>
                    <th rowSpan={2} className="py-2 px-2 text-center border-r border-slate-200 w-24">Rate of Dep. (%)</th>
                    <th rowSpan={2} className="py-2 px-3 text-right border-r border-slate-200">Depreciation Charged During the Year</th>
                    <th rowSpan={2} className="py-2 px-3 text-right">W.D.V as on 31.12.{d.year}</th>
                  </tr>
                  <tr className="bg-slate-50 text-slate-700 border-b border-slate-200 text-[11px]">
                    <th className="py-1.5 px-3 text-right border-r border-slate-200">Opening Balance</th>
                    <th className="py-1.5 px-3 text-right border-r border-slate-200">Addition During the Year</th>
                    <th className="py-1.5 px-3 text-right border-r border-slate-200">Disposals During the Year</th>
                    <th className="py-1.5 px-3 text-right border-r border-slate-200">Balance as on 31.12.{d.year}</th>
                  </tr>
                </thead>
                <tbody className="text-slate-800">
                  <tr className="border-b border-slate-200 hover:bg-slate-50/70">
                    <td className="py-2 px-2 text-center font-mono text-slate-600 border-r border-slate-200">1</td>
                    <td className="py-2 px-3 text-slate-900 border-r border-slate-200">Property Plant &amp; Equipment</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-900 border-r border-slate-200">{fmtCell(d.schedOpeningBalance)}</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-900 border-r border-slate-200">{fmtCell(d.schedAddition)}</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-900 border-r border-slate-200">{fmtCell(d.schedDisposals)}</td>
                    <td className="py-2 px-3 text-right font-mono text-emerald-800 font-semibold border-r border-slate-200">
                      {onUpdatePrimaryDriver ? (
                        <InlineDriverCell
                          value={d.schedGrossAssetBase}
                          isEditMode={isEditMode}
                          onSave={val => onUpdatePrimaryDriver(d.year, 'fixedAssetsGross', val)}
                          title="Fixed User Input: Fixed Assets (Gross Asset Base)"
                        />
                      ) : fmtCell(d.schedBalanceClosing)}
                    </td>
                    <td className="py-2 px-2 text-center font-mono text-slate-700 border-r border-slate-200">{d.schedRatePct}%</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-900 border-r border-slate-200">{fmtCell(d.schedDepreciation)}</td>
                    <td className="py-2 px-3 text-right font-mono text-slate-900 font-semibold">{fmtCell(d.schedWDV)}</td>
                  </tr>
                  <tr className="bg-sky-50/75 font-bold text-sky-900">
                    <td className="py-2 px-2 border-r border-slate-200"></td>
                    <td className="py-2 px-3 text-center border-r border-slate-200">Total</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.schedOpeningBalance)}</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.schedAddition)}</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.schedDisposals)}</td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.schedBalanceClosing)}</td>
                    <td className="py-2 px-2 border-r border-slate-200"></td>
                    <td className="py-2 px-3 text-right font-mono border-r border-slate-200">{fmtCell(d.schedDepreciation)}</td>
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

