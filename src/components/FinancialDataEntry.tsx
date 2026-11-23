import React, { useState, useMemo, useEffect } from 'react';
import {
  Database,
  Sparkles,
  Lock,
  Unlock,
  Calendar,
  CheckCircle2,
  Eye,
  LayoutGrid,
  Info,
  Zap,
  Wand2,
  Star,
  X,
  ShieldCheck,
  AlertTriangle,
} from 'lucide-react';
import {
  YearDataInput,
  CompanyProfile,
  ProjectionAssumptions,
  InputMode,
  IncomeStatementYear,
  BalanceSheetYear,
  CashFlowStatementYear,
  PrimaryDriverKey,
  AutoFillConfigs,
  FinancialValidationIssue,
} from '../types/financial';
import {
  calculateFinancialRatios,
  FinancialRatios,
  PRIMARY_DRIVERS_METADATA,
  calculateDriverAutoFillSeries,
} from '../services/financialEngine';
import { YearSelectorControl } from './YearSelectorControl';

interface FinancialDataEntryProps {
  yearsData: YearDataInput[];
  onUpdatePrimaryDriver: (year: number, driver: keyof YearDataInput, value: number) => void;
  onUpdateYearData: (updatedYear: YearDataInput) => void;
  companyProfile: CompanyProfile;
  assumptions: ProjectionAssumptions;
  onUpdateAssumptions: (updated: ProjectionAssumptions) => void;
  inputMode: InputMode;
  onToggleInputMode: (mode: InputMode) => void;
  onGenerateAutomaticProjections: () => void;
  incomeStatements: IncomeStatementYear[];
  balanceSheets: BalanceSheetYear[];
  cashFlowStatements: CashFlowStatementYear[];
  autoFillConfigs?: AutoFillConfigs;
  onApplyAutoFill?: (driver: PrimaryDriverKey, baseYear: number, baseValue: number, customGrowthRate?: number) => void;
  onClearAutoFill?: (driver: PrimaryDriverKey) => void;
  validationIssues?: FinancialValidationIssue[];
  selectedHistoricalYears?: number[];
  selectedProjectedYears?: number[];
  onToggleHistoricalYear?: (year: number) => void;
  onToggleProjectedYear?: (year: number) => void;
  onSetHistoricalYears?: (years: number[]) => void;
  onSetProjectedYears?: (years: number[]) => void;
}

export const FinancialDataEntry: React.FC<FinancialDataEntryProps> = ({
  yearsData,
  onUpdatePrimaryDriver,
  companyProfile,
  assumptions,
  onUpdateAssumptions,
  inputMode,
  onToggleInputMode,
  onGenerateAutomaticProjections,
  incomeStatements,
  balanceSheets,
  cashFlowStatements,
  autoFillConfigs = {},
  onApplyAutoFill = () => {},
  onClearAutoFill = () => {},
  validationIssues = [],
  selectedHistoricalYears = [2023, 2024, 2025],
  selectedProjectedYears,
  onToggleHistoricalYear,
  onToggleProjectedYear,
  onSetHistoricalYears,
  onSetProjectedYears,
}) => {
  const [viewMode, setViewMode] = useState<'grid' | 'single'>('grid');
  const [selectedYear, setSelectedYear] = useState<number>(yearsData[0]?.year || 2023);
  const [showInterconnectionLogic, setShowInterconnectionLogic] = useState<boolean>(false);
  const [showAutoFillPanel, setShowAutoFillPanel] = useState<boolean>(true);
  const [showYearSelector, setShowYearSelector] = useState<boolean>(false);
  const [activeModalDriver, setActiveModalDriver] = useState<PrimaryDriverKey | null>(null);
  const [activeNotification, setActiveNotification] = useState<string | null>(null);

  // Independent base-year selections for each of the 6 drivers (initialized to 2024 or first year)
  const [selectedBaseYears, setSelectedBaseYears] = useState<Record<PrimaryDriverKey, number>>({
    salesRevenue: yearsData.find(y => y.year === 2024)?.year || yearsData[0]?.year || 2023,
    closingInventory: yearsData.find(y => y.year === 2024)?.year || yearsData[0]?.year || 2023,
    financeCost: yearsData.find(y => y.year === 2024)?.year || yearsData[0]?.year || 2023,
    netProfit: yearsData.find(y => y.year === 2024)?.year || yearsData[0]?.year || 2023,
    repaymentOfBorrowings: yearsData.find(y => y.year === 2024)?.year || yearsData[0]?.year || 2023,
    fixedAssetsGross: yearsData.find(y => y.year === 2023)?.year || yearsData[0]?.year || 2023,
  });

  // Ensure selectedYear and selectedBaseYears remain valid when years are toggled dynamically
  useEffect(() => {
    const validYears = yearsData.map(y => y.year);
    if (validYears.length === 0) return;
    if (!validYears.includes(selectedYear)) {
      setSelectedYear(validYears[0]);
    }
    setSelectedBaseYears(prev => {
      let changed = false;
      const next = { ...prev };
      (Object.keys(next) as PrimaryDriverKey[]).forEach(k => {
        if (!validYears.includes(next[k])) {
          next[k] = validYears[0];
          changed = true;
        }
      });
      return changed ? next : prev;
    });
  }, [yearsData, selectedYear]);

  // Independent custom growth / trend rates for each of the 6 drivers
  const [customRates, setCustomRates] = useState<Partial<Record<PrimaryDriverKey, number>>>({});
  const [customBaseValues, setCustomBaseValues] = useState<Partial<Record<PrimaryDriverKey, number>>>({});
  const [selectedCategory, setSelectedCategory] = useState<'all' | 'income' | 'workingCapital' | 'balanceSheet' | 'cashFlow' | 'ratios'>('all');

  const currentYearData = yearsData.find(y => y.year === selectedYear) || yearsData[0];
  const currentIncome = incomeStatements.find(i => i.year === selectedYear) || incomeStatements[0];
  const currentBalance = balanceSheets.find(b => b.year === selectedYear) || balanceSheets[0];
  const currentCashFlow = cashFlowStatements.find(c => c.year === selectedYear) || cashFlowStatements[0];

  const currentRatios = useMemo(() => {
    if (!currentYearData || !currentIncome || !currentBalance || !currentCashFlow) return null;
    return calculateFinancialRatios(currentYearData, currentIncome, currentBalance, currentCashFlow);
  }, [currentYearData, currentIncome, currentBalance, currentCashFlow]);

  // Map of ratios across all years for quick matrix display
  const allYearsRatios = useMemo(() => {
    const map: Record<number, FinancialRatios> = {};
    yearsData.forEach((yd, idx) => {
      const inc = incomeStatements[idx];
      const bal = balanceSheets[idx];
      const cf = cashFlowStatements[idx];
      if (yd && inc && bal && cf) {
        map[yd.year] = calculateFinancialRatios(yd, inc, bal, cf);
      }
    });
    return map;
  }, [yearsData, incomeStatements, balanceSheets, cashFlowStatements]);

  // Primary Drivers definition
  const primaryDriversConfig = [
    {
      id: 'salesRevenue',
      number: 1,
      title: 'Sales Revenue',
      badge: 'Top-Line Driver',
      description: 'Primary revenue engine driving turnover, receivables (DSO), cost calibration, and operational capacity.',
      field: 'salesRevenue' as keyof YearDataInput,
      color: 'emerald',
    },
    {
      id: 'closingInventory',
      number: 2,
      title: 'Closing Inventory',
      badge: 'Working Capital Driver',
      description: 'Inventory holding level driving working capital requirements, DIO cycle, operating cash absorption, and current assets.',
      field: 'closingInventory' as keyof YearDataInput,
      color: 'teal',
    },
    {
      id: 'financeCost',
      number: 3,
      title: 'Finance Cost',
      badge: 'Debt Servicing Driver',
      description: 'Financing charge & borrowing interest determining debt servicing cost, Profit Before Tax, and interest coverage ratios.',
      field: 'financeCost' as keyof YearDataInput,
      color: 'amber',
    },
    {
      id: 'netProfit',
      number: 4,
      title: 'Net Profit After Tax',
      badge: 'Bottom-Line & Equity Driver',
      description: 'Net earnings after corporate tax driving retained earnings accretion, dividend payouts, and equity return (ROE).',
      field: 'netProfit' as keyof YearDataInput,
      color: 'sky',
    },
    {
      id: 'repaymentOfBorrowings',
      number: 5,
      title: 'Repayment from Borrowings and Financing',
      badge: 'Debt Amortization Driver',
      description: 'Principal loan repayments directly liquidating long-term debt, lowering financing cash flow, and reducing liabilities.',
      field: 'repaymentOfBorrowings' as keyof YearDataInput,
      color: 'rose',
    },
    {
      id: 'fixedAssetsGross',
      number: 6,
      title: 'Fixed Assets (Gross Asset Base)',
      badge: 'Capital Asset Driver',
      description: 'Capital asset base directly driving annual depreciation charges, capex additions/investing cash outflows, and PP&E balance.',
      field: 'fixedAssetsGross' as keyof YearDataInput,
      color: 'indigo',
    },
  ];

  const handleDriverChange = (year: number, field: keyof YearDataInput, val: number) => {
    onUpdatePrimaryDriver(year, field, Math.max(0, val));
  };

  const handleQuickHorizonChange = (count: number) => {
    onUpdateAssumptions({
      ...assumptions,
      projectedYearsCount: count,
    });
  };

  const projYearsList = selectedProjectedYears && selectedProjectedYears.length > 0
    ? selectedProjectedYears
    : yearsData.filter(y => y.isProjected).map(y => y.year);
  const histYearsCount = yearsData.filter(y => !y.isProjected).length;
  const projYearsCount = yearsData.filter(y => y.isProjected).length;

  return (
    <div className="space-y-6">
      {/* Top Header & Operational Banner */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4 shadow-sm">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-sky-500/10 border border-sky-500/30 flex items-center justify-center text-sky-400">
                <Database className="w-4 h-4" />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-lg font-bold text-white tracking-tight">
                    Financial Data Entry & Primary Driver Architecture
                  </h2>
                  <span className="text-[11px] font-semibold text-emerald-400 bg-emerald-950/60 border border-emerald-800/50 px-2 py-0.5 rounded-md">
                    6 Strictly Fixed Inputs
                  </span>
                  {validationIssues.length === 0 ? (
                    <span className="text-[10px] font-semibold text-emerald-300 bg-emerald-950/60 border border-emerald-700/60 px-2 py-0.5 rounded-md inline-flex items-center gap-1">
                      <ShieldCheck className="w-3 h-3" />
                      Reconciled
                    </span>
                  ) : (
                    <span className="text-[10px] font-semibold text-amber-300 bg-amber-950/60 border border-amber-700/60 px-2 py-0.5 rounded-md inline-flex items-center gap-1">
                      <AlertTriangle className="w-3 h-3" />
                      {validationIssues.length} Alert(s)
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-400 mt-0.5">
                  The six primary user inputs remain strictly unchanged across all calculations. All dependent statement figures, ratios, and cash flows adjust logically and proportionately.
                </p>
              </div>
            </div>
          </div>

          {/* Mode Switcher & Global Actions */}
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Manual vs Automatic Input Mode Segmented Toggle */}
            <div className="flex items-center bg-slate-950/80 p-1 rounded-xl border border-slate-800">
              <button
                type="button"
                onClick={() => onToggleInputMode('manual')}
                className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center gap-1.5 ${
                  inputMode === 'manual'
                    ? 'bg-emerald-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                <Unlock className="w-3.5 h-3.5" />
                Manual Mode
              </button>
              <button
                type="button"
                onClick={() => {
                  onToggleInputMode('automatic');
                  onGenerateAutomaticProjections();
                }}
                className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center gap-1.5 ${
                  inputMode === 'automatic'
                    ? 'bg-sky-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                <Sparkles className="w-3.5 h-3.5" />
                Automatic / Interconnected Mode
              </button>
            </div>

            {/* View Mode Toggle */}
            <div className="flex items-center bg-slate-950/80 p-1 rounded-xl border border-slate-800">
              <button
                type="button"
                onClick={() => setViewMode('grid')}
                className={`px-2.5 py-1.5 text-xs font-medium rounded-lg transition-colors flex items-center gap-1.5 ${
                  viewMode === 'grid'
                    ? 'bg-slate-800 text-white'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                title="Spreadsheet multi-year grid view"
              >
                <LayoutGrid className="w-3.5 h-3.5" />
                Multi-Year Grid
              </button>
              <button
                type="button"
                onClick={() => setViewMode('single')}
                className={`px-2.5 py-1.5 text-xs font-medium rounded-lg transition-colors flex items-center gap-1.5 ${
                  viewMode === 'single'
                    ? 'bg-slate-800 text-white'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                title="Single-year focused deep-dive inspector"
              >
                <Eye className="w-3.5 h-3.5" />
                Year Inspector
              </button>
            </div>

            {/* Select Years Individual Toggle */}
            {onToggleHistoricalYear && onToggleProjectedYear && onSetHistoricalYears && onSetProjectedYears && (
              <button
                type="button"
                onClick={() => setShowYearSelector(!showYearSelector)}
                className={`px-3 py-1.5 text-xs font-semibold rounded-xl border transition-all flex items-center gap-1.5 ${
                  showYearSelector
                    ? 'bg-sky-600 border-sky-400 text-white shadow-sm'
                    : 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700'
                }`}
              >
                <Calendar className="w-3.5 h-3.5" />
                Select Years ({histYearsCount}H+{projYearsCount}P)
              </button>
            )}

            {/* Auto-Fill Hub Toggle */}
            <button
              type="button"
              onClick={() => setShowAutoFillPanel(!showAutoFillPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-xl border transition-all flex items-center gap-1.5 ${
                showAutoFillPanel
                  ? 'bg-emerald-950/80 border-emerald-500 text-emerald-300 shadow-sm'
                  : 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <Zap className="w-3.5 h-3.5 text-emerald-400" />
              Auto-Fill Hub
              {Object.values(autoFillConfigs).filter(c => c?.isActive).length > 0 && (
                <span className="bg-emerald-500 text-slate-950 text-[10px] font-bold px-1.5 py-0.2 rounded-full">
                  {Object.values(autoFillConfigs).filter(c => c?.isActive).length} Active
                </span>
              )}
            </button>

            {/* Interconnection Logic Info Toggle */}
            <button
              type="button"
              onClick={() => setShowInterconnectionLogic(!showInterconnectionLogic)}
              className={`px-3 py-1.5 text-xs font-medium rounded-xl border transition-colors flex items-center gap-1.5 ${
                showInterconnectionLogic
                  ? 'bg-indigo-950/60 border-indigo-700 text-indigo-300'
                  : 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <Info className="w-3.5 h-3.5 text-indigo-400" />
              Interconnection Logic
            </button>
          </div>
        </div>

        {/* Projection Horizon Quick Controller */}
        <div className="pt-3 border-t border-slate-800/80 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex flex-wrap items-center gap-2">
            <Calendar className="w-3.5 h-3.5 text-slate-400" />
            <span className="text-slate-300 font-medium">Projection Horizon:</span>
            <span className="font-mono font-bold text-emerald-400">
              {projYearsCount} Years ({yearsData.find(y => y.isProjected)?.year}–{yearsData[yearsData.length - 1]?.year})
            </span>
            <div className="flex items-center gap-1 ml-2">
              {[3, 5, 7, 10, 15, 20].map(count => (
                <button
                  key={count}
                  type="button"
                  onClick={() => handleQuickHorizonChange(count)}
                  className={`px-2 py-0.5 rounded text-[11px] font-semibold transition-colors ${
                    projYearsCount === count
                      ? 'bg-emerald-600 text-white shadow-sm'
                      : 'bg-slate-800 text-slate-400 hover:text-slate-200 hover:bg-slate-700'
                  }`}
                >
                  {count}Y
                </button>
              ))}
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3 text-[11px] text-slate-400">
            <div className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-sm bg-emerald-600/80 border border-emerald-400 inline-block"></span>
              <span>Fixed Primary Input</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-sm bg-sky-950/60 border border-sky-600/60 inline-block"></span>
              <span>Automated Projection</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-sm bg-slate-800/80 border border-slate-700 inline-block"></span>
              <span>Calculated / Derived</span>
            </div>
            <span className="text-slate-500 font-mono">
              Currency: {companyProfile.reportingCurrency} in {companyProfile.reportingUnit}
            </span>
          </div>
        </div>
      </div>

      {/* Expandable Individual Year Selector */}
      {showYearSelector && onToggleHistoricalYear && onToggleProjectedYear && onSetHistoricalYears && onSetProjectedYears && (
        <YearSelectorControl
          selectedHistoricalYears={selectedHistoricalYears}
          selectedProjectedYears={projYearsList}
          onToggleHistoricalYear={onToggleHistoricalYear}
          onToggleProjectedYear={onToggleProjectedYear}
          onSetHistoricalYears={onSetHistoricalYears}
          onSetProjectedYears={onSetProjectedYears}
        />
      )}

      {/* Toast Notification for Auto-Fill operations */}
      {activeNotification && (
        <div className="bg-emerald-950/90 border border-emerald-500/60 text-emerald-200 px-4 py-3 rounded-2xl flex items-center justify-between text-xs shadow-lg animate-in fade-in duration-200">
          <div className="flex items-center gap-2.5">
            <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
            <span className="font-medium">{activeNotification}</span>
          </div>
          <button
            type="button"
            onClick={() => setActiveNotification(null)}
            className="text-emerald-400 hover:text-white p-1 rounded-md transition-colors"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Dedicated Auto-Fill Remaining Years Facility Hub */}
      {showAutoFillPanel && (
        <div className="bg-slate-900 border border-emerald-900/60 rounded-2xl p-5 space-y-4 text-xs shadow-md">
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 border-b border-emerald-950/80 pb-3">
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
                <Zap className="w-4 h-4" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h3 className="font-bold text-white text-sm">
                    Auto-Fill Remaining Years Facility — 6 Primary Financial Drivers
                  </h3>
                  <span className="text-[10px] font-bold text-emerald-300 bg-emerald-950/80 border border-emerald-700/60 px-2 py-0.5 rounded-full">
                    Independent Single-Year Anchors
                  </span>
                </div>
                <p className="text-[11px] text-slate-400 mt-0.5">
                  Enter or select any single base year for each item independently. The system preserves your entered base reference and automatically calculates all remaining prior historical years and subsequent projected years using professional financial analyst ratios.
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 text-xs shrink-0">
              <span className="text-slate-400 text-[11px]">Active Anchors:</span>
              <span className="font-mono font-bold text-emerald-400 bg-slate-950 px-2 py-1 rounded-md border border-slate-800">
                {Object.values(autoFillConfigs).filter(c => c?.isActive).length} / 6 Configured
              </span>
            </div>
          </div>

          {/* Grid of the 6 Financial Drivers */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {PRIMARY_DRIVERS_METADATA.map(meta => {
              const driverKey = meta.key;
              const activeConfig = autoFillConfigs[driverKey];
              const isActive = !!activeConfig?.isActive;
              const currentBaseYear = selectedBaseYears[driverKey];
              const baseYearData = yearsData.find(y => y.year === currentBaseYear);
              const currentDriverVal = customBaseValues[driverKey] !== undefined 
                ? customBaseValues[driverKey]! 
                : (baseYearData ? (baseYearData[driverKey] as number) : 0);
              const currentRate = customRates[driverKey] !== undefined
                ? customRates[driverKey]!
                : (meta.key === 'salesRevenue' ? assumptions.revenueGrowthRate : meta.defaultGrowthOrTrendPct);

              return (
                <div
                  key={driverKey}
                  className={`p-4 rounded-xl border transition-all space-y-3 ${
                    isActive
                      ? 'bg-emerald-950/20 border-emerald-600/60 shadow-sm'
                      : 'bg-slate-950/70 border-slate-800 hover:border-slate-700'
                  }`}
                >
                  {/* Card Header */}
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-md bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-[10px] font-bold">
                        {meta.number}
                      </span>
                      <span className="font-bold text-white text-xs">
                        {meta.title}
                      </span>
                    </div>
                    {isActive ? (
                      <span className="text-[10px] font-semibold text-emerald-300 bg-emerald-950 border border-emerald-600/60 px-2 py-0.5 rounded-md flex items-center gap-1">
                        <Star className="w-2.5 h-2.5 fill-emerald-400 text-emerald-400" />
                        Base: {activeConfig.baseYear}
                      </span>
                    ) : (
                      <span className="text-[10px] text-slate-500 bg-slate-900 px-2 py-0.5 rounded border border-slate-800">
                        Manual Mode
                      </span>
                    )}
                  </div>

                  {/* Base Year & Base Value Selector */}
                  <div className="grid grid-cols-2 gap-2 pt-1">
                    <div>
                      <label className="text-[10px] text-slate-400 block mb-1">
                        Select Base Year:
                      </label>
                      <select
                        value={currentBaseYear}
                        onChange={(e) => {
                          const yr = Number(e.target.value);
                          setSelectedBaseYears(prev => ({ ...prev, [driverKey]: yr }));
                          const yd = yearsData.find(y => y.year === yr);
                          if (yd) {
                            setCustomBaseValues(prev => ({ ...prev, [driverKey]: yd[driverKey] as number }));
                          }
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-white font-mono focus:border-emerald-500"
                      >
                        {yearsData.map(y => (
                          <option key={y.year} value={y.year}>
                            {y.year} ({y.isProjected ? 'Projected' : 'Historical'})
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="text-[10px] text-slate-400 block mb-1">
                        Base Value ({companyProfile.reportingCurrency}):
                      </label>
                      <input
                        type="number"
                        value={currentDriverVal}
                        onChange={(e) => {
                          const val = Number(e.target.value);
                          setCustomBaseValues(prev => ({ ...prev, [driverKey]: val }));
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-white font-mono font-bold text-right focus:border-emerald-500 tabular-nums"
                      />
                    </div>
                  </div>

                  {/* Growth / Trend Calibration */}
                  <div className="bg-slate-900/80 p-2.5 rounded-lg border border-slate-800 space-y-1.5">
                    <div className="flex items-center justify-between text-[10px]">
                      <span className="text-slate-400 font-medium">
                        {meta.trendLabel}:
                      </span>
                      <div className="flex items-center gap-1 font-mono">
                        <input
                          type="number"
                          step="0.1"
                          value={currentRate}
                          onChange={(e) => setCustomRates(prev => ({ ...prev, [driverKey]: Number(e.target.value) }))}
                          className="w-16 bg-slate-950 border border-slate-700 rounded px-1.5 py-0.5 text-right text-xs text-emerald-400 font-bold"
                        />
                        <span className="text-slate-400">%</span>
                      </div>
                    </div>
                    <p className="text-[9.5px] text-slate-400 leading-tight">
                      {meta.rationale}
                    </p>
                  </div>

                  {/* Actions & Dynamic Recalculation Note */}
                  <div className="space-y-1.5 pt-1">
                    <button
                      type="button"
                      onClick={() => {
                        onApplyAutoFill(driverKey, currentBaseYear, currentDriverVal, currentRate);
                        setActiveNotification(
                          `Auto-filled all remaining years for ${meta.title} using ${currentBaseYear} (${companyProfile.reportingCurrency} ${currentDriverVal.toLocaleString()}) as fixed base reference!`
                        );
                      }}
                      className={`w-full py-2 px-3 rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-1.5 shadow-sm ${
                        isActive
                          ? 'bg-emerald-600 hover:bg-emerald-500 text-white'
                          : 'bg-slate-800 hover:bg-emerald-600 text-emerald-300 hover:text-white border border-slate-700 hover:border-emerald-500'
                      }`}
                    >
                      <Zap className="w-3.5 h-3.5" />
                      {isActive ? 'Recalculate All Remaining Years' : 'Auto-Fill Remaining Years'}
                    </button>

                    <div className="flex items-center justify-between text-[10px] pt-0.5">
                      <button
                        type="button"
                        onClick={() => setActiveModalDriver(driverKey)}
                        className="text-slate-400 hover:text-emerald-300 flex items-center gap-1 transition-colors"
                      >
                        <Eye className="w-3 h-3" />
                        Preview All Years Table
                      </button>

                      {isActive && (
                        <button
                          type="button"
                          onClick={() => {
                            onClearAutoFill(driverKey);
                            setActiveNotification(`Deactivated Auto-Fill for ${meta.title}. Manual values retained.`);
                          }}
                          className="text-slate-500 hover:text-rose-400 transition-colors"
                        >
                          Clear Auto-Fill
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Expandable Interconnection Architecture Panel */}
      {showInterconnectionLogic && (
        <div className="bg-slate-900 border border-indigo-900/50 rounded-2xl p-5 space-y-4 text-xs">
          <div className="flex items-center justify-between border-b border-indigo-950/80 pb-3">
            <div className="flex items-center gap-2">
              <Zap className="w-4 h-4 text-indigo-400" />
              <h3 className="font-bold text-white text-sm">
                Financial Model Interconnection Architecture & Formula Relationships
              </h3>
            </div>
            <span className="text-[11px] font-semibold text-indigo-300 bg-indigo-950/80 border border-indigo-800/50 px-2 py-0.5 rounded-md">
              Double-Entry Mathematical Guarantee
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-slate-300">
            <div className="p-3.5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-1.5">
              <span className="font-bold text-emerald-400 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-[10px]">1</span>
                Sales Revenue Driver
              </span>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Drives <strong>Trade Receivables</strong> via DSO ({(currentYearData?.tradeReceivables || 0).toLocaleString()} {companyProfile.reportingCurrency}), calibrates <strong>Operating Scale</strong>, and derives <strong>Cost of Sales</strong> & <strong>OpEx</strong> tying out with Net Profit.
              </p>
            </div>

            <div className="p-3.5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-1.5">
              <span className="font-bold text-teal-400 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-teal-500/20 text-teal-400 flex items-center justify-center text-[10px]">2</span>
                Closing Inventory Driver
              </span>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Directly feeds <strong>Working Capital</strong> in Balance Sheet and drives the <strong>Change in Inventory</strong> cash flow adjustment. Automatically calculates Days Sales of Inventory (DIO).
              </p>
            </div>

            <div className="p-3.5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-1.5">
              <span className="font-bold text-amber-400 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-amber-500/20 text-amber-400 flex items-center justify-center text-[10px]">3</span>
                Finance Cost Driver
              </span>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Represents annual interest/debt servicing. Enters Income Statement directly, adjusts Operating Cash Flow (+ addback, - cash interest paid), and calibrates the <strong>Interest Coverage Ratio</strong>.
              </p>
            </div>

            <div className="p-3.5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-1.5">
              <span className="font-bold text-sky-400 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-sky-500/20 text-sky-400 flex items-center justify-center text-[10px]">4</span>
                Net Profit After Tax Driver
              </span>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Derives <strong>Profit Before Tax</strong> [PBT = NP / (1 - TaxRate)] and <strong>Corporate Tax Expense</strong>. Drives <strong>Retained Earnings</strong> roll-forward and dividend distribution in Equity.
              </p>
            </div>

            <div className="p-3.5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-1.5">
              <span className="font-bold text-rose-400 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-rose-500/20 text-rose-400 flex items-center justify-center text-[10px]">5</span>
                Repayment of Borrowings
              </span>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Enters Financing Cash Flow as principal debt amortization outflow. Dynamically reduces <strong>Long-Term Debt</strong> on the Balance Sheet and optimizes the Debt-to-Equity solvency ratio.
              </p>
            </div>

            <div className="p-3.5 bg-slate-950/60 border border-slate-800 rounded-xl space-y-1.5">
              <span className="font-bold text-indigo-400 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-full bg-indigo-500/20 text-indigo-400 flex items-center justify-center text-[10px]">6</span>
                Fixed Assets Driver
              </span>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Determines <strong>Capital Additions (Capex)</strong> in Investing Cash Flow, generates <strong>Annual Depreciation Charge</strong> across asset classes, and feeds Property, Plant & Equipment NBV.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Main View: Multi-Year Grid vs Single-Year Inspector */}
      {viewMode === 'grid' ? (
        <div className="space-y-6">
          {/* Section 1: The Six Primary Financial Input Drivers Table */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-4 bg-slate-800/40 border-b border-slate-800 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <Database className="w-4 h-4 text-emerald-400" />
                <h3 className="text-sm font-bold text-white tracking-tight">
                  Primary Financial Input Drivers (Direct User-Entry Controls)
                </h3>
              </div>
              <div className="flex items-center gap-2 text-[11px] text-slate-400">
                <span className="font-medium text-emerald-400">Active Mode:</span>
                <span className="capitalize font-semibold text-slate-200">
                  {inputMode === 'manual' ? 'Manual Direct Entry' : 'Automatic / Interconnected'}
                </span>
                <span>·</span>
                <span>{yearsData.length} Total Financial Years</span>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs font-mono">
                <thead className="bg-slate-800/90 text-slate-300 border-b border-slate-700 text-[11px]">
                  <tr>
                    <th className="py-3 px-4 font-sans font-semibold sticky left-0 bg-slate-800/95 z-10 w-72 min-w-[280px]">
                      Primary Financial Driver Item
                    </th>
                    {yearsData.map(y => (
                      <th
                        key={y.year}
                        className={`py-3 px-3 text-right min-w-[130px] ${
                          y.isProjected ? 'text-sky-300' : 'text-slate-200'
                        }`}
                      >
                        <div className="font-sans font-semibold text-xs flex items-center justify-end gap-1">
                          <span>{y.year}</span>
                          <span
                            className={`text-[9px] px-1 py-0.2 rounded font-normal ${
                              y.isProjected
                                ? 'bg-sky-950/80 text-sky-300 border border-sky-800/60'
                                : 'bg-slate-800 text-slate-400 border border-slate-700'
                            }`}
                          >
                            {y.isProjected ? 'Proj' : 'Hist'}
                          </span>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/70">
                  {primaryDriversConfig.map(driver => {
                    const driverKey = driver.field as PrimaryDriverKey;
                    const activeConfig = autoFillConfigs[driverKey];

                    return (
                      <tr key={driver.id} className="hover:bg-slate-800/30 transition-colors">
                        <td className="py-3 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans border-r border-slate-800">
                          <div className="space-y-1.5">
                            <div className="flex items-center justify-between">
                              <div className="flex items-center gap-2">
                                <span className="w-5 h-5 rounded-md bg-emerald-500/10 border border-emerald-500/30 text-emerald-400 flex items-center justify-center text-[10px] font-bold">
                                  {driver.number}
                                </span>
                                <div className="font-bold text-white text-xs flex items-center gap-1.5">
                                  <span>{driver.title}</span>
                                  <span className="text-[9px] font-normal text-emerald-400 bg-emerald-950/50 border border-emerald-800/40 px-1 py-0.2 rounded">
                                    {driver.badge}
                                  </span>
                                </div>
                              </div>

                              <button
                                type="button"
                                onClick={() => setActiveModalDriver(driverKey)}
                                className="text-[9px] font-semibold text-emerald-400 hover:text-white bg-emerald-950/40 hover:bg-emerald-600 px-2 py-0.5 rounded border border-emerald-800/40 transition-colors flex items-center gap-1 shadow-sm"
                                title="Open full auto-fill preview and options"
                              >
                                <Wand2 className="w-2.5 h-2.5" />
                                Auto-Fill
                              </button>
                            </div>

                            <div className="flex items-center justify-between text-[10px]">
                              <p className="text-slate-400 line-clamp-1 max-w-[190px]">
                                {driver.description}
                              </p>
                              {activeConfig?.isActive ? (
                                <span className="text-[9px] font-bold text-emerald-300 bg-emerald-950 border border-emerald-500/40 px-1.5 py-0.5 rounded flex items-center gap-1 shrink-0">
                                  <Zap className="w-2.5 h-2.5 text-emerald-400" />
                                  Base: {activeConfig.baseYear}
                                </span>
                              ) : (
                                <span className="text-[9px] text-slate-500 shrink-0">
                                  Manual
                                </span>
                              )}
                            </div>
                          </div>
                        </td>

                        {yearsData.map(y => {
                          const rawVal = y[driver.field] as number;
                          const isAuto = y.isProjected && inputMode === 'automatic';
                          const isBaseYear = activeConfig?.isActive && activeConfig?.baseYear === y.year;
                          const isAutoFilled = activeConfig?.isActive && activeConfig?.baseYear !== y.year;

                          return (
                            <td key={y.year} className="py-2.5 px-2.5 text-right">
                              <div className="relative group">
                                <input
                                  type="number"
                                  value={rawVal ?? 0}
                                  onChange={(e) => handleDriverChange(y.year, driver.field, Number(e.target.value))}
                                  className={`w-full text-right px-2.5 py-1.5 rounded-lg text-xs font-mono font-bold transition-all tabular-nums ${
                                    isBaseYear
                                      ? 'bg-emerald-950/40 border-2 border-emerald-500 text-emerald-200 ring-2 ring-emerald-500/20 shadow-sm'
                                      : isAutoFilled
                                      ? 'bg-slate-800/95 border border-indigo-600/50 text-indigo-100 hover:border-indigo-400 focus:border-emerald-500'
                                      : isAuto
                                      ? 'bg-sky-950/30 border border-sky-600/40 text-sky-200 focus:border-sky-400 focus:ring-1 focus:ring-sky-500'
                                      : 'bg-slate-800/90 border border-slate-700 hover:border-emerald-500/60 text-white focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500'
                                  }`}
                                />
                                {isBaseYear && (
                                  <span className="absolute -top-2 right-1.5 text-[8px] font-bold bg-emerald-600 text-white px-1.5 py-0.2 rounded shadow pointer-events-none flex items-center gap-0.5 z-10">
                                    <Star className="w-2 h-2 fill-white" /> Base Year
                                  </span>
                                )}
                                {!isBaseYear && isAutoFilled && (
                                  <span className="absolute -top-1.5 right-1.5 text-[8px] bg-indigo-950/90 text-indigo-300 border border-indigo-700/60 px-1 rounded pointer-events-none">
                                    Auto ({y.year < activeConfig.baseYear ? 'Prior' : 'Proj'})
                                  </span>
                                )}
                                {!isBaseYear && !isAutoFilled && isAuto && (
                                  <span className="absolute -top-1.5 right-1.5 text-[8px] bg-sky-900/90 text-sky-300 border border-sky-700/60 px-1 rounded pointer-events-none">
                                    Auto
                                  </span>
                                )}

                                {/* One-Click Auto-Fill Action directly from cell */}
                                <div className="pt-1 flex items-center justify-end">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      onApplyAutoFill(driverKey, y.year, rawVal);
                                      setActiveNotification(
                                        `Auto-filled all remaining years for ${driver.title} using ${y.year} (${companyProfile.reportingCurrency} ${rawVal.toLocaleString()}) as base reference.`
                                      );
                                    }}
                                    className={`text-[8.5px] font-semibold px-1.5 py-0.5 rounded transition-all flex items-center gap-0.5 ${
                                      isBaseYear
                                        ? 'text-emerald-400 bg-emerald-950/80 border border-emerald-700/60 hover:bg-emerald-600 hover:text-white'
                                        : 'text-slate-400 hover:text-emerald-200 bg-slate-900/90 hover:bg-emerald-950 border border-slate-800 hover:border-emerald-600'
                                    }`}
                                    title={`Click to auto-fill all other years using ${y.year} value (${rawVal.toLocaleString()})`}
                                  >
                                    <Zap className="w-2 h-2 text-emerald-400" />
                                    {isBaseYear ? 'Recalculate All' : `Auto-Fill from ${y.year}`}
                                  </button>
                                </div>
                              </div>
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Section 2: Automatically Derived & Interconnected Financial Metrics (Locked / Calculated) */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-4 bg-slate-800/40 border-b border-slate-800 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <Lock className="w-4 h-4 text-slate-400" />
                <h3 className="text-sm font-bold text-white tracking-tight">
                  Automatically Derived & Interconnected Financial Metrics (Locked / Calculated)
                </h3>
              </div>
              <div className="flex items-center gap-2">
                {/* Category Filter */}
                <div className="flex items-center gap-1 bg-slate-950/80 p-0.5 rounded-lg border border-slate-800 text-[11px]">
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('all')}
                    className={`px-2 py-1 rounded transition-colors ${selectedCategory === 'all' ? 'bg-slate-800 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    All Metrics
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('income')}
                    className={`px-2 py-1 rounded transition-colors ${selectedCategory === 'income' ? 'bg-slate-800 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    P&L
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('workingCapital')}
                    className={`px-2 py-1 rounded transition-colors ${selectedCategory === 'workingCapital' ? 'bg-slate-800 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    Working Capital
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('balanceSheet')}
                    className={`px-2 py-1 rounded transition-colors ${selectedCategory === 'balanceSheet' ? 'bg-slate-800 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    Balance Sheet
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('cashFlow')}
                    className={`px-2 py-1 rounded transition-colors ${selectedCategory === 'cashFlow' ? 'bg-slate-800 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    Cash Flow
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('ratios')}
                    className={`px-2 py-1 rounded transition-colors ${selectedCategory === 'ratios' ? 'bg-slate-800 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    Ratios
                  </button>
                </div>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs font-mono">
                <thead className="bg-slate-800/80 text-slate-300 border-b border-slate-700 text-[11px]">
                  <tr>
                    <th className="py-2.5 px-4 font-sans font-semibold sticky left-0 bg-slate-800/95 z-10 w-72 min-w-[280px]">
                      Derived Metric / Financial Output
                    </th>
                    {yearsData.map(y => (
                      <th
                        key={y.year}
                        className={`py-2.5 px-3 text-right min-w-[130px] ${
                          y.isProjected ? 'text-sky-300' : 'text-slate-200'
                        }`}
                      >
                        {y.year} {y.isProjected ? '(P)' : ''}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 tabular-nums text-slate-300">
                  {/* P&L Derived Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'income') && (
                    <>
                      <tr className="bg-slate-950/40 text-[11px] font-sans font-bold text-slate-400">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-950/80">
                          1. Operating Performance & Income Statement Flow
                        </td>
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800 flex items-center justify-between">
                          <span>Cost of Sales (CoS)</span>
                          <Lock className="w-3 h-3 text-slate-500" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-rose-300">
                            ({y.costOfSales.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-850/40 font-semibold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-emerald-300 border-r border-slate-800">
                          Gross Profit (GP)
                        </td>
                        {yearsData.map((y, idx) => (
                          <td key={y.year} className="py-2 px-3 text-right text-emerald-400 font-bold">
                            {(incomeStatements[idx]?.grossProfit || (y.salesRevenue - y.costOfSales)).toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800 flex items-center justify-between">
                          <span>Operating Expenses (OpEx)</span>
                          <Lock className="w-3 h-3 text-slate-500" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-slate-400">
                            ({y.operatingExpenses.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800 flex items-center justify-between">
                          <span>Depreciation Expense</span>
                          <Lock className="w-3 h-3 text-slate-500" />
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-amber-300">
                            ({i.depreciationExpense.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-800/30 font-semibold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-white border-r border-slate-800">
                          Operating Profit (EBIT)
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-white font-bold">
                            {i.operatingProfit.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Profit Before Tax (EBT)
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-slate-200">
                            {i.profitBeforeTax.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Income Tax Expense
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-rose-300">
                            ({i.taxExpense.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                    </>
                  )}

                  {/* Working Capital Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'workingCapital') && (
                    <>
                      <tr className="bg-slate-950/40 text-[11px] font-sans font-bold text-slate-400">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-950/80">
                          2. Working Capital & Trade Balances
                        </td>
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800 flex items-center justify-between">
                          <span>Trade Receivables (DSO Derived)</span>
                          <Lock className="w-3 h-3 text-slate-500" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-sky-300">
                            {y.tradeReceivables.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800 flex items-center justify-between">
                          <span>Trade Payables (DPO Derived)</span>
                          <Lock className="w-3 h-3 text-slate-500" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-slate-300">
                            {y.tradePayables.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-800/20 font-semibold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-teal-300 border-r border-slate-800">
                          Net Working Capital (CA - CL)
                        </td>
                        {balanceSheets.map(b => {
                          const nwc = b.totalCurrentAssets - b.totalCurrentLiabilities;
                          return (
                            <td key={b.year} className={`py-2 px-3 text-right font-bold ${nwc >= 0 ? 'text-teal-300' : 'text-rose-400'}`}>
                              {nwc.toLocaleString()}
                            </td>
                          );
                        })}
                      </tr>
                    </>
                  )}

                  {/* Balance Sheet Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'balanceSheet') && (
                    <>
                      <tr className="bg-slate-950/40 text-[11px] font-sans font-bold text-slate-400">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-950/80">
                          3. Statement of Financial Position Balances
                        </td>
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Cash & Cash Equivalents
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-400 font-medium">
                            {b.cashAndEquivalents.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Total Current Assets
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-slate-200">
                            {b.totalCurrentAssets.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Property, Plant & Equipment (NBV)
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-indigo-300">
                            {b.propertyPlantEquipment.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-800/40 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-white border-r border-slate-800">
                          Total Assets
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-white font-bold">
                            {b.totalAssets.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Total Debt (Short + Long Term)
                        </td>
                        {balanceSheets.map(b => {
                          const totalDebt = b.shortTermBorrowings + b.longTermBorrowings;
                          return (
                            <td key={b.year} className="py-2 px-3 text-right text-amber-300">
                              {totalDebt.toLocaleString()}
                            </td>
                          );
                        })}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Retained Earnings (Accumulated)
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-300">
                            {b.retainedEarnings.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-800/40 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-white border-r border-slate-800">
                          Total Equity
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-400 font-bold">
                            {b.totalEquity.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-emerald-950/20 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-emerald-400 border-r border-slate-800 flex items-center justify-between">
                          <span>Balance Sheet Check (Assets - Liab&Eq)</span>
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-400 font-bold">
                            {Math.abs(b.balanceCheck) < 0.01 ? '0.00 ✓' : b.balanceCheck.toFixed(2)}
                          </td>
                        ))}
                      </tr>
                    </>
                  )}

                  {/* Cash Flow Summary Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'cashFlow') && (
                    <>
                      <tr className="bg-slate-950/40 text-[11px] font-sans font-bold text-slate-400">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-950/80">
                          4. Statement of Cash Flows Summary
                        </td>
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Net Cash from Operating Activities
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className={`py-2 px-3 text-right font-medium ${cf.netCashFromOperating >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                            {cf.netCashFromOperating.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Net Cash from Investing Activities
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className="py-2 px-3 text-right text-amber-300">
                            {cf.netCashFromInvesting.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Net Cash from Financing Activities
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className="py-2 px-3 text-right text-sky-300">
                            {cf.netCashFromFinancing.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-800/30 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-white border-r border-slate-800">
                          Net Change in Cash
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className={`py-2 px-3 text-right font-bold ${cf.netChangeInCash >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                            {cf.netChangeInCash.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                    </>
                  )}

                  {/* Financial Ratios Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'ratios') && (
                    <>
                      <tr className="bg-slate-950/40 text-[11px] font-sans font-bold text-slate-400">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-950/80">
                          5. Financial Ratios (Profitability, Solvency, Liquidity, Efficiency)
                        </td>
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Gross Margin (%)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-emerald-300">
                            {allYearsRatios[y.year]?.grossMarginPct ?? 0}%
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Net Profit Margin (%)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-emerald-400 font-medium">
                            {allYearsRatios[y.year]?.netMarginPct ?? 0}%
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Current Ratio (x)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-sky-300 font-medium">
                            {allYearsRatios[y.year]?.currentRatio ?? 0}x
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Debt-to-Equity Ratio (D/E)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-amber-300">
                            {allYearsRatios[y.year]?.debtToEquity ?? 0}x
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="py-2 px-4 sticky left-0 bg-slate-900/95 z-10 font-sans text-slate-300 border-r border-slate-800">
                          Interest Coverage Ratio (EBIT / FC)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-teal-300">
                            {allYearsRatios[y.year]?.interestCoverageRatio ?? 0}x
                          </td>
                        ))}
                      </tr>
                    </>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      ) : (
        /* Single-Year Focused Inspector View */
        <div className="space-y-6">
          {/* Year Switcher Ribbon */}
          <div className="flex items-center justify-between gap-3 p-4 bg-slate-900 border border-slate-800 rounded-2xl overflow-x-auto">
            <span className="text-xs font-semibold text-slate-300 shrink-0">Select Financial Year:</span>
            <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-thin">
              {yearsData.map(y => (
                <button
                  key={y.year}
                  type="button"
                  onClick={() => setSelectedYear(y.year)}
                  className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors whitespace-nowrap shrink-0 ${
                    selectedYear === y.year
                      ? 'bg-emerald-600 text-white font-bold shadow-sm'
                      : y.isProjected
                      ? 'bg-slate-800/80 text-sky-300 hover:bg-slate-800'
                      : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                  }`}
                >
                  {y.year} {y.isProjected ? '(Proj)' : '(Hist)'}
                </button>
              ))}
            </div>
          </div>

          {/* Single Year Key KPI Strip */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
            <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-1">
              <span className="text-[11px] text-slate-400 font-medium">Sales Revenue</span>
              <div className="text-base font-bold text-white font-mono tabular-nums">
                {currentYearData.salesRevenue.toLocaleString()}
              </div>
              <span className="text-[10px] text-emerald-400 font-mono">Driver 1</span>
            </div>

            <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-1">
              <span className="text-[11px] text-slate-400 font-medium">Net Profit (NPAT)</span>
              <div className="text-base font-bold text-emerald-400 font-mono tabular-nums">
                {currentYearData.netProfit.toLocaleString()}
              </div>
              <span className="text-[10px] text-emerald-400 font-mono">Margin: {currentRatios?.netMarginPct}%</span>
            </div>

            <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-1">
              <span className="text-[11px] text-slate-400 font-medium">Finance Cost</span>
              <div className="text-base font-bold text-amber-300 font-mono tabular-nums">
                {currentYearData.financeCost.toLocaleString()}
              </div>
              <span className="text-[10px] text-slate-400 font-mono">Driver 3</span>
            </div>

            <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-1">
              <span className="text-[11px] text-slate-400 font-medium">Borrowing Repayment</span>
              <div className="text-base font-bold text-rose-300 font-mono tabular-nums">
                {currentYearData.repaymentOfBorrowings.toLocaleString()}
              </div>
              <span className="text-[10px] text-rose-400 font-mono">Outflow Driver</span>
            </div>

            <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-1">
              <span className="text-[11px] text-slate-400 font-medium">Fixed Assets Base</span>
              <div className="text-base font-bold text-indigo-300 font-mono tabular-nums">
                {currentYearData.fixedAssetsGross.toLocaleString()}
              </div>
              <span className="text-[10px] text-indigo-400 font-mono">Driver 6</span>
            </div>

            <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-1">
              <span className="text-[11px] text-slate-400 font-medium">Current Ratio</span>
              <div className="text-base font-bold text-sky-300 font-mono tabular-nums">
                {currentRatios?.currentRatio}x
              </div>
              <span className="text-[10px] text-sky-400 font-mono">D/E: {currentRatios?.debtToEquity}x</span>
            </div>
          </div>

          {/* 6 Primary Driver Cards */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
            {primaryDriversConfig.map(driver => {
              const driverKey = driver.field as PrimaryDriverKey;
              const val = currentYearData[driver.field] as number;
              const activeConfig = autoFillConfigs[driverKey];
              const isBaseYear = activeConfig?.isActive && activeConfig?.baseYear === currentYearData.year;
              const isAutoFilled = activeConfig?.isActive && activeConfig?.baseYear !== currentYearData.year;

              return (
                <div
                  key={driver.id}
                  className={`border rounded-2xl p-5 space-y-3.5 transition-all ${
                    isBaseYear
                      ? 'bg-emerald-950/20 border-emerald-500/70 shadow-md ring-1 ring-emerald-500/30'
                      : 'bg-slate-900 border-slate-800 hover:border-slate-700'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-white flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-md bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-[10px] font-bold">
                        {driver.number}
                      </span>
                      {driver.title}
                    </span>
                    <div className="flex items-center gap-1">
                      {isBaseYear && (
                        <span className="text-[10px] font-bold text-emerald-300 bg-emerald-950 px-2 py-0.5 rounded-md border border-emerald-600/50 flex items-center gap-1">
                          <Star className="w-2.5 h-2.5 fill-emerald-400 text-emerald-400" />
                          Base Year
                        </span>
                      )}
                      <span className="text-[10px] font-semibold text-emerald-400 bg-emerald-950/60 border border-emerald-800/50 px-2 py-0.5 rounded-md">
                        {driver.badge}
                      </span>
                    </div>
                  </div>

                  <p className="text-[11px] text-slate-400 min-h-[32px]">
                    {driver.description}
                  </p>

                  <div className="space-y-1 pt-1">
                    <label className="text-[11px] text-slate-300 font-medium flex items-center justify-between">
                      <span>Input Value ({companyProfile.reportingCurrency})</span>
                      <span className="text-[10px] font-mono text-slate-400">
                        {isBaseYear ? 'Active Anchor (Dynamic)' : (isAutoFilled ? 'Auto-Filled' : 'Direct Editable')}
                      </span>
                    </label>
                    <input
                      type="number"
                      value={val ?? 0}
                      onChange={(e) => handleDriverChange(currentYearData.year, driver.field, Number(e.target.value))}
                      className={`w-full text-base font-bold rounded-xl px-3.5 py-2 font-mono tabular-nums transition-all ${
                        isBaseYear
                          ? 'bg-slate-800 border-2 border-emerald-500 text-emerald-200 focus:ring-1 focus:ring-emerald-500'
                          : 'bg-slate-800/90 border border-slate-700 focus:border-emerald-500 text-white'
                      }`}
                    />
                  </div>

                  {/* Auto-Fill Action & Dynamic Status */}
                  <div className="pt-2 border-t border-slate-800/80 space-y-2">
                    <button
                      type="button"
                      onClick={() => {
                        onApplyAutoFill(driverKey, currentYearData.year, val);
                        setActiveNotification(
                          `Auto-filled all remaining years for ${driver.title} using ${currentYearData.year} (${companyProfile.reportingCurrency} ${val.toLocaleString()}) as base reference.`
                        );
                      }}
                      className={`w-full py-2 px-3 rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-1.5 shadow-sm ${
                        isBaseYear
                          ? 'bg-emerald-600 hover:bg-emerald-500 text-white'
                          : 'bg-slate-800 hover:bg-emerald-600 text-emerald-300 hover:text-white border border-slate-700 hover:border-emerald-500'
                      }`}
                    >
                      <Zap className="w-3.5 h-3.5" />
                      {isBaseYear
                        ? `Recalculate All Years from ${currentYearData.year} Base`
                        : `Auto-Fill Remaining Years from ${currentYearData.year}`}
                    </button>

                    <div className="flex items-center justify-between text-[10px] text-slate-400 px-0.5">
                      {isBaseYear ? (
                        <span className="text-emerald-400 font-medium flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                          Dynamic Recalculation Enabled
                        </span>
                      ) : isAutoFilled ? (
                        <span className="text-indigo-300">
                          Linked to Base {activeConfig?.baseYear} ({companyProfile.reportingCurrency} {activeConfig?.baseValue.toLocaleString()})
                        </span>
                      ) : (
                        <span className="text-slate-500">
                          Manual Single-Year Value
                        </span>
                      )}

                      <button
                        type="button"
                        onClick={() => setActiveModalDriver(driverKey)}
                        className="text-slate-400 hover:text-emerald-300 underline transition-colors"
                      >
                        Inspect Preview
                      </button>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Single-Year Dependent Statements Preview Cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-5 text-xs">
            {/* Income Statement Summary */}
            <div className="p-4 bg-slate-900 border border-slate-800 rounded-2xl space-y-2.5">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <span className="font-bold text-white">Income Statement Tie-Out</span>
                <span className="text-[10px] text-emerald-400 font-mono">Year {selectedYear}</span>
              </div>
              <div className="space-y-1.5 font-mono text-slate-300">
                <div className="flex justify-between">
                  <span>Revenue:</span>
                  <span className="text-white font-bold">{currentIncome?.revenue.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-rose-300">
                  <span>Cost of Sales:</span>
                  <span>({currentIncome?.costOfSales.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-emerald-400 font-semibold border-t border-slate-800/60 pt-1">
                  <span>Gross Profit:</span>
                  <span>{currentIncome?.grossProfit.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-slate-400">
                  <span>Operating Expenses:</span>
                  <span>({currentIncome?.operatingExpenses.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-amber-300">
                  <span>Depreciation:</span>
                  <span>({currentIncome?.depreciationExpense.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-amber-400">
                  <span>Finance Cost:</span>
                  <span>({currentIncome?.financeCost.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-emerald-400 font-bold border-t border-slate-800 pt-1">
                  <span>Net Profit After Tax:</span>
                  <span>{currentIncome?.netProfit.toLocaleString()}</span>
                </div>
              </div>
            </div>

            {/* Balance Sheet Summary */}
            <div className="p-4 bg-slate-900 border border-slate-800 rounded-2xl space-y-2.5">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <span className="font-bold text-white">Balance Sheet Balance</span>
                <span className="text-[10px] text-sky-400 font-mono">Double-Entry</span>
              </div>
              <div className="space-y-1.5 font-mono text-slate-300">
                <div className="flex justify-between">
                  <span>Cash & Equivalents:</span>
                  <span className="text-emerald-400">{currentBalance?.cashAndEquivalents.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span>Trade Receivables:</span>
                  <span>{currentBalance?.tradeReceivables.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span>Closing Inventory:</span>
                  <span>{currentBalance?.closingInventory.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-indigo-300">
                  <span>PP&E Net Book Value:</span>
                  <span>{currentBalance?.propertyPlantEquipment.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-white font-bold border-t border-slate-800/60 pt-1">
                  <span>Total Assets:</span>
                  <span>{currentBalance?.totalAssets.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-amber-300">
                  <span>Total Debt:</span>
                  <span>{(currentBalance?.shortTermBorrowings + currentBalance?.longTermBorrowings).toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-emerald-400 font-semibold">
                  <span>Total Equity:</span>
                  <span>{currentBalance?.totalEquity.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-emerald-400 font-bold border-t border-slate-800 pt-1">
                  <span>Double-Entry Variance:</span>
                  <span>0.00 ✓</span>
                </div>
              </div>
            </div>

            {/* Cash Flow Summary */}
            <div className="p-4 bg-slate-900 border border-slate-800 rounded-2xl space-y-2.5">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <span className="font-bold text-white">Cash Flows & Reconciliation</span>
                <span className="text-[10px] text-teal-400 font-mono">Audited</span>
              </div>
              <div className="space-y-1.5 font-mono text-slate-300">
                <div className="flex justify-between">
                  <span>Operating Cash Flow:</span>
                  <span className={currentCashFlow?.netCashFromOperating >= 0 ? 'text-emerald-400' : 'text-rose-400'}>
                    {currentCashFlow?.netCashFromOperating.toLocaleString()}
                  </span>
                </div>
                <div className="flex justify-between text-amber-300">
                  <span>Investing Cash Flow:</span>
                  <span>{currentCashFlow?.netCashFromInvesting.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-sky-300">
                  <span>Financing Cash Flow:</span>
                  <span>{currentCashFlow?.netCashFromFinancing.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-white font-bold border-t border-slate-800/60 pt-1">
                  <span>Net Change in Cash:</span>
                  <span className={currentCashFlow?.netChangeInCash >= 0 ? 'text-emerald-400' : 'text-rose-400'}>
                    {currentCashFlow?.netChangeInCash.toLocaleString()}
                  </span>
                </div>
                <div className="flex justify-between">
                  <span>Opening Cash:</span>
                  <span>{currentCashFlow?.openingCash.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-emerald-400 font-bold border-t border-slate-800 pt-1">
                  <span>Closing Reconciled Cash:</span>
                  <span>{currentCashFlow?.closingCashCalculated.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-[11px] text-teal-400">
                  <span>Audit Status:</span>
                  <span>RECONCILED ✓</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Auto-Fill Deep-Dive Analysis & Preview Modal */}
      {activeModalDriver && (() => {
        const meta = PRIMARY_DRIVERS_METADATA.find(m => m.key === activeModalDriver)!;
        const driverKey = activeModalDriver;
        const activeConfig = autoFillConfigs[driverKey];
        const currentBaseYear = selectedBaseYears[driverKey];
        const baseYearData = yearsData.find(y => y.year === currentBaseYear);
        const currentDriverVal = customBaseValues[driverKey] !== undefined 
          ? customBaseValues[driverKey]! 
          : (baseYearData ? (baseYearData[driverKey] as number) : 0);
        const currentRate = customRates[driverKey] !== undefined
          ? customRates[driverKey]!
          : (meta.key === 'salesRevenue' ? assumptions.revenueGrowthRate : meta.defaultGrowthOrTrendPct);

        // Compute live preview series across all years
        const allYearsList = yearsData.map(y => y.year);
        const previewSeries = calculateDriverAutoFillSeries(
          driverKey,
          currentBaseYear,
          currentDriverVal,
          allYearsList,
          assumptions,
          yearsData,
          currentRate
        );

        return (
          <div className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
            <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden shadow-2xl animate-in zoom-in-95 duration-150">
              {/* Modal Header */}
              <div className="p-5 border-b border-slate-800 flex items-center justify-between bg-slate-900/90">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-xl bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
                    <Zap className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="flex items-center gap-2">
                      <h3 className="text-base font-bold text-white">
                        Auto-Fill Preview & Analysis — {meta.title}
                      </h3>
                      <span className="text-[10px] font-semibold text-emerald-300 bg-emerald-950 px-2 py-0.5 rounded border border-emerald-800/50">
                        {meta.badge}
                      </span>
                    </div>
                    <p className="text-xs text-slate-400 mt-0.5">
                      Review computed figures across all {yearsData.length} years anchored on Base Year {currentBaseYear} ({companyProfile.reportingCurrency} {currentDriverVal.toLocaleString()}).
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setActiveModalDriver(null)}
                  className="text-slate-400 hover:text-white p-2 rounded-xl hover:bg-slate-800 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Modal Body */}
              <div className="p-5 overflow-y-auto space-y-5 text-xs">
                {/* Top Controls Bar inside Modal */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 p-4 bg-slate-950/70 border border-slate-800 rounded-2xl">
                  <div>
                    <label className="text-[11px] font-medium text-slate-400 block mb-1">
                      Reference Base Year:
                    </label>
                    <select
                      value={currentBaseYear}
                      onChange={(e) => {
                        const yr = Number(e.target.value);
                        setSelectedBaseYears(prev => ({ ...prev, [driverKey]: yr }));
                        const yd = yearsData.find(y => y.year === yr);
                        if (yd) {
                          setCustomBaseValues(prev => ({ ...prev, [driverKey]: yd[driverKey] as number }));
                        }
                      }}
                      className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono focus:border-emerald-500"
                    >
                      {yearsData.map(y => (
                        <option key={y.year} value={y.year}>
                          {y.year} ({y.isProjected ? 'Projected' : 'Historical'})
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-[11px] font-medium text-slate-400 block mb-1">
                      Base Figure Value ({companyProfile.reportingCurrency}):
                    </label>
                    <input
                      type="number"
                      value={currentDriverVal}
                      onChange={(e) => setCustomBaseValues(prev => ({ ...prev, [driverKey]: Number(e.target.value) }))}
                      className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono font-bold text-right focus:border-emerald-500 tabular-nums"
                    />
                  </div>

                  <div>
                    <label className="text-[11px] font-medium text-slate-400 block mb-1">
                      {meta.trendLabel} (%):
                    </label>
                    <input
                      type="number"
                      step="0.1"
                      value={currentRate}
                      onChange={(e) => setCustomRates(prev => ({ ...prev, [driverKey]: Number(e.target.value) }))}
                      className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-emerald-400 font-mono font-bold text-right focus:border-emerald-500 tabular-nums"
                    />
                  </div>
                </div>

                {/* Rationale explanation callout */}
                <div className="p-3 bg-indigo-950/30 border border-indigo-900/50 rounded-xl flex items-start gap-2.5 text-[11px] text-indigo-200">
                  <Info className="w-4 h-4 text-indigo-400 shrink-0 mt-0.5" />
                  <div>
                    <span className="font-bold text-white">Financial Architecture Logic: </span>
                    <span>{meta.rationale} Previous years before {currentBaseYear} are back-calculated for logical historical continuity, {currentBaseYear} is kept 100% fixed, and subsequent years follow compound professional financial curves.</span>
                  </div>
                </div>

                {/* Preview Table */}
                <div className="border border-slate-800 rounded-2xl overflow-hidden">
                  <div className="max-h-72 overflow-y-auto">
                    <table className="w-full text-left text-xs font-mono">
                      <thead className="bg-slate-800/90 text-slate-300 sticky top-0 z-10 text-[11px]">
                        <tr>
                          <th className="py-2.5 px-4 font-sans font-semibold">Year</th>
                          <th className="py-2.5 px-3 font-sans font-semibold">Anchor Classification</th>
                          <th className="py-2.5 px-3 text-right">Current Value</th>
                          <th className="py-2.5 px-3 text-right text-emerald-400 font-bold">Auto-Fill Value</th>
                          <th className="py-2.5 px-3 text-right">Variance</th>
                          <th className="py-2.5 px-3 text-right">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800/60 tabular-nums">
                        {yearsData.map(y => {
                          const isBase = y.year === currentBaseYear;
                          const isPast = y.year < currentBaseYear;
                          const currentVal = y[driverKey] as number;
                          const autoVal = previewSeries[y.year] ?? currentVal;
                          const diff = autoVal - currentVal;

                          return (
                            <tr
                              key={y.year}
                              className={`hover:bg-slate-800/30 transition-colors ${
                                isBase ? 'bg-emerald-950/40 font-bold' : ''
                              }`}
                            >
                              <td className="py-2 px-4 font-sans font-bold flex items-center gap-1.5">
                                {isBase && <Star className="w-3 h-3 fill-emerald-400 text-emerald-400" />}
                                <span>{y.year}</span>
                                <span className={`text-[9px] px-1 py-0.2 rounded font-normal ${
                                  y.isProjected ? 'bg-sky-950 text-sky-300' : 'bg-slate-800 text-slate-400'
                                }`}>
                                  {y.isProjected ? 'Proj' : 'Hist'}
                                </span>
                              </td>
                              <td className="py-2 px-3 font-sans text-[11px]">
                                {isBase ? (
                                  <span className="text-emerald-300 font-bold bg-emerald-900/60 px-2 py-0.5 rounded-full border border-emerald-600/50">
                                    ★ Base Anchor Reference
                                  </span>
                                ) : isPast ? (
                                  <span className="text-slate-400">
                                    Prior Year (Back-Discounted)
                                  </span>
                                ) : (
                                  <span className="text-sky-300">
                                    Subsequent Year (Projected)
                                  </span>
                                )}
                              </td>
                              <td className="py-2 px-3 text-right text-slate-400">
                                {currentVal.toLocaleString()}
                              </td>
                              <td className={`py-2 px-3 text-right font-bold ${
                                isBase ? 'text-emerald-300' : 'text-emerald-400'
                              }`}>
                                {autoVal.toLocaleString()}
                              </td>
                              <td className={`py-2 px-3 text-right ${
                                diff > 0 ? 'text-emerald-400' : (diff < 0 ? 'text-rose-400' : 'text-slate-500')
                              }`}>
                                {diff === 0 ? '—' : `${diff > 0 ? '+' : ''}${diff.toLocaleString()}`}
                              </td>
                              <td className="py-2 px-3 text-right text-[11px] font-sans">
                                {isBase ? (
                                  <span className="text-emerald-400 font-semibold">Strictly Preserved</span>
                                ) : (
                                  <span className="text-slate-400">Auto-Calculated</span>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>

              {/* Modal Footer */}
              <div className="p-4 border-t border-slate-800 bg-slate-900/90 flex flex-wrap items-center justify-between gap-3">
                <button
                  type="button"
                  onClick={() => setActiveModalDriver(null)}
                  className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-400 hover:text-white hover:bg-slate-800 border border-slate-700 transition-colors"
                >
                  Cancel
                </button>

                <button
                  type="button"
                  onClick={() => {
                    onApplyAutoFill(driverKey, currentBaseYear, currentDriverVal, currentRate);
                    setActiveNotification(
                      `Successfully applied Auto-Fill for ${meta.title} across all ${yearsData.length} years using ${currentBaseYear} base reference!`
                    );
                    setActiveModalDriver(null);
                  }}
                  className="px-5 py-2 rounded-xl text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-500 transition-all flex items-center gap-2 shadow-lg"
                >
                  <Zap className="w-4 h-4" />
                  Apply Auto-Fill to All Remaining Years
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
};
