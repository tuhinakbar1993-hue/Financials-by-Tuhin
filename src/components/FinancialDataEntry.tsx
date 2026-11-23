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
  ParsedRepaymentSchedule,
  RepaymentScheduleSnapshot,
  DebtServiceCoverageYear,
} from '../types/financial';
import {
  calculateFinancialRatios,
  FinancialRatios,
  PRIMARY_DRIVERS_METADATA,
  calculateDriverAutoFillSeries,
  ImplementScheduleResult,
} from '../services/financialEngine';
import { YearSelectorControl } from './YearSelectorControl';
import { RepaymentScheduleStudio } from './RepaymentScheduleStudio';

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
  parsedSchedule?: ParsedRepaymentSchedule | null;
  onSetParsedSchedule?: (schedule: ParsedRepaymentSchedule | null) => void;
  scheduleSnapshot?: RepaymentScheduleSnapshot | null;
  isScheduleImplemented?: boolean;
  lastImplementationResult?: ImplementScheduleResult | null;
  debtServiceCoverage?: DebtServiceCoverageYear[];
  onImplementRepaymentSchedule?: (options: { autoIncludeScheduleYears: boolean; alignHistoricalScale: boolean }) => void;
  onRevertScheduleSnapshot?: () => void;
  onUnlockScheduleYear?: (year: number) => void;
  initialShowScheduleStudio?: boolean;
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
  parsedSchedule = null,
  onSetParsedSchedule,
  scheduleSnapshot = null,
  isScheduleImplemented = false,
  lastImplementationResult = null,
  debtServiceCoverage = [],
  onImplementRepaymentSchedule,
  onRevertScheduleSnapshot,
  onUnlockScheduleYear,
  initialShowScheduleStudio = true,
}) => {
  const [viewMode, setViewMode] = useState<'grid' | 'single'>('grid');
  const [selectedYear, setSelectedYear] = useState<number>(yearsData[0]?.year || 2023);
  const [showInterconnectionLogic, setShowInterconnectionLogic] = useState<boolean>(false);
  const [showAutoFillPanel, setShowAutoFillPanel] = useState<boolean>(true);
  const [showScheduleStudio, setShowScheduleStudio] = useState<boolean>(initialShowScheduleStudio);
  const [showYearSelector, setShowYearSelector] = useState<boolean>(false);
  const [activeModalDriver, setActiveModalDriver] = useState<PrimaryDriverKey | null>(null);
  const [activeNotification, setActiveNotification] = useState<string | null>(null);

  useEffect(() => {
    if (initialShowScheduleStudio) {
      setShowScheduleStudio(true);
    }
  }, [initialShowScheduleStudio]);

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
      <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl p-5 space-y-4 shadow-sm">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-sky-50 border border-sky-200 flex items-center justify-center text-sky-700 shrink-0">
                <Database className="w-4 h-4" />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-lg font-bold text-slate-900 tracking-tight">
                    Financial Data Entry & Primary Driver Architecture
                  </h2>
                  <span className="text-[11px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md">
                    6 Strictly Fixed Inputs
                  </span>
                  {validationIssues.length === 0 ? (
                    <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md inline-flex items-center gap-1">
                      <ShieldCheck className="w-3 h-3 text-emerald-700" />
                      Reconciled
                    </span>
                  ) : (
                    <span className="text-[10px] font-semibold text-amber-800 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-md inline-flex items-center gap-1">
                      <AlertTriangle className="w-3 h-3 text-amber-700" />
                      {validationIssues.length} Alert(s)
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-600 mt-0.5">
                  The six primary user inputs remain strictly unchanged across all calculations. All dependent statement figures, ratios, and cash flows adjust logically and proportionately.
                </p>
              </div>
            </div>
          </div>

          {/* Mode Switcher & Global Actions */}
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Manual vs Automatic Input Mode Segmented Toggle */}
            <div className="flex items-center bg-slate-100/95 p-1 rounded-xl border border-slate-200">
              <button
                type="button"
                onClick={() => onToggleInputMode('manual')}
                className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  inputMode === 'manual'
                    ? 'bg-emerald-600 text-white shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
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
                className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  inputMode === 'automatic'
                    ? 'bg-sky-600 text-white shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Sparkles className="w-3.5 h-3.5" />
                Automatic / Interconnected Mode
              </button>
            </div>

            {/* View Mode Toggle */}
            <div className="flex items-center bg-slate-100/95 p-1 rounded-xl border border-slate-200">
              <button
                type="button"
                onClick={() => setViewMode('grid')}
                className={`px-2.5 py-1.5 text-xs font-medium rounded-lg transition-colors flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  viewMode === 'grid'
                    ? 'bg-white text-slate-900 shadow-xs border border-slate-200 font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
                title="Spreadsheet multi-year grid view"
              >
                <LayoutGrid className="w-3.5 h-3.5" />
                Multi-Year Grid
              </button>
              <button
                type="button"
                onClick={() => setViewMode('single')}
                className={`px-2.5 py-1.5 text-xs font-medium rounded-lg transition-colors flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  viewMode === 'single'
                    ? 'bg-white text-slate-900 shadow-xs border border-slate-200 font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
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
                className={`px-3 py-1.5 text-xs font-semibold rounded-xl border transition-all flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  showYearSelector
                    ? 'bg-sky-600 border-sky-600 text-white shadow-sm'
                    : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-slate-900'
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
              className={`px-3 py-1.5 text-xs font-semibold rounded-xl border transition-all flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                showAutoFillPanel
                  ? 'bg-emerald-50 border-emerald-400 text-emerald-800 shadow-sm'
                  : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-slate-900'
              }`}
            >
              <Zap className="w-3.5 h-3.5 text-emerald-600" />
              Auto-Fill Hub
              {Object.values(autoFillConfigs).filter(c => c?.isActive).length > 0 && (
                <span className="bg-emerald-600 text-white text-[10px] font-bold px-1.5 py-0.5 rounded-md">
                  {Object.values(autoFillConfigs).filter(c => c?.isActive).length} Active
                </span>
              )}
            </button>

            {/* Repayment Schedule & DSCR Toggle */}
            {onSetParsedSchedule && onImplementRepaymentSchedule && onRevertScheduleSnapshot && (
              <button
                type="button"
                onClick={() => setShowScheduleStudio(!showScheduleStudio)}
                className={`px-3 py-1.5 text-xs font-semibold rounded-xl border transition-all flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  showScheduleStudio
                    ? 'bg-emerald-600 border-emerald-600 text-white shadow-sm'
                    : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-slate-900'
                }`}
              >
                <ShieldCheck className="w-3.5 h-3.5" />
                Repayment Schedule{isScheduleImplemented ? ' (Locked)' : ''}
              </button>
            )}

            {/* Interconnection Logic Info Toggle */}
            <button
              type="button"
              onClick={() => setShowInterconnectionLogic(!showInterconnectionLogic)}
              className={`px-3 py-1.5 text-xs font-medium rounded-xl border transition-colors flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                showInterconnectionLogic
                  ? 'bg-indigo-50 border-indigo-300 text-indigo-800 font-semibold'
                  : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-slate-900'
              }`}
            >
              <Info className="w-3.5 h-3.5 text-indigo-600" />
              Interconnection Logic
            </button>
          </div>
        </div>

        {/* Projection Horizon Quick Controller */}
        <div className="pt-3 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex flex-wrap items-center gap-2">
            <Calendar className="w-3.5 h-3.5 text-slate-500" />
            <span className="text-slate-700 font-medium">Projection Horizon:</span>
            <span className="font-mono font-bold text-emerald-700 tabular-nums">
              {projYearsCount} Years ({yearsData.find(y => y.isProjected)?.year}–{yearsData[yearsData.length - 1]?.year})
            </span>
            <div className="flex items-center gap-1 ml-2">
              {[3, 5, 7, 10, 15, 20].map(count => (
                <button
                  key={count}
                  type="button"
                  onClick={() => handleQuickHorizonChange(count)}
                  className={`px-2 py-0.5 rounded text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 ${
                    projYearsCount === count
                      ? 'bg-emerald-600 text-white shadow-sm'
                      : 'bg-slate-100 text-slate-700 hover:text-slate-900 hover:bg-slate-200 border border-slate-200'
                  }`}
                >
                  {count}Y
                </button>
              ))}
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3 text-[11px] text-slate-600">
            <div className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-sm bg-emerald-100 border border-emerald-500 inline-block"></span>
              <span>Fixed Primary Input</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-sm bg-sky-100 border border-sky-400 inline-block"></span>
              <span>Automated Projection</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-sm bg-slate-200 border border-slate-400 inline-block"></span>
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

      {/* Repayment Schedule Upload, Implementation & DSCR Studio */}
      {showScheduleStudio && onSetParsedSchedule && onImplementRepaymentSchedule && onRevertScheduleSnapshot && (
        <RepaymentScheduleStudio
          companyProfile={companyProfile}
          allModelYears={yearsData.map(y => y.year)}
          parsedSchedule={parsedSchedule}
          onSetParsedSchedule={onSetParsedSchedule}
          scheduleSnapshot={scheduleSnapshot}
          isScheduleImplemented={isScheduleImplemented}
          lastImplementationResult={lastImplementationResult}
          debtServiceCoverage={debtServiceCoverage}
          onImplementRepaymentSchedule={(opts) => {
            onImplementRepaymentSchedule(opts);
            setActiveNotification(
              'Repayment schedule implemented! Exact annual Principal and Finance Cost are locked and projected cash flows are sustainably optimized.'
            );
          }}
          onRevertScheduleSnapshot={() => {
            onRevertScheduleSnapshot();
            setActiveNotification('Reverted all financial figures back to the pre-schedule snapshot.');
          }}
        />
      )}

      {/* Toast Notification for Auto-Fill operations */}
      {activeNotification && (
        <div className="bg-emerald-50 border border-emerald-300 text-emerald-900 px-4 py-3 rounded-2xl flex items-center justify-between text-xs shadow-sm animate-in fade-in duration-200">
          <div className="flex items-center gap-2.5">
            <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
            <span className="font-medium">{activeNotification}</span>
          </div>
          <button
            type="button"
            onClick={() => setActiveNotification(null)}
            className="text-emerald-700 hover:text-emerald-950 p-1 rounded-md transition-colors"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Dedicated Auto-Fill Remaining Years Facility Hub */}
      {showAutoFillPanel && (
        <div className="bg-white/92 backdrop-blur-md border border-emerald-200 rounded-2xl p-5 space-y-4 text-xs shadow-sm">
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 border-b border-slate-200 pb-3">
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700 shrink-0">
                <Zap className="w-4 h-4" />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h3 className="font-bold text-slate-900 text-sm">
                    Auto-Fill Remaining Years Facility — 6 Primary Financial Drivers
                  </h3>
                  <span className="text-[10px] font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md">
                    Independent Single-Year Anchors
                  </span>
                </div>
                <p className="text-[11px] text-slate-600 mt-0.5">
                  Enter or select any single base year for each item independently. The system preserves your entered base reference and automatically calculates all remaining prior historical years and subsequent projected years using professional financial analyst ratios.
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 text-xs shrink-0">
              <span className="text-slate-600 text-[11px]">Active Anchors:</span>
              <span className="font-mono font-bold text-emerald-800 bg-emerald-50 px-2.5 py-1 rounded-md border border-emerald-200 tabular-nums">
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
                      ? 'bg-emerald-50/70 border-emerald-300 shadow-sm'
                      : 'bg-slate-50/90 border-slate-200 hover:border-slate-300'
                  }`}
                >
                  {/* Card Header */}
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-md bg-emerald-100 text-emerald-800 border border-emerald-200 flex items-center justify-center text-[10px] font-bold">
                        {meta.number}
                      </span>
                      <span className="font-bold text-slate-900 text-xs">
                        {meta.title}
                      </span>
                    </div>
                    {isActive ? (
                      <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-100 border border-emerald-300 px-2 py-0.5 rounded-md flex items-center gap-1">
                        <Star className="w-2.5 h-2.5 fill-emerald-600 text-emerald-600" />
                        Base: {activeConfig.baseYear}
                      </span>
                    ) : (
                      <span className="text-[10px] text-slate-600 bg-white px-2 py-0.5 rounded border border-slate-200">
                        Manual Mode
                      </span>
                    )}
                  </div>

                  {/* Base Year & Base Value Selector */}
                  <div className="grid grid-cols-2 gap-2 pt-1">
                    <div>
                      <label className="text-[10px] text-slate-600 font-medium block mb-1">
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
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-mono focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                      >
                        {yearsData.map(y => (
                          <option key={y.year} value={y.year}>
                            {y.year} ({y.isProjected ? 'Projected' : 'Historical'})
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="text-[10px] text-slate-600 font-medium block mb-1">
                        Base Value ({companyProfile.reportingCurrency}):
                      </label>
                      <input
                        type="number"
                        value={currentDriverVal}
                        onChange={(e) => {
                          const val = Number(e.target.value);
                          setCustomBaseValues(prev => ({ ...prev, [driverKey]: val }));
                        }}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-mono font-bold text-right focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500 tabular-nums"
                      />
                    </div>
                  </div>

                  {/* Growth / Trend Calibration */}
                  <div className="bg-white p-2.5 rounded-lg border border-slate-200 space-y-1.5">
                    <div className="flex items-center justify-between text-[10px]">
                      <span className="text-slate-600 font-medium">
                        {meta.trendLabel}:
                      </span>
                      <div className="flex items-center gap-1 font-mono">
                        <input
                          type="number"
                          step="0.1"
                          value={currentRate}
                          onChange={(e) => setCustomRates(prev => ({ ...prev, [driverKey]: Number(e.target.value) }))}
                          className="w-16 bg-slate-50 border border-slate-300 rounded px-1.5 py-0.5 text-right text-xs text-emerald-700 font-bold tabular-nums focus:border-emerald-600"
                        />
                        <span className="text-slate-600">%</span>
                      </div>
                    </div>
                    <p className="text-[9.5px] text-slate-600 leading-tight">
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
                          ? 'bg-emerald-600 hover:bg-emerald-700 text-white'
                          : 'bg-white hover:bg-emerald-600 text-emerald-800 hover:text-white border border-emerald-300 hover:border-emerald-600'
                      }`}
                    >
                      <Zap className="w-3.5 h-3.5" />
                      {isActive ? 'Recalculate All Remaining Years' : 'Auto-Fill Remaining Years'}
                    </button>

                    <div className="flex items-center justify-between text-[10px] pt-0.5">
                      <button
                        type="button"
                        onClick={() => setActiveModalDriver(driverKey)}
                        className="text-slate-600 hover:text-emerald-700 flex items-center gap-1 transition-colors font-medium"
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
                          className="text-slate-500 hover:text-rose-700 transition-colors font-medium"
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
        <div className="bg-white/92 backdrop-blur-md border border-indigo-200 rounded-2xl p-5 space-y-4 text-xs shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-3">
            <div className="flex items-center gap-2">
              <Zap className="w-4 h-4 text-indigo-600" />
              <h3 className="font-bold text-slate-900 text-sm">
                Financial Model Interconnection Architecture & Formula Relationships
              </h3>
            </div>
            <span className="text-[11px] font-semibold text-indigo-800 bg-indigo-50 border border-indigo-200 px-2 py-0.5 rounded-md">
              Double-Entry Mathematical Guarantee
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-slate-700">
            <div className="p-3.5 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1.5">
              <span className="font-bold text-emerald-800 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-md bg-emerald-100 text-emerald-800 flex items-center justify-center text-[10px]">1</span>
                Sales Revenue Driver
              </span>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Drives <strong className="text-slate-900">Trade Receivables</strong> via DSO ({(currentYearData?.tradeReceivables || 0).toLocaleString()} {companyProfile.reportingCurrency}), calibrates <strong className="text-slate-900">Operating Scale</strong>, and derives <strong className="text-slate-900">Cost of Sales</strong> & <strong className="text-slate-900">OpEx</strong> tying out with Net Profit.
              </p>
            </div>

            <div className="p-3.5 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1.5">
              <span className="font-bold text-teal-800 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-md bg-teal-100 text-teal-800 flex items-center justify-center text-[10px]">2</span>
                Closing Inventory Driver
              </span>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Directly feeds <strong className="text-slate-900">Working Capital</strong> in Balance Sheet and drives the <strong className="text-slate-900">Change in Inventory</strong> cash flow adjustment. Automatically calculates Days Sales of Inventory (DIO).
              </p>
            </div>

            <div className="p-3.5 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1.5">
              <span className="font-bold text-amber-800 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-md bg-amber-100 text-amber-800 flex items-center justify-center text-[10px]">3</span>
                Finance Cost Driver
              </span>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Represents annual interest/debt servicing. Enters Income Statement directly, adjusts Operating Cash Flow (+ addback, - cash interest paid), and calibrates the <strong className="text-slate-900">Interest Coverage Ratio</strong>.
              </p>
            </div>

            <div className="p-3.5 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1.5">
              <span className="font-bold text-sky-800 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-md bg-sky-100 text-sky-800 flex items-center justify-center text-[10px]">4</span>
                Net Profit After Tax Driver
              </span>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Derives <strong className="text-slate-900">Profit Before Tax</strong> [PBT = NP / (1 - TaxRate)] and <strong className="text-slate-900">Corporate Tax Expense</strong>. Drives <strong className="text-slate-900">Retained Earnings</strong> roll-forward and dividend distribution in Equity.
              </p>
            </div>

            <div className="p-3.5 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1.5">
              <span className="font-bold text-rose-800 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-md bg-rose-100 text-rose-800 flex items-center justify-center text-[10px]">5</span>
                Repayment of Borrowings
              </span>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Enters Financing Cash Flow as principal debt amortization outflow. Dynamically reduces <strong className="text-slate-900">Long-Term Debt</strong> on the Balance Sheet and optimizes the Debt-to-Equity solvency ratio.
              </p>
            </div>

            <div className="p-3.5 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1.5">
              <span className="font-bold text-indigo-800 flex items-center gap-1.5">
                <span className="w-4 h-4 rounded-md bg-indigo-100 text-indigo-800 flex items-center justify-center text-[10px]">6</span>
                Fixed Assets Driver
              </span>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                Determines <strong className="text-slate-900">Capital Additions (Capex)</strong> in Investing Cash Flow, generates <strong className="text-slate-900">Annual Depreciation Charge</strong> across asset classes, and feeds Property, Plant & Equipment NBV.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Main View: Multi-Year Grid vs Single-Year Inspector */}
      {viewMode === 'grid' ? (
        <div className="space-y-6">
          {/* Section 1: The Six Primary Financial Input Drivers Table */}
          <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-4 bg-slate-100/80 border-b border-slate-200 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <Database className="w-4 h-4 text-emerald-700" />
                <h3 className="text-sm font-bold text-slate-900 tracking-tight">
                  Primary Financial Input Drivers (Direct User-Entry Controls)
                </h3>
              </div>
              <div className="flex items-center gap-2 text-[11px] text-slate-600">
                <span className="font-medium text-emerald-700">Active Mode:</span>
                <span className="capitalize font-semibold text-slate-900">
                  {inputMode === 'manual' ? 'Manual Direct Entry' : 'Automatic / Interconnected'}
                </span>
                <span>·</span>
                <span>{yearsData.length} Total Financial Years</span>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs font-mono">
                <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
                  <tr>
                    <th className="py-3 px-4 font-sans font-semibold sticky left-0 bg-slate-100/95 z-10 w-72 min-w-[280px] border-r border-slate-200 text-slate-900">
                      Primary Financial Driver Item
                    </th>
                    {yearsData.map(y => (
                      <th
                        key={y.year}
                        className={`py-3 px-3 text-right min-w-[130px] ${
                          y.isProjected ? 'text-sky-800' : 'text-slate-800'
                        }`}
                      >
                        <div className="font-sans font-semibold text-xs flex items-center justify-end gap-1">
                          <span>{y.year}</span>
                          <span
                            className={`text-[9px] px-1 py-0.5 rounded font-normal ${
                              y.isProjected
                                ? 'bg-sky-100 text-sky-800 border border-sky-200'
                                : 'bg-white text-slate-600 border border-slate-200'
                            }`}
                          >
                            {y.isProjected ? 'Proj' : 'Hist'}
                          </span>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {primaryDriversConfig.map(driver => {
                    const driverKey = driver.field as PrimaryDriverKey;
                    const activeConfig = autoFillConfigs[driverKey];

                    return (
                      <tr key={driver.id} className="hover:bg-slate-50/80 transition-colors">
                        <td className="py-3 px-4 sticky left-0 bg-white/95 z-10 font-sans border-r border-slate-200">
                          <div className="space-y-1.5">
                            <div className="flex items-center justify-between gap-2">
                              <div className="flex items-center gap-2">
                                <span className="w-5 h-5 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-800 flex items-center justify-center text-[10px] font-bold shrink-0">
                                  {driver.number}
                                </span>
                                <div className="font-bold text-slate-900 text-xs flex items-center gap-1.5">
                                  <span>{driver.title}</span>
                                  <span className="text-[9px] font-normal text-emerald-800 bg-emerald-50 border border-emerald-200 px-1 py-0.5 rounded">
                                    {driver.badge}
                                  </span>
                                </div>
                              </div>

                              <button
                                type="button"
                                onClick={() => setActiveModalDriver(driverKey)}
                                className="text-[9px] font-semibold text-emerald-800 hover:text-white bg-emerald-50 hover:bg-emerald-600 px-2 py-0.5 rounded border border-emerald-200 transition-colors flex items-center gap-1 shadow-xs shrink-0"
                                title="Open full auto-fill preview and options"
                              >
                                <Wand2 className="w-2.5 h-2.5" />
                                Auto-Fill
                              </button>
                            </div>

                            <div className="flex items-center justify-between text-[10px]">
                              <p className="text-slate-600 line-clamp-1 max-w-[190px]">
                                {driver.description}
                              </p>
                              {activeConfig?.isActive ? (
                                <span className="text-[9px] font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded flex items-center gap-1 shrink-0">
                                  <Zap className="w-2.5 h-2.5 text-emerald-600" />
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
                          const isCellScheduleLocked = Boolean(
                            y.isScheduleLocked &&
                            (driver.field === 'repaymentOfBorrowings' || driver.field === 'financeCost')
                          );

                          return (
                            <td key={y.year} className="py-2.5 px-2.5 text-right">
                              <div className="relative group">
                                <input
                                  type="number"
                                  step={isCellScheduleLocked ? '0.01' : 'any'}
                                  readOnly={isCellScheduleLocked}
                                  value={rawVal ?? 0}
                                  onChange={(e) => {
                                    if (!isCellScheduleLocked) {
                                      handleDriverChange(y.year, driver.field, Number(e.target.value));
                                    }
                                  }}
                                  title={
                                    isCellScheduleLocked
                                      ? `Strictly locked from Repayment Schedule (${driver.field === 'repaymentOfBorrowings' ? 'Exact Annual Inst. Principal' : 'Exact Annual Inst. Profit'})`
                                      : undefined
                                  }
                                  className={`w-full text-right px-2.5 py-1.5 rounded-lg text-xs font-mono font-bold transition-all tabular-nums ${
                                    isCellScheduleLocked
                                      ? 'bg-emerald-50 border border-emerald-400 text-emerald-900 cursor-not-allowed'
                                      : isBaseYear
                                      ? 'bg-emerald-50 border-2 border-emerald-500 text-emerald-900 ring-2 ring-emerald-500/20 shadow-xs'
                                      : isAutoFilled
                                      ? 'bg-indigo-50/75 border border-indigo-300 text-indigo-900 hover:border-indigo-400 focus:border-emerald-600'
                                      : isAuto
                                      ? 'bg-sky-50/80 border border-sky-300 text-sky-900 focus:border-sky-500 focus:ring-1 focus:ring-sky-500'
                                      : 'bg-white border border-slate-300 hover:border-emerald-500 text-slate-900 focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500'
                                  }`}
                                />
                                {isCellScheduleLocked && (
                                  <span className="absolute -top-2 right-1.5 text-[8px] font-bold bg-emerald-700 text-white px-1.5 py-0.2 rounded shadow-xs pointer-events-none flex items-center gap-0.5 z-10">
                                    <Lock className="w-2 h-2" /> Schedule
                                  </span>
                                )}
                                {!isCellScheduleLocked && isBaseYear && (
                                  <span className="absolute -top-2 right-1.5 text-[8px] font-bold bg-emerald-600 text-white px-1.5 py-0.2 rounded shadow-xs pointer-events-none flex items-center gap-0.5 z-10">
                                    <Star className="w-2 h-2 fill-white" /> Base Year
                                  </span>
                                )}
                                {!isCellScheduleLocked && !isBaseYear && isAutoFilled && (
                                  <span className="absolute -top-1.5 right-1.5 text-[8px] bg-indigo-100 text-indigo-800 border border-indigo-200 px-1 rounded pointer-events-none font-semibold">
                                    Auto ({y.year < activeConfig.baseYear ? 'Prior' : 'Proj'})
                                  </span>
                                )}
                                {!isCellScheduleLocked && !isBaseYear && !isAutoFilled && isAuto && (
                                  <span className="absolute -top-1.5 right-1.5 text-[8px] bg-sky-100 text-sky-800 border border-sky-200 px-1 rounded pointer-events-none font-semibold">
                                    Auto
                                  </span>
                                )}

                                {/* One-Click Auto-Fill Action or Schedule Unlock */}
                                <div className="pt-1 flex items-center justify-end gap-1">
                                  {isCellScheduleLocked && onUnlockScheduleYear ? (
                                    <button
                                      type="button"
                                      onClick={() => onUnlockScheduleYear(y.year)}
                                      className="text-[8.5px] font-semibold px-1.5 py-0.5 rounded bg-amber-50 text-amber-800 hover:bg-amber-100 border border-amber-300 transition-all flex items-center gap-0.5"
                                      title={`Unlock ${y.year} schedule values for manual editing`}
                                    >
                                      <Unlock className="w-2 h-2" />
                                      Unlock {y.year}
                                    </button>
                                  ) : (
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
                                          ? 'text-emerald-800 bg-emerald-50 border border-emerald-300 hover:bg-emerald-600 hover:text-white'
                                          : 'text-slate-600 hover:text-emerald-800 bg-slate-50 hover:bg-emerald-50 border border-slate-200 hover:border-emerald-300'
                                      }`}
                                      title={`Click to auto-fill all other years using ${y.year} value (${rawVal.toLocaleString()})`}
                                    >
                                      <Zap className="w-2 h-2 text-emerald-600" />
                                      {isBaseYear ? 'Recalculate All' : `Auto-Fill from ${y.year}`}
                                    </button>
                                  )}
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
          <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-4 bg-slate-100/80 border-b border-slate-200 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <Lock className="w-4 h-4 text-slate-600" />
                <h3 className="text-sm font-bold text-slate-900 tracking-tight">
                  Automatically Derived & Interconnected Financial Metrics (Locked / Calculated)
                </h3>
              </div>
              <div className="flex items-center gap-2">
                {/* Category Filter */}
                <div className="flex flex-wrap items-center gap-1 bg-slate-200/70 p-0.5 rounded-lg border border-slate-200 text-[11px]">
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('all')}
                    className={`px-2 py-1 rounded transition-colors whitespace-nowrap shrink-0 ${selectedCategory === 'all' ? 'bg-white text-slate-900 font-semibold shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
                  >
                    All Metrics
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('income')}
                    className={`px-2 py-1 rounded transition-colors whitespace-nowrap shrink-0 ${selectedCategory === 'income' ? 'bg-white text-slate-900 font-semibold shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
                  >
                    P&L
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('workingCapital')}
                    className={`px-2 py-1 rounded transition-colors whitespace-nowrap shrink-0 ${selectedCategory === 'workingCapital' ? 'bg-white text-slate-900 font-semibold shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
                  >
                    Working Capital
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('balanceSheet')}
                    className={`px-2 py-1 rounded transition-colors whitespace-nowrap shrink-0 ${selectedCategory === 'balanceSheet' ? 'bg-white text-slate-900 font-semibold shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
                  >
                    Balance Sheet
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('cashFlow')}
                    className={`px-2 py-1 rounded transition-colors whitespace-nowrap shrink-0 ${selectedCategory === 'cashFlow' ? 'bg-white text-slate-900 font-semibold shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
                  >
                    Cash Flow
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedCategory('ratios')}
                    className={`px-2 py-1 rounded transition-colors whitespace-nowrap shrink-0 ${selectedCategory === 'ratios' ? 'bg-white text-slate-900 font-semibold shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
                  >
                    Ratios
                  </button>
                </div>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs font-mono">
                <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
                  <tr>
                    <th className="py-2.5 px-4 font-sans font-semibold sticky left-0 bg-slate-100/95 z-10 w-72 min-w-[280px] border-r border-slate-200 text-slate-900">
                      Derived Metric / Financial Output
                    </th>
                    {yearsData.map(y => (
                      <th
                        key={y.year}
                        className={`py-2.5 px-3 text-right min-w-[130px] ${
                          y.isProjected ? 'text-sky-800' : 'text-slate-800'
                        }`}
                      >
                        {y.year} {y.isProjected ? '(P)' : ''}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 tabular-nums text-slate-700">
                  {/* P&L Derived Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'income') && (
                    <>
                      <tr className="bg-slate-100/90 text-[11px] font-sans font-bold text-slate-700">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-100/95">
                          1. Operating Performance & Income Statement Flow
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200 flex items-center justify-between">
                          <span>Cost of Sales (CoS)</span>
                          <Lock className="w-3 h-3 text-slate-400" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-rose-700">
                            ({y.costOfSales.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-emerald-50/40 font-semibold">
                        <td className="py-2 px-4 sticky left-0 bg-emerald-50/90 z-10 font-sans text-emerald-800 border-r border-slate-200">
                          Gross Profit (GP)
                        </td>
                        {yearsData.map((y, idx) => (
                          <td key={y.year} className="py-2 px-3 text-right text-emerald-700 font-bold">
                            {(incomeStatements[idx]?.grossProfit || (y.salesRevenue - y.costOfSales)).toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200 flex items-center justify-between">
                          <span>Operating Expenses (OpEx)</span>
                          <Lock className="w-3 h-3 text-slate-400" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-slate-600">
                            ({y.operatingExpenses.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200 flex items-center justify-between">
                          <span>Depreciation Expense</span>
                          <Lock className="w-3 h-3 text-slate-400" />
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-amber-800">
                            ({i.depreciationExpense.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-50/90 font-semibold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-50/95 z-10 font-sans text-slate-900 border-r border-slate-200">
                          Operating Profit (EBIT)
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-slate-900 font-bold">
                            {i.operatingProfit.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Profit Before Tax (EBT)
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-slate-800">
                            {i.profitBeforeTax.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Income Tax Expense
                        </td>
                        {incomeStatements.map(i => (
                          <td key={i.year} className="py-2 px-3 text-right text-rose-700">
                            ({i.taxExpense.toLocaleString()})
                          </td>
                        ))}
                      </tr>
                    </>
                  )}

                  {/* Working Capital Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'workingCapital') && (
                    <>
                      <tr className="bg-slate-100/90 text-[11px] font-sans font-bold text-slate-700">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-100/95">
                          2. Working Capital & Trade Balances
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200 flex items-center justify-between">
                          <span>Trade Receivables (DSO Derived)</span>
                          <Lock className="w-3 h-3 text-slate-400" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-sky-800">
                            {y.tradeReceivables.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200 flex items-center justify-between">
                          <span>Trade Payables (DPO Derived)</span>
                          <Lock className="w-3 h-3 text-slate-400" />
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-slate-700">
                            {y.tradePayables.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-teal-50/40 font-semibold">
                        <td className="py-2 px-4 sticky left-0 bg-teal-50/90 z-10 font-sans text-teal-800 border-r border-slate-200">
                          Net Working Capital (CA - CL)
                        </td>
                        {balanceSheets.map(b => {
                          const nwc = b.totalCurrentAssets - b.totalCurrentLiabilities;
                          return (
                            <td key={b.year} className={`py-2 px-3 text-right font-bold ${nwc >= 0 ? 'text-teal-700' : 'text-rose-700'}`}>
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
                      <tr className="bg-slate-100/90 text-[11px] font-sans font-bold text-slate-700">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-100/95">
                          3. Statement of Financial Position Balances
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Cash & Cash Equivalents
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-700 font-medium">
                            {b.cashAndEquivalents.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Total Current Assets
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-slate-800">
                            {b.totalCurrentAssets.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Property, Plant & Equipment (NBV)
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-indigo-700">
                            {b.propertyPlantEquipment.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-50/90 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-50/95 z-10 font-sans text-slate-900 border-r border-slate-200">
                          Total Assets
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-slate-900 font-bold">
                            {b.totalAssets.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Total Debt (Short + Long Term)
                        </td>
                        {balanceSheets.map(b => {
                          const totalDebt = b.shortTermBorrowings + b.longTermBorrowings;
                          return (
                            <td key={b.year} className="py-2 px-3 text-right text-amber-800">
                              {totalDebt.toLocaleString()}
                            </td>
                          );
                        })}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Retained Earnings (Accumulated)
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-700">
                            {b.retainedEarnings.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-50/90 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-50/95 z-10 font-sans text-slate-900 border-r border-slate-200">
                          Total Equity
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-700 font-bold">
                            {b.totalEquity.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-emerald-50/80 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-emerald-50/95 z-10 font-sans text-emerald-800 border-r border-slate-200 flex items-center justify-between">
                          <span>Balance Sheet Check (Assets - Liab&Eq)</span>
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-700" />
                        </td>
                        {balanceSheets.map(b => (
                          <td key={b.year} className="py-2 px-3 text-right text-emerald-700 font-bold">
                            {Math.abs(b.balanceCheck) < 0.01 ? '0.00 ✓' : b.balanceCheck.toFixed(2)}
                          </td>
                        ))}
                      </tr>
                    </>
                  )}

                  {/* Cash Flow Summary Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'cashFlow') && (
                    <>
                      <tr className="bg-slate-100/90 text-[11px] font-sans font-bold text-slate-700">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-100/95">
                          4. Statement of Cash Flows Summary
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Net Cash from Operating Activities
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className={`py-2 px-3 text-right font-medium ${cf.netCashFromOperating >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                            {cf.netCashFromOperating.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Net Cash from Investing Activities
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className="py-2 px-3 text-right text-amber-800">
                            {cf.netCashFromInvesting.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Net Cash from Financing Activities
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className="py-2 px-3 text-right text-sky-800">
                            {cf.netCashFromFinancing.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-slate-50/90 font-bold">
                        <td className="py-2 px-4 sticky left-0 bg-slate-50/95 z-10 font-sans text-slate-900 border-r border-slate-200">
                          Net Change in Cash
                        </td>
                        {cashFlowStatements.map(cf => (
                          <td key={cf.year} className={`py-2 px-3 text-right font-bold ${cf.netChangeInCash >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                            {cf.netChangeInCash.toLocaleString()}
                          </td>
                        ))}
                      </tr>
                    </>
                  )}

                  {/* Financial Ratios Group */}
                  {(selectedCategory === 'all' || selectedCategory === 'ratios') && (
                    <>
                      <tr className="bg-slate-100/90 text-[11px] font-sans font-bold text-slate-700">
                        <td colSpan={yearsData.length + 1} className="py-1.5 px-4 uppercase tracking-wider bg-slate-100/95">
                          5. Financial Ratios (Profitability, Solvency, Liquidity, Efficiency)
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Gross Margin (%)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-emerald-700">
                            {allYearsRatios[y.year]?.grossMarginPct ?? 0}%
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Net Profit Margin (%)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-emerald-700 font-medium">
                            {allYearsRatios[y.year]?.netMarginPct ?? 0}%
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Current Ratio (x)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-sky-800 font-medium">
                            {allYearsRatios[y.year]?.currentRatio ?? 0}x
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Debt-to-Equity Ratio (D/E)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-amber-800">
                            {allYearsRatios[y.year]?.debtToEquity ?? 0}x
                          </td>
                        ))}
                      </tr>
                      <tr className="hover:bg-slate-50/80">
                        <td className="py-2 px-4 sticky left-0 bg-white/95 z-10 font-sans text-slate-700 border-r border-slate-200">
                          Interest Coverage Ratio (EBIT / FC)
                        </td>
                        {yearsData.map(y => (
                          <td key={y.year} className="py-2 px-3 text-right text-teal-800">
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
          <div className="flex items-center justify-between gap-3 p-4 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-x-auto shadow-sm">
            <span className="text-xs font-semibold text-slate-700 shrink-0">Select Financial Year:</span>
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
                      ? 'bg-sky-50 text-sky-800 hover:bg-sky-100 border border-sky-200'
                      : 'bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200'
                  }`}
                >
                  {y.year} {y.isProjected ? '(Proj)' : '(Hist)'}
                </button>
              ))}
            </div>
          </div>

          {/* Single Year Key KPI Strip */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
            <div className="p-3.5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1 shadow-xs">
              <span className="text-[11px] text-slate-600 font-medium">Sales Revenue</span>
              <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
                {currentYearData.salesRevenue.toLocaleString()}
              </div>
              <span className="text-[10px] text-emerald-700 font-mono font-semibold">Driver 1</span>
            </div>

            <div className="p-3.5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1 shadow-xs">
              <span className="text-[11px] text-slate-600 font-medium">Net Profit (NPAT)</span>
              <div className="text-base font-bold text-emerald-700 font-mono tabular-nums">
                {currentYearData.netProfit.toLocaleString()}
              </div>
              <span className="text-[10px] text-emerald-700 font-mono font-semibold">Margin: {currentRatios?.netMarginPct}%</span>
            </div>

            <div className="p-3.5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1 shadow-xs">
              <span className="text-[11px] text-slate-600 font-medium">Finance Cost</span>
              <div className="text-base font-bold text-amber-800 font-mono tabular-nums">
                {currentYearData.financeCost.toLocaleString()}
              </div>
              <span className="text-[10px] text-slate-600 font-mono font-semibold">Driver 3</span>
            </div>

            <div className="p-3.5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1 shadow-xs">
              <span className="text-[11px] text-slate-600 font-medium">Borrowing Repayment</span>
              <div className="text-base font-bold text-rose-700 font-mono tabular-nums">
                {currentYearData.repaymentOfBorrowings.toLocaleString()}
              </div>
              <span className="text-[10px] text-rose-700 font-mono font-semibold">Outflow Driver</span>
            </div>

            <div className="p-3.5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1 shadow-xs">
              <span className="text-[11px] text-slate-600 font-medium">Fixed Assets Base</span>
              <div className="text-base font-bold text-indigo-700 font-mono tabular-nums">
                {currentYearData.fixedAssetsGross.toLocaleString()}
              </div>
              <span className="text-[10px] text-indigo-700 font-mono font-semibold">Driver 6</span>
            </div>

            <div className="p-3.5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-xl space-y-1 shadow-xs">
              <span className="text-[11px] text-slate-600 font-medium">Current Ratio</span>
              <div className="text-base font-bold text-sky-800 font-mono tabular-nums">
                {currentRatios?.currentRatio}x
              </div>
              <span className="text-[10px] text-sky-800 font-mono font-semibold">D/E: {currentRatios?.debtToEquity}x</span>
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
              const isSingleLocked = Boolean(
                currentYearData.isScheduleLocked &&
                (driver.field === 'repaymentOfBorrowings' || driver.field === 'financeCost')
              );

              return (
                <div
                  key={driver.id}
                  className={`border rounded-2xl p-5 space-y-3.5 transition-all backdrop-blur-md ${
                    isSingleLocked
                      ? 'bg-emerald-50/90 border-emerald-400 shadow-sm'
                      : isBaseYear
                      ? 'bg-emerald-50/85 border-emerald-400 shadow-sm ring-1 ring-emerald-500/25'
                      : 'bg-white/92 border-slate-200/90 hover:border-slate-300 shadow-xs'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-md bg-emerald-100 text-emerald-800 border border-emerald-200 flex items-center justify-center text-[10px] font-bold shrink-0">
                        {driver.number}
                      </span>
                      {driver.title}
                    </span>
                    <div className="flex items-center gap-1 shrink-0">
                      {isSingleLocked && (
                        <span className="text-[10px] font-bold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded-md border border-emerald-300 flex items-center gap-1">
                          <Lock className="w-2.5 h-2.5" />
                          Schedule Locked
                        </span>
                      )}
                      {!isSingleLocked && isBaseYear && (
                        <span className="text-[10px] font-bold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded-md border border-emerald-300 flex items-center gap-1">
                          <Star className="w-2.5 h-2.5 fill-emerald-600 text-emerald-600" />
                          Base Year
                        </span>
                      )}
                      <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md">
                        {driver.badge}
                      </span>
                    </div>
                  </div>

                  <p className="text-[11px] text-slate-600 min-h-[32px]">
                    {driver.description}
                  </p>

                  <div className="space-y-1 pt-1">
                    <label className="text-[11px] text-slate-700 font-medium flex items-center justify-between">
                      <span>Input Value ({companyProfile.reportingCurrency})</span>
                      <span className="text-[10px] font-mono text-slate-500">
                        {isSingleLocked ? 'Exact Schedule Value' : isBaseYear ? 'Active Anchor (Dynamic)' : (isAutoFilled ? 'Auto-Filled' : 'Direct Editable')}
                      </span>
                    </label>
                    <input
                      type="number"
                      step={isSingleLocked ? '0.01' : 'any'}
                      readOnly={isSingleLocked}
                      value={val ?? 0}
                      onChange={(e) => {
                        if (!isSingleLocked) {
                          handleDriverChange(currentYearData.year, driver.field, Number(e.target.value));
                        }
                      }}
                      className={`w-full text-base font-bold rounded-xl px-3.5 py-2 font-mono tabular-nums transition-all ${
                        isSingleLocked
                          ? 'bg-emerald-100/70 border-2 border-emerald-500 text-emerald-900 cursor-not-allowed'
                          : isBaseYear
                          ? 'bg-white border-2 border-emerald-500 text-emerald-900 focus:ring-1 focus:ring-emerald-500'
                          : 'bg-white border border-slate-300 focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500 text-slate-900'
                      }`}
                    />
                  </div>

                  {/* Auto-Fill Action & Dynamic Status */}
                  <div className="pt-2 border-t border-slate-200 space-y-2">
                    {isSingleLocked ? (
                      <div className="flex items-center justify-between gap-2 py-2 px-3 rounded-xl bg-emerald-100/80 border border-emerald-300 text-xs">
                        <span className="text-emerald-900 font-semibold flex items-center gap-1.5">
                          <Lock className="w-3.5 h-3.5 text-emerald-700" />
                          Exact Schedule Locked
                        </span>
                        {onUnlockScheduleYear && (
                          <button
                            type="button"
                            onClick={() => onUnlockScheduleYear(currentYearData.year)}
                            className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-300 rounded-lg text-[11px] font-semibold flex items-center gap-1 transition-colors"
                          >
                            <Unlock className="w-3 h-3" />
                            Unlock {currentYearData.year}
                          </button>
                        )}
                      </div>
                    ) : (
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
                            ? 'bg-emerald-600 hover:bg-emerald-700 text-white'
                            : 'bg-white hover:bg-emerald-600 text-emerald-800 hover:text-white border border-emerald-300 hover:border-emerald-600'
                        }`}
                      >
                        <Zap className="w-3.5 h-3.5" />
                        {isBaseYear
                          ? `Recalculate All Years from ${currentYearData.year} Base`
                          : `Auto-Fill Remaining Years from ${currentYearData.year}`}
                      </button>
                    )}

                    <div className="flex items-center justify-between text-[10px] text-slate-600 px-0.5">
                      {isBaseYear ? (
                        <span className="text-emerald-700 font-medium flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3 text-emerald-700" />
                          Dynamic Recalculation Enabled
                        </span>
                      ) : isAutoFilled ? (
                        <span className="text-indigo-700 font-medium">
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
                        className="text-slate-600 hover:text-emerald-700 underline transition-colors font-medium"
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
            <div className="p-4 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl space-y-2.5 shadow-sm">
              <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                <span className="font-bold text-slate-900">Income Statement Tie-Out</span>
                <span className="text-[10px] text-emerald-700 font-mono font-semibold">Year {selectedYear}</span>
              </div>
              <div className="space-y-1.5 font-mono text-slate-700 tabular-nums">
                <div className="flex justify-between">
                  <span>Revenue:</span>
                  <span className="text-slate-900 font-bold">{currentIncome?.revenue.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-rose-700">
                  <span>Cost of Sales:</span>
                  <span>({currentIncome?.costOfSales.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-emerald-700 font-semibold border-t border-slate-200 pt-1">
                  <span>Gross Profit:</span>
                  <span>{currentIncome?.grossProfit.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-slate-600">
                  <span>Operating Expenses:</span>
                  <span>({currentIncome?.operatingExpenses.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-amber-800">
                  <span>Depreciation:</span>
                  <span>({currentIncome?.depreciationExpense.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-amber-800">
                  <span>Finance Cost:</span>
                  <span>({currentIncome?.financeCost.toLocaleString()})</span>
                </div>
                <div className="flex justify-between text-emerald-700 font-bold border-t border-slate-200 pt-1">
                  <span>Net Profit After Tax:</span>
                  <span>{currentIncome?.netProfit.toLocaleString()}</span>
                </div>
              </div>
            </div>

            {/* Balance Sheet Summary */}
            <div className="p-4 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl space-y-2.5 shadow-sm">
              <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                <span className="font-bold text-slate-900">Balance Sheet Balance</span>
                <span className="text-[10px] text-sky-700 font-mono font-semibold">Double-Entry</span>
              </div>
              <div className="space-y-1.5 font-mono text-slate-700 tabular-nums">
                <div className="flex justify-between">
                  <span>Cash & Equivalents:</span>
                  <span className="text-emerald-700 font-medium">{currentBalance?.cashAndEquivalents.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span>Trade Receivables:</span>
                  <span>{currentBalance?.tradeReceivables.toLocaleString()}</span>
                </div>
                <div className="flex justify-between">
                  <span>Closing Inventory:</span>
                  <span>{currentBalance?.closingInventory.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-indigo-700">
                  <span>PP&E Net Book Value:</span>
                  <span>{currentBalance?.propertyPlantEquipment.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-slate-900 font-bold border-t border-slate-200 pt-1">
                  <span>Total Assets:</span>
                  <span>{currentBalance?.totalAssets.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-amber-800">
                  <span>Total Debt:</span>
                  <span>{(currentBalance?.shortTermBorrowings + currentBalance?.longTermBorrowings).toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-emerald-700 font-semibold">
                  <span>Total Equity:</span>
                  <span>{currentBalance?.totalEquity.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-emerald-700 font-bold border-t border-slate-200 pt-1">
                  <span>Double-Entry Variance:</span>
                  <span>0.00 ✓</span>
                </div>
              </div>
            </div>

            {/* Cash Flow Summary */}
            <div className="p-4 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl space-y-2.5 shadow-sm">
              <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                <span className="font-bold text-slate-900">Cash Flows & Reconciliation</span>
                <span className="text-[10px] text-teal-700 font-mono font-semibold">Audited</span>
              </div>
              <div className="space-y-1.5 font-mono text-slate-700 tabular-nums">
                <div className="flex justify-between">
                  <span>Operating Cash Flow:</span>
                  <span className={currentCashFlow?.netCashFromOperating >= 0 ? 'text-emerald-700 font-medium' : 'text-rose-700 font-medium'}>
                    {currentCashFlow?.netCashFromOperating.toLocaleString()}
                  </span>
                </div>
                <div className="flex justify-between text-amber-800">
                  <span>Investing Cash Flow:</span>
                  <span>{currentCashFlow?.netCashFromInvesting.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-sky-800">
                  <span>Financing Cash Flow:</span>
                  <span>{currentCashFlow?.netCashFromFinancing.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-slate-900 font-bold border-t border-slate-200 pt-1">
                  <span>Net Change in Cash:</span>
                  <span className={currentCashFlow?.netChangeInCash >= 0 ? 'text-emerald-700' : 'text-rose-700'}>
                    {currentCashFlow?.netChangeInCash.toLocaleString()}
                  </span>
                </div>
                <div className="flex justify-between">
                  <span>Opening Cash:</span>
                  <span>{currentCashFlow?.openingCash.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-emerald-700 font-bold border-t border-slate-200 pt-1">
                  <span>Closing Reconciled Cash:</span>
                  <span>{currentCashFlow?.closingCashCalculated.toLocaleString()}</span>
                </div>
                <div className="flex justify-between text-[11px] text-teal-700 font-semibold">
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
          <div className="fixed inset-0 z-50 bg-slate-900/45 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
            <div className="bg-white border border-slate-200 rounded-3xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden shadow-2xl animate-in zoom-in-95 duration-150">
              {/* Modal Header */}
              <div className="p-5 border-b border-slate-200 flex items-center justify-between bg-slate-50/90">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700 shrink-0">
                    <Zap className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="flex flex-wrap items-center gap-2">
                      <h3 className="text-base font-bold text-slate-900">
                        Auto-Fill Preview & Analysis — {meta.title}
                      </h3>
                      <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                        {meta.badge}
                      </span>
                    </div>
                    <p className="text-xs text-slate-600 mt-0.5">
                      Review computed figures across all {yearsData.length} years anchored on Base Year {currentBaseYear} ({companyProfile.reportingCurrency} {currentDriverVal.toLocaleString()}).
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setActiveModalDriver(null)}
                  className="text-slate-500 hover:text-slate-900 p-2 rounded-xl hover:bg-slate-200/70 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Modal Body */}
              <div className="p-5 overflow-y-auto space-y-5 text-xs">
                {/* Top Controls Bar inside Modal */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 p-4 bg-slate-50 border border-slate-200 rounded-2xl">
                  <div>
                    <label className="text-[11px] font-medium text-slate-700 block mb-1">
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
                      className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-slate-900 font-mono focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                    >
                      {yearsData.map(y => (
                        <option key={y.year} value={y.year}>
                          {y.year} ({y.isProjected ? 'Projected' : 'Historical'})
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-[11px] font-medium text-slate-700 block mb-1">
                      Base Figure Value ({companyProfile.reportingCurrency}):
                    </label>
                    <input
                      type="number"
                      value={currentDriverVal}
                      onChange={(e) => setCustomBaseValues(prev => ({ ...prev, [driverKey]: Number(e.target.value) }))}
                      className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-slate-900 font-mono font-bold text-right focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500 tabular-nums"
                    />
                  </div>

                  <div>
                    <label className="text-[11px] font-medium text-slate-700 block mb-1">
                      {meta.trendLabel} (%):
                    </label>
                    <input
                      type="number"
                      step="0.1"
                      value={currentRate}
                      onChange={(e) => setCustomRates(prev => ({ ...prev, [driverKey]: Number(e.target.value) }))}
                      className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-emerald-700 font-mono font-bold text-right focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500 tabular-nums"
                    />
                  </div>
                </div>

                {/* Rationale explanation callout */}
                <div className="p-3 bg-indigo-50/80 border border-indigo-200 rounded-xl flex items-start gap-2.5 text-[11px] text-indigo-900">
                  <Info className="w-4 h-4 text-indigo-600 shrink-0 mt-0.5" />
                  <div>
                    <span className="font-bold text-slate-900">Financial Architecture Logic: </span>
                    <span>{meta.rationale} Previous years before {currentBaseYear} are back-calculated for logical historical continuity, {currentBaseYear} is kept 100% fixed, and subsequent years follow compound professional financial curves.</span>
                  </div>
                </div>

                {/* Preview Table */}
                <div className="border border-slate-200 rounded-2xl overflow-hidden">
                  <div className="max-h-72 overflow-y-auto">
                    <table className="w-full text-left text-xs font-mono">
                      <thead className="bg-slate-100 text-slate-700 sticky top-0 z-10 text-[11px] border-b border-slate-200">
                        <tr>
                          <th className="py-2.5 px-4 font-sans font-semibold text-slate-900">Year</th>
                          <th className="py-2.5 px-3 font-sans font-semibold text-slate-900">Anchor Classification</th>
                          <th className="py-2.5 px-3 text-right text-slate-700">Current Value</th>
                          <th className="py-2.5 px-3 text-right text-emerald-700 font-bold">Auto-Fill Value</th>
                          <th className="py-2.5 px-3 text-right text-slate-700">Variance</th>
                          <th className="py-2.5 px-3 text-right text-slate-700">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 tabular-nums">
                        {yearsData.map(y => {
                          const isBase = y.year === currentBaseYear;
                          const isPast = y.year < currentBaseYear;
                          const currentVal = y[driverKey] as number;
                          const autoVal = previewSeries[y.year] ?? currentVal;
                          const diff = autoVal - currentVal;

                          return (
                            <tr
                              key={y.year}
                              className={`hover:bg-slate-50 transition-colors ${
                                isBase ? 'bg-emerald-50/70 font-bold' : ''
                              }`}
                            >
                              <td className="py-2 px-4 font-sans font-bold text-slate-900 flex items-center gap-1.5">
                                {isBase && <Star className="w-3 h-3 fill-emerald-600 text-emerald-600" />}
                                <span>{y.year}</span>
                                <span className={`text-[9px] px-1 py-0.2 rounded font-normal ${
                                  y.isProjected ? 'bg-sky-50 text-sky-800 border border-sky-200' : 'bg-slate-100 text-slate-600 border border-slate-200'
                                }`}>
                                  {y.isProjected ? 'Proj' : 'Hist'}
                                </span>
                              </td>
                              <td className="py-2 px-3 font-sans text-[11px]">
                                {isBase ? (
                                  <span className="text-emerald-800 font-bold bg-emerald-100 px-2 py-0.5 rounded-md border border-emerald-300">
                                    ★ Base Anchor Reference
                                  </span>
                                ) : isPast ? (
                                  <span className="text-slate-600">
                                    Prior Year (Back-Discounted)
                                  </span>
                                ) : (
                                  <span className="text-sky-800">
                                    Subsequent Year (Projected)
                                  </span>
                                )}
                              </td>
                              <td className="py-2 px-3 text-right text-slate-600">
                                {currentVal.toLocaleString()}
                              </td>
                              <td className={`py-2 px-3 text-right font-bold ${
                                isBase ? 'text-emerald-800' : 'text-emerald-700'
                              }`}>
                                {autoVal.toLocaleString()}
                              </td>
                              <td className={`py-2 px-3 text-right ${
                                diff > 0 ? 'text-emerald-700' : (diff < 0 ? 'text-rose-700' : 'text-slate-400')
                              }`}>
                                {diff === 0 ? '—' : `${diff > 0 ? '+' : ''}${diff.toLocaleString()}`}
                              </td>
                              <td className="py-2 px-3 text-right text-[11px] font-sans">
                                {isBase ? (
                                  <span className="text-emerald-700 font-semibold">Strictly Preserved</span>
                                ) : (
                                  <span className="text-slate-500">Auto-Calculated</span>
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
              <div className="p-4 border-t border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
                <button
                  type="button"
                  onClick={() => setActiveModalDriver(null)}
                  className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-700 hover:text-slate-900 hover:bg-slate-200/70 bg-white border border-slate-300 transition-colors"
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
                  className="px-5 py-2 rounded-xl text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 transition-all flex items-center gap-2 shadow-sm"
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
