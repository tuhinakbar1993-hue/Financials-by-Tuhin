import React from 'react';
import { Sliders, TrendingUp, Calendar, Zap, Shield, RotateCcw } from 'lucide-react';
import { ProjectionAssumptions } from '../types/financial';
import { YearSelectorControl } from './YearSelectorControl';

interface ProjectionEngineViewProps {
  assumptions: ProjectionAssumptions;
  onUpdateAssumptions: (updated: ProjectionAssumptions) => void;
  onApplyAndRecalculate: () => void;
  onResetDefaults: () => void;
  selectedHistoricalYears?: number[];
  selectedProjectedYears?: number[];
  onToggleHistoricalYear?: (year: number) => void;
  onToggleProjectedYear?: (year: number) => void;
  onSetHistoricalYears?: (years: number[]) => void;
  onSetProjectedYears?: (years: number[]) => void;
}

export const ProjectionEngineView: React.FC<ProjectionEngineViewProps> = ({
  assumptions,
  onUpdateAssumptions,
  onApplyAndRecalculate,
  onResetDefaults,
  selectedHistoricalYears = [2023, 2024, 2025],
  selectedProjectedYears,
  onToggleHistoricalYear,
  onToggleProjectedYear,
  onSetHistoricalYears,
  onSetProjectedYears,
}) => {
  const handleChange = (field: keyof ProjectionAssumptions, val: number) => {
    onUpdateAssumptions({
      ...assumptions,
      [field]: val,
    });
  };

  const sortedHist = [...selectedHistoricalYears].sort((a, b) => a - b);
  const maxHist = sortedHist[sortedHist.length - 1] || 2025;
  const projYears = selectedProjectedYears && selectedProjectedYears.length > 0
    ? [...selectedProjectedYears].sort((a, b) => a - b)
    : Array.from({ length: assumptions.projectedYearsCount }, (_, i) => maxHist + 1 + i);

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 p-5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl shadow-sm">
        <div>
          <div className="flex items-center gap-2">
            <Sliders className="w-5 h-5 text-amber-600" />
            <h2 className="text-lg font-bold text-slate-900 tracking-tight">
              Financial Projection, Horizon & Year Selection Engine
            </h2>
            <span className="text-[11px] font-semibold text-amber-900 bg-amber-50 border border-amber-300 px-2 py-0.5 rounded-md">
              Section 76.8
            </span>
          </div>
          <p className="text-xs text-slate-600 mt-1">
            Configure previous ({sortedHist.length}Y) and projected ({projYears.length}Y) financial years individually, plus revenue escalation, cost structures, and working capital cycles across all five financial statements.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={onResetDefaults}
            className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 rounded-lg border border-slate-300 transition-colors flex items-center gap-1.5"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Reset Baseline
          </button>
          <button
            onClick={onApplyAndRecalculate}
            className="px-4 py-1.5 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-500 rounded-lg shadow-sm transition-colors flex items-center gap-1.5"
          >
            <Zap className="w-3.5 h-3.5" />
            Run Projection Engine
          </button>
        </div>
      </div>

      {onToggleHistoricalYear && onToggleProjectedYear && onSetHistoricalYears && onSetProjectedYears && (
        <YearSelectorControl
          selectedHistoricalYears={sortedHist}
          selectedProjectedYears={projYears}
          onToggleHistoricalYear={onToggleHistoricalYear}
          onToggleProjectedYear={onToggleProjectedYear}
          onSetHistoricalYears={onSetHistoricalYears}
          onSetProjectedYears={onSetProjectedYears}
        />
      )}

      {/* Assumptions Cards Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
        {/* Card 1: Time Horizon */}
        <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl p-5 space-y-4 shadow-sm">
          <div className="flex items-center gap-2 text-slate-900 text-xs font-bold border-b border-slate-200 pb-2">
            <Calendar className="w-4 h-4 text-emerald-600" />
            Projection Horizon Configuration
          </div>
          <div className="space-y-3 text-xs">
            <div className="space-y-1">
              <label className="text-slate-700 font-medium flex justify-between">
                <span>Select Projected Years Horizon (1 to 20 Years)</span>
                <span className="font-mono text-emerald-700 font-bold">{projYears.length} Years</span>
              </label>
              <input
                type="range"
                min={1}
                max={20}
                value={Math.min(20, Math.max(1, projYears.length))}
                onChange={(e) => handleChange('projectedYearsCount', Number(e.target.value))}
                className="w-full accent-emerald-600"
              />
              <div className="flex justify-between text-[10px] text-slate-500 font-mono">
                <span>1 Year</span>
                <span>5 Years</span>
                <span>10 Years</span>
                <span>15 Years</span>
                <span>20 Years</span>
              </div>
            </div>

            {/* Quick buttons */}
            <div className="flex flex-wrap gap-1 pt-1">
              {[1, 3, 5, 7, 10, 12, 15, 20].map(yr => (
                <button
                  key={yr}
                  type="button"
                  onClick={() => handleChange('projectedYearsCount', yr)}
                  className={`px-2 py-0.5 rounded text-[11px] font-semibold transition-colors border ${
                    projYears.length === yr
                      ? 'bg-emerald-600 text-white border-emerald-600'
                      : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  {yr}Y
                </button>
              ))}
            </div>

            <div className="space-y-1 bg-slate-50/90 border border-slate-200 rounded-lg p-3">
              <div className="flex justify-between text-slate-700 font-medium">
                <span>Selected Previous Years ({sortedHist.length})</span>
                <span className="font-mono text-emerald-700 font-bold">{sortedHist.join(', ')}</span>
              </div>
              <div className="flex justify-between text-slate-700 font-medium pt-1">
                <span>Selected Projected Years ({projYears.length})</span>
                <span className="font-mono text-sky-700 font-bold truncate max-w-[170px]" title={projYears.join(', ')}>
                  {projYears[0]}–{projYears[projYears.length - 1]}
                </span>
              </div>
            </div>
          </div>
        </div>

        {/* Card 2: Revenue & Margin Projections */}
        <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl p-5 space-y-4 shadow-sm">
          <div className="flex items-center gap-2 text-slate-900 text-xs font-bold border-b border-slate-200 pb-2">
            <TrendingUp className="w-4 h-4 text-sky-600" />
            Revenue & Profit Margins
          </div>
          <div className="space-y-3 text-xs">
            <div className="space-y-1">
              <label className="text-slate-700 font-medium flex justify-between">
                <span>Annual Revenue Growth Rate</span>
                <span className="font-mono text-sky-700 font-bold">{assumptions.revenueGrowthRate}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.revenueGrowthRate}
                onChange={(e) => handleChange('revenueGrowthRate', Number(e.target.value))}
                className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900 font-mono tabular-nums focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            </div>

            <div className="space-y-1">
              <label className="text-slate-700 font-medium flex justify-between">
                <span>Target Gross Profit Margin (%)</span>
                <span className="font-mono text-slate-800 font-semibold">{assumptions.grossMarginPct}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.grossMarginPct}
                onChange={(e) => handleChange('grossMarginPct', Number(e.target.value))}
                className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900 font-mono tabular-nums focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            </div>

            <div className="space-y-1">
              <label className="text-slate-700 font-medium flex justify-between">
                <span>Operating Expenses (% of Revenue)</span>
                <span className="font-mono text-slate-800 font-semibold">{assumptions.opexPctOfRevenue}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.opexPctOfRevenue}
                onChange={(e) => handleChange('opexPctOfRevenue', Number(e.target.value))}
                className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900 font-mono tabular-nums focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            </div>
          </div>
        </div>

        {/* Card 3: Working Capital & Capex Cycles */}
        <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl p-5 space-y-4 shadow-sm">
          <div className="flex items-center gap-2 text-slate-900 text-xs font-bold border-b border-slate-200 pb-2">
            <Shield className="w-4 h-4 text-teal-600" />
            Working Capital & Capex Metrics
          </div>
          <div className="space-y-3 text-xs">
            <div className="grid grid-cols-3 gap-2">
              <div className="space-y-1">
                <label className="text-[11px] text-slate-600 font-medium">DSO (Receivables)</label>
                <input
                  type="number"
                  value={assumptions.receivableDays}
                  onChange={(e) => handleChange('receivableDays', Number(e.target.value))}
                  className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-slate-900 font-mono tabular-nums text-center focus:outline-none focus:ring-1 focus:ring-emerald-500"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[11px] text-slate-600 font-medium">DIO (Inventory)</label>
                <input
                  type="number"
                  value={assumptions.inventoryDays}
                  onChange={(e) => handleChange('inventoryDays', Number(e.target.value))}
                  className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-slate-900 font-mono tabular-nums text-center focus:outline-none focus:ring-1 focus:ring-emerald-500"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[11px] text-slate-600 font-medium">DPO (Payables)</label>
                <input
                  type="number"
                  value={assumptions.payableDays}
                  onChange={(e) => handleChange('payableDays', Number(e.target.value))}
                  className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-slate-900 font-mono tabular-nums text-center focus:outline-none focus:ring-1 focus:ring-emerald-500"
                />
              </div>
            </div>

            <div className="space-y-1">
              <label className="text-slate-700 font-medium flex justify-between">
                <span>Capex Additions (% of Revenue)</span>
                <span className="font-mono text-slate-800 font-semibold">{assumptions.capexPctOfRevenue}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.capexPctOfRevenue}
                onChange={(e) => handleChange('capexPctOfRevenue', Number(e.target.value))}
                className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900 font-mono tabular-nums focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            </div>

            <div className="space-y-1">
              <label className="text-slate-700 font-medium flex justify-between">
                <span>Dividend Payout Ratio (%)</span>
                <span className="font-mono text-slate-800 font-semibold">{assumptions.dividendPayoutRatio}%</span>
              </label>
              <input
                type="number"
                value={assumptions.dividendPayoutRatio}
                onChange={(e) => handleChange('dividendPayoutRatio', Number(e.target.value))}
                className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900 font-mono tabular-nums focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            </div>

            <div className="space-y-1 pt-1 border-t border-slate-200/80">
              <label className="text-slate-800 font-semibold flex justify-between">
                <span>Loan Rescheduling Interest Rate (% p.a.)</span>
                <span className="font-mono text-amber-800 font-bold">{(assumptions.loanInterestRatePct ?? 15.5).toFixed(2)}%</span>
              </label>
              <input
                type="number"
                step="0.05"
                min={0}
                max={45}
                value={assumptions.loanInterestRatePct ?? 15.5}
                onChange={(e) => handleChange('loanInterestRatePct', Math.max(0, Number(e.target.value) || 0))}
                aria-label="Loan Rescheduling Interest Rate (%)"
                className="w-full bg-amber-50/50 border border-amber-300 rounded-lg px-3 py-1.5 text-slate-900 font-mono font-bold tabular-nums focus:outline-none focus:ring-1 focus:ring-amber-500"
              />
            </div>
          </div>
        </div>
      </div>

      {/* Auto-Fill Feature Highlight Card */}
      <div className="bg-gradient-to-r from-emerald-50/90 via-white/95 to-white/95 backdrop-blur-md border border-emerald-200 rounded-2xl p-5 space-y-3 shadow-sm">
        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-100 border border-emerald-200 flex items-center justify-center text-emerald-700">
              <Zap className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                Automatic Year-Wise Auto-Fill Facility
                <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-100 border border-emerald-300 px-2 py-0.5 rounded-md">
                  Active in Data Entry
                </span>
              </h3>
              <p className="text-xs text-slate-600 mt-0.5">
                Each of the six primary drivers (Sales Revenue, Closing Inventory, Finance Cost, Net Profit After Tax, Repayment of Borrowings, Fixed Assets) supports independent single-year base entry with automatic calculation of all previous and subsequent years.
              </p>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 pt-2 text-[11px] font-mono">
          <div className="bg-white/90 p-2.5 rounded-xl border border-slate-200 space-y-0.5">
            <span className="text-slate-500 text-[10px]">Driver 1</span>
            <div className="font-bold text-emerald-700 truncate">Sales Revenue</div>
            <div className="text-[10px] text-slate-500 font-sans">Compound growth</div>
          </div>
          <div className="bg-white/90 p-2.5 rounded-xl border border-slate-200 space-y-0.5">
            <span className="text-slate-500 text-[10px]">Driver 2</span>
            <div className="font-bold text-teal-700 truncate">Closing Inventory</div>
            <div className="text-[10px] text-slate-500 font-sans">DIO turnover scaling</div>
          </div>
          <div className="bg-white/90 p-2.5 rounded-xl border border-slate-200 space-y-0.5">
            <span className="text-slate-500 text-[10px]">Driver 3</span>
            <div className="font-bold text-amber-800 truncate">Finance Cost</div>
            <div className="text-[10px] text-slate-500 font-sans">Debt deleveraging</div>
          </div>
          <div className="bg-white/90 p-2.5 rounded-xl border border-slate-200 space-y-0.5">
            <span className="text-slate-500 text-[10px]">Driver 4</span>
            <div className="font-bold text-sky-700 truncate">Net Profit</div>
            <div className="text-[10px] text-slate-500 font-sans">Operating leverage</div>
          </div>
          <div className="bg-white/90 p-2.5 rounded-xl border border-slate-200 space-y-0.5">
            <span className="text-slate-500 text-[10px]">Driver 5</span>
            <div className="font-bold text-rose-700 truncate">Repay Borrowings</div>
            <div className="text-[10px] text-slate-500 font-sans">Facility amortization</div>
          </div>
          <div className="bg-white/90 p-2.5 rounded-xl border border-slate-200 space-y-0.5">
            <span className="text-slate-500 text-[10px]">Driver 6</span>
            <div className="font-bold text-indigo-700 truncate">Fixed Assets</div>
            <div className="text-[10px] text-slate-500 font-sans">Capex reinvestment</div>
          </div>
        </div>
      </div>
    </div>
  );
};
