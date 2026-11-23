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
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 p-5 bg-slate-900 border border-slate-800 rounded-2xl">
        <div>
          <div className="flex items-center gap-2">
            <Sliders className="w-5 h-5 text-amber-400" />
            <h2 className="text-lg font-bold text-white tracking-tight">
              Financial Projection, Horizon & Year Selection Engine
            </h2>
            <span className="text-[11px] font-semibold text-amber-300 bg-amber-950/60 border border-amber-800/50 px-2 py-0.5 rounded-md">
              Section 76.8
            </span>
          </div>
          <p className="text-xs text-slate-400 mt-1">
            Configure previous ({sortedHist.length}Y) and projected ({projYears.length}Y) financial years individually, plus revenue escalation, cost structures, and working capital cycles across all five financial statements.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={onResetDefaults}
            className="px-3 py-1.5 text-xs font-medium text-slate-300 bg-slate-800 hover:bg-slate-700 rounded-lg border border-slate-700 transition-colors flex items-center gap-1.5"
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
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
          <div className="flex items-center gap-2 text-slate-200 text-xs font-bold border-b border-slate-800 pb-2">
            <Calendar className="w-4 h-4 text-emerald-400" />
            Projection Horizon Configuration
          </div>
          <div className="space-y-3 text-xs">
            <div className="space-y-1">
              <label className="text-slate-300 font-medium flex justify-between">
                <span>Select Projected Years Horizon (1 to 20 Years)</span>
                <span className="font-mono text-emerald-400 font-bold">{projYears.length} Years</span>
              </label>
              <input
                type="range"
                min={1}
                max={20}
                value={Math.min(20, Math.max(1, projYears.length))}
                onChange={(e) => handleChange('projectedYearsCount', Number(e.target.value))}
                className="w-full accent-emerald-500"
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
                  className={`px-2 py-0.5 rounded text-[11px] font-semibold transition-colors ${
                    projYears.length === yr
                      ? 'bg-emerald-600 text-white'
                      : 'bg-slate-800 text-slate-400 hover:text-slate-200'
                  }`}
                >
                  {yr}Y
                </button>
              ))}
            </div>

            <div className="space-y-1 bg-slate-950/60 border border-slate-800 rounded-lg p-3">
              <div className="flex justify-between text-slate-300 font-medium">
                <span>Selected Previous Years ({sortedHist.length})</span>
                <span className="font-mono text-emerald-400 font-bold">{sortedHist.join(', ')}</span>
              </div>
              <div className="flex justify-between text-slate-300 font-medium pt-1">
                <span>Selected Projected Years ({projYears.length})</span>
                <span className="font-mono text-sky-400 font-bold truncate max-w-[170px]" title={projYears.join(', ')}>
                  {projYears[0]}–{projYears[projYears.length - 1]}
                </span>
              </div>
            </div>
          </div>
        </div>

        {/* Card 2: Revenue & Margin Projections */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
          <div className="flex items-center gap-2 text-slate-200 text-xs font-bold border-b border-slate-800 pb-2">
            <TrendingUp className="w-4 h-4 text-sky-400" />
            Revenue & Profit Margins
          </div>
          <div className="space-y-3 text-xs">
            <div className="space-y-1">
              <label className="text-slate-300 font-medium flex justify-between">
                <span>Annual Revenue Growth Rate</span>
                <span className="font-mono text-sky-400 font-bold">{assumptions.revenueGrowthRate}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.revenueGrowthRate}
                onChange={(e) => handleChange('revenueGrowthRate', Number(e.target.value))}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-white font-mono tabular-nums"
              />
            </div>

            <div className="space-y-1">
              <label className="text-slate-300 font-medium flex justify-between">
                <span>Target Gross Profit Margin (%)</span>
                <span className="font-mono text-slate-200">{assumptions.grossMarginPct}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.grossMarginPct}
                onChange={(e) => handleChange('grossMarginPct', Number(e.target.value))}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-white font-mono tabular-nums"
              />
            </div>

            <div className="space-y-1">
              <label className="text-slate-300 font-medium flex justify-between">
                <span>Operating Expenses (% of Revenue)</span>
                <span className="font-mono text-slate-200">{assumptions.opexPctOfRevenue}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.opexPctOfRevenue}
                onChange={(e) => handleChange('opexPctOfRevenue', Number(e.target.value))}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-white font-mono tabular-nums"
              />
            </div>
          </div>
        </div>

        {/* Card 3: Working Capital & Capex Cycles */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
          <div className="flex items-center gap-2 text-slate-200 text-xs font-bold border-b border-slate-800 pb-2">
            <Shield className="w-4 h-4 text-teal-400" />
            Working Capital & Capex Metrics
          </div>
          <div className="space-y-3 text-xs">
            <div className="grid grid-cols-3 gap-2">
              <div className="space-y-1">
                <label className="text-[11px] text-slate-400">DSO (Receivables)</label>
                <input
                  type="number"
                  value={assumptions.receivableDays}
                  onChange={(e) => handleChange('receivableDays', Number(e.target.value))}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-2 py-1 text-white font-mono tabular-nums text-center"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[11px] text-slate-400">DIO (Inventory)</label>
                <input
                  type="number"
                  value={assumptions.inventoryDays}
                  onChange={(e) => handleChange('inventoryDays', Number(e.target.value))}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-2 py-1 text-white font-mono tabular-nums text-center"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[11px] text-slate-400">DPO (Payables)</label>
                <input
                  type="number"
                  value={assumptions.payableDays}
                  onChange={(e) => handleChange('payableDays', Number(e.target.value))}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-2 py-1 text-white font-mono tabular-nums text-center"
                />
              </div>
            </div>

            <div className="space-y-1">
              <label className="text-slate-300 font-medium flex justify-between">
                <span>Capex Additions (% of Revenue)</span>
                <span className="font-mono text-slate-200">{assumptions.capexPctOfRevenue}%</span>
              </label>
              <input
                type="number"
                step="0.1"
                value={assumptions.capexPctOfRevenue}
                onChange={(e) => handleChange('capexPctOfRevenue', Number(e.target.value))}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-white font-mono tabular-nums"
              />
            </div>

            <div className="space-y-1">
              <label className="text-slate-300 font-medium flex justify-between">
                <span>Dividend Payout Ratio (%)</span>
                <span className="font-mono text-slate-200">{assumptions.dividendPayoutRatio}%</span>
              </label>
              <input
                type="number"
                value={assumptions.dividendPayoutRatio}
                onChange={(e) => handleChange('dividendPayoutRatio', Number(e.target.value))}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-white font-mono tabular-nums"
              />
            </div>
          </div>
        </div>
      </div>

      {/* Auto-Fill Feature Highlight Card */}
      <div className="bg-gradient-to-r from-emerald-950/40 via-slate-900 to-slate-900 border border-emerald-800/40 rounded-2xl p-5 space-y-3">
        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
              <Zap className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-white flex items-center gap-2">
                Automatic Year-Wise Auto-Fill Facility
                <span className="text-[10px] font-semibold text-emerald-400 bg-emerald-950/80 border border-emerald-700/50 px-2 py-0.5 rounded-md">
                  Active in Data Entry
                </span>
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                Each of the six primary drivers (Sales Revenue, Closing Inventory, Finance Cost, Net Profit After Tax, Repayment of Borrowings, Fixed Assets) supports independent single-year base entry with automatic calculation of all previous and subsequent years.
              </p>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 pt-2 text-[11px] font-mono">
          <div className="bg-slate-950/70 p-2.5 rounded-xl border border-slate-800 space-y-0.5">
            <span className="text-slate-400 text-[10px]">Driver 1</span>
            <div className="font-bold text-emerald-300 truncate">Sales Revenue</div>
            <div className="text-[10px] text-slate-500 font-sans">Compound growth</div>
          </div>
          <div className="bg-slate-950/70 p-2.5 rounded-xl border border-slate-800 space-y-0.5">
            <span className="text-slate-400 text-[10px]">Driver 2</span>
            <div className="font-bold text-teal-300 truncate">Closing Inventory</div>
            <div className="text-[10px] text-slate-500 font-sans">DIO turnover scaling</div>
          </div>
          <div className="bg-slate-950/70 p-2.5 rounded-xl border border-slate-800 space-y-0.5">
            <span className="text-slate-400 text-[10px]">Driver 3</span>
            <div className="font-bold text-amber-300 truncate">Finance Cost</div>
            <div className="text-[10px] text-slate-500 font-sans">Debt deleveraging</div>
          </div>
          <div className="bg-slate-950/70 p-2.5 rounded-xl border border-slate-800 space-y-0.5">
            <span className="text-slate-400 text-[10px]">Driver 4</span>
            <div className="font-bold text-sky-300 truncate">Net Profit</div>
            <div className="text-[10px] text-slate-500 font-sans">Operating leverage</div>
          </div>
          <div className="bg-slate-950/70 p-2.5 rounded-xl border border-slate-800 space-y-0.5">
            <span className="text-slate-400 text-[10px]">Driver 5</span>
            <div className="font-bold text-rose-300 truncate">Repay Borrowings</div>
            <div className="text-[10px] text-slate-500 font-sans">Facility amortization</div>
          </div>
          <div className="bg-slate-950/70 p-2.5 rounded-xl border border-slate-800 space-y-0.5">
            <span className="text-slate-400 text-[10px]">Driver 6</span>
            <div className="font-bold text-indigo-300 truncate">Fixed Assets</div>
            <div className="text-[10px] text-slate-500 font-sans">Capex reinvestment</div>
          </div>
        </div>
      </div>
    </div>
  );
};
