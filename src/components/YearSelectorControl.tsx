import React, { useState } from 'react';
import { Calendar, Plus, X, Check, History, TrendingUp } from 'lucide-react';

interface YearSelectorControlProps {
  selectedHistoricalYears: number[];
  selectedProjectedYears: number[];
  onToggleHistoricalYear: (year: number) => void;
  onToggleProjectedYear: (year: number) => void;
  onSetHistoricalYears: (years: number[]) => void;
  onSetProjectedYears: (years: number[]) => void;
  compact?: boolean;
}

export const YearSelectorControl: React.FC<YearSelectorControlProps> = ({
  selectedHistoricalYears,
  selectedProjectedYears,
  onToggleHistoricalYear,
  onToggleProjectedYear,
  onSetHistoricalYears,
  onSetProjectedYears,
  compact = false,
}) => {
  const [customHistYear, setCustomHistYear] = useState<string>('');
  const [customProjYear, setCustomProjYear] = useState<string>('');

  const sortedHist = [...selectedHistoricalYears].sort((a, b) => a - b);
  const sortedProj = [...selectedProjectedYears].sort((a, b) => a - b);

  // Quick candidate historical years (2018-2026 union selected)
  const histCandidates = Array.from(
    new Set([2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, ...sortedHist])
  )
    .filter(y => !sortedProj.includes(y))
    .sort((a, b) => a - b);

  const maxHist = sortedHist[sortedHist.length - 1] || 2025;

  // Quick candidate projected years (maxHist+1 through maxHist+15 union selected)
  const projCandidates = Array.from(
    new Set([
      ...Array.from({ length: 15 }, (_, i) => maxHist + 1 + i),
      ...sortedProj,
    ])
  )
    .filter(y => !sortedHist.includes(y))
    .sort((a, b) => a - b);

  const handleAddCustomHistorical = (e: React.FormEvent) => {
    e.preventDefault();
    const yr = Number(customHistYear.trim());
    if (!Number.isInteger(yr) || yr < 1990 || yr > 2100) return;
    if (!sortedHist.includes(yr) && !sortedProj.includes(yr)) {
      onToggleHistoricalYear(yr);
    }
    setCustomHistYear('');
  };

  const handleAddCustomProjected = (e: React.FormEvent) => {
    e.preventDefault();
    const yr = Number(customProjYear.trim());
    if (!Number.isInteger(yr) || yr < 1990 || yr > 2100) return;
    if (!sortedProj.includes(yr) && !sortedHist.includes(yr)) {
      onToggleProjectedYear(yr);
    }
    setCustomProjYear('');
  };

  return (
    <div className={`bg-slate-900/90 border border-slate-800 rounded-2xl ${compact ? 'p-3.5 space-y-3' : 'p-5 space-y-4'} shadow-sm`}>
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800/80 pb-2.5">
        <div className="flex items-center gap-2">
          <Calendar className="w-4 h-4 text-emerald-400" />
          <h3 className="text-xs sm:text-sm font-bold text-white tracking-tight">
            Flexible Year-by-Year Selection (Previous & Projected)
          </h3>
        </div>
        <div className="flex items-center gap-2 text-[11px] font-mono">
          <span className="px-2 py-0.5 rounded bg-slate-800 text-slate-200 border border-slate-700">
            Previous: <strong className="text-emerald-400">{sortedHist.length}Y</strong>
          </span>
          <span className="px-2 py-0.5 rounded bg-sky-950/60 text-sky-200 border border-sky-800/60">
            Projected: <strong className="text-sky-400">{sortedProj.length}Y</strong>
          </span>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        {/* Previous (Historical) Years Individual Selector */}
        <div className="bg-slate-950/70 border border-slate-800/90 rounded-xl p-3.5 space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-1.5 text-xs font-bold text-emerald-300">
              <History className="w-3.5 h-3.5" />
              <span>Previous Financial Years (Select Individually)</span>
            </div>
            <div className="flex items-center gap-1 text-[10px]">
              <button
                type="button"
                onClick={() => onSetHistoricalYears([2024, 2025].filter(y => !sortedProj.includes(y)))}
                className="px-1.5 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
              >
                2Y
              </button>
              <button
                type="button"
                onClick={() => onSetHistoricalYears([2023, 2024, 2025].filter(y => !sortedProj.includes(y)))}
                className="px-1.5 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
              >
                3Y
              </button>
              <button
                type="button"
                onClick={() => onSetHistoricalYears([2021, 2022, 2023, 2024, 2025].filter(y => !sortedProj.includes(y)))}
                className="px-1.5 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
              >
                5Y
              </button>
            </div>
          </div>

          <div className="flex flex-wrap gap-1.5">
            {histCandidates.map(yr => {
              const isSelected = sortedHist.includes(yr);
              const isLastOne = isSelected && sortedHist.length === 1;
              return (
                <button
                  key={yr}
                  type="button"
                  disabled={isLastOne}
                  onClick={() => onToggleHistoricalYear(yr)}
                  title={isLastOne ? 'At least 1 previous year must remain selected' : `Click to ${isSelected ? 'deselect' : 'select'} ${yr}`}
                  className={`px-2 py-1 rounded-lg text-[11px] font-mono font-semibold transition-all flex items-center gap-1 border ${
                    isSelected
                      ? 'bg-emerald-600/90 text-white border-emerald-400/60 shadow-sm'
                      : 'bg-slate-900 text-slate-400 border-slate-800 hover:border-slate-600 hover:text-slate-200'
                  } ${isLastOne ? 'opacity-90 cursor-not-allowed' : 'cursor-pointer'}`}
                >
                  {isSelected && <Check className="w-2.5 h-2.5" />}
                  <span>{yr}</span>
                  {isSelected && !isLastOne && (
                    <X className="w-2.5 h-2.5 opacity-70 hover:opacity-100" />
                  )}
                </button>
              );
            })}
          </div>

          <form onSubmit={handleAddCustomHistorical} className="flex items-center gap-1.5 pt-1">
            <input
              type="number"
              min={1990}
              max={2100}
              placeholder="Add year (e.g. 2017)"
              value={customHistYear}
              onChange={e => setCustomHistYear(e.target.value)}
              className="w-36 bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1 text-[11px] font-mono text-white placeholder:text-slate-500 focus:outline-none focus:border-emerald-500"
            />
            <button
              type="submit"
              className="px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-emerald-600 text-emerald-300 hover:text-white border border-slate-700 text-[11px] font-semibold transition-colors flex items-center gap-1"
            >
              <Plus className="w-3 h-3" />
              Add Previous
            </button>
          </form>
        </div>

        {/* Projected Years Individual Selector */}
        <div className="bg-slate-950/70 border border-slate-800/90 rounded-xl p-3.5 space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-1.5 text-xs font-bold text-sky-300">
              <TrendingUp className="w-3.5 h-3.5" />
              <span>Projected Financial Years (Select Individually)</span>
            </div>
            <div className="flex items-center gap-1 text-[10px]">
              {[3, 5, 10, 15].map(cnt => (
                <button
                  key={cnt}
                  type="button"
                  onClick={() => onSetProjectedYears(Array.from({ length: cnt }, (_, i) => maxHist + 1 + i))}
                  className="px-1.5 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
                >
                  {cnt}Y
                </button>
              ))}
            </div>
          </div>

          <div className="flex flex-wrap gap-1.5">
            {projCandidates.map(yr => {
              const isSelected = sortedProj.includes(yr);
              const isLastOne = isSelected && sortedProj.length === 1;
              return (
                <button
                  key={yr}
                  type="button"
                  disabled={isLastOne}
                  onClick={() => onToggleProjectedYear(yr)}
                  title={isLastOne ? 'At least 1 projected year must remain selected' : `Click to ${isSelected ? 'deselect' : 'select'} ${yr}`}
                  className={`px-2 py-1 rounded-lg text-[11px] font-mono font-semibold transition-all flex items-center gap-1 border ${
                    isSelected
                      ? 'bg-sky-600/90 text-white border-sky-400/60 shadow-sm'
                      : 'bg-slate-900 text-slate-400 border-slate-800 hover:border-slate-600 hover:text-slate-200'
                  } ${isLastOne ? 'opacity-90 cursor-not-allowed' : 'cursor-pointer'}`}
                >
                  {isSelected && <Check className="w-2.5 h-2.5" />}
                  <span>{yr}</span>
                  {isSelected && !isLastOne && (
                    <X className="w-2.5 h-2.5 opacity-70 hover:opacity-100" />
                  )}
                </button>
              );
            })}
          </div>

          <form onSubmit={handleAddCustomProjected} className="flex items-center gap-1.5 pt-1">
            <input
              type="number"
              min={1990}
              max={2100}
              placeholder="Add year (e.g. 2042)"
              value={customProjYear}
              onChange={e => setCustomProjYear(e.target.value)}
              className="w-36 bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1 text-[11px] font-mono text-white placeholder:text-slate-500 focus:outline-none focus:border-sky-500"
            />
            <button
              type="submit"
              className="px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-sky-600 text-sky-300 hover:text-white border border-slate-700 text-[11px] font-semibold transition-colors flex items-center gap-1"
            >
              <Plus className="w-3 h-3" />
              Add Projected
            </button>
          </form>
        </div>
      </div>
    </div>
  );
};
