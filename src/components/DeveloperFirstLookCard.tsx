import React, { useState } from 'react';
import { Mail, GraduationCap, Building, Calendar, ChevronRight, ChevronDown } from 'lucide-react';
import { ProjectionAssumptions, CompanyProfile } from '../types/financial';
import { YearSelectorControl } from './YearSelectorControl';

interface DeveloperFirstLookCardProps {
  assumptions: ProjectionAssumptions;
  onUpdateAssumptions: (updated: ProjectionAssumptions) => void;
  onOpenDeveloperModal: () => void;
  companyProfile: CompanyProfile;
  baseYear: number;
  selectedHistoricalYears?: number[];
  selectedProjectedYears?: number[];
  onToggleHistoricalYear?: (year: number) => void;
  onToggleProjectedYear?: (year: number) => void;
  onSetHistoricalYears?: (years: number[]) => void;
  onSetProjectedYears?: (years: number[]) => void;
}

export const DeveloperFirstLookCard: React.FC<DeveloperFirstLookCardProps> = ({
  assumptions,
  onUpdateAssumptions,
  onOpenDeveloperModal,
  baseYear,
  selectedHistoricalYears = [2023, 2024, 2025],
  selectedProjectedYears,
  onToggleHistoricalYear,
  onToggleProjectedYear,
  onSetHistoricalYears,
  onSetProjectedYears,
}) => {
  const [showYearCustomizer, setShowYearCustomizer] = useState(false);
  const quickHorizons = [1, 3, 5, 7, 10, 12, 15, 20];

  const handleSetHorizon = (years: number) => {
    onUpdateAssumptions({
      ...assumptions,
      projectedYearsCount: years,
    });
  };

  const projYears = selectedProjectedYears && selectedProjectedYears.length > 0
    ? [...selectedProjectedYears].sort((a, b) => a - b)
    : Array.from({ length: assumptions.projectedYearsCount }, (_, i) => baseYear + 1 + i);
  const startProjectedYear = projYears[0] ?? baseYear + 1;
  const endProjectedYear = projYears[projYears.length - 1] ?? baseYear + assumptions.projectedYearsCount;
  const sortedHist = [...selectedHistoricalYears].sort((a, b) => a - b);

  return (
    <div className="relative overflow-hidden rounded-2xl bg-white/88 backdrop-blur-md border border-slate-200/90 shadow-sm">
      <div className="relative p-5 sm:p-6 lg:p-7 space-y-4">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-6">
          
          {/* Developer Identity Block (First Look Feature) */}
          <div className="flex flex-col sm:flex-row items-center sm:items-start gap-5">
            {/* Developer Photograph with Executive Framing */}
            <div className="relative shrink-0 group cursor-pointer" onClick={onOpenDeveloperModal}>
              <div className="relative p-1 rounded-2xl bg-gradient-to-tr from-emerald-500/40 via-slate-200 to-teal-400/50 shadow-sm ring-1 ring-emerald-500/20 group-hover:ring-emerald-500/50 transition-all">
                <img
                  src="/assets/aistudio/Tuhin_Pic_6c5894de.jpeg"
                  alt="Mohammad Towhidul Akbar Tuhin"
                  className="w-24 h-24 sm:w-28 sm:h-28 rounded-xl object-cover"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = '/public/assets/aistudio/Tuhin_Pic_6c5894de.jpeg';
                  }}
                />
              </div>
              <span className="mt-1.5 block text-center text-[10px] font-bold tracking-wider text-emerald-700 uppercase font-mono">
                Architect
              </span>
            </div>

            {/* Developer Bio & Qualifications */}
            <div className="text-center sm:text-left space-y-2">
              <div className="flex flex-wrap items-center justify-center sm:justify-start gap-2.5">
                <h1 className="text-lg sm:text-xl font-extrabold tracking-tight text-slate-900 font-sans">
                  Mohammad Towhidul Akbar Tuhin
                </h1>
                <span className="text-[11px] font-mono font-semibold text-emerald-700">
                  · Lead Software Architect
                </span>
              </div>

              {/* Designation & Organization */}
              <div className="space-y-0.5">
                <p className="text-xs font-semibold text-slate-800 flex items-center justify-center sm:justify-start gap-1.5">
                  <Building className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                  <span>Officer, Sammilito Islami Bank PLC</span>
                </p>
                <p className="text-[11px] text-slate-500">
                  (Transferee of Assets & Liabilities of Union Bank PLC)
                </p>
              </div>

              {/* Academic Degrees */}
              <div className="flex flex-wrap items-center justify-center sm:justify-start gap-x-3 gap-y-1 text-xs text-slate-700 pt-0.5">
                <span className="inline-flex items-center gap-1.5 font-medium">
                  <GraduationCap className="w-3.5 h-3.5 text-sky-600 shrink-0" />
                  BBA, Heriot-Watt University, UK
                </span>
                <span className="text-slate-300 hidden sm:inline" aria-hidden="true">·</span>
                <span className="inline-flex items-center gap-1.5 font-medium">
                  <GraduationCap className="w-3.5 h-3.5 text-sky-600 shrink-0" />
                  MBA, University of Chittagong
                </span>
              </div>

              {/* Contact Link */}
              <div className="pt-1 flex items-center justify-center sm:justify-start gap-3">
                <a
                  href="mailto:tuhinakbar1993@gmail.com"
                  className="inline-flex items-center gap-1.5 text-[11px] font-semibold text-emerald-700 hover:text-emerald-600 transition-colors"
                >
                  <Mail className="w-3.5 h-3.5" />
                  tuhinakbar1993@gmail.com
                </a>
                <span className="text-slate-300" aria-hidden="true">·</span>
                <button
                  onClick={onOpenDeveloperModal}
                  className="text-[11px] font-medium text-slate-600 hover:text-slate-900 underline underline-offset-2 transition-colors flex items-center gap-1"
                >
                  <span>View Full Profile & Credentials</span>
                  <ChevronRight className="w-3 h-3 text-slate-400" />
                </button>
              </div>
            </div>
          </div>

          {/* Interactive Projection Horizon & Individual Year Selector Control */}
          <div className="w-full lg:w-auto bg-slate-50/90 border border-slate-200/90 rounded-xl p-4 space-y-3 shrink-0">
            <div className="flex items-center justify-between gap-4">
              <div className="flex items-center gap-2">
                <Calendar className="w-4 h-4 text-emerald-600" />
                <span className="text-xs font-bold text-slate-900 tracking-tight uppercase font-mono">
                  Reporting Horizon
                </span>
              </div>
              <span className="text-xs font-mono font-bold text-emerald-800 bg-emerald-50 px-2.5 py-0.5 rounded border border-emerald-200">
                {sortedHist.length}H + {projYears.length}P ({startProjectedYear}–{endProjectedYear})
              </span>
            </div>

            {/* Quick Horizon Buttons */}
            <div className="flex flex-wrap items-center gap-1.5">
              {quickHorizons.map(years => (
                <button
                  key={years}
                  onClick={() => handleSetHorizon(years)}
                  className={`px-2.5 py-1 text-xs font-mono font-bold rounded-lg transition-all ${
                    projYears.length === years
                      ? 'bg-emerald-600 text-white shadow-sm'
                      : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 hover:text-slate-900'
                  }`}
                >
                  {years}Y
                </button>
              ))}
            </div>

            {/* Custom Slider / Fine Tuning */}
            <div className="space-y-1.5 pt-1">
              <div className="flex items-center justify-between text-[11px] text-slate-600">
                <span className="font-medium">Previous: <strong className="text-emerald-700 font-mono">{sortedHist.join(', ')}</strong></span>
                <span className="font-mono text-slate-900 font-semibold">{projYears.length} Projected Years</span>
              </div>
              <input
                type="range"
                min={1}
                max={20}
                value={Math.min(20, Math.max(1, projYears.length))}
                onChange={(e) => handleSetHorizon(Number(e.target.value))}
                className="w-full h-1.5 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-emerald-600 hover:accent-emerald-500 transition-all"
              />
              <div className="flex items-center justify-between text-[10px] text-slate-500 font-mono">
                <span>1Y · 5Y · 10Y · 15Y · 20Y</span>
                {onToggleHistoricalYear && onToggleProjectedYear && onSetHistoricalYears && onSetProjectedYears && (
                  <button
                    type="button"
                    onClick={() => setShowYearCustomizer(!showYearCustomizer)}
                    className="text-emerald-700 hover:text-emerald-600 font-sans font-semibold inline-flex items-center gap-1 transition-colors"
                  >
                    <span>{showYearCustomizer ? 'Hide Individual Years' : 'Select Individual Years'}</span>
                    <ChevronDown className={`w-3 h-3 transition-transform ${showYearCustomizer ? 'rotate-180' : ''}`} />
                  </button>
                )}
              </div>
            </div>
          </div>

        </div>

        {/* Expandable Individual Year Selector */}
        {showYearCustomizer && onToggleHistoricalYear && onToggleProjectedYear && onSetHistoricalYears && onSetProjectedYears && (
          <YearSelectorControl
            selectedHistoricalYears={sortedHist}
            selectedProjectedYears={projYears}
            onToggleHistoricalYear={onToggleHistoricalYear}
            onToggleProjectedYear={onToggleProjectedYear}
            onSetHistoricalYears={onSetHistoricalYears}
            onSetProjectedYears={onSetProjectedYears}
            compact
          />
        )}
      </div>
    </div>
  );
};
