import React from 'react';
import { Download, FileSpreadsheet, FileText, Settings, Database, Sliders, Scale, Layers, ChevronDown } from 'lucide-react';
import { CompanyProfile, PreviousComparativePair } from '../types/financial';

interface HeaderNavProps {
  activeTab: string;
  setActiveTab: (tab: string) => void;
  companyProfile: CompanyProfile;
  onOpenProfileModal: () => void;
  onOpenDeveloperModal: () => void;
  onExportPackageExcel: () => void;
  onExportPackageWord: () => void;
  onExportPackagePdf: () => void;
  previousComparativePairs?: PreviousComparativePair[];
  onExportPreviousPair?: (pairId: string, format: 'xlsx' | 'docx' | 'pdf') => void;
  onExportPreviousExcel?: () => void;
  onExportPreviousWord?: () => void;
  onExportPreviousPdf?: () => void;
  projectedYearsRangeLabel?: string;
  onExportProjectedExcel?: () => void;
  onExportProjectedWord?: () => void;
  onExportProjectedPdf?: () => void;
}

export const HeaderNav: React.FC<HeaderNavProps> = ({
  activeTab,
  setActiveTab,
  companyProfile,
  onOpenProfileModal,
  onOpenDeveloperModal,
  onExportPackageExcel,
  onExportPackageWord,
  onExportPackagePdf,
  previousComparativePairs = [],
  onExportPreviousPair,
  onExportPreviousExcel,
  onExportPreviousWord,
  onExportPreviousPdf,
  projectedYearsRangeLabel,
  onExportProjectedExcel,
  onExportProjectedWord,
  onExportProjectedPdf,
}) => {
  const [showExportMenu, setShowExportMenu] = React.useState(false);

  const navItems = [
    { id: 'statements', label: 'Statements & Reports', icon: Layers, accent: 'text-emerald-700' },
    { id: 'data-entry', label: 'Financial Data Entry', icon: Database, accent: 'text-sky-700' },
    { id: 'configuration', label: 'Cash Flow Config', icon: Settings, accent: 'text-indigo-700' },
    { id: 'projections', label: 'Projections Engine', icon: Sliders, accent: 'text-amber-700' },
    { id: 'reconciliation', label: 'Audit Matrix', icon: Scale, accent: 'text-teal-700' },
  ];

  return (
    <header className="sticky top-0 z-40 bg-white/82 backdrop-blur-xl border-b border-slate-200/90 shadow-sm text-slate-900 transition-all">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between gap-4 h-16">
          
          {/* Brand Identity / Lead Masthead */}
          <div className="flex items-center gap-3 shrink-0">
            <button 
              onClick={() => setActiveTab('statements')}
              className="text-left group flex items-center gap-3 focus:outline-none"
            >
              <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 flex items-center justify-center text-white font-extrabold text-base shadow-sm ring-1 ring-emerald-500/30 transition-all">
                T
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-base font-bold tracking-tight text-slate-900 whitespace-nowrap group-hover:text-emerald-700 transition-colors">
                    Financials with Tuhin
                  </span>
                  <span className="text-[10px] font-mono font-semibold text-emerald-700">
                    Sec 76
                  </span>
                </div>
                <span className="text-[11px] text-slate-500 tracking-normal block -mt-0.5 truncate max-w-[210px] font-sans">
                  {companyProfile.companyName} · {companyProfile.reportingCurrency}
                </span>
              </div>
            </button>
          </div>

          {/* Navigation Links: Clean Frosted Studio Navigation */}
          <nav className="hidden lg:flex items-center gap-1 bg-slate-100/90 p-1 rounded-xl border border-slate-200/90">
            {navItems.map((item) => {
              const Icon = item.icon;
              const isActive = activeTab === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => setActiveTab(item.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition-all flex items-center gap-2 ${
                    isActive
                      ? 'bg-white text-slate-900 shadow-sm ring-1 ring-slate-200 font-bold'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
                  }`}
                >
                  <Icon className={`w-3.5 h-3.5 ${isActive ? item.accent : 'text-slate-400'}`} />
                  <span>{item.label}</span>
                </button>
              );
            })}
          </nav>

          {/* Mobile / Compact Navigation */}
          <div className="flex lg:hidden items-center gap-1 text-xs">
            {navItems.map((item) => {
              const Icon = item.icon;
              const isActive = activeTab === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => setActiveTab(item.id)}
                  className={`p-2 rounded-lg transition-colors ${
                    isActive ? 'bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200' : 'text-slate-500 hover:text-slate-900 hover:bg-slate-100'
                  }`}
                  title={item.label}
                >
                  <Icon className="w-4 h-4" />
                </button>
              );
            })}
          </div>

          {/* Actions & Developer Identity */}
          <div className="flex items-center gap-2.5 shrink-0">
            {/* Entity Setup Button */}
            <button
              onClick={onOpenProfileModal}
              title="Configure Company Entity & Accounting Framework"
              className="px-3 py-1.5 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 rounded-lg border border-slate-200 hover:border-slate-300 transition-all whitespace-nowrap hidden sm:inline-flex items-center gap-1.5 shadow-sm"
            >
              <Settings className="w-3.5 h-3.5 text-slate-500" />
              <span>Entity Setup</span>
            </button>

            {/* Combined Package Export Dropdown */}
            <div className="relative">
              <button
                onClick={() => setShowExportMenu(!showExportMenu)}
                className="px-3.5 py-1.5 text-xs font-semibold text-white bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 rounded-lg shadow-sm transition-all whitespace-nowrap shrink-0 flex items-center gap-2"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export Package</span>
                <ChevronDown className="w-3 h-3 opacity-90" />
              </button>

              {showExportMenu && (
                <div 
                  className="absolute right-0 mt-2 w-96 max-h-[82vh] overflow-y-auto bg-white/98 backdrop-blur-xl border border-slate-200 rounded-xl shadow-2xl p-2.5 z-50 text-xs text-slate-800 space-y-2.5"
                  onMouseLeave={() => setShowExportMenu(false)}
                >
                  {/* 1. Separate Previous Financials — Comparative Pairwise Downloads */}
                  {(onExportPreviousExcel || previousComparativePairs.length > 0) && (
                    <div className="bg-emerald-50/75 border border-emerald-200/90 rounded-lg p-2.5 space-y-2">
                      <div className="flex items-center justify-between gap-2">
                        <div>
                          <span className="text-[11px] font-bold text-emerald-900 block">
                            1. Previous Financials (Comparative Sets)
                          </span>
                          <span className="text-[10px] text-emerald-700">
                            Exact 2-year comparative downloads (31.12.YYYY)
                          </span>
                        </div>
                      </div>

                      {previousComparativePairs.length > 0 && onExportPreviousPair && (
                        <div className="space-y-1.5">
                          {previousComparativePairs.map((pair) => (
                            <div
                              key={pair.id}
                              className="flex items-center justify-between gap-2 bg-white/95 border border-emerald-200/80 rounded-md px-2 py-1.5"
                            >
                              <span className="font-semibold text-slate-800 font-mono text-[11px] whitespace-nowrap">
                                ({pair.label})
                              </span>
                              <div className="grid grid-cols-3 gap-1 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => { setShowExportMenu(false); onExportPreviousPair(pair.id, 'xlsx'); }}
                                  className="px-2 py-1 rounded bg-emerald-50 hover:bg-emerald-600 text-emerald-700 hover:text-white border border-emerald-200 font-mono text-[10px] font-bold transition-colors"
                                >
                                  .XLSX
                                </button>
                                <button
                                  type="button"
                                  onClick={() => { setShowExportMenu(false); onExportPreviousPair(pair.id, 'docx'); }}
                                  className="px-2 py-1 rounded bg-sky-50 hover:bg-sky-600 text-sky-700 hover:text-white border border-sky-200 font-mono text-[10px] font-bold transition-colors"
                                >
                                  .DOCX
                                </button>
                                <button
                                  type="button"
                                  onClick={() => { setShowExportMenu(false); onExportPreviousPair(pair.id, 'pdf'); }}
                                  className="px-2 py-1 rounded bg-rose-50 hover:bg-rose-600 text-rose-700 hover:text-white border border-rose-200 font-mono text-[10px] font-bold transition-colors"
                                >
                                  .PDF
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}

                      {onExportPreviousExcel && (
                        <div className="pt-1 border-t border-emerald-200/70">
                          <div className="text-[10px] font-semibold text-emerald-800 mb-1">
                            Download All Comparative Pairs ({previousComparativePairs.length} {previousComparativePairs.length === 1 ? 'file' : 'separate files'}):
                          </div>
                          <div className="grid grid-cols-3 gap-1.5">
                            <button
                              type="button"
                              onClick={() => { setShowExportMenu(false); onExportPreviousExcel?.(); }}
                              className="py-1 rounded bg-emerald-600 hover:bg-emerald-700 text-white font-mono text-[10px] font-bold transition-colors"
                            >
                              All .XLSX
                            </button>
                            <button
                              type="button"
                              onClick={() => { setShowExportMenu(false); onExportPreviousWord?.(); }}
                              className="py-1 rounded bg-sky-600 hover:bg-sky-700 text-white font-mono text-[10px] font-bold transition-colors"
                            >
                              All .DOCX
                            </button>
                            <button
                              type="button"
                              onClick={() => { setShowExportMenu(false); onExportPreviousPdf?.(); }}
                              className="py-1 rounded bg-rose-600 hover:bg-rose-700 text-white font-mono text-[10px] font-bold transition-colors"
                            >
                              All .PDF
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* 2. Separate Projected Financials — All Selected Years Together */}
                  {onExportProjectedExcel && (
                    <div className="bg-sky-50/75 border border-sky-200/90 rounded-lg p-2.5 space-y-1.5">
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-[11px] font-bold text-sky-900">
                          2. Projected Financials {projectedYearsRangeLabel ? `(${projectedYearsRangeLabel})` : ''}
                        </span>
                        <span className="text-[10px] font-mono text-sky-700">Notes 1–13</span>
                      </div>
                      <p className="text-[10px] text-sky-700">
                        All selected projected years combined in exact format
                      </p>
                      <div className="grid grid-cols-3 gap-1.5">
                        <button
                          type="button"
                          onClick={() => { setShowExportMenu(false); onExportProjectedExcel?.(); }}
                          className="py-1 rounded bg-emerald-50 hover:bg-emerald-600 text-emerald-700 hover:text-white border border-emerald-200 font-mono text-[10px] font-bold transition-colors"
                        >
                          .XLSX
                        </button>
                        <button
                          type="button"
                          onClick={() => { setShowExportMenu(false); onExportProjectedWord?.(); }}
                          className="py-1 rounded bg-sky-50 hover:bg-sky-600 text-sky-700 hover:text-white border border-sky-200 font-mono text-[10px] font-bold transition-colors"
                        >
                          .DOCX
                        </button>
                        <button
                          type="button"
                          onClick={() => { setShowExportMenu(false); onExportProjectedPdf?.(); }}
                          className="py-1 rounded bg-rose-50 hover:bg-rose-600 text-rose-700 hover:text-white border border-rose-200 font-mono text-[10px] font-bold transition-colors"
                        >
                          .PDF
                        </button>
                      </div>
                    </div>
                  )}

                  {/* 3. Complete Section 76 Package (Preserved) */}
                  <div className="pt-1 border-t border-slate-200 space-y-1">
                    <div className="px-1 py-1">
                      <div className="flex items-center justify-between">
                        <p className="font-bold text-slate-900 tracking-tight">3. Complete Financial Package (Sec 76)</p>
                      </div>
                      <p className="text-[11px] text-slate-500">All 5 Statements + Fixed Assets Schedule + Notes</p>
                    </div>

                    <button
                      type="button"
                      onClick={() => {
                        setShowExportMenu(false);
                        onExportPackageExcel();
                      }}
                      className="w-full text-left px-2.5 py-2 hover:bg-slate-100 rounded-lg text-slate-800 flex items-center gap-3 transition-colors group"
                    >
                      <div className="w-8 h-8 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center shrink-0">
                        <FileSpreadsheet className="w-4 h-4 text-emerald-700" />
                      </div>
                      <div className="flex-1">
                        <div className="font-semibold text-slate-900 flex items-center justify-between">
                          <span>Excel Workbook</span>
                          <span className="text-[10px] font-mono text-emerald-700 font-bold">.XLSX</span>
                        </div>
                        <div className="text-[11px] text-slate-500">Coordinated multi-sheet workbook</div>
                      </div>
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setShowExportMenu(false);
                        onExportPackageWord();
                      }}
                      className="w-full text-left px-2.5 py-2 hover:bg-slate-100 rounded-lg text-slate-800 flex items-center gap-3 transition-colors group"
                    >
                      <div className="w-8 h-8 rounded-lg bg-sky-50 border border-sky-200 flex items-center justify-center shrink-0">
                        <FileText className="w-4 h-4 text-sky-700" />
                      </div>
                      <div className="flex-1">
                        <div className="font-semibold text-slate-900 flex items-center justify-between">
                          <span>Word Document</span>
                          <span className="text-[10px] font-mono text-sky-700 font-bold">.DOCX</span>
                        </div>
                        <div className="text-[11px] text-slate-500">Formatted executive statutory report</div>
                      </div>
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setShowExportMenu(false);
                        onExportPackagePdf();
                      }}
                      className="w-full text-left px-2.5 py-2 hover:bg-slate-100 rounded-lg text-slate-800 flex items-center gap-3 transition-colors group"
                    >
                      <div className="w-8 h-8 rounded-lg bg-rose-50 border border-rose-200 flex items-center justify-center shrink-0">
                        <Download className="w-4 h-4 text-rose-700" />
                      </div>
                      <div className="flex-1">
                        <div className="font-semibold text-slate-900 flex items-center justify-between">
                          <span>PDF Document</span>
                          <span className="text-[10px] font-mono text-rose-700 font-bold">.PDF</span>
                        </div>
                        <div className="text-[11px] text-slate-500">Auditor-ready publication format</div>
                      </div>
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Developer Avatar & Credential Badge */}
            <button
              onClick={onOpenDeveloperModal}
              title="Software Author & Architect: Mohammad Towhidul Akbar Tuhin"
              className="flex items-center gap-2 pl-1.5 group cursor-pointer text-left"
            >
              <img
                src="/assets/aistudio/Tuhin_Pic_6c5894de.jpeg"
                alt="Mohammad Towhidul Akbar Tuhin"
                className="w-8 h-8 rounded-lg object-cover ring-1 ring-emerald-600/40 shadow-sm group-hover:ring-emerald-600 transition-all"
                onError={(e) => {
                  (e.target as HTMLImageElement).src = '/public/assets/aistudio/Tuhin_Pic_6c5894de.jpeg';
                }}
              />
              <div className="hidden xl:block text-left text-xs leading-tight">
                <span className="font-semibold text-slate-900 group-hover:text-emerald-700 transition-colors block truncate max-w-[160px]">
                  Mohammad T. A. Tuhin
                </span>
                <span className="text-[10px] text-emerald-700 font-mono block truncate max-w-[160px]">
                  Officer, Sammilito Islami Bank
                </span>
              </div>
            </button>
          </div>
        </div>
      </div>
    </header>
  );
};

