import React from 'react';
import { Download, FileSpreadsheet, FileText, Settings, Database, Sliders, Scale, Layers, ChevronDown, Sparkles } from 'lucide-react';
import { CompanyProfile } from '../types/financial';

interface HeaderNavProps {
  activeTab: string;
  setActiveTab: (tab: string) => void;
  companyProfile: CompanyProfile;
  onOpenProfileModal: () => void;
  onOpenDeveloperModal: () => void;
  onExportPackageExcel: () => void;
  onExportPackageWord: () => void;
  onExportPackagePdf: () => void;
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
}) => {
  const [showExportMenu, setShowExportMenu] = React.useState(false);

  const navItems = [
    { id: 'statements', label: 'Statements & Reports', icon: Layers, accent: 'text-emerald-400' },
    { id: 'data-entry', label: 'Financial Data Entry', icon: Database, accent: 'text-sky-400' },
    { id: 'configuration', label: 'Cash Flow Config', icon: Settings, accent: 'text-indigo-400' },
    { id: 'projections', label: 'Projections Engine', icon: Sliders, accent: 'text-amber-400' },
    { id: 'reconciliation', label: 'Audit Matrix', icon: Scale, accent: 'text-teal-400' },
  ];

  return (
    <header className="sticky top-0 z-40 bg-slate-950/85 backdrop-blur-xl border-b border-slate-800/80 shadow-lg shadow-black/20 text-slate-100 transition-all">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between gap-4 h-16">
          
          {/* Brand Identity / Lead Masthead */}
          <div className="flex items-center gap-3 shrink-0">
            <button 
              onClick={() => setActiveTab('statements')}
              className="text-left group flex items-center gap-3 focus:outline-none"
            >
              <div className="relative">
                <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-700 flex items-center justify-center text-white font-extrabold text-base shadow-md shadow-emerald-950/60 ring-1 ring-emerald-400/30 group-hover:ring-emerald-400/60 transition-all">
                  T
                </div>
                <div className="absolute -bottom-0.5 -right-0.5 w-2.5 h-2.5 rounded-full bg-emerald-400 ring-2 ring-slate-950" />
              </div>
              <div>
                <div className="flex items-center gap-1.5">
                  <span className="text-base font-bold tracking-tight text-white whitespace-nowrap group-hover:text-emerald-300 transition-colors">
                    Financials with Tuhin
                  </span>
                  <span className="text-[10px] uppercase font-mono font-bold tracking-wider text-emerald-400/90 bg-emerald-950/80 border border-emerald-800/60 px-1.5 py-0.2 rounded">
                    Sec 76
                  </span>
                </div>
                <span className="text-[11px] text-slate-400 tracking-normal block -mt-0.5 truncate max-w-[210px] font-sans">
                  {companyProfile.companyName} · {companyProfile.reportingCurrency}
                </span>
              </div>
            </button>
          </div>

          {/* Navigation Links: Segmented Terminal Navigation */}
          <nav className="hidden lg:flex items-center gap-1 bg-slate-900/90 p-1 rounded-xl border border-slate-800/80 shadow-inner">
            {navItems.map((item) => {
              const Icon = item.icon;
              const isActive = activeTab === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => setActiveTab(item.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition-all flex items-center gap-2 ${
                    isActive
                      ? 'bg-slate-800 text-white shadow-sm ring-1 ring-slate-700/80 font-bold'
                      : 'text-slate-400 hover:text-slate-200 hover:bg-slate-850/60'
                  }`}
                >
                  <Icon className={`w-3.5 h-3.5 ${isActive ? item.accent : 'text-slate-500'}`} />
                  <span>{item.label}</span>
                  {isActive && (
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                  )}
                </button>
              );
            })}
          </nav>

          {/* Mobile / Compact Navigation (Fallback for medium screens) */}
          <div className="flex lg:hidden items-center gap-1 text-xs">
            {navItems.slice(0, 3).map((item) => {
              const Icon = item.icon;
              const isActive = activeTab === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => setActiveTab(item.id)}
                  className={`p-2 rounded-lg transition-colors ${
                    isActive ? 'bg-slate-800 text-emerald-400' : 'text-slate-400 hover:text-white'
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
              className="px-3 py-1.5 text-xs font-medium text-slate-300 bg-slate-900 hover:bg-slate-850 rounded-lg border border-slate-800/90 hover:border-slate-700 transition-all whitespace-nowrap hidden sm:inline-flex items-center gap-1.5 shadow-sm"
            >
              <Settings className="w-3.5 h-3.5 text-slate-400" />
              <span>Entity Setup</span>
            </button>

            {/* Combined Package Export Dropdown */}
            <div className="relative">
              <button
                onClick={() => setShowExportMenu(!showExportMenu)}
                className="px-3.5 py-1.5 text-xs font-semibold text-white bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 rounded-lg shadow-md shadow-emerald-950/50 transition-all whitespace-nowrap shrink-0 flex items-center gap-2 ring-1 ring-emerald-400/30 hover:ring-emerald-400/50"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Export Package</span>
                <ChevronDown className="w-3 h-3 opacity-80" />
              </button>

              {showExportMenu && (
                <div 
                  className="absolute right-0 mt-2 w-72 bg-slate-900/95 backdrop-blur-xl border border-slate-700/80 rounded-xl shadow-2xl p-2 z-50 text-xs animate-in fade-in slide-in-from-top-2 duration-150 ring-1 ring-black/40"
                  onMouseLeave={() => setShowExportMenu(false)}
                >
                  <div className="px-3 py-2 border-b border-slate-800/80 mb-1">
                    <div className="flex items-center justify-between">
                      <p className="font-bold text-white tracking-tight">Complete Financial Package</p>
                      <span className="text-[10px] font-mono text-emerald-400 bg-emerald-950 px-1.5 py-0.5 rounded border border-emerald-800/50">
                        Section 76
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-400 mt-0.5">All 5 Statements + Fixed Assets Schedule + Notes</p>
                  </div>

                  <button
                    onClick={() => {
                      setShowExportMenu(false);
                      onExportPackageExcel();
                    }}
                    className="w-full text-left px-3 py-2.5 hover:bg-slate-800/90 rounded-lg text-slate-200 flex items-center gap-3 transition-colors group"
                  >
                    <div className="w-8 h-8 rounded-lg bg-emerald-950/80 border border-emerald-800/60 flex items-center justify-center shrink-0 group-hover:border-emerald-600 transition-colors">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                    </div>
                    <div className="flex-1">
                      <div className="font-semibold text-white flex items-center justify-between">
                        <span>Excel Workbook</span>
                        <span className="text-[10px] font-mono text-emerald-400 bg-emerald-950/60 px-1 rounded">.XLSX</span>
                      </div>
                      <div className="text-[11px] text-slate-400">6 coordinated formatted worksheets</div>
                    </div>
                  </button>

                  <button
                    onClick={() => {
                      setShowExportMenu(false);
                      onExportPackageWord();
                    }}
                    className="w-full text-left px-3 py-2.5 hover:bg-slate-800/90 rounded-lg text-slate-200 flex items-center gap-3 transition-colors group"
                  >
                    <div className="w-8 h-8 rounded-lg bg-sky-950/80 border border-sky-800/60 flex items-center justify-center shrink-0 group-hover:border-sky-600 transition-colors">
                      <FileText className="w-4 h-4 text-sky-400" />
                    </div>
                    <div className="flex-1">
                      <div className="font-semibold text-white flex items-center justify-between">
                        <span>Word Document</span>
                        <span className="text-[10px] font-mono text-sky-400 bg-sky-950/60 px-1 rounded">.DOCX</span>
                      </div>
                      <div className="text-[11px] text-slate-400">Formatted executive statutory report</div>
                    </div>
                  </button>

                  <button
                    onClick={() => {
                      setShowExportMenu(false);
                      onExportPackagePdf();
                    }}
                    className="w-full text-left px-3 py-2.5 hover:bg-slate-800/90 rounded-lg text-slate-200 flex items-center gap-3 transition-colors group"
                  >
                    <div className="w-8 h-8 rounded-lg bg-rose-950/80 border border-rose-800/60 flex items-center justify-center shrink-0 group-hover:border-rose-600 transition-colors">
                      <Download className="w-4 h-4 text-rose-400" />
                    </div>
                    <div className="flex-1">
                      <div className="font-semibold text-white flex items-center justify-between">
                        <span>PDF Document</span>
                        <span className="text-[10px] font-mono text-rose-400 bg-rose-950/60 px-1 rounded">.PDF</span>
                      </div>
                      <div className="text-[11px] text-slate-400">Auditor-ready publication format with cover</div>
                    </div>
                  </button>
                </div>
              )}
            </div>

            {/* Developer Avatar & Credential Badge */}
            <button
              onClick={onOpenDeveloperModal}
              title="Software Author & Architect: Mohammad Towhidul Akbar Tuhin"
              className="flex items-center gap-2 pl-1.5 group cursor-pointer text-left"
            >
              <div className="relative">
                <img
                  src="/assets/aistudio/Tuhin_Pic_6c5894de.jpeg"
                  alt="Mohammad Towhidul Akbar Tuhin"
                  className="w-8 h-8 rounded-lg object-cover ring-1 ring-emerald-500/50 shadow-sm group-hover:ring-emerald-400 transition-all"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = '/public/assets/aistudio/Tuhin_Pic_6c5894de.jpeg';
                  }}
                />
                <span className="absolute -bottom-0.5 -right-0.5 w-2 h-2 rounded-full bg-emerald-400 ring-2 ring-slate-950" />
              </div>
              <div className="hidden xl:block text-left text-xs leading-tight">
                <span className="font-semibold text-slate-200 group-hover:text-emerald-300 transition-colors block truncate max-w-[160px]">
                  Mohammad T. A. Tuhin
                </span>
                <span className="text-[10px] text-emerald-400/90 font-mono block truncate max-w-[160px]">
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

