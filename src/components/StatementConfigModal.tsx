import React, { useState, useMemo } from 'react';
import {
  Settings,
  Plus,
  Trash2,
  ArrowUp,
  ArrowDown,
  CheckSquare,
  Square,
  RefreshCw,
  Search,
  Sparkles,
  BookmarkCheck,
  Check,
} from 'lucide-react';
import { StatementConfigItem } from '../types/financial';
import { generateCustomElementUnevenSeries } from '../services/financialEngine';

interface StatementConfigViewProps {
  configItems: StatementConfigItem[];
  onUpdateConfig: (items: StatementConfigItem[]) => void;
  onResetDefaults: () => void;
  onSaveAsDefault?: () => void;
  onResetFactoryDefaults?: () => void;
  hasSavedDefault?: boolean;
  savedDefaultAt?: string | null;
  years?: number[];
}

type StatementFilter = 'all' | 'income' | 'balanceSheet' | 'cashFlow' | 'equity' | 'fixedAssets';

const STATEMENT_TABS: { id: StatementFilter; label: string }[] = [
  { id: 'all', label: 'All Statements' },
  { id: 'income', label: 'Income Statement' },
  { id: 'balanceSheet', label: 'Balance Sheet' },
  { id: 'cashFlow', label: 'Cash Flows' },
  { id: 'equity', label: 'Changes in Equity' },
  { id: 'fixedAssets', label: 'Fixed Assets' },
];

const STATEMENT_SECTIONS: Record<Exclude<StatementFilter, 'all'>, string[]> = {
  income: ['Revenue & Gross Margin', 'Operating Expenses', 'Non-Operating & Tax'],
  balanceSheet: ['Current Assets', 'Non-Current Assets', 'Current Liabilities', 'Non-Current Liabilities', 'Equity'],
  cashFlow: ['Operating Activities', 'Investing Activities', 'Financing Activities'],
  equity: ['Share Capital', 'Retained Earnings', 'Reserves & OCI'],
  fixedAssets: ['Asset Categories'],
};

const STATEMENT_BADGE: Record<StatementConfigItem['statement'], string> = {
  income: 'Income Statement',
  balanceSheet: 'Balance Sheet',
  cashFlow: 'Cash Flows',
  equity: 'Equity',
  fixedAssets: 'Fixed Assets',
};

export const StatementConfigView: React.FC<StatementConfigViewProps> = ({
  configItems,
  onUpdateConfig,
  onResetDefaults,
  onSaveAsDefault,
  onResetFactoryDefaults,
  hasSavedDefault = false,
  savedDefaultAt = null,
  years = [],
}) => {
  const [selectedStatement, setSelectedStatement] = useState<StatementFilter>('all');
  const [selectedSection, setSelectedSection] = useState<string>('All');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [newItemLabel, setNewItemLabel] = useState('');
  const [newItemStatement, setNewItemStatement] =
    useState<StatementConfigItem['statement']>('cashFlow');
  const [newItemSection, setNewItemSection] = useState('Operating Activities');
  const [newItemSign, setNewItemSign] = useState<'positive' | 'negative'>('positive');
  const [expandedCustomId, setExpandedCustomId] = useState<string | null>(null);
  const [savedNotice, setSavedNotice] = useState<boolean>(false);

  const handleSaveDefaultClick = () => {
    if (!onSaveAsDefault) return;
    onSaveAsDefault();
    setSavedNotice(true);
    setTimeout(() => setSavedNotice(false), 3200);
  };

  const availableSections = useMemo(() => {
    const scoped =
      selectedStatement === 'all'
        ? configItems
        : configItems.filter(i => i.statement === selectedStatement);
    const unique = Array.from(new Set(scoped.map(i => i.section)));
    return ['All', ...unique];
  }, [configItems, selectedStatement]);

  const filteredItems = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return configItems
      .filter(item => {
        if (selectedStatement !== 'all' && item.statement !== selectedStatement) return false;
        if (selectedSection !== 'All' && item.section !== selectedSection) return false;
        if (
          q &&
          !item.label.toLowerCase().includes(q) &&
          !item.section.toLowerCase().includes(q) &&
          !STATEMENT_BADGE[item.statement].toLowerCase().includes(q)
        ) {
          return false;
        }
        return true;
      })
      .sort((a, b) => {
        if (a.statement !== b.statement) return a.statement.localeCompare(b.statement);
        return a.order - b.order;
      });
  }, [configItems, selectedStatement, selectedSection, searchQuery]);

  const handleSelectStatementTab = (tab: StatementFilter) => {
    setSelectedStatement(tab);
    setSelectedSection('All');
    if (tab !== 'all') {
      setNewItemStatement(tab);
      setNewItemSection(STATEMENT_SECTIONS[tab][0]);
    }
  };

  const toggleVisibility = (id: string) => {
    onUpdateConfig(
      configItems.map(item =>
        item.id === id ? { ...item, isVisible: !item.isVisible } : item
      )
    );
  };

  const handleSelectScoped = (select: boolean, optionalOnly = false) => {
    const q = searchQuery.trim().toLowerCase();
    onUpdateConfig(
      configItems.map(item => {
        const matchesStatement =
          selectedStatement === 'all' || item.statement === selectedStatement;
        const matchesSection = selectedSection === 'All' || item.section === selectedSection;
        const matchesQuery =
          !q ||
          item.label.toLowerCase().includes(q) ||
          item.section.toLowerCase().includes(q) ||
          STATEMENT_BADGE[item.statement].toLowerCase().includes(q);
        if (!matchesStatement || !matchesSection || !matchesQuery) return item;
        if (optionalOnly && item.isMandatory) return item;
        return { ...item, isVisible: select };
      })
    );
  };

  const handleLabelChange = (id: string, newLabel: string) => {
    onUpdateConfig(
      configItems.map(item => (item.id === id ? { ...item, label: newLabel } : item))
    );
  };

  const handleCustomYearValueChange = (id: string, yr: number, value: number) => {
    onUpdateConfig(
      configItems.map(item =>
        item.id === id
          ? {
              ...item,
              customValues: {
                ...(item.customValues || {}),
                [yr]: Number.isFinite(value) ? value : 0,
              },
            }
          : item
      )
    );
  };

  const handleCustomSignChange = (id: string, sign: 'positive' | 'negative') => {
    onUpdateConfig(
      configItems.map(item => (item.id === id ? { ...item, sign } : item))
    );
  };

  const handleFillCustomUnevenRatios = (item: StatementConfigItem) => {
    if (years.length === 0) return;
    const sortedYears = [...years].sort((a, b) => a - b);
    const firstNonZeroYear = sortedYears.find(yr => Math.abs(item.customValues?.[yr] || 0) > 0);
    const baseYear = firstNonZeroYear ?? sortedYears[0];
    const baseVal = firstNonZeroYear ? Math.abs(item.customValues![firstNonZeroYear]) : 250000;
    const labelKey = item.label || item.id;
    const seed = labelKey.split('').reduce((acc, ch, i) => (acc + ch.charCodeAt(0) * (i + 1)) % 97, 17);
    const series = generateCustomElementUnevenSeries(baseVal, sortedYears, baseYear, 8.2, seed);
    onUpdateConfig(
      configItems.map(ci => (ci.id === item.id ? { ...ci, customValues: series } : ci))
    );
  };

  const handleMove = (id: string, direction: 'up' | 'down') => {
    const idx = configItems.findIndex(i => i.id === id);
    if (idx === -1) return;
    const targetIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (targetIdx < 0 || targetIdx >= configItems.length) return;

    const newItems = [...configItems];
    const temp = newItems[idx];
    newItems[idx] = newItems[targetIdx];
    newItems[targetIdx] = temp;

    const reindexed = newItems.map((item, i) => ({ ...item, order: i + 1 }));
    onUpdateConfig(reindexed);
  };

  const handleAddCustomItem = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newItemLabel.trim()) return;

    const newId = `custom_${Date.now()}`;
    const newItem: StatementConfigItem = {
      id: newId,
      statement: newItemStatement,
      section: newItemSection,
      label: newItemLabel.trim(),
      isMandatory: false,
      isVisible: true,
      order: configItems.length + 1,
      isCustom: true,
      sign: newItemSign,
      customValues: Object.fromEntries(years.map(y => [y, 0])),
    };

    onUpdateConfig([...configItems, newItem]);
    setNewItemLabel('');
    setExpandedCustomId(newId);
  };

  const handleDeleteCustomItem = (id: string) => {
    onUpdateConfig(configItems.filter(i => i.id !== id));
    if (expandedCustomId === id) setExpandedCustomId(null);
  };

  const totalSelectedCount = filteredItems.filter(i => i.isVisible).length;

  return (
    <div className="space-y-6">
      {/* Configuration Header */}
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 p-5 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl shadow-sm">
        <div>
          <div className="flex items-center gap-2 flex-wrap">
            <Settings className="w-5 h-5 text-indigo-600" />
            <h2 className="text-lg font-bold text-slate-900 tracking-tight">
              Financial Statement Element Selection &amp; Configuration
            </h2>
            <span className="text-[11px] font-semibold text-indigo-800 bg-indigo-50 border border-indigo-200 px-2 py-0.5 rounded-md">
              {totalSelectedCount} / {filteredItems.length} Selected
            </span>
            {hasSavedDefault && (
              <span
                className="text-[11px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md flex items-center gap-1"
                title={savedDefaultAt ? `User default saved at ${new Date(savedDefaultAt).toLocaleString()}` : 'Loaded from user-saved default configuration'}
              >
                <BookmarkCheck className="w-3 h-3 text-emerald-600" />
                Saved Default Active
              </span>
            )}
          </div>
          <p className="text-xs text-slate-600 mt-1">
            Select or deselect any line item across all 5 financial statements, customize labels, add custom items with year-wise amounts, and recalculate linked subtotals, ratios, previews, and exports automatically.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0">
          {onSaveAsDefault && (
            <button
              type="button"
              onClick={handleSaveDefaultClick}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg border transition-colors flex items-center gap-1.5 shadow-xs ${
                savedNotice
                  ? 'bg-emerald-600 text-white border-emerald-600'
                  : 'bg-indigo-600 hover:bg-indigo-700 text-white border-indigo-600'
              }`}
              title="Save current item selections, order, and custom labels as the automatic default for future launches and New Case resets"
            >
              {savedNotice ? <Check className="w-3.5 h-3.5" /> : <BookmarkCheck className="w-3.5 h-3.5" />}
              {savedNotice ? 'Saved as Default!' : 'Save Selections as Default'}
            </button>
          )}
          <button
            type="button"
            onClick={onResetDefaults}
            className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 rounded-lg border border-slate-300 transition-colors flex items-center gap-1.5"
            title={hasSavedDefault ? 'Reload your previously saved default selections' : 'Reset to factory default selections'}
          >
            <RefreshCw className="w-3.5 h-3.5" />
            {hasSavedDefault ? 'Restore Saved Default' : 'Reset Defaults'}
          </button>
          {hasSavedDefault && onResetFactoryDefaults && (
            <button
              type="button"
              onClick={onResetFactoryDefaults}
              className="px-3 py-1.5 text-xs font-medium text-rose-700 bg-white hover:bg-rose-50 rounded-lg border border-slate-300 hover:border-rose-300 transition-colors"
              title="Clear your saved default profile and restore original factory defaults"
            >
              Factory Defaults
            </button>
          )}
        </div>
      </div>

      {/* Statement Tabs + Section Filters */}
      <div className="flex flex-col gap-3 p-4 bg-white/90 backdrop-blur-md border border-slate-200/90 rounded-xl text-xs shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-lg border border-slate-200 flex-wrap">
            {STATEMENT_TABS.map(tab => (
              <button
                key={tab.id}
                onClick={() => handleSelectStatementTab(tab.id)}
                className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors ${
                  selectedStatement === tab.id
                    ? 'bg-emerald-600 text-white font-semibold'
                    : 'text-slate-700 hover:text-slate-900'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search elements..."
                aria-label="Search financial statement elements"
                className="pl-8 pr-3 py-1.5 w-44 sm:w-52 bg-white border border-slate-300 rounded-lg text-xs text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500"
              />
            </div>
            <button
              type="button"
              onClick={() => handleSelectScoped(true, false)}
              className="px-2.5 py-1.5 text-slate-700 hover:text-slate-900 bg-white hover:bg-slate-50 rounded-lg border border-slate-300 transition-colors flex items-center gap-1 font-medium"
            >
              <CheckSquare className="w-3.5 h-3.5 text-emerald-600" />
              Select All
            </button>
            <button
              type="button"
              onClick={() => handleSelectScoped(false, false)}
              className="px-2.5 py-1.5 text-slate-700 hover:text-rose-800 bg-white hover:bg-rose-50 rounded-lg border border-slate-300 hover:border-rose-300 transition-colors flex items-center gap-1 font-medium"
            >
              <Square className="w-3.5 h-3.5 text-rose-500" />
              Deselect All
            </button>
            <button
              type="button"
              onClick={() => handleSelectScoped(false, true)}
              className="px-2.5 py-1.5 text-slate-700 hover:text-slate-900 bg-white hover:bg-slate-50 rounded-lg border border-slate-300 transition-colors flex items-center gap-1"
            >
              <Square className="w-3.5 h-3.5 text-slate-500" />
              Deselect Optional
            </button>
          </div>
        </div>

        {availableSections.length > 2 && (
          <div className="flex items-center gap-2 flex-wrap pt-1 border-t border-slate-200/70">
            <span className="text-slate-600 font-medium">Section:</span>
            <div className="flex items-center gap-1 flex-wrap">
              {availableSections.map(sec => (
                <button
                  key={sec}
                  onClick={() => setSelectedSection(sec)}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors border ${
                    selectedSection === sec
                      ? 'bg-indigo-600 text-white border-indigo-600'
                      : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
                  }`}
                >
                  {sec}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Add Custom Item Widget */}
      <form
        onSubmit={handleAddCustomItem}
        className="p-4 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-xl flex flex-wrap items-center gap-2.5 text-xs shadow-sm"
      >
        <div className="flex items-center gap-2 text-slate-900 font-semibold shrink-0">
          <Plus className="w-4 h-4 text-emerald-600" />
          Add Custom Element:
        </div>
        <select
          value={newItemStatement}
          onChange={e => {
            const stmt = e.target.value as StatementConfigItem['statement'];
            setNewItemStatement(stmt);
            setNewItemSection(STATEMENT_SECTIONS[stmt][0]);
          }}
          className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500 shrink-0"
        >
          <option value="income">Income Statement</option>
          <option value="balanceSheet">Balance Sheet</option>
          <option value="cashFlow">Cash Flow Statement</option>
          <option value="equity">Statement of Changes in Equity</option>
          <option value="fixedAssets">Schedule of Fixed Assets</option>
        </select>
        <select
          value={newItemSection}
          onChange={e => setNewItemSection(e.target.value)}
          className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500 shrink-0"
        >
          {STATEMENT_SECTIONS[newItemStatement].map(sec => (
            <option key={sec} value={sec}>
              {sec}
            </option>
          ))}
        </select>
        <input
          type="text"
          placeholder="Custom element label (e.g. Sukuk Issuance Proceeds)"
          value={newItemLabel}
          onChange={e => setNewItemLabel(e.target.value)}
          className="flex-1 min-w-[200px] bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
        />
        <select
          value={newItemSign}
          onChange={e => setNewItemSign(e.target.value as 'positive' | 'negative')}
          className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500 shrink-0"
          title="Contribution sign in subtotals"
        >
          <option value="positive">+ Positive / Inflow</option>
          <option value="negative">- Negative / Outflow</option>
        </select>
        <button
          type="submit"
          className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-medium rounded-lg transition-colors shrink-0"
        >
          Add Item
        </button>
      </form>

      {/* Item List Table */}
      <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 uppercase tracking-wider text-[11px]">
              <tr>
                <th className="py-3 px-4 w-14 text-center font-semibold">Include</th>
                <th className="py-3 px-4 w-16 text-center font-semibold">Order</th>
                <th className="py-3 px-4 w-36 font-semibold">Statement</th>
                <th className="py-3 px-4 font-semibold text-slate-900">Statement Line Item Label</th>
                <th className="py-3 px-4 w-44 font-semibold">Section</th>
                <th className="py-3 px-4 w-28 font-semibold">Role</th>
                <th className="py-3 px-4 w-32 text-right font-semibold">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200/70 font-sans">
              {filteredItems.map((item, idx) => (
                <React.Fragment key={item.id}>
                  <tr
                    className={`hover:bg-slate-50 transition-colors ${
                      !item.isVisible ? 'opacity-55 bg-slate-100/60' : ''
                    }`}
                  >
                    <td className="py-2.5 px-4 text-center">
                      <input
                        type="checkbox"
                        checked={item.isVisible}
                        onChange={() => toggleVisibility(item.id)}
                        aria-label={`Include ${item.label}`}
                        className="w-4 h-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer accent-emerald-600"
                      />
                    </td>

                    <td className="py-2.5 px-4 text-center text-slate-500 font-mono">
                      {item.order}
                    </td>

                    <td className="py-2.5 px-4">
                      <span className="text-[11px] font-medium text-slate-700 bg-slate-100 border border-slate-200 px-2 py-0.5 rounded">
                        {STATEMENT_BADGE[item.statement]}
                      </span>
                    </td>

                    <td className="py-2.5 px-4">
                      <input
                        type="text"
                        value={item.label}
                        onChange={e => handleLabelChange(item.id, e.target.value)}
                        className="w-full bg-transparent hover:bg-slate-50 focus:bg-white border border-transparent focus:border-emerald-500 rounded px-2 py-1 text-slate-900 font-medium transition-colors focus:outline-none"
                      />
                    </td>

                    <td className="py-2.5 px-4">
                      <span className="text-[11px] text-slate-600">{item.section}</span>
                    </td>

                    <td className="py-2.5 px-4">
                      {item.isCustom ? (
                        <span className="text-[10px] font-semibold text-sky-800 bg-sky-50 border border-sky-300 px-2 py-0.5 rounded">
                          Custom
                        </span>
                      ) : item.isMandatory ? (
                        <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 px-2 py-0.5 rounded">
                          Core
                        </span>
                      ) : (
                        <span className="text-[10px] text-slate-500">Optional</span>
                      )}
                    </td>

                    <td className="py-2.5 px-4 text-right">
                      <div className="flex items-center justify-end gap-1">
                        {item.isCustom && years.length > 0 && (
                          <button
                            type="button"
                            onClick={() =>
                              setExpandedCustomId(expandedCustomId === item.id ? null : item.id)
                            }
                            className="px-2 py-0.5 text-[10px] font-semibold text-sky-700 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded transition-colors"
                          >
                            {expandedCustomId === item.id ? 'Hide Values' : 'Values'}
                          </button>
                        )}
                        <button
                          onClick={() => handleMove(item.id, 'up')}
                          disabled={idx === 0}
                          title="Move Up"
                          className="p-1 text-slate-500 hover:text-slate-900 disabled:opacity-30 rounded hover:bg-slate-100"
                        >
                          <ArrowUp className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => handleMove(item.id, 'down')}
                          disabled={idx === filteredItems.length - 1}
                          title="Move Down"
                          className="p-1 text-slate-500 hover:text-slate-900 disabled:opacity-30 rounded hover:bg-slate-100"
                        >
                          <ArrowDown className="w-3.5 h-3.5" />
                        </button>
                        {item.isCustom && (
                          <button
                            onClick={() => handleDeleteCustomItem(item.id)}
                            title="Delete Custom Item"
                            className="p-1 text-rose-600 hover:text-rose-700 rounded hover:bg-rose-50"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>

                  {item.isCustom && expandedCustomId === item.id && years.length > 0 && (
                    <tr className="bg-sky-50/50">
                      <td colSpan={7} className="px-6 py-3">
                        <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                          <span className="text-[11px] font-semibold text-sky-900">
                            Year-Wise Amounts for "{item.label}"
                          </span>
                          <div className="flex items-center gap-2">
                            <button
                              type="button"
                              onClick={() => handleFillCustomUnevenRatios(item)}
                              className="px-2.5 py-0.5 text-[10px] font-semibold bg-emerald-600 hover:bg-emerald-500 text-white rounded inline-flex items-center gap-1 transition-colors"
                              title="Populate logically varied uneven ratios across all years (>= 1.2% YoY variance)"
                            >
                              <Sparkles className="w-3 h-3" />
                              Fill Uneven Ratios
                            </button>
                            <select
                              value={item.sign || 'positive'}
                              onChange={e =>
                                handleCustomSignChange(
                                  item.id,
                                  e.target.value as 'positive' | 'negative'
                                )
                              }
                              className="text-[11px] bg-white border border-sky-200 rounded px-2 py-0.5 text-slate-800"
                            >
                              <option value="positive">Positive / Inflow (+)</option>
                              <option value="negative">Negative / Outflow (-)</option>
                            </select>
                          </div>
                        </div>
                        <div className="grid grid-cols-3 sm:grid-cols-6 md:grid-cols-9 gap-2">
                          {years.map(yr => (
                            <label key={yr} className="flex flex-col gap-1 text-[10px] text-slate-600">
                              <span className="font-mono font-semibold">{yr}</span>
                              <input
                                type="number"
                                value={item.customValues?.[yr] ?? 0}
                                onChange={e =>
                                  handleCustomYearValueChange(
                                    item.id,
                                    yr,
                                    parseFloat(e.target.value)
                                  )
                                }
                                className="w-full px-2 py-1 bg-white border border-sky-200 rounded font-mono text-right text-slate-900 focus:outline-none focus:ring-1 focus:ring-sky-500"
                              />
                            </label>
                          ))}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
