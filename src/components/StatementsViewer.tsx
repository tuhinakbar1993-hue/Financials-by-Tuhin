import React, { useState } from 'react';
import {
  FileSpreadsheet,
  FileText,
  Download,
  AlertTriangle,
  TrendingUp,
  Filter,
  Edit3,
  RotateCcw,
  Landmark,
  Scale,
  Layers,
  Building2,
  ShieldCheck,
  Calendar,
  CalendarCheck2,
  ArrowUpRight,
} from 'lucide-react';
import {
  CompanyProfile,
  IncomeStatementYear,
  BalanceSheetYear,
  CashFlowStatementYear,
  EquityStatementYear,
  ScheduleOfFixedAssetsYear,
  FinancialNoteItem,
  StatementConfigItem,
  ProjectionAssumptions,
  StatementCellOverrides,
  StatementLabelOverrides,
  StatementFormatMode,
  FinancialValidationIssue,
  PreviousFormatYearData,
  PreviousComparativePair,
  ProjectedFormatYearData,
  YearDataInput,
  DebtServiceCoverageYear,
  ParsedRepaymentSchedule,
} from '../types/financial';
import {
  exportIncomeStatementExcel,
  exportBalanceSheetExcel,
  exportCashFlowExcel,
  exportEquityStatementExcel,
  exportFixedAssetsExcel,
  exportNotesExcel,
  exportDocumentWord,
  exportDocumentPdf,
  exportCompletePackageExcel,
  exportCompletePackageWord,
  exportCompletePackagePdf,
  exportAllPreviousComparativePairsExcel,
  exportAllPreviousComparativePairsWord,
  exportAllPreviousComparativePairsPdf,
  exportPreviousFinancialsExcel,
  exportPreviousFinancialsWord,
  exportPreviousFinancialsPdf,
  exportProjectedFinancialsExcel,
  exportProjectedFinancialsWord,
  exportProjectedFinancialsPdf,
} from '../services/exportService';
import { PreviousFormatSuiteView, ProjectedFormatSuiteView } from './FormatStatementsViews';
import { YearSelectorControl } from './YearSelectorControl';

interface StatementsViewerProps {
  companyProfile: CompanyProfile;
  incomeData: IncomeStatementYear[];
  balanceData: BalanceSheetYear[];
  cashFlowData: CashFlowStatementYear[];
  equityData: EquityStatementYear[];
  fixedAssetsData: ScheduleOfFixedAssetsYear[];
  notes: FinancialNoteItem[];
  statementConfig: StatementConfigItem[];
  assumptions?: ProjectionAssumptions;
  onUpdateAssumptions?: (updated: ProjectionAssumptions) => void;
  cellOverrides?: StatementCellOverrides;
  labelOverrides?: StatementLabelOverrides;
  onUpdateStatementValue?: (statement: string, field: string, year: number, val: number) => void;
  onUpdateStatementLabel?: (statement: string, field: string, newLabel: string) => void;
  onResetOverrides?: () => void;
  onUpdateNotes?: (updatedNotes: FinancialNoteItem[]) => void;
  onUpdateFixedAssetsCategory?: (catIndex: number, field: string, val: number | string) => void;
  previousFormatData?: PreviousFormatYearData[];
  previousComparativePairs?: PreviousComparativePair[];
  getPairScopedNotes?: (pairData: PreviousFormatYearData[]) => FinancialNoteItem[];
  projectedFormatData?: ProjectedFormatYearData[];
  formatMode?: StatementFormatMode;
  onChangeFormatMode?: (mode: StatementFormatMode) => void;
  validationIssues?: FinancialValidationIssue[];
  selectedHistoricalYears?: number[];
  selectedProjectedYears?: number[];
  onToggleHistoricalYear?: (year: number) => void;
  onToggleProjectedYear?: (year: number) => void;
  onSetHistoricalYears?: (years: number[]) => void;
  onSetProjectedYears?: (years: number[]) => void;
  onUpdatePrimaryDriver?: (year: number, driver: keyof YearDataInput, val: number) => void;
  debtServiceCoverage?: DebtServiceCoverageYear[];
  isScheduleImplemented?: boolean;
  parsedSchedule?: ParsedRepaymentSchedule | null;
  onJumpToRepaymentSchedule?: () => void;
}

export const StatementsViewer: React.FC<StatementsViewerProps> = ({
  companyProfile,
  incomeData,
  balanceData,
  cashFlowData,
  equityData,
  fixedAssetsData,
  notes,
  assumptions,
  onUpdateAssumptions,
  cellOverrides = {},
  labelOverrides = {},
  onUpdateStatementValue = () => {},
  onUpdateStatementLabel = () => {},
  onResetOverrides = () => {},
  onUpdateNotes = () => {},
  onUpdateFixedAssetsCategory = () => {},
  previousFormatData = [],
  previousComparativePairs = [],
  getPairScopedNotes,
  projectedFormatData = [],
  formatMode: controlledFormatMode,
  onChangeFormatMode,
  validationIssues = [],
  selectedHistoricalYears = [2023, 2024, 2025],
  selectedProjectedYears,
  onToggleHistoricalYear,
  onToggleProjectedYear,
  onSetHistoricalYears,
  onSetProjectedYears,
  onUpdatePrimaryDriver,
  debtServiceCoverage = [],
  isScheduleImplemented = false,
  parsedSchedule = null,
  onJumpToRepaymentSchedule,
}) => {
  const [activeStatement, setActiveStatement] = useState<string>('cashFlow');
  const [yearViewMode, setYearViewMode] = useState<'all' | 'historical' | 'projected' | 'comparative5'>('all');
  const [isEditMode, setIsEditMode] = useState<boolean>(false);
  const [internalFormatMode, setInternalFormatMode] = useState<StatementFormatMode>('section76');
  const [showYearSelector, setShowYearSelector] = useState<boolean>(false);

  const formatMode = controlledFormatMode ?? internalFormatMode;
  const setFormatMode = onChangeFormatMode ?? setInternalFormatMode;

  const histCount = incomeData.filter(d => !d.isProjected).length;
  const projCount = incomeData.filter(d => d.isProjected).length;
  const projYearsList = selectedProjectedYears && selectedProjectedYears.length > 0
    ? selectedProjectedYears
    : incomeData.filter(d => d.isProjected).map(d => d.year);

  // Filter year data based on yearViewMode
  const filterYears = <T extends { year: number; isProjected: boolean }>(items: T[]): T[] => {
    switch (yearViewMode) {
      case 'historical':
        return items.filter(i => !i.isProjected);
      case 'projected':
        return items.filter(i => i.isProjected);
      case 'comparative5': {
        const hist = items.filter(i => !i.isProjected).slice(-2);
        const proj = items.filter(i => i.isProjected).slice(0, 3);
        return [...hist, ...proj];
      }
      case 'all':
      default:
        return items;
    }
  };

  const filteredIncome = filterYears(incomeData);
  const filteredBalance = filterYears(balanceData);
  const filteredCashFlow = filterYears(cashFlowData);
  const filteredEquity = filterYears(equityData);
  const filteredFixedAssets = filterYears(fixedAssetsData);

  // Helper to get effective cell value (checks cellOverrides first, falls back to original)
  const getVal = (stmt: string, field: string, year: number, fallback: number): number => {
    const key = `${stmt}_${field}`;
    if (cellOverrides[key] && cellOverrides[key][year] !== undefined) {
      return cellOverrides[key][year];
    }
    return fallback;
  };

  // Helper to get effective label (checks labelOverrides first, falls back to default)
  const getLbl = (stmt: string, field: string, defaultLabel: string): string => {
    const key = `${stmt}_${field}`;
    if (labelOverrides[key] !== undefined) {
      return labelOverrides[key];
    }
    return defaultLabel;
  };

  // Individual Export Handlers (exports updated values & labels!)
  const handleExportExcel = () => {
    switch (activeStatement) {
      case 'income':
        exportIncomeStatementExcel(companyProfile, filteredIncome);
        break;
      case 'balanceSheet':
        exportBalanceSheetExcel(companyProfile, filteredBalance);
        break;
      case 'cashFlow':
        exportCashFlowExcel(companyProfile, filteredCashFlow);
        break;
      case 'equity':
        exportEquityStatementExcel(companyProfile, filteredEquity);
        break;
      case 'fixedAssets':
        exportFixedAssetsExcel(companyProfile, filteredFixedAssets);
        break;
      case 'notes':
        exportNotesExcel(companyProfile, notes);
        break;
      case 'complete':
        exportCompletePackageExcel(
          companyProfile,
          filteredIncome,
          filteredBalance,
          filteredCashFlow,
          filteredEquity,
          filteredFixedAssets,
          notes,
          undefined,
          isScheduleImplemented ? debtServiceCoverage : undefined
        );
        break;
    }
  };

  const handleExportWord = () => {
    if (activeStatement === 'complete') {
      exportCompletePackageWord(
        companyProfile,
        filteredIncome,
        filteredBalance,
        filteredCashFlow,
        filteredEquity,
        filteredFixedAssets,
        notes,
        undefined,
        isScheduleImplemented ? debtServiceCoverage : undefined
      );
      return;
    }
    const { title, headers, rows } = getTableDataForExport(activeStatement);
    exportDocumentWord(companyProfile, title, headers, rows, `${activeStatement}_Statement.docx`);
  };

  const handleExportPdf = () => {
    if (activeStatement === 'complete') {
      exportCompletePackagePdf(
        companyProfile,
        filteredIncome,
        filteredBalance,
        filteredCashFlow,
        filteredEquity,
        filteredFixedAssets,
        notes,
        undefined,
        isScheduleImplemented ? debtServiceCoverage : undefined
      );
      return;
    }
    const { title, headers, rows } = getTableDataForExport(activeStatement);
    exportDocumentPdf(companyProfile, title, headers, rows, `${activeStatement}_Statement.pdf`);
  };

  function getTableDataForExport(stmt: string): { title: string; headers: string[]; rows: (string | number)[][] } {
    switch (stmt) {
      case 'income':
        return {
          title: 'Statement of Comprehensive Income',
          headers: ['Metric', ...filteredIncome.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)],
          rows: [
            [getLbl('income', 'revenue', 'Sales Revenue'), ...filteredIncome.map(d => getVal('income', 'revenue', d.year, d.revenue))],
            [getLbl('income', 'costOfSales', 'Cost of Sales'), ...filteredIncome.map(d => -getVal('income', 'costOfSales', d.year, d.costOfSales))],
            [getLbl('income', 'grossProfit', 'Gross Profit'), ...filteredIncome.map(d => getVal('income', 'grossProfit', d.year, d.grossProfit))],
            [getLbl('income', 'operatingExpenses', 'Operating Expenses'), ...filteredIncome.map(d => -getVal('income', 'operatingExpenses', d.year, d.operatingExpenses))],
            [getLbl('income', 'depreciationExpense', 'Depreciation Expense'), ...filteredIncome.map(d => -getVal('income', 'depreciationExpense', d.year, d.depreciationExpense))],
            [getLbl('income', 'operatingProfit', 'Operating Profit (EBIT)'), ...filteredIncome.map(d => getVal('income', 'operatingProfit', d.year, d.operatingProfit))],
            [getLbl('income', 'financeCost', 'Finance Cost'), ...filteredIncome.map(d => -getVal('income', 'financeCost', d.year, d.financeCost))],
            [getLbl('income', 'profitBeforeTax', 'Profit Before Tax'), ...filteredIncome.map(d => getVal('income', 'profitBeforeTax', d.year, d.profitBeforeTax))],
            [getLbl('income', 'netProfit', 'Net Profit After Tax'), ...filteredIncome.map(d => getVal('income', 'netProfit', d.year, d.netProfit))],
          ],
        };
      case 'balanceSheet':
        return {
          title: 'Statement of Financial Position',
          headers: ['Line Item', ...filteredBalance.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)],
          rows: [
            [getLbl('bs', 'cashAndEquivalents', 'Cash and Cash Equivalents'), ...filteredBalance.map(d => getVal('bs', 'cashAndEquivalents', d.year, d.cashAndEquivalents))],
            [getLbl('bs', 'tradeReceivables', 'Trade Receivables'), ...filteredBalance.map(d => getVal('bs', 'tradeReceivables', d.year, d.tradeReceivables))],
            [getLbl('bs', 'closingInventory', 'Closing Inventory'), ...filteredBalance.map(d => getVal('bs', 'closingInventory', d.year, d.closingInventory))],
            [getLbl('bs', 'totalCurrentAssets', 'Total Current Assets'), ...filteredBalance.map(d => getVal('bs', 'totalCurrentAssets', d.year, d.totalCurrentAssets))],
            [getLbl('bs', 'propertyPlantEquipment', 'Property, Plant & Equipment (NBV)'), ...filteredBalance.map(d => getVal('bs', 'propertyPlantEquipment', d.year, d.propertyPlantEquipment))],
            [getLbl('bs', 'totalAssets', 'TOTAL ASSETS'), ...filteredBalance.map(d => getVal('bs', 'totalAssets', d.year, d.totalAssets))],
            [getLbl('bs', 'totalCurrentLiabilities', 'Total Current Liabilities'), ...filteredBalance.map(d => getVal('bs', 'totalCurrentLiabilities', d.year, d.totalCurrentLiabilities))],
            [getLbl('bs', 'longTermBorrowings', 'Long-Term Borrowings'), ...filteredBalance.map(d => getVal('bs', 'longTermBorrowings', d.year, d.longTermBorrowings))],
            [getLbl('bs', 'totalLiabilities', 'TOTAL LIABILITIES'), ...filteredBalance.map(d => getVal('bs', 'totalLiabilities', d.year, d.totalLiabilities))],
            [getLbl('bs', 'totalEquity', 'TOTAL EQUITY'), ...filteredBalance.map(d => getVal('bs', 'totalEquity', d.year, d.totalEquity))],
            [getLbl('bs', 'totalLiabilitiesAndEquity', 'TOTAL LIABILITIES & EQUITY'), ...filteredBalance.map(d => getVal('bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity))],
          ],
        };
      case 'cashFlow':
      default:
        return {
          title: 'Statement of Cash Flows (Indirect Method)',
          headers: ['Cash Flow Item', ...filteredCashFlow.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)],
          rows: [
            [getLbl('cf', 'profitBeforeTax', 'Net Profit Before Tax'), ...filteredCashFlow.map(d => getVal('cf', 'profitBeforeTax', d.year, d.profitBeforeTax))],
            [getLbl('cf', 'adjDepreciation', 'Depreciation Adjustment'), ...filteredCashFlow.map(d => getVal('cf', 'adjDepreciation', d.year, d.adjDepreciation))],
            [getLbl('cf', 'adjFinanceCost', 'Finance Cost Adjustment'), ...filteredCashFlow.map(d => getVal('cf', 'adjFinanceCost', d.year, d.adjFinanceCost))],
            [getLbl('cf', 'operatingProfitBeforeWC', 'Operating Profit Before Working Capital'), ...filteredCashFlow.map(d => getVal('cf', 'operatingProfitBeforeWC', d.year, d.operatingProfitBeforeWC))],
            [getLbl('cf', 'netCashFromOperating', 'NET CASH FROM OPERATING ACTIVITIES'), ...filteredCashFlow.map(d => getVal('cf', 'netCashFromOperating', d.year, d.netCashFromOperating))],
            [getLbl('cf', 'purchasePPE', 'Purchase of PPE (Capex)'), ...filteredCashFlow.map(d => getVal('cf', 'purchasePPE', d.year, d.purchasePPE))],
            [getLbl('cf', 'netCashFromInvesting', 'NET CASH FROM / (USED IN) INVESTING'), ...filteredCashFlow.map(d => getVal('cf', 'netCashFromInvesting', d.year, d.netCashFromInvesting))],
            [getLbl('cf', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing'), ...filteredCashFlow.map(d => getVal('cf', 'financingProceedsBorrowings', d.year, d.financingProceedsBorrowings))],
            [getLbl('cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing'), ...filteredCashFlow.map(d => -getVal('cf', 'repaymentOfBorrowings', d.year, d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings)))],
            [getLbl('cf', 'financingDividendsPaid', 'Dividends Paid to Shareholders'), ...filteredCashFlow.map(d => getVal('cf', 'financingDividendsPaid', d.year, d.financingDividendsPaid))],
            [getLbl('cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Financing Cash Flow (Total)'), ...filteredCashFlow.map(d => getVal('cf', 'directlyEnteredFinancingCashFlow', d.year, d.directlyEnteredFinancingCashFlow))],
            [getLbl('cf', 'netCashFromFinancing', 'NET CASH FROM / (USED IN) FINANCING'), ...filteredCashFlow.map(d => getVal('cf', 'netCashFromFinancing', d.year, d.netCashFromFinancing))],
            [getLbl('cf', 'netChangeInCash', 'Net Increase / (Decrease) in Cash'), ...filteredCashFlow.map(d => getVal('cf', 'netChangeInCash', d.year, d.netChangeInCash))],
            [getLbl('cf', 'closingCashCalculated', 'Closing Cash & Cash Equivalents (Calculated)'), ...filteredCashFlow.map(d => getVal('cf', 'closingCashCalculated', d.year, d.closingCashCalculated))],
            [getLbl('cf', 'balanceSheetCash', 'Stated Cash in Balance Sheet'), ...filteredCashFlow.map(d => getVal('cf', 'balanceSheetCash', d.year, d.balanceSheetCash))],
          ],
        };
    }
  }

  const statementTabs = [
    { id: 'cashFlow', label: '1. Cash Flows', fullLabel: 'Statement of Cash Flows', icon: Scale },
    { id: 'income', label: '2. Income Statement', fullLabel: 'Statement of Comprehensive Income', icon: TrendingUp },
    { id: 'balanceSheet', label: '3. Balance Sheet', fullLabel: 'Statement of Financial Position', icon: Landmark },
    { id: 'equity', label: '4. Equity Movement', fullLabel: 'Statement of Changes in Equity', icon: Layers },
    { id: 'fixedAssets', label: '5. Fixed Assets', fullLabel: 'Schedule of Fixed Assets', icon: Building2 },
    { id: 'notes', label: '6. Notes (1-12)', fullLabel: 'Notes to Financials', icon: FileText },
    { id: 'complete', label: 'Full Package', fullLabel: 'Complete Financial Package', icon: FileSpreadsheet },
  ];

  const handleDownloadPair = (pair: PreviousComparativePair, fmt: 'xlsx' | 'docx' | 'pdf') => {
    const scoped = getPairScopedNotes ? getPairScopedNotes(pair.data) : notes;
    if (fmt === 'xlsx') exportPreviousFinancialsExcel(companyProfile, pair.data, scoped);
    else if (fmt === 'docx') void exportPreviousFinancialsWord(companyProfile, pair.data, scoped);
    else exportPreviousFinancialsPdf(companyProfile, pair.data, scoped);
  };

  return (
    <div className="space-y-6">
      {/* Top Banner, Format Switcher, Edit Toggle & Independent Download Options */}
      <div className="flex flex-col gap-4 p-5 bg-white/92 backdrop-blur-md border border-slate-200 rounded-2xl shadow-sm">
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <TrendingUp className="w-5 h-5 text-emerald-700" />
              <h2 className="text-lg font-bold text-slate-900 tracking-tight">
                Five Standard Financial Statements &amp; Schedules
              </h2>
              <span className="text-[11px] font-medium text-slate-600">
                · Section 76 &amp; Editable ·{' '}
                {validationIssues.length === 0 ? (
                  <span className="text-emerald-700 font-semibold inline-flex items-center gap-1">
                    <ShieldCheck className="w-3.5 h-3.5" />
                    All {histCount + projCount}Y Reconciled
                  </span>
                ) : (
                  <span className="text-amber-700 font-semibold inline-flex items-center gap-1">
                    <AlertTriangle className="w-3.5 h-3.5" />
                    {validationIssues.length} Audit Alert(s)
                  </span>
                )}
              </span>
            </div>
            <p className="text-xs text-slate-600 mt-1">
              {companyProfile.companyName} · {companyProfile.accountingFramework} · {companyProfile.reportingCurrency} in {companyProfile.reportingUnit}
            </p>
          </div>

          {/* Global Controls: Year Selector, Edit Mode Toggle & Export Strip */}
          <div className="flex flex-wrap items-center gap-2">
            {onJumpToRepaymentSchedule && (
              <button
                type="button"
                onClick={onJumpToRepaymentSchedule}
                className={`px-3 py-1.5 text-xs font-semibold rounded-lg border transition-all flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  isScheduleImplemented
                    ? 'bg-emerald-600 hover:bg-emerald-700 text-white border-emerald-600 shadow-sm'
                    : 'bg-white hover:bg-emerald-50 text-emerald-700 border-slate-300 hover:border-emerald-400'
                }`}
                title="Open Repayment Schedule Upload & DSCR Studio"
              >
                <CalendarCheck2 className="w-3.5 h-3.5" />
                <span>{isScheduleImplemented ? 'Schedule Locked' : 'Repayment Schedule'}</span>
                <ArrowUpRight className="w-3 h-3" />
              </button>
            )}

            {onToggleHistoricalYear && onToggleProjectedYear && onSetHistoricalYears && onSetProjectedYears && (
              <button
                type="button"
                onClick={() => setShowYearSelector(!showYearSelector)}
                className={`px-3 py-1.5 text-xs font-semibold rounded-lg border transition-all flex items-center gap-1.5 whitespace-nowrap shrink-0 ${
                  showYearSelector
                    ? 'bg-sky-600 text-white border-sky-600 shadow-sm'
                    : 'bg-white hover:bg-slate-50 text-slate-700 border-slate-300'
                }`}
              >
                <Calendar className="w-3.5 h-3.5" />
                <span>Years ({histCount}H + {projCount}P)</span>
              </button>
            )}

            {/* Universal Edit Mode Button */}
            <button
              onClick={() => setIsEditMode(!isEditMode)}
              className={`px-3.5 py-1.5 text-xs font-semibold rounded-lg border transition-all flex items-center gap-1.5 shadow-sm whitespace-nowrap shrink-0 ${
                isEditMode
                  ? 'bg-amber-500 text-white border-amber-500 ring-2 ring-amber-400/30'
                  : 'bg-white hover:bg-slate-50 text-slate-700 border-slate-300 hover:text-slate-900'
              }`}
              title="Toggle inline editing for every line item title and financial number"
            >
              <Edit3 className="w-3.5 h-3.5" />
              {isEditMode ? 'Finish Editing (Active)' : 'Edit Values & Labels'}
            </button>

            {isEditMode && formatMode === 'section76' && (
              <button
                onClick={onResetOverrides}
                className="px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:text-rose-700 bg-white hover:bg-rose-50 rounded-lg border border-slate-300 transition-colors flex items-center gap-1 whitespace-nowrap shrink-0"
                title="Reset any custom numbers back to the calculation engine defaults"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                Reset
              </button>
            )}

            {/* Individual Statement Export Options for Section 76 */}
            {formatMode === 'section76' && (
              <div className="flex items-center gap-1 pl-2 border-l border-slate-200">
                <button
                  onClick={handleExportExcel}
                  className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-emerald-50 hover:text-emerald-800 rounded-lg border border-slate-300 transition-colors flex items-center gap-1.5 whitespace-nowrap shrink-0"
                  title="Export this statement to Microsoft Excel (.xlsx)"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                  Excel
                </button>
                <button
                  onClick={handleExportWord}
                  className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-blue-50 hover:text-blue-800 rounded-lg border border-slate-300 transition-colors flex items-center gap-1.5 whitespace-nowrap shrink-0"
                  title="Export this statement to Microsoft Word (.docx)"
                >
                  <FileText className="w-3.5 h-3.5 text-blue-600" />
                  Word
                </button>
                <button
                  onClick={handleExportPdf}
                  className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-rose-50 hover:text-rose-800 rounded-lg border border-slate-300 transition-colors flex items-center gap-1.5 whitespace-nowrap shrink-0"
                  title="Export this statement to Adobe PDF (.pdf)"
                >
                  <Download className="w-3.5 h-3.5 text-rose-600" />
                  PDF
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Statement Presentation Format Bar */}
        <div className="pt-3 border-t border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-2.5">
          <div className="flex flex-wrap items-center gap-1.5">
            <span className="text-[11px] font-semibold text-slate-600 mr-1">Presentation Format:</span>
            <button
              type="button"
              onClick={() => setFormatMode('section76')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all border whitespace-nowrap shrink-0 ${
                formatMode === 'section76'
                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              Section 76 Standard (All Years)
            </button>
            <button
              type="button"
              onClick={() => setFormatMode('previous_format')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all border whitespace-nowrap shrink-0 ${
                formatMode === 'previous_format'
                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              Previous Financials Format ({histCount}Y)
            </button>
            <button
              type="button"
              onClick={() => setFormatMode('projected_format')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all border whitespace-nowrap shrink-0 ${
                formatMode === 'projected_format'
                  ? 'bg-sky-600 text-white border-sky-600 shadow-sm'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              Projected Financials Format ({projCount}Y)
            </button>
          </div>
          <span className="text-[11px] text-slate-500 font-mono">
            Fixed Inputs Preserved: Sales · Inventory · Finance Cost · NPAT · Borrowings Repay · Gross FA
          </span>
        </div>

        {/* Always-Accessible Quick Separate Downloads Strip (Previous Comparative Sets & Projected All-Years) */}
        <div className="pt-3 border-t border-slate-200/80 grid grid-cols-1 xl:grid-cols-2 gap-3">
          {/* Option 1: Previous Financials Comparative Sets */}
          <div className="p-3 bg-emerald-50/55 border border-emerald-200 rounded-xl flex flex-col gap-2">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="text-xs font-bold text-emerald-950">
                Previous Financials Downloads ({previousComparativePairs.length} Comparative {previousComparativePairs.length === 1 ? 'Set' : 'Sets'})
              </div>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsExcel(companyProfile, previousComparativePairs, pairData => (getPairScopedNotes ? getPairScopedNotes(pairData) : notes))}
                  className="px-2 py-1 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-semibold transition-colors whitespace-nowrap shrink-0"
                  title="Download all Previous Comparative sets as Excel workbooks"
                >
                  All .XLSX
                </button>
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsWord(companyProfile, previousComparativePairs, pairData => (getPairScopedNotes ? getPairScopedNotes(pairData) : notes))}
                  className="px-2 py-1 rounded-md bg-blue-600 hover:bg-blue-700 text-white text-[10px] font-semibold transition-colors whitespace-nowrap shrink-0"
                  title="Download all Previous Comparative sets as Word documents"
                >
                  All .DOCX
                </button>
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsPdf(companyProfile, previousComparativePairs, pairData => (getPairScopedNotes ? getPairScopedNotes(pairData) : notes))}
                  className="px-2 py-1 rounded-md bg-rose-600 hover:bg-rose-700 text-white text-[10px] font-semibold transition-colors whitespace-nowrap shrink-0"
                  title="Download all Previous Comparative sets as PDF statements"
                >
                  All .PDF
                </button>
              </div>
            </div>
            {previousComparativePairs.length > 0 && (
              <div className="flex flex-wrap gap-2">
                {previousComparativePairs.map(pair => (
                  <div
                    key={pair.id}
                    className="flex items-center gap-1.5 bg-white border border-emerald-200/90 rounded-lg px-2 py-1 text-[11px]"
                  >
                    <span className="font-semibold text-slate-800">{pair.label}</span>
                    <button
                      type="button"
                      onClick={() => handleDownloadPair(pair, 'xlsx')}
                      className="px-1.5 py-0.5 rounded bg-emerald-50 hover:bg-emerald-600 text-emerald-800 hover:text-white font-mono text-[10px] font-semibold transition-colors"
                    >
                      XLS
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDownloadPair(pair, 'docx')}
                      className="px-1.5 py-0.5 rounded bg-blue-50 hover:bg-blue-600 text-blue-800 hover:text-white font-mono text-[10px] font-semibold transition-colors"
                    >
                      DOC
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDownloadPair(pair, 'pdf')}
                      className="px-1.5 py-0.5 rounded bg-rose-50 hover:bg-rose-600 text-rose-800 hover:text-white font-mono text-[10px] font-semibold transition-colors"
                    >
                      PDF
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Option 2: Projected Financials All Selected Years */}
          <div className="p-3 bg-sky-50/55 border border-sky-200 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div>
              <div className="text-xs font-bold text-sky-950">
                Projected Financials Download ({projectedFormatData.length} Selected {projectedFormatData.length === 1 ? 'Year' : 'Years'})
              </div>
              <div className="text-[11px] text-slate-600">
                All selected projected years combined in exact landscape format
              </div>
            </div>
            <div className="flex items-center gap-1.5 shrink-0">
              <button
                type="button"
                onClick={() => exportProjectedFinancialsExcel(companyProfile, projectedFormatData)}
                className="px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-semibold flex items-center gap-1 shadow-xs transition-colors whitespace-nowrap shrink-0"
              >
                <FileSpreadsheet className="w-3 h-3" />
                .XLSX
              </button>
              <button
                type="button"
                onClick={() => exportProjectedFinancialsWord(companyProfile, projectedFormatData)}
                className="px-2.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-semibold flex items-center gap-1 shadow-xs transition-colors whitespace-nowrap shrink-0"
              >
                <FileText className="w-3 h-3" />
                .DOCX
              </button>
              <button
                type="button"
                onClick={() => exportProjectedFinancialsPdf(companyProfile, projectedFormatData)}
                className="px-2.5 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-[11px] font-semibold flex items-center gap-1 shadow-xs transition-colors whitespace-nowrap shrink-0"
              >
                <Download className="w-3 h-3" />
                .PDF
              </button>
            </div>
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

      {/* Edit Mode Notification Banner */}
      {isEditMode && (
        <div className="p-3.5 bg-amber-50/95 border border-amber-300 rounded-xl text-xs text-amber-950 flex items-center justify-between gap-3 animate-in fade-in duration-200 shadow-xs">
          <div className="flex items-center gap-2">
            <Edit3 className="w-4 h-4 text-amber-600 shrink-0" />
            <span>
              <strong>Direct Editing Enabled:</strong> Edit primary fixed inputs or statement cells directly in the tables below. All dependent calculations and exports update immediately.
            </span>
          </div>
          <button
            onClick={() => setIsEditMode(false)}
            className="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded font-medium text-[11px] shrink-0"
          >
            Done Editing
          </button>
        </div>
      )}

      {/* Exact PDF Compliance Format Modes */}
      {formatMode === 'previous_format' && (
        <PreviousFormatSuiteView
          companyProfile={companyProfile}
          prevData={previousFormatData}
          comparativePairs={previousComparativePairs}
          getPairScopedNotes={getPairScopedNotes}
          notes={notes}
          isEditMode={isEditMode}
          onUpdatePrimaryDriver={onUpdatePrimaryDriver}
        />
      )}

      {formatMode === 'projected_format' && (
        <ProjectedFormatSuiteView
          companyProfile={companyProfile}
          projData={projectedFormatData}
          isEditMode={isEditMode}
          onUpdatePrimaryDriver={onUpdatePrimaryDriver}
        />
      )}

      {/* Existing Section 76 7-Tab Suite */}
      {formatMode === 'section76' && (
        <>
          {/* Statement Tabs Bar & Horizon Selector */}
          <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 pb-2">
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 max-w-full scrollbar-thin">
              {statementTabs.map(tab => {
                const Icon = tab.icon;
                const isActive = activeStatement === tab.id;
                return (
                  <button
                    key={tab.id}
                    onClick={() => setActiveStatement(tab.id)}
                    className={`px-3 py-2 text-xs font-semibold rounded-lg transition-all whitespace-nowrap shrink-0 flex items-center gap-2 ${
                      isActive
                        ? 'bg-emerald-600 text-white shadow-sm'
                        : 'bg-white/90 text-slate-700 hover:bg-slate-50 hover:text-slate-900 border border-slate-200'
                    }`}
                  >
                    <Icon className={`w-3.5 h-3.5 ${isActive ? 'text-white' : 'text-slate-500'}`} />
                    <span>{tab.label}</span>
                  </button>
                );
              })}
            </div>

            {/* Year Filter Controls & Horizon Selector */}
            {activeStatement !== 'notes' && (
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex items-center gap-1.5 text-xs bg-white/92 border border-slate-200 p-1 rounded-lg shadow-xs">
                  <Filter className="w-3.5 h-3.5 text-slate-500 ml-1.5" />
                  <button
                    onClick={() => setYearViewMode('all')}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                      yearViewMode === 'all' ? 'bg-emerald-50 text-emerald-800 font-bold' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    All ({histCount}H+{projCount}P)
                  </button>
                  <button
                    onClick={() => setYearViewMode('historical')}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                      yearViewMode === 'historical' ? 'bg-emerald-50 text-emerald-800 font-bold' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Historical ({histCount}Y)
                  </button>
                  <button
                    onClick={() => setYearViewMode('comparative5')}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                      yearViewMode === 'comparative5' ? 'bg-emerald-50 text-emerald-800 font-bold' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Comparative (5Y)
                  </button>
                  <button
                    onClick={() => setYearViewMode('projected')}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                      yearViewMode === 'projected' ? 'bg-emerald-50 text-emerald-800 font-bold' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Projected ({projCount}Y)
                  </button>
                </div>

                {/* Direct Number of Projected Years Selector */}
                {assumptions && onUpdateAssumptions && (
                  <div className="flex items-center gap-1.5 text-xs bg-white/92 border border-slate-200 px-2.5 py-1 rounded-lg shadow-xs">
                    <span className="text-slate-600 text-[11px] font-medium">Horizon:</span>
                    <select
                      value={projCount}
                      onChange={(e) => onUpdateAssumptions({ ...assumptions, projectedYearsCount: Number(e.target.value) })}
                      className="bg-slate-50 text-emerald-700 font-bold rounded px-1.5 py-0.5 border border-slate-300 text-xs focus:outline-none focus:ring-1 focus:ring-emerald-500"
                    >
                      {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 18, 20].map(y => (
                        <option key={y} value={y}>{y} {y === 1 ? 'Year' : 'Years'}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Render Statement Body */}
          {activeStatement === 'cashFlow' && (
            <div className="space-y-6">
              <StatementOfCashFlowsTable
                cashFlowData={filteredCashFlow}
                currency={companyProfile.reportingCurrency}
                companyProfile={companyProfile}
                isEditMode={isEditMode}
                getVal={getVal}
                getLbl={getLbl}
                onUpdateValue={onUpdateStatementValue}
                onUpdateLabel={onUpdateStatementLabel}
              />
              {isScheduleImplemented && debtServiceCoverage.some(d => d.totalDebtService > 0) && (
                <DebtServiceCoverageSummaryPanel
                  coverage={debtServiceCoverage}
                  companyProfile={companyProfile}
                  parsedSchedule={parsedSchedule}
                  onJumpToRepaymentSchedule={onJumpToRepaymentSchedule}
                />
              )}
            </div>
          )}

          {activeStatement === 'income' && (
            <IncomeStatementTable
              incomeData={filteredIncome}
              currency={companyProfile.reportingCurrency}
              companyProfile={companyProfile}
              isEditMode={isEditMode}
              getVal={getVal}
              getLbl={getLbl}
              onUpdateValue={onUpdateStatementValue}
              onUpdateLabel={onUpdateStatementLabel}
            />
          )}

          {activeStatement === 'balanceSheet' && (
            <BalanceSheetTable
              balanceData={filteredBalance}
              currency={companyProfile.reportingCurrency}
              companyProfile={companyProfile}
              isEditMode={isEditMode}
              getVal={getVal}
              getLbl={getLbl}
              onUpdateValue={onUpdateStatementValue}
              onUpdateLabel={onUpdateStatementLabel}
            />
          )}

          {activeStatement === 'equity' && (
            <EquityStatementTable
              equityData={filteredEquity}
              currency={companyProfile.reportingCurrency}
              companyProfile={companyProfile}
              isEditMode={isEditMode}
              getVal={getVal}
              getLbl={getLbl}
              onUpdateValue={onUpdateStatementValue}
              onUpdateLabel={onUpdateStatementLabel}
            />
          )}

          {activeStatement === 'fixedAssets' && (
            <FixedAssetsScheduleTable
              fixedAssetsData={filteredFixedAssets}
              currency={companyProfile.reportingCurrency}
              companyProfile={companyProfile}
              isEditMode={isEditMode}
              onUpdateCategory={onUpdateFixedAssetsCategory}
            />
          )}

          {activeStatement === 'notes' && (
            <NotesViewer
              notes={notes}
              currency={companyProfile.reportingCurrency}
              isEditMode={isEditMode}
              onUpdateNotes={onUpdateNotes}
            />
          )}

          {activeStatement === 'complete' && (
            <CompletePackagePreview
              companyProfile={companyProfile}
              incomeData={filteredIncome}
              balanceData={filteredBalance}
              cashFlowData={filteredCashFlow}
              equityData={filteredEquity}
              fixedAssetsData={filteredFixedAssets}
              notes={notes}
              debtServiceCoverage={isScheduleImplemented ? debtServiceCoverage : []}
              onExportExcel={handleExportExcel}
              onExportWord={handleExportWord}
              onExportPdf={handleExportPdf}
            />
          )}
        </>
      )}
    </div>
  );
};

// ----------------- STATEMENT MASTHEAD COMPONENT ----------------- //
interface StatementMastheadProps {
  title: string;
  subtitle: string;
  standard: string;
  companyProfile?: CompanyProfile;
}

const StatementMasthead: React.FC<StatementMastheadProps> = ({
  title,
  subtitle,
  standard,
  companyProfile,
}) => {
  return (
    <div className="p-4 sm:p-5 bg-slate-50/85 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
      <div>
        <div className="flex items-center gap-2 text-[11px] font-mono">
          <span className="uppercase tracking-wider text-emerald-700 font-bold">
            {companyProfile?.companyName || 'Corporate Entity'}
          </span>
          <span className="text-slate-300" aria-hidden="true">·</span>
          <span className="text-slate-600">
            {companyProfile?.accountingFramework || 'IFRS / IAS Standards'}
          </span>
        </div>
        <h3 className="text-base sm:text-lg font-bold text-slate-900 tracking-tight mt-0.5 font-sans">
          {title}
        </h3>
        <p className="text-[11px] text-slate-600 mt-0.5">
          {subtitle} {companyProfile ? `· Reported in ${companyProfile.reportingCurrency} (${companyProfile.reportingUnit})` : ''}
        </p>
      </div>
      <div className="flex items-center gap-2 shrink-0">
        <span className="text-[11px] font-mono font-bold uppercase tracking-wider text-emerald-700">
          {standard}
        </span>
      </div>
    </div>
  );
};

function fmtStatementNum(val: number): string {
  const absVal = Math.abs(val);
  if (absVal < 0.005) return '0';
  const hasCents = Math.abs(absVal - Math.round(absVal)) >= 0.005;
  return hasCents
    ? absVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : Math.round(absVal).toLocaleString('en-US');
}

// ----------------- EDITABLE CELL COMPONENT ----------------- //
interface EditableCellProps {
  val: number;
  isEditMode: boolean;
  onSave: (newVal: number) => void;
  className?: string;
  isNegative?: boolean;
}

const EditableCell: React.FC<EditableCellProps> = ({
  val,
  isEditMode,
  onSave,
  className = '',
  isNegative = false,
}) => {
  const [localVal, setLocalVal] = useState<string>(String(val));
  const [isEditingSingle, setIsEditingSingle] = useState(false);

  React.useEffect(() => {
    setLocalVal(String(val));
  }, [val]);

  const handleBlur = () => {
    setIsEditingSingle(false);
    const num = Number(localVal);
    if (!isNaN(num)) {
      onSave(num);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      handleBlur();
    }
  };

  if (isEditMode || isEditingSingle) {
    return (
      <input
        type="number"
        value={localVal}
        onChange={(e) => setLocalVal(e.target.value)}
        onBlur={handleBlur}
        onKeyDown={handleKeyDown}
        autoFocus={isEditingSingle}
        className="w-24 sm:w-28 bg-white border border-emerald-600 rounded px-1.5 py-0.5 text-right font-mono text-xs text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500 shadow-xs"
      />
    );
  }

  return (
    <span
      onClick={() => setIsEditingSingle(true)}
      title="Click to edit value directly"
      className={`cursor-pointer hover:underline hover:text-emerald-700 transition-colors font-mono tabular-nums ${className}`}
    >
      {isNegative || val < 0 ? `(${fmtStatementNum(val)})` : fmtStatementNum(val)}
    </span>
  );
};

// ----------------- EDITABLE LABEL COMPONENT ----------------- //
interface EditableLabelProps {
  label: string;
  isEditMode: boolean;
  onSave: (newLabel: string) => void;
  className?: string;
}

const EditableLabel: React.FC<EditableLabelProps> = ({
  label,
  isEditMode,
  onSave,
  className = '',
}) => {
  const [localLabel, setLocalLabel] = useState(label);
  const [isEditingSingle, setIsEditingSingle] = useState(false);

  React.useEffect(() => {
    setLocalLabel(label);
  }, [label]);

  const handleBlur = () => {
    setIsEditingSingle(false);
    if (localLabel.trim()) {
      onSave(localLabel.trim());
    }
  };

  if (isEditMode || isEditingSingle) {
    return (
      <input
        type="text"
        value={localLabel}
        onChange={(e) => setLocalLabel(e.target.value)}
        onBlur={handleBlur}
        onKeyDown={(e) => e.key === 'Enter' && handleBlur()}
        autoFocus={isEditingSingle}
        className="w-full bg-white border border-emerald-600 rounded px-1.5 py-0.5 text-left text-xs font-medium text-slate-900 focus:outline-none focus:ring-1 focus:ring-emerald-500 shadow-xs"
      />
    );
  }

  return (
    <span
      onClick={() => setIsEditingSingle(true)}
      title="Click to edit row label"
      className={`cursor-pointer hover:text-emerald-700 transition-colors ${className}`}
    >
      {label}
    </span>
  );
};

// ----------------- STATEMENT OF CASH FLOWS TABLE (SECTION 76.3) ----------------- //
function StatementOfCashFlowsTable({
  cashFlowData,
  currency,
  companyProfile,
  isEditMode,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  cashFlowData: CashFlowStatementYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  return (
    <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
      <StatementMasthead
        title="Statement of Cash Flows (Indirect Method)"
        subtitle="Comprehensive cash flow movements from Operating, Investing, and Financing activities"
        standard="Section 76.3 & IAS 7"
        companyProfile={companyProfile}
      />

      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs font-mono">
          <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
            <tr>
              <th className="py-3 px-4 font-sans font-semibold text-slate-900 tracking-wide">
                Activity / Financial Cash Flow Account
              </th>
              {cashFlowData.map(d => (
                <th key={d.year} className={`py-2.5 px-3 text-right ${d.isProjected ? 'bg-sky-50/90 text-sky-900' : 'text-slate-800'}`}>
                  <div className="flex flex-col items-end">
                    <span className="font-bold text-xs">{d.year}</span>
                    <span className={`text-[9px] font-sans uppercase tracking-wider font-semibold ${
                      d.isProjected ? 'text-sky-700' : 'text-slate-500'
                    }`}>
                      {d.isProjected ? 'Forecast' : 'Audited'}
                    </span>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200/70 tabular-nums">
            {/* A. Operating Activities */}
            <tr className="bg-emerald-50/75 font-sans font-bold text-emerald-800">
              <td colSpan={cashFlowData.length + 1} className="py-2.5 px-4 text-xs tracking-wider uppercase">
                A. Cash Flows from Operating Activities
              </td>
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-800 font-medium">
                <EditableLabel
                  label={getLbl('cf', 'profitBeforeTax', 'Net Profit Before Tax')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'profitBeforeTax', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-800">
                  <EditableCell
                    val={getVal('cf', 'profitBeforeTax', d.year, d.profitBeforeTax)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'profitBeforeTax', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'adjDepreciation', 'Adjustments for Depreciation')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'adjDepreciation', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'adjDepreciation', d.year, d.adjDepreciation)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'adjDepreciation', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'adjAmortization', 'Adjustments for Amortization')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'adjAmortization', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'adjAmortization', d.year, d.adjAmortization)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'adjAmortization', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'adjFinanceCost', 'Finance Cost Adjustment')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'adjFinanceCost', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'adjFinanceCost', d.year, d.adjFinanceCost)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'adjFinanceCost', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'adjInterestIncome', 'Interest & Investment Income')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'adjInterestIncome', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'adjInterestIncome', d.year, d.adjInterestIncome)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'adjInterestIncome', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-slate-100/80 font-semibold text-slate-800">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('cf', 'operatingProfitBeforeWC', 'Operating Profit Before Working Capital Changes')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'operatingProfitBeforeWC', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-emerald-700 font-bold">
                  <EditableCell
                    val={getVal('cf', 'operatingProfitBeforeWC', d.year, d.operatingProfitBeforeWC)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'operatingProfitBeforeWC', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'changeInInventory', '(Increase) / Decrease in Inventory')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'changeInInventory', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'changeInInventory', d.year, d.changeInInventory)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'changeInInventory', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'changeInReceivables', '(Increase) / Decrease in Trade Receivables')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'changeInReceivables', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'changeInReceivables', d.year, d.changeInReceivables)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'changeInReceivables', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'changeInPayables', 'Increase / (Decrease) in Trade Payables')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'changeInPayables', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'changeInPayables', d.year, d.changeInPayables)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'changeInPayables', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'incomeTaxPaid', 'Income Tax Paid')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'incomeTaxPaid', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-rose-700">
                  <EditableCell
                    val={getVal('cf', 'incomeTaxPaid', d.year, d.incomeTaxPaid)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'incomeTaxPaid', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-emerald-50/60 font-bold border-t border-emerald-200">
              <td className="py-2.5 px-4 font-sans text-emerald-800">
                <EditableLabel
                  label={getLbl('cf', 'netCashFromOperating', 'Net Cash Generated from / (Used in) Operating Activities')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'netCashFromOperating', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className={`py-2.5 px-3 text-right ${d.netCashFromOperating >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                  <EditableCell
                    val={getVal('cf', 'netCashFromOperating', d.year, d.netCashFromOperating)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'netCashFromOperating', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {/* B. Investing Activities */}
            <tr className="bg-sky-50/80 font-sans font-bold text-sky-800">
              <td colSpan={cashFlowData.length + 1} className="py-2.5 px-4 text-xs tracking-wider uppercase">
                B. Cash Flows from Investing Activities
              </td>
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'purchasePPE', 'Purchase of Property, Plant & Equipment (Capex)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'purchasePPE', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-rose-700">
                  <EditableCell
                    val={getVal('cf', 'purchasePPE', d.year, d.purchasePPE)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'purchasePPE', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'salePPE', 'Proceeds from Sale of Fixed Assets')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'salePPE', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'salePPE', d.year, d.salePPE)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'salePPE', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-sky-50/60 font-bold border-t border-sky-200">
              <td className="py-2.5 px-4 font-sans text-sky-800">
                <EditableLabel
                  label={getLbl('cf', 'netCashFromInvesting', 'Net Cash Generated from / (Used in) Investing Activities')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'netCashFromInvesting', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className={`py-2.5 px-3 text-right ${d.netCashFromInvesting >= 0 ? 'text-sky-700' : 'text-rose-700'}`}>
                  <EditableCell
                    val={getVal('cf', 'netCashFromInvesting', d.year, d.netCashFromInvesting)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'netCashFromInvesting', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {/* C. Financing Activities */}
            <tr className="bg-amber-50/85 font-sans font-bold text-amber-900">
              <td colSpan={cashFlowData.length + 1} className="py-2.5 px-4 text-xs tracking-wider uppercase flex items-center justify-between">
                <span>C. Cash Flows from Financing Activities</span>
                <span className="text-[10px] font-semibold normal-case text-amber-900 bg-amber-100 px-2 py-0.5 rounded border border-amber-300">
                  Mandatory Direct User Input
                </span>
              </td>
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'financingProceedsBorrowings', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-800">
                  <EditableCell
                    val={getVal('cf', 'financingProceedsBorrowings', d.year, d.financingProceedsBorrowings)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'financingProceedsBorrowings', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-rose-50/65 font-medium">
              <td className="py-2.5 px-4 font-sans text-rose-800 pl-6 flex items-center gap-1.5">
                <EditableLabel
                  label={getLbl('cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing (Outflow) *')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'repaymentOfBorrowings', lbl)}
                />
                <span className="text-[10px] bg-rose-100 text-rose-800 border border-rose-300 px-1.5 py-0.2 rounded font-semibold">
                  User Input
                </span>
              </td>
              {cashFlowData.map(d => {
                const repVal = getVal('cf', 'repaymentOfBorrowings', d.year, d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings));
                return (
                  <td key={d.year} className="py-2.5 px-3 text-right text-rose-700 font-bold">
                    <EditableCell
                      val={repVal}
                      isEditMode={isEditMode}
                      onSave={(val) => onUpdateValue('cf', 'repaymentOfBorrowings', d.year, val)}
                      isNegative
                    />
                  </td>
                );
              })}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700 pl-6">
                <EditableLabel
                  label={getLbl('cf', 'financingDividendsPaid', 'Dividends Paid to Shareholders')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'financingDividendsPaid', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-rose-700">
                  <EditableCell
                    val={getVal('cf', 'financingDividendsPaid', d.year, d.financingDividendsPaid)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'financingDividendsPaid', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-amber-50/75 font-semibold border-t border-amber-200">
              <td className="py-2 px-4 font-sans text-amber-900">
                <EditableLabel
                  label={getLbl('cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Cash Flow from Financing Activities *')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'directlyEnteredFinancingCashFlow', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-amber-800 font-bold">
                  <EditableCell
                    val={getVal('cf', 'directlyEnteredFinancingCashFlow', d.year, d.directlyEnteredFinancingCashFlow)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'directlyEnteredFinancingCashFlow', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-amber-50/90 font-bold border-t border-amber-200">
              <td className="py-2.5 px-4 font-sans text-amber-900">
                <EditableLabel
                  label={getLbl('cf', 'netCashFromFinancing', 'Net Cash Generated from / (Used in) Financing Activities')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'netCashFromFinancing', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2.5 px-3 text-right text-amber-900">
                  <EditableCell
                    val={getVal('cf', 'netCashFromFinancing', d.year, d.netCashFromFinancing)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'netCashFromFinancing', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {/* D. Net Increase / Decrease in Cash */}
            <tr className="bg-slate-100 font-bold text-slate-900 border-t-2 border-slate-300">
              <td className="py-3 px-4 font-sans">
                <EditableLabel
                  label={getLbl('cf', 'netChangeInCash', 'Net Increase / (Decrease) in Cash and Cash Equivalents')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'netChangeInCash', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className={`py-3 px-3 text-right ${d.netChangeInCash >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                  <EditableCell
                    val={getVal('cf', 'netChangeInCash', d.year, d.netChangeInCash)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'netChangeInCash', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600">
                <EditableLabel
                  label={getLbl('cf', 'openingCash', 'Opening Cash and Cash Equivalents')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'openingCash', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('cf', 'openingCash', d.year, d.openingCash)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'openingCash', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="font-bold text-emerald-800 bg-emerald-50/85 border-b-4 border-double border-emerald-600 border-t border-emerald-200">
              <td className="py-3 px-4 font-sans text-sm">
                <EditableLabel
                  label={getLbl('cf', 'closingCashCalculated', 'Closing Cash and Cash Equivalents (Calculated)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'closingCashCalculated', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-3 px-3 text-right text-emerald-800 font-bold text-sm">
                  <EditableCell
                    val={getVal('cf', 'closingCashCalculated', d.year, d.closingCashCalculated)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'closingCashCalculated', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="font-semibold text-sky-800 bg-sky-50/40">
              <td className="py-2.5 px-4 font-sans">
                <EditableLabel
                  label={getLbl('cf', 'balanceSheetCash', 'Stated Cash in Statement of Financial Position')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'balanceSheetCash', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2.5 px-3 text-right">
                  <EditableCell
                    val={getVal('cf', 'balanceSheetCash', d.year, d.balanceSheetCash)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'balanceSheetCash', d.year, val)}
                  />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ----------------- INCOME STATEMENT TABLE ----------------- //
function IncomeStatementTable({
  incomeData,
  currency,
  companyProfile,
  isEditMode,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  incomeData: IncomeStatementYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  return (
    <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
      <StatementMasthead
        title="Statement of Comprehensive Income (Profit & Loss)"
        subtitle="Operating revenues, cost of sales, operating margins, finance costs, and net earnings"
        standard="IAS 1 & Section 76.1"
        companyProfile={companyProfile}
      />
      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs font-mono">
          <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
            <tr>
              <th className="py-3 px-4 font-sans font-semibold text-slate-900 tracking-wide">
                Financial Revenue / Expense Account
              </th>
              {incomeData.map(d => (
                <th key={d.year} className={`py-2.5 px-3 text-right ${d.isProjected ? 'bg-sky-50/90 text-sky-900' : 'text-slate-800'}`}>
                  <div className="flex flex-col items-end">
                    <span className="font-bold text-xs">{d.year}</span>
                    <span className={`text-[9px] font-sans uppercase tracking-wider font-semibold ${
                      d.isProjected ? 'text-sky-700' : 'text-slate-500'
                    }`}>
                      {d.isProjected ? 'Forecast' : 'Audited'}
                    </span>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200/70 tabular-nums">
            <tr className="font-semibold text-slate-900 hover:bg-slate-50/70">
              <td className="py-2.5 px-4 font-sans">
                <EditableLabel
                  label={getLbl('income', 'revenue', 'Sales Revenue')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'revenue', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2.5 px-3 text-right text-emerald-700 font-bold">
                  <EditableCell
                    val={getVal('income', 'revenue', d.year, d.revenue)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'revenue', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('income', 'costOfSales', 'Cost of Goods Sold / Cost of Sales')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'costOfSales', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-rose-700">
                  <EditableCell
                    val={getVal('income', 'costOfSales', d.year, d.costOfSales)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'costOfSales', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-slate-100/80 font-bold text-slate-900">
              <td className="py-2.5 px-4 font-sans">
                <EditableLabel
                  label={getLbl('income', 'grossProfit', 'Gross Profit')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'grossProfit', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2.5 px-3 text-right text-slate-900">
                  <EditableCell
                    val={getVal('income', 'grossProfit', d.year, d.grossProfit)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'grossProfit', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('income', 'operatingExpenses', 'Operating Expenses (Admin & Selling)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'operatingExpenses', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-rose-700">
                  <EditableCell
                    val={getVal('income', 'operatingExpenses', d.year, d.operatingExpenses)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'operatingExpenses', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('income', 'depreciationExpense', 'Depreciation Expense (Schedule of Fixed Assets)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'depreciationExpense', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-rose-700">
                  <EditableCell
                    val={getVal('income', 'depreciationExpense', d.year, d.depreciationExpense)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'depreciationExpense', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-sky-50/75 font-bold text-sky-800">
              <td className="py-2.5 px-4 font-sans">
                <EditableLabel
                  label={getLbl('income', 'operatingProfit', 'Operating Profit (EBIT)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'operatingProfit', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2.5 px-3 text-right">
                  <EditableCell
                    val={getVal('income', 'operatingProfit', d.year, d.operatingProfit)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'operatingProfit', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-amber-50/75 font-semibold">
              <td className="py-2 px-4 font-sans text-amber-900 pl-6">
                <EditableLabel
                  label={getLbl('income', 'financeCost', 'Finance Cost (Mandatory Direct Entry)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'financeCost', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-amber-800">
                  <EditableCell
                    val={getVal('income', 'financeCost', d.year, d.financeCost)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'financeCost', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="font-bold text-slate-900 hover:bg-slate-50/70">
              <td className="py-2.5 px-4 font-sans">
                <EditableLabel
                  label={getLbl('income', 'profitBeforeTax', 'Profit Before Tax (EBT)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'profitBeforeTax', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-2.5 px-3 text-right">
                  <EditableCell
                    val={getVal('income', 'profitBeforeTax', d.year, d.profitBeforeTax)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'profitBeforeTax', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-emerald-50/85 font-bold text-emerald-800 border-t-2 border-emerald-300 border-b-4 border-double border-emerald-600">
              <td className="py-3 px-4 font-sans text-emerald-800 text-sm">
                <EditableLabel
                  label={getLbl('income', 'netProfit', 'Net Profit After Tax')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('income', 'netProfit', lbl)}
                />
              </td>
              {incomeData.map(d => (
                <td key={d.year} className="py-3 px-3 text-right text-emerald-800 text-sm font-bold">
                  <EditableCell
                    val={getVal('income', 'netProfit', d.year, d.netProfit)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('income', 'netProfit', d.year, val)}
                  />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ----------------- BALANCE SHEET TABLE ----------------- //
function BalanceSheetTable({
  balanceData,
  currency,
  companyProfile,
  isEditMode,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  balanceData: BalanceSheetYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  return (
    <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
      <StatementMasthead
        title="Statement of Financial Position (Balance Sheet)"
        subtitle="Assets, liabilities, and shareholders equity solvency schedule"
        standard="IAS 1 & Section 76.2"
        companyProfile={companyProfile}
      />
      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs font-mono">
          <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
            <tr>
              <th className="py-3 px-4 font-sans font-semibold text-slate-900 tracking-wide">
                Financial Asset / Liability / Equity Item
              </th>
              {balanceData.map(d => (
                <th key={d.year} className={`py-2.5 px-3 text-right ${d.isProjected ? 'bg-sky-50/90 text-sky-900' : 'text-slate-800'}`}>
                  <div className="flex flex-col items-end">
                    <span className="font-bold text-xs">{d.year}</span>
                    <span className={`text-[9px] font-sans uppercase tracking-wider font-semibold ${
                      d.isProjected ? 'text-sky-700' : 'text-slate-500'
                    }`}>
                      {d.isProjected ? 'Forecast' : 'Audited'}
                    </span>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200/70 tabular-nums">
            <tr className="bg-sky-50/75 font-sans font-bold text-sky-800">
              <td colSpan={balanceData.length + 1} className="py-2 px-4 uppercase text-[11px]">Current Assets</td>
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-800 font-medium">
                <EditableLabel
                  label={getLbl('bs', 'cashAndEquivalents', 'Cash and Cash Equivalents')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'cashAndEquivalents', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-sky-700 font-medium">
                  <EditableCell
                    val={getVal('bs', 'cashAndEquivalents', d.year, d.cashAndEquivalents)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'cashAndEquivalents', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700">
                <EditableLabel
                  label={getLbl('bs', 'tradeReceivables', 'Trade Receivables')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'tradeReceivables', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('bs', 'tradeReceivables', d.year, d.tradeReceivables)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'tradeReceivables', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700">
                <EditableLabel
                  label={getLbl('bs', 'closingInventory', 'Closing Inventory')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'closingInventory', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('bs', 'closingInventory', d.year, d.closingInventory)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'closingInventory', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-slate-100/80 font-semibold text-slate-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('bs', 'totalCurrentAssets', 'Total Current Assets')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalCurrentAssets', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-sky-800">
                  <EditableCell
                    val={getVal('bs', 'totalCurrentAssets', d.year, d.totalCurrentAssets)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalCurrentAssets', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-sky-50/75 font-sans font-bold text-sky-800">
              <td colSpan={balanceData.length + 1} className="py-2 px-4 uppercase text-[11px]">Non-Current Assets</td>
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-800 font-medium">
                <EditableLabel
                  label={getLbl('bs', 'propertyPlantEquipment', 'Property, Plant and Equipment (NBV)')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'propertyPlantEquipment', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-emerald-700 font-medium">
                  <EditableCell
                    val={getVal('bs', 'propertyPlantEquipment', d.year, d.propertyPlantEquipment)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'propertyPlantEquipment', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-emerald-50/85 font-bold text-emerald-800 border-t-2 border-emerald-300 border-b-4 border-double border-emerald-600">
              <td className="py-3 px-4 font-sans text-sm">
                <EditableLabel
                  label={getLbl('bs', 'totalAssets', 'TOTAL ASSETS')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalAssets', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-3 px-3 text-right text-sm font-bold text-emerald-800">
                  <EditableCell
                    val={getVal('bs', 'totalAssets', d.year, d.totalAssets)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalAssets', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {/* Liabilities */}
            <tr className="bg-rose-50/75 font-sans font-bold text-rose-800">
              <td colSpan={balanceData.length + 1} className="py-2 px-4 uppercase text-[11px]">Liabilities</td>
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700">
                <EditableLabel
                  label={getLbl('bs', 'tradePayables', 'Trade Payables')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'tradePayables', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('bs', 'tradePayables', d.year, d.tradePayables)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'tradePayables', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700">
                <EditableLabel
                  label={getLbl('bs', 'longTermBorrowings', 'Long-Term Borrowings')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'longTermBorrowings', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('bs', 'longTermBorrowings', d.year, d.longTermBorrowings)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'longTermBorrowings', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {/* Equity */}
            <tr className="bg-teal-50/75 font-sans font-bold text-teal-800">
              <td colSpan={balanceData.length + 1} className="py-2 px-4 uppercase text-[11px]">Shareholders Equity</td>
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700">
                <EditableLabel
                  label={getLbl('bs', 'shareCapital', 'Share Capital')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'shareCapital', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('bs', 'shareCapital', d.year, d.shareCapital)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'shareCapital', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700">
                <EditableLabel
                  label={getLbl('bs', 'retainedEarnings', 'Retained Earnings')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'retainedEarnings', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('bs', 'retainedEarnings', d.year, d.retainedEarnings)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'retainedEarnings', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-emerald-50/85 font-bold text-emerald-800 border-t-2 border-emerald-300 border-b-4 border-double border-emerald-600">
              <td className="py-3 px-4 font-sans text-sm">
                <EditableLabel
                  label={getLbl('bs', 'totalLiabilitiesAndEquity', 'TOTAL LIABILITIES & EQUITY')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalLiabilitiesAndEquity', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-3 px-3 text-right text-sm font-bold text-emerald-800">
                  <EditableCell
                    val={getVal('bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalLiabilitiesAndEquity', d.year, val)}
                  />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ----------------- CHANGES IN EQUITY TABLE ----------------- //
function EquityStatementTable({
  equityData,
  currency,
  companyProfile,
  isEditMode,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  equityData: EquityStatementYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  return (
    <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
      <StatementMasthead
        title="Statement of Changes in Equity"
        subtitle="Share capital movements, retained earnings rollforward, and closing shareholders equity"
        standard="IAS 1 & Section 76.5"
        companyProfile={companyProfile}
      />
      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs font-mono">
          <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
            <tr>
              <th className="py-3 px-4 font-sans font-semibold text-slate-900 tracking-wide">
                Equity Classification / Component
              </th>
              {equityData.map(d => (
                <th key={d.year} className={`py-2.5 px-3 text-right ${d.isProjected ? 'bg-sky-50/90 text-sky-900' : 'text-slate-800'}`}>
                  <div className="flex flex-col items-end">
                    <span className="font-bold text-xs">{d.year}</span>
                    <span className={`text-[9px] font-sans uppercase tracking-wider font-semibold ${
                      d.isProjected ? 'text-sky-700' : 'text-slate-500'
                    }`}>
                      {d.isProjected ? 'Forecast' : 'Audited'}
                    </span>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200/70 tabular-nums">
            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-700 font-medium">
                <EditableLabel
                  label={getLbl('eq', 'openingShareCapital', 'Opening Share Capital')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'openingShareCapital', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('eq', 'openingShareCapital', d.year, d.openingShareCapital)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'openingShareCapital', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('eq', 'shareCapitalIssued', 'Share Capital Issued during Year')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'shareCapitalIssued', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-700">
                  <EditableCell
                    val={getVal('eq', 'shareCapitalIssued', d.year, d.shareCapitalIssued)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'shareCapitalIssued', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('eq', 'netProfitForYear', 'Net Profit for the Year')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'netProfitForYear', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-emerald-700 font-medium">
                  <EditableCell
                    val={getVal('eq', 'netProfitForYear', d.year, d.netProfitForYear)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'netProfitForYear', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="hover:bg-slate-50/70">
              <td className="py-2 px-4 font-sans text-slate-600 pl-6">
                <EditableLabel
                  label={getLbl('eq', 'dividendsPaid', 'Dividends Paid to Shareholders')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'dividendsPaid', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-rose-700">
                  <EditableCell
                    val={getVal('eq', 'dividendsPaid', d.year, d.dividendsPaid)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'dividendsPaid', d.year, val)}
                    isNegative
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-emerald-50/85 font-bold text-emerald-800 border-t-2 border-emerald-300 border-b-4 border-double border-emerald-600 text-sm">
              <td className="py-3 px-4 font-sans">
                <EditableLabel
                  label={getLbl('eq', 'totalClosingEquity', 'TOTAL CLOSING SHAREHOLDERS EQUITY')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'totalClosingEquity', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-3 px-3 text-right font-bold text-emerald-800">
                  <EditableCell
                    val={getVal('eq', 'totalClosingEquity', d.year, d.totalClosingEquity)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'totalClosingEquity', d.year, val)}
                  />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ----------------- SCHEDULE OF FIXED ASSETS TABLE ----------------- //
function FixedAssetsScheduleTable({
  fixedAssetsData,
  currency,
  companyProfile,
  isEditMode,
  onUpdateCategory,
}: {
  fixedAssetsData: ScheduleOfFixedAssetsYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  onUpdateCategory: (catIndex: number, field: string, val: number | string) => void;
}) {
  if (fixedAssetsData.length === 0) return null;
  const latest = fixedAssetsData[fixedAssetsData.length - 1];

  return (
    <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl overflow-hidden shadow-sm">
      <StatementMasthead
        title={`Schedule of Fixed Assets (Property, Plant & Equipment) — FY ${latest.year}`}
        subtitle={`Asset classes, capital additions, disposals, depreciation, and carrying net book values (${latest.isProjected ? 'Projected Period' : 'Historical Period'})`}
        standard="IAS 16 & Section 76.6"
        companyProfile={companyProfile}
      />

      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs font-mono">
          <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
            <tr>
              <th className="py-3 px-4 font-sans font-semibold text-slate-900">Asset Class</th>
              <th className="py-3 px-2 text-right font-semibold">Dep %</th>
              <th className="py-3 px-3 text-right font-semibold">Opening Cost</th>
              <th className="py-3 px-3 text-right text-emerald-700 font-semibold">Additions</th>
              <th className="py-3 px-3 text-right text-rose-700 font-semibold">Disposals</th>
              <th className="py-3 px-3 text-right text-slate-900 font-semibold">Closing Cost</th>
              <th className="py-3 px-3 text-right font-semibold">Opening Dep</th>
              <th className="py-3 px-3 text-right text-rose-700 font-semibold">Charge</th>
              <th className="py-3 px-3 text-right text-slate-900 font-semibold">Closing Dep</th>
              <th className="py-3 px-4 text-right text-emerald-700 font-bold">Net Book Value</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200/70 tabular-nums">
            {latest.categories.map((c, idx) => (
              <tr key={c.name} className="hover:bg-slate-50">
                <td className="py-2.5 px-4 font-sans text-slate-900 font-medium">
                  {isEditMode ? (
                    <input
                      type="text"
                      value={c.name}
                      onChange={(e) => onUpdateCategory(idx, 'name', e.target.value)}
                      className="bg-white border border-emerald-600 rounded px-1.5 py-0.5 text-xs text-slate-900 w-full focus:outline-none focus:ring-1 focus:ring-emerald-500"
                    />
                  ) : (
                    c.name
                  )}
                </td>
                <td className="py-2.5 px-2 text-right text-slate-600">{c.depRate}%</td>
                <td className="py-2.5 px-3 text-right text-slate-700">
                  <EditableCell
                    val={c.costOpening}
                    isEditMode={isEditMode}
                    onSave={(v) => onUpdateCategory(idx, 'costOpening', v)}
                  />
                </td>
                <td className="py-2.5 px-3 text-right text-emerald-700">
                  <EditableCell
                    val={c.additions}
                    isEditMode={isEditMode}
                    onSave={(v) => onUpdateCategory(idx, 'additions', v)}
                  />
                </td>
                <td className="py-2.5 px-3 text-right text-rose-700">
                  <EditableCell
                    val={c.disposals}
                    isEditMode={isEditMode}
                    onSave={(v) => onUpdateCategory(idx, 'disposals', v)}
                  />
                </td>
                <td className="py-2.5 px-3 text-right text-slate-900 font-semibold">
                  {c.costClosing.toLocaleString()}
                </td>
                <td className="py-2.5 px-3 text-right text-slate-600">
                  <EditableCell
                    val={c.depOpening}
                    isEditMode={isEditMode}
                    onSave={(v) => onUpdateCategory(idx, 'depOpening', v)}
                  />
                </td>
                <td className="py-2.5 px-3 text-right text-rose-700">
                  <EditableCell
                    val={c.depCharge}
                    isEditMode={isEditMode}
                    onSave={(v) => onUpdateCategory(idx, 'depCharge', v)}
                  />
                </td>
                <td className="py-2.5 px-3 text-right text-slate-700">
                  {c.depClosing.toLocaleString()}
                </td>
                <td className="py-2.5 px-4 text-right text-emerald-700 font-bold">
                  {c.netBookValue.toLocaleString()}
                </td>
              </tr>
            ))}

            {/* Total Carrying Net Book Value row with Double Underline */}
            <tr className="bg-emerald-50/85 font-bold text-emerald-800 border-t-2 border-emerald-300 border-b-4 border-double border-emerald-600 text-xs">
              <td className="py-3 px-4 font-sans font-bold text-slate-900">
                TOTAL PROPERTY, PLANT & EQUIPMENT
              </td>
              <td className="py-3 px-2 text-right text-slate-500">-</td>
              <td className="py-3 px-3 text-right text-slate-700 font-semibold">{latest.categories.reduce((s, c) => s + c.costOpening, 0).toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-emerald-700 font-bold">{latest.categories.reduce((s, c) => s + c.additions, 0).toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-rose-700">{latest.categories.reduce((s, c) => s + c.disposals, 0).toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-slate-900 font-bold">{latest.categories.reduce((s, c) => s + c.costClosing, 0).toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-slate-600">{latest.categories.reduce((s, c) => s + c.depOpening, 0).toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-rose-700">{latest.categories.reduce((s, c) => s + c.depCharge, 0).toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-slate-900">{latest.categories.reduce((s, c) => s + c.depClosing, 0).toLocaleString()}</td>
              <td className="py-3 px-4 text-right text-emerald-800 font-extrabold text-sm font-mono">
                {latest.categories.reduce((s, c) => s + c.netBookValue, 0).toLocaleString()}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ----------------- NOTES VIEWER ----------------- //
function NotesViewer({
  notes,
  currency,
  isEditMode,
  onUpdateNotes,
}: {
  notes: FinancialNoteItem[];
  currency: string;
  isEditMode: boolean;
  onUpdateNotes: (updated: FinancialNoteItem[]) => void;
}) {
  const handleNoteTextChange = (noteNumber: number, text: string) => {
    onUpdateNotes(
      notes.map(n => n.noteNumber === noteNumber ? { ...n, textDescription: text } : n)
    );
  };

  const handleNoteTitleChange = (noteNumber: number, title: string) => {
    onUpdateNotes(
      notes.map(n => n.noteNumber === noteNumber ? { ...n, title } : n)
    );
  };

  return (
    <div className="space-y-4">
      <div className="p-4 bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl flex items-center justify-between shadow-sm">
        <div>
          <h3 className="text-sm font-bold text-slate-900">
            Notes to the Financial Statements (Notes 1 to 12)
          </h3>
          <p className="text-xs text-slate-600">
            Mandatory disclosures, accounting policies, and cash flow reconciliations (All editable)
          </p>
        </div>
        <span className="text-xs text-emerald-700 font-semibold">12 Statutory Notes</span>
      </div>

      <div className="space-y-4">
        {notes.map(note => (
          <div key={note.noteNumber} className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl p-5 space-y-3 shadow-sm">
            <div className="flex items-center justify-between border-b border-slate-200 pb-2">
              <h4 className="text-xs font-bold text-slate-900 flex items-center gap-2 flex-1">
                <span className="w-5 h-5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-300 flex items-center justify-center text-[10px] shrink-0">
                  {note.noteNumber}
                </span>
                {isEditMode ? (
                  <input
                    type="text"
                    value={note.title}
                    onChange={(e) => handleNoteTitleChange(note.noteNumber, e.target.value)}
                    className="bg-white border border-emerald-600 rounded px-2 py-0.5 text-xs font-bold text-slate-900 flex-1 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                ) : (
                  <span>Note {note.noteNumber}: {note.title}</span>
                )}
              </h4>
              <span className="text-[10px] text-slate-600 bg-slate-100 border border-slate-200 px-2 py-0.5 rounded ml-2 shrink-0">
                {note.category}
              </span>
            </div>

            {isEditMode ? (
              <textarea
                value={note.textDescription}
                onChange={(e) => handleNoteTextChange(note.noteNumber, e.target.value)}
                rows={3}
                className="w-full bg-white border border-emerald-600 rounded p-2 text-xs text-slate-900 font-sans focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            ) : (
              <p className="text-xs text-slate-700 leading-relaxed font-sans">
                {note.textDescription}
              </p>
            )}

            {note.tableData && (
              <div className="overflow-x-auto pt-2">
                <table className="w-full text-left text-xs font-mono">
                  <thead className="bg-slate-100 text-slate-700 border-b border-slate-200 text-[11px]">
                    <tr>
                      {note.tableData.headers.map((h, i) => (
                        <th key={i} className={`py-2 px-3 font-semibold ${i === 0 ? 'text-left font-sans text-slate-900' : 'text-right'}`}>
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200/70 tabular-nums">
                    {note.tableData.rows.map((r, ri) => (
                      <tr key={ri} className="hover:bg-slate-50">
                        <td className="py-2 px-3 font-sans text-slate-700">{r.label}</td>
                        {r.values.map((v, vi) => (
                          <td key={vi} className="py-2 px-3 text-right text-slate-800">{v}</td>
                        ))}
                      </tr>
                    ))}
                    {note.tableData.totalRow && (
                      <tr className="bg-emerald-50/75 font-bold border-t border-emerald-200">
                        <td className="py-2 px-3 font-sans text-slate-900">{note.tableData.totalRow.label}</td>
                        {note.tableData.totalRow.values.map((v, vi) => (
                          <td key={vi} className="py-2 px-3 text-right text-emerald-800">{v}</td>
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
  );
}

// ----------------- DEBT SERVICING & REPAYMENT CAPACITY (DSCR) PANEL ----------------- //
function DebtServiceCoverageSummaryPanel({
  coverage,
  companyProfile,
  parsedSchedule,
  onJumpToRepaymentSchedule,
}: {
  coverage: DebtServiceCoverageYear[];
  companyProfile: CompanyProfile;
  parsedSchedule?: ParsedRepaymentSchedule | null;
  onJumpToRepaymentSchedule?: () => void;
}) {
  const activeRows = coverage.filter(r => r.totalDebtService > 0);
  if (activeRows.length === 0) return null;

  return (
    <div className="bg-white/92 backdrop-blur-md border border-emerald-300 rounded-2xl overflow-hidden shadow-sm">
      <div className="p-4 border-b border-emerald-200 bg-emerald-50/75 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <span className="text-[10px] font-mono uppercase tracking-widest text-emerald-800 font-bold">
            Implemented Repayment Schedule Coverage · {parsedSchedule?.metadata?.sourceFileName || companyProfile.companyName}
          </span>
          <h4 className="text-sm font-bold text-slate-900 mt-0.5">
            Debt Servicing &amp; Repayment Capacity Verification (DSCR)
          </h4>
        </div>
        {onJumpToRepaymentSchedule && (
          <button
            type="button"
            onClick={onJumpToRepaymentSchedule}
            className="px-3 py-1.5 bg-white hover:bg-slate-50 text-emerald-800 border border-emerald-300 rounded-lg text-xs font-semibold flex items-center gap-1.5 self-start sm:self-auto shadow-xs"
          >
            <CalendarCheck2 className="w-3.5 h-3.5" />
            <span>Inspect Schedule Ledger</span>
            <ArrowUpRight className="w-3 h-3" />
          </button>
        )}
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs font-mono">
          <thead className="bg-slate-100/95 text-slate-700 border-b border-slate-200 text-[11px]">
            <tr>
              <th className="py-2.5 px-4 font-sans font-semibold text-slate-900">Year</th>
              <th className="py-2.5 px-3 text-right font-semibold">EBITDA</th>
              <th className="py-2.5 px-3 text-right font-semibold">Pre-Fin OCF</th>
              <th className="py-2.5 px-3 text-right text-amber-800 font-semibold">Inst. Principal</th>
              <th className="py-2.5 px-3 text-right text-sky-800 font-semibold">Finance Cost (Profit)</th>
              <th className="py-2.5 px-3 text-right text-slate-900 font-semibold">Total Debt Service</th>
              <th className="py-2.5 px-3 text-right font-semibold">EBITDA DSCR</th>
              <th className="py-2.5 px-3 text-right font-semibold">Cash DSCR</th>
              <th className="py-2.5 px-3 text-right font-semibold">Closing Cash</th>
              <th className="py-2.5 px-4 text-right font-semibold">Status</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200/70 tabular-nums">
            {activeRows.map(row => (
              <tr key={row.year} className="hover:bg-slate-50">
                <td className="py-2 px-4 font-bold text-slate-900">
                  {row.year} <span className="text-[10px] text-slate-500 font-normal">{row.isProjected ? '(P)' : '(H)'}</span>
                </td>
                <td className="py-2 px-3 text-right text-slate-800">{fmtStatementNum(row.ebitda)}</td>
                <td className="py-2 px-3 text-right text-slate-800">{fmtStatementNum(row.operatingCashFlowBeforeFinance)}</td>
                <td className="py-2 px-3 text-right text-amber-800 font-semibold">{fmtStatementNum(row.scheduledPrincipal)}</td>
                <td className="py-2 px-3 text-right text-sky-800 font-semibold">{fmtStatementNum(row.scheduledFinanceCost)}</td>
                <td className="py-2 px-3 text-right text-slate-900 font-bold">{fmtStatementNum(row.totalDebtService)}</td>
                <td className={`py-2 px-3 text-right font-bold ${row.ebitdaDscr >= 1.25 ? 'text-emerald-700' : 'text-amber-800'}`}>
                  {row.ebitdaDscr.toFixed(2)}x
                </td>
                <td className={`py-2 px-3 text-right font-bold ${row.cashDscr >= 1.15 ? 'text-emerald-700' : 'text-amber-800'}`}>
                  {row.cashDscr.toFixed(2)}x
                </td>
                <td className="py-2 px-3 text-right text-emerald-700 font-medium">{fmtStatementNum(row.closingCash)}</td>
                <td className="py-2 px-4 text-right font-sans">
                  <span className={`inline-block text-[10px] font-bold px-2 py-0.5 rounded border ${
                    row.capacityStatus === 'STRONG'
                      ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                      : row.capacityStatus === 'ADEQUATE'
                      ? 'bg-sky-50 text-sky-800 border-sky-300'
                      : 'bg-amber-50 text-amber-900 border-amber-300'
                  }`}>
                    {row.capacityStatus}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ----------------- COMPLETE PACKAGE PREVIEW ----------------- //
function CompletePackagePreview({
  companyProfile,
  incomeData,
  balanceData,
  cashFlowData,
  equityData,
  fixedAssetsData,
  notes,
  debtServiceCoverage = [],
  onExportExcel,
  onExportWord,
  onExportPdf,
}: {
  companyProfile: CompanyProfile;
  incomeData: IncomeStatementYear[];
  balanceData: BalanceSheetYear[];
  cashFlowData: CashFlowStatementYear[];
  equityData: EquityStatementYear[];
  fixedAssetsData: ScheduleOfFixedAssetsYear[];
  notes: FinancialNoteItem[];
  debtServiceCoverage?: DebtServiceCoverageYear[];
  onExportExcel: () => void;
  onExportWord: () => void;
  onExportPdf: () => void;
}) {
  const hasDscr = debtServiceCoverage.some(d => d.totalDebtService > 0);
  return (
    <div className="bg-white/92 backdrop-blur-md border border-slate-200/90 rounded-2xl p-6 space-y-6 shadow-sm">
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <span className="text-[11px] font-semibold text-emerald-700 uppercase tracking-wider block">
            Section 76.11 Export Engine
          </span>
          <h3 className="text-xl font-bold text-slate-900 tracking-tight">
            Export Complete Financial Statements Package
          </h3>
          <p className="text-xs text-slate-600 mt-1">
            Generate one master consolidated report containing all five standard financial statements, fixed assets schedule, and twelve notes. All edited figures and custom labels are preserved in the export.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={onExportExcel}
            className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors flex items-center gap-2"
          >
            <FileSpreadsheet className="w-4 h-4" />
            Excel Workbook ({hasDscr ? '7' : '6'} Sheets)
          </button>
          <button
            onClick={onExportWord}
            className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors flex items-center gap-2"
          >
            <FileText className="w-4 h-4" />
            Word Document (.docx)
          </button>
          <button
            onClick={onExportPdf}
            className="px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors flex items-center gap-2"
          >
            <Download className="w-4 h-4" />
            PDF Package (.pdf)
          </button>
        </div>
      </div>

      {/* Package Contents Breakdown Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
        <div className="p-4 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1">
          <span className="text-slate-900 font-bold block">1. Income Statement</span>
          <p className="text-slate-600">Revenue, cost of sales, operating expenses, EBIT, finance cost, and net profit.</p>
        </div>
        <div className="p-4 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1">
          <span className="text-slate-900 font-bold block">2. Balance Sheet</span>
          <p className="text-slate-600">Statement of Financial Position: Current & Non-Current Assets, Liabilities, and Equity.</p>
        </div>
        <div className="p-4 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1">
          <span className="text-slate-900 font-bold block">3. Statement of Cash Flows</span>
          <p className="text-slate-600">Indirect method, operating, capex investing, and directly entered financing cash flows.</p>
        </div>
        <div className="p-4 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1">
          <span className="text-slate-900 font-bold block">4. Statement of Changes in Equity</span>
          <p className="text-slate-600">Share capital movement, retained earnings, comprehensive reserves, and dividends.</p>
        </div>
        <div className="p-4 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1">
          <span className="text-slate-900 font-bold block">5. Schedule of Fixed Assets</span>
          <p className="text-slate-600">Category-wise gross carrying amounts, additions, disposals, depreciation, and NBV.</p>
        </div>
        <div className="p-4 bg-slate-50/90 border border-slate-200 rounded-xl space-y-1">
          <span className="text-slate-900 font-bold block">6. Notes to Financials (1 to 12)</span>
          <p className="text-slate-600">Accounting policies, statutory disclosures, and cash flow balance sheet audit reconciliations.</p>
        </div>
        {hasDscr && (
          <div className="p-4 bg-emerald-50/75 border border-emerald-200 rounded-xl space-y-1 md:col-span-2 lg:col-span-3">
            <span className="text-emerald-800 font-bold block">7. Debt Servicing &amp; DSCR Schedule (Repayment Schedule Active)</span>
            <p className="text-slate-700">Annual Scheduled Principal, Scheduled Profit (Finance Cost), EBITDA DSCR, Cash DSCR, and Closing Cash sustainability metrics.</p>
          </div>
        )}
      </div>
    </div>
  );
}
