import React, { useState, useMemo } from 'react';
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
  CheckSquare,
  Square,
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
  FinancialYearEndOption,
} from '../types/financial';
import {
  exportIncomeStatementExcel,
  exportBalanceSheetExcel,
  exportCashFlowExcel,
  exportEquityStatementExcel,
  exportFixedAssetsExcel,
  exportNotesExcel,
  exportNotesWord,
  exportNotesPdf,
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
  resolveFACategoriesForExport,
  hasAnyNonZeroValue,
} from '../services/exportService';
import {
  getCustomItemYearValue,
  FIXED_ASSET_CATEGORY_CONFIG_IDS,
  generateNotesToFinancialStatements,
  normalizeFinancialYearEnd,
} from '../services/financialEngine';
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
  onUpdateStatementConfig?: (items: StatementConfigItem[]) => void;
  onResetStatementConfig?: () => void;
  assumptions?: ProjectionAssumptions;
  onUpdateAssumptions?: (updated: ProjectionAssumptions) => void;
  cellOverrides?: StatementCellOverrides;
  labelOverrides?: StatementLabelOverrides;
  onUpdateStatementValue?: (statement: string, field: string, year: number, val: number) => void;
  onUpdateStatementLabel?: (statement: string, field: string, newLabel: string) => void;
  onResetOverrides?: () => void;
  onUpdateNotes?: (updatedNotes: FinancialNoteItem[]) => void;
  onUpdateFixedAssetsCategory?: (year: number, catIndex: number, field: string, val: number) => void;
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
  onChangeFinancialYearEnd?: (fyEnd: FinancialYearEndOption) => void;
}

export const StatementsViewer: React.FC<StatementsViewerProps> = ({
  companyProfile,
  incomeData,
  balanceData,
  cashFlowData,
  equityData,
  fixedAssetsData,
  notes,
  statementConfig = [],
  onUpdateStatementConfig,
  onResetStatementConfig,
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
  onChangeFinancialYearEnd,
}) => {
  const [activeStatement, setActiveStatement] = useState<string>('cashFlow');
  const [yearViewMode, setYearViewMode] = useState<'all' | 'historical' | 'projected' | 'comparative5'>('all');
  const [isEditMode, setIsEditMode] = useState<boolean>(false);
  const [isCustomizeMode, setIsCustomizeMode] = useState<boolean>(false);
  const [internalFormatMode, setInternalFormatMode] = useState<StatementFormatMode>('section76');
  const [showYearSelector, setShowYearSelector] = useState<boolean>(false);

  const formatMode = controlledFormatMode ?? internalFormatMode;
  const setFormatMode = onChangeFormatMode ?? setInternalFormatMode;

  const isElemVisible = (id: string): boolean => {
    const found = statementConfig.find(c => c.id === id);
    return found ? found.isVisible : true;
  };

  const toggleConfigElement = (id: string) => {
    if (!onUpdateStatementConfig) return;
    onUpdateStatementConfig(
      statementConfig.map(c => (c.id === id ? { ...c, isVisible: !c.isVisible } : c))
    );
  };

  const activeConfigStatement: StatementConfigItem['statement'] | null =
    activeStatement === 'income' ||
    activeStatement === 'balanceSheet' ||
    activeStatement === 'cashFlow' ||
    activeStatement === 'equity' ||
    activeStatement === 'fixedAssets'
      ? activeStatement
      : null;

  const handleBulkSelectElements = (select: boolean, scope: 'active' | 'all' = 'active') => {
    if (!onUpdateStatementConfig) return;
    const targetStmt = scope === 'active' ? activeConfigStatement : null;
    onUpdateStatementConfig(
      statementConfig.map(c =>
        !targetStmt || c.statement === targetStmt ? { ...c, isVisible: select } : c
      )
    );
  };

  const getCustomItems = (stmt: StatementConfigItem['statement'], section?: string) =>
    statementConfig
      .filter(c => c.isCustom && c.statement === stmt && (!section || c.section === section))
      .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

  const exportSyncOptions = { statementConfig, labelOverrides, cellOverrides };
  const totalConfigCount = statementConfig.length;
  const selectedConfigCount = statementConfig.filter(c => c.isVisible).length;

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
  const activeYearSet = new Set(filteredIncome.map(d => d.year));
  const filteredDscr = debtServiceCoverage.filter(d => activeYearSet.has(d.year));

  const filteredNotes = useMemo(() => {
    if (yearViewMode === 'all' || filteredIncome.length === 0) {
      return notes;
    }
    const baseOpts = { statementConfig, labelOverrides, cellOverrides };
    const unfilteredDefaults = generateNotesToFinancialStatements(
      companyProfile,
      incomeData.map(y => y.year),
      incomeData,
      balanceData,
      cashFlowData,
      fixedAssetsData,
      baseOpts
    );
    const scoped = generateNotesToFinancialStatements(
      companyProfile,
      filteredIncome.map(y => y.year),
      filteredIncome,
      filteredBalance,
      filteredCashFlow,
      filteredFixedAssets,
      baseOpts
    );
    return scoped.map(sn => {
      const userNote = notes.find(n => n.noteNumber === sn.noteNumber);
      const defNote = unfilteredDefaults.find(n => n.noteNumber === sn.noteNumber);
      if (!userNote || !defNote) return sn;
      return {
        ...sn,
        title: userNote.title !== defNote.title ? userNote.title : sn.title,
        textDescription:
          userNote.textDescription !== defNote.textDescription
            ? userNote.textDescription
            : sn.textDescription,
      };
    });
  }, [
    yearViewMode,
    notes,
    companyProfile,
    incomeData,
    balanceData,
    cashFlowData,
    fixedAssetsData,
    filteredIncome,
    filteredBalance,
    filteredCashFlow,
    filteredFixedAssets,
    statementConfig,
    labelOverrides,
    cellOverrides,
  ]);

  // Helper to get effective cell value (checks cellOverrides first, falls back to original)
  const getVal = (stmt: string, field: string, year: number, fallback: number): number => {
    const key = `${stmt}_${field}`;
    if (cellOverrides[key] && cellOverrides[key][year] !== undefined) {
      return cellOverrides[key][year];
    }
    return fallback;
  };

  // Helper to get effective label (checks labelOverrides first, then statementConfig custom labels, falls back to default)
  const getLbl = (stmt: string, field: string, defaultLabel: string, configId?: string): string => {
    const key = `${stmt}_${field}`;
    if (labelOverrides[key] !== undefined) {
      return labelOverrides[key];
    }
    if (configId) {
      const cfg = statementConfig.find(c => c.id === configId);
      if (cfg && cfg.label) return cfg.label;
    }
    return defaultLabel;
  };

  // Individual Export Handlers (exports updated values, visible selections & labels!)
  const handleExportExcel = () => {
    switch (activeStatement) {
      case 'income':
        exportIncomeStatementExcel(companyProfile, filteredIncome, exportSyncOptions);
        break;
      case 'balanceSheet':
        exportBalanceSheetExcel(companyProfile, filteredBalance, exportSyncOptions);
        break;
      case 'cashFlow':
        exportCashFlowExcel(companyProfile, filteredCashFlow, exportSyncOptions);
        break;
      case 'equity':
        exportEquityStatementExcel(companyProfile, filteredEquity, exportSyncOptions);
        break;
      case 'fixedAssets':
        exportFixedAssetsExcel(companyProfile, filteredFixedAssets, exportSyncOptions);
        break;
      case 'notes':
        exportNotesExcel(companyProfile, filteredNotes);
        break;
      case 'complete':
        exportCompletePackageExcel(
          companyProfile,
          filteredIncome,
          filteredBalance,
          filteredCashFlow,
          filteredEquity,
          filteredFixedAssets,
          filteredNotes,
          undefined,
          isScheduleImplemented ? filteredDscr : undefined,
          exportSyncOptions
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
        filteredNotes,
        undefined,
        isScheduleImplemented ? filteredDscr : undefined,
        exportSyncOptions
      );
      return;
    }
    if (activeStatement === 'notes') {
      exportNotesWord(companyProfile, filteredNotes);
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
        filteredNotes,
        undefined,
        isScheduleImplemented ? filteredDscr : undefined,
        exportSyncOptions
      );
      return;
    }
    if (activeStatement === 'notes') {
      exportNotesPdf(companyProfile, filteredNotes);
      return;
    }
    const { title, headers, rows } = getTableDataForExport(activeStatement);
    exportDocumentPdf(companyProfile, title, headers, rows, `${activeStatement}_Statement.pdf`);
  };

  function getTableDataForExport(stmt: string): { title: string; headers: string[]; rows: (string | number)[][] } {
    const pushIfNonZero = (target: (string | number)[][], label: string, vals: number[]) => {
      if (hasAnyNonZeroValue(vals)) target.push([label, ...vals]);
    };
    switch (stmt) {
      case 'income': {
        const rows: (string | number)[][] = [];
        const stdIncomeSections = ['Revenue & Gross Margin', 'Operating Expenses', 'Non-Operating & Tax'];
        const visCustRev = getCustomItems('income', 'Revenue & Gross Margin').filter(c => c.isVisible);
        const visCustOpex = getCustomItems('income', 'Operating Expenses').filter(c => c.isVisible);
        const visCustNonOp = statementConfig
          .filter(
            c =>
              c.isCustom &&
              c.isVisible &&
              c.statement === 'income' &&
              (c.section === 'Non-Operating & Tax' || !stdIncomeSections.includes(c.section))
          )
          .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
        if (isElemVisible('inc_revenue')) {
          pushIfNonZero(
            rows,
            getLbl('income', 'revenue', 'Sales Revenue', 'inc_revenue'),
            filteredIncome.map(d => {
              const custSum = visCustRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
              return getVal('income', 'revenue', d.year, Math.round((d.revenue - custSum) * 100) / 100);
            })
          );
        }
        visCustRev.forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredIncome.map(d => getCustomItemYearValue(ci, d.year)));
        });
        if (isElemVisible('inc_costOfSales')) {
          pushIfNonZero(rows, getLbl('income', 'costOfSales', 'Cost of Sales', 'inc_costOfSales'), filteredIncome.map(d => -getVal('income', 'costOfSales', d.year, d.costOfSales)));
        }
        rows.push([getLbl('income', 'grossProfit', 'Gross Profit'), ...filteredIncome.map(d => getVal('income', 'grossProfit', d.year, d.grossProfit))]);
        if (isElemVisible('inc_operatingExpenses')) {
          pushIfNonZero(
            rows,
            getLbl('income', 'operatingExpenses', 'Operating Expenses', 'inc_operatingExpenses'),
            filteredIncome.map(d => {
              const custSum = visCustOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
              return -getVal('income', 'operatingExpenses', d.year, Math.round((d.operatingExpenses + custSum) * 100) / 100);
            })
          );
        }
        visCustOpex.forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredIncome.map(d => getCustomItemYearValue(ci, d.year)));
        });
        if (isElemVisible('inc_depreciationExpense')) {
          pushIfNonZero(rows, getLbl('income', 'depreciationExpense', 'Depreciation Expense', 'inc_depreciationExpense'), filteredIncome.map(d => -getVal('income', 'depreciationExpense', d.year, d.depreciationExpense)));
        }
        rows.push([getLbl('income', 'operatingProfit', 'Operating Profit (EBIT)'), ...filteredIncome.map(d => getVal('income', 'operatingProfit', d.year, d.operatingProfit))]);
        if (isElemVisible('inc_interestIncome')) {
          pushIfNonZero(
            rows,
            getLbl('income', 'interestIncome', 'Interest & Investment Income', 'inc_interestIncome'),
            filteredIncome.map(d => {
              const custSum = visCustNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
              return getVal('income', 'interestIncome', d.year, Math.round((d.interestIncome - custSum) * 100) / 100);
            })
          );
        }
        if (isElemVisible('inc_financeCost')) {
          pushIfNonZero(rows, getLbl('income', 'financeCost', 'Finance Cost', 'inc_financeCost'), filteredIncome.map(d => -getVal('income', 'financeCost', d.year, d.financeCost)));
        }
        visCustNonOp.forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredIncome.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('income', 'profitBeforeTax', 'Profit Before Tax (EBT)'), ...filteredIncome.map(d => getVal('income', 'profitBeforeTax', d.year, d.profitBeforeTax))]);
        if (isElemVisible('inc_taxExpense')) {
          pushIfNonZero(rows, getLbl('income', 'taxExpense', 'Income Tax Expense', 'inc_taxExpense'), filteredIncome.map(d => -getVal('income', 'taxExpense', d.year, d.taxExpense)));
        }
        rows.push([getLbl('income', 'netProfit', 'Net Profit After Tax'), ...filteredIncome.map(d => getVal('income', 'netProfit', d.year, d.netProfit))]);
        return {
          title: 'Statement of Comprehensive Income',
          headers: ['Metric', ...filteredIncome.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)],
          rows,
        };
      }
      case 'balanceSheet': {
        const rows: (string | number)[][] = [];
        if (isElemVisible('bs_cashAndEquivalents')) pushIfNonZero(rows, getLbl('bs', 'cashAndEquivalents', 'Cash and Cash Equivalents', 'bs_cashAndEquivalents'), filteredBalance.map(d => getVal('bs', 'cashAndEquivalents', d.year, d.cashAndEquivalents)));
        if (isElemVisible('bs_tradeReceivables')) pushIfNonZero(rows, getLbl('bs', 'tradeReceivables', 'Trade Receivables', 'bs_tradeReceivables'), filteredBalance.map(d => getVal('bs', 'tradeReceivables', d.year, d.tradeReceivables)));
        if (isElemVisible('bs_otherReceivables')) pushIfNonZero(rows, getLbl('bs', 'otherReceivables', 'Other Receivables & Prepayments', 'bs_otherReceivables'), filteredBalance.map(d => getVal('bs', 'otherReceivables', d.year, d.otherReceivables)));
        if (isElemVisible('bs_closingInventory')) pushIfNonZero(rows, getLbl('bs', 'closingInventory', 'Closing Inventory', 'bs_closingInventory'), filteredBalance.map(d => getVal('bs', 'closingInventory', d.year, d.closingInventory)));
        if (isElemVisible('bs_otherCurrentAssets')) pushIfNonZero(rows, getLbl('bs', 'otherCurrentAssets', 'Other Current Assets', 'bs_otherCurrentAssets'), filteredBalance.map(d => getVal('bs', 'otherCurrentAssets', d.year, d.otherCurrentAssets)));
        getCustomItems('balanceSheet', 'Current Assets').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredBalance.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('bs', 'totalCurrentAssets', 'Total Current Assets'), ...filteredBalance.map(d => getVal('bs', 'totalCurrentAssets', d.year, d.totalCurrentAssets))]);
        if (isElemVisible('bs_propertyPlantEquipment')) pushIfNonZero(rows, getLbl('bs', 'propertyPlantEquipment', 'Property, Plant & Equipment (NBV)', 'bs_propertyPlantEquipment'), filteredBalance.map(d => getVal('bs', 'propertyPlantEquipment', d.year, d.propertyPlantEquipment)));
        if (isElemVisible('bs_intangibleAssets')) pushIfNonZero(rows, getLbl('bs', 'intangibleAssets', 'Intangible Assets', 'bs_intangibleAssets'), filteredBalance.map(d => getVal('bs', 'intangibleAssets', d.year, d.intangibleAssets)));
        if (isElemVisible('bs_otherNonCurrentAssets')) pushIfNonZero(rows, getLbl('bs', 'otherNonCurrentAssets', 'Other Non-Current Assets', 'bs_otherNonCurrentAssets'), filteredBalance.map(d => getVal('bs', 'otherNonCurrentAssets', d.year, d.otherNonCurrentAssets)));
        getCustomItems('balanceSheet', 'Non-Current Assets').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredBalance.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('bs', 'totalNonCurrentAssets', 'Total Non-Current Assets'), ...filteredBalance.map(d => getVal('bs', 'totalNonCurrentAssets', d.year, d.totalNonCurrentAssets))]);
        rows.push([getLbl('bs', 'totalAssets', 'TOTAL ASSETS'), ...filteredBalance.map(d => getVal('bs', 'totalAssets', d.year, d.totalAssets))]);
        if (isElemVisible('bs_tradePayables')) pushIfNonZero(rows, getLbl('bs', 'tradePayables', 'Trade Payables', 'bs_tradePayables'), filteredBalance.map(d => getVal('bs', 'tradePayables', d.year, d.tradePayables)));
        if (isElemVisible('bs_otherPayables')) pushIfNonZero(rows, getLbl('bs', 'otherPayables', 'Other Payables & Accruals', 'bs_otherPayables'), filteredBalance.map(d => getVal('bs', 'otherPayables', d.year, d.otherPayables)));
        if (isElemVisible('bs_shortTermBorrowings')) pushIfNonZero(rows, getLbl('bs', 'shortTermBorrowings', 'Short-Term Borrowings', 'bs_shortTermBorrowings'), filteredBalance.map(d => getVal('bs', 'shortTermBorrowings', d.year, d.shortTermBorrowings)));
        if (isElemVisible('bs_currentTaxPayable')) pushIfNonZero(rows, getLbl('bs', 'currentTaxPayable', 'Current Tax Payable', 'bs_currentTaxPayable'), filteredBalance.map(d => getVal('bs', 'currentTaxPayable', d.year, d.currentTaxPayable)));
        getCustomItems('balanceSheet', 'Current Liabilities').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredBalance.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('bs', 'totalCurrentLiabilities', 'Total Current Liabilities'), ...filteredBalance.map(d => getVal('bs', 'totalCurrentLiabilities', d.year, d.totalCurrentLiabilities))]);
        if (isElemVisible('bs_longTermBorrowings')) pushIfNonZero(rows, getLbl('bs', 'longTermBorrowings', 'Long-Term Borrowings', 'bs_longTermBorrowings'), filteredBalance.map(d => getVal('bs', 'longTermBorrowings', d.year, d.longTermBorrowings)));
        if (isElemVisible('bs_otherNonCurrentLiabilities')) pushIfNonZero(rows, getLbl('bs', 'otherNonCurrentLiabilities', 'Other Non-Current Liabilities', 'bs_otherNonCurrentLiabilities'), filteredBalance.map(d => getVal('bs', 'otherNonCurrentLiabilities', d.year, d.otherNonCurrentLiabilities)));
        getCustomItems('balanceSheet', 'Non-Current Liabilities').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredBalance.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('bs', 'totalNonCurrentLiabilities', 'Total Non-Current Liabilities'), ...filteredBalance.map(d => getVal('bs', 'totalNonCurrentLiabilities', d.year, d.totalNonCurrentLiabilities))]);
        rows.push([getLbl('bs', 'totalLiabilities', 'TOTAL LIABILITIES'), ...filteredBalance.map(d => getVal('bs', 'totalLiabilities', d.year, d.totalLiabilities))]);
        if (isElemVisible('bs_shareCapital')) pushIfNonZero(rows, getLbl('bs', 'shareCapital', 'Share Capital', 'bs_shareCapital'), filteredBalance.map(d => getVal('bs', 'shareCapital', d.year, d.shareCapital)));
        if (isElemVisible('bs_retainedEarnings')) pushIfNonZero(rows, getLbl('bs', 'retainedEarnings', 'Retained Earnings', 'bs_retainedEarnings'), filteredBalance.map(d => getVal('bs', 'retainedEarnings', d.year, d.retainedEarnings)));
        if (isElemVisible('bs_reserves')) pushIfNonZero(rows, getLbl('bs', 'reserves', 'Reserves & Surplus', 'bs_reserves'), filteredBalance.map(d => getVal('bs', 'reserves', d.year, d.reserves)));
        getCustomItems('balanceSheet', 'Equity').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredBalance.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('bs', 'totalEquity', 'TOTAL EQUITY'), ...filteredBalance.map(d => getVal('bs', 'totalEquity', d.year, d.totalEquity))]);
        rows.push([getLbl('bs', 'totalLiabilitiesAndEquity', 'TOTAL LIABILITIES & EQUITY'), ...filteredBalance.map(d => getVal('bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity))]);
        rows.push([
          'Balance Check',
          ...filteredBalance.map(d => {
            const totA = getVal('bs', 'totalAssets', d.year, d.totalAssets);
            const totLE = getVal('bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity);
            return Math.round((totA - totLE) * 100) / 100;
          }),
        ]);
        return {
          title: 'Statement of Financial Position',
          headers: ['Line Item', ...filteredBalance.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)],
          rows,
        };
      }
      case 'equity': {
        const rows: (string | number)[][] = [];
        if (isElemVisible('eq_openingShareCapital')) pushIfNonZero(rows, getLbl('eq', 'openingShareCapital', 'Opening Share Capital', 'eq_openingShareCapital'), filteredEquity.map(d => getVal('eq', 'openingShareCapital', d.year, d.openingShareCapital)));
        if (isElemVisible('eq_shareCapitalIssued')) pushIfNonZero(rows, getLbl('eq', 'shareCapitalIssued', 'Share Capital Issued during Year', 'eq_shareCapitalIssued'), filteredEquity.map(d => getVal('eq', 'shareCapitalIssued', d.year, d.shareCapitalIssued)));
        rows.push([getLbl('eq', 'closingShareCapital', 'Closing Share Capital'), ...filteredEquity.map(d => getVal('eq', 'closingShareCapital', d.year, d.closingShareCapital))]);
        if (isElemVisible('eq_openingRetainedEarnings')) pushIfNonZero(rows, getLbl('eq', 'openingRetainedEarnings', 'Opening Retained Earnings', 'eq_openingRetainedEarnings'), filteredEquity.map(d => getVal('eq', 'openingRetainedEarnings', d.year, d.openingRetainedEarnings)));
        if (isElemVisible('eq_netProfitForYear')) pushIfNonZero(rows, getLbl('eq', 'netProfitForYear', 'Net Profit for the Year', 'eq_netProfitForYear'), filteredEquity.map(d => getVal('eq', 'netProfitForYear', d.year, d.netProfitForYear)));
        if (isElemVisible('eq_dividendsPaid')) pushIfNonZero(rows, getLbl('eq', 'dividendsPaid', 'Dividends Paid to Shareholders', 'eq_dividendsPaid'), filteredEquity.map(d => -getVal('eq', 'dividendsPaid', d.year, d.dividendsPaid)));
        rows.push([getLbl('eq', 'closingRetainedEarnings', 'Closing Retained Earnings'), ...filteredEquity.map(d => getVal('eq', 'closingRetainedEarnings', d.year, d.closingRetainedEarnings))]);
        const openResVals = filteredEquity.map(d => getVal('eq', 'openingReserves', d.year, d.openingReserves));
        const ociVals = filteredEquity.map(d => getVal('eq', 'otherComprehensiveIncome', d.year, d.otherComprehensiveIncome));
        const closeResVals = filteredEquity.map(d => getVal('eq', 'closingReserves', d.year, d.closingReserves));
        const showOpenRes = isElemVisible('eq_openingReserves') && hasAnyNonZeroValue(openResVals);
        const showOci = isElemVisible('eq_otherComprehensiveIncome') && hasAnyNonZeroValue(ociVals);
        if (showOpenRes || showOci || hasAnyNonZeroValue(closeResVals)) {
          if (showOpenRes) rows.push([getLbl('eq', 'openingReserves', 'Opening Reserves', 'eq_openingReserves'), ...openResVals]);
          if (showOci) rows.push([getLbl('eq', 'otherComprehensiveIncome', 'Other Comprehensive Income', 'eq_otherComprehensiveIncome'), ...ociVals]);
          rows.push([getLbl('eq', 'closingReserves', 'Closing Reserves'), ...closeResVals]);
        }
        getCustomItems('equity').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredEquity.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('eq', 'totalOpeningEquity', 'TOTAL OPENING EQUITY'), ...filteredEquity.map(d => getVal('eq', 'totalOpeningEquity', d.year, d.totalOpeningEquity))]);
        rows.push([getLbl('eq', 'totalClosingEquity', 'TOTAL CLOSING SHAREHOLDERS EQUITY'), ...filteredEquity.map(d => getVal('eq', 'totalClosingEquity', d.year, d.totalClosingEquity))]);
        return {
          title: 'Statement of Changes in Equity',
          headers: ['Equity Component', ...filteredEquity.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)],
          rows,
        };
      }
      case 'fixedAssets': {
        const latest = filteredFixedAssets[filteredFixedAssets.length - 1];
        if (!latest) {
          return { title: 'Schedule of Fixed Assets', headers: ['Asset Class'], rows: [] };
        }
        const cats = resolveFACategoriesForExport(latest, exportSyncOptions);
        return {
          title: `Schedule of Fixed Assets — FY ${latest.year}`,
          headers: ['Asset Class', 'Dep %', 'Opening Cost', 'Additions', 'Disposals', 'Closing Cost', 'Opening Dep', 'Charge', 'Dep. Disposals', 'Closing Dep', 'Net Book Value'],
          rows: [
            ...cats.map(c => [
              c.name,
              `${c.depRate}%`,
              c.costOpening,
              c.additions,
              c.disposals,
              c.costClosing,
              c.depOpening,
              c.depCharge,
              c.depDisposals,
              c.depClosing,
              c.netBookValue,
            ]),
            [
              'TOTAL PROPERTY, PLANT & EQUIPMENT',
              '-',
              latest.totalCostOpening,
              latest.totalAdditions,
              latest.totalDisposals,
              latest.totalCostClosing,
              latest.totalDepOpening,
              latest.totalDepCharge,
              latest.totalDepDisposals,
              latest.totalDepClosing,
              latest.totalNetBookValue,
            ],
          ],
        };
      }
      case 'notes': {
        return {
          title: 'Notes to the Financial Statements (Notes 1 to 12)',
          headers: ['Note #', 'Title', 'Category', 'Disclosure Summary'],
          rows: filteredNotes.map(n => [`Note ${n.noteNumber}`, n.title, n.category, n.textDescription]),
        };
      }
      case 'cashFlow':
      default: {
        const rows: (string | number)[][] = [];
        if (isElemVisible('cf_pbt')) pushIfNonZero(rows, getLbl('cf', 'profitBeforeTax', 'Net Profit Before Tax', 'cf_pbt'), filteredCashFlow.map(d => getVal('cf', 'profitBeforeTax', d.year, d.profitBeforeTax)));
        if (isElemVisible('cf_dep')) pushIfNonZero(rows, getLbl('cf', 'adjDepreciation', 'Adjustments for Depreciation', 'cf_dep'), filteredCashFlow.map(d => getVal('cf', 'adjDepreciation', d.year, d.adjDepreciation)));
        if (isElemVisible('cf_amort')) pushIfNonZero(rows, getLbl('cf', 'adjAmortization', 'Adjustments for Amortization', 'cf_amort'), filteredCashFlow.map(d => getVal('cf', 'adjAmortization', d.year, d.adjAmortization)));
        if (isElemVisible('cf_finance')) pushIfNonZero(rows, getLbl('cf', 'adjFinanceCost', 'Finance Cost Adjustment', 'cf_finance'), filteredCashFlow.map(d => getVal('cf', 'adjFinanceCost', d.year, d.adjFinanceCost)));
        if (isElemVisible('cf_interest_inc')) pushIfNonZero(rows, getLbl('cf', 'adjInterestIncome', 'Interest & Investment Income', 'cf_interest_inc'), filteredCashFlow.map(d => getVal('cf', 'adjInterestIncome', d.year, d.adjInterestIncome)));
        if (isElemVisible('cf_disposal_gain_loss')) pushIfNonZero(rows, getLbl('cf', 'adjLossGainDisposal', 'Loss / (Gain) on Disposal of Fixed Assets', 'cf_disposal_gain_loss'), filteredCashFlow.map(d => getVal('cf', 'adjLossGainDisposal', d.year, d.adjLossGainDisposal)));
        rows.push([getLbl('cf', 'operatingProfitBeforeWC', 'Operating Profit Before Working Capital Changes'), ...filteredCashFlow.map(d => getVal('cf', 'operatingProfitBeforeWC', d.year, d.operatingProfitBeforeWC))]);
        if (isElemVisible('cf_wc_inventory')) pushIfNonZero(rows, getLbl('cf', 'changeInInventory', '(Increase) / Decrease in Inventory', 'cf_wc_inventory'), filteredCashFlow.map(d => getVal('cf', 'changeInInventory', d.year, d.changeInInventory)));
        if (isElemVisible('cf_wc_receivables')) pushIfNonZero(rows, getLbl('cf', 'changeInReceivables', '(Increase) / Decrease in Trade Receivables', 'cf_wc_receivables'), filteredCashFlow.map(d => getVal('cf', 'changeInReceivables', d.year, d.changeInReceivables)));
        if (isElemVisible('cf_wc_other_rec')) pushIfNonZero(rows, getLbl('cf', 'changeInOtherReceivables', '(Increase) / Decrease in Other Receivables', 'cf_wc_other_rec'), filteredCashFlow.map(d => getVal('cf', 'changeInOtherReceivables', d.year, d.changeInOtherReceivables)));
        if (isElemVisible('cf_wc_other_ca')) pushIfNonZero(rows, getLbl('cf', 'changeInOtherCA', '(Increase) / Decrease in Other Current Assets', 'cf_wc_other_ca'), filteredCashFlow.map(d => getVal('cf', 'changeInOtherCA', d.year, d.changeInOtherCA ?? 0)));
        if (isElemVisible('cf_wc_payables')) pushIfNonZero(rows, getLbl('cf', 'changeInPayables', 'Increase / (Decrease) in Trade Payables', 'cf_wc_payables'), filteredCashFlow.map(d => getVal('cf', 'changeInPayables', d.year, d.changeInPayables)));
        if (isElemVisible('cf_wc_other_pay')) pushIfNonZero(rows, getLbl('cf', 'changeInOtherPayables', 'Increase / (Decrease) in Other Payables', 'cf_wc_other_pay'), filteredCashFlow.map(d => getVal('cf', 'changeInOtherPayables', d.year, d.changeInOtherPayables)));
        getCustomItems('cashFlow', 'Operating Activities').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredCashFlow.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('cf', 'cashGeneratedFromOperations', 'Cash Generated from Operations'), ...filteredCashFlow.map(d => getVal('cf', 'cashGeneratedFromOperations', d.year, d.cashGeneratedFromOperations))]);
        if (isElemVisible('cf_tax_paid')) pushIfNonZero(rows, getLbl('cf', 'incomeTaxPaid', 'Income Tax Paid', 'cf_tax_paid'), filteredCashFlow.map(d => getVal('cf', 'incomeTaxPaid', d.year, d.incomeTaxPaid)));
        if (isElemVisible('cf_finance_paid')) pushIfNonZero(rows, getLbl('cf', 'financeCostPaid', 'Finance Cost Paid', 'cf_finance_paid'), filteredCashFlow.map(d => getVal('cf', 'financeCostPaid', d.year, d.financeCostPaid)));
        rows.push([getLbl('cf', 'netCashFromOperating', 'NET CASH FROM OPERATING ACTIVITIES'), ...filteredCashFlow.map(d => getVal('cf', 'netCashFromOperating', d.year, d.netCashFromOperating))]);
        if (isElemVisible('cf_pur_ppe')) pushIfNonZero(rows, getLbl('cf', 'purchasePPE', 'Purchase of PPE (Capex)', 'cf_pur_ppe'), filteredCashFlow.map(d => getVal('cf', 'purchasePPE', d.year, d.purchasePPE)));
        if (isElemVisible('cf_sale_ppe')) pushIfNonZero(rows, getLbl('cf', 'salePPE', 'Proceeds from Sale of Fixed Assets', 'cf_sale_ppe'), filteredCashFlow.map(d => getVal('cf', 'salePPE', d.year, d.salePPE)));
        if (isElemVisible('cf_pur_intangible')) pushIfNonZero(rows, getLbl('cf', 'purchaseIntangibles', 'Purchase of Intangible Assets', 'cf_pur_intangible'), filteredCashFlow.map(d => getVal('cf', 'purchaseIntangibles', d.year, d.purchaseIntangibles)));
        if (isElemVisible('cf_interest_rec')) pushIfNonZero(rows, getLbl('cf', 'interestReceived', 'Interest Received', 'cf_interest_rec'), filteredCashFlow.map(d => getVal('cf', 'interestReceived', d.year, d.interestReceived)));
        if (isElemVisible('cf_dividends_rec')) pushIfNonZero(rows, getLbl('cf', 'dividendsReceived', 'Dividends Received', 'cf_dividends_rec'), filteredCashFlow.map(d => getVal('cf', 'dividendsReceived', d.year, d.dividendsReceived)));
        getCustomItems('cashFlow', 'Investing Activities').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredCashFlow.map(d => getCustomItemYearValue(ci, d.year)));
        });
        rows.push([getLbl('cf', 'netCashFromInvesting', 'NET CASH FROM / (USED IN) INVESTING'), ...filteredCashFlow.map(d => getVal('cf', 'netCashFromInvesting', d.year, d.netCashFromInvesting))]);
        if (isElemVisible('cf_shares_issue')) pushIfNonZero(rows, getLbl('cf', 'financingProceedsShares', 'Proceeds from Issue of Share Capital', 'cf_shares_issue'), filteredCashFlow.map(d => getVal('cf', 'financingProceedsShares', d.year, d.financingProceedsShares)));
        if (isElemVisible('cf_borrow_proceeds')) pushIfNonZero(rows, getLbl('cf', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing', 'cf_borrow_proceeds'), filteredCashFlow.map(d => getVal('cf', 'financingProceedsBorrowings', d.year, d.financingProceedsBorrowings)));
        // Mandatory inclusion of 'Repayment of Borrowings & Financing' overrides nil-value omission
        rows.push([
          getLbl('cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing', 'cf_borrow_repay'),
          ...filteredCashFlow.map(d =>
            isElemVisible('cf_borrow_repay')
              ? -getVal('cf', 'repaymentOfBorrowings', d.year, d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings))
              : 0
          ),
        ]);
        if (isElemVisible('cf_dividends_paid')) pushIfNonZero(rows, getLbl('cf', 'financingDividendsPaid', 'Dividends Paid to Shareholders', 'cf_dividends_paid'), filteredCashFlow.map(d => getVal('cf', 'financingDividendsPaid', d.year, d.financingDividendsPaid)));
        getCustomItems('cashFlow', 'Financing Activities').filter(c => c.isVisible).forEach(ci => {
          pushIfNonZero(rows, ci.label, filteredCashFlow.map(d => getCustomItemYearValue(ci, d.year)));
        });
        if (isElemVisible('cf_direct_financing')) pushIfNonZero(rows, getLbl('cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Financing Cash Flow', 'cf_direct_financing'), filteredCashFlow.map(d => getVal('cf', 'directlyEnteredFinancingCashFlow', d.year, d.directlyEnteredFinancingCashFlow)));
        rows.push([getLbl('cf', 'netCashFromFinancing', 'NET CASH FROM / (USED IN) FINANCING'), ...filteredCashFlow.map(d => getVal('cf', 'netCashFromFinancing', d.year, d.netCashFromFinancing))]);
        rows.push([getLbl('cf', 'netChangeInCash', 'Net Increase / (Decrease) in Cash'), ...filteredCashFlow.map(d => getVal('cf', 'netChangeInCash', d.year, d.netChangeInCash))]);
        rows.push([getLbl('cf', 'openingCash', 'Opening Cash and Cash Equivalents'), ...filteredCashFlow.map(d => getVal('cf', 'openingCash', d.year, d.openingCash))]);
        rows.push([getLbl('cf', 'closingCashCalculated', 'Closing Cash & Cash Equivalents (Calculated)'), ...filteredCashFlow.map(d => getVal('cf', 'closingCashCalculated', d.year, d.closingCashCalculated))]);
        rows.push([getLbl('cf', 'balanceSheetCash', 'Stated Cash in Balance Sheet'), ...filteredCashFlow.map(d => getVal('cf', 'balanceSheetCash', d.year, d.balanceSheetCash))]);
        rows.push([
          'Reconciliation Difference',
          ...filteredCashFlow.map(d => {
            const closeVal = getVal('cf', 'closingCashCalculated', d.year, d.closingCashCalculated);
            const bsCashVal = getVal('cf', 'balanceSheetCash', d.year, d.balanceSheetCash);
            return Math.round((closeVal - bsCashVal) * 100) / 100;
          }),
        ]);
        rows.push([
          'Audit Reconciliation Status',
          ...filteredCashFlow.map(d => {
            const closeVal = getVal('cf', 'closingCashCalculated', d.year, d.closingCashCalculated);
            const bsCashVal = getVal('cf', 'balanceSheetCash', d.year, d.balanceSheetCash);
            return Math.abs(closeVal - bsCashVal) < 0.01 ? 'RECONCILED' : 'VARIANCE_IDENTIFIED';
          }),
        ]);
        return {
          title: 'Statement of Cash Flows (Indirect Method)',
          headers: ['Cash Flow Item', ...filteredCashFlow.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)],
          rows,
        };
      }
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
    if (fmt === 'xlsx') exportPreviousFinancialsExcel(companyProfile, pair.data, scoped, exportSyncOptions);
    else if (fmt === 'docx') void exportPreviousFinancialsWord(companyProfile, pair.data, scoped, exportSyncOptions);
    else exportPreviousFinancialsPdf(companyProfile, pair.data, scoped, exportSyncOptions);
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
              {companyProfile.companyName?.trim() || 'Untitled Entity'} · {companyProfile.accountingFramework} · {companyProfile.reportingCurrency} in {companyProfile.reportingUnit} · FY End: {normalizeFinancialYearEnd(companyProfile.financialYearEnd)}
            </p>
          </div>

          {/* Global Controls: Year Selector, FY End, Edit Mode Toggle & Export Strip */}
          <div className="flex flex-wrap items-center gap-2">
            {onChangeFinancialYearEnd && (
              <div
                className="flex items-center bg-slate-100 border border-slate-300 rounded-lg p-0.5 text-[11px] font-semibold shrink-0"
                title="Switch Financial Year-End between June 30 and December 31"
              >
                {(['June 30', 'December 31'] as FinancialYearEndOption[]).map(opt => {
                  const active = normalizeFinancialYearEnd(companyProfile.financialYearEnd) === opt;
                  return (
                    <button
                      key={opt}
                      type="button"
                      onClick={() => onChangeFinancialYearEnd(opt)}
                      className={`px-2 py-1 rounded-md transition-colors whitespace-nowrap ${
                        active ? 'bg-emerald-600 text-white shadow-xs' : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {opt}
                    </button>
                  );
                })}
              </div>
            )}
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

            {/* Customize Statement Elements Button */}
            {formatMode === 'section76' && onUpdateStatementConfig && (
              <button
                type="button"
                onClick={() => setIsCustomizeMode(!isCustomizeMode)}
                className={`px-3 py-1.5 text-xs font-semibold rounded-lg border transition-all flex items-center gap-1.5 shadow-sm whitespace-nowrap shrink-0 ${
                  isCustomizeMode
                    ? 'bg-indigo-600 text-white border-indigo-600 ring-2 ring-indigo-400/30'
                    : 'bg-white hover:bg-indigo-50 text-indigo-700 border-slate-300 hover:border-indigo-400'
                }`}
                title="Select or deselect individual financial statement elements and auto-recalculate"
              >
                <CheckSquare className="w-3.5 h-3.5" />
                <span>
                  {isCustomizeMode
                    ? `Done Selecting (${selectedConfigCount}/${totalConfigCount})`
                    : `Customize Elements (${selectedConfigCount}/${totalConfigCount})`}
                </span>
              </button>
            )}

            {isCustomizeMode && formatMode === 'section76' && onUpdateStatementConfig && (
              <>
                <button
                  type="button"
                  onClick={() => handleBulkSelectElements(true, 'active')}
                  className="px-2.5 py-1.5 text-xs font-medium text-emerald-800 hover:text-emerald-950 bg-white hover:bg-emerald-50 rounded-lg border border-slate-300 hover:border-emerald-300 transition-colors flex items-center gap-1 whitespace-nowrap shrink-0"
                  title={activeConfigStatement ? 'Select all elements in this statement' : 'Select all statement elements'}
                >
                  <CheckSquare className="w-3.5 h-3.5 text-emerald-600" />
                  Select All
                </button>
                <button
                  type="button"
                  onClick={() => handleBulkSelectElements(false, 'active')}
                  className="px-2.5 py-1.5 text-xs font-medium text-rose-700 hover:text-rose-900 bg-white hover:bg-rose-50 rounded-lg border border-slate-300 hover:border-rose-300 transition-colors flex items-center gap-1 whitespace-nowrap shrink-0"
                  title={activeConfigStatement ? 'Deselect all elements in this statement' : 'Deselect all statement elements'}
                >
                  <Square className="w-3.5 h-3.5 text-rose-500" />
                  Deselect All
                </button>
              </>
            )}

            {isCustomizeMode && formatMode === 'section76' && onResetStatementConfig && (
              <button
                type="button"
                onClick={onResetStatementConfig}
                className="px-2.5 py-1.5 text-xs font-medium text-indigo-700 hover:text-indigo-900 bg-white hover:bg-indigo-50 rounded-lg border border-slate-300 transition-colors flex items-center gap-1 whitespace-nowrap shrink-0"
                title="Restore all default financial statement elements and labels"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                Reset Defaults
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
                  onClick={() => exportAllPreviousComparativePairsExcel(companyProfile, previousComparativePairs, pairData => (getPairScopedNotes ? getPairScopedNotes(pairData) : notes), exportSyncOptions)}
                  className="px-2 py-1 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-semibold transition-colors whitespace-nowrap shrink-0"
                  title="Download all Previous Comparative sets as Excel workbooks"
                >
                  All .XLSX
                </button>
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsWord(companyProfile, previousComparativePairs, pairData => (getPairScopedNotes ? getPairScopedNotes(pairData) : notes), exportSyncOptions)}
                  className="px-2 py-1 rounded-md bg-blue-600 hover:bg-blue-700 text-white text-[10px] font-semibold transition-colors whitespace-nowrap shrink-0"
                  title="Download all Previous Comparative sets as Word documents"
                >
                  All .DOCX
                </button>
                <button
                  type="button"
                  onClick={() => exportAllPreviousComparativePairsPdf(companyProfile, previousComparativePairs, pairData => (getPairScopedNotes ? getPairScopedNotes(pairData) : notes), exportSyncOptions)}
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
                onClick={() => exportProjectedFinancialsExcel(companyProfile, projectedFormatData, exportSyncOptions)}
                className="px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-semibold flex items-center gap-1 shadow-xs transition-colors whitespace-nowrap shrink-0"
              >
                <FileSpreadsheet className="w-3 h-3" />
                .XLSX
              </button>
              <button
                type="button"
                onClick={() => exportProjectedFinancialsWord(companyProfile, projectedFormatData, exportSyncOptions)}
                className="px-2.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-semibold flex items-center gap-1 shadow-xs transition-colors whitespace-nowrap shrink-0"
              >
                <FileText className="w-3 h-3" />
                .DOCX
              </button>
              <button
                type="button"
                onClick={() => exportProjectedFinancialsPdf(companyProfile, projectedFormatData, exportSyncOptions)}
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
          statementConfig={statementConfig}
          onUpdateStatementConfig={onUpdateStatementConfig}
          onResetStatementConfig={onResetStatementConfig}
          labelOverrides={labelOverrides}
          cellOverrides={cellOverrides}
        />
      )}

      {formatMode === 'projected_format' && (
        <ProjectedFormatSuiteView
          companyProfile={companyProfile}
          projData={projectedFormatData}
          isEditMode={isEditMode}
          onUpdatePrimaryDriver={onUpdatePrimaryDriver}
          statementConfig={statementConfig}
          onUpdateStatementConfig={onUpdateStatementConfig}
          onResetStatementConfig={onResetStatementConfig}
          labelOverrides={labelOverrides}
          cellOverrides={cellOverrides}
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

          {/* Element Selection Info Banner */}
          {isCustomizeMode && (
            <div className="p-3.5 bg-indigo-50/95 border border-indigo-300 rounded-xl text-xs text-indigo-950 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-xs">
              <div className="flex items-center gap-2">
                <CheckSquare className="w-4 h-4 text-indigo-600 shrink-0" />
                <span>
                  <strong>Element Selection Mode ({selectedConfigCount}/{totalConfigCount} active):</strong> Check or uncheck any statement row below (such as <em>Repayment of Borrowings &amp; Financing</em>). All subtotals, linked statements, ratios, DSCR, previews, and exports recalculate automatically.
                </span>
              </div>
              <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                <button
                  type="button"
                  onClick={() => handleBulkSelectElements(true, 'all')}
                  className="px-2.5 py-1 bg-white hover:bg-emerald-50 text-emerald-800 border border-emerald-300 rounded font-medium text-[11px]"
                >
                  Select All (All Statements)
                </button>
                <button
                  type="button"
                  onClick={() => handleBulkSelectElements(false, 'all')}
                  className="px-2.5 py-1 bg-white hover:bg-rose-50 text-rose-700 border border-rose-300 rounded font-medium text-[11px]"
                >
                  Deselect All (All Statements)
                </button>
                <button
                  type="button"
                  onClick={() => setIsCustomizeMode(false)}
                  className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded font-medium text-[11px]"
                >
                  Done Selecting
                </button>
              </div>
            </div>
          )}

          {/* Render Statement Body */}
          {activeStatement === 'cashFlow' && (
            <div className="space-y-6">
              <StatementOfCashFlowsTable
                cashFlowData={filteredCashFlow}
                currency={companyProfile.reportingCurrency}
                companyProfile={companyProfile}
                isEditMode={isEditMode}
                isCustomizeMode={isCustomizeMode}
                isElemVisible={isElemVisible}
                onToggleElem={toggleConfigElement}
                customItems={getCustomItems('cashFlow')}
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
              isCustomizeMode={isCustomizeMode}
              isElemVisible={isElemVisible}
              onToggleElem={toggleConfigElement}
              customItems={getCustomItems('income')}
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
              isCustomizeMode={isCustomizeMode}
              isElemVisible={isElemVisible}
              onToggleElem={toggleConfigElement}
              customItems={getCustomItems('balanceSheet')}
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
              isCustomizeMode={isCustomizeMode}
              isElemVisible={isElemVisible}
              onToggleElem={toggleConfigElement}
              customItems={getCustomItems('equity')}
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
              isCustomizeMode={isCustomizeMode}
              isElemVisible={isElemVisible}
              onToggleElem={toggleConfigElement}
              customItems={getCustomItems('fixedAssets')}
              getLbl={getLbl}
              onUpdateCategory={onUpdateFixedAssetsCategory}
              onUpdateValue={onUpdateStatementValue}
              onUpdateLabel={onUpdateStatementLabel}
            />
          )}

          {activeStatement === 'notes' && (
            <NotesViewer
              notes={filteredNotes}
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
              notes={filteredNotes}
              debtServiceCoverage={isScheduleImplemented ? filteredDscr : []}
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
            {companyProfile?.companyName?.trim() || 'Untitled Entity'}
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
  isCustomizeMode,
  isElemVisible,
  onToggleElem,
  customItems,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  cashFlowData: CashFlowStatementYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  isCustomizeMode: boolean;
  isElemVisible: (id: string) => boolean;
  onToggleElem: (id: string) => void;
  customItems: StatementConfigItem[];
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string, configId?: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  const renderToggle = (configId: string) =>
    isCustomizeMode ? (
      <input
        type="checkbox"
        checked={isElemVisible(configId)}
        onChange={() => onToggleElem(configId)}
        title="Include or exclude this element from calculation and exports"
        className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
      />
    ) : null;

  const renderStandardRow = (
    configId: string,
    field: string,
    defaultLabel: string,
    getFallback: (d: CashFlowStatementYear) => number,
    opts?: {
      indent?: boolean;
      isNegative?: boolean;
      highlightClass?: string;
      cellTextClass?: string;
      badge?: string;
    }
  ) => {
    const visible = isElemVisible(configId);
    if (!visible && !isCustomizeMode) return null;
    const indentClass = opts?.indent ? 'pl-6' : '';
    const rowBg = !visible
      ? 'opacity-45 bg-slate-50/90'
      : opts?.highlightClass || 'hover:bg-slate-50/70';

    return (
      <tr key={configId} className={rowBg}>
        <td className={`py-2 px-4 font-sans ${indentClass} ${opts?.badge ? 'flex items-center gap-1.5 text-rose-800' : 'text-slate-700'}`}>
          <div className="flex items-center gap-2">
            {renderToggle(configId)}
            <EditableLabel
              label={getLbl('cf', field, defaultLabel, configId)}
              isEditMode={isEditMode}
              onSave={(lbl) => onUpdateLabel('cf', field, lbl)}
            />
            {opts?.badge && (
              <span className="text-[10px] bg-rose-100 text-rose-800 border border-rose-300 px-1.5 py-0.2 rounded font-semibold shrink-0">
                {opts.badge}
              </span>
            )}
          </div>
        </td>
        {cashFlowData.map(d => (
          <td key={d.year} className={`py-2 px-3 text-right ${opts?.cellTextClass || 'text-slate-700'}`}>
            <EditableCell
              val={getVal('cf', field, d.year, getFallback(d))}
              isEditMode={isEditMode}
              onSave={(val) => onUpdateValue('cf', field, d.year, val)}
              isNegative={opts?.isNegative}
            />
          </td>
        ))}
      </tr>
    );
  };

  const renderCustomRows = (section: string) =>
    customItems
      .filter(ci => ci.section === section && (ci.isVisible || isCustomizeMode))
      .map(ci => (
        <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50/90' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
          <td className="py-2 px-4 font-sans text-indigo-900 pl-6">
            <div className="flex items-center gap-2">
              {isCustomizeMode && (
                <input
                  type="checkbox"
                  checked={ci.isVisible}
                  onChange={() => onToggleElem(ci.id)}
                  className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
                />
              )}
              <EditableLabel
                label={ci.label}
                isEditMode={isEditMode}
                onSave={(lbl) => onUpdateLabel('custom', ci.id, lbl)}
              />
              <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1.5 py-0.5 rounded shrink-0">
                Custom
              </span>
            </div>
          </td>
          {cashFlowData.map(d => (
            <td key={d.year} className={`py-2 px-3 text-right ${ci.sign === 'negative' ? 'text-rose-700' : 'text-indigo-800'}`}>
              <EditableCell
                val={getCustomItemYearValue(ci, d.year)}
                isEditMode={isEditMode}
                onSave={(val) => onUpdateValue('custom', ci.id, d.year, val)}
                isNegative={ci.sign === 'negative'}
              />
            </td>
          ))}
        </tr>
      ));

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

            {renderStandardRow('cf_pbt', 'profitBeforeTax', 'Net Profit Before Tax', d => d.profitBeforeTax, { cellTextClass: 'text-slate-800 font-medium' })}
            {renderStandardRow('cf_dep', 'adjDepreciation', 'Adjustments for Depreciation', d => d.adjDepreciation, { indent: true })}
            {renderStandardRow('cf_amort', 'adjAmortization', 'Adjustments for Amortization', d => d.adjAmortization, { indent: true })}
            {renderStandardRow('cf_finance', 'adjFinanceCost', 'Finance Cost Adjustment', d => d.adjFinanceCost, { indent: true })}
            {renderStandardRow('cf_interest_inc', 'adjInterestIncome', 'Interest & Investment Income', d => d.adjInterestIncome, { indent: true })}
            {renderStandardRow('cf_disposal_gain_loss', 'adjLossGainDisposal', 'Loss / (Gain) on Disposal of Fixed Assets', d => d.adjLossGainDisposal, { indent: true })}

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

            {renderStandardRow('cf_wc_inventory', 'changeInInventory', '(Increase) / Decrease in Inventory', d => d.changeInInventory, { indent: true })}
            {renderStandardRow('cf_wc_receivables', 'changeInReceivables', '(Increase) / Decrease in Trade Receivables', d => d.changeInReceivables, { indent: true })}
            {renderStandardRow('cf_wc_other_rec', 'changeInOtherReceivables', '(Increase) / Decrease in Other Receivables', d => d.changeInOtherReceivables, { indent: true })}
            {renderStandardRow('cf_wc_other_ca', 'changeInOtherCA', '(Increase) / Decrease in Other Current Assets', d => d.changeInOtherCA ?? 0, { indent: true })}
            {renderStandardRow('cf_wc_payables', 'changeInPayables', 'Increase / (Decrease) in Trade Payables', d => d.changeInPayables, { indent: true })}
            {renderStandardRow('cf_wc_other_pay', 'changeInOtherPayables', 'Increase / (Decrease) in Other Payables', d => d.changeInOtherPayables, { indent: true })}
            {renderCustomRows('Operating Activities')}

            <tr className="bg-slate-50/90 font-semibold text-slate-800">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('cf', 'cashGeneratedFromOperations', 'Cash Generated from Operations')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('cf', 'cashGeneratedFromOperations', lbl)}
                />
              </td>
              {cashFlowData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right text-slate-900 font-semibold">
                  <EditableCell
                    val={getVal('cf', 'cashGeneratedFromOperations', d.year, d.cashGeneratedFromOperations)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('cf', 'cashGeneratedFromOperations', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {renderStandardRow('cf_tax_paid', 'incomeTaxPaid', 'Income Tax Paid', d => d.incomeTaxPaid, { indent: true, isNegative: true, cellTextClass: 'text-rose-700' })}
            {renderStandardRow('cf_finance_paid', 'financeCostPaid', 'Finance Cost Paid', d => d.financeCostPaid, { indent: true, isNegative: true, cellTextClass: 'text-rose-700' })}

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

            {renderStandardRow('cf_pur_ppe', 'purchasePPE', 'Purchase of Property, Plant & Equipment (Capex)', d => d.purchasePPE, { indent: true, isNegative: true, cellTextClass: 'text-rose-700' })}
            {renderStandardRow('cf_sale_ppe', 'salePPE', 'Proceeds from Sale of Fixed Assets', d => d.salePPE, { indent: true })}
            {renderStandardRow('cf_pur_intangible', 'purchaseIntangibles', 'Purchase of Intangible Assets', d => d.purchaseIntangibles, { indent: true, isNegative: true, cellTextClass: 'text-rose-700' })}
            {renderStandardRow('cf_interest_rec', 'interestReceived', 'Interest Received', d => d.interestReceived, { indent: true })}
            {renderStandardRow('cf_dividends_rec', 'dividendsReceived', 'Dividends Received', d => d.dividendsReceived, { indent: true })}
            {renderCustomRows('Investing Activities')}

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
              <td colSpan={cashFlowData.length + 1} className="py-2.5 px-4 text-xs tracking-wider uppercase">
                <div className="flex items-center justify-between">
                  <span>C. Cash Flows from Financing Activities</span>
                  <span className="text-[10px] font-semibold normal-case text-amber-900 bg-amber-100 px-2 py-0.5 rounded border border-amber-300">
                    Mandatory Direct User Input
                  </span>
                </div>
              </td>
            </tr>

            {renderStandardRow('cf_shares_issue', 'financingProceedsShares', 'Proceeds from Issue of Share Capital', d => d.financingProceedsShares, { indent: true, cellTextClass: 'text-slate-800' })}
            {renderStandardRow('cf_borrow_proceeds', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing', d => d.financingProceedsBorrowings, { indent: true, cellTextClass: 'text-slate-800' })}
            {renderStandardRow('cf_borrow_repay', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing (Outflow) *', d => d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings), { indent: true, isNegative: true, highlightClass: 'bg-rose-50/65 font-medium', cellTextClass: 'text-rose-700 font-bold', badge: 'User Input' })}
            {renderStandardRow('cf_dividends_paid', 'financingDividendsPaid', 'Dividends Paid to Shareholders', d => d.financingDividendsPaid, { indent: true, isNegative: true, cellTextClass: 'text-rose-700' })}
            {renderCustomRows('Financing Activities')}

            {(isElemVisible('cf_direct_financing') || isCustomizeMode) && (
              <tr className={!isElemVisible('cf_direct_financing') ? 'opacity-45 bg-slate-50/90' : 'bg-amber-50/75 font-semibold border-t border-amber-200'}>
                <td className="py-2 px-4 font-sans text-amber-900">
                  <div className="flex items-center gap-2">
                    {renderToggle('cf_direct_financing')}
                    <EditableLabel
                      label={getLbl('cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Cash Flow from Financing Activities *', 'cf_direct_financing')}
                      isEditMode={isEditMode}
                      onSave={(lbl) => onUpdateLabel('cf', 'directlyEnteredFinancingCashFlow', lbl)}
                    />
                  </div>
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
            )}

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
  isCustomizeMode,
  isElemVisible,
  onToggleElem,
  customItems,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  incomeData: IncomeStatementYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  isCustomizeMode: boolean;
  isElemVisible: (id: string) => boolean;
  onToggleElem: (id: string) => void;
  customItems: StatementConfigItem[];
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string, configId?: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  const renderToggle = (configId: string) =>
    isCustomizeMode ? (
      <input
        type="checkbox"
        checked={isElemVisible(configId)}
        onChange={() => onToggleElem(configId)}
        title="Include or exclude this element from calculation and exports"
        className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
      />
    ) : null;

  const renderIncomeRow = (
    configId: string,
    field: string,
    defaultLabel: string,
    getFallback: (d: IncomeStatementYear) => number,
    opts?: { indent?: boolean; isNegative?: boolean; rowClass?: string; cellClass?: string }
  ) => {
    const visible = isElemVisible(configId);
    if (!visible && !isCustomizeMode) return null;
    return (
      <tr key={configId} className={!visible ? 'opacity-45 bg-slate-50/90' : opts?.rowClass || 'hover:bg-slate-50/70'}>
        <td className={`py-2 px-4 font-sans ${opts?.indent ? 'pl-6 text-slate-600' : 'text-slate-900'}`}>
          <div className="flex items-center gap-2">
            {renderToggle(configId)}
            <EditableLabel
              label={getLbl('income', field, defaultLabel, configId)}
              isEditMode={isEditMode}
              onSave={(lbl) => onUpdateLabel('income', field, lbl)}
            />
          </div>
        </td>
        {incomeData.map(d => (
          <td key={d.year} className={`py-2 px-3 text-right ${opts?.cellClass || 'text-rose-700'}`}>
            <EditableCell
              val={getVal('income', field, d.year, getFallback(d))}
              isEditMode={isEditMode}
              onSave={(val) => onUpdateValue('income', field, d.year, val)}
              isNegative={opts?.isNegative}
            />
          </td>
        ))}
      </tr>
    );
  };

  const stdIncomeSections = ['Revenue & Gross Margin', 'Operating Expenses', 'Non-Operating & Tax'];
  const visCustRev = customItems.filter(ci => ci.isVisible && ci.section === 'Revenue & Gross Margin');
  const visCustOpex = customItems.filter(ci => ci.isVisible && ci.section === 'Operating Expenses');
  const visCustNonOp = customItems.filter(
    ci => ci.isVisible && (ci.section === 'Non-Operating & Tax' || !stdIncomeSections.includes(ci.section))
  );

  const renderCustomRows = (section: string, includeUnscoped = false) =>
    customItems
      .filter(
        ci =>
          (ci.section === section || (includeUnscoped && !stdIncomeSections.includes(ci.section))) &&
          (ci.isVisible || isCustomizeMode)
      )
      .map(ci => (
        <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50/90' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
          <td className="py-2 px-4 font-sans text-indigo-900 pl-6">
            <div className="flex items-center gap-2">
              {isCustomizeMode && (
                <input
                  type="checkbox"
                  checked={ci.isVisible}
                  onChange={() => onToggleElem(ci.id)}
                  className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
                />
              )}
              <EditableLabel
                label={ci.label}
                isEditMode={isEditMode}
                onSave={(lbl) => onUpdateLabel('custom', ci.id, lbl)}
              />
              <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1.5 py-0.5 rounded shrink-0">
                Custom
              </span>
            </div>
          </td>
          {incomeData.map(d => (
            <td key={d.year} className={`py-2 px-3 text-right ${ci.sign === 'negative' ? 'text-rose-700' : 'text-indigo-800'}`}>
              <EditableCell
                val={getCustomItemYearValue(ci, d.year)}
                isEditMode={isEditMode}
                onSave={(val) => onUpdateValue('custom', ci.id, d.year, val)}
                isNegative={ci.sign === 'negative'}
              />
            </td>
          ))}
        </tr>
      ));

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
            {renderIncomeRow('inc_revenue', 'revenue', 'Sales Revenue', d => {
              const custSum = visCustRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
              return Math.round((d.revenue - custSum) * 100) / 100;
            }, {
              rowClass: 'font-semibold text-slate-900 hover:bg-slate-50/70',
              cellClass: 'text-emerald-700 font-bold',
            })}
            {renderCustomRows('Revenue & Gross Margin')}
            {renderIncomeRow('inc_costOfSales', 'costOfSales', 'Cost of Goods Sold / Cost of Sales', d => d.costOfSales, {
              indent: true,
              isNegative: true,
              cellClass: 'text-rose-700',
            })}

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

            {renderIncomeRow('inc_operatingExpenses', 'operatingExpenses', 'Operating Expenses (Admin & Selling)', d => {
              const custSum = visCustOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
              return Math.round((d.operatingExpenses + custSum) * 100) / 100;
            }, {
              indent: true,
              isNegative: true,
              cellClass: 'text-rose-700',
            })}
            {renderCustomRows('Operating Expenses')}
            {renderIncomeRow('inc_depreciationExpense', 'depreciationExpense', 'Depreciation Expense (Schedule of Fixed Assets)', d => d.depreciationExpense, {
              indent: true,
              isNegative: true,
              cellClass: 'text-rose-700',
            })}

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

            {renderIncomeRow('inc_interestIncome', 'interestIncome', 'Interest & Investment Income', d => {
              const custSum = visCustNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
              return Math.round((d.interestIncome - custSum) * 100) / 100;
            }, {
              indent: true,
              cellClass: 'text-emerald-700',
            })}
            {renderIncomeRow('inc_financeCost', 'financeCost', 'Finance Cost (Mandatory Direct Entry)', d => d.financeCost, {
              indent: true,
              isNegative: true,
              rowClass: 'bg-amber-50/75 font-semibold',
              cellClass: 'text-amber-800',
            })}
            {renderCustomRows('Non-Operating & Tax', true)}

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

            {renderIncomeRow('inc_taxExpense', 'taxExpense', 'Income Tax Expense', d => d.taxExpense, {
              indent: true,
              isNegative: true,
              cellClass: 'text-rose-700',
            })}

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
  isCustomizeMode,
  isElemVisible,
  onToggleElem,
  customItems,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  balanceData: BalanceSheetYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  isCustomizeMode: boolean;
  isElemVisible: (id: string) => boolean;
  onToggleElem: (id: string) => void;
  customItems: StatementConfigItem[];
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string, configId?: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  const renderToggle = (configId: string) =>
    isCustomizeMode ? (
      <input
        type="checkbox"
        checked={isElemVisible(configId)}
        onChange={() => onToggleElem(configId)}
        title="Include or exclude this element from calculation and exports"
        className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
      />
    ) : null;

  const renderBSRow = (
    configId: string,
    field: string,
    defaultLabel: string,
    getFallback: (d: BalanceSheetYear) => number,
    cellClass = 'text-slate-700',
    labelClass = 'text-slate-700'
  ) => {
    const visible = isElemVisible(configId);
    if (!visible && !isCustomizeMode) return null;
    return (
      <tr key={configId} className={!visible ? 'opacity-45 bg-slate-50/90' : 'hover:bg-slate-50/70'}>
        <td className={`py-2 px-4 font-sans ${labelClass}`}>
          <div className="flex items-center gap-2">
            {renderToggle(configId)}
            <EditableLabel
              label={getLbl('bs', field, defaultLabel, configId)}
              isEditMode={isEditMode}
              onSave={(lbl) => onUpdateLabel('bs', field, lbl)}
            />
          </div>
        </td>
        {balanceData.map(d => (
          <td key={d.year} className={`py-2 px-3 text-right ${cellClass}`}>
            <EditableCell
              val={getVal('bs', field, d.year, getFallback(d))}
              isEditMode={isEditMode}
              onSave={(val) => onUpdateValue('bs', field, d.year, val)}
            />
          </td>
        ))}
      </tr>
    );
  };

  const renderCustomRows = (section: string) =>
    customItems
      .filter(ci => ci.section === section && (ci.isVisible || isCustomizeMode))
      .map(ci => (
        <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50/90' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
          <td className="py-2 px-4 font-sans text-indigo-900">
            <div className="flex items-center gap-2">
              {isCustomizeMode && (
                <input
                  type="checkbox"
                  checked={ci.isVisible}
                  onChange={() => onToggleElem(ci.id)}
                  className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
                />
              )}
              <EditableLabel
                label={ci.label}
                isEditMode={isEditMode}
                onSave={(lbl) => onUpdateLabel('custom', ci.id, lbl)}
              />
              <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1.5 py-0.5 rounded shrink-0">
                Custom
              </span>
            </div>
          </td>
          {balanceData.map(d => (
            <td key={d.year} className="py-2 px-3 text-right text-indigo-800">
              <EditableCell
                val={getCustomItemYearValue(ci, d.year)}
                isEditMode={isEditMode}
                onSave={(val) => onUpdateValue('custom', ci.id, d.year, val)}
                isNegative={ci.sign === 'negative'}
              />
            </td>
          ))}
        </tr>
      ));

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

            {renderBSRow('bs_cashAndEquivalents', 'cashAndEquivalents', 'Cash and Cash Equivalents', d => d.cashAndEquivalents, 'text-sky-700 font-medium', 'text-slate-800 font-medium')}
            {renderBSRow('bs_tradeReceivables', 'tradeReceivables', 'Trade Receivables', d => d.tradeReceivables)}
            {renderBSRow('bs_otherReceivables', 'otherReceivables', 'Other Receivables & Prepayments', d => d.otherReceivables)}
            {renderBSRow('bs_closingInventory', 'closingInventory', 'Closing Inventory', d => d.closingInventory)}
            {renderBSRow('bs_otherCurrentAssets', 'otherCurrentAssets', 'Other Current Assets', d => d.otherCurrentAssets)}
            {renderCustomRows('Current Assets')}

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

            {renderBSRow('bs_propertyPlantEquipment', 'propertyPlantEquipment', 'Property, Plant and Equipment (NBV)', d => d.propertyPlantEquipment, 'text-emerald-700 font-medium', 'text-slate-800 font-medium')}
            {renderBSRow('bs_intangibleAssets', 'intangibleAssets', 'Intangible Assets', d => d.intangibleAssets)}
            {renderBSRow('bs_otherNonCurrentAssets', 'otherNonCurrentAssets', 'Other Non-Current Assets', d => d.otherNonCurrentAssets)}
            {renderCustomRows('Non-Current Assets')}

            <tr className="bg-slate-100/80 font-semibold text-slate-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('bs', 'totalNonCurrentAssets', 'Total Non-Current Assets')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalNonCurrentAssets', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-slate-800">
                  <EditableCell
                    val={getVal('bs', 'totalNonCurrentAssets', d.year, d.totalNonCurrentAssets)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalNonCurrentAssets', d.year, val)}
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

            {/* Current Liabilities */}
            <tr className="bg-rose-50/75 font-sans font-bold text-rose-800">
              <td colSpan={balanceData.length + 1} className="py-2 px-4 uppercase text-[11px]">Current Liabilities</td>
            </tr>

            {renderBSRow('bs_tradePayables', 'tradePayables', 'Trade Payables', d => d.tradePayables)}
            {renderBSRow('bs_otherPayables', 'otherPayables', 'Other Payables & Accruals', d => d.otherPayables)}
            {renderBSRow('bs_shortTermBorrowings', 'shortTermBorrowings', 'Short-Term Borrowings', d => d.shortTermBorrowings)}
            {renderBSRow('bs_currentTaxPayable', 'currentTaxPayable', 'Current Tax Payable', d => d.currentTaxPayable)}
            {renderCustomRows('Current Liabilities')}

            <tr className="bg-slate-100/80 font-semibold text-slate-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('bs', 'totalCurrentLiabilities', 'Total Current Liabilities')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalCurrentLiabilities', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-rose-800">
                  <EditableCell
                    val={getVal('bs', 'totalCurrentLiabilities', d.year, d.totalCurrentLiabilities)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalCurrentLiabilities', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {/* Non-Current Liabilities */}
            <tr className="bg-rose-50/75 font-sans font-bold text-rose-800">
              <td colSpan={balanceData.length + 1} className="py-2 px-4 uppercase text-[11px]">Non-Current Liabilities</td>
            </tr>

            {renderBSRow('bs_longTermBorrowings', 'longTermBorrowings', 'Long-Term Borrowings', d => d.longTermBorrowings)}
            {renderBSRow('bs_otherNonCurrentLiabilities', 'otherNonCurrentLiabilities', 'Other Non-Current Liabilities', d => d.otherNonCurrentLiabilities)}
            {renderCustomRows('Non-Current Liabilities')}

            <tr className="bg-slate-100/80 font-semibold text-slate-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('bs', 'totalNonCurrentLiabilities', 'Total Non-Current Liabilities')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalNonCurrentLiabilities', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-slate-800">
                  <EditableCell
                    val={getVal('bs', 'totalNonCurrentLiabilities', d.year, d.totalNonCurrentLiabilities)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalNonCurrentLiabilities', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            <tr className="bg-rose-50/70 font-bold text-rose-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('bs', 'totalLiabilities', 'TOTAL LIABILITIES')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalLiabilities', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-rose-900">
                  <EditableCell
                    val={getVal('bs', 'totalLiabilities', d.year, d.totalLiabilities)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalLiabilities', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {/* Equity */}
            <tr className="bg-teal-50/75 font-sans font-bold text-teal-800">
              <td colSpan={balanceData.length + 1} className="py-2 px-4 uppercase text-[11px]">Shareholders Equity</td>
            </tr>

            {renderBSRow('bs_shareCapital', 'shareCapital', 'Share Capital', d => d.shareCapital)}
            {renderBSRow('bs_retainedEarnings', 'retainedEarnings', 'Retained Earnings', d => d.retainedEarnings)}
            {renderBSRow('bs_reserves', 'reserves', 'Reserves & Surplus', d => d.reserves)}
            {renderCustomRows('Equity')}

            <tr className="bg-teal-50/75 font-bold text-teal-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('bs', 'totalEquity', 'TOTAL EQUITY')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('bs', 'totalEquity', lbl)}
                />
              </td>
              {balanceData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-teal-900">
                  <EditableCell
                    val={getVal('bs', 'totalEquity', d.year, d.totalEquity)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('bs', 'totalEquity', d.year, val)}
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
  isCustomizeMode,
  isElemVisible,
  onToggleElem,
  customItems,
  getVal,
  getLbl,
  onUpdateValue,
  onUpdateLabel,
}: {
  equityData: EquityStatementYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  isCustomizeMode: boolean;
  isElemVisible: (id: string) => boolean;
  onToggleElem: (id: string) => void;
  customItems: StatementConfigItem[];
  getVal: (stmt: string, field: string, year: number, fallback: number) => number;
  getLbl: (stmt: string, field: string, defaultLabel: string, configId?: string) => string;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  const renderToggle = (configId: string) =>
    isCustomizeMode ? (
      <input
        type="checkbox"
        checked={isElemVisible(configId)}
        onChange={() => onToggleElem(configId)}
        title="Include or exclude this element from calculation and exports"
        className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
      />
    ) : null;

  const renderEqRow = (
    configId: string,
    field: string,
    defaultLabel: string,
    getFallback: (d: EquityStatementYear) => number,
    opts?: { indent?: boolean; isNegative?: boolean; cellClass?: string; labelClass?: string }
  ) => {
    const visible = isElemVisible(configId);
    if (!visible && !isCustomizeMode) return null;
    return (
      <tr key={configId} className={!visible ? 'opacity-45 bg-slate-50/90' : 'hover:bg-slate-50/70'}>
        <td className={`py-2 px-4 font-sans ${opts?.indent ? 'pl-6 text-slate-600' : opts?.labelClass || 'text-slate-700 font-medium'}`}>
          <div className="flex items-center gap-2">
            {renderToggle(configId)}
            <EditableLabel
              label={getLbl('eq', field, defaultLabel, configId)}
              isEditMode={isEditMode}
              onSave={(lbl) => onUpdateLabel('eq', field, lbl)}
            />
          </div>
        </td>
        {equityData.map(d => (
          <td key={d.year} className={`py-2 px-3 text-right ${opts?.cellClass || 'text-slate-700'}`}>
            <EditableCell
              val={getVal('eq', field, d.year, getFallback(d))}
              isEditMode={isEditMode}
              onSave={(val) => onUpdateValue('eq', field, d.year, val)}
              isNegative={opts?.isNegative}
            />
          </td>
        ))}
      </tr>
    );
  };

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
            {renderEqRow('eq_openingShareCapital', 'openingShareCapital', 'Opening Share Capital', d => d.openingShareCapital)}
            {renderEqRow('eq_shareCapitalIssued', 'shareCapitalIssued', 'Share Capital Issued during Year', d => d.shareCapitalIssued, { indent: true })}

            <tr className="bg-slate-100/75 font-semibold text-slate-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('eq', 'closingShareCapital', 'Closing Share Capital')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'closingShareCapital', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-slate-800">
                  <EditableCell
                    val={getVal('eq', 'closingShareCapital', d.year, d.closingShareCapital)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'closingShareCapital', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {renderEqRow('eq_openingRetainedEarnings', 'openingRetainedEarnings', 'Opening Retained Earnings', d => d.openingRetainedEarnings)}
            {renderEqRow('eq_netProfitForYear', 'netProfitForYear', 'Net Profit for the Year', d => d.netProfitForYear, { indent: true, cellClass: 'text-emerald-700 font-medium' })}
            {renderEqRow('eq_dividendsPaid', 'dividendsPaid', 'Dividends Paid to Shareholders', d => d.dividendsPaid, { indent: true, isNegative: true, cellClass: 'text-rose-700' })}

            <tr className="bg-slate-100/75 font-semibold text-slate-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('eq', 'closingRetainedEarnings', 'Closing Retained Earnings')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'closingRetainedEarnings', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-slate-800">
                  <EditableCell
                    val={getVal('eq', 'closingRetainedEarnings', d.year, d.closingRetainedEarnings)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'closingRetainedEarnings', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {renderEqRow('eq_openingReserves', 'openingReserves', 'Opening Reserves', d => d.openingReserves)}
            {renderEqRow('eq_otherComprehensiveIncome', 'otherComprehensiveIncome', 'Other Comprehensive Income', d => d.otherComprehensiveIncome, { indent: true })}

            <tr className="bg-slate-100/75 font-semibold text-slate-900">
              <td className="py-2 px-4 font-sans">
                <EditableLabel
                  label={getLbl('eq', 'closingReserves', 'Closing Reserves')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'closingReserves', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2 px-3 text-right font-bold text-slate-800">
                  <EditableCell
                    val={getVal('eq', 'closingReserves', d.year, d.closingReserves)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'closingReserves', d.year, val)}
                  />
                </td>
              ))}
            </tr>

            {customItems
              .filter(ci => ci.isVisible || isCustomizeMode)
              .map(ci => (
                <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50/90' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                  <td className="py-2 px-4 font-sans text-indigo-900">
                    <div className="flex items-center gap-2">
                      {isCustomizeMode && (
                        <input
                          type="checkbox"
                          checked={ci.isVisible}
                          onChange={() => onToggleElem(ci.id)}
                          className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
                        />
                      )}
                      <EditableLabel
                        label={ci.label}
                        isEditMode={isEditMode}
                        onSave={(lbl) => onUpdateLabel('custom', ci.id, lbl)}
                      />
                      <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1.5 py-0.5 rounded shrink-0">
                        Custom
                      </span>
                    </div>
                  </td>
                  {equityData.map(d => (
                    <td key={d.year} className={`py-2 px-3 text-right ${ci.sign === 'negative' ? 'text-rose-700' : 'text-indigo-800'}`}>
                      <EditableCell
                        val={getCustomItemYearValue(ci, d.year)}
                        isEditMode={isEditMode}
                        onSave={(val) => onUpdateValue('custom', ci.id, d.year, val)}
                        isNegative={ci.sign === 'negative'}
                      />
                    </td>
                  ))}
                </tr>
              ))}

            <tr className="bg-slate-100/90 font-semibold text-slate-900 border-t border-slate-300">
              <td className="py-2.5 px-4 font-sans">
                <EditableLabel
                  label={getLbl('eq', 'totalOpeningEquity', 'TOTAL OPENING EQUITY')}
                  isEditMode={isEditMode}
                  onSave={(lbl) => onUpdateLabel('eq', 'totalOpeningEquity', lbl)}
                />
              </td>
              {equityData.map(d => (
                <td key={d.year} className="py-2.5 px-3 text-right font-bold text-slate-800">
                  <EditableCell
                    val={getVal('eq', 'totalOpeningEquity', d.year, d.totalOpeningEquity)}
                    isEditMode={isEditMode}
                    onSave={(val) => onUpdateValue('eq', 'totalOpeningEquity', d.year, val)}
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
  isCustomizeMode,
  isElemVisible,
  onToggleElem,
  customItems,
  getLbl,
  onUpdateCategory,
  onUpdateValue,
  onUpdateLabel,
}: {
  fixedAssetsData: ScheduleOfFixedAssetsYear[];
  currency: string;
  companyProfile?: CompanyProfile;
  isEditMode: boolean;
  isCustomizeMode: boolean;
  isElemVisible: (id: string) => boolean;
  onToggleElem: (id: string) => void;
  customItems: StatementConfigItem[];
  getLbl: (stmt: string, field: string, defaultLabel: string, configId?: string) => string;
  onUpdateCategory: (year: number, catIndex: number, field: string, val: number) => void;
  onUpdateValue: (stmt: string, field: string, year: number, val: number) => void;
  onUpdateLabel: (stmt: string, field: string, newLabel: string) => void;
}) {
  if (fixedAssetsData.length === 0) return null;
  const latest = fixedAssetsData[fixedAssetsData.length - 1];
  const standardCategories = latest.allCategories || latest.categories.slice(0, FIXED_ASSET_CATEGORY_CONFIG_IDS.length);

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
              <th className="py-3 px-3 text-right text-emerald-700 font-semibold">Dep. Disposals</th>
              <th className="py-3 px-3 text-right text-slate-900 font-semibold">Closing Dep</th>
              <th className="py-3 px-4 text-right text-emerald-700 font-bold">Net Book Value</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200/70 tabular-nums">
            {standardCategories.map((c, idx) => {
              const cfgId = FIXED_ASSET_CATEGORY_CONFIG_IDS[idx];
              const visible = cfgId ? isElemVisible(cfgId) : true;
              if (!visible && !isCustomizeMode) return null;
              const effectiveName = getLbl('fa', String(idx), c.name, cfgId);

              return (
                <tr key={cfgId || c.name} className={!visible ? 'opacity-45 bg-slate-50/90' : 'hover:bg-slate-50'}>
                  <td className="py-2.5 px-4 font-sans text-slate-900 font-medium">
                    <div className="flex items-center gap-2">
                      {isCustomizeMode && cfgId && (
                        <input
                          type="checkbox"
                          checked={visible}
                          onChange={() => onToggleElem(cfgId)}
                          title="Include or exclude this asset class from calculations and exports"
                          className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
                        />
                      )}
                      <EditableLabel
                        label={effectiveName}
                        isEditMode={isEditMode}
                        onSave={(lbl) => onUpdateLabel('fa', String(idx), lbl)}
                      />
                    </div>
                  </td>
                  <td className="py-2.5 px-2 text-right text-slate-600">
                    <EditableCell
                      val={c.depRate}
                      isEditMode={isEditMode}
                      onSave={(v) => onUpdateCategory(latest.year, idx, 'depRate', v)}
                    />
                    %
                  </td>
                  <td className="py-2.5 px-3 text-right text-slate-700">
                    <EditableCell
                      val={c.costOpening}
                      isEditMode={isEditMode}
                      onSave={(v) => onUpdateCategory(latest.year, idx, 'costOpening', v)}
                    />
                  </td>
                  <td className="py-2.5 px-3 text-right text-emerald-700">
                    <EditableCell
                      val={c.additions}
                      isEditMode={isEditMode}
                      onSave={(v) => onUpdateCategory(latest.year, idx, 'additions', v)}
                    />
                  </td>
                  <td className="py-2.5 px-3 text-right text-rose-700">
                    <EditableCell
                      val={c.disposals}
                      isEditMode={isEditMode}
                      onSave={(v) => onUpdateCategory(latest.year, idx, 'disposals', v)}
                    />
                  </td>
                  <td className="py-2.5 px-3 text-right text-slate-900 font-semibold">
                    {c.costClosing.toLocaleString()}
                  </td>
                  <td className="py-2.5 px-3 text-right text-slate-600">
                    <EditableCell
                      val={c.depOpening}
                      isEditMode={isEditMode}
                      onSave={(v) => onUpdateCategory(latest.year, idx, 'depOpening', v)}
                    />
                  </td>
                  <td className="py-2.5 px-3 text-right text-rose-700">
                    <EditableCell
                      val={c.depCharge}
                      isEditMode={isEditMode}
                      onSave={(v) => onUpdateCategory(latest.year, idx, 'depCharge', v)}
                    />
                  </td>
                  <td className="py-2.5 px-3 text-right text-emerald-700">
                    <EditableCell
                      val={c.depDisposals}
                      isEditMode={isEditMode}
                      onSave={(v) => onUpdateCategory(latest.year, idx, 'depDisposals', v)}
                    />
                  </td>
                  <td className="py-2.5 px-3 text-right text-slate-700">
                    {c.depClosing.toLocaleString()}
                  </td>
                  <td className="py-2.5 px-4 text-right text-emerald-700 font-bold">
                    {c.netBookValue.toLocaleString()}
                  </td>
                </tr>
              );
            })}

            {customItems
              .filter(ci => ci.isVisible || isCustomizeMode)
              .map(ci => {
                const val = getCustomItemYearValue(ci, latest.year);
                return (
                  <tr key={ci.id} className={!ci.isVisible ? 'opacity-45 bg-slate-50/90' : 'bg-indigo-50/35 hover:bg-indigo-50/60'}>
                    <td className="py-2.5 px-4 font-sans text-indigo-900 font-medium">
                      <div className="flex items-center gap-2">
                        {isCustomizeMode && (
                          <input
                            type="checkbox"
                            checked={ci.isVisible}
                            onChange={() => onToggleElem(ci.id)}
                            className="w-3.5 h-3.5 accent-indigo-600 rounded cursor-pointer shrink-0"
                          />
                        )}
                        <EditableLabel
                          label={ci.label}
                          isEditMode={isEditMode}
                          onSave={(lbl) => onUpdateLabel('custom', ci.id, lbl)}
                        />
                        <span className="text-[9px] font-mono uppercase bg-indigo-100 text-indigo-800 border border-indigo-200 px-1.5 py-0.5 rounded shrink-0">
                          Custom
                        </span>
                      </div>
                    </td>
                    <td className="py-2.5 px-2 text-right text-slate-500">0%</td>
                    <td className="py-2.5 px-3 text-right text-indigo-800">
                      <EditableCell
                        val={val}
                        isEditMode={isEditMode}
                        onSave={(v) => onUpdateValue('custom', ci.id, latest.year, v)}
                      />
                    </td>
                    <td className="py-2.5 px-3 text-right text-slate-500">0</td>
                    <td className="py-2.5 px-3 text-right text-slate-500">0</td>
                    <td className="py-2.5 px-3 text-right text-indigo-900 font-semibold">{val.toLocaleString()}</td>
                    <td className="py-2.5 px-3 text-right text-slate-500">0</td>
                    <td className="py-2.5 px-3 text-right text-slate-500">0</td>
                    <td className="py-2.5 px-3 text-right text-slate-500">0</td>
                    <td className="py-2.5 px-3 text-right text-slate-500">0</td>
                    <td className="py-2.5 px-4 text-right text-indigo-800 font-bold">{val.toLocaleString()}</td>
                  </tr>
                );
              })}

            {/* Total Carrying Net Book Value row with Double Underline */}
            <tr className="bg-emerald-50/85 font-bold text-emerald-800 border-t-2 border-emerald-300 border-b-4 border-double border-emerald-600 text-xs">
              <td className="py-3 px-4 font-sans font-bold text-slate-900">
                TOTAL PROPERTY, PLANT & EQUIPMENT
              </td>
              <td className="py-3 px-2 text-right text-slate-500">-</td>
              <td className="py-3 px-3 text-right text-slate-700 font-semibold">{latest.totalCostOpening.toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-emerald-700 font-bold">{latest.totalAdditions.toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-rose-700">{latest.totalDisposals.toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-slate-900 font-bold">{latest.totalCostClosing.toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-slate-600">{latest.totalDepOpening.toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-rose-700">{latest.totalDepCharge.toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-emerald-700">{latest.totalDepDisposals.toLocaleString()}</td>
              <td className="py-3 px-3 text-right text-slate-900">{latest.totalDepClosing.toLocaleString()}</td>
              <td className="py-3 px-4 text-right text-emerald-800 font-extrabold text-sm font-mono">
                {latest.totalNetBookValue.toLocaleString()}
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
            Implemented Repayment Schedule Coverage · {parsedSchedule?.metadata?.sourceFileName || companyProfile.companyName?.trim() || 'Untitled Entity'}
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
