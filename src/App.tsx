/**
 * Financials with Tuhin
 * Section 76: Statement of Cash Flows and Expanded Financial Statement Generation & Export Facility
 * Lead Developer & Financial Architect: Mohammad Towhidul Akbar Tuhin
 * Officer, Sammilito Islami Bank PLC (Transferee of Assets & Liabilities of Union Bank PLC)
 * BBA, Heriot Watt University, UK · MBA, University of Chittagong
 */

import React, { useState, useMemo } from 'react';
import { HeaderNav } from './components/HeaderNav';
import { DeveloperFirstLookCard } from './components/DeveloperFirstLookCard';
import { StatementsViewer } from './components/StatementsViewer';
import { FinancialDataEntry } from './components/FinancialDataEntry';
import { StatementConfigView } from './components/StatementConfigModal';
import { ProjectionEngineView } from './components/ProjectionEngineView';
import { CashFlowReconciliationAudit } from './components/CashFlowReconciliationAudit';
import { CompanyProfileModal } from './components/CompanyProfileModal';
import { DeveloperProfileModal } from './components/DeveloperProfileModal';
import {
  INITIAL_COMPANY_PROFILE,
  DEFAULT_PROJECTION_ASSUMPTIONS,
  INITIAL_HISTORICAL_INPUTS,
  DEFAULT_STATEMENT_CONFIG_ITEMS,
  generateProjectedYears,
  generateYearDataForTargetYear,
  calculateScheduleOfFixedAssets,
  calculateIncomeStatement,
  calculateCashFlowStatement,
  calculateEquityStatement,
  calculateBalanceSheet,
  generateNotesToFinancialStatements,
  deriveDependentFinancials,
  calculateDriverAutoFillSeries,
  recalculateAllYearsWithUpdatedDriver,
  buildPreviousFormatData,
  buildProjectedFormatData,
  validateFinancialModel,
} from './services/financialEngine';
import {
  exportCompletePackageExcel,
  exportCompletePackageWord,
  exportCompletePackagePdf,
  exportPreviousFinancialsExcel,
  exportPreviousFinancialsWord,
  exportPreviousFinancialsPdf,
  exportProjectedFinancialsExcel,
  exportProjectedFinancialsWord,
  exportProjectedFinancialsPdf,
} from './services/exportService';
import {
  CompanyProfile,
  YearDataInput,
  ProjectionAssumptions,
  StatementConfigItem,
  StatementCellOverrides,
  StatementLabelOverrides,
  FinancialNoteItem,
  InputMode,
  PrimaryDriverKey,
  AutoFillConfigs,
  StatementFormatMode,
} from './types/financial';

export default function App() {
  const [activeTab, setActiveTab] = useState<string>('statements');
  const [formatMode, setFormatMode] = useState<StatementFormatMode>('section76');
  const [companyProfile, setCompanyProfile] = useState<CompanyProfile>(INITIAL_COMPANY_PROFILE);
  const [assumptions, setAssumptions] = useState<ProjectionAssumptions>(DEFAULT_PROJECTION_ASSUMPTIONS);
  const [inputMode, setInputMode] = useState<InputMode>('manual');
  const [historicalYearsStore, setHistoricalYearsStore] = useState<Record<number, YearDataInput>>(() =>
    Object.fromEntries(INITIAL_HISTORICAL_INPUTS.map(item => [item.year, item]))
  );
  const [customProjectedOverrides, setCustomProjectedOverrides] = useState<Record<number, YearDataInput>>({});
  const [statementConfig, setStatementConfig] = useState<StatementConfigItem[]>(DEFAULT_STATEMENT_CONFIG_ITEMS);

  // Independent Auto-Fill Configurations for the 6 primary financial items (Section 76 Intelligent Year-Wise Filling)
  const [autoFillConfigs, setAutoFillConfigs] = useState<AutoFillConfigs>({});

  // Statement-level cell & label overrides for universal inline editing
  const [cellOverrides, setCellOverrides] = useState<StatementCellOverrides>({});
  const [labelOverrides, setLabelOverrides] = useState<StatementLabelOverrides>({});
  const [customNotes, setCustomNotes] = useState<FinancialNoteItem[] | null>(null);

  // Modals state
  const [isProfileModalOpen, setIsProfileModalOpen] = useState(false);
  const [isDeveloperModalOpen, setIsDeveloperModalOpen] = useState(false);

  // Derived: Active Historical Years based on user's flexible individual year selection
  const activeHistoricalInputs = useMemo(() => {
    const selectedHist = assumptions.selectedHistoricalYears && assumptions.selectedHistoricalYears.length > 0
      ? [...assumptions.selectedHistoricalYears].sort((a, b) => a - b)
      : INITIAL_HISTORICAL_INPUTS.map(h => h.year);

    const existingList = Object.values(historicalYearsStore).sort((a, b) => a.year - b.year);
    const resolved: YearDataInput[] = [];
    for (const yr of selectedHist) {
      if (historicalYearsStore[yr]) {
        resolved.push({ ...historicalYearsStore[yr], isProjected: false });
      } else {
        resolved.push(
          generateYearDataForTargetYear(
            yr,
            false,
            resolved.length > 0 ? resolved : existingList,
            assumptions
          )
        );
      }
    }
    return resolved;
  }, [assumptions, historicalYearsStore]);

  // Derived: Generate Projected Years based on user's selected projected years / horizon
  const baseYearInput = activeHistoricalInputs[activeHistoricalInputs.length - 1] || INITIAL_HISTORICAL_INPUTS[INITIAL_HISTORICAL_INPUTS.length - 1];
  const generatedProjected = useMemo(() => {
    const raw = generateProjectedYears(baseYearInput, assumptions);
    return raw.map(y => customProjectedOverrides[y.year] || y);
  }, [baseYearInput, assumptions, customProjectedOverrides]);

  // Combined All Years (Selected Historical + Selected Projected)
  const allYearsInputs = useMemo(() => {
    return [...activeHistoricalInputs, ...generatedProjected];
  }, [activeHistoricalInputs, generatedProjected]);

  // Five Standard Financial Statements + Schedule of Fixed Assets Calculations
  const fixedAssetsSchedule = useMemo(() => {
    return calculateScheduleOfFixedAssets(allYearsInputs);
  }, [allYearsInputs]);

  const incomeStatements = useMemo(() => {
    return calculateIncomeStatement(allYearsInputs, fixedAssetsSchedule);
  }, [allYearsInputs, fixedAssetsSchedule]);

  const cashFlowStatements = useMemo(() => {
    return calculateCashFlowStatement(allYearsInputs, incomeStatements, fixedAssetsSchedule);
  }, [allYearsInputs, incomeStatements, fixedAssetsSchedule]);

  const equityStatements = useMemo(() => {
    return calculateEquityStatement(allYearsInputs, incomeStatements);
  }, [allYearsInputs, incomeStatements]);

  const balanceSheets = useMemo(() => {
    return calculateBalanceSheet(allYearsInputs, equityStatements, fixedAssetsSchedule, cashFlowStatements);
  }, [allYearsInputs, equityStatements, fixedAssetsSchedule, cashFlowStatements]);

  const defaultNotes = useMemo(() => {
    return generateNotesToFinancialStatements(
      companyProfile,
      allYearsInputs.map(y => y.year),
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      fixedAssetsSchedule
    );
  }, [companyProfile, allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, fixedAssetsSchedule]);

  const effectiveNotes = customNotes || defaultNotes;

  // Exact Compliance Format Datasets (Previous Financials & Projected Financials)
  const previousFormatData = useMemo(() => {
    return buildPreviousFormatData(
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule
    );
  }, [allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, equityStatements, fixedAssetsSchedule]);

  const projectedFormatData = useMemo(() => {
    return buildProjectedFormatData(
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule
    );
  }, [allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, equityStatements, fixedAssetsSchedule]);

  // Model-Wide Accounting & Fixed-Driver Validation
  const validationIssues = useMemo(() => {
    return validateFinancialModel(
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule
    );
  }, [allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, equityStatements, fixedAssetsSchedule]);

  // Unified Assumptions Updater (keeps Horizon slider & selectedProjectedYears synchronized)
  const handleUpdateAssumptions = (updated: ProjectionAssumptions) => {
    const countChanged = updated.projectedYearsCount !== assumptions.projectedYearsCount;
    const sameProjArray = updated.selectedProjectedYears === assumptions.selectedProjectedYears;
    const histYears = updated.selectedHistoricalYears && updated.selectedHistoricalYears.length > 0
      ? [...updated.selectedHistoricalYears].sort((a, b) => a - b)
      : activeHistoricalInputs.map(y => y.year);
    const maxHist = histYears[histYears.length - 1] || 2025;

    let nextProjYears = updated.selectedProjectedYears;
    if (countChanged && sameProjArray) {
      const targetCount = Math.max(1, updated.projectedYearsCount);
      nextProjYears = Array.from({ length: targetCount }, (_, i) => maxHist + 1 + i);
    }

    setAssumptions({
      ...updated,
      historicalYearsCount: histYears.length,
      selectedHistoricalYears: histYears,
      selectedProjectedYears: nextProjYears,
      projectedYearsCount: nextProjYears ? nextProjYears.length : updated.projectedYearsCount,
    });
  };

  // Individual Year Selection Handlers (Previous & Projected)
  const handleToggleHistoricalYear = (year: number) => {
    const current = activeHistoricalInputs.map(y => y.year);
    if (current.includes(year)) {
      if (current.length <= 1) return;
      const nextHist = current.filter(y => y !== year).sort((a, b) => a - b);
      setAssumptions(prev => ({
        ...prev,
        historicalYearsCount: nextHist.length,
        selectedHistoricalYears: nextHist,
      }));
      return;
    }

    if (!historicalYearsStore[year]) {
      const generated = generateYearDataForTargetYear(year, false, activeHistoricalInputs, assumptions);
      setHistoricalYearsStore(prev => ({ ...prev, [year]: generated }));
    }
    const nextHist = [...current, year].sort((a, b) => a - b);
    const currentProj = (assumptions.selectedProjectedYears || generatedProjected.map(y => y.year)).filter(y => y !== year);
    const nextProj = currentProj.length > 0 ? currentProj : [nextHist[nextHist.length - 1] + 1];

    setAssumptions(prev => ({
      ...prev,
      historicalYearsCount: nextHist.length,
      selectedHistoricalYears: nextHist,
      selectedProjectedYears: nextProj,
      projectedYearsCount: nextProj.length,
    }));
  };

  const handleToggleProjectedYear = (year: number) => {
    const current = assumptions.selectedProjectedYears || generatedProjected.map(y => y.year);
    if (current.includes(year)) {
      if (current.length <= 1) return;
      const nextProj = current.filter(y => y !== year).sort((a, b) => a - b);
      setAssumptions(prev => ({
        ...prev,
        selectedProjectedYears: nextProj,
        projectedYearsCount: nextProj.length,
      }));
      return;
    }

    const nextProj = [...current, year].sort((a, b) => a - b);
    setAssumptions(prev => ({
      ...prev,
      selectedProjectedYears: nextProj,
      projectedYearsCount: nextProj.length,
    }));
  };

  const handleSetHistoricalYears = (years: number[]) => {
    const uniqueSorted = Array.from(new Set(years)).sort((a, b) => a - b);
    if (uniqueSorted.length === 0) return;

    setHistoricalYearsStore(prev => {
      const nextStore = { ...prev };
      const seedList = Object.values(nextStore).sort((a, b) => a.year - b.year);
      for (const yr of uniqueSorted) {
        if (!nextStore[yr]) {
          nextStore[yr] = generateYearDataForTargetYear(yr, false, seedList, assumptions);
          seedList.push(nextStore[yr]);
          seedList.sort((a, b) => a.year - b.year);
        }
      }
      return nextStore;
    });

    const filteredProj = (assumptions.selectedProjectedYears || generatedProjected.map(y => y.year)).filter(
      y => !uniqueSorted.includes(y)
    );
    const maxHist = uniqueSorted[uniqueSorted.length - 1];
    const nextProj = filteredProj.length > 0 ? filteredProj : [maxHist + 1];

    setAssumptions(prev => ({
      ...prev,
      historicalYearsCount: uniqueSorted.length,
      selectedHistoricalYears: uniqueSorted,
      selectedProjectedYears: nextProj,
      projectedYearsCount: nextProj.length,
    }));
  };

  const handleSetProjectedYears = (years: number[]) => {
    const histSet = new Set(activeHistoricalInputs.map(y => y.year));
    const uniqueSorted = Array.from(new Set(years)).filter(y => !histSet.has(y)).sort((a, b) => a - b);
    if (uniqueSorted.length === 0) return;

    setAssumptions(prev => ({
      ...prev,
      selectedProjectedYears: uniqueSorted,
      projectedYearsCount: uniqueSorted.length,
    }));
  };

  // Handler: Update Year Input
  const handleUpdateYearData = (updated: YearDataInput) => {
    if (!updated.isProjected) {
      setHistoricalYearsStore(prev => ({
        ...prev,
        [updated.year]: { ...updated, isProjected: false },
      }));
    } else {
      setCustomProjectedOverrides(prev => ({
        ...prev,
        [updated.year]: updated,
      }));
    }
  };

  // Handler: Apply Auto-Fill to all remaining years for a specific primary financial driver
  const handleApplyAutoFill = (
    driver: PrimaryDriverKey,
    baseYear: number,
    baseValue: number,
    customGrowthRate?: number
  ) => {
    // 1. Calculate the complete series across all years (both prior historical and future projected)
    const allYearsList = allYearsInputs.map(y => y.year);
    const series = calculateDriverAutoFillSeries(
      driver,
      baseYear,
      baseValue,
      allYearsList,
      assumptions,
      allYearsInputs,
      customGrowthRate
    );

    // 2. Re-derive all dependent financials chronologically maintaining 100% double-entry & cash flow integrity
    const updatedAllYears = recalculateAllYearsWithUpdatedDriver(
      allYearsInputs,
      driver,
      series,
      assumptions
    );

    // 3. Update both historical inputs store and custom projected overrides
    setHistoricalYearsStore(prev => {
      const nextStore = { ...prev };
      updatedAllYears.filter(y => !y.isProjected).forEach(y => {
        nextStore[y.year] = y;
      });
      return nextStore;
    });
    setCustomProjectedOverrides(prev => {
      const nextOverrides = { ...prev };
      updatedAllYears.filter(y => y.isProjected).forEach(y => {
        nextOverrides[y.year] = y;
      });
      return nextOverrides;
    });

    // 4. Save active Auto-Fill configuration for dynamic recalculation
    setAutoFillConfigs(prev => ({
      ...prev,
      [driver]: {
        driver,
        baseYear,
        baseValue,
        isActive: true,
        growthRatePct: customGrowthRate,
      },
    }));
  };

  // Handler: Clear / Deactivate Auto-Fill for a driver
  const handleClearAutoFill = (driver: PrimaryDriverKey) => {
    setAutoFillConfigs(prev => {
      const copy = { ...prev };
      delete copy[driver];
      return copy;
    });
  };

  // Handler: Update Any of the 6 Primary Financial Drivers (Interconnected Calculation Engine)
  const handleUpdatePrimaryDriver = (year: number, driver: keyof YearDataInput, val: number) => {
    const isPrimaryDriver = (
      driver === 'salesRevenue' ||
      driver === 'closingInventory' ||
      driver === 'financeCost' ||
      driver === 'netProfit' ||
      driver === 'repaymentOfBorrowings' ||
      driver === 'fixedAssetsGross'
    );

    // Dynamic Recalculation when Auto-Fill anchor base year is changed
    if (isPrimaryDriver) {
      const driverKey = driver as PrimaryDriverKey;
      const activeConfig = autoFillConfigs[driverKey];
      if (activeConfig && activeConfig.isActive && activeConfig.baseYear === year) {
        handleApplyAutoFill(driverKey, year, val, activeConfig.growthRatePct);
        return;
      }
    }

    const yearIdx = allYearsInputs.findIndex(y => y.year === year);
    if (yearIdx === -1) return;

    const currentYearObj = allYearsInputs[yearIdx];
    const prevYearObj = yearIdx > 0 ? allYearsInputs[yearIdx - 1] : null;

    const updatedDrivers = {
      year: currentYearObj.year,
      isProjected: currentYearObj.isProjected,
      salesRevenue: driver === 'salesRevenue' ? val : currentYearObj.salesRevenue,
      closingInventory: driver === 'closingInventory' ? val : currentYearObj.closingInventory,
      financeCost: driver === 'financeCost' ? val : currentYearObj.financeCost,
      netProfit: driver === 'netProfit' ? val : currentYearObj.netProfit,
      repaymentOfBorrowings: driver === 'repaymentOfBorrowings' ? val : currentYearObj.repaymentOfBorrowings,
      fixedAssetsGross: driver === 'fixedAssetsGross' ? val : currentYearObj.fixedAssetsGross,
    };

    const recomputedYear = deriveDependentFinancials(updatedDrivers, prevYearObj, assumptions, currentYearObj);

    if (!currentYearObj.isProjected) {
      setHistoricalYearsStore(prev => ({
        ...prev,
        [year]: { ...recomputedYear, isProjected: false },
      }));
    } else {
      setCustomProjectedOverrides(prev => ({
        ...prev,
        [year]: recomputedYear,
      }));
    }
  };

  // Handler: Reset custom projected overrides and generate automatic interconnected projection
  const handleGenerateAutomaticProjections = () => {
    setCustomProjectedOverrides({});
    setInputMode('automatic');
  };

  // Handler: Direct inline editing of ANY statement cell value
  const handleUpdateStatementValue = (statement: string, field: string, year: number, val: number) => {
    const key = `${statement}_${field}`;
    setCellOverrides(prev => ({
      ...prev,
      [key]: {
        ...(prev[key] || {}),
        [year]: val,
      },
    }));

    const fieldMapping: Record<string, keyof YearDataInput> = {
      'revenue': 'salesRevenue',
      'salesRevenue': 'salesRevenue',
      'costOfSales': 'costOfSales',
      'operatingExpenses': 'operatingExpenses',
      'financeCost': 'financeCost',
      'adjFinanceCost': 'financeCost',
      'netProfit': 'netProfit',
      'cashFlowFromFinancing': 'cashFlowFromFinancing',
      'directlyEnteredFinancingCashFlow': 'cashFlowFromFinancing',
      'repaymentOfBorrowings': 'repaymentOfBorrowings',
      'financingRepaymentBorrowings': 'repaymentOfBorrowings',
      'financingProceedsBorrowings': 'financingProceedsBorrowings',
      'financingProceedsShares': 'financingProceedsShares',
      'purchasePPE': 'fixedAssetsAdditions',
      'salePPE': 'fixedAssetsDisposals',
      'closingInventory': 'closingInventory',
      'tradeReceivables': 'tradeReceivables',
      'otherReceivables': 'otherReceivables',
      'tradePayables': 'tradePayables',
      'otherPayables': 'otherPayables',
      'cashAndEquivalents': 'cashAndEquivalentsBS',
      'balanceSheetCash': 'cashAndEquivalentsBS',
      'shortTermBorrowings': 'shortTermBorrowings',
      'longTermBorrowings': 'longTermBorrowings',
      'shareCapital': 'shareCapital',
      'dividendsPaid': 'dividendsPaid',
    };

    const targetField = fieldMapping[field];
    if (targetField) {
      const yearObj = allYearsInputs.find(y => y.year === year);
      if (yearObj) {
        const updatedYear: YearDataInput = {
          ...yearObj,
          [targetField]: field === 'purchasePPE' || field === 'dividendsPaid' ? Math.abs(val) : val,
        };

        if (targetField === 'repaymentOfBorrowings') {
          const absVal = Math.abs(val);
          const diff = absVal - (yearObj.repaymentOfBorrowings || 0);
          updatedYear.repaymentOfBorrowings = absVal;
          updatedYear.financingRepaymentBorrowings = -absVal;
          updatedYear.cashFlowFromFinancing =
            (updatedYear.financingProceedsShares || 0) +
            (updatedYear.financingProceedsBorrowings || 0) -
            absVal -
            updatedYear.dividendsPaid;
          updatedYear.longTermBorrowings = Math.max(0, (yearObj.longTermBorrowings || 0) - diff);
          updatedYear.cashAndEquivalentsBS = Math.max(0, (yearObj.cashAndEquivalentsBS || 0) - diff);
          updatedYear.financeCost = Math.max(200, Math.round((updatedYear.longTermBorrowings + updatedYear.shortTermBorrowings) * 0.085));
        }

        handleUpdateYearData(updatedYear);
      }
    }
  };

  // Handler: Direct inline editing of ANY statement item label
  const handleUpdateStatementLabel = (statement: string, field: string, newLabel: string) => {
    setLabelOverrides(prev => ({
      ...prev,
      [`${statement}_${field}`]: newLabel,
    }));
  };

  // Reset all custom overrides back to pure calculation engine defaults
  const handleResetOverrides = () => {
    setCellOverrides({});
    setLabelOverrides({});
    setCustomNotes(null);
  };

  // Handler: Export Complete Package (Section 76)
  const handleExportPackageExcel = () => {
    exportCompletePackageExcel(
      companyProfile,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule,
      effectiveNotes
    );
  };

  const handleExportPackageWord = () => {
    exportCompletePackageWord(
      companyProfile,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule,
      effectiveNotes
    );
  };

  const handleExportPackagePdf = () => {
    exportCompletePackagePdf(
      companyProfile,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule,
      effectiveNotes
    );
  };

  // Handlers: Export Previous & Projected Exact Format Suites
  const handleExportPreviousExcel = () => exportPreviousFinancialsExcel(companyProfile, previousFormatData, effectiveNotes);
  const handleExportPreviousWord = () => exportPreviousFinancialsWord(companyProfile, previousFormatData, effectiveNotes);
  const handleExportPreviousPdf = () => exportPreviousFinancialsPdf(companyProfile, previousFormatData, effectiveNotes);
  const handleExportProjectedExcel = () => exportProjectedFinancialsExcel(companyProfile, projectedFormatData);
  const handleExportProjectedWord = () => exportProjectedFinancialsWord(companyProfile, projectedFormatData);
  const handleExportProjectedPdf = () => exportProjectedFinancialsPdf(companyProfile, projectedFormatData);

  const selectedHistYears = activeHistoricalInputs.map(y => y.year);
  const selectedProjYears = generatedProjected.map(y => y.year);

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      {/* 3-Zone Header Contract */}
      <HeaderNav
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        companyProfile={companyProfile}
        onOpenProfileModal={() => setIsProfileModalOpen(true)}
        onOpenDeveloperModal={() => setIsDeveloperModalOpen(true)}
        onExportPackageExcel={handleExportPackageExcel}
        onExportPackageWord={handleExportPackageWord}
        onExportPackagePdf={handleExportPackagePdf}
        onExportPreviousExcel={handleExportPreviousExcel}
        onExportPreviousWord={handleExportPreviousWord}
        onExportPreviousPdf={handleExportPreviousPdf}
        onExportProjectedExcel={handleExportProjectedExcel}
        onExportProjectedWord={handleExportProjectedWord}
        onExportProjectedPdf={handleExportProjectedPdf}
      />

      {/* Main Workspace Body */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        {/* First Look Developer Showcase Card with Interactive Horizon & Year Selector */}
        <DeveloperFirstLookCard
          assumptions={assumptions}
          onUpdateAssumptions={handleUpdateAssumptions}
          onOpenDeveloperModal={() => setIsDeveloperModalOpen(true)}
          companyProfile={companyProfile}
          baseYear={baseYearInput.year}
          selectedHistoricalYears={selectedHistYears}
          selectedProjectedYears={selectedProjYears}
          onToggleHistoricalYear={handleToggleHistoricalYear}
          onToggleProjectedYear={handleToggleProjectedYear}
          onSetHistoricalYears={handleSetHistoricalYears}
          onSetProjectedYears={handleSetProjectedYears}
        />

        {activeTab === 'statements' && (
          <StatementsViewer
            companyProfile={companyProfile}
            incomeData={incomeStatements}
            balanceData={balanceSheets}
            cashFlowData={cashFlowStatements}
            equityData={equityStatements}
            fixedAssetsData={fixedAssetsSchedule}
            notes={effectiveNotes}
            statementConfig={statementConfig}
            assumptions={assumptions}
            onUpdateAssumptions={handleUpdateAssumptions}
            cellOverrides={cellOverrides}
            labelOverrides={labelOverrides}
            onUpdateStatementValue={handleUpdateStatementValue}
            onUpdateStatementLabel={handleUpdateStatementLabel}
            onResetOverrides={handleResetOverrides}
            onUpdateNotes={setCustomNotes}
            previousFormatData={previousFormatData}
            projectedFormatData={projectedFormatData}
            formatMode={formatMode}
            onChangeFormatMode={setFormatMode}
            validationIssues={validationIssues}
            selectedHistoricalYears={selectedHistYears}
            selectedProjectedYears={selectedProjYears}
            onToggleHistoricalYear={handleToggleHistoricalYear}
            onToggleProjectedYear={handleToggleProjectedYear}
            onSetHistoricalYears={handleSetHistoricalYears}
            onSetProjectedYears={handleSetProjectedYears}
            onUpdatePrimaryDriver={handleUpdatePrimaryDriver}
          />
        )}

        {activeTab === 'data-entry' && (
          <FinancialDataEntry
            yearsData={allYearsInputs}
            onUpdatePrimaryDriver={handleUpdatePrimaryDriver}
            onUpdateYearData={handleUpdateYearData}
            companyProfile={companyProfile}
            assumptions={assumptions}
            onUpdateAssumptions={handleUpdateAssumptions}
            inputMode={inputMode}
            onToggleInputMode={setInputMode}
            onGenerateAutomaticProjections={handleGenerateAutomaticProjections}
            incomeStatements={incomeStatements}
            balanceSheets={balanceSheets}
            cashFlowStatements={cashFlowStatements}
            autoFillConfigs={autoFillConfigs}
            onApplyAutoFill={handleApplyAutoFill}
            onClearAutoFill={handleClearAutoFill}
            validationIssues={validationIssues}
            selectedHistoricalYears={selectedHistYears}
            selectedProjectedYears={selectedProjYears}
            onToggleHistoricalYear={handleToggleHistoricalYear}
            onToggleProjectedYear={handleToggleProjectedYear}
            onSetHistoricalYears={handleSetHistoricalYears}
            onSetProjectedYears={handleSetProjectedYears}
          />
        )}

        {activeTab === 'configuration' && (
          <StatementConfigView
            configItems={statementConfig}
            onUpdateConfig={setStatementConfig}
            onResetDefaults={() => setStatementConfig(DEFAULT_STATEMENT_CONFIG_ITEMS)}
          />
        )}

        {activeTab === 'projections' && (
          <ProjectionEngineView
            assumptions={assumptions}
            onUpdateAssumptions={handleUpdateAssumptions}
            onApplyAndRecalculate={() => setActiveTab('statements')}
            onResetDefaults={() => setAssumptions(DEFAULT_PROJECTION_ASSUMPTIONS)}
            selectedHistoricalYears={selectedHistYears}
            selectedProjectedYears={selectedProjYears}
            onToggleHistoricalYear={handleToggleHistoricalYear}
            onToggleProjectedYear={handleToggleProjectedYear}
            onSetHistoricalYears={handleSetHistoricalYears}
            onSetProjectedYears={handleSetProjectedYears}
          />
        )}

        {activeTab === 'reconciliation' && (
          <CashFlowReconciliationAudit
            cashFlowYears={cashFlowStatements}
            balanceSheetYears={balanceSheets}
            companyProfile={companyProfile}
            validationIssues={validationIssues}
          />
        )}
      </main>

      {/* Modals */}
      <CompanyProfileModal
        isOpen={isProfileModalOpen}
        onClose={() => setIsProfileModalOpen(false)}
        profile={companyProfile}
        onSave={setCompanyProfile}
      />

      <DeveloperProfileModal
        isOpen={isDeveloperModalOpen}
        onClose={() => setIsDeveloperModalOpen(false)}
      />

      {/* Quiet Professional Footer */}
      <footer className="border-t border-slate-900 bg-slate-950 py-6 text-xs text-slate-400">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-slate-300">Financials with Tuhin</span>
            <span aria-hidden="true">·</span>
            <span>Section 76 Compliance Architecture</span>
            <span aria-hidden="true">·</span>
            <span>Five Standard Statements & Notes</span>
          </div>

          <div className="flex items-center gap-4 text-[11px]">
            <button
              onClick={() => setIsDeveloperModalOpen(true)}
              className="text-emerald-400 hover:text-emerald-300 transition-colors"
            >
              Developer: Mohammad Towhidul Akbar Tuhin (Officer, Sammilito Islami Bank PLC)
            </button>
            <span aria-hidden="true">·</span>
            <span>BBA (Heriot-Watt, UK) · MBA (CU)</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
