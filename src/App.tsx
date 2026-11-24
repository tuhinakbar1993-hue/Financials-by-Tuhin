/**
 * Financials with Tuhin
 * Section 76: Statement of Cash Flows and Expanded Financial Statement Generation & Export Facility
 * Lead Developer & Financial Architect: Mohammad Towhidul Akbar Tuhin
 * Officer, Sammilito Islami Bank PLC (Transferee of Assets & Liabilities of Union Bank PLC)
 * BBA, Heriot Watt University, UK · MBA, University of Chittagong
 */

import React, { useState, useMemo, useEffect, useCallback } from 'react';
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
  buildPreviousComparativePairs,
  buildPreviousPairScopedNotes,
  buildProjectedFormatData,
  validateFinancialModel,
  calculateDebtServiceCoverage,
  implementRepaymentScheduleIntoModel,
  ImplementScheduleResult,
  STATEMENT_CONFIG_TO_OVERRIDE_KEY,
  FIXED_ASSET_CATEGORY_CONFIG_IDS,
} from './services/financialEngine';
import { rebuildScheduleSummaries } from './services/repaymentScheduleParser';
import {
  exportCompletePackageExcel,
  exportCompletePackageWord,
  exportCompletePackagePdf,
  exportPreviousFinancialsExcel,
  exportPreviousFinancialsWord,
  exportPreviousFinancialsPdf,
  exportAllPreviousComparativePairsExcel,
  exportAllPreviousComparativePairsWord,
  exportAllPreviousComparativePairsPdf,
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
  ParsedRepaymentSchedule,
  RepaymentScheduleSnapshot,
  PreviousFormatYearData,
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
  const [noteEdits, setNoteEdits] = useState<Record<number, { title?: string; textDescription?: string }>>({});

  // Modals state
  const [isProfileModalOpen, setIsProfileModalOpen] = useState(false);
  const [isDeveloperModalOpen, setIsDeveloperModalOpen] = useState(false);

  // Repayment Schedule Upload, Implementation & Pre-Schedule Snapshot State
  const [parsedSchedule, setParsedSchedule] = useState<ParsedRepaymentSchedule | null>(null);
  const [scheduleSnapshot, setScheduleSnapshot] = useState<RepaymentScheduleSnapshot | null>(null);
  const [isScheduleImplemented, setIsScheduleImplemented] = useState<boolean>(false);
  const [lastImplementationResult, setLastImplementationResult] = useState<ImplementScheduleResult | null>(null);
  const [openScheduleStudio, setOpenScheduleStudio] = useState<boolean>(false);

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
    return raw.map((y, idx) => {
      const resolved = customProjectedOverrides[y.year] || y;
      if (idx === 0 && (resolved.financingProceedsBorrowings || 0) !== 0) {
        const removed = resolved.financingProceedsBorrowings || 0;
        return {
          ...resolved,
          financingProceedsBorrowings: 0,
          cashFlowFromFinancing: Math.round(((resolved.cashFlowFromFinancing || 0) - removed + Number.EPSILON) * 100) / 100,
        };
      }
      return resolved;
    });
  }, [baseYearInput, assumptions, customProjectedOverrides]);

  // Combined All Years (Selected Historical + Selected Projected)
  const allYearsInputs = useMemo(() => {
    return [...activeHistoricalInputs, ...generatedProjected];
  }, [activeHistoricalInputs, generatedProjected]);

  const calcOptions = useMemo(
    () => ({ statementConfig, cellOverrides }),
    [statementConfig, cellOverrides]
  );

  // Five Standard Financial Statements + Schedule of Fixed Assets Calculations
  const fixedAssetsSchedule = useMemo(() => {
    return calculateScheduleOfFixedAssets(allYearsInputs, calcOptions);
  }, [allYearsInputs, calcOptions]);

  const incomeStatements = useMemo(() => {
    return calculateIncomeStatement(allYearsInputs, fixedAssetsSchedule, calcOptions);
  }, [allYearsInputs, fixedAssetsSchedule, calcOptions]);

  const cashFlowStatements = useMemo(() => {
    return calculateCashFlowStatement(allYearsInputs, incomeStatements, fixedAssetsSchedule, calcOptions);
  }, [allYearsInputs, incomeStatements, fixedAssetsSchedule, calcOptions]);

  const equityStatements = useMemo(() => {
    return calculateEquityStatement(allYearsInputs, incomeStatements, calcOptions);
  }, [allYearsInputs, incomeStatements, calcOptions]);

  const balanceSheets = useMemo(() => {
    return calculateBalanceSheet(
      allYearsInputs,
      equityStatements,
      fixedAssetsSchedule,
      cashFlowStatements,
      calcOptions,
      incomeStatements
    );
  }, [allYearsInputs, equityStatements, fixedAssetsSchedule, cashFlowStatements, calcOptions, incomeStatements]);

  const defaultNotes = useMemo(() => {
    return generateNotesToFinancialStatements(
      companyProfile,
      allYearsInputs.map(y => y.year),
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      fixedAssetsSchedule,
      { statementConfig, labelOverrides, cellOverrides }
    );
  }, [companyProfile, allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, fixedAssetsSchedule, statementConfig, labelOverrides, cellOverrides]);

  const effectiveNotes = useMemo(() => {
    return defaultNotes.map(note => {
      const edit = noteEdits[note.noteNumber];
      return edit ? { ...note, ...edit } : note;
    });
  }, [defaultNotes, noteEdits]);

  // Exact Compliance Format Datasets (Previous Financials & Projected Financials)
  const previousFormatData = useMemo(() => {
    return buildPreviousFormatData(
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule,
      calcOptions
    );
  }, [allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, equityStatements, fixedAssetsSchedule, calcOptions]);

  const previousComparativePairs = useMemo(() => {
    return buildPreviousComparativePairs(previousFormatData);
  }, [previousFormatData]);

  const getPairScopedNotes = useCallback((pairData: PreviousFormatYearData[]) => {
    return buildPreviousPairScopedNotes(
      companyProfile,
      pairData.map(d => d.year),
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      fixedAssetsSchedule,
      effectiveNotes,
      defaultNotes,
      { statementConfig, labelOverrides, cellOverrides }
    );
  }, [
    companyProfile,
    allYearsInputs,
    incomeStatements,
    balanceSheets,
    cashFlowStatements,
    fixedAssetsSchedule,
    effectiveNotes,
    defaultNotes,
    statementConfig,
    labelOverrides,
    cellOverrides,
  ]);

  const projectedFormatData = useMemo(() => {
    return buildProjectedFormatData(
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule,
      calcOptions
    );
  }, [allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, equityStatements, fixedAssetsSchedule, calcOptions]);

  // Model-Wide Accounting & Fixed-Driver Validation
  const validationIssues = useMemo(() => {
    return validateFinancialModel(
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule,
      calcOptions
    );
  }, [allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, equityStatements, fixedAssetsSchedule, calcOptions]);

  // Debt Service Coverage & Repayment Capacity Analysis
  const debtServiceCoverage = useMemo(() => {
    return calculateDebtServiceCoverage(
      allYearsInputs,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      isScheduleImplemented ? parsedSchedule : null,
      calcOptions
    );
  }, [allYearsInputs, incomeStatements, balanceSheets, cashFlowStatements, isScheduleImplemented, parsedSchedule, calcOptions]);

  // Synchronize parsed schedule year-wise aggregation when FY end or model years change
  const modelYearsKey = useMemo(() => allYearsInputs.map(y => y.year).join(','), [allYearsInputs]);
  useEffect(() => {
    if (!parsedSchedule) return;
    const modelYears = modelYearsKey ? modelYearsKey.split(',').map(Number) : [];
    const rebuilt = rebuildScheduleSummaries(
      parsedSchedule.rows,
      parsedSchedule.metadata,
      companyProfile.financialYearEnd,
      modelYears
    );
    const prevKey = JSON.stringify({
      sums: parsedSchedule.annualSummaries,
      warns: parsedSchedule.validationWarnings,
    });
    const nextKey = JSON.stringify({
      sums: rebuilt.annualSummaries,
      warns: rebuilt.validationWarnings,
    });
    if (prevKey !== nextKey) {
      setParsedSchedule(rebuilt);
    }
  }, [companyProfile.financialYearEnd, modelYearsKey, parsedSchedule]);

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

  // Handler: Update Year Input (also clears matching statement cellOverrides so Data Entry edits take precedence)
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
    const keysToClear = [
      'income_revenue',
      'income_financeCost',
      'income_netProfit',
      'bs_closingInventory',
      'bs_tradeReceivables',
      'bs_otherReceivables',
      'bs_otherCurrentAssets',
      'bs_intangibleAssets',
      'bs_tradePayables',
      'bs_otherPayables',
      'bs_shortTermBorrowings',
      'bs_longTermBorrowings',
      'bs_shareCapital',
      'bs_cashAndEquivalents',
      'cf_directlyEnteredFinancingCashFlow',
      'cf_repaymentOfBorrowings',
      'cf_financingRepaymentBorrowings',
      'cf_financingProceedsBorrowings',
      'cf_financingProceedsShares',
      'cf_adjFinanceCost',
      'cf_financeCostPaid',
      'eq_dividendsPaid',
    ];
    setCellOverrides(prev => {
      let changed = false;
      const next = { ...prev };
      for (const k of keysToClear) {
        if (next[k]?.[updated.year] !== undefined) {
          const yrCopy = { ...next[k] };
          delete yrCopy[updated.year];
          if (Object.keys(yrCopy).length === 0) delete next[k];
          else next[k] = yrCopy;
          changed = true;
        }
      }
      return changed ? next : prev;
    });
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

    // 3. Preserve schedule-locked years for financeCost / repaymentOfBorrowings
    const guardedYears = updatedAllYears.map((y, idx) => {
      const orig = allYearsInputs[idx];
      if (orig?.isScheduleLocked && (driver === 'financeCost' || driver === 'repaymentOfBorrowings')) {
        return orig;
      }
      return y;
    });

    // 4. Update both historical inputs store and custom projected overrides
    setHistoricalYearsStore(prev => {
      const nextStore = { ...prev };
      guardedYears.filter(y => !y.isProjected).forEach(y => {
        nextStore[y.year] = y;
      });
      return nextStore;
    });
    setCustomProjectedOverrides(prev => {
      const nextOverrides = { ...prev };
      guardedYears.filter(y => y.isProjected).forEach(y => {
        nextOverrides[y.year] = y;
      });
      return nextOverrides;
    });

    // 5. Save active Auto-Fill configuration for dynamic recalculation
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

    const yearIdx = allYearsInputs.findIndex(y => y.year === year);
    if (yearIdx === -1) return;

    const currentYearObj = allYearsInputs[yearIdx];
    if (currentYearObj.isScheduleLocked && (driver === 'financeCost' || driver === 'repaymentOfBorrowings')) {
      return;
    }

    // Dynamic Recalculation when Auto-Fill anchor base year is changed
    if (isPrimaryDriver) {
      const driverKey = driver as PrimaryDriverKey;
      const activeConfig = autoFillConfigs[driverKey];
      if (activeConfig && activeConfig.isActive && activeConfig.baseYear === year) {
        handleApplyAutoFill(driverKey, year, val, activeConfig.growthRatePct);
        return;
      }
    }

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

  // Handler: Implement Uploaded Repayment Schedule & Optimize Financial Statements
  const handleImplementRepaymentSchedule = (options: {
    autoIncludeScheduleYears: boolean;
    alignHistoricalScale: boolean;
  }) => {
    if (!parsedSchedule || !parsedSchedule.isValid) return;

    if (!scheduleSnapshot) {
      setScheduleSnapshot({
        capturedAt: new Date().toISOString(),
        historicalYearsStore: JSON.parse(JSON.stringify(historicalYearsStore)),
        customProjectedOverrides: JSON.parse(JSON.stringify(customProjectedOverrides)),
        assumptions: JSON.parse(JSON.stringify(assumptions)),
        autoFillConfigs: JSON.parse(JSON.stringify(autoFillConfigs)),
        companyProfile: JSON.parse(JSON.stringify(companyProfile)),
        cellOverrides: JSON.parse(JSON.stringify(cellOverrides)),
      });
    }

    const res = implementRepaymentScheduleIntoModel({
      schedule: parsedSchedule,
      historicalYearsStore,
      customProjectedOverrides,
      assumptions,
      autoIncludeScheduleYears: options.autoIncludeScheduleYears,
      alignHistoricalScale: options.alignHistoricalScale,
    });

    setHistoricalYearsStore(res.historicalYearsStore);
    setCustomProjectedOverrides(res.customProjectedOverrides);
    setAssumptions(res.assumptions);
    setInputMode('automatic');

    // Clear autoFillConfigs on schedule-driven metrics so schedule locks remain sovereign
    setAutoFillConfigs(prev => {
      const next = { ...prev };
      delete next.financeCost;
      delete next.repaymentOfBorrowings;
      return next;
    });

    // Clear stale statement cellOverrides for years covered by the implemented schedule
    const coveredSet = new Set(res.coveredYears);
    setCellOverrides(prev => {
      const next: StatementCellOverrides = {};
      for (const [k, yrMap] of Object.entries(prev)) {
        const filteredYrMap: Record<number, number> = {};
        for (const [yrStr, val] of Object.entries(yrMap)) {
          if (!coveredSet.has(Number(yrStr))) {
            filteredYrMap[Number(yrStr)] = val;
          }
        }
        if (Object.keys(filteredYrMap).length > 0) {
          next[k] = filteredYrMap;
        }
      }
      return next;
    });

    setIsScheduleImplemented(true);
    setLastImplementationResult(res);
  };

  // Handler: Revert to Pre-Schedule Snapshot
  const handleRevertScheduleSnapshot = () => {
    if (!scheduleSnapshot) return;
    setHistoricalYearsStore(scheduleSnapshot.historicalYearsStore);
    setCustomProjectedOverrides(scheduleSnapshot.customProjectedOverrides);
    setAssumptions(scheduleSnapshot.assumptions);
    setAutoFillConfigs(scheduleSnapshot.autoFillConfigs);
    setCompanyProfile(scheduleSnapshot.companyProfile);
    setCellOverrides(scheduleSnapshot.cellOverrides);
    setScheduleSnapshot(null);
    setIsScheduleImplemented(false);
    setLastImplementationResult(null);
  };

  // Handler: Explicitly Unlock Schedule Protection for a Single Financial Year
  const handleUnlockScheduleYear = (year: number) => {
    const target = allYearsInputs.find(y => y.year === year);
    if (!target) return;
    const unlocked: YearDataInput = {
      ...target,
      isScheduleLocked: false,
      schedulePrincipalLocked: undefined,
      scheduleFinanceCostLocked: undefined,
    };
    if (!target.isProjected) {
      setHistoricalYearsStore(prev => ({ ...prev, [year]: { ...unlocked, isProjected: false } }));
    } else {
      setCustomProjectedOverrides(prev => ({ ...prev, [year]: unlocked }));
    }
  };

  const handleJumpToRepaymentSchedule = () => {
    setOpenScheduleStudio(true);
    setActiveTab('data-entry');
  };

  // Handler: Reset custom projected overrides and generate automatic interconnected projection
  const handleGenerateAutomaticProjections = () => {
    setCustomProjectedOverrides({});
    setInputMode('automatic');
  };

  // Handler: Direct inline editing of ANY statement cell value
  const handleUpdateStatementValue = (statement: string, field: string, year: number, val: number) => {
    const yearObj = allYearsInputs.find(y => y.year === year);

    // Custom statement item value edit (id passed as field when statement === 'custom' or field starts with 'custom_')
    if (statement === 'custom' || field.startsWith('custom_')) {
      const targetId = field;
      setStatementConfig(prev =>
        prev.map(item =>
          item.id === targetId
            ? {
                ...item,
                customValues: {
                  ...(item.customValues || {}),
                  [year]: item.sign === 'negative' ? Math.abs(val) : val,
                },
              }
            : item
        )
      );
      return;
    }

    const fieldMapping: Record<string, keyof YearDataInput> = {
      revenue: 'salesRevenue',
      salesRevenue: 'salesRevenue',
      financeCost: 'financeCost',
      adjFinanceCost: 'financeCost',
      netProfit: 'netProfit',
      netProfitForYear: 'netProfit',
      cashFlowFromFinancing: 'cashFlowFromFinancing',
      directlyEnteredFinancingCashFlow: 'cashFlowFromFinancing',
      repaymentOfBorrowings: 'repaymentOfBorrowings',
      financingRepaymentBorrowings: 'repaymentOfBorrowings',
      financingProceedsBorrowings: 'financingProceedsBorrowings',
      financingProceedsShares: 'financingProceedsShares',
      shareCapitalIssued: 'financingProceedsShares',
      closingInventory: 'closingInventory',
      tradeReceivables: 'tradeReceivables',
      otherReceivables: 'otherReceivables',
      otherCurrentAssets: 'otherCurrentAssets',
      intangibleAssets: 'intangibleAssets',
      tradePayables: 'tradePayables',
      otherPayables: 'otherPayables',
      shortTermBorrowings: 'shortTermBorrowings',
      longTermBorrowings: 'longTermBorrowings',
      shareCapital: 'shareCapital',
      dividendsPaid: 'dividendsPaid',
    };

    const targetField = fieldMapping[field];
    const firstProjYear = generatedProjected[0]?.year;

    // Restrict fresh Loan Received / Proceeds from Borrowings in the first projected year (existing loan rescheduling)
    if (targetField === 'financingProceedsBorrowings' && firstProjYear !== undefined && year === firstProjYear && val !== 0) {
      return;
    }

    // Check Repayment Schedule lock BEFORE modifying state
    if (
      yearObj?.isScheduleLocked &&
      (targetField === 'repaymentOfBorrowings' || targetField === 'financeCost' || field === 'financeCostPaid')
    ) {
      return;
    }

    const key = `${statement}_${field}`;

    if (targetField && yearObj) {
      const updatedYear: YearDataInput = {
        ...yearObj,
        [targetField]: field === 'dividendsPaid' ? Math.abs(val) : val,
      };

      if (targetField === 'repaymentOfBorrowings') {
        const absVal = Math.abs(val);
        const diff = absVal - (yearObj.repaymentOfBorrowings || 0);
        const effectiveBorrowProceeds =
          firstProjYear !== undefined && year === firstProjYear ? 0 : (updatedYear.financingProceedsBorrowings || 0);
        updatedYear.repaymentOfBorrowings = absVal;
        updatedYear.financingRepaymentBorrowings = -absVal;
        updatedYear.financingProceedsBorrowings = effectiveBorrowProceeds;
        updatedYear.cashFlowFromFinancing =
          (updatedYear.financingProceedsShares || 0) +
          effectiveBorrowProceeds -
          absVal -
          updatedYear.dividendsPaid;
        updatedYear.longTermBorrowings = Math.max(0, (yearObj.longTermBorrowings || 0) - diff);
        updatedYear.cashAndEquivalentsBS = Math.max(0, (yearObj.cashAndEquivalentsBS || 0) - diff);
        updatedYear.financeCost = Math.max(
          200,
          Math.round((updatedYear.longTermBorrowings + updatedYear.shortTermBorrowings) * 0.085)
        );
      }

      handleUpdateYearData(updatedYear);
      return;
    }

    // Derived / statement-only fields (costOfSales, operatingExpenses, depreciationExpense, interestIncome, taxExpense, subtotals, cf adjustments, working capital deltas, eq opening/OCI, bs cash override, etc.)
    setCellOverrides(prev => ({
      ...prev,
      [key]: {
        ...(prev[key] || {}),
        [year]: val,
      },
    }));
  };

  // Handler: Direct inline editing of Fixed Assets category cells
  const handleUpdateFixedAssetsCategory = (
    year: number,
    categoryIndex: number,
    field: string,
    value: number
  ) => {
    const key = `fa_${categoryIndex}_${field}`;
    setCellOverrides(prev => ({
      ...prev,
      [key]: {
        ...(prev[key] || {}),
        [year]: value,
      },
    }));
  };

  // Handler: Direct inline editing of ANY statement item label (syncs with statementConfig)
  const handleUpdateStatementLabel = (statement: string, field: string, newLabel: string) => {
    if (statement === 'custom') {
      setStatementConfig(prev =>
        prev.map(item => (item.id === field ? { ...item, label: newLabel } : item))
      );
      return;
    }

    setLabelOverrides(prev => ({
      ...prev,
      [`${statement}_${field}`]: newLabel,
    }));

    const matchedCfgId = Object.entries(STATEMENT_CONFIG_TO_OVERRIDE_KEY).find(
      ([, v]) => v.stmt === statement && v.field === field
    )?.[0];
    if (matchedCfgId) {
      setStatementConfig(prev =>
        prev.map(item => (item.id === matchedCfgId ? { ...item, label: newLabel } : item))
      );
    } else if (statement === 'fa') {
      const idxNum = Number(field);
      const faCfgId = Number.isInteger(idxNum) ? FIXED_ASSET_CATEGORY_CONFIG_IDS[idxNum] : undefined;
      if (faCfgId) {
        setStatementConfig(prev =>
          prev.map(item => (item.id === faCfgId ? { ...item, label: newLabel } : item))
        );
      }
    }
  };

  // Handler: Update statementConfig and sync any customized labels into labelOverrides
  const handleUpdateStatementConfig = (nextConfig: StatementConfigItem[]) => {
    setStatementConfig(nextConfig);
    setLabelOverrides(prev => {
      const next = { ...prev };
      nextConfig.forEach(item => {
        const mapping = STATEMENT_CONFIG_TO_OVERRIDE_KEY[item.id];
        const defItem = DEFAULT_STATEMENT_CONFIG_ITEMS.find(d => d.id === item.id);
        if (mapping && defItem) {
          const overrideKey = `${mapping.stmt}_${mapping.field}`;
          if (item.label !== defItem.label) {
            next[overrideKey] = item.label;
          } else if (next[overrideKey] && next[overrideKey] === defItem.label) {
            delete next[overrideKey];
          }
          return;
        }
        const faIdx = FIXED_ASSET_CATEGORY_CONFIG_IDS.indexOf(
          item.id as (typeof FIXED_ASSET_CATEGORY_CONFIG_IDS)[number]
        );
        if (faIdx !== -1 && defItem) {
          const overrideKey = `fa_${faIdx}`;
          if (item.label !== defItem.label) {
            next[overrideKey] = item.label;
          } else if (next[overrideKey]) {
            delete next[overrideKey];
          }
        }
      });
      return next;
    });
  };

  const handleResetStatementConfig = () => {
    setStatementConfig(DEFAULT_STATEMENT_CONFIG_ITEMS);
    setLabelOverrides(prev => {
      const next = { ...prev };
      Object.values(STATEMENT_CONFIG_TO_OVERRIDE_KEY).forEach(mapping => {
        delete next[`${mapping.stmt}_${mapping.field}`];
      });
      FIXED_ASSET_CATEGORY_CONFIG_IDS.forEach((_, idx) => {
        delete next[`fa_${idx}`];
      });
      return next;
    });
  };

  // Handler: Update Note title/text edits while keeping live numeric tables & Note 12 calculations dynamic
  const handleUpdateNotes = (updatedNotes: FinancialNoteItem[]) => {
    setNoteEdits(prev => {
      const next = { ...prev };
      updatedNotes.forEach(un => {
        const def = defaultNotes.find(d => d.noteNumber === un.noteNumber);
        if (!def) return;
        const editEntry: { title?: string; textDescription?: string } = {};
        if (un.title !== def.title) editEntry.title = un.title;
        if (un.textDescription !== def.textDescription) editEntry.textDescription = un.textDescription;
        if (editEntry.title !== undefined || editEntry.textDescription !== undefined) {
          next[un.noteNumber] = editEntry;
        } else {
          delete next[un.noteNumber];
        }
      });
      return next;
    });
  };

  // Reset all custom overrides back to pure calculation engine defaults
  const handleResetOverrides = () => {
    setCellOverrides({});
    setLabelOverrides({});
    setNoteEdits({});
    const defaultLabelMap = new Map(DEFAULT_STATEMENT_CONFIG_ITEMS.map(d => [d.id, d.label]));
    setStatementConfig(prev =>
      prev.map(item =>
        !item.isCustom && defaultLabelMap.has(item.id)
          ? { ...item, label: defaultLabelMap.get(item.id)! }
          : item
      )
    );
  };

  const exportSyncOptions = useMemo(
    () => ({ statementConfig, labelOverrides, cellOverrides }),
    [statementConfig, labelOverrides, cellOverrides]
  );

  // Handler: Export Complete Package (Section 76)
  const handleExportPackageExcel = () => {
    exportCompletePackageExcel(
      companyProfile,
      incomeStatements,
      balanceSheets,
      cashFlowStatements,
      equityStatements,
      fixedAssetsSchedule,
      effectiveNotes,
      undefined,
      isScheduleImplemented ? debtServiceCoverage : undefined,
      exportSyncOptions
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
      effectiveNotes,
      undefined,
      isScheduleImplemented ? debtServiceCoverage : undefined,
      exportSyncOptions
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
      effectiveNotes,
      undefined,
      isScheduleImplemented ? debtServiceCoverage : undefined,
      exportSyncOptions
    );
  };

  // Handlers: Export Previous Comparative Pairs & Projected Exact Format Suites
  const handleExportPreviousExcel = () =>
    exportAllPreviousComparativePairsExcel(
      companyProfile,
      previousComparativePairs,
      getPairScopedNotes,
      exportSyncOptions
    );
  const handleExportPreviousWord = () =>
    exportAllPreviousComparativePairsWord(
      companyProfile,
      previousComparativePairs,
      getPairScopedNotes,
      exportSyncOptions
    );
  const handleExportPreviousPdf = () =>
    exportAllPreviousComparativePairsPdf(
      companyProfile,
      previousComparativePairs,
      getPairScopedNotes,
      exportSyncOptions
    );

  const handleExportPreviousPair = (pairId: string, format: 'xlsx' | 'docx' | 'pdf') => {
    const pair = previousComparativePairs.find(p => p.id === pairId);
    if (!pair) return;
    const scopedNotes = getPairScopedNotes(pair.data);
    if (format === 'xlsx')
      exportPreviousFinancialsExcel(companyProfile, pair.data, scopedNotes, exportSyncOptions);
    else if (format === 'docx')
      void exportPreviousFinancialsWord(companyProfile, pair.data, scopedNotes, exportSyncOptions);
    else exportPreviousFinancialsPdf(companyProfile, pair.data, scopedNotes, exportSyncOptions);
  };

  const handleExportProjectedExcel = () =>
    exportProjectedFinancialsExcel(companyProfile, projectedFormatData, exportSyncOptions);
  const handleExportProjectedWord = () =>
    exportProjectedFinancialsWord(companyProfile, projectedFormatData, exportSyncOptions);
  const handleExportProjectedPdf = () =>
    exportProjectedFinancialsPdf(companyProfile, projectedFormatData, exportSyncOptions);

  const selectedHistYears = activeHistoricalInputs.map(y => y.year);
  const selectedProjYears = generatedProjected.map(y => y.year);
  const projectedRangeLabel = selectedProjYears.length > 0
    ? (selectedProjYears.length === 1
        ? `${selectedProjYears[0]}`
        : `${Math.min(...selectedProjYears)}–${Math.max(...selectedProjYears)}`)
    : '';

  return (
    <div className="min-h-screen relative text-slate-900 flex flex-col font-sans bg-[#f4f6f8]">
      {/* Fixed Global Architectural Sunlit Backdrop ("Background.png") */}
      <div className="fixed inset-0 z-0 pointer-events-none overflow-hidden" aria-hidden="true">
        <img
          src="/assets/aistudio/Background_be955592.png"
          alt=""
          className="w-full h-full object-cover object-center select-none"
          onError={(e) => {
            (e.target as HTMLImageElement).style.display = 'none';
          }}
        />
        <div className="absolute inset-0 bg-gradient-to-br from-white/35 via-white/15 to-emerald-50/15" />
      </div>

      <div className="relative z-10 min-h-screen flex flex-col">
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
          previousComparativePairs={previousComparativePairs}
          onExportPreviousPair={handleExportPreviousPair}
          onExportPreviousExcel={handleExportPreviousExcel}
          onExportPreviousWord={handleExportPreviousWord}
          onExportPreviousPdf={handleExportPreviousPdf}
          projectedYearsRangeLabel={projectedRangeLabel}
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
              onUpdateStatementConfig={handleUpdateStatementConfig}
              onResetStatementConfig={handleResetStatementConfig}
              assumptions={assumptions}
              onUpdateAssumptions={handleUpdateAssumptions}
              cellOverrides={cellOverrides}
              labelOverrides={labelOverrides}
              onUpdateStatementValue={handleUpdateStatementValue}
              onUpdateStatementLabel={handleUpdateStatementLabel}
              onUpdateFixedAssetsCategory={handleUpdateFixedAssetsCategory}
              onResetOverrides={handleResetOverrides}
              onUpdateNotes={handleUpdateNotes}
              previousFormatData={previousFormatData}
              previousComparativePairs={previousComparativePairs}
              getPairScopedNotes={getPairScopedNotes}
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
              debtServiceCoverage={debtServiceCoverage}
              isScheduleImplemented={isScheduleImplemented}
              parsedSchedule={parsedSchedule}
              onJumpToRepaymentSchedule={handleJumpToRepaymentSchedule}
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
              parsedSchedule={parsedSchedule}
              onSetParsedSchedule={setParsedSchedule}
              scheduleSnapshot={scheduleSnapshot}
              isScheduleImplemented={isScheduleImplemented}
              lastImplementationResult={lastImplementationResult}
              debtServiceCoverage={debtServiceCoverage}
              onImplementRepaymentSchedule={handleImplementRepaymentSchedule}
              onRevertScheduleSnapshot={handleRevertScheduleSnapshot}
              onUnlockScheduleYear={handleUnlockScheduleYear}
              initialShowScheduleStudio={openScheduleStudio || isScheduleImplemented || parsedSchedule !== null}
            />
          )}

          {activeTab === 'configuration' && (
            <StatementConfigView
              configItems={statementConfig}
              onUpdateConfig={handleUpdateStatementConfig}
              onResetDefaults={handleResetStatementConfig}
              years={allYearsInputs.map(y => y.year)}
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
              debtServiceCoverage={debtServiceCoverage}
              isScheduleImplemented={isScheduleImplemented}
              parsedSchedule={parsedSchedule}
              onJumpToRepaymentSchedule={handleJumpToRepaymentSchedule}
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
        <footer className="border-t border-slate-200/90 bg-white/80 backdrop-blur-md py-5 text-xs text-slate-600 mt-8">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex flex-wrap items-center gap-2">
              <span className="font-bold text-slate-900">Financials with Tuhin</span>
              <span aria-hidden="true">·</span>
              <span>Section 76 Compliance Architecture</span>
              <span aria-hidden="true">·</span>
              <span>Five Standard Statements & Notes</span>
            </div>

            <div className="flex flex-wrap items-center gap-4 text-[11px]">
              <button
                onClick={() => setIsDeveloperModalOpen(true)}
                className="font-semibold text-emerald-700 hover:text-emerald-600 transition-colors"
              >
                Developer: Mohammad Towhidul Akbar Tuhin (Officer, Sammilito Islami Bank PLC)
              </button>
              <span aria-hidden="true">·</span>
              <span className="text-slate-500">BBA (Heriot-Watt, UK) · MBA (CU)</span>
            </div>
          </div>
        </footer>
      </div>
    </div>
  );
}
