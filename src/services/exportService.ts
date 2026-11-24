import * as XLSX from 'xlsx';
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  TextRun,
  HeadingLevel,
  AlignmentType,
  WidthType,
  BorderStyle,
  PageBreak,
} from 'docx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  CompanyProfile,
  IncomeStatementYear,
  BalanceSheetYear,
  CashFlowStatementYear,
  EquityStatementYear,
  ScheduleOfFixedAssetsYear,
  FinancialNoteItem,
  PreviousFormatYearData,
  ProjectedFormatYearData,
  DebtServiceCoverageYear,
  StatementConfigItem,
  StatementLabelOverrides,
} from '../types/financial';
import {
  FIXED_ASSET_CATEGORY_CONFIG_IDS,
  getCustomItemYearValue,
  resolveEffectiveLabel,
} from './financialEngine';

export interface ExportSyncOptions {
  statementConfig?: StatementConfigItem[];
  labelOverrides?: StatementLabelOverrides;
}

function resolveFilenameAndOptions(
  defaultFilename: string,
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions
): { filename: string; options?: ExportSyncOptions } {
  if (typeof filenameOrOptions === 'string') {
    return { filename: filenameOrOptions, options: explicitOptions };
  }
  return { filename: defaultFilename, options: filenameOrOptions ?? explicitOptions };
}

function isVis(options: ExportSyncOptions | undefined, configId: string): boolean {
  const item = options?.statementConfig?.find(c => c.id === configId);
  return item ? item.isVisible : true;
}

function lbl(
  options: ExportSyncOptions | undefined,
  stmt: string,
  field: string,
  defaultLabel: string
): string {
  return resolveEffectiveLabel(stmt, field, defaultLabel, options?.labelOverrides, options?.statementConfig);
}

function getCustomRows(
  options: ExportSyncOptions | undefined,
  statement: StatementConfigItem['statement'],
  section?: string
): StatementConfigItem[] {
  return (options?.statementConfig || []).filter(
    item =>
      item.isCustom &&
      item.isVisible &&
      item.statement === statement &&
      (!section || item.section === section)
  );
}

function getUnscopedCustomIncomeRows(options: ExportSyncOptions | undefined): StatementConfigItem[] {
  return (options?.statementConfig || []).filter(
    item =>
      item.isCustom &&
      item.isVisible &&
      item.statement === 'income' &&
      item.section !== 'Revenue & Gross Margin' &&
      item.section !== 'Operating Expenses' &&
      item.section !== 'Non-Operating & Tax'
  );
}

/**
 * Formats a list of 1-indexed Excel row numbers into a compact formula expression
 * e.g. [7,8,9,10] -> "SUM(B7:B10)", [7,9] -> "B7+B9", [] -> "0"
 */
function excelSumRows(col: string, rowNums: number[]): string {
  if (rowNums.length === 0) return '0';
  const sorted = [...rowNums].sort((a, b) => a - b);
  const parts: string[] = [];
  let start = sorted[0];
  let end = sorted[0];
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === end + 1) {
      end = sorted[i];
    } else {
      parts.push(start === end ? `${col}${start}` : `SUM(${col}${start}:${col}${end})`);
      start = sorted[i];
      end = sorted[i];
    }
  }
  parts.push(start === end ? `${col}${start}` : `SUM(${col}${start}:${col}${end})`);
  return parts.join('+');
}

function excelDiffExpr(col: string, addRows: number[], subRows: number[]): string {
  const addPart = addRows.length > 0 ? excelSumRows(col, addRows) : '';
  const subPart = subRows.map(r => `-${col}${r}`).join('');
  if (!addPart && !subPart) return '0';
  if (!addPart) return `0${subPart}`;
  return `${addPart}${subPart}`;
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------- EXCEL FORMULA & WORKSHEET HELPERS ---------------- //

type ExcelCellValue = string | number | XLSX.CellObject;

function colLetter(colIdxZeroBased: number): string {
  return XLSX.utils.encode_col(colIdxZeroBased);
}

function fmtExportNum(val: number): string {
  if (Math.abs(val) < 0.005) return '-';
  const absVal = Math.abs(val);
  const hasCents = Math.abs(absVal - Math.round(absVal)) >= 0.005;
  const formatted = hasCents
    ? absVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : Math.round(absVal).toLocaleString('en-US');
  return val < 0 ? `(${formatted})` : formatted;
}

/**
 * Builds an editable Excel formula cell with cached numeric value
 * so spreadsheets evaluate dynamically in Excel, LibreOffice, and Google Sheets.
 */
function fCell(formula: string, cachedValue: number): XLSX.CellObject {
  const num = Number.isFinite(cachedValue) ? cachedValue : 0;
  const hasCents = Math.abs(num - Math.round(num)) >= 0.005;
  return {
    t: 'n',
    v: num,
    f: formula,
    z: hasCents ? '#,##0.00;(#,##0.00);"-"' : '#,##0;(#,##0);"-"',
  };
}

function nCell(value: number): XLSX.CellObject {
  const num = Number.isFinite(value) ? value : 0;
  const hasCents = Math.abs(num - Math.round(num)) >= 0.005;
  return {
    t: 'n',
    v: num,
    z: hasCents ? '#,##0.00;(#,##0.00);"-"' : '#,##0;(#,##0);"-"',
  };
}

function applyColumnWidths(ws: XLSX.WorkSheet, widths: number[]) {
  ws['!cols'] = widths.map(wch => ({ wch }));
}

function createTableDataForStatement(
  title: string,
  company: CompanyProfile,
  headers: string[],
  rows: ExcelCellValue[][]
): ExcelCellValue[][] {
  return [
    [company.companyName.toUpperCase()],
    [`${title.toUpperCase()} (${company.reportingCurrency} in ${company.reportingUnit})`],
    [`Accounting Framework: ${company.accountingFramework} | Entity Type: ${company.entityType.toUpperCase()}`],
    [`Report Generated by Financials with Tuhin | Prepared by: ${company.preparedBy}`],
    [],
    headers,
    ...rows,
  ];
}

// Rows 1..6 are masthead + headers -> first data row is Excel Row 7 (1-indexed)
const STD_DATA_START_ROW = 7;

function buildIncomeStatementSheet(
  company: CompanyProfile,
  data: IncomeStatementYear[],
  title = 'Income Statement',
  options?: ExportSyncOptions
): XLSX.WorkSheet {
  const headers = ['Financial Metric', ...data.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const rows: ExcelCellValue[][] = [];
  const nextRow = () => STD_DATA_START_ROW + rows.length;

  const customRev = getCustomRows(options, 'income', 'Revenue & Gross Margin');
  const customOpex = getCustomRows(options, 'income', 'Operating Expenses');
  const customNonOp = [...getCustomRows(options, 'income', 'Non-Operating & Tax'), ...getUnscopedCustomIncomeRows(options)];

  const gpInputs: number[] = [];
  if (isVis(options, 'inc_revenue')) {
    const revRow = nextRow();
    rows.push([
      lbl(options, 'income', 'revenue', 'Sales Revenue'),
      ...data.map(d => {
        const custSum = customRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
        return nCell(d.revenue - custSum);
      }),
    ]);
    gpInputs.push(revRow);
    customRev.forEach(ci => {
      const cr = nextRow();
      rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
      gpInputs.push(cr);
    });
  } else {
    customRev.forEach(ci => {
      rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
    });
  }

  if (isVis(options, 'inc_costOfSales')) {
    const cosRow = nextRow();
    rows.push([lbl(options, 'income', 'costOfSales', 'Cost of Sales'), ...data.map(d => nCell(-d.costOfSales))]);
    gpInputs.push(cosRow);
  }

  const gpRow = nextRow();
  rows.push([
    lbl(options, 'income', 'grossProfit', 'Gross Profit'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(excelSumRows(c, gpInputs), d.grossProfit);
    }),
  ]);

  const ebitInputs: number[] = [gpRow];
  if (isVis(options, 'inc_operatingExpenses')) {
    const opexRow = nextRow();
    rows.push([
      lbl(options, 'income', 'operatingExpenses', 'Operating Expenses'),
      ...data.map(d => {
        const custSum = customOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
        return nCell(-(d.operatingExpenses + custSum));
      }),
    ]);
    ebitInputs.push(opexRow);
    customOpex.forEach(ci => {
      const cr = nextRow();
      rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
      ebitInputs.push(cr);
    });
  } else {
    customOpex.forEach(ci => {
      rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
    });
  }

  if (isVis(options, 'inc_depreciationExpense')) {
    const depRow = nextRow();
    rows.push([lbl(options, 'income', 'depreciationExpense', 'Depreciation Expense'), ...data.map(d => nCell(-d.depreciationExpense))]);
    ebitInputs.push(depRow);
  }

  const ebitRow = nextRow();
  rows.push([
    lbl(options, 'income', 'operatingProfit', 'Operating Profit (EBIT)'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(excelSumRows(c, ebitInputs), d.operatingProfit);
    }),
  ]);

  const pbtInputs: number[] = [ebitRow];
  if (isVis(options, 'inc_interestIncome')) {
    const intRow = nextRow();
    rows.push([
      lbl(options, 'income', 'interestIncome', 'Interest & Investment Income'),
      ...data.map(d => {
        const custSum = customNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
        return nCell(d.interestIncome - custSum);
      }),
    ]);
    pbtInputs.push(intRow);
    customNonOp.forEach(ci => {
      const cr = nextRow();
      rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
      pbtInputs.push(cr);
    });
  } else {
    customNonOp.forEach(ci => {
      rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
    });
  }

  if (isVis(options, 'inc_financeCost')) {
    const fcRow = nextRow();
    rows.push([lbl(options, 'income', 'financeCost', 'Finance Cost (Mandatory Direct Entry)'), ...data.map(d => nCell(-d.financeCost))]);
    pbtInputs.push(fcRow);
  }

  const pbtRow = nextRow();
  rows.push([
    lbl(options, 'income', 'profitBeforeTax', 'Profit Before Tax (EBT)'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(excelSumRows(c, pbtInputs), d.profitBeforeTax);
    }),
  ]);

  const npInputs: number[] = [pbtRow];
  if (isVis(options, 'inc_taxExpense')) {
    const taxRow = nextRow();
    rows.push([lbl(options, 'income', 'taxExpense', 'Income Tax Expense'), ...data.map(d => nCell(-d.taxExpense))]);
    npInputs.push(taxRow);
  }

  rows.push([
    lbl(options, 'income', 'netProfit', 'Net Profit After Tax'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(excelSumRows(c, npInputs), d.netProfit);
    }),
  ]);

  const sheetData = createTableDataForStatement(title, company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [42, ...data.map(() => 18)]);
  return ws;
}

function buildBalanceSheetWorksheet(
  company: CompanyProfile,
  data: BalanceSheetYear[],
  title = 'Statement of Financial Position',
  options?: ExportSyncOptions
): XLSX.WorkSheet {
  const headers = ['Statement Line Item', ...data.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const rows: ExcelCellValue[][] = [];
  const nextRow = () => STD_DATA_START_ROW + rows.length;

  rows.push(['ASSETS', ...data.map(() => '')]);

  const caRows: number[] = [];
  if (isVis(options, 'bs_cashAndEquivalents')) {
    caRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'cashAndEquivalents', 'Cash and Cash Equivalents'), ...data.map(d => nCell(d.cashAndEquivalents))]);
  }
  if (isVis(options, 'bs_tradeReceivables')) {
    caRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'tradeReceivables', 'Trade Receivables'), ...data.map(d => nCell(d.tradeReceivables))]);
  }
  if (isVis(options, 'bs_otherReceivables')) {
    caRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'otherReceivables', 'Other Receivables & Prepayments'), ...data.map(d => nCell(d.otherReceivables))]);
  }
  if (isVis(options, 'bs_closingInventory')) {
    caRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'closingInventory', 'Closing Inventory'), ...data.map(d => nCell(d.closingInventory))]);
  }
  if (isVis(options, 'bs_otherCurrentAssets')) {
    caRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'otherCurrentAssets', 'Other Current Assets'), ...data.map(d => nCell(d.otherCurrentAssets))]);
  }
  getCustomRows(options, 'balanceSheet', 'Current Assets').forEach(ci => {
    caRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const totalCaRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalCurrentAssets', 'Total Current Assets'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), caRows), d.totalCurrentAssets)),
  ]);

  const ncaRows: number[] = [];
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    ncaRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'propertyPlantEquipment', 'Property, Plant and Equipment (NBV)'), ...data.map(d => nCell(d.propertyPlantEquipment))]);
  }
  if (isVis(options, 'bs_intangibleAssets')) {
    ncaRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'intangibleAssets', 'Intangible Assets'), ...data.map(d => nCell(d.intangibleAssets))]);
  }
  if (isVis(options, 'bs_otherNonCurrentAssets')) {
    ncaRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'otherNonCurrentAssets', 'Other Non-Current Assets'), ...data.map(d => nCell(d.otherNonCurrentAssets))]);
  }
  getCustomRows(options, 'balanceSheet', 'Non-Current Assets').forEach(ci => {
    ncaRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const totalNcaRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalNonCurrentAssets', 'Total Non-Current Assets'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), ncaRows), d.totalNonCurrentAssets)),
  ]);

  const totalAssetsRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalAssets', 'TOTAL ASSETS'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(`${c}${totalCaRow}+${c}${totalNcaRow}`, d.totalAssets);
    }),
  ]);

  rows.push([]);
  rows.push(['LIABILITIES & SHAREHOLDERS EQUITY', ...data.map(() => '')]);

  const clRows: number[] = [];
  if (isVis(options, 'bs_tradePayables')) {
    clRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'tradePayables', 'Trade Payables'), ...data.map(d => nCell(d.tradePayables))]);
  }
  if (isVis(options, 'bs_otherPayables')) {
    clRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'otherPayables', 'Other Current Payables'), ...data.map(d => nCell(d.otherPayables))]);
  }
  if (isVis(options, 'bs_shortTermBorrowings')) {
    clRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'shortTermBorrowings', 'Short-Term Borrowings'), ...data.map(d => nCell(d.shortTermBorrowings))]);
  }
  if (isVis(options, 'bs_currentTaxPayable')) {
    clRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'currentTaxPayable', 'Current Tax Payable'), ...data.map(d => nCell(d.currentTaxPayable))]);
  }
  getCustomRows(options, 'balanceSheet', 'Current Liabilities').forEach(ci => {
    clRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const totalClRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalCurrentLiabilities', 'Total Current Liabilities'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), clRows), d.totalCurrentLiabilities)),
  ]);

  const nclRows: number[] = [];
  if (isVis(options, 'bs_longTermBorrowings')) {
    nclRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'longTermBorrowings', 'Long-Term Borrowings'), ...data.map(d => nCell(d.longTermBorrowings))]);
  }
  if (isVis(options, 'bs_otherNonCurrentLiabilities')) {
    nclRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'otherNonCurrentLiabilities', 'Other Non-Current Liabilities'), ...data.map(d => nCell(d.otherNonCurrentLiabilities))]);
  }
  getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities').forEach(ci => {
    nclRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const totalNclRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalNonCurrentLiabilities', 'Total Non-Current Liabilities'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), nclRows), d.totalNonCurrentLiabilities)),
  ]);

  const totalLiabRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalLiabilities', 'TOTAL LIABILITIES'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(`${c}${totalClRow}+${c}${totalNclRow}`, d.totalLiabilities);
    }),
  ]);

  const eqRows: number[] = [];
  if (isVis(options, 'bs_shareCapital')) {
    eqRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'shareCapital', 'Share Capital'), ...data.map(d => nCell(d.shareCapital))]);
  }
  if (isVis(options, 'bs_retainedEarnings')) {
    eqRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'retainedEarnings', 'Retained Earnings'), ...data.map(d => nCell(d.retainedEarnings))]);
  }
  if (isVis(options, 'bs_reserves')) {
    eqRows.push(nextRow());
    rows.push([lbl(options, 'bs', 'reserves', 'Reserves'), ...data.map(d => nCell(d.reserves))]);
  }
  getCustomRows(options, 'balanceSheet', 'Equity').forEach(ci => {
    eqRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const totalEqRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalEquity', 'TOTAL EQUITY'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), eqRows), d.totalEquity)),
  ]);

  const totalLERow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalLiabilitiesAndEquity', 'TOTAL LIABILITIES & EQUITY'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(`${c}${totalLiabRow}+${c}${totalEqRow}`, d.totalLiabilitiesAndEquity);
    }),
  ]);

  rows.push([
    'Balance Check (Variance)',
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(`${c}${totalAssetsRow}-${c}${totalLERow}`, d.balanceCheck);
    }),
  ]);

  const sheetData = createTableDataForStatement(title, company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [42, ...data.map(() => 18)]);
  return ws;
}

function buildCashFlowWorksheet(
  company: CompanyProfile,
  data: CashFlowStatementYear[],
  title = 'Statement of Cash Flows',
  options?: ExportSyncOptions
): XLSX.WorkSheet {
  const headers = ['Cash Flow Activity / Item', ...data.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const rows: ExcelCellValue[][] = [];
  const nextRow = () => STD_DATA_START_ROW + rows.length;

  rows.push(['CASH FLOWS FROM OPERATING ACTIVITIES (Indirect Method)', ...data.map(() => '')]);
  const nonCashRows: number[] = [];
  if (isVis(options, 'cf_pbt')) {
    nonCashRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'profitBeforeTax', 'Net Profit Before Tax'), ...data.map(d => nCell(d.profitBeforeTax))]);
  }
  if (isVis(options, 'cf_dep')) {
    nonCashRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'adjDepreciation', 'Adjustments for Depreciation'), ...data.map(d => nCell(d.adjDepreciation))]);
  }
  if (isVis(options, 'cf_amort')) {
    nonCashRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'adjAmortization', 'Adjustments for Amortization'), ...data.map(d => nCell(d.adjAmortization))]);
  }
  if (isVis(options, 'cf_finance')) {
    nonCashRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'adjFinanceCost', 'Finance Cost Adjustment'), ...data.map(d => nCell(d.adjFinanceCost))]);
  }
  if (isVis(options, 'cf_interest_inc')) {
    nonCashRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'adjInterestIncome', 'Interest & Investment Income'), ...data.map(d => nCell(d.adjInterestIncome))]);
  }
  if (isVis(options, 'cf_disposal_gain_loss')) {
    nonCashRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'adjLossGainDisposal', 'Loss / (Gain) on Disposal of Fixed Assets'), ...data.map(d => nCell(d.adjLossGainDisposal))]);
  }

  const opBeforeWcRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'operatingProfitBeforeWC', 'Operating Profit Before Working Capital Changes'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), nonCashRows), d.operatingProfitBeforeWC)),
  ]);

  const wcRows: number[] = [opBeforeWcRow];
  if (isVis(options, 'cf_wc_inventory')) {
    wcRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'changeInInventory', '(Increase) / Decrease in Inventory'), ...data.map(d => nCell(d.changeInInventory))]);
  }
  if (isVis(options, 'cf_wc_receivables')) {
    wcRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'changeInReceivables', '(Increase) / Decrease in Trade Receivables'), ...data.map(d => nCell(d.changeInReceivables))]);
  }
  if (isVis(options, 'cf_wc_other_rec')) {
    wcRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'changeInOtherReceivables', '(Increase) / Decrease in Other Receivables'), ...data.map(d => nCell(d.changeInOtherReceivables))]);
  }
  if (isVis(options, 'cf_wc_other_ca')) {
    wcRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'changeInOtherCA', '(Increase) / Decrease in Other Current Assets'), ...data.map(d => nCell(d.changeInOtherCA ?? 0))]);
  }
  if (isVis(options, 'cf_wc_payables')) {
    wcRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'changeInPayables', 'Increase / (Decrease) in Trade Payables'), ...data.map(d => nCell(d.changeInPayables))]);
  }
  if (isVis(options, 'cf_wc_other_pay')) {
    wcRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'changeInOtherPayables', 'Increase / (Decrease) in Other Payables'), ...data.map(d => nCell(d.changeInOtherPayables))]);
  }
  getCustomRows(options, 'cashFlow', 'Operating Activities').forEach(ci => {
    wcRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const cashGenOpsRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'cashGeneratedFromOperations', 'Cash Generated from Operations'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), wcRows), d.cashGeneratedFromOperations)),
  ]);

  const opEndRows: number[] = [cashGenOpsRow];
  if (isVis(options, 'cf_tax_paid')) {
    opEndRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'incomeTaxPaid', 'Income Tax Paid'), ...data.map(d => nCell(d.incomeTaxPaid))]);
  }
  if (isVis(options, 'cf_finance_paid')) {
    opEndRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'financeCostPaid', 'Finance Cost Paid'), ...data.map(d => nCell(d.financeCostPaid))]);
  }

  const netOpCashRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'netCashFromOperating', 'NET CASH GENERATED FROM OPERATING ACTIVITIES'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), opEndRows), d.netCashFromOperating)),
  ]);

  rows.push([]);
  rows.push(['CASH FLOWS FROM INVESTING ACTIVITIES', ...data.map(() => '')]);
  const invRows: number[] = [];
  if (isVis(options, 'cf_pur_ppe')) {
    invRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'purchasePPE', 'Purchase of Property, Plant & Equipment (Capex)'), ...data.map(d => nCell(d.purchasePPE))]);
  }
  if (isVis(options, 'cf_sale_ppe')) {
    invRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'salePPE', 'Proceeds from Sale of Fixed Assets'), ...data.map(d => nCell(d.salePPE))]);
  }
  if (isVis(options, 'cf_pur_intangible')) {
    invRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'purchaseIntangibles', 'Purchase of Intangible Assets'), ...data.map(d => nCell(d.purchaseIntangibles))]);
  }
  if (isVis(options, 'cf_interest_rec')) {
    invRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'interestReceived', 'Interest Received'), ...data.map(d => nCell(d.interestReceived))]);
  }
  if (isVis(options, 'cf_dividends_rec')) {
    invRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'dividendsReceived', 'Dividends Received'), ...data.map(d => nCell(d.dividendsReceived))]);
  }
  getCustomRows(options, 'cashFlow', 'Investing Activities').forEach(ci => {
    invRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const netInvCashRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'netCashFromInvesting', 'NET CASH FROM / (USED IN) INVESTING ACTIVITIES'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), invRows), d.netCashFromInvesting)),
  ]);

  rows.push([]);
  rows.push(['CASH FLOWS FROM FINANCING ACTIVITIES (Mandatory Direct User Input)', ...data.map(() => '')]);

  let directFinRow: number | null = null;
  if (isVis(options, 'cf_direct_financing')) {
    directFinRow = nextRow();
    rows.push([lbl(options, 'cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Financing Cash Flow'), ...data.map(d => nCell(d.directlyEnteredFinancingCashFlow))]);
  }

  const finBreakdownRows: number[] = [];
  if (isVis(options, 'cf_shares_issue')) {
    finBreakdownRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'financingProceedsShares', 'Proceeds from Issue of Share Capital'), ...data.map(d => nCell(d.financingProceedsShares))]);
  }
  if (isVis(options, 'cf_borrow_proceeds')) {
    finBreakdownRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing'), ...data.map(d => nCell(d.financingProceedsBorrowings))]);
  }
  if (isVis(options, 'cf_borrow_repay')) {
    finBreakdownRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing'), ...data.map(d => nCell(-(d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings))))]);
  }
  if (isVis(options, 'cf_dividends_paid')) {
    finBreakdownRows.push(nextRow());
    rows.push([lbl(options, 'cf', 'financingDividendsPaid', 'Dividends Paid'), ...data.map(d => nCell(d.financingDividendsPaid))]);
  }
  const customFinRows = getCustomRows(options, 'cashFlow', 'Financing Activities');
  customFinRows.forEach(ci => {
    finBreakdownRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  const netFinCashRow = nextRow();
  const allStdFinVisible =
    isVis(options, 'cf_shares_issue') &&
    isVis(options, 'cf_borrow_proceeds') &&
    isVis(options, 'cf_borrow_repay') &&
    isVis(options, 'cf_dividends_paid') &&
    customFinRows.length === 0;
  rows.push([
    lbl(options, 'cf', 'netCashFromFinancing', 'NET CASH FROM / (USED IN) FINANCING ACTIVITIES'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      if (directFinRow === null) {
        return fCell(excelSumRows(c, finBreakdownRows), d.netCashFromFinancing);
      }
      if (allStdFinVisible && Math.abs(d.netCashFromFinancing - d.directlyEnteredFinancingCashFlow) < 0.01) {
        return fCell(`${c}${directFinRow}`, d.netCashFromFinancing);
      }
      if (d.isFinancingBreakdownComplete) {
        return fCell(excelSumRows(c, finBreakdownRows), d.netCashFromFinancing);
      }
      return nCell(d.netCashFromFinancing);
    }),
  ]);

  rows.push([]);
  const netChangeRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'netChangeInCash', 'NET INCREASE / (DECREASE) IN CASH & EQUIVALENTS'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(`${c}${netOpCashRow}+${c}${netInvCashRow}+${c}${netFinCashRow}`, d.netChangeInCash);
    }),
  ]);

  const openCashRow = nextRow();
  rows.push([lbl(options, 'cf', 'openingCash', 'Opening Cash & Cash Equivalents'), ...data.map(d => nCell(d.openingCash))]);

  const closeCashRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'closingCashCalculated', 'Closing Cash & Cash Equivalents (Calculated)'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(`${c}${netChangeRow}+${c}${openCashRow}`, d.closingCashCalculated);
    }),
  ]);

  const bsCashRow = nextRow();
  rows.push([lbl(options, 'cf', 'balanceSheetCash', 'Stated Cash in Statement of Financial Position'), ...data.map(d => nCell(d.balanceSheetCash))]);

  rows.push([
    'Reconciliation Variance (Calculated - Stated)',
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      return fCell(`${c}${closeCashRow}-${c}${bsCashRow}`, d.reconciliationDifference);
    }),
  ]);
  rows.push(['Audit Reconciliation Status', ...data.map(d => d.reconciliationStatus)]);

  const sheetData = createTableDataForStatement(title, company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [48, ...data.map(() => 18)]);
  return ws;
}

function buildEquityWorksheet(
  company: CompanyProfile,
  data: EquityStatementYear[],
  title = 'Statement of Changes in Equity',
  options?: ExportSyncOptions
): XLSX.WorkSheet {
  const headers = ['Equity Component', ...data.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const rows: ExcelCellValue[][] = [];
  const nextRow = () => STD_DATA_START_ROW + rows.length;

  const scRows: number[] = [];
  let openScRow: number | null = null;
  if (isVis(options, 'eq_openingShareCapital')) {
    openScRow = nextRow();
    scRows.push(openScRow);
    rows.push([lbl(options, 'eq', 'openingShareCapital', 'Opening Share Capital'), ...data.map(d => nCell(d.openingShareCapital))]);
  }
  if (isVis(options, 'eq_shareCapitalIssued')) {
    scRows.push(nextRow());
    rows.push([lbl(options, 'eq', 'shareCapitalIssued', 'Share Capital Issued during Year'), ...data.map(d => nCell(d.shareCapitalIssued))]);
  }
  const closeScRow = nextRow();
  rows.push([
    lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), scRows), d.closingShareCapital)),
  ]);
  rows.push([]);

  const reRows: number[] = [];
  let openReRow: number | null = null;
  if (isVis(options, 'eq_openingRetainedEarnings')) {
    openReRow = nextRow();
    reRows.push(openReRow);
    rows.push([lbl(options, 'eq', 'openingRetainedEarnings', 'Opening Retained Earnings'), ...data.map(d => nCell(d.openingRetainedEarnings))]);
  }
  if (isVis(options, 'eq_netProfitForYear')) {
    reRows.push(nextRow());
    rows.push([lbl(options, 'eq', 'netProfitForYear', 'Net Profit for the Year'), ...data.map(d => nCell(d.netProfitForYear))]);
  }
  if (isVis(options, 'eq_dividendsPaid')) {
    reRows.push(nextRow());
    rows.push([lbl(options, 'eq', 'dividendsPaid', 'Dividends Paid to Shareholders'), ...data.map(d => nCell(-d.dividendsPaid))]);
  }
  const closeReRow = nextRow();
  rows.push([
    lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), reRows), d.closingRetainedEarnings)),
  ]);
  rows.push([]);

  const resRows: number[] = [];
  let openResRow: number | null = null;
  if (isVis(options, 'eq_openingReserves')) {
    openResRow = nextRow();
    resRows.push(openResRow);
    rows.push([lbl(options, 'eq', 'openingReserves', 'Opening Reserves'), ...data.map(d => nCell(d.openingReserves))]);
  }
  if (isVis(options, 'eq_otherComprehensiveIncome')) {
    resRows.push(nextRow());
    rows.push([lbl(options, 'eq', 'otherComprehensiveIncome', 'Other Comprehensive Income / Transfers'), ...data.map(d => nCell(d.otherComprehensiveIncome))]);
  }
  const closeResRow = nextRow();
  rows.push([
    lbl(options, 'eq', 'closingReserves', 'Closing Reserves'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), resRows), d.closingReserves)),
  ]);

  const customEqRows: number[] = [];
  getCustomRows(options, 'equity').forEach(ci => {
    customEqRows.push(nextRow());
    rows.push([ci.label, ...data.map(d => nCell(getCustomItemYearValue(ci, d.year)))]);
  });

  rows.push([]);
  const openEqSumRows = [openScRow, openReRow, openResRow].filter((x): x is number => x !== null);
  rows.push([
    lbl(options, 'eq', 'totalOpeningEquity', 'TOTAL OPENING EQUITY'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), openEqSumRows), d.totalOpeningEquity)),
  ]);
  const closeEqSumRows = [closeScRow, closeReRow, closeResRow, ...customEqRows];
  rows.push([
    lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'),
    ...data.map((d, idx) => fCell(excelSumRows(colLetter(idx + 1), closeEqSumRows), d.totalClosingEquity)),
  ]);

  const sheetData = createTableDataForStatement(title, company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [42, ...data.map(() => 18)]);
  return ws;
}

function resolveFACategoriesForExport(
  latest: ScheduleOfFixedAssetsYear,
  options?: ExportSyncOptions
) {
  if (!options?.statementConfig || options.statementConfig.length === 0) {
    return latest.categories;
  }
  if (latest.allCategories && latest.allCategories.length > 0) {
    const visibleStd = latest.allCategories
      .map((cat, idx) => {
        const cfgId = FIXED_ASSET_CATEGORY_CONFIG_IDS[idx];
        const cfgItem = cfgId ? options.statementConfig?.find(c => c.id === cfgId) : undefined;
        return {
          cat: cfgItem ? { ...cat, name: cfgItem.label } : cat,
          visible: cfgId ? isVis(options, cfgId) : true,
        };
      })
      .filter(entry => entry.visible)
      .map(entry => entry.cat);
    const customRows = getCustomRows(options, 'fixedAssets').map(ci => {
      const val = getCustomItemYearValue(ci, latest.year);
      return {
        name: ci.label,
        costOpening: val,
        additions: 0,
        disposals: 0,
        costClosing: val,
        depOpening: 0,
        depCharge: 0,
        depDisposals: 0,
        depClosing: 0,
        netBookValue: val,
        depRate: 0,
      };
    });
    return [...visibleStd, ...customRows];
  }
  return latest.categories;
}

function buildFixedAssetsWorksheet(
  company: CompanyProfile,
  data: ScheduleOfFixedAssetsYear[],
  options?: ExportSyncOptions
): XLSX.WorkSheet {
  const latest = data[data.length - 1];
  const headers = ['Asset Class', 'Opening Cost', 'Additions', 'Disposals', 'Closing Cost', 'Opening Dep.', 'Dep. Charge', 'Dep. Disposals', 'Closing Dep.', 'Net Book Value'];
  if (!latest) {
    const sheetData = createTableDataForStatement('Schedule of Fixed Assets', company, headers, []);
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    applyColumnWidths(ws, [38, 16, 14, 14, 16, 16, 14, 15, 16, 18]);
    return ws;
  }
  const categories = resolveFACategoriesForExport(latest, options);
  const r = STD_DATA_START_ROW;
  const rows: ExcelCellValue[][] = categories.map((c, idx) => {
    const rowNum = r + idx;
    return [
      c.name,
      nCell(c.costOpening),
      nCell(c.additions),
      nCell(c.disposals),
      fCell(`B${rowNum}+C${rowNum}-D${rowNum}`, c.costClosing),
      nCell(c.depOpening),
      nCell(c.depCharge),
      nCell(c.depDisposals),
      fCell(`F${rowNum}+G${rowNum}-H${rowNum}`, c.depClosing),
      fCell(`E${rowNum}-I${rowNum}`, c.netBookValue),
    ];
  });

  if (categories.length > 0) {
    const endRow = r + categories.length - 1;
    rows.push([
      'TOTAL SCHEDULE OF FIXED ASSETS',
      fCell(`SUM(B${r}:B${endRow})`, latest.totalCostOpening),
      fCell(`SUM(C${r}:C${endRow})`, latest.totalAdditions),
      fCell(`SUM(D${r}:D${endRow})`, latest.totalDisposals),
      fCell(`SUM(E${r}:E${endRow})`, latest.totalCostClosing),
      fCell(`SUM(F${r}:F${endRow})`, latest.totalDepOpening),
      fCell(`SUM(G${r}:G${endRow})`, latest.totalDepCharge),
      fCell(`SUM(H${r}:H${endRow})`, latest.totalDepDisposals),
      fCell(`SUM(I${r}:I${endRow})`, latest.totalDepClosing),
      fCell(`SUM(J${r}:J${endRow})`, latest.totalNetBookValue),
    ]);
  } else {
    rows.push([
      'TOTAL SCHEDULE OF FIXED ASSETS',
      nCell(0), nCell(0), nCell(0), nCell(0),
      nCell(0), nCell(0), nCell(0), nCell(0), nCell(0),
    ]);
  }

  const sheetData = createTableDataForStatement(`Schedule of Fixed Assets (${latest.year})`, company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [38, 16, 14, 14, 16, 16, 14, 15, 16, 18]);
  return ws;
}

// ---------------- STANDARD SECTION 76 EXCEL EXPORTS ---------------- //

export function exportIncomeStatementExcel(
  company: CompanyProfile,
  data: IncomeStatementYear[],
  filenameOrOptions: string | ExportSyncOptions = 'Income_Statement.xlsx',
  explicitOptions?: ExportSyncOptions
) {
  const { filename, options } = resolveFilenameAndOptions('Income_Statement.xlsx', filenameOrOptions, explicitOptions);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildIncomeStatementSheet(company, data, 'Income Statement', options), 'Income Statement');
  XLSX.writeFile(wb, filename);
}

export function exportBalanceSheetExcel(
  company: CompanyProfile,
  data: BalanceSheetYear[],
  filenameOrOptions: string | ExportSyncOptions = 'Statement_of_Financial_Position.xlsx',
  explicitOptions?: ExportSyncOptions
) {
  const { filename, options } = resolveFilenameAndOptions('Statement_of_Financial_Position.xlsx', filenameOrOptions, explicitOptions);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildBalanceSheetWorksheet(company, data, 'Statement of Financial Position', options), 'Balance Sheet');
  XLSX.writeFile(wb, filename);
}

export function exportCashFlowExcel(
  company: CompanyProfile,
  data: CashFlowStatementYear[],
  filenameOrOptions: string | ExportSyncOptions = 'Statement_of_Cash_Flows.xlsx',
  explicitOptions?: ExportSyncOptions
) {
  const { filename, options } = resolveFilenameAndOptions('Statement_of_Cash_Flows.xlsx', filenameOrOptions, explicitOptions);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildCashFlowWorksheet(company, data, 'Statement of Cash Flows', options), 'Cash Flow Statement');
  XLSX.writeFile(wb, filename);
}

export function exportEquityStatementExcel(
  company: CompanyProfile,
  data: EquityStatementYear[],
  filenameOrOptions: string | ExportSyncOptions = 'Statement_of_Changes_in_Equity.xlsx',
  explicitOptions?: ExportSyncOptions
) {
  const { filename, options } = resolveFilenameAndOptions('Statement_of_Changes_in_Equity.xlsx', filenameOrOptions, explicitOptions);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildEquityWorksheet(company, data, 'Statement of Changes in Equity', options), 'Equity Statement');
  XLSX.writeFile(wb, filename);
}

export function exportFixedAssetsExcel(
  company: CompanyProfile,
  data: ScheduleOfFixedAssetsYear[],
  filenameOrOptions: string | ExportSyncOptions = 'Schedule_of_Fixed_Assets.xlsx',
  explicitOptions?: ExportSyncOptions
) {
  const { filename, options } = resolveFilenameAndOptions('Schedule_of_Fixed_Assets.xlsx', filenameOrOptions, explicitOptions);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildFixedAssetsWorksheet(company, data, options), 'Fixed Assets Schedule');
  XLSX.writeFile(wb, filename);
}

export function exportNotesExcel(
  company: CompanyProfile,
  notes: FinancialNoteItem[],
  filename = 'Notes_to_the_Financial_Statements.xlsx'
) {
  const wb = XLSX.utils.book_new();
  const rows: (string | number)[][] = [
    [company.companyName.toUpperCase()],
    ['NOTES TO THE FINANCIAL STATEMENTS'],
    [`Framework: ${company.accountingFramework} | Currency: ${company.reportingCurrency}`],
    [],
  ];

  notes.forEach(note => {
    rows.push([`Note ${note.noteNumber}: ${note.title} [${note.category}]`]);
    rows.push([note.textDescription]);
    if (note.tableData) {
      rows.push(note.tableData.headers);
      note.tableData.rows.forEach(r => {
        rows.push([r.label, ...r.values]);
      });
      if (note.tableData.totalRow) {
        rows.push([note.tableData.totalRow.label, ...note.tableData.totalRow.values]);
      }
    }
    rows.push([]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  applyColumnWidths(ws, [45, 22, 22, 22]);
  XLSX.utils.book_append_sheet(wb, ws, 'Statutory Notes');
  XLSX.writeFile(wb, filename);
}

function buildDebtServiceCoverageWorksheet(
  company: CompanyProfile,
  dscrData: DebtServiceCoverageYear[]
): XLSX.WorkSheet {
  const headers = [
    'Year',
    'Period',
    'EBITDA',
    'Pre-Fin OCF',
    'Scheduled Principal (Inst. Principal)',
    'Scheduled Finance Cost (Inst. Profit)',
    'Total Debt Service',
    'EBITDA DSCR (x)',
    'Cash DSCR (x)',
    'Closing Cash',
    'Capacity Status',
  ];
  const r = STD_DATA_START_ROW;
  const rows: ExcelCellValue[][] = dscrData.map((d, idx) => {
    const rowNum = r + idx;
    return [
      d.year,
      d.isProjected ? 'Projected' : 'Historical',
      nCell(d.ebitda),
      nCell(d.operatingCashFlowBeforeFinance),
      nCell(d.scheduledPrincipal),
      nCell(d.scheduledFinanceCost),
      fCell(`E${rowNum}+F${rowNum}`, d.totalDebtService),
      Number(d.ebitdaDscr.toFixed(2)),
      Number(d.cashDscr.toFixed(2)),
      nCell(d.closingCash),
      d.capacityStatus,
    ];
  });
  const sheetData = createTableDataForStatement('Debt Servicing & Repayment Capacity (DSCR)', company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [10, 12, 16, 16, 26, 26, 20, 14, 14, 18, 16]);
  return ws;
}

// COMPLETE COMBINED EXCEL PACKAGE (Section 76.11.A)
export function exportCompletePackageExcel(
  company: CompanyProfile,
  incomeData: IncomeStatementYear[],
  balanceData: BalanceSheetYear[],
  cashFlowData: CashFlowStatementYear[],
  equityData: EquityStatementYear[],
  fixedAssetsData: ScheduleOfFixedAssetsYear[],
  notes: FinancialNoteItem[],
  filename = `${company.companyName.replace(/\s+/g, '_')}_Complete_Financial_Package.xlsx`,
  debtServiceCoverage?: DebtServiceCoverageYear[],
  options?: ExportSyncOptions
) {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, buildIncomeStatementSheet(company, incomeData, 'Income Statement', options), '1. Income Statement');
  XLSX.utils.book_append_sheet(wb, buildBalanceSheetWorksheet(company, balanceData, 'Statement of Financial Position', options), '2. Balance Sheet');
  XLSX.utils.book_append_sheet(wb, buildCashFlowWorksheet(company, cashFlowData, 'Statement of Cash Flows', options), '3. Cash Flows');
  XLSX.utils.book_append_sheet(wb, buildEquityWorksheet(company, equityData, 'Statement of Changes in Equity', options), '4. Changes in Equity');
  XLSX.utils.book_append_sheet(wb, buildFixedAssetsWorksheet(company, fixedAssetsData, options), '5. Fixed Assets');

  // 6. Notes to Financial Statements Sheet
  const notesRows: (string | number)[][] = [
    [company.companyName.toUpperCase()],
    ['NOTES TO THE FINANCIAL STATEMENTS (Statutory Disclosures 1 to 12)'],
    [],
  ];
  notes.forEach(note => {
    notesRows.push([`Note ${note.noteNumber}: ${note.title}`]);
    notesRows.push([note.textDescription]);
    if (note.tableData) {
      notesRows.push(note.tableData.headers);
      note.tableData.rows.forEach(r => notesRows.push([r.label, ...r.values]));
      if (note.tableData.totalRow) {
        notesRows.push([note.tableData.totalRow.label, ...note.tableData.totalRow.values]);
      }
    }
    notesRows.push([]);
  });
  const wsNotes = XLSX.utils.aoa_to_sheet(notesRows);
  applyColumnWidths(wsNotes, [45, 22, 22, 22]);
  XLSX.utils.book_append_sheet(wb, wsNotes, '6. Notes to Financials');

  const activeDscr = (debtServiceCoverage || []).filter(d => d.totalDebtService > 0);
  if (activeDscr.length > 0) {
    XLSX.utils.book_append_sheet(wb, buildDebtServiceCoverageWorksheet(company, activeDscr), '7. Debt Servicing & DSCR');
  }

  XLSX.writeFile(wb, filename);
}

// ---------------- WORD EXPORTS (docx) ---------------- //

function createDocxTable(headers: string[], rows: (string | number)[][], leftColsCount = 1): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: '334155' },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: '334155' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
      left: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: h, bold: true, size: 18, color: '111827' })],
            alignment: i < leftColsCount ? AlignmentType.LEFT : AlignmentType.RIGHT,
          })],
          shading: { fill: 'F1F5F9' },
        })),
      }),
      ...rows.map(row => {
        const label = String(row[0] ?? '');
        const isBoldRow =
          label.startsWith('Total') ||
          label.startsWith('TOTAL') ||
          label === 'Taka' ||
          label === 'Subtotal' ||
          label === 'Gross Profit' ||
          label === 'Net Operating Profit' ||
          label === 'Net Profit After Income Tax' ||
          label.includes('Net Cash') ||
          label.includes('NET CASH') ||
          label.includes('Capital after drawings') ||
          label.includes('Total Operating Profit') ||
          label.includes('Net Profit/(Loss)') ||
          label.includes('Closing Cash');

        return new TableRow({
          children: row.map((cell, idx) => new TableCell({
            children: [new Paragraph({
              children: [new TextRun({
                text: typeof cell === 'number'
                  ? fmtExportNum(cell)
                  : String(cell),
                size: 17,
                bold: isBoldRow,
                color: '1F2937',
              })],
              alignment: idx < leftColsCount ? AlignmentType.LEFT : AlignmentType.RIGHT,
            })],
            shading: isBoldRow ? { fill: 'F8FAFC' } : undefined,
          })),
        });
      }),
    ],
  });
}

export async function exportDocumentWord(
  company: CompanyProfile,
  title: string,
  headers: string[],
  rows: (string | number)[][],
  filename = 'Financial_Report.docx'
) {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: company.companyName,
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: title, bold: true, size: 24, color: '1E3A8A' }),
            ],
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `Reporting Currency: ${company.reportingCurrency} (${company.reportingUnit}) | Framework: ${company.accountingFramework}`, italics: true, size: 18 }),
            ],
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `Prepared by: ${company.preparedBy} · Financials with Tuhin Engine`, size: 16, color: '4B5563' }),
            ],
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: '' }),
          createDocxTable(headers, rows),
          new Paragraph({ text: '' }),
          new Paragraph({
            children: [
              new TextRun({ text: 'CONFIDENTIAL & PROPRIETARY — STATUTORY FINANCIAL REPORT', size: 14, color: '9CA3AF' }),
            ],
            alignment: AlignmentType.RIGHT,
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, filename);
}

// Shared synchronized Section 76 table-row builders for Word & PDF complete packages
function buildSyncIncomeTextRows(incomeData: IncomeStatementYear[], options?: ExportSyncOptions): (string | number)[][] {
  const customRev = getCustomRows(options, 'income', 'Revenue & Gross Margin');
  const customOpex = getCustomRows(options, 'income', 'Operating Expenses');
  const customNonOp = [...getCustomRows(options, 'income', 'Non-Operating & Tax'), ...getUnscopedCustomIncomeRows(options)];
  const rows: (string | number)[][] = [];
  if (isVis(options, 'inc_revenue')) {
    rows.push([
      lbl(options, 'income', 'revenue', 'Sales Revenue'),
      ...incomeData.map(d => d.revenue - customRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)),
    ]);
  }
  customRev.forEach(ci => rows.push([ci.label, ...incomeData.map(d => getCustomItemYearValue(ci, d.year))]));
  if (isVis(options, 'inc_costOfSales')) {
    rows.push([lbl(options, 'income', 'costOfSales', 'Cost of Sales'), ...incomeData.map(d => -d.costOfSales)]);
  }
  rows.push([lbl(options, 'income', 'grossProfit', 'Gross Profit'), ...incomeData.map(d => d.grossProfit)]);
  if (isVis(options, 'inc_operatingExpenses')) {
    rows.push([
      lbl(options, 'income', 'operatingExpenses', 'Operating Expenses'),
      ...incomeData.map(d => -(d.operatingExpenses + customOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))),
    ]);
  }
  customOpex.forEach(ci => rows.push([ci.label, ...incomeData.map(d => getCustomItemYearValue(ci, d.year))]));
  if (isVis(options, 'inc_depreciationExpense')) {
    rows.push([lbl(options, 'income', 'depreciationExpense', 'Depreciation Expense'), ...incomeData.map(d => -d.depreciationExpense)]);
  }
  rows.push([lbl(options, 'income', 'operatingProfit', 'Operating Profit (EBIT)'), ...incomeData.map(d => d.operatingProfit)]);
  if (isVis(options, 'inc_interestIncome')) {
    rows.push([
      lbl(options, 'income', 'interestIncome', 'Interest & Investment Income'),
      ...incomeData.map(d => d.interestIncome - customNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)),
    ]);
  }
  customNonOp.forEach(ci => rows.push([ci.label, ...incomeData.map(d => getCustomItemYearValue(ci, d.year))]));
  if (isVis(options, 'inc_financeCost')) {
    rows.push([lbl(options, 'income', 'financeCost', 'Finance Cost (Mandatory Direct Entry)'), ...incomeData.map(d => -d.financeCost)]);
  }
  rows.push([lbl(options, 'income', 'profitBeforeTax', 'Profit Before Tax (EBT)'), ...incomeData.map(d => d.profitBeforeTax)]);
  if (isVis(options, 'inc_taxExpense')) {
    rows.push([lbl(options, 'income', 'taxExpense', 'Income Tax Expense'), ...incomeData.map(d => -d.taxExpense)]);
  }
  rows.push([lbl(options, 'income', 'netProfit', 'NET PROFIT AFTER TAX'), ...incomeData.map(d => d.netProfit)]);
  return rows;
}

function buildSyncBalanceTextRows(balanceData: BalanceSheetYear[], options?: ExportSyncOptions): (string | number)[][] {
  const rows: (string | number)[][] = [];
  if (isVis(options, 'bs_cashAndEquivalents')) rows.push([lbl(options, 'bs', 'cashAndEquivalents', 'Cash and Cash Equivalents'), ...balanceData.map(d => d.cashAndEquivalents)]);
  if (isVis(options, 'bs_tradeReceivables')) rows.push([lbl(options, 'bs', 'tradeReceivables', 'Trade Receivables'), ...balanceData.map(d => d.tradeReceivables)]);
  if (isVis(options, 'bs_otherReceivables')) rows.push([lbl(options, 'bs', 'otherReceivables', 'Other Receivables & Prepayments'), ...balanceData.map(d => d.otherReceivables)]);
  if (isVis(options, 'bs_closingInventory')) rows.push([lbl(options, 'bs', 'closingInventory', 'Closing Inventory'), ...balanceData.map(d => d.closingInventory)]);
  if (isVis(options, 'bs_otherCurrentAssets')) rows.push([lbl(options, 'bs', 'otherCurrentAssets', 'Other Current Assets'), ...balanceData.map(d => d.otherCurrentAssets)]);
  getCustomRows(options, 'balanceSheet', 'Current Assets').forEach(ci => rows.push([ci.label, ...balanceData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'bs', 'totalCurrentAssets', 'TOTAL CURRENT ASSETS'), ...balanceData.map(d => d.totalCurrentAssets)]);
  if (isVis(options, 'bs_propertyPlantEquipment')) rows.push([lbl(options, 'bs', 'propertyPlantEquipment', 'Property, Plant and Equipment (NBV)'), ...balanceData.map(d => d.propertyPlantEquipment)]);
  if (isVis(options, 'bs_intangibleAssets')) rows.push([lbl(options, 'bs', 'intangibleAssets', 'Intangible Assets'), ...balanceData.map(d => d.intangibleAssets)]);
  if (isVis(options, 'bs_otherNonCurrentAssets')) rows.push([lbl(options, 'bs', 'otherNonCurrentAssets', 'Other Non-Current Assets'), ...balanceData.map(d => d.otherNonCurrentAssets)]);
  getCustomRows(options, 'balanceSheet', 'Non-Current Assets').forEach(ci => rows.push([ci.label, ...balanceData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'bs', 'totalNonCurrentAssets', 'TOTAL NON-CURRENT ASSETS'), ...balanceData.map(d => d.totalNonCurrentAssets)]);
  rows.push([lbl(options, 'bs', 'totalAssets', 'TOTAL ASSETS'), ...balanceData.map(d => d.totalAssets)]);
  if (isVis(options, 'bs_tradePayables')) rows.push([lbl(options, 'bs', 'tradePayables', 'Trade Payables'), ...balanceData.map(d => d.tradePayables)]);
  if (isVis(options, 'bs_otherPayables')) rows.push([lbl(options, 'bs', 'otherPayables', 'Other Payables & Accruals'), ...balanceData.map(d => d.otherPayables)]);
  if (isVis(options, 'bs_shortTermBorrowings')) rows.push([lbl(options, 'bs', 'shortTermBorrowings', 'Short-Term Borrowings'), ...balanceData.map(d => d.shortTermBorrowings)]);
  if (isVis(options, 'bs_currentTaxPayable')) rows.push([lbl(options, 'bs', 'currentTaxPayable', 'Current Tax Payable'), ...balanceData.map(d => d.currentTaxPayable)]);
  getCustomRows(options, 'balanceSheet', 'Current Liabilities').forEach(ci => rows.push([ci.label, ...balanceData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'bs', 'totalCurrentLiabilities', 'Total Current Liabilities'), ...balanceData.map(d => d.totalCurrentLiabilities)]);
  if (isVis(options, 'bs_longTermBorrowings')) rows.push([lbl(options, 'bs', 'longTermBorrowings', 'Long-Term Borrowings'), ...balanceData.map(d => d.longTermBorrowings)]);
  if (isVis(options, 'bs_otherNonCurrentLiabilities')) rows.push([lbl(options, 'bs', 'otherNonCurrentLiabilities', 'Other Non-Current Liabilities'), ...balanceData.map(d => d.otherNonCurrentLiabilities)]);
  getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities').forEach(ci => rows.push([ci.label, ...balanceData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'bs', 'totalNonCurrentLiabilities', 'Total Non-Current Liabilities'), ...balanceData.map(d => d.totalNonCurrentLiabilities)]);
  rows.push([lbl(options, 'bs', 'totalLiabilities', 'TOTAL LIABILITIES'), ...balanceData.map(d => d.totalLiabilities)]);
  if (isVis(options, 'bs_shareCapital')) rows.push([lbl(options, 'bs', 'shareCapital', 'Share Capital'), ...balanceData.map(d => d.shareCapital)]);
  if (isVis(options, 'bs_retainedEarnings')) rows.push([lbl(options, 'bs', 'retainedEarnings', 'Retained Earnings'), ...balanceData.map(d => d.retainedEarnings)]);
  if (isVis(options, 'bs_reserves')) rows.push([lbl(options, 'bs', 'reserves', 'Reserves'), ...balanceData.map(d => d.reserves)]);
  getCustomRows(options, 'balanceSheet', 'Equity').forEach(ci => rows.push([ci.label, ...balanceData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'bs', 'totalEquity', 'TOTAL EQUITY'), ...balanceData.map(d => d.totalEquity)]);
  rows.push([lbl(options, 'bs', 'totalLiabilitiesAndEquity', 'TOTAL LIABILITIES & EQUITY'), ...balanceData.map(d => d.totalLiabilitiesAndEquity)]);
  rows.push(['Balance Check', ...balanceData.map(d => d.balanceCheck)]);
  return rows;
}

function buildSyncCashFlowTextRows(cashFlowData: CashFlowStatementYear[], options?: ExportSyncOptions): (string | number)[][] {
  const rows: (string | number)[][] = [];
  if (isVis(options, 'cf_pbt')) rows.push([lbl(options, 'cf', 'profitBeforeTax', 'Net Profit Before Tax'), ...cashFlowData.map(d => d.profitBeforeTax)]);
  if (isVis(options, 'cf_dep')) rows.push([lbl(options, 'cf', 'adjDepreciation', 'Adjustments for Depreciation'), ...cashFlowData.map(d => d.adjDepreciation)]);
  if (isVis(options, 'cf_amort')) rows.push([lbl(options, 'cf', 'adjAmortization', 'Adjustments for Amortization'), ...cashFlowData.map(d => d.adjAmortization)]);
  if (isVis(options, 'cf_finance')) rows.push([lbl(options, 'cf', 'adjFinanceCost', 'Finance Cost Adjustment'), ...cashFlowData.map(d => d.adjFinanceCost)]);
  if (isVis(options, 'cf_interest_inc')) rows.push([lbl(options, 'cf', 'adjInterestIncome', 'Interest & Investment Income'), ...cashFlowData.map(d => d.adjInterestIncome)]);
  if (isVis(options, 'cf_disposal_gain_loss')) rows.push([lbl(options, 'cf', 'adjLossGainDisposal', 'Loss / (Gain) on Disposal of Fixed Assets'), ...cashFlowData.map(d => d.adjLossGainDisposal)]);
  rows.push([lbl(options, 'cf', 'operatingProfitBeforeWC', 'Operating Profit Before Working Capital'), ...cashFlowData.map(d => d.operatingProfitBeforeWC)]);
  if (isVis(options, 'cf_wc_inventory')) rows.push([lbl(options, 'cf', 'changeInInventory', '(Increase) / Decrease in Inventory'), ...cashFlowData.map(d => d.changeInInventory)]);
  if (isVis(options, 'cf_wc_receivables')) rows.push([lbl(options, 'cf', 'changeInReceivables', '(Increase) / Decrease in Trade Receivables'), ...cashFlowData.map(d => d.changeInReceivables)]);
  if (isVis(options, 'cf_wc_other_rec')) rows.push([lbl(options, 'cf', 'changeInOtherReceivables', '(Increase) / Decrease in Other Receivables'), ...cashFlowData.map(d => d.changeInOtherReceivables)]);
  if (isVis(options, 'cf_wc_other_ca')) rows.push([lbl(options, 'cf', 'changeInOtherCA', '(Increase) / Decrease in Other Current Assets'), ...cashFlowData.map(d => d.changeInOtherCA ?? 0)]);
  if (isVis(options, 'cf_wc_payables')) rows.push([lbl(options, 'cf', 'changeInPayables', 'Increase / (Decrease) in Trade Payables'), ...cashFlowData.map(d => d.changeInPayables)]);
  if (isVis(options, 'cf_wc_other_pay')) rows.push([lbl(options, 'cf', 'changeInOtherPayables', 'Increase / (Decrease) in Other Payables'), ...cashFlowData.map(d => d.changeInOtherPayables)]);
  getCustomRows(options, 'cashFlow', 'Operating Activities').forEach(ci => rows.push([ci.label, ...cashFlowData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'cf', 'cashGeneratedFromOperations', 'Cash Generated from Operations'), ...cashFlowData.map(d => d.cashGeneratedFromOperations)]);
  if (isVis(options, 'cf_tax_paid')) rows.push([lbl(options, 'cf', 'incomeTaxPaid', 'Income Tax Paid'), ...cashFlowData.map(d => d.incomeTaxPaid)]);
  if (isVis(options, 'cf_finance_paid')) rows.push([lbl(options, 'cf', 'financeCostPaid', 'Finance Cost Paid'), ...cashFlowData.map(d => d.financeCostPaid)]);
  rows.push([lbl(options, 'cf', 'netCashFromOperating', 'NET CASH FROM OPERATING ACTIVITIES'), ...cashFlowData.map(d => d.netCashFromOperating)]);
  if (isVis(options, 'cf_pur_ppe')) rows.push([lbl(options, 'cf', 'purchasePPE', 'Purchase of PPE (Capex)'), ...cashFlowData.map(d => d.purchasePPE)]);
  if (isVis(options, 'cf_sale_ppe')) rows.push([lbl(options, 'cf', 'salePPE', 'Proceeds from Sale of Fixed Assets'), ...cashFlowData.map(d => d.salePPE)]);
  if (isVis(options, 'cf_pur_intangible')) rows.push([lbl(options, 'cf', 'purchaseIntangibles', 'Purchase of Intangible Assets'), ...cashFlowData.map(d => d.purchaseIntangibles)]);
  if (isVis(options, 'cf_interest_rec')) rows.push([lbl(options, 'cf', 'interestReceived', 'Interest Received'), ...cashFlowData.map(d => d.interestReceived)]);
  if (isVis(options, 'cf_dividends_rec')) rows.push([lbl(options, 'cf', 'dividendsReceived', 'Dividends Received'), ...cashFlowData.map(d => d.dividendsReceived)]);
  getCustomRows(options, 'cashFlow', 'Investing Activities').forEach(ci => rows.push([ci.label, ...cashFlowData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'cf', 'netCashFromInvesting', 'NET CASH FROM / (USED IN) INVESTING'), ...cashFlowData.map(d => d.netCashFromInvesting)]);
  if (isVis(options, 'cf_shares_issue')) rows.push([lbl(options, 'cf', 'financingProceedsShares', 'Proceeds from Issue of Share Capital'), ...cashFlowData.map(d => d.financingProceedsShares)]);
  if (isVis(options, 'cf_borrow_proceeds')) rows.push([lbl(options, 'cf', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing'), ...cashFlowData.map(d => d.financingProceedsBorrowings)]);
  if (isVis(options, 'cf_borrow_repay')) rows.push([lbl(options, 'cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing'), ...cashFlowData.map(d => -(d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings)))]);
  if (isVis(options, 'cf_dividends_paid')) rows.push([lbl(options, 'cf', 'financingDividendsPaid', 'Dividends Paid'), ...cashFlowData.map(d => d.financingDividendsPaid)]);
  getCustomRows(options, 'cashFlow', 'Financing Activities').forEach(ci => rows.push([ci.label, ...cashFlowData.map(d => getCustomItemYearValue(ci, d.year))]));
  if (isVis(options, 'cf_direct_financing')) rows.push([lbl(options, 'cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Financing Cash Flow'), ...cashFlowData.map(d => d.directlyEnteredFinancingCashFlow)]);
  rows.push([lbl(options, 'cf', 'netCashFromFinancing', 'NET CASH FROM / (USED IN) FINANCING'), ...cashFlowData.map(d => d.netCashFromFinancing)]);
  rows.push([lbl(options, 'cf', 'netChangeInCash', 'Net Change in Cash and Equivalents'), ...cashFlowData.map(d => d.netChangeInCash)]);
  rows.push([lbl(options, 'cf', 'openingCash', 'Opening Cash & Cash Equivalents'), ...cashFlowData.map(d => d.openingCash)]);
  rows.push([lbl(options, 'cf', 'closingCashCalculated', 'Closing Cash (Calculated)'), ...cashFlowData.map(d => d.closingCashCalculated)]);
  rows.push([lbl(options, 'cf', 'balanceSheetCash', 'Stated Balance Sheet Cash'), ...cashFlowData.map(d => d.balanceSheetCash)]);
  rows.push(['Reconciliation Difference', ...cashFlowData.map(d => d.reconciliationDifference)]);
  rows.push(['Audit Reconciliation Status', ...cashFlowData.map(d => d.reconciliationStatus)]);
  return rows;
}

function buildSyncEquityTextRows(equityData: EquityStatementYear[], options?: ExportSyncOptions): (string | number)[][] {
  const rows: (string | number)[][] = [];
  if (isVis(options, 'eq_openingShareCapital')) rows.push([lbl(options, 'eq', 'openingShareCapital', 'Opening Share Capital'), ...equityData.map(d => d.openingShareCapital)]);
  if (isVis(options, 'eq_shareCapitalIssued')) rows.push([lbl(options, 'eq', 'shareCapitalIssued', 'Share Capital Issued'), ...equityData.map(d => d.shareCapitalIssued)]);
  rows.push([lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'), ...equityData.map(d => d.closingShareCapital)]);
  if (isVis(options, 'eq_openingRetainedEarnings')) rows.push([lbl(options, 'eq', 'openingRetainedEarnings', 'Opening Retained Earnings'), ...equityData.map(d => d.openingRetainedEarnings)]);
  if (isVis(options, 'eq_netProfitForYear')) rows.push([lbl(options, 'eq', 'netProfitForYear', 'Net Profit for the Year'), ...equityData.map(d => d.netProfitForYear)]);
  if (isVis(options, 'eq_dividendsPaid')) rows.push([lbl(options, 'eq', 'dividendsPaid', 'Dividends Paid'), ...equityData.map(d => -d.dividendsPaid)]);
  rows.push([lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'), ...equityData.map(d => d.closingRetainedEarnings)]);
  if (isVis(options, 'eq_openingReserves')) rows.push([lbl(options, 'eq', 'openingReserves', 'Opening Reserves'), ...equityData.map(d => d.openingReserves)]);
  if (isVis(options, 'eq_otherComprehensiveIncome')) rows.push([lbl(options, 'eq', 'otherComprehensiveIncome', 'Other Comprehensive Income'), ...equityData.map(d => d.otherComprehensiveIncome)]);
  rows.push([lbl(options, 'eq', 'closingReserves', 'Closing Reserves'), ...equityData.map(d => d.closingReserves)]);
  getCustomRows(options, 'equity').forEach(ci => rows.push([ci.label, ...equityData.map(d => getCustomItemYearValue(ci, d.year))]));
  rows.push([lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'), ...equityData.map(d => d.totalClosingEquity)]);
  return rows;
}

// COMPLETE WORD PACKAGE (Section 76.11.B)
export async function exportCompletePackageWord(
  company: CompanyProfile,
  incomeData: IncomeStatementYear[],
  balanceData: BalanceSheetYear[],
  cashFlowData: CashFlowStatementYear[],
  equityData: EquityStatementYear[],
  fixedAssetsData: ScheduleOfFixedAssetsYear[],
  notes: FinancialNoteItem[],
  filename = `${company.companyName.replace(/\s+/g, '_')}_Complete_Package.docx`,
  debtServiceCoverage?: DebtServiceCoverageYear[],
  options?: ExportSyncOptions
) {
  const isHeaders = ['Item', ...incomeData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const isRows = buildSyncIncomeTextRows(incomeData, options);

  const bsHeaders = ['Item', ...balanceData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const bsRows = buildSyncBalanceTextRows(balanceData, options);

  const cfHeaders = ['Item', ...cashFlowData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const cfRows = buildSyncCashFlowTextRows(cashFlowData, options);

  const eqHeaders = ['Equity Component', ...equityData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const eqRows = buildSyncEquityTextRows(equityData, options);

  const latestFA = fixedAssetsData[fixedAssetsData.length - 1];
  const faHeaders = ['Asset Class', 'Opening Cost', 'Additions', 'Disposals', 'Closing Cost', 'Closing Dep.', 'Net Book Value'];
  const faCats = latestFA ? resolveFACategoriesForExport(latestFA, options) : [];
  const faRows: (string | number)[][] = faCats.map(c => [
    c.name, c.costOpening, c.additions, c.disposals, c.costClosing, c.depClosing, c.netBookValue
  ]);
  if (latestFA) {
    faRows.push(['TOTAL', latestFA.totalCostOpening, latestFA.totalAdditions, latestFA.totalDisposals, latestFA.totalCostClosing, latestFA.totalDepClosing, latestFA.totalNetBookValue]);
  }

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: company.companyName, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: 'COMPLETE FINANCIAL STATEMENTS PACKAGE', heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: `Framework: ${company.accountingFramework} | Entity: ${company.entityType.toUpperCase()}`, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: `Lead Systems Architect: ${company.preparedBy}`, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: '' }),

          new Paragraph({ text: '1. Income Statement', heading: HeadingLevel.HEADING_2 }),
          createDocxTable(isHeaders, isRows),
          new Paragraph({ text: '' }),

          new Paragraph({ text: '2. Statement of Financial Position', heading: HeadingLevel.HEADING_2 }),
          createDocxTable(bsHeaders, bsRows),
          new Paragraph({ text: '' }),

          new Paragraph({ text: '3. Statement of Cash Flows (Section 76 Compliance)', heading: HeadingLevel.HEADING_2 }),
          createDocxTable(cfHeaders, cfRows),
          new Paragraph({ text: '' }),

          new Paragraph({ text: '4. Statement of Changes in Equity', heading: HeadingLevel.HEADING_2 }),
          createDocxTable(eqHeaders, eqRows),
          new Paragraph({ text: '' }),

          new Paragraph({ text: `5. Schedule of Fixed Assets (${latestFA?.year ?? ''})`, heading: HeadingLevel.HEADING_2 }),
          createDocxTable(faHeaders, faRows),
          new Paragraph({ text: '' }),

          new Paragraph({ text: '6. Summary of Significant Accounting Policies & Notes', heading: HeadingLevel.HEADING_2 }),
          ...notes.flatMap(n => [
            new Paragraph({ text: `Note ${n.noteNumber}: ${n.title}`, heading: HeadingLevel.HEADING_3 }),
            new Paragraph({ text: n.textDescription }),
            ...(n.tableData ? [createDocxTable(n.tableData.headers, [
              ...n.tableData.rows.map(r => [r.label, ...r.values]),
              ...(n.tableData.totalRow ? [[n.tableData.totalRow.label, ...n.tableData.totalRow.values]] : []),
            ])] : []),
            new Paragraph({ text: '' }),
          ]),

          ...((debtServiceCoverage || []).some(d => d.totalDebtService > 0)
            ? [
                new Paragraph({ text: '7. Debt Servicing & Repayment Capacity Analysis (DSCR)', heading: HeadingLevel.HEADING_2 }),
                createDocxTable(
                  ['Year', 'EBITDA', 'Pre-Fin OCF', 'Inst. Principal', 'Finance Cost', 'Total Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status'],
                  (debtServiceCoverage || [])
                    .filter(d => d.totalDebtService > 0)
                    .map(d => [
                      `${d.year}${d.isProjected ? ' (P)' : ''}`,
                      d.ebitda,
                      d.operatingCashFlowBeforeFinance,
                      d.scheduledPrincipal,
                      d.scheduledFinanceCost,
                      d.totalDebtService,
                      `${d.ebitdaDscr.toFixed(2)}x`,
                      `${d.cashDscr.toFixed(2)}x`,
                      d.closingCash,
                      d.capacityStatus,
                    ])
                ),
              ]
            : []),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, filename);
}

// ---------------- PDF EXPORTS (jspdf + jspdf-autotable) ---------------- //

export function exportDocumentPdf(
  company: CompanyProfile,
  title: string,
  headers: string[],
  rows: (string | number)[][],
  filename = 'Financial_Report.pdf'
) {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, 297, 24, 'F');

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(14);
  doc.setFont('helvetica', 'bold');
  doc.text(company.companyName.toUpperCase(), 14, 10);

  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.text(`${title} · ${company.accountingFramework} Standards · Currency: ${company.reportingCurrency} (${company.reportingUnit})`, 14, 17);
  doc.text(`Financials with Tuhin | Mohammad Towhidul Akbar Tuhin`, 297 - 14, 17, { align: 'right' });

  autoTable(doc, {
    head: [headers],
    body: rows.map(r => r.map(c => typeof c === 'number' ? fmtExportNum(c) : String(c))),
    startY: 28,
    styles: { fontSize: 8, cellPadding: 2, font: 'helvetica' },
    headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    margin: { left: 14, right: 14 },
    didDrawPage: (data) => {
      doc.setFontSize(8);
      doc.setTextColor(148, 163, 184);
      doc.text(
        `Financials with Tuhin · Page ${data.pageNumber} of ${doc.getNumberOfPages()}`,
        14,
        210 - 8
      );
      doc.text(
        `Certified System Export · Date: ${new Date().toISOString().split('T')[0]}`,
        297 - 14,
        210 - 8,
        { align: 'right' }
      );
    },
  });

  doc.save(filename);
}

// COMPLETE COMBINED PDF PACKAGE (Section 76.11.C)
export function exportCompletePackagePdf(
  company: CompanyProfile,
  incomeData: IncomeStatementYear[],
  balanceData: BalanceSheetYear[],
  cashFlowData: CashFlowStatementYear[],
  equityData: EquityStatementYear[],
  fixedAssetsData: ScheduleOfFixedAssetsYear[],
  notes: FinancialNoteItem[],
  filename = `${company.companyName.replace(/\s+/g, '_')}_Complete_Financial_Package.pdf`,
  debtServiceCoverage?: DebtServiceCoverageYear[],
  options?: ExportSyncOptions
) {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  // Cover Page
  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, 297, 210, 'F');

  doc.setDrawColor(59, 130, 246);
  doc.setLineWidth(1);
  doc.rect(10, 10, 277, 190);

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(26);
  doc.setFont('helvetica', 'bold');
  doc.text(company.companyName.toUpperCase(), 148, 65, { align: 'center' });

  doc.setFontSize(16);
  doc.setTextColor(147, 197, 253);
  doc.text('ANNUAL STATUTORY FINANCIAL STATEMENTS & AUDIT SCHEDULES', 148, 80, { align: 'center' });

  doc.setFontSize(11);
  doc.setTextColor(203, 213, 225);
  doc.text(`Five Standard Financial Statements + Complete Notes to the Financial Statements`, 148, 92, { align: 'center' });
  doc.text(`Compliant with Section 76 & International Accounting Framework (${company.accountingFramework})`, 148, 100, { align: 'center' });

  doc.setDrawColor(71, 85, 105);
  doc.line(60, 115, 237, 115);

  doc.setFontSize(10);
  doc.setTextColor(148, 163, 184);
  doc.text(`Reporting Entity Type: ${company.entityType.toUpperCase()}`, 148, 125, { align: 'center' });
  doc.text(`Corporate Address: ${company.registeredAddress}`, 148, 133, { align: 'center' });
  doc.text(`Reporting Currency: ${company.reportingCurrency} (${company.reportingUnit}) | Multi-Year Horizon: ${incomeData[0]?.year ?? ''} - ${incomeData[incomeData.length - 1]?.year ?? ''}`, 148, 141, { align: 'center' });

  doc.setTextColor(52, 211, 153);
  doc.text(`Lead Financial Architect: ${company.preparedBy}`, 148, 160, { align: 'center' });
  doc.setTextColor(148, 163, 184);
  doc.text(`Independent Statutory Auditors: ${company.auditorName}`, 148, 168, { align: 'center' });

  // Page 2: Income Statement
  doc.addPage();
  renderPdfHeader(doc, company, '1. STATEMENT OF COMPREHENSIVE INCOME (INCOME STATEMENT)');
  const isHeaders = ['Line Item', ...incomeData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const isRows = buildSyncIncomeTextRows(incomeData, options);
  autoTable(doc, {
    head: [isHeaders],
    body: isRows.map(r => r.map(c => typeof c === 'number' ? fmtExportNum(c) : String(c))),
    startY: 28,
    styles: { fontSize: 8, cellPadding: 2 },
    headStyles: { fillColor: [30, 41, 59] },
  });

  // Page 3: Statement of Financial Position
  doc.addPage();
  renderPdfHeader(doc, company, '2. STATEMENT OF FINANCIAL POSITION (BALANCE SHEET)');
  const bsHeaders = ['Item', ...balanceData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const bsRows = buildSyncBalanceTextRows(balanceData, options);
  autoTable(doc, {
    head: [bsHeaders],
    body: bsRows.map(r => r.map(c => typeof c === 'number' ? fmtExportNum(c) : String(c))),
    startY: 28,
    styles: { fontSize: 7.5, cellPadding: 1.8 },
    headStyles: { fillColor: [30, 41, 59] },
  });

  // Page 4: Statement of Cash Flows
  doc.addPage();
  renderPdfHeader(doc, company, '3. STATEMENT OF CASH FLOWS (INDIRECT METHOD - SECTION 76)');
  const cfHeaders = ['Item', ...cashFlowData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const cfRows = buildSyncCashFlowTextRows(cashFlowData, options);
  autoTable(doc, {
    head: [cfHeaders],
    body: cfRows.map(r => r.map(c => typeof c === 'number' ? fmtExportNum(c) : String(c))),
    startY: 28,
    styles: { fontSize: 7, cellPadding: 1.6 },
    headStyles: { fillColor: [30, 41, 59] },
  });

  // Page 5: Statement of Changes in Equity
  doc.addPage();
  renderPdfHeader(doc, company, '4. STATEMENT OF CHANGES IN EQUITY');
  const eqHeaders = ['Component', ...equityData.map(d => `${d.year}${d.isProjected ? ' (P)' : ''}`)];
  const eqRows = buildSyncEquityTextRows(equityData, options);
  autoTable(doc, {
    head: [eqHeaders],
    body: eqRows.map(r => r.map(c => typeof c === 'number' ? fmtExportNum(c) : String(c))),
    startY: 28,
    styles: { fontSize: 8, cellPadding: 2 },
    headStyles: { fillColor: [30, 41, 59] },
  });

  // Page 6: Schedule of Fixed Assets
  doc.addPage();
  const latestFA = fixedAssetsData[fixedAssetsData.length - 1];
  renderPdfHeader(doc, company, `5. SCHEDULE OF FIXED ASSETS (${latestFA?.year ?? ''})`);
  const faHeaders = ['Category', 'Opening Cost', 'Additions', 'Disposals', 'Closing Cost', 'Opening Dep.', 'Charge', 'Closing Dep.', 'Net Book Value'];
  const faCats = latestFA ? resolveFACategoriesForExport(latestFA, options) : [];
  const faRows: (string | number)[][] = faCats.map(c => [
    c.name, c.costOpening, c.additions, c.disposals, c.costClosing, c.depOpening, c.depCharge, c.depClosing, c.netBookValue
  ]);
  if (latestFA) {
    faRows.push(['TOTAL PPE', latestFA.totalCostOpening, latestFA.totalAdditions, latestFA.totalDisposals, latestFA.totalCostClosing, latestFA.totalDepOpening, latestFA.totalDepCharge, latestFA.totalDepClosing, latestFA.totalNetBookValue]);
  }
  autoTable(doc, {
    head: [faHeaders],
    body: faRows.map(r => r.map(c => typeof c === 'number' ? fmtExportNum(c) : String(c))),
    startY: 28,
    styles: { fontSize: 8, cellPadding: 2 },
    headStyles: { fillColor: [30, 41, 59] },
  });

  // Page 7: Notes to the Financial Statements
  doc.addPage();
  renderPdfHeader(doc, company, '6. NOTES TO THE FINANCIAL STATEMENTS (STATUTORY DISCLOSURES)');
  let yPos = 30;
  notes.forEach((note) => {
    if (yPos > 175) {
      doc.addPage();
      renderPdfHeader(doc, company, '6. NOTES TO THE FINANCIAL STATEMENTS (CONTINUED)');
      yPos = 30;
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(30, 41, 59);
    doc.text(`Note ${note.noteNumber}: ${note.title} [${note.category}]`, 14, yPos);
    yPos += 5;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(71, 85, 105);
    const splitText = doc.splitTextToSize(note.textDescription, 269);
    doc.text(splitText, 14, yPos);
    yPos += splitText.length * 4 + 4;
  });

  const activeDscr = (debtServiceCoverage || []).filter(d => d.totalDebtService > 0);
  if (activeDscr.length > 0) {
    doc.addPage();
    renderPdfHeader(doc, company, '7. DEBT SERVICING & REPAYMENT CAPACITY ANALYSIS (DSCR)');
    autoTable(doc, {
      head: [['Year', 'EBITDA', 'Pre-Fin OCF', 'Inst. Principal', 'Finance Cost', 'Total Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status']],
      body: activeDscr.map(d => [
        `${d.year}${d.isProjected ? ' (P)' : ''}`,
        fmtExportNum(d.ebitda),
        fmtExportNum(d.operatingCashFlowBeforeFinance),
        fmtExportNum(d.scheduledPrincipal),
        fmtExportNum(d.scheduledFinanceCost),
        fmtExportNum(d.totalDebtService),
        `${d.ebitdaDscr.toFixed(2)}x`,
        `${d.cashDscr.toFixed(2)}x`,
        fmtExportNum(d.closingCash),
        d.capacityStatus,
      ]),
      startY: 28,
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: { fillColor: [30, 41, 59] },
    });
  }

  const totalPages = doc.getNumberOfPages();
  for (let i = 2; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(148, 163, 184);
    doc.text(`Financials with Tuhin · Page ${i} of ${totalPages}`, 14, 210 - 8);
    doc.text(`Certified Export · All statements reconciled under Section 76`, 297 - 14, 210 - 8, { align: 'right' });
  }

  doc.save(filename);
}

function renderPdfHeader(doc: jsPDF, company: CompanyProfile, title: string) {
  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, 297, 22, 'F');

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(11);
  doc.setFont('helvetica', 'bold');
  doc.text(company.companyName.toUpperCase(), 14, 9);

  doc.setFontSize(8);
  doc.setFont('helvetica', 'normal');
  doc.text(`${title} · ${company.reportingCurrency} (${company.reportingUnit}) · ${company.accountingFramework}`, 14, 16);
  doc.text(`Financials with Tuhin | Mohammad Towhidul Akbar Tuhin`, 297 - 14, 16, { align: 'right' });
}

// ============================================================================
// PREVIOUS FINANCIALS EXACT-FORMAT EXPORTS (.xlsx, .docx, .pdf)
// Strictly preserves layout, particulars, row order, and descending 31.12.YYYY
// columns from "Previous Financials.pdf"
// ============================================================================

export function sanitizeFilePrefix(name: string): string {
  return (name || 'Financial_Statements')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

export function buildPreviousPairFilename(
  company: CompanyProfile,
  prevData: PreviousFormatYearData[],
  ext: 'xlsx' | 'docx' | 'pdf'
): string {
  const prefix = sanitizeFilePrefix(company.companyName);
  if (prevData.length === 0) return `${prefix}_Previous_Financials.${ext}`;
  const years = prevData.map(d => d.year);
  const minYear = Math.min(...years);
  const maxYear = Math.max(...years);
  if (minYear === maxYear) {
    return `${prefix}_Previous_Financials_${minYear}.${ext}`;
  }
  return `${prefix}_Previous_Financials_${minYear}_${maxYear}.${ext}`;
}

export function buildProjectedFilename(
  company: CompanyProfile,
  projData: ProjectedFormatYearData[],
  ext: 'xlsx' | 'docx' | 'pdf'
): string {
  const prefix = sanitizeFilePrefix(company.companyName);
  if (projData.length === 0) return `${prefix}_Projected_Financials.${ext}`;
  const years = projData.map(d => d.year);
  const minYear = Math.min(...years);
  const maxYear = Math.max(...years);
  if (minYear === maxYear) {
    return `${prefix}_Projected_Financials_${minYear}.${ext}`;
  }
  return `${prefix}_Projected_Financials_${minYear}_${maxYear}.${ext}`;
}

function getCurrencyUnitCaption(company: CompanyProfile): string {
  const ccy = company.reportingCurrency === '৳' ? 'BDT' : company.reportingCurrency;
  if (company.reportingUnit === 'Units') {
    return `(Figure in ${ccy})`;
  }
  return `(Figure in ${ccy} · ${company.reportingUnit})`;
}

function fmtExactNum(val: number): string {
  return fmtExportNum(val);
}

function buildPreviousTextTables(
  prevData: PreviousFormatYearData[],
  options?: ExportSyncOptions
) {
  const bsRows: (string | number)[][] = [
    ['PROPERTIES & ASSETS', ...prevData.map(() => '')],
    ...(isVis(options, 'bs_propertyPlantEquipment')
      ? [['Fixed Assets (at cost less depreciation)', ...prevData.map(d => d.fixedAssetsWDV)]]
      : []),
    ['Current Assets', ...prevData.map(d => d.totalCurrentAssets)],
    ...(isVis(options, 'bs_otherReceivables')
      ? [['Advances, Deposits & Pre-payments', ...prevData.map(d => d.advancesDepositsPrepayments)]]
      : []),
    ...(isVis(options, 'bs_closingInventory')
      ? [['Stock & Stores', ...prevData.map(d => d.stockAndStores)]]
      : []),
    ...(isVis(options, 'bs_tradeReceivables')
      ? [['Bills Receivable', ...prevData.map(d => d.billsReceivable)]]
      : []),
    ...(isVis(options, 'bs_cashAndEquivalents')
      ? [[lbl(options, 'bs', 'cashAndEquivalents', 'Cash & Cash Equivalents'), ...prevData.map(d => d.cashAndCashEquivalents)]]
      : []),
    ['Taka', ...prevData.map(d => d.totalAssetsTaka)],
    ['CAPITAL & LIABILITIES', ...prevData.map(() => '')],
    ['Opening Capital', ...prevData.map(d => d.openingCapital)],
    ['Add. Net Profit for the year', ...prevData.map(d => d.netProfitForYear)],
    ['Subtotal', ...prevData.map(d => d.capitalSubtotal)],
    ...(isVis(options, 'eq_dividendsPaid')
      ? [['Less: Drawings', ...prevData.map(d => d.drawings)]]
      : []),
    ['Capital after drawings', ...prevData.map(d => d.capitalAfterDrawings)],
    ['Current Liabilities', ...prevData.map(d => d.totalCurrentLiabilities)],
    ...(isVis(options, 'bs_otherPayables')
      ? [['For Expenses', ...prevData.map(d => d.forExpenses)]]
      : []),
    ...(isVis(options, 'bs_tradePayables')
      ? [['Accounts Payable', ...prevData.map(d => d.accountsPayable)]]
      : []),
    ...(isVis(options, 'bs_shortTermBorrowings')
      ? [['Short Term Loan', ...prevData.map(d => d.shortTermLoan)]]
      : []),
    ...(isVis(options, 'bs_currentTaxPayable')
      ? [['Provision for Taxation', ...prevData.map(d => d.provisionForTaxation)]]
      : []),
    ['Total', ...prevData.map(d => d.totalCapitalAndLiabilities)],
  ];

  const isRows: (string | number)[][] = [
    ...(isVis(options, 'inc_revenue')
      ? [[lbl(options, 'income', 'revenue', 'Sales'), ...prevData.map(d => d.sales)]]
      : []),
    ...(isVis(options, 'inc_costOfSales')
      ? [[`Less- ${lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold')}`, ...prevData.map(d => d.costOfGoodsSold)]]
      : []),
    ['Gross Profit', ...prevData.map(d => d.grossProfit)],
    ['Less- Administrative & Selling Expenses', ...prevData.map(d => d.adminAndSellingTotal)],
    ...(isVis(options, 'inc_operatingExpenses')
      ? [
          ['Salary & Allowances', ...prevData.map(d => d.salaryAndAllowances)],
          ['Telephone & Mobile Bill', ...prevData.map(d => d.telephoneAndMobileBill)],
          ['Postage & Courier', ...prevData.map(d => d.postageAndCourier)],
          ['Electric Charges', ...prevData.map(d => d.electricCharges)],
          ['Entertainment', ...prevData.map(d => d.entertainment)],
          ['Printing & Stationery', ...prevData.map(d => d.printingAndStationery)],
          ['General Expenses', ...prevData.map(d => d.generalExpenses)],
          ['Licence & Renewals', ...prevData.map(d => d.licenceAndRenewals)],
          ['Audit & Professional Fee', ...prevData.map(d => d.auditAndProfessionalFee)],
          ['Travelling & Conveyance', ...prevData.map(d => d.travellingAndConveyance)],
          ['News Paper & Periodicals', ...prevData.map(d => d.newspaperAndPeriodicals)],
          ['Selling Exp.', ...prevData.map(d => d.sellingExp)],
        ]
      : []),
    ...(isVis(options, 'inc_depreciationExpense')
      ? [['Depreciation', ...prevData.map(d => d.depreciation)]]
      : []),
    ['Net Operating Profit', ...prevData.map(d => d.netOperatingProfit)],
    ...(isVis(options, 'inc_interestIncome')
      ? [['Add: Other Income', ...prevData.map(d => d.otherIncome)]]
      : []),
    ['Subtotal', ...prevData.map(d => d.operatingSubtotal)],
    ...(isVis(options, 'inc_financeCost')
      ? [['Less: Financial Exp.', ...prevData.map(d => d.financialExp)]]
      : []),
    ...(isVis(options, 'inc_taxExpense')
      ? [['Less- Income Tax Provision', ...prevData.map(d => d.incomeTaxProvision)]]
      : []),
    ['Net Profit After Income Tax', ...prevData.map(d => d.netProfitAfterIncomeTax)],
  ];

  const cfRows: (string | number)[][] = [
    ['a) CASH FLOWS FROM OPERATING ACTIVITIES:', ...prevData.map(() => '')],
    ...(isVis(options, 'cf_pbt')
      ? [['Net Income', ...prevData.map(d => d.cfNetIncome)]]
      : []),
    ...(isVis(options, 'cf_dep')
      ? [['Add: Depreciation of Tangible Fixed Assets', ...prevData.map(d => d.cfDepreciation)]]
      : []),
    ...(isVis(options, 'cf_wc_receivables')
      ? [['Add: (Increase) or decrease in Current Assets', ...prevData.map(d => d.cfChangeInCurrentAssets)]]
      : []),
    ['Add: (Increase) or decrease in Construction Working Progress', ...prevData.map(d => d.cfChangeInCWIP)],
    ...(isVis(options, 'cf_wc_payables')
      ? [['Add: Increase or (decrease) in Current Liabilities', ...prevData.map(d => d.cfChangeInCurrentLiabilities)]]
      : []),
    ['Net Cash from operating activities', ...prevData.map(d => d.cfNetCashOperating)],
    ['b) CASH FLOWS FROM INVESTING ACTIVITIES:', ...prevData.map(() => '')],
    ...(isVis(options, 'cf_pur_ppe')
      ? [['Acquisition of Tangible Fixed Assets', ...prevData.map(d => d.cfAcquisitionPPE)]]
      : []),
    ...(isVis(options, 'cf_pur_intangible')
      ? [['Disposal of Intangible Assets', ...prevData.map(d => d.cfDisposalIntangible)]]
      : []),
    ['Net Cash used in investing Activities', ...prevData.map(d => d.cfNetCashInvesting)],
    ['c) CASH FLOWS FROM FINANCING ACTIVITIES:', ...prevData.map(() => '')],
    ...(isVis(options, 'cf_dividends_paid')
      ? [['Drawings', ...prevData.map(d => d.cfDrawings)]]
      : []),
    ['Opening Balance', ...prevData.map(d => d.cfOpeningBalance)],
    ['Net Cash used in Financing Activities', ...prevData.map(d => d.cfNetCashFinancing)],
    ['d) Net Increase decrease in cash & cash equivalents (A+B+C)', ...prevData.map(d => d.cfNetIncreaseCash)],
    ['e) Cash & Cash Equivalents at beginning of period', ...prevData.map(d => d.cfCashBeginning)],
    ['f) Cash & Cash Equivalents at end of period (D+E)', ...prevData.map(d => d.cfCashEnding)],
  ];

  const eqRows: (string | number)[][] = [
    ...(isVis(options, 'eq_openingShareCapital')
      ? [['Opening Share Capital', ...prevData.map(d => d.eqOpeningShareCapital)]]
      : []),
    ...(isVis(options, 'eq_shareCapitalIssued')
      ? [['Share Capital Issued', ...prevData.map(d => d.eqShareCapitalIssued)]]
      : []),
    ['Closing Share Capital', ...prevData.map(d => d.eqClosingShareCapital)],
    ...(isVis(options, 'eq_openingRetainedEarnings')
      ? [['Opening Retained Earnings', ...prevData.map(d => d.eqOpeningRetainedEarnings)]]
      : []),
    ...(isVis(options, 'eq_netProfitForYear')
      ? [['Net Profit for the Year', ...prevData.map(d => d.eqNetProfitForYear)]]
      : []),
    ...(isVis(options, 'eq_dividendsPaid')
      ? [['Dividends Paid', ...prevData.map(d => -d.eqDividendsPaid)]]
      : []),
    ['Closing Retained Earnings', ...prevData.map(d => d.eqClosingRetainedEarnings)],
    ...(isVis(options, 'eq_openingReserves')
      ? [['Closing Reserves', ...prevData.map(d => d.eqClosingReserves)]]
      : []),
    ['TOTAL CLOSING EQUITY', ...prevData.map(d => d.eqTotalClosingEquity)],
  ];

  return { bsRows, isRows, cfRows, eqRows };
}

export function exportPreviousFinancialsExcel(
  company: CompanyProfile,
  prevData: PreviousFormatYearData[],
  notes: FinancialNoteItem[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions
) {
  if (prevData.length === 0) return;
  const { filename: resolvedFilename, options } = resolveFilenameAndOptions(
    buildPreviousPairFilename(company, prevData, 'xlsx'),
    filenameOrOptions,
    explicitOptions
  );
  const wb = XLSX.utils.book_new();
  const maxYear = prevData[0].year; // prevData sorted descending
  const unitLabel = getCurrencyUnitCaption(company);
  const colHeaders = ['Particulars', ...prevData.map(d => d.dateLabel)];
  const widths = [46, ...prevData.map(() => 18)];

  // 1. STATEMENT OF FINANCIAL POSITION (dynamic Excel row tracking)
  const bsRows: ExcelCellValue[][] = [
    [company.companyName.toUpperCase()],
    ['STATEMENT OF FINANCIAL POSITION'],
    [`AS AT 31 DECEMBER, ${maxYear}`],
    [unitLabel],
    [],
    colHeaders,
    ['PROPERTIES & ASSETS', ...prevData.map(() => '')],
  ];
  const nextBsRow = () => bsRows.length + 1;
  let fixedAssetsExcelRow: number | null = null;
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    fixedAssetsExcelRow = nextBsRow();
    bsRows.push(['Fixed Assets (at cost less depreciation)', ...prevData.map(d => nCell(d.fixedAssetsWDV))]);
  }
  const caHeaderIdx = bsRows.length;
  const caExcelRow = caHeaderIdx + 1;
  bsRows.push(['Current Assets', ...prevData.map(d => nCell(d.totalCurrentAssets))]);
  const caChildRows: number[] = [];
  if (isVis(options, 'bs_otherReceivables')) {
    caChildRows.push(nextBsRow());
    bsRows.push(['Advances, Deposits & Pre-payments', ...prevData.map(d => nCell(d.advancesDepositsPrepayments))]);
  }
  if (isVis(options, 'bs_closingInventory')) {
    caChildRows.push(nextBsRow());
    bsRows.push(['Stock & Stores', ...prevData.map(d => nCell(d.stockAndStores))]);
  }
  if (isVis(options, 'bs_tradeReceivables')) {
    caChildRows.push(nextBsRow());
    bsRows.push(['Bills Receivable', ...prevData.map(d => nCell(d.billsReceivable))]);
  }
  if (isVis(options, 'bs_cashAndEquivalents')) {
    caChildRows.push(nextBsRow());
    bsRows.push([lbl(options, 'bs', 'cashAndEquivalents', 'Cash & Cash Equivalents'), ...prevData.map(d => nCell(d.cashAndCashEquivalents))]);
  }
  bsRows[caHeaderIdx] = [
    'Current Assets',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), caChildRows), d.totalCurrentAssets)),
  ];
  const totalAssetsInputs = fixedAssetsExcelRow !== null ? [fixedAssetsExcelRow, caExcelRow] : [caExcelRow];
  bsRows.push([
    'Taka',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), totalAssetsInputs), d.totalAssetsTaka)),
  ]);

  bsRows.push(['CAPITAL & LIABILITIES', ...prevData.map(() => '')]);
  const openCapRow = nextBsRow();
  bsRows.push(['Opening Capital', ...prevData.map(d => nCell(d.openingCapital))]);
  const netProfitBsRow = nextBsRow();
  bsRows.push(['Add. Net Profit for the year', ...prevData.map(d => nCell(d.netProfitForYear))]);
  const capSubRow = nextBsRow();
  bsRows.push([
    'Subtotal',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), [openCapRow, netProfitBsRow]), d.capitalSubtotal)),
  ]);
  let drawingsBsRow: number | null = null;
  if (isVis(options, 'eq_dividendsPaid')) {
    drawingsBsRow = nextBsRow();
    bsRows.push(['Less: Drawings', ...prevData.map(d => nCell(d.drawings))]);
  }
  const capAfterDrawRow = nextBsRow();
  bsRows.push([
    'Capital after drawings',
    ...prevData.map((d, i) =>
      fCell(excelDiffExpr(colLetter(i + 1), [capSubRow], drawingsBsRow !== null ? [drawingsBsRow] : []), d.capitalAfterDrawings)
    ),
  ]);
  const clHeaderIdx = bsRows.length;
  const clExcelRow = clHeaderIdx + 1;
  bsRows.push(['Current Liabilities', ...prevData.map(d => nCell(d.totalCurrentLiabilities))]);
  const clChildRows: number[] = [];
  if (isVis(options, 'bs_otherPayables')) {
    clChildRows.push(nextBsRow());
    bsRows.push(['For Expenses', ...prevData.map(d => nCell(d.forExpenses))]);
  }
  if (isVis(options, 'bs_tradePayables')) {
    clChildRows.push(nextBsRow());
    bsRows.push(['Accounts Payable', ...prevData.map(d => nCell(d.accountsPayable))]);
  }
  if (isVis(options, 'bs_shortTermBorrowings')) {
    clChildRows.push(nextBsRow());
    bsRows.push(['Short Term Loan', ...prevData.map(d => nCell(d.shortTermLoan))]);
  }
  if (isVis(options, 'bs_currentTaxPayable')) {
    clChildRows.push(nextBsRow());
    bsRows.push(['Provision for Taxation', ...prevData.map(d => nCell(d.provisionForTaxation))]);
  }
  bsRows[clHeaderIdx] = [
    'Current Liabilities',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), clChildRows), d.totalCurrentLiabilities)),
  ];
  bsRows.push([
    'Total',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), [capAfterDrawRow, clExcelRow]), d.totalCapitalAndLiabilities)),
  ]);
  const wsBS = XLSX.utils.aoa_to_sheet(bsRows);
  applyColumnWidths(wsBS, widths);
  XLSX.utils.book_append_sheet(wb, wsBS, '1. Financial Position');

  // 2. STATEMENT OF COMPREHENSIVE INCOME (dynamic Excel row tracking)
  const isRows: ExcelCellValue[][] = [
    [company.companyName.toUpperCase()],
    ['Statement of Comprehensive Income'],
    [`For the year ended 31 December, ${maxYear}`],
    [unitLabel],
    [],
    colHeaders,
  ];
  const nextIsRow = () => isRows.length + 1;
  let salesExcelRow: number | null = null;
  if (isVis(options, 'inc_revenue')) {
    salesExcelRow = nextIsRow();
    isRows.push([lbl(options, 'income', 'revenue', 'Sales'), ...prevData.map(d => nCell(d.sales))]);
  }
  let cogsExcelRow: number | null = null;
  if (isVis(options, 'inc_costOfSales')) {
    cogsExcelRow = nextIsRow();
    isRows.push([`Less- ${lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold')}`, ...prevData.map(d => nCell(d.costOfGoodsSold))]);
  }
  const gpExcelRow = nextIsRow();
  isRows.push([
    'Gross Profit',
    ...prevData.map((d, i) =>
      fCell(
        excelDiffExpr(
          colLetter(i + 1),
          salesExcelRow !== null ? [salesExcelRow] : [],
          cogsExcelRow !== null ? [cogsExcelRow] : []
        ),
        d.grossProfit
      )
    ),
  ]);
  const adminHeaderIdx = isRows.length;
  const adminExcelRow = adminHeaderIdx + 1;
  isRows.push(['Less- Administrative & Selling Expenses', ...prevData.map(d => nCell(d.adminAndSellingTotal))]);
  const adminChildRows: number[] = [];
  if (isVis(options, 'inc_operatingExpenses')) {
    const opexLines: [string, keyof PreviousFormatYearData][] = [
      ['Salary & Allowances', 'salaryAndAllowances'],
      ['Telephone & Mobile Bill', 'telephoneAndMobileBill'],
      ['Postage & Courier', 'postageAndCourier'],
      ['Electric Charges', 'electricCharges'],
      ['Entertainment', 'entertainment'],
      ['Printing & Stationery', 'printingAndStationery'],
      ['General Expenses', 'generalExpenses'],
      ['Licence & Renewals', 'licenceAndRenewals'],
      ['Audit & Professional Fee', 'auditAndProfessionalFee'],
      ['Travelling & Conveyance', 'travellingAndConveyance'],
      ['News Paper & Periodicals', 'newspaperAndPeriodicals'],
      ['Selling Exp.', 'sellingExp'],
    ];
    opexLines.forEach(([lineLabel, field]) => {
      adminChildRows.push(nextIsRow());
      isRows.push([lineLabel, ...prevData.map(d => nCell(Number(d[field])))]);
    });
  }
  if (isVis(options, 'inc_depreciationExpense')) {
    adminChildRows.push(nextIsRow());
    isRows.push(['Depreciation', ...prevData.map(d => nCell(d.depreciation))]);
  }
  isRows[adminHeaderIdx] = [
    'Less- Administrative & Selling Expenses',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), adminChildRows), d.adminAndSellingTotal)),
  ];
  const netOpProfitRow = nextIsRow();
  isRows.push([
    'Net Operating Profit',
    ...prevData.map((d, i) => fCell(excelDiffExpr(colLetter(i + 1), [gpExcelRow], [adminExcelRow]), d.netOperatingProfit)),
  ]);
  let otherIncRow: number | null = null;
  if (isVis(options, 'inc_interestIncome')) {
    otherIncRow = nextIsRow();
    isRows.push(['Add: Other Income', ...prevData.map(d => nCell(d.otherIncome))]);
  }
  const opSubRow = nextIsRow();
  isRows.push([
    'Subtotal',
    ...prevData.map((d, i) =>
      fCell(
        excelSumRows(colLetter(i + 1), otherIncRow !== null ? [netOpProfitRow, otherIncRow] : [netOpProfitRow]),
        d.operatingSubtotal
      )
    ),
  ]);
  const taxAndFinSubRows: number[] = [];
  if (isVis(options, 'inc_financeCost')) {
    taxAndFinSubRows.push(nextIsRow());
    isRows.push(['Less: Financial Exp.', ...prevData.map(d => nCell(d.financialExp))]);
  }
  if (isVis(options, 'inc_taxExpense')) {
    taxAndFinSubRows.push(nextIsRow());
    isRows.push(['Less- Income Tax Provision', ...prevData.map(d => nCell(d.incomeTaxProvision))]);
  }
  isRows.push([
    'Net Profit After Income Tax',
    ...prevData.map((d, i) =>
      fCell(excelDiffExpr(colLetter(i + 1), [opSubRow], taxAndFinSubRows), d.netProfitAfterIncomeTax)
    ),
  ]);
  const wsIS = XLSX.utils.aoa_to_sheet(isRows);
  applyColumnWidths(wsIS, widths);
  XLSX.utils.book_append_sheet(wb, wsIS, '2. Comprehensive Income');

  // 3. STATEMENT OF CASH FLOWS (dynamic Excel row tracking)
  const cfRows: ExcelCellValue[][] = [
    [company.companyName.toUpperCase()],
    ['STATEMENT OF CASH FLOWS'],
    [`For the year ended 31 December ${maxYear}`],
    [unitLabel],
    [],
    colHeaders,
    ['a) CASH FLOWS FROM OPERATING ACTIVITIES:', ...prevData.map(() => '')],
  ];
  const nextCfRow = () => cfRows.length + 1;
  const cfOpChildRows: number[] = [];
  if (isVis(options, 'cf_pbt')) {
    cfOpChildRows.push(nextCfRow());
    cfRows.push(['Net Income', ...prevData.map(d => nCell(d.cfNetIncome))]);
  }
  if (isVis(options, 'cf_dep')) {
    cfOpChildRows.push(nextCfRow());
    cfRows.push(['Add: Depreciation of Tangible Fixed Assets', ...prevData.map(d => nCell(d.cfDepreciation))]);
  }
  if (isVis(options, 'cf_wc_receivables')) {
    cfOpChildRows.push(nextCfRow());
    cfRows.push(['Add: (Increase) or decrease in Current Assets', ...prevData.map(d => nCell(d.cfChangeInCurrentAssets))]);
  }
  cfOpChildRows.push(nextCfRow());
  cfRows.push(['Add: (Increase) or decrease in Construction Working Progress', ...prevData.map(d => nCell(d.cfChangeInCWIP))]);
  if (isVis(options, 'cf_wc_payables')) {
    cfOpChildRows.push(nextCfRow());
    cfRows.push(['Add: Increase or (decrease) in Current Liabilities', ...prevData.map(d => nCell(d.cfChangeInCurrentLiabilities))]);
  }
  const cfNetOpRow = nextCfRow();
  cfRows.push([
    'Net Cash from operating activities',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfOpChildRows), d.cfNetCashOperating)),
  ]);

  cfRows.push(['b) CASH FLOWS FROM INVESTING ACTIVITIES:', ...prevData.map(() => '')]);
  const cfInvChildRows: number[] = [];
  if (isVis(options, 'cf_pur_ppe')) {
    cfInvChildRows.push(nextCfRow());
    cfRows.push(['Acquisition of Tangible Fixed Assets', ...prevData.map(d => nCell(d.cfAcquisitionPPE))]);
  }
  if (isVis(options, 'cf_pur_intangible')) {
    cfInvChildRows.push(nextCfRow());
    cfRows.push(['Disposal of Intangible Assets', ...prevData.map(d => nCell(d.cfDisposalIntangible))]);
  }
  const cfNetInvRow = nextCfRow();
  cfRows.push([
    'Net Cash used in investing Activities',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfInvChildRows), d.cfNetCashInvesting)),
  ]);

  cfRows.push(['c) CASH FLOWS FROM FINANCING ACTIVITIES:', ...prevData.map(() => '')]);
  const cfFinChildRows: number[] = [];
  if (isVis(options, 'cf_dividends_paid')) {
    cfFinChildRows.push(nextCfRow());
    cfRows.push(['Drawings', ...prevData.map(d => nCell(d.cfDrawings))]);
  }
  cfFinChildRows.push(nextCfRow());
  cfRows.push(['Opening Balance', ...prevData.map(d => nCell(d.cfOpeningBalance))]);
  const cfNetFinRow = nextCfRow();
  cfRows.push([
    'Net Cash used in Financing Activities',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfFinChildRows), d.cfNetCashFinancing)),
  ]);

  const cfNetIncRow = nextCfRow();
  cfRows.push([
    'd) Net Increase decrease in cash & cash equivalents (A+B+C)',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), [cfNetOpRow, cfNetInvRow, cfNetFinRow]), d.cfNetIncreaseCash)),
  ]);
  const cfBegCashRow = nextCfRow();
  cfRows.push(['e) Cash & Cash Equivalents at beginning of period', ...prevData.map(d => nCell(d.cfCashBeginning))]);
  cfRows.push([
    'f) Cash & Cash Equivalents at end of period (D+E)',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), [cfNetIncRow, cfBegCashRow]), d.cfCashEnding)),
  ]);
  const wsCF = XLSX.utils.aoa_to_sheet(cfRows);
  applyColumnWidths(wsCF, [56, ...prevData.map(() => 18)]);
  XLSX.utils.book_append_sheet(wb, wsCF, '3. Cash Flows');

  // 4. STATEMENT OF CHANGES IN EQUITY (dynamic Excel row tracking)
  const eqRows: ExcelCellValue[][] = [
    [company.companyName.toUpperCase()],
    ['STATEMENT OF CHANGES IN EQUITY'],
    [`For the year ended 31 December ${maxYear}`],
    [unitLabel],
    [],
    colHeaders,
  ];
  const nextEqRow = () => eqRows.length + 1;
  const scInputs: number[] = [];
  if (isVis(options, 'eq_openingShareCapital')) {
    scInputs.push(nextEqRow());
    eqRows.push(['Opening Share Capital', ...prevData.map(d => nCell(d.eqOpeningShareCapital))]);
  }
  if (isVis(options, 'eq_shareCapitalIssued')) {
    scInputs.push(nextEqRow());
    eqRows.push(['Share Capital Issued', ...prevData.map(d => nCell(d.eqShareCapitalIssued))]);
  }
  const closeScRow = nextEqRow();
  eqRows.push([
    'Closing Share Capital',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), scInputs), d.eqClosingShareCapital)),
  ]);
  const reInputs: number[] = [];
  if (isVis(options, 'eq_openingRetainedEarnings')) {
    reInputs.push(nextEqRow());
    eqRows.push(['Opening Retained Earnings', ...prevData.map(d => nCell(d.eqOpeningRetainedEarnings))]);
  }
  if (isVis(options, 'eq_netProfitForYear')) {
    reInputs.push(nextEqRow());
    eqRows.push(['Net Profit for the Year', ...prevData.map(d => nCell(d.eqNetProfitForYear))]);
  }
  if (isVis(options, 'eq_dividendsPaid')) {
    reInputs.push(nextEqRow());
    eqRows.push(['Dividends Paid', ...prevData.map(d => nCell(-d.eqDividendsPaid))]);
  }
  const closeReRow = nextEqRow();
  eqRows.push([
    'Closing Retained Earnings',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), reInputs), d.eqClosingRetainedEarnings)),
  ]);
  const eqTotalInputs = [closeScRow, closeReRow];
  if (isVis(options, 'eq_openingReserves')) {
    eqTotalInputs.push(nextEqRow());
    eqRows.push(['Closing Reserves', ...prevData.map(d => nCell(d.eqClosingReserves))]);
  }
  eqRows.push([
    'TOTAL CLOSING EQUITY',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), eqTotalInputs), d.eqTotalClosingEquity)),
  ]);
  const wsEQ = XLSX.utils.aoa_to_sheet(eqRows);
  applyColumnWidths(wsEQ, widths);
  XLSX.utils.book_append_sheet(wb, wsEQ, '4. Changes in Equity');

  // 5. 1.00 FIXED ASSETS SCHEDULE & NOTES
  const faAsc = [...prevData].sort((a, b) => a.year - b.year);
  const faRows: ExcelCellValue[][] = [
    [company.companyName.toUpperCase()],
    [`For the Year ended 31 December ${maxYear}`],
    ['1.00 FIXED ASSETS SCHEDULE'],
    [],
  ];
  faAsc.forEach(d => {
    const startIdx = faRows.length + 1;
    faRows.push([
      'Particulars',
      `Opening at 01.01.${d.year}`,
      'Addition during the year',
      `Total at 31.12.${d.year}`,
      'Rate %',
      'Depreciation charged for the year',
      `W.D.V. at 31.12.${d.year}`,
    ]);
    const itemRow = startIdx + 1;
    faRows.push([
      'Fixed Assets',
      nCell(d.faOpening),
      nCell(d.faAddition - d.faDisposal),
      fCell(`B${itemRow}+C${itemRow}`, d.faTotal),
      nCell(d.faRatePct),
      nCell(d.faDepreciationCharged),
      nCell(d.faWDV),
    ]);
    faRows.push([
      'Taka',
      fCell(`B${itemRow}`, d.faOpening),
      fCell(`C${itemRow}`, d.faAddition - d.faDisposal),
      fCell(`D${itemRow}`, d.faTotal),
      '',
      fCell(`F${itemRow}`, d.faDepreciationCharged),
      fCell(`G${itemRow}`, d.faWDV),
    ]);
    faRows.push([]);
  });

  faRows.push(['NOTES TO THE FINANCIAL STATEMENTS:']);
  notes.forEach(note => {
    faRows.push([`Note ${note.noteNumber}: ${note.title}`]);
    faRows.push([note.textDescription]);
    if (note.tableData) {
      faRows.push(note.tableData.headers);
      note.tableData.rows.forEach(r => faRows.push([r.label, ...r.values]));
      if (note.tableData.totalRow) {
        faRows.push([note.tableData.totalRow.label, ...note.tableData.totalRow.values]);
      }
    }
    faRows.push([]);
  });

  const wsFA = XLSX.utils.aoa_to_sheet(faRows);
  applyColumnWidths(wsFA, [34, 20, 22, 20, 12, 26, 20]);
  XLSX.utils.book_append_sheet(wb, wsFA, '5. Fixed Assets & Notes');

  XLSX.writeFile(wb, resolvedFilename);
}

export async function exportPreviousFinancialsWord(
  company: CompanyProfile,
  prevData: PreviousFormatYearData[],
  notes: FinancialNoteItem[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions
) {
  if (prevData.length === 0) return;
  const { filename: resolvedFilename, options } = resolveFilenameAndOptions(
    buildPreviousPairFilename(company, prevData, 'docx'),
    filenameOrOptions,
    explicitOptions
  );
  const maxYear = prevData[0].year;
  const unitLabel = getCurrencyUnitCaption(company);
  const colHeaders = ['Particulars', ...prevData.map(d => d.dateLabel)];

  const { bsRows, isRows, cfRows, eqRows } = buildPreviousTextTables(prevData, options);
  const faAsc = [...prevData].sort((a, b) => a.year - b.year);

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: company.companyName.toUpperCase(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'STATEMENT OF FINANCIAL POSITION', bold: true, size: 24 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `AS AT 31 DECEMBER, ${maxYear}`, bold: true, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: unitLabel, size: 16 })], alignment: AlignmentType.RIGHT }),
          createDocxTable(colHeaders, bsRows),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName.toUpperCase(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'Statement of Comprehensive Income', bold: true, size: 24 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `For the year ended 31 December, ${maxYear}`, bold: true, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: unitLabel, size: 16 })], alignment: AlignmentType.RIGHT }),
          createDocxTable(colHeaders, isRows),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName.toUpperCase(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'STATEMENT OF CASH FLOWS', bold: true, size: 24 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `For the year ended 31 December ${maxYear}`, bold: true, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: unitLabel, size: 16 })], alignment: AlignmentType.RIGHT }),
          createDocxTable(colHeaders, cfRows),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName.toUpperCase(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'STATEMENT OF CHANGES IN EQUITY', bold: true, size: 24 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `For the year ended 31 December ${maxYear}`, bold: true, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: unitLabel, size: 16 })], alignment: AlignmentType.RIGHT }),
          createDocxTable(colHeaders, eqRows),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName.toUpperCase(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `For the Year ended 31 December ${maxYear}`, bold: true, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: '1.00 FIXED ASSETS SCHEDULE', bold: true, size: 20 })] }),
          ...faAsc.flatMap(d => [
            createDocxTable(
              ['Particulars', `Opening at 01.01.${d.year}`, 'Addition during the year', `Total at 31.12.${d.year}`, 'Rate %', 'Depreciation charged for the year', `W.D.V. at 31.12.${d.year}`],
              [
                ['Fixed Assets', d.faOpening, d.faAddition - d.faDisposal, d.faTotal, d.faRatePct, d.faDepreciationCharged, d.faWDV],
                ['Taka', d.faOpening, d.faAddition - d.faDisposal, d.faTotal, '', d.faDepreciationCharged, d.faWDV],
              ]
            ),
            new Paragraph({ text: '' }),
          ]),

          new Paragraph({ children: [new TextRun({ text: 'NOTES TO THE FINANCIAL STATEMENTS:', bold: true, size: 20 })] }),
          ...notes.flatMap(n => [
            new Paragraph({ children: [new TextRun({ text: `Note ${n.noteNumber}: ${n.title}`, bold: true, size: 18 })] }),
            new Paragraph({ children: [new TextRun({ text: n.textDescription, size: 17 })] }),
            ...(n.tableData
              ? [
                  createDocxTable(n.tableData.headers, [
                    ...n.tableData.rows.map(r => [r.label, ...r.values]),
                    ...(n.tableData.totalRow ? [[n.tableData.totalRow.label, ...n.tableData.totalRow.values]] : []),
                  ]),
                ]
              : []),
            new Paragraph({ text: '' }),
          ]),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, resolvedFilename);
}

export function exportPreviousFinancialsPdf(
  company: CompanyProfile,
  prevData: PreviousFormatYearData[],
  notes: FinancialNoteItem[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions
) {
  if (prevData.length === 0) return;
  const { filename: resolvedFilename, options } = resolveFilenameAndOptions(
    buildPreviousPairFilename(company, prevData, 'pdf'),
    filenameOrOptions,
    explicitOptions
  );
  const orientation = prevData.length > 4 ? 'landscape' : 'portrait';
  const doc = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
  const pageW = orientation === 'landscape' ? 297 : 210;
  const pageH = orientation === 'landscape' ? 210 : 297;
  const maxYear = prevData[0].year;
  const unitLabel = getCurrencyUnitCaption(company);
  const colHeaders = ['Particulars', ...prevData.map(d => d.dateLabel)];

  const { bsRows, isRows, cfRows, eqRows } = buildPreviousTextTables(prevData, options);

  const drawClassicHeader = (stmtTitle: string, subtitle: string, showUnitCaption = true) => {
    doc.setTextColor(15, 23, 42);
    doc.setFont('times', 'bold');
    doc.setFontSize(13);
    doc.text(company.companyName.toUpperCase(), pageW / 2, 16, { align: 'center' });
    doc.setFontSize(12);
    doc.text(stmtTitle, pageW / 2, 23, { align: 'center' });
    doc.setFontSize(9.5);
    doc.text(subtitle, pageW / 2, 29, { align: 'center' });
    if (showUnitCaption) {
      doc.setFont('times', 'normal');
      doc.setFontSize(8.5);
      doc.text(unitLabel, pageW - 14, 34.5, { align: 'right' });
    }
  };

  const boldLabels = new Set([
    'Taka', 'Subtotal', 'Capital after drawings', 'Total', 'Current Assets', 'Current Liabilities',
    'Gross Profit', 'Less- Administrative & Selling Expenses', 'Net Operating Profit', 'Net Profit After Income Tax',
    'Net Cash from operating activities', 'Net Cash used in investing Activities', 'Net Cash used in Financing Activities',
    'f) Cash & Cash Equivalents at end of period (D+E)', 'TOTAL CLOSING EQUITY',
  ]);

  const renderTablePage = (stmtTitle: string, subtitle: string, rows: (string | number)[][], isFirst = false) => {
    if (!isFirst) doc.addPage();
    drawClassicHeader(stmtTitle, subtitle, true);
    autoTable(doc, {
      head: [colHeaders],
      body: rows.map(r => r.map((c, idx) => idx === 0 ? String(c) : (typeof c === 'number' ? fmtExactNum(c) : String(c)))),
      startY: 36,
      theme: 'grid',
      styles: { font: 'times', fontSize: 8.8, cellPadding: 1.8, textColor: [15, 23, 42], lineColor: [51, 65, 85], lineWidth: 0.2 },
      headStyles: { fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', halign: 'center', lineColor: [51, 65, 85], lineWidth: 0.25 },
      columnStyles: {
        0: { halign: 'left', cellWidth: orientation === 'landscape' ? 115 : 92 },
        ...Object.fromEntries(prevData.map((_, i) => [i + 1, { halign: 'right' }])),
      },
      didParseCell: (data) => {
        if (data.section === 'body') {
          const rawRow = data.row.raw;
          const rowLabel = Array.isArray(rawRow) ? String(rawRow[0] ?? '') : '';
          if (boldLabels.has(rowLabel) || rowLabel.startsWith('a)') || rowLabel.startsWith('b)') || rowLabel.startsWith('c)')) {
            data.cell.styles.fontStyle = 'bold';
          }
          if (data.column.index === 0 && (rowLabel === 'Taka' || rowLabel === 'Total')) {
            data.cell.styles.halign = 'right';
          }
        }
      },
      margin: { left: 14, right: 14 },
    });
  };

  renderTablePage('STATEMENT OF FINANCIAL POSITION', `AS AT 31 DECEMBER, ${maxYear}`, bsRows, true);
  renderTablePage('Statement of Comprehensive Income', `For the year ended 31 December, ${maxYear}`, isRows);
  renderTablePage('STATEMENT OF CASH FLOWS', `For the year ended 31 December ${maxYear}`, cfRows);
  renderTablePage('STATEMENT OF CHANGES IN EQUITY', `For the year ended 31 December ${maxYear}`, eqRows);

  // Page 5: 1.00 FIXED ASSETS SCHEDULE & NOTES
  doc.addPage();
  doc.setTextColor(15, 23, 42);
  doc.setFont('times', 'bold');
  doc.setFontSize(13);
  doc.text(company.companyName.toUpperCase(), pageW / 2, 16, { align: 'center' });
  doc.setFontSize(9.5);
  doc.text(`For the Year ended 31 December ${maxYear}`, pageW / 2, 23, { align: 'center' });
  doc.setFontSize(10);
  doc.text('1.00 FIXED ASSETS SCHEDULE', 14, 31);

  const faAsc = [...prevData].sort((a, b) => a.year - b.year);
  let yCursor = 34;
  faAsc.forEach((d) => {
    autoTable(doc, {
      head: [[
        'Particulars',
        `Opening at 01.01.${d.year}`,
        'Addition during the year',
        `Total at 31.12.${d.year}`,
        'Rate %',
        'Depreciation charged for the year',
        `W.D.V. at 31.12.${d.year}`,
      ]],
      body: [
        ['Fixed Assets', fmtExactNum(d.faOpening), fmtExactNum(d.faAddition - d.faDisposal), fmtExactNum(d.faTotal), `${d.faRatePct}`, fmtExactNum(d.faDepreciationCharged), fmtExactNum(d.faWDV)],
        ['Taka', fmtExactNum(d.faOpening), fmtExactNum(d.faAddition - d.faDisposal), fmtExactNum(d.faTotal), '', fmtExactNum(d.faDepreciationCharged), fmtExactNum(d.faWDV)],
      ],
      startY: yCursor,
      theme: 'grid',
      styles: { font: 'times', fontSize: 8.2, cellPadding: 1.6, textColor: [15, 23, 42], lineColor: [51, 65, 85], lineWidth: 0.2 },
      headStyles: { fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', halign: 'center', lineColor: [51, 65, 85], lineWidth: 0.25 },
      columnStyles: {
        0: { halign: 'left' },
        1: { halign: 'right' },
        2: { halign: 'right' },
        3: { halign: 'right' },
        4: { halign: 'center' },
        5: { halign: 'right' },
        6: { halign: 'right' },
      },
      didParseCell: (data) => {
        if (data.section === 'body' && data.row.index === 1) {
          data.cell.styles.fontStyle = 'bold';
        }
      },
      margin: { left: 14, right: 14 },
    });
    yCursor = ((doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? yCursor + 22) + 6;
  });

  if (yCursor > pageH - 40) {
    doc.addPage();
    yCursor = 18;
  }
  doc.setFont('times', 'bold');
  doc.setFontSize(10);
  doc.text('NOTES TO THE FINANCIAL STATEMENTS:', 14, yCursor + 4);
  yCursor += 10;

  notes.forEach(note => {
    if (yCursor > pageH - 35) {
      doc.addPage();
      yCursor = 18;
    }
    doc.setFont('times', 'bold');
    doc.setFontSize(9);
    doc.text(`Note ${note.noteNumber}: ${note.title}`, 14, yCursor);
    yCursor += 4.5;
    doc.setFont('times', 'normal');
    doc.setFontSize(8.5);
    const lines = doc.splitTextToSize(note.textDescription, pageW - 28);
    doc.text(lines, 14, yCursor);
    yCursor += lines.length * 3.8 + 2;

    if (note.tableData) {
      if (yCursor > pageH - 40) {
        doc.addPage();
        yCursor = 18;
      }
      const nCols = note.tableData.headers.length;
      const bodyRows = [
        ...note.tableData.rows.map(r => [r.label, ...r.values.map(v => String(v))]),
        ...(note.tableData.totalRow ? [[note.tableData.totalRow.label, ...note.tableData.totalRow.values.map(v => String(v))]] : []),
      ];
      autoTable(doc, {
        head: [note.tableData.headers],
        body: bodyRows,
        startY: yCursor,
        theme: 'grid',
        styles: { font: 'times', fontSize: 8, cellPadding: 1.4, textColor: [15, 23, 42], lineColor: [148, 163, 184], lineWidth: 0.15 },
        headStyles: { fillColor: [248, 250, 252], textColor: [15, 23, 42], fontStyle: 'bold', halign: 'left' },
        columnStyles: Object.fromEntries(
          Array.from({ length: nCols }, (_, idx) => [idx, { halign: idx === 0 ? 'left' : 'right' }])
        ),
        didParseCell: (data) => {
          if (data.section === 'body' && note.tableData?.totalRow && data.row.index === bodyRows.length - 1) {
            data.cell.styles.fontStyle = 'bold';
          }
        },
        margin: { left: 14, right: 14 },
      });
      yCursor = ((doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? yCursor + 18) + 5;
    } else {
      yCursor += 2;
    }
  });

  doc.save(resolvedFilename);
}

// Batch exporters for all overlapping Previous Comparative Pairs (e.g. 2023 & 2024 + 2024 & 2025)
export function exportAllPreviousComparativePairsExcel(
  company: CompanyProfile,
  pairs: { earlierYear: number; laterYear: number; data: PreviousFormatYearData[] }[],
  getScopedNotes: (pairData: PreviousFormatYearData[]) => FinancialNoteItem[],
  options?: ExportSyncOptions
) {
  pairs.forEach((pair, idx) => {
    setTimeout(() => {
      exportPreviousFinancialsExcel(company, pair.data, getScopedNotes(pair.data), options);
    }, idx * 250);
  });
}

export function exportAllPreviousComparativePairsWord(
  company: CompanyProfile,
  pairs: { earlierYear: number; laterYear: number; data: PreviousFormatYearData[] }[],
  getScopedNotes: (pairData: PreviousFormatYearData[]) => FinancialNoteItem[],
  options?: ExportSyncOptions
) {
  pairs.forEach((pair, idx) => {
    setTimeout(() => {
      void exportPreviousFinancialsWord(company, pair.data, getScopedNotes(pair.data), options);
    }, idx * 250);
  });
}

export function exportAllPreviousComparativePairsPdf(
  company: CompanyProfile,
  pairs: { earlierYear: number; laterYear: number; data: PreviousFormatYearData[] }[],
  getScopedNotes: (pairData: PreviousFormatYearData[]) => FinancialNoteItem[],
  options?: ExportSyncOptions
) {
  pairs.forEach((pair, idx) => {
    setTimeout(() => {
      exportPreviousFinancialsPdf(company, pair.data, getScopedNotes(pair.data), options);
    }, idx * 250);
  });
}

// ============================================================================
// PROJECTED FINANCIALS EXACT-FORMAT EXPORTS (.xlsx, .docx, .pdf)
// Strictly preserves layout, Notes column, Amount In Taka / YYYY headers,
// Proprietor signature, Notes 1..13.00, and Schedule-A PPE from "Projected Financials.pdf"
// ============================================================================

function getProjectedAmountCaption(company: CompanyProfile): string {
  return company.reportingCurrency === '৳' ? 'Amount In Taka' : `Amount In ${company.reportingCurrency}`;
}

function buildProjectedTextTables(
  projData: ProjectedFormatYearData[],
  options?: ExportSyncOptions
) {
  const salesLbl = lbl(options, 'income', 'revenue', 'Sales');
  const cogsLbl = `Less: ${lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold')}`;
  const opexLbl = lbl(options, 'income', 'operatingExpenses', 'Administrative & Selling Expenses');
  const trRecLbl = lbl(options, 'bs', 'tradeReceivables', 'Trade & Other Receivables');
  const cashLbl = lbl(options, 'bs', 'cashAndEquivalents', 'Cash & Cash Equivalents');

  const bsRows: (string | number)[][] = [
    ['Assets', '', ...projData.map(() => '')],
    ['Non Current Assets', '', ...projData.map(() => '')],
    ...(isVis(options, 'bs_propertyPlantEquipment')
      ? [['Property Plant & Equiptments', 'Schedule-A', ...projData.map(d => d.propertyPlantEquipments)]]
      : []),
    ['Total Non Current Assets', '', ...projData.map(d => d.propertyPlantEquipments)],
    ['Current Assets', '', ...projData.map(() => '')],
    ...(isVis(options, 'bs_closingInventory')
      ? [['Inventories', '3.00', ...projData.map(d => d.inventories)]]
      : []),
    ...(isVis(options, 'bs_tradeReceivables')
      ? [[trRecLbl, '4.00', ...projData.map(d => d.tradeAndOtherReceivables)]]
      : []),
    ...(isVis(options, 'bs_otherReceivables')
      ? [['Advance, Deposit & Prepayments', '5.00', ...projData.map(d => d.advanceDepositPrepayments)]]
      : []),
    ...(isVis(options, 'bs_cashAndEquivalents')
      ? [[cashLbl, '6.00', ...projData.map(d => d.cashAndCashEquivalents)]]
      : []),
    ['Total Current Assets', '', ...projData.map(d => d.totalCurrentAssets)],
    ['Total Assets', '', ...projData.map(d => d.totalAssets)],
    ['Equity & Liabilities', '', ...projData.map(() => '')],
    ['Capital', '', ...projData.map(d => d.capital)],
    ['Additional Capital', '', ...projData.map(d => d.additionalCapital)],
    ['Profit during the year', '', ...projData.map(d => d.profitDuringYear)],
    ['Subtotal', '', ...projData.map(d => d.equitySubtotal)],
    ...(isVis(options, 'eq_dividendsPaid')
      ? [['Drawings', '', ...projData.map(d => d.drawings)]]
      : []),
    ['Closing Capital', '', ...projData.map(d => d.closingCapital)],
    ['Non-Current Liabilities', '', ...projData.map(() => '')],
    ...(isVis(options, 'bs_longTermBorrowings')
      ? [['Long Term Loan', '7.00', ...projData.map(d => d.longTermLoan)]]
      : []),
    ['Current Liabilities', '', ...projData.map(() => '')],
    ...(isVis(options, 'bs_tradePayables')
      ? [['Trade & Other Payables', '8.00', ...projData.map(d => d.tradeAndOtherPayables)]]
      : []),
    ...(isVis(options, 'bs_currentTaxPayable')
      ? [['Provisions For Tax', '9.00', ...projData.map(d => d.provisionsForTax)]]
      : []),
    ...(isVis(options, 'bs_shortTermBorrowings')
      ? [['Short Term Loan', '10.00', ...projData.map(d => d.shortTermLoan)]]
      : []),
    ...(isVis(options, 'bs_otherPayables')
      ? [['Liabilities for exp.', '', ...projData.map(d => d.liabilitiesForExp)]]
      : []),
    ['Total Current Liabilities', '', ...projData.map(d => d.totalCurrentLiabilities)],
    ['Total Equity & Liabilities', '', ...projData.map(d => d.totalEquityAndLiabilities)],
  ];

  const plRows: (string | number)[][] = [
    ...(isVis(options, 'inc_revenue') ? [[salesLbl, '', ...projData.map(d => d.sales)]] : []),
    ...(isVis(options, 'inc_costOfSales') ? [[cogsLbl, '11.00', ...projData.map(d => d.costOfGoodsSold)]] : []),
    ['Gross Profit', '', ...projData.map(d => d.grossProfit)],
    ...(isVis(options, 'inc_interestIncome') ? [['Add; Others Income', '', ...projData.map(d => d.othersIncome)]] : []),
    ['Subtotal', '', ...projData.map(d => d.grossProfitWithOtherIncome)],
    ['Less: Operating Expenses:', '', ...projData.map(() => '')],
    ...(isVis(options, 'inc_operatingExpenses') ? [[opexLbl, '12.00', ...projData.map(d => d.adminAndSellingExpenses)]] : []),
    ['Total Operating Profit/(Loss)', '', ...projData.map(d => d.totalOperatingProfit)],
    ...(isVis(options, 'inc_financeCost') ? [['Less: Financial Expenses', '13.00', ...projData.map(d => d.financialExpenses)]] : []),
    ['Net Profit/(Loss) Before Income Tax', '', ...projData.map(d => d.netProfitBeforeTax)],
    ...(isVis(options, 'inc_taxExpense') ? [['Less: Provision For Income Tax', '', ...projData.map(d => d.provisionForIncomeTax)]] : []),
    ['Net Profit/(Loss) After Income Tax', '', ...projData.map(d => d.netProfitAfterTax)],
  ];

  const cfRows: (string | number)[][] = [
    ...(isVis(options, 'cf_pbt') ? [['Net Income', ...projData.map(d => d.cfNetIncome)]] : []),
    ...(isVis(options, 'cf_wc_inventory') ? [['(Increase)/Decrease In Inventory', ...projData.map(d => d.cfChangeInventory)]] : []),
    ...(isVis(options, 'cf_wc_receivables') ? [['(Increase)/Decrease In Trade & Other Receivables', ...projData.map(d => d.cfChangeReceivables)]] : []),
    ...(isVis(options, 'cf_wc_other_rec') ? [['(Increase)/Decrease in Advance, Deposit & Prepayments', ...projData.map(d => d.cfChangeAdvanceDeposit)]] : []),
    ...(isVis(options, 'cf_wc_other_ca') ? [['Increase/(Decrease) in Current Assets', ...projData.map(d => d.cfChangeCurrentAssets)]] : []),
    ['Increase/(Decrease) in Current Liabilities', ...projData.map(d => d.cfChangeCurrentLiabilities)],
    ...(isVis(options, 'cf_wc_payables') ? [['Increase/(Decrease) in Trade & Other Payables', ...projData.map(d => d.cfChangeTradePayables)]] : []),
    ...(isVis(options, 'cf_tax_paid') ? [['Increase/(Decrease) In Provisions For Tax', ...projData.map(d => d.cfChangeProvisionsTax)]] : []),
    ...(isVis(options, 'cf_wc_other_pay') ? [['Increase/(Decrease) Liabilities For Expense', ...projData.map(d => d.cfChangeLiabilitiesExp)]] : []),
    ...(isVis(options, 'cf_dep') ? [['Depreciation', ...projData.map(d => d.cfDepreciation)]] : []),
    ['Net Cash Used by Operating Activities', ...projData.map(d => d.cfNetCashOperating)],
    ['2. CASH FLOW FROM INVESTING ACTIVITIES:', ...projData.map(() => '')],
    ['Inflows:', ...projData.map(() => '')],
    ...(isVis(options, 'cf_sale_ppe') ? [['Sale Of Property, Plant & Equipment', ...projData.map(d => d.cfSalePPE)]] : []),
    ['Capital', ...projData.map(d => d.cfInvestingCapitalInflow)],
    ['Outflows:', ...projData.map(() => '')],
    ['Preliminary Expenses', ...projData.map(d => d.cfPreliminaryExpenses)],
    ...(isVis(options, 'cf_pur_ppe') ? [['Purchase Property, Plant & Equipment', ...projData.map(d => d.cfPurchasePPE)]] : []),
    ['Purchase Investment In Debt Or Securities', ...projData.map(d => d.cfPurchaseInvestment)],
    ['Net Cash Used by Investing Activities', ...projData.map(d => d.cfNetCashInvesting)],
    ['3. CASH FLOW FROM FINANCIAL ACTIVITIES:', ...projData.map(() => '')],
    ['Inflows:', ...projData.map(() => '')],
    ...(isVis(options, 'cf_borrow_proceeds') ? [['Loan Received', ...projData.map(d => d.cfLoanReceived)]] : []),
    ['Additional Capital', ...projData.map(d => d.cfAdditionalCapital)],
    ['Issuance Of Debt (Bond & Notes)', ...projData.map(d => d.cfIssuanceOfDebt)],
    ['Outflows:', ...projData.map(() => '')],
    ...(isVis(options, 'cf_dividends_paid') ? [['To Stockholders as Dividend', ...projData.map(d => d.cfDividendsToStockholders)]] : []),
    ['Drawings / Dividend', ...projData.map(d => d.cfDrawingsOutflow)],
    ...(isVis(options, 'cf_borrow_repay') ? [['To Redeem Long Term Debt (Treasury Stock)', ...projData.map(d => d.cfRedeemLongTermDebt)]] : []),
    ['Issuance Of Bonds Payable', ...projData.map(d => d.cfIssuanceBondsPayable)],
    ['Net Cash Provided by Financial Activities', ...projData.map(d => d.cfNetCashFinancing)],
    ['Net Cash Increase', ...projData.map(d => d.cfNetCashIncrease)],
    ['Opening Cash & Bank Balance', ...projData.map(d => d.cfOpeningCash)],
    ['Closing Cash & Bank Balance', ...projData.map(d => d.cfClosingCash)],
  ];

  const quantNotesRows: (string | number)[][] = [
    ['3.00', 'Inventories', ...projData.map(d => d.inventories)],
    ['4.00', trRecLbl, ...projData.map(() => '')],
    ['', 'As Per Last Year', ...projData.map(d => d.note4ReceivablesOpening)],
    ['', 'Add: Credit Sales During the Year', ...projData.map(d => d.note4ReceivablesAdded)],
    ['', 'Subtotal', ...projData.map(d => d.note4ReceivablesOpening + d.note4ReceivablesAdded)],
    ['', 'Less: Received During the Year', ...projData.map(d => d.note4ReceivablesReceived)],
    ['', `Closing ${trRecLbl}`, ...projData.map(d => d.tradeAndOtherReceivables)],
    ['5.00', 'Advance, Deposit & Prepayments', ...projData.map(() => '')],
    ['', 'As Per Last Year', ...projData.map(d => d.note5AdvanceOpening)],
    ['', 'Add: During the Year', ...projData.map(d => d.note5AdvanceAdded)],
    ['', 'Subtotal', ...projData.map(d => d.note5AdvanceOpening + d.note5AdvanceAdded)],
    ['', 'Less: Adjustment', ...projData.map(d => d.note5AdvanceAdjusted)],
    ['', 'Closing Advance, Deposit & Prepayments', ...projData.map(d => d.advanceDepositPrepayments)],
    ['6.00', cashLbl, ...projData.map(() => '')],
    ['', 'Cash In Hand & Bank', ...projData.map(d => d.cashAndCashEquivalents)],
    ['7.00', 'Long Term Loan', ...projData.map(d => d.longTermLoan)],
    ['8.00', 'Trade & Other Payables', ...projData.map(() => '')],
    ['', 'As Per Last Year', ...projData.map(d => d.note8PayablesOpening)],
    ['', 'Add: During the Year', ...projData.map(d => d.note8PayablesAdded)],
    ['', 'Subtotal', ...projData.map(d => d.note8PayablesOpening + d.note8PayablesAdded)],
    ['', 'Less: Payment During the Year', ...projData.map(d => d.note8PayablesPaid)],
    ['', 'Closing Trade & Other Payables', ...projData.map(d => d.tradeAndOtherPayables)],
    ['9.00', 'Provisions For Tax', ...projData.map(() => '')],
    ['', 'As Per Last Year', ...projData.map(d => d.note9TaxOpening)],
    ['', 'Add: During the Year', ...projData.map(d => d.note9TaxAdded)],
    ['', 'Subtotal', ...projData.map(d => d.note9TaxOpening + d.note9TaxAdded)],
    ['', 'Less: Adjustment', ...projData.map(d => d.note9TaxAdjusted)],
    ['', 'Closing Provisions For Tax', ...projData.map(d => d.provisionsForTax)],
    ['10.00', 'Short Term Loan', ...projData.map(() => '')],
    ['', 'Bank Loan', ...projData.map(d => d.shortTermLoan)],
    ['11.00', lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold'), ...projData.map(() => '')],
    ['', 'Opening Stock', ...projData.map(d => d.note11OpeningStock)],
    ['', 'Add: Purchase', ...projData.map(d => d.note11Purchase)],
    ['', 'Subtotal', ...projData.map(d => d.note11OpeningStock + d.note11Purchase)],
    ['', 'Less: Closing Stock', ...projData.map(d => d.note11ClosingStock)],
    ['', 'Material Consumed', ...projData.map(d => d.note11OpeningStock + d.note11Purchase - d.note11ClosingStock)],
    ['', 'Add: Direct Expenses', ...projData.map(() => '')],
    ['', 'Labour & Wages', ...projData.map(d => d.note11LabourWages)],
    ['', 'Electricity Bill', ...projData.map(d => d.note11ElectricityBill)],
    ['', 'Carriage Inward', ...projData.map(d => d.note11CarriageInward)],
    ['', 'Other Direct Expenses', ...projData.map(d => d.note11OtherDirectExpenses)],
    ['', 'Total Direct Expenses', ...projData.map(d => d.note11DirectExpensesTotal)],
    ['', 'Cost of Goods Sold', ...projData.map(d => d.costOfGoodsSold)],
    ['12.00', opexLbl, ...projData.map(d => d.adminAndSellingExpenses)],
    ['13.00', 'Financial Expenses', ...projData.map(() => '')],
    ['', 'Bank Charges & Interest', ...projData.map(d => d.financialExpenses)],
  ];

  return { bsRows, plRows, cfRows, quantNotesRows, salesLbl, trRecLbl, opexLbl };
}

export function exportProjectedFinancialsExcel(
  company: CompanyProfile,
  projData: ProjectedFormatYearData[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions
) {
  if (projData.length === 0) return;
  const { filename: resolvedFilename, options } = resolveFilenameAndOptions(
    buildProjectedFilename(company, projData, 'xlsx'),
    filenameOrOptions,
    explicitOptions
  );
  const wb = XLSX.utils.book_new();
  const maxYear = projData[0].year; // projData sorted descending
  const amtLabel = getProjectedAmountCaption(company);
  const yearStartColIdx = 2; // Col A: Particulars, Col B: Notes, Col C+: years
  const widths = [42, 14, ...projData.map(() => 18)];

  const salesLbl = lbl(options, 'income', 'revenue', 'Sales');
  const cogsLbl = `Less: ${lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold')}`;
  const opexLbl = lbl(options, 'income', 'operatingExpenses', 'Administrative & Selling Expenses');
  const trRecLbl = lbl(options, 'bs', 'tradeReceivables', 'Trade & Other Receivables');
  const cashLbl = lbl(options, 'bs', 'cashAndEquivalents', 'Cash & Cash Equivalents');

  // 1. Statement of Financial Position (dynamic Excel row tracking)
  const bsRows: ExcelCellValue[][] = [
    [company.companyName],
    ['Statement of Financial Position'],
    [`As at 31st December, ${maxYear}(Projected)`],
    [],
    ['Particulars', 'Notes', ...projData.map(() => amtLabel)],
    ['', '', ...projData.map(d => d.year)],
    [],
    ['Assets', '', ...projData.map(() => '')],
    ['Non Current Assets', '', ...projData.map(() => '')],
  ];
  const nextBsRow = () => bsRows.length + 1;
  let ppeRow: number | null = null;
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    ppeRow = nextBsRow();
    bsRows.push(['Property Plant & Equiptments', 'Schedule-A', ...projData.map(d => nCell(d.propertyPlantEquipments))]);
  }
  const totalNcaRow = nextBsRow();
  bsRows.push([
    'Total Non Current Assets',
    '',
    ...projData.map((d, i) =>
      fCell(excelSumRows(colLetter(yearStartColIdx + i), ppeRow !== null ? [ppeRow] : []), d.propertyPlantEquipments)
    ),
  ]);
  bsRows.push(['Current Assets', '', ...projData.map(() => '')]);
  const caRows: number[] = [];
  if (isVis(options, 'bs_closingInventory')) {
    caRows.push(nextBsRow());
    bsRows.push(['Inventories', '3.00', ...projData.map(d => nCell(d.inventories))]);
  }
  if (isVis(options, 'bs_tradeReceivables')) {
    caRows.push(nextBsRow());
    bsRows.push([trRecLbl, '4.00', ...projData.map(d => nCell(d.tradeAndOtherReceivables))]);
  }
  if (isVis(options, 'bs_otherReceivables')) {
    caRows.push(nextBsRow());
    bsRows.push(['Advance, Deposit & Prepayments', '5.00', ...projData.map(d => nCell(d.advanceDepositPrepayments))]);
  }
  if (isVis(options, 'bs_cashAndEquivalents')) {
    caRows.push(nextBsRow());
    bsRows.push([cashLbl, '6.00', ...projData.map(d => nCell(d.cashAndCashEquivalents))]);
  }
  const totalCaRow = nextBsRow();
  bsRows.push([
    'Total Current Assets',
    '',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(yearStartColIdx + i), caRows), d.totalCurrentAssets)),
  ]);
  bsRows.push([
    'Total Assets',
    '',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(yearStartColIdx + i), [totalNcaRow, totalCaRow]), d.totalAssets)),
  ]);
  bsRows.push(['Equity & Liabilities', '', ...projData.map(() => '')]);
  const capRow = nextBsRow();
  bsRows.push(['Capital', '', ...projData.map(d => nCell(d.capital))]);
  const addCapRow = nextBsRow();
  bsRows.push(['Additional Capital', '', ...projData.map(d => nCell(d.additionalCapital))]);
  const profRow = nextBsRow();
  bsRows.push(['Profit during the year', '', ...projData.map(d => nCell(d.profitDuringYear))]);
  const eqSubRow = nextBsRow();
  bsRows.push([
    'Subtotal',
    '',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(yearStartColIdx + i), [capRow, addCapRow, profRow]), d.equitySubtotal)),
  ]);
  let drawRow: number | null = null;
  if (isVis(options, 'eq_dividendsPaid')) {
    drawRow = nextBsRow();
    bsRows.push(['Drawings', '', ...projData.map(d => nCell(d.drawings))]);
  }
  const closingCapRow = nextBsRow();
  bsRows.push([
    'Closing Capital',
    '',
    ...projData.map((d, i) =>
      fCell(excelDiffExpr(colLetter(yearStartColIdx + i), [eqSubRow], drawRow !== null ? [drawRow] : []), d.closingCapital)
    ),
  ]);
  bsRows.push(['Non-Current Liabilities', '', ...projData.map(() => '')]);
  let ltLoanRow: number | null = null;
  if (isVis(options, 'bs_longTermBorrowings')) {
    ltLoanRow = nextBsRow();
    bsRows.push(['Long Term Loan', '7.00', ...projData.map(d => nCell(d.longTermLoan))]);
  }
  bsRows.push(['Current Liabilities', '', ...projData.map(() => '')]);
  const clRows: number[] = [];
  if (isVis(options, 'bs_tradePayables')) {
    clRows.push(nextBsRow());
    bsRows.push(['Trade & Other Payables', '8.00', ...projData.map(d => nCell(d.tradeAndOtherPayables))]);
  }
  if (isVis(options, 'bs_currentTaxPayable')) {
    clRows.push(nextBsRow());
    bsRows.push(['Provisions For Tax', '9.00', ...projData.map(d => nCell(d.provisionsForTax))]);
  }
  if (isVis(options, 'bs_shortTermBorrowings')) {
    clRows.push(nextBsRow());
    bsRows.push(['Short Term Loan', '10.00', ...projData.map(d => nCell(d.shortTermLoan))]);
  }
  if (isVis(options, 'bs_otherPayables')) {
    clRows.push(nextBsRow());
    bsRows.push(['Liabilities for exp.', '', ...projData.map(d => nCell(d.liabilitiesForExp))]);
  }
  const totalClRow = nextBsRow();
  bsRows.push([
    'Total Current Liabilities',
    '',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(yearStartColIdx + i), clRows), d.totalCurrentLiabilities)),
  ]);
  const totalEqLiabInputs = ltLoanRow !== null ? [closingCapRow, ltLoanRow, totalClRow] : [closingCapRow, totalClRow];
  bsRows.push([
    'Total Equity & Liabilities',
    '',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(yearStartColIdx + i), totalEqLiabInputs), d.totalEquityAndLiabilities)),
  ]);
  bsRows.push([]);
  bsRows.push(['Proprietor', '', 'Signed in terms of our separate report of even date.']);
  const wsBS = XLSX.utils.aoa_to_sheet(bsRows);
  applyColumnWidths(wsBS, widths);
  XLSX.utils.book_append_sheet(wb, wsBS, '1. Financial Position');

  // 2. Statement of Profit or Loss & Other Comprehensive Income (dynamic Excel row tracking)
  const plRows: ExcelCellValue[][] = [
    [company.companyName],
    ['Statement of Profit or Loss & Other Comprehensive Income'],
    [`For the year ended 31st December, ${maxYear}(Projected)`],
    [],
    ['Particulars', 'Notes', ...projData.map(() => amtLabel)],
    ['', '', ...projData.map(d => d.year)],
    [],
  ];
  const nextPlRow = () => plRows.length + 1;
  let plSalesRow: number | null = null;
  if (isVis(options, 'inc_revenue')) {
    plSalesRow = nextPlRow();
    plRows.push([salesLbl, '', ...projData.map(d => nCell(d.sales))]);
  }
  let plCogsRow: number | null = null;
  if (isVis(options, 'inc_costOfSales')) {
    plCogsRow = nextPlRow();
    plRows.push([cogsLbl, '11.00', ...projData.map(d => nCell(d.costOfGoodsSold))]);
  }
  const plGpRow = nextPlRow();
  plRows.push([
    'Gross Profit',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelDiffExpr(
          colLetter(yearStartColIdx + i),
          plSalesRow !== null ? [plSalesRow] : [],
          plCogsRow !== null ? [plCogsRow] : []
        ),
        d.grossProfit
      )
    ),
  ]);
  let plOtherIncRow: number | null = null;
  if (isVis(options, 'inc_interestIncome')) {
    plOtherIncRow = nextPlRow();
    plRows.push(['Add; Others Income', '', ...projData.map(d => nCell(d.othersIncome))]);
  }
  const plSubRow = nextPlRow();
  plRows.push([
    'Subtotal',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelSumRows(colLetter(yearStartColIdx + i), plOtherIncRow !== null ? [plGpRow, plOtherIncRow] : [plGpRow]),
        d.grossProfitWithOtherIncome
      )
    ),
  ]);
  plRows.push(['Less: Operating Expenses:', '', ...projData.map(() => '')]);
  let plOpexRow: number | null = null;
  if (isVis(options, 'inc_operatingExpenses')) {
    plOpexRow = nextPlRow();
    plRows.push([opexLbl, '12.00', ...projData.map(d => nCell(d.adminAndSellingExpenses))]);
  }
  const plOpProfitRow = nextPlRow();
  plRows.push([
    'Total Operating Profit/(Loss)',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelDiffExpr(colLetter(yearStartColIdx + i), [plSubRow], plOpexRow !== null ? [plOpexRow] : []),
        d.totalOperatingProfit
      )
    ),
  ]);
  let plFinRow: number | null = null;
  if (isVis(options, 'inc_financeCost')) {
    plFinRow = nextPlRow();
    plRows.push(['Less: Financial Expenses', '13.00', ...projData.map(d => nCell(d.financialExpenses))]);
  }
  const plPbtRow = nextPlRow();
  plRows.push([
    'Net Profit/(Loss) Before Income Tax',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelDiffExpr(colLetter(yearStartColIdx + i), [plOpProfitRow], plFinRow !== null ? [plFinRow] : []),
        d.netProfitBeforeTax
      )
    ),
  ]);
  let plTaxRow: number | null = null;
  if (isVis(options, 'inc_taxExpense')) {
    plTaxRow = nextPlRow();
    plRows.push(['Less: Provision For Income Tax', '', ...projData.map(d => nCell(d.provisionForIncomeTax))]);
  }
  plRows.push([
    'Net Profit/(Loss) After Income Tax',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelDiffExpr(colLetter(yearStartColIdx + i), [plPbtRow], plTaxRow !== null ? [plTaxRow] : []),
        d.netProfitAfterTax
      )
    ),
  ]);
  plRows.push([]);
  plRows.push(['Proprietor', '', 'Signed in terms of our separate report of even date.']);
  const wsPL = XLSX.utils.aoa_to_sheet(plRows);
  applyColumnWidths(wsPL, widths);
  XLSX.utils.book_append_sheet(wb, wsPL, '2. Profit or Loss');

  // 3. Statement Of Cash Flows (dynamic Excel row tracking)
  const cfRows: ExcelCellValue[][] = [
    [company.companyName],
    [company.registeredAddress],
    ['Statement Of Cash Flows'],
    [`For the year ended 31st December, ${maxYear}(Projected)`],
    [],
    ['1. CASH FLOW FROM OPERATING ACTIVITIES:', ...projData.map(d => d.year)],
  ];
  const nextCfRow = () => cfRows.length + 1;
  const cfOpRows: number[] = [];
  if (isVis(options, 'cf_pbt')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['Net Income', ...projData.map(d => nCell(d.cfNetIncome))]);
  }
  if (isVis(options, 'cf_wc_inventory')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['(Increase)/Decrease In Inventory', ...projData.map(d => nCell(d.cfChangeInventory))]);
  }
  if (isVis(options, 'cf_wc_receivables')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['(Increase)/Decrease In Trade & Other Receivables', ...projData.map(d => nCell(d.cfChangeReceivables))]);
  }
  if (isVis(options, 'cf_wc_other_rec')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['(Increase)/Decrease in Advance, Deposit & Prepayments', ...projData.map(d => nCell(d.cfChangeAdvanceDeposit))]);
  }
  if (isVis(options, 'cf_wc_other_ca')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['Increase/(Decrease) in Current Assets', ...projData.map(d => nCell(d.cfChangeCurrentAssets))]);
  }
  cfOpRows.push(nextCfRow());
  cfRows.push(['Increase/(Decrease) in Current Liabilities', ...projData.map(d => nCell(d.cfChangeCurrentLiabilities))]);
  if (isVis(options, 'cf_wc_payables')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['Increase/(Decrease) in Trade & Other Payables', ...projData.map(d => nCell(d.cfChangeTradePayables))]);
  }
  if (isVis(options, 'cf_tax_paid')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['Increase/(Decrease) In Provisions For Tax', ...projData.map(d => nCell(d.cfChangeProvisionsTax))]);
  }
  if (isVis(options, 'cf_wc_other_pay')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['Increase/(Decrease) Liabilities For Expense', ...projData.map(d => nCell(d.cfChangeLiabilitiesExp))]);
  }
  if (isVis(options, 'cf_dep')) {
    cfOpRows.push(nextCfRow());
    cfRows.push(['Depreciation', ...projData.map(d => nCell(d.cfDepreciation))]);
  }
  const cfNetOpRow = nextCfRow();
  cfRows.push([
    'Net Cash Used by Operating Activities',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfOpRows), d.cfNetCashOperating)),
  ]);

  cfRows.push(['2. CASH FLOW FROM INVESTING ACTIVITIES:', ...projData.map(() => '')]);
  cfRows.push(['Inflows:', ...projData.map(() => '')]);
  const cfInvRows: number[] = [];
  if (isVis(options, 'cf_sale_ppe')) {
    cfInvRows.push(nextCfRow());
    cfRows.push(['Sale Of Property, Plant & Equipment', ...projData.map(d => nCell(d.cfSalePPE))]);
  }
  cfInvRows.push(nextCfRow());
  cfRows.push(['Capital', ...projData.map(d => nCell(d.cfInvestingCapitalInflow))]);
  cfRows.push(['Outflows:', ...projData.map(() => '')]);
  cfInvRows.push(nextCfRow());
  cfRows.push(['Preliminary Expenses', ...projData.map(d => nCell(d.cfPreliminaryExpenses))]);
  if (isVis(options, 'cf_pur_ppe')) {
    cfInvRows.push(nextCfRow());
    cfRows.push(['Purchase Property, Plant & Equipment', ...projData.map(d => nCell(d.cfPurchasePPE))]);
  }
  cfInvRows.push(nextCfRow());
  cfRows.push(['Purchase Investment In Debt Or Securities', ...projData.map(d => nCell(d.cfPurchaseInvestment))]);
  const cfNetInvRow = nextCfRow();
  cfRows.push([
    'Net Cash Used by Investing Activities',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfInvRows), d.cfNetCashInvesting)),
  ]);

  cfRows.push(['3. CASH FLOW FROM FINANCIAL ACTIVITIES:', ...projData.map(() => '')]);
  cfRows.push(['Inflows:', ...projData.map(() => '')]);
  const cfFinRows: number[] = [];
  if (isVis(options, 'cf_borrow_proceeds')) {
    cfFinRows.push(nextCfRow());
    cfRows.push(['Loan Received', ...projData.map(d => nCell(d.cfLoanReceived))]);
  }
  cfFinRows.push(nextCfRow());
  cfRows.push(['Additional Capital', ...projData.map(d => nCell(d.cfAdditionalCapital))]);
  cfFinRows.push(nextCfRow());
  cfRows.push(['Issuance Of Debt (Bond & Notes)', ...projData.map(d => nCell(d.cfIssuanceOfDebt))]);
  cfRows.push(['Outflows:', ...projData.map(() => '')]);
  if (isVis(options, 'cf_dividends_paid')) {
    cfFinRows.push(nextCfRow());
    cfRows.push(['To Stockholders as Dividend', ...projData.map(d => nCell(d.cfDividendsToStockholders))]);
  }
  cfFinRows.push(nextCfRow());
  cfRows.push(['Drawings / Dividend', ...projData.map(d => nCell(d.cfDrawingsOutflow))]);
  if (isVis(options, 'cf_borrow_repay')) {
    cfFinRows.push(nextCfRow());
    cfRows.push(['To Redeem Long Term Debt (Treasury Stock)', ...projData.map(d => nCell(d.cfRedeemLongTermDebt))]);
  }
  cfFinRows.push(nextCfRow());
  cfRows.push(['Issuance Of Bonds Payable', ...projData.map(d => nCell(d.cfIssuanceBondsPayable))]);
  const cfNetFinRow = nextCfRow();
  cfRows.push([
    'Net Cash Provided by Financial Activities',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfFinRows), d.cfNetCashFinancing)),
  ]);

  const cfNetIncRow = nextCfRow();
  cfRows.push([
    'Net Cash Increase',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), [cfNetOpRow, cfNetInvRow, cfNetFinRow]), d.cfNetCashIncrease)),
  ]);
  const cfOpenCashRow = nextCfRow();
  cfRows.push(['Opening Cash & Bank Balance', ...projData.map(d => nCell(d.cfOpeningCash))]);
  cfRows.push([
    'Closing Cash & Bank Balance',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), [cfNetIncRow, cfOpenCashRow]), d.cfClosingCash)),
  ]);
  const wsCF = XLSX.utils.aoa_to_sheet(cfRows);
  applyColumnWidths(wsCF, [50, ...projData.map(() => 18)]);
  XLSX.utils.book_append_sheet(wb, wsCF, '3. Cash Flows');

  // 4. Notes to the Financial Statements (Notes 1, 2.1-2.6, 3.00-13.00)
  const { quantNotesRows } = buildProjectedTextTables(projData, options);
  const notesRows: ExcelCellValue[][] = [
    [company.companyName],
    [company.registeredAddress],
    ['Notes to the Financial Statements'],
    [`For the year ended 31st December, ${maxYear}(Projected)`],
    [],
    ['Note No.', 'Particulars'],
    ['1', 'Reporting Entity:'],
    ['', `${company.companyName} is a ${company.entityType.replace(/_/g, ' ')} entity and its office is situated at ${company.registeredAddress}.`],
    ['2', 'SIGNIFICANT ACCOUNTING POLICIES:'],
    ['2.1', `Basis of Accounting: Prepared on a going concern basis using accrual method of accounting in accordance with ${company.accountingFramework} and applicable statutory provisions.`],
    ['2.2', 'Accounting Policies: Selected and applied consistently for significant transactions and events to all periods presented.'],
    ['2.3', 'Property, Plant and Equipment: Stated at cost less accumulated depreciation. Depreciation is computed over estimated useful lives consistently.'],
    ['2.4', `Reporting Period: Financial Statements cover the annual reporting period ending ${company.financialYearEnd}.`],
    ['2.5', `Reporting currency and level of precision: Figures represent ${company.reportingCurrency} (${company.reportingUnit}) rounded off cleanly.`],
    ['2.6', 'Comparative Information: Disclosed for all selected years and narratives relevant for understanding the financial statements.'],
    [],
    ['Note No.', 'Particulars', ...projData.map(() => amtLabel)],
    ['', '', ...projData.map(d => d.year)],
    ...quantNotesRows.map(row => row.map((cell, idx) => (idx >= 2 && typeof cell === 'number' ? nCell(cell) : cell))),
  ];
  const wsNotes = XLSX.utils.aoa_to_sheet(notesRows);
  applyColumnWidths(wsNotes, [12, 42, ...projData.map(() => 18)]);
  XLSX.utils.book_append_sheet(wb, wsNotes, '4. Notes to Financials');

  // 5. Schedule of Property Plant & Equipments (Schedule-A)
  const projAsc = [...projData].sort((a, b) => a.year - b.year);
  const schedRows: ExcelCellValue[][] = [
    [company.companyName],
    [company.registeredAddress],
    ['Schedule of Property Plant & Equipments'],
    [`As at 31st December, ${maxYear}(Projected)`],
    [],
  ];
  projAsc.forEach(d => {
    const hdrIdx = schedRows.length + 1;
    schedRows.push([
      'SL No.',
      'Particulars',
      'Cost: Opening Balance',
      'Addition During the Year',
      'Disposals During the Year',
      `Balance as on 31.12.${d.year}`,
      'Rate of Dep. (%)',
      'Depreciation Charged During the Year',
      `W.D.V as on 31.12.${d.year}`,
    ]);
    const itemRow = hdrIdx + 1;
    schedRows.push([
      1,
      'Property Plant & Equipment',
      nCell(d.schedOpeningBalance),
      nCell(d.schedAddition),
      nCell(d.schedDisposals),
      fCell(`C${itemRow}+D${itemRow}-E${itemRow}`, d.schedBalanceClosing),
      nCell(d.schedRatePct),
      nCell(d.schedDepreciation),
      nCell(d.schedWDV),
    ]);
    schedRows.push([
      '',
      'Total',
      fCell(`C${itemRow}`, d.schedOpeningBalance),
      fCell(`D${itemRow}`, d.schedAddition),
      fCell(`E${itemRow}`, d.schedDisposals),
      fCell(`F${itemRow}`, d.schedBalanceClosing),
      '',
      fCell(`H${itemRow}`, d.schedDepreciation),
      fCell(`I${itemRow}`, d.schedWDV),
    ]);
    schedRows.push([]);
  });
  const wsSched = XLSX.utils.aoa_to_sheet(schedRows);
  applyColumnWidths(wsSched, [8, 28, 18, 20, 20, 20, 14, 26, 20]);
  XLSX.utils.book_append_sheet(wb, wsSched, '5. Schedule-A PPE');

  XLSX.writeFile(wb, resolvedFilename);
}

export async function exportProjectedFinancialsWord(
  company: CompanyProfile,
  projData: ProjectedFormatYearData[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions
) {
  if (projData.length === 0) return;
  const { filename: resolvedFilename, options } = resolveFilenameAndOptions(
    buildProjectedFilename(company, projData, 'docx'),
    filenameOrOptions,
    explicitOptions
  );
  const maxYear = projData[0].year;
  const amtLabel = getProjectedAmountCaption(company);
  const bsHeaders = ['Particulars', 'Notes', ...projData.map(d => `${amtLabel}\n${d.year}`)];
  const cfHeaders = ['1. CASH FLOW FROM OPERATING ACTIVITIES:', ...projData.map(d => String(d.year))];
  const projAsc = [...projData].sort((a, b) => a.year - b.year);

  const { bsRows, plRows, cfRows, quantNotesRows } = buildProjectedTextTables(projData, options);

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: company.companyName, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'Statement of Financial Position', bold: true, size: 22 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `As at 31st December, ${maxYear}(Projected)`, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ text: '' }),
          createDocxTable(bsHeaders, bsRows, 2),
          new Paragraph({ text: '' }),
          new Paragraph({ children: [new TextRun({ text: '.............................\nProprietor', bold: true, size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: 'Signed in terms of our separate report of even date.', size: 16 })], alignment: AlignmentType.CENTER }),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'Statement of Profit or Loss & Other Comprehensive Income', bold: true, size: 22 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `For the year ended 31st December, ${maxYear}(Projected)`, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ text: '' }),
          createDocxTable(bsHeaders, plRows, 2),
          new Paragraph({ text: '' }),
          new Paragraph({ children: [new TextRun({ text: '.............................\nProprietor', bold: true, size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: 'Signed in terms of our separate report of even date.', size: 16 })], alignment: AlignmentType.CENTER }),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: company.registeredAddress, size: 16 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'Statement Of Cash Flows', bold: true, size: 22 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `For the year ended 31st December, ${maxYear}(Projected)`, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ text: '' }),
          createDocxTable(cfHeaders, cfRows, 1),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: company.registeredAddress, size: 16 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'Notes to the Financial Statements', bold: true, size: 22 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `For the year ended 31st December, ${maxYear}(Projected)`, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ text: '' }),
          new Paragraph({ children: [new TextRun({ text: '1  Reporting Entity:', bold: true, size: 18 })] }),
          new Paragraph({ children: [new TextRun({ text: `${company.companyName} is a ${company.entityType.replace(/_/g, ' ')} entity and its office is situated at ${company.registeredAddress}.`, size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: '2  SIGNIFICANT ACCOUNTING POLICIES:', bold: true, size: 18 })] }),
          new Paragraph({ children: [new TextRun({ text: `2.1 Basis of Accounting: Prepared on a going concern basis using accrual method of accounting in accordance with ${company.accountingFramework} and applicable statutory provisions.`, size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: '2.2 Accounting Policies: Selected and applied consistently for significant transactions and events to all periods presented in these financial statements.', size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: '2.3 Property, Plant and Equipment: Stated at cost less accumulated depreciation over estimated useful lives consistently.', size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: `2.4 Reporting Period: Financial statements cover the annual reporting period ending ${company.financialYearEnd}.`, size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: `2.5 Reporting currency and level of precision: Figures represent ${company.reportingCurrency} (${company.reportingUnit}) rounded cleanly.`, size: 16 })] }),
          new Paragraph({ children: [new TextRun({ text: '2.6 Comparative Information: Disclosed for all selected years and narratives relevant for understanding the financial statements.', size: 16 })] }),
          new Paragraph({ text: '' }),
          createDocxTable(['Note No.', 'Particulars', ...projData.map(d => `${amtLabel}\n${d.year}`)], quantNotesRows, 2),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: company.companyName, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: company.registeredAddress, size: 16 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: 'Schedule of Property Plant & Equipments', bold: true, size: 22 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `As at 31st December, ${maxYear}(Projected)`, size: 18 })], alignment: AlignmentType.CENTER }),
          new Paragraph({ text: '' }),
          ...projAsc.flatMap(d => [
            createDocxTable(
              ['SL No.', 'Particulars', 'Opening Balance', 'Addition During the Year', 'Disposals During the Year', `Balance as on 31.12.${d.year}`, 'Rate (%)', 'Depreciation Charged During the Year', `W.D.V as on 31.12.${d.year}`],
              [
                [1, 'Property Plant & Equipment', d.schedOpeningBalance, d.schedAddition, d.schedDisposals, d.schedBalanceClosing, `${d.schedRatePct}%`, d.schedDepreciation, d.schedWDV],
                ['', 'Total', d.schedOpeningBalance, d.schedAddition, d.schedDisposals, d.schedBalanceClosing, '', d.schedDepreciation, d.schedWDV],
              ],
              2
            ),
            new Paragraph({ text: '' }),
          ]),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, resolvedFilename);
}

export function exportProjectedFinancialsPdf(
  company: CompanyProfile,
  projData: ProjectedFormatYearData[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions
) {
  if (projData.length === 0) return;
  const { filename: resolvedFilename, options } = resolveFilenameAndOptions(
    buildProjectedFilename(company, projData, 'pdf'),
    filenameOrOptions,
    explicitOptions
  );
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const maxYear = projData[0].year;
  const amtLabel = getProjectedAmountCaption(company);
  const shortAmtLabel = projData.length >= 10
    ? (company.reportingCurrency === '৳' ? 'In Taka' : `In ${company.reportingCurrency}`)
    : amtLabel;
  const projAsc = [...projData].sort((a, b) => a.year - b.year);

  const nCols = projData.length;
  const usableW = 277; // 297mm - 20mm margins
  const partW = nCols > 15 ? 46 : nCols >= 12 ? 50 : nCols >= 8 ? 56 : 66;
  const notesW = nCols > 15 ? 13 : nCols >= 12 ? 15 : 18;
  const yearW = Math.max(11, (usableW - partW - notesW) / nCols);
  const bodyFontSize = nCols > 15 ? 5.1 : nCols >= 12 ? 5.6 : nCols >= 8 ? 6.6 : 7.6;
  const headFontSize = nCols > 15 ? 5.0 : nCols >= 12 ? 5.5 : nCols >= 8 ? 6.5 : 7.5;
  const cellPad = nCols > 15 ? 0.6 : nCols >= 12 ? 0.8 : 1.1;

  const cfPartW = nCols > 15 ? 60 : nCols >= 12 ? 66 : 82;
  const cfYearW = Math.max(11, (usableW - cfPartW) / nCols);

  const drawProjHeader = (title: string, subtitle: string, includeAddress = false) => {
    doc.setTextColor(15, 23, 42);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11.5);
    doc.text(company.companyName, 148.5, 11.5, { align: 'center' });
    let y = 16;
    if (includeAddress) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text(company.registeredAddress, 148.5, y, { align: 'center' });
      y += 4.5;
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10.5);
    doc.text(title, 148.5, y, { align: 'center' });
    y += 4.2;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.text(subtitle, 148.5, y, { align: 'center' });
    return y + 3.5;
  };

  const drawProprietorFooter = (yPos: number) => {
    let safeY = yPos + 10;
    if (safeY > 192) {
      doc.addPage();
      safeY = 24;
    } else {
      safeY = Math.min(190, Math.max(safeY, 170));
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(15, 23, 42);
    doc.text('...........................', 16, safeY);
    doc.text('Proprietor', 22, safeY + 4);
    doc.setFont('helvetica', 'normal');
    doc.text('Signed in terms of our separate report of even date.', 148.5, safeY + 6, { align: 'center' });
  };

  const twoRowHead = [
    [
      { content: 'Particulars', rowSpan: 2, styles: { valign: 'middle' as const, halign: 'center' as const } },
      { content: 'Notes', rowSpan: 2, styles: { valign: 'middle' as const, halign: 'center' as const } },
      ...projData.map(() => ({ content: shortAmtLabel, styles: { halign: 'center' as const } })),
    ],
    projData.map(d => ({ content: String(d.year), styles: { halign: 'center' as const } })),
  ];

  const {
    bsRows: bsBody,
    plRows: plBody,
    cfRows: cfBody,
    quantNotesRows: notesBody,
    salesLbl,
    trRecLbl,
    opexLbl,
  } = buildProjectedTextTables(projData, options);

  // Page 1: Statement of Financial Position
  const startY1 = drawProjHeader('Statement of Financial Position', `As at 31st December, ${maxYear}(Projected)`);
  const bsBoldRows = new Set([
    'Assets', 'Non Current Assets', 'Total Non Current Assets', 'Current Assets', 'Total Current Assets',
    'Total Assets', 'Equity & Liabilities', 'Subtotal', 'Closing Capital', 'Non-Current Liabilities',
    'Current Liabilities', 'Total Current Liabilities', 'Total Equity & Liabilities',
  ]);
  const bsRuleRows = new Set([
    'Total Non Current Assets', 'Total Current Assets', 'Total Assets',
    'Subtotal', 'Closing Capital', 'Total Current Liabilities', 'Total Equity & Liabilities',
  ]);
  autoTable(doc, {
    head: twoRowHead,
    body: bsBody.map(r => r.map((c, idx) => idx < 2 ? String(c) : (typeof c === 'number' ? fmtExactNum(c) : String(c)))),
    startY: startY1,
    theme: 'plain',
    styles: { fontSize: bodyFontSize, cellPadding: cellPad, textColor: [15, 23, 42], overflow: 'ellipsize' },
    headStyles: { fontSize: headFontSize, fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', lineWidth: 0.2, lineColor: [30, 41, 59] },
    columnStyles: {
      0: { halign: 'left', cellWidth: partW },
      1: { halign: 'center', cellWidth: notesW },
      ...Object.fromEntries(projData.map((_, i) => [i + 2, { halign: 'right', cellWidth: yearW }])),
    },
    didParseCell: (data) => {
      if (data.section === 'body') {
        const rowLabel = Array.isArray(data.row.raw) ? String(data.row.raw[0] ?? '') : '';
        if (bsBoldRows.has(rowLabel)) data.cell.styles.fontStyle = 'bold';
        if (data.column.index === 0 && (rowLabel === 'Total Assets' || rowLabel === 'Total Equity & Liabilities')) {
          data.cell.styles.halign = 'right';
        }
        if (data.column.index >= 2 && bsRuleRows.has(rowLabel)) {
          data.cell.styles.lineWidth = { top: 0.25, bottom: (rowLabel === 'Total Assets' || rowLabel === 'Total Equity & Liabilities') ? 0.55 : 0.25, left: 0, right: 0 };
          data.cell.styles.lineColor = [15, 23, 42];
        }
      }
    },
    margin: { left: 10, right: 10 },
  });
  drawProprietorFooter(((doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? 170));

  // Page 2: Statement of Profit or Loss
  doc.addPage();
  const startY2 = drawProjHeader('Statement of Profit or Loss & Other Comprehensive Income', `For the year ended 31st December, ${maxYear}(Projected)`);
  const plRuleRows = new Set([
    'Gross Profit', 'Subtotal', 'Total Operating Profit/(Loss)',
    'Net Profit/(Loss) Before Income Tax', 'Net Profit/(Loss) After Income Tax',
  ]);
  autoTable(doc, {
    head: twoRowHead,
    body: plBody.map(r => r.map((c, idx) => idx < 2 ? String(c) : (typeof c === 'number' ? fmtExactNum(c) : String(c)))),
    startY: startY2,
    theme: 'plain',
    styles: { fontSize: Math.max(bodyFontSize, 5.8), cellPadding: Math.max(cellPad, 1.0), textColor: [15, 23, 42], overflow: 'ellipsize' },
    headStyles: { fontSize: headFontSize, fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', lineWidth: 0.2, lineColor: [30, 41, 59] },
    columnStyles: {
      0: { halign: 'left', cellWidth: partW },
      1: { halign: 'center', cellWidth: notesW },
      ...Object.fromEntries(projData.map((_, i) => [i + 2, { halign: 'right', cellWidth: yearW }])),
    },
    didParseCell: (data) => {
      if (data.section === 'body') {
        const rowLabel = Array.isArray(data.row.raw) ? String(data.row.raw[0] ?? '') : '';
        if (rowLabel === salesLbl || rowLabel === 'Less: Operating Expenses:' || plRuleRows.has(rowLabel)) {
          data.cell.styles.fontStyle = 'bold';
        }
        if (data.column.index >= 2 && plRuleRows.has(rowLabel)) {
          data.cell.styles.lineWidth = { top: 0.25, bottom: rowLabel === 'Net Profit/(Loss) After Income Tax' ? 0.55 : 0.25, left: 0, right: 0 };
          data.cell.styles.lineColor = [15, 23, 42];
        }
      }
    },
    margin: { left: 10, right: 10 },
  });
  drawProprietorFooter(((doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? 150));

  // Page 3: Statement Of Cash Flows
  doc.addPage();
  const startY3 = drawProjHeader('Statement Of Cash Flows', `For the year ended 31st December, ${maxYear}(Projected)`, true);
  const cfHead = ['1. CASH FLOW FROM OPERATING ACTIVITIES:', ...projData.map(d => String(d.year))];
  const cfTotalRows = new Set([
    'Net Cash Used by Operating Activities',
    'Net Cash Used by Investing Activities',
    'Net Cash Provided by Financial Activities',
    'Net Cash Increase',
    'Opening Cash & Bank Balance',
    'Closing Cash & Bank Balance',
  ]);
  autoTable(doc, {
    head: [cfHead],
    body: cfBody.map(r => r.map((c, idx) => idx === 0 ? String(c) : (typeof c === 'number' ? fmtExactNum(c) : String(c)))),
    startY: startY3,
    theme: 'plain',
    styles: { fontSize: bodyFontSize, cellPadding: Math.max(0.6, cellPad * 0.9), textColor: [15, 23, 42], overflow: 'ellipsize' },
    headStyles: { fontSize: headFontSize, fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', lineWidth: 0.25, lineColor: [30, 41, 59] },
    columnStyles: {
      0: { halign: 'left', cellWidth: cfPartW },
      ...Object.fromEntries(projData.map((_, i) => [i + 1, { halign: 'right', cellWidth: cfYearW }])),
    },
    didParseCell: (data) => {
      if (data.section === 'body') {
        const rowLabel = Array.isArray(data.row.raw) ? String(data.row.raw[0] ?? '') : '';
        if (cfTotalRows.has(rowLabel) || rowLabel.startsWith('2.') || rowLabel.startsWith('3.') || rowLabel === 'Inflows:' || rowLabel === 'Outflows:') {
          data.cell.styles.fontStyle = 'bold';
        }
        if (data.column.index >= 1 && (rowLabel === 'Net Cash Used by Operating Activities' || rowLabel === 'Closing Cash & Bank Balance')) {
          data.cell.styles.lineWidth = { top: 0.25, bottom: rowLabel === 'Closing Cash & Bank Balance' ? 0.55 : 0.25, left: 0, right: 0 };
          data.cell.styles.lineColor = [15, 23, 42];
        }
      }
    },
    margin: { left: 10, right: 10 },
  });

  // Page 4: Narrative Notes to the Financial Statements (Notes 1 & 2.1-2.6)
  doc.addPage();
  const startYNotes1 = drawProjHeader('Notes to the Financial Statements', `For the year ended 31st December, ${maxYear}(Projected)`, true);
  autoTable(doc, {
    head: [['Note No.', 'Particulars']],
    body: [
      ['1', 'Reporting Entity:'],
      ['', `${company.companyName} is a ${company.entityType.replace(/_/g, ' ')} entity and its office is situated at ${company.registeredAddress}.`],
      ['2', 'SIGNIFICANT ACCOUNTING POLICIES:'],
      ['2.1', `Basis of Accounting:\nThe Financial Statements have been prepared on a going concern basis using accrual method of accounting and in accordance with ${company.accountingFramework} and applicable statutory requirements.`],
      ['2.2', 'Accounting Policies:\nThe Accounting policies were selected and applied consistently for significant transactions and events to all periods presented in these financial statements.'],
      ['2.3', 'Property, Plant and Equipment:\nItems of PPE are stated at cost less accumulated depreciation. Depreciation is computed over the estimated useful lives consistently.'],
      ['2.4', `Reporting Period:\nFinancial Statements of the entity cover the annual reporting period ending ${company.financialYearEnd} consistently.`],
      ['2.5', `Reporting currency and level of precision:\nThe figures in the Financial Statements represent ${company.reportingCurrency} (${company.reportingUnit}) currency rounded off cleanly.`],
      ['2.6', 'Comparative Information:\nComparative Information has been disclosed in respect of all selected years for numerical information, narratives, and descriptive schedules.'],
    ],
    startY: startYNotes1 + 2,
    theme: 'plain',
    styles: { fontSize: 8.5, cellPadding: 2, textColor: [15, 23, 42] },
    headStyles: { fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', lineWidth: 0.25, lineColor: [30, 41, 59] },
    columnStyles: {
      0: { halign: 'center', cellWidth: 22, fontStyle: 'bold' },
      1: { halign: 'left' },
    },
    didParseCell: (data) => {
      if (data.section === 'body') {
        const raw = Array.isArray(data.row.raw) ? data.row.raw : [];
        const no = String(raw[0] ?? '');
        const txt = String(raw[1] ?? '');
        if (no === '1' || no === '2' || txt.startsWith(company.companyName)) {
          data.cell.styles.fontStyle = 'bold';
        }
      }
    },
    margin: { left: 14, right: 14 },
  });

  // Page 5+: Full Quantitative Notes 3.00 - 13.00 (including all roll-forward lines)
  doc.addPage();
  const startY4 = drawProjHeader('Notes to the Financial Statements', `For the year ended 31st December, ${maxYear}(Projected)`, true);
  const noteBoldSet = new Set([
    'Subtotal', `Closing ${trRecLbl}`, 'Closing Advance, Deposit & Prepayments',
    'Cash In Hand & Bank', 'Long Term Loan', 'Closing Trade & Other Payables', 'Closing Provisions For Tax',
    'Bank Loan', 'Material Consumed', 'Total Direct Expenses', 'Cost of Goods Sold',
    opexLbl, 'Inventories',
  ]);
  autoTable(doc, {
    head: [
      [
        { content: 'Note No.', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
        { content: 'Particulars', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
        ...projData.map(() => ({ content: shortAmtLabel, styles: { halign: 'center' as const } })),
      ],
      projData.map(d => ({ content: String(d.year), styles: { halign: 'center' as const } })),
    ],
    body: notesBody.map(r => r.map((c, idx) => idx < 2 ? String(c) : (typeof c === 'number' ? fmtExactNum(c) : String(c)))),
    startY: startY4,
    theme: 'plain',
    styles: { fontSize: bodyFontSize, cellPadding: cellPad, textColor: [15, 23, 42], overflow: 'ellipsize' },
    headStyles: { fontSize: headFontSize, fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', lineWidth: 0.2, lineColor: [30, 41, 59] },
    columnStyles: {
      0: { halign: 'center', cellWidth: notesW },
      1: { halign: 'left', cellWidth: partW },
      ...Object.fromEntries(projData.map((_, i) => [i + 2, { halign: 'right', cellWidth: yearW }])),
    },
    didParseCell: (data) => {
      if (data.section === 'body') {
        const raw = Array.isArray(data.row.raw) ? data.row.raw : [];
        const noteNo = String(raw[0] ?? '');
        const rowLabel = String(raw[1] ?? '');
        if (noteNo !== '' || noteBoldSet.has(rowLabel) || rowLabel === 'Add: Direct Expenses') {
          data.cell.styles.fontStyle = 'bold';
        }
        if (data.column.index >= 2 && noteBoldSet.has(rowLabel)) {
          data.cell.styles.lineWidth = { top: 0.2, bottom: 0.35, left: 0, right: 0 };
          data.cell.styles.lineColor = [15, 23, 42];
        }
      }
    },
    margin: { left: 10, right: 10 },
  });

  // Final Page+: Schedule of Property Plant & Equipments (Schedule-A)
  doc.addPage();
  let ySched = drawProjHeader('Schedule of Property Plant & Equipments', `As at 31st December, ${maxYear}(Projected)`, true);
  projAsc.forEach(d => {
    if (ySched > 175) {
      doc.addPage();
      ySched = 16;
    }
    autoTable(doc, {
      head: [
        [
          { content: 'SL No.', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
          { content: 'Particulars', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
          { content: 'Cost', colSpan: 4, styles: { halign: 'center' } },
          { content: 'Rate of Dep. (%)', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
          { content: 'Depreciation Charged During the Year', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
          { content: `W.D.V as on 31.12.${d.year}`, rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
        ],
        ['Opening Balance', 'Addition During the Year', 'Disposals During the Year', `Balance as on 31.12.${d.year}`],
      ],
      body: [
        ['1', 'Property Plant & Equipment', fmtExactNum(d.schedOpeningBalance), fmtExactNum(d.schedAddition), fmtExactNum(d.schedDisposals), fmtExactNum(d.schedBalanceClosing), `${d.schedRatePct}%`, fmtExactNum(d.schedDepreciation), fmtExactNum(d.schedWDV)],
        ['', 'Total', fmtExactNum(d.schedOpeningBalance), fmtExactNum(d.schedAddition), fmtExactNum(d.schedDisposals), fmtExactNum(d.schedBalanceClosing), '', fmtExactNum(d.schedDepreciation), fmtExactNum(d.schedWDV)],
      ],
      startY: ySched,
      theme: 'grid',
      styles: { fontSize: 7.5, cellPadding: 1.5, textColor: [15, 23, 42], lineColor: [51, 65, 85], lineWidth: 0.2 },
      headStyles: { fillColor: [255, 255, 255], textColor: [15, 23, 42], fontStyle: 'bold', halign: 'center', lineColor: [51, 65, 85], lineWidth: 0.25 },
      columnStyles: {
        0: { halign: 'center', cellWidth: 12 },
        1: { halign: 'left', cellWidth: 48 },
        2: { halign: 'right' },
        3: { halign: 'right' },
        4: { halign: 'right' },
        5: { halign: 'right' },
        6: { halign: 'center', cellWidth: 20 },
        7: { halign: 'right' },
        8: { halign: 'right' },
      },
      didParseCell: (data) => {
        if (data.section === 'body' && data.row.index === 1) {
          data.cell.styles.fontStyle = 'bold';
        }
      },
      margin: { left: 10, right: 10 },
    });
    ySched = ((doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? ySched + 24) + 5;
  });

  doc.save(resolvedFilename);
}
