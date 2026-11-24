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
  StatementCellOverrides,
} from '../types/financial';
import {
  FIXED_ASSET_CATEGORY_CONFIG_IDS,
  getCustomItemYearValue,
  resolveEffectiveLabel,
  buildProjectedQuantNoteRows,
  ProjectedQuantNoteRow,
} from './financialEngine';

export interface ExportSyncOptions {
  statementConfig?: StatementConfigItem[];
  labelOverrides?: StatementLabelOverrides;
  cellOverrides?: StatementCellOverrides;
}

export function isNilExportValue(val: unknown): boolean {
  if (val === null || val === undefined) return true;
  if (typeof val === 'number') {
    return !Number.isFinite(val) || Math.abs(val) < 0.005;
  }
  if (typeof val === 'string') {
    const t = val.trim();
    if (t === '' || t === '-' || t === '—' || t === 'N/A') return true;
    const stripped = t.replace(/[(),%\s]/g, '');
    if (stripped === '' || stripped === '-') return true;
    const parsed = Number(stripped);
    return Number.isFinite(parsed) && Math.abs(parsed) < 0.005;
  }
  return false;
}

export function hasAnyNonZeroValue(values: readonly unknown[]): boolean {
  return values.some(v => !isNilExportValue(v));
}

function filterNoteTableForExport(tableData: FinancialNoteItem['tableData']) {
  if (!tableData) return null;
  const nonNilRows = tableData.rows.filter(r => hasAnyNonZeroValue(r.values));
  const hasNonNilTotal = Boolean(tableData.totalRow && hasAnyNonZeroValue(tableData.totalRow.values));
  if (nonNilRows.length === 0 && !hasNonNilTotal) return null;
  return {
    headers: tableData.headers,
    rows: nonNilRows,
    totalRow: hasNonNilTotal || nonNilRows.length > 0 ? tableData.totalRow : undefined,
  };
}

function filterProjectedQuantNotesForExport(noteRows: ProjectedQuantNoteRow[]): ProjectedQuantNoteRow[] {
  const result: ProjectedQuantNoteRow[] = [];
  let i = 0;
  while (i < noteRows.length) {
    const row = noteRows[i];
    if (row.noteNo !== '' && row.isSectionHeader) {
      // Collect entire top-level note block until next noteNo !== ''
      let j = i + 1;
      while (j < noteRows.length && noteRows[j].noteNo === '') {
        j++;
      }
      const children = noteRows.slice(i + 1, j);
      const filteredChildren: ProjectedQuantNoteRow[] = [];
      let k = 0;
      while (k < children.length) {
        const ch = children[k];
        if (ch.isSectionHeader) {
          // Sub-header like "Add: Direct Expenses" within Note 11.00
          let m = k + 1;
          while (m < children.length && !children[m].isSectionHeader && !children[m].label.startsWith('Cost of Goods Sold')) {
            if (children[m].label === 'Total Direct Expenses') {
              m++;
              break;
            }
            m++;
          }
          const subGroup = children.slice(k + 1, m);
          const subDetails = subGroup.filter(sg => !sg.isSubtotal && hasAnyNonZeroValue(sg.values));
          if (subDetails.length > 0) {
            filteredChildren.push(ch);
            subGroup.forEach(sg => {
              if (sg.isSubtotal || hasAnyNonZeroValue(sg.values)) filteredChildren.push(sg);
            });
          }
          k = m;
        } else {
          if (hasAnyNonZeroValue(ch.values)) {
            filteredChildren.push(ch);
          }
          k++;
        }
      }
      const hasAnyValueInBlock = filteredChildren.some(c => !c.isSectionHeader && hasAnyNonZeroValue(c.values));
      if (hasAnyValueInBlock) {
        result.push(row, ...filteredChildren);
      }
      i = j;
    } else {
      if (hasAnyNonZeroValue(row.values)) {
        result.push(row);
      }
      i++;
    }
  }
  return result;
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

function hasOv(
  options: ExportSyncOptions | undefined,
  stmt: string,
  field: string,
  year: number
): boolean {
  return options?.cellOverrides?.[`${stmt}_${field}`]?.[year] !== undefined;
}

function ovVal(
  options: ExportSyncOptions | undefined,
  stmt: string,
  field: string,
  year: number,
  fallback: number
): number {
  const val = options?.cellOverrides?.[`${stmt}_${field}`]?.[year];
  return val !== undefined ? val : fallback;
}

function getCustomRows(
  options: ExportSyncOptions | undefined,
  statement: StatementConfigItem['statement'],
  section?: string
): StatementConfigItem[] {
  return (options?.statementConfig || [])
    .filter(
      item =>
        item.isCustom &&
        item.isVisible &&
        item.statement === statement &&
        (!section || item.section === section)
    )
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
}

function getUnscopedCustomIncomeRows(options: ExportSyncOptions | undefined): StatementConfigItem[] {
  return (options?.statementConfig || [])
    .filter(
      item =>
        item.isCustom &&
        item.isVisible &&
        item.statement === 'income' &&
        item.section !== 'Revenue & Gross Margin' &&
        item.section !== 'Operating Expenses' &&
        item.section !== 'Non-Operating & Tax'
    )
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
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
    const revVals = data.map(d => {
      const custSum = customRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
      return ovVal(options, 'income', 'revenue', d.year, d.revenue - custSum);
    });
    if (hasAnyNonZeroValue(revVals)) {
      const revRow = nextRow();
      rows.push([lbl(options, 'income', 'revenue', 'Sales Revenue'), ...revVals.map(v => nCell(v))]);
      gpInputs.push(revRow);
    }
  }
  customRev.forEach(ci => {
    const vals = data.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) {
      const cr = nextRow();
      rows.push([ci.label, ...vals.map(v => nCell(v))]);
      gpInputs.push(cr);
    }
  });

  if (isVis(options, 'inc_costOfSales')) {
    const cosVals = data.map(d => -ovVal(options, 'income', 'costOfSales', d.year, d.costOfSales));
    if (hasAnyNonZeroValue(cosVals)) {
      const cosRow = nextRow();
      rows.push([lbl(options, 'income', 'costOfSales', 'Cost of Sales'), ...cosVals.map(v => nCell(v))]);
      gpInputs.push(cosRow);
    }
  }

  const gpRow = nextRow();
  rows.push([
    lbl(options, 'income', 'grossProfit', 'Gross Profit'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'income', 'grossProfit', d.year, d.grossProfit);
      return hasOv(options, 'income', 'grossProfit', d.year) ? nCell(val) : fCell(excelSumRows(c, gpInputs), val);
    }),
  ]);

  const ebitInputs: number[] = [gpRow];
  if (isVis(options, 'inc_operatingExpenses')) {
    const opexVals = data.map(d => {
      const custSum = customOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
      return -ovVal(options, 'income', 'operatingExpenses', d.year, d.operatingExpenses + custSum);
    });
    if (hasAnyNonZeroValue(opexVals)) {
      const opexRow = nextRow();
      rows.push([lbl(options, 'income', 'operatingExpenses', 'Operating Expenses'), ...opexVals.map(v => nCell(v))]);
      ebitInputs.push(opexRow);
    }
  }
  customOpex.forEach(ci => {
    const vals = data.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) {
      const cr = nextRow();
      rows.push([ci.label, ...vals.map(v => nCell(v))]);
      ebitInputs.push(cr);
    }
  });

  if (isVis(options, 'inc_depreciationExpense')) {
    const depVals = data.map(d => -ovVal(options, 'income', 'depreciationExpense', d.year, d.depreciationExpense));
    if (hasAnyNonZeroValue(depVals)) {
      const depRow = nextRow();
      rows.push([lbl(options, 'income', 'depreciationExpense', 'Depreciation Expense'), ...depVals.map(v => nCell(v))]);
      ebitInputs.push(depRow);
    }
  }

  const ebitRow = nextRow();
  rows.push([
    lbl(options, 'income', 'operatingProfit', 'Operating Profit (EBIT)'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'income', 'operatingProfit', d.year, d.operatingProfit);
      return hasOv(options, 'income', 'operatingProfit', d.year) ? nCell(val) : fCell(excelSumRows(c, ebitInputs), val);
    }),
  ]);

  const pbtInputs: number[] = [ebitRow];
  if (isVis(options, 'inc_interestIncome')) {
    const intVals = data.map(d => {
      const custSum = customNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0);
      return ovVal(options, 'income', 'interestIncome', d.year, d.interestIncome - custSum);
    });
    if (hasAnyNonZeroValue(intVals)) {
      const intRow = nextRow();
      rows.push([lbl(options, 'income', 'interestIncome', 'Interest & Investment Income'), ...intVals.map(v => nCell(v))]);
      pbtInputs.push(intRow);
    }
  }

  if (isVis(options, 'inc_financeCost')) {
    const fcVals = data.map(d => -ovVal(options, 'income', 'financeCost', d.year, d.financeCost));
    if (hasAnyNonZeroValue(fcVals)) {
      const fcRow = nextRow();
      rows.push([lbl(options, 'income', 'financeCost', 'Finance Cost (Mandatory Direct Entry)'), ...fcVals.map(v => nCell(v))]);
      pbtInputs.push(fcRow);
    }
  }

  customNonOp.forEach(ci => {
    const vals = data.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) {
      const cr = nextRow();
      rows.push([ci.label, ...vals.map(v => nCell(v))]);
      pbtInputs.push(cr);
    }
  });

  const pbtRow = nextRow();
  rows.push([
    lbl(options, 'income', 'profitBeforeTax', 'Profit Before Tax (EBT)'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'income', 'profitBeforeTax', d.year, d.profitBeforeTax);
      return hasOv(options, 'income', 'profitBeforeTax', d.year) ? nCell(val) : fCell(excelSumRows(c, pbtInputs), val);
    }),
  ]);

  const npInputs: number[] = [pbtRow];
  if (isVis(options, 'inc_taxExpense')) {
    const taxVals = data.map(d => -ovVal(options, 'income', 'taxExpense', d.year, d.taxExpense));
    if (hasAnyNonZeroValue(taxVals)) {
      const taxRow = nextRow();
      rows.push([lbl(options, 'income', 'taxExpense', 'Income Tax Expense'), ...taxVals.map(v => nCell(v))]);
      npInputs.push(taxRow);
    }
  }

  rows.push([
    lbl(options, 'income', 'netProfit', 'Net Profit After Tax'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'income', 'netProfit', d.year, d.netProfit);
      return hasOv(options, 'income', 'netProfit', d.year) ? nCell(val) : fCell(excelSumRows(c, npInputs), val);
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

  const pushBsRowIfNonZero = (rowTracker: number[], label: string, vals: number[]) => {
    if (!hasAnyNonZeroValue(vals)) return;
    rowTracker.push(nextRow());
    rows.push([label, ...vals.map(v => nCell(v))]);
  };

  const caRows: number[] = [];
  if (isVis(options, 'bs_cashAndEquivalents')) {
    pushBsRowIfNonZero(caRows, lbl(options, 'bs', 'cashAndEquivalents', 'Cash and Cash Equivalents'), data.map(d => ovVal(options, 'bs', 'cashAndEquivalents', d.year, d.cashAndEquivalents)));
  }
  if (isVis(options, 'bs_tradeReceivables')) {
    pushBsRowIfNonZero(caRows, lbl(options, 'bs', 'tradeReceivables', 'Trade Receivables'), data.map(d => ovVal(options, 'bs', 'tradeReceivables', d.year, d.tradeReceivables)));
  }
  if (isVis(options, 'bs_otherReceivables')) {
    pushBsRowIfNonZero(caRows, lbl(options, 'bs', 'otherReceivables', 'Other Receivables & Prepayments'), data.map(d => ovVal(options, 'bs', 'otherReceivables', d.year, d.otherReceivables)));
  }
  if (isVis(options, 'bs_closingInventory')) {
    pushBsRowIfNonZero(caRows, lbl(options, 'bs', 'closingInventory', 'Closing Inventory'), data.map(d => ovVal(options, 'bs', 'closingInventory', d.year, d.closingInventory)));
  }
  if (isVis(options, 'bs_otherCurrentAssets')) {
    pushBsRowIfNonZero(caRows, lbl(options, 'bs', 'otherCurrentAssets', 'Other Current Assets'), data.map(d => ovVal(options, 'bs', 'otherCurrentAssets', d.year, d.otherCurrentAssets)));
  }
  getCustomRows(options, 'balanceSheet', 'Current Assets').forEach(ci => {
    pushBsRowIfNonZero(caRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
  });

  const totalCaRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalCurrentAssets', 'Total Current Assets'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'bs', 'totalCurrentAssets', d.year, d.totalCurrentAssets);
      return hasOv(options, 'bs', 'totalCurrentAssets', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), caRows), val);
    }),
  ]);

  const ncaRows: number[] = [];
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    pushBsRowIfNonZero(ncaRows, lbl(options, 'bs', 'propertyPlantEquipment', 'Property, Plant and Equipment (NBV)'), data.map(d => ovVal(options, 'bs', 'propertyPlantEquipment', d.year, d.propertyPlantEquipment)));
  }
  if (isVis(options, 'bs_intangibleAssets')) {
    pushBsRowIfNonZero(ncaRows, lbl(options, 'bs', 'intangibleAssets', 'Intangible Assets'), data.map(d => ovVal(options, 'bs', 'intangibleAssets', d.year, d.intangibleAssets)));
  }
  if (isVis(options, 'bs_otherNonCurrentAssets')) {
    pushBsRowIfNonZero(ncaRows, lbl(options, 'bs', 'otherNonCurrentAssets', 'Other Non-Current Assets'), data.map(d => ovVal(options, 'bs', 'otherNonCurrentAssets', d.year, d.otherNonCurrentAssets)));
  }
  getCustomRows(options, 'balanceSheet', 'Non-Current Assets').forEach(ci => {
    pushBsRowIfNonZero(ncaRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
  });

  const totalNcaRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalNonCurrentAssets', 'Total Non-Current Assets'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'bs', 'totalNonCurrentAssets', d.year, d.totalNonCurrentAssets);
      return hasOv(options, 'bs', 'totalNonCurrentAssets', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), ncaRows), val);
    }),
  ]);

  const totalAssetsRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalAssets', 'TOTAL ASSETS'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'bs', 'totalAssets', d.year, d.totalAssets);
      return hasOv(options, 'bs', 'totalAssets', d.year) ? nCell(val) : fCell(`${c}${totalCaRow}+${c}${totalNcaRow}`, val);
    }),
  ]);

  rows.push([]);
  rows.push(['LIABILITIES & SHAREHOLDERS EQUITY', ...data.map(() => '')]);

  const clRows: number[] = [];
  if (isVis(options, 'bs_tradePayables')) {
    pushBsRowIfNonZero(clRows, lbl(options, 'bs', 'tradePayables', 'Trade Payables'), data.map(d => ovVal(options, 'bs', 'tradePayables', d.year, d.tradePayables)));
  }
  if (isVis(options, 'bs_otherPayables')) {
    pushBsRowIfNonZero(clRows, lbl(options, 'bs', 'otherPayables', 'Other Current Payables'), data.map(d => ovVal(options, 'bs', 'otherPayables', d.year, d.otherPayables)));
  }
  if (isVis(options, 'bs_shortTermBorrowings')) {
    pushBsRowIfNonZero(clRows, lbl(options, 'bs', 'shortTermBorrowings', 'Short-Term Borrowings'), data.map(d => ovVal(options, 'bs', 'shortTermBorrowings', d.year, d.shortTermBorrowings)));
  }
  if (isVis(options, 'bs_currentTaxPayable')) {
    pushBsRowIfNonZero(clRows, lbl(options, 'bs', 'currentTaxPayable', 'Current Tax Payable'), data.map(d => ovVal(options, 'bs', 'currentTaxPayable', d.year, d.currentTaxPayable)));
  }
  getCustomRows(options, 'balanceSheet', 'Current Liabilities').forEach(ci => {
    pushBsRowIfNonZero(clRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
  });

  const totalClRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalCurrentLiabilities', 'Total Current Liabilities'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'bs', 'totalCurrentLiabilities', d.year, d.totalCurrentLiabilities);
      return hasOv(options, 'bs', 'totalCurrentLiabilities', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), clRows), val);
    }),
  ]);

  const nclRows: number[] = [];
  if (isVis(options, 'bs_longTermBorrowings')) {
    pushBsRowIfNonZero(nclRows, lbl(options, 'bs', 'longTermBorrowings', 'Long-Term Borrowings'), data.map(d => ovVal(options, 'bs', 'longTermBorrowings', d.year, d.longTermBorrowings)));
  }
  if (isVis(options, 'bs_otherNonCurrentLiabilities')) {
    pushBsRowIfNonZero(nclRows, lbl(options, 'bs', 'otherNonCurrentLiabilities', 'Other Non-Current Liabilities'), data.map(d => ovVal(options, 'bs', 'otherNonCurrentLiabilities', d.year, d.otherNonCurrentLiabilities)));
  }
  getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities').forEach(ci => {
    pushBsRowIfNonZero(nclRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
  });

  const totalNclRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalNonCurrentLiabilities', 'Total Non-Current Liabilities'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'bs', 'totalNonCurrentLiabilities', d.year, d.totalNonCurrentLiabilities);
      return hasOv(options, 'bs', 'totalNonCurrentLiabilities', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), nclRows), val);
    }),
  ]);

  const totalLiabRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalLiabilities', 'TOTAL LIABILITIES'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'bs', 'totalLiabilities', d.year, d.totalLiabilities);
      return hasOv(options, 'bs', 'totalLiabilities', d.year) ? nCell(val) : fCell(`${c}${totalClRow}+${c}${totalNclRow}`, val);
    }),
  ]);

  const eqRows: number[] = [];
  if (isVis(options, 'bs_shareCapital')) {
    pushBsRowIfNonZero(eqRows, lbl(options, 'bs', 'shareCapital', 'Share Capital'), data.map(d => ovVal(options, 'bs', 'shareCapital', d.year, d.shareCapital)));
  }
  if (isVis(options, 'bs_retainedEarnings')) {
    pushBsRowIfNonZero(eqRows, lbl(options, 'bs', 'retainedEarnings', 'Retained Earnings'), data.map(d => ovVal(options, 'bs', 'retainedEarnings', d.year, d.retainedEarnings)));
  }
  if (isVis(options, 'bs_reserves')) {
    pushBsRowIfNonZero(eqRows, lbl(options, 'bs', 'reserves', 'Reserves'), data.map(d => ovVal(options, 'bs', 'reserves', d.year, d.reserves)));
  }
  getCustomRows(options, 'balanceSheet', 'Equity').forEach(ci => {
    pushBsRowIfNonZero(eqRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
  });

  const totalEqRow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalEquity', 'TOTAL EQUITY'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'bs', 'totalEquity', d.year, d.totalEquity);
      return hasOv(options, 'bs', 'totalEquity', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), eqRows), val);
    }),
  ]);

  const totalLERow = nextRow();
  rows.push([
    lbl(options, 'bs', 'totalLiabilitiesAndEquity', 'TOTAL LIABILITIES & EQUITY'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity);
      return hasOv(options, 'bs', 'totalLiabilitiesAndEquity', d.year) ? nCell(val) : fCell(`${c}${totalLiabRow}+${c}${totalEqRow}`, val);
    }),
  ]);

  rows.push([
    'Balance Check (Variance)',
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const totA = ovVal(options, 'bs', 'totalAssets', d.year, d.totalAssets);
      const totLE = ovVal(options, 'bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity);
      return fCell(`${c}${totalAssetsRow}-${c}${totalLERow}`, Math.round((totA - totLE) * 100) / 100);
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

  const pushCfRowIfNonZero = (rowTracker: number[], label: string, vals: number[], mandatory = false) => {
    if (!mandatory && !hasAnyNonZeroValue(vals)) return;
    rowTracker.push(nextRow());
    rows.push([label, ...vals.map(v => nCell(v))]);
  };

  rows.push(['CASH FLOWS FROM OPERATING ACTIVITIES (Indirect Method)', ...data.map(() => '')]);
  const nonCashRows: number[] = [];
  if (isVis(options, 'cf_pbt')) {
    pushCfRowIfNonZero(nonCashRows, lbl(options, 'cf', 'profitBeforeTax', 'Net Profit Before Tax'), data.map(d => ovVal(options, 'cf', 'profitBeforeTax', d.year, d.profitBeforeTax)));
  }
  if (isVis(options, 'cf_dep')) {
    pushCfRowIfNonZero(nonCashRows, lbl(options, 'cf', 'adjDepreciation', 'Adjustments for Depreciation'), data.map(d => ovVal(options, 'cf', 'adjDepreciation', d.year, d.adjDepreciation)));
  }
  if (isVis(options, 'cf_amort')) {
    pushCfRowIfNonZero(nonCashRows, lbl(options, 'cf', 'adjAmortization', 'Adjustments for Amortization'), data.map(d => ovVal(options, 'cf', 'adjAmortization', d.year, d.adjAmortization)));
  }
  if (isVis(options, 'cf_finance')) {
    pushCfRowIfNonZero(nonCashRows, lbl(options, 'cf', 'adjFinanceCost', 'Finance Cost Adjustment'), data.map(d => ovVal(options, 'cf', 'adjFinanceCost', d.year, d.adjFinanceCost)));
  }
  if (isVis(options, 'cf_interest_inc')) {
    pushCfRowIfNonZero(nonCashRows, lbl(options, 'cf', 'adjInterestIncome', 'Interest & Investment Income'), data.map(d => ovVal(options, 'cf', 'adjInterestIncome', d.year, d.adjInterestIncome)));
  }
  if (isVis(options, 'cf_disposal_gain_loss')) {
    pushCfRowIfNonZero(nonCashRows, lbl(options, 'cf', 'adjLossGainDisposal', 'Loss / (Gain) on Disposal of Fixed Assets'), data.map(d => ovVal(options, 'cf', 'adjLossGainDisposal', d.year, d.adjLossGainDisposal)));
  }

  const opBeforeWcRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'operatingProfitBeforeWC', 'Operating Profit Before Working Capital Changes'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'cf', 'operatingProfitBeforeWC', d.year, d.operatingProfitBeforeWC);
      return hasOv(options, 'cf', 'operatingProfitBeforeWC', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), nonCashRows), val);
    }),
  ]);

  const wcRows: number[] = [opBeforeWcRow];
  if (isVis(options, 'cf_wc_inventory')) {
    pushCfRowIfNonZero(wcRows, lbl(options, 'cf', 'changeInInventory', '(Increase) / Decrease in Inventory'), data.map(d => ovVal(options, 'cf', 'changeInInventory', d.year, d.changeInInventory)));
  }
  if (isVis(options, 'cf_wc_receivables')) {
    pushCfRowIfNonZero(wcRows, lbl(options, 'cf', 'changeInReceivables', '(Increase) / Decrease in Trade Receivables'), data.map(d => ovVal(options, 'cf', 'changeInReceivables', d.year, d.changeInReceivables)));
  }
  if (isVis(options, 'cf_wc_other_rec')) {
    pushCfRowIfNonZero(wcRows, lbl(options, 'cf', 'changeInOtherReceivables', '(Increase) / Decrease in Other Receivables'), data.map(d => ovVal(options, 'cf', 'changeInOtherReceivables', d.year, d.changeInOtherReceivables)));
  }
  if (isVis(options, 'cf_wc_other_ca')) {
    pushCfRowIfNonZero(wcRows, lbl(options, 'cf', 'changeInOtherCA', '(Increase) / Decrease in Other Current Assets'), data.map(d => ovVal(options, 'cf', 'changeInOtherCA', d.year, d.changeInOtherCA ?? 0)));
  }
  if (isVis(options, 'cf_wc_payables')) {
    pushCfRowIfNonZero(wcRows, lbl(options, 'cf', 'changeInPayables', 'Increase / (Decrease) in Trade Payables'), data.map(d => ovVal(options, 'cf', 'changeInPayables', d.year, d.changeInPayables)));
  }
  if (isVis(options, 'cf_wc_other_pay')) {
    pushCfRowIfNonZero(wcRows, lbl(options, 'cf', 'changeInOtherPayables', 'Increase / (Decrease) in Other Payables'), data.map(d => ovVal(options, 'cf', 'changeInOtherPayables', d.year, d.changeInOtherPayables)));
  }
  getCustomRows(options, 'cashFlow', 'Operating Activities').forEach(ci => {
    pushCfRowIfNonZero(wcRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
  });

  const cashGenOpsRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'cashGeneratedFromOperations', 'Cash Generated from Operations'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'cf', 'cashGeneratedFromOperations', d.year, d.cashGeneratedFromOperations);
      return hasOv(options, 'cf', 'cashGeneratedFromOperations', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), wcRows), val);
    }),
  ]);

  const opEndRows: number[] = [cashGenOpsRow];
  if (isVis(options, 'cf_tax_paid')) {
    pushCfRowIfNonZero(opEndRows, lbl(options, 'cf', 'incomeTaxPaid', 'Income Tax Paid'), data.map(d => ovVal(options, 'cf', 'incomeTaxPaid', d.year, d.incomeTaxPaid)));
  }
  if (isVis(options, 'cf_finance_paid')) {
    pushCfRowIfNonZero(opEndRows, lbl(options, 'cf', 'financeCostPaid', 'Finance Cost Paid'), data.map(d => ovVal(options, 'cf', 'financeCostPaid', d.year, d.financeCostPaid)));
  }

  const netOpCashRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'netCashFromOperating', 'NET CASH GENERATED FROM OPERATING ACTIVITIES'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'cf', 'netCashFromOperating', d.year, d.netCashFromOperating);
      return hasOv(options, 'cf', 'netCashFromOperating', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), opEndRows), val);
    }),
  ]);

  rows.push([]);
  rows.push(['CASH FLOWS FROM INVESTING ACTIVITIES', ...data.map(() => '')]);
  const invRows: number[] = [];
  if (isVis(options, 'cf_pur_ppe')) {
    pushCfRowIfNonZero(invRows, lbl(options, 'cf', 'purchasePPE', 'Purchase of Property, Plant & Equipment (Capex)'), data.map(d => ovVal(options, 'cf', 'purchasePPE', d.year, d.purchasePPE)));
  }
  if (isVis(options, 'cf_sale_ppe')) {
    pushCfRowIfNonZero(invRows, lbl(options, 'cf', 'salePPE', 'Proceeds from Sale of Fixed Assets'), data.map(d => ovVal(options, 'cf', 'salePPE', d.year, d.salePPE)));
  }
  if (isVis(options, 'cf_pur_intangible')) {
    pushCfRowIfNonZero(invRows, lbl(options, 'cf', 'purchaseIntangibles', 'Purchase of Intangible Assets'), data.map(d => ovVal(options, 'cf', 'purchaseIntangibles', d.year, d.purchaseIntangibles)));
  }
  if (isVis(options, 'cf_interest_rec')) {
    pushCfRowIfNonZero(invRows, lbl(options, 'cf', 'interestReceived', 'Interest Received'), data.map(d => ovVal(options, 'cf', 'interestReceived', d.year, d.interestReceived)));
  }
  if (isVis(options, 'cf_dividends_rec')) {
    pushCfRowIfNonZero(invRows, lbl(options, 'cf', 'dividendsReceived', 'Dividends Received'), data.map(d => ovVal(options, 'cf', 'dividendsReceived', d.year, d.dividendsReceived)));
  }
  getCustomRows(options, 'cashFlow', 'Investing Activities').forEach(ci => {
    pushCfRowIfNonZero(invRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
  });

  const netInvCashRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'netCashFromInvesting', 'NET CASH FROM / (USED IN) INVESTING ACTIVITIES'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'cf', 'netCashFromInvesting', d.year, d.netCashFromInvesting);
      return hasOv(options, 'cf', 'netCashFromInvesting', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), invRows), val);
    }),
  ]);

  rows.push([]);
  rows.push(['CASH FLOWS FROM FINANCING ACTIVITIES (Mandatory Direct User Input)', ...data.map(() => '')]);

  let directFinRow: number | null = null;
  if (isVis(options, 'cf_direct_financing')) {
    const directVals = data.map(d => ovVal(options, 'cf', 'directlyEnteredFinancingCashFlow', d.year, d.directlyEnteredFinancingCashFlow));
    if (hasAnyNonZeroValue(directVals)) {
      directFinRow = nextRow();
      rows.push([lbl(options, 'cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Financing Cash Flow'), ...directVals.map(v => nCell(v))]);
    }
  }

  const finBreakdownRows: number[] = [];
  if (isVis(options, 'cf_shares_issue')) {
    pushCfRowIfNonZero(finBreakdownRows, lbl(options, 'cf', 'financingProceedsShares', 'Proceeds from Issue of Share Capital'), data.map(d => ovVal(options, 'cf', 'financingProceedsShares', d.year, d.financingProceedsShares)));
  }
  if (isVis(options, 'cf_borrow_proceeds')) {
    pushCfRowIfNonZero(finBreakdownRows, lbl(options, 'cf', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing'), data.map(d => ovVal(options, 'cf', 'financingProceedsBorrowings', d.year, d.financingProceedsBorrowings)));
  }
  // Mandatory inclusion of 'Repayment of Borrowings & Financing' takes precedence over nil-value omission
  pushCfRowIfNonZero(
    finBreakdownRows,
    lbl(options, 'cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing'),
    data.map(d =>
      isVis(options, 'cf_borrow_repay')
        ? -ovVal(options, 'cf', 'repaymentOfBorrowings', d.year, d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings))
        : 0
    ),
    true
  );
  if (isVis(options, 'cf_dividends_paid')) {
    pushCfRowIfNonZero(finBreakdownRows, lbl(options, 'cf', 'financingDividendsPaid', 'Dividends Paid'), data.map(d => ovVal(options, 'cf', 'financingDividendsPaid', d.year, d.financingDividendsPaid)));
  }
  const customFinRows = getCustomRows(options, 'cashFlow', 'Financing Activities');
  customFinRows.forEach(ci => {
    pushCfRowIfNonZero(finBreakdownRows, ci.label, data.map(d => getCustomItemYearValue(ci, d.year)));
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
      const val = ovVal(options, 'cf', 'netCashFromFinancing', d.year, d.netCashFromFinancing);
      if (hasOv(options, 'cf', 'netCashFromFinancing', d.year)) {
        return nCell(val);
      }
      if (directFinRow === null) {
        return fCell(excelSumRows(c, finBreakdownRows), val);
      }
      if (allStdFinVisible && Math.abs(val - d.directlyEnteredFinancingCashFlow) < 0.01) {
        return fCell(`${c}${directFinRow}`, val);
      }
      if (d.isFinancingBreakdownComplete) {
        return fCell(excelSumRows(c, finBreakdownRows), val);
      }
      return nCell(val);
    }),
  ]);

  rows.push([]);
  const netChangeRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'netChangeInCash', 'NET INCREASE / (DECREASE) IN CASH & EQUIVALENTS'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'cf', 'netChangeInCash', d.year, d.netChangeInCash);
      return hasOv(options, 'cf', 'netChangeInCash', d.year)
        ? nCell(val)
        : fCell(`${c}${netOpCashRow}+${c}${netInvCashRow}+${c}${netFinCashRow}`, val);
    }),
  ]);

  const openCashRow = nextRow();
  rows.push([lbl(options, 'cf', 'openingCash', 'Opening Cash & Cash Equivalents'), ...data.map(d => nCell(ovVal(options, 'cf', 'openingCash', d.year, d.openingCash)))]);

  const closeCashRow = nextRow();
  rows.push([
    lbl(options, 'cf', 'closingCashCalculated', 'Closing Cash & Cash Equivalents (Calculated)'),
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const val = ovVal(options, 'cf', 'closingCashCalculated', d.year, d.closingCashCalculated);
      return hasOv(options, 'cf', 'closingCashCalculated', d.year)
        ? nCell(val)
        : fCell(`${c}${netChangeRow}+${c}${openCashRow}`, val);
    }),
  ]);

  const bsCashRow = nextRow();
  rows.push([lbl(options, 'cf', 'balanceSheetCash', 'Stated Cash in Statement of Financial Position'), ...data.map(d => nCell(ovVal(options, 'cf', 'balanceSheetCash', d.year, d.balanceSheetCash)))]);

  rows.push([
    'Reconciliation Variance (Calculated - Stated)',
    ...data.map((d, idx) => {
      const c = colLetter(idx + 1);
      const closeVal = ovVal(options, 'cf', 'closingCashCalculated', d.year, d.closingCashCalculated);
      const bsCashVal = ovVal(options, 'cf', 'balanceSheetCash', d.year, d.balanceSheetCash);
      return fCell(`${c}${closeCashRow}-${c}${bsCashRow}`, Math.round((closeVal - bsCashVal) * 100) / 100);
    }),
  ]);
  rows.push([
    'Audit Reconciliation Status',
    ...data.map(d => {
      const closeVal = ovVal(options, 'cf', 'closingCashCalculated', d.year, d.closingCashCalculated);
      const bsCashVal = ovVal(options, 'cf', 'balanceSheetCash', d.year, d.balanceSheetCash);
      return Math.abs(closeVal - bsCashVal) < 0.01 ? 'RECONCILED' : 'VARIANCE_IDENTIFIED';
    }),
  ]);

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
    const vals = data.map(d => ovVal(options, 'eq', 'openingShareCapital', d.year, d.openingShareCapital));
    if (hasAnyNonZeroValue(vals)) {
      openScRow = nextRow();
      scRows.push(openScRow);
      rows.push([lbl(options, 'eq', 'openingShareCapital', 'Opening Share Capital'), ...vals.map(v => nCell(v))]);
    }
  }
  if (isVis(options, 'eq_shareCapitalIssued')) {
    const vals = data.map(d => ovVal(options, 'eq', 'shareCapitalIssued', d.year, d.shareCapitalIssued));
    if (hasAnyNonZeroValue(vals)) {
      scRows.push(nextRow());
      rows.push([lbl(options, 'eq', 'shareCapitalIssued', 'Share Capital Issued during Year'), ...vals.map(v => nCell(v))]);
    }
  }
  const closeScRow = nextRow();
  rows.push([
    lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'eq', 'closingShareCapital', d.year, d.closingShareCapital);
      return hasOv(options, 'eq', 'closingShareCapital', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), scRows), val);
    }),
  ]);
  rows.push([]);

  const reRows: number[] = [];
  let openReRow: number | null = null;
  if (isVis(options, 'eq_openingRetainedEarnings')) {
    const vals = data.map(d => ovVal(options, 'eq', 'openingRetainedEarnings', d.year, d.openingRetainedEarnings));
    if (hasAnyNonZeroValue(vals)) {
      openReRow = nextRow();
      reRows.push(openReRow);
      rows.push([lbl(options, 'eq', 'openingRetainedEarnings', 'Opening Retained Earnings'), ...vals.map(v => nCell(v))]);
    }
  }
  if (isVis(options, 'eq_netProfitForYear')) {
    const vals = data.map(d => ovVal(options, 'eq', 'netProfitForYear', d.year, d.netProfitForYear));
    if (hasAnyNonZeroValue(vals)) {
      reRows.push(nextRow());
      rows.push([lbl(options, 'eq', 'netProfitForYear', 'Net Profit for the Year'), ...vals.map(v => nCell(v))]);
    }
  }
  if (isVis(options, 'eq_dividendsPaid')) {
    const vals = data.map(d => -ovVal(options, 'eq', 'dividendsPaid', d.year, d.dividendsPaid));
    if (hasAnyNonZeroValue(vals)) {
      reRows.push(nextRow());
      rows.push([lbl(options, 'eq', 'dividendsPaid', 'Dividends Paid to Shareholders'), ...vals.map(v => nCell(v))]);
    }
  }
  const closeReRow = nextRow();
  rows.push([
    lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'eq', 'closingRetainedEarnings', d.year, d.closingRetainedEarnings);
      return hasOv(options, 'eq', 'closingRetainedEarnings', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), reRows), val);
    }),
  ]);

  const resRows: number[] = [];
  let openResRow: number | null = null;
  const openResVals = data.map(d => ovVal(options, 'eq', 'openingReserves', d.year, d.openingReserves));
  const ociVals = data.map(d => ovVal(options, 'eq', 'otherComprehensiveIncome', d.year, d.otherComprehensiveIncome));
  const closeResVals = data.map(d => ovVal(options, 'eq', 'closingReserves', d.year, d.closingReserves));
  const showOpenRes = isVis(options, 'eq_openingReserves') && hasAnyNonZeroValue(openResVals);
  const showOci = isVis(options, 'eq_otherComprehensiveIncome') && hasAnyNonZeroValue(ociVals);
  const showResBlock = showOpenRes || showOci || hasAnyNonZeroValue(closeResVals);
  let closeResRow: number | null = null;
  if (showResBlock) {
    rows.push([]);
    if (showOpenRes) {
      openResRow = nextRow();
      resRows.push(openResRow);
      rows.push([lbl(options, 'eq', 'openingReserves', 'Opening Reserves'), ...openResVals.map(v => nCell(v))]);
    }
    if (showOci) {
      resRows.push(nextRow());
      rows.push([lbl(options, 'eq', 'otherComprehensiveIncome', 'Other Comprehensive Income / Transfers'), ...ociVals.map(v => nCell(v))]);
    }
    closeResRow = nextRow();
    rows.push([
      lbl(options, 'eq', 'closingReserves', 'Closing Reserves'),
      ...data.map((d, idx) => {
        const val = ovVal(options, 'eq', 'closingReserves', d.year, d.closingReserves);
        return hasOv(options, 'eq', 'closingReserves', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), resRows), val);
      }),
    ]);
  }

  const customEqRows: number[] = [];
  getCustomRows(options, 'equity').forEach(ci => {
    const vals = data.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) {
      customEqRows.push(nextRow());
      rows.push([ci.label, ...vals.map(v => nCell(v))]);
    }
  });

  rows.push([]);
  const openEqSumRows = [openScRow, openReRow, openResRow].filter((x): x is number => x !== null);
  rows.push([
    lbl(options, 'eq', 'totalOpeningEquity', 'TOTAL OPENING EQUITY'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'eq', 'totalOpeningEquity', d.year, d.totalOpeningEquity);
      return hasOv(options, 'eq', 'totalOpeningEquity', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), openEqSumRows), val);
    }),
  ]);
  const closeEqSumRows = [closeScRow, closeReRow, ...(closeResRow !== null ? [closeResRow] : []), ...customEqRows];
  rows.push([
    lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'),
    ...data.map((d, idx) => {
      const val = ovVal(options, 'eq', 'totalClosingEquity', d.year, d.totalClosingEquity);
      return hasOv(options, 'eq', 'totalClosingEquity', d.year) ? nCell(val) : fCell(excelSumRows(colLetter(idx + 1), closeEqSumRows), val);
    }),
  ]);

  const sheetData = createTableDataForStatement(title, company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [42, ...data.map(() => 18)]);
  return ws;
}

export function resolveFACategoriesForExport(
  latest: ScheduleOfFixedAssetsYear,
  options?: ExportSyncOptions
) {
  const stdSource =
    latest.allCategories && latest.allCategories.length > 0
      ? latest.allCategories
      : latest.categories;
  const visibleStd = stdSource
    .map((cat, idx) => {
      const cfgId = FIXED_ASSET_CATEGORY_CONFIG_IDS[idx];
      const name = resolveEffectiveLabel('fa', String(idx), cat.name, options?.labelOverrides, options?.statementConfig);
      return {
        cat: { ...cat, name },
        visible: cfgId && options?.statementConfig ? isVis(options, cfgId) : true,
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
  return [...visibleStd, ...customRows].filter(c =>
    hasAnyNonZeroValue([
      c.costOpening,
      c.additions,
      c.disposals,
      c.costClosing,
      c.depOpening,
      c.depCharge,
      c.depDisposals,
      c.depClosing,
      c.netBookValue,
    ])
  );
}

function buildFixedAssetsWorksheet(
  company: CompanyProfile,
  data: ScheduleOfFixedAssetsYear[],
  options?: ExportSyncOptions
): XLSX.WorkSheet {
  const latest = data[data.length - 1];
  const headers = ['Asset Class', 'Dep %', 'Opening Cost', 'Additions', 'Disposals', 'Closing Cost', 'Opening Dep.', 'Dep. Charge', 'Dep. Disposals', 'Closing Dep.', 'Net Book Value'];
  if (!latest) {
    const sheetData = createTableDataForStatement('Schedule of Fixed Assets', company, headers, []);
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    applyColumnWidths(ws, [38, 10, 16, 14, 14, 16, 16, 14, 15, 16, 18]);
    return ws;
  }
  const categories = resolveFACategoriesForExport(latest, options);
  const r = STD_DATA_START_ROW;
  const rows: ExcelCellValue[][] = categories.map((c, idx) => {
    const rowNum = r + idx;
    return [
      c.name,
      `${c.depRate}%`,
      nCell(c.costOpening),
      nCell(c.additions),
      nCell(c.disposals),
      fCell(`C${rowNum}+D${rowNum}-E${rowNum}`, c.costClosing),
      nCell(c.depOpening),
      nCell(c.depCharge),
      nCell(c.depDisposals),
      fCell(`G${rowNum}+H${rowNum}-I${rowNum}`, c.depClosing),
      fCell(`F${rowNum}-J${rowNum}`, c.netBookValue),
    ];
  });

  if (categories.length > 0) {
    const endRow = r + categories.length - 1;
    rows.push([
      'TOTAL SCHEDULE OF FIXED ASSETS',
      '-',
      fCell(`SUM(C${r}:C${endRow})`, latest.totalCostOpening),
      fCell(`SUM(D${r}:D${endRow})`, latest.totalAdditions),
      fCell(`SUM(E${r}:E${endRow})`, latest.totalDisposals),
      fCell(`SUM(F${r}:F${endRow})`, latest.totalCostClosing),
      fCell(`SUM(G${r}:G${endRow})`, latest.totalDepOpening),
      fCell(`SUM(H${r}:H${endRow})`, latest.totalDepCharge),
      fCell(`SUM(I${r}:I${endRow})`, latest.totalDepDisposals),
      fCell(`SUM(J${r}:J${endRow})`, latest.totalDepClosing),
      fCell(`SUM(K${r}:K${endRow})`, latest.totalNetBookValue),
    ]);
  } else {
    rows.push([
      'TOTAL SCHEDULE OF FIXED ASSETS',
      '-',
      nCell(0), nCell(0), nCell(0), nCell(0),
      nCell(0), nCell(0), nCell(0), nCell(0), nCell(0),
    ]);
  }

  const sheetData = createTableDataForStatement(`Schedule of Fixed Assets (${latest.year})`, company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [38, 10, 16, 14, 14, 16, 16, 14, 15, 16, 18]);
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
    const filteredTbl = filterNoteTableForExport(note.tableData);
    if (filteredTbl) {
      rows.push(filteredTbl.headers);
      filteredTbl.rows.forEach(r => {
        rows.push([r.label, ...r.values]);
      });
      if (filteredTbl.totalRow) {
        rows.push([filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]);
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
    const filteredTbl = filterNoteTableForExport(note.tableData);
    if (filteredTbl) {
      notesRows.push(filteredTbl.headers);
      filteredTbl.rows.forEach(r => notesRows.push([r.label, ...r.values]));
      if (filteredTbl.totalRow) {
        notesRows.push([filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]);
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
          label === 'Operating Profit (EBIT)' ||
          label === 'Profit Before Tax (EBT)' ||
          label === 'Net Profit After Tax' ||
          label === 'NET PROFIT AFTER TAX' ||
          label === 'Net Operating Profit' ||
          label === 'Net Profit After Income Tax' ||
          label === 'Operating Profit Before Working Capital Changes' ||
          label === 'Cash Generated from Operations' ||
          label.includes('Net Cash') ||
          label.includes('NET CASH') ||
          label.includes('NET INCREASE') ||
          label.includes('Net Change in Cash') ||
          label.includes('Capital after drawings') ||
          label.includes('Total Operating Profit') ||
          label.includes('Net Profit/(Loss)') ||
          label.includes('Closing Cash') ||
          label.startsWith('Closing Share') ||
          label.startsWith('Closing Retained') ||
          label.startsWith('Closing Reserves') ||
          label.startsWith('Closing Capital');

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
  const pushIfNonZero = (label: string, vals: number[]) => {
    if (hasAnyNonZeroValue(vals)) rows.push([label, ...vals]);
  };
  if (isVis(options, 'inc_revenue')) {
    pushIfNonZero(
      lbl(options, 'income', 'revenue', 'Sales Revenue'),
      incomeData.map(d => ovVal(options, 'income', 'revenue', d.year, d.revenue - customRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)))
    );
  }
  customRev.forEach(ci => pushIfNonZero(ci.label, incomeData.map(d => getCustomItemYearValue(ci, d.year))));
  if (isVis(options, 'inc_costOfSales')) {
    pushIfNonZero(lbl(options, 'income', 'costOfSales', 'Cost of Sales'), incomeData.map(d => -ovVal(options, 'income', 'costOfSales', d.year, d.costOfSales)));
  }
  rows.push([lbl(options, 'income', 'grossProfit', 'Gross Profit'), ...incomeData.map(d => ovVal(options, 'income', 'grossProfit', d.year, d.grossProfit))]);
  if (isVis(options, 'inc_operatingExpenses')) {
    pushIfNonZero(
      lbl(options, 'income', 'operatingExpenses', 'Operating Expenses'),
      incomeData.map(d => -ovVal(options, 'income', 'operatingExpenses', d.year, d.operatingExpenses + customOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)))
    );
  }
  customOpex.forEach(ci => pushIfNonZero(ci.label, incomeData.map(d => getCustomItemYearValue(ci, d.year))));
  if (isVis(options, 'inc_depreciationExpense')) {
    pushIfNonZero(lbl(options, 'income', 'depreciationExpense', 'Depreciation Expense'), incomeData.map(d => -ovVal(options, 'income', 'depreciationExpense', d.year, d.depreciationExpense)));
  }
  rows.push([lbl(options, 'income', 'operatingProfit', 'Operating Profit (EBIT)'), ...incomeData.map(d => ovVal(options, 'income', 'operatingProfit', d.year, d.operatingProfit))]);
  if (isVis(options, 'inc_interestIncome')) {
    pushIfNonZero(
      lbl(options, 'income', 'interestIncome', 'Interest & Investment Income'),
      incomeData.map(d => ovVal(options, 'income', 'interestIncome', d.year, d.interestIncome - customNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)))
    );
  }
  if (isVis(options, 'inc_financeCost')) {
    pushIfNonZero(lbl(options, 'income', 'financeCost', 'Finance Cost (Mandatory Direct Entry)'), incomeData.map(d => -ovVal(options, 'income', 'financeCost', d.year, d.financeCost)));
  }
  customNonOp.forEach(ci => pushIfNonZero(ci.label, incomeData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'income', 'profitBeforeTax', 'Profit Before Tax (EBT)'), ...incomeData.map(d => ovVal(options, 'income', 'profitBeforeTax', d.year, d.profitBeforeTax))]);
  if (isVis(options, 'inc_taxExpense')) {
    pushIfNonZero(lbl(options, 'income', 'taxExpense', 'Income Tax Expense'), incomeData.map(d => -ovVal(options, 'income', 'taxExpense', d.year, d.taxExpense)));
  }
  rows.push([lbl(options, 'income', 'netProfit', 'Net Profit After Tax'), ...incomeData.map(d => ovVal(options, 'income', 'netProfit', d.year, d.netProfit))]);
  return rows;
}

function buildSyncBalanceTextRows(balanceData: BalanceSheetYear[], options?: ExportSyncOptions): (string | number)[][] {
  const rows: (string | number)[][] = [];
  const pushIfNonZero = (label: string, vals: number[]) => {
    if (hasAnyNonZeroValue(vals)) rows.push([label, ...vals]);
  };
  if (isVis(options, 'bs_cashAndEquivalents')) pushIfNonZero(lbl(options, 'bs', 'cashAndEquivalents', 'Cash and Cash Equivalents'), balanceData.map(d => ovVal(options, 'bs', 'cashAndEquivalents', d.year, d.cashAndEquivalents)));
  if (isVis(options, 'bs_tradeReceivables')) pushIfNonZero(lbl(options, 'bs', 'tradeReceivables', 'Trade Receivables'), balanceData.map(d => ovVal(options, 'bs', 'tradeReceivables', d.year, d.tradeReceivables)));
  if (isVis(options, 'bs_otherReceivables')) pushIfNonZero(lbl(options, 'bs', 'otherReceivables', 'Other Receivables & Prepayments'), balanceData.map(d => ovVal(options, 'bs', 'otherReceivables', d.year, d.otherReceivables)));
  if (isVis(options, 'bs_closingInventory')) pushIfNonZero(lbl(options, 'bs', 'closingInventory', 'Closing Inventory'), balanceData.map(d => ovVal(options, 'bs', 'closingInventory', d.year, d.closingInventory)));
  if (isVis(options, 'bs_otherCurrentAssets')) pushIfNonZero(lbl(options, 'bs', 'otherCurrentAssets', 'Other Current Assets'), balanceData.map(d => ovVal(options, 'bs', 'otherCurrentAssets', d.year, d.otherCurrentAssets)));
  getCustomRows(options, 'balanceSheet', 'Current Assets').forEach(ci => pushIfNonZero(ci.label, balanceData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'bs', 'totalCurrentAssets', 'Total Current Assets'), ...balanceData.map(d => ovVal(options, 'bs', 'totalCurrentAssets', d.year, d.totalCurrentAssets))]);
  if (isVis(options, 'bs_propertyPlantEquipment')) pushIfNonZero(lbl(options, 'bs', 'propertyPlantEquipment', 'Property, Plant and Equipment (NBV)'), balanceData.map(d => ovVal(options, 'bs', 'propertyPlantEquipment', d.year, d.propertyPlantEquipment)));
  if (isVis(options, 'bs_intangibleAssets')) pushIfNonZero(lbl(options, 'bs', 'intangibleAssets', 'Intangible Assets'), balanceData.map(d => ovVal(options, 'bs', 'intangibleAssets', d.year, d.intangibleAssets)));
  if (isVis(options, 'bs_otherNonCurrentAssets')) pushIfNonZero(lbl(options, 'bs', 'otherNonCurrentAssets', 'Other Non-Current Assets'), balanceData.map(d => ovVal(options, 'bs', 'otherNonCurrentAssets', d.year, d.otherNonCurrentAssets)));
  getCustomRows(options, 'balanceSheet', 'Non-Current Assets').forEach(ci => pushIfNonZero(ci.label, balanceData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'bs', 'totalNonCurrentAssets', 'Total Non-Current Assets'), ...balanceData.map(d => ovVal(options, 'bs', 'totalNonCurrentAssets', d.year, d.totalNonCurrentAssets))]);
  rows.push([lbl(options, 'bs', 'totalAssets', 'TOTAL ASSETS'), ...balanceData.map(d => ovVal(options, 'bs', 'totalAssets', d.year, d.totalAssets))]);
  if (isVis(options, 'bs_tradePayables')) pushIfNonZero(lbl(options, 'bs', 'tradePayables', 'Trade Payables'), balanceData.map(d => ovVal(options, 'bs', 'tradePayables', d.year, d.tradePayables)));
  if (isVis(options, 'bs_otherPayables')) pushIfNonZero(lbl(options, 'bs', 'otherPayables', 'Other Payables & Accruals'), balanceData.map(d => ovVal(options, 'bs', 'otherPayables', d.year, d.otherPayables)));
  if (isVis(options, 'bs_shortTermBorrowings')) pushIfNonZero(lbl(options, 'bs', 'shortTermBorrowings', 'Short-Term Borrowings'), balanceData.map(d => ovVal(options, 'bs', 'shortTermBorrowings', d.year, d.shortTermBorrowings)));
  if (isVis(options, 'bs_currentTaxPayable')) pushIfNonZero(lbl(options, 'bs', 'currentTaxPayable', 'Current Tax Payable'), balanceData.map(d => ovVal(options, 'bs', 'currentTaxPayable', d.year, d.currentTaxPayable)));
  getCustomRows(options, 'balanceSheet', 'Current Liabilities').forEach(ci => pushIfNonZero(ci.label, balanceData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'bs', 'totalCurrentLiabilities', 'Total Current Liabilities'), ...balanceData.map(d => ovVal(options, 'bs', 'totalCurrentLiabilities', d.year, d.totalCurrentLiabilities))]);
  if (isVis(options, 'bs_longTermBorrowings')) pushIfNonZero(lbl(options, 'bs', 'longTermBorrowings', 'Long-Term Borrowings'), balanceData.map(d => ovVal(options, 'bs', 'longTermBorrowings', d.year, d.longTermBorrowings)));
  if (isVis(options, 'bs_otherNonCurrentLiabilities')) pushIfNonZero(lbl(options, 'bs', 'otherNonCurrentLiabilities', 'Other Non-Current Liabilities'), balanceData.map(d => ovVal(options, 'bs', 'otherNonCurrentLiabilities', d.year, d.otherNonCurrentLiabilities)));
  getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities').forEach(ci => pushIfNonZero(ci.label, balanceData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'bs', 'totalNonCurrentLiabilities', 'Total Non-Current Liabilities'), ...balanceData.map(d => ovVal(options, 'bs', 'totalNonCurrentLiabilities', d.year, d.totalNonCurrentLiabilities))]);
  rows.push([lbl(options, 'bs', 'totalLiabilities', 'TOTAL LIABILITIES'), ...balanceData.map(d => ovVal(options, 'bs', 'totalLiabilities', d.year, d.totalLiabilities))]);
  if (isVis(options, 'bs_shareCapital')) pushIfNonZero(lbl(options, 'bs', 'shareCapital', 'Share Capital'), balanceData.map(d => ovVal(options, 'bs', 'shareCapital', d.year, d.shareCapital)));
  if (isVis(options, 'bs_retainedEarnings')) pushIfNonZero(lbl(options, 'bs', 'retainedEarnings', 'Retained Earnings'), balanceData.map(d => ovVal(options, 'bs', 'retainedEarnings', d.year, d.retainedEarnings)));
  if (isVis(options, 'bs_reserves')) pushIfNonZero(lbl(options, 'bs', 'reserves', 'Reserves'), balanceData.map(d => ovVal(options, 'bs', 'reserves', d.year, d.reserves)));
  getCustomRows(options, 'balanceSheet', 'Equity').forEach(ci => pushIfNonZero(ci.label, balanceData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'bs', 'totalEquity', 'TOTAL EQUITY'), ...balanceData.map(d => ovVal(options, 'bs', 'totalEquity', d.year, d.totalEquity))]);
  rows.push([lbl(options, 'bs', 'totalLiabilitiesAndEquity', 'TOTAL LIABILITIES & EQUITY'), ...balanceData.map(d => ovVal(options, 'bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity))]);
  rows.push([
    'Balance Check',
    ...balanceData.map(d => {
      const totA = ovVal(options, 'bs', 'totalAssets', d.year, d.totalAssets);
      const totLE = ovVal(options, 'bs', 'totalLiabilitiesAndEquity', d.year, d.totalLiabilitiesAndEquity);
      return Math.round((totA - totLE) * 100) / 100;
    }),
  ]);
  return rows;
}

export function buildSyncCashFlowTextRows(cashFlowData: CashFlowStatementYear[], options?: ExportSyncOptions): (string | number)[][] {
  const rows: (string | number)[][] = [];
  const pushIfNonZero = (label: string, vals: number[]) => {
    if (hasAnyNonZeroValue(vals)) rows.push([label, ...vals]);
  };
  if (isVis(options, 'cf_pbt')) pushIfNonZero(lbl(options, 'cf', 'profitBeforeTax', 'Net Profit Before Tax'), cashFlowData.map(d => ovVal(options, 'cf', 'profitBeforeTax', d.year, d.profitBeforeTax)));
  if (isVis(options, 'cf_dep')) pushIfNonZero(lbl(options, 'cf', 'adjDepreciation', 'Adjustments for Depreciation'), cashFlowData.map(d => ovVal(options, 'cf', 'adjDepreciation', d.year, d.adjDepreciation)));
  if (isVis(options, 'cf_amort')) pushIfNonZero(lbl(options, 'cf', 'adjAmortization', 'Adjustments for Amortization'), cashFlowData.map(d => ovVal(options, 'cf', 'adjAmortization', d.year, d.adjAmortization)));
  if (isVis(options, 'cf_finance')) pushIfNonZero(lbl(options, 'cf', 'adjFinanceCost', 'Finance Cost Adjustment'), cashFlowData.map(d => ovVal(options, 'cf', 'adjFinanceCost', d.year, d.adjFinanceCost)));
  if (isVis(options, 'cf_interest_inc')) pushIfNonZero(lbl(options, 'cf', 'adjInterestIncome', 'Interest & Investment Income'), cashFlowData.map(d => ovVal(options, 'cf', 'adjInterestIncome', d.year, d.adjInterestIncome)));
  if (isVis(options, 'cf_disposal_gain_loss')) pushIfNonZero(lbl(options, 'cf', 'adjLossGainDisposal', 'Loss / (Gain) on Disposal of Fixed Assets'), cashFlowData.map(d => ovVal(options, 'cf', 'adjLossGainDisposal', d.year, d.adjLossGainDisposal)));
  rows.push([lbl(options, 'cf', 'operatingProfitBeforeWC', 'Operating Profit Before Working Capital Changes'), ...cashFlowData.map(d => ovVal(options, 'cf', 'operatingProfitBeforeWC', d.year, d.operatingProfitBeforeWC))]);
  if (isVis(options, 'cf_wc_inventory')) pushIfNonZero(lbl(options, 'cf', 'changeInInventory', '(Increase) / Decrease in Inventory'), cashFlowData.map(d => ovVal(options, 'cf', 'changeInInventory', d.year, d.changeInInventory)));
  if (isVis(options, 'cf_wc_receivables')) pushIfNonZero(lbl(options, 'cf', 'changeInReceivables', '(Increase) / Decrease in Trade Receivables'), cashFlowData.map(d => ovVal(options, 'cf', 'changeInReceivables', d.year, d.changeInReceivables)));
  if (isVis(options, 'cf_wc_other_rec')) pushIfNonZero(lbl(options, 'cf', 'changeInOtherReceivables', '(Increase) / Decrease in Other Receivables'), cashFlowData.map(d => ovVal(options, 'cf', 'changeInOtherReceivables', d.year, d.changeInOtherReceivables)));
  if (isVis(options, 'cf_wc_other_ca')) pushIfNonZero(lbl(options, 'cf', 'changeInOtherCA', '(Increase) / Decrease in Other Current Assets'), cashFlowData.map(d => ovVal(options, 'cf', 'changeInOtherCA', d.year, d.changeInOtherCA ?? 0)));
  if (isVis(options, 'cf_wc_payables')) pushIfNonZero(lbl(options, 'cf', 'changeInPayables', 'Increase / (Decrease) in Trade Payables'), cashFlowData.map(d => ovVal(options, 'cf', 'changeInPayables', d.year, d.changeInPayables)));
  if (isVis(options, 'cf_wc_other_pay')) pushIfNonZero(lbl(options, 'cf', 'changeInOtherPayables', 'Increase / (Decrease) in Other Payables'), cashFlowData.map(d => ovVal(options, 'cf', 'changeInOtherPayables', d.year, d.changeInOtherPayables)));
  getCustomRows(options, 'cashFlow', 'Operating Activities').forEach(ci => pushIfNonZero(ci.label, cashFlowData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'cf', 'cashGeneratedFromOperations', 'Cash Generated from Operations'), ...cashFlowData.map(d => ovVal(options, 'cf', 'cashGeneratedFromOperations', d.year, d.cashGeneratedFromOperations))]);
  if (isVis(options, 'cf_tax_paid')) pushIfNonZero(lbl(options, 'cf', 'incomeTaxPaid', 'Income Tax Paid'), cashFlowData.map(d => ovVal(options, 'cf', 'incomeTaxPaid', d.year, d.incomeTaxPaid)));
  if (isVis(options, 'cf_finance_paid')) pushIfNonZero(lbl(options, 'cf', 'financeCostPaid', 'Finance Cost Paid'), cashFlowData.map(d => ovVal(options, 'cf', 'financeCostPaid', d.year, d.financeCostPaid)));
  rows.push([lbl(options, 'cf', 'netCashFromOperating', 'NET CASH FROM OPERATING ACTIVITIES'), ...cashFlowData.map(d => ovVal(options, 'cf', 'netCashFromOperating', d.year, d.netCashFromOperating))]);
  if (isVis(options, 'cf_pur_ppe')) pushIfNonZero(lbl(options, 'cf', 'purchasePPE', 'Purchase of PPE (Capex)'), cashFlowData.map(d => ovVal(options, 'cf', 'purchasePPE', d.year, d.purchasePPE)));
  if (isVis(options, 'cf_sale_ppe')) pushIfNonZero(lbl(options, 'cf', 'salePPE', 'Proceeds from Sale of Fixed Assets'), cashFlowData.map(d => ovVal(options, 'cf', 'salePPE', d.year, d.salePPE)));
  if (isVis(options, 'cf_pur_intangible')) pushIfNonZero(lbl(options, 'cf', 'purchaseIntangibles', 'Purchase of Intangible Assets'), cashFlowData.map(d => ovVal(options, 'cf', 'purchaseIntangibles', d.year, d.purchaseIntangibles)));
  if (isVis(options, 'cf_interest_rec')) pushIfNonZero(lbl(options, 'cf', 'interestReceived', 'Interest Received'), cashFlowData.map(d => ovVal(options, 'cf', 'interestReceived', d.year, d.interestReceived)));
  if (isVis(options, 'cf_dividends_rec')) pushIfNonZero(lbl(options, 'cf', 'dividendsReceived', 'Dividends Received'), cashFlowData.map(d => ovVal(options, 'cf', 'dividendsReceived', d.year, d.dividendsReceived)));
  getCustomRows(options, 'cashFlow', 'Investing Activities').forEach(ci => pushIfNonZero(ci.label, cashFlowData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'cf', 'netCashFromInvesting', 'NET CASH FROM / (USED IN) INVESTING'), ...cashFlowData.map(d => ovVal(options, 'cf', 'netCashFromInvesting', d.year, d.netCashFromInvesting))]);
  if (isVis(options, 'cf_shares_issue')) pushIfNonZero(lbl(options, 'cf', 'financingProceedsShares', 'Proceeds from Issue of Share Capital'), cashFlowData.map(d => ovVal(options, 'cf', 'financingProceedsShares', d.year, d.financingProceedsShares)));
  if (isVis(options, 'cf_borrow_proceeds')) pushIfNonZero(lbl(options, 'cf', 'financingProceedsBorrowings', 'Proceeds from Borrowings & Financing'), cashFlowData.map(d => ovVal(options, 'cf', 'financingProceedsBorrowings', d.year, d.financingProceedsBorrowings)));
  // Mandatory inclusion of 'Repayment of Borrowings & Financing' overrides nil-value omission
  rows.push([
    lbl(options, 'cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing'),
    ...cashFlowData.map(d =>
      isVis(options, 'cf_borrow_repay')
        ? -ovVal(options, 'cf', 'repaymentOfBorrowings', d.year, d.repaymentOfBorrowings || Math.abs(d.financingRepaymentBorrowings))
        : 0
    ),
  ]);
  if (isVis(options, 'cf_dividends_paid')) pushIfNonZero(lbl(options, 'cf', 'financingDividendsPaid', 'Dividends Paid'), cashFlowData.map(d => ovVal(options, 'cf', 'financingDividendsPaid', d.year, d.financingDividendsPaid)));
  getCustomRows(options, 'cashFlow', 'Financing Activities').forEach(ci => pushIfNonZero(ci.label, cashFlowData.map(d => getCustomItemYearValue(ci, d.year))));
  if (isVis(options, 'cf_direct_financing')) pushIfNonZero(lbl(options, 'cf', 'directlyEnteredFinancingCashFlow', 'Directly Entered Financing Cash Flow'), cashFlowData.map(d => ovVal(options, 'cf', 'directlyEnteredFinancingCashFlow', d.year, d.directlyEnteredFinancingCashFlow)));
  rows.push([lbl(options, 'cf', 'netCashFromFinancing', 'NET CASH FROM / (USED IN) FINANCING'), ...cashFlowData.map(d => ovVal(options, 'cf', 'netCashFromFinancing', d.year, d.netCashFromFinancing))]);
  rows.push([lbl(options, 'cf', 'netChangeInCash', 'Net Change in Cash and Equivalents'), ...cashFlowData.map(d => ovVal(options, 'cf', 'netChangeInCash', d.year, d.netChangeInCash))]);
  rows.push([lbl(options, 'cf', 'openingCash', 'Opening Cash & Cash Equivalents'), ...cashFlowData.map(d => ovVal(options, 'cf', 'openingCash', d.year, d.openingCash))]);
  rows.push([lbl(options, 'cf', 'closingCashCalculated', 'Closing Cash (Calculated)'), ...cashFlowData.map(d => ovVal(options, 'cf', 'closingCashCalculated', d.year, d.closingCashCalculated))]);
  rows.push([lbl(options, 'cf', 'balanceSheetCash', 'Stated Balance Sheet Cash'), ...cashFlowData.map(d => ovVal(options, 'cf', 'balanceSheetCash', d.year, d.balanceSheetCash))]);
  rows.push([
    'Reconciliation Difference',
    ...cashFlowData.map(d => {
      const closeVal = ovVal(options, 'cf', 'closingCashCalculated', d.year, d.closingCashCalculated);
      const bsCashVal = ovVal(options, 'cf', 'balanceSheetCash', d.year, d.balanceSheetCash);
      return Math.round((closeVal - bsCashVal) * 100) / 100;
    }),
  ]);
  rows.push([
    'Audit Reconciliation Status',
    ...cashFlowData.map(d => {
      const closeVal = ovVal(options, 'cf', 'closingCashCalculated', d.year, d.closingCashCalculated);
      const bsCashVal = ovVal(options, 'cf', 'balanceSheetCash', d.year, d.balanceSheetCash);
      return Math.abs(closeVal - bsCashVal) < 0.01 ? 'RECONCILED' : 'VARIANCE_IDENTIFIED';
    }),
  ]);
  return rows;
}

function buildSyncEquityTextRows(equityData: EquityStatementYear[], options?: ExportSyncOptions): (string | number)[][] {
  const rows: (string | number)[][] = [];
  const pushIfNonZero = (label: string, vals: number[]) => {
    if (hasAnyNonZeroValue(vals)) rows.push([label, ...vals]);
  };
  if (isVis(options, 'eq_openingShareCapital')) pushIfNonZero(lbl(options, 'eq', 'openingShareCapital', 'Opening Share Capital'), equityData.map(d => ovVal(options, 'eq', 'openingShareCapital', d.year, d.openingShareCapital)));
  if (isVis(options, 'eq_shareCapitalIssued')) pushIfNonZero(lbl(options, 'eq', 'shareCapitalIssued', 'Share Capital Issued'), equityData.map(d => ovVal(options, 'eq', 'shareCapitalIssued', d.year, d.shareCapitalIssued)));
  rows.push([lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'), ...equityData.map(d => ovVal(options, 'eq', 'closingShareCapital', d.year, d.closingShareCapital))]);
  if (isVis(options, 'eq_openingRetainedEarnings')) pushIfNonZero(lbl(options, 'eq', 'openingRetainedEarnings', 'Opening Retained Earnings'), equityData.map(d => ovVal(options, 'eq', 'openingRetainedEarnings', d.year, d.openingRetainedEarnings)));
  if (isVis(options, 'eq_netProfitForYear')) pushIfNonZero(lbl(options, 'eq', 'netProfitForYear', 'Net Profit for the Year'), equityData.map(d => ovVal(options, 'eq', 'netProfitForYear', d.year, d.netProfitForYear)));
  if (isVis(options, 'eq_dividendsPaid')) pushIfNonZero(lbl(options, 'eq', 'dividendsPaid', 'Dividends Paid'), equityData.map(d => -ovVal(options, 'eq', 'dividendsPaid', d.year, d.dividendsPaid)));
  rows.push([lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'), ...equityData.map(d => ovVal(options, 'eq', 'closingRetainedEarnings', d.year, d.closingRetainedEarnings))]);
  const openResVals = equityData.map(d => ovVal(options, 'eq', 'openingReserves', d.year, d.openingReserves));
  const ociVals = equityData.map(d => ovVal(options, 'eq', 'otherComprehensiveIncome', d.year, d.otherComprehensiveIncome));
  const closeResVals = equityData.map(d => ovVal(options, 'eq', 'closingReserves', d.year, d.closingReserves));
  const showOpenRes = isVis(options, 'eq_openingReserves') && hasAnyNonZeroValue(openResVals);
  const showOci = isVis(options, 'eq_otherComprehensiveIncome') && hasAnyNonZeroValue(ociVals);
  if (showOpenRes || showOci || hasAnyNonZeroValue(closeResVals)) {
    if (showOpenRes) rows.push([lbl(options, 'eq', 'openingReserves', 'Opening Reserves'), ...openResVals]);
    if (showOci) rows.push([lbl(options, 'eq', 'otherComprehensiveIncome', 'Other Comprehensive Income'), ...ociVals]);
    rows.push([lbl(options, 'eq', 'closingReserves', 'Closing Reserves'), ...closeResVals]);
  }
  getCustomRows(options, 'equity').forEach(ci => pushIfNonZero(ci.label, equityData.map(d => getCustomItemYearValue(ci, d.year))));
  rows.push([lbl(options, 'eq', 'totalOpeningEquity', 'TOTAL OPENING EQUITY'), ...equityData.map(d => ovVal(options, 'eq', 'totalOpeningEquity', d.year, d.totalOpeningEquity))]);
  rows.push([lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'), ...equityData.map(d => ovVal(options, 'eq', 'totalClosingEquity', d.year, d.totalClosingEquity))]);
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
  const faHeaders = ['Asset Class', 'Dep %', 'Opening Cost', 'Additions', 'Disposals', 'Closing Cost', 'Opening Dep.', 'Dep. Charge', 'Dep. Disposals', 'Closing Dep.', 'Net Book Value'];
  const faCats = latestFA ? resolveFACategoriesForExport(latestFA, options) : [];
  const faRows: (string | number)[][] = faCats.map(c => [
    c.name, `${c.depRate}%`, c.costOpening, c.additions, c.disposals, c.costClosing, c.depOpening, c.depCharge, c.depDisposals, c.depClosing, c.netBookValue
  ]);
  if (latestFA) {
    faRows.push(['TOTAL', '-', latestFA.totalCostOpening, latestFA.totalAdditions, latestFA.totalDisposals, latestFA.totalCostClosing, latestFA.totalDepOpening, latestFA.totalDepCharge, latestFA.totalDepDisposals, latestFA.totalDepClosing, latestFA.totalNetBookValue]);
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
          ...notes.flatMap(n => {
            const filteredTbl = filterNoteTableForExport(n.tableData);
            return [
              new Paragraph({ text: `Note ${n.noteNumber}: ${n.title}`, heading: HeadingLevel.HEADING_3 }),
              new Paragraph({ text: n.textDescription }),
              ...(filteredTbl ? [createDocxTable(filteredTbl.headers, [
                ...filteredTbl.rows.map(r => [r.label, ...r.values]),
                ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
              ])] : []),
              new Paragraph({ text: '' }),
            ];
          }),

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

export async function exportNotesWord(
  company: CompanyProfile,
  notes: FinancialNoteItem[],
  filename = 'Notes_to_the_Financial_Statements.docx'
) {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: company.companyName, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: 'NOTES TO THE FINANCIAL STATEMENTS (Notes 1 to 12)', heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
          new Paragraph({
            text: `Reporting Currency: ${company.reportingCurrency} (${company.reportingUnit}) | Framework: ${company.accountingFramework}`,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: '' }),
          ...notes.flatMap(n => {
            const filteredTbl = filterNoteTableForExport(n.tableData);
            return [
              new Paragraph({ text: `Note ${n.noteNumber}: ${n.title} [${n.category}]`, heading: HeadingLevel.HEADING_3 }),
              new Paragraph({ text: n.textDescription }),
              ...(filteredTbl
                ? [
                    createDocxTable(filteredTbl.headers, [
                      ...filteredTbl.rows.map(r => [r.label, ...r.values]),
                      ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
                    ]),
                  ]
                : []),
              new Paragraph({ text: '' }),
            ];
          }),
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

export function exportNotesPdf(
  company: CompanyProfile,
  notes: FinancialNoteItem[],
  filename = 'Notes_to_the_Financial_Statements.pdf'
) {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  renderPdfHeader(doc, company, 'NOTES TO THE FINANCIAL STATEMENTS (STATUTORY DISCLOSURES 1 TO 12)');
  let yPos = 30;
  notes.forEach((note) => {
    if (yPos > 170) {
      doc.addPage();
      renderPdfHeader(doc, company, 'NOTES TO THE FINANCIAL STATEMENTS (CONTINUED)');
      yPos = 30;
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9.5);
    doc.setTextColor(30, 41, 59);
    doc.text(`Note ${note.noteNumber}: ${note.title} [${note.category}]`, 14, yPos);
    yPos += 5;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(71, 85, 105);
    const splitText = doc.splitTextToSize(note.textDescription, 269);
    doc.text(splitText, 14, yPos);
    yPos += splitText.length * 3.8 + 2.5;

    const filteredTbl = filterNoteTableForExport(note.tableData);
    if (filteredTbl) {
      if (yPos > 165) {
        doc.addPage();
        renderPdfHeader(doc, company, 'NOTES TO THE FINANCIAL STATEMENTS (CONTINUED)');
        yPos = 30;
      }
      const tableBody = [
        ...filteredTbl.rows.map(r => [r.label, ...r.values]),
        ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
      ];
      autoTable(doc, {
        head: [filteredTbl.headers],
        body: tableBody,
        startY: yPos,
        styles: { fontSize: 7.5, cellPadding: 1.5, font: 'helvetica' },
        headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [248, 250, 252] },
        margin: { left: 14, right: 14 },
      });
      const finalY = (doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? yPos + 14;
      yPos = finalY + 5;
    } else {
      yPos += 2;
    }
  });

  const totalPages = doc.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(148, 163, 184);
    doc.text(`Financials with Tuhin · Page ${i} of ${totalPages}`, 14, 210 - 8);
    doc.text(`Certified System Export · Date: ${new Date().toISOString().split('T')[0]}`, 297 - 14, 210 - 8, { align: 'right' });
  }

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
  const faHeaders = ['Category', 'Dep %', 'Opening Cost', 'Additions', 'Disposals', 'Closing Cost', 'Opening Dep.', 'Charge', 'Dep. Disposals', 'Closing Dep.', 'Net Book Value'];
  const faCats = latestFA ? resolveFACategoriesForExport(latestFA, options) : [];
  const faRows: (string | number)[][] = faCats.map(c => [
    c.name, `${c.depRate}%`, c.costOpening, c.additions, c.disposals, c.costClosing, c.depOpening, c.depCharge, c.depDisposals, c.depClosing, c.netBookValue
  ]);
  if (latestFA) {
    faRows.push(['TOTAL PPE', '-', latestFA.totalCostOpening, latestFA.totalAdditions, latestFA.totalDisposals, latestFA.totalCostClosing, latestFA.totalDepOpening, latestFA.totalDepCharge, latestFA.totalDepDisposals, latestFA.totalDepClosing, latestFA.totalNetBookValue]);
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
    if (yPos > 170) {
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
    yPos += splitText.length * 3.8 + 2.5;

    const filteredTbl = filterNoteTableForExport(note.tableData);
    if (filteredTbl) {
      if (yPos > 165) {
        doc.addPage();
        renderPdfHeader(doc, company, '6. NOTES TO THE FINANCIAL STATEMENTS (CONTINUED)');
        yPos = 30;
      }
      const tableBody = [
        ...filteredTbl.rows.map(r => [r.label, ...r.values]),
        ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
      ];
      autoTable(doc, {
        head: [filteredTbl.headers],
        body: tableBody,
        startY: yPos,
        styles: { fontSize: 7.5, cellPadding: 1.5, font: 'helvetica' },
        headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [248, 250, 252] },
        margin: { left: 14, right: 14 },
      });
      const finalY = (doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? yPos + 14;
      yPos = finalY + 5;
    } else {
      yPos += 2;
    }
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

export function buildPreviousTextTables(
  prevData: PreviousFormatYearData[],
  options?: ExportSyncOptions
) {
  const pushIfNonZero = (target: (string | number)[][], label: string, vals: number[]) => {
    if (hasAnyNonZeroValue(vals)) target.push([label, ...vals]);
  };

  const customBsCa = getCustomRows(options, 'balanceSheet', 'Current Assets');
  const customBsNca = getCustomRows(options, 'balanceSheet', 'Non-Current Assets');
  const customBsCl = getCustomRows(options, 'balanceSheet', 'Current Liabilities');
  const customBsNcl = getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities');
  const customBsEq = getCustomRows(options, 'balanceSheet', 'Equity');

  const customIncRev = getCustomRows(options, 'income', 'Revenue & Gross Margin');
  const customIncOpex = getCustomRows(options, 'income', 'Operating Expenses');
  const customIncNonOp = [...getCustomRows(options, 'income', 'Non-Operating & Tax'), ...getUnscopedCustomIncomeRows(options)];

  const customCfOp = getCustomRows(options, 'cashFlow', 'Operating Activities');
  const customCfInv = getCustomRows(options, 'cashFlow', 'Investing Activities');
  const customCfFin = getCustomRows(options, 'cashFlow', 'Financing Activities');
  const customEq = getCustomRows(options, 'equity');

  const bsRows: (string | number)[][] = [
    ['PROPERTIES & ASSETS', ...prevData.map(() => '')],
  ];
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'propertyPlantEquipment', 'Fixed Assets (at cost less depreciation)'), prevData.map(d => d.fixedAssetsWDV));
  }
  customBsNca.forEach(ci => pushIfNonZero(bsRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  bsRows.push(['Current Assets', ...prevData.map(d => d.totalCurrentAssets)]);
  if (isVis(options, 'bs_otherReceivables')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'otherReceivables', 'Advances, Deposits & Pre-payments'), prevData.map(d => d.advancesDepositsPrepayments));
  }
  if (isVis(options, 'bs_closingInventory')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'closingInventory', 'Stock & Stores'), prevData.map(d => d.stockAndStores));
  }
  if (isVis(options, 'bs_tradeReceivables')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'tradeReceivables', 'Bills Receivable'), prevData.map(d => d.billsReceivable));
  }
  if (isVis(options, 'bs_cashAndEquivalents')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'cashAndEquivalents', 'Cash & Cash Equivalents'), prevData.map(d => d.cashAndCashEquivalents));
  }
  customBsCa.forEach(ci => pushIfNonZero(bsRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  bsRows.push(['Taka', ...prevData.map(d => d.totalAssetsTaka)]);
  bsRows.push(['CAPITAL & LIABILITIES', ...prevData.map(() => '')]);
  pushIfNonZero(bsRows, 'Opening Capital', prevData.map(d => d.openingCapital));
  pushIfNonZero(bsRows, 'Add. Net Profit for the year', prevData.map(d => d.netProfitForYear));
  bsRows.push(['Subtotal', ...prevData.map(d => d.capitalSubtotal)]);
  if (isVis(options, 'eq_dividendsPaid')) {
    pushIfNonZero(bsRows, `Less: ${lbl(options, 'eq', 'dividendsPaid', 'Drawings')}`, prevData.map(d => d.drawings));
  }
  bsRows.push(['Capital after drawings', ...prevData.map(d => d.capitalAfterDrawings)]);
  customBsEq.forEach(ci => pushIfNonZero(bsRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  customBsNcl.forEach(ci => pushIfNonZero(bsRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  bsRows.push(['Current Liabilities', ...prevData.map(d => d.totalCurrentLiabilities)]);
  if (isVis(options, 'bs_otherPayables')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'otherPayables', 'For Expenses'), prevData.map(d => d.forExpenses));
  }
  if (isVis(options, 'bs_tradePayables')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'tradePayables', 'Accounts Payable'), prevData.map(d => d.accountsPayable));
  }
  if (isVis(options, 'bs_shortTermBorrowings')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'shortTermBorrowings', 'Short Term Loan'), prevData.map(d => d.shortTermLoan));
  }
  if (isVis(options, 'bs_currentTaxPayable')) {
    pushIfNonZero(bsRows, lbl(options, 'bs', 'currentTaxPayable', 'Provision for Taxation'), prevData.map(d => d.provisionForTaxation));
  }
  customBsCl.forEach(ci => pushIfNonZero(bsRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  bsRows.push(['Total', ...prevData.map(d => d.totalCapitalAndLiabilities)]);

  const isRows: (string | number)[][] = [];
  if (isVis(options, 'inc_revenue')) {
    pushIfNonZero(
      isRows,
      lbl(options, 'income', 'revenue', 'Sales'),
      prevData.map(d => d.sales - customIncRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))
    );
  }
  customIncRev.forEach(ci => pushIfNonZero(isRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  if (isVis(options, 'inc_costOfSales')) {
    pushIfNonZero(isRows, `Less- ${lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold')}`, prevData.map(d => d.costOfGoodsSold));
  }
  isRows.push(['Gross Profit', ...prevData.map(d => d.grossProfit)]);
  isRows.push(['Less- Administrative & Selling Expenses', ...prevData.map(d => d.adminAndSellingTotal)]);
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
      pushIfNonZero(isRows, lineLabel, prevData.map(d => Number(d[field])));
    });
  }
  customIncOpex.forEach(ci => pushIfNonZero(isRows, ci.label, prevData.map(d => -getCustomItemYearValue(ci, d.year))));
  if (isVis(options, 'inc_depreciationExpense')) {
    pushIfNonZero(isRows, lbl(options, 'income', 'depreciationExpense', 'Depreciation'), prevData.map(d => d.depreciation));
  }
  isRows.push(['Net Operating Profit', ...prevData.map(d => d.netOperatingProfit)]);
  if (isVis(options, 'inc_interestIncome')) {
    pushIfNonZero(
      isRows,
      `Add: ${lbl(options, 'income', 'interestIncome', 'Other Income')}`,
      prevData.map(d => d.otherIncome - customIncNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))
    );
  }
  customIncNonOp.forEach(ci => pushIfNonZero(isRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  isRows.push(['Subtotal', ...prevData.map(d => d.operatingSubtotal)]);
  if (isVis(options, 'inc_financeCost')) {
    pushIfNonZero(isRows, `Less: ${lbl(options, 'income', 'financeCost', 'Financial Exp.')}`, prevData.map(d => d.financialExp));
  }
  if (isVis(options, 'inc_taxExpense')) {
    pushIfNonZero(isRows, `Less- ${lbl(options, 'income', 'taxExpense', 'Income Tax Provision')}`, prevData.map(d => d.incomeTaxProvision));
  }
  isRows.push(['Net Profit After Income Tax', ...prevData.map(d => d.netProfitAfterIncomeTax)]);

  const cfRows: (string | number)[][] = [
    ['a) CASH FLOWS FROM OPERATING ACTIVITIES:', ...prevData.map(() => '')],
  ];
  if (isVis(options, 'cf_pbt')) {
    pushIfNonZero(cfRows, lbl(options, 'cf', 'profitBeforeTax', 'Net Income'), prevData.map(d => d.cfNetIncome));
  }
  if (isVis(options, 'cf_dep')) {
    pushIfNonZero(cfRows, lbl(options, 'cf', 'adjDepreciation', 'Add: Depreciation of Tangible Fixed Assets'), prevData.map(d => d.cfDepreciation));
  }
  if (isVis(options, 'cf_wc_receivables')) {
    pushIfNonZero(cfRows, lbl(options, 'cf', 'changeInReceivables', 'Add: (Increase) or decrease in Current Assets'), prevData.map(d => d.cfChangeInCurrentAssets));
  }
  pushIfNonZero(cfRows, 'Add: (Increase) or decrease in Construction Working Progress', prevData.map(d => d.cfChangeInCWIP));
  if (isVis(options, 'cf_wc_payables')) {
    pushIfNonZero(cfRows, lbl(options, 'cf', 'changeInPayables', 'Add: Increase or (decrease) in Current Liabilities'), prevData.map(d => d.cfChangeInCurrentLiabilities));
  }
  customCfOp.forEach(ci => pushIfNonZero(cfRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  cfRows.push(['Net Cash from operating activities', ...prevData.map(d => d.cfNetCashOperating)]);

  cfRows.push(['b) CASH FLOWS FROM INVESTING ACTIVITIES:', ...prevData.map(() => '')]);
  if (isVis(options, 'cf_pur_ppe')) {
    pushIfNonZero(cfRows, lbl(options, 'cf', 'purchasePPE', 'Acquisition of Tangible Fixed Assets'), prevData.map(d => d.cfAcquisitionPPE));
  }
  if (isVis(options, 'cf_pur_intangible')) {
    pushIfNonZero(cfRows, lbl(options, 'cf', 'purchaseIntangibles', 'Disposal of Intangible Assets'), prevData.map(d => d.cfDisposalIntangible));
  }
  customCfInv.forEach(ci => pushIfNonZero(cfRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  cfRows.push(['Net Cash used in investing Activities', ...prevData.map(d => d.cfNetCashInvesting)]);

  cfRows.push(['c) CASH FLOWS FROM FINANCING ACTIVITIES:', ...prevData.map(() => '')]);
  if (isVis(options, 'cf_dividends_paid')) {
    pushIfNonZero(cfRows, lbl(options, 'cf', 'financingDividendsPaid', 'Drawings'), prevData.map(d => d.cfDrawings));
  }
  customCfFin.forEach(ci => pushIfNonZero(cfRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  pushIfNonZero(cfRows, 'Opening Balance', prevData.map(d => d.cfOpeningBalance));
  cfRows.push(['Net Cash used in Financing Activities', ...prevData.map(d => d.cfNetCashFinancing)]);
  cfRows.push(['d) Net Increase decrease in cash & cash equivalents (A+B+C)', ...prevData.map(d => d.cfNetIncreaseCash)]);
  cfRows.push(['e) Cash & Cash Equivalents at beginning of period', ...prevData.map(d => d.cfCashBeginning)]);
  cfRows.push(['f) Cash & Cash Equivalents at end of period (D+E)', ...prevData.map(d => d.cfCashEnding)]);

  const eqRows: (string | number)[][] = [];
  if (isVis(options, 'eq_openingShareCapital')) {
    pushIfNonZero(eqRows, lbl(options, 'eq', 'openingShareCapital', 'Opening Share Capital'), prevData.map(d => d.eqOpeningShareCapital));
  }
  if (isVis(options, 'eq_shareCapitalIssued')) {
    pushIfNonZero(eqRows, lbl(options, 'eq', 'shareCapitalIssued', 'Share Capital Issued'), prevData.map(d => d.eqShareCapitalIssued));
  }
  eqRows.push([lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'), ...prevData.map(d => d.eqClosingShareCapital)]);
  if (isVis(options, 'eq_openingRetainedEarnings')) {
    pushIfNonZero(eqRows, lbl(options, 'eq', 'openingRetainedEarnings', 'Opening Retained Earnings'), prevData.map(d => d.eqOpeningRetainedEarnings));
  }
  if (isVis(options, 'eq_netProfitForYear')) {
    pushIfNonZero(eqRows, lbl(options, 'eq', 'netProfitForYear', 'Net Profit for the Year'), prevData.map(d => d.eqNetProfitForYear));
  }
  if (isVis(options, 'eq_dividendsPaid')) {
    pushIfNonZero(eqRows, lbl(options, 'eq', 'dividendsPaid', 'Dividends Paid'), prevData.map(d => -d.eqDividendsPaid));
  }
  eqRows.push([lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'), ...prevData.map(d => d.eqClosingRetainedEarnings)]);
  if (isVis(options, 'eq_openingReserves') || isVis(options, 'eq_otherComprehensiveIncome')) {
    pushIfNonZero(eqRows, lbl(options, 'eq', 'closingReserves', 'Closing Reserves'), prevData.map(d => d.eqClosingReserves));
  }
  customEq.forEach(ci => pushIfNonZero(eqRows, ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year))));
  eqRows.push([lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'), ...prevData.map(d => d.eqTotalClosingEquity)]);

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

  const customBsCa = getCustomRows(options, 'balanceSheet', 'Current Assets');
  const customBsNca = getCustomRows(options, 'balanceSheet', 'Non-Current Assets');
  const customBsCl = getCustomRows(options, 'balanceSheet', 'Current Liabilities');
  const customBsNcl = getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities');
  const customBsEq = getCustomRows(options, 'balanceSheet', 'Equity');

  const customIncRev = getCustomRows(options, 'income', 'Revenue & Gross Margin');
  const customIncOpex = getCustomRows(options, 'income', 'Operating Expenses');
  const customIncNonOp = [...getCustomRows(options, 'income', 'Non-Operating & Tax'), ...getUnscopedCustomIncomeRows(options)];

  const customCfOp = getCustomRows(options, 'cashFlow', 'Operating Activities');
  const customCfInv = getCustomRows(options, 'cashFlow', 'Investing Activities');
  const customCfFin = getCustomRows(options, 'cashFlow', 'Financing Activities');
  const customEq = getCustomRows(options, 'equity');

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
  const pushBsIfNonZero = (label: string, vals: number[], trackRows: number[]) => {
    if (hasAnyNonZeroValue(vals)) {
      trackRows.push(nextBsRow());
      bsRows.push([label, ...vals.map(v => nCell(v))]);
    }
  };
  const assetInputs: number[] = [];
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    pushBsIfNonZero(lbl(options, 'bs', 'propertyPlantEquipment', 'Fixed Assets (at cost less depreciation)'), prevData.map(d => d.fixedAssetsWDV), assetInputs);
  }
  customBsNca.forEach(ci => {
    pushBsIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), assetInputs);
  });
  const caHeaderIdx = bsRows.length;
  const caExcelRow = caHeaderIdx + 1;
  bsRows.push(['Current Assets', ...prevData.map(d => nCell(d.totalCurrentAssets))]);
  const caChildRows: number[] = [];
  if (isVis(options, 'bs_otherReceivables')) {
    pushBsIfNonZero(lbl(options, 'bs', 'otherReceivables', 'Advances, Deposits & Pre-payments'), prevData.map(d => d.advancesDepositsPrepayments), caChildRows);
  }
  if (isVis(options, 'bs_closingInventory')) {
    pushBsIfNonZero(lbl(options, 'bs', 'closingInventory', 'Stock & Stores'), prevData.map(d => d.stockAndStores), caChildRows);
  }
  if (isVis(options, 'bs_tradeReceivables')) {
    pushBsIfNonZero(lbl(options, 'bs', 'tradeReceivables', 'Bills Receivable'), prevData.map(d => d.billsReceivable), caChildRows);
  }
  if (isVis(options, 'bs_cashAndEquivalents')) {
    pushBsIfNonZero(lbl(options, 'bs', 'cashAndEquivalents', 'Cash & Cash Equivalents'), prevData.map(d => d.cashAndCashEquivalents), caChildRows);
  }
  customBsCa.forEach(ci => {
    pushBsIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), caChildRows);
  });
  bsRows[caHeaderIdx] = [
    'Current Assets',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), caChildRows), d.totalCurrentAssets)),
  ];
  assetInputs.push(caExcelRow);
  bsRows.push([
    'Taka',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), assetInputs), d.totalAssetsTaka)),
  ]);

  bsRows.push(['CAPITAL & LIABILITIES', ...prevData.map(() => '')]);
  const capSubInputs: number[] = [];
  pushBsIfNonZero('Opening Capital', prevData.map(d => d.openingCapital), capSubInputs);
  pushBsIfNonZero('Add. Net Profit for the year', prevData.map(d => d.netProfitForYear), capSubInputs);
  const capSubRow = nextBsRow();
  bsRows.push([
    'Subtotal',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), capSubInputs), d.capitalSubtotal)),
  ]);
  let drawingsBsRow: number | null = null;
  if (isVis(options, 'eq_dividendsPaid')) {
    const drawVals = prevData.map(d => d.drawings);
    if (hasAnyNonZeroValue(drawVals)) {
      drawingsBsRow = nextBsRow();
      bsRows.push([`Less: ${lbl(options, 'eq', 'dividendsPaid', 'Drawings')}`, ...drawVals.map(v => nCell(v))]);
    }
  }
  const capAfterDrawRow = nextBsRow();
  bsRows.push([
    'Capital after drawings',
    ...prevData.map((d, i) =>
      fCell(excelDiffExpr(colLetter(i + 1), [capSubRow], drawingsBsRow !== null ? [drawingsBsRow] : []), d.capitalAfterDrawings)
    ),
  ]);
  const totalCapLiabInputs = [capAfterDrawRow];
  customBsEq.forEach(ci => {
    pushBsIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), totalCapLiabInputs);
  });
  customBsNcl.forEach(ci => {
    pushBsIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), totalCapLiabInputs);
  });
  const clHeaderIdx = bsRows.length;
  const clExcelRow = clHeaderIdx + 1;
  bsRows.push(['Current Liabilities', ...prevData.map(d => nCell(d.totalCurrentLiabilities))]);
  const clChildRows: number[] = [];
  if (isVis(options, 'bs_otherPayables')) {
    pushBsIfNonZero(lbl(options, 'bs', 'otherPayables', 'For Expenses'), prevData.map(d => d.forExpenses), clChildRows);
  }
  if (isVis(options, 'bs_tradePayables')) {
    pushBsIfNonZero(lbl(options, 'bs', 'tradePayables', 'Accounts Payable'), prevData.map(d => d.accountsPayable), clChildRows);
  }
  if (isVis(options, 'bs_shortTermBorrowings')) {
    pushBsIfNonZero(lbl(options, 'bs', 'shortTermBorrowings', 'Short Term Loan'), prevData.map(d => d.shortTermLoan), clChildRows);
  }
  if (isVis(options, 'bs_currentTaxPayable')) {
    pushBsIfNonZero(lbl(options, 'bs', 'currentTaxPayable', 'Provision for Taxation'), prevData.map(d => d.provisionForTaxation), clChildRows);
  }
  customBsCl.forEach(ci => {
    pushBsIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), clChildRows);
  });
  bsRows[clHeaderIdx] = [
    'Current Liabilities',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), clChildRows), d.totalCurrentLiabilities)),
  ];
  totalCapLiabInputs.push(clExcelRow);
  bsRows.push([
    'Total',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), totalCapLiabInputs), d.totalCapitalAndLiabilities)),
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
  const pushIsIfNonZero = (label: string, vals: number[], trackRows: number[]) => {
    if (hasAnyNonZeroValue(vals)) {
      trackRows.push(nextIsRow());
      isRows.push([label, ...vals.map(v => nCell(v))]);
    }
  };
  const gpAddRows: number[] = [];
  if (isVis(options, 'inc_revenue')) {
    pushIsIfNonZero(
      lbl(options, 'income', 'revenue', 'Sales'),
      prevData.map(d => d.sales - customIncRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)),
      gpAddRows
    );
  }
  customIncRev.forEach(ci => {
    pushIsIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), gpAddRows);
  });
  let cogsExcelRow: number | null = null;
  if (isVis(options, 'inc_costOfSales')) {
    const cogsVals = prevData.map(d => d.costOfGoodsSold);
    if (hasAnyNonZeroValue(cogsVals)) {
      cogsExcelRow = nextIsRow();
      isRows.push([`Less- ${lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold')}`, ...cogsVals.map(v => nCell(v))]);
    }
  }
  const gpExcelRow = nextIsRow();
  isRows.push([
    'Gross Profit',
    ...prevData.map((d, i) =>
      fCell(
        excelDiffExpr(
          colLetter(i + 1),
          gpAddRows,
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
      pushIsIfNonZero(lineLabel, prevData.map(d => Number(d[field])), adminChildRows);
    });
  }
  customIncOpex.forEach(ci => {
    pushIsIfNonZero(ci.label, prevData.map(d => -getCustomItemYearValue(ci, d.year)), adminChildRows);
  });
  if (isVis(options, 'inc_depreciationExpense')) {
    pushIsIfNonZero(lbl(options, 'income', 'depreciationExpense', 'Depreciation'), prevData.map(d => d.depreciation), adminChildRows);
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
  const opSubAddRows: number[] = [netOpProfitRow];
  if (isVis(options, 'inc_interestIncome')) {
    pushIsIfNonZero(
      `Add: ${lbl(options, 'income', 'interestIncome', 'Other Income')}`,
      prevData.map(d => d.otherIncome - customIncNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)),
      opSubAddRows
    );
  }
  customIncNonOp.forEach(ci => {
    pushIsIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), opSubAddRows);
  });
  const opSubRow = nextIsRow();
  isRows.push([
    'Subtotal',
    ...prevData.map((d, i) =>
      fCell(
        excelSumRows(colLetter(i + 1), opSubAddRows),
        d.operatingSubtotal
      )
    ),
  ]);
  const taxAndFinSubRows: number[] = [];
  if (isVis(options, 'inc_financeCost')) {
    pushIsIfNonZero(`Less: ${lbl(options, 'income', 'financeCost', 'Financial Exp.')}`, prevData.map(d => d.financialExp), taxAndFinSubRows);
  }
  if (isVis(options, 'inc_taxExpense')) {
    pushIsIfNonZero(`Less- ${lbl(options, 'income', 'taxExpense', 'Income Tax Provision')}`, prevData.map(d => d.incomeTaxProvision), taxAndFinSubRows);
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
  const pushCfIfNonZero = (label: string, vals: number[], trackRows: number[]) => {
    if (hasAnyNonZeroValue(vals)) {
      trackRows.push(nextCfRow());
      cfRows.push([label, ...vals.map(v => nCell(v))]);
    }
  };
  const cfOpChildRows: number[] = [];
  if (isVis(options, 'cf_pbt')) {
    pushCfIfNonZero(lbl(options, 'cf', 'profitBeforeTax', 'Net Income'), prevData.map(d => d.cfNetIncome), cfOpChildRows);
  }
  if (isVis(options, 'cf_dep')) {
    pushCfIfNonZero(lbl(options, 'cf', 'adjDepreciation', 'Add: Depreciation of Tangible Fixed Assets'), prevData.map(d => d.cfDepreciation), cfOpChildRows);
  }
  if (isVis(options, 'cf_wc_receivables')) {
    pushCfIfNonZero(lbl(options, 'cf', 'changeInReceivables', 'Add: (Increase) or decrease in Current Assets'), prevData.map(d => d.cfChangeInCurrentAssets), cfOpChildRows);
  }
  pushCfIfNonZero('Add: (Increase) or decrease in Construction Working Progress', prevData.map(d => d.cfChangeInCWIP), cfOpChildRows);
  if (isVis(options, 'cf_wc_payables')) {
    pushCfIfNonZero(lbl(options, 'cf', 'changeInPayables', 'Add: Increase or (decrease) in Current Liabilities'), prevData.map(d => d.cfChangeInCurrentLiabilities), cfOpChildRows);
  }
  customCfOp.forEach(ci => {
    pushCfIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), cfOpChildRows);
  });
  const cfNetOpRow = nextCfRow();
  cfRows.push([
    'Net Cash from operating activities',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfOpChildRows), d.cfNetCashOperating)),
  ]);

  cfRows.push(['b) CASH FLOWS FROM INVESTING ACTIVITIES:', ...prevData.map(() => '')]);
  const cfInvChildRows: number[] = [];
  if (isVis(options, 'cf_pur_ppe')) {
    pushCfIfNonZero(lbl(options, 'cf', 'purchasePPE', 'Acquisition of Tangible Fixed Assets'), prevData.map(d => d.cfAcquisitionPPE), cfInvChildRows);
  }
  if (isVis(options, 'cf_pur_intangible')) {
    pushCfIfNonZero(lbl(options, 'cf', 'purchaseIntangibles', 'Disposal of Intangible Assets'), prevData.map(d => d.cfDisposalIntangible), cfInvChildRows);
  }
  customCfInv.forEach(ci => {
    pushCfIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), cfInvChildRows);
  });
  const cfNetInvRow = nextCfRow();
  cfRows.push([
    'Net Cash used in investing Activities',
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfInvChildRows), d.cfNetCashInvesting)),
  ]);

  cfRows.push(['c) CASH FLOWS FROM FINANCING ACTIVITIES:', ...prevData.map(() => '')]);
  const cfFinChildRows: number[] = [];
  if (isVis(options, 'cf_dividends_paid')) {
    pushCfIfNonZero(lbl(options, 'cf', 'financingDividendsPaid', 'Drawings'), prevData.map(d => d.cfDrawings), cfFinChildRows);
  }
  customCfFin.forEach(ci => {
    pushCfIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), cfFinChildRows);
  });
  pushCfIfNonZero('Opening Balance', prevData.map(d => d.cfOpeningBalance), cfFinChildRows);
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
  const pushEqIfNonZero = (label: string, vals: number[], trackRows: number[]) => {
    if (hasAnyNonZeroValue(vals)) {
      trackRows.push(nextEqRow());
      eqRows.push([label, ...vals.map(v => nCell(v))]);
    }
  };
  const scInputs: number[] = [];
  if (isVis(options, 'eq_openingShareCapital')) {
    pushEqIfNonZero(lbl(options, 'eq', 'openingShareCapital', 'Opening Share Capital'), prevData.map(d => d.eqOpeningShareCapital), scInputs);
  }
  if (isVis(options, 'eq_shareCapitalIssued')) {
    pushEqIfNonZero(lbl(options, 'eq', 'shareCapitalIssued', 'Share Capital Issued'), prevData.map(d => d.eqShareCapitalIssued), scInputs);
  }
  const closeScRow = nextEqRow();
  eqRows.push([
    lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'),
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), scInputs), d.eqClosingShareCapital)),
  ]);
  const reInputs: number[] = [];
  if (isVis(options, 'eq_openingRetainedEarnings')) {
    pushEqIfNonZero(lbl(options, 'eq', 'openingRetainedEarnings', 'Opening Retained Earnings'), prevData.map(d => d.eqOpeningRetainedEarnings), reInputs);
  }
  if (isVis(options, 'eq_netProfitForYear')) {
    pushEqIfNonZero(lbl(options, 'eq', 'netProfitForYear', 'Net Profit for the Year'), prevData.map(d => d.eqNetProfitForYear), reInputs);
  }
  if (isVis(options, 'eq_dividendsPaid')) {
    pushEqIfNonZero(lbl(options, 'eq', 'dividendsPaid', 'Dividends Paid'), prevData.map(d => -d.eqDividendsPaid), reInputs);
  }
  const closeReRow = nextEqRow();
  eqRows.push([
    lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'),
    ...prevData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), reInputs), d.eqClosingRetainedEarnings)),
  ]);
  const eqTotalInputs = [closeScRow, closeReRow];
  if (isVis(options, 'eq_openingReserves') || isVis(options, 'eq_otherComprehensiveIncome')) {
    pushEqIfNonZero(lbl(options, 'eq', 'closingReserves', 'Closing Reserves'), prevData.map(d => d.eqClosingReserves), eqTotalInputs);
  }
  customEq.forEach(ci => {
    pushEqIfNonZero(ci.label, prevData.map(d => getCustomItemYearValue(ci, d.year)), eqTotalInputs);
  });
  eqRows.push([
    lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'),
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
    const filteredTbl = filterNoteTableForExport(note.tableData);
    if (filteredTbl) {
      faRows.push(filteredTbl.headers);
      filteredTbl.rows.forEach(r => faRows.push([r.label, ...r.values]));
      if (filteredTbl.totalRow) {
        faRows.push([filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]);
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
          ...notes.flatMap(n => {
            const filteredTbl = filterNoteTableForExport(n.tableData);
            return [
              new Paragraph({ children: [new TextRun({ text: `Note ${n.noteNumber}: ${n.title}`, bold: true, size: 18 })] }),
              new Paragraph({ children: [new TextRun({ text: n.textDescription, size: 17 })] }),
              ...(filteredTbl
                ? [
                    createDocxTable(filteredTbl.headers, [
                      ...filteredTbl.rows.map(r => [r.label, ...r.values]),
                      ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
                    ]),
                  ]
                : []),
              new Paragraph({ text: '' }),
            ];
          }),
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
    'f) Cash & Cash Equivalents at end of period (D+E)',
    lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'),
    lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'),
    lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'),
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

    const filteredTbl = filterNoteTableForExport(note.tableData);
    if (filteredTbl) {
      if (yCursor > pageH - 40) {
        doc.addPage();
        yCursor = 18;
      }
      const nCols = filteredTbl.headers.length;
      const bodyRows = [
        ...filteredTbl.rows.map(r => [r.label, ...r.values.map(v => String(v))]),
        ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values.map(v => String(v))]] : []),
      ];
      autoTable(doc, {
        head: [filteredTbl.headers],
        body: bodyRows,
        startY: yCursor,
        theme: 'grid',
        styles: { font: 'times', fontSize: 8, cellPadding: 1.4, textColor: [15, 23, 42], lineColor: [148, 163, 184], lineWidth: 0.15 },
        headStyles: { fillColor: [248, 250, 252], textColor: [15, 23, 42], fontStyle: 'bold', halign: 'left' },
        columnStyles: Object.fromEntries(
          Array.from({ length: nCols }, (_, idx) => [idx, { halign: idx === 0 ? 'left' : 'right' }])
        ),
        didParseCell: (data) => {
          if (data.section === 'body' && filteredTbl.totalRow && data.row.index === bodyRows.length - 1) {
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

export function buildProjectedTextTables(
  projData: ProjectedFormatYearData[],
  options?: ExportSyncOptions
) {
  const salesLbl = lbl(options, 'income', 'revenue', 'Sales');
  const cogsLbl = `Less: ${lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold')}`;
  const opexLbl = lbl(options, 'income', 'operatingExpenses', 'Administrative & Selling Expenses');
  const trRecLbl = lbl(options, 'bs', 'tradeReceivables', 'Trade & Other Receivables');
  const cashLbl = lbl(options, 'bs', 'cashAndEquivalents', 'Cash & Cash Equivalents');
  const invLbl = lbl(options, 'bs', 'closingInventory', 'Inventories');
  const advLbl = lbl(options, 'bs', 'otherReceivables', 'Advance, Deposit & Prepayments');
  const ppeLbl = lbl(options, 'bs', 'propertyPlantEquipment', 'Property Plant & Equiptments');
  const ltLoanLbl = lbl(options, 'bs', 'longTermBorrowings', 'Long Term Loan');
  const trPayLbl = lbl(options, 'bs', 'tradePayables', 'Trade & Other Payables');
  const taxPayLbl = lbl(options, 'bs', 'currentTaxPayable', 'Provisions For Tax');
  const stLoanLbl = lbl(options, 'bs', 'shortTermBorrowings', 'Short Term Loan');
  const expLiabLbl = lbl(options, 'bs', 'otherPayables', 'Liabilities for exp.');
  const otherIncLbl = `Add; ${lbl(options, 'income', 'interestIncome', 'Others Income')}`;
  const finCostLbl = `Less: ${lbl(options, 'income', 'financeCost', 'Financial Expenses')}`;
  const taxExpLbl = `Less: ${lbl(options, 'income', 'taxExpense', 'Provision For Income Tax')}`;
  const repayLbl = lbl(options, 'cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing');

  const customBsCa = getCustomRows(options, 'balanceSheet', 'Current Assets');
  const customBsNca = getCustomRows(options, 'balanceSheet', 'Non-Current Assets');
  const customBsCl = getCustomRows(options, 'balanceSheet', 'Current Liabilities');
  const customBsNcl = getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities');
  const customBsEq = getCustomRows(options, 'balanceSheet', 'Equity');

  const customIncRev = getCustomRows(options, 'income', 'Revenue & Gross Margin');
  const customIncOpex = getCustomRows(options, 'income', 'Operating Expenses');
  const customIncNonOp = [...getCustomRows(options, 'income', 'Non-Operating & Tax'), ...getUnscopedCustomIncomeRows(options)];

  const customCfOp = getCustomRows(options, 'cashFlow', 'Operating Activities');
  const customCfInv = getCustomRows(options, 'cashFlow', 'Investing Activities');
  const customCfFin = getCustomRows(options, 'cashFlow', 'Financing Activities');

  const closingCapLbl = lbl(options, 'eq', 'closingShareCapital', 'Closing Capital / Equity');
  const cogsBaseLbl = lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold');

  const showAdminOpexRow = isVis(options, 'inc_operatingExpenses') || isVis(options, 'inc_depreciationExpense');

  const pushProj3IfNonZero = (target: (string | number)[][], label: string, noteRef: string, vals: number[]) => {
    if (hasAnyNonZeroValue(vals)) target.push([label, noteRef, ...vals]);
  };
  const pushCfIfNonZero = (target: (string | number)[][], label: string, vals: number[]) => {
    if (hasAnyNonZeroValue(vals)) target.push([label, ...vals]);
  };

  // Balance Sheet
  const bsRows: (string | number)[][] = [
    ['Assets', '', ...projData.map(() => '')],
  ];
  const ncaDetailRows: (string | number)[][] = [];
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    pushProj3IfNonZero(
      ncaDetailRows,
      ppeLbl,
      'Schedule-A',
      projData.map(d => d.propertyPlantEquipments - customBsNca.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))
    );
  }
  customBsNca.forEach(ci => pushProj3IfNonZero(ncaDetailRows, ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year))));
  if (ncaDetailRows.length > 0) {
    bsRows.push(['Non Current Assets', '', ...projData.map(() => '')], ...ncaDetailRows);
  }
  bsRows.push(['Total Non-Current Assets', '', ...projData.map(d => d.propertyPlantEquipments)]);

  const caDetailRows: (string | number)[][] = [];
  if (isVis(options, 'bs_closingInventory')) pushProj3IfNonZero(caDetailRows, invLbl, '3.00', projData.map(d => d.inventories));
  if (isVis(options, 'bs_tradeReceivables')) pushProj3IfNonZero(caDetailRows, trRecLbl, '4.00', projData.map(d => d.tradeAndOtherReceivables));
  if (isVis(options, 'bs_otherReceivables')) pushProj3IfNonZero(caDetailRows, advLbl, '5.00', projData.map(d => d.advanceDepositPrepayments));
  if (isVis(options, 'bs_cashAndEquivalents')) pushProj3IfNonZero(caDetailRows, cashLbl, '6.00', projData.map(d => d.cashAndCashEquivalents));
  customBsCa.forEach(ci => pushProj3IfNonZero(caDetailRows, ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year))));
  if (caDetailRows.length > 0) {
    bsRows.push(['Current Assets', '', ...projData.map(() => '')], ...caDetailRows);
  }
  bsRows.push(['Total Current Assets', '', ...projData.map(d => d.totalCurrentAssets)]);
  bsRows.push(['Total Assets', '', ...projData.map(d => d.totalAssets)]);

  bsRows.push(['Equity & Liabilities', '', ...projData.map(() => '')]);
  pushProj3IfNonZero(bsRows, 'Capital', '', projData.map(d => d.capital));
  pushProj3IfNonZero(bsRows, 'Additional Capital', '', projData.map(d => d.additionalCapital));
  pushProj3IfNonZero(bsRows, 'Profit during the year', '', projData.map(d => d.profitDuringYear));
  bsRows.push(['Subtotal', '', ...projData.map(d => d.equitySubtotal)]);
  if (isVis(options, 'eq_dividendsPaid')) {
    pushProj3IfNonZero(bsRows, lbl(options, 'eq', 'dividendsPaid', 'Drawings'), '', projData.map(d => d.drawings));
  }
  bsRows.push([closingCapLbl, '', ...projData.map(d => d.closingCapital)]);
  customBsEq.forEach(ci => pushProj3IfNonZero(bsRows, ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year))));

  const nclDetailRows: (string | number)[][] = [];
  if (isVis(options, 'bs_longTermBorrowings')) {
    pushProj3IfNonZero(
      nclDetailRows,
      ltLoanLbl,
      '7.00',
      projData.map(d => d.longTermLoan - customBsNcl.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))
    );
  }
  customBsNcl.forEach(ci => pushProj3IfNonZero(nclDetailRows, ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year))));
  if (nclDetailRows.length > 0) {
    bsRows.push(['Non-Current Liabilities', '', ...projData.map(() => '')], ...nclDetailRows);
  }

  const clDetailRows: (string | number)[][] = [];
  if (isVis(options, 'bs_tradePayables')) pushProj3IfNonZero(clDetailRows, trPayLbl, '8.00', projData.map(d => d.tradeAndOtherPayables));
  if (isVis(options, 'bs_currentTaxPayable')) pushProj3IfNonZero(clDetailRows, taxPayLbl, '9.00', projData.map(d => d.provisionsForTax));
  if (isVis(options, 'bs_shortTermBorrowings')) pushProj3IfNonZero(clDetailRows, stLoanLbl, '10.00', projData.map(d => d.shortTermLoan));
  if (isVis(options, 'bs_otherPayables')) pushProj3IfNonZero(clDetailRows, expLiabLbl, '', projData.map(d => d.liabilitiesForExp));
  customBsCl.forEach(ci => pushProj3IfNonZero(clDetailRows, ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year))));
  if (clDetailRows.length > 0) {
    bsRows.push(['Current Liabilities', '', ...projData.map(() => '')], ...clDetailRows);
  }
  bsRows.push(['Total Current Liabilities', '', ...projData.map(d => d.totalCurrentLiabilities)]);
  bsRows.push(['Total Equity & Liabilities', '', ...projData.map(d => d.totalEquityAndLiabilities)]);

  // Profit or Loss
  const plRows: (string | number)[][] = [];
  if (isVis(options, 'inc_revenue')) {
    pushProj3IfNonZero(
      plRows,
      salesLbl,
      '',
      projData.map(d => d.sales - customIncRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))
    );
  }
  customIncRev.forEach(ci => pushProj3IfNonZero(plRows, ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year))));
  if (isVis(options, 'inc_costOfSales')) pushProj3IfNonZero(plRows, cogsLbl, '11.00', projData.map(d => d.costOfGoodsSold));
  plRows.push(['Gross Profit', '', ...projData.map(d => d.grossProfit)]);
  if (isVis(options, 'inc_interestIncome')) {
    pushProj3IfNonZero(
      plRows,
      otherIncLbl,
      '',
      projData.map(d => d.othersIncome - customIncNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))
    );
  }
  customIncNonOp.forEach(ci => pushProj3IfNonZero(plRows, ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year))));
  plRows.push(['Total Gross Profit & Other Income', '', ...projData.map(d => d.grossProfitWithOtherIncome)]);
  const plOpexDetailRows: (string | number)[][] = [];
  if (showAdminOpexRow) {
    pushProj3IfNonZero(
      plOpexDetailRows,
      opexLbl,
      '12.00',
      projData.map(d => d.adminAndSellingExpenses + customIncOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0))
    );
  }
  customIncOpex.forEach(ci => pushProj3IfNonZero(plOpexDetailRows, ci.label, '', projData.map(d => -getCustomItemYearValue(ci, d.year))));
  if (plOpexDetailRows.length > 0) {
    plRows.push(['Less: Operating Expenses:', '', ...projData.map(() => '')], ...plOpexDetailRows);
  }
  plRows.push(['Total Operating Profit/(Loss)', '', ...projData.map(d => d.totalOperatingProfit)]);
  if (isVis(options, 'inc_financeCost')) pushProj3IfNonZero(plRows, finCostLbl, '13.00', projData.map(d => d.financialExpenses));
  plRows.push(['Net Profit/(Loss) Before Income Tax', '', ...projData.map(d => d.netProfitBeforeTax)]);
  if (isVis(options, 'inc_taxExpense')) pushProj3IfNonZero(plRows, taxExpLbl, '', projData.map(d => d.provisionForIncomeTax));
  plRows.push(['Net Profit/(Loss) After Income Tax', '', ...projData.map(d => d.netProfitAfterTax)]);

  // Cash Flows
  const cfRows: (string | number)[][] = [];
  if (isVis(options, 'cf_pbt')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'profitBeforeTax', 'Net Income'), projData.map(d => d.cfNetIncome));
  if (isVis(options, 'cf_wc_inventory')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'changeInInventory', '(Increase)/Decrease In Inventory'), projData.map(d => d.cfChangeInventory));
  if (isVis(options, 'cf_wc_receivables')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'changeInReceivables', '(Increase)/Decrease In Trade & Other Receivables'), projData.map(d => d.cfChangeReceivables));
  if (isVis(options, 'cf_wc_other_rec')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'changeInOtherReceivables', '(Increase)/Decrease in Advance, Deposit & Prepayments'), projData.map(d => d.cfChangeAdvanceDeposit));
  if (isVis(options, 'cf_wc_other_ca')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'changeInOtherCA', 'Increase/(Decrease) in Current Assets'), projData.map(d => d.cfChangeCurrentAssets));
  pushCfIfNonZero(cfRows, 'Increase/(Decrease) in Current Liabilities', projData.map(d => d.cfChangeCurrentLiabilities));
  if (isVis(options, 'cf_wc_payables')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'changeInPayables', 'Increase/(Decrease) in Trade & Other Payables'), projData.map(d => d.cfChangeTradePayables));
  if (isVis(options, 'cf_tax_paid')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'incomeTaxPaid', 'Increase/(Decrease) In Provisions For Tax'), projData.map(d => d.cfChangeProvisionsTax));
  if (isVis(options, 'cf_wc_other_pay')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'changeInOtherPayables', 'Increase/(Decrease) Liabilities For Expense'), projData.map(d => d.cfChangeLiabilitiesExp));
  if (isVis(options, 'cf_dep')) pushCfIfNonZero(cfRows, lbl(options, 'cf', 'adjDepreciation', 'Depreciation'), projData.map(d => d.cfDepreciation));
  customCfOp.forEach(ci => pushCfIfNonZero(cfRows, ci.label, projData.map(d => getCustomItemYearValue(ci, d.year))));
  cfRows.push(['Net Cash Used by Operating Activities', ...projData.map(d => d.cfNetCashOperating)]);

  cfRows.push(['2. CASH FLOW FROM INVESTING ACTIVITIES:', ...projData.map(() => '')]);
  const invInflowRows: (string | number)[][] = [];
  if (isVis(options, 'cf_sale_ppe')) pushCfIfNonZero(invInflowRows, lbl(options, 'cf', 'salePPE', 'Sale Of Property, Plant & Equipment'), projData.map(d => d.cfSalePPE));
  pushCfIfNonZero(invInflowRows, 'Capital', projData.map(d => d.cfInvestingCapitalInflow));
  if (invInflowRows.length > 0) {
    cfRows.push(['Inflows:', ...projData.map(() => '')], ...invInflowRows);
  }
  const invOutflowRows: (string | number)[][] = [];
  pushCfIfNonZero(invOutflowRows, 'Preliminary Expenses', projData.map(d => d.cfPreliminaryExpenses));
  if (isVis(options, 'cf_pur_ppe')) pushCfIfNonZero(invOutflowRows, lbl(options, 'cf', 'purchasePPE', 'Purchase Property, Plant & Equipment'), projData.map(d => d.cfPurchasePPE));
  pushCfIfNonZero(invOutflowRows, 'Purchase Investment In Debt Or Securities', projData.map(d => d.cfPurchaseInvestment));
  customCfInv.forEach(ci => pushCfIfNonZero(invOutflowRows, ci.label, projData.map(d => getCustomItemYearValue(ci, d.year))));
  if (invOutflowRows.length > 0) {
    cfRows.push(['Outflows:', ...projData.map(() => '')], ...invOutflowRows);
  }
  cfRows.push(['Net Cash Used by Investing Activities', ...projData.map(d => d.cfNetCashInvesting)]);

  cfRows.push(['3. CASH FLOW FROM FINANCIAL ACTIVITIES:', ...projData.map(() => '')]);
  const finInflowRows: (string | number)[][] = [];
  if (isVis(options, 'cf_borrow_proceeds')) pushCfIfNonZero(finInflowRows, lbl(options, 'cf', 'financingProceedsBorrowings', 'Loan Received'), projData.map(d => d.cfLoanReceived));
  if (isVis(options, 'cf_shares_issue')) pushCfIfNonZero(finInflowRows, 'Additional Capital', projData.map(d => d.cfAdditionalCapital));
  pushCfIfNonZero(finInflowRows, 'Issuance Of Debt (Bond & Notes)', projData.map(d => d.cfIssuanceOfDebt));
  if (finInflowRows.length > 0) {
    cfRows.push(['Inflows:', ...projData.map(() => '')], ...finInflowRows);
  }
  const finOutflowRows: (string | number)[][] = [];
  if (isVis(options, 'cf_dividends_paid')) {
    pushCfIfNonZero(finOutflowRows, 'To Stockholders as Dividend', projData.map(d => d.cfDividendsToStockholders));
    pushCfIfNonZero(finOutflowRows, lbl(options, 'cf', 'financingDividendsPaid', 'Drawings / Dividend'), projData.map(d => d.cfDrawingsOutflow));
  }
  // Mandatory inclusion of 'Repayment of Borrowings & Financing' overrides nil-value omission
  finOutflowRows.push([repayLbl, ...projData.map(d => (isVis(options, 'cf_borrow_repay') ? d.cfRedeemLongTermDebt : 0))]);
  pushCfIfNonZero(finOutflowRows, 'Issuance Of Bonds Payable', projData.map(d => d.cfIssuanceBondsPayable));
  customCfFin.forEach(ci => pushCfIfNonZero(finOutflowRows, ci.label, projData.map(d => getCustomItemYearValue(ci, d.year))));
  if (finOutflowRows.length > 0) {
    cfRows.push(['Outflows:', ...projData.map(() => '')], ...finOutflowRows);
  }
  cfRows.push(['Net Cash Provided by Financial Activities', ...projData.map(d => d.cfNetCashFinancing)]);
  cfRows.push(['Net Cash Increase', ...projData.map(d => d.cfNetCashIncrease)]);
  cfRows.push(['Opening Cash & Bank Balance', ...projData.map(d => d.cfOpeningCash)]);
  cfRows.push(['Closing Cash & Bank Balance', ...projData.map(d => d.cfClosingCash)]);

  const quantNoteObjs = filterProjectedQuantNotesForExport(buildProjectedQuantNoteRows(projData, options));
  const quantNotesRows: (string | number)[][] = quantNoteObjs.map(r => [
    r.noteNo,
    r.label,
    ...r.values.map(v => (v === null ? '' : v)),
  ]);

  return {
    bsRows,
    plRows,
    cfRows,
    quantNotesRows,
    salesLbl,
    trRecLbl,
    opexLbl,
    closingCapLbl,
    cogsBaseLbl,
    invLbl,
    advLbl,
    ltLoanLbl,
    trPayLbl,
    taxPayLbl,
  };
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
  const invLbl = lbl(options, 'bs', 'closingInventory', 'Inventories');
  const advLbl = lbl(options, 'bs', 'otherReceivables', 'Advance, Deposit & Prepayments');
  const ppeLbl = lbl(options, 'bs', 'propertyPlantEquipment', 'Property Plant & Equiptments');
  const ltLoanLbl = lbl(options, 'bs', 'longTermBorrowings', 'Long Term Loan');
  const trPayLbl = lbl(options, 'bs', 'tradePayables', 'Trade & Other Payables');
  const taxPayLbl = lbl(options, 'bs', 'currentTaxPayable', 'Provisions For Tax');
  const stLoanLbl = lbl(options, 'bs', 'shortTermBorrowings', 'Short Term Loan');
  const expLiabLbl = lbl(options, 'bs', 'otherPayables', 'Liabilities for exp.');
  const otherIncLbl = `Add; ${lbl(options, 'income', 'interestIncome', 'Others Income')}`;
  const finCostLbl = `Less: ${lbl(options, 'income', 'financeCost', 'Financial Expenses')}`;
  const taxExpLbl = `Less: ${lbl(options, 'income', 'taxExpense', 'Provision For Income Tax')}`;
  const repayLbl = lbl(options, 'cf', 'repaymentOfBorrowings', 'Repayment of Borrowings & Financing');

  const customBsCa = getCustomRows(options, 'balanceSheet', 'Current Assets');
  const customBsNca = getCustomRows(options, 'balanceSheet', 'Non-Current Assets');
  const customBsCl = getCustomRows(options, 'balanceSheet', 'Current Liabilities');
  const customBsNcl = getCustomRows(options, 'balanceSheet', 'Non-Current Liabilities');
  const customBsEq = getCustomRows(options, 'balanceSheet', 'Equity');

  const customIncRev = getCustomRows(options, 'income', 'Revenue & Gross Margin');
  const customIncOpex = getCustomRows(options, 'income', 'Operating Expenses');
  const customIncNonOp = [...getCustomRows(options, 'income', 'Non-Operating & Tax'), ...getUnscopedCustomIncomeRows(options)];

  const customCfOp = getCustomRows(options, 'cashFlow', 'Operating Activities');
  const customCfInv = getCustomRows(options, 'cashFlow', 'Investing Activities');
  const customCfFin = getCustomRows(options, 'cashFlow', 'Financing Activities');

  const showAdminOpexRow = isVis(options, 'inc_operatingExpenses') || isVis(options, 'inc_depreciationExpense');

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
  ];
  const nextBsRow = () => bsRows.length + 1;
  const ncaCandidates: { label: string; noteRef: string; vals: number[] }[] = [];
  if (isVis(options, 'bs_propertyPlantEquipment')) {
    const ppeVals = projData.map(d => d.propertyPlantEquipments - customBsNca.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0));
    if (hasAnyNonZeroValue(ppeVals)) ncaCandidates.push({ label: ppeLbl, noteRef: 'Schedule-A', vals: ppeVals });
  }
  customBsNca.forEach(ci => {
    const vals = projData.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) ncaCandidates.push({ label: ci.label, noteRef: '', vals });
  });
  const ncaRows: number[] = [];
  if (ncaCandidates.length > 0) {
    bsRows.push(['Non Current Assets', '', ...projData.map(() => '')]);
    ncaCandidates.forEach(item => {
      ncaRows.push(nextBsRow());
      bsRows.push([item.label, item.noteRef, ...item.vals.map(v => nCell(v))]);
    });
  }
  const totalNcaRow = nextBsRow();
  bsRows.push([
    'Total Non-Current Assets',
    '',
    ...projData.map((d, i) =>
      fCell(excelSumRows(colLetter(yearStartColIdx + i), ncaRows), d.propertyPlantEquipments)
    ),
  ]);

  const caCandidates: { label: string; noteRef: string; vals: number[] }[] = [];
  if (isVis(options, 'bs_closingInventory')) {
    const vals = projData.map(d => d.inventories);
    if (hasAnyNonZeroValue(vals)) caCandidates.push({ label: invLbl, noteRef: '3.00', vals });
  }
  if (isVis(options, 'bs_tradeReceivables')) {
    const vals = projData.map(d => d.tradeAndOtherReceivables);
    if (hasAnyNonZeroValue(vals)) caCandidates.push({ label: trRecLbl, noteRef: '4.00', vals });
  }
  if (isVis(options, 'bs_otherReceivables')) {
    const vals = projData.map(d => d.advanceDepositPrepayments);
    if (hasAnyNonZeroValue(vals)) caCandidates.push({ label: advLbl, noteRef: '5.00', vals });
  }
  if (isVis(options, 'bs_cashAndEquivalents')) {
    const vals = projData.map(d => d.cashAndCashEquivalents);
    if (hasAnyNonZeroValue(vals)) caCandidates.push({ label: cashLbl, noteRef: '6.00', vals });
  }
  customBsCa.forEach(ci => {
    const vals = projData.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) caCandidates.push({ label: ci.label, noteRef: '', vals });
  });
  const caRows: number[] = [];
  if (caCandidates.length > 0) {
    bsRows.push(['Current Assets', '', ...projData.map(() => '')]);
    caCandidates.forEach(item => {
      caRows.push(nextBsRow());
      bsRows.push([item.label, item.noteRef, ...item.vals.map(v => nCell(v))]);
    });
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
  const eqSubInputs: number[] = [];
  const pushBs3IfNonZero = (label: string, vals: number[], trackRows: number[]) => {
    if (hasAnyNonZeroValue(vals)) {
      trackRows.push(nextBsRow());
      bsRows.push([label, '', ...vals.map(v => nCell(v))]);
    }
  };
  pushBs3IfNonZero('Capital', projData.map(d => d.capital), eqSubInputs);
  pushBs3IfNonZero('Additional Capital', projData.map(d => d.additionalCapital), eqSubInputs);
  pushBs3IfNonZero('Profit during the year', projData.map(d => d.profitDuringYear), eqSubInputs);
  const eqSubRow = nextBsRow();
  bsRows.push([
    'Subtotal',
    '',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(yearStartColIdx + i), eqSubInputs), d.equitySubtotal)),
  ]);
  let drawRow: number | null = null;
  if (isVis(options, 'eq_dividendsPaid')) {
    const drawVals = projData.map(d => d.drawings);
    if (hasAnyNonZeroValue(drawVals)) {
      drawRow = nextBsRow();
      bsRows.push([lbl(options, 'eq', 'dividendsPaid', 'Drawings'), '', ...drawVals.map(v => nCell(v))]);
    }
  }
  const closingCapRow = nextBsRow();
  bsRows.push([
    lbl(options, 'eq', 'closingShareCapital', 'Closing Capital / Equity'),
    '',
    ...projData.map((d, i) =>
      fCell(excelDiffExpr(colLetter(yearStartColIdx + i), [eqSubRow], drawRow !== null ? [drawRow] : []), d.closingCapital)
    ),
  ]);
  const eqExtraRows: number[] = [];
  customBsEq.forEach(ci => {
    pushBs3IfNonZero(ci.label, projData.map(d => getCustomItemYearValue(ci, d.year)), eqExtraRows);
  });

  const nclCandidates: { label: string; noteRef: string; vals: number[] }[] = [];
  if (isVis(options, 'bs_longTermBorrowings')) {
    const vals = projData.map(d => d.longTermLoan - customBsNcl.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0));
    if (hasAnyNonZeroValue(vals)) nclCandidates.push({ label: ltLoanLbl, noteRef: '7.00', vals });
  }
  customBsNcl.forEach(ci => {
    const vals = projData.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) nclCandidates.push({ label: ci.label, noteRef: '', vals });
  });
  const nclRows: number[] = [];
  if (nclCandidates.length > 0) {
    bsRows.push(['Non-Current Liabilities', '', ...projData.map(() => '')]);
    nclCandidates.forEach(item => {
      nclRows.push(nextBsRow());
      bsRows.push([item.label, item.noteRef, ...item.vals.map(v => nCell(v))]);
    });
  }

  const clCandidates: { label: string; noteRef: string; vals: number[] }[] = [];
  if (isVis(options, 'bs_tradePayables')) {
    const vals = projData.map(d => d.tradeAndOtherPayables);
    if (hasAnyNonZeroValue(vals)) clCandidates.push({ label: trPayLbl, noteRef: '8.00', vals });
  }
  if (isVis(options, 'bs_currentTaxPayable')) {
    const vals = projData.map(d => d.provisionsForTax);
    if (hasAnyNonZeroValue(vals)) clCandidates.push({ label: taxPayLbl, noteRef: '9.00', vals });
  }
  if (isVis(options, 'bs_shortTermBorrowings')) {
    const vals = projData.map(d => d.shortTermLoan);
    if (hasAnyNonZeroValue(vals)) clCandidates.push({ label: stLoanLbl, noteRef: '10.00', vals });
  }
  if (isVis(options, 'bs_otherPayables')) {
    const vals = projData.map(d => d.liabilitiesForExp);
    if (hasAnyNonZeroValue(vals)) clCandidates.push({ label: expLiabLbl, noteRef: '', vals });
  }
  customBsCl.forEach(ci => {
    const vals = projData.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) clCandidates.push({ label: ci.label, noteRef: '', vals });
  });
  const clRows: number[] = [];
  if (clCandidates.length > 0) {
    bsRows.push(['Current Liabilities', '', ...projData.map(() => '')]);
    clCandidates.forEach(item => {
      clRows.push(nextBsRow());
      bsRows.push([item.label, item.noteRef, ...item.vals.map(v => nCell(v))]);
    });
  }
  const totalClRow = nextBsRow();
  bsRows.push([
    'Total Current Liabilities',
    '',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(yearStartColIdx + i), clRows), d.totalCurrentLiabilities)),
  ]);
  const totalEqLiabInputs = [closingCapRow, ...eqExtraRows, ...nclRows, totalClRow];
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
  const pushPlIfNonZero = (label: string, noteRef: string, vals: number[], trackRows: number[]) => {
    if (hasAnyNonZeroValue(vals)) {
      trackRows.push(nextPlRow());
      plRows.push([label, noteRef, ...vals.map(v => nCell(v))]);
    }
  };
  const plGpAddRows: number[] = [];
  if (isVis(options, 'inc_revenue')) {
    pushPlIfNonZero(
      salesLbl,
      '',
      projData.map(d => d.sales - customIncRev.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)),
      plGpAddRows
    );
  }
  customIncRev.forEach(ci => {
    pushPlIfNonZero(ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year)), plGpAddRows);
  });
  let plCogsRow: number | null = null;
  if (isVis(options, 'inc_costOfSales')) {
    const cogsVals = projData.map(d => d.costOfGoodsSold);
    if (hasAnyNonZeroValue(cogsVals)) {
      plCogsRow = nextPlRow();
      plRows.push([cogsLbl, '11.00', ...cogsVals.map(v => nCell(v))]);
    }
  }
  const plGpRow = nextPlRow();
  plRows.push([
    'Gross Profit',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelDiffExpr(
          colLetter(yearStartColIdx + i),
          plGpAddRows,
          plCogsRow !== null ? [plCogsRow] : []
        ),
        d.grossProfit
      )
    ),
  ]);
  const plSubAddRows: number[] = [plGpRow];
  if (isVis(options, 'inc_interestIncome')) {
    pushPlIfNonZero(
      otherIncLbl,
      '',
      projData.map(d => d.othersIncome - customIncNonOp.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0)),
      plSubAddRows
    );
  }
  customIncNonOp.forEach(ci => {
    pushPlIfNonZero(ci.label, '', projData.map(d => getCustomItemYearValue(ci, d.year)), plSubAddRows);
  });
  const plSubRow = nextPlRow();
  plRows.push([
    'Total Gross Profit & Other Income',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelSumRows(colLetter(yearStartColIdx + i), plSubAddRows),
        d.grossProfitWithOtherIncome
      )
    ),
  ]);
  const plOpexCandidates: { label: string; noteRef: string; vals: number[] }[] = [];
  if (showAdminOpexRow) {
    const vals = projData.map(d => d.adminAndSellingExpenses + customIncOpex.reduce((s, ci) => s + getCustomItemYearValue(ci, d.year), 0));
    if (hasAnyNonZeroValue(vals)) plOpexCandidates.push({ label: opexLbl, noteRef: '12.00', vals });
  }
  customIncOpex.forEach(ci => {
    const vals = projData.map(d => -getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) plOpexCandidates.push({ label: ci.label, noteRef: '', vals });
  });
  const plOpexRows: number[] = [];
  if (plOpexCandidates.length > 0) {
    plRows.push(['Less: Operating Expenses:', '', ...projData.map(() => '')]);
    plOpexCandidates.forEach(item => {
      plOpexRows.push(nextPlRow());
      plRows.push([item.label, item.noteRef, ...item.vals.map(v => nCell(v))]);
    });
  }
  const plOpProfitRow = nextPlRow();
  plRows.push([
    'Total Operating Profit/(Loss)',
    '',
    ...projData.map((d, i) =>
      fCell(
        excelDiffExpr(colLetter(yearStartColIdx + i), [plSubRow], plOpexRows),
        d.totalOperatingProfit
      )
    ),
  ]);
  let plFinRow: number | null = null;
  if (isVis(options, 'inc_financeCost')) {
    const finVals = projData.map(d => d.financialExpenses);
    if (hasAnyNonZeroValue(finVals)) {
      plFinRow = nextPlRow();
      plRows.push([finCostLbl, '13.00', ...finVals.map(v => nCell(v))]);
    }
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
    const taxVals = projData.map(d => d.provisionForIncomeTax);
    if (hasAnyNonZeroValue(taxVals)) {
      plTaxRow = nextPlRow();
      plRows.push([taxExpLbl, '', ...taxVals.map(v => nCell(v))]);
    }
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
  const pushCfExcelIfNonZero = (label: string, vals: number[], trackRows: number[]) => {
    if (hasAnyNonZeroValue(vals)) {
      trackRows.push(nextCfRow());
      cfRows.push([label, ...vals.map(v => nCell(v))]);
    }
  };
  const cfOpRows: number[] = [];
  if (isVis(options, 'cf_pbt')) pushCfExcelIfNonZero(lbl(options, 'cf', 'profitBeforeTax', 'Net Income'), projData.map(d => d.cfNetIncome), cfOpRows);
  if (isVis(options, 'cf_wc_inventory')) pushCfExcelIfNonZero(lbl(options, 'cf', 'changeInInventory', '(Increase)/Decrease In Inventory'), projData.map(d => d.cfChangeInventory), cfOpRows);
  if (isVis(options, 'cf_wc_receivables')) pushCfExcelIfNonZero(lbl(options, 'cf', 'changeInReceivables', '(Increase)/Decrease In Trade & Other Receivables'), projData.map(d => d.cfChangeReceivables), cfOpRows);
  if (isVis(options, 'cf_wc_other_rec')) pushCfExcelIfNonZero(lbl(options, 'cf', 'changeInOtherReceivables', '(Increase)/Decrease in Advance, Deposit & Prepayments'), projData.map(d => d.cfChangeAdvanceDeposit), cfOpRows);
  if (isVis(options, 'cf_wc_other_ca')) pushCfExcelIfNonZero(lbl(options, 'cf', 'changeInOtherCA', 'Increase/(Decrease) in Current Assets'), projData.map(d => d.cfChangeCurrentAssets), cfOpRows);
  pushCfExcelIfNonZero('Increase/(Decrease) in Current Liabilities', projData.map(d => d.cfChangeCurrentLiabilities), cfOpRows);
  if (isVis(options, 'cf_wc_payables')) pushCfExcelIfNonZero(lbl(options, 'cf', 'changeInPayables', 'Increase/(Decrease) in Trade & Other Payables'), projData.map(d => d.cfChangeTradePayables), cfOpRows);
  if (isVis(options, 'cf_tax_paid')) pushCfExcelIfNonZero(lbl(options, 'cf', 'incomeTaxPaid', 'Increase/(Decrease) In Provisions For Tax'), projData.map(d => d.cfChangeProvisionsTax), cfOpRows);
  if (isVis(options, 'cf_wc_other_pay')) pushCfExcelIfNonZero(lbl(options, 'cf', 'changeInOtherPayables', 'Increase/(Decrease) Liabilities For Expense'), projData.map(d => d.cfChangeLiabilitiesExp), cfOpRows);
  if (isVis(options, 'cf_dep')) pushCfExcelIfNonZero(lbl(options, 'cf', 'adjDepreciation', 'Depreciation'), projData.map(d => d.cfDepreciation), cfOpRows);
  customCfOp.forEach(ci => pushCfExcelIfNonZero(ci.label, projData.map(d => getCustomItemYearValue(ci, d.year)), cfOpRows));
  const cfNetOpRow = nextCfRow();
  cfRows.push([
    'Net Cash Used by Operating Activities',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfOpRows), d.cfNetCashOperating)),
  ]);

  cfRows.push(['2. CASH FLOW FROM INVESTING ACTIVITIES:', ...projData.map(() => '')]);
  const cfInvRows: number[] = [];
  const invInflowCandidates: { label: string; vals: number[] }[] = [];
  if (isVis(options, 'cf_sale_ppe')) {
    const vals = projData.map(d => d.cfSalePPE);
    if (hasAnyNonZeroValue(vals)) invInflowCandidates.push({ label: lbl(options, 'cf', 'salePPE', 'Sale Of Property, Plant & Equipment'), vals });
  }
  const invCapVals = projData.map(d => d.cfInvestingCapitalInflow);
  if (hasAnyNonZeroValue(invCapVals)) invInflowCandidates.push({ label: 'Capital', vals: invCapVals });
  if (invInflowCandidates.length > 0) {
    cfRows.push(['Inflows:', ...projData.map(() => '')]);
    invInflowCandidates.forEach(item => {
      cfInvRows.push(nextCfRow());
      cfRows.push([item.label, ...item.vals.map(v => nCell(v))]);
    });
  }
  const invOutflowCandidates: { label: string; vals: number[] }[] = [];
  const prelimVals = projData.map(d => d.cfPreliminaryExpenses);
  if (hasAnyNonZeroValue(prelimVals)) invOutflowCandidates.push({ label: 'Preliminary Expenses', vals: prelimVals });
  if (isVis(options, 'cf_pur_ppe')) {
    const vals = projData.map(d => d.cfPurchasePPE);
    if (hasAnyNonZeroValue(vals)) invOutflowCandidates.push({ label: lbl(options, 'cf', 'purchasePPE', 'Purchase Property, Plant & Equipment'), vals });
  }
  const purInvVals = projData.map(d => d.cfPurchaseInvestment);
  if (hasAnyNonZeroValue(purInvVals)) invOutflowCandidates.push({ label: 'Purchase Investment In Debt Or Securities', vals: purInvVals });
  customCfInv.forEach(ci => {
    const vals = projData.map(d => getCustomItemYearValue(ci, d.year));
    if (hasAnyNonZeroValue(vals)) invOutflowCandidates.push({ label: ci.label, vals });
  });
  if (invOutflowCandidates.length > 0) {
    cfRows.push(['Outflows:', ...projData.map(() => '')]);
    invOutflowCandidates.forEach(item => {
      cfInvRows.push(nextCfRow());
      cfRows.push([item.label, ...item.vals.map(v => nCell(v))]);
    });
  }
  const cfNetInvRow = nextCfRow();
  cfRows.push([
    'Net Cash Used by Investing Activities',
    ...projData.map((d, i) => fCell(excelSumRows(colLetter(i + 1), cfInvRows), d.cfNetCashInvesting)),
  ]);

  cfRows.push(['3. CASH FLOW FROM FINANCIAL ACTIVITIES:', ...projData.map(() => '')]);
  const cfFinRows: number[] = [];
  const finInflowCandidates: { label: string; vals: number[] }[] = [];
  if (isVis(options, 'cf_borrow_proceeds')) {
    const vals = projData.map(d => d.cfLoanReceived);
    if (hasAnyNonZeroValue(vals)) finInflowCandidates.push({ label: lbl(options, 'cf', 'financingProceedsBorrowings', 'Loan Received'), vals });
  }
  if (isVis(options, 'cf_shares_issue')) {
    const vals = projData.map(d => d.cfAdditionalCapital);
    if (hasAnyNonZeroValue(vals)) finInflowCandidates.push({ label: 'Additional Capital', vals });
  }
  const debtIssVals = projData.map(d => d.cfIssuanceOfDebt);
  if (hasAnyNonZeroValue(debtIssVals)) finInflowCandidates.push({ label: 'Issuance Of Debt (Bond & Notes)', vals: debtIssVals });
  if (finInflowCandidates.length > 0) {
    cfRows.push(['Inflows:', ...projData.map(() => '')]);
    finInflowCandidates.forEach(item => {
      cfFinRows.push(nextCfRow());
      cfRows.push([item.label, ...item.vals.map(v => nCell(v))]);
    });
  }

  cfRows.push(['Outflows:', ...projData.map(() => '')]);
  if (isVis(options, 'cf_dividends_paid')) {
    pushCfExcelIfNonZero('To Stockholders as Dividend', projData.map(d => d.cfDividendsToStockholders), cfFinRows);
    pushCfExcelIfNonZero(lbl(options, 'cf', 'financingDividendsPaid', 'Drawings / Dividend'), projData.map(d => d.cfDrawingsOutflow), cfFinRows);
  }
  // Mandatory inclusion of 'Repayment of Borrowings & Financing' overrides nil-value omission
  cfFinRows.push(nextCfRow());
  cfRows.push([repayLbl, ...projData.map(d => nCell(isVis(options, 'cf_borrow_repay') ? d.cfRedeemLongTermDebt : 0))]);
  pushCfExcelIfNonZero('Issuance Of Bonds Payable', projData.map(d => d.cfIssuanceBondsPayable), cfFinRows);
  customCfFin.forEach(ci => {
    pushCfExcelIfNonZero(ci.label, projData.map(d => getCustomItemYearValue(ci, d.year)), cfFinRows);
  });
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
    closingCapLbl,
    cogsBaseLbl,
    invLbl,
    advLbl,
    ltLoanLbl,
    trPayLbl,
    taxPayLbl,
  } = buildProjectedTextTables(projData, options);

  // Page 1: Statement of Financial Position
  const startY1 = drawProjHeader('Statement of Financial Position', `As at 31st December, ${maxYear}(Projected)`);
  const bsBoldRows = new Set([
    'Assets', 'Non Current Assets', 'Total Non-Current Assets', 'Current Assets', 'Total Current Assets',
    'Total Assets', 'Equity & Liabilities', 'Subtotal', closingCapLbl, 'Non-Current Liabilities',
    'Current Liabilities', 'Total Current Liabilities', 'Total Equity & Liabilities',
  ]);
  const bsRuleRows = new Set([
    'Total Non-Current Assets', 'Total Current Assets', 'Total Assets',
    'Subtotal', closingCapLbl, 'Total Current Liabilities', 'Total Equity & Liabilities',
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
    'Gross Profit', 'Total Gross Profit & Other Income', 'Total Operating Profit/(Loss)',
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
    'Subtotal', `Closing ${trRecLbl}`, `Closing ${advLbl}`,
    'Cash In Hand & Bank', ltLoanLbl, `Closing ${trPayLbl}`, `Closing ${taxPayLbl}`,
    'Bank Loan', 'Material Consumed', 'Total Direct Expenses', cogsBaseLbl,
    opexLbl, invLbl, 'Bank Charges & Interest',
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
