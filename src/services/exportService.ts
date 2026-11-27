import * as XLSX from 'xlsx';
import XLSXStyle from 'xlsx-js-style';
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  TextRun,
  AlignmentType,
  WidthType,
  BorderStyle,
  PageBreak,
  PageOrientation,
  convertMillimetersToTwip,
  VerticalAlign,
  ShadingType,
  TableLayoutType,
  HeightRule,
  Footer,
  PageNumber,
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
  ParsedRepaymentSchedule,
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
  formatFyEndDateCode,
  formatFyStartDateCode,
  formatFyAsAtUpper,
  formatFyYearEndedTitle,
  formatFyYearEndedNoComma,
  formatFyProjectedAsAt,
  formatFyProjectedYearEnded,
  normalizeFinancialYearEnd,
} from './financialEngine';

const safeCompanyName = (c: CompanyProfile): string => c.companyName?.trim() || 'Corporate Entity';
const safeCompanyUpper = (c: CompanyProfile): string => (c.companyName?.trim() || 'FINANCIAL STATEMENTS').toUpperCase();
const safeAddress = (c: CompanyProfile): string => c.registeredAddress?.trim() || 'Registered Corporate Office';

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
  if (typeof document === 'undefined') return;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

const exportedDocxBuffersForTest: { filename: string; size: number }[] = [];
export function getCapturedDocxExportsForTest() {
  return exportedDocxBuffersForTest;
}
export function clearCapturedDocxExportsForTest() {
  exportedDocxBuffersForTest.length = 0;
}

async function emitDocxDocument(doc: Document, filename: string) {
  if (typeof document === 'undefined') {
    const buf = await Packer.toBuffer(doc);
    exportedDocxBuffersForTest.push({ filename, size: buf.byteLength });
    return;
  }
  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, filename);
}

const getXlsxStyleEngine = (): typeof XLSX => {
  const mod = XLSXStyle as unknown as { writeFile?: typeof XLSX.writeFile; default?: typeof XLSX };
  return (mod && typeof mod.writeFile === 'function' ? mod : mod?.default) as typeof XLSX;
};

function writeWorkbookStyled(wb: XLSX.WorkBook, filename: string) {
  if (typeof window === 'undefined') {
    XLSX.writeFile(wb, filename);
    return;
  }
  getXlsxStyleEngine().writeFile(wb, filename);
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

type ExcelBorderSide = { style: 'thin' | 'medium' | 'double'; color: { rgb: string } };
type ExcelCellStyle = {
  font?: { name?: string; sz?: number; bold?: boolean; italic?: boolean; color?: { rgb: string } };
  fill?: { patternType: 'solid'; fgColor: { rgb: string } };
  alignment?: { horizontal?: 'left' | 'center' | 'right'; vertical?: 'top' | 'center' | 'bottom'; wrapText?: boolean };
  border?: {
    top?: ExcelBorderSide;
    bottom?: ExcelBorderSide;
    left?: ExcelBorderSide;
    right?: ExcelBorderSide;
  };
};

type StyledCellObj = XLSX.CellObject & { s?: ExcelCellStyle };

function setWorksheetPrintSetup(
  ws: XLSX.WorkSheet,
  orientation: 'landscape' | 'portrait',
  marginMm = 14,
  fitToWidth = 1
) {
  const marginIn = Number((marginMm / 25.4).toFixed(3));
  (ws as Record<string, unknown>)['!pageSetup'] = {
    paperSize: 9, // ISO A4
    orientation,
    fitToWidth,
    fitToHeight: 0,
    scale: 100,
  };
  (ws as Record<string, unknown>)['!margins'] = {
    left: marginIn,
    right: marginIn,
    top: marginIn,
    bottom: marginIn,
    header: 0.25,
    footer: 0.25,
  };
}

function setRowHeightPt(ws: XLSX.WorkSheet, rowZeroBased: number, hpt: number) {
  if (!ws['!rows']) ws['!rows'] = [];
  ws['!rows'][rowZeroBased] = { hpt };
}

function addWorksheetMerge(ws: XLSX.WorkSheet, sR: number, sC: number, eR: number, eC: number) {
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: sR, c: sC }, e: { r: eR, c: eC } });
}

function applyCellStyle(ws: XLSX.WorkSheet, r: number, c: number, style: ExcelCellStyle, createIfMissing = true) {
  const addr = XLSX.utils.encode_cell({ r, c });
  let cell = ws[addr] as StyledCellObj | undefined;
  if (!cell) {
    if (!createIfMissing) return;
    cell = { t: 's', v: '' };
    ws[addr] = cell;
  }
  cell.s = style;
}

function getRowStringVal(ws: XLSX.WorkSheet, r: number, c = 0): string {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = ws[addr] as XLSX.CellObject | undefined;
  return cell && cell.v !== undefined && cell.v !== null ? String(cell.v) : '';
}

function getSheetRowCount(ws: XLSX.WorkSheet): number {
  const ref = ws['!ref'];
  if (!ref) return 0;
  return XLSX.utils.decode_range(ref).e.r + 1;
}

const BORDER_THIN_SLATE: ExcelBorderSide = { style: 'thin', color: { rgb: 'CBD5E1' } };
const BORDER_THIN_GRID: ExcelBorderSide = { style: 'thin', color: { rgb: '334155' } };
const BORDER_THIN_DARK: ExcelBorderSide = { style: 'thin', color: { rgb: '0F172A' } };
const BORDER_DOUBLE_DARK: ExcelBorderSide = { style: 'double', color: { rgb: '0F172A' } };

const FULL_THIN_SLATE = {
  top: BORDER_THIN_SLATE,
  bottom: BORDER_THIN_SLATE,
  left: BORDER_THIN_SLATE,
  right: BORDER_THIN_SLATE,
};

const FULL_THIN_GRID = {
  top: BORDER_THIN_GRID,
  bottom: BORDER_THIN_GRID,
  left: BORDER_THIN_GRID,
  right: BORDER_THIN_GRID,
};

function styleSection76Sheet(
  ws: XLSX.WorkSheet,
  colCount: number,
  opts?: { centerFirstCols?: number; centerAllHeaders?: boolean }
) {
  setWorksheetPrintSetup(ws, 'landscape', 14, 1);
  const totalRows = getSheetRowCount(ws);
  if (totalRows === 0) return;

  // Masthead Rows 0..3 matching PDF #0F172A dark slate banner
  for (let r = 0; r <= 3; r++) {
    if (colCount > 1) addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 25 : 17);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: {
          name: 'Helvetica',
          sz: r === 0 ? 13 : r === 1 ? 10 : 8.5,
          bold: r === 0 || r === 1,
          color: { rgb: r <= 1 ? 'FFFFFF' : 'CBD5E1' },
        },
        fill: { patternType: 'solid', fgColor: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      });
    }
  }
  setRowHeightPt(ws, 4, 8);

  // Header Row 5 matching PDF #1E293B dark navy header
  const headerRow = 5;
  setRowHeightPt(ws, headerRow, 22);
  const leftCols = opts?.centerFirstCols ?? 1;
  for (let c = 0; c < colCount; c++) {
    applyCellStyle(ws, headerRow, c, {
      font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: 'FFFFFF' } },
      fill: { patternType: 'solid', fgColor: { rgb: '1E293B' } },
      alignment: {
        horizontal: opts?.centerAllHeaders ? 'center' : c < leftCols ? 'left' : 'right',
        vertical: 'center',
        wrapText: true,
      },
      border: FULL_THIN_GRID,
    });
  }

  let zebraIdx = 0;
  for (let r = 6; r < totalRows; r++) {
    const label = getRowStringVal(ws, r, 0).trim();
    const col1 = getRowStringVal(ws, r, 1).trim();
    if (!label && !col1) {
      setRowHeightPt(ws, r, 9);
      continue;
    }

    const isGrandTotal =
      label.startsWith('TOTAL ') ||
      label === 'Net Profit After Tax' ||
      label === 'NET CASH GENERATED FROM OPERATING ACTIVITIES' ||
      label === 'NET CASH FROM / (USED IN) INVESTING ACTIVITIES' ||
      label === 'NET CASH FROM / (USED IN) FINANCING ACTIVITIES' ||
      label === 'Closing Cash & Cash Equivalents (Calculated)' ||
      label === 'Balance Check (Variance)' ||
      label === 'Reconciliation Variance (Calculated - Stated)' ||
      label === 'Total :';

    const isSectionHeader =
      label === 'ASSETS' ||
      label === 'LIABILITIES & SHAREHOLDERS EQUITY' ||
      label.startsWith('CASH FLOWS FROM ');

    const isSubtotal =
      !isGrandTotal &&
      !isSectionHeader &&
      (label === 'Gross Profit' ||
        label === 'Operating Profit (EBIT)' ||
        label === 'Profit Before Tax (EBT)' ||
        label.startsWith('Total ') ||
        label.startsWith('Closing Share') ||
        label.startsWith('Closing Retained') ||
        label.startsWith('Closing Reserves') ||
        label.startsWith('Operating Profit Before') ||
        label.startsWith('Cash Generated') ||
        label.startsWith('NET INCREASE'));

    const isStatusRow = label === 'Audit Reconciliation Status';

    setRowHeightPt(ws, r, isGrandTotal || isSectionHeader ? 20 : 17.5);
    const fillHex = isGrandTotal
      ? 'ECFDF5'
      : isSectionHeader
      ? 'E2E8F0'
      : isSubtotal
      ? 'F1F5F9'
      : zebraIdx % 2 === 1
      ? 'F8FAFC'
      : 'FFFFFF';
    if (!isSectionHeader) zebraIdx++;

    for (let c = 0; c < colCount; c++) {
      const cellVal = getRowStringVal(ws, r, c).trim();
      const alignH: 'left' | 'center' | 'right' =
        c === 0
          ? 'left'
          : opts?.centerFirstCols && c < opts.centerFirstCols
          ? 'center'
          : isStatusRow
          ? 'center'
          : cellVal.endsWith('%') || cellVal.endsWith('x')
          ? 'right'
          : 'right';
      const cellFill =
        isStatusRow && c > 0
          ? cellVal === 'RECONCILED'
            ? 'ECFDF5'
            : 'FEF2F2'
          : fillHex;
      const cellColor =
        isStatusRow && c > 0
          ? cellVal === 'RECONCILED'
            ? '065F46'
            : '991B1B'
          : '0F172A';
      applyCellStyle(ws, r, c, {
        font: {
          name: 'Helvetica',
          sz: 8.5,
          bold: isGrandTotal || isSubtotal || isSectionHeader || isStatusRow,
          color: { rgb: cellColor },
        },
        fill: { patternType: 'solid', fgColor: { rgb: cellFill } },
        alignment: { horizontal: alignH, vertical: 'center', wrapText: c === colCount - 1 && colCount >= 10 },
        border: isGrandTotal
          ? {
              top: BORDER_THIN_DARK,
              bottom: BORDER_DOUBLE_DARK,
              left: BORDER_THIN_SLATE,
              right: BORDER_THIN_SLATE,
            }
          : FULL_THIN_SLATE,
      });
    }
  }
}

function buildSection76CoverWorksheet(company: CompanyProfile, horizonText: string): XLSX.WorkSheet {
  const rows: ExcelCellValue[][] = [
    [''],
    [''],
    [safeCompanyUpper(company)],
    ['ANNUAL STATUTORY FINANCIAL STATEMENTS & AUDIT SCHEDULES'],
    ['Five Standard Financial Statements + Complete Notes to the Financial Statements'],
    [`Compliant with Section 76 & International Accounting Framework (${company.accountingFramework})`],
    [''],
    [`Reporting Entity Type: ${company.entityType.toUpperCase()} | Registration: ${company.registrationNumber || 'N/A'} | TIN: ${company.taxId || 'N/A'}`],
    [`Corporate Address: ${safeAddress(company)}`],
    [`Reporting Currency: ${company.reportingCurrency} (${company.reportingUnit}) | FY End: ${normalizeFinancialYearEnd(company.financialYearEnd)} | Horizon: ${horizonText}`],
    [''],
    [`Lead Financial Architect: ${company.preparedBy}`],
    [`Independent Statutory Auditors: ${company.auditorName || 'Chartered Accountants'}`],
    [''],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  applyColumnWidths(ws, [28, 28, 28, 28, 28]);
  setWorksheetPrintSetup(ws, 'landscape', 10, 1);
  for (let r = 0; r < rows.length; r++) {
    addWorksheetMerge(ws, r, 0, r, 4);
    setRowHeightPt(ws, r, r === 2 ? 38 : r === 3 ? 28 : 20);
    const fontSz = r === 2 ? 22 : r === 3 ? 14 : r <= 5 ? 11 : 10;
    const textRgb = r === 2 ? 'FFFFFF' : r === 3 ? '93C5FD' : r === 11 ? '34D399' : r >= 7 ? 'CBD5E1' : 'E2E8F0';
    for (let c = 0; c < 5; c++) {
      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: fontSz, bold: r <= 3 || r === 11, color: { rgb: textRgb } },
        fill: { patternType: 'solid', fgColor: { rgb: '0F172A' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: {
          top: r === 0 ? { style: 'medium', color: { rgb: '3B82F6' } } : undefined,
          bottom: r === rows.length - 1 ? { style: 'medium', color: { rgb: '3B82F6' } } : undefined,
          left: c === 0 ? { style: 'medium', color: { rgb: '3B82F6' } } : undefined,
          right: c === 4 ? { style: 'medium', color: { rgb: '3B82F6' } } : undefined,
        },
      });
    }
  }
  return ws;
}

function styleSection76NotesSheet(ws: XLSX.WorkSheet, colCount = 4) {
  setWorksheetPrintSetup(ws, 'landscape', 14, 1);
  const totalRows = getSheetRowCount(ws);
  if (totalRows === 0) return;

  for (let r = 0; r <= 2; r++) {
    addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 24 : 18);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: r === 0 ? 12.5 : 9.5, bold: r <= 1, color: { rgb: r <= 1 ? 'FFFFFF' : 'CBD5E1' } },
        fill: { patternType: 'solid', fgColor: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      });
    }
  }

  let inTable = false;
  let tableZebra = 0;
  for (let r = 3; r < totalRows; r++) {
    const col0 = getRowStringVal(ws, r, 0).trim();
    const col1 = getRowStringVal(ws, r, 1).trim();
    if (!col0 && !col1) {
      inTable = false;
      tableZebra = 0;
      setRowHeightPt(ws, r, 9);
      continue;
    }
    if (col0.startsWith('Note ')) {
      inTable = false;
      addWorksheetMerge(ws, r, 0, r, colCount - 1);
      setRowHeightPt(ws, r, 19);
      for (let c = 0; c < colCount; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Helvetica', sz: 9.5, bold: true, color: { rgb: '1E293B' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'E2E8F0' } },
          alignment: { horizontal: 'left', vertical: 'center' },
        });
      }
      continue;
    }
    if (!inTable && !col1 && col0.length > 40) {
      addWorksheetMerge(ws, r, 0, r, colCount - 1);
      setRowHeightPt(ws, r, 26);
      applyCellStyle(ws, r, 0, {
        font: { name: 'Helvetica', sz: 8.5, color: { rgb: '475569' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      });
      continue;
    }
    if (!inTable && col1) {
      inTable = true;
      tableZebra = 0;
      setRowHeightPt(ws, r, 20);
      for (let c = 0; c < colCount; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '1E293B' } },
          alignment: { horizontal: c === 0 ? 'left' : 'right', vertical: 'center' },
          border: FULL_THIN_GRID,
        });
      }
      continue;
    }
    if (inTable) {
      const nextCol0 = r + 1 < totalRows ? getRowStringVal(ws, r + 1, 0).trim() : '';
      const isTotal = col0.toLowerCase().startsWith('total') || nextCol0 === '';
      setRowHeightPt(ws, r, 17);
      const fillHex = isTotal ? 'F1F5F9' : tableZebra % 2 === 1 ? 'F8FAFC' : 'FFFFFF';
      tableZebra++;
      for (let c = 0; c < colCount; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Helvetica', sz: 8.5, bold: isTotal, color: { rgb: '0F172A' } },
          fill: { patternType: 'solid', fgColor: { rgb: fillHex } },
          alignment: { horizontal: c === 0 ? 'left' : 'right', vertical: 'center' },
          border: FULL_THIN_SLATE,
        });
      }
    }
  }
}

function stylePreviousFormatSheet(
  ws: XLSX.WorkSheet,
  colCount: number,
  orientation: 'portrait' | 'landscape',
  boldLabels: Set<string>
) {
  setWorksheetPrintSetup(ws, orientation, 14, 1);
  const totalRows = getSheetRowCount(ws);
  if (totalRows === 0) return;

  for (let r = 0; r <= 3; r++) {
    if (colCount > 1) addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 22 : r === 1 ? 19 : 16);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: {
          name: 'Times New Roman',
          sz: r === 0 ? 13 : r === 1 ? 12 : r === 2 ? 9.5 : 8.5,
          bold: r <= 2,
          color: { rgb: '0F172A' },
        },
        alignment: { horizontal: r === 3 ? 'right' : 'center', vertical: 'center' },
      });
    }
  }
  setRowHeightPt(ws, 4, 7);

  // Header row 5 (white background, bold centered, grid borders matching Previous Financials PDF)
  setRowHeightPt(ws, 5, 22);
  for (let c = 0; c < colCount; c++) {
    applyCellStyle(ws, 5, c, {
      font: { name: 'Times New Roman', sz: 9.5, bold: true, color: { rgb: '0F172A' } },
      fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: FULL_THIN_GRID,
    });
  }

  const doubleBottomLabels = new Set([
    'Taka',
    'Total',
    'Net Profit After Income Tax',
    'f) Cash & Cash Equivalents at end of period (D+E)',
    'TOTAL CLOSING EQUITY',
  ]);

  for (let r = 6; r < totalRows; r++) {
    const label = getRowStringVal(ws, r, 0).trim();
    const isBold =
      boldLabels.has(label) ||
      label === 'PROPERTIES & ASSETS' ||
      label === 'CAPITAL & LIABILITIES' ||
      label.startsWith('a)') ||
      label.startsWith('b)') ||
      label.startsWith('c)');
    const isDoubleBottom = doubleBottomLabels.has(label);
    setRowHeightPt(ws, r, 18);
    for (let c = 0; c < colCount; c++) {
      const col0Align: 'left' | 'right' =
        c === 0 ? (label === 'Taka' || label === 'Total' ? 'right' : 'left') : 'right';
      applyCellStyle(ws, r, c, {
        font: { name: 'Times New Roman', sz: 9, bold: isBold, color: { rgb: '0F172A' } },
        fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
        alignment: { horizontal: col0Align, vertical: 'center' },
        border: {
          top: BORDER_THIN_GRID,
          bottom: isDoubleBottom ? BORDER_DOUBLE_DARK : BORDER_THIN_GRID,
          left: BORDER_THIN_GRID,
          right: BORDER_THIN_GRID,
        },
      });
    }
  }
}

function stylePreviousFaNotesSheet(ws: XLSX.WorkSheet, orientation: 'portrait' | 'landscape') {
  setWorksheetPrintSetup(ws, orientation, 14, 1);
  const totalRows = getSheetRowCount(ws);
  const colCount = 7;
  for (let r = 0; r <= 2; r++) {
    addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 22 : 18);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: { name: 'Times New Roman', sz: r === 0 ? 13 : 10, bold: true, color: { rgb: '0F172A' } },
        alignment: { horizontal: r < 2 ? 'center' : 'left', vertical: 'center' },
      });
    }
  }

  let mode: 'fa' | 'notes_head' | 'note_tbl' = 'fa';
  let noteCols = 3;
  for (let r = 3; r < totalRows; r++) {
    const col0 = getRowStringVal(ws, r, 0).trim();
    const col1 = getRowStringVal(ws, r, 1).trim();
    if (!col0 && !col1) {
      if (mode === 'note_tbl') mode = 'notes_head';
      setRowHeightPt(ws, r, 8);
      continue;
    }
    if (col0 === 'NOTES TO THE FINANCIAL STATEMENTS:') {
      mode = 'notes_head';
      addWorksheetMerge(ws, r, 0, r, colCount - 1);
      setRowHeightPt(ws, r, 20);
      applyCellStyle(ws, r, 0, {
        font: { name: 'Times New Roman', sz: 10.5, bold: true, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      });
      continue;
    }
    if (mode === 'fa') {
      const isHdr = col0 === 'Particulars';
      const isTaka = col0 === 'Taka';
      setRowHeightPt(ws, r, isHdr ? 24 : 18);
      for (let c = 0; c < colCount; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Times New Roman', sz: 8.5, bold: isHdr || isTaka, color: { rgb: '0F172A' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
          alignment: {
            horizontal: isHdr ? 'center' : c === 0 ? 'left' : c === 4 ? 'center' : 'right',
            vertical: 'center',
            wrapText: isHdr,
          },
          border: {
            top: BORDER_THIN_GRID,
            bottom: isTaka ? BORDER_DOUBLE_DARK : BORDER_THIN_GRID,
            left: BORDER_THIN_GRID,
            right: BORDER_THIN_GRID,
          },
        });
      }
      continue;
    }
    if (col0.startsWith('Note ')) {
      mode = 'notes_head';
      addWorksheetMerge(ws, r, 0, r, colCount - 1);
      setRowHeightPt(ws, r, 18);
      applyCellStyle(ws, r, 0, {
        font: { name: 'Times New Roman', sz: 9.5, bold: true, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      });
      continue;
    }
    if (mode === 'notes_head' && !col1) {
      addWorksheetMerge(ws, r, 0, r, colCount - 1);
      setRowHeightPt(ws, r, 24);
      applyCellStyle(ws, r, 0, {
        font: { name: 'Times New Roman', sz: 8.5, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      });
      continue;
    }
    if (mode === 'notes_head' && col1) {
      mode = 'note_tbl';
      noteCols = 1;
      for (let c = 1; c < colCount; c++) {
        if (getRowStringVal(ws, r, c) !== '') noteCols = c + 1;
      }
      setRowHeightPt(ws, r, 19);
      for (let c = 0; c < noteCols; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Times New Roman', sz: 8.5, bold: true, color: { rgb: '0F172A' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'F8FAFC' } },
          alignment: { horizontal: c === 0 ? 'left' : 'right', vertical: 'center' },
          border: FULL_THIN_SLATE,
        });
      }
      continue;
    }
    if (mode === 'note_tbl') {
      const nextCol0 = r + 1 < totalRows ? getRowStringVal(ws, r + 1, 0).trim() : '';
      const isTotal = col0.toLowerCase().startsWith('total') || nextCol0 === '';
      setRowHeightPt(ws, r, 17);
      for (let c = 0; c < noteCols; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Times New Roman', sz: 8.5, bold: isTotal, color: { rgb: '0F172A' } },
          alignment: { horizontal: c === 0 ? 'left' : 'right', vertical: 'center' },
          border: {
            top: BORDER_THIN_SLATE,
            bottom: isTotal ? BORDER_DOUBLE_DARK : BORDER_THIN_SLATE,
            left: BORDER_THIN_SLATE,
            right: BORDER_THIN_SLATE,
          },
        });
      }
    }
  }
}

function styleProjectedStatementSheet(
  ws: XLSX.WorkSheet,
  kind: 'bs' | 'pl' | 'cf',
  colCount: number,
  boldLabels: Set<string>,
  ruleLabels: Set<string>,
  doubleBottomLabels: Set<string>
) {
  setWorksheetPrintSetup(ws, 'landscape', 10, 1);
  const totalRows = getSheetRowCount(ws);
  if (totalRows === 0) return;

  const titleRowCount = kind === 'cf' ? 4 : 3;
  for (let r = 0; r < titleRowCount; r++) {
    if (colCount > 1) addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 20 : 16);
    const isAddr = kind === 'cf' && r === 1;
    const isSub = r === titleRowCount - 1;
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: {
          name: 'Helvetica',
          sz: r === 0 ? 11.5 : isAddr ? 8 : isSub ? 8.5 : 10.5,
          bold: r === 0 || (!isAddr && !isSub),
          color: { rgb: '0F172A' },
        },
        alignment: { horizontal: 'center', vertical: 'center' },
      });
    }
  }

  if (kind === 'bs' || kind === 'pl') {
    // 2-row merged header on Excel Rows 5..6 (0-based 4..5) matching PDF twoRowHead
    addWorksheetMerge(ws, 4, 0, 5, 0);
    addWorksheetMerge(ws, 4, 1, 5, 1);
    setRowHeightPt(ws, 4, 17);
    setRowHeightPt(ws, 5, 17);
    for (let r = 4; r <= 5; r++) {
      for (let c = 0; c < colCount; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: '0F172A' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: {
            top: BORDER_THIN_GRID,
            bottom: BORDER_THIN_GRID,
            left: BORDER_THIN_GRID,
            right: BORDER_THIN_GRID,
          },
        });
      }
    }
    setRowHeightPt(ws, 6, 7);
  } else {
    // Cash flow header on Row 6 (0-based 5) - unmerged so year columns stay distinct
    setRowHeightPt(ws, 5, 20);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, 5, c, {
        font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: '0F172A' } },
        fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
        alignment: { horizontal: c === 0 ? 'left' : 'right', vertical: 'center' },
        border: { top: BORDER_THIN_GRID, bottom: BORDER_THIN_GRID, left: BORDER_THIN_GRID, right: BORDER_THIN_GRID },
      });
    }
  }

  const startDataR = kind === 'cf' ? 6 : 7;
  const numStartCol = kind === 'cf' ? 1 : 2;
  for (let r = startDataR; r < totalRows; r++) {
    const label = getRowStringVal(ws, r, 0).trim();
    const col2Text = getRowStringVal(ws, r, 2).trim();
    if (!label && !col2Text) {
      setRowHeightPt(ws, r, 9);
      continue;
    }
    if (label === 'Proprietor') {
      setRowHeightPt(ws, r, 20);
      applyCellStyle(ws, r, 0, {
        font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'bottom' },
        border: { top: BORDER_THIN_DARK },
      });
      if (colCount > 3) addWorksheetMerge(ws, r, 2, r, colCount - 1);
      applyCellStyle(ws, r, 2, {
        font: { name: 'Helvetica', sz: 8.5, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'center', vertical: 'bottom' },
      });
      continue;
    }

    const isBold = boldLabels.has(label);
    const hasRule = ruleLabels.has(label);
    const hasDouble = doubleBottomLabels.has(label);
    setRowHeightPt(ws, r, hasDouble ? 19 : 16.5);
    for (let c = 0; c < colCount; c++) {
      const col0Align: 'left' | 'center' | 'right' =
        c === 0
          ? label === 'Total Assets' || label === 'Total Equity & Liabilities'
            ? 'right'
            : 'left'
          : kind !== 'cf' && c === 1
          ? 'center'
          : 'right';

      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: 8.5, bold: isBold, color: { rgb: '0F172A' } },
        alignment: { horizontal: col0Align, vertical: 'center' },
        border:
          c >= numStartCol && (hasRule || hasDouble)
            ? {
                top: BORDER_THIN_DARK,
                bottom: hasDouble ? BORDER_DOUBLE_DARK : BORDER_THIN_DARK,
              }
            : undefined,
      });
    }
  }
}

function styleProjectedNotesSheet(ws: XLSX.WorkSheet, colCount: number, noteBoldSet: Set<string>) {
  setWorksheetPrintSetup(ws, 'landscape', 12, 1);
  const totalRows = getSheetRowCount(ws);
  if (totalRows === 0) return;

  for (let r = 0; r <= 3; r++) {
    addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 20 : 16);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: {
          name: 'Helvetica',
          sz: r === 0 ? 11.5 : r === 1 ? 8 : r === 2 ? 10.5 : 8.5,
          bold: r === 0 || r === 2,
          color: { rgb: '0F172A' },
        },
        alignment: { horizontal: 'center', vertical: 'center' },
      });
    }
  }

  // Narrative Notes Header (row 5) & rows 6..14
  addWorksheetMerge(ws, 5, 1, 5, colCount - 1);
  setRowHeightPt(ws, 5, 19);
  for (let c = 0; c < colCount; c++) {
    applyCellStyle(ws, 5, c, {
      font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: '0F172A' } },
      alignment: { horizontal: c === 0 ? 'center' : 'left', vertical: 'center' },
      border: { top: BORDER_THIN_GRID, bottom: BORDER_THIN_GRID },
    });
  }
  for (let r = 6; r <= 14; r++) {
    addWorksheetMerge(ws, r, 1, r, colCount - 1);
    const no = getRowStringVal(ws, r, 0).trim();
    const isHdr = no === '1' || no === '2';
    setRowHeightPt(ws, r, isHdr ? 18 : 24);
    applyCellStyle(ws, r, 0, {
      font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: '0F172A' } },
      alignment: { horizontal: 'center', vertical: 'center' },
    });
    applyCellStyle(ws, r, 1, {
      font: { name: 'Helvetica', sz: 8.5, bold: isHdr || no === '', color: { rgb: '0F172A' } },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    });
  }

  // Quantitative Notes 2-row merged header on rows 16..17
  addWorksheetMerge(ws, 16, 0, 17, 0);
  addWorksheetMerge(ws, 16, 1, 17, 1);
  setRowHeightPt(ws, 16, 17);
  setRowHeightPt(ws, 17, 17);
  for (let r = 16; r <= 17; r++) {
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: '0F172A' } },
        fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: FULL_THIN_GRID,
      });
    }
  }

  for (let r = 18; r < totalRows; r++) {
    const noteNo = getRowStringVal(ws, r, 0).trim();
    const label = getRowStringVal(ws, r, 1).trim();
    const isBold = noteNo !== '' || noteBoldSet.has(label) || label === 'Add: Direct Expenses';
    const hasRule = noteBoldSet.has(label);
    setRowHeightPt(ws, r, 16.5);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: 8.5, bold: isBold, color: { rgb: '0F172A' } },
        alignment: { horizontal: c === 0 ? 'center' : c === 1 ? 'left' : 'right', vertical: 'center' },
        border: c >= 2 && hasRule ? { top: BORDER_THIN_DARK, bottom: BORDER_DOUBLE_DARK } : undefined,
      });
    }
  }
}

function styleProjectedScheduleASheet(ws: XLSX.WorkSheet, headerTopRows: number[]) {
  setWorksheetPrintSetup(ws, 'landscape', 10, 1);
  const totalRows = getSheetRowCount(ws);
  const colCount = 9;
  for (let r = 0; r <= 3; r++) {
    addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 20 : 16);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: {
          name: 'Helvetica',
          sz: r === 0 ? 11.5 : r === 1 ? 8 : r === 2 ? 10.5 : 8.5,
          bold: r === 0 || r === 2,
          color: { rgb: '0F172A' },
        },
        alignment: { horizontal: 'center', vertical: 'center' },
      });
    }
  }
  for (const rTop of headerTopRows) {
    const rSub = rTop + 1;
    // Merge 2-row spans: SL No. (0), Particulars (1), Cost across 2..5, Rate (6), Dep (7), WDV (8)
    addWorksheetMerge(ws, rTop, 0, rSub, 0);
    addWorksheetMerge(ws, rTop, 1, rSub, 1);
    addWorksheetMerge(ws, rTop, 2, rTop, 5);
    addWorksheetMerge(ws, rTop, 6, rSub, 6);
    addWorksheetMerge(ws, rTop, 7, rSub, 7);
    addWorksheetMerge(ws, rTop, 8, rSub, 8);
    setRowHeightPt(ws, rTop, 18);
    setRowHeightPt(ws, rSub, 22);
    for (let r = rTop; r <= rSub; r++) {
      for (let c = 0; c < colCount; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Helvetica', sz: 8, bold: true, color: { rgb: '0F172A' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: FULL_THIN_GRID,
        });
      }
    }
    const itemR = rSub + 1;
    const totR = rSub + 2;
    for (const r of [itemR, totR]) {
      if (r >= totalRows) continue;
      const isTot = r === totR;
      setRowHeightPt(ws, r, 18);
      for (let c = 0; c < colCount; c++) {
        applyCellStyle(ws, r, c, {
          font: { name: 'Helvetica', sz: 8, bold: isTot, color: { rgb: '0F172A' } },
          alignment: {
            horizontal: c === 0 || c === 6 ? 'center' : c === 1 ? 'left' : 'right',
            vertical: 'center',
          },
          border: {
            top: BORDER_THIN_GRID,
            bottom: isTot ? BORDER_DOUBLE_DARK : BORDER_THIN_GRID,
            left: BORDER_THIN_GRID,
            right: BORDER_THIN_GRID,
          },
        });
      }
    }
  }
}

function styleRepaymentScheduleSheet(ws: XLSX.WorkSheet) {
  setWorksheetPrintSetup(ws, 'landscape', 14, 1);
  const totalRows = getSheetRowCount(ws);
  const colCount = 9;
  if (totalRows === 0) return;

  for (let r = 0; r <= 1; r++) {
    addWorksheetMerge(ws, r, 0, r, colCount - 1);
    setRowHeightPt(ws, r, r === 0 ? 24 : 18);
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: r === 0 ? 12.5 : 9.5, bold: true, color: { rgb: r === 0 ? 'FFFFFF' : 'CBD5E1' } },
        fill: { patternType: 'solid', fgColor: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      });
    }
  }
  // Metadata rows 2..5
  for (let r = 2; r <= 5; r++) {
    addWorksheetMerge(ws, r, 1, r, 4);
    addWorksheetMerge(ws, r, 6, r, 8);
    setRowHeightPt(ws, r, 18);
    for (let c = 0; c < colCount; c++) {
      const isLabel = c === 0 || c === 5;
      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: 8.5, bold: isLabel, color: { rgb: '0F172A' } },
        fill: { patternType: 'solid', fgColor: { rgb: isLabel ? 'F8FAFC' : 'FFFFFF' } },
        alignment: { horizontal: isLabel ? 'left' : 'right', vertical: 'center' },
        border: FULL_THIN_GRID,
      });
    }
  }
  setRowHeightPt(ws, 6, 8);

  // Table header row 7
  setRowHeightPt(ws, 7, 22);
  for (let c = 0; c < colCount; c++) {
    applyCellStyle(ws, 7, c, {
      font: { name: 'Helvetica', sz: 8.5, bold: true, color: { rgb: 'FFFFFF' } },
      fill: { patternType: 'solid', fgColor: { rgb: '1E293B' } },
      alignment: { horizontal: c === 0 ? 'left' : c === 1 ? 'center' : 'right', vertical: 'center', wrapText: true },
      border: FULL_THIN_GRID,
    });
  }

  let zebra = 0;
  for (let r = 8; r < totalRows; r++) {
    const label = getRowStringVal(ws, r, 0).trim();
    if (!label) {
      setRowHeightPt(ws, r, 9);
      continue;
    }
    if (label.startsWith('Prepared By :')) {
      addWorksheetMerge(ws, r, 0, r, colCount - 1);
      setRowHeightPt(ws, r, 18);
      applyCellStyle(ws, r, 0, {
        font: { name: 'Helvetica', sz: 8, italic: true, color: { rgb: '475569' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      });
      continue;
    }
    const isTotal = label === 'Total :';
    setRowHeightPt(ws, r, isTotal ? 20 : 17);
    const fillHex = isTotal ? 'ECFDF5' : zebra % 2 === 1 ? 'F8FAFC' : 'FFFFFF';
    zebra++;
    for (let c = 0; c < colCount; c++) {
      applyCellStyle(ws, r, c, {
        font: { name: 'Helvetica', sz: 8.5, bold: isTotal, color: { rgb: '0F172A' } },
        fill: { patternType: 'solid', fgColor: { rgb: fillHex } },
        alignment: { horizontal: c === 0 ? 'left' : c === 1 ? 'center' : 'right', vertical: 'center' },
        border: isTotal
          ? { top: BORDER_THIN_DARK, bottom: BORDER_DOUBLE_DARK, left: BORDER_THIN_SLATE, right: BORDER_THIN_SLATE }
          : FULL_THIN_SLATE,
      });
    }
  }
}

function createTableDataForStatement(
  title: string,
  company: CompanyProfile,
  headers: string[],
  rows: ExcelCellValue[][]
): ExcelCellValue[][] {
  return [
    [safeCompanyUpper(company)],
    [`${title.toUpperCase()} (${company.reportingCurrency} in ${company.reportingUnit})`],
    [`Accounting Framework: ${company.accountingFramework} | Entity Type: ${company.entityType.toUpperCase()} | FY End: ${normalizeFinancialYearEnd(company.financialYearEnd)}`],
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
  styleSection76Sheet(ws, headers.length);
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
  styleSection76Sheet(ws, headers.length);
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
  styleSection76Sheet(ws, headers.length);
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
  styleSection76Sheet(ws, headers.length);
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
    styleSection76Sheet(ws, headers.length, { centerFirstCols: 2 });
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
  styleSection76Sheet(ws, headers.length, { centerFirstCols: 2 });
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
  writeWorkbookStyled(wb, filename);
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
  writeWorkbookStyled(wb, filename);
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
  writeWorkbookStyled(wb, filename);
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
  writeWorkbookStyled(wb, filename);
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
  writeWorkbookStyled(wb, filename);
}

export function exportNotesExcel(
  company: CompanyProfile,
  notes: FinancialNoteItem[],
  filename = 'Notes_to_the_Financial_Statements.xlsx'
) {
  const wb = XLSX.utils.book_new();
  const rows: (string | number)[][] = [
    [safeCompanyUpper(company)],
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
  styleSection76NotesSheet(ws, 4);
  XLSX.utils.book_append_sheet(wb, ws, 'Statutory Notes');
  writeWorkbookStyled(wb, filename);
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
    'Rescheduling Assessment',
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
      d.reschedulingContext || '',
    ];
  });
  const sheetData = createTableDataForStatement('Debt Servicing & Repayment Capacity (DSCR)', company, headers, rows);
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [10, 12, 16, 16, 26, 26, 20, 14, 14, 18, 16, 48]);
  styleSection76Sheet(ws, headers.length, { centerFirstCols: 2, centerAllHeaders: true });
  return ws;
}

function buildRepaymentScheduleMetadataRows(schedule: ParsedRepaymentSchedule): (string | number)[][] {
  const repaymentCount = schedule.rows.filter(r => !r.isDisbursement).length;
  const firstRepRow = schedule.rows.find(r => !r.isDisbursement);
  const lastRepRow = [...schedule.rows].reverse().find(r => !r.isDisbursement);
  const invAmt = schedule.metadata.investmentMade ?? schedule.grandTotalPrincipal;
  return [
    [
      'Investment Made :',
      invAmt,
      '',
      '',
      '',
      'Disbursement Date :',
      schedule.metadata.disbursementDate || schedule.rows[0]?.emiDateRaw || '',
      '',
      '',
    ],
    [
      'Interval :',
      schedule.metadata.interval || 'Monthly',
      '',
      '',
      '',
      '1st Installment Date :',
      schedule.metadata.firstInstallmentDate || firstRepRow?.emiDateRaw || '',
      '',
      '',
    ],
    [
      'Number of Installment :',
      schedule.metadata.numberOfInstallments ?? repaymentCount,
      '',
      '',
      '',
      'Expiry Date :',
      schedule.metadata.expiryDate || lastRepRow?.emiDateRaw || '',
      '',
      '',
    ],
    [
      'Gestation Period :',
      schedule.metadata.gestationPeriod || '0 Days',
      '',
      '',
      '',
      'Last Installment size minimized :',
      schedule.metadata.lastInstallmentMinimized || 'No',
      '',
      '',
    ],
  ];
}

function buildRepaymentScheduleWorksheet(
  company: CompanyProfile,
  schedule: ParsedRepaymentSchedule
): XLSX.WorkSheet {
  const metaRows = buildRepaymentScheduleMetadataRows(schedule);
  const headers = [
    'Event Type.',
    'EMI Date',
    'Inst. Principal',
    'Inst. Profit',
    'Gestation Payment',
    'Over due Gest.',
    'Overdue Profit',
    'Balance',
    'Total',
  ];
  const sheetRows: ExcelCellValue[][] = [
    [safeCompanyUpper(company)],
    ['Repayment Schedule'],
    ...metaRows.map(mr =>
      mr.map((cell, idx) => (idx === 1 && typeof cell === 'number' ? nCell(cell) : cell))
    ),
    [],
    headers,
  ];

  const dataStartRow = sheetRows.length + 1;
  schedule.rows.forEach((row, idx) => {
    const currRowNum = dataStartRow + idx;
    const prevRowNum = currRowNum - 1;
    const isFirstDisb = idx === 0 || Boolean(row.isDisbursement);
    const balCell: ExcelCellValue = isFirstDisb
      ? nCell(row.balance)
      : fCell(`MAX(0,H${prevRowNum}-C${currRowNum})`, row.balance);
    const totCell: ExcelCellValue = row.isDisbursement
      ? nCell(0)
      : fCell(`SUM(C${currRowNum}:G${currRowNum})`, row.totalInstallment);

    sheetRows.push([
      row.eventType,
      row.emiDateRaw,
      nCell(row.instPrincipal),
      nCell(row.instProfit),
      nCell(row.gestationPayment || 0),
      nCell(row.overdueGestation || 0),
      nCell(row.overdueProfit || 0),
      balCell,
      totCell,
    ]);
  });

  if (schedule.rows.length > 0) {
    const endRow = dataStartRow + schedule.rows.length - 1;
    sheetRows.push([
      'Total :',
      '',
      fCell(`SUM(C${dataStartRow}:C${endRow})`, schedule.grandTotalPrincipal),
      fCell(`SUM(D${dataStartRow}:D${endRow})`, schedule.grandTotalProfit),
      fCell(`SUM(E${dataStartRow}:E${endRow})`, schedule.grandTotalGestationPayment ?? 0),
      fCell(`SUM(F${dataStartRow}:F${endRow})`, schedule.grandTotalOverdueGestation ?? 0),
      fCell(`SUM(G${dataStartRow}:G${endRow})`, schedule.grandTotalOverdueProfit ?? 0),
      nCell(schedule.closingBalance ?? 0),
      fCell(`SUM(I${dataStartRow}:I${endRow})`, schedule.grandTotalDebtService),
    ]);
  }

  sheetRows.push([]);
  sheetRows.push([
    schedule.metadata.preparedByNote ||
      'Prepared By : ABABIL, Developed By : Millennium Information Solution Limited',
  ]);

  const ws = XLSX.utils.aoa_to_sheet(sheetRows);
  applyColumnWidths(ws, [20, 14, 18, 16, 18, 16, 16, 18, 18]);
  styleRepaymentScheduleSheet(ws);
  return ws;
}

function buildAnnualScheduleSummaryWorksheet(
  company: CompanyProfile,
  schedule: ParsedRepaymentSchedule
): XLSX.WorkSheet {
  const headers = [
    'Financial Year',
    'Installments',
    'Inst. Principal',
    'Inst. Profit',
    'Gestation Payment',
    'Over due Gest.',
    'Overdue Profit',
    'Total Debt Service',
    'Closing Balance',
  ];
  const r = STD_DATA_START_ROW;
  const rows: ExcelCellValue[][] = schedule.annualSummaries.map((s, idx) => {
    const rowNum = r + idx;
    return [
      `FY ${s.financialYear}`,
      s.installmentCount,
      nCell(s.totalPrincipal),
      nCell(s.totalProfit),
      nCell(s.totalGestationPayment ?? 0),
      nCell(s.totalOverdueGestation ?? 0),
      nCell(s.totalOverdueProfit ?? 0),
      fCell(`SUM(C${rowNum}:G${rowNum})`, s.totalDebtService),
      nCell(s.closingBalance),
    ];
  });
  if (schedule.annualSummaries.length > 0) {
    const endRow = r + schedule.annualSummaries.length - 1;
    rows.push([
      'Total :',
      fCell(`SUM(B${r}:B${endRow})`, schedule.annualSummaries.reduce((sum, a) => sum + a.installmentCount, 0)),
      fCell(`SUM(C${r}:C${endRow})`, schedule.grandTotalPrincipal),
      fCell(`SUM(D${r}:D${endRow})`, schedule.grandTotalProfit),
      fCell(`SUM(E${r}:E${endRow})`, schedule.grandTotalGestationPayment ?? 0),
      fCell(`SUM(F${r}:F${endRow})`, schedule.grandTotalOverdueGestation ?? 0),
      fCell(`SUM(G${r}:G${endRow})`, schedule.grandTotalOverdueProfit ?? 0),
      fCell(`SUM(H${r}:H${endRow})`, schedule.grandTotalDebtService),
      nCell(schedule.closingBalance ?? 0),
    ]);
  }
  const sheetData = createTableDataForStatement(
    'Annualized Repayment Schedule Summary',
    company,
    headers,
    rows
  );
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  applyColumnWidths(ws, [16, 14, 18, 16, 18, 16, 16, 18, 18]);
  styleSection76Sheet(ws, headers.length, { centerFirstCols: 2 });
  return ws;
}

export function exportRepaymentScheduleExcel(
  company: CompanyProfile,
  schedule: ParsedRepaymentSchedule,
  dscrData?: DebtServiceCoverageYear[],
  filename = `${sanitizeFilePrefix(company.companyName)}_Repayment_Schedule.xlsx`
) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildRepaymentScheduleWorksheet(company, schedule), '1. Repayment Schedule');
  XLSX.utils.book_append_sheet(wb, buildAnnualScheduleSummaryWorksheet(company, schedule), '2. Annual Debt Summary');
  const activeDscr = (dscrData || []).filter(d => d.totalDebtService > 0 || d.ebitda !== 0);
  if (activeDscr.length > 0) {
    XLSX.utils.book_append_sheet(wb, buildDebtServiceCoverageWorksheet(company, activeDscr), '3. DSCR Justification');
  }
  writeWorkbookStyled(wb, filename);
}

export async function exportRepaymentScheduleWord(
  company: CompanyProfile,
  schedule: ParsedRepaymentSchedule,
  dscrData?: DebtServiceCoverageYear[],
  filename = `${sanitizeFilePrefix(company.companyName)}_Repayment_Schedule.docx`
) {
  const usableW = docxUsableWidthTwips('landscape', 14);
  const metaRows = buildRepaymentScheduleMetadataRows(schedule);
  const schedHeaders = [
    'Event Type.',
    'EMI Date',
    'Inst. Principal',
    'Inst. Profit',
    'Gestation Payment',
    'Over due Gest.',
    'Overdue Profit',
    'Balance',
    'Total',
  ];
  const schedRows: (string | number)[][] = [
    ...schedule.rows.map(row => [
      row.eventType,
      row.emiDateRaw,
      row.instPrincipal,
      row.instProfit,
      row.gestationPayment || 0,
      row.overdueGestation || 0,
      row.overdueProfit || 0,
      row.balance,
      row.totalInstallment,
    ]),
    [
      'Total :',
      '',
      schedule.grandTotalPrincipal,
      schedule.grandTotalProfit,
      schedule.grandTotalGestationPayment ?? 0,
      schedule.grandTotalOverdueGestation ?? 0,
      schedule.grandTotalOverdueProfit ?? 0,
      schedule.closingBalance ?? 0,
      schedule.grandTotalDebtService,
    ],
  ];

  const activeDscr = (dscrData || []).filter(d => d.totalDebtService > 0);
  const doc = new Document({
    sections: [
      {
        properties: docxPageProps('landscape', 14),
        footers: {
          default: docxFooter('Repayment Schedule & DSCR Justification'),
        },
        children: [
          docxSectionHeaderBanner(company, 'REPAYMENT SCHEDULE (EXACT AUTOMATED INSTALLMENT LEDGER)', usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          createDocxTable(
            ['Facility Parameter', 'Value', 'Schedule Dates', 'Date / Setting'],
            metaRows.map(mr => [
              String(mr[0]),
              typeof mr[1] === 'number' ? fmtExportNum(mr[1]) : String(mr[1]),
              String(mr[5]),
              String(mr[6]),
            ]),
            {
              theme: 'previous_grid',
              font: 'Helvetica',
              fontSizePt: 7.5,
              colWidthsTwips: [
                convertMillimetersToTwip(46),
                convertMillimetersToTwip(45),
                convertMillimetersToTwip(56),
                convertMillimetersToTwip(45),
              ],
              customAligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.LEFT, AlignmentType.RIGHT],
              boldColIndices: [0, 2],
              hideHeader: true,
            }
          ),
          new Paragraph({ spacing: { before: 80, after: 60 }, children: [] }),
          createDocxTable(schedHeaders, schedRows, {
            theme: 'section76',
            font: 'Helvetica',
            fontSizePt: 7.2,
            headFontSizePt: 7.2,
            totalWidthTwips: usableW,
            colWidthsTwips: [
              convertMillimetersToTwip(32),
              convertMillimetersToTwip(22),
              convertMillimetersToTwip(31),
              convertMillimetersToTwip(30),
              convertMillimetersToTwip(31),
              convertMillimetersToTwip(29),
              convertMillimetersToTwip(29),
              convertMillimetersToTwip(33),
              convertMillimetersToTwip(32),
            ],
            customAligns: [
              AlignmentType.LEFT,
              AlignmentType.CENTER,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
            ],
            doubleBottomLabels: new Set(['Total :']),
          }),
          new Paragraph({
            spacing: { before: 80, after: 100 },
            children: [
              new TextRun({
                text:
                  schedule.metadata.preparedByNote ||
                  'Prepared By : ABABIL, Developed By : Millennium Information Solution Limited',
                font: 'Helvetica',
                size: 15,
                italics: true,
                color: '475569',
              }),
            ],
          }),
          new Paragraph({ children: [new PageBreak()] }),
          docxSectionHeaderBanner(company, 'ANNUALIZED REPAYMENT SCHEDULE AGGREGATION', usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          createDocxTable(
            ['Financial Year', 'Installments', 'Inst. Principal', 'Inst. Profit', 'Total Debt Service', 'Closing Balance'],
            [
              ...schedule.annualSummaries.map(s => [
                `FY ${s.financialYear}`,
                s.installmentCount,
                s.totalPrincipal,
                s.totalProfit,
                s.totalDebtService,
                s.closingBalance,
              ]),
              [
                'Total :',
                schedule.annualSummaries.reduce((sum, a) => sum + a.installmentCount, 0),
                schedule.grandTotalPrincipal,
                schedule.grandTotalProfit,
                schedule.grandTotalDebtService,
                schedule.closingBalance ?? 0,
              ],
            ],
            {
              theme: 'section76',
              font: 'Helvetica',
              fontSizePt: 8,
              totalWidthTwips: usableW,
              customAligns: [
                AlignmentType.LEFT,
                AlignmentType.CENTER,
                AlignmentType.RIGHT,
                AlignmentType.RIGHT,
                AlignmentType.RIGHT,
                AlignmentType.RIGHT,
              ],
              doubleBottomLabels: new Set(['Total :']),
            }
          ),
          ...(activeDscr.length > 0
            ? [
                new Paragraph({ children: [new PageBreak()] }),
                docxSectionHeaderBanner(
                  company,
                  'DEBT SERVICING & RESCHEDULING JUSTIFICATION (HISTORICAL VS PROJECTED DSCR)',
                  usableW
                ),
                new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
                createDocxTable(
                  ['Year', 'EBITDA', 'Pre-Fin OCF', 'Inst. Principal', 'Finance Cost', 'Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status', 'Assessment'],
                  activeDscr.map(d => [
                    `${d.year}${d.isProjected ? ' (P)' : ' (H)'}`,
                    d.ebitda,
                    d.operatingCashFlowBeforeFinance,
                    d.scheduledPrincipal,
                    d.scheduledFinanceCost,
                    d.totalDebtService,
                    `${d.ebitdaDscr.toFixed(2)}x`,
                    `${d.cashDscr.toFixed(2)}x`,
                    d.closingCash,
                    d.capacityStatus,
                    d.reschedulingContext || '',
                  ]),
                  {
                    theme: 'section76',
                    font: 'Helvetica',
                    fontSizePt: 7.2,
                    headFontSizePt: 7.2,
                    totalWidthTwips: usableW,
                    leftColsCount: 1,
                    customAligns: [
                      AlignmentType.LEFT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.CENTER,
                      AlignmentType.LEFT,
                    ],
                  }
                ),
              ]
            : []),
        ],
      },
    ],
  });

  await emitDocxDocument(doc, filename);
}

export function exportRepaymentSchedulePdf(
  company: CompanyProfile,
  schedule: ParsedRepaymentSchedule,
  dscrData?: DebtServiceCoverageYear[],
  filename = `${sanitizeFilePrefix(company.companyName)}_Repayment_Schedule.pdf`
) {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  renderPdfHeader(doc, company, 'REPAYMENT SCHEDULE (EXACT AUTOMATED INSTALLMENT LEDGER)');
  const metaRows = buildRepaymentScheduleMetadataRows(schedule);

  autoTable(doc, {
    body: metaRows.map(mr => [
      String(mr[0]),
      typeof mr[1] === 'number' ? fmtExportNum(mr[1]) : String(mr[1]),
      String(mr[5]),
      String(mr[6]),
    ]),
    startY: 26,
    theme: 'grid',
    styles: { fontSize: 7.5, cellPadding: 1.5, textColor: [15, 23, 42] },
    columnStyles: {
      0: { fontStyle: 'bold', cellWidth: 46 },
      1: { halign: 'right', cellWidth: 45 },
      2: { fontStyle: 'bold', cellWidth: 56 },
      3: { halign: 'right', cellWidth: 45 },
    },
    margin: { left: 14, right: 14 },
  });

  const metaEndY = ((doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 50) + 4;
  const schedBody: string[][] = [
    ...schedule.rows.map(row => [
      row.eventType,
      row.emiDateRaw,
      fmtExportNum(row.instPrincipal),
      fmtExportNum(row.instProfit),
      fmtExportNum(row.gestationPayment || 0),
      fmtExportNum(row.overdueGestation || 0),
      fmtExportNum(row.overdueProfit || 0),
      fmtExportNum(row.balance),
      fmtExportNum(row.totalInstallment),
    ]),
    [
      'Total :',
      '',
      fmtExportNum(schedule.grandTotalPrincipal),
      fmtExportNum(schedule.grandTotalProfit),
      fmtExportNum(schedule.grandTotalGestationPayment ?? 0),
      fmtExportNum(schedule.grandTotalOverdueGestation ?? 0),
      fmtExportNum(schedule.grandTotalOverdueProfit ?? 0),
      fmtExportNum(schedule.closingBalance ?? 0),
      fmtExportNum(schedule.grandTotalDebtService),
    ],
  ];

  autoTable(doc, {
    head: [['Event Type.', 'EMI Date', 'Inst. Principal', 'Inst. Profit', 'Gestation Payment', 'Over due Gest.', 'Overdue Profit', 'Balance', 'Total']],
    body: schedBody,
    startY: metaEndY,
    styles: { fontSize: 7.2, cellPadding: 1.4, font: 'helvetica' },
    headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    columnStyles: {
      0: { halign: 'left' },
      1: { halign: 'center' },
      2: { halign: 'right' },
      3: { halign: 'right' },
      4: { halign: 'right' },
      5: { halign: 'right' },
      6: { halign: 'right' },
      7: { halign: 'right' },
      8: { halign: 'right' },
    },
    didParseCell: (data) => {
      if (data.section === 'body' && data.row.index === schedBody.length - 1) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.fillColor = [236, 253, 245];
      }
    },
    margin: { left: 14, right: 14 },
  });

  const endSchedY = ((doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 180) + 5;
  if (endSchedY < 194) {
    doc.setFontSize(7.5);
    doc.setTextColor(71, 85, 105);
    doc.text(
      schedule.metadata.preparedByNote ||
        'Prepared By : ABABIL, Developed By : Millennium Information Solution Limited',
      14,
      endSchedY
    );
  }

  const activeDscr = (dscrData || []).filter(d => d.totalDebtService > 0);
  if (activeDscr.length > 0) {
    doc.addPage();
    renderPdfHeader(doc, company, 'DEBT SERVICING & RESCHEDULING JUSTIFICATION (HISTORICAL VS PROJECTED DSCR)');
    autoTable(doc, {
      head: [['Year', 'EBITDA', 'Pre-Fin OCF', 'Inst. Principal', 'Finance Cost', 'Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status', 'Assessment']],
      body: activeDscr.map(d => [
        `${d.year}${d.isProjected ? ' (P)' : ' (H)'}`,
        fmtExportNum(d.ebitda),
        fmtExportNum(d.operatingCashFlowBeforeFinance),
        fmtExportNum(d.scheduledPrincipal),
        fmtExportNum(d.scheduledFinanceCost),
        fmtExportNum(d.totalDebtService),
        `${d.ebitdaDscr.toFixed(2)}x`,
        `${d.cashDscr.toFixed(2)}x`,
        fmtExportNum(d.closingCash),
        d.capacityStatus,
        d.reschedulingContext || '',
      ]),
      startY: 28,
      styles: { fontSize: 7.2, cellPadding: 1.6 },
      headStyles: { fillColor: [30, 41, 59] },
    });
  }

  const totalPages = doc.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(148, 163, 184);
    doc.text(`Financials with Tuhin · Page ${i} of ${totalPages}`, 14, 210 - 8);
    doc.text(`Repayment Schedule & DSCR Justification`, 297 - 14, 210 - 8, { align: 'right' });
  }

  doc.save(filename);
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
  filename = `${sanitizeFilePrefix(company.companyName)}_Complete_Financial_Package.xlsx`,
  debtServiceCoverage?: DebtServiceCoverageYear[],
  options?: ExportSyncOptions,
  parsedSchedule?: ParsedRepaymentSchedule | null
) {
  const wb = XLSX.utils.book_new();

  const horizonText = incomeData.length > 0 ? `${incomeData[0].year} - ${incomeData[incomeData.length - 1].year}` : '';
  XLSX.utils.book_append_sheet(wb, buildSection76CoverWorksheet(company, horizonText), '0. Cover & Index');
  XLSX.utils.book_append_sheet(wb, buildIncomeStatementSheet(company, incomeData, 'Income Statement', options), '1. Income Statement');
  XLSX.utils.book_append_sheet(wb, buildBalanceSheetWorksheet(company, balanceData, 'Statement of Financial Position', options), '2. Balance Sheet');
  XLSX.utils.book_append_sheet(wb, buildCashFlowWorksheet(company, cashFlowData, 'Statement of Cash Flows', options), '3. Cash Flows');
  XLSX.utils.book_append_sheet(wb, buildEquityWorksheet(company, equityData, 'Statement of Changes in Equity', options), '4. Changes in Equity');
  XLSX.utils.book_append_sheet(wb, buildFixedAssetsWorksheet(company, fixedAssetsData, options), '5. Fixed Assets');

  // 6. Notes to Financial Statements Sheet
  const notesRows: (string | number)[][] = [
    [safeCompanyUpper(company)],
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
  styleSection76NotesSheet(wsNotes, 4);
  XLSX.utils.book_append_sheet(wb, wsNotes, '6. Notes to Financials');

  const activeDscr = (debtServiceCoverage || []).filter(d => d.totalDebtService > 0);
  if (activeDscr.length > 0) {
    XLSX.utils.book_append_sheet(wb, buildDebtServiceCoverageWorksheet(company, activeDscr), '7. Debt Servicing & DSCR');
  }

  if (parsedSchedule && parsedSchedule.rows.length > 0) {
    const schedTab = activeDscr.length > 0 ? '8. Repayment Schedule' : '7. Repayment Schedule';
    XLSX.utils.book_append_sheet(wb, buildRepaymentScheduleWorksheet(company, parsedSchedule), schedTab);
  }

  writeWorkbookStyled(wb, filename);
}

// ---------------- WORD EXPORTS (docx) ---------------- //

interface DocxMultiHeaderCell {
  content: string;
  rowSpan?: number;
  colSpan?: number;
  align?: typeof AlignmentType[keyof typeof AlignmentType];
}

interface DocxTableOptions {
  leftColsCount?: number;
  theme?: 'section76' | 'previous_grid' | 'projected_plain' | 'projected_grid';
  font?: 'Helvetica' | 'Times New Roman';
  fontSizePt?: number;
  headFontSizePt?: number;
  colWidthsTwips?: number[];
  totalWidthTwips?: number;
  boldRowLabels?: Set<string>;
  doubleBottomLabels?: Set<string>;
  ruleLabels?: Set<string>;
  rightAlignCol0Labels?: Set<string>;
  customAligns?: (typeof AlignmentType[keyof typeof AlignmentType])[];
  multiHeaderRows?: DocxMultiHeaderCell[][];
  labelColIndex?: number;
  boldColIndices?: number[];
  hideHeader?: boolean;
}

const DOCX_NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const DOCX_NONE_BORDERS = {
  top: DOCX_NO_BORDER,
  bottom: DOCX_NO_BORDER,
  left: DOCX_NO_BORDER,
  right: DOCX_NO_BORDER,
  insideHorizontal: DOCX_NO_BORDER,
  insideVertical: DOCX_NO_BORDER,
};

function docxPageProps(orientation: 'landscape' | 'portrait', marginMm = 14) {
  const m = convertMillimetersToTwip(marginMm);
  return {
    page: {
      size: {
        width: convertMillimetersToTwip(210),
        height: convertMillimetersToTwip(297),
        orientation: orientation === 'landscape' ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
      },
      margin: {
        top: m,
        bottom: m,
        left: m,
        right: m,
        header: convertMillimetersToTwip(6),
        footer: convertMillimetersToTwip(7),
      },
    },
  };
}

function docxUsableWidthTwips(orientation: 'landscape' | 'portrait', marginMm = 14): number {
  return convertMillimetersToTwip((orientation === 'landscape' ? 297 : 210) - 2 * marginMm);
}

function docxFooter(rightLabel: string): Footer {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            children: ['Financials with Tuhin · Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES, `   |   ${rightLabel}`],
            font: 'Helvetica',
            size: 15,
            color: '94A3B8',
          }),
        ],
      }),
    ],
  });
}

function docxSectionHeaderBanner(company: CompanyProfile, title: string, totalWidthTwips: number): Table {
  const colW = [Math.round(totalWidthTwips * 0.65), totalWidthTwips - Math.round(totalWidthTwips * 0.65)];
  return new Table({
    width: { size: totalWidthTwips, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: colW,
    borders: DOCX_NONE_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: colW[0], type: WidthType.DXA },
            columnSpan: 2,
            shading: { type: ShadingType.CLEAR, color: 'auto', fill: '0F172A' },
            margins: { top: 90, bottom: 25, left: 140, right: 140 },
            borders: { top: DOCX_NO_BORDER, bottom: DOCX_NO_BORDER, left: DOCX_NO_BORDER, right: DOCX_NO_BORDER },
            children: [
              new Paragraph({
                children: [new TextRun({ text: safeCompanyUpper(company), font: 'Helvetica', size: 22, bold: true, color: 'FFFFFF' })],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: colW[0], type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, color: 'auto', fill: '0F172A' },
            margins: { top: 20, bottom: 90, left: 140, right: 80 },
            borders: { top: DOCX_NO_BORDER, bottom: DOCX_NO_BORDER, left: DOCX_NO_BORDER, right: DOCX_NO_BORDER },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `${title} · ${company.reportingCurrency} (${company.reportingUnit}) · ${company.accountingFramework} · FY End: ${normalizeFinancialYearEnd(company.financialYearEnd)}`,
                    font: 'Helvetica',
                    size: 16,
                    color: 'CBD5E1',
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: colW[1], type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, color: 'auto', fill: '0F172A' },
            margins: { top: 20, bottom: 90, left: 80, right: 140 },
            borders: { top: DOCX_NO_BORDER, bottom: DOCX_NO_BORDER, left: DOCX_NO_BORDER, right: DOCX_NO_BORDER },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: 'Financials with Tuhin | Mohammad Towhidul Akbar Tuhin',
                    font: 'Helvetica',
                    size: 16,
                    color: 'CBD5E1',
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function docxCoverPage(company: CompanyProfile, horizonText: string, totalWidthTwips: number): Table {
  const blueBorder = { style: BorderStyle.SINGLE, size: 16, color: '3B82F6' };
  return new Table({
    width: { size: totalWidthTwips, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: [totalWidthTwips],
    borders: {
      top: blueBorder,
      bottom: blueBorder,
      left: blueBorder,
      right: blueBorder,
      insideHorizontal: DOCX_NO_BORDER,
      insideVertical: DOCX_NO_BORDER,
    },
    rows: [
      new TableRow({
        height: { value: convertMillimetersToTwip(186), rule: HeightRule.ATLEAST },
        children: [
          new TableCell({
            width: { size: totalWidthTwips, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            shading: { type: ShadingType.CLEAR, color: 'auto', fill: '0F172A' },
            margins: { top: 360, bottom: 360, left: 320, right: 320 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 180 },
                children: [new TextRun({ text: safeCompanyUpper(company), font: 'Helvetica', size: 52, bold: true, color: 'FFFFFF' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 160 },
                children: [new TextRun({ text: 'ANNUAL STATUTORY FINANCIAL STATEMENTS & AUDIT SCHEDULES', font: 'Helvetica', size: 32, bold: true, color: '93C5FD' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 80 },
                children: [new TextRun({ text: 'Five Standard Financial Statements + Complete Notes to the Financial Statements', font: 'Helvetica', size: 22, color: 'CBD5E1' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 280 },
                children: [new TextRun({ text: `Compliant with Section 76 & International Accounting Framework (${company.accountingFramework})`, font: 'Helvetica', size: 22, color: 'CBD5E1' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 },
                children: [new TextRun({ text: `Reporting Entity Type: ${company.entityType.toUpperCase()}`, font: 'Helvetica', size: 20, color: '94A3B8' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 },
                children: [new TextRun({ text: `Corporate Address: ${safeAddress(company)}`, font: 'Helvetica', size: 20, color: '94A3B8' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 260 },
                children: [new TextRun({ text: `Reporting Currency: ${company.reportingCurrency} (${company.reportingUnit}) | FY End: ${normalizeFinancialYearEnd(company.financialYearEnd)} | Horizon: ${horizonText}`, font: 'Helvetica', size: 20, color: '94A3B8' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 },
                children: [new TextRun({ text: `Lead Financial Architect: ${company.preparedBy}`, font: 'Helvetica', size: 20, bold: true, color: '34D399' })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: `Independent Statutory Auditors: ${company.auditorName}`, font: 'Helvetica', size: 20, color: '94A3B8' })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function docxProprietorSignatureBlock(usableWidthTwips: number): Table {
  const leftW = Math.round(usableWidthTwips * 0.35);
  const rightW = usableWidthTwips - leftW;
  return new Table({
    width: { size: usableWidthTwips, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: [leftW, rightW],
    borders: DOCX_NONE_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: leftW, type: WidthType.DXA },
            borders: { top: DOCX_NO_BORDER, bottom: DOCX_NO_BORDER, left: DOCX_NO_BORDER, right: DOCX_NO_BORDER },
            children: [
              new Paragraph({ children: [new TextRun({ text: '...........................', font: 'Helvetica', size: 17, bold: true, color: '0F172A' })] }),
              new Paragraph({ children: [new TextRun({ text: 'Proprietor', font: 'Helvetica', size: 17, bold: true, color: '0F172A' })] }),
            ],
          }),
          new TableCell({
            width: { size: rightW, type: WidthType.DXA },
            verticalAlign: VerticalAlign.BOTTOM,
            borders: { top: DOCX_NO_BORDER, bottom: DOCX_NO_BORDER, left: DOCX_NO_BORDER, right: DOCX_NO_BORDER },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: 'Signed in terms of our separate report of even date.', font: 'Helvetica', size: 17, color: '0F172A' })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function computeDocxColWidths(colCount: number, leftCols: number, totalWidthTwips: number): number[] {
  if (colCount <= 0) return [];
  if (colCount === 1) return [totalWidthTwips];
  if (leftCols === 2 && colCount >= 3) {
    const labelW = Math.round(totalWidthTwips * 0.28);
    const noteW = Math.round(totalWidthTwips * 0.08);
    const remW = Math.floor((totalWidthTwips - labelW - noteW) / (colCount - 2));
    const arr = [labelW, noteW, ...Array.from({ length: colCount - 2 }, () => remW)];
    arr[colCount - 1] += totalWidthTwips - arr.reduce((s, w) => s + w, 0);
    return arr;
  }
  const firstRatio = colCount >= 10 ? 0.22 : colCount >= 6 ? 0.30 : 0.48;
  const firstW = Math.round(totalWidthTwips * firstRatio);
  const otherW = Math.floor((totalWidthTwips - firstW) / (colCount - 1));
  const arr = [firstW, ...Array.from({ length: colCount - 1 }, () => otherW)];
  arr[colCount - 1] += totalWidthTwips - arr.reduce((s, w) => s + w, 0);
  return arr;
}

function createDocxTable(
  headers: string[],
  rows: (string | number)[][],
  leftColsCountOrOpts: number | DocxTableOptions = 1
): Table {
  const opts: DocxTableOptions =
    typeof leftColsCountOrOpts === 'number' ? { leftColsCount: leftColsCountOrOpts } : leftColsCountOrOpts;

  const theme = opts.theme ?? 'section76';
  const font = opts.font ?? (theme === 'previous_grid' ? 'Times New Roman' : 'Helvetica');
  const bodySize = Math.round((opts.fontSizePt ?? (theme === 'previous_grid' ? 8.5 : 8)) * 2);
  const headSize = Math.round((opts.headFontSizePt ?? opts.fontSizePt ?? (theme === 'previous_grid' ? 8.8 : 8)) * 2);
  const leftCols = opts.leftColsCount ?? 1;
  const labelColIdx = opts.labelColIndex ?? 0;
  const totalWidthTwips = opts.totalWidthTwips ?? docxUsableWidthTwips('landscape', 14);
  const colWidths =
    opts.colWidthsTwips && opts.colWidthsTwips.length === headers.length
      ? opts.colWidthsTwips
      : computeDocxColWidths(headers.length, leftCols, totalWidthTwips);

  const isCompact = (opts.fontSizePt ?? 8) <= 6.2;
  const pad = isCompact
    ? { top: 28, bottom: 28, left: 40, right: 40 }
    : { top: 48, bottom: 48, left: 72, right: 72 };

  const gridBorder = { style: BorderStyle.SINGLE, size: 4, color: '334155' };
  const slateBorder = { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1' };
  const darkSingleBorder = { style: BorderStyle.SINGLE, size: 6, color: '0F172A' };
  const darkDoubleBorder = { style: BorderStyle.DOUBLE, size: 6, color: '0F172A' };

  const tableBorders =
    theme === 'projected_plain'
      ? DOCX_NONE_BORDERS
      : theme === 'previous_grid' || theme === 'projected_grid'
      ? {
          top: gridBorder,
          bottom: gridBorder,
          left: gridBorder,
          right: gridBorder,
          insideHorizontal: gridBorder,
          insideVertical: gridBorder,
        }
      : {
          top: darkSingleBorder,
          bottom: darkSingleBorder,
          left: slateBorder,
          right: slateBorder,
          insideHorizontal: slateBorder,
          insideVertical: slateBorder,
        };

  const headFill = theme === 'section76' ? '1E293B' : 'FFFFFF';
  const headColor = theme === 'section76' ? 'FFFFFF' : '0F172A';
  const headBorderObj =
    theme === 'section76'
      ? { top: gridBorder, bottom: gridBorder, left: gridBorder, right: gridBorder }
      : { top: darkSingleBorder, bottom: darkSingleBorder, left: darkSingleBorder, right: darkSingleBorder };

  const headerTableRows: TableRow[] = opts.hideHeader
    ? []
    : opts.multiHeaderRows
    ? opts.multiHeaderRows.map(hdrRow => {
        let colCursor = 0;
        return new TableRow({
          tableHeader: true,
          children: hdrRow.map(cellSpec => {
            const span = cellSpec.colSpan ?? 1;
            let cellW = 0;
            for (let k = 0; k < span; k++) {
              cellW += colWidths[colCursor + k] ?? colWidths[colWidths.length - 1] ?? 1200;
            }
            colCursor += span;
            const lines = String(cellSpec.content).split('\n');
            return new TableCell({
              width: { size: cellW, type: WidthType.DXA },
              columnSpan: cellSpec.colSpan,
              rowSpan: cellSpec.rowSpan,
              verticalAlign: VerticalAlign.CENTER,
              shading: { type: ShadingType.CLEAR, color: 'auto', fill: headFill },
              margins: pad,
              borders: headBorderObj,
              children: lines.map(
                ln =>
                  new Paragraph({
                    alignment: cellSpec.align ?? AlignmentType.CENTER,
                    children: [new TextRun({ text: ln, font, size: headSize, bold: true, color: headColor })],
                  })
              ),
            });
          }),
        });
      })
    : [
        new TableRow({
          tableHeader: true,
          children: headers.map((h, i) => {
            const cellAlign =
              theme === 'previous_grid' || theme === 'projected_grid'
                ? AlignmentType.CENTER
                : opts.customAligns?.[i] ?? (i < leftCols ? AlignmentType.LEFT : AlignmentType.RIGHT);
            const lines = String(h).split('\n');
            return new TableCell({
              width: { size: colWidths[i] ?? 1400, type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              shading: { type: ShadingType.CLEAR, color: 'auto', fill: headFill },
              margins: pad,
              borders: headBorderObj,
              children: lines.map(
                ln =>
                  new Paragraph({
                    alignment: cellAlign,
                    children: [new TextRun({ text: ln, font, size: headSize, bold: true, color: headColor })],
                  })
              ),
            });
          }),
        }),
      ];

  let zebra = 0;
  const bodyTableRows: TableRow[] = rows.map(row => {
    const label = String(row[labelColIdx] ?? '').trim();
    const firstColLabel = String(row[0] ?? '').trim();

    const isGrandTotal =
      label.startsWith('TOTAL ') ||
      label === 'TOTAL' ||
      label === 'Net Profit After Tax' ||
      label === 'NET CASH GENERATED FROM OPERATING ACTIVITIES' ||
      label === 'NET CASH FROM / (USED IN) INVESTING ACTIVITIES' ||
      label === 'NET CASH FROM / (USED IN) FINANCING ACTIVITIES' ||
      label === 'Closing Cash & Cash Equivalents (Calculated)' ||
      label === 'Balance Check (Variance)' ||
      label === 'Reconciliation Variance (Calculated - Stated)' ||
      label === 'Total :';

    const isSectionHeader =
      label === 'ASSETS' ||
      label === 'LIABILITIES & SHAREHOLDERS EQUITY' ||
      label.startsWith('CASH FLOWS FROM ');

    const isDoubleBottom = opts.doubleBottomLabels?.has(label) || (theme === 'section76' && isGrandTotal);
    const hasRule = opts.ruleLabels?.has(label) || isDoubleBottom;

    const isBoldRow =
      opts.boldRowLabels?.has(label) ||
      isGrandTotal ||
      isSectionHeader ||
      (theme === 'section76' &&
        (label.startsWith('Total') ||
          label.startsWith('TOTAL') ||
          label === 'Gross Profit' ||
          label === 'Operating Profit (EBIT)' ||
          label === 'Profit Before Tax (EBT)' ||
          label === 'Operating Profit Before Working Capital Changes' ||
          label === 'Cash Generated from Operations' ||
          label.includes('Net Cash') ||
          label.includes('NET CASH') ||
          label.includes('NET INCREASE') ||
          label.includes('Net Change in Cash') ||
          label.includes('Closing Cash') ||
          label.startsWith('Closing Share') ||
          label.startsWith('Closing Retained') ||
          label.startsWith('Closing Reserves'))) ||
      (opts.labelColIndex === 1 && firstColLabel !== '');

    const rowFill =
      theme === 'section76'
        ? isGrandTotal
          ? 'ECFDF5'
          : isSectionHeader
          ? 'E2E8F0'
          : isBoldRow
          ? 'F1F5F9'
          : zebra % 2 === 1
          ? 'F8FAFC'
          : undefined
        : undefined;
    if (!isSectionHeader) zebra++;

    return new TableRow({
      children: Array.from({ length: colWidths.length }, (_, idx) => {
        const cellVal = row[idx];
        const cellText =
          typeof cellVal === 'number'
            ? fmtExportNum(cellVal)
            : cellVal === undefined || cellVal === null
            ? ''
            : String(cellVal);

        let cellAlign = opts.customAligns?.[idx] ?? (idx < leftCols ? AlignmentType.LEFT : AlignmentType.RIGHT);
        if (idx === 0 && opts.rightAlignCol0Labels?.has(firstColLabel)) {
          cellAlign = AlignmentType.RIGHT;
        } else if (theme === 'projected_plain' && opts.labelColIndex === 1 && idx === 0) {
          cellAlign = AlignmentType.CENTER;
        }

        const isStatusReconciled = label === 'Audit Reconciliation Status' && idx > 0 && cellText === 'RECONCILED';
        const isStatusVar = label === 'Audit Reconciliation Status' && idx > 0 && cellText === 'VARIANCE_IDENTIFIED';
        const cellFill = isStatusReconciled ? 'ECFDF5' : isStatusVar ? 'FEF2F2' : rowFill;
        const textColor = isStatusReconciled ? '065F46' : isStatusVar ? '991B1B' : '0F172A';
        const cellBold = isBoldRow || Boolean(opts.boldColIndices?.includes(idx));

        const numStart = opts.labelColIndex === 1 ? 2 : leftCols;
        const cellBorders =
          theme === 'projected_plain'
            ? idx >= numStart && hasRule
              ? {
                  top: darkSingleBorder,
                  bottom: isDoubleBottom ? darkDoubleBorder : darkSingleBorder,
                  left: DOCX_NO_BORDER,
                  right: DOCX_NO_BORDER,
                }
              : { top: DOCX_NO_BORDER, bottom: DOCX_NO_BORDER, left: DOCX_NO_BORDER, right: DOCX_NO_BORDER }
            : theme === 'previous_grid' || theme === 'projected_grid'
            ? {
                top: gridBorder,
                bottom: isDoubleBottom ? darkDoubleBorder : gridBorder,
                left: gridBorder,
                right: gridBorder,
              }
            : isDoubleBottom
            ? {
                top: darkSingleBorder,
                bottom: darkDoubleBorder,
                left: slateBorder,
                right: slateBorder,
              }
            : undefined;

        return new TableCell({
          width: { size: colWidths[idx] ?? 1400, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          shading: cellFill ? { type: ShadingType.CLEAR, color: 'auto', fill: cellFill } : undefined,
          margins: pad,
          borders: cellBorders,
          children: [
            new Paragraph({
              alignment: cellAlign,
              children: [
                new TextRun({
                  text: cellText,
                  font,
                  size: bodySize,
                  bold: cellBold,
                  color: textColor,
                }),
              ],
            }),
          ],
        });
      }),
    });
  });

  return new Table({
    width: { size: totalWidthTwips, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: colWidths,
    borders: tableBorders,
    rows: [...headerTableRows, ...bodyTableRows],
  });
}

export async function exportDocumentWord(
  company: CompanyProfile,
  title: string,
  headers: string[],
  rows: (string | number)[][],
  filename = 'Financial_Report.docx'
) {
  const usableW = docxUsableWidthTwips('landscape', 14);
  const doc = new Document({
    sections: [
      {
        properties: docxPageProps('landscape', 14),
        footers: {
          default: docxFooter(`Certified System Export · Date: ${new Date().toISOString().split('T')[0]}`),
        },
        children: [
          docxSectionHeaderBanner(company, title.toUpperCase(), usableW),
          new Paragraph({ spacing: { before: 100, after: 80 }, children: [] }),
          createDocxTable(headers, rows, {
            theme: 'section76',
            font: 'Helvetica',
            fontSizePt: 8,
            totalWidthTwips: usableW,
          }),
        ],
      },
    ],
  });

  await emitDocxDocument(doc, filename);
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
  filename = `${sanitizeFilePrefix(company.companyName)}_Complete_Package.docx`,
  debtServiceCoverage?: DebtServiceCoverageYear[],
  options?: ExportSyncOptions,
  parsedSchedule?: ParsedRepaymentSchedule | null
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

  const usableW = docxUsableWidthTwips('landscape', 14);
  const coverW = docxUsableWidthTwips('landscape', 10);
  const horizonText = `${incomeData[0]?.year ?? ''} - ${incomeData[incomeData.length - 1]?.year ?? ''}`;
  const doc = new Document({
    sections: [
      {
        properties: docxPageProps('landscape', 10),
        children: [docxCoverPage(company, horizonText, coverW)],
      },
      {
        properties: docxPageProps('landscape', 14),
        footers: {
          default: docxFooter('Certified Export · All statements reconciled under Section 76'),
        },
        children: [
          docxSectionHeaderBanner(company, '1. STATEMENT OF COMPREHENSIVE INCOME (INCOME STATEMENT)', usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          createDocxTable(isHeaders, isRows, { theme: 'section76', font: 'Helvetica', fontSizePt: 8, totalWidthTwips: usableW }),
          new Paragraph({ children: [new PageBreak()] }),

          docxSectionHeaderBanner(company, '2. STATEMENT OF FINANCIAL POSITION (BALANCE SHEET)', usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          createDocxTable(bsHeaders, bsRows, { theme: 'section76', font: 'Helvetica', fontSizePt: 7.5, totalWidthTwips: usableW }),
          new Paragraph({ children: [new PageBreak()] }),

          docxSectionHeaderBanner(company, '3. STATEMENT OF CASH FLOWS (INDIRECT METHOD - SECTION 76)', usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          createDocxTable(cfHeaders, cfRows, { theme: 'section76', font: 'Helvetica', fontSizePt: 7.2, totalWidthTwips: usableW }),
          new Paragraph({ children: [new PageBreak()] }),

          docxSectionHeaderBanner(company, '4. STATEMENT OF CHANGES IN EQUITY', usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          createDocxTable(eqHeaders, eqRows, { theme: 'section76', font: 'Helvetica', fontSizePt: 8, totalWidthTwips: usableW }),
          new Paragraph({ children: [new PageBreak()] }),

          docxSectionHeaderBanner(company, `5. SCHEDULE OF FIXED ASSETS (${latestFA?.year ?? ''})`, usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          createDocxTable(faHeaders, faRows, {
            theme: 'section76',
            font: 'Helvetica',
            fontSizePt: 7.5,
            leftColsCount: 2,
            totalWidthTwips: usableW,
            customAligns: [
              AlignmentType.LEFT,
              AlignmentType.CENTER,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
              AlignmentType.RIGHT,
            ],
          }),
          new Paragraph({ children: [new PageBreak()] }),

          docxSectionHeaderBanner(company, '6. NOTES TO THE FINANCIAL STATEMENTS (STATUTORY DISCLOSURES)', usableW),
          new Paragraph({ spacing: { before: 80, after: 60 }, children: [] }),
          ...notes.flatMap(n => {
            const filteredTbl = filterNoteTableForExport(n.tableData);
            return [
              new Paragraph({
                spacing: { before: 80, after: 30 },
                children: [new TextRun({ text: `Note ${n.noteNumber}: ${n.title} [${n.category}]`, font: 'Helvetica', size: 18, bold: true, color: '1E293B' })],
              }),
              new Paragraph({
                spacing: { after: 50 },
                children: [new TextRun({ text: n.textDescription, font: 'Helvetica', size: 16, color: '475569' })],
              }),
              ...(filteredTbl
                ? [
                    createDocxTable(
                      filteredTbl.headers,
                      [
                        ...filteredTbl.rows.map(r => [r.label, ...r.values]),
                        ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
                      ],
                      { theme: 'section76', font: 'Helvetica', fontSizePt: 7.5, totalWidthTwips: usableW }
                    ),
                  ]
                : []),
              new Paragraph({ spacing: { after: 60 }, children: [] }),
            ];
          }),

          ...((debtServiceCoverage || []).some(d => d.totalDebtService > 0)
            ? [
                new Paragraph({ children: [new PageBreak()] }),
                docxSectionHeaderBanner(company, '7. DEBT SERVICING & REPAYMENT CAPACITY ANALYSIS (DSCR)', usableW),
                new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
                createDocxTable(
                  ['Year', 'EBITDA', 'Pre-Fin OCF', 'Inst. Principal', 'Finance Cost', 'Total Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status', 'Assessment'],
                  (debtServiceCoverage || [])
                    .filter(d => d.totalDebtService > 0)
                    .map(d => [
                      `${d.year}${d.isProjected ? ' (P)' : ' (H)'}`,
                      d.ebitda,
                      d.operatingCashFlowBeforeFinance,
                      d.scheduledPrincipal,
                      d.scheduledFinanceCost,
                      d.totalDebtService,
                      `${d.ebitdaDscr.toFixed(2)}x`,
                      `${d.cashDscr.toFixed(2)}x`,
                      d.closingCash,
                      d.capacityStatus,
                      d.reschedulingContext || '',
                    ]),
                  {
                    theme: 'section76',
                    font: 'Helvetica',
                    fontSizePt: 7.2,
                    totalWidthTwips: usableW,
                    customAligns: [
                      AlignmentType.LEFT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.CENTER,
                      AlignmentType.LEFT,
                    ],
                  }
                ),
              ]
            : []),

          ...(parsedSchedule && parsedSchedule.rows.length > 0
            ? [
                new Paragraph({ children: [new PageBreak()] }),
                docxSectionHeaderBanner(
                  company,
                  `${(debtServiceCoverage || []).some(d => d.totalDebtService > 0) ? '8' : '7'}. REPAYMENT SCHEDULE (EXACT AUTOMATED INSTALLMENT LEDGER)`,
                  usableW
                ),
                new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
                createDocxTable(
                  ['Facility Parameter', 'Value', 'Schedule Dates', 'Date / Setting'],
                  buildRepaymentScheduleMetadataRows(parsedSchedule).map(mr => [
                    String(mr[0]),
                    typeof mr[1] === 'number' ? fmtExportNum(mr[1]) : String(mr[1]),
                    String(mr[5]),
                    String(mr[6]),
                  ]),
                  {
                    theme: 'previous_grid',
                    font: 'Helvetica',
                    fontSizePt: 7.4,
                    colWidthsTwips: [
                      convertMillimetersToTwip(46),
                      convertMillimetersToTwip(45),
                      convertMillimetersToTwip(56),
                      convertMillimetersToTwip(45),
                    ],
                    customAligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.LEFT, AlignmentType.RIGHT],
                    boldColIndices: [0, 2],
                    hideHeader: true,
                  }
                ),
                new Paragraph({ spacing: { before: 80, after: 60 }, children: [] }),
                createDocxTable(
                  ['Event Type.', 'EMI Date', 'Inst. Principal', 'Inst. Profit', 'Gestation Payment', 'Over due Gest.', 'Overdue Profit', 'Balance', 'Total'],
                  [
                    ...parsedSchedule.rows.map(r => [
                      r.eventType,
                      r.emiDateRaw,
                      r.instPrincipal,
                      r.instProfit,
                      r.gestationPayment || 0,
                      r.overdueGestation || 0,
                      r.overdueProfit || 0,
                      r.balance,
                      r.totalInstallment,
                    ]),
                    [
                      'Total :',
                      '',
                      parsedSchedule.grandTotalPrincipal,
                      parsedSchedule.grandTotalProfit,
                      parsedSchedule.grandTotalGestationPayment ?? 0,
                      parsedSchedule.grandTotalOverdueGestation ?? 0,
                      parsedSchedule.grandTotalOverdueProfit ?? 0,
                      parsedSchedule.closingBalance ?? 0,
                      parsedSchedule.grandTotalDebtService,
                    ],
                  ],
                  {
                    theme: 'section76',
                    font: 'Helvetica',
                    fontSizePt: 7.0,
                    totalWidthTwips: usableW,
                    customAligns: [
                      AlignmentType.LEFT,
                      AlignmentType.CENTER,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                    ],
                    doubleBottomLabels: new Set(['Total :']),
                  }
                ),
                new Paragraph({
                  spacing: { before: 60 },
                  children: [
                    new TextRun({
                      text:
                        parsedSchedule.metadata.preparedByNote ||
                        'Prepared By : ABABIL, Developed By : Millennium Information Solution Limited',
                      font: 'Helvetica',
                      size: 15,
                      italics: true,
                      color: '475569',
                    }),
                  ],
                }),
              ]
            : []),
        ],
      },
    ],
  });

  await emitDocxDocument(doc, filename);
}

export async function exportNotesWord(
  company: CompanyProfile,
  notes: FinancialNoteItem[],
  filename = 'Notes_to_the_Financial_Statements.docx'
) {
  const usableW = docxUsableWidthTwips('landscape', 14);
  const doc = new Document({
    sections: [
      {
        properties: docxPageProps('landscape', 14),
        footers: {
          default: docxFooter(`Certified System Export · Date: ${new Date().toISOString().split('T')[0]}`),
        },
        children: [
          docxSectionHeaderBanner(company, 'NOTES TO THE FINANCIAL STATEMENTS (STATUTORY DISCLOSURES 1 TO 12)', usableW),
          new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
          ...notes.flatMap(n => {
            const filteredTbl = filterNoteTableForExport(n.tableData);
            return [
              new Paragraph({
                spacing: { before: 80, after: 30 },
                children: [new TextRun({ text: `Note ${n.noteNumber}: ${n.title} [${n.category}]`, font: 'Helvetica', size: 19, bold: true, color: '1E293B' })],
              }),
              new Paragraph({
                spacing: { after: 50 },
                children: [new TextRun({ text: n.textDescription, font: 'Helvetica', size: 16, color: '475569' })],
              }),
              ...(filteredTbl
                ? [
                    createDocxTable(
                      filteredTbl.headers,
                      [
                        ...filteredTbl.rows.map(r => [r.label, ...r.values]),
                        ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
                      ],
                      { theme: 'section76', font: 'Helvetica', fontSizePt: 7.5, totalWidthTwips: usableW }
                    ),
                  ]
                : []),
              new Paragraph({ spacing: { after: 60 }, children: [] }),
            ];
          }),
        ],
      },
    ],
  });

  await emitDocxDocument(doc, filename);
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
  doc.text(safeCompanyUpper(company), 14, 10);

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
  filename = `${sanitizeFilePrefix(company.companyName)}_Complete_Financial_Package.pdf`,
  debtServiceCoverage?: DebtServiceCoverageYear[],
  options?: ExportSyncOptions,
  parsedSchedule?: ParsedRepaymentSchedule | null
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
  doc.text(safeCompanyUpper(company), 148, 65, { align: 'center' });

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
  doc.text(`Corporate Address: ${safeAddress(company)}`, 148, 133, { align: 'center' });
  doc.text(`Reporting Currency: ${company.reportingCurrency} (${company.reportingUnit}) | FY End: ${normalizeFinancialYearEnd(company.financialYearEnd)} | Horizon: ${incomeData[0]?.year ?? ''} - ${incomeData[incomeData.length - 1]?.year ?? ''}`, 148, 141, { align: 'center' });

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
      head: [['Year', 'EBITDA', 'Pre-Fin OCF', 'Inst. Principal', 'Finance Cost', 'Total Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status', 'Assessment']],
      body: activeDscr.map(d => [
        `${d.year}${d.isProjected ? ' (P)' : ' (H)'}`,
        fmtExportNum(d.ebitda),
        fmtExportNum(d.operatingCashFlowBeforeFinance),
        fmtExportNum(d.scheduledPrincipal),
        fmtExportNum(d.scheduledFinanceCost),
        fmtExportNum(d.totalDebtService),
        `${d.ebitdaDscr.toFixed(2)}x`,
        `${d.cashDscr.toFixed(2)}x`,
        fmtExportNum(d.closingCash),
        d.capacityStatus,
        d.reschedulingContext || '',
      ]),
      startY: 28,
      styles: { fontSize: 7.4, cellPadding: 1.8 },
      headStyles: { fillColor: [30, 41, 59] },
    });
  }

  if (parsedSchedule && parsedSchedule.rows.length > 0) {
    doc.addPage();
    renderPdfHeader(
      doc,
      company,
      `${activeDscr.length > 0 ? '8' : '7'}. REPAYMENT SCHEDULE (EXACT AUTOMATED INSTALLMENT LEDGER)`
    );
    const metaRows = buildRepaymentScheduleMetadataRows(parsedSchedule);
    autoTable(doc, {
      body: metaRows.map(mr => [
        String(mr[0]),
        typeof mr[1] === 'number' ? fmtExportNum(mr[1]) : String(mr[1]),
        String(mr[5]),
        String(mr[6]),
      ]),
      startY: 26,
      theme: 'grid',
      styles: { fontSize: 7.4, cellPadding: 1.4, textColor: [15, 23, 42] },
      columnStyles: {
        0: { fontStyle: 'bold', cellWidth: 46 },
        1: { halign: 'right', cellWidth: 45 },
        2: { fontStyle: 'bold', cellWidth: 56 },
        3: { halign: 'right', cellWidth: 45 },
      },
      margin: { left: 14, right: 14 },
    });
    const metaEndY = ((doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 50) + 4;
    const schedBody: string[][] = [
      ...parsedSchedule.rows.map(r => [
        r.eventType,
        r.emiDateRaw,
        fmtExportNum(r.instPrincipal),
        fmtExportNum(r.instProfit),
        fmtExportNum(r.gestationPayment || 0),
        fmtExportNum(r.overdueGestation || 0),
        fmtExportNum(r.overdueProfit || 0),
        fmtExportNum(r.balance),
        fmtExportNum(r.totalInstallment),
      ]),
      [
        'Total :',
        '',
        fmtExportNum(parsedSchedule.grandTotalPrincipal),
        fmtExportNum(parsedSchedule.grandTotalProfit),
        fmtExportNum(parsedSchedule.grandTotalGestationPayment ?? 0),
        fmtExportNum(parsedSchedule.grandTotalOverdueGestation ?? 0),
        fmtExportNum(parsedSchedule.grandTotalOverdueProfit ?? 0),
        fmtExportNum(parsedSchedule.closingBalance ?? 0),
        fmtExportNum(parsedSchedule.grandTotalDebtService),
      ],
    ];
    autoTable(doc, {
      head: [['Event Type.', 'EMI Date', 'Inst. Principal', 'Inst. Profit', 'Gestation Payment', 'Over due Gest.', 'Overdue Profit', 'Balance', 'Total']],
      body: schedBody,
      startY: metaEndY,
      styles: { fontSize: 7.0, cellPadding: 1.3 },
      headStyles: { fillColor: [30, 41, 59] },
      didParseCell: (data) => {
        if (data.section === 'body' && data.row.index === schedBody.length - 1) {
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.fillColor = [236, 253, 245];
        }
      },
      margin: { left: 14, right: 14 },
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
  doc.text(safeCompanyUpper(company), 14, 9);

  doc.setFontSize(8);
  doc.setFont('helvetica', 'normal');
  doc.text(`${title} · ${company.reportingCurrency} (${company.reportingUnit}) · ${company.accountingFramework} · FY End: ${normalizeFinancialYearEnd(company.financialYearEnd)}`, 14, 16);
  doc.text(`Financials with Tuhin | Mohammad Towhidul Akbar Tuhin`, 297 - 14, 16, { align: 'right' });
}

// ============================================================================
// PREVIOUS FINANCIALS EXACT-FORMAT EXPORTS (.xlsx, .docx, .pdf)
// Strictly preserves layout, particulars, row order, and descending FY-end
// columns from "Previous Financials.pdf"
// ============================================================================

export function sanitizeFilePrefix(name: string): string {
  return (name?.trim() || 'Financial_Statements')
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '') || 'Financial_Statements';
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
    [safeCompanyUpper(company)],
    ['STATEMENT OF FINANCIAL POSITION'],
    [formatFyAsAtUpper(maxYear, company.financialYearEnd)],
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
  const prevOrientation = prevData.length > 4 ? 'landscape' : 'portrait';
  const prevBoldLabels = new Set([
    'Taka', 'Subtotal', 'Capital after drawings', 'Total', 'Current Assets', 'Current Liabilities',
    'Gross Profit', 'Less- Administrative & Selling Expenses', 'Net Operating Profit', 'Net Profit After Income Tax',
    'Net Cash from operating activities', 'Net Cash used in investing Activities', 'Net Cash used in Financing Activities',
    'f) Cash & Cash Equivalents at end of period (D+E)',
    lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'),
    lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'),
    lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'),
  ]);

  const wsBS = XLSX.utils.aoa_to_sheet(bsRows);
  applyColumnWidths(wsBS, widths);
  stylePreviousFormatSheet(wsBS, colHeaders.length, prevOrientation, prevBoldLabels);
  XLSX.utils.book_append_sheet(wb, wsBS, '1. Financial Position');

  // 2. STATEMENT OF COMPREHENSIVE INCOME (dynamic Excel row tracking)
  const isRows: ExcelCellValue[][] = [
    [safeCompanyUpper(company)],
    ['Statement of Comprehensive Income'],
    [formatFyYearEndedTitle(maxYear, company.financialYearEnd)],
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
  stylePreviousFormatSheet(wsIS, colHeaders.length, prevOrientation, prevBoldLabels);
  XLSX.utils.book_append_sheet(wb, wsIS, '2. Comprehensive Income');

  // 3. STATEMENT OF CASH FLOWS (dynamic Excel row tracking)
  const cfRows: ExcelCellValue[][] = [
    [safeCompanyUpper(company)],
    ['STATEMENT OF CASH FLOWS'],
    [formatFyYearEndedNoComma(maxYear, company.financialYearEnd)],
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
  stylePreviousFormatSheet(wsCF, colHeaders.length, prevOrientation, prevBoldLabels);
  XLSX.utils.book_append_sheet(wb, wsCF, '3. Cash Flows');

  // 4. STATEMENT OF CHANGES IN EQUITY (dynamic Excel row tracking)
  const eqRows: ExcelCellValue[][] = [
    [safeCompanyUpper(company)],
    ['STATEMENT OF CHANGES IN EQUITY'],
    [formatFyYearEndedNoComma(maxYear, company.financialYearEnd)],
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
  stylePreviousFormatSheet(wsEQ, colHeaders.length, prevOrientation, prevBoldLabels);
  XLSX.utils.book_append_sheet(wb, wsEQ, '4. Changes in Equity');

  // 5. 1.00 FIXED ASSETS SCHEDULE & NOTES
  const faAsc = [...prevData].sort((a, b) => a.year - b.year);
  const faRows: ExcelCellValue[][] = [
    [safeCompanyUpper(company)],
    [formatFyYearEndedNoComma(maxYear, company.financialYearEnd)],
    ['1.00 FIXED ASSETS SCHEDULE'],
    [],
  ];
  faAsc.forEach(d => {
    const startIdx = faRows.length + 1;
    faRows.push([
      'Particulars',
      `Opening at ${formatFyStartDateCode(d.year, company.financialYearEnd)}`,
      'Addition during the year',
      `Total at ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
      'Rate %',
      'Depreciation charged for the year',
      `W.D.V. at ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
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
  stylePreviousFaNotesSheet(wsFA, prevOrientation);
  XLSX.utils.book_append_sheet(wb, wsFA, '5. Fixed Assets & Notes');

  writeWorkbookStyled(wb, resolvedFilename);
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

  const orientation: 'portrait' | 'landscape' = prevData.length > 4 ? 'landscape' : 'portrait';
  const usableW = docxUsableWidthTwips(orientation, 14);
  const col0W = convertMillimetersToTwip(orientation === 'landscape' ? 115 : 92);
  const yearRemW = Math.floor((usableW - col0W) / prevData.length);
  const prevColWidths = [col0W, ...prevData.map((_, i) => (i === prevData.length - 1 ? usableW - col0W - yearRemW * (prevData.length - 1) : yearRemW))];

  const prevBoldLabels = new Set([
    'PROPERTIES & ASSETS', 'CAPITAL & LIABILITIES',
    'Taka', 'Subtotal', 'Capital after drawings', 'Total', 'Current Assets', 'Current Liabilities',
    'Gross Profit', 'Less- Administrative & Selling Expenses', 'Net Operating Profit', 'Net Profit After Income Tax',
    'Net Cash from operating activities', 'Net Cash used in investing Activities', 'Net Cash used in Financing Activities',
    'a) CASH FLOWS FROM OPERATING ACTIVITIES:', 'b) CASH FLOWS FROM INVESTING ACTIVITIES:', 'c) CASH FLOWS FROM FINANCING ACTIVITIES:',
    'f) Cash & Cash Equivalents at end of period (D+E)',
    lbl(options, 'eq', 'closingShareCapital', 'Closing Share Capital'),
    lbl(options, 'eq', 'closingRetainedEarnings', 'Closing Retained Earnings'),
    lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'),
  ]);
  const prevDoubleBottom = new Set([
    'Taka',
    'Total',
    'Net Profit After Income Tax',
    'f) Cash & Cash Equivalents at end of period (D+E)',
    lbl(options, 'eq', 'totalClosingEquity', 'TOTAL CLOSING EQUITY'),
  ]);
  const rightAlignZero = new Set(['Taka', 'Total']);

  const classicHeader = (stmtTitle: string, subtitle: string, showUnitCaption = true): Paragraph[] => [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new TextRun({ text: safeCompanyUpper(company), font: 'Times New Roman', size: 26, bold: true, color: '0F172A' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 30 },
      children: [new TextRun({ text: stmtTitle, font: 'Times New Roman', size: 24, bold: true, color: '0F172A' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 20 },
      children: [new TextRun({ text: subtitle, font: 'Times New Roman', size: 19, bold: true, color: '0F172A' })],
    }),
    ...(showUnitCaption
      ? [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 40 },
            children: [new TextRun({ text: unitLabel, font: 'Times New Roman', size: 17, color: '0F172A' })],
          }),
        ]
      : []),
  ];

  const prevTableOpts: DocxTableOptions = {
    theme: 'previous_grid',
    font: 'Times New Roman',
    fontSizePt: 8.8,
    headFontSizePt: 9,
    colWidthsTwips: prevColWidths,
    totalWidthTwips: usableW,
    boldRowLabels: prevBoldLabels,
    doubleBottomLabels: prevDoubleBottom,
    rightAlignCol0Labels: rightAlignZero,
  };

  const doc = new Document({
    sections: [
      {
        properties: docxPageProps(orientation, 14),
        children: [
          ...classicHeader('STATEMENT OF FINANCIAL POSITION', formatFyAsAtUpper(maxYear, company.financialYearEnd), true),
          createDocxTable(colHeaders, bsRows, prevTableOpts),

          new Paragraph({ children: [new PageBreak()] }),
          ...classicHeader('Statement of Comprehensive Income', formatFyYearEndedTitle(maxYear, company.financialYearEnd), true),
          createDocxTable(colHeaders, isRows, prevTableOpts),

          new Paragraph({ children: [new PageBreak()] }),
          ...classicHeader('STATEMENT OF CASH FLOWS', formatFyYearEndedNoComma(maxYear, company.financialYearEnd), true),
          createDocxTable(colHeaders, cfRows, prevTableOpts),

          new Paragraph({ children: [new PageBreak()] }),
          ...classicHeader('STATEMENT OF CHANGES IN EQUITY', formatFyYearEndedNoComma(maxYear, company.financialYearEnd), true),
          createDocxTable(colHeaders, eqRows, prevTableOpts),

          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 40 },
            children: [new TextRun({ text: safeCompanyUpper(company), font: 'Times New Roman', size: 26, bold: true, color: '0F172A' })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 60 },
            children: [new TextRun({ text: formatFyYearEndedNoComma(maxYear, company.financialYearEnd), font: 'Times New Roman', size: 19, bold: true, color: '0F172A' })],
          }),
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: '1.00 FIXED ASSETS SCHEDULE', font: 'Times New Roman', size: 20, bold: true, color: '0F172A' })],
          }),
          ...faAsc.flatMap(d => [
            createDocxTable(
              [
                'Particulars',
                `Opening at ${formatFyStartDateCode(d.year, company.financialYearEnd)}`,
                'Addition during the year',
                `Total at ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
                'Rate %',
                'Depreciation charged for the year',
                `W.D.V. at ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
              ],
              [
                ['Fixed Assets', d.faOpening, d.faAddition - d.faDisposal, d.faTotal, d.faRatePct, d.faDepreciationCharged, d.faWDV],
                ['Taka', d.faOpening, d.faAddition - d.faDisposal, d.faTotal, '', d.faDepreciationCharged, d.faWDV],
              ],
              {
                theme: 'previous_grid',
                font: 'Times New Roman',
                fontSizePt: 8.2,
                headFontSizePt: 8.2,
                totalWidthTwips: usableW,
                boldRowLabels: new Set(['Taka']),
                doubleBottomLabels: new Set(['Taka']),
                customAligns: [
                  AlignmentType.LEFT,
                  AlignmentType.RIGHT,
                  AlignmentType.RIGHT,
                  AlignmentType.RIGHT,
                  AlignmentType.CENTER,
                  AlignmentType.RIGHT,
                  AlignmentType.RIGHT,
                ],
              }
            ),
            new Paragraph({ spacing: { after: 80 }, children: [] }),
          ]),

          new Paragraph({
            spacing: { before: 80, after: 60 },
            children: [new TextRun({ text: 'NOTES TO THE FINANCIAL STATEMENTS:', font: 'Times New Roman', size: 20, bold: true, color: '0F172A' })],
          }),
          ...notes.flatMap(n => {
            const filteredTbl = filterNoteTableForExport(n.tableData);
            return [
              new Paragraph({
                spacing: { before: 60, after: 20 },
                children: [new TextRun({ text: `Note ${n.noteNumber}: ${n.title}`, font: 'Times New Roman', size: 18, bold: true, color: '0F172A' })],
              }),
              new Paragraph({
                spacing: { after: 40 },
                children: [new TextRun({ text: n.textDescription, font: 'Times New Roman', size: 17, color: '0F172A' })],
              }),
              ...(filteredTbl
                ? [
                    createDocxTable(
                      filteredTbl.headers,
                      [
                        ...filteredTbl.rows.map(r => [r.label, ...r.values]),
                        ...(filteredTbl.totalRow ? [[filteredTbl.totalRow.label, ...filteredTbl.totalRow.values]] : []),
                      ],
                      {
                        theme: 'previous_grid',
                        font: 'Times New Roman',
                        fontSizePt: 8,
                        totalWidthTwips: usableW,
                        boldRowLabels: filteredTbl.totalRow ? new Set([filteredTbl.totalRow.label]) : undefined,
                        doubleBottomLabels: filteredTbl.totalRow ? new Set([filteredTbl.totalRow.label]) : undefined,
                      }
                    ),
                  ]
                : []),
              new Paragraph({ spacing: { after: 60 }, children: [] }),
            ];
          }),
        ],
      },
    ],
  });

  await emitDocxDocument(doc, resolvedFilename);
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
    doc.text(safeCompanyUpper(company), pageW / 2, 16, { align: 'center' });
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

  renderTablePage('STATEMENT OF FINANCIAL POSITION', formatFyAsAtUpper(maxYear, company.financialYearEnd), bsRows, true);
  renderTablePage('Statement of Comprehensive Income', formatFyYearEndedTitle(maxYear, company.financialYearEnd), isRows);
  renderTablePage('STATEMENT OF CASH FLOWS', formatFyYearEndedNoComma(maxYear, company.financialYearEnd), cfRows);
  renderTablePage('STATEMENT OF CHANGES IN EQUITY', formatFyYearEndedNoComma(maxYear, company.financialYearEnd), eqRows);

  // Page 5: 1.00 FIXED ASSETS SCHEDULE & NOTES
  doc.addPage();
  doc.setTextColor(15, 23, 42);
  doc.setFont('times', 'bold');
  doc.setFontSize(13);
  doc.text(safeCompanyUpper(company), pageW / 2, 16, { align: 'center' });
  doc.setFontSize(9.5);
  doc.text(formatFyYearEndedNoComma(maxYear, company.financialYearEnd), pageW / 2, 23, { align: 'center' });
  doc.setFontSize(10);
  doc.text('1.00 FIXED ASSETS SCHEDULE', 14, 31);

  const faAsc = [...prevData].sort((a, b) => a.year - b.year);
  let yCursor = 34;
  faAsc.forEach((d) => {
    autoTable(doc, {
      head: [[
        'Particulars',
        `Opening at ${formatFyStartDateCode(d.year, company.financialYearEnd)}`,
        'Addition during the year',
        `Total at ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
        'Rate %',
        'Depreciation charged for the year',
        `W.D.V. at ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
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
  explicitOptions?: ExportSyncOptions,
  parsedSchedule?: ParsedRepaymentSchedule | null,
  debtServiceCoverage?: DebtServiceCoverageYear[]
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
    [safeCompanyName(company)],
    ['Statement of Financial Position'],
    [formatFyProjectedAsAt(maxYear, company.financialYearEnd)],
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
  const closingCapLbl = lbl(options, 'eq', 'closingShareCapital', 'Closing Capital / Equity');
  const cogsBaseLbl = lbl(options, 'income', 'costOfSales', 'Cost of Goods Sold');
  const projBsBold = new Set([
    'Assets', 'Non Current Assets', 'Total Non-Current Assets', 'Current Assets', 'Total Current Assets',
    'Total Assets', 'Equity & Liabilities', 'Subtotal', closingCapLbl, 'Non-Current Liabilities',
    'Current Liabilities', 'Total Current Liabilities', 'Total Equity & Liabilities',
  ]);
  const projBsRule = new Set([
    'Total Non-Current Assets', 'Total Current Assets', 'Total Assets',
    'Subtotal', closingCapLbl, 'Total Current Liabilities', 'Total Equity & Liabilities',
  ]);
  const wsBS = XLSX.utils.aoa_to_sheet(bsRows);
  applyColumnWidths(wsBS, widths);
  styleProjectedStatementSheet(
    wsBS,
    'bs',
    2 + projData.length,
    projBsBold,
    projBsRule,
    new Set(['Total Assets', 'Total Equity & Liabilities'])
  );
  XLSX.utils.book_append_sheet(wb, wsBS, '1. Financial Position');

  // 2. Statement of Profit or Loss & Other Comprehensive Income (dynamic Excel row tracking)
  const plRows: ExcelCellValue[][] = [
    [safeCompanyName(company)],
    ['Statement of Profit or Loss & Other Comprehensive Income'],
    [formatFyProjectedYearEnded(maxYear, company.financialYearEnd)],
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
  const projPlRule = new Set([
    'Gross Profit', 'Total Gross Profit & Other Income', 'Total Operating Profit/(Loss)',
    'Net Profit/(Loss) Before Income Tax', 'Net Profit/(Loss) After Income Tax',
  ]);
  const projPlBold = new Set([salesLbl, 'Less: Operating Expenses:', ...projPlRule]);
  const wsPL = XLSX.utils.aoa_to_sheet(plRows);
  applyColumnWidths(wsPL, widths);
  styleProjectedStatementSheet(
    wsPL,
    'pl',
    2 + projData.length,
    projPlBold,
    projPlRule,
    new Set(['Net Profit/(Loss) After Income Tax'])
  );
  XLSX.utils.book_append_sheet(wb, wsPL, '2. Profit or Loss');

  // 3. Statement Of Cash Flows (dynamic Excel row tracking)
  const cfRows: ExcelCellValue[][] = [
    [safeCompanyName(company)],
    [safeAddress(company)],
    ['Statement Of Cash Flows'],
    [formatFyProjectedYearEnded(maxYear, company.financialYearEnd)],
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
  const projCfBold = new Set([
    'Net Cash Used by Operating Activities',
    'Net Cash Used by Investing Activities',
    'Net Cash Provided by Financial Activities',
    'Net Cash Increase',
    'Opening Cash & Bank Balance',
    'Closing Cash & Bank Balance',
    '2. CASH FLOW FROM INVESTING ACTIVITIES:',
    '3. CASH FLOW FROM FINANCIAL ACTIVITIES:',
    'Inflows:',
    'Outflows:',
  ]);
  const wsCF = XLSX.utils.aoa_to_sheet(cfRows);
  applyColumnWidths(wsCF, [50, ...projData.map(() => 18)]);
  styleProjectedStatementSheet(
    wsCF,
    'cf',
    1 + projData.length,
    projCfBold,
    new Set(['Net Cash Used by Operating Activities']),
    new Set(['Closing Cash & Bank Balance'])
  );
  XLSX.utils.book_append_sheet(wb, wsCF, '3. Cash Flows');

  // 4. Notes to the Financial Statements (Notes 1, 2.1-2.6, 3.00-13.00)
  const { quantNotesRows } = buildProjectedTextTables(projData, options);
  const notesRows: ExcelCellValue[][] = [
    [safeCompanyName(company)],
    [safeAddress(company)],
    ['Notes to the Financial Statements'],
    [formatFyProjectedYearEnded(maxYear, company.financialYearEnd)],
    [],
    ['Note No.', 'Particulars'],
    ['1', 'Reporting Entity:'],
    ['', `${safeCompanyName(company)} is a ${company.entityType.replace(/_/g, ' ')} entity and its office is situated at ${safeAddress(company)}.`],
    ['2', 'SIGNIFICANT ACCOUNTING POLICIES:'],
    ['2.1', `Basis of Accounting: Prepared on a going concern basis using accrual method of accounting in accordance with ${company.accountingFramework} and applicable statutory provisions.`],
    ['2.2', 'Accounting Policies: Selected and applied consistently for significant transactions and events to all periods presented.'],
    ['2.3', 'Property, Plant and Equipment: Stated at cost less accumulated depreciation. Depreciation is computed over estimated useful lives consistently.'],
    ['2.4', `Reporting Period: Financial Statements cover the annual reporting period ending ${normalizeFinancialYearEnd(company.financialYearEnd)}.`],
    ['2.5', `Reporting currency and level of precision: Figures represent ${company.reportingCurrency} (${company.reportingUnit}) rounded off cleanly.`],
    ['2.6', 'Comparative Information: Disclosed for all selected years and narratives relevant for understanding the financial statements.'],
    [],
    ['Note No.', 'Particulars', ...projData.map(() => amtLabel)],
    ['', '', ...projData.map(d => d.year)],
    ...quantNotesRows.map(row => row.map((cell, idx) => (idx >= 2 && typeof cell === 'number' ? nCell(cell) : cell))),
  ];
  const projNoteBold = new Set([
    'Subtotal', `Closing ${trRecLbl}`, `Closing ${advLbl}`,
    'Cash In Hand & Bank', ltLoanLbl, `Closing ${trPayLbl}`, `Closing ${taxPayLbl}`,
    'Bank Loan', 'Material Consumed', 'Total Direct Expenses', cogsBaseLbl,
    opexLbl, invLbl, 'Bank Charges & Interest',
  ]);
  const wsNotes = XLSX.utils.aoa_to_sheet(notesRows);
  applyColumnWidths(wsNotes, [12, 42, ...projData.map(() => 18)]);
  styleProjectedNotesSheet(wsNotes, 2 + projData.length, projNoteBold);
  XLSX.utils.book_append_sheet(wb, wsNotes, '4. Notes to Financials');

  // 5. Schedule of Property Plant & Equipments (Schedule-A)
  const projAsc = [...projData].sort((a, b) => a.year - b.year);
  const schedRows: ExcelCellValue[][] = [
    [safeCompanyName(company)],
    [safeAddress(company)],
    ['Schedule of Property Plant & Equipments'],
    [formatFyProjectedAsAt(maxYear, company.financialYearEnd)],
    [],
  ];
  const schedTopHeaderRows: number[] = [];
  projAsc.forEach(d => {
    const topIdx = schedRows.length;
    schedTopHeaderRows.push(topIdx);
    schedRows.push([
      'SL No.',
      'Particulars',
      'Cost',
      '',
      '',
      '',
      'Rate of Dep. (%)',
      'Depreciation Charged During the Year',
      `W.D.V as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
    ]);
    schedRows.push([
      '',
      '',
      'Opening Balance',
      'Addition During the Year',
      'Disposals During the Year',
      `Balance as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`,
      '',
      '',
      '',
    ]);
    const itemRow = topIdx + 3;
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
  styleProjectedScheduleASheet(wsSched, schedTopHeaderRows);
  XLSX.utils.book_append_sheet(wb, wsSched, '5. Schedule-A PPE');

  const activeDscr = (debtServiceCoverage || []).filter(d => d.totalDebtService > 0);
  if (parsedSchedule && parsedSchedule.rows.length > 0) {
    XLSX.utils.book_append_sheet(wb, buildRepaymentScheduleWorksheet(company, parsedSchedule), '6. Repayment Schedule');
  }
  if (activeDscr.length > 0) {
    const dscrTab = parsedSchedule && parsedSchedule.rows.length > 0 ? '7. DSCR Justification' : '6. DSCR Justification';
    XLSX.utils.book_append_sheet(wb, buildDebtServiceCoverageWorksheet(company, activeDscr), dscrTab);
  }

  writeWorkbookStyled(wb, resolvedFilename);
}

export async function exportProjectedFinancialsWord(
  company: CompanyProfile,
  projData: ProjectedFormatYearData[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions,
  parsedSchedule?: ParsedRepaymentSchedule | null,
  debtServiceCoverage?: DebtServiceCoverageYear[]
) {
  if (projData.length === 0) return;
  const { filename: resolvedFilename, options } = resolveFilenameAndOptions(
    buildProjectedFilename(company, projData, 'docx'),
    filenameOrOptions,
    explicitOptions
  );
  const maxYear = projData[0].year;
  const amtLabel = getProjectedAmountCaption(company);
  const shortAmtLabel = projData.length >= 10
    ? (company.reportingCurrency === '৳' ? 'In Taka' : `In ${company.reportingCurrency}`)
    : amtLabel;
  const projAsc = [...projData].sort((a, b) => a.year - b.year);

  const usableMm = 277; // A4 Landscape 297mm - 2*10mm margins
  const usableW = convertMillimetersToTwip(usableMm);
  const nCols = projData.length;
  const partMm = nCols > 15 ? 46 : nCols >= 12 ? 50 : nCols >= 8 ? 56 : 66;
  const notesMm = nCols > 15 ? 13 : nCols >= 12 ? 15 : 18;
  const yearMm = Math.max(11, (usableMm - partMm - notesMm) / nCols);
  const bodyFontSize = nCols > 15 ? 5.1 : nCols >= 12 ? 5.6 : nCols >= 8 ? 6.6 : 7.6;
  const headFontSize = nCols > 15 ? 5.0 : nCols >= 12 ? 5.5 : nCols >= 8 ? 6.5 : 7.5;
  const cfPartMm = nCols > 15 ? 60 : nCols >= 12 ? 66 : 82;
  const cfYearMm = Math.max(11, (usableMm - cfPartMm) / nCols);

  const bsColWidths = [
    convertMillimetersToTwip(partMm),
    convertMillimetersToTwip(notesMm),
    ...projData.map(() => convertMillimetersToTwip(yearMm)),
  ];
  const noteColWidths = [
    convertMillimetersToTwip(notesMm),
    convertMillimetersToTwip(partMm),
    ...projData.map(() => convertMillimetersToTwip(yearMm)),
  ];
  const cfColWidths = [
    convertMillimetersToTwip(cfPartMm),
    ...projData.map(() => convertMillimetersToTwip(cfYearMm)),
  ];

  const {
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
  } = buildProjectedTextTables(projData, options);

  const projHeader = (title: string, subtitle: string, includeAddress = false): Paragraph[] => [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 30 },
      children: [new TextRun({ text: safeCompanyName(company), font: 'Helvetica', size: 23, bold: true, color: '0F172A' })],
    }),
    ...(includeAddress
      ? [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 25 },
            children: [new TextRun({ text: safeAddress(company), font: 'Helvetica', size: 16, color: '0F172A' })],
          }),
        ]
      : []),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 25 },
      children: [new TextRun({ text: title, font: 'Helvetica', size: 21, bold: true, color: '0F172A' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
      children: [new TextRun({ text: subtitle, font: 'Helvetica', size: 17, color: '0F172A' })],
    }),
  ];

  const bsHeaders = ['Particulars', 'Notes', ...projData.map(d => `${shortAmtLabel}\n${d.year}`)];
  const bsMultiHeader: DocxMultiHeaderCell[][] = [
    [
      { content: 'Particulars', rowSpan: 2, align: AlignmentType.CENTER },
      { content: 'Notes', rowSpan: 2, align: AlignmentType.CENTER },
      ...projData.map(() => ({ content: shortAmtLabel, align: AlignmentType.CENTER })),
    ],
    projData.map(d => ({ content: String(d.year), align: AlignmentType.CENTER })),
  ];
  const bsBoldRows = new Set([
    'Assets', 'Non Current Assets', 'Total Non-Current Assets', 'Current Assets', 'Total Current Assets',
    'Total Assets', 'Equity & Liabilities', 'Subtotal', closingCapLbl, 'Non-Current Liabilities',
    'Current Liabilities', 'Total Current Liabilities', 'Total Equity & Liabilities',
  ]);
  const bsRuleRows = new Set([
    'Total Non-Current Assets', 'Total Current Assets', 'Total Assets',
    'Subtotal', closingCapLbl, 'Total Current Liabilities', 'Total Equity & Liabilities',
  ]);

  const plRuleRows = new Set([
    'Gross Profit', 'Total Gross Profit & Other Income', 'Total Operating Profit/(Loss)',
    'Net Profit/(Loss) Before Income Tax', 'Net Profit/(Loss) After Income Tax',
  ]);
  const plBoldRows = new Set([salesLbl, 'Less: Operating Expenses:', ...plRuleRows]);

  const cfHeaders = ['1. CASH FLOW FROM OPERATING ACTIVITIES:', ...projData.map(d => String(d.year))];
  const cfBoldRows = new Set([
    'Net Cash Used by Operating Activities',
    'Net Cash Used by Investing Activities',
    'Net Cash Provided by Financial Activities',
    'Net Cash Increase',
    'Opening Cash & Bank Balance',
    'Closing Cash & Bank Balance',
    '2. CASH FLOW FROM INVESTING ACTIVITIES:',
    '3. CASH FLOW FROM FINANCIAL ACTIVITIES:',
    'Inflows:',
    'Outflows:',
  ]);

  const noteBoldSet = new Set([
    'Subtotal', `Closing ${trRecLbl}`, `Closing ${advLbl}`,
    'Cash In Hand & Bank', ltLoanLbl, `Closing ${trPayLbl}`, `Closing ${taxPayLbl}`,
    'Bank Loan', 'Material Consumed', 'Total Direct Expenses', cogsBaseLbl,
    opexLbl, invLbl, 'Bank Charges & Interest', 'Add: Direct Expenses',
  ]);
  const noteMultiHeader: DocxMultiHeaderCell[][] = [
    [
      { content: 'Note No.', rowSpan: 2, align: AlignmentType.CENTER },
      { content: 'Particulars', rowSpan: 2, align: AlignmentType.CENTER },
      ...projData.map(() => ({ content: shortAmtLabel, align: AlignmentType.CENTER })),
    ],
    projData.map(d => ({ content: String(d.year), align: AlignmentType.CENTER })),
  ];

  const schedColsW = [
    convertMillimetersToTwip(12),
    convertMillimetersToTwip(48),
    convertMillimetersToTwip(32),
    convertMillimetersToTwip(32),
    convertMillimetersToTwip(32),
    convertMillimetersToTwip(32),
    convertMillimetersToTwip(20),
    convertMillimetersToTwip(35),
    convertMillimetersToTwip(34),
  ];

  const doc = new Document({
    sections: [
      {
        properties: docxPageProps('landscape', 10),
        children: [
          ...projHeader('Statement of Financial Position', formatFyProjectedAsAt(maxYear, company.financialYearEnd)),
          createDocxTable(bsHeaders, bsRows, {
            theme: 'projected_plain',
            font: 'Helvetica',
            fontSizePt: bodyFontSize,
            headFontSizePt: headFontSize,
            colWidthsTwips: bsColWidths,
            totalWidthTwips: usableW,
            leftColsCount: 2,
            multiHeaderRows: bsMultiHeader,
            boldRowLabels: bsBoldRows,
            ruleLabels: bsRuleRows,
            doubleBottomLabels: new Set(['Total Assets', 'Total Equity & Liabilities']),
            rightAlignCol0Labels: new Set(['Total Assets', 'Total Equity & Liabilities']),
            customAligns: [AlignmentType.LEFT, AlignmentType.CENTER, ...projData.map(() => AlignmentType.RIGHT)],
          }),
          new Paragraph({ spacing: { before: 120 }, children: [] }),
          docxProprietorSignatureBlock(usableW),

          new Paragraph({ children: [new PageBreak()] }),
          ...projHeader('Statement of Profit or Loss & Other Comprehensive Income', formatFyProjectedYearEnded(maxYear, company.financialYearEnd)),
          createDocxTable(bsHeaders, plRows, {
            theme: 'projected_plain',
            font: 'Helvetica',
            fontSizePt: Math.max(bodyFontSize, 5.8),
            headFontSizePt: headFontSize,
            colWidthsTwips: bsColWidths,
            totalWidthTwips: usableW,
            leftColsCount: 2,
            multiHeaderRows: bsMultiHeader,
            boldRowLabels: plBoldRows,
            ruleLabels: plRuleRows,
            doubleBottomLabels: new Set(['Net Profit/(Loss) After Income Tax']),
            customAligns: [AlignmentType.LEFT, AlignmentType.CENTER, ...projData.map(() => AlignmentType.RIGHT)],
          }),
          new Paragraph({ spacing: { before: 120 }, children: [] }),
          docxProprietorSignatureBlock(usableW),

          new Paragraph({ children: [new PageBreak()] }),
          ...projHeader('Statement Of Cash Flows', formatFyProjectedYearEnded(maxYear, company.financialYearEnd), true),
          createDocxTable(cfHeaders, cfRows, {
            theme: 'projected_plain',
            font: 'Helvetica',
            fontSizePt: bodyFontSize,
            headFontSizePt: headFontSize,
            colWidthsTwips: cfColWidths,
            totalWidthTwips: usableW,
            leftColsCount: 1,
            boldRowLabels: cfBoldRows,
            ruleLabels: new Set(['Net Cash Used by Operating Activities', 'Closing Cash & Bank Balance']),
            doubleBottomLabels: new Set(['Closing Cash & Bank Balance']),
          }),

          new Paragraph({ children: [new PageBreak()] }),
          ...projHeader('Notes to the Financial Statements', formatFyProjectedYearEnded(maxYear, company.financialYearEnd), true),
          createDocxTable(
            ['Note No.', 'Particulars'],
            [
              ['1', 'Reporting Entity:'],
              ['', `${safeCompanyName(company)} is a ${company.entityType.replace(/_/g, ' ')} entity and its office is situated at ${safeAddress(company)}.`],
              ['2', 'SIGNIFICANT ACCOUNTING POLICIES:'],
              ['2.1', `Basis of Accounting: The Financial Statements have been prepared on a going concern basis using accrual method of accounting and in accordance with ${company.accountingFramework} and applicable statutory requirements.`],
              ['2.2', 'Accounting Policies: The Accounting policies were selected and applied consistently for significant transactions and events to all periods presented in these financial statements.'],
              ['2.3', 'Property, Plant and Equipment: Items of PPE are stated at cost less accumulated depreciation. Depreciation is computed over the estimated useful lives consistently.'],
              ['2.4', `Reporting Period: Financial Statements of the entity cover the annual reporting period ending ${normalizeFinancialYearEnd(company.financialYearEnd)} consistently.`],
              ['2.5', `Reporting currency and level of precision: The figures in the Financial Statements represent ${company.reportingCurrency} (${company.reportingUnit}) currency rounded off cleanly.`],
              ['2.6', 'Comparative Information: Comparative Information has been disclosed in respect of all selected years for numerical information, narratives, and descriptive schedules.'],
            ],
            {
              theme: 'projected_plain',
              font: 'Helvetica',
              fontSizePt: 8.5,
              colWidthsTwips: [convertMillimetersToTwip(22), usableW - convertMillimetersToTwip(22)],
              totalWidthTwips: usableW,
              leftColsCount: 2,
              boldColIndices: [0],
              boldRowLabels: new Set(['1', '2']),
              customAligns: [AlignmentType.CENTER, AlignmentType.LEFT],
            }
          ),

          new Paragraph({ children: [new PageBreak()] }),
          ...projHeader('Notes to the Financial Statements', formatFyProjectedYearEnded(maxYear, company.financialYearEnd), true),
          createDocxTable(
            ['Note No.', 'Particulars', ...projData.map(d => `${shortAmtLabel}\n${d.year}`)],
            quantNotesRows,
            {
              theme: 'projected_plain',
              font: 'Helvetica',
              fontSizePt: bodyFontSize,
              headFontSizePt: headFontSize,
              colWidthsTwips: noteColWidths,
              totalWidthTwips: usableW,
              leftColsCount: 2,
              labelColIndex: 1,
              multiHeaderRows: noteMultiHeader,
              boldRowLabels: noteBoldSet,
              ruleLabels: noteBoldSet,
              customAligns: [AlignmentType.CENTER, AlignmentType.LEFT, ...projData.map(() => AlignmentType.RIGHT)],
            }
          ),

          new Paragraph({ children: [new PageBreak()] }),
          ...projHeader('Schedule of Property Plant & Equipments', formatFyProjectedAsAt(maxYear, company.financialYearEnd), true),
          ...projAsc.flatMap(d => [
            createDocxTable(
              ['SL No.', 'Particulars', 'Opening Balance', 'Addition During the Year', 'Disposals During the Year', `Balance as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`, 'Rate (%)', 'Depreciation Charged During the Year', `W.D.V as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`],
              [
                [1, 'Property Plant & Equipment', d.schedOpeningBalance, d.schedAddition, d.schedDisposals, d.schedBalanceClosing, `${d.schedRatePct}%`, d.schedDepreciation, d.schedWDV],
                ['', 'Total', d.schedOpeningBalance, d.schedAddition, d.schedDisposals, d.schedBalanceClosing, '', d.schedDepreciation, d.schedWDV],
              ],
              {
                theme: 'projected_grid',
                font: 'Helvetica',
                fontSizePt: 7.5,
                headFontSizePt: 7.5,
                colWidthsTwips: schedColsW,
                totalWidthTwips: usableW,
                leftColsCount: 2,
                labelColIndex: 1,
                boldRowLabels: new Set(['Total']),
                doubleBottomLabels: new Set(['Total']),
                customAligns: [
                  AlignmentType.CENTER,
                  AlignmentType.LEFT,
                  AlignmentType.RIGHT,
                  AlignmentType.RIGHT,
                  AlignmentType.RIGHT,
                  AlignmentType.RIGHT,
                  AlignmentType.CENTER,
                  AlignmentType.RIGHT,
                  AlignmentType.RIGHT,
                ],
                multiHeaderRows: [
                  [
                    { content: 'SL No.', rowSpan: 2, align: AlignmentType.CENTER },
                    { content: 'Particulars', rowSpan: 2, align: AlignmentType.CENTER },
                    { content: 'Cost', colSpan: 4, align: AlignmentType.CENTER },
                    { content: 'Rate of Dep. (%)', rowSpan: 2, align: AlignmentType.CENTER },
                    { content: 'Depreciation Charged During the Year', rowSpan: 2, align: AlignmentType.CENTER },
                    { content: `W.D.V as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`, rowSpan: 2, align: AlignmentType.CENTER },
                  ],
                  [
                    { content: 'Opening Balance', align: AlignmentType.CENTER },
                    { content: 'Addition During the Year', align: AlignmentType.CENTER },
                    { content: 'Disposals During the Year', align: AlignmentType.CENTER },
                    { content: `Balance as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`, align: AlignmentType.CENTER },
                  ],
                ],
              }
            ),
            new Paragraph({ spacing: { after: 70 }, children: [] }),
          ]),

          ...(parsedSchedule && parsedSchedule.rows.length > 0
            ? [
                new Paragraph({ children: [new PageBreak()] }),
                ...projHeader('Repayment Schedule', safeCompanyName(company)),
                createDocxTable(
                  ['Facility Parameter', 'Value', 'Schedule Dates', 'Date / Setting'],
                  buildRepaymentScheduleMetadataRows(parsedSchedule).map(mr => [
                    String(mr[0]),
                    typeof mr[1] === 'number' ? fmtExportNum(mr[1]) : String(mr[1]),
                    String(mr[5]),
                    String(mr[6]),
                  ]),
                  {
                    theme: 'previous_grid',
                    font: 'Helvetica',
                    fontSizePt: 7.4,
                    colWidthsTwips: [
                      convertMillimetersToTwip(46),
                      convertMillimetersToTwip(45),
                      convertMillimetersToTwip(56),
                      convertMillimetersToTwip(45),
                    ],
                    customAligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.LEFT, AlignmentType.RIGHT],
                    boldColIndices: [0, 2],
                    hideHeader: true,
                  }
                ),
                new Paragraph({ spacing: { after: 60 }, children: [] }),
                createDocxTable(
                  ['Event Type.', 'EMI Date', 'Inst. Principal', 'Inst. Profit', 'Gestation Payment', 'Over due Gest.', 'Overdue Profit', 'Balance', 'Total'],
                  [
                    ...parsedSchedule.rows.map(r => [
                      r.eventType,
                      r.emiDateRaw,
                      r.instPrincipal,
                      r.instProfit,
                      r.gestationPayment || 0,
                      r.overdueGestation || 0,
                      r.overdueProfit || 0,
                      r.balance,
                      r.totalInstallment,
                    ]),
                    [
                      'Total :',
                      '',
                      parsedSchedule.grandTotalPrincipal,
                      parsedSchedule.grandTotalProfit,
                      parsedSchedule.grandTotalGestationPayment ?? 0,
                      parsedSchedule.grandTotalOverdueGestation ?? 0,
                      parsedSchedule.grandTotalOverdueProfit ?? 0,
                      parsedSchedule.closingBalance ?? 0,
                      parsedSchedule.grandTotalDebtService,
                    ],
                  ],
                  {
                    theme: 'section76',
                    font: 'Helvetica',
                    fontSizePt: 6.8,
                    headFontSizePt: 6.8,
                    totalWidthTwips: usableW,
                    customAligns: [
                      AlignmentType.LEFT,
                      AlignmentType.CENTER,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                    ],
                    doubleBottomLabels: new Set(['Total :']),
                  }
                ),
                new Paragraph({
                  spacing: { before: 60 },
                  children: [
                    new TextRun({
                      text:
                        parsedSchedule.metadata.preparedByNote ||
                        'Prepared By : ABABIL, Developed By : Millennium Information Solution Limited',
                      font: 'Helvetica',
                      size: 15,
                      italics: true,
                      color: '475569',
                    }),
                  ],
                }),
              ]
            : []),

          ...((debtServiceCoverage || []).some(d => d.totalDebtService > 0)
            ? [
                new Paragraph({ children: [new PageBreak()] }),
                ...projHeader('Debt Servicing & Rescheduling Justification (DSCR)', formatFyProjectedAsAt(maxYear, company.financialYearEnd)),
                createDocxTable(
                  ['Year', 'EBITDA', 'Pre-Fin OCF', 'Inst. Principal', 'Finance Cost', 'Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status', 'Assessment'],
                  (debtServiceCoverage || [])
                    .filter(d => d.totalDebtService > 0)
                    .map(d => [
                      `${d.year}${d.isProjected ? ' (P)' : ' (H)'}`,
                      d.ebitda,
                      d.operatingCashFlowBeforeFinance,
                      d.scheduledPrincipal,
                      d.scheduledFinanceCost,
                      d.totalDebtService,
                      `${d.ebitdaDscr.toFixed(2)}x`,
                      `${d.cashDscr.toFixed(2)}x`,
                      d.closingCash,
                      d.capacityStatus,
                      d.reschedulingContext || '',
                    ]),
                  {
                    theme: 'section76',
                    font: 'Helvetica',
                    fontSizePt: 7.2,
                    headFontSizePt: 7.2,
                    totalWidthTwips: usableW,
                    customAligns: [
                      AlignmentType.LEFT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.RIGHT,
                      AlignmentType.CENTER,
                      AlignmentType.LEFT,
                    ],
                  }
                ),
              ]
            : []),
        ],
      },
    ],
  });

  await emitDocxDocument(doc, resolvedFilename);
}

export function exportProjectedFinancialsPdf(
  company: CompanyProfile,
  projData: ProjectedFormatYearData[],
  filenameOrOptions?: string | ExportSyncOptions,
  explicitOptions?: ExportSyncOptions,
  parsedSchedule?: ParsedRepaymentSchedule | null,
  debtServiceCoverage?: DebtServiceCoverageYear[]
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
    doc.text(safeCompanyName(company), 148.5, 11.5, { align: 'center' });
    let y = 16;
    if (includeAddress) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text(safeAddress(company), 148.5, y, { align: 'center' });
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
  const startY1 = drawProjHeader('Statement of Financial Position', formatFyProjectedAsAt(maxYear, company.financialYearEnd));
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
  const startY2 = drawProjHeader('Statement of Profit or Loss & Other Comprehensive Income', formatFyProjectedYearEnded(maxYear, company.financialYearEnd));
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
  const startY3 = drawProjHeader('Statement Of Cash Flows', formatFyProjectedYearEnded(maxYear, company.financialYearEnd), true);
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
  const startYNotes1 = drawProjHeader('Notes to the Financial Statements', formatFyProjectedYearEnded(maxYear, company.financialYearEnd), true);
  autoTable(doc, {
    head: [['Note No.', 'Particulars']],
    body: [
      ['1', 'Reporting Entity:'],
      ['', `${safeCompanyName(company)} is a ${company.entityType.replace(/_/g, ' ')} entity and its office is situated at ${safeAddress(company)}.`],
      ['2', 'SIGNIFICANT ACCOUNTING POLICIES:'],
      ['2.1', `Basis of Accounting:\nThe Financial Statements have been prepared on a going concern basis using accrual method of accounting and in accordance with ${company.accountingFramework} and applicable statutory requirements.`],
      ['2.2', 'Accounting Policies:\nThe Accounting policies were selected and applied consistently for significant transactions and events to all periods presented in these financial statements.'],
      ['2.3', 'Property, Plant and Equipment:\nItems of PPE are stated at cost less accumulated depreciation. Depreciation is computed over the estimated useful lives consistently.'],
      ['2.4', `Reporting Period:\nFinancial Statements of the entity cover the annual reporting period ending ${normalizeFinancialYearEnd(company.financialYearEnd)} consistently.`],
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
        if (no === '1' || no === '2' || txt.startsWith(safeCompanyName(company))) {
          data.cell.styles.fontStyle = 'bold';
        }
      }
    },
    margin: { left: 14, right: 14 },
  });

  // Page 5+: Full Quantitative Notes 3.00 - 13.00 (including all roll-forward lines)
  doc.addPage();
  const startY4 = drawProjHeader('Notes to the Financial Statements', formatFyProjectedYearEnded(maxYear, company.financialYearEnd), true);
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
  let ySched = drawProjHeader('Schedule of Property Plant & Equipments', formatFyProjectedAsAt(maxYear, company.financialYearEnd), true);
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
          { content: `W.D.V as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`, rowSpan: 2, styles: { valign: 'middle', halign: 'center' } },
        ],
        ['Opening Balance', 'Addition During the Year', 'Disposals During the Year', `Balance as on ${formatFyEndDateCode(d.year, company.financialYearEnd)}`],
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

  if (parsedSchedule && parsedSchedule.rows.length > 0) {
    doc.addPage();
    const schedStartY = drawProjHeader('Repayment Schedule', safeCompanyName(company));
    const metaRows = buildRepaymentScheduleMetadataRows(parsedSchedule);
    autoTable(doc, {
      body: metaRows.map(mr => [
        String(mr[0]),
        typeof mr[1] === 'number' ? fmtExactNum(mr[1]) : String(mr[1]),
        String(mr[5]),
        String(mr[6]),
      ]),
      startY: schedStartY,
      theme: 'grid',
      styles: { fontSize: 7.4, cellPadding: 1.4, textColor: [15, 23, 42] },
      columnStyles: {
        0: { fontStyle: 'bold', cellWidth: 46 },
        1: { halign: 'right', cellWidth: 45 },
        2: { fontStyle: 'bold', cellWidth: 56 },
        3: { halign: 'right', cellWidth: 45 },
      },
      margin: { left: 10, right: 10 },
    });
    const metaEndY = ((doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? schedStartY + 24) + 3;
    const schedBody: string[][] = [
      ...parsedSchedule.rows.map(r => [
        r.eventType,
        r.emiDateRaw,
        fmtExactNum(r.instPrincipal),
        fmtExactNum(r.instProfit),
        fmtExactNum(r.gestationPayment || 0),
        fmtExactNum(r.overdueGestation || 0),
        fmtExactNum(r.overdueProfit || 0),
        fmtExactNum(r.balance),
        fmtExactNum(r.totalInstallment),
      ]),
      [
        'Total :',
        '',
        fmtExactNum(parsedSchedule.grandTotalPrincipal),
        fmtExactNum(parsedSchedule.grandTotalProfit),
        fmtExactNum(parsedSchedule.grandTotalGestationPayment ?? 0),
        fmtExactNum(parsedSchedule.grandTotalOverdueGestation ?? 0),
        fmtExactNum(parsedSchedule.grandTotalOverdueProfit ?? 0),
        fmtExactNum(parsedSchedule.closingBalance ?? 0),
        fmtExactNum(parsedSchedule.grandTotalDebtService),
      ],
    ];
    autoTable(doc, {
      head: [['Event Type.', 'EMI Date', 'Inst. Principal', 'Inst. Profit', 'Gestation Payment', 'Over due Gest.', 'Overdue Profit', 'Balance', 'Total']],
      body: schedBody,
      startY: metaEndY,
      styles: { fontSize: 6.8, cellPadding: 1.2 },
      headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontStyle: 'bold' },
      didParseCell: (data) => {
        if (data.section === 'body' && data.row.index === schedBody.length - 1) {
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.fillColor = [236, 253, 245];
        }
      },
      margin: { left: 10, right: 10 },
    });
  }

  const activeDscr = (debtServiceCoverage || []).filter(d => d.totalDebtService > 0);
  if (activeDscr.length > 0) {
    doc.addPage();
    const dscrY = drawProjHeader('Debt Servicing & Rescheduling Justification (DSCR)', formatFyProjectedAsAt(maxYear, company.financialYearEnd));
    autoTable(doc, {
      head: [['Year', 'EBITDA', 'Pre-Fin OCF', 'Principal', 'Finance Cost', 'Debt Service', 'EBITDA DSCR', 'Cash DSCR', 'Closing Cash', 'Status', 'Assessment']],
      body: activeDscr.map(d => [
        `${d.year}${d.isProjected ? ' (P)' : ' (H)'}`,
        fmtExactNum(d.ebitda),
        fmtExactNum(d.operatingCashFlowBeforeFinance),
        fmtExactNum(d.scheduledPrincipal),
        fmtExactNum(d.scheduledFinanceCost),
        fmtExactNum(d.totalDebtService),
        `${d.ebitdaDscr.toFixed(2)}x`,
        `${d.cashDscr.toFixed(2)}x`,
        fmtExactNum(d.closingCash),
        d.capacityStatus,
        d.reschedulingContext || '',
      ]),
      startY: dscrY,
      styles: { fontSize: 7.2, cellPadding: 1.5 },
      headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontStyle: 'bold' },
      margin: { left: 10, right: 10 },
    });
  }

  doc.save(resolvedFilename);
}
