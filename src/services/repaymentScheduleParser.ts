import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import {
  AnnualRepaymentSummary,
  ParsedRepaymentSchedule,
  RepaymentInstallmentRow,
  RepaymentScheduleMetadata,
} from '../types/financial';

if (typeof window !== 'undefined' && pdfjsLib.GlobalWorkerOptions) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
    'pdfjs-dist/build/pdf.worker.min.mjs',
    import.meta.url
  ).toString();
}

const r2 = (n: number): number => Math.round((n + Number.EPSILON) * 100) / 100;

/**
 * Parses numbers including Indian/Bangladeshi lakh/crore grouping ("24,30,778.97"),
 * Western commas ("2,430,778.97"), currency prefixes, and accounting parentheses.
 */
export function parseScheduleNumber(raw: unknown): number {
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? r2(raw) : 0;
  }
  const s = String(raw ?? '').trim();
  if (!s || s === '-' || s === '—' || s === '--') return 0;
  const isNeg = (s.startsWith('(') && s.endsWith(')')) || s.startsWith('-');
  const cleaned = s
    .replace(/[()৳$€£,\s]/g, '')
    .replace(/^(BDT|Tk\.?|USD|EUR|GBP|INR|AED)/i, '')
    .replace(/^-/, '');
  if (!cleaned) return 0;
  const num = Number.parseFloat(cleaned);
  if (!Number.isFinite(num)) return 0;
  return r2(isNeg ? -num : num);
}

/**
 * Determines which Financial Year a date belongs to based on the Entity Profile's
 * Financial Year End string (e.g. "31 December", "30 June", "31 March").
 */
export function getFinancialYearForDate(dateISO: string, financialYearEnd = '31 December'): number {
  const parts = dateISO.split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return new Date().getFullYear();
  const [yr, mo, dy] = parts;

  const fyLower = (financialYearEnd || '31 December').toLowerCase();
  let fyEndMonth = 12;
  let fyEndDay = 31;

  if (fyLower.includes('june') || fyLower.includes('jun') || fyLower.includes('/06') || fyLower.includes('-06')) {
    fyEndMonth = 6;
    fyEndDay = 30;
  } else if (fyLower.includes('march') || fyLower.includes('mar') || fyLower.includes('/03') || fyLower.includes('-03')) {
    fyEndMonth = 3;
    fyEndDay = 31;
  } else if (fyLower.includes('sept') || fyLower.includes('sep') || fyLower.includes('/09') || fyLower.includes('-09')) {
    fyEndMonth = 9;
    fyEndDay = 30;
  }

  if (fyEndMonth === 12) return yr;
  // If date falls after FY end month/day, it belongs to next fiscal year
  if (mo > fyEndMonth || (mo === fyEndMonth && dy > fyEndDay)) {
    return yr + 1;
  }
  return yr;
}

const MONTH_MAP: Record<string, number> = {
  jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3,
  apr: 4, april: 4, may: 5, jun: 6, june: 6,
  jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9, september: 9,
  oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
};

function isValidYMD(year: number, month: number, day: number): boolean {
  if (year < 1990 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) return false;
  const d = new Date(Date.UTC(year, month - 1, day));
  return d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day;
}

function toISO(year: number, month: number, day: number): string {
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

/**
 * Infers DD/MM/YYYY vs MM/DD/YYYY across all slash/dash dates in a schedule,
 * evaluating both out-of-range tokens (>12) and monotonic sequence progression (~30d monthly steps).
 */
export function detectSlashDateFormat(
  rawDates: string[],
  intervalHint?: string
): 'DD/MM/YYYY' | 'MM/DD/YYYY' | 'YYYY-MM-DD' {
  const slashTokens: { a: number; b: number; y: number }[] = [];
  let hasISO = false;

  for (const raw of rawDates) {
    const s = raw.trim();
    const isoM = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (isoM) {
      hasISO = true;
      continue;
    }
    const sm = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if (sm) {
      const a = Number(sm[1]);
      const b = Number(sm[2]);
      let y = Number(sm[3]);
      if (y < 100) y += 2000;
      slashTokens.push({ a, b, y });
    }
  }

  if (hasISO && slashTokens.length === 0) return 'YYYY-MM-DD';
  if (slashTokens.length === 0) return 'DD/MM/YYYY';

  // Hard constraint: if any a > 12 => must be DD/MM/YYYY; if any b > 12 => must be MM/DD/YYYY
  const aExceeds12 = slashTokens.some(t => t.a > 12);
  const bExceeds12 = slashTokens.some(t => t.b > 12);
  if (aExceeds12 && !bExceeds12) return 'DD/MM/YYYY';
  if (bExceeds12 && !aExceeds12) return 'MM/DD/YYYY';

  // Score sequential monotonicity & expected interval spacing
  const scoreInterpretation = (isDMY: boolean): number => {
    const timestamps: number[] = [];
    for (const t of slashTokens) {
      const day = isDMY ? t.a : t.b;
      const month = isDMY ? t.b : t.a;
      if (!isValidYMD(t.y, month, day)) return -1e9;
      timestamps.push(Date.UTC(t.y, month - 1, day));
    }
    if (timestamps.length < 2) return 0;

    let score = 0;
    const isWeekly = /week/i.test(intervalHint || '');
    const isQuarterly = /quarter/i.test(intervalHint || '');
    const isYearly = /year|annual/i.test(intervalHint || '');

    for (let i = 1; i < timestamps.length; i++) {
      const diffDays = (timestamps[i] - timestamps[i - 1]) / (1000 * 60 * 60 * 24);
      if (diffDays > 0) score += 2;
      else if (diffDays < 0) score -= 5;

      if (isWeekly) {
        if (diffDays >= 5 && diffDays <= 9) score += 3;
      } else if (isQuarterly) {
        if (diffDays >= 80 && diffDays <= 100) score += 3;
      } else if (isYearly) {
        if (diffDays >= 350 && diffDays <= 380) score += 3;
      } else {
        // Default monthly check (~28-31 days)
        if (diffDays >= 26 && diffDays <= 33) score += 4;
        else if (diffDays >= 1 && diffDays <= 5) score -= 2;
      }
    }
    return score;
  };

  const dmyScore = scoreInterpretation(true);
  const mdyScore = scoreInterpretation(false);
  return dmyScore >= mdyScore ? 'DD/MM/YYYY' : 'MM/DD/YYYY';
}

export function parseScheduleDate(
  raw: unknown,
  preferredFormat: 'DD/MM/YYYY' | 'MM/DD/YYYY' | 'YYYY-MM-DD' | 'EXCEL_SERIAL' = 'DD/MM/YYYY'
): string | null {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return toISO(raw.getFullYear(), raw.getMonth() + 1, raw.getDate());
  }
  if (typeof raw === 'number' && Number.isFinite(raw) && raw > 20000 && raw < 90000) {
    // Excel serial date (epoch Dec 30, 1899)
    const utcDays = Math.floor(raw - 25569);
    const d = new Date(utcDays * 86400 * 1000);
    if (!isNaN(d.getTime())) {
      return toISO(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
    }
  }

  const s = String(raw ?? '').trim();
  if (!s) return null;

  // YYYY-MM-DD
  const isoM = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (isoM) {
    const [, y, m, d] = isoM.map(Number);
    return isValidYMD(y, m, d) ? toISO(y, m, d) : null;
  }

  // DD-MMM-YYYY or DD MMM YYYY
  const namedM = s.match(/^(\d{1,2})[-\s/]([A-Za-z]{3,9})[-\s/,]+(\d{2,4})$/);
  if (namedM) {
    const day = Number(namedM[1]);
    const month = MONTH_MAP[namedM[2].toLowerCase()];
    let year = Number(namedM[3]);
    if (year < 100) year += 2000;
    if (month && isValidYMD(year, month, day)) return toISO(year, month, day);
  }

  // DD/MM/YYYY or MM/DD/YYYY
  const sm = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (sm) {
    const a = Number(sm[1]);
    const b = Number(sm[2]);
    let year = Number(sm[3]);
    if (year < 100) year += 2000;

    if (preferredFormat === 'MM/DD/YYYY') {
      if (isValidYMD(year, a, b)) return toISO(year, a, b);
      if (isValidYMD(year, b, a)) return toISO(year, b, a);
    } else {
      if (isValidYMD(year, b, a)) return toISO(year, b, a);
      if (isValidYMD(year, a, b)) return toISO(year, a, b);
    }
  }

  return null;
}

export interface RebuildScheduleOptions {
  recalculateRunningBalances?: boolean;
  recalculateRowTotals?: boolean;
}

/**
 * Rebuilds year-wise Principal (`Repayment from Borrowings`) and Profit (`Finance Cost`)
 * summaries, auto-computes running balances & row totals when needed/requested,
 * and runs structural & mathematical validation checks.
 */
export function rebuildScheduleSummaries(
  rows: RepaymentInstallmentRow[],
  metadata: RepaymentScheduleMetadata,
  financialYearEnd: string,
  selectedModelYears: number[],
  dateFormatOverride?: 'DD/MM/YYYY' | 'MM/DD/YYYY',
  options?: RebuildScheduleOptions
): ParsedRepaymentSchedule {
  const effectiveFmt = dateFormatOverride || (metadata.dateFormatDetected === 'MM/DD/YYYY' ? 'MM/DD/YYYY' : 'DD/MM/YYYY');
  const hasAnyUploadedBalance = rows.some(r => !r.isDisbursement && r.balance > 0);
  const shouldRecalcBalances = Boolean(options?.recalculateRunningBalances) || !hasAnyUploadedBalance;
  const shouldRecalcRowTotals = Boolean(options?.recalculateRowTotals);

  const mappedRows: RepaymentInstallmentRow[] = rows.map((row, idx) => {
    let iso = row.emiDateISO;
    if (dateFormatOverride && row.emiDateRaw) {
      const reparsed = parseScheduleDate(row.emiDateRaw, dateFormatOverride);
      if (reparsed) iso = reparsed;
    }
    const fy = getFinancialYearForDate(iso, financialYearEnd);
    const instPrincipal = r2(Math.max(0, row.instPrincipal));
    const instProfit = r2(Math.max(0, row.instProfit));
    const gestationPayment = r2(Math.max(0, row.gestationPayment || 0));
    const overdueGestation = r2(Math.max(0, row.overdueGestation || 0));
    const overdueProfit = r2(Math.max(0, row.overdueProfit || 0));
    const computedRowTotal = row.isDisbursement
      ? 0
      : r2(instPrincipal + instProfit + gestationPayment + overdueGestation + overdueProfit);
    const totalInstallment = shouldRecalcRowTotals || (!row.isDisbursement && row.totalInstallment <= 0)
      ? computedRowTotal
      : r2(row.totalInstallment);
    return {
      ...row,
      installmentNo: row.isDisbursement ? 0 : (row.installmentNo || idx + 1),
      emiDateISO: iso,
      financialYear: fy,
      instPrincipal,
      instProfit,
      gestationPayment,
      overdueGestation,
      overdueProfit,
      totalInstallment,
    };
  });

  // Sort chronologically by ISO date & installment number
  mappedRows.sort((a, b) => {
    if (a.emiDateISO !== b.emiDateISO) return a.emiDateISO.localeCompare(b.emiDateISO);
    if (a.isDisbursement && !b.isDisbursement) return -1;
    if (!a.isDisbursement && b.isDisbursement) return 1;
    return a.installmentNo - b.installmentNo;
  });

  // Auto-calculate running Balance if requested or missing from uploaded data
  const disbRow = mappedRows.find(r => r.isDisbursement);
  const rawTotalPrincipal = r2(mappedRows.filter(r => !r.isDisbursement).reduce((s, r) => s + r.instPrincipal, 0));
  const openingPrincipal = r2(
    (metadata.investmentMade && metadata.investmentMade > 0 ? metadata.investmentMade : 0) ||
    (disbRow && disbRow.balance > 0 ? disbRow.balance : 0) ||
    rawTotalPrincipal
  );

  let runningBalance = openingPrincipal;
  let repCounter = 0;
  const balanceDiscrepancyCount = { count: 0 };
  const rowTotalDiscrepancyCount = { count: 0 };

  const normalizedRows = mappedRows.map(row => {
    if (row.isDisbursement) {
      const disbBal = shouldRecalcBalances
        ? openingPrincipal
        : (row.balance > 0 ? r2(row.balance) : openingPrincipal);
      runningBalance = disbBal;
      return {
        ...row,
        installmentNo: 0,
        instPrincipal: 0,
        instProfit: 0,
        gestationPayment: 0,
        overdueGestation: 0,
        overdueProfit: 0,
        balance: disbBal,
        totalInstallment: 0,
      };
    }

    repCounter += 1;
    const expectedBal = r2(Math.max(0, runningBalance - row.instPrincipal));
    const finalBal = shouldRecalcBalances ? expectedBal : r2(row.balance);
    if (!shouldRecalcBalances && Math.abs(expectedBal - row.balance) > 1.0) {
      balanceDiscrepancyCount.count += 1;
    }
    const expectedRowTot = r2(
      row.instPrincipal + row.instProfit + row.gestationPayment + row.overdueGestation + row.overdueProfit
    );
    if (!shouldRecalcRowTotals && row.totalInstallment > 0 && Math.abs(expectedRowTot - row.totalInstallment) > 1.0) {
      rowTotalDiscrepancyCount.count += 1;
    }
    runningBalance = expectedBal;

    return {
      ...row,
      installmentNo: repCounter,
      balance: finalBal,
    };
  });

  const selectedSet = new Set(selectedModelYears);
  const byYear = new Map<number, AnnualRepaymentSummary>();

  for (const row of normalizedRows) {
    const existing = byYear.get(row.financialYear) || {
      financialYear: row.financialYear,
      installmentCount: 0,
      totalPrincipal: 0,
      totalProfit: 0,
      totalGestationPayment: 0,
      totalOverdueGestation: 0,
      totalOverdueProfit: 0,
      totalDebtService: 0,
      disbursementAmount: 0,
      closingBalance: 0,
      isSelectedInModel: selectedSet.has(row.financialYear),
    };

    if (row.isDisbursement) {
      existing.disbursementAmount = r2(existing.disbursementAmount + (row.balance > 0 ? row.balance : (metadata.investmentMade || 0)));
      if (existing.closingBalance === 0 && row.balance > 0) {
        existing.closingBalance = r2(row.balance);
      }
    } else {
      existing.installmentCount += 1;
      existing.totalPrincipal = r2(existing.totalPrincipal + row.instPrincipal);
      existing.totalProfit = r2(existing.totalProfit + row.instProfit);
      existing.totalGestationPayment = r2((existing.totalGestationPayment || 0) + row.gestationPayment);
      existing.totalOverdueGestation = r2((existing.totalOverdueGestation || 0) + row.overdueGestation);
      existing.totalOverdueProfit = r2((existing.totalOverdueProfit || 0) + row.overdueProfit);
      existing.totalDebtService = r2(
        existing.totalPrincipal +
        existing.totalProfit +
        (existing.totalGestationPayment || 0) +
        (existing.totalOverdueGestation || 0) +
        (existing.totalOverdueProfit || 0)
      );
      existing.closingBalance = r2(row.balance);
    }

    byYear.set(row.financialYear, existing);
  }

  // If disbursement happened separately in metadata & year had no disbursement row, attribute Investment Made to disbursement FY
  if (metadata.investmentMade && metadata.investmentMade > 0) {
    const hasDisbAmount = Array.from(byYear.values()).some(s => s.disbursementAmount > 0);
    if (!hasDisbAmount) {
      const disbISO = metadata.disbursementDate ? parseScheduleDate(metadata.disbursementDate, effectiveFmt) : null;
      const firstYear = disbISO
        ? getFinancialYearForDate(disbISO, financialYearEnd)
        : (normalizedRows[0]?.financialYear ?? new Date().getFullYear());
      const entry = byYear.get(firstYear);
      if (entry) {
        entry.disbursementAmount = r2(metadata.investmentMade);
      }
    }
  }

  const annualSummaries = Array.from(byYear.values()).sort((a, b) => a.financialYear - b.financialYear);
  const repaymentOnlyRows = normalizedRows.filter(r => !r.isDisbursement);

  const grandTotalPrincipal = r2(repaymentOnlyRows.reduce((s, r) => s + r.instPrincipal, 0));
  const grandTotalProfit = r2(repaymentOnlyRows.reduce((s, r) => s + r.instProfit, 0));
  const grandTotalGestationPayment = r2(repaymentOnlyRows.reduce((s, r) => s + r.gestationPayment, 0));
  const grandTotalOverdueGestation = r2(repaymentOnlyRows.reduce((s, r) => s + r.overdueGestation, 0));
  const grandTotalOverdueProfit = r2(repaymentOnlyRows.reduce((s, r) => s + r.overdueProfit, 0));
  const grandTotalDebtService = r2(
    grandTotalPrincipal +
    grandTotalProfit +
    grandTotalGestationPayment +
    grandTotalOverdueGestation +
    grandTotalOverdueProfit
  );
  const finalClosingBalance = normalizedRows.length > 0 ? r2(normalizedRows[normalizedRows.length - 1].balance) : 0;

  const validationErrors: string[] = [];
  const validationWarnings: string[] = [];

  if (repaymentOnlyRows.length === 0) {
    validationErrors.push('No valid repayment installment rows containing EMI Date, Inst. Principal, and Inst. Profit were found.');
  }

  const missingDateRows = repaymentOnlyRows.filter(r => !r.emiDateISO || !/^\d{4}-\d{2}-\d{2}$/.test(r.emiDateISO));
  if (missingDateRows.length > 0) {
    validationErrors.push(`${missingDateRows.length} installment row(s) have missing or unreadable EMI Dates.`);
  }

  // Verify stated schedule grand totals against extracted rows
  if (metadata.statedTotalPrincipal !== undefined && metadata.statedTotalPrincipal > 0) {
    const diff = Math.abs(grandTotalPrincipal - metadata.statedTotalPrincipal);
    if (diff > 1.0) {
      validationWarnings.push(
        `Extracted Principal sum (${grandTotalPrincipal.toLocaleString(undefined, { maximumFractionDigits: 2 })}) differs from stated Schedule Total (${metadata.statedTotalPrincipal.toLocaleString(undefined, { maximumFractionDigits: 2 })}) by ${diff.toFixed(2)}. Review rows or click Auto-Recalculate.`
      );
    }
  }

  if (metadata.statedTotalProfit !== undefined && metadata.statedTotalProfit > 0) {
    const diff = Math.abs(grandTotalProfit - metadata.statedTotalProfit);
    if (diff > 1.0) {
      validationWarnings.push(
        `Extracted Profit sum (${grandTotalProfit.toLocaleString(undefined, { maximumFractionDigits: 2 })}) differs from stated Schedule Total (${metadata.statedTotalProfit.toLocaleString(undefined, { maximumFractionDigits: 2 })}) by ${diff.toFixed(2)}. Review rows or click Auto-Recalculate.`
      );
    }
  }

  if (balanceDiscrepancyCount.count > 0) {
    validationWarnings.push(
      `${balanceDiscrepancyCount.count} row(s) show running Balance variance against (Previous Balance - Inst. Principal). Click "Auto-Calculate Balances & Totals" to normalize.`
    );
  }
  if (rowTotalDiscrepancyCount.count > 0) {
    validationWarnings.push(
      `${rowTotalDiscrepancyCount.count} row(s) show row Total variance against component sums. Click "Auto-Calculate Balances & Totals" to normalize.`
    );
  }

  if (metadata.investmentMade && metadata.investmentMade > 0) {
    const loanDiff = Math.abs(grandTotalPrincipal - metadata.investmentMade);
    if (loanDiff > 5 && repaymentOnlyRows.length > 0) {
      validationWarnings.push(
        `Total extracted Inst. Principal (${grandTotalPrincipal.toLocaleString(undefined, { maximumFractionDigits: 2 })}) differs from Investment Made (${metadata.investmentMade.toLocaleString(undefined, { maximumFractionDigits: 2 })}) — schedule may represent a partial or restructured schedule.`
      );
    }
  }

  // Check if any schedule years are not yet in user's selected financial years
  const missingModelYears = annualSummaries.filter(s => !selectedSet.has(s.financialYear)).map(s => s.financialYear);
  if (missingModelYears.length > 0) {
    validationWarnings.push(
      `Schedule covers year(s) ${missingModelYears.join(', ')} not currently selected in the reporting horizon. Implementing will auto-include these years.`
    );
  }

  return {
    metadata: {
      ...metadata,
      investmentMade: metadata.investmentMade ?? openingPrincipal,
      gestationPeriod: metadata.gestationPeriod ?? '0 Days',
      lastInstallmentMinimized: metadata.lastInstallmentMinimized ?? 'No',
      dateFormatDetected: effectiveFmt,
    },
    rows: normalizedRows,
    annualSummaries,
    grandTotalPrincipal,
    grandTotalProfit,
    grandTotalGestationPayment,
    grandTotalOverdueGestation,
    grandTotalOverdueProfit,
    grandTotalDebtService,
    closingBalance: finalClosingBalance,
    validationErrors,
    validationWarnings,
    isValid: validationErrors.length === 0 && repaymentOnlyRows.length > 0,
  };
}

/**
 * Automatically recalculates every row's Total Installment, sequential Running Balance,
 * annual aggregations, and stated totals after inline edits or structural adjustments.
 */
export function recalculateRepaymentSchedule(
  schedule: ParsedRepaymentSchedule,
  financialYearEnd: string,
  selectedModelYears: number[],
  dateFormatOverride?: 'DD/MM/YYYY' | 'MM/DD/YYYY'
): ParsedRepaymentSchedule {
  const rebuilt = rebuildScheduleSummaries(
    schedule.rows,
    schedule.metadata,
    financialYearEnd,
    selectedModelYears,
    dateFormatOverride,
    { recalculateRunningBalances: true, recalculateRowTotals: true }
  );
  return {
    ...rebuilt,
    metadata: {
      ...rebuilt.metadata,
      statedTotalPrincipal: rebuilt.grandTotalPrincipal,
      statedTotalProfit: rebuilt.grandTotalProfit,
    },
  };
}

/**
 * Generates an exact 9-column ABABIL-format amortization Repayment Schedule
 * with automated EMI/installment calculation, running balance, and zero final balance.
 */
export function generateStandardRepaymentSchedule(params: {
  investmentMade: number;
  annualProfitRatePct: number;
  numberOfInstallments: number;
  interval?: 'Monthly' | 'Quarterly' | 'Half-Yearly' | 'Yearly';
  disbursementDateRaw?: string;
  firstInstallmentDateRaw?: string;
  gestationPeriodDays?: number;
  financialYearEnd?: string;
  selectedModelYears?: number[];
}): ParsedRepaymentSchedule {
  const principal = r2(Math.max(1000, params.investmentMade));
  const interval = params.interval || 'Monthly';
  const periodsPerYear = interval === 'Quarterly' ? 4 : interval === 'Half-Yearly' ? 2 : interval === 'Yearly' ? 1 : 12;
  const nInst = Math.max(1, Math.round(params.numberOfInstallments));
  const ratePerPeriod = (Math.max(0, params.annualProfitRatePct) / 100) / periodsPerYear;
  const fyEnd = params.financialYearEnd || '31 December';
  const disbRaw = params.disbursementDateRaw || '01/01/2027';
  const disbISO = parseScheduleDate(disbRaw, 'DD/MM/YYYY') || '2027-01-01';
  const firstRaw = params.firstInstallmentDateRaw || disbRaw;
  const firstISO = parseScheduleDate(firstRaw, 'DD/MM/YYYY') || disbISO;

  const [fy0, fm0, fd0] = firstISO.split('-').map(Number);
  const addMonthsFormat = (baseY: number, baseM: number, baseD: number, monthsToAdd: number): { raw: string; iso: string } => {
    const dt = new Date(Date.UTC(baseY, baseM - 1 + monthsToAdd, baseD));
    const yy = dt.getUTCFullYear();
    const mm = dt.getUTCMonth() + 1;
    const dd = dt.getUTCDate();
    return {
      raw: `${String(dd).padStart(2, '0')}/${String(mm).padStart(2, '0')}/${yy}`,
      iso: `${yy}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`,
    };
  };

  const stepMonths = 12 / periodsPerYear;
  const emi = ratePerPeriod > 0
    ? r2((principal * ratePerPeriod * Math.pow(1 + ratePerPeriod, nInst)) / (Math.pow(1 + ratePerPeriod, nInst) - 1))
    : r2(principal / nInst);

  const rows: RepaymentInstallmentRow[] = [
    {
      id: 'gen_0',
      installmentNo: 0,
      eventType: 'Disbursement',
      emiDateRaw: disbRaw,
      emiDateISO: disbISO,
      financialYear: getFinancialYearForDate(disbISO, fyEnd),
      instPrincipal: 0,
      instProfit: 0,
      gestationPayment: 0,
      overdueGestation: 0,
      overdueProfit: 0,
      balance: principal,
      totalInstallment: 0,
      isDisbursement: true,
    },
  ];

  let balance = principal;
  let lastDateRaw = firstRaw;
  for (let i = 1; i <= nInst; i++) {
    const dt = addMonthsFormat(fy0, fm0, fd0, i * stepMonths);
    lastDateRaw = dt.raw;
    const instProfit = r2(balance * ratePerPeriod);
    let instPrincipal = i === nInst ? r2(balance) : r2(Math.min(balance, Math.max(0, emi - instProfit)));
    if (i === nInst) instPrincipal = r2(balance);
    const nextBalance = r2(Math.max(0, balance - instPrincipal));
    const totalInstallment = r2(instPrincipal + instProfit);
    balance = nextBalance;
    rows.push({
      id: `gen_${i}`,
      installmentNo: i,
      eventType: `Repayment ${i}`,
      emiDateRaw: dt.raw,
      emiDateISO: dt.iso,
      financialYear: getFinancialYearForDate(dt.iso, fyEnd),
      instPrincipal,
      instProfit,
      gestationPayment: 0,
      overdueGestation: 0,
      overdueProfit: 0,
      balance: nextBalance,
      totalInstallment,
      isDisbursement: false,
    });
  }

  const totP = r2(rows.filter(r => !r.isDisbursement).reduce((s, r) => s + r.instPrincipal, 0));
  const totPr = r2(rows.filter(r => !r.isDisbursement).reduce((s, r) => s + r.instProfit, 0));

  return rebuildScheduleSummaries(
    rows,
    {
      sourceFileName: 'Generated_Rescheduled_Repayment_Schedule.pdf',
      parserMethod: 'generated',
      investmentMade: principal,
      interval,
      numberOfInstallments: nInst,
      gestationPeriod: `${params.gestationPeriodDays ?? 0} Days`,
      disbursementDate: disbRaw,
      firstInstallmentDate: firstRaw,
      expiryDate: lastDateRaw,
      lastInstallmentMinimized: 'No',
      annualProfitRatePct: params.annualProfitRatePct,
      statedTotalPrincipal: totP,
      statedTotalProfit: totPr,
      dateFormatDetected: 'DD/MM/YYYY',
      uploadedAt: new Date().toISOString(),
    },
    fyEnd,
    params.selectedModelYears || [],
    'DD/MM/YYYY',
    { recalculateRunningBalances: true, recalculateRowTotals: true }
  );
}

// ---------------- TABLE GRID PARSER (SHARED FOR PDF, EXCEL, CSV, PASTE) ---------------- //

interface GridCell {
  x: number;
  text: string;
}

function normalizeHeader(h: string): string {
  return h.toLowerCase().replace(/[^a-z0-9]/g, '');
}

type ColumnRole =
  | 'eventType'
  | 'emiDate'
  | 'instPrincipal'
  | 'instProfit'
  | 'gestationPayment'
  | 'overdueGestation'
  | 'overdueProfit'
  | 'balance'
  | 'total';

function matchColumnRole(headerText: string): ColumnRole | null {
  const norm = normalizeHeader(headerText);
  if (!norm) return null;
  if (norm.includes('event') || norm === 'sl' || norm === 'slno' || norm === 'instno' || norm === 'no') return 'eventType';
  if (norm.includes('emidate') || norm.includes('duedate') || norm.includes('installmentdate') || norm.includes('repaymentdate') || norm === 'date') return 'emiDate';
  if (norm.includes('instprincipal') || norm === 'principal' || norm.includes('principalpaid') || norm.includes('principalamount')) return 'instPrincipal';
  if (norm.includes('instprofit') || norm === 'profit' || norm === 'interest' || norm.includes('profitamount') || norm.includes('interestamount') || norm.includes('markup')) return 'instProfit';
  if (norm.includes('gestationpayment') || norm === 'gestation') return 'gestationPayment';
  if (norm.includes('overduegest')) return 'overdueGestation';
  if (norm.includes('overdueprofit')) return 'overdueProfit';
  if (norm === 'balance' || norm.includes('closingbalance') || norm.includes('outstanding') || norm.includes('principalbalance')) return 'balance';
  if (norm === 'total' || norm === 'emi' || norm === 'installment' || norm.includes('totalpayment')) return 'total';
  return null;
}

function parseRowsFromStringMatrix(
  matrix: string[][],
  metadataBase: Omit<RepaymentScheduleMetadata, 'dateFormatDetected' | 'uploadedAt'>,
  financialYearEnd: string,
  selectedModelYears: number[]
): ParsedRepaymentSchedule {
  let headerIdx = -1;
  let colMap: Partial<Record<ColumnRole, number>> = {};

  for (let r = 0; r < Math.min(matrix.length, 40); r++) {
    const row = matrix[r];
    const candidateMap: Partial<Record<ColumnRole, number>> = {};
    row.forEach((cell, cIdx) => {
      const role = matchColumnRole(cell);
      if (role && candidateMap[role] === undefined) {
        candidateMap[role] = cIdx;
      }
    });
    if (candidateMap.emiDate !== undefined && (candidateMap.instPrincipal !== undefined || candidateMap.instProfit !== undefined)) {
      headerIdx = r;
      colMap = candidateMap;
      break;
    }
  }

  // Fallback headerless CSV/paste: expect Date, Principal, Profit [, Balance, Total]
  const rawRowsToParse: {
    eventType: string;
    dateRaw: string;
    instPrincipal: number;
    instProfit: number;
    gestationPayment: number;
    overdueGestation: number;
    overdueProfit: number;
    balance: number;
    total: number;
    isDisbursement: boolean;
  }[] = [];

  let metaInvestment = metadataBase.investmentMade;
  let metaInterval = metadataBase.interval;
  let metaNumInst = metadataBase.numberOfInstallments;
  let metaGestationPeriod = metadataBase.gestationPeriod;
  let metaDisbDate = metadataBase.disbursementDate;
  let metaFirstDate = metadataBase.firstInstallmentDate;
  let metaExpiryDate = metadataBase.expiryDate;
  let metaLastMin = metadataBase.lastInstallmentMinimized;
  let statedTotalPrincipal = metadataBase.statedTotalPrincipal;
  let statedTotalProfit = metadataBase.statedTotalProfit;

  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r];
    const joined = row.join(' ').trim();
    if (!joined) continue;

    // Extract metadata before header
    if (headerIdx === -1 || r < headerIdx) {
      const invM = joined.match(/investment\s*made\s*:?\s*([\d,]+\.?\d*)/i);
      if (invM) metaInvestment = parseScheduleNumber(invM[1]);
      const intM = joined.match(/interval\s*:?\s*([A-Za-z]+)/i);
      if (intM) metaInterval = intM[1];
      const numM = joined.match(/number\s*of\s*installment\s*:?\s*(\d+)/i);
      if (numM) metaNumInst = Number(numM[1]);
      const gestM = joined.match(/gestation\s*period\s*:?\s*([0-9]+\s*(?:days?|months?|years?)?)/i);
      if (gestM) metaGestationPeriod = gestM[1].trim();
      const disbM = joined.match(/disbursement\s*date\s*:?\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{2}-\d{2})/i);
      if (disbM) metaDisbDate = disbM[1];
      const firstM = joined.match(/1st\s*installment\s*date\s*:?\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{2}-\d{2})/i);
      if (firstM) metaFirstDate = firstM[1];
      const expM = joined.match(/expiry\s*date\s*:?\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{2}-\d{2})/i);
      if (expM) metaExpiryDate = expM[1];
      const minM = joined.match(/last\s*installment\s*size\s*minimized\s*:?\s*(yes|no)/i);
      if (minM) metaLastMin = minM[1];
      if (headerIdx !== -1) continue;
    }

    // Skip repeated headers on multi-page exports
    if (r !== headerIdx && row.some(c => matchColumnRole(c) === 'instPrincipal') && row.some(c => matchColumnRole(c) === 'instProfit')) {
      continue;
    }

    // Check Total footer row
    if (/^\s*total\s*:?\s*$/i.test(row[0] || '') || /\btotal\s*:/i.test(joined)) {
      if (colMap.instPrincipal !== undefined && row[colMap.instPrincipal]) {
        const val = parseScheduleNumber(row[colMap.instPrincipal]);
        if (val > 0) statedTotalPrincipal = val;
      }
      if (colMap.instProfit !== undefined && row[colMap.instProfit]) {
        const val = parseScheduleNumber(row[colMap.instProfit]);
        if (val > 0) statedTotalProfit = val;
      }
      continue;
    }

    if (headerIdx !== -1 && r > headerIdx) {
      const eventType = colMap.eventType !== undefined ? (row[colMap.eventType] || '').trim() : `Repayment ${rawRowsToParse.length + 1}`;
      const dateRaw = colMap.emiDate !== undefined ? (row[colMap.emiDate] || '').trim() : '';
      const isDisb = /disburs/i.test(eventType);
      const instPrincipal = colMap.instPrincipal !== undefined ? parseScheduleNumber(row[colMap.instPrincipal]) : 0;
      const instProfit = colMap.instProfit !== undefined ? parseScheduleNumber(row[colMap.instProfit]) : 0;
      const gestationPayment = colMap.gestationPayment !== undefined ? parseScheduleNumber(row[colMap.gestationPayment]) : 0;
      const overdueGestation = colMap.overdueGestation !== undefined ? parseScheduleNumber(row[colMap.overdueGestation]) : 0;
      const overdueProfit = colMap.overdueProfit !== undefined ? parseScheduleNumber(row[colMap.overdueProfit]) : 0;
      const balance = colMap.balance !== undefined ? parseScheduleNumber(row[colMap.balance]) : 0;
      const total = colMap.total !== undefined ? parseScheduleNumber(row[colMap.total]) : r2(instPrincipal + instProfit + gestationPayment);

      if (!dateRaw && instPrincipal === 0 && instProfit === 0) continue;
      if (!isDisb && instPrincipal === 0 && instProfit === 0 && balance === 0) continue;

      rawRowsToParse.push({
        eventType: eventType || `Repayment ${rawRowsToParse.length + 1}`,
        dateRaw,
        instPrincipal,
        instProfit,
        gestationPayment,
        overdueGestation,
        overdueProfit,
        balance,
        total,
        isDisbursement: isDisb,
      });
    } else if (headerIdx === -1) {
      // Headerless heuristic fallback: scan cells for a date + numeric columns
      const dateCellIdx = row.findIndex(c => /^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/.test(c.trim()) || /^\d{1,2}[-\s/][A-Za-z]{3,9}[-\s/]\d{2,4}$/.test(c.trim()));
      if (dateCellIdx !== -1) {
        const numsAfter = row.slice(dateCellIdx + 1).map(c => parseScheduleNumber(c));
        const eventType = dateCellIdx > 0 ? row[0].trim() : `Repayment ${rawRowsToParse.length + 1}`;
        const isDisb = /disburs/i.test(eventType);
        const instPrincipal = isDisb ? 0 : (numsAfter[0] ?? 0);
        const instProfit = isDisb ? 0 : (numsAfter[1] ?? 0);
        const balance = numsAfter.length >= 6 ? numsAfter[numsAfter.length - 2] : (numsAfter[2] ?? 0);
        const total = numsAfter.length >= 3 ? numsAfter[numsAfter.length - 1] : r2(instPrincipal + instProfit);
        rawRowsToParse.push({
          eventType,
          dateRaw: row[dateCellIdx].trim(),
          instPrincipal,
          instProfit,
          gestationPayment: 0,
          overdueGestation: 0,
          overdueProfit: 0,
          balance,
          total,
          isDisbursement: isDisb,
        });
      }
    }
  }

  const detectedDateFmt = detectSlashDateFormat(
    rawRowsToParse.map(r => r.dateRaw).filter(Boolean),
    metaInterval
  );

  let repCount = 0;
  const rows: RepaymentInstallmentRow[] = rawRowsToParse
    .map((r, i) => {
      const iso = parseScheduleDate(r.dateRaw, detectedDateFmt);
      if (!iso) return null;
      if (!r.isDisbursement) repCount += 1;
      return {
        id: `inst_${i + 1}`,
        installmentNo: r.isDisbursement ? 0 : repCount,
        eventType: r.eventType,
        emiDateRaw: r.dateRaw,
        emiDateISO: iso,
        financialYear: getFinancialYearForDate(iso, financialYearEnd),
        instPrincipal: r.instPrincipal,
        instProfit: r.instProfit,
        gestationPayment: r.gestationPayment,
        overdueGestation: r.overdueGestation,
        overdueProfit: r.overdueProfit,
        balance: r.balance,
        totalInstallment: r.total,
        isDisbursement: r.isDisbursement,
      };
    })
    .filter((x): x is NonNullable<typeof x> => x !== null);

  return rebuildScheduleSummaries(
    rows,
    {
      ...metadataBase,
      investmentMade: metaInvestment,
      interval: metaInterval,
      numberOfInstallments: metaNumInst,
      gestationPeriod: metaGestationPeriod,
      disbursementDate: metaDisbDate,
      firstInstallmentDate: metaFirstDate,
      expiryDate: metaExpiryDate,
      lastInstallmentMinimized: metaLastMin,
      statedTotalPrincipal,
      statedTotalProfit,
      dateFormatDetected: detectedDateFmt,
      uploadedAt: new Date().toISOString(),
    },
    financialYearEnd,
    selectedModelYears
  );
}

// ---------------- PDF EXTRACTION (pdfjs-dist geometric column mapper) ---------------- //

export async function parseRepaymentSchedulePDF(
  file: File,
  financialYearEnd: string,
  selectedModelYears: number[]
): Promise<ParsedRepaymentSchedule> {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;

  interface PdfTextNode {
    x: number;
    y: number;
    str: string;
  }

  const allMatrixRows: string[][] = [];
  let headerColumnRanges: { role: ColumnRole; xMin: number; xMax: number; xCenter: number }[] = [];
  let metaInvestment: number | undefined;
  let metaInterval: string | undefined;
  let metaNumInst: number | undefined;
  let metaGestationPeriod: string | undefined;
  let metaDisbDate: string | undefined;
  let metaFirstDate: string | undefined;
  let metaExpiryDate: string | undefined;
  let metaLastMin: string | undefined;
  let metaPreparedByNote: string | undefined;
  let statedTotalPrincipal: number | undefined;
  let statedTotalProfit: number | undefined;

  const extractNodesIntoMatrix = (pageNodes: PdfTextNode[]) => {
    // Cluster nodes into horizontal rows by Y coordinate (within +-3.5pt)
    const rowBuckets: { y: number; nodes: GridCell[] }[] = [];
    for (const node of pageNodes) {
      if (!node.str.trim()) continue;
      let bucket = rowBuckets.find(b => Math.abs(b.y - node.y) <= 3.5);
      if (!bucket) {
        bucket = { y: node.y, nodes: [] };
        rowBuckets.push(bucket);
      }
      bucket.nodes.push({ x: node.x, text: node.str.trim() });
    }

    // Sort top-to-bottom (higher PDF Y = top of page)
    rowBuckets.sort((a, b) => b.y - a.y);

    for (const bucket of rowBuckets) {
      bucket.nodes.sort((a, b) => a.x - b.x);
      // Merge horizontally adjacent text tokens (<10pt gap) unless they align to distinct columns
      const mergedCells: GridCell[] = [];
      for (const cell of bucket.nodes) {
        const last = mergedCells[mergedCells.length - 1];
        if (last && cell.x - (last.x + last.text.length * 3.8) < 10) {
          last.text = `${last.text} ${cell.text}`.trim();
        } else {
          mergedCells.push({ ...cell });
        }
      }

      const lineJoined = mergedCells.map(c => c.text).join(' ');

      // Parse metadata lines at top of PDF
      const invM = lineJoined.match(/Investment\s*Made\s*:\s*([\d,]+\.?\d*)/i);
      if (invM) metaInvestment = parseScheduleNumber(invM[1]);
      const intM = lineJoined.match(/Interval\s*:\s*([A-Za-z]+)/i);
      if (intM) metaInterval = intM[1];
      const numM = lineJoined.match(/Number\s*of\s*Installment\s*:\s*(\d+)/i);
      if (numM) metaNumInst = Number(numM[1]);
      const gestM = lineJoined.match(/Gestation\s*Period\s*:\s*([0-9]+\s*(?:Days?|Months?|Years?)?)/i);
      if (gestM) metaGestationPeriod = gestM[1].trim();
      const disbM = lineJoined.match(/Disbursement\s*Date\s*:\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{2}-\d{2})/i);
      if (disbM) metaDisbDate = disbM[1];
      const firstM = lineJoined.match(/1st\s*Installment\s*Date\s*:\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{2}-\d{2})/i);
      if (firstM) metaFirstDate = firstM[1];
      const expM = lineJoined.match(/Expiry\s*Date\s*:\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{2}-\d{2})/i);
      if (expM) metaExpiryDate = expM[1];
      const minM = lineJoined.match(/Last\s*Installment\s*size\s*minimized\s*:\s*(Yes|No)/i);
      if (minM) metaLastMin = minM[1];
      if (/Prepared\s*By\s*:/i.test(lineJoined)) {
        metaPreparedByNote = lineJoined.trim();
      }

      // Check if this row is a table header row
      const matchedCols: { role: ColumnRole; x: number; text: string }[] = [];
      for (const c of mergedCells) {
        const role = matchColumnRole(c.text);
        if (role) matchedCols.push({ role, x: c.x, text: c.text });
      }
      if (matchedCols.some(m => m.role === 'emiDate') && (matchedCols.some(m => m.role === 'instPrincipal') || matchedCols.some(m => m.role === 'instProfit'))) {
        matchedCols.sort((a, b) => a.x - b.x);
        headerColumnRanges = matchedCols.map((col, idx) => {
          const prevX = idx > 0 ? (matchedCols[idx - 1].x + col.x) / 2 : -1e6;
          const nextX = idx < matchedCols.length - 1 ? (col.x + matchedCols[idx + 1].x) / 2 : 1e6;
          return {
            role: col.role,
            xMin: prevX,
            xMax: nextX,
            xCenter: col.x,
          };
        });

        if (allMatrixRows.length === 0) {
          allMatrixRows.push(headerColumnRanges.map(h => h.role));
        }
        continue;
      }

      if (headerColumnRanges.length > 0) {
        // Assign every node by nearest column X-range so empty columns never shift data left!
        const rowArr = new Array<string>(headerColumnRanges.length).fill('');
        for (const rawNode of bucket.nodes) {
          let bestColIdx = 0;
          let bestDist = Infinity;
          headerColumnRanges.forEach((col, idx) => {
            const dist = Math.abs(rawNode.x - col.xCenter);
            if (dist < bestDist) {
              bestDist = dist;
              bestColIdx = idx;
            }
          });
          rowArr[bestColIdx] = rowArr[bestColIdx]
            ? `${rowArr[bestColIdx]} ${rawNode.text}`.trim()
            : rawNode.text;
        }

        if (/^\s*total\s*:?/i.test(rowArr[0] || '') || /\btotal\s*:/i.test(lineJoined)) {
          const pIdx = headerColumnRanges.findIndex(h => h.role === 'instPrincipal');
          const prIdx = headerColumnRanges.findIndex(h => h.role === 'instProfit');
          if (pIdx >= 0 && rowArr[pIdx]) statedTotalPrincipal = parseScheduleNumber(rowArr[pIdx]);
          if (prIdx >= 0 && rowArr[prIdx]) statedTotalProfit = parseScheduleNumber(rowArr[prIdx]);
        } else {
          allMatrixRows.push(rowArr);
        }
      }
    }
  };

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    const nodes: PdfTextNode[] = tc.items
      .map(item => {
        const anyItem = item as { str?: string; transform?: number[] };
        return {
          x: anyItem.transform?.[4] ?? 0,
          y: anyItem.transform?.[5] ?? 0,
          str: anyItem.str ?? '',
        };
      })
      .filter(n => n.str.trim().length > 0);
    extractNodesIntoMatrix(nodes);
  }

  return parseRowsFromStringMatrix(
    allMatrixRows,
    {
      sourceFileName: file.name,
      parserMethod: 'pdf-text',
      investmentMade: metaInvestment,
      interval: metaInterval,
      numberOfInstallments: metaNumInst,
      gestationPeriod: metaGestationPeriod,
      disbursementDate: metaDisbDate,
      firstInstallmentDate: metaFirstDate,
      expiryDate: metaExpiryDate,
      lastInstallmentMinimized: metaLastMin,
      preparedByNote: metaPreparedByNote,
      statedTotalPrincipal,
      statedTotalProfit,
    },
    financialYearEnd,
    selectedModelYears
  );
}

// ---------------- EXCEL & CSV EXTRACTION ---------------- //

export async function parseRepaymentScheduleSpreadsheet(
  file: File,
  financialYearEnd: string,
  selectedModelYears: number[]
): Promise<ParsedRepaymentSchedule> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { cellDates: true });
  const firstSheet = wb.Sheets[wb.SheetNames[0]];
  const rawMatrix = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, raw: true });

  const stringMatrix: string[][] = rawMatrix.map(row =>
    (row || []).map(cell => {
      if (cell instanceof Date && !isNaN(cell.getTime())) {
        return toISO(cell.getFullYear(), cell.getMonth() + 1, cell.getDate());
      }
      return String(cell ?? '').trim();
    })
  );

  const isCsv = file.name.toLowerCase().endsWith('.csv');
  return parseRowsFromStringMatrix(
    stringMatrix,
    {
      sourceFileName: file.name,
      parserMethod: isCsv ? 'csv' : 'excel',
    },
    financialYearEnd,
    selectedModelYears
  );
}

// ---------------- TEXT / PASTE FALLBACK PARSER ---------------- //

export function parseRepaymentScheduleText(
  rawText: string,
  financialYearEnd: string,
  selectedModelYears: number[],
  sourceName = 'Manual Schedule Paste'
): ParsedRepaymentSchedule {
  const lines = rawText
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(Boolean);

  const matrix = lines.map(line => {
    if (line.includes('\t')) return line.split('\t').map(s => s.trim());
    if (line.includes(',')) {
      // Respect quoted numbers like "26,297.85" in CSV
      const cells: string[] = [];
      let cur = '';
      let inQuotes = false;
      for (const ch of line) {
        if (ch === '"') {
          inQuotes = !inQuotes;
        } else if (ch === ',' && !inQuotes) {
          cells.push(cur.trim());
          cur = '';
        } else {
          cur += ch;
        }
      }
      cells.push(cur.trim());
      return cells;
    }
    return line.split(/\s{2,}/).map(s => s.trim());
  });

  return parseRowsFromStringMatrix(
    matrix,
    {
      sourceFileName: sourceName,
      parserMethod: 'manual-paste',
    },
    financialYearEnd,
    selectedModelYears
  );
}

// ---------------- SERVER-SIDE GEMINI OCR FALLBACK ---------------- //

export async function parseRepaymentScheduleViaServerAI(
  file: File,
  financialYearEnd: string,
  selectedModelYears: number[]
): Promise<ParsedRepaymentSchedule> {
  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  const base64Data = btoa(binary);

  const resp = await fetch('/api/repayment-schedule/parse', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      fileName: file.name,
      mimeType: file.type || (file.name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image/png'),
      base64Data,
    }),
  });

  const payload = await resp.json();
  if (!resp.ok) {
    throw new Error(payload?.error || `Server extraction returned status ${resp.status}`);
  }

  const rawRows: {
    eventType?: string;
    emiDate: string;
    instPrincipal: number;
    instProfit: number;
    gestationPayment?: number;
    overdueGestation?: number;
    overdueProfit?: number;
    balance?: number;
    total?: number;
    isDisbursement?: boolean;
  }[] = Array.isArray(payload.rows) ? payload.rows : [];

  const detectedFmt = detectSlashDateFormat(
    rawRows.map(r => r.emiDate).filter(Boolean),
    payload.interval
  );

  let repCount = 0;
  const rows: RepaymentInstallmentRow[] = rawRows
    .map((r, i) => {
      const iso = parseScheduleDate(r.emiDate, detectedFmt);
      if (!iso) return null;
      const isDisb = Boolean(r.isDisbursement) || /disburs/i.test(r.eventType || '');
      if (!isDisb) repCount += 1;
      const p = r2(parseScheduleNumber(r.instPrincipal));
      const pr = r2(parseScheduleNumber(r.instProfit));
      const g = r2(parseScheduleNumber(r.gestationPayment ?? 0));
      const og = r2(parseScheduleNumber(r.overdueGestation ?? 0));
      const op = r2(parseScheduleNumber(r.overdueProfit ?? 0));
      const bal = r2(parseScheduleNumber(r.balance ?? 0));
      const tot = r.total ? r2(parseScheduleNumber(r.total)) : r2(p + pr + g + og + op);
      return {
        id: `ai_inst_${i + 1}`,
        installmentNo: isDisb ? 0 : repCount,
        eventType: r.eventType || (isDisb ? 'Disbursement' : `Repayment ${repCount}`),
        emiDateRaw: r.emiDate,
        emiDateISO: iso,
        financialYear: getFinancialYearForDate(iso, financialYearEnd),
        instPrincipal: p,
        instProfit: pr,
        gestationPayment: g,
        overdueGestation: og,
        overdueProfit: op,
        balance: bal,
        totalInstallment: tot,
        isDisbursement: isDisb,
      };
    })
    .filter((x): x is NonNullable<typeof x> => x !== null);

  return rebuildScheduleSummaries(
    rows,
    {
      sourceFileName: file.name,
      parserMethod: 'gemini-ocr',
      investmentMade: payload.investmentMade ? parseScheduleNumber(payload.investmentMade) : undefined,
      interval: payload.interval,
      numberOfInstallments: payload.numberOfInstallments,
      gestationPeriod: payload.gestationPeriod,
      disbursementDate: payload.disbursementDate,
      firstInstallmentDate: payload.firstInstallmentDate,
      expiryDate: payload.expiryDate,
      lastInstallmentMinimized: payload.lastInstallmentMinimized,
      statedTotalPrincipal: payload.statedTotalPrincipal ? parseScheduleNumber(payload.statedTotalPrincipal) : undefined,
      statedTotalProfit: payload.statedTotalProfit ? parseScheduleNumber(payload.statedTotalProfit) : undefined,
      dateFormatDetected: detectedFmt,
      uploadedAt: new Date().toISOString(),
    },
    financialYearEnd,
    selectedModelYears
  );
}

// ---------------- UNIFIED UPLOAD ORCHESTRATOR WITH SAFE FALLBACK ---------------- //

export async function parseUploadedRepaymentSchedule(
  file: File,
  financialYearEnd: string,
  selectedModelYears: number[]
): Promise<ParsedRepaymentSchedule> {
  const lower = file.name.toLowerCase();

  if (lower.endsWith('.xlsx') || lower.endsWith('.xls') || lower.endsWith('.csv')) {
    return parseRepaymentScheduleSpreadsheet(file, financialYearEnd, selectedModelYears);
  }

  if (lower.endsWith('.txt')) {
    const txt = await file.text();
    return parseRepaymentScheduleText(txt, financialYearEnd, selectedModelYears, file.name);
  }

  if (lower.endsWith('.pdf')) {
    try {
      const pdfParsed = await parseRepaymentSchedulePDF(file, financialYearEnd, selectedModelYears);
      if (pdfParsed.rows.filter(r => !r.isDisbursement).length > 0) {
        return pdfParsed;
      }
      // Scanned PDF fallback via server Gemini endpoint
      return await parseRepaymentScheduleViaServerAI(file, financialYearEnd, selectedModelYears);
    } catch (pdfErr) {
      // Attempt server Gemini fallback if client PDF reader hit an issue
      try {
        return await parseRepaymentScheduleViaServerAI(file, financialYearEnd, selectedModelYears);
      } catch (serverErr) {
        const msg = serverErr instanceof Error ? serverErr.message : String(pdfErr);
        throw new Error(
          `Unable to extract schedule rows from PDF (${msg}). Use the Excel/CSV upload or Paste Schedule Table safe fallback below.`
        );
      }
    }
  }

  if (/\.(png|jpe?g|webp)$/i.test(lower)) {
    return parseRepaymentScheduleViaServerAI(file, financialYearEnd, selectedModelYears);
  }

  throw new Error('Unsupported file type. Please upload PDF, Excel (.xlsx/.xls), CSV, or Image (.png/.jpg), or paste table data using the safe fallback.');
}

// ---------------- ATTACHED CLIENT ABABIL SCHEDULE (60 INSTALLMENTS, 2027-2032) ---------------- //

const ABABIL_ATTACHED_ROWS: [string, string, number, number, number, number][] = [
  ['Disbursement', '01/01/2027', 0, 0, 2430778.97, 0],
  ['Repayment 1', '01/02/2027', 26297.85, 32444.15, 2404481.12, 58742.00],
  ['Repayment 2', '01/03/2027', 29754.64, 28987.36, 2374726.48, 58742.00],
  ['Repayment 3', '01/04/2027', 27046.00, 31696.00, 2347680.48, 58742.00],
  ['Repayment 4', '01/05/2027', 28417.79, 30324.21, 2319262.69, 58742.00],
  ['Repayment 5', '01/06/2027', 27786.29, 30955.71, 2291476.40, 58742.00],
  ['Repayment 6', '01/07/2027', 29143.76, 29598.24, 2262332.64, 58742.00],
  ['Repayment 7', '01/08/2027', 28546.14, 30195.86, 2233786.50, 58742.00],
  ['Repayment 8', '01/09/2027', 28927.16, 29814.84, 2204859.34, 58742.00],
  ['Repayment 9', '01/10/2027', 30262.57, 28479.43, 2174596.77, 58742.00],
  ['Repayment 10', '01/11/2027', 29717.17, 29024.83, 2144879.60, 58742.00],
  ['Repayment 11', '01/12/2027', 31037.31, 27704.69, 2113842.29, 58742.00],
  ['Repayment 12', '01/01/2028', 30528.08, 28213.92, 2083314.21, 58742.00],
  ['Repayment 13', '01/02/2028', 30935.54, 27806.46, 2052378.67, 58742.00],
  ['Repayment 14', '01/03/2028', 33115.77, 25626.23, 2019262.90, 58742.00],
  ['Repayment 15', '01/04/2028', 31790.45, 26951.55, 1987472.45, 58742.00],
  ['Repayment 16', '01/05/2028', 33070.48, 25671.52, 1954401.97, 58742.00],
  ['Repayment 17', '01/06/2028', 32656.16, 26085.84, 1921745.81, 58742.00],
  ['Repayment 18', '01/07/2028', 33919.45, 24822.55, 1887826.36, 58742.00],
  ['Repayment 19', '01/08/2028', 33544.76, 25197.24, 1854281.60, 58742.00],
  ['Repayment 20', '01/09/2028', 33992.49, 24749.51, 1820289.11, 58742.00],
  ['Repayment 21', '01/10/2028', 35229.93, 23512.07, 1785059.18, 58742.00],
  ['Repayment 22', '01/11/2028', 34916.42, 23825.58, 1750142.76, 58742.00],
  ['Repayment 23', '01/12/2028', 36135.99, 22606.01, 1714006.77, 58742.00],
  ['Repayment 24', '01/01/2029', 35864.77, 22877.23, 1678142.00, 58742.00],
  ['Repayment 25', '01/02/2029', 36343.47, 22398.53, 1641798.53, 58742.00],
  ['Repayment 26', '01/03/2029', 38949.21, 19792.79, 1602849.32, 58742.00],
  ['Repayment 27', '01/04/2029', 37348.41, 21393.59, 1565500.91, 58742.00],
  ['Repayment 28', '01/05/2029', 38520.95, 20221.05, 1526979.96, 58742.00],
  ['Repayment 29', '01/06/2029', 38361.06, 20380.94, 1488618.90, 58742.00],
  ['Repayment 30', '01/07/2029', 39514.01, 19227.99, 1449104.89, 58742.00],
  ['Repayment 31', '01/08/2029', 39400.48, 19341.52, 1409704.41, 58742.00],
  ['Repayment 32', '01/09/2029', 39926.36, 18815.64, 1369778.05, 58742.00],
  ['Repayment 33', '01/10/2029', 41049.03, 17692.97, 1328729.02, 58742.00],
  ['Repayment 34', '01/11/2029', 41007.16, 17734.84, 1287721.86, 58742.00],
  ['Repayment 35', '01/12/2029', 42108.93, 16633.07, 1245612.93, 58742.00],
  ['Repayment 36', '01/01/2030', 42116.53, 16625.47, 1203496.40, 58742.00],
  ['Repayment 37', '01/02/2030', 42678.67, 16063.33, 1160817.73, 58742.00],
  ['Repayment 38', '01/03/2030', 44747.70, 13994.30, 1116070.03, 58742.00],
  ['Repayment 39', '01/04/2030', 43845.57, 14896.43, 1072224.46, 58742.00],
  ['Repayment 40', '01/05/2030', 44892.43, 13849.57, 1027332.03, 58742.00],
  ['Repayment 41', '01/06/2030', 45029.97, 13712.03, 982302.06, 58742.00],
  ['Repayment 42', '01/07/2030', 46053.93, 12688.07, 936248.13, 58742.00],
  ['Repayment 43', '01/08/2030', 46245.69, 12496.31, 890002.44, 58742.00],
  ['Repayment 44', '01/09/2030', 46862.94, 11879.06, 843139.50, 58742.00],
  ['Repayment 45', '01/10/2030', 47851.45, 10890.55, 795288.05, 58742.00],
  ['Repayment 46', '01/11/2030', 48127.11, 10614.89, 747160.94, 58742.00],
  ['Repayment 47', '01/12/2030', 49091.17, 9650.83, 698069.77, 58742.00],
  ['Repayment 48', '01/01/2031', 49424.71, 9317.29, 648645.06, 58742.00],
  ['Repayment 49', '01/02/2031', 50084.39, 8657.61, 598560.67, 58742.00],
  ['Repayment 50', '01/03/2031', 51526.02, 7215.98, 547034.65, 58742.00],
  ['Repayment 51', '01/04/2031', 51440.61, 7301.39, 495594.04, 58742.00],
  ['Repayment 52', '01/05/2031', 52340.58, 6401.42, 443253.46, 58742.00],
  ['Repayment 53', '01/06/2031', 52825.80, 5916.20, 390427.66, 58742.00],
  ['Repayment 54', '01/07/2031', 53698.98, 5043.02, 336728.68, 58742.00],
  ['Repayment 55', '01/08/2031', 54247.61, 4494.39, 282481.07, 58742.00],
  ['Repayment 56', '01/09/2031', 54971.66, 3770.34, 227509.41, 58742.00],
  ['Repayment 57', '01/10/2031', 55803.34, 2938.66, 171706.07, 58742.00],
  ['Repayment 58', '01/11/2031', 56450.20, 2291.80, 115255.87, 58742.00],
  ['Repayment 59', '01/12/2031', 57253.28, 1488.72, 58002.59, 58742.00],
  ['Repayment 60', '01/01/2032', 58002.59, 773.41, 0.00, 58776.00],
];

export function getAttachedReferenceRepaymentSchedule(
  financialYearEnd = '31 December',
  selectedModelYears: number[] = []
): ParsedRepaymentSchedule {
  const rows: RepaymentInstallmentRow[] = ABABIL_ATTACHED_ROWS.map(([eventType, dateRaw, p, pr, bal, tot], idx) => {
    const isDisb = eventType === 'Disbursement';
    const iso = parseScheduleDate(dateRaw, 'DD/MM/YYYY')!;
    return {
      id: `ababil_${idx}`,
      installmentNo: isDisb ? 0 : idx,
      eventType,
      emiDateRaw: dateRaw,
      emiDateISO: iso,
      financialYear: getFinancialYearForDate(iso, financialYearEnd),
      instPrincipal: p,
      instProfit: pr,
      gestationPayment: 0,
      overdueGestation: 0,
      overdueProfit: 0,
      balance: bal,
      totalInstallment: tot,
      isDisbursement: isDisb,
    };
  });

  return rebuildScheduleSummaries(
    rows,
    {
      sourceFileName: 'Repayment_Schedule_ABABIL_Reference.pdf',
      parserMethod: 'ababil-reference',
      investmentMade: 2430778.97,
      interval: 'Monthly',
      numberOfInstallments: 96,
      gestationPeriod: '0 Days',
      disbursementDate: '01/01/2027',
      firstInstallmentDate: '01/01/2027',
      expiryDate: '01/01/2032',
      lastInstallmentMinimized: 'No',
      preparedByNote: 'Prepared By : ABABIL, Developed By : Millennium Information Solution Limited',
      statedTotalPrincipal: 2430778.97,
      statedTotalProfit: 1093775.03,
      dateFormatDetected: 'DD/MM/YYYY',
      uploadedAt: new Date().toISOString(),
    },
    financialYearEnd,
    selectedModelYears
  );
}
