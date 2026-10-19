import { read, utils, write } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING } from './projectMapping';

// Canonical bucket list matching required format for Lead Levels and Enquiry Levels:
// Open, Site Visit Scheduled, Site Visited, Revisited, Cold, Warm, Hot, Discard, Booked
export const DEFAULT_BUCKET_LIST = [
  'Open',
  'Site Visit Scheduled',
  'Site Visited',
  'Revisited',
  'Cold',
  'Warm',
  'Hot',
  'Discard',
  'Booked'
];

export const DEFAULT_PRESALES_BUCKET_LIST = [
  'Open',
  'Site Visit Scheduled',
  'Site Visited',
  'Revisited',
  'Cold',
  'Warm',
  'Hot',
  'Discard',
  'Booked'
];

export function normalizeBucketCasing(val: string): string {
  const trimmed = val.trim();
  const lower = trimmed.toLowerCase();
  const canonicalMap: Record<string, string> = {
    'open': 'Open',
    'site visit scheduled': 'Site Visit Scheduled',
    'site visited': 'Site Visited',
    'visit scheduled': 'Site Visit Scheduled',
    'scheduled': 'Site Visit Scheduled',
    'visited': 'Site Visited',
    'revisited': 'Revisited',
    're-visited': 'Revisited',
    're visit': 'Revisited',
    're-visit': 'Revisited',
    'cold': 'Cold',
    'warm': 'Warm',
    'hot': 'Hot',
    'discard': 'Discard',
    'drop': 'Discard',
    'junk': 'Discard',
    'lost': 'Discard',
    'booked': 'Booked',
    'booking': 'Booked'
  };
  return canonicalMap[lower] || (trimmed.length > 0 ? (trimmed.charAt(0).toUpperCase() + trimmed.slice(1)) : trimmed);
}

export interface DetectedColumn {
  colIndex: number;
  headerName: string;
  role: 'sales' | 'bucket' | 'project' | 'source' | 'agency' | 'other';
  roleLabel: string;
  uniqueValues: string[];
  valueCounts: Record<string, number>;
  totalCount: number;
}

export interface BucketReportAnalysis {
  headerIndex: number;
  detectedColumns: DetectedColumn[];
  suggestedSalesColIdx: number;
  suggestedBucketColIdx: number;
  suggestedAgencyColIdx?: number;
  suggestedProjectColIdx: number;
  suggestedDateColIdx?: number;
  suggestedMobileColIdx?: number;
  detectedProject: string;
  detectedStartDate: string;
  detectedEndDate: string;
  detectedStartDateYMD?: string;
  detectedEndDateYMD?: string;
  defaultTitle: string;
  totalRows: number;
}

export interface VisitsReportAnalysis {
  headerIndex: number;
  detectedColumns: DetectedColumn[];
  suggestedProjectColIdx: number;
  suggestedMobileColIdx: number;
  suggestedRevisitDateColIdx: number;
  suggestedDateColIdx: number;
  detectedProject: string;
  totalRows: number;
}

export interface BucketSubRowData {
  agency: string;
  grandTotal: number;
  bucketCounts: Record<string, number>;
  visitsDone?: number;
}

export interface BucketRowData {
  salesUser: string;
  grandTotal: number;
  bucketCounts: Record<string, number>;
  isGroupHeader?: boolean;
  subRows?: BucketSubRowData[];
  visitsDone?: number;
}

export interface BucketTableSummary {
  reportTitle: string;
  columns: string[]; // ['Telecaller' | 'Sales', 'Grand Total', ...buckets, (optional) 'Visits Done']
  buckets: string[];
  rows: BucketRowData[];
  columnTotals: {
    grandTotal: number;
    bucketTotals: Record<string, number>;
    visitsDoneTotal?: number;
  };
  dimensionLabel?: string;
  hasVisitsDone?: boolean;
}

function getCellValue(cell: any): string {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'object' && cell.v !== undefined) return String(cell.v);
  return String(cell);
}

/**
 * Checks if a cell value represents an empty, missing, NA, or null value
 */
export function isBlankValue(val: any): boolean {
  if (val === null || val === undefined) return true;
  const s = String(typeof val === 'object' && val.v !== undefined ? val.v : val).trim();
  if (s === '' || s === '-' || s === '--') return true;
  const lower = s.toLowerCase();
  return (
    lower === 'na' ||
    lower === 'n/a' ||
    lower === 'n / a' ||
    lower === 'null' ||
    lower === 'undefined' ||
    lower === 'none' ||
    lower === '(blank)' ||
    lower === '#n/a' ||
    lower === '#na'
  );
}

/**
 * Normalizes a record value according to report rules:
 * - ONLY for the "Lead Level" column: blank, NA, null, "-", etc. are normalized to "Open"
 * - For Agency Name and all other columns: blank, NA, null, "-", etc. are normalized to "(blank)"
 */
export function normalizeRecordValue(val: any, isLeadLevel: boolean = false): string {
  if (isBlankValue(val)) {
    return isLeadLevel ? 'Open' : '(blank)';
  }
  const s = String(typeof val === 'object' && val.v !== undefined ? val.v : val).trim();
  return s;
}

export function formatToDDMMYYYY(dateStr: string): string {
  if (!dateStr) return '';
  const dmyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmyMatch) {
    const d = dmyMatch[1].padStart(2, '0');
    const m = dmyMatch[2].padStart(2, '0');
    const y = dmyMatch[3];
    return `${d}-${m}-${y}`;
  }
  const ymdMatch = dateStr.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymdMatch) {
    const y = ymdMatch[1];
    const m = ymdMatch[2].padStart(2, '0');
    const d = ymdMatch[3].padStart(2, '0');
    return `${d}-${m}-${y}`;
  }
  const dateObj = new Date(dateStr);
  if (!isNaN(dateObj.getTime())) {
    const d = String(dateObj.getDate()).padStart(2, '0');
    const m = String(dateObj.getMonth() + 1).padStart(2, '0');
    const y = dateObj.getFullYear();
    return `${d}-${m}-${y}`;
  }
  return dateStr;
}

/**
 * Robust date parser handling Excel serial numbers, Date instances, ISO strings,
 * DD-MM-YYYY, DD/MM/YYYY, DD Mon YYYY, with or without time.
 */
export function parseAnyDate(val: any): Date | null {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  if (typeof val === 'number') {
    // Check if it's an Excel serial date number (typically between 20000 and 90000)
    if (val > 20000 && val < 90000) {
      // Excel epoch: Dec 30 1899
      const utcMs = Math.round((val - 25569) * 86400 * 1000);
      const d = new Date(utcMs);
      return isNaN(d.getTime()) ? null : d;
    }
  }
  const str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'null' || str.toLowerCase() === 'na' || str.toLowerCase() === 'n/a') return null;

  // Number formatted as string
  const numVal = Number(str);
  if (!isNaN(numVal) && numVal > 20000 && numVal < 90000 && !str.includes('-') && !str.includes('/') && !str.includes(':')) {
    const utcMs = Math.round((numVal - 25569) * 86400 * 1000);
    const d = new Date(utcMs);
    return isNaN(d.getTime()) ? null : d;
  }

  // Handle DD-MM-YYYY or DD/MM/YYYY with optional time
  const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmy) {
    const day = parseInt(dmy[1], 10);
    const month = parseInt(dmy[2], 10) - 1;
    const year = parseInt(dmy[3], 10);
    const hour = dmy[4] ? parseInt(dmy[4], 10) : 0;
    const min = dmy[5] ? parseInt(dmy[5], 10) : 0;
    const sec = dmy[6] ? parseInt(dmy[6], 10) : 0;
    const d = new Date(year, month, day, hour, min, sec);
    return isNaN(d.getTime()) ? null : d;
  }

  // Handle YYYY-MM-DD
  const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (ymd) {
    const year = parseInt(ymd[1], 10);
    const month = parseInt(ymd[2], 10) - 1;
    const day = parseInt(ymd[3], 10);
    const hour = ymd[4] ? parseInt(ymd[4], 10) : 0;
    const min = ymd[5] ? parseInt(ymd[5], 10) : 0;
    const sec = ymd[6] ? parseInt(ymd[6], 10) : 0;
    const d = new Date(year, month, day, hour, min, sec);
    return isNaN(d.getTime()) ? null : d;
  }

  // Handle DD MMM YYYY (e.g. 15 Oct 2026 or 15-Sept-2026)
  const mWord = str.match(/^(\d{1,2})[\s\-]+([a-zA-Z]+)[\s\-]+(\d{4})(?:\s+(\d{1,2}):(\d{1,2}))?/);
  if (mWord) {
    const day = parseInt(mWord[1], 10);
    const mStr = mWord[2].toLowerCase().slice(0, 3);
    const year = parseInt(mWord[3], 10);
    const hour = mWord[4] ? parseInt(mWord[4], 10) : 0;
    const min = mWord[5] ? parseInt(mWord[5], 10) : 0;
    const months: Record<string, number> = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
    };
    if (months[mStr] !== undefined) {
      return new Date(year, months[mStr], day, hour, min);
    }
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

export function parseDateFromCell(cell: any): Date | null {
  return parseAnyDate(cell);
}

/**
 * Normalizes phone numbers to standard 10-digit mobile number string.
 * Strips country codes (+91, 91), leading zeros, spaces, hyphens, scientific notation.
 */
export function normalizePhoneNumber(val: any): string {
  if (val === null || val === undefined) return '';
  let s = String(typeof val === 'object' && val.v !== undefined ? val.v : val).trim();
  if (!s) return '';
  
  // Handle scientific notation e.g. "9.87654E+09" or "9.87654e9"
  if (/^[0-9.]+[eE]\+[0-9]+$/.test(s)) {
    const num = Number(s);
    if (!isNaN(num)) s = Math.round(num).toString();
  }
  // Strip decimals if it ends with .0
  s = s.replace(/\.0+$/, '');
  // Extract all digits
  const digits = s.replace(/\D/g, '');
  if (!digits) return '';
  // If digits are >= 10, take the last 10 digits
  if (digits.length >= 10) {
    return digits.slice(-10);
  }
  return digits;
}

/**
 * Checks if an agency name contains the current project name.
 * e.g., if project is "Legacy Highpoint", returns true for:
 * - "Visited Data | Legacy Highpoint"
 * - "Highpoint"
 * - "Visited Data | Highpoint"
 * but false for:
 * - "Visited Data | Legacy Kairos"
 * - "Google Ads"
 */
export function isAgencyForCurrentProject(agencyName: string, currentProject: string): boolean {
  if (!agencyName || !currentProject) return false;
  const a = agencyName.toLowerCase().trim();
  const p = currentProject.toLowerCase().trim();
  if (a.includes(p)) return true;
  // Strip "legacy " prefix from project name if present
  const coreP = p.replace(/^legacy\s+/, '').trim();
  if (coreP.length >= 3 && a.includes(coreP)) return true;
  return false;
}

/**
 * Strict date comparison: returns true only if dateA is strictly after dateB.
 * If both contain timestamps (hours or minutes > 0), compares full millisecond timestamps.
 * Otherwise, compares calendar days.
 */
export function isDateStrictlyAfter(dateA: Date, dateB: Date): boolean {
  if (!dateA || !dateB || isNaN(dateA.getTime()) || isNaN(dateB.getTime())) return false;
  const hasTimeA = dateA.getHours() !== 0 || dateA.getMinutes() !== 0 || dateA.getSeconds() !== 0;
  const hasTimeB = dateB.getHours() !== 0 || dateB.getMinutes() !== 0 || dateB.getSeconds() !== 0;
  if (hasTimeA && hasTimeB) {
    return dateA.getTime() > dateB.getTime();
  }
  const dayA = new Date(dateA.getFullYear(), dateA.getMonth(), dateA.getDate()).getTime();
  const dayB = new Date(dateB.getFullYear(), dateB.getMonth(), dateB.getDate()).getTime();
  return dayA > dayB;
}

/**
 * Calculates the previous completed Monday to Sunday week range
 */
export function getDefaultWeekRange(): { startDate: string; endDate: string; displayRange: string } {
  const now = new Date();
  const dayOfWeek = now.getDay(); // 0 is Sunday, 1 is Monday...
  const prevSunday = new Date(now);
  const diffToSunday = dayOfWeek === 0 ? 7 : dayOfWeek;
  prevSunday.setDate(now.getDate() - diffToSunday);
  prevSunday.setHours(0, 0, 0, 0);

  const prevMonday = new Date(prevSunday);
  prevMonday.setDate(prevSunday.getDate() - 6);

  const formatYMD = (d: Date) => {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  const startFormatted = formatToDDMMYYYY(formatYMD(prevMonday));
  const endFormatted = formatToDDMMYYYY(formatYMD(prevSunday));

  return {
    startDate: formatYMD(prevMonday),
    endDate: formatYMD(prevSunday),
    displayRange: `(${startFormatted} to ${endFormatted})`
  };
}

/**
 * Parses the Excel file and detects columns having most records as same
 */
export async function detectBucketReportFields(
  file: File, 
  manualStart?: string, 
  manualEnd?: string,
  reportType: 'sales' | 'presales' | 'presales_visits' = 'sales'
): Promise<BucketReportAnalysis> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'array', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][];

        if (!rawRows || rawRows.length === 0) {
          throw new Error("The selected Excel file is empty.");
        }

        // Clean cell values
        const range = utils.decode_range(sheet['!ref'] || 'A1');
        for (let R = 0; R < rawRows.length; R++) {
          const row = rawRows[R];
          if (!row || !Array.isArray(row)) continue;
          const maxC = Math.max(row.length, range.e.c + 1);
          for (let C = 0; C < maxC; C++) {
            const cell_ref = utils.encode_cell({ c: C, r: R });
            const originalCell = sheet[cell_ref];
            if (originalCell) {
              let val = row[C];
              if (val === undefined || val === null) {
                val = originalCell.v !== undefined ? originalCell.v : (originalCell.w !== undefined ? originalCell.w : '');
              }
              if (originalCell.f && String(originalCell.f).toUpperCase().includes('HYPERLINK')) {
                const m = String(originalCell.f).match(/.+,\s*\"?([^\"\)]+)\"?\s*\)/i);
                if (m) val = m[1];
              }
              row[C] = val;
            }
          }
        }

        // Identify Header Row
        let headerIndex = -1;
        const knownHeaderKeywords = [
          'sales executive', 'presales executive', 'presales', 'enquiry level', 'project name', 'visit source', 'channel partner',
          'call status', 'configuration', 'sales', 'assigned to', 'stage', 'sub stage', 'lead sub stage',
          'bucket', 'status', 'project', 'source', 'lead source', 'date', 'customer name', 'lead name', 'caller', 'telecaller', 'mobile no'
        ];

        let bestHeaderScore = -1;
        for (let i = 0; i < Math.min(15, rawRows.length); i++) {
          const row = rawRows[i];
          if (!Array.isArray(row)) continue;
          const strCells = row.map(c => getCellValue(c).toLowerCase().trim()).filter(Boolean);
          if (strCells.length < 2) continue;

          let score = 0;
          for (const cellText of strCells) {
            if (knownHeaderKeywords.some(kw => cellText.includes(kw))) {
              score += 3;
            } else if (cellText.length > 0 && cellText.length < 30) {
              score += 1;
            }
          }

          if (score > bestHeaderScore) {
            bestHeaderScore = score;
            headerIndex = i;
          }
        }

        if (headerIndex === -1) {
          headerIndex = 0;
        }

        const headerRow = rawRows[headerIndex] || [];
        const numCols = Math.max(headerRow.length, range.e.c + 1);
        const dataRows = rawRows.slice(headerIndex + 1).filter(r => Array.isArray(r) && r.some(c => getCellValue(c).trim().length > 0));

        const totalRows = dataRows.length;
        const detectedColumns: DetectedColumn[] = [];

        let suggestedSalesColIdx = -1;
        let suggestedBucketColIdx = -1;
        let suggestedAgencyColIdx = -1;
        let suggestedProjectColIdx = -1;
        let suggestedDateColIdx = -1;
        let suggestedMobileColIdx = -1;
        let detectedProject = '';

        const isPresales = reportType === 'presales' || reportType === 'presales_visits';

        const salesAliases = isPresales
          ? [
              'telecaller', 'caller', 'tele-caller', 'presales executive', 'presales', 'presales user', 'presale',
              'assigned to', 'assigned_to', 'owner', 'executive', 'agent', 'sales executive', 'user'
            ]
          : [
              'sales executive', 'sales', 'sales person', 'sales rep', 'assigned to', 'assigned_to', 'owner', 'executive',
              'agent', 'caller', 'user'
            ];

        const agencyAliases = [
          'agency me', 'agency name', 'agency', 'agency_name', 'agency-name'
        ];

        const bucketAliases = isPresales
          ? [
              'ai lead level', 'lead level', 'enquiry level', 'enquiry_level', 'enquirylevel',
              'lead sub stage', 'sub stage', 'sub_stage', 'lead sub-stage', 'bucket',
              'stage', 'status', 'lead status', 'disposition', 'lead state'
            ]
          : [
              'enquiry level', 'enquiry_level', 'enquirylevel',
              'lead sub stage', 'sub stage', 'sub_stage', 'lead sub-stage', 'bucket',
              'stage', 'status', 'lead status', 'disposition', 'lead state'
            ];
        const projectAliases = [
          'project name', 'project', 'site', 'project (af)', 'project(af)', 'property'
        ];
        const sourceAliases = [
          'visit source', 'lead source', 'source', 'sub source', 'channel', 'enquiry source'
        ];

        // Process each column
        for (let c = 0; c < numCols; c++) {
          const rawHeader = getCellValue(headerRow[c]).trim();
          if (!rawHeader) continue;

          const lowerHeader = rawHeader.toLowerCase();

          // Track mobile column
          if (lowerHeader === 'mobile no' || lowerHeader === 'mobile number' || lowerHeader === 'mobile' || lowerHeader === 'phone') {
            if (suggestedMobileColIdx === -1) suggestedMobileColIdx = c;
          }

          // Exclude columns that shouldn't be dropdown filter fields
          const isExcludedHeader = [
            's.no', 's no', 'sr no', 'sno', 'client id', 'mobile no', 'mobile number', 'secondary mobile no',
            'whatsapp no', 'firm phone', 'email id', 'customer name', 'min budget', 'max budget',
            'remark', 'source description', 'source expiry', 're enquiry count', 'revisit count'
          ].some(kw => lowerHeader === kw || lowerHeader.startsWith(kw));

          if (isExcludedHeader) continue;

          // Track Date column for date range detection
          if (lowerHeader === 'lead date' || lowerHeader === 'date' || lowerHeader === 'created at' || lowerHeader === 'created date') {
            if (suggestedDateColIdx === -1) {
              suggestedDateColIdx = c;
            }
            continue;
          }

          // Skip any date or timestamp columns as dropdown filters
          if (lowerHeader.includes('date') || lowerHeader.includes('time') || lowerHeader.endsWith(' at')) {
            continue;
          }

          const isAiLeadLevel = lowerHeader === 'ai lead level' || lowerHeader === 'ai_lead_level' || lowerHeader.includes('ai lead level');
          const isExactLeadLevel = !isAiLeadLevel && (lowerHeader === 'lead level' || lowerHeader === 'lead_level' || lowerHeader === 'leadlevel');
          const isEnquiryLevel = !isAiLeadLevel && (
            isExactLeadLevel || 
            lowerHeader === 'enquiry level' || 
            lowerHeader.includes('enquiry level') || 
            lowerHeader.includes('lead level') || 
            bucketAliases.some(b => lowerHeader === b)
          );
          const values: string[] = [];
          const counts: Record<string, number> = {};

          for (const row of dataRows) {
            const rawVal = getCellValue(row[c]);
            let val = normalizeRecordValue(rawVal, isEnquiryLevel);
            if (isEnquiryLevel) {
              val = normalizeBucketCasing(val);
            }
            values.push(val);
            counts[val] = (counts[val] || 0) + 1;
          }

          const totalNonEmpty = values.length;
          if (totalNonEmpty === 0) continue;

          // Order unique values naturally, keeping (blank) at the end if present
          const rawUnique = Object.keys(counts);
          const uniqueVals = rawUnique.filter(v => v !== '(blank)');
          if (rawUnique.includes('(blank)')) {
            uniqueVals.push('(blank)');
          }
          const uniqueCount = uniqueVals.length;

          // Check if column has categorical property
          const isCategorical = uniqueCount <= 50 || (totalRows > 15 && (uniqueCount / totalNonEmpty) <= 0.65);

          // Check average character length
          const avgLength = values.reduce((sum, v) => sum + v.length, 0) / (totalNonEmpty || 1);
          if (avgLength > 50) continue;

          // Exclude pure numerical IDs or serial numbers
          const isNumericId = uniqueVals.every(v => /^\d+$/.test(v)) && uniqueCount > 15;
          if (isNumericId) continue;

          // Determine role and precise label
          let role: 'sales' | 'bucket' | 'project' | 'source' | 'agency' | 'other' = 'other';
          let roleLabel = rawHeader;

          const hasCanonicalBucketValue = uniqueVals.some(v => 
            ['open', 'site visit scheduled', 'site visited', 'cold', 'warm', 'hot', 'discard', 'booked', 'scheduled', 'revisited'].includes(v.toLowerCase())
          );
          const hasKnownSalesValue = uniqueVals.some(v => 
            USER_PROJECT_MAPPING[v] || USER_TEAM_MAPPING[v] || ['alex dmello', 'amol patil', 'prasad patne', 'bhavya jain', 'manisha singh', 'smita kad'].includes(v.toLowerCase())
          );
          const hasProjectValue = uniqueVals.some(v => 
            ['legacy ekam', 'legacy milestone', 'legacy kairos', 'aqua life', 'milestone', 'kairos', 'ekam', 'highpoint', 'legacy highpoint'].includes(v.toLowerCase())
          );

          if (isPresales && (lowerHeader === 'telecaller' || lowerHeader === 'caller' || lowerHeader === 'tele-caller')) {
            role = 'sales';
            roleLabel = 'Telecaller (Rows)';
            suggestedSalesColIdx = c;
          } else if (isPresales && (lowerHeader === 'agency me' || agencyAliases.some(a => lowerHeader === a))) {
            role = 'agency';
            roleLabel = 'Agency Name (Sub-Rows)';
            if (suggestedAgencyColIdx === -1 || lowerHeader === 'agency me' || lowerHeader === 'agency name') {
              suggestedAgencyColIdx = c;
            }
          } else if (isExactLeadLevel || (isPresales && !isAiLeadLevel && lowerHeader.includes('lead level'))) {
            role = 'bucket';
            roleLabel = 'Lead Level (Buckets)';
            suggestedBucketColIdx = c;
          } else if (isAiLeadLevel) {
            role = 'other';
            roleLabel = 'AI Lead Level';
          } else if (lowerHeader === 'enquiry level' || isEnquiryLevel || hasCanonicalBucketValue) {
            role = 'bucket';
            roleLabel = isPresales ? 'Lead Level (Buckets)' : 'Enquiry Level / Buckets (Columns)';
            if (suggestedBucketColIdx === -1 || lowerHeader === 'enquiry level' || lowerHeader === 'lead level') {
              suggestedBucketColIdx = c;
            }
          } else if (
            (isPresales && (lowerHeader.includes('presale') || lowerHeader.includes('telecaller') || lowerHeader.includes('caller'))) ||
            (!isPresales && (lowerHeader === 'sales executive' || lowerHeader === 'sales')) ||
            (salesAliases.some(a => lowerHeader === a) && !lowerHeader.includes('sourcing') && !lowerHeader.includes('source')) ||
            hasKnownSalesValue
          ) {
            role = 'sales';
            roleLabel = isPresales ? 'Telecaller (Rows)' : 'Sales Executive (Rows)';
            if (suggestedSalesColIdx === -1) {
              suggestedSalesColIdx = c;
            } else if (isPresales && (lowerHeader === 'telecaller' || lowerHeader === 'caller')) {
              suggestedSalesColIdx = c;
            } else if (!isPresales && lowerHeader === 'sales executive') {
              suggestedSalesColIdx = c;
            }
          } else if (projectAliases.some(a => lowerHeader === a || lowerHeader.includes(a)) || hasProjectValue) {
            role = 'project';
            roleLabel = 'Project Name';
            if (suggestedProjectColIdx === -1 || lowerHeader === 'project name') {
              suggestedProjectColIdx = c;
              if (uniqueVals.length > 0) {
                const sortedProjects = [...uniqueVals].sort((a, b) => counts[b] - counts[a]);
                detectedProject = sortedProjects[0];
              }
            }
          } else if (sourceAliases.some(a => lowerHeader === a || lowerHeader.includes(a))) {
            role = 'source';
            roleLabel = isPresales ? 'Lead Source' : 'Visit Source';
          } else if (lowerHeader.includes('channel partner')) {
            roleLabel = 'Channel Partner';
          } else if (lowerHeader.includes('call status') || lowerHeader.includes('latest call status')) {
            roleLabel = 'Call Status';
          } else if (lowerHeader.includes('configuration')) {
            roleLabel = 'Configuration';
          } else if (lowerHeader.includes('campaign name')) {
            roleLabel = 'Campaign Name';
          } else if (lowerHeader.includes('sourcing manager')) {
            roleLabel = 'Sourcing Manager';
          } else if (lowerHeader.includes('source type')) {
            roleLabel = 'Source Type';
          } else if (lowerHeader.includes('source executive')) {
            roleLabel = 'Source Executive';
          } else if (lowerHeader.includes('enquiry type') || lowerHeader.includes('lead type')) {
            roleLabel = 'Lead Type';
          } else if (lowerHeader.includes('buying purpose')) {
            roleLabel = 'Buying Purpose';
          } else if (lowerHeader.includes('preferred location')) {
            roleLabel = 'Preferred Location';
          } else if (lowerHeader.includes('booking plan')) {
            roleLabel = 'Booking Plan Within';
          } else if (lowerHeader.includes('age range')) {
            roleLabel = 'Age Range';
          }

          if (isCategorical || role !== 'other') {
            let sortedUniqueValues = [...uniqueVals].sort((a, b) => counts[b] - counts[a]);

            if (role === 'bucket') {
              const cleanUniqueVals = uniqueVals.filter(v => v.toLowerCase() !== '(blank)' && !isBlankValue(v));
              const targetBucketList = DEFAULT_BUCKET_LIST;
              const lowerCanonical = targetBucketList.map(b => b.toLowerCase());

              targetBucketList.forEach(cb => {
                if (!counts[cb]) counts[cb] = 0;
              });

              const extraBuckets = cleanUniqueVals.filter(v => 
                !lowerCanonical.includes(v.toLowerCase())
              );

              sortedUniqueValues = [...targetBucketList, ...extraBuckets];
            }

            detectedColumns.push({
              colIndex: c,
              headerName: rawHeader,
              role,
              roleLabel,
              uniqueValues: sortedUniqueValues,
              valueCounts: counts,
              totalCount: totalNonEmpty
            });
          }
        }

        // Fallback for sales/bucket if not detected
        if (suggestedSalesColIdx === -1 && detectedColumns.length > 0) {
          const cand = detectedColumns.find(col => col.role === 'sales') || detectedColumns[0];
          suggestedSalesColIdx = cand.colIndex;
        }

        // Priority for Presales: Explicitly ensure 'Lead Level' (not AI Lead Level) is chosen as bucket
        if (isPresales) {
          const exactLeadLevelCol = detectedColumns.find(col => {
            const h = col.headerName.toLowerCase().trim();
            return (h === 'lead level' || h === 'lead_level' || h === 'leadlevel' || (!h.includes('ai') && h.includes('lead level')));
          });
          if (exactLeadLevelCol) {
            suggestedBucketColIdx = exactLeadLevelCol.colIndex;
            exactLeadLevelCol.role = 'bucket';
            exactLeadLevelCol.roleLabel = 'Lead Level (Buckets)';
            detectedColumns.forEach(col => {
              const h = col.headerName.toLowerCase().trim();
              if (h.includes('ai lead level') && col.colIndex !== exactLeadLevelCol.colIndex) {
                col.role = 'other';
                col.roleLabel = col.headerName;
              }
            });
          }
        }

        if (suggestedBucketColIdx === -1 && detectedColumns.length > 1) {
          const cand = detectedColumns.find(col => col.role === 'bucket' && col.colIndex !== suggestedSalesColIdx) || detectedColumns[1];
          suggestedBucketColIdx = cand.colIndex;
        }

        // Canonical bucket ordering and blank removal for the chosen bucket column
        const bucketColObj = detectedColumns.find(col => col.colIndex === suggestedBucketColIdx);
        if (bucketColObj) {
          bucketColObj.role = 'bucket';
          bucketColObj.roleLabel = isPresales ? 'Lead Level (Buckets)' : 'Enquiry Level / Buckets (Columns)';
          
          const cleanVals = bucketColObj.uniqueValues.filter(v => v.toLowerCase() !== '(blank)' && !isBlankValue(v));
          const lowerCanonical = DEFAULT_BUCKET_LIST.map(b => b.toLowerCase());
          
          DEFAULT_BUCKET_LIST.forEach(cb => {
            if (!bucketColObj.valueCounts[cb]) {
              bucketColObj.valueCounts[cb] = 0;
            }
          });

          const extraVals = cleanVals.filter(v => !lowerCanonical.includes(v.toLowerCase()));
          bucketColObj.uniqueValues = [...DEFAULT_BUCKET_LIST, ...extraVals];
        }

        // Sort detected columns so that Sales/Telecaller, Agency, Bucket, Project, Source appear first
        detectedColumns.sort((a, b) => {
          const order = { sales: 1, agency: 2, bucket: 3, project: 4, source: 5, other: 6 };
          return (order[a.role] || 99) - (order[b.role] || 99);
        });

        // Check file name or content for Project
        if (!detectedProject) {
          const filenameLower = file.name.toLowerCase();
          if (filenameLower.includes('highpoint')) detectedProject = 'Legacy Highpoint';
          else if (filenameLower.includes('ekam')) detectedProject = 'Legacy Ekam';
          else if (filenameLower.includes('milestone')) detectedProject = 'Legacy Milestone';
          else if (filenameLower.includes('kairos')) detectedProject = 'Legacy Kairos';
          else if (filenameLower.includes('aqua')) detectedProject = 'Legacy Aqua Life';
          else detectedProject = 'Legacy Ekam';
        }

        // Detect Dates from the Date column if available
        let detectedStartDate = '';
        let detectedEndDate = '';
        let detectedStartDateYMD = '';
        let detectedEndDateYMD = '';

        if (suggestedDateColIdx !== -1) {
          const parsedDates: Date[] = [];
          for (const row of dataRows) {
            const dt = parseDateFromCell(row[suggestedDateColIdx]);
            if (dt && !isNaN(dt.getTime())) {
              parsedDates.push(dt);
            }
          }

          if (parsedDates.length > 0) {
            parsedDates.sort((a, b) => a.getTime() - b.getTime());
            const minDate = parsedDates[0];
            const maxDate = parsedDates[parsedDates.length - 1];

            const dayOfWk = maxDate.getDay();
            const wkSunday = new Date(maxDate);
            if (dayOfWk !== 0) {
              wkSunday.setDate(maxDate.getDate() + (7 - dayOfWk));
            }
            const wkMonday = new Date(wkSunday);
            wkMonday.setDate(wkSunday.getDate() - 6);

            const formatYMD = (d: Date) => {
              const year = d.getFullYear();
              const month = String(d.getMonth() + 1).padStart(2, '0');
              const day = String(d.getDate()).padStart(2, '0');
              return `${year}-${month}-${day}`;
            };

            detectedStartDateYMD = formatYMD(wkMonday);
            detectedEndDateYMD = formatYMD(wkSunday);
            detectedStartDate = formatToDDMMYYYY(detectedStartDateYMD);
            detectedEndDate = formatToDDMMYYYY(detectedEndDateYMD);
          }
        }

        if (!detectedStartDate || !detectedEndDate) {
          const defaultWeek = getDefaultWeekRange();
          detectedStartDateYMD = defaultWeek.startDate;
          detectedEndDateYMD = defaultWeek.endDate;
          detectedStartDate = formatToDDMMYYYY(defaultWeek.startDate);
          detectedEndDate = formatToDDMMYYYY(defaultWeek.endDate);
        }

        const startDisp = manualStart ? formatToDDMMYYYY(manualStart) : detectedStartDate;
        const endDisp = manualEnd ? formatToDDMMYYYY(manualEnd) : detectedEndDate;

        let defaultTitle: string;
        if (isPresales) {
          if (manualStart || manualEnd) {
            defaultTitle = `Presales Leads Report | ${detectedProject} | Week Report (${startDisp} to ${endDisp})`;
          } else {
            const now = new Date();
            const monthsArr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
            const dayStr = String(now.getDate()).padStart(2, '0');
            const monStr = monthsArr[now.getMonth()];
            const yrStr = now.getFullYear();
            let hours = now.getHours();
            const ampm = hours >= 12 ? 'PM' : 'AM';
            hours = hours % 12;
            hours = hours ? hours : 12;
            const minStr = String(now.getMinutes()).padStart(2, '0');
            const timeStr = `${dayStr} ${monStr} ${yrStr} ${String(hours).padStart(2, '0')}:${minStr} ${ampm}`;
            defaultTitle = `Presales Leads Report | ${detectedProject} | ${timeStr} | All Time`;
          }
        } else {
          defaultTitle = `Site Visits Report | ${detectedProject} | Week Report (${startDisp} to ${endDisp})`;
        }

        resolve({
          headerIndex,
          detectedColumns,
          suggestedSalesColIdx,
          suggestedBucketColIdx,
          suggestedAgencyColIdx,
          suggestedProjectColIdx,
          suggestedDateColIdx,
          suggestedMobileColIdx,
          detectedProject,
          detectedStartDate: startDisp,
          detectedEndDate: endDisp,
          detectedStartDateYMD,
          detectedEndDateYMD,
          defaultTitle,
          totalRows
        });
      } catch (err: any) {
        console.error("Error detecting bucket report columns:", err);
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Parses the Visits Excel file and detects categorical columns for filtering (Project Name, Visit Source, etc.)
 */
export async function detectVisitsReportFields(file: File): Promise<VisitsReportAnalysis> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'array', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][];

        if (!rawRows || rawRows.length === 0) {
          throw new Error("The selected Visits Excel file is empty.");
        }

        const range = utils.decode_range(sheet['!ref'] || 'A1');
        for (let R = 0; R < rawRows.length; R++) {
          const row = rawRows[R];
          if (!row || !Array.isArray(row)) continue;
          const maxC = Math.max(row.length, range.e.c + 1);
          for (let C = 0; C < maxC; C++) {
            const cell_ref = utils.encode_cell({ c: C, r: R });
            const originalCell = sheet[cell_ref];
            if (originalCell) {
              let val = row[C];
              if (val === undefined || val === null) {
                val = originalCell.v !== undefined ? originalCell.v : (originalCell.w !== undefined ? originalCell.w : '');
              }
              row[C] = val;
            }
          }
        }

        // Identify header row
        let headerIndex = -1;
        const knownKeywords = [
          's.no', 'client id', 'project name', 'mobile no', 'revisit date', 'revisit count',
          'sales executive', 'enquiry level', 'call status', 'digital agency', 'visit source', 'channel partner'
        ];
        let bestScore = -1;
        for (let i = 0; i < Math.min(15, rawRows.length); i++) {
          const row = rawRows[i];
          if (!Array.isArray(row)) continue;
          const strCells = row.map(c => getCellValue(c).toLowerCase().trim()).filter(Boolean);
          let score = 0;
          for (const cell of strCells) {
            if (knownKeywords.some(kw => cell === kw || cell.includes(kw))) {
              score += 3;
            } else if (cell.length > 0 && cell.length < 30) {
              score += 1;
            }
          }
          if (score > bestScore) {
            bestScore = score;
            headerIndex = i;
          }
        }
        if (headerIndex === -1) headerIndex = 0;

        const headerRow = rawRows[headerIndex] || [];
        const numCols = Math.max(headerRow.length, range.e.c + 1);
        const dataRows = rawRows.slice(headerIndex + 1).filter(r => Array.isArray(r) && r.some(c => getCellValue(c).trim().length > 0));

        let suggestedProjectColIdx = -1;
        let suggestedMobileColIdx = -1;
        let suggestedRevisitDateColIdx = -1;
        let suggestedDateColIdx = -1;
        let detectedProject = '';

        const detectedColumns: DetectedColumn[] = [];

        // Excluded headers from being dropdown filters
        const excludedHeaders = [
          's.no', 'sr no', 'sno', 'client id', 'mobile no', 'mobile number', 'secondary mobile no',
          'whatsapp no', 'firm phone', 'email id', 'customer name', 'min budget', 'max budget',
          'remark', 'source description', 'revisit date', 'date', 'follow-up date', 'expiry date',
          'source expiry', 'created at', 'updated at', 're enquiry count', 'revisit count'
        ];

        for (let c = 0; c < numCols; c++) {
          const rawHeader = getCellValue(headerRow[c]).trim();
          if (!rawHeader) continue;
          const lowerHeader = rawHeader.toLowerCase();

          if (lowerHeader === 'mobile no' || lowerHeader === 'mobile number' || lowerHeader === 'mobile' || lowerHeader === 'phone') {
            if (suggestedMobileColIdx === -1) suggestedMobileColIdx = c;
          }
          if (lowerHeader === 'revisit date' || lowerHeader === 're-visit date') {
            if (suggestedRevisitDateColIdx === -1) suggestedRevisitDateColIdx = c;
          }
          if (lowerHeader === 'date' || lowerHeader === 'visit date') {
            if (suggestedDateColIdx === -1) suggestedDateColIdx = c;
          }

          if (excludedHeaders.some(kw => lowerHeader === kw || lowerHeader.startsWith(kw))) {
            continue;
          }

          if (lowerHeader.includes('date') || lowerHeader.includes('time') || lowerHeader.endsWith(' at')) {
            continue;
          }

          const values: string[] = [];
          const counts: Record<string, number> = {};

          for (const row of dataRows) {
            const rawVal = getCellValue(row[c]);
            const val = normalizeRecordValue(rawVal, false);
            values.push(val);
            counts[val] = (counts[val] || 0) + 1;
          }

          const totalNonEmpty = values.length;
          if (totalNonEmpty === 0) continue;

          const rawUnique = Object.keys(counts);
          const uniqueVals = rawUnique.filter(v => v !== '(blank)');
          if (rawUnique.includes('(blank)')) {
            uniqueVals.push('(blank)');
          }
          const uniqueCount = uniqueVals.length;

          // Categorical check
          const isCategorical = uniqueCount <= 50 || (dataRows.length > 15 && (uniqueCount / totalNonEmpty) <= 0.65);
          const avgLength = values.reduce((sum, v) => sum + v.length, 0) / (totalNonEmpty || 1);
          if (avgLength > 50) continue;

          let role: 'sales' | 'bucket' | 'project' | 'source' | 'agency' | 'other' = 'other';
          let roleLabel = rawHeader;

          if (lowerHeader === 'project name' || lowerHeader === 'project' || lowerHeader.includes('project')) {
            role = 'project';
            roleLabel = 'Project Name (Visits)';
            if (suggestedProjectColIdx === -1 || lowerHeader === 'project name') {
              suggestedProjectColIdx = c;
              if (uniqueVals.length > 0) {
                const sortedProjects = [...uniqueVals].sort((a, b) => counts[b] - counts[a]);
                detectedProject = sortedProjects[0];
              }
            }
          } else if (lowerHeader.includes('visit source') || lowerHeader === 'source') {
            role = 'source';
            roleLabel = 'Visit Source';
          } else if (lowerHeader.includes('digital agency')) {
            role = 'agency';
            roleLabel = 'Digital Agency';
          } else if (lowerHeader.includes('sales executive') || lowerHeader === 'sales') {
            role = 'sales';
            roleLabel = 'Sales Executive';
          } else if (lowerHeader.includes('enquiry level')) {
            role = 'bucket';
            roleLabel = 'Enquiry Level';
          } else if (lowerHeader.includes('channel partner')) {
            roleLabel = 'Channel Partner';
          } else if (lowerHeader.includes('call status')) {
            roleLabel = 'Call Status';
          } else if (lowerHeader.includes('configuration')) {
            roleLabel = 'Configuration';
          } else if (lowerHeader.includes('sourcing manager')) {
            roleLabel = 'Sourcing Manager';
          }

          if (isCategorical || role !== 'other') {
            const sortedUniqueValues = [...uniqueVals].sort((a, b) => counts[b] - counts[a]);
            detectedColumns.push({
              colIndex: c,
              headerName: rawHeader,
              role,
              roleLabel,
              uniqueValues: sortedUniqueValues,
              valueCounts: counts,
              totalCount: totalNonEmpty
            });
          }
        }

        // Sort detected columns so Project Name comes first
        detectedColumns.sort((a, b) => {
          const order = { project: 1, source: 2, agency: 3, sales: 4, bucket: 5, other: 6 };
          return (order[a.role] || 99) - (order[b.role] || 99);
        });

        resolve({
          headerIndex,
          detectedColumns,
          suggestedProjectColIdx,
          suggestedMobileColIdx,
          suggestedRevisitDateColIdx,
          suggestedDateColIdx,
          detectedProject,
          totalRows: dataRows.length
        });
      } catch (err: any) {
        console.error("Error detecting visits report columns:", err);
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

interface VisitRecord {
  mobile: string;
  revisitDate: Date | null;
  visitDate: Date | null;
  projectName: string;
}

/**
 * Computes the aggregated table data according to user selections.
 * If visitsFile is provided, calculates the "Visits Done" column based on mobile numbers and revisit date logic.
 */
export async function computeBucketReportTable(
  file: File,
  options: {
    salesColIdx: number;
    bucketColIdx: number;
    agencyColIdx?: number;
    selectedSales: string[];
    selectedBuckets: string[];
    selectedAgencies?: string[];
    columnFilters: Record<number, string[]>;
    reportTitle: string;
    headerIndex: number;
    dimensionLabel?: string;
    reportType?: 'sales' | 'presales' | 'presales_visits';
    visitsFile?: File | null;
    visitsColumnFilters?: Record<number, string[]>;
  }
): Promise<BucketTableSummary> {
  return new Promise(async (resolve, reject) => {
    try {
      const { 
        salesColIdx, 
        bucketColIdx, 
        agencyColIdx, 
        selectedSales, 
        selectedBuckets, 
        columnFilters, 
        reportTitle, 
        headerIndex, 
        reportType,
        visitsFile,
        visitsColumnFilters
      } = options;

      const isPresales = reportType === 'presales' || reportType === 'presales_visits';
      const wantsVisitsDone = reportType === 'presales_visits' || !!visitsFile;
      const dimensionLabel = options.dimensionLabel || (isPresales ? 'Telecaller' : 'Sales');

      // 1. Read Leads Excel file
      const leadsArrayBuffer = await file.arrayBuffer();
      const leadsWorkbook = read(leadsArrayBuffer, { type: 'array', cellDates: true });
      const leadsSheet = leadsWorkbook.Sheets[leadsWorkbook.SheetNames[0]];
      const rawRows = utils.sheet_to_json(leadsSheet, { header: 1, raw: true }) as any[][];

      const headerRow = rawRows[headerIndex] || [];
      const dataRows = rawRows.slice(headerIndex + 1);

      const hasAgencyHierarchy = isPresales && agencyColIdx !== undefined && agencyColIdx !== -1;

      // Identify key columns in Leads Excel for Visits calculation
      let leadsMobileColIdx = -1;
      let leadsDateColIdx = -1;
      let leadsProjectColIdx = -1;

      headerRow.forEach((cell, idx) => {
        const h = getCellValue(cell).toLowerCase().trim();
        if (leadsMobileColIdx === -1 && (h === 'mobile no' || h === 'mobile number' || h === 'mobile' || h === 'phone' || h.includes('mobile') || h.includes('contact'))) {
          if (!h.includes('secondary') && !h.includes('firm') && !h.includes('whatsapp')) {
            leadsMobileColIdx = idx;
          }
        }
        if (leadsDateColIdx === -1 && (h === 'lead date' || h === 'date' || h === 'created at' || h === 'created date' || h.includes('lead date') || h === 'enquiry date')) {
          if (!h.includes('follow') && !h.includes('expiry') && !h.includes('revisit')) {
            leadsDateColIdx = idx;
          }
        }
        if (leadsProjectColIdx === -1 && (h === 'project name' || h === 'project' || h === 'site')) {
          leadsProjectColIdx = idx;
        }
      });

      // 2. Parse Visits Excel file if provided
      const visitsByMobile = new Map<string, VisitRecord[]>();
      let hasValidVisitsData = false;
      let visitMobileColIdx = -1;
      let visitRevisitDateColIdx = -1;
      let visitDateColIdx = -1;
      let visitProjectColIdx = -1;

      if (wantsVisitsDone && visitsFile) {
        try {
          const visitsArrayBuffer = await visitsFile.arrayBuffer();
          const visitsWorkbook = read(visitsArrayBuffer, { type: 'array', cellDates: true });
          const visitsSheet = visitsWorkbook.Sheets[visitsWorkbook.SheetNames[0]];
          const rawVisitsRows = utils.sheet_to_json(visitsSheet, { header: 1, raw: true }) as any[][];

          if (rawVisitsRows && rawVisitsRows.length > 0) {
            // Find visits header row
            let visitsHeaderIdx = 0;
            for (let i = 0; i < Math.min(10, rawVisitsRows.length); i++) {
              const r = rawVisitsRows[i];
              if (Array.isArray(r) && r.some(c => {
                const s = getCellValue(c).toLowerCase().trim();
                return s === 'mobile no' || s === 'client id' || s === 'revisit date' || s === 'sales executive';
              })) {
                visitsHeaderIdx = i;
                break;
              }
            }

            const visitsHeaderRow = rawVisitsRows[visitsHeaderIdx] || [];

            visitsHeaderRow.forEach((cell, idx) => {
              const h = getCellValue(cell).toLowerCase().trim();
              if (visitMobileColIdx === -1 && (h === 'mobile no' || h === 'mobile number' || h === 'mobile' || h === 'phone')) {
                visitMobileColIdx = idx;
              }
              if (visitRevisitDateColIdx === -1 && (h === 'revisit date' || h === 're-visit date' || h.includes('revisit date') || h.includes('re-visit date'))) {
                visitRevisitDateColIdx = idx;
              }
              if (visitDateColIdx === -1 && (h === 'date' || h === 'visit date') && !h.includes('follow') && !h.includes('revisit') && !h.includes('expiry') && !h.includes('created')) {
                visitDateColIdx = idx;
              }
              if (visitProjectColIdx === -1 && (h === 'project name' || h === 'project' || h === 'site')) {
                visitProjectColIdx = idx;
              }
            });

            // Fallback for visits mobile column if not exact
            if (visitMobileColIdx === -1) {
              visitMobileColIdx = visitsHeaderRow.findIndex(c => {
                const h = getCellValue(c).toLowerCase().trim();
                return h.includes('mobile') && !h.includes('secondary') && !h.includes('whatsapp');
              });
            }

            for (let i = visitsHeaderIdx + 1; i < rawVisitsRows.length; i++) {
              const vRow = rawVisitsRows[i];
              if (!Array.isArray(vRow) || vRow.length === 0) continue;

              // Apply visitsColumnFilters (including Project selection)
              if (visitsColumnFilters) {
                let passesVisitsFilters = true;
                for (const colIdxStr of Object.keys(visitsColumnFilters)) {
                  const cIdx = parseInt(colIdxStr, 10);
                  const allowed = visitsColumnFilters[cIdx];
                  if (allowed && allowed.length > 0) {
                    const val = normalizeRecordValue(getCellValue(vRow[cIdx]), false);
                    if (!allowed.includes(val)) {
                      passesVisitsFilters = false;
                      break;
                    }
                  }
                }
                if (!passesVisitsFilters) continue;
              }

              const rawMobile = visitMobileColIdx !== -1 ? vRow[visitMobileColIdx] : null;
              const normMobile = normalizePhoneNumber(rawMobile);
              if (!normMobile) continue;

              const revisitDate = visitRevisitDateColIdx !== -1 ? parseAnyDate(vRow[visitRevisitDateColIdx]) : null;
              const visitDate = visitDateColIdx !== -1 ? parseAnyDate(vRow[visitDateColIdx]) : null;
              const projectName = visitProjectColIdx !== -1 ? getCellValue(vRow[visitProjectColIdx]).trim() : '';

              const rec: VisitRecord = { mobile: normMobile, revisitDate, visitDate, projectName };
              const existing = visitsByMobile.get(normMobile) || [];
              existing.push(rec);
              visitsByMobile.set(normMobile, existing);
            }

            hasValidVisitsData = true;
          }
        } catch (vErr) {
          console.error("Error processing visits file:", vErr);
        }
      }

      // Determine current project name from Leads Excel filters or Visits Excel filters
      let currentProjectName = '';
      if (leadsProjectColIdx !== -1 && columnFilters[leadsProjectColIdx] && columnFilters[leadsProjectColIdx].length === 1) {
        currentProjectName = columnFilters[leadsProjectColIdx][0];
      }
      if (!currentProjectName && visitProjectColIdx !== -1 && visitsColumnFilters?.[visitProjectColIdx]?.length === 1) {
        currentProjectName = visitsColumnFilters[visitProjectColIdx][0];
      }
      if (!currentProjectName && reportTitle) {
        const parts = reportTitle.split('|').map(p => p.trim());
        if (parts.length >= 2 && parts[1]) {
          currentProjectName = parts[1];
        }
      }

      // Clean selectedBuckets: remove any '(blank)' and sort strictly by canonical DEFAULT_BUCKET_LIST order
      const lowerCanonical = DEFAULT_BUCKET_LIST.map(b => b.toLowerCase());
      const cleanedBuckets = (selectedBuckets || [])
        .filter(b => b.toLowerCase() !== '(blank)' && !isBlankValue(b))
        .sort((a, b) => {
          const idxA = lowerCanonical.indexOf(a.toLowerCase());
          const idxB = lowerCanonical.indexOf(b.toLowerCase());
          if (idxA !== -1 && idxB !== -1) return idxA - idxB;
          if (idxA !== -1) return -1;
          if (idxB !== -1) return 1;
          return a.localeCompare(b);
        });

      // Map of salesUser -> bucket -> count
      const userBucketMap: Record<string, Record<string, number>> = {};
      // Map of salesUser -> agency -> bucket -> count
      const userAgencyBucketMap: Record<string, Record<string, Record<string, number>>> = {};
      // Map of salesUser -> agency -> Map of mobile -> Array of leadDates
      const userAgencyLeadsMap: Record<string, Record<string, Map<string, (Date | null)[]>>> = {};
      
      // Initialize all selected sales users with 0 for all cleaned buckets
      selectedSales.forEach(user => {
        userBucketMap[user] = {};
        userAgencyBucketMap[user] = {};
        userAgencyLeadsMap[user] = {};
        cleanedBuckets.forEach(b => {
          userBucketMap[user][b] = 0;
        });
      });

      for (const row of dataRows) {
        if (!Array.isArray(row)) continue;

        // Apply all active column filters
        let passesFilters = true;
        for (const colIdxStr of Object.keys(columnFilters)) {
          const cIdx = parseInt(colIdxStr, 10);
          const allowed = columnFilters[cIdx];
          if (allowed && allowed.length > 0) {
            const val = normalizeRecordValue(getCellValue(row[cIdx]), cIdx === bucketColIdx);
            if (!allowed.includes(val)) {
              passesFilters = false;
              break;
            }
          }
        }
        if (!passesFilters) continue;

        const rawUser = normalizeRecordValue(getCellValue(row[salesColIdx]), false);

        // Find matching sales user
        const matchedUser = selectedSales.find(u => u.toLowerCase() === rawUser.toLowerCase());
        if (!matchedUser) continue;

        // If current project name not yet determined, derive from row
        if (!currentProjectName && leadsProjectColIdx !== -1) {
          const pVal = getCellValue(row[leadsProjectColIdx]).trim();
          if (pVal) currentProjectName = pVal;
        }

        // Normalize Agency Name: blank/NA/null/- is '(blank)'
        let rawAgency = '(blank)';
        if (hasAgencyHierarchy) {
          rawAgency = normalizeRecordValue(getCellValue(row[agencyColIdx!]), false);
          if (options.selectedAgencies && options.selectedAgencies.length > 0) {
            const matchedAgencyFilter = options.selectedAgencies.some(a => a.toLowerCase() === rawAgency.toLowerCase());
            if (!matchedAgencyFilter) continue;
          }
        }

        // Normalize bucket: blank/NA/null/- is considered Open for Enquiry Level & Lead Level
        let rawBucket = normalizeRecordValue(getCellValue(row[bucketColIdx]), true);
        rawBucket = normalizeBucketCasing(rawBucket);

        // Match bucket case-insensitively or normalized
        const matchedBucket = cleanedBuckets.find(b => b.toLowerCase().trim() === rawBucket.toLowerCase().trim());
        if (matchedBucket) {
          if (!userBucketMap[matchedUser]) {
            userBucketMap[matchedUser] = {};
          }
          userBucketMap[matchedUser][matchedBucket] = (userBucketMap[matchedUser][matchedBucket] || 0) + 1;

          if (hasAgencyHierarchy) {
            if (!userAgencyBucketMap[matchedUser]) {
              userAgencyBucketMap[matchedUser] = {};
            }
            if (!userAgencyBucketMap[matchedUser][rawAgency]) {
              userAgencyBucketMap[matchedUser][rawAgency] = {};
              cleanedBuckets.forEach(b => {
                userAgencyBucketMap[matchedUser][rawAgency][b] = 0;
              });
            }
            userAgencyBucketMap[matchedUser][rawAgency][matchedBucket] = (userAgencyBucketMap[matchedUser][rawAgency][matchedBucket] || 0) + 1;
          }
        }

        // Track lead mobile and lead date for Visits Done calculation
        if (wantsVisitsDone) {
          const rawMobile = leadsMobileColIdx !== -1 ? row[leadsMobileColIdx] : null;
          const leadMobile = normalizePhoneNumber(rawMobile);
          const leadDate = leadsDateColIdx !== -1 ? parseAnyDate(row[leadsDateColIdx]) : null;

          if (leadMobile) {
            if (!userAgencyLeadsMap[matchedUser]) userAgencyLeadsMap[matchedUser] = {};
            if (!userAgencyLeadsMap[matchedUser][rawAgency]) userAgencyLeadsMap[matchedUser][rawAgency] = new Map();
            const dateList = userAgencyLeadsMap[matchedUser][rawAgency].get(leadMobile) || [];
            dateList.push(leadDate);
            userAgencyLeadsMap[matchedUser][rawAgency].set(leadMobile, dateList);
          }
        }
      }

      // Sort telecallers / sales users alphabetically
      const sortedUsers = [...selectedSales].sort((a, b) => a.localeCompare(b));

      // Build rows
      const rows: BucketRowData[] = sortedUsers.map(user => {
        const counts = userBucketMap[user] || {};
        let grandTotal = 0;
        cleanedBuckets.forEach(b => {
          grandTotal += (counts[b] || 0);
        });

        if (hasAgencyHierarchy) {
          const agencyMap = userAgencyBucketMap[user] || {};
          const agencyKeys = Object.keys(agencyMap);

          // Sort agencies: '(blank)' comes first, then alphabetically
          agencyKeys.sort((a, b) => {
            if (a.toLowerCase() === '(blank)' && b.toLowerCase() !== '(blank)') return -1;
            if (b.toLowerCase() === '(blank)' && a.toLowerCase() !== '(blank)') return 1;
            return a.localeCompare(b);
          });

          const subRows: BucketSubRowData[] = agencyKeys
            .map(agency => {
              const bCounts = agencyMap[agency] || {};
              let agencyGrandTotal = 0;
              cleanedBuckets.forEach(b => {
                agencyGrandTotal += (bCounts[b] || 0);
              });

              // Calculate Visits Done for this subrow
              let agencyVisitsDone = 0;
              if (wantsVisitsDone && hasValidVisitsData) {
                const isSameProjectAgency = isAgencyForCurrentProject(agency, currentProjectName);
                const mobileMap = userAgencyLeadsMap[user]?.[agency];

                if (mobileMap) {
                  for (const [mobile, leadDates] of mobileMap.entries()) {
                    const visits = visitsByMobile.get(mobile);
                    if (!visits || visits.length === 0) continue;

                    if (isSameProjectAgency) {
                      // Same project agency rule:
                      // Consider visit only if Revisit Date is after Lead Date in Leads Excel.
                      // Don't consider if Revisit Date is empty or before Lead Date.
                      const validLeadDates = leadDates.filter((d): d is Date => d !== null && !isNaN(d.getTime()));
                      if (validLeadDates.length === 0) continue;
                      const minLeadDate = new Date(Math.min(...validLeadDates.map(d => d.getTime())));

                      for (const v of visits) {
                        if (v.revisitDate && isDateStrictlyAfter(v.revisitDate, minLeadDate)) {
                          agencyVisitsDone++;
                        }
                      }
                    } else {
                      // Regular agency rule:
                      // Visits will be counted using mobile number. Count matching visits.
                      agencyVisitsDone += visits.length;
                    }
                  }
                }
              }

              return {
                agency,
                grandTotal: agencyGrandTotal,
                bucketCounts: bCounts,
                visitsDone: wantsVisitsDone ? agencyVisitsDone : undefined
              };
            })
            .filter(sr => sr.grandTotal > 0 || agencyKeys.length === 1);

          const userVisitsDone = subRows.reduce((sum, s) => sum + (s.visitsDone || 0), 0);

          return {
            salesUser: user,
            grandTotal,
            bucketCounts: counts,
            isGroupHeader: true,
            subRows,
            visitsDone: wantsVisitsDone ? userVisitsDone : undefined
          };
        }

        // Telecaller without subrows
        let userDirectVisitsDone = 0;
        if (wantsVisitsDone && hasValidVisitsData) {
          const mobileMap = userAgencyLeadsMap[user]?.['(blank)'];
          if (mobileMap) {
            for (const [mobile, leadDates] of mobileMap.entries()) {
              const visits = visitsByMobile.get(mobile);
              if (!visits || visits.length === 0) continue;
              userDirectVisitsDone += visits.length;
            }
          }
        }

        return {
          salesUser: user,
          grandTotal,
          bucketCounts: counts,
          visitsDone: wantsVisitsDone ? userDirectVisitsDone : undefined
        };
      }).filter(r => r.grandTotal > 0);

      // Compute Column Totals
      let overallGrandTotal = 0;
      let overallVisitsDone = 0;
      const bucketTotals: Record<string, number> = {};
      cleanedBuckets.forEach(b => {
        bucketTotals[b] = 0;
      });

      rows.forEach(r => {
        overallGrandTotal += r.grandTotal;
        if (wantsVisitsDone && r.visitsDone !== undefined) {
          overallVisitsDone += r.visitsDone;
        }
        cleanedBuckets.forEach(b => {
          bucketTotals[b] += (r.bucketCounts[b] || 0);
        });
      });

      const columns = [dimensionLabel, 'Grand Total', ...cleanedBuckets];
      if (wantsVisitsDone) {
        columns.push('Visits Done');
      }

      resolve({
        reportTitle,
        columns,
        buckets: cleanedBuckets,
        rows,
        columnTotals: {
          grandTotal: overallGrandTotal,
          bucketTotals,
          visitsDoneTotal: wantsVisitsDone ? overallVisitsDone : undefined
        },
        dimensionLabel,
        hasVisitsDone: wantsVisitsDone
      });
    } catch (err: any) {
      reject(err);
    }
  });
}

/**
 * Renders the pixel-perfect table container as PNG matching the clean report design.
 * If summary.hasVisitsDone is true, includes "Visits Done" column with distinct visual separation.
 */
export async function generateBucketReportImage(summary: BucketTableSummary): Promise<string> {
  const container = document.createElement('div');
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    backgroundColor: '#ffffff',
    padding: '24px',
    fontFamily: 'Arial, Helvetica, sans-serif',
    color: '#000000',
    zIndex: '-9999',
    pointerEvents: 'none',
    boxSizing: 'border-box',
    display: 'inline-block'
  });

  const { reportTitle, buckets, rows, columnTotals, hasVisitsDone } = summary;
  const dimHeader = summary.dimensionLabel || 'Sales';
  const totalCols = 2 + buckets.length + (hasVisitsDone ? 1 : 0);

  const headerCellsHtml = buckets.map(b => `
    <th style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; font-weight: bold; text-align: center; color: #000000; background-color: #ffffff; white-space: nowrap; line-height: 1.2;">${b}</th>
  `).join('');

  const visitsHeaderHtml = hasVisitsDone ? `
    <th style="border: 1px solid #000000; border-left: 3px double #000000; padding: 6px 14px; font-size: 13.5px; font-weight: bold; text-align: center; color: #000000; background-color: #f8fafc; white-space: nowrap; line-height: 1.2;">Visits Done</th>
  ` : '';

  const rowCellsHtml = rows.map(r => {
    const rVisitsCell = hasVisitsDone ? `
      <td style="border: 1px solid #000000; border-left: 3px double #000000; padding: 6px 14px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000; background-color: #f8fafc;">${r.visitsDone ?? 0}</td>
    ` : '';

    if (r.subRows && r.subRows.length > 0) {
      const headerRowHtml = `
        <tr style="background-color: #ffffff;">
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: left; font-weight: bold; color: #000000; white-space: nowrap;">${r.salesUser}</td>
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${r.grandTotal}</td>
          ${buckets.map(b => `
            <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${r.bucketCounts[b] || 0}</td>
          `).join('')}
          ${rVisitsCell}
        </tr>
      `;
      const subRowsHtml = r.subRows.map(sub => {
        const subVisitsCell = hasVisitsDone ? `
          <td style="border: 1px solid #000000; border-left: 3px double #000000; padding: 5px 14px; font-size: 13px; text-align: center; font-weight: normal; color: #000000; background-color: #fafbfc;">${sub.visitsDone ?? 0}</td>
        ` : '';

        return `
          <tr style="background-color: #ffffff;">
            <td style="border: 1px solid #000000; padding: 5px 12px; font-size: 13px; text-align: left; font-weight: normal; color: #000000; white-space: nowrap;">${sub.agency}</td>
            <td style="border: 1px solid #000000; padding: 5px 12px; font-size: 13px; text-align: center; font-weight: normal; color: #000000;">${sub.grandTotal}</td>
            ${buckets.map(b => `
              <td style="border: 1px solid #000000; padding: 5px 12px; font-size: 13px; text-align: center; font-weight: normal; color: #000000;">${sub.bucketCounts[b] || 0}</td>
            `).join('')}
            ${subVisitsCell}
          </tr>
        `;
      }).join('');
      return headerRowHtml + subRowsHtml;
    }

    return `
      <tr style="background-color: #ffffff;">
        <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: left; font-weight: normal; color: #000000; white-space: nowrap;">${r.salesUser}</td>
        <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: normal; color: #000000;">${r.grandTotal}</td>
        ${buckets.map(b => `
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: normal; color: #000000;">${r.bucketCounts[b] || 0}</td>
        `).join('')}
        ${rVisitsCell}
      </tr>
    `;
  }).join('');

  const totalBucketCellsHtml = buckets.map(b => `
    <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${columnTotals.bucketTotals[b] || 0}</td>
  `).join('');

  const totalVisitsCellHtml = hasVisitsDone ? `
    <td style="border: 1px solid #000000; border-left: 3px double #000000; padding: 6px 14px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000; background-color: #f8fafc;">${columnTotals.visitsDoneTotal || 0}</td>
  ` : '';

  container.innerHTML = `
    <table style="border-collapse: collapse; border: 2px solid #000000; font-family: Arial, Helvetica, sans-serif; background-color: #ffffff; width: auto; min-width: 750px;">
      <thead>
        <!-- Main Merged Title Header -->
        <tr>
          <th colspan="${totalCols}" style="border: 1px solid #000000; padding: 8px 14px; font-size: 15px; font-weight: bold; text-align: center; color: #000000; background-color: #ffffff; letter-spacing: 0.2px;">
            ${reportTitle}
          </th>
        </tr>
        <!-- Columns Header -->
        <tr>
          <th style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; font-weight: bold; text-align: left; color: #000000; background-color: #ffffff; min-width: 140px;">${dimHeader}</th>
          <th style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; font-weight: bold; text-align: center; color: #000000; background-color: #ffffff; min-width: 90px; white-space: nowrap;">Grand Total</th>
          ${headerCellsHtml}
          ${visitsHeaderHtml}
        </tr>
      </thead>
      <tbody>
        ${rowCellsHtml}
        <!-- Grand Total Footer Row -->
        <tr>
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: left; font-weight: bold; color: #000000;">Grand Total</td>
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${columnTotals.grandTotal}</td>
          ${totalBucketCellsHtml}
          ${totalVisitsCellHtml}
        </tr>
      </tbody>
    </table>
  `;

  document.body.appendChild(container);
  await new Promise(resolve => setTimeout(resolve, 400));

  try {
    const dataUrl = await toPng(container, {
      quality: 1.0,
      pixelRatio: 2,
      backgroundColor: '#ffffff'
    });
    return dataUrl;
  } finally {
    if (document.body.contains(container)) {
      document.body.removeChild(container);
    }
  }
}

/**
 * Creates formatted Excel (.xlsx) file with Excel SUM formulas for Grand Totals and Visits Done
 */
export function generateBucketReportExcel(summary: BucketTableSummary): ArrayBuffer {
  const { reportTitle, buckets, rows, columnTotals, hasVisitsDone } = summary;
  const dimHeader = summary.dimensionLabel || 'Sales';
  const wb = utils.book_new();

  // Excel AOA
  const aoa: any[][] = [];

  // Row 1: Merged Title
  aoa.push([reportTitle]);

  // Row 2: Headers
  const headerRow = [dimHeader, 'Grand Total', ...buckets];
  if (hasVisitsDone) {
    headerRow.push('Visits Done');
  }
  aoa.push(headerRow);

  let currentExcelRow = 3; // 1-indexed in Excel (Row 3 starts first data row)
  const telecallerExcelRows: number[] = [];
  const visitsColLetter = hasVisitsDone ? utils.encode_col(2 + buckets.length) : '';

  rows.forEach(r => {
    if (r.subRows && r.subRows.length > 0) {
      const telecallerRowIdx = currentExcelRow;
      telecallerExcelRows.push(telecallerRowIdx);
      const firstSubRow = currentExcelRow + 1;
      const lastSubRow = currentExcelRow + r.subRows.length;

      // Telecaller Row: Grand Total formula sums subrows
      const telecallerGrandTotalFormula = `SUM(B${firstSubRow}:B${lastSubRow})`;

      const telecallerRowArr: any[] = [
        r.salesUser,
        { f: telecallerGrandTotalFormula, v: r.grandTotal, t: 'n' },
        ...buckets.map((b, bIdx) => {
          const colLetter = utils.encode_col(2 + bIdx);
          return { f: `SUM(${colLetter}${firstSubRow}:${colLetter}${lastSubRow})`, v: r.bucketCounts[b] || 0, t: 'n' };
        })
      ];

      if (hasVisitsDone) {
        telecallerRowArr.push({
          f: `SUM(${visitsColLetter}${firstSubRow}:${visitsColLetter}${lastSubRow})`,
          v: r.visitsDone || 0,
          t: 'n'
        });
      }

      aoa.push(telecallerRowArr);
      currentExcelRow++;

      // Subrows
      const lastBucketColLetter = utils.encode_col(1 + buckets.length);
      r.subRows.forEach(sub => {
        const subRowIdx = currentExcelRow;
        const subGrandTotalFormula = `SUM(C${subRowIdx}:${lastBucketColLetter}${subRowIdx})`;
        const subRowArr: any[] = [
          sub.agency,
          { f: subGrandTotalFormula, v: sub.grandTotal, t: 'n' },
          ...buckets.map(b => sub.bucketCounts[b] || 0)
        ];
        if (hasVisitsDone) {
          subRowArr.push(sub.visitsDone || 0);
        }
        aoa.push(subRowArr);
        currentExcelRow++;
      });
    } else {
      const rowIdx = currentExcelRow;
      telecallerExcelRows.push(rowIdx);
      const lastBucketColLetter = utils.encode_col(1 + buckets.length);
      const formula = `SUM(C${rowIdx}:${lastBucketColLetter}${rowIdx})`;
      const singleRowArr: any[] = [
        r.salesUser,
        { f: formula, v: r.grandTotal, t: 'n' },
        ...buckets.map(b => r.bucketCounts[b] || 0)
      ];
      if (hasVisitsDone) {
        singleRowArr.push(r.visitsDone || 0);
      }
      aoa.push(singleRowArr);
      currentExcelRow++;
    }
  });

  // Footer Row: Grand Total
  const grandTotalRowIdx = currentExcelRow;
  let grandTotalFormula = '';
  if (telecallerExcelRows.length > 0) {
    grandTotalFormula = `SUM(${telecallerExcelRows.map(r => `B${r}`).join(',')})`;
  } else {
    grandTotalFormula = `SUM(B3:B${grandTotalRowIdx - 1})`;
  }

  const footerRow: any[] = [
    'Grand Total',
    { f: grandTotalFormula, v: columnTotals.grandTotal, t: 'n' }
  ];

  buckets.forEach((b, bIdx) => {
    const colLetter = utils.encode_col(2 + bIdx);
    const bucketFormula = telecallerExcelRows.length > 0
      ? `SUM(${telecallerExcelRows.map(r => `${colLetter}${r}`).join(',')})`
      : `SUM(${colLetter}3:${colLetter}${grandTotalRowIdx - 1})`;
    footerRow.push({ f: bucketFormula, v: columnTotals.bucketTotals[b] || 0, t: 'n' });
  });

  if (hasVisitsDone) {
    const visitsTotalFormula = telecallerExcelRows.length > 0
      ? `SUM(${telecallerExcelRows.map(r => `${visitsColLetter}${r}`).join(',')})`
      : `SUM(${visitsColLetter}3:${visitsColLetter}${grandTotalRowIdx - 1})`;
    footerRow.push({ f: visitsTotalFormula, v: columnTotals.visitsDoneTotal || 0, t: 'n' });
  }

  aoa.push(footerRow);

  const ws = utils.aoa_to_sheet(aoa);

  // Set merge for Row 1
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 1 + buckets.length + (hasVisitsDone ? 1 : 0) } }
  ];

  // Set column widths
  const cols = [
    { wch: 28 }, // Telecaller / Agency
    { wch: 14 }, // Grand Total
    ...buckets.map(b => ({ wch: Math.max(b.length + 3, 12) }))
  ];
  if (hasVisitsDone) {
    cols.push({ wch: 14 }); // Visits Done
  }
  ws['!cols'] = cols;

  const sheetName = summary.dimensionLabel === 'Telecaller' || summary.dimensionLabel === 'Presales' ? 'Presales Leads Report' : 'Site Visits Report';
  utils.book_append_sheet(wb, ws, sheetName);
  const buffer = write(wb, { bookType: 'xlsx', type: 'array' });
  return buffer;
}

/**
 * Creates PDF file embedding the clean table
 */
export async function generateBucketReportPDF(summary: BucketTableSummary, pngDataUrl: string): Promise<string> {
  const pdf = new jsPDF({
    orientation: 'landscape',
    unit: 'pt',
    format: 'a4'
  });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const imgWidth = img.width;
      const imgHeight = img.height;

      const margin = 40;
      const maxW = pageWidth - margin * 2;
      const maxH = pageHeight - margin * 2;

      let renderW = maxW;
      let renderH = (imgHeight * renderW) / imgWidth;

      if (renderH > maxH) {
        renderH = maxH;
        renderW = (imgWidth * renderH) / imgHeight;
      }

      const x = (pageWidth - renderW) / 2;
      const y = (pageHeight - renderH) / 2;

      pdf.addImage(pngDataUrl, 'PNG', x, y, renderW, renderH);
      resolve(pdf.output('datauristring'));
    };
    img.src = pngDataUrl;
  });
}

/**
 * Main processor function for Bucket Report (Site Visits, Presales Leads, or Presales Leads with Visits)
 */
export async function processBucketSiteVisitFile(
  file: File,
  options: {
    salesColIdx: number;
    bucketColIdx: number;
    agencyColIdx?: number;
    selectedSales: string[];
    selectedBuckets: string[];
    selectedAgencies?: string[];
    columnFilters: Record<number, string[]>;
    reportTitle: string;
    headerIndex: number;
    dimensionLabel?: string;
    reportType?: 'sales' | 'presales' | 'presales_visits';
    visitsFile?: File | null;
    visitsColumnFilters?: Record<number, string[]>;
  }
): Promise<ProcessResponse> {
  const summary = await computeBucketReportTable(file, options);
  const pngDataUrl = await generateBucketReportImage(summary);
  const pdfDataUrl = await generateBucketReportPDF(summary, pngDataUrl);
  const xlsxBuffer = generateBucketReportExcel(summary);

  const safeTitle = summary.reportTitle.replace(/[^a-z0-9]/gi, '_').toLowerCase();

  const zip = new JSZip();
  zip.file(`${safeTitle}.png`, pngDataUrl.split(',')[1], { base64: true });
  zip.file(`${safeTitle}.pdf`, pdfDataUrl.split(',')[1], { base64: true });
  zip.file(`${safeTitle}.xlsx`, xlsxBuffer);

  const zipBlob = await zip.generateAsync({ type: 'blob' });

  const images: GeneratedImage[] = [
    {
      project_name: summary.reportTitle,
      image_url: pngDataUrl,
      filename: `${safeTitle}.png`
    },
    {
      project_name: summary.reportTitle,
      image_url: pdfDataUrl,
      filename: `${safeTitle}.pdf`
    }
  ];

  return {
    images,
    zip_url: URL.createObjectURL(zipBlob),
    message: summary.hasVisitsDone 
      ? "Bucket Report with Visits Done generated successfully."
      : "Bucket Report generated successfully."
  };
}
