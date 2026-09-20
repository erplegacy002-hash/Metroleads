import { read, utils, write } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING } from './projectMapping';

// Canonical bucket list matching required format for Lead Levels and Enquiry Levels:
// Open, Site Visit Scheduled, Site Visited, Cold, Warm, Hot, Discard, Booked
export const DEFAULT_BUCKET_LIST = [
  'Open',
  'Site Visit Scheduled',
  'Site Visited',
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
    'cold': 'Cold',
    'warm': 'Warm',
    'hot': 'Hot',
    'discard': 'Discard',
    'booked': 'Booked',
    'revisited': 'Revisited',
    're-visited': 'Revisited',
    'lost': 'Lost',
    'junk': 'Junk',
    'drop': 'Drop'
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
  detectedProject: string;
  detectedStartDate: string;
  detectedEndDate: string;
  detectedStartDateYMD?: string;
  detectedEndDateYMD?: string;
  defaultTitle: string;
  totalRows: number;
}

export interface BucketSubRowData {
  agency: string;
  grandTotal: number;
  bucketCounts: Record<string, number>;
}

export interface BucketRowData {
  salesUser: string;
  grandTotal: number;
  bucketCounts: Record<string, number>;
  isGroupHeader?: boolean;
  subRows?: BucketSubRowData[];
}

export interface BucketTableSummary {
  reportTitle: string;
  columns: string[]; // ['Telecaller' | 'Sales', 'Grand Total', ...buckets]
  buckets: string[];
  rows: BucketRowData[];
  columnTotals: {
    grandTotal: number;
    bucketTotals: Record<string, number>;
  };
  dimensionLabel?: string;
}

function getCellValue(cell: any): string {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'object' && cell.v !== undefined) return String(cell.v);
  return String(cell);
}

/**
 * Checks if a cell value represents an empty, missing, NA, or null value:
 * - null or undefined
 * - empty string or whitespace only ("   ")
 * - "-" or "--"
 * - "NA", "N/A", "na", "n/a", "N / A", "n / a"
 * - "null", "NULL"
 * - "undefined"
 * - "(blank)"
 * - "none", "NONE"
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
 * Parses date cell into Date object
 */
export function parseDateFromCell(cell: any): Date | null {
  if (!cell) return null;
  if (cell instanceof Date && !isNaN(cell.getTime())) return cell;
  const str = getCellValue(cell).trim();
  if (!str || str === '-' || str.toLowerCase() === 'null') return null;

  // e.g. "13 Sept 2026", "13 Sep 2026", "13-Sep-2026"
  const mWord = str.match(/^(\d{1,2})[\s\-]+([a-zA-Z]+)[\s\-]+(\d{4})/);
  if (mWord) {
    const d = parseInt(mWord[1], 10);
    const mStr = mWord[2].toLowerCase();
    const y = parseInt(mWord[3], 10);
    const months: Record<string, number> = {
      jan: 0, january: 0,
      feb: 1, february: 1,
      mar: 2, march: 2,
      apr: 3, april: 3,
      may: 4,
      jun: 5, june: 5,
      jul: 6, july: 6,
      aug: 7, august: 7,
      sep: 8, sept: 8, september: 8,
      oct: 9, october: 9,
      nov: 10, november: 10,
      dec: 11, december: 11
    };
    if (months[mStr] !== undefined) {
      return new Date(y, months[mStr], d);
    }
  }

  // e.g. "13-09-2026" or "13/09/2026"
  const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) {
    return new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
  }

  // e.g. "2026-09-13"
  const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymd) {
    return new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10));
  }

  const d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  return null;
}

/**
 * Calculates the previous completed Monday to Sunday week range
 */
export function getDefaultWeekRange(): { startDate: string; endDate: string; displayRange: string } {
  const now = new Date();
  // Find previous Sunday
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
  reportType: 'sales' | 'presales' = 'sales'
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
          'bucket', 'status', 'project', 'source', 'lead source', 'date', 'customer name', 'lead name', 'caller', 'telecaller'
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
        let detectedProject = '';

        const isPresales = reportType === 'presales';

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

          // Exclude columns that shouldn't be dropdown filter fields
          const isExcludedHeader = [
            's.no', 's no', 'sr no', 'sno', 'client id', 'mobile no', 'secondary mobile no',
            'whatsapp no', 'firm phone', 'email id', 'customer name', 'min budget', 'max budget',
            'remark', 'source description', 'source expiry', 're enquiry count', 'revisit count'
          ].some(kw => lowerHeader === kw || lowerHeader.startsWith(kw));

          if (isExcludedHeader) continue;

          // Track Date column for date range detection
          if (lowerHeader === 'lead date' || lowerHeader === 'date' || lowerHeader === 'created at') {
            if (suggestedDateColIdx === -1) {
              suggestedDateColIdx = c;
            }
            continue;
          }

          // Skip any date or timestamp columns as dropdown filters
          if (lowerHeader.includes('date') || lowerHeader.includes('time') || lowerHeader.endsWith(' at')) {
            continue;
          }

          const isExactLeadLevel = lowerHeader === 'lead level' || lowerHeader === 'lead_level' || lowerHeader === 'leadlevel';
          const isAiLeadLevel = lowerHeader === 'ai lead level' || lowerHeader === 'ai_lead_level' || lowerHeader.includes('ai lead level');
          const isEnquiryLevel = isExactLeadLevel || lowerHeader === 'enquiry level' || lowerHeader.includes('enquiry level') || (!isAiLeadLevel && lowerHeader.includes('lead level')) || bucketAliases.some(b => lowerHeader === b);
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

          // Check if column has categorical property (distinct values are limited or ratio is low)
          const isCategorical = uniqueCount <= 50 || (totalRows > 15 && (uniqueCount / totalNonEmpty) <= 0.65);

          // Check average character length (filter out long remarks/notes)
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
            ['legacy ekam', 'legacy milestone', 'legacy kairos', 'aqua life', 'milestone', 'kairos', 'ekam'].includes(v.toLowerCase())
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
            // User requested: "I want Lead Level to be used not AI Lead Level"
            role = 'other';
            roleLabel = 'AI Lead Level';
          } else if (lowerHeader === 'enquiry level' || (!isPresales && (isEnquiryLevel || hasCanonicalBucketValue))) {
            role = 'bucket';
            roleLabel = isPresales ? 'Lead Level (Buckets)' : 'Enquiry Level / Buckets (Columns)';
            if (suggestedBucketColIdx === -1 || lowerHeader === 'enquiry level') {
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
              // Ensure 'Open' is present in unique values
              if (!uniqueVals.includes('Open')) {
                uniqueVals.push('Open');
                if (!counts['Open']) counts['Open'] = 0;
              }

              // Ensure canonical buckets appear in standard order
              const targetBucketList = isPresales ? DEFAULT_PRESALES_BUCKET_LIST : DEFAULT_BUCKET_LIST;
              const lowerCanonical = targetBucketList.map(b => b.toLowerCase());
              targetBucketList.forEach(cb => {
                if (!uniqueVals.some(v => v.toLowerCase() === cb.toLowerCase())) {
                  uniqueVals.push(cb);
                  if (!counts[cb]) counts[cb] = 0;
                }
              });

              sortedUniqueValues = [...uniqueVals].sort((a, b) => {
                const idxA = lowerCanonical.indexOf(a.toLowerCase());
                const idxB = lowerCanonical.indexOf(b.toLowerCase());
                if (idxA !== -1 && idxB !== -1) return idxA - idxB;
                if (idxA !== -1) return -1;
                if (idxB !== -1) return 1;
                return (counts[b] || 0) - (counts[a] || 0);
              });
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
            // Demote any AI Lead Level column
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

        // Sort detected columns so that Sales/Telecaller, Agency, Bucket, Project, Source appear first
        detectedColumns.sort((a, b) => {
          const order = { sales: 1, agency: 2, bucket: 3, project: 4, source: 5, other: 6 };
          return (order[a.role] || 99) - (order[b.role] || 99);
        });

        // Check file name or content for Project
        if (!detectedProject) {
          const filenameLower = file.name.toLowerCase();
          if (filenameLower.includes('ekam')) detectedProject = 'Legacy Ekam';
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

            // If max date is known, derive the Monday to Sunday week
            const dayOfWk = maxDate.getDay(); // 0 is Sunday, 1 is Monday...
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
 * Computes the aggregated table data according to user selections
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
    reportType?: 'sales' | 'presales';
  }
): Promise<BucketTableSummary> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'array', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][];

        const { salesColIdx, bucketColIdx, agencyColIdx, selectedSales, selectedBuckets, columnFilters, reportTitle, headerIndex, reportType } = options;
        const isPresales = reportType === 'presales';
        const dimensionLabel = options.dimensionLabel || (isPresales ? 'Telecaller' : 'Sales');

        const dataRows = rawRows.slice(headerIndex + 1);

        const hasAgencyHierarchy = isPresales && agencyColIdx !== undefined && agencyColIdx !== -1;

        // Map of salesUser -> bucket -> count
        const userBucketMap: Record<string, Record<string, number>> = {};
        // Map of salesUser -> agency -> bucket -> count
        const userAgencyBucketMap: Record<string, Record<string, Record<string, number>>> = {};
        
        // Initialize all selected sales users with 0 for all selected buckets
        selectedSales.forEach(user => {
          userBucketMap[user] = {};
          userAgencyBucketMap[user] = {};
          selectedBuckets.forEach(b => {
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
          const matchedBucket = selectedBuckets.find(b => b.toLowerCase().trim() === rawBucket.toLowerCase().trim());
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
                selectedBuckets.forEach(b => {
                  userAgencyBucketMap[matchedUser][rawAgency][b] = 0;
                });
              }
              userAgencyBucketMap[matchedUser][rawAgency][matchedBucket] = (userAgencyBucketMap[matchedUser][rawAgency][matchedBucket] || 0) + 1;
            }
          }
        }

        // Sort telecallers / sales users alphabetically
        const sortedUsers = [...selectedSales].sort((a, b) => a.localeCompare(b));

        // Build rows
        const rows: BucketRowData[] = sortedUsers.map(user => {
          const counts = userBucketMap[user] || {};
          let grandTotal = 0;
          selectedBuckets.forEach(b => {
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
                selectedBuckets.forEach(b => {
                  agencyGrandTotal += (bCounts[b] || 0);
                });
                return {
                  agency,
                  grandTotal: agencyGrandTotal,
                  bucketCounts: bCounts
                };
              })
              .filter(sr => sr.grandTotal > 0 || agencyKeys.length === 1);

            return {
              salesUser: user,
              grandTotal,
              bucketCounts: counts,
              isGroupHeader: true,
              subRows
            };
          }

          return {
            salesUser: user,
            grandTotal,
            bucketCounts: counts
          };
        });

        // Compute Column Totals
        let overallGrandTotal = 0;
        const bucketTotals: Record<string, number> = {};
        selectedBuckets.forEach(b => {
          bucketTotals[b] = 0;
        });

        rows.forEach(r => {
          overallGrandTotal += r.grandTotal;
          selectedBuckets.forEach(b => {
            bucketTotals[b] += (r.bucketCounts[b] || 0);
          });
        });

        resolve({
          reportTitle,
          columns: [dimensionLabel, 'Grand Total', ...selectedBuckets],
          buckets: selectedBuckets,
          rows,
          columnTotals: {
            grandTotal: overallGrandTotal,
            bucketTotals
          },
          dimensionLabel
        });
      } catch (err: any) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Renders the pixel-perfect table container as PNG matching image.png
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

  const { reportTitle, buckets, rows, columnTotals } = summary;
  const dimHeader = summary.dimensionLabel || 'Sales';
  const totalCols = 2 + buckets.length;

  const headerCellsHtml = buckets.map(b => `
    <th style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; font-weight: bold; text-align: center; color: #000000; background-color: #ffffff; white-space: nowrap; line-height: 1.2;">${b}</th>
  `).join('');

  const rowCellsHtml = rows.map(r => {
    if (r.subRows && r.subRows.length > 0) {
      const headerRowHtml = `
        <tr style="background-color: #ffffff;">
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: left; font-weight: bold; color: #000000; white-space: nowrap;">${r.salesUser}</td>
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${r.grandTotal}</td>
          ${buckets.map(b => `
            <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${r.bucketCounts[b] || 0}</td>
          `).join('')}
        </tr>
      `;
      const subRowsHtml = r.subRows.map(sub => `
        <tr style="background-color: #ffffff;">
          <td style="border: 1px solid #000000; padding: 5px 12px; font-size: 13px; text-align: left; font-weight: normal; color: #000000; white-space: nowrap;">${sub.agency}</td>
          <td style="border: 1px solid #000000; padding: 5px 12px; font-size: 13px; text-align: center; font-weight: normal; color: #000000;">${sub.grandTotal}</td>
          ${buckets.map(b => `
            <td style="border: 1px solid #000000; padding: 5px 12px; font-size: 13px; text-align: center; font-weight: normal; color: #000000;">${sub.bucketCounts[b] || 0}</td>
          `).join('')}
        </tr>
      `).join('');
      return headerRowHtml + subRowsHtml;
    }

    return `
      <tr style="background-color: #ffffff;">
        <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: left; font-weight: normal; color: #000000; white-space: nowrap;">${r.salesUser}</td>
        <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: normal; color: #000000;">${r.grandTotal}</td>
        ${buckets.map(b => `
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: normal; color: #000000;">${r.bucketCounts[b] || 0}</td>
        `).join('')}
      </tr>
    `;
  }).join('');

  const totalBucketCellsHtml = buckets.map(b => `
    <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${columnTotals.bucketTotals[b] || 0}</td>
  `).join('');

  container.innerHTML = `
    <table style="border-collapse: collapse; border: 2px solid #000000; font-family: Arial, Helvetica, sans-serif; background-color: #ffffff; width: auto; min-width: 750px;">
      <thead>
        <!-- Main Merged Title Header matching image.png -->
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
        </tr>
      </thead>
      <tbody>
        ${rowCellsHtml}
        <!-- Grand Total Footer Row -->
        <tr>
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: left; font-weight: bold; color: #000000;">Grand Total</td>
          <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: bold; color: #000000;">${columnTotals.grandTotal}</td>
          ${totalBucketCellsHtml}
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
 * Creates formatted Excel (.xlsx) file with Excel SUM formulas for Grand Totals
 */
export function generateBucketReportExcel(summary: BucketTableSummary): ArrayBuffer {
  const { reportTitle, buckets, rows, columnTotals } = summary;
  const dimHeader = summary.dimensionLabel || 'Sales';
  const wb = utils.book_new();

  // Excel AOA
  const aoa: any[][] = [];

  // Row 1: Merged Title
  aoa.push([reportTitle]);

  // Row 2: Headers
  aoa.push([dimHeader, 'Grand Total', ...buckets]);

  let currentExcelRow = 3; // 1-indexed in Excel (Row 3 starts first data row)
  const telecallerExcelRows: number[] = [];

  rows.forEach(r => {
    if (r.subRows && r.subRows.length > 0) {
      const telecallerRowIdx = currentExcelRow;
      telecallerExcelRows.push(telecallerRowIdx);
      const firstSubRow = currentExcelRow + 1;
      const lastSubRow = currentExcelRow + r.subRows.length;

      // Telecaller Row: Grand Total formula sums subrows
      const lastColLetter = utils.encode_col(1 + buckets.length);
      const telecallerGrandTotalFormula = `SUM(B${firstSubRow}:B${lastSubRow})`;

      const telecallerRowArr: any[] = [
        r.salesUser,
        { f: telecallerGrandTotalFormula, v: r.grandTotal, t: 'n' },
        ...buckets.map((b, bIdx) => {
          const colLetter = utils.encode_col(2 + bIdx);
          return { f: `SUM(${colLetter}${firstSubRow}:${colLetter}${lastSubRow})`, v: r.bucketCounts[b] || 0, t: 'n' };
        })
      ];
      aoa.push(telecallerRowArr);
      currentExcelRow++;

      // Subrows
      r.subRows.forEach(sub => {
        const subRowIdx = currentExcelRow;
        const subGrandTotalFormula = `SUM(C${subRowIdx}:${lastColLetter}${subRowIdx})`;
        aoa.push([
          sub.agency,
          { f: subGrandTotalFormula, v: sub.grandTotal, t: 'n' },
          ...buckets.map(b => sub.bucketCounts[b] || 0)
        ]);
        currentExcelRow++;
      });
    } else {
      const rowIdx = currentExcelRow;
      telecallerExcelRows.push(rowIdx);
      const lastColLetter = utils.encode_col(1 + buckets.length);
      const formula = `SUM(C${rowIdx}:${lastColLetter}${rowIdx})`;
      aoa.push([
        r.salesUser,
        { f: formula, v: r.grandTotal, t: 'n' },
        ...buckets.map(b => r.bucketCounts[b] || 0)
      ]);
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
  aoa.push(footerRow);

  const ws = utils.aoa_to_sheet(aoa);

  // Set merge for Row 1
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 1 + buckets.length } }
  ];

  // Set column widths
  ws['!cols'] = [
    { wch: 28 }, // Telecaller / Agency
    { wch: 14 }, // Grand Total
    ...buckets.map(b => ({ wch: Math.max(b.length + 3, 12) }))
  ];

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

  // Load image to get aspect ratio
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
 * Main processor function for Bucket Report (Site Visits or Presales Leads)
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
    reportType?: 'sales' | 'presales';
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
    message: "Bucket Report generated successfully."
  };
}
