import { read, utils, write } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING } from './projectMapping';

// Canonical bucket list matching image format
export const DEFAULT_BUCKET_LIST = [
  'Open',
  'Site Visit Scheduled',
  'Site Visited',
  'Revisited',
  'Cold',
  'Warm',
  'Hot',
  'Discard'
];

export interface DetectedColumn {
  colIndex: number;
  headerName: string;
  role: 'sales' | 'bucket' | 'project' | 'source' | 'other';
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

export interface BucketRowData {
  salesUser: string;
  grandTotal: number;
  bucketCounts: Record<string, number>;
}

export interface BucketTableSummary {
  reportTitle: string;
  columns: string[]; // ['Sales', 'Grand Total', ...buckets]
  buckets: string[];
  rows: BucketRowData[];
  columnTotals: {
    grandTotal: number;
    bucketTotals: Record<string, number>;
  };
}

function getCellValue(cell: any): string {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'object' && cell.v !== undefined) return String(cell.v);
  return String(cell);
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
export async function detectBucketReportFields(file: File, manualStart?: string, manualEnd?: string): Promise<BucketReportAnalysis> {
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
          'sales executive', 'enquiry level', 'project name', 'visit source', 'channel partner',
          'call status', 'configuration', 'sales', 'assigned to', 'stage', 'sub stage', 'lead sub stage',
          'bucket', 'status', 'project', 'source', 'lead source', 'date', 'customer name', 'lead name'
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
        let suggestedProjectColIdx = -1;
        let suggestedDateColIdx = -1;
        let detectedProject = '';

        const salesAliases = [
          'sales executive', 'sales', 'assigned to', 'assigned_to', 'owner', 'executive',
          'agent', 'sales person', 'caller', 'user'
        ];
        const bucketAliases = [
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
          if (lowerHeader === 'date' || lowerHeader === 'created at') {
            if (suggestedDateColIdx === -1) {
              suggestedDateColIdx = c;
            }
            continue;
          }

          // Skip any date or timestamp columns as dropdown filters
          if (lowerHeader.includes('date') || lowerHeader.includes('time') || lowerHeader.endsWith(' at')) {
            continue;
          }

          const isEnquiryLevel = lowerHeader === 'enquiry level' || lowerHeader.includes('enquiry level') || bucketAliases.some(b => lowerHeader === b);
          const values: string[] = [];
          const counts: Record<string, number> = {};

          for (const row of dataRows) {
            let rawVal = getCellValue(row[c]).trim();
            // Requirement: "consider blank Enquiry Level as Open for the report consideration"
            if (isEnquiryLevel) {
              if (!rawVal || rawVal === '-' || rawVal.toLowerCase() === 'null' || rawVal.toLowerCase() === 'undefined') {
                rawVal = 'Open';
              }
            }
            if (rawVal && rawVal !== '-' && rawVal.toLowerCase() !== 'null' && rawVal.toLowerCase() !== 'undefined') {
              values.push(rawVal);
              counts[rawVal] = (counts[rawVal] || 0) + 1;
            }
          }

          const totalNonEmpty = values.length;
          if (totalNonEmpty === 0) continue;

          const uniqueVals = Object.keys(counts);
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
          let role: 'sales' | 'bucket' | 'project' | 'source' | 'other' = 'other';
          let roleLabel = rawHeader;

          const hasCanonicalBucketValue = uniqueVals.some(v => 
            ['open', 'cold', 'warm', 'hot', 'site visited', 'scheduled', 'revisited', 'discard'].includes(v.toLowerCase())
          );
          const hasKnownSalesValue = uniqueVals.some(v => 
            USER_PROJECT_MAPPING[v] || USER_TEAM_MAPPING[v] || ['alex dmello', 'amol patil', 'prasad patne'].includes(v.toLowerCase())
          );
          const hasProjectValue = uniqueVals.some(v => 
            ['legacy ekam', 'legacy milestone', 'legacy kairos', 'aqua life', 'milestone', 'kairos', 'ekam'].includes(v.toLowerCase())
          );

          if (lowerHeader === 'enquiry level' || isEnquiryLevel || hasCanonicalBucketValue) {
            role = 'bucket';
            roleLabel = 'Enquiry Level / Buckets (Columns)';
            if (suggestedBucketColIdx === -1 || lowerHeader === 'enquiry level') {
              suggestedBucketColIdx = c;
            }
          } else if (lowerHeader === 'sales executive' || (salesAliases.some(a => lowerHeader === a) && !lowerHeader.includes('sourcing') && !lowerHeader.includes('source')) || hasKnownSalesValue) {
            role = 'sales';
            roleLabel = 'Sales Executive (Rows)';
            if (suggestedSalesColIdx === -1 || lowerHeader === 'sales executive') {
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
            roleLabel = 'Visit Source';
          } else if (lowerHeader.includes('channel partner')) {
            roleLabel = 'Channel Partner';
          } else if (lowerHeader.includes('call status')) {
            roleLabel = 'Call Status';
          } else if (lowerHeader.includes('configuration')) {
            roleLabel = 'Configuration';
          } else if (lowerHeader.includes('sourcing manager')) {
            roleLabel = 'Sourcing Manager';
          } else if (lowerHeader.includes('source type')) {
            roleLabel = 'Source Type';
          } else if (lowerHeader.includes('source executive')) {
            roleLabel = 'Source Executive';
          } else if (lowerHeader.includes('enquiry type')) {
            roleLabel = 'Enquiry Type';
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

              // Ensure all canonical buckets appear in standard order
              const lowerCanonical = DEFAULT_BUCKET_LIST.map(b => b.toLowerCase());
              DEFAULT_BUCKET_LIST.forEach(cb => {
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
        if (suggestedBucketColIdx === -1 && detectedColumns.length > 1) {
          const cand = detectedColumns.find(col => col.role === 'bucket' && col.colIndex !== suggestedSalesColIdx) || detectedColumns[1];
          suggestedBucketColIdx = cand.colIndex;
        }

        // Sort detected columns so that Sales, Bucket, Project, Source appear first
        detectedColumns.sort((a, b) => {
          const order = { sales: 1, bucket: 2, project: 3, source: 4, other: 5 };
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
        const defaultTitle = `Site Visits Report | ${detectedProject} | Week Report (${startDisp} to ${endDisp})`;

        resolve({
          headerIndex,
          detectedColumns,
          suggestedSalesColIdx,
          suggestedBucketColIdx,
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
    selectedSales: string[];
    selectedBuckets: string[];
    columnFilters: Record<number, string[]>;
    reportTitle: string;
    headerIndex: number;
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

        const { salesColIdx, bucketColIdx, selectedSales, selectedBuckets, columnFilters, reportTitle, headerIndex } = options;

        const dataRows = rawRows.slice(headerIndex + 1);

        // Map of salesUser -> bucket -> count
        const userBucketMap: Record<string, Record<string, number>> = {};
        
        // Initialize all selected sales users with 0 for all selected buckets
        selectedSales.forEach(user => {
          userBucketMap[user] = {};
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
              let val = getCellValue(row[cIdx]).trim();
              if (cIdx === bucketColIdx && (!val || val === '-' || val.toLowerCase() === 'null' || val.toLowerCase() === 'undefined')) {
                val = 'Open';
              }
              if (val && !allowed.includes(val)) {
                passesFilters = false;
                break;
              }
            }
          }
          if (!passesFilters) continue;

          const rawUser = getCellValue(row[salesColIdx]).trim();
          if (!rawUser || rawUser === '-') continue;

          // Find matching sales user
          const matchedUser = selectedSales.find(u => u.toLowerCase() === rawUser.toLowerCase());
          if (!matchedUser) continue;

          let rawBucket = getCellValue(row[bucketColIdx]).trim();
          // Requirement: "consider blank Enquiry Level as Open for the report consideration"
          if (!rawBucket || rawBucket === '-' || rawBucket.toLowerCase() === 'null' || rawBucket.toLowerCase() === 'undefined') {
            rawBucket = 'Open';
          }

          // Match bucket case-insensitively or normalized
          const matchedBucket = selectedBuckets.find(b => b.toLowerCase() === rawBucket.toLowerCase());
          if (matchedBucket) {
            if (!userBucketMap[matchedUser]) {
              userBucketMap[matchedUser] = {};
            }
            userBucketMap[matchedUser][matchedBucket] = (userBucketMap[matchedUser][matchedBucket] || 0) + 1;
          }
        }

        // Build rows
        const rows: BucketRowData[] = selectedSales.map(user => {
          const counts = userBucketMap[user] || {};
          let grandTotal = 0;
          selectedBuckets.forEach(b => {
            grandTotal += (counts[b] || 0);
          });
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
          columns: ['Sales', 'Grand Total', ...selectedBuckets],
          buckets: selectedBuckets,
          rows,
          columnTotals: {
            grandTotal: overallGrandTotal,
            bucketTotals
          }
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
  const totalCols = 2 + buckets.length;

  const headerCellsHtml = buckets.map(b => `
    <th style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; font-weight: bold; text-align: center; color: #000000; background-color: #ffffff; white-space: nowrap; line-height: 1.2;">${b}</th>
  `).join('');

  const rowCellsHtml = rows.map(r => `
    <tr>
      <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: left; font-weight: normal; color: #000000; white-space: nowrap;">${r.salesUser}</td>
      <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: normal; color: #000000;">${r.grandTotal}</td>
      ${buckets.map(b => `
        <td style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; text-align: center; font-weight: normal; color: #000000;">${r.bucketCounts[b] || 0}</td>
      `).join('')}
    </tr>
  `).join('');

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
          <th style="border: 1px solid #000000; padding: 6px 12px; font-size: 13.5px; font-weight: bold; text-align: left; color: #000000; background-color: #ffffff; min-width: 140px;">Sales</th>
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
  const wb = utils.book_new();

  // Excel AOA
  const aoa: any[][] = [];

  // Row 1: Merged Title
  aoa.push([reportTitle]);

  // Row 2: Headers
  aoa.push(['Sales', 'Grand Total', ...buckets]);

  // Row 3+: Data rows
  const firstDataRow = 3; // 1-indexed in Excel: Row 3
  const lastDataRow = 2 + rows.length;

  rows.forEach((r, idx) => {
    const excelRowNum = firstDataRow + idx;
    // Grand Total formula: SUM of all bucket columns from C to last col
    const lastColLetter = utils.encode_col(1 + buckets.length);
    const formula = `SUM(C${excelRowNum}:${lastColLetter}${excelRowNum})`;

    const rowArray: any[] = [
      r.salesUser,
      { f: formula, v: r.grandTotal, t: 'n' },
      ...buckets.map(b => r.bucketCounts[b] || 0)
    ];
    aoa.push(rowArray);
  });

  // Footer Row: Grand Total
  const totalRowNum = lastDataRow + 1;
  const grandTotalColLetter = 'B';
  const grandTotalFormula = `SUM(B${firstDataRow}:B${lastDataRow})`;

  const footerRow: any[] = [
    'Grand Total',
    { f: grandTotalFormula, v: columnTotals.grandTotal, t: 'n' }
  ];

  buckets.forEach((b, bIdx) => {
    const colLetter = utils.encode_col(2 + bIdx);
    const bucketFormula = `SUM(${colLetter}${firstDataRow}:${colLetter}${lastDataRow})`;
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
    { wch: 22 }, // Sales
    { wch: 14 }, // Grand Total
    ...buckets.map(b => ({ wch: Math.max(b.length + 3, 12) }))
  ];

  utils.book_append_sheet(wb, ws, 'Site Visits Report');
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
 * Main processor function for Bucket Report - Site Visit
 */
export async function processBucketSiteVisitFile(
  file: File,
  options: {
    salesColIdx: number;
    bucketColIdx: number;
    selectedSales: string[];
    selectedBuckets: string[];
    columnFilters: Record<number, string[]>;
    reportTitle: string;
    headerIndex: number;
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
