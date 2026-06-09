import { read, utils, write } from 'xlsx';
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, DEFAULT_SITE } from './projectMapping';

// --- Helpers ---

function getCellValue(cell: any): string {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'object' && cell.v !== undefined) return String(cell.v);
  return String(cell);
}

function parseDate(val: any): Date | null {
  if (typeof val === 'object' && val !== null && !(val instanceof Date) && val.v !== undefined) val = val.v;
  if (!val) return null;
  let date: Date | undefined;

  if (val instanceof Date) {
    // Extract year, month, date components in UTC/Naive format to prevent local/UTC timezone misalignment shifts
    const y = val.getUTCFullYear();
    const m = val.getUTCMonth();
    const d = val.getUTCDate();
    date = new Date(y, m, d);
  } else if (typeof val === 'number') {
    date = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else if (typeof val === 'string') {
    const v = val.trim();
    if (v.match(/^\d{4}-\d{2}-\d{2}$/)) {
       const [y, m, d] = v.split('-').map(Number);
       date = new Date(y, m - 1, d);
    } else {
        const dmyMatch = v.match(/^(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{2,4})/);
        if (dmyMatch) {
            const part1 = parseInt(dmyMatch[1], 10);
            const part2 = parseInt(dmyMatch[2], 10);
            let year = parseInt(dmyMatch[3], 10);
            if (year < 100) year += 2000;

            let day = part1;
            let month = part2;
            if (part1 > 12) {
                day = part1;
                month = part2;
            } else if (part2 > 12) {
                day = part2;
                month = part1;
            }
            date = new Date(year, month - 1, day);
        } else {
             const dMmmYMatch = v.match(/^(\d{1,2})[\/\-\.\s]([A-Za-z]{3,12})[\/\-\.\s](\d{2,4})/);
             if (dMmmYMatch) {
                 const day = parseInt(dMmmYMatch[1], 10);
                 const monthStr = dMmmYMatch[2].toLowerCase().substring(0, 3);
                 const yearStr = dMmmYMatch[3];
                 let year = parseInt(yearStr, 10);
                 if (year < 100) year += 2000;
                 
                 const months: {[key:string]: number} = {jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11};
                 if (months[monthStr] !== undefined) {
                     date = new Date(year, months[monthStr], day);
                 }
             } else {
                const d = new Date(v);
                if (!isNaN(d.getTime())) date = d;
             }
        }
    }
  }

  if (date && !isNaN(date.getTime())) {
    date.setHours(0, 0, 0, 0); 
    return date;
  }
  return null;
}

function formatDate(date: Date): string {
  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  }).toUpperCase();
}

function findColumnIndex(row: any[], aliases: string[]): number {
  if (!row || !Array.isArray(row)) return -1;
  const normalizedRow = row.map(cell => {
    if (cell === null || cell === undefined) return '';
    return String(cell).toLowerCase().replace(/[\s_\-\(\)]+/g, '');
  });

  const cleanAliases = aliases.map(a => a.toLowerCase().replace(/[\s_\-\(\)]+/g, ''));

  // 1. Exact clean match
  for (const cleanAlias of cleanAliases) {
    const idx = normalizedRow.indexOf(cleanAlias);
    if (idx !== -1) return idx;
  }

  // 2. Fuzzy includes match (prioritize longer aliases first to avoid greedy matching)
  const sortedAliases = [...cleanAliases].sort((a, b) => b.length - a.length);
  for (const cleanAlias of sortedAliases) {
    if (cleanAlias.length < 3) continue; // skip very short words
    for (let i = 0; i < normalizedRow.length; i++) {
      if (normalizedRow[i].includes(cleanAlias)) {
        return i;
      }
    }
  }

  return -1;
}

// User performance aggregates container
interface UserPerfStats {
  user: string;
  siteVisitDone: number;
  revisitDone: number;
  booked: number;
  total: number;
  conversionRatio: number; // percentage
  avgLeadAge: number; // in days
  leadAgeSum: number; // for calculation
  leadAgeCount: number; // for grand total calculation
}

// --- PDF Generator ---

async function generateUserPerformancePDF(
  projectName: string,
  statsList: UserPerfStats[],
  startDate: string,
  endDate: string
): Promise<string> {
  const doc = new jsPDF({ orientation: 'landscape', format: 'a4' });

  // Custom styling to match user performance report image
  const reportTitle = `${projectName === 'Overall' ? 'User Performance Report' : `User Performance Report (Project: ${projectName})`}`;
  
  // Table head
  const headRow = ['User', 'Site Visit Done', 'Re-Visit Done', 'Booked', 'Total', 'Conversion Ratio', 'Avg. Lead Age'];

  // Calculations for totals row
  let grandSVD = 0;
  let grandRVD = 0;
  let grandBooked = 0;
  let grandTotal = 0;
  let grandLeadAgeSum = 0;
  let grandLeadAgeCount = 0;

  statsList.forEach(stat => {
    grandSVD += stat.siteVisitDone;
    grandRVD += stat.revisitDone;
    grandBooked += stat.booked;
    grandTotal += stat.total;
    grandLeadAgeSum += stat.leadAgeSum;
    grandLeadAgeCount += stat.leadAgeCount;
  });

  const grandConvRatio = grandTotal > 0 ? Number(((grandBooked / grandTotal) * 100).toFixed(2)) : 0;
  const grandAvgLeadAge = grandLeadAgeCount > 0 ? Math.round(grandLeadAgeSum / grandLeadAgeCount) : 0;

  const tableBody = statsList.map(stat => [
    stat.user,
    stat.siteVisitDone,
    stat.revisitDone,
    stat.booked,
    stat.total,
    `${stat.conversionRatio}%`,
    `${stat.avgLeadAge} days`
  ]);

  // Add highly visible Total row
  tableBody.push([
    'Total',
    grandSVD,
    grandRVD,
    grandBooked,
    grandTotal,
    `${grandConvRatio}%`,
    `${grandAvgLeadAge} days`
  ]);

  // Large outer table border styling container to match the image precisely
  autoTable(doc, {
    startY: 40,
    head: [headRow],
    body: tableBody,
    theme: 'grid',
    styles: { 
      fontSize: 10, 
      cellPadding: 6,
      lineColor: [0, 0, 0],
      lineWidth: 0.2,
      textColor: [0, 0, 0],
      font: 'helvetica'
    },
    headStyles: {
      fillColor: [245, 245, 245],
      textColor: [0, 0, 0],
      fontStyle: 'bold',
      halign: 'center',
      valign: 'middle'
    },
    columnStyles: {
      0: { halign: 'left', fontStyle: 'normal' },
      1: { halign: 'center' },
      2: { halign: 'center' },
      3: { halign: 'center' },
      4: { halign: 'center' },
      5: { halign: 'center', fontStyle: 'bold' },
      6: { halign: 'center' }
    },
    willDrawCell: function(data) {
      // Bold the last row (Total)
      if (data.row.index === tableBody.length - 1) {
        doc.setFont("helvetica", "bold");
      }
    },
    didDrawPage: function (data) {
      // Top Report Box to match the screenshot layout
      const pageSize = doc.internal.pageSize;
      const pageWidth = pageSize.width ? pageSize.width : pageSize.getWidth();
      const pageHeight = pageSize.height ? pageSize.height : pageSize.getHeight();

      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.5);
      
      // Draw a neat bounding box around content at the top
      doc.rect(14, 10, pageWidth - 28, 25);
      
      // Title
      doc.setFontSize(14);
      doc.setFont("helvetica", "bold");
      doc.text(reportTitle.toUpperCase(), pageWidth / 2, 20, { align: "center" });

      // Dates centered or split inside the bounding box
      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");
      doc.text(`Start Date : ${startDate}`, 20, 30);
      doc.text(`End Date : ${endDate}`, pageWidth - 20, 30, { align: "right" });

      // Footer
      doc.setFontSize(8);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(120, 120, 120);
      doc.text(`Page ${data.pageNumber}`, pageWidth - 14, pageHeight - 10, { align: "right" });
    },
    margin: { top: 45, left: 14, right: 14 }
  });

  return doc.output('datauristring');
}

// --- Helper to read and process raw excel array buffers asynchronously ---
async function parseExcelFile(file: File): Promise<any[][] | null> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'array', cellDates: true, cellNF: true, cellFormula: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][];

        if (!rawRows || rawRows.length === 0) {
          resolve(null);
          return;
        }

        // Clean cell values / hyperlinks
        for (let R = 0; R < rawRows.length; R++) {
            const row = rawRows[R];
            if (!row || !Array.isArray(row)) continue;
            const range = utils.decode_range(sheet['!ref'] || 'A1');
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
                         if (m) {
                             val = m[1];
                         }
                     }
                     row[C] = val;
                 }
            }
        }
        resolve(rawRows);
      } catch (err) {
        console.error("Error parsing file", file.name, err);
        resolve(null);
      }
    };
    reader.onerror = () => resolve(null);
    reader.readAsArrayBuffer(file);
  });
}

interface SheetInfo {
  headerIndex: number;
  assignedToIdx: number;
  projectIdx: number;
  createdDateIdx: number;
  visitDateIdx: number;
  sourceIdx: number;
}

interface UserStatsAccumulator {
  siteVisitDone: number;
  revisitDone: number;
  booked: number;
  leadAgeSum: number;
  leadAgeCount: number;
}

const assignedAliases = ['sales executive', 'assigned to', 'assigned_to', 'owner', 'agent', 'executive', 'allocated to', 'sales person', 'sourcing manager', 'closing manager', 'user', 'employee'];
const bookedAgentAliases = ['executive name', 'executive_name', 'executive', 'sales executive', ...assignedAliases];
const projectAliases = ['project name', 'project', 'site name', 'project (af)', 'project(af)', 'project (af', 'project(af', 'site', 'project_name'];
const createdDateAliases = ['created at', 'created on', 'created time', 'created date', 'lead date', 'created_on', 'created_time', 'date of lead', 'entry date'];
const dateAliases = ['site visit date', 'visit date 1', 'visit_date_1', 'visit date', 'visit_date', 'date of visit', 'visited date', 'revisit date 1', 'revisit_date_1', 'revisit date', 'booking date', 'booked date', 'date', 'entry date'];
const sourceAliases = ['visit source', 'lead source', 'lead source (f)', 'source', 'source of lead', 'enquiry source'];

function analyzeSheet(rawRows: any[][], isBookedSheet: boolean): SheetInfo {
  let headerIndex = -1;
  let assignedToIdx = -1;
  let projectIdx = -1;
  let createdDateIdx = -1;
  let visitDateIdx = -1;
  let sourceIdx = -1;

  const userAliases = isBookedSheet ? bookedAgentAliases : assignedAliases;

  for (let i = 0; i < Math.min(100, rawRows.length); i++) {
    const row = rawRows[i];
    if (!Array.isArray(row)) continue;

    const aIdx = findColumnIndex(row, userAliases);

    if (aIdx !== -1) {
      headerIndex = i;
      assignedToIdx = aIdx;
      projectIdx = findColumnIndex(row, projectAliases);
      createdDateIdx = findColumnIndex(row, createdDateAliases);
      visitDateIdx = findColumnIndex(row, dateAliases);
      sourceIdx = findColumnIndex(row, sourceAliases);
      break;
    }
  }

  return { headerIndex, assignedToIdx, projectIdx, createdDateIdx, visitDateIdx, sourceIdx };
}

// --- Main File Processor ---

export async function processUserPerformanceFile(
  files: File[] | File,
  manualStartDate?: string,
  manualEndDate?: string,
  sourceFilter: string = 'All'
): Promise<ProcessResponse> {
  const filesArray = Array.isArray(files) ? files : [files];
  
  let visitFile: File | null = null;
  let revisitFile: File | null = null;
  let bookedFile: File | null = null;

  filesArray.forEach(f => {
    const name = f.name.toLowerCase();
    if (name.includes('revisit') || name.includes('re-visit')) {
      revisitFile = f;
    } else if (name.includes('booked') || name.includes('booking')) {
      bookedFile = f;
    } else if (name.includes('visit') || name.includes('site_visit') || name.includes('site-visit')) {
      visitFile = f;
    }
  });

  // Fallback match by index if names don't clearly state the roles
  const unmatched = filesArray.filter(f => f !== visitFile && f !== revisitFile && f !== bookedFile);
  if (unmatched.length > 0) {
    if (!visitFile) visitFile = unmatched.shift() || null;
    if (!revisitFile) revisitFile = unmatched.shift() || null;
    if (!bookedFile) bookedFile = unmatched.shift() || null;
  }

  const startFilter = manualStartDate ? parseDate(manualStartDate) : null;
  const endFilter = manualEndDate ? parseDate(manualEndDate) : null;

  let globalMinDate: Date | null = null;
  let globalMaxDate: Date | null = null;

  const accumulatedData: Record<string, Record<string, UserStatsAccumulator>> = {};
  const userNormalizationMap: Record<string, string> = {}; // lowercase -> Title Case representation

  function normalizeUser(u: string): string {
    const trimmed = u.trim().replace(/\s+/g, ' ');
    const lower = trimmed.toLowerCase();
    if (!userNormalizationMap[lower]) {
      // Format to clean Title Case: "deepesh khare" -> "Deepesh Khare"
      const capitalized = trimmed.split(' ')
        .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
        .join(' ');
      userNormalizationMap[lower] = capitalized;
    }
    return userNormalizationMap[lower];
  }

  function normalizeProject(p: string): string {
    const val = p.toLowerCase().trim();
    if (val.includes('milestone')) return 'Milestone';
    if (val.includes('kairos')) return 'Kairos';
    if (val.includes('aqua') || val.includes('aqualife')) return 'Aqua Life';
    if (val.includes('statement')) return 'Statement';
    if (val.includes('ekam')) return 'Legacy Ekam';
    return p.trim();
  }

  function getOrCreateUserStats(project: string, user: string): UserStatsAccumulator {
    if (!accumulatedData[project]) {
      accumulatedData[project] = {};
    }
    if (!accumulatedData[project][user]) {
      accumulatedData[project][user] = {
        siteVisitDone: 0,
        revisitDone: 0,
        booked: 0,
        leadAgeSum: 0,
        leadAgeCount: 0
      };
    }
    return accumulatedData[project][user];
  }

  const processRows = async (file: File, type: 'visit' | 'revisit' | 'booked') => {
    const rawRows = await parseExcelFile(file);
    if (!rawRows || rawRows.length === 0) return;

    const isBooked = type === 'booked';
    const info = analyzeSheet(rawRows, isBooked);
    if (info.headerIndex === -1) return;

    for (let i = info.headerIndex + 1; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (!row || row.length === 0) continue;

      const isExcluded = row.some(cell => {
         if (!cell) return false;
         const s = String(cell).toLowerCase().trim();
         return s === 'test' || s.includes('ramesh bodke');
      });
      if (isExcluded) continue;

      const rawAssigned = info.assignedToIdx !== -1 ? row[info.assignedToIdx] : '';
      const rawUser = rawAssigned ? String(rawAssigned).trim() : '';
      if (!rawUser || rawUser === '-' || rawUser.toLowerCase() === 'unassigned' || rawUser.toLowerCase() === 'total' || rawUser.toLowerCase() === 'grand total' || rawUser.toLowerCase() === 'sum') continue;

      const user = normalizeUser(rawUser);

      // Extract and resolve Project Name safely from Project column, fallback to USER_PROJECT_MAPPING
      let rawProject = info.projectIdx !== -1 ? String(row[info.projectIdx]).trim() : '';
      if (rawProject === '-' || rawProject.toLowerCase() === 'unassigned' || !rawProject) {
        rawProject = '';
      }

      let project = 'Unspecified Project';
      if (rawProject) {
        project = normalizeProject(rawProject);
      } else {
        const lowerUser = user.toLowerCase();
        let foundMapping = false;
        for (const [k, v] of Object.entries(USER_PROJECT_MAPPING)) {
          const lowerKey = k.toLowerCase();
          if (lowerUser === lowerKey || lowerUser.includes(lowerKey) || lowerKey.includes(lowerUser)) {
            project = v;
            foundMapping = true;
            break;
          }
        }
        if (!foundMapping) {
          project = DEFAULT_SITE;
        }
      }

      if (sourceFilter !== 'All' && info.sourceIdx !== -1) {
         const rSource = getCellValue(row[info.sourceIdx]).toLowerCase();
         if (!rSource.includes(sourceFilter.toLowerCase())) {
           continue;
         }
      }

      let activeDate = info.visitDateIdx !== -1 ? parseDate(row[info.visitDateIdx]) : null;
      if (!activeDate && info.createdDateIdx !== -1) {
         activeDate = parseDate(row[info.createdDateIdx]);
      }

      if (activeDate) {
        if (startFilter && activeDate < startFilter) continue;
        if (endFilter && activeDate > endFilter) continue;

        if (!globalMinDate || activeDate < globalMinDate) globalMinDate = activeDate;
        if (!globalMaxDate || activeDate > globalMaxDate) globalMaxDate = activeDate;
      }

      const stats = getOrCreateUserStats(project, user);
      if (type === 'visit') {
        stats.siteVisitDone++;
        const createdDate = info.createdDateIdx !== -1 ? parseDate(row[info.createdDateIdx]) : null;
        if (createdDate) {
          const today = new Date();
          today.setHours(0, 0, 0, 0);
          const diff = today.getTime() - createdDate.getTime();
          const age = Math.max(0, Math.floor(diff / (1000 * 60 * 60 * 24)));
          stats.leadAgeSum += age;
          stats.leadAgeCount++;
        }
      } else if (type === 'revisit') {
        stats.revisitDone++;
      } else if (type === 'booked') {
        stats.booked++;
      }
    }
  };

  if (visitFile) await processRows(visitFile, 'visit');
  if (revisitFile) await processRows(revisitFile, 'revisit');
  if (bookedFile) await processRows(bookedFile, 'booked');

  const projectsList = Object.keys(accumulatedData).sort();
  if (projectsList.length === 0) {
    throw new Error("No record found matching the selected dates or source filter. Check if your spreadsheet filenames contain 'visit', 'revisit', or 'booked'.");
  }

  const finalGlobalStart = manualStartDate || (globalMinDate ? formatDate(globalMinDate) : "-");
  const finalGlobalEnd = manualEndDate || (globalMaxDate ? formatDate(globalMaxDate) : "-");

  const getProjectStatsList = (proj: string): UserPerfStats[] => {
    const usersMap = accumulatedData[proj];
    const userKeys = Object.keys(usersMap);
    const list = userKeys.map(user => {
      const acc = usersMap[user];
      const total = acc.siteVisitDone + acc.revisitDone;
      const convRatio = total > 0 ? Number(((acc.booked / total) * 100).toFixed(2)) : 0;
      const avgAge = acc.leadAgeCount > 0 ? Math.round(acc.leadAgeSum / acc.leadAgeCount) : 0;

      return {
        user,
        siteVisitDone: acc.siteVisitDone,
        revisitDone: acc.revisitDone,
        booked: acc.booked,
        total,
        conversionRatio: convRatio,
        avgLeadAge: avgAge,
        leadAgeSum: acc.leadAgeSum,
        leadAgeCount: acc.leadAgeCount
      };
    });

    list.sort((a, b) => {
      if (b.conversionRatio !== a.conversionRatio) {
        return b.conversionRatio - a.conversionRatio;
      }
      if (b.total !== a.total) {
        return b.total - a.total;
      }
      return a.user.localeCompare(b.user);
    });

    return list;
  };

  // Build overall cumulative stats summed across all projects
  const overallStatsMap: Record<string, UserStatsAccumulator> = {};
  for (const proj of Object.keys(accumulatedData)) {
    const projUsers = accumulatedData[proj];
    for (const user of Object.keys(projUsers)) {
      if (!overallStatsMap[user]) {
        overallStatsMap[user] = {
          siteVisitDone: 0,
          revisitDone: 0,
          booked: 0,
          leadAgeSum: 0,
          leadAgeCount: 0
        };
      }
      overallStatsMap[user].siteVisitDone += projUsers[user].siteVisitDone;
      overallStatsMap[user].revisitDone += projUsers[user].revisitDone;
      overallStatsMap[user].booked += projUsers[user].booked;
      overallStatsMap[user].leadAgeSum += projUsers[user].leadAgeSum;
      overallStatsMap[user].leadAgeCount += projUsers[user].leadAgeCount;
    }
  }

  const overallStatsList: UserPerfStats[] = Object.keys(overallStatsMap).map(user => {
    const acc = overallStatsMap[user];
    const total = acc.siteVisitDone + acc.revisitDone;
    const convRatio = total > 0 ? Number(((acc.booked / total) * 100).toFixed(2)) : 0;
    const avgAge = acc.leadAgeCount > 0 ? Math.round(acc.leadAgeSum / acc.leadAgeCount) : 0;
    return {
      user,
      siteVisitDone: acc.siteVisitDone,
      revisitDone: acc.revisitDone,
      booked: acc.booked,
      total,
      conversionRatio: convRatio,
      avgLeadAge: avgAge,
      leadAgeSum: acc.leadAgeSum,
      leadAgeCount: acc.leadAgeCount
    };
  });

  overallStatsList.sort((a, b) => {
    if (b.conversionRatio !== a.conversionRatio) {
      return b.conversionRatio - a.conversionRatio;
    }
    if (b.total !== a.total) {
      return b.total - a.total;
    }
    return a.user.localeCompare(b.user);
  });

  const zip = new JSZip();
  const images: GeneratedImage[] = [];
  const sheetsData: Record<string, any[][]> = {};

  const formatExcelAOA = (title: string, stats: UserPerfStats[]) => {
    const rowsList: any[][] = [
      [title.toUpperCase()],
      [`Start Date : ${finalGlobalStart}`, "", "", "", "", "", `End Date : ${finalGlobalEnd}`],
      [],
      ['User', 'Site Visit Done', 'Re-Visit Done', 'Booked', 'Total', 'Conversion Ratio', 'Avg. Lead Age']
    ];

    let totSVD = 0;
    let totRVD = 0;
    let totB = 0;
    let totT = 0;
    let totAgeSum = 0;
    let totAgeCount = 0;

    const firstDataRow = 5;
    const lastDataRow = 4 + stats.length;
    const totalRow = 5 + stats.length;

    stats.forEach((s, idx) => {
      totSVD += s.siteVisitDone;
      totRVD += s.revisitDone;
      totB += s.booked;
      totT += s.total;
      totAgeSum += s.leadAgeSum;
      totAgeCount += s.leadAgeCount;
      
      const rNum = firstDataRow + idx;
      
      rowsList.push([
        s.user,
        s.siteVisitDone,
        s.revisitDone,
        s.booked,
        { f: "B" + rNum + "+C" + rNum, v: s.total, t: 'n' },
        { f: "IF(E" + rNum + ">0, ROUND((D" + rNum + "/E" + rNum + ")*100, 2), 0) & \"%\"", v: s.conversionRatio + "%", t: 's' },
        { v: s.avgLeadAge, t: 'n', z: '0" days"' }
      ]);
    });

    const totConv = totT > 0 ? Number(((totB / totT) * 100).toFixed(2)) : 0;
    const totAvgAge = totAgeCount > 0 ? Math.round(totAgeSum / totAgeCount) : 0;

    rowsList.push([
      'Total',
      { f: "SUM(B5:B" + lastDataRow + ")", v: totSVD, t: 'n' },
      { f: "SUM(C5:C" + lastDataRow + ")", v: totRVD, t: 'n' },
      { f: "SUM(D5:D" + lastDataRow + ")", v: totB, t: 'n' },
      { f: "SUM(E5:E" + lastDataRow + ")", v: totT, t: 'n' },
      { f: "IF(E" + totalRow + ">0, ROUND((D" + totalRow + "/E" + totalRow + ")*100, 2), 0) & \"%\"", v: totConv + "%", t: 's' },
      { f: "ROUND(IF(COUNT(G5:G" + lastDataRow + ")>0, AVERAGE(G5:G" + lastDataRow + "), 0), 0) & \" days\"", v: totAvgAge + " days", t: 's' }
    ]);

    // Footnotes and logic explanations
    rowsList.push([]);
    rowsList.push(["LOGICAL FORMULAS & CALCULATION DEFINITIONS:"]);
    rowsList.push(["1. Conversion Ratio (%)", "Formula: (Booked / Total) * 100", "Where: Total = Site Visit Done + Re-Visit Done"]);
    rowsList.push(["2. Avg. Lead Age (days)", "Formula: Average of (Today's Date - Lead Created At Date)", "Where: Today's date refers to the current system date"]);
    rowsList.push(["3. Totals", "Site Visit Done: Excel SUM() of column B for all user rows"]);
    rowsList.push(["", "Re-Visit Done: Excel SUM() of column C for all user rows"]);
    rowsList.push(["", "Booked: Excel SUM() of column D for all user rows"]);
    rowsList.push(["", "Total Actions: Excel SUM() of column E for all user rows"]);
    rowsList.push(["", "Grand Conversion Ratio: calculated as Booked Total / Actions Total * 100 using Excel IF() & ROUND() formulas"]);
    rowsList.push(["", "Grand Average Lead Age: calculated as AVERAGE() of column G for all user rows using Excel ROUND() & AVERAGE() formulas"]);

    return rowsList;
  };

  // Compile sheets and PDFs. Always generate 'Overall Summary' first, followed by project breakdowns
  const allProjectsToGenerate = ['Overall', ...projectsList];

  for (const proj of allProjectsToGenerate) {
    const pStatsList = proj === 'Overall' ? overallStatsList : getProjectStatsList(proj);

    const pPdfDataUrl = await generateUserPerformancePDF(
      proj,
      pStatsList,
      finalGlobalStart,
      finalGlobalEnd
    );

    const safeProjFilename = proj.replace(/[^a-z0-9]/gi, '_').toLowerCase();

    images.push({
      project_name: proj === 'Overall' ? 'Overall Summary' : `Project: ${proj}`,
      image_url: pPdfDataUrl,
      filename: `${safeProjFilename}_user_performance.pdf`
    });
    zip.file(`${safeProjFilename}_user_performance.pdf`, pPdfDataUrl.split(',')[1], { base64: true });

    sheetsData[proj.substring(0, 30)] = formatExcelAOA(`User Performance - ${proj}`, pStatsList);
  }

  const wb = utils.book_new();
  Object.keys(sheetsData).forEach(sheetName => {
    const ws = utils.aoa_to_sheet(sheetsData[sheetName]);
    utils.book_append_sheet(wb, ws, sheetName);
  });
  const excelBuffer = write(wb, { bookType: 'xlsx', type: 'array' });
  zip.file('user_performance_report.xlsx', excelBuffer);

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  return {
    images,
    zip_url: URL.createObjectURL(zipBlob),
    message: "Success"
  };
}

