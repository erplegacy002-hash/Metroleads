import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING, DEFAULT_SITE } from './projectMapping';

// --- Helpers ---

function parseDate(val: any): Date | null {
  if (!val) return null;
  let date: Date | undefined;

  if (val instanceof Date) {
    date = val;
  } else if (typeof val === 'number') {
    date = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else if (typeof val === 'string') {
    const v = val.trim();
    if (v.match(/^\d{4}-\d{2}-\d{2}$/)) {
       const [y, m, d] = v.split('-').map(Number);
       date = new Date(y, m - 1, d);
    } else {
        const dmyMatch = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
        if (dmyMatch) {
            const day = parseInt(dmyMatch[1], 10);
            const month = parseInt(dmyMatch[2], 10) - 1;
            let year = parseInt(dmyMatch[3], 10);
            if (year < 100) year += 2000;
            date = new Date(year, month, day);
        } else {
             // Try DD-MMM-YYYY (e.g. 12-Jan-2024)
             const dMmmYMatch = v.match(/^(\d{1,2})[\/\-\s]([A-Za-z]{3})[\/\-\s](\d{2,4})/);
             if (dMmmYMatch) {
                 const day = parseInt(dMmmYMatch[1], 10);
                 const monthStr = dMmmYMatch[2].toLowerCase();
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

function determineSource(cpData: any, sourceData: any, subSourceData: any): string {
  if (cpData && String(cpData).trim().length > 0) {
    return 'Channel Partner';
  }

  const rawSource = sourceData ? String(sourceData).trim() : '';
  const rawSubSource = subSourceData ? String(subSourceData).trim() : '';

  const checkKeywords = (str: string): string | null => {
    const s = str.toLowerCase();
    if (s.includes('channel partner')) return 'Channel Partner';
    if (s.includes('walk-in') || s.includes('walkin')) return 'Walk-In';
    if (s.includes('digital') || s.includes('facebook') || s.includes('instagram') || s.includes('website') || s.includes('google') || s.includes('whatsapp') || s.includes('popup')) return 'Digital';
    if (s.includes('offer')) return 'Offer';
    if (s.includes('refer') || s.includes('referral') || s.includes('reference')) return 'Referral';
    if (s.includes('hoarding') || s.includes('hoardings') || s.includes('incoming call - rahatne') || s.includes('site branding')) return 'Hoarding';
    return null;
  };

  // Check Lead Source first
  if (rawSource.length > 0) {
    const match = checkKeywords(rawSource);
    if (match) return match;
  }

  // Check Sub Source second
  if (rawSubSource.length > 0) {
    const match = checkKeywords(rawSubSource);
    if (match) return match;
  }

  // Fallback
  if (rawSource.length > 0) return rawSource;
  return '-';
}

// --- Image Generation: Main List ---

async function generateWeeklyListImage(siteName: string, rows: any[], reportTitle: string, startDate: string, endDate: string): Promise<string> {
  const container = document.createElement('div');
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '1250px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: 'sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });
  
  if (rows.length === 0) return '';

  const headerHtml = `
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 40px;">Sr. No.</th>
    <th style="padding: 6px 4px; text-align: left; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Visitor Name</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 80px;">Team</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 100px;">Source</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 140px;">CP Firm Name</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 85px;">Visit Date</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 85px;">2nd Visit</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 85px;">3rd Visit</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 85px;">4th Visit</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 85px;">5th Visit</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 110px;">State</th>
  `;

  const rowsHtml = rows.map((row, index) => `
    <tr>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${index + 1}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: left; color: #000000; font-weight: 500;">${row.name}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.team}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.source}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.cpFirmName}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.date}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.date2}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.date3}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.date4}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.date5}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.state}</td>
    </tr>
  `).join('');

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 800; color: #000000; text-transform: uppercase;">SITE VISIT</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 18px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${reportTitle}</div>
      </div>

      <div style="padding: 5px 2px; display: flex; justify-content: space-between; font-size: 11px; font-weight: 700; color: #000000;">
        <span>Start Date: ${startDate}</span>
        <span>End Date: ${endDate}</span>
      </div>

      <table style="width: 100%; border-collapse: collapse; background-color: #ffffff;">
        <thead><tr>${headerHtml}</tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  `;

  document.body.appendChild(container);
  await new Promise(resolve => setTimeout(resolve, 600));

  try {
    return await toPng(container, { quality: 0.95, pixelRatio: 2 });
  } finally {
    if (document.body.contains(container)) document.body.removeChild(container);
  }
}

// --- Image Generation: Summary ---

interface TeamCounts {
  presales: number;
  salesGre: number;
}

async function generateWeeklySummaryImage(
  siteName: string, 
  rows: any[], 
  summaryStats: Record<string, TeamCounts>, 
  sourceStats: Record<string, TeamCounts>, 
  reportTitle: string, 
  startDate: string, 
  endDate: string
): Promise<string> {
  const container = document.createElement('div');
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '450px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: 'sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });

  // Calculate Metrics for Header Box
  const totalRows = rows.length; // Unique entries in the report
  const countDate2 = rows.filter(r => r.date2 && r.date2 !== '-').length;
  const countDate3 = rows.filter(r => r.date3 && r.date3 !== '-').length;
  const countDate4 = rows.filter(r => r.date4 && r.date4 !== '-').length;
  const countDate5 = rows.filter(r => r.date5 && r.date5 !== '-').length;
  const totalRevisits = countDate2 + countDate3 + countDate4 + countDate5; // Sum of 2nd, 3rd, 4th, 5th visits
  const totalVisits = totalRows + totalRevisits; // Total Footfall
  const totalBookings = (summaryStats['Booked']?.presales || 0) + (summaryStats['Booked']?.salesGre || 0);

  // Table Footer Calculation
  const totalStatusPresales = Object.values(summaryStats).reduce((a, b) => a + b.presales, 0);
  const totalStatusSalesGre = Object.values(summaryStats).reduce((a, b) => a + b.salesGre, 0);
  
  // Revisit should NOT be included in "Total Visits" count for sources table footer if it's derived from existing visits
  const totalSourcePresales = Object.entries(sourceStats)
    .filter(([key]) => key !== 'Revisit')
    .reduce((a, [_, b]) => a + b.presales, 0);
    
  const totalSourceSalesGre = Object.entries(sourceStats)
    .filter(([key]) => key !== 'Revisit')
    .reduce((a, [_, b]) => a + b.salesGre, 0);

  // Define display order for States
  const mandatoryStates = ["Visit Scheduled", "Revisit Done", "Booked"];
  
  // Include ALL states present in stats, do not exclude "New Lead", "Interested", etc.
  const allStateKeys = Array.from(new Set([...mandatoryStates, ...Object.keys(summaryStats)]));
  
  allStateKeys.sort((a, b) => {
    const idxA = mandatoryStates.indexOf(a);
    const idxB = mandatoryStates.indexOf(b);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.localeCompare(b);
  });

  const summaryRowsHtml = allStateKeys.map(state => {
    const counts = summaryStats[state] || { presales: 0, salesGre: 0 };
    // Filter out rows with 0 in both columns
    if (counts.presales === 0 && counts.salesGre === 0) return '';
    return `
    <tr>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: left; color: #000000;">${state}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000;">${counts.presales}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000;">${counts.salesGre}</td>
    </tr>`;
  }).join('');

  // Define display order for Sources
  const mandatorySources = ["Digital", "Channel Partner", "Referral", "Offer", "Walk-In", "Hoarding"];
  const otherSources = Object.keys(sourceStats).filter(k => !mandatorySources.includes(k) && k !== 'Revisit');
  const finalSourceOrder = [...mandatorySources, ...otherSources, "Revisit"];

  const sourceRowsHtml = finalSourceOrder.map(source => {
    const counts = sourceStats[source] || { presales: 0, salesGre: 0 };
    // Filter out rows with 0 in both columns
    if (counts.presales === 0 && counts.salesGre === 0) return '';
    return `
    <tr>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: left; color: #000000;">${source}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000;">${counts.presales}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000;">${counts.salesGre}</td>
    </tr>`;
  }).join('');

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 800; color: #000000; text-transform: uppercase;">SUMMARY REPORT</div>
        <div style="width: 100px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 16px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 100px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${reportTitle}</div>
      </div>

      <div style="padding: 5px 2px; display: flex; justify-content: space-between; font-size: 11px; font-weight: 700; color: #000000; padding-left: 15px; padding-right: 15px;">
        <span>Start Date: ${startDate}</span>
        <span>End Date: ${endDate}</span>
      </div>

      <!-- Summary Metrics Box -->
      <div style="margin: 15px 15px 20px 15px; border: 1px solid #000000; padding: 12px 0; display: flex; justify-content: space-between;">
          <div style="flex: 1; text-align: center; border-right: 1px solid #e5e7eb; display: flex; flex-direction: column; align-items: center; justify-content: flex-start;">
              <div style="font-size: 10px; font-weight: 800; color: #4b5563; text-transform: uppercase; margin-bottom: 5px; line-height: 1.3; min-height: 26px; display: flex; align-items: flex-end;">Total No. of<br>Visits</div>
              <div style="font-size: 18px; font-weight: 900; color: #000000;">${totalVisits}</div>
          </div>
          <div style="flex: 1; text-align: center; border-right: 1px solid #e5e7eb; display: flex; flex-direction: column; align-items: center; justify-content: flex-start;">
              <div style="font-size: 10px; font-weight: 800; color: #4b5563; text-transform: uppercase; margin-bottom: 5px; line-height: 1.3; min-height: 26px; display: flex; align-items: flex-end;">Total No. of<br>Revisits</div>
              <div style="font-size: 18px; font-weight: 900; color: #000000;">${totalRevisits}</div>
          </div>
          <div style="flex: 1; text-align: center; display: flex; flex-direction: column; align-items: center; justify-content: flex-start;">
              <div style="font-size: 10px; font-weight: 800; color: #4b5563; text-transform: uppercase; margin-bottom: 5px; line-height: 1.3; min-height: 26px; display: flex; align-items: flex-end;">Bookings</div>
              <div style="font-size: 18px; font-weight: 900; color: #000000;">${totalBookings}</div>
          </div>
      </div>

      <div style="padding: 0 15px 15px 15px;">
        <!-- Lead Status Summary -->
        <div style="font-size: 12px; font-weight: 800; color: #000000; text-transform: uppercase; margin-bottom: 6px;">LEAD STATUS SUMMARY</div>
        <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; margin-bottom: 20px;">
          <thead>
            <tr>
              <th style="padding: 8px 12px; text-align: left; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">State</th>
              <th style="padding: 8px 12px; text-align: center; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Presales</th>
              <th style="padding: 8px 12px; text-align: center; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Sales + GRE</th>
            </tr>
          </thead>
          <tbody>
            ${summaryRowsHtml}
            <tr>
               <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: right; font-weight: 700; color: #000000; background-color: #f9fafb;">Total</td>
               <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000; background-color: #f9fafb;">${totalStatusPresales}</td>
               <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000; background-color: #f9fafb;">${totalStatusSalesGre}</td>
            </tr>
          </tbody>
        </table>

        <!-- Source Summary -->
        <div style="font-size: 12px; font-weight: 800; color: #000000; text-transform: uppercase; margin-bottom: 6px;">SOURCE SUMMARY</div>
        <table style="width: 100%; border-collapse: collapse; background-color: #ffffff;">
          <thead>
            <tr>
              <th style="padding: 8px 12px; text-align: left; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Source</th>
              <th style="padding: 8px 12px; text-align: center; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Presales</th>
              <th style="padding: 8px 12px; text-align: center; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Sales + GRE</th>
            </tr>
          </thead>
          <tbody>
            ${sourceRowsHtml}
            <tr>
               <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: right; font-weight: 700; color: #000000; background-color: #f9fafb;">Total Visits</td>
               <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000; background-color: #f9fafb;">${totalSourcePresales}</td>
               <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000; background-color: #f9fafb;">${totalSourceSalesGre}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `;

  document.body.appendChild(container);
  await new Promise(resolve => setTimeout(resolve, 600));

  try {
    return await toPng(container, { quality: 0.95, pixelRatio: 2 });
  } finally {
    if (document.body.contains(container)) document.body.removeChild(container);
  }
}

// --- Main Processor ---

function findColumnIndex(row: any[], aliases: string[]): number {
  for (const alias of aliases) {
    const idx = row.findIndex(c => c && String(c).toLowerCase().trim() === alias);
    if (idx !== -1) return idx;
  }
  return -1;
}

export async function processWeeklySiteVisitFile(file: File, manualStartDate?: string, manualEndDate?: string): Promise<ProcessResponse> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'array', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][];

        if (!rawRows || rawRows.length === 0) throw new Error("Excel file is empty.");

        // Detect Columns
        let headerIndex = -1;
        let nameIdx = -1, stateIdx = -1, assignedToIdx = -1;
        let visitDateIdx = -1, visitDate2Idx = -1, visitDate3Idx = -1, visitDate4Idx = -1, visitDate5Idx = -1;
        let cpFirmNameIdx = -1, leadSourceIdx = -1, subSourceIdx = -1;
        let projectIdx = -1;

        const nameAliases = ['name', 'visitor name', 'lead name', 'customer name', 'full name', 'client name'];
        const stateAliases = ['lead state', 'state', 'region', 'location'];
        const assignedAliases = ['assigned to', 'assigned_to', 'owner', 'agent', 'executive', 'sales executive', 'allocated to', 'sales person', 'sourcing manager', 'closing manager'];
        
        const dateAliases = ['visit date', 'visit_date', 'date of visit', 'date', 'visited date', 'created time', 'created on', 'entry date'];
        const date2Aliases = ['2nd visit date', 'second visit date', 'visit date 2', '2nd_visit_date'];
        const date3Aliases = ['3rd visit date', 'third visit date', 'visit date 3', '3rd_visit_date'];
        const date4Aliases = ['4th site visit date', '4th site visit date (az)', '4th visit date', 'fourth visit date', 'visit date 4', '4th_visit_date', '4th site visit', '4th visit', 'visit 4'];
        const date5Aliases = ['5th site visit date', '5th visit date', 'fifth visit date', 'visit date 5', '5th_visit_date', '5th site visit', '5th visit', 'visit 5'];
        
        const cpFirmAliases = ['cp firm name', 'cp firm name (v)', 'cp name', 'channel partner firm name'];
        const leadSourceAliases = ['lead source', 'lead source (f)', 'source', 'source of lead', 'enquiry source'];
        const subSourceAliases = ['sub source', 'sub source (u)', 'sub_source', 'subsource'];
        
        const projectAliases = ['project', 'project name', 'project (af)', 'project(af)', 'project (af', 'project(af'];

        for (let i = 0; i < Math.min(100, rawRows.length); i++) {
          const row = rawRows[i];
          if (!Array.isArray(row)) continue;

          const nIdx = findColumnIndex(row, nameAliases);
          const sIdx = findColumnIndex(row, stateAliases);
          const aIdx = findColumnIndex(row, assignedAliases);
          
          const dIdx = findColumnIndex(row, dateAliases);
          const d2Idx = findColumnIndex(row, date2Aliases);
          const d3Idx = findColumnIndex(row, date3Aliases);
          const d4Idx = findColumnIndex(row, date4Aliases);
          const d5Idx = findColumnIndex(row, date5Aliases);
          
          const cpIdx = findColumnIndex(row, cpFirmAliases);
          const lsIdx = findColumnIndex(row, leadSourceAliases);
          const ssIdx = findColumnIndex(row, subSourceAliases);
          
          const pIdx = findColumnIndex(row, projectAliases);

          if (nIdx !== -1 && aIdx !== -1) {
            headerIndex = i;
            nameIdx = nIdx;
            stateIdx = sIdx;
            assignedToIdx = aIdx;
            visitDateIdx = dIdx;
            visitDate2Idx = d2Idx;
            visitDate3Idx = d3Idx;
            visitDate4Idx = d4Idx;
            visitDate5Idx = d5Idx;
            cpFirmNameIdx = cpIdx;
            leadSourceIdx = lsIdx;
            subSourceIdx = ssIdx;
            projectIdx = pIdx;
            break;
          }
        }

        if (headerIndex === -1) throw new Error("Could not find required columns (Name, Assigned To, etc).");

        const startFilter = manualStartDate ? parseDate(manualStartDate) : null;
        const endFilter = manualEndDate ? parseDate(manualEndDate) : null;

        const normalizedMapping: Record<string, string> = {};
        Object.keys(USER_PROJECT_MAPPING).forEach(k => {
          normalizedMapping[k.toLowerCase().trim()] = USER_PROJECT_MAPPING[k];
        });

        const normalizedTeamMapping: Record<string, string> = {};
        Object.keys(USER_TEAM_MAPPING).forEach(k => {
          normalizedTeamMapping[k.toLowerCase().trim()] = USER_TEAM_MAPPING[k];
        });

        const sites: Record<string, any[]> = {};
        
        for (let i = headerIndex + 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0) continue;

          // GLOBAL EXCLUSION CHECK
          // Exclude rows if any cell contains "metro", "test", or "ramesh bodke"
          const isExcluded = row.some(cell => {
             if (!cell) return false;
             const s = String(cell).toLowerCase();
             return s.includes('metro') || s.includes('test') || s.includes('ramesh bodke');
          });
          if (isExcluded) continue;

          const rawAssigned = row[assignedToIdx];
          const assignedStr = rawAssigned ? String(rawAssigned).trim() : "Unassigned";
          const assignedLower = assignedStr.toLowerCase();

          // Initialize siteName and team
          let siteName = DEFAULT_SITE;
          let team = '-';

          // Special logic for specific users: Manisha Singh, Smita Kad, Sejal Satav
          const specificUsers = ["manisha singh", "smita kad", "sejal satav"];
          const isSpecificUser = specificUsers.some(u => assignedLower.includes(u));

          if (isSpecificUser) {
             // Look into 'Project' column for keywords
             const rawProject = projectIdx !== -1 ? row[projectIdx] : '';
             const projectVal = String(rawProject).toLowerCase().trim();
             
             if (projectVal.includes('kairos')) siteName = 'Kairos';
             else if (projectVal.includes('aqua') || projectVal.includes('aqualife')) siteName = 'Aqua Life';
             else if (projectVal.includes('milestone')) siteName = 'Milestone';
             else if (projectVal.includes('statement')) siteName = 'Statement';
             
             // Try to assign team based on user mapping if available, otherwise default
             let matchedUserKey = Object.keys(normalizedTeamMapping).find(k => assignedLower.includes(k));
             if (matchedUserKey) {
                team = normalizedTeamMapping[matchedUserKey];
             }
          } else {
              // Standard Fuzzy match / Check mapping
              let matchedUserKey = Object.keys(normalizedMapping).find(k => {
                return assignedLower === k || assignedLower.includes(k) || k.includes(assignedLower);
              });
              
              if (matchedUserKey) {
                siteName = normalizedMapping[matchedUserKey];
                 // Try to find team using matched key
                if (normalizedTeamMapping[matchedUserKey]) {
                    team = normalizedTeamMapping[matchedUserKey];
                }
              }
          }

          const name = row[nameIdx] ? String(row[nameIdx]).trim() : '-';
          const nameLower = name.toLowerCase();
          // Filter if name contains 'test'
          if (nameLower.includes('test')) continue;

          let state = (stateIdx !== -1 && row[stateIdx]) ? String(row[stateIdx]).trim() : '-';
          if (state.toLowerCase() === 're_visit_done') state = 'Revisit Done';
          // REMOVED BOOKED FILTER: Now including Booked records

           // --- Date Logic ---
          const d1 = visitDateIdx !== -1 ? parseDate(row[visitDateIdx]) : null;
          const d2 = visitDate2Idx !== -1 ? parseDate(row[visitDate2Idx]) : null;
          const d3 = visitDate3Idx !== -1 ? parseDate(row[visitDate3Idx]) : null;
          const d4 = visitDate4Idx !== -1 ? parseDate(row[visitDate4Idx]) : null;
          const d5 = visitDate5Idx !== -1 ? parseDate(row[visitDate5Idx]) : null;

          let selectedDate: Date | null = null;

          if (startFilter && endFilter) {
              const datesToCheck = [d1, d2, d3, d4, d5].filter(d => d !== null) as Date[];
              const datesInRange = datesToCheck.filter(d => d >= startFilter && d <= endFilter);
              
              if (datesInRange.length === 0) continue; 
              
              datesInRange.sort((a,b) => b.getTime() - a.getTime());
              selectedDate = datesInRange[0];
          } else {
              const datesToCheck = [d1, d2, d3, d4, d5].filter(d => d !== null) as Date[];
              if (datesToCheck.length > 0) {
                   datesToCheck.sort((a,b) => b.getTime() - a.getTime());
                   selectedDate = datesToCheck[0];
              }
          }

          const cpData = cpFirmNameIdx !== -1 ? row[cpFirmNameIdx] : null;
          const leadSourceData = leadSourceIdx !== -1 ? row[leadSourceIdx] : null;
          const subSourceData = subSourceIdx !== -1 ? row[subSourceIdx] : null;

          const source = determineSource(cpData, leadSourceData, subSourceData);
          const cpFirmName = cpData ? String(cpData).trim() : '-';

          if (!sites[siteName]) sites[siteName] = [];
          
          sites[siteName].push({
            name,
            state,
            team,
            source,
            cpFirmName,
            date: d1 ? formatDate(d1) : '-',
            date2: d2 ? formatDate(d2) : '-',
            date3: d3 ? formatDate(d3) : '-',
            date4: d4 ? formatDate(d4) : '-',
            date5: d5 ? formatDate(d5) : '-',
            rawDateVal: selectedDate,
            sortDate: d1 // Store 1st visit date for sorting
          });
        }

        const images: GeneratedImage[] = [];
        const zip = new JSZip();
        const siteKeys = Object.keys(sites);

        if (siteKeys.length === 0) throw new Error("No matching records found based on criteria.");

        const manualStartFormatted = startFilter ? formatDate(startFilter) : null;
        const manualEndFormatted = endFilter ? formatDate(endFilter) : null;

        let dateLabel = "WEEKLY REPORT";

        for (const site of siteKeys) {
          const rows = sites[site];
          
          // Calculate Date Range for Header (Sort rawDateVal to get min/max independent of row order)
          const validRawDates = rows.map(r => r.rawDateVal).filter(d => d) as Date[];
          validRawDates.sort((a, b) => a.getTime() - b.getTime());
          const startDateVal = validRawDates.length > 0 ? validRawDates[0] : null;
          const endDateVal = validRawDates.length > 0 ? validRawDates[validRawDates.length - 1] : null;

          // Sort Rows by Visit Date (d1)
          rows.sort((a, b) => {
            const da = a.sortDate;
            const db = b.sortDate;
            if (!da && !db) return 0;
            if (!da) return 1;
            if (!db) return -1;
            return da > db ? 1 : -1;
          });
          
          const autoStartDateStr = startDateVal ? formatDate(startDateVal) : "-";
          const autoEndDateStr = endDateVal ? formatDate(endDateVal) : "-";
          
          const finalStartDateStr = manualStartFormatted || autoStartDateStr;
          const finalEndDateStr = manualEndFormatted || autoEndDateStr;

          const summaryStats: Record<string, TeamCounts> = {};
          const sourceStats: Record<string, TeamCounts> = {};

          const incrementStats = (stats: Record<string, TeamCounts>, key: string, isPresales: boolean) => {
            if (!stats[key]) stats[key] = { presales: 0, salesGre: 0 };
            if (isPresales) {
                stats[key].presales++;
            } else {
                stats[key].salesGre++;
            }
          };
          
          rows.forEach(r => {
             const isPresales = r.team === 'Presales';
             incrementStats(summaryStats, r.state, isPresales);
             incrementStats(sourceStats, r.source, isPresales);

             // Revisit Logic: Check 5th, 4th, 3rd, 2nd
             let isRevisit = false;
             if (r.date5 && r.date5 !== '-') {
                 isRevisit = true;
             } else if (r.date4 && r.date4 !== '-') {
                 isRevisit = true;
             } else if (r.date3 && r.date3 !== '-') {
                 isRevisit = true;
             } else if (r.date2 && r.date2 !== '-') {
                 isRevisit = true;
             }
             if (isRevisit) {
                 incrementStats(sourceStats, 'Revisit', isPresales);
             }
          });

          const listDataUrl = await generateWeeklyListImage(site, rows, dateLabel, finalStartDateStr, finalEndDateStr);
          const listFilename = `${site.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_weekly_visit.png`;
          images.push({ project_name: site, image_url: listDataUrl, filename: listFilename });
          zip.file(listFilename, listDataUrl.split(',')[1], { base64: true });

          const summaryDataUrl = await generateWeeklySummaryImage(site, rows, summaryStats, sourceStats, dateLabel, finalStartDateStr, finalEndDateStr);
          const summaryFilename = `${site.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_weekly_summary.png`;
          images.push({ project_name: `Summary - ${site}`, image_url: summaryDataUrl, filename: summaryFilename });
          zip.file(summaryFilename, summaryDataUrl.split(',')[1], { base64: true });
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        resolve({ images, zip_url: URL.createObjectURL(zipBlob), message: "Success" });

      } catch (err: any) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}