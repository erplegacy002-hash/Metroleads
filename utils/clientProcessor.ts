import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING } from './projectMapping';

// --- Configuration ---

/**
 * Categorizes a call status. 
 * If it matches known success keywords, it's Answered.
 * Otherwise, it's counted as Missed (ensuring total count is accurate).
 */
function isAnswered(status: any): boolean {
  if (status === null || status === undefined) return false;
  const s = String(status).toLowerCase().trim();
  
  // Explicit "Answered" markers
  const answeredKeywords = [
    'connected', 'talked', 'answer', 'converted', 'sale', 
    'interested', 'visit', 'meeting', 'demo', 'deal', 'follow'
  ];

  const matchesAnswered = answeredKeywords.some(kw => s.includes(kw));
  
  // Ensure we don't treat "Not Answered" as Answered
  if (matchesAnswered) {
    const isFalsePositive = s.includes('not ') || s.includes('no ');
    if (isFalsePositive && !s.includes('connected')) return false;
    return true;
  }

  return false;
}

/**
 * Handles durations from Excel "Editing Mode" (raw values).
 */
function parseDurationRaw(val: any): number {
  if (val === null || val === undefined) return 0;
  
  if (typeof val === 'number') {
    // Excel Time Serial check (fraction of 24h)
    if (val > 0 && val < 1) {
      return Math.round(val * 86400);
    }
    return val;
  }
  
  if (typeof val === 'string') {
    const v = val.trim();
    if (v.includes(':')) {
      const parts = v.split(':').map(Number);
      if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
      if (parts.length === 2) return parts[0] * 60 + parts[1];
    }
    const parsed = parseFloat(v);
    return isNaN(parsed) ? 0 : parsed;
  }
  return 0;
}

/**
 * Extracts a date from filename in format YYYY-MM-DD
 */
function extractDateFromFilename(filename: string): string {
  const datePattern = /(\d{4})-(\d{2})-(\d{2})/;
  const match = filename.match(datePattern);
  
  if (match) {
    const [_, year, month, day] = match;
    const dateObj = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    return dateObj.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    }).toUpperCase();
  }

  return new Date().toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  }).toUpperCase();
}

function formatWithHourLogic(totalSeconds: number): string {
  if (totalSeconds >= 3600) {
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    return `${h}:${m.toString().padStart(2, '0')} hrs`;
  } else {
    const m = Math.floor(totalSeconds / 60);
    const s = Math.floor(totalSeconds % 60);
    return `${m}:${s.toString().padStart(2, '0')} mins`;
  }
}

// --- Image Generation ---

async function generateTableImage(siteName: string, rows: any[], displayDate: string): Promise<string> {
  const container = document.createElement('div');
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '780px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: 'sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });
  
  if (rows.length === 0) return '';

  const displayHeaders = [
    "User Name", "Answered", "Call Duration (Answered)", 
    "Missed", "Call Duration (Missed)", "Total Call Duration", 
    "Total Count", "Average Call"
  ];

  const headerHtml = displayHeaders.map(h => {
    let formattedHeader = h;
    if (h.includes('(') && h.includes(')')) {
      formattedHeader = h.replace(' (', '<br/>(');
    } else if (h === "Total Call Duration") {
      formattedHeader = "Total Call<br/>Duration";
    }
    const alignment = h === "User Name" ? 'left' : 'center';
    return `<th style="padding: 6px 4px; text-align: ${alignment}; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; vertical-align: bottom; line-height: 1.1;">
      ${formattedHeader}
    </th>`;
  }).join('');

  const rowsHtml = rows.map((row) => `
    <tr>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: left; color: #000000; font-weight: 500;">${row['User Name']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Answered']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Call Duration (Answered)']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Missed']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Call Duration (Missed)']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Total Call Duration']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Total Count']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Average Call']}</td>
    </tr>
  `).join('');

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center; border-bottom: 1px solid #000000;">
        <div style="font-size: 14px; font-weight: 800; color: #000000; text-transform: uppercase;">CALLING REPORT</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 18px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${displayDate}</div>
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

// --- Main Processor ---

export async function processFile(file: File): Promise<ProcessResponse> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'array', cellDates: true, cellNF: true, cellFormula: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][];
        
        if (!rawRows || rawRows.length === 0) throw new Error("Excel file is empty.");

        // Detect Columns
        let headerIndex = -1;
        let userColIdx = -1, dispositionColIdx = -1, durationColIdx = -1;

        // "Terminated on" is prioritized for the username as requested.
        // We REMOVE 'project name' from userAliases to ensure it doesn't pick that up instead.
        const userAliases = ['terminated on', 'assigned to', 'user', 'agent', 'employee'];
        const durationAliases = ['duration', 'talk time', 'bill sec', 'call duration', 'time'];
        const dispAliases = ['disposition', 'status', 'call result', 'result'];

        for (let i = 0; i < Math.min(50, rawRows.length); i++) {
          const row = rawRows[i];
          if (!Array.isArray(row)) continue;
          
          // Find User Column (Check prioritized list)
          const uIdx = row.findIndex(c => {
            if (!c) return false;
            const cleanCell = String(c).toLowerCase().trim();
            return userAliases.some(alias => cleanCell === alias || cleanCell.includes(alias));
          });

          const dIdx = row.findIndex(c => c && dispAliases.some(a => String(c).toLowerCase().includes(a)));
          
          if (uIdx !== -1 && dIdx !== -1) {
            headerIndex = i;
            userColIdx = uIdx;
            dispositionColIdx = dIdx;
            durationColIdx = row.findIndex(c => c && durationAliases.some(a => String(c).toLowerCase().includes(a)));
            break;
          }
        }

        if (headerIndex === -1) throw new Error("Required columns (Terminated on/User and Disposition) not found.");

        const jsonData = utils.sheet_to_json(sheet, { range: headerIndex, raw: true, defval: "" });
        const headerRow = rawRows[headerIndex];
        const userKey = String(headerRow[userColIdx]);
        const dispositionKey = String(headerRow[dispositionColIdx]);
        const durationKey = durationColIdx !== -1 ? String(headerRow[durationColIdx]) : null;

        const normalizedMapping: Record<string, string> = {};
        Object.keys(USER_PROJECT_MAPPING).forEach(k => {
          normalizedMapping[k.toLowerCase().trim()] = USER_PROJECT_MAPPING[k];
        });

        const sites: Record<string, Record<string, any>> = {};

        jsonData.forEach((row: any) => {
          const rawVal = row[userKey];
          if (!rawVal) return;
          const userStr = String(rawVal).trim();
          
          // Fuzzy match against our known user list
          const userLower = userStr.toLowerCase();
          const matchedUserKey = Object.keys(normalizedMapping).find(k => {
             // Exact match or contains (to handle 'Agent - Name' or variations)
             return userLower === k || userLower.includes(k) || k.includes(userLower);
          });

          if (!matchedUserKey) return;

          const siteName = normalizedMapping[matchedUserKey];
          const disposition = row[dispositionKey];
          const durationRaw = durationKey ? parseDurationRaw(row[durationKey]) : 0;

          if (!sites[siteName]) sites[siteName] = {};
          if (!sites[siteName][matchedUserKey]) {
            // Find original casing for display
            const originalName = Object.keys(USER_PROJECT_MAPPING).find(k => k.toLowerCase() === matchedUserKey) || matchedUserKey;
            sites[siteName][matchedUserKey] = { 
              displayName: originalName,
              answered: 0, missed: 0, durAns: 0, durMiss: 0 
            };
          }

          const stats = sites[siteName][matchedUserKey];
          if (isAnswered(disposition)) {
            stats.answered++;
            stats.durAns += durationRaw;
          } else {
            stats.missed++;
            stats.durMiss += durationRaw;
          }
        });

        const reportDate = extractDateFromFilename(file.name);
        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        const siteKeys = Object.keys(sites);
        if (siteKeys.length === 0) throw new Error("No matching users from the mapping found in the data. Please check the 'Terminated on' column content.");

        for (const site of siteKeys) {
          const userStats = sites[site];
          const rows = Object.values(userStats).map((s: any) => {
            const totalSec = s.durAns + s.durMiss;
            const avgSec = s.answered > 0 ? (s.durAns / s.answered) : 0;
            return {
              "User Name": `User - ${s.displayName}`,
              "Answered": s.answered,
              "Call Duration (Answered)": formatWithHourLogic(s.durAns),
              "Missed": s.missed,
              "Call Duration (Missed)": formatWithHourLogic(s.durMiss),
              "Total Call Duration": formatWithHourLogic(totalSec),
              "Total Count": s.answered + s.missed,
              "Average Call": `${Math.floor(avgSec / 60)}:${Math.floor(avgSec % 60).toString().padStart(2, '0')}`,
              "rawTotal": totalSec
            };
          }).sort((a, b) => b.rawTotal - a.rawTotal);

          const dataUrl = await generateTableImage(site, rows, reportDate);
          const filename = `${site.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_summary.png`;
          images.push({ project_name: site, image_url: dataUrl, filename });
          zip.file(filename, dataUrl.split(',')[1], { base64: true });
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
