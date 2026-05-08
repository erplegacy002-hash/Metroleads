import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING } from './projectMapping';

// --- Configuration ---

/**
 * Categorizes a call status from the 'Disposition' column. 
 * If it matches known success keywords, it's Answered.
 * Otherwise, it's counted as Missed (ensuring total count is accurate).
 */
// Function isAnswered removed to support strict "Answered" vs "Missed" counting logic.

/**
 * Handles durations from Excel "Editing Mode" (raw values).
 */

function getCellValue(cell: any): string {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'object' && cell.v !== undefined) return String(cell.v);
  return String(cell);
}

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
 * Parses and formats a date value (string, number, or Date object)
 * Returns string in DD MMM YYYY format (e.g., 26 DEC 2025)
 */
function parseAndFormatDate(val: any): string | null {
  if (!val) return null;
  let date: Date | undefined;

  if (val instanceof Date) {
    date = val;
  } else if (typeof val === 'number') {
    // Excel serial date (approximate)
    // 25569 is offset between 1900-01-01 and 1970-01-01
    date = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else if (typeof val === 'string') {
    const v = val.trim();
    
    // Check for DD/MM/YY or DD/MM/YYYY format specifically (e.g. 25/12/25)
    // Regex matches starts with DD/MM/YY or DD/MM/YYYY, optionally followed by time
    const dmyMatch = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    
    if (dmyMatch) {
      const day = parseInt(dmyMatch[1], 10);
      const month = parseInt(dmyMatch[2], 10) - 1; // JS months are 0-indexed
      let year = parseInt(dmyMatch[3], 10);
      
      // Handle 2 digit years (assume 2000s)
      if (year < 100) year += 2000;
      
      date = new Date(year, month, day);
    } else {
      // Try standard parsing
      const d = new Date(v);
      if (!isNaN(d.getTime())) {
        date = d;
      }
    }
  }

  if (date && !isNaN(date.getTime())) {
    return date.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    }).toUpperCase();
  }
  return null;
}

/**
 * Formats seconds into a string using 'hour logic'.
 */
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

/**
 * Formats seconds into "X hrs Y mins" format.
 */
function formatHrsMins(totalSeconds: number): string {
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  return `${h} hrs ${m} mins`;
}

// --- Image Generation ---

async function generateTableImage(siteName: string, rows: any[], displayDate: string): Promise<string> {
  const container = document.createElement('div');
  // Widen the container slightly to accommodate the new column
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '850px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: "'Calibri', sans-serif",
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
    return `<th style="padding: 6px 4px; text-align: ${alignment}; border: 1px solid #000000; font-size: 11px; font-weight: 900; font-family: 'Arial', sans-serif; color: #000000; background-color: #f3f4f6; vertical-align: bottom; line-height: 1.1;">
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
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000; font-weight: 400;">${row['Total Call Duration']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000; font-weight: 600;">${row['Total Count']}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row['Average Call']}</td>
    </tr>
  `).join('');

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box; font-family: 'Calibri', sans-serif;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 900; font-family: 'Arial', sans-serif; color: #000000; text-transform: uppercase;">CALLING REPORT</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 18px; font-weight: 900; font-family: 'Arial', sans-serif; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; font-family: 'Arial', sans-serif; color: #000000;">${displayDate}</div>
      </div>
      <div style="height: 10px;"></div>
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
        // CellNF: true and CellFormula: true enable 'editing mode' values
        const workbook = read(data, { type: 'array', cellDates: true, cellNF: true, cellFormula: true });
        const sheetName = workbook.SheetNames.find(n => n.toLowerCase() === 'call logs');
        const sheet = sheetName ? workbook.Sheets[sheetName] : null;
        if (!sheet) throw new Error("Sheet 'Call logs' not found in the Excel file.");
        const rawRows = utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][];
        
        if (!rawRows || rawRows.length === 0) throw new Error("Excel file is empty.");
        for(let R = 0; R < rawRows.length; R++) {
            const row = rawRows[R];
            if (!row || !Array.isArray(row)) continue;
            const range = utils.decode_range(sheet['!ref'] || 'A1');
            const maxC = Math.max(row.length, range.e.c + 1);
            for(let C = 0; C < maxC; C++) {
                 const cell_ref = utils.encode_cell({ c: C, r: R });
                 const originalCell = sheet[cell_ref];
                 if (originalCell && originalCell.l && originalCell.l.Target) {
                     row[C] = { v: row[C] !== null ? row[C] : '', t: originalCell.t || 's', l: originalCell.l };
                 }
            }
        }


        // Detect Columns
        let headerIndex = -1;
        let userColIdx = -1, dispositionColIdx = -1, durationColIdx = -1, dateColIdx = -1;

        // "Terminated on" is the prioritized target for identifying the agent/user.
        const userAliases = ['terminated on', 'assigned to', 'agent', 'user', 'employee'];
        const durationAliases = ['duration', 'talk time', 'bill sec', 'call duration', 'time'];
        const dispAliases = ['disposition', 'status', 'call result', 'result'];
        const dateAliases = ['call date', 'date', 'call_date', 'calldate'];

        for (let i = 0; i < Math.min(50, rawRows.length); i++) {
          const row = rawRows[i];
          if (!Array.isArray(row)) continue;
          
          // Identify the 'Terminated on' column specifically
          const uIdx = row.findIndex(c => {
            if (!c) return false;
            const cleanCell = String(c).toLowerCase().trim();
            // Explicit check for 'Terminated on' first
            return cleanCell === 'terminated on' || cleanCell.includes('terminated on');
          });

          // Identify Disposition
          const dIdx = row.findIndex(c => {
            if (!c) return false;
            const cleanCell = String(c).toLowerCase().trim();
            return cleanCell === 'disposition' || cleanCell.includes('disposition');
          });
          
          if (uIdx !== -1 && dIdx !== -1) {
            headerIndex = i;
            userColIdx = uIdx;
            dispositionColIdx = dIdx;
            durationColIdx = row.findIndex(c => c && durationAliases.some(a => String(c).toLowerCase().includes(a)));
            dateColIdx = row.findIndex(c => c && dateAliases.some(a => String(c).toLowerCase().trim() === a || String(c).toLowerCase().includes(a)));
            break;
          }
        }

        // Fallback if strict 'Terminated on' wasn't found - search more broadly
        if (headerIndex === -1) {
           for (let i = 0; i < Math.min(50, rawRows.length); i++) {
            const row = rawRows[i];
            if (!Array.isArray(row)) continue;
            const uIdx = row.findIndex(c => c && userAliases.some(a => String(c).toLowerCase().trim() === a || String(c).toLowerCase().includes(a)));
            const dIdx = row.findIndex(c => c && dispAliases.some(a => String(c).toLowerCase().includes(a)));
            if (uIdx !== -1 && dIdx !== -1) {
              headerIndex = i;
              userColIdx = uIdx;
              dispositionColIdx = dIdx;
              durationColIdx = row.findIndex(c => c && durationAliases.some(a => String(c).toLowerCase().includes(a)));
              dateColIdx = row.findIndex(c => c && dateAliases.some(a => String(c).toLowerCase().trim() === a || String(c).toLowerCase().includes(a)));
              break;
            }
          }
        }

        if (headerIndex === -1) throw new Error("Could not find 'Terminated on' or 'Disposition' columns.");

        const jsonData = utils.sheet_to_json(sheet, { range: headerIndex, raw: true, defval: "" });
        const headerRow = rawRows[headerIndex];
        const userKey = String(headerRow[userColIdx]);
        const dispositionKey = String(headerRow[dispositionColIdx]);
        const durationKey = durationColIdx !== -1 ? String(headerRow[durationColIdx]) : null;
        const dateKey = dateColIdx !== -1 ? String(headerRow[dateColIdx]) : null;

        const normalizedMapping: Record<string, string> = {};
        Object.keys(USER_PROJECT_MAPPING).forEach(k => {
          normalizedMapping[k.toLowerCase().trim()] = USER_PROJECT_MAPPING[k];
        });

        const sites: Record<string, Record<string, any>> = {};
        let detectedDate: string | null = null;

        jsonData.forEach((row: any) => {
          // Try to grab date from the first available row in "Call Date" column
          // We only take the first valid date we find
          if (!detectedDate && dateKey && row[dateKey]) {
            detectedDate = parseAndFormatDate(row[dateKey]);
          }

          const rawVal = row[userKey];
          if (rawVal === undefined || rawVal === null) return;
          const userStr = String(rawVal).trim();
          if (!userStr) return;
          
          // Fuzzy match against our known user list
          const userLower = userStr.toLowerCase();
          const matchedUserKey = Object.keys(normalizedMapping).find(k => {
             return userLower === k || userLower.includes(k) || k.includes(userLower);
          });

          if (!matchedUserKey) return;

          const siteName = normalizedMapping[matchedUserKey];
          const disposition = row[dispositionKey];
          const durationRaw = durationKey ? parseDurationRaw(row[durationKey]) : 0;

          if (!sites[siteName]) sites[siteName] = {};
          if (!sites[siteName][matchedUserKey]) {
            const originalName = Object.keys(USER_PROJECT_MAPPING).find(k => k.toLowerCase() === matchedUserKey) || matchedUserKey;
            sites[siteName][matchedUserKey] = { 
              displayName: originalName,
              answered: 0, missed: 0, durAns: 0, durMiss: 0 
            };
          }

          const stats = sites[siteName][matchedUserKey];
          
          const disp = String(disposition).trim().toLowerCase();
          
          if (disp === 'answered') {
            stats.answered++;
            stats.durAns += durationRaw;
          } else if (disp === 'missed') {
            stats.missed++;
            stats.durMiss += durationRaw;
          }
        });

        // Use ONLY detected date from "Call Date" column.
        // If not found, use a placeholder or empty string to indicate missing data.
        const reportDate = detectedDate || "DATE MISSING"; 

        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        const siteKeys = Object.keys(sites);
        if (siteKeys.length === 0) throw new Error("No matching users found in the 'Terminated on' column. Please check if the names match the project mapping.");

        for (const site of siteKeys) {
          const userStats = sites[site];
          const rows = Object.values(userStats).map((s: any) => {
            const totalSec = s.durAns + s.durMiss;
            const avgSec = s.answered > 0 ? (s.durAns / s.answered) : 0;
            const totalCount = s.answered + s.missed;

            return {
              "User Name": `User - ${s.displayName}`,
              "Answered": s.answered,
              "Call Duration (Answered)": formatWithHourLogic(s.durAns),
              "Missed": s.missed,
              "Call Duration (Missed)": formatWithHourLogic(s.durMiss),
              "Total Call Duration": formatWithHourLogic(totalSec),
              "Total Count": totalCount,
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
