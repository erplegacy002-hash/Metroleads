import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING } from './projectMapping';

// --- Helpers ---

function parseAndFormatDate(val: any): string | null {
  if (!val) return null;
  let date: Date | undefined;

  if (val instanceof Date) {
    date = val;
  } else if (typeof val === 'number') {
    // Excel serial date
    date = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else if (typeof val === 'string') {
    const v = val.trim();
    // Try DD/MM/YYYY
    const dmyMatch = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dmyMatch) {
      const day = parseInt(dmyMatch[1], 10);
      const month = parseInt(dmyMatch[2], 10) - 1;
      let year = parseInt(dmyMatch[3], 10);
      if (year < 100) year += 2000;
      date = new Date(year, month, day);
    } else {
      const d = new Date(v);
      if (!isNaN(d.getTime())) date = d;
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

// --- Image Generation: Daily List (No Date Column) ---

async function generateDailyListImage(siteName: string, rows: any[], dateLabel: string): Promise<string> {
  const container = document.createElement('div');
  
  // Reduced width from 850px to 600px for a more compact look
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '600px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: 'sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });
  
  if (rows.length === 0) return '';

  const headerHtml = `
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 50px;">Sr. No.</th>
    <th style="padding: 6px 4px; text-align: left; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Visitor Name</th>
    <th style="padding: 6px 4px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 150px;">State</th>
  `;

  const rowsHtml = rows.map((row, index) => `
    <tr>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${index + 1}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: left; color: #000000; font-weight: 500;">${row.name}</td>
      <td style="padding: 5px 6px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.state}</td>
    </tr>
  `).join('');

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 800; color: #000000; text-transform: uppercase;">DAILY SITE VISIT REPORT</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 18px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${dateLabel}</div>
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

// --- Image Generation: Summary ---

async function generateDailySummaryImage(siteName: string, summaryStats: Record<string, number>, dateLabel: string): Promise<string> {
  const container = document.createElement('div');
  
  // Smaller width for the summary card
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '400px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: 'sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });

  const total = Object.values(summaryStats).reduce((a, b) => a + b, 0);

  const summaryRowsHtml = Object.entries(summaryStats).map(([state, count]) => `
    <tr>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: left; color: #000000;">${state}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000;">${count}</td>
    </tr>
  `).join('');

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 800; color: #000000; text-transform: uppercase;">DAILY LEAD STATUS SUMMARY</div>
        <div style="width: 100px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 16px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 100px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${dateLabel}</div>
      </div>
      <div style="height: 15px;"></div>
      <table style="width: 100%; border-collapse: collapse; background-color: #ffffff;">
        <thead>
          <tr>
            <th style="padding: 8px 12px; text-align: left; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">State</th>
            <th style="padding: 8px 12px; text-align: center; border: 1px solid #000000; font-size: 12px; font-weight: 700; color: #000000; background-color: #f3f4f6;">Count</th>
          </tr>
        </thead>
        <tbody>
          ${summaryRowsHtml}
          <tr>
             <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: right; font-weight: 700; color: #000000; background-color: #f9fafb;">Total</td>
             <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; font-weight: 700; color: #000000; background-color: #f9fafb;">${total}</td>
          </tr>
        </tbody>
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

export async function processDailySiteVisitFile(file: File): Promise<ProcessResponse> {
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
        let nameIdx = -1, stateIdx = -1, assignedToIdx = -1, visitDateIdx = -1;

        const nameAliases = ['name', 'visitor name', 'lead name', 'customer name'];
        const stateAliases = ['lead state', 'state', 'region', 'location'];
        const assignedAliases = ['assigned to', 'assigned_to', 'owner', 'agent'];
        const dateAliases = ['visit date', 'visit_date', 'date of visit', 'date', 'visited date'];

        for (let i = 0; i < Math.min(50, rawRows.length); i++) {
          const row = rawRows[i];
          if (!Array.isArray(row)) continue;

          const nIdx = row.findIndex(c => c && nameAliases.some(a => String(c).toLowerCase().trim() === a));
          const sIdx = row.findIndex(c => c && stateAliases.some(a => String(c).toLowerCase().trim() === a));
          const aIdx = row.findIndex(c => c && assignedAliases.some(a => String(c).toLowerCase().trim() === a));
          const dIdx = row.findIndex(c => c && dateAliases.some(a => String(c).toLowerCase().trim() === a));

          if (nIdx !== -1 && aIdx !== -1) {
            headerIndex = i;
            nameIdx = nIdx;
            stateIdx = sIdx;
            assignedToIdx = aIdx;
            visitDateIdx = dIdx;
            break;
          }
        }

        if (headerIndex === -1) throw new Error("Could not find required columns (Name, Assigned To, etc).");

        // --- EXTRACT DATE FROM FIRST ROW ---
        // We do this immediately to find the report date
        let reportDateStr = "";
        if (visitDateIdx !== -1) {
            // Find the first non-empty date starting from data rows
            for (let i = headerIndex + 1; i < rawRows.length; i++) {
                const row = rawRows[i];
                if (!row) continue;
                const val = row[visitDateIdx];
                const formatted = parseAndFormatDate(val);
                if (formatted) {
                    reportDateStr = formatted;
                    break; 
                }
            }
        }

        // Process Data
        const normalizedMapping: Record<string, string> = {};
        Object.keys(USER_PROJECT_MAPPING).forEach(k => {
          normalizedMapping[k.toLowerCase().trim()] = USER_PROJECT_MAPPING[k];
        });

        const sites: Record<string, any[]> = {};

        for (let i = headerIndex + 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0) continue;

          const rawAssigned = row[assignedToIdx];
          if (!rawAssigned) continue;
          
          const assignedStr = String(rawAssigned).trim();
          const assignedLower = assignedStr.toLowerCase();

          // Fuzzy match / Check mapping
          const matchedUserKey = Object.keys(normalizedMapping).find(k => {
            return assignedLower === k || assignedLower.includes(k) || k.includes(assignedLower);
          });

          if (!matchedUserKey) continue;

          const siteName = normalizedMapping[matchedUserKey];
          const name = row[nameIdx] ? String(row[nameIdx]).trim() : '-';
          
          // --- Filter: Ignore "Test" names ---
          if (name.toLowerCase() === 'test') continue;

          let state = (stateIdx !== -1 && row[stateIdx]) ? String(row[stateIdx]).trim() : '-';
          
          // --- Format: Handle "re_visit_done" ---
          if (state.toLowerCase() === 're_visit_done') {
            state = 'Revisit Done';
          }

          // --- Filter: Ignore "Booked" state ---
          if (state.toLowerCase() === 'booked') continue;

          if (!sites[siteName]) sites[siteName] = [];
          
          sites[siteName].push({
            name,
            state
          });
        }

        const images: GeneratedImage[] = [];
        const zip = new JSZip();
        const siteKeys = Object.keys(sites);

        if (siteKeys.length === 0) throw new Error("No matching records found based on Project Mapping.");

        for (const site of siteKeys) {
          const rows = sites[site];
          
          // Sort rows by State (Alphabetical)
          rows.sort((a, b) => a.state.localeCompare(b.state));

          // Calculate Summary Stats (State -> Count)
          const summaryStats: Record<string, number> = {};
          rows.forEach(r => {
             const s = r.state;
             summaryStats[s] = (summaryStats[s] || 0) + 1;
          });

          // 1. Generate Daily List Image
          // We pass reportDateStr (extracted from the first valid row) as the dateLabel
          const listDataUrl = await generateDailyListImage(site, rows, reportDateStr);
          const listFilename = `${site.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_daily_visit.png`;
          images.push({ project_name: site, image_url: listDataUrl, filename: listFilename });
          zip.file(listFilename, listDataUrl.split(',')[1], { base64: true });

          // 2. Generate Daily Summary Image
          const summaryDataUrl = await generateDailySummaryImage(site, summaryStats, reportDateStr);
          const summaryFilename = `${site.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_daily_summary.png`;
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