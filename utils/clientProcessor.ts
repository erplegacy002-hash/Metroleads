import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING } from './projectMapping';

// --- Logic Helpers ---

function isAnswered(status: string): boolean {
  const s = status.toLowerCase();
  
  if (s.includes('not ') || 
      s.includes('fail') || 
      s.includes('busy') || 
      s.includes('voice') || 
      s.includes('missed') || 
      s.includes('wrong') || 
      s.includes('invalid') ||
      s.includes('switched off') ||
      s.includes('not reachable') ||
      s.includes('no answer')) {
    return false;
  }

  return s.includes('connected') || 
         s.includes('talked') || 
         s.includes('answer') || 
         s.includes('converted') || 
         s.includes('sale') || 
         s.includes('interested') ||
         s.includes('visit') || 
         s.includes('meeting') || 
         s.includes('demo') || 
         s.includes('deal') || 
         s.includes('follow');
}

function parseDurationRaw(val: any): number {
  if (!val) return 0;
  if (typeof val === 'number') return val;
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
 * Custom formatting logic as requested:
 * 1. If total duration is less than 1 hour (<3600s), show in minutes (e.g., "30 mins").
 * 2. If 1 hour or more, use decimal hours but apply "Clock Logic":
 *    If the decimal part is >= 0.60, treat every 0.60 as an additional 1.00 hour.
 *    Example: 1.60 standard becomes 2.00 clock.
 */
function formatClockDuration(totalSeconds: number): string {
  if (totalSeconds < 3600) {
    const totalMinutes = Math.floor(totalSeconds / 60);
    return `${totalMinutes} mins`;
  }

  let hours = totalSeconds / 3600;
  // Convert to 2-decimal rounded number for the "Clock Logic" processing
  let rounded = Number(hours.toFixed(2));
  
  let intPart = Math.floor(rounded);
  let decPart = rounded - intPart;

  // Rollover logic: if decimal part is >= 0.60, carry over to the hour
  // Using decPart > 0.59 to handle floating point precision for exactly 0.60
  if (decPart > 0.599) {
    const extraHours = Math.floor(decPart / 0.6);
    const remainder = decPart % 0.6;
    rounded = intPart + extraHours + remainder;
  }

  return rounded.toFixed(2) + " hrs";
}

// --- Image Generation ---

async function generateTableImage(siteName: string, rows: any[]): Promise<string> {
  const container = document.createElement('div');
  
  Object.assign(container.style, {
    position: 'absolute',
    top: '0',
    left: '0',
    zIndex: '-1000', 
    width: '1450px',
    backgroundColor: '#ffffff',
    padding: '40px',
    fontFamily: "Arial, Helvetica, sans-serif", 
    color: '#000000',
    display: 'block',
    visibility: 'visible'
  });
  
  if (rows.length === 0) return '';

  const displayHeaders = [
    "User Name", 
    "Answered", 
    "Call Duration (Answered)", 
    "Missed", 
    "Call Duration (Missed)", 
    "Total Call Duration", 
    "Total Count",
    "Average Call"
  ];

  const headerHtml = displayHeaders.map(h => {
    let formattedHeader = h;
    if (h.includes('(') && h.includes(')')) {
      formattedHeader = h.replace(' (', '<br/>(');
    }

    return `<th style="padding: 12px 10px; text-align: right; border: 1px solid #000; font-size: 15px; font-weight: 600; color: #000; background-color: #fff; vertical-align: bottom;">
      ${h === "User Name" ? '<div style="text-align: left;">' + formattedHeader + '</div>' : formattedHeader}
    </th>`;
  }).join('');

  const rowsHtml = rows.map((row) => {
    return `
      <tr>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: left; color: #000;">${row['User Name']}</td>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: right; color: #000;">${row['Answered']}</td>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: right; color: #000;">${row['Call Duration (Answered)']}</td>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: right; color: #000;">${row['Missed']}</td>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: right; color: #000;">${row['Call Duration (Missed)']}</td>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: right; color: #000;">${row['Total Call Duration']}</td>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: right; color: #000;">${row['Total Count']}</td>
        <td style="padding: 8px 10px; border: 1px solid #000; font-size: 15px; text-align: right; color: #000;">${row['Average Call']}</td>
      </tr>
    `;
  }).join('');

  container.innerHTML = `
    <div style="background-color: white; color: black; font-family: Arial, sans-serif;">
      <div style="text-align: center; border: 1px solid #000; border-bottom: none; padding: 15px; background-color: #fff;">
        <h2 style="margin: 0; font-size: 24px; font-weight: normal;">Today's Calling Report - ${siteName}</h2>
      </div>
      <table style="width: 100%; border-collapse: collapse; text-align: right;">
        <thead>
          <tr>${headerHtml}</tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>
    </div>
  `;

  document.body.appendChild(container);
  await new Promise(resolve => setTimeout(resolve, 200));

  try {
    const dataUrl = await toPng(container, { 
      quality: 1.0, 
      pixelRatio: 2,
      backgroundColor: '#ffffff',
      skipAutoScale: true,
      cacheBust: true,
      fontEmbedCSS: '', 
      filter: (node) => node.tagName !== 'LINK'
    });
    return dataUrl;
  } catch (err) {
    console.error("Image generation failed:", err);
    return '';
  } finally {
    if (document.body.contains(container)) {
      document.body.removeChild(container);
    }
  }
}

// --- Main Processor ---

export async function processFile(file: File): Promise<ProcessResponse> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        const rawRows = utils.sheet_to_json(sheet, { header: 1 }) as any[][];
        if (!rawRows || rawRows.length === 0) throw new Error("Excel file appears to be empty.");

        let headerIndex = -1;
        let projectKeyIndex = -1; 
        let dispositionKeyIndex = -1;
        let durationKeyIndex = -1;

        const possibleProjectNames = ['assigned to', 'user', 'agent', 'project name'];
        const possibleDurationNames = ['duration', 'talk time', 'bill sec', 'call duration', 'time'];

        for (let i = 0; i < Math.min(20, rawRows.length); i++) {
           const row = rawRows[i];
           if (!Array.isArray(row)) continue;

           const pIdx = row.findIndex(cell => cell && typeof cell === 'string' && possibleProjectNames.some(t => cell.toLowerCase().trim() === t || cell.toLowerCase().includes(t)));
           let dIdx = row.findIndex(cell => cell && typeof cell === 'string' && cell.toLowerCase().includes('disposition'));
           if (dIdx === -1) {
              dIdx = row.findIndex(cell => cell && typeof cell === 'string' && (cell.toLowerCase().includes('call result') || cell.toLowerCase().trim() === 'status'));
           }

           if (pIdx !== -1 && dIdx !== -1) {
             headerIndex = i;
             projectKeyIndex = pIdx;
             dispositionKeyIndex = dIdx;
             durationKeyIndex = row.findIndex(cell => cell && typeof cell === 'string' && possibleDurationNames.some(t => cell.toLowerCase().includes(t)));
             break;
           }
        }

        if (headerIndex === -1) {
          throw new Error("Could not find required columns: 'Assigned To' and 'Disposition'.");
        }

        const jsonData = utils.sheet_to_json(sheet, { range: headerIndex, raw: false, defval: "" });
        const headerRow = rawRows[headerIndex];
        const projectKey = String(headerRow[projectKeyIndex]).trim();
        const dispositionKey = String(headerRow[dispositionKeyIndex]).trim();
        const durationKey = durationKeyIndex !== -1 ? String(headerRow[durationKeyIndex]).trim() : null;

        const sites: Record<string, Record<string, { answered: number, missed: number, durationAnsweredRaw: number, durationMissedRaw: number }>> = {};

        const normalizedMapping: Record<string, string> = {};
        Object.keys(USER_PROJECT_MAPPING).forEach(key => {
          normalizedMapping[key.toLowerCase()] = USER_PROJECT_MAPPING[key];
        });

        jsonData.forEach((row: any) => {
          if (!row) return;

          const rawUser = row[projectKey];
          if (!rawUser) return;
          const userName = String(rawUser).trim();
          
          const siteName = normalizedMapping[userName.toLowerCase()];
          if (!siteName) return; 

          const disposition = String(row[dispositionKey] || '').trim();
          const durationVal = durationKey ? row[durationKey] : 0;
          const durationRaw = parseDurationRaw(durationVal);

          if (!sites[siteName]) sites[siteName] = {};
          if (!sites[siteName][userName]) {
            sites[siteName][userName] = { 
              answered: 0, 
              missed: 0, 
              durationAnsweredRaw: 0, 
              durationMissedRaw: 0 
            };
          }

          const stats = sites[siteName][userName];
          
          if (isAnswered(disposition)) {
            stats.answered += 1;
            stats.durationAnsweredRaw += durationRaw;
          } else {
            stats.missed += 1;
            stats.durationMissedRaw += durationRaw;
          }
        });

        const siteNames = Object.keys(sites);
        if (siteNames.length === 0) throw new Error("No matching data found.");

        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        for (const site of siteNames) {
          const userStats = sites[site];
          const rows: any[] = [];

          Object.keys(userStats).forEach(user => {
            const stat = userStats[user];
            
            // Format durations with the custom "Clock & Unit Logic" requested
            const clockAnswered = formatClockDuration(stat.durationAnsweredRaw);
            const clockMissed = formatClockDuration(stat.durationMissedRaw);
            const clockTotal = formatClockDuration(stat.durationAnsweredRaw + stat.durationMissedRaw);
            
            const totalCount = stat.answered + stat.missed;
            
            // Average Call remains as standard minutes for calculation clarity
            const minutesAnswered = stat.durationAnsweredRaw / 60;
            const avgCallMinutes = stat.answered > 0 ? (minutesAnswered / stat.answered) : 0;
            
            rows.push({
              "User Name": `User - ${user}`,
              "Answered": stat.answered,
              "Call Duration (Answered)": clockAnswered,
              "Missed": stat.missed,
              "Call Duration (Missed)": clockMissed,
              "Total Call Duration": clockTotal,
              "Total Count": totalCount,
              "Average Call": avgCallMinutes.toFixed(2)
            });
          });

          rows.sort((a, b) => a["User Name"].localeCompare(b["User Name"]));

          const dataUrl = await generateTableImage(site, rows);
          const cleanName = site.replace(/[^a-z0-9]/gi, '_').toLowerCase();
          const filename = `${cleanName}_summary.png`;

          images.push({
            project_name: site,
            image_url: dataUrl,
            filename: filename
          });

          zip.file(filename, dataUrl.split(',')[1], { base64: true });
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const zipUrl = URL.createObjectURL(zipBlob);

        resolve({ images, zip_url: zipUrl, message: "Success" });

      } catch (err: any) {
        console.error("Processing Error:", err);
        reject(err);
      }
    };
    
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsArrayBuffer(file);
  });
}
