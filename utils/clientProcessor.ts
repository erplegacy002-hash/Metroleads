import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING } from './projectMapping';

// --- Configuration ---

// Using absolute paths starting with / ensures Vite/Vercel finds them in the public folder
const PROJECT_LOGOS: Record<string, string> = {
  "Aqua Life": "/aqualife.png",
  "Kairos": "/kairos.png",
  "Statement": "/statement.png",
  "Milestone": "/milestone.png"
};

const logoDataCache: Record<string, string> = {};

// --- Logic Helpers ---

async function getBase64FromUrl(url: string): Promise<string> {
  if (logoDataCache[url]) return logoDataCache[url];
  
  try {
    // Ensure we fetch from the current origin for local assets
    const fetchUrl = url.startsWith('http') ? url : `${window.location.origin}${url.startsWith('/') ? '' : '/'}${url}`;
    const response = await fetch(fetchUrl);
    
    if (!response.ok) {
      console.warn(`Could not find logo file: ${fetchUrl} (Status: ${response.status})`);
      return "";
    }
    
    const blob = await response.blob();
    const base64: string = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });

    if (base64 && base64.length > 100) {
      logoDataCache[url] = base64;
      return base64;
    }
    return "";
  } catch (err) {
    console.warn(`Failed to load logo: ${url}. Proceeding with text fallback.`, err);
    return "";
  }
}

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

function formatClockDuration(totalSeconds: number): string {
  if (totalSeconds < 3600) {
    const totalMinutes = Math.floor(totalSeconds / 60);
    return `${totalMinutes} mins`;
  }

  let hours = totalSeconds / 3600;
  let rounded = Number(hours.toFixed(2));
  
  let intPart = Math.floor(rounded);
  let decPart = rounded - intPart;

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
    position: 'fixed',
    top: '0',
    left: '0',
    width: '1280px', 
    backgroundColor: '#ffffff', 
    padding: '40px',
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });
  
  if (rows.length === 0) return '';

  const displayHeaders = [
    "User Name", 
    "Answered", 
    "Call Duration (Answered)", 
    "Missed", 
    "Call Duration (Missed)", 
    "Total Call Duration", 
    "Total Call Count",
    "Average Call"
  ];

  const headerHtml = displayHeaders.map(h => {
    let formattedHeader = h;
    if (h.includes('(') && h.includes(')')) {
      formattedHeader = h.replace(' (', '<br/>(');
    }

    return `<th style="padding: 14px 10px; text-align: right; border: 1px solid #000000; font-size: 16px; font-weight: 700; color: #000000; background-color: #f3f4f6; vertical-align: bottom; line-height: 1.2; font-family: sans-serif;">
      ${h === "User Name" ? '<div style="text-align: left;">' + formattedHeader + '</div>' : formattedHeader}
    </th>`;
  }).join('');

  const rowsHtml = rows.map((row) => {
    return `
      <tr>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: left; color: #000000; font-weight: 500; font-family: sans-serif;">${row['User Name']}</td>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: right; color: #000000; font-family: sans-serif;">${row['Answered']}</td>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: right; color: #000000; font-family: sans-serif;">${row['Call Duration (Answered)']}</td>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: right; color: #000000; font-family: sans-serif;">${row['Missed']}</td>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: right; color: #000000; font-family: sans-serif;">${row['Call Duration (Missed)']}</td>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: right; color: #000000; font-weight: 500; font-family: sans-serif;">${row['Total Call Duration']}</td>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: right; color: #000000; font-family: sans-serif;">${row['Total Call Count']}</td>
        <td style="padding: 12px 10px; border: 1px solid #000000; font-size: 15px; text-align: right; color: #000000; font-family: sans-serif;">${row['Average Call']}</td>
      </tr>
    `;
  }).join('');

  const now = new Date();
  const timestamp = now.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  }) + ' ' + now.toLocaleTimeString('en-GB', {
    hour: '2-digit',
    minute: '2-digit'
  });

  const localLogoPath = PROJECT_LOGOS[siteName] || "";
  const logoBase64 = localLogoPath ? await getBase64FromUrl(localLogoPath) : "";

  // Maintaining requested logo sizes
  const logoHeight = (siteName === "Milestone" || siteName === "Kairos") ? '70px' : '65px';

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 2px solid #000000; box-sizing: border-box;">
      <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #000000; padding: 25px 35px; background-color: #000000;">
        <h2 style="margin: 0; font-size: 26px; font-weight: 800; color: #ffffff; text-transform: uppercase; text-align: left; flex: 1; letter-spacing: 0.5px; font-family: sans-serif;">CALLING REPORT AS ON ${timestamp}</h2>
        <div style="flex-shrink: 0; display: flex; align-items: center; justify-content: flex-end; padding-left: 20px;">
          ${logoBase64 
            ? `<img src="${logoBase64}" alt="${siteName}" style="height: ${logoHeight}; width: auto; max-width: 400px; object-fit: contain; background-color: #000000; padding: 6px; border-radius: 4px;" />` 
            : `<div style="color: #ffffff; font-size: 24px; font-weight: 900; padding: 10px 20px; border: 3px solid #ffffff; border-radius: 4px; letter-spacing: 1px; font-family: sans-serif;">${siteName.toUpperCase()}</div>`
          }
        </div>
      </div>
      <table style="width: 100%; border-collapse: collapse; background-color: #ffffff;">
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
  
  // Wait for images and fonts to be ready
  const imgs = container.getElementsByTagName('img');
  const imagePromises = Array.from(imgs).map(img => {
    if (img.complete) return Promise.resolve();
    return new Promise((resolve) => {
      img.onload = resolve;
      img.onerror = resolve;
    });
  });

  // Ensure fonts are loaded so table sizing is accurate
  if ('fonts' in document) {
    await (document as any).fonts.ready;
  }
  
  await Promise.all(imagePromises);
  
  // Extra buffer for Vercel/Headless rendering environments
  await new Promise(resolve => setTimeout(resolve, 1000));

  try {
    const dataUrl = await toPng(container, { 
      quality: 0.95, 
      pixelRatio: 2, // Higher pixel ratio for crisp text
      backgroundColor: '#ffffff', 
      cacheBust: true,
      skipFonts: false
    });
    
    if (!dataUrl || dataUrl.length < 5000) throw new Error("Blank capture detected.");
    return dataUrl;
  } catch (err: any) {
    console.error("Capture Error:", err);
    // Fallback attempt
    return await toPng(container);
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
          throw new Error("Columns not found. Ensure 'Assigned To' and 'Disposition' are present.");
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
        if (siteNames.length === 0) throw new Error("No data found for the predefined users.");

        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        for (const site of siteNames) {
          const userStats = sites[site];
          const rows: any[] = [];

          Object.keys(userStats).forEach(user => {
            const stat = userStats[user];
            const totalSeconds = stat.durationAnsweredRaw + stat.durationMissedRaw;

            const clockAnswered = formatClockDuration(stat.durationAnsweredRaw);
            const clockMissed = formatClockDuration(stat.durationMissedRaw);
            const clockTotal = formatClockDuration(totalSeconds);
            
            const totalCount = stat.answered + stat.missed;
            const minutesAnswered = stat.durationAnsweredRaw / 60;
            const avgCallMinutes = stat.answered > 0 ? (minutesAnswered / stat.answered) : 0;
            
            rows.push({
              "User Name": `User - ${user}`,
              "Answered": stat.answered,
              "Call Duration (Answered)": clockAnswered,
              "Missed": stat.missed,
              "Call Duration (Missed)": clockMissed,
              "Total Call Duration": clockTotal,
              "Total Call Count": totalCount,
              "Average Call": avgCallMinutes.toFixed(2),
              "rawTotalSeconds": totalSeconds
            });
          });

          rows.sort((a, b) => b.rawTotalSeconds - a.rawTotalSeconds);

          const dataUrl = await generateTableImage(site, rows);
          if (!dataUrl) continue;

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