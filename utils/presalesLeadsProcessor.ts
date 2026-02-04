import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING, DEFAULT_SITE } from './projectMapping';

// --- Image Generation ---

interface AggregatedLead {
  user: string;
  source: string;
  group: string;
  state: string;
  count: number;
}

async function generatePresalesLeadsImage(
  siteName: string, 
  rows: AggregatedLead[], 
  reportTitle: string
): Promise<string> {
  const container = document.createElement('div');
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '800px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: 'sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });
  
  if (rows.length === 0) return '';

  const headerHtml = `
    <th style="padding: 6px 8px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 50px;">Sr. No.</th>
    <th style="padding: 6px 8px; text-align: left; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6;">User (Assigned to)</th>
    <th style="padding: 6px 8px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 120px;">Lead Source</th>
    <th style="padding: 6px 8px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 100px;">Lead Group</th>
    <th style="padding: 6px 8px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 120px;">Lead State</th>
    <th style="padding: 6px 8px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 80px;">Lead Count</th>
  `;

  const rowsHtml = rows.map((row, index) => `
    <tr>
      <td style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${index + 1}</td>
      <td style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: left; color: #000000; font-weight: 500;">${row.user}</td>
      <td style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.source}</td>
      <td style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.group}</td>
      <td style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.state}</td>
      <td style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000; font-weight: 700;">${row.count}</td>
    </tr>
  `).join('');

  // Calculate total leads
  const totalLeads = rows.reduce((acc, curr) => acc + curr.count, 0);

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 800; color: #000000; text-transform: uppercase;">PRESALES LEADS</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 18px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${reportTitle}</div>
      </div>

      <div style="height: 15px;"></div>

      <table style="width: 100%; border-collapse: collapse; background-color: #ffffff;">
        <thead><tr>${headerHtml}</tr></thead>
        <tbody>
          ${rowsHtml}
          <tr>
            <td colspan="5" style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: right; color: #000000; font-weight: 800; background-color: #f9fafb;">Total</td>
            <td style="padding: 5px 8px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000; font-weight: 800; background-color: #f9fafb;">${totalLeads}</td>
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

function findColumnIndex(row: any[], aliases: string[]): number {
  for (const alias of aliases) {
    const idx = row.findIndex(c => c && String(c).toLowerCase().trim() === alias);
    if (idx !== -1) return idx;
  }
  return -1;
}

export async function processPresalesLeadsFile(file: File): Promise<ProcessResponse> {
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
        let assignedToIdx = -1;
        let leadSourceIdx = -1;
        let leadStateIdx = -1;
        let leadGroupIdx = -1;

        const assignedAliases = ['assigned to', 'assigned_to', 'owner', 'agent', 'executive', 'sales executive', 'allocated to', 'sales person', 'sourcing manager', 'closing manager'];
        const sourceAliases = ['lead source', 'lead source (f)', 'source', 'source of lead', 'enquiry source'];
        const stateAliases = ['lead state', 'state', 'region', 'location'];
        const groupAliases = ['lead group', 'group', 'classification', 'category', 'quality', 'lead quality'];

        for (let i = 0; i < Math.min(100, rawRows.length); i++) {
          const row = rawRows[i];
          if (!Array.isArray(row)) continue;

          const aIdx = findColumnIndex(row, assignedAliases);
          const sIdx = findColumnIndex(row, sourceAliases);
          const stIdx = findColumnIndex(row, stateAliases);
          
          if (aIdx !== -1 && (sIdx !== -1 || stIdx !== -1)) {
            headerIndex = i;
            assignedToIdx = aIdx;
            leadSourceIdx = sIdx;
            leadStateIdx = stIdx;
            leadGroupIdx = findColumnIndex(row, groupAliases);
            break;
          }
        }

        if (headerIndex === -1) throw new Error("Could not find required columns (Assigned To, Lead Source/State).");

        // Aggregation: Key = User|Source|Group|State
        const aggregation: Record<string, number> = {};

        for (let i = headerIndex + 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0) continue;

          // Check for 'test'
          const isTestRow = row.some(cell => {
             if (!cell) return false;
             const s = String(cell).toLowerCase();
             return s.includes('test') || s.includes('testing');
          });
          if (isTestRow) continue;

          // Get User Name
          const rawAssigned = row[assignedToIdx];
          const assignedStr = rawAssigned ? String(rawAssigned).trim() : "Unassigned";
          
          let userName = assignedStr;
          // Normalize user name casing using the mapping keys if available
          const assignedLower = assignedStr.toLowerCase();
          const matchedKey = Object.keys(USER_PROJECT_MAPPING).find(k => k.toLowerCase() === assignedLower); 
          if (matchedKey) {
             userName = matchedKey;
          }

          // Extract other fields
          const leadSource = leadSourceIdx !== -1 && row[leadSourceIdx] ? String(row[leadSourceIdx]).trim() : '-';
          const leadState = leadStateIdx !== -1 && row[leadStateIdx] ? String(row[leadStateIdx]).trim() : '-';
          const leadGroup = leadGroupIdx !== -1 && row[leadGroupIdx] ? String(row[leadGroupIdx]).trim() : '-';

          // Create Key
          const key = `${userName}||${leadSource}||${leadGroup}||${leadState}`;

          if (!aggregation[key]) {
              aggregation[key] = 0;
          }
          aggregation[key]++;
        }

        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        const rows: AggregatedLead[] = [];

        // Convert map to array
        Object.entries(aggregation).forEach(([key, count]) => {
            const [user, source, group, state] = key.split('||');
            rows.push({
                user,
                source,
                group,
                state,
                count
            });
        });

        // Sort: User (Asc) -> Count (Desc)
        rows.sort((a, b) => {
            const userCompare = a.user.localeCompare(b.user);
            if (userCompare !== 0) return userCompare;
            return b.count - a.count;
        });

        if (rows.length === 0) throw new Error("No valid records found.");

        const reportTitle = "SUMMARY REPORT";
        const siteName = "Consolidated Leads";

        // Generate Image
        const imgDataUrl = await generatePresalesLeadsImage(siteName, rows, reportTitle);
        const filename = `presales_leads_summary.png`;
        
        images.push({ project_name: "Presales Leads Summary", image_url: imgDataUrl, filename: filename });
        zip.file(filename, imgDataUrl.split(',')[1], { base64: true });

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        resolve({ images, zip_url: URL.createObjectURL(zipBlob), message: "Success" });

      } catch (err: any) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}