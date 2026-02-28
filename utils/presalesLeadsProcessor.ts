import { read, utils } from 'xlsx';
import JSZip from 'jszip';
import { toPng } from 'html-to-image';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING, DEFAULT_SITE } from './projectMapping';

// --- Helpers ---

function determineSource(cpData: any, sourceData: any, subSourceData: any, pageNameData: any): string {
  // 1. If CP Firm Name exists and is not empty/hyphen, it is a Channel Partner
  if (cpData && String(cpData).trim().length > 0 && String(cpData).trim() !== '-') {
    return 'Channel Partner';
  }

  const rawSource = sourceData ? String(sourceData).trim() : '';
  const rawSubSource = subSourceData ? String(subSourceData).trim() : '';
  const rawPageName = pageNameData ? String(pageNameData).trim() : '';

  const checkKeywords = (str: string): string | null => {
    const s = str.toLowerCase();
    if (s.includes('channel partner') || s.includes('cp')) return 'Channel Partner';
    if (s.includes('walk-in') || s.includes('walkin') || s.includes('walk in')) return 'Walk-In';
    if (s.includes('digital') || s.includes('facebook') || s.includes('instagram') || s.includes('website') || s.includes('google') || s.includes('whatsapp') || s.includes('popup') || s.includes('online')) return 'Digital';
    if (s.includes('offer')) return 'Offer';
    if (s.includes('refer') || s.includes('referral') || s.includes('reference')) return 'Referral';
    if (s.includes('hoarding') || s.includes('hoardings') || s.includes('signage') || s.includes('board') || s.includes('site branding')) return 'Hoarding';
    if (s.includes('data') || s.includes('database') || s.includes('cold call')) return 'Data Calling';
    return null;
  };

  // 2. Check Lead Source first
  if (rawSource.length > 0) {
    const match = checkKeywords(rawSource);
    if (match) return match;
  }

  // 3. Check Sub Source second
  if (rawSubSource.length > 0) {
    const match = checkKeywords(rawSubSource);
    if (match) return match;
  }

  // 4. Check Page Name third
  if (rawPageName.length > 0) {
    const match = checkKeywords(rawPageName);
    if (match) return match;
  }

  // 5. Fallback
  if (rawSource.length > 0) return rawSource;
  return '-';
}

// --- PDF Generation ---

interface AggregatedLead {
  user: string;
  state: string;
  pageName: string;
  source: string;
  count: number;
}

async function generatePresalesLeadsImage(
  siteName: string, 
  rows: AggregatedLead[], 
  reportTitle: string
): Promise<string> {
  const container = document.createElement('div');
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.top = '0';
  container.style.width = '1000px'; // Fixed width for consistent output
  container.style.backgroundColor = '#ffffff';

  const totalLeads = rows.reduce((acc, curr) => acc + curr.count, 0);

  const rowsHtml = rows.map((row, index) => `
    <tr>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${index + 1}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 11px; text-align: left; color: #000000;">${row.user}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.state}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.pageName}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.source}</td>
      <td style="padding: 8px 12px; border: 1px solid #000000; font-size: 11px; text-align: center; font-weight: 700; color: #000000;">${row.count}</td>
    </tr>
  `).join('');

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box; font-family: 'Cormorant Garamond', serif;">
      <div style="padding: 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 900; color: #000000; text-transform: uppercase;">PRESALES LEADS</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 18px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 6px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${reportTitle}</div>
      </div>
      
      <table style="width: 100%; border-collapse: collapse; background-color: #ffffff;">
        <thead>
          <tr style="background-color: #f3f4f6;">
            <th style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; color: #000000; font-weight: 900;">Sr. No.</th>
            <th style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: left; color: #000000; font-weight: 900;">User (Assigned to)</th>
            <th style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; color: #000000; font-weight: 900;">Lead State</th>
            <th style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; color: #000000; font-weight: 900;">Page Name</th>
            <th style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; color: #000000; font-weight: 900;">Lead Source</th>
            <th style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; color: #000000; font-weight: 900;">Lead Count</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
          <tr style="background-color: #f9fafb; font-weight: 900;">
            <td colspan="4" style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: right; color: #000000;">TOTAL</td>
            <td colspan="2" style="padding: 10px 12px; border: 1px solid #000000; font-size: 12px; text-align: center; color: #000000;">${totalLeads}</td>
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
        let pageNameIdx = -1;
        let cpFirmNameIdx = -1;
        let subSourceIdx = -1;

        const assignedAliases = ['assigned to', 'assigned_to', 'owner', 'agent', 'executive', 'sales executive', 'allocated to', 'sales person', 'sourcing manager', 'closing manager'];
        const sourceAliases = ['lead source', 'lead source (f)', 'source', 'source of lead', 'enquiry source'];
        const subSourceAliases = ['sub source', 'sub source (u)', 'sub_source', 'subsource'];
        const stateAliases = ['lead state', 'state', 'region', 'location'];
        const pageNameAliases = ['page name', 'page_name', 'page', 'campaign', 'campaign name', 'ad set name', 'ad set', 'form name'];
        const cpFirmAliases = ['cp firm name', 'cp firm name (v)', 'cp name', 'channel partner firm name'];

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
            pageNameIdx = findColumnIndex(row, pageNameAliases);
            cpFirmNameIdx = findColumnIndex(row, cpFirmAliases);
            subSourceIdx = findColumnIndex(row, subSourceAliases);
            break;
          }
        }

        if (headerIndex === -1) throw new Error("Could not find required columns (Assigned To, Lead Source/State).");

        // Aggregation: Key = User|State|PageName|NormalizedSource
        const aggregation: Record<string, number> = {};

        for (let i = headerIndex + 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0) continue;

          // Check for 'test' in any column
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
          // Normalize user name using the mapping keys if available
          const assignedLower = assignedStr.toLowerCase();
          const matchedKey = Object.keys(USER_PROJECT_MAPPING).find(k => k.toLowerCase() === assignedLower); 
          if (matchedKey) {
             userName = matchedKey;
          }

          // Extract other fields
          const leadState = leadStateIdx !== -1 && row[leadStateIdx] ? String(row[leadStateIdx]).trim() : '-';
          const pageName = pageNameIdx !== -1 && row[pageNameIdx] ? String(row[pageNameIdx]).trim() : '-';
          
          // Determine Source using Keyword Logic
          const leadSourceData = leadSourceIdx !== -1 ? row[leadSourceIdx] : null;
          const cpData = cpFirmNameIdx !== -1 ? row[cpFirmNameIdx] : null;
          const subSourceData = subSourceIdx !== -1 ? row[subSourceIdx] : null;
          const pageNameData = pageNameIdx !== -1 ? row[pageNameIdx] : null;
          
          const normalizedSource = determineSource(cpData, leadSourceData, subSourceData, pageNameData);

          // Create Key
          const key = `${userName}||${leadState}||${pageName}||${normalizedSource}`;

          if (!aggregation[key]) {
              aggregation[key] = 0;
          }
          aggregation[key]++;
        }

        const rows: AggregatedLead[] = [];

        // Convert map to array
        Object.entries(aggregation).forEach(([key, count]) => {
            const [user, state, pageName, source] = key.split('||');
            rows.push({
                user,
                state,
                pageName,
                source,
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

        // Generate Image directly
        const imageDataUrl = await generatePresalesLeadsImage(siteName, rows, reportTitle);
        
        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        const filename = `presales_leads_summary.png`;
        
        images.push({ project_name: "Presales Leads Summary", image_url: imageDataUrl, filename: filename });
        zip.file(filename, imageDataUrl.split(',')[1], { base64: true });

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        resolve({ images, zip_url: URL.createObjectURL(zipBlob), message: "Success" });

      } catch (err: any) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}