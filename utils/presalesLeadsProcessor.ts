import { read, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { jsPDF } from "jspdf";
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING, DEFAULT_SITE } from './projectMapping';

// --- Helpers ---

function determineSource(cpData: any, sourceData: any, subSourceData: any): string {
  // 1. If CP Firm Name exists and is not empty/hyphen, it is a Channel Partner
  if (cpData && String(cpData).trim().length > 0 && String(cpData).trim() !== '-') {
    return 'Channel Partner';
  }

  const rawSource = sourceData ? String(sourceData).trim() : '';
  const rawSubSource = subSourceData ? String(subSourceData).trim() : '';

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

  // 4. Fallback
  if (rawSource.length > 0) return rawSource;
  return '-';
}

// --- Image Generation ---

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
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '1000px', 
    backgroundColor: '#ffffff', 
    padding: '20px', 
    fontFamily: 'sans-serif',
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });
  
  if (rows.length === 0) return '';

  const headerHtml = `
    <th style="padding: 8px 10px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 50px;">Sr. No.</th>
    <th style="padding: 8px 10px; text-align: left; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6;">User (Assigned to)</th>
    <th style="padding: 8px 10px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 120px;">Lead State</th>
    <th style="padding: 8px 10px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 220px;">Page Name</th>
    <th style="padding: 8px 10px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 140px;">Lead Source</th>
    <th style="padding: 8px 10px; text-align: center; border: 1px solid #000000; font-size: 11px; font-weight: 700; color: #000000; background-color: #f3f4f6; width: 80px;">Lead Count</th>
  `;

  const rowsHtml = rows.map((row, index) => `
    <tr>
      <td style="padding: 6px 10px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${index + 1}</td>
      <td style="padding: 6px 10px; border: 1px solid #000000; font-size: 11px; text-align: left; color: #000000; font-weight: 500;">${row.user}</td>
      <td style="padding: 6px 10px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.state}</td>
      <td style="padding: 6px 10px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000; word-break: break-word;">${row.pageName}</td>
      <td style="padding: 6px 10px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000;">${row.source}</td>
      <td style="padding: 6px 10px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000; font-weight: 700;">${row.count}</td>
    </tr>
  `).join('');

  // Calculate total leads
  const totalLeads = rows.reduce((acc, curr) => acc + curr.count, 0);

  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 15px 20px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 14px; font-weight: 800; color: #000000; text-transform: uppercase;">PRESALES LEADS</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 8px auto;"></div>
        <div style="font-size: 18px; font-weight: 900; color: #000000; text-transform: uppercase;">${siteName}</div>
        <div style="width: 150px; height: 1px; background-color: #000000; margin: 8px auto;"></div>
        <div style="font-size: 12px; font-weight: 700; color: #000000;">${reportTitle}</div>
      </div>

      <div style="height: 10px;"></div>

      <table style="width: 100%; border-collapse: collapse; background-color: #ffffff;">
        <thead><tr>${headerHtml}</tr></thead>
        <tbody>
          ${rowsHtml}
          <tr>
            <td colspan="5" style="padding: 8px 10px; border: 1px solid #000000; font-size: 11px; text-align: right; color: #000000; font-weight: 800; background-color: #f9fafb;">Total</td>
            <td style="padding: 8px 10px; border: 1px solid #000000; font-size: 11px; text-align: center; color: #000000; font-weight: 800; background-color: #f9fafb;">${totalLeads}</td>
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
          
          const normalizedSource = determineSource(cpData, leadSourceData, subSourceData);

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

        // Generate Image from HTML
        const imgDataUrl = await generatePresalesLeadsImage(siteName, rows, reportTitle);
        
        // Convert to PDF
        const img = new Image();
        img.src = imgDataUrl;
        await new Promise((r) => { img.onload = r; });
        
        const pdf = new jsPDF({
            orientation: img.width > img.height ? 'l' : 'p',
            unit: 'px',
            format: [img.width, img.height]
        });
        
        pdf.addImage(imgDataUrl, 'PNG', 0, 0, img.width, img.height);
        const pdfDataUrl = pdf.output('datauristring');

        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        const filename = `presales_leads_summary.pdf`;
        
        // Push the PDF Data URL. 
        // Note: The frontend ImageGallery needs to handle PDF icons for .pdf extension (already handled in previous code)
        images.push({ project_name: "Presales Leads Summary", image_url: pdfDataUrl, filename: filename });
        zip.file(filename, pdfDataUrl.split(',')[1], { base64: true });

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        resolve({ images, zip_url: URL.createObjectURL(zipBlob), message: "Success" });

      } catch (err: any) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}