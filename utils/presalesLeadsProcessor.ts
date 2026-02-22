import { read, utils } from 'xlsx';
import JSZip from 'jszip';
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
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

async function generatePresalesLeadsPDF(
  siteName: string, 
  rows: AggregatedLead[], 
  reportTitle: string
): Promise<string> {
  const doc = new jsPDF();

  // --- Header ---
  const pageWidth = doc.internal.pageSize.width;
  
  // Title
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.text("PRESALES LEADS", pageWidth / 2, 15, { align: "center" });
  
  doc.setLineWidth(0.5);
  doc.line(pageWidth / 2 - 25, 18, pageWidth / 2 + 25, 18);
  
  // Site Name
  doc.setFontSize(18);
  doc.setFont("helvetica", "bold");
  doc.text(siteName.toUpperCase(), pageWidth / 2, 26, { align: "center" });

  doc.setLineWidth(0.5);
  doc.line(pageWidth / 2 - 25, 29, pageWidth / 2 + 25, 29);

  // Report Title
  doc.setFontSize(12);
  doc.setFont("helvetica", "bold");
  doc.text(reportTitle, pageWidth / 2, 36, { align: "center" });

  // --- Table Data ---
  
  // Calculate total leads
  const totalLeads = rows.reduce((acc, curr) => acc + curr.count, 0);

  const tableBody = rows.map((row, index) => [
    index + 1,
    row.user,
    row.state,
    row.pageName,
    row.source,
    row.count
  ]);

  // Add Total Row
  tableBody.push([
    "",
    "",
    "",
    "",
    "TOTAL",
    totalLeads
  ]);

  // --- Generate Table ---
  autoTable(doc, {
    startY: 42,
    head: [['Sr. No.', 'User (Assigned to)', 'Lead State', 'Page Name', 'Lead Source', 'Lead Count']],
    body: tableBody,
    theme: 'grid',
    styles: {
        fontSize: 10,
        font: "helvetica",
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.1,
    },
    headStyles: {
      fillColor: [243, 244, 246], // #f3f4f6
      textColor: [0, 0, 0],
      fontStyle: 'bold',
      halign: 'center',
      valign: 'middle',
    },
    columnStyles: {
      0: { halign: 'center', cellWidth: 20 },
      1: { halign: 'left' },
      2: { halign: 'center' },
      3: { halign: 'center' },
      4: { halign: 'center' },
      5: { halign: 'center', fontStyle: 'bold', cellWidth: 30 }
    },
    didParseCell: function (data: any) {
        // Identify the last row (Total row)
        if (data.row.index === tableBody.length - 1) {
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.fillColor = [249, 250, 251]; // #f9fafb
            
            // Align "TOTAL" label to the right
            if (data.column.index === 4) { 
               data.cell.styles.halign = 'right';
            }
        }
    }
  });

  return doc.output('datauristring');
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

        // Generate PDF directly
        const pdfDataUrl = await generatePresalesLeadsPDF(siteName, rows, reportTitle);
        
        const images: GeneratedImage[] = [];
        const zip = new JSZip();

        const filename = `presales_leads_summary.pdf`;
        
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